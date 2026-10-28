/**
 * dsh-teacher-kit — host half.
 *
 * 教师助手的宿主侧：注册 agent 工具 + 本地 HTTP 接口。
 *
 * 三条硬约束（都来自需求定案，改动前先读理由）：
 *  1. 纯本地：除模型对话外不联网、不上传任何文件内容。
 *  2. 绝不抛异常 —— 插件 apply() 抛错会让整个 Web 壳白屏，所以每处注册
 *     单独 try/catch，失败只降级、不中断。
 *  3. 零外部安装：文档读写走 DSH 自带 Python（python-docx/pptx/openpyxl），
 *     PDF 导出与逐页预览走 DSH 自带的 @deepseek-ai/libreoffice-kit。
 *
 * @module dsh-teacher-kit
 */
import { spawn, spawnSync } from 'node:child_process'
import { randomUUID } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { basename, dirname, join, resolve, isAbsolute } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { TASKS, TASK_KEYS, buildGuide, findTask } from './guide.js'

export const name = 'dsh-teacher-kit'
export const VERSION = '0.18.0'
/**
 * 不声明硬依赖：inject 为空时插件立即 mount，服务一出现就注册（见 withService）。
 * 若这里写死 ['webServer']，在没有 webServer 的组合里插件根本不会加载、面板静默消失。
 */
export const inject = []

const API = '/teacher-kit/api'
const LOCAL_HEADER = 'x-teacher-kit'

/**
 * 用 defineTool 包一下。能解析到平台包就用它；解析不到时**自己**做完该做的转换。
 *
 * 为什么不能简单退回字面量：真 registry 注册前会跑 assertSupportedJsonSchema，
 * schema.type 只认 object/array/string/number/integer/boolean/null。author 形态的
 * parameters（{inputPath:{type:'string',required:true}}）和 {type:'json'} 都会被拒，
 * 三个工具会全部注册失败且用户无感（见 reviews/step1-skeleton.md S2）。
 */
const JSON_SCHEMA_TYPES = new Set(['object', 'array', 'string', 'number', 'integer', 'boolean', 'null'])

function valueSpecToSchema(spec) {
  const value = spec && typeof spec === 'object' ? spec : {}
  if (!value.type || value.type === 'json') return {}
  if (!JSON_SCHEMA_TYPES.has(value.type)) return {}
  const schema = { type: value.type }
  if (typeof value.description === 'string' && value.description) schema.description = value.description
  if (Array.isArray(value.enum)) schema.enum = value.enum.slice()
  if (value.type === 'array' && value.items) schema.items = valueSpecToSchema(value.items)
  return schema
}

function parameterSpecToSchema(parameters) {
  const properties = {}
  const required = []
  for (const [key, spec] of Object.entries(parameters || {})) {
    properties[key] = valueSpecToSchema(spec)
    if (spec && spec.required) required.push(key)
  }
  const schema = { type: 'object', properties, additionalProperties: false }
  if (required.length > 0) schema.required = required
  return schema
}

/**
 * 真 defineTool 的参数校验只允许「有 required 就必须是 true」，
 * 写 `required: false` 会被判 unsupported JSON schema（真机三条工具全灭）。
 * 所以送进去之前统一把 required:false 摘掉。
 */
function normalizeParameters(parameters) {
  if (!parameters || typeof parameters !== 'object') return parameters
  const clean = {}
  for (const [key, spec] of Object.entries(parameters)) {
    if (spec && typeof spec === 'object') {
      const copy = { ...spec }
      if (copy.required === false) delete copy.required
      clean[key] = copy
    } else {
      clean[key] = spec
    }
  }
  return clean
}

function fallbackDefineTool(definition) {
  const output = definition.output && typeof definition.output === 'object' ? definition.output : {}
  const declared = output.schema
  const schema = declared && declared.type === 'json' ? {} : (declared && typeof declared === 'object' ? declared : {})
  return Object.assign({}, definition, {
    parameters: parameterSpecToSchema(definition.parameters),
    output: Object.assign({}, output, { schema })
  })
}

let defineTool = fallbackDefineTool
try {
  const mod = await import('@deepseek-ai/dsh-tools')
  if (mod && typeof mod.defineTool === 'function') defineTool = mod.defineTool
} catch { /* 平台包不可直接解析时用本地回退（形状等价） */ }

/** 送进 defineTool 之前的统一整形：参数摘掉 required:false。 */
function defineToolSafe(definition) {
  const input = { ...definition, parameters: normalizeParameters(definition.parameters) }
  return defineTool(input)
}

/** 工具注册结果记在这里，面板 /status 能看见"工具没挂上"，不再静默。 */
const registration = { tools: 0, routes: 0, issues: [] }

const renderJson = (_args, value) => [{ type: 'text', text: JSON.stringify(value, null, 2) }]

// ── Python 探测 ────────────────────────────────────────────────────────────

function pythonCandidates() {
  const home = homedir()
  const list = []
  if (process.env.DSH_PYTHON) list.push(process.env.DSH_PYTHON)
  list.push(join(home, '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'python.exe'))
  list.push(join(home, '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'bin', 'python3'))
  list.push('python')
  return list
}

/**
 * 探测一个可执行文件能不能跑。
 *
 * **必须是异步的**：以前用 spawnSync，面板点「重新体检」会把整个 Web 壳（也就是老师眼前
 * 整个聊天界面）冻住 ~0.85s，最坏（python 卡住）20 秒 —— 期间任何点击都没反应。
 */
function probeAsync(executable, args, timeout) {
  return new Promise((resolve) => {
    let settled = false
    let text = ''
    let child
    const finish = (value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolve(value)
    }
    const timer = setTimeout(() => {
      try { child?.kill() } catch { /* 已经退出 */ }
      finish(undefined)
    }, timeout)
    try {
      child = spawn(executable, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
    } catch {
      finish(undefined)
      return
    }
    const collect = (chunk) => { if (text.length < 4096) text += String(chunk) }
    child.stdout?.on('data', collect)
    child.stderr?.on('data', collect)
    child.on('error', () => finish(undefined))
    child.on('close', (code) => {
      const out = text.trim()
      if (code !== 0 && out === '') return finish(undefined)
      finish({ path: executable, version: out.split(/\r?\n/)[0].slice(0, 160) })
    })
  })
}

// 报的是真 Python 版本 + 三个库的版本（原来只 print 了 python-docx 的版本，面板上
// 会显示成 "Python 1.2.0"，老师看了以为环境不对）。
const PY_CHECK = 'import sys, docx, pptx, openpyxl; '
  + 'print("Python " + sys.version.split()[0] + " · python-docx " + docx.__version__ '
  + '+ " · python-pptx " + pptx.__version__ + " · openpyxl " + openpyxl.__version__)'

let runtimeCache = { at: 0, value: undefined }

/**
 * 找 Python。**这个探测必须有缓存**：调一次引擎要起两个子进程（探测 ~0.7s + 引擎 ~0.5s），
 * 以前每读一次画像都付两遍，老师点一下面板要等一秒多。
 * python 的位置在一次会话里不会变，所以正结果长缓存；负结果只缓存 10 秒
 * （刚装好 Python 的老师不用重启 DSH 就能用上）。force=true 走真探测（/status 强刷用）。
 */
let pythonCache = { at: 0, value: undefined }
const PYTHON_HIT_TTL = 10 * 60_000
const PYTHON_MISS_TTL = 10_000

let pythonPending

/**
 * 找 Python（异步）。**并发去重**：面板打开时 /status 与 /profile 会同时要 Python，
 * 不去重就同时起五六个 python 进程各跑一遍 import 检查（老师机器上更慢）。
 */
function pythonReport(force = false) {
  const now = Date.now()
  if (!force && pythonCache.at) {
    const ttl = pythonCache.value ? PYTHON_HIT_TTL : PYTHON_MISS_TTL
    if (now - pythonCache.at < ttl) return Promise.resolve(pythonCache.value)
  }
  if (pythonPending) return pythonPending
  pythonPending = (async () => {
    const at = Date.now()
    try {
      for (const candidate of pythonCandidates()) {
        if (candidate !== 'python' && !existsSync(candidate)) continue
        const found = await probeAsync(candidate, ['-c', PY_CHECK], 20_000)
        if (found) {
          pythonCache = { at, value: found }
          return found
        }
      }
      pythonCache = { at, value: undefined }
      return undefined
    } finally {
      pythonPending = undefined
    }
  })()
  return pythonPending
}

// ── LibreOffice kit（PDF 导出 + 逐页预览）──────────────────────────────────

/**
 * 找 libreoffice-kit 的入口文件。**必须先试 app.asar.unpacked**。
 *
 * 真机踩过：裸包名 '@deepseek-ai/libreoffice-kit' 能 import 成功（asar 里的 JS 可读），
 * 但它内部按自己的 __dirname 去找平台包里的 libreoffice-kit.exe，路径落在 app.asar 里
 * → spawn ...app.asar\...\libreoffice-kit.exe ENOENT，导出 PDF 直接失败。
 * unpacked 目录里才有 .exe（electron-builder 的 asarUnpack 规则），所以顺序反过来。
 */
function kitSpecifiers() {
  const list = []
  const bases = []
  if (process.resourcesPath) bases.push(process.resourcesPath)
  const exeDir = process.execPath ? dirname(process.execPath) : undefined
  if (exeDir) {
    bases.push(join(exeDir, 'resources'))
    bases.push(dirname(exeDir))
  }
  // 从纯 node 跑自检 / 从命令行调插件时，上面两条都指不到 DSH 的 resources。
  // 允许显式指路（自检用），再从环境变量里捞出任何指向 app.asar.unpacked 的路径反推资源根
  // （打包好的 DSH 会把原生模块的 unpacked 路径写进环境变量，比如 VIPSHOME）。
  for (const key of ['TEACHER_KIT_RESOURCES', 'DSH_RESOURCES', 'DSH_RESOURCES_PATH']) {
    const value = process.env[key]
    if (value) bases.push(value)
  }
  for (const value of Object.values(process.env)) {
    if (typeof value !== 'string' || !value.includes('app.asar.unpacked')) continue
    // Windows 会把这些路径写成 \\?\D:\...\ 的扩展长度形式，import 不认，先削掉
    const root = value.slice(0, value.indexOf('app.asar.unpacked'))
      .replace(/^\\\\\?\\/, '')
      .replace(/[\\/]+$/, '')
    if (root) bases.push(root)
  }
  for (const base of bases) {
    for (const mid of ['dsh' + '/node_modules', 'node_modules']) {
      list.push(join(base, 'app.asar.unpacked', mid, '@deepseek-ai', 'libreoffice-kit', 'lib', 'index.js'))
    }
  }
  // 最后才退回裸包名（开发态 / 从源码跑时走这条）
  list.push('@deepseek-ai/libreoffice-kit')
  return list
}

let kitCache = { state: 'idle', module: undefined, from: undefined, error: undefined }
// 失败也要能重试：启动那一秒的一次瞬时失败（asar 还没就绪、杀软锁文件）不该让 PDF 整个进程生命周期不可用（审查 P2-10）。
const KIT_FAILED_TTL = 60_000

async function loadKit() {
  if (kitCache.state === 'ready') return kitCache
  if (kitCache.state === 'failed' && Date.now() - (kitCache.at || 0) < KIT_FAILED_TTL) return kitCache
  for (const spec of kitSpecifiers()) {
    try {
      const target = spec.startsWith('/') || /^[A-Za-z]:/.test(spec) ? pathToFileURL(spec).href : spec
      const mod = await import(target)
      if (mod && typeof mod.createConverter === 'function') {
        kitCache = { state: 'ready', module: mod, from: spec, error: undefined }
        return kitCache
      }
    } catch { /* 换下一个候选 */ }
  }
  kitCache = { state: 'failed', at: Date.now(), module: undefined, from: undefined, error: '未找到 @deepseek-ai/libreoffice-kit' }
  return kitCache
}

let converterPromise
let converterHolder

/** 单例 converter：kit 每次转换都另起私有进程，复用实例是官方推荐用法。 */
async function getConverter() {
  if (converterPromise) return converterPromise
  converterPromise = (async () => {
    const kit = await loadKit()
    if (kit.state !== 'ready') throw new Error(kit.error || 'libreoffice-kit 不可用')
    const converter = await kit.module.createConverter({ timeoutMs: 120_000 })
    converterHolder = { converter }
    return converter
  })()
  try {
    return await converterPromise
  } catch (error) {
    converterPromise = undefined
    throw error
  }
}

async function disposeConverter() {
  const holder = converterHolder
  converterHolder = undefined
  converterPromise = undefined
  if (!holder) return
  try { await holder.converter.dispose() } catch { /* 卸载时尽力而为 */ }
}

/** 运行时体检（带缓存，默认 30s；/status 用 60s）；不启动 converter，只确认模块在。 */
async function runtimeReport(force, ttlMs = 30_000) {
  const now = Date.now()
  if (!force && runtimeCache.value && now - runtimeCache.at < ttlMs) return runtimeCache.value
  const python = await pythonReport(force)
  const kit = await loadKit()
  const value = {
    ok: Boolean(python) && kit.state === 'ready',
    version: VERSION,
    python: python ? { path: python.path, version: python.version } : null,
    libreoffice: kit.state === 'ready' ? { via: kit.from } : null,
    canConvertPdf: kit.state === 'ready',
    canPreviewPages: kit.state === 'ready',
    engine: { cli: ENGINE_CLI, present: existsSync(ENGINE_CLI) },
    surface: {
      toolsRegistered: registration.tools,
      routesRegistered: registration.routes,
      issues: registration.issues.slice()
    },
    hint: python ? undefined : '未找到 DSH 自带 Python：文档读写不可用。'
  }
  runtimeCache = { at: now, value }
  return value
}

// ── Python 引擎桥（engine/cli.py，单行 JSON 协议）────────────────────────────

const PLUGIN_ROOT = dirname(dirname(fileURLToPath(import.meta.url)))
const ENGINE_CLI = join(PLUGIN_ROOT, 'engine', 'cli.py')

/**
 * 调一次引擎命令。永不抛异常：任何失败都折叠成 {ok:false, error:{code,message}}，
 * 面板与工具都只认这一种形状，不需要 try/catch 各自处理。
 */
/**
 * 超时后把引擎进程整棵树杀掉（审查 R-16）。
 *
 * Windows 下子进程 kill() 只杀 python 本体，引擎刚起的 soffice 还在跑并占着文件锁；
 * 而 taskkill 自己也可能失败（被安全软件拦、进程已经退了、pid 被复用），
 * 所以 taskkill 起不来 / 退出码非 0 时必须回头再 kill 一次本体，不能什么都不做就报超时。
 */
function killEngineTree(child) {
  if (!child || !child.pid) return
  const killBody = () => { try { child.kill() } catch { /* 已经退了 */ } }
  if (process.platform !== 'win32') return killBody()
  let killer
  try {
    killer = spawn('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' })
  } catch {
    return killBody()
  }
  killer.on('error', killBody)
  killer.on('exit', (code) => { if (code !== 0) killBody() })
}

function runEngine(command, args, options = {}) {
  const started = Date.now()
  return new Promise(async (done) => {
    const python = await pythonReport()
    if (!python) {
      done({ ok: false, error: { code: 'python-missing', message: '没找到 DSH 自带的 Python，文档处理不可用' } })
      return
    }
    if (!existsSync(ENGINE_CLI)) {
      done({ ok: false, error: { code: 'engine-missing', message: '插件里的 engine/cli.py 不见了：' + ENGINE_CLI } })
      return
    }
    let child
    // 参数走 stdin（--args -）而不是 argv：Windows 单条命令行上限约 32K 字符，组卷那种大 blocks
    // 会把 spawn 直接撑死成 ENAMETOOLONG，而 /engine 本身允许 8MB body。写失败不算错，
    // 真正的报错从 stderr 与退出码里读（见下面 finish 的两条分支）。
    const payload = JSON.stringify(args ?? {})
    try {
      const env = Object.assign({}, process.env, { PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1', PYTHONDONTWRITEBYTECODE: '1' })
      // 画像路径统一成绝对路径再交给引擎：相对路径在两边（宿主 cwd / 引擎 cwd）解析结果不同，
      // 宿主会一直看着 A、引擎一直在写 B，缓存与面板就永远停在旧画像（审查 P2-11）。
      if (process.env.TEACHER_KIT_PROFILE) env.TEACHER_KIT_PROFILE = resolve(process.env.TEACHER_KIT_PROFILE)
      child = spawn(python.path, [ENGINE_CLI, command, '--args', '-'], {
        cwd: PLUGIN_ROOT,
        env,
        windowsHide: true
      })
    } catch (error) {
      done({ ok: false, error: { code: 'engine-spawn', message: error instanceof Error ? error.message : String(error) } })
      return
    }

    try {
      child.stdin.on('error', () => { /* 引擎提前退出时写 stdin 会 EPIPE，忽略即可 */ })
      child.stdin.end(payload, 'utf8')
    } catch { /* ignore */ }

    let stdout = ''
    let stderr = ''
    let settled = false
    const finish = (value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      done(Object.assign({ command, ms: Date.now() - started }, value))
    }
    // 超时上限/下限都夹住：传 "abc"/0/负数 会让 setTimeout 立刻触发或按 1ms 算，
    // 传 1e12 等于关掉超时（审查 P2-2）。
    const rawTimeout = Number(options.timeoutMs)
    const timeoutMs = Number.isFinite(rawTimeout) ? Math.min(Math.max(Math.round(rawTimeout), 1_000), 600_000) : 180_000
    const timer = setTimeout(() => {
      killEngineTree(child)
      finish({ ok: false, error: { code: 'engine-timeout', message: '这一步超过 ' + Math.round(timeoutMs / 1000) + ' 秒还没做完，已中止（大文件可以先拆小一点再试）' } })
    }, timeoutMs)

    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk) => { stdout += chunk })
    child.stderr.on('data', (chunk) => { stderr += chunk })
    child.on('error', (error) => finish({ ok: false, error: { code: 'engine-spawn', message: error.message } }))
    child.on('close', (code) => {
      const line = stdout.trim().split(/\r?\n/).filter((item) => item.trim() !== '').pop()
      if (!line) {
        finish({
          ok: false,
          error: {
            code: 'engine-empty',
            message: '引擎没有输出（退出码 ' + code + '）',
            hint: stderr.trim().slice(-600) || undefined
          }
        })
        return
      }
      try {
        const parsed = JSON.parse(line)
        if (parsed && parsed.error && typeof parsed.error.message === 'string') {
          parsed.error.message = scrub(parsed.error.message)
          if (typeof parsed.error.hint === 'string') parsed.error.hint = scrub(parsed.error.hint)
        }
        finish(parsed && typeof parsed === 'object' ? parsed : { ok: false, error: { code: 'engine-protocol', message: '引擎输出不是对象' } })
      } catch {
        finish({
          ok: false,
          error: { code: 'engine-protocol', message: '引擎输出不是 JSON', hint: scrub((line.slice(0, 300) + ' ' + stderr.trim().slice(-300)).trim()) }
        })
      }
    })
  })
}

// ── 画像缓存 ───────────────────────────────────────────────────────────────

/**
 * 画像文件在哪。必须与 engine/teacher_kit/profile.py 的 profile_path() 一一对应：
 * 这里只用它取 mtime/size 当缓存 key，**读内容仍然交给引擎**（老师手改的画像也得认）。
 */
function teacherHomePath() {
  const override = process.env.TEACHER_KIT_HOME
  if (override) return override
  const home = homedir()
  const documents = join(home, 'Documents')
  const base = existsSync(documents) ? documents : home
  return join(base, '教师助手')
}

function profileFilePath() {
  const override = process.env.TEACHER_KIT_PROFILE
  if (override) return resolve(override)
  const primary = join(homedir(), '.dsh', 'teacher-kit', 'profile.json')
  if (existsSync(dirname(primary))) return primary
  return join(teacherHomePath(), 'profile.json')
}

/**
 * 缓存 key = 路径 + mtimeMs + size。
 * 文件不存在时给 "...|missing" —— 这样"先没有、后来才存了画像"也会换 key，自动失效。
 */
function profileCacheKey() {
  const file = profileFilePath()
  try {
    const info = statSync(file)
    return file + '|' + info.mtimeMs + '|' + info.size
  } catch {
    return file + '|missing'
  }
}

let profileCache = { key: undefined, value: undefined }

/**
 * 读画像（进程内缓存）。返回形状与 runEngine 完全一致，面板与 /launch 都不用改。
 *
 * 只在**成功**时缓存：引擎报错（画像文件坏了、权限没了）不粘住，下一把还去 spawn，
 * 老师修好文件就能立刻读到。任何异常都退回真 spawn，绝不让请求变成 500。
 */
async function profileSummary(force = false) {
  let key
  try { key = profileCacheKey() } catch { key = undefined }
  if (!force && key !== undefined && profileCache.value !== undefined && profileCache.key === key) {
    return { ok: true, command: 'profile:get', data: profileCache.value, ms: 0, cached: true }
  }
  const result = await runEngine('profile:get', {})
  if (key !== undefined && result && result.ok === true && result.data && typeof result.data === 'object') {
    profileCache = { key, value: result.data }
  }
  return result
}

/** 刚存过画像（面板 / 工具）→ 立刻作废，下一次读就是新内容。 */
function invalidateProfileCache() {
  profileCache = { key: undefined, value: undefined }
}

// ── HTTP 小工具 ────────────────────────────────────────────────────────────

function writeJson(res, code, payload) {
  if (!res || res.writableEnded || res.destroyed) return
  const body = JSON.stringify(payload)
  res.writeHead(code, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body)
  })
  res.end(body)
}

/**
 * 面板内的失败回包统一成 {ok:false, error:{code, message, hint?}}。
 * /engine 一直是这个形状；/list-dir、/reveal 早先只回 {ok:false, error:"英文代码"}，
 * 面板把 "dir-must-be-absolute" 原样甩给老师（persona P1-11）。
 */
function panelFailure(res, status, code, message, hint) {
  const error = { code, message: scrub(message || code) }
  if (hint) error.hint = hint
  writeJson(res, status, { ok: false, error })
}

/**
 * 把 Node / 宿主抛出来的英文技术错翻成老师能看懂的话（审查 P2-12）。
 * 只认最常见的几种，不认识的保留原文（总比假装成功强），但仍然要 scrub。
 */
function friendlyText(raw) {
  const text = String(raw || '')
  if (!text) return '这一步没做成，请再试一次'
  if (/was disposed|disposed|session .* closed/i.test(text)) return '这个会话已经结束了，重新开一个新对话再试'
  if (/ENOENT/.test(text)) return '找不到这个文件（可能已经被移走或删掉了）'
  if (/EACCES|EPERM/.test(text)) return '没有权限操作这个位置（换到「桌面」或「文档」再试）'
  if (/EBUSY/.test(text)) return '这个文件正被别的程序打开，先关掉它再试'
  if (/ENOSPC/.test(text)) return '磁盘空间不够了'
  if (/^inputPath-/.test(text)) return '没有给要预览的文件'
  return scrub(text)
}

/** 路由里读请求体那几种失败（超大/不是 JSON/客户端中途断开）说人话。 */
function bodyFailure(res, error) {
  const message = error instanceof Error ? error.message : String(error)
  if (message === 'body-too-large') return panelFailure(res, 413, message, '这次的内容太大了（超过上限），先在文档里精简一下再试')
  if (message === 'invalid-json') return panelFailure(res, 400, message, '请求内容不是合法的 JSON')
  if (message === 'body-aborted') return panelFailure(res, 400, message, '请求发到一半断开了，请再点一次')
  return null
}

/** 回给面板的文字里不出现插件自己的安装路径（用户看不懂，也没必要知道）。 */
function scrub(text) {
  if (typeof text !== 'string') return text
  // Python 的 str(OSError) 会把路径里的反斜杠双写（D:\\a\\b），只 split 单写形式匹配不上，
  // 绝对路径（含用户名）就漏到面板上了（审查 P2-8）。两种形态都替换。
  const doubled = PLUGIN_ROOT.replace(/\\/g, '\\\\')
  return text.split(PLUGIN_ROOT).join('<插件目录>').split(doubled).join('<插件目录>').split(PLUGIN_ROOT.replace(/\\/g, '/')).join('<插件目录>')
}

/**
 * 导出 PDF 自己的失败：code 留给面板分支，message 直接是老师能看懂的中文。
 *
 * 用 teacherKit 标记，路由就能把「我们写好的中文」和「Node / LibreOffice 的
 * 英文 errno」分开：前者原样回，后者一律翻译，绝不让 EEXIST、ENOENT 漏到面板
 * （persona P1-1 是全接口唯一的英文报错点）。
 */
function pdfFailure(code, message, hint, status = 400) {
  const error = new Error(message)
  error.code = code
  error.status = status
  if (hint) error.hint = hint
  error.teacherKit = true
  return error
}

/**
 * /export-pdf 的失败信封：一律 panelFailure，而且**只给中文**。
 *
 * 老师不需要知道 errno，她需要知道"现在该做什么"：关掉阅读器、换个位置、
 * 清一下磁盘。所以下面把 errno 翻成动作，连原始 message 都不外带。
 */
function failExportPdf(res, error) {
  const raw = error instanceof Error ? error.message : String(error)
  if (raw === 'body-too-large') return panelFailure(res, 413, 'body-too-large', '这次的请求太大了')
  if (raw === 'invalid-json') return panelFailure(res, 400, 'invalid-json', '请求内容不是合法的 JSON')
  if (error && error.teacherKit) return panelFailure(res, error.status || 400, error.code, raw, error.hint)
  if (/libreoffice/i.test(raw)) {
    return panelFailure(res, 500, 'pdf-engine-missing', '这台电脑上的 PDF 转换组件没找到', '重启一次 DSH 再试；急着用的话可以先用 Word「另存为 PDF」')
  }
  const errno = String((error && error.code) || '')
  const friendly = {
    EEXIST: ['output-exists', '导出位置已经有一份同名文件，而且删不掉它', '先关掉正在看这份 PDF 的程序（阅读器 / Word / 打印预览），再点一次'],
    EPERM: ['output-busy', '这个 PDF 正被别的程序打开，先关掉它再导', '阅读器、Word、打印预览都可能占着它'],
    EBUSY: ['output-busy', '这个 PDF 正被别的程序打开，先关掉它再导', '阅读器、Word、打印预览都可能占着它'],
    EACCES: ['output-denied', '这个位置没有写入权限', '换一个位置，比如「桌面」或「文档」'],
    ENOENT: ['path-not-found', '找不到要导出的文件（可能已经被移走或删掉了）', '重新选一次文件，或者先把它另存到「文档」里'],
    ENOSPC: ['disk-full', '磁盘空间不够了，PDF 写不进去', '清一清「下载」里的旧文件，或者换个盘再导'],
    EROFS: ['read-only', '这个位置是只读的，写不进去', '换一个位置，比如「文档」']
  }[errno]
  if (friendly) return panelFailure(res, 400, friendly[0], friendly[1], friendly[2])
  panelFailure(res, 500, 'export-pdf-failed', 'PDF 没导出来', '确认文件没被占用、磁盘还有空间，然后再点一次')
}

const LOOPBACK = /^(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$/

/**
 * 面板请求判定：自定义头 + 同源 + 浏览器侧 same-origin。三道都不跳过：
 * 头挡跨站表单/图片探测，sec-fetch-site 挡同站子域被利用，Host 挡 DNS rebinding。
 * 这不是身份凭证（本机进程照样能发），但把浏览器侧的攻击面关死。
 */
function isPanelRequest(req) {
  if (!req.headers[LOCAL_HEADER]) return false
  const host = req.headers.host
  if (host && !LOOPBACK.test(host)) return false
  const site = req.headers['sec-fetch-site']
  if (site && site !== 'same-origin' && site !== 'none') return false
  const origin = req.headers.origin
  if (!origin) return true
  // 有 Origin 就必须能和 Host 对上；Host 缺失时按不可信处理（不放行）。
  try {
    return Boolean(host) && new URL(origin).host === host
  } catch {
    return false
  }
}

/** /status?refresh=1（面板上的"重新体检"）绕过 60 秒缓存。 */
function wantsRefresh(req) {
  const raw = String(req?.url || '')
  const at = raw.indexOf('?')
  if (at < 0) return false
  try {
    const value = new URLSearchParams(raw.slice(at + 1)).get('refresh')
    return value === '1' || value === 'true'
  } catch {
    return false
  }
}

function readJsonBody(req, limit = 1 << 20) {
  return new Promise((resolve, reject) => {
    let size = 0
    const chunks = []
    req.on('data', (chunk) => {
      size += chunk.length
      if (size > limit) { reject(new Error('body-too-large')); req.destroy(); return }
      chunks.push(chunk)
    })
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8')
      if (!raw) return resolve({})
      try { resolve(JSON.parse(raw)) } catch { reject(new Error('invalid-json')) }
    })
    req.on('error', reject)
    // 客户端发一半关掉标签页时只有 aborted/close，没有 end：不接住这个 Promise 永不 settle，
    // 请求闭包与 socket 一起泄漏（审查 P2-3）。
    req.on('aborted', () => reject(new Error('body-aborted')))
    req.on('close', () => reject(new Error('body-aborted')))
  })
}

function absPath(value, label) {
  if (typeof value !== 'string' || value.trim() === '') throw new Error(label + '-required')
  const full = resolve(value)
  if (!isAbsolute(full)) throw new Error(label + '-must-be-absolute')
  if (!existsSync(full)) throw new Error(label + '-not-found')
  return full
}

/** 刚写完的文件在 Windows 上偶尔会晚一拍才可见（杀软扫描）；等到 600ms 再判死。 */
async function waitForFile(fullPath, budgetMs = 600) {
  const until = Date.now() + budgetMs
  while (Date.now() < until) {
    if (existsSync(fullPath)) return true
    await new Promise((done) => setTimeout(done, 60))
  }
  return existsSync(fullPath)
}

/**
 * 一次渲染 = PDF + 逐页 PNG。
 * 这是"生成前能看见、印前能核对"的唯一实现路径（DSH 无 WYSIWYG 编辑）。
 */
const PREVIEW_DIR = '.teacher-kit-preview'
const OWNED_MARKER = '.teacher-kit-owned'

/**
 * 预览目录只允许两种：源文件旁的 .teacher-kit-preview（默认），或调用方给出
 * 且**已经带本插件标记**的目录。递归删除前必须见到标记，避免插件成为
 * "删别人目录"的通道（reviews/step1-skeleton.md M4）。
 */
const MARKER_TEXT = 'created by dsh-teacher-kit'
function hasOwnMarker(dir) {
  try { return readFileSync(join(dir, OWNED_MARKER), 'utf8').trim() === MARKER_TEXT } catch { return false }
}
function previewDirFor(input, workDir) {
  if (!workDir) {
    const dir = join(dirname(input), PREVIEW_DIR)
    mkdirSync(dir, { recursive: true })
    writeFileSync(join(dir, OWNED_MARKER), MARKER_TEXT + '\n', 'utf8')
    return dir
  }
  const dir = resolve(workDir)
  if (!hasOwnMarker(dir)) {
    // 光看文件名后缀就放行，等于允许"删掉任意一个叫 .teacher-kit-preview 的目录"（审查 P2-13）。
    // 没标记的目录只在两种情况下接手：目录不存在（我们自己建），
    // 或者名字像预览目录且**里面是空的**（老师手建的、还没有任何东西会被删）。
    const shaped = dir.replace(/[\\/]+$/, '').endsWith(PREVIEW_DIR)
    let empty = false
    let missing = true
    try {
      const stat = statSync(dir)
      missing = false
      empty = stat.isDirectory() && readdirSync(dir).length === 0
    } catch { missing = true }
    if (!missing && !(shaped && empty)) throw new Error('workDir-not-allowed')
  }
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, OWNED_MARKER), MARKER_TEXT + '\n', 'utf8')
  return dir
}

async function renderDocument(inputPath, workDir, pdfOverride) {
  const candidate = typeof inputPath === 'string' ? resolve(inputPath) : ''
  if (candidate && !existsSync(candidate)) await waitForFile(candidate)
  const input = absPath(inputPath, 'inputPath')
  const outDir = previewDirFor(input, workDir)
  const pdfPath = join(outDir, basenameWithoutExt(input) + '.pdf')
  rmSync(pdfPath, { force: true })
  // 逐页图按源文件名分目录：同一个预览目录里先渲染课件、再渲染试卷时，
  // 不能把前一份的页图删掉（面板上还挂着那些图的路径）。
  const pagesRoot = join(outDir, 'pages')
  const pagesDir = join(pagesRoot, basenameWithoutExt(input))
  mkdirSync(pagesRoot, { recursive: true })
  writeFileSync(join(pagesRoot, OWNED_MARKER), 'created by dsh-teacher-kit\n', 'utf8')
  if (existsSync(join(pagesRoot, OWNED_MARKER)) || existsSync(join(outDir, OWNED_MARKER))) {
    rmSync(pagesDir, { recursive: true, force: true })
  }

  const converter = await getConverter()
  // 面板刚导出过 PDF 时直接拿那份来出图：同一份文档不必让 LibreOffice 转两遍
  // （8 开卷子转换约 1–2 秒，老师等着呢）。
  // 只信"源文件同目录下的 .pdf"：pdfPath 来自请求体，放任它就能把任意文件当 PDF 读进预览（审查 P2）。
  const trustedPdf = (() => {
    if (typeof pdfOverride !== 'string' || !pdfOverride.trim()) return ''
    const candidatePdf = resolve(pdfOverride)
    if (!/\.pdf$/i.test(candidatePdf)) return ''
    if (dirname(candidatePdf) !== dirname(input)) return ''
    return existsSync(candidatePdf) ? candidatePdf : ''
  })()
  const reuse = !!trustedPdf
  const pdf = reuse
    ? { outputPath: trustedPdf, missingFonts: [] }
    : await converter.convert({ inputPath: input, outputPath: pdfPath })
  const sourcePdf = reuse ? trustedPdf : pdfPath
  // 逐页图从**导出后的 PDF** 渲染，而不是再让 LibreOffice 读一遍源文件：
  // 它渲染多栏文档时会在正文区域画一圈边框（PDF 里没有），预览和成品不一致，
  // 老师会以为印出来带框。从 PDF 出图，预览就是成品。
  const rendered = await converter.renderImages({ inputPath: sourcePdf, outputDir: pagesDir })
  return {
    ok: true,
    input,
    pdf: sourcePdf,
    pageCount: rendered.pageCount,
    dpi: rendered.dpi,
    missingFonts: rendered.missingFonts || pdf.missingFonts || [],
    pages: (rendered.images || []).map((image) => ({
      page: image.page,
      file: image.path,
      width: image.width,
      height: image.height
    })),
    warnings: (rendered.missingFonts || []).length > 0
      ? ['本机缺少字体：' + (rendered.missingFonts || []).join('、') + '（打开后版面可能与本机预览不同）']
      : []
  }
}

const PDF_EXTENSION = /\.pdf$/i

/**
 * 定出这次导出要写哪个 PDF，并把"转换前能做完的事"做完。
 *
 * 为什么单独抽出来（persona P1-1，教师实测）：老师改完一个错字会**再导一次**。
 * 旧实现把输出路径也丢给 absPath（它要求文件已存在），于是
 *   - 默认目标已经有一份同名 pdf → LibreOffice kit 用 'wx' 独占创建 → 英文 EEXIST；
 *   - 传一个新的 outputPath → outputPath-not-found。
 * 现在：输入必须存在；输出按"将来要写的文件"处理 ——
 *   - 默认 = 源文件旁边同名 .pdf；
 *   - 传了 outputPath 就按它写（不要求已存在），父目录缺了递归建出来；
 *   - 同名 pdf 已存在就先删掉（第二次导出 = 覆盖，返回值带 overwritten，面板可提示
 *     「已更新上次那份」）。删在转换之前是必须的：kit 的 open(..., 'wx') 见到旧文件
 *     必报 EEXIST，留着它等于这条路永远只能走一次；
 *   - 输出就是输入、或名字不是 .pdf，直接给中文错误。
 * 全同步、不碰 LibreOffice，smoke 里可以直接测（真转一遍太慢，也不该由 smoke 承担）。
 */
export function resolvePdfTarget(inputPath, outputPath) {
  let input
  try {
    input = absPath(inputPath, 'inputPath')
  } catch (error) {
    const code = error instanceof Error ? error.message : String(error)
    if (code === 'inputPath-required') {
      throw pdfFailure('inputPath-required', '没有指定要导出哪份文件', '先在面板里选一份 Word / PPT / Excel 文件')
    }
    if (code === 'inputPath-not-found') {
      throw pdfFailure('inputPath-not-found', '找不到这份文件（可能已经被移走或删掉了）', '重新选一次文件，或者先把它另存到「文档」里')
    }
    throw pdfFailure('inputPath-invalid', '这份文件的路径没法用', '点「浏览…」从磁盘里重新选一次')
  }

  const explicit = typeof outputPath === 'string' && outputPath.trim() !== ''
  let target
  if (explicit) {
    const raw = outputPath.trim()
    if (!isAbsolute(raw)) {
      throw pdfFailure('outputPath-must-be-absolute', '导出位置要写完整路径', '点「浏览…」从磁盘里选一个位置，就不用自己拼路径')
    }
    target = resolve(raw)
  } else {
    target = join(dirname(input), basenameWithoutExt(input) + '.pdf')
  }
  // 输出和输入是同一个文件：既会写坏源文件，LibreOffice 也会直接拒转。
  // Windows 上「同一份文件」不能靠字符串比：C:\\a\\X.docx 与 c:\\a\\x.docx 是同一个文件。
  if (samePath(target, input)) {
    throw pdfFailure('output-same-as-input', '导出的位置不能就是原文件本身', '换一个名字，比如在文件名后面加「-打印版」')
  }
  if (explicit && !PDF_EXTENSION.test(target)) {
    throw pdfFailure('output-not-pdf', '导出文件的名字要以 .pdf 结尾', '比如「第三单元-打印版.pdf」')
  }

  const parent = dirname(target)
  if (!existsSync(parent)) {
    try {
      mkdirSync(parent, { recursive: true })
    } catch {
      throw pdfFailure('output-dir-uncreatable', '这个位置建不出文件夹', '换一个位置，比如「桌面」或「文档」')
    }
  }

  let overwritten = false
  if (existsSync(target)) {
    let isDir = false
    try { isDir = statSync(target).isDirectory() } catch { isDir = false }
    if (isDir) {
      throw pdfFailure('output-is-directory', '这个导出位置是一个文件夹，不是文件', '换一个文件名，比如在末尾加上「-打印版.pdf」')
    }
    overwritten = true
  }
  return { input, target, overwritten }
}

/**
 * 导出 PDF：docx/pptx/xlsx → .pdf（走内置 LibreOffice）。
 * 幂等：同一份文件连点两次 → 第二次直接覆盖同名 pdf，返回 ok 且 overwritten: true。
 */
async function exportPdfFile(inputPath, outputPath) {
  const { input, target, overwritten } = resolvePdfTarget(inputPath, outputPath)
  const converter = await getConverter()
  // 先转到同目录的临时名字，成功了再顶掉旧 PDF。
  // 旧实现是「先 rmSync 目标再转」，一旦转换失败（字体缺失、文件被锁、素材坏），
  // 老师手里上一版 PDF 就永久没了 —— 打印课前那一下最伤。
  // 临时名必须以 .pdf 结尾：转换器按扩展名决定输出格式，.part 会被直接拒。
  const temp = join(dirname(target), '.' + basename(target, '.pdf') + '.tk-' + process.pid + '-' + Date.now() + '.pdf')
  let pdf
  try {
    pdf = await converter.convert({ inputPath: input, outputPath: temp })
  } catch (error) {
    try { rmSync(temp, { force: true }) } catch { /* ignore */ }
    throw error
  }
  try {
    if (existsSync(target)) rmSync(target, { force: true })
    renameSync(temp, target)
  } catch {
    try { rmSync(temp, { force: true }) } catch { /* ignore */ }
    throw pdfFailure('output-busy', '这个 PDF 正被别的程序打开，换不掉', '先关掉正在看这份 PDF 的阅读器 / Word / 打印预览，再点一次')
  }
  return { ok: true, pdf: target, overwritten, missingFonts: pdf.missingFonts || [] }
}

/** Windows 上路径大小写不敏感，比「是不是同一个文件」要归一化；其它平台按字面比。 */
function samePath(a, b) {
  if (a === b) return true
  if (process.platform !== 'win32') return false
  return a.toLowerCase() === b.toLowerCase()
}

/** 在资源管理器里选中文件（面板上的"打开…所在文件夹"）。 */
function revealInExplorer(target) {
  const full = absPath(target, 'path')
  let isDir = false
  try { isDir = statSync(full).isDirectory() } catch { isDir = false }
  if (process.platform === 'win32') {
    // explorer 的 /select 必须和路径在**同一个** argv 里，拆开会当成两个路径。
    spawn('explorer.exe', ['/select,' + full], { detached: true, stdio: 'ignore' }).unref()
  } else if (process.platform === 'darwin') {
    spawn('open', ['-R', full], { detached: true, stdio: 'ignore' }).unref()
  } else {
    spawn('xdg-open', [isDir ? full : dirname(full)], { detached: true, stdio: 'ignore' }).unref()
  }
  return { ok: true, revealed: full }
}

function basenameWithoutExt(file) {
  const base = file.split(/[\\/]/).pop() || 'document'
  return base.replace(/\.[^.]+$/, '')
}

// ── 本地文件浏览（面板的文件选择器） ──────────────────────────────────────
//
// 老师的文件在哪？桌面、文档、下载、U 盘。让他在面板里从上一次的位置
// 一层一层点进去，比让他敲 "D:\资料\七上\Unit 3 单元卷.docx" 现实得多。
// 只读：只列名字/大小/时间，不读文件内容，不写任何东西。

const FILE_KINDS = {
  docx: ['.docx', '.docm', '.dotx', '.doc'],
  pptx: ['.pptx', '.pptm', '.potx', '.ppt'],
  xlsx: ['.xlsx', '.xlsm', '.xltx', '.xls', '.csv'],
  pdf: ['.pdf'],
  image: ['.png', '.jpg', '.jpeg', '.webp', '.gif'],
  text: ['.txt', '.md']
}

/** kinds: ['docx'] / ['docx','pptx'] / ['all']；不传 = 全都列（目录永远都列）。 */
function kindsToExtensions(kinds) {
  if (!Array.isArray(kinds) || kinds.length === 0) return null
  const wanted = new Set()
  let all = false
  for (const kind of kinds) {
    const key = String(kind || '').toLowerCase()
    if (key === 'all') { all = true; continue }
    for (const ext of FILE_KINDS[key] || [key.startsWith('.') ? key : '.' + key]) wanted.add(ext)
  }
  if (all) return null
  return wanted.size ? wanted : null
}

function quickPlaces() {
  const home = homedir()
  const candidates = [
    ['桌面', join(home, 'Desktop')],
    ['文档', join(home, 'Documents')],
    ['下载', join(home, 'Downloads')],
    ['教师助手产物', join(home, 'Documents', '教师助手')]
  ]
  return candidates
    .filter(([, path]) => {
      try { return statSync(path).isDirectory() } catch { return false }
    })
    .map(([label, path]) => ({ label, path }))
}

/**
 * 列一层目录。默认从"文档"开始（老师的东西基本在这儿）。
 * 只回一层：面板要的是"点进去"，不是一次吐 5000 个文件。
 */
function listDirectory(dir, kinds) {
  const home = homedir()
  const documents = join(home, 'Documents')
  const fallback = (() => { try { return statSync(documents).isDirectory() ? documents : home } catch { return home } })()
  if (dir && !isAbsolute(String(dir))) throw new Error('dir-must-be-absolute')
  const base = dir ? resolve(String(dir)) : fallback
  let info
  try { info = statSync(base) } catch { throw new Error('dir-not-found') }
  if (!info.isDirectory()) throw new Error('not-a-directory')

  const allowed = kindsToExtensions(kinds)
  const entries = []
  let truncated = false
  let raw = []
  try { raw = readdirSync(base, { withFileTypes: true }) } catch { throw new Error('dir-unreadable') }
  for (const item of raw) {
    // 隐藏目录与插件自己的预览目录不进列表（老师不需要看见）
    if (item.name.startsWith('.') || item.name.startsWith('~$')) continue
    const full = join(base, item.name)
    if (item.isDirectory()) {
      entries.push({ name: item.name, path: full, dir: true, ext: '', size: 0, mtime: 0 })
      continue
    }
    if (!item.isFile()) continue
    const ext = (item.name.match(/\.[^.]+$/) || [''])[0].toLowerCase()
    if (allowed && !allowed.has(ext)) continue
    let size = 0
    let mtime = 0
    try {
      const stat = statSync(full)
      size = stat.size
      mtime = Math.round(stat.mtimeMs)
    } catch { continue }
    entries.push({ name: item.name, path: full, dir: false, ext, size, mtime })
  }
  entries.sort((left, right) => {
    if (left.dir !== right.dir) return left.dir ? -1 : 1
    return left.name.localeCompare(right.name, 'zh-Hans-CN')
  })
  const LIMIT = 400
  if (entries.length > LIMIT) {
    truncated = true
    entries.length = LIMIT
  }
  const parent = dirname(base)
  return {
    ok: true,
    dir: base,
    parent: parent === base ? null : parent,
    quick: quickPlaces(),
    entries,
    truncated
  }
}

// ── apply ──────────────────────────────────────────────────────────────────

/**
 * 启动记录：装坏 / 服务没就绪 / 路由被跳过时，用户在界面上只看到 404，
 * 一声不响最难查。所以每次 apply 与每处注册都往 ~/.dsh/teacher-kit/boot.log 记一行
 * （只保留最近 100 行，不含任何文件内容）。
 */
function bootLog(line) {
  try {
    const dir = join(homedir(), '.dsh', 'teacher-kit')
    mkdirSync(dir, { recursive: true })
    const file = join(dir, 'boot.log')
    let kept = ''
    try {
      const lines = readFileSync(file, 'utf8').split(/\r?\n/).filter(Boolean)
      kept = lines.slice(-99).join('\n')
    } catch { /* 第一次启动没有旧文件 */ }
    const stamp = new Date().toISOString().replace('T', ' ').slice(0, 19)
    const body = (kept ? kept + '\n' : '') + stamp + ' ' + line + '\n'
    writeFileSync(file, body, 'utf8')
  } catch { /* 记日志失败不能影响启动 */ }
}

export function apply(ctx) {
  bootLog('apply start v' + VERSION)
  const disposers = []
  const guard = (label, run) => {
    try {
      const dispose = run()
      if (typeof dispose === 'function') disposers.push(dispose)
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      registration.issues.push({ what: label, message })
      bootLog(label + ' 注册失败：' + message)
      try { ctx.logger?.warn?.('[teacher-kit] ' + label + ' 注册失败：' + message) } catch { /* ignore */ }
    }
  }

  /**
   * 等服务就绪再注册。
   *
   * 真机踩过：inject 为空时 apply 立即执行，那一刻 webServer / tools 可能还没被别的插件
   * 提供，ctx.get() 拿不到 → 注册被静默跳过（表现：面板一片 404，日志一声不响）。
   * ctx.inject(names, cb) 会在服务出现后回调；老宿主没有 inject 时退回立即执行。
   */
  const registered = new Set()
  const withService = (names, label, run) => {
    const attempt = (serviceCtx) => {
      if (registered.has(label)) return
      bootLog(label + ' 服务就绪，注册')
      // 不能先 add 再 run：各注册回调在服务还没拿到时 return undefined，
      // 那种情况本来该等下一次注入再试，先 add 就把重试机会永久吃掉了（审查 P2-5）。
      let dispose
      try {
        dispose = run(serviceCtx ?? ctx)
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        registration.issues.push({ what: label, message })
        bootLog(label + ' 注册失败：' + message)
        try { ctx.logger?.warn?.('[teacher-kit] ' + label + ' 注册失败：' + message) } catch { /* ignore */ }
        return
      }
      if (typeof dispose === 'function') disposers.push(dispose)
      if (dispose === undefined) {
        bootLog(label + ' 服务还没就绪，等下次注入')
        return
      }
      registered.add(label)
    }
    if (typeof ctx?.inject === 'function') {
      try {
        ctx.inject(names, (serviceCtx) => attempt(serviceCtx))
        return
      } catch (error) {
        registration.issues.push({ what: label, message: 'ctx.inject 不可用：' + (error instanceof Error ? error.message : String(error)) })
      }
    }
    attempt(ctx)
  }

  // register 是**方法**，必须绑在服务对象上调用：真机里把 server.register 拆下来传进来，
  // this 变成 undefined，内部读 this.exact / this.layers 直接 TypeError（面板 404 的元凶之一）。
  const countRoute = (server, entry) => {
    const dispose = server.register(entry)
    registration.routes += 1
    return dispose
  }
  const getService = (name) => {
    try { return ctx?.get?.(name) } catch { return undefined }
  }
  const countTool = (tools, definition) => {
    const dispose = tools.register(definition)
    registration.tools += 1
    return dispose
  }

  // agent 工具：模型能自己体检
  withService(['tools'], 'tool:teacher_kit_status', (serviceCtx) => {
    const tools = serviceCtx?.tools ?? getService('tools')
    if (!tools || typeof tools.register !== 'function') return undefined
    return countTool(tools, defineToolSafe({
      name: 'teacher_kit_status',
      description: '体检本机能否跑教师助手的文档能力（DSH 自带 Python 文档库 + 内置 LibreOffice）。用户问"插件能不能用/为什么转不了 PDF"时先调它。',
      parameters: {},
      output: { schema: { type: 'json' }, render: renderJson },
      async execute() { return await runtimeReport(true) }
    }))
  })

  // agent 工具：把一个 docx/pptx 渲染成 PDF + 逐页 PNG（预览与印前核对的地基）
  withService(['tools'], 'tool:teacher_kit_render', (serviceCtx) => {
    const tools = serviceCtx?.tools ?? getService('tools')
    if (!tools || typeof tools.register !== 'function') return undefined
    return countTool(tools, defineToolSafe({
      name: 'teacher_kit_render',
      description: '把一个 docx/pptx/xlsx 渲染成 PDF 并导出逐页 PNG，返回页数、缺字体与每页图片路径。用于"先看一眼再交付"和印前核对。',
      parameters: {
        inputPath: { type: 'string', required: true, description: '待渲染文件的绝对路径' },
        workDir: { type: 'string', required: false, description: '输出目录（绝对路径）；省略则放在源文件旁的 .teacher-kit-preview' }
      },
      output: { schema: { type: 'json' }, render: renderJson },
      async execute(args) {
        try {
          return await renderDocument(args.inputPath, args.workDir)
        } catch (error) {
          return { ok: false, error: error instanceof Error ? error.message : String(error) }
        }
      }
    }))
  })

  // agent 工具：引擎总入口（具体业务命令在这上面长出来）
  withService(['tools'], 'tool:teacher_kit_engine', (serviceCtx) => {
    const tools = serviceCtx?.tools ?? getService('tools')
    if (!tools || typeof tools.register !== 'function') return undefined
    return countTool(tools, defineToolSafe({
      name: 'teacher_kit_engine',
      description: '调用教师助手文档引擎。常用：selfcheck（体检）、papers（纸张清单）、inspect（读 docx/pptx 版式事实）、exam:build（出试卷）、bank:build（组卷）、handout:build（听写单）、plan:build（教案）、lesson:deck（课件）、docx:new / pptx:new（从零造文件）、wechat:render（公众号推文出 HTML+docx）、docflow:build（公文办理单）、profile:get / profile:set / profile:reset（学校画像，reset 是真清空）、template:make（把旧卷变成空白模板）、preflight（印前检查）。args 是命令参数对象。命令拼错时引擎会回 unknown-command 并列出全部可用命令，照它重试一次即可。',
      parameters: {
        command: { type: 'string', required: true, description: '引擎命令名，如 inspect / docx:new / pptx:new / papers / selfcheck' },
        args: { type: 'json', required: false, description: '命令参数对象，例如 {"path":"D:\\a.docx"} 或 {"outputPath":"D:\\out.docx","blocks":[...]}' }
      },
      output: { schema: { type: 'json' }, render: renderJson },
      async execute(args) {
        const input = args && typeof args === 'object' ? args : {}
        const command = typeof input.command === 'string' ? input.command.trim() : ''
        if (!command) {
          return { ok: false, error: 'command-required', hint: '告诉我要跑哪个引擎命令（selfcheck / inspect / docx:new …）' }
        }
        try {
          return await runEngine(command, input.args ?? {})
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error)
          return { ok: false, error: scrub(message) }
        }
      }
    }))
  })

  // agent 工具：读/记"学校画像"（学习适应的落地处）
  withService(['tools'], 'tool:teacher_kit_profile', (serviceCtx) => {
    const tools = serviceCtx?.tools ?? getService('tools')
    if (!tools || typeof tools.register !== 'function') return undefined
    return countTool(tools, defineToolSafe({
      name: 'teacher_kit_profile',
      description: '读写本机的"学校画像"：学校名、学段学科、常教年级、试卷纸型与页脚写法、教案版式、课件风格、公众号语气、公文单位。开始一个任务前先用 action=get 看已有预设（别重复问老师）；对话里问清后，用 action=set 把这次学到的记下来，下次自动带上。',
      parameters: {
        action: { type: 'string', required: true, enum: ['get', 'set', 'reset'], description: 'get 读画像 / set 记画像 / reset 清空画像（老师说「忘掉我学校的设置」时用）' },
        patch: { type: 'json', required: false, description: '要记的内容，按段给：{school:{name,stage,subject,grades,teacher}, paper:{key,footer}, plan:{layout}, deck:{theme,titleSizePt,bodySizePt}, wechat:{tone,signature}, docflow:{unit,dept}}' },
        note: { type: 'string', required: false, description: '这次为什么记它（写进画像历史，方便以后回溯）' },
        task: { type: 'string', required: false, description: '哪类任务学到的，如 exam / plan / deck / wechat / docflow' },
        learn: { type: 'boolean', required: false, description: 'true = 老师明确同意把这次的版式/风格存成默认' }
      },
      output: { schema: { type: 'json' }, render: renderJson },
      async execute(args) {
        const input = args && typeof args === 'object' ? args : {}
        const action = typeof input.action === 'string' ? input.action.trim() : ''
        try {
          if (action === 'get' || action === '') return await runEngine('profile:get', {})
          if (action === 'reset') {
            const cleared = await runEngine('profile:reset', {})
            invalidateProfileCache()
            return cleared
          }
          if (action === 'set') {
            const patch = input.patch && typeof input.patch === 'object' ? input.patch : {}
            const saved = await runEngine('profile:set', {
              patch,
              note: typeof input.note === 'string' ? input.note : '',
              task: typeof input.task === 'string' ? input.task : '',
              learn: input.learn === true
            })
            // 模型刚记下画像：宿主缓存也要跟着作废，别让面板还看到旧的
            invalidateProfileCache()
            return saved
          }
          return { ok: false, error: 'unknown-action', hint: 'action 只认 get / set / reset' }
        } catch (error) {
          return { ok: false, error: scrub(error instanceof Error ? error.message : String(error)) }
        }
      }
    }))
  })

  withService(['webServer'], 'route:engine', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/engine',
      handler: async (req, res) => {
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 POST 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求内容要按 JSON 发（面板版本与后台对不上，刷新页面再试）')
        }
        try {
          const body = await readJsonBody(req, 8 << 20)  // 组卷的 blocks 可能不小
          const command = body.command
          if (typeof command !== 'string' || command.trim() === '') {
            return panelFailure(res, 400, 'command-required', '没告诉后台要跑哪条引擎命令')
          }
          writeJson(res, 200, await runEngine(command.trim(), body.args ?? {}, { timeoutMs: body.timeoutMs }))
        } catch (error) {
          if (bodyFailure(res, error)) return
          writeJson(res, 400, { ok: false, error: { code: 'engine-request', message: friendlyText(error instanceof Error ? error.message : String(error)) } })
        }
      }
    })
  })

  /**
   * 对话式引导的核心：把"这次要做什么"变成一条老师口吻的 user 消息投进当前会话。
   *
   * 为什么必须走这条路：命令执行结果只是会话里的一个 flow 节点，模型收不到；
   * agent.followup 才是把消息作为 user 投进会话并唤醒模型（官方 /goal 命令也这么做）。
   * 面板因此只做"选任务"，问与做全在对话里。
   */
  /** 投递失败的原因 → 老师能看懂的一句话（R-15）。 */
  const LAUNCH_FAIL_TEXT = {
    'no-agents': '这个版本的 DSH 没开放对话通道，没法自动把话放进去',
    'no-session': '已经找不到这个对话了（可能刚关掉或换了会话），重新开一个再点一次',
    'no-channel': '这个对话暂时不接收自动消息，稍等一会儿再点，或者复制下面的话自己粘'
  }

  const deliverGuide = (sessionId, text) => {
    const agents = (() => { try { return ctx?.agents } catch { return undefined } })() ?? getService('agents')
    if (!agents || typeof agents.get !== 'function') return { ok: false, reason: 'no-agents' }
    let agent
    try { agent = agents.get(sessionId) } catch { agent = undefined }
    if (!agent) return { ok: false, reason: 'no-session' }
    // 消息形状照 createUserMessage 的产物写（id + role + content + source），宿主侧不再依赖 SDK 导出
    const message = { id: randomUUID(), role: 'user', content: [{ type: 'text', text }], source: { kind: 'user' } }
    if (typeof agent.followup === 'function') {
      agent.followup(message)
      return { ok: true, delivered: 'followup' }
    }
    if (typeof agent.inject === 'function') {
      agent.inject(message)
      return { ok: true, delivered: 'inject' }
    }
    return { ok: false, reason: 'no-channel' }
  }

  withService(['webServer'], 'route:launch', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/launch',
      handler: async (req, res) => {
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 POST 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求内容要按 JSON 发（面板版本与后台对不上，刷新页面再试）')
        }
        let text = ''
        try {
          const body = await readJsonBody(req, 256 << 10)
          const taskKey = typeof body.task === 'string' ? body.task.trim() : ''
          const task = findTask(taskKey)
          if (!task) {
            return panelFailure(res, 400, 'unknown-task', '不认识这个任务：' + (taskKey || '（空）'), '可选任务：' + TASK_KEYS.join(' / '))
          }
          /**
           * sessionId 来自请求体，本机任何进程（以及同源的别的插件前端）都能伪造 —— 见审查 P2-1。
           * 这里不做身份校验（宿主没给"这个面板属于哪个会话"的信息），只把形状夹死：
           * 会话号是 uuid 形态、长度有限，且必须真能在 agents 里查到。
           * 能投进去的也只可能是我们自己拼的引导文案（task 白名单 + note 500 字上限），
           * 最坏情况是别人往某个会话里塞一段教学引导，不是任意内容。
           */
          const rawSessionId = typeof body.sessionId === 'string' ? body.sessionId.trim() : ''
          const sessionId = /^[A-Za-z0-9_-]{6,128}$/.test(rawSessionId) ? rawSessionId : ''
          if (!sessionId) return panelFailure(res, 400, 'session-required', '面板没拿到会话号，没法把话放进对话里', '刷新页面再试；还不行就复制面板给的那段话自己粘到输入框')
          let profileText = ''
          try {
            // profileSummary 回的是与 runEngine 同形的信封（可能来自缓存），画像本体在 data 里
            const summary = await profileSummary()
            profileText = typeof summary?.data?.text === 'string' ? summary.data.text : ''
          } catch { profileText = '' }
          text = buildGuide(taskKey, profileText, typeof body.note === 'string' ? body.note : '')
          const delivered = deliverGuide(sessionId, text)
          if (delivered.ok) {
            writeJson(res, 200, { ok: true, task: taskKey, delivered: delivered.delivered, text })
          } else {
            // 投不进对话（宿主没给 agents / 会话还没起）时不能装成功：把话还给面板，老师可以自己粘。
            // reason 留成机器可读的短码（面板与测试按它分支），error 给老师一句人话（R-15 的信封口径）。
            writeJson(res, 200, {
              ok: false,
              task: taskKey,
              reason: delivered.reason,
              error: { code: delivered.reason, message: LAUNCH_FAIL_TEXT[delivered.reason] || '没能把这段话放进对话里' },
              text,
              hint: '没能自动开始对话，可以点"复制这段话"贴到输入框'
            })
          }
        } catch (error) {
          if (bodyFailure(res, error)) return
          // text 要留着：面板会退化成"复制这段话自己粘"，老师至少还有内容可用
          writeJson(res, 400, { ok: false, error: friendlyText(error instanceof Error ? error.message : String(error)), text })
        }
      }
    })
  })

  withService(['webServer'], 'route:profile', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/profile',
      handler: async (req, res) => {
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        try {
          if (req.method === 'GET') {
            const summary = await profileSummary()
            const tasks = TASKS.map((t) => ({ key: t.key, label: t.label, brief: t.brief }))
            // 「读失败」与「还没记过」必须分开：混成 exists:false 会让 /launch 拿空画像去重问老师。
            if (!summary || summary.ok === false) {
              const err = (summary && summary.error) || {}
              return writeJson(res, 200, {
                ok: true,
                degraded: true,
                profileError: err.message || '画像暂时读不出来',
                profile: { text: '', exists: false },
                tasks
              })
            }
            return writeJson(res, 200, { ok: true, profile: summary.data ?? { text: '', exists: false }, tasks })
          }
          if (req.method === 'POST') {
            const body = await readJsonBody(req, 256 << 10)
            const patch = body.patch && typeof body.patch === 'object' ? body.patch : {}
            const saved = await runEngine('profile:set', {
              patch,
              note: typeof body.note === 'string' ? body.note : '',
              learn: body.learn === true,
              task: typeof body.task === 'string' ? body.task : ''
            })
            // 刚存过：立刻作废，下一次读（可能是同一个面板马上再拉一次）必须是新画像
            invalidateProfileCache()
            // 引擎失败（比如画像文件被占用、写盘没权限）不能包成成功：
            // 老师会以为「记住了」，下次却按空画像重问一遍，或更糟——按旧习惯生成。
            if (!saved || saved.ok === false) {
              const err = (saved && saved.error) || {}
              return writeJson(res, err.code === 'profile-write-failed' ? 500 : 502, {
                ok: false,
                code: err.code || 'profile-save-failed',
                error: err.message || '画像没存进去',
                hint: err.hint || '再存一次；还不行的话把 ~/.dsh/teacher-kit/profile.json 删掉重设'
              })
            }
            return writeJson(res, 200, { ok: true, profile: saved.data ?? saved })
          }
          return writeJson(res, 405, { ok: false, error: 'method-not-allowed' })
        } catch (error) {
          if (bodyFailure(res, error)) return
          writeJson(res, 400, { ok: false, error: friendlyText(error instanceof Error ? error.message : String(error)) })
        }
      }
    })
  })

  withService(['webServer'], 'route:status', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/status',
      handler: async (req, res) => {
        if (req.method !== 'GET') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 GET 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        // 面板每次打开都会打这一枪：60 秒内复用上次体检，?refresh=1 才真去探测
        try { writeJson(res, 200, await runtimeReport(wantsRefresh(req), 60_000)) } catch (error) {
          writeJson(res, 500, { ok: false, error: error instanceof Error ? error.message : String(error) })
        }
      }
    })
  })

  withService(['webServer'], 'route:render', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/render',
      handler: async (req, res) => {
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 POST 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求内容要按 JSON 发（面板版本与后台对不上，刷新页面再试）')
        }
        try {
          const body = await readJsonBody(req)
          writeJson(res, 200, await renderDocument(body.inputPath, body.workDir, body.pdfPath))
        } catch (error) {
          if (bodyFailure(res, error)) return
          writeJson(res, 400, { ok: false, error: friendlyText(error instanceof Error ? error.message : String(error)) })
        }
      }
    })
  })

  withService(['webServer'], 'route:export-pdf', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/export-pdf',
      handler: async (req, res) => {
        // 失败一律走 {ok:false, error:{code,message,hint}}：这个接口曾经唯一次
        // 把 Node 的英文 EEXIST 甩给老师（persona P1-1），现在连 403/405 也统一。
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这个接口只接受 POST 请求')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是从面板发出来的，已经拦掉了')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求要带 application/json')
        }
        try {
          const body = await readJsonBody(req)
          writeJson(res, 200, await exportPdfFile(body.path, body.outputPath))
        } catch (error) {
          failExportPdf(res, error)
        }
      }
    })
  })

  withService(['webServer'], 'route:reveal', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/reveal',
      handler: async (req, res) => {
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 POST 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求内容要按 JSON 发（面板版本与后台对不上，刷新页面再试）')
        }
        try {
          const body = await readJsonBody(req)
          writeJson(res, 200, revealInExplorer(body.path))
        } catch (error) {
          const raw = error instanceof Error ? error.message : String(error)
          if (raw === 'body-too-large') return panelFailure(res, 413, raw, '这次的请求太大了')
          if (raw === 'path-required') {
            return panelFailure(res, 400, raw, '没有指定要打开的文件', '先做出一份文件，再点它下面的「打开…所在文件夹」')
          }
          if (raw === 'path-not-found') {
            return panelFailure(res, 400, raw, '找不到这个文件（可能已经被移走或删掉了）', '重新生成一次，或者自己去「文档」里找')
          }
          panelFailure(res, 400, 'reveal-failed', raw, '可以直接在资源管理器里打开「文档」文件夹找')
        }
      }
    })
  })

  withService(['webServer'], 'route:list-dir', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/list-dir',
      handler: async (req, res) => {
        if (req.method !== 'POST') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 POST 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        if (!(req.headers['content-type'] ?? '').toLowerCase().startsWith('application/json')) {
          return panelFailure(res, 415, 'json-required', '请求内容要按 JSON 发（面板版本与后台对不上，刷新页面再试）')
        }
        try {
          const body = await readJsonBody(req)
          writeJson(res, 200, listDirectory(body.dir, body.kinds))
        } catch (error) {
          const code = error instanceof Error ? error.message : String(error)
          const friendly = {
            'dir-not-found': ['找不到这个文件夹', '点「↑ 上一级」退回去，或者从「桌面 / 文档 / 下载」重新进去'],
            'not-a-directory': ['这不是一个文件夹', '选一个文件夹再进来'],
            'dir-must-be-absolute': ['请选一个文件夹（要写完整路径，如 D:\\教学\\第三单元）', '点「浏览…」从磁盘根目录一层层点进去就不会错'],
            'dir-unreadable': ['这个文件夹没有权限读', '换一个位置（比如桌面），或者把文件复制到「文档」里再选']
          }[code]
          if (friendly) return panelFailure(res, 400, code, friendly[0], friendly[1])
          panelFailure(res, 500, 'list-dir-failed', code, '换一个文件夹试试')
        }
      }
    })
  })

  /**
   * 读一张本地图片（预览缩略图）。
   * 浏览器加载 <img src> 带不了自定义头，所以面板用 fetch(带 x-teacher-kit) 拿到
   * bytes 再转 blob URL —— 这样"面板能看到自己刚做出来的卷子"不需要开静态目录，
   * 也就不会把用户任意文件暴露成一个 URL。
   */
  withService(['webServer'], 'route:file', (serviceCtx) => {
    const server = serviceCtx?.webServer ?? getService('webServer')
    if (!server || typeof server.register !== 'function') return undefined
    return countRoute(server, {
      kind: 'exact',
      path: API + '/file',
      handler: async (req, res) => {
        if (req.method !== 'GET') return panelFailure(res, 405, 'method-not-allowed', '这一步只接受 GET 请求（面板版本与后台对不上，刷新页面再试）')
        if (!isPanelRequest(req)) return panelFailure(res, 403, 'forbidden', '这个请求不是教师助手面板发来的，已拒绝（本机其它网页不能替你读写文件）')
        try {
          const url = new URL(req.url || '/', 'http://127.0.0.1')
          const full = absPath(url.searchParams.get('path') || '', 'path')
          const ext = (full.match(/\.[^.]+$/) || [''])[0].toLowerCase()
          const types = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.webp': 'image/webp',
            '.gif': 'image/gif'
          }
          const type = types[ext]
          if (!type) return panelFailure(res, 415, 'not-an-image', '这个文件不是图片，没法在这里预览')
          // 先看大小再读：几百 MB 的图整块读进宿主进程会直接 OOM（审查 P2-7）
          const stat = statSync(full)
          if (stat.size > 16 << 20) return panelFailure(res, 413, 'image-too-large', '这张图太大了（超过 16MB），没法在这里预览')
          const bytes = readFileSync(full)
          res.writeHead(200, {
            'content-type': type,
            'cache-control': 'no-store',
            'content-length': bytes.length
          })
          res.end(bytes)
        } catch (error) {
          writeJson(res, 400, { ok: false, error: friendlyText(error instanceof Error ? error.message : String(error)) })
        }
      }
    })
  })

  const cleanup = () => {
    for (const dispose of disposers.splice(0)) {
      try { dispose() } catch { /* 卸载时尽力而为 */ }
    }
    void disposeConverter()
  }
  // ctx.effect 是 apply 里最后一处"外部给的函数"：它不存在或自己抛错都不能带走整个壳。
  try {
    if (typeof ctx?.effect === 'function') {
      ctx.effect(() => cleanup, 'dsh-teacher-kit: host surface')
    } else {
      registration.issues.push({ what: 'ctx.effect', message: '宿主没有 effect 服务（卸载时不会自动清理）' })
    }
  } catch (error) {
    registration.issues.push({ what: 'ctx.effect', message: error instanceof Error ? error.message : String(error) })
  }
}