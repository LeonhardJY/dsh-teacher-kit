/**
 * 响应性能自测（发布前量「点一下要等多久」）。
 *
 * 用法：node scripts/perf-check.mjs [标签]
 *   - 标签只用来命名落盘的对比数据（.perf/<标签>.json），比如 before / after。
 *   - 全程进程内跑：直接 import lib/index.js、拿到注册的路由 handler，用假 ctx +
 *     假 agents + 内存 res 调它，不连真实实例、不写老师真实画像
 *     （TEACHER_KIT_PROFILE 指到 .perf/work/profile.json）。
 *
 * 量三件老师真会点的事：点开面板（GET /status、GET /profile）、点任务（POST /launch），
 * 外加一次裸的引擎往返（python cli.py profile:get）当基线 —— 好判断剩下的时间是不是
 * 本来就该花在起子进程上。
 *
 * 末尾再把缓存正确性过一遍：响应结构不变、外部改文件能失效、文件后来才出现能失效、
 * profile:set 能失效、/status 的 ?refresh=1 能绕过。任何一步 500 都算失败。
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const label = (process.argv[2] || 'run').replace(/[^A-Za-z0-9_-]/g, '')
const perfDir = join(root, '.perf')
const work = join(perfDir, 'work')
const profileFile = join(work, 'profile.json')
// 画像隔离：绝不能让自检写到老师真实的 ~/.dsh/teacher-kit/profile.json
process.env.TEACHER_KIT_PROFILE = profileFile
rmSync(work, { recursive: true, force: true })
mkdirSync(work, { recursive: true })
mkdirSync(perfDir, { recursive: true })

let failures = 0
const ok = (text) => console.log('  ✓ ' + text)
const bad = (text, detail) => { failures += 1; console.log('  ✗ ' + text + (detail ? ' — ' + detail : '')) }
const check = (cond, text, detail) => (cond ? ok(text) : bad(text, detail))

// ── 假宿主：只给路由注册口 + 一个能投递的会话 ──────────────────────────────
const routes = []
const guides = []
const fakeAgents = {
  get: (id) => (id === 'perf-session' ? { followup: (message) => guides.push(message) } : undefined)
}
const hostCtx = {
  get: (service) => service === 'webServer'
    ? { register: (route) => { routes.push(route); return () => {} } }
    : service === 'agents'
      ? fakeAgents
      : service === 'tools'
        ? { register: () => () => {} }
        : undefined,
  effect: (fn) => { fn(); return () => {} },
  logger: { warn: () => {} }
}

const host = await import(pathToFileURL(join(root, 'lib', 'index.js')).href)
try { host.apply(hostCtx) } catch (error) { bad('宿主 apply 不抛异常', error && error.message) }
const routeOf = (path) => routes.find((route) => route.path === path)
const API = '/teacher-kit/api'

function fakeRes() {
  const res = { code: 0, headers: {}, body: '' }
  res.writeHead = (code, headers) => { res.code = code; res.headers = headers || {} }
  res.end = (body) => { res.body = body || '' }
  return res
}
const HEADERS = { 'x-teacher-kit': '1', host: '127.0.0.1:52860' }

function parseBody(res) {
  try { return JSON.parse(res.body) } catch { return undefined }
}

/** 调一次 GET 路由，返回 {ms, code, payload}。 */
async function get(pathname, query) {
  const route = routeOf(API + pathname)
  if (!route) return { ms: 0, code: 0, payload: undefined, missing: true }
  const res = fakeRes()
  const started = performance.now()
  await route.handler({ method: 'GET', url: API + pathname + (query || ''), headers: { ...HEADERS } }, res)
  return { ms: performance.now() - started, code: res.code, payload: parseBody(res) }
}

/** 调一次 POST 路由（假 req 走 on('data'/'end')，跟真 http 一样）。 */
async function post(pathname, body) {
  const route = routeOf(API + pathname)
  if (!route) return { ms: 0, code: 0, payload: undefined, missing: true }
  const res = fakeRes()
  const payload = Buffer.from(JSON.stringify(body === undefined ? {} : body))
  const started = performance.now()
  await route.handler({
    method: 'POST',
    url: API + pathname,
    headers: { ...HEADERS, 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(payload)
      if (event === 'end') handler()
    }
  }, res)
  return { ms: performance.now() - started, code: res.code, payload: parseBody(res) }
}

/** 连打 n 次，返回每次毫秒数。 */
async function series(n, fn) {
  const out = []
  for (let i = 0; i < n; i += 1) out.push(await fn(i))
  return out
}
const median = (list) => {
  const sorted = [...list].sort((a, b) => a - b)
  const mid = Math.floor(sorted.length / 2)
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2
}
const ms = (value) => Math.round(value * 10) / 10

// ── 引擎基线：裸起一次 python 子进程（不经过任何缓存）──────────────────────
const pythonCandidates = [
  process.env.DSH_PYTHON,
  join(homedir(), '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'python.exe'),
  join(homedir(), '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'bin', 'python3'),
  'python'
].filter(Boolean)
const pythonPath = pythonCandidates.find((candidate) => candidate === 'python' || existsSync(candidate))
const engineCli = join(root, 'engine', 'cli.py')

function spawnOnce(args) {
  return new Promise((done) => {
    const started = performance.now()
    let settled = false
    const finish = (value) => { if (!settled) { settled = true; done(value) } }
    let child
    try {
      child = spawn(pythonPath, args, {
        cwd: root,
        env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1', PYTHONDONTWRITEBYTECODE: '1' },
        windowsHide: true
      })
    } catch { return finish({ ms: performance.now() - started, out: '' }) }
    let out = ''
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk) => { out += chunk })
    child.on('error', () => finish({ ms: performance.now() - started, out }))
    child.on('close', () => finish({ ms: performance.now() - started, out }))
  })
}

const PY_CHECK = 'import sys, docx, pptx, openpyxl; print(sys.version.split()[0])'
const engineBaseline = []
const probeBaseline = []
if (pythonPath && existsSync(engineCli)) {
  for (let i = 0; i < 3; i += 1) {
    const run = await spawnOnce([engineCli, 'profile:get', '--args', '{}'])
    engineBaseline.push(run.ms)
    if (i < 2) probeBaseline.push((await spawnOnce(['-c', PY_CHECK])).ms)
  }
} else {
  bad('找到 python + engine/cli.py（性能基线）', pythonPath || 'no python')
}

// ── 冷/热计时：每个接口连打 3 次，第 1 次是冷（进程内首次、缓存空），后两次是热 ──
const statusTimes = await series(3, async () => (await get('/status')).ms)
const profileTimes = await series(3, async () => (await get('/profile')).ms)
const launchTimes = await series(3, async () => (await post('/launch', { task: 'exam', sessionId: 'perf-session' })).ms)

function row(name, times) {
  return '  ' + name.padEnd(30) + times.map((value) => (ms(value) + 'ms').padStart(10)).join('') + (ms(median(times)) + 'ms').padStart(11)
}
console.log('')
console.log('性能自测（标签 ' + label + '）—— 每次毫秒')
console.log('  ' + '指标'.padEnd(29) + '冷(第1次)'.padStart(11) + '热(第2次)'.padStart(11) + '热(第3次)'.padStart(11) + '中位数'.padStart(11))
console.log(row('GET /status', statusTimes))
console.log(row('GET /profile', profileTimes))
console.log(row('POST /launch', launchTimes))
if (engineBaseline.length) console.log(row('引擎往返 profile:get（基线）', engineBaseline))
if (probeBaseline.length) console.log(row('python 体检探测（基线）', probeBaseline))
console.log('')
console.log('  注：「首次」= 本进程里第一次调这个接口；接口之间共享 python 探测与画像缓存，')
console.log('      越靠后的「首次」越便宜 —— 真实使用也是这个顺序（面板先 /status、再 /profile，')
console.log('      老师点任务才 /launch），真正付子进程钱的只有第一次。')

// ── 正确性：结构、失效、异常都不许把请求变成 500 ──────────────────────────
console.log('缓存正确性')
const profileCold = await get('/profile')
const profileWarm = await get('/profile')
check(profileCold.code === 200 && profileWarm.code === 200, 'GET /profile 冷/热都 200', profileCold.code + '/' + profileWarm.code)
check(
  JSON.stringify(profileCold.payload) === JSON.stringify(profileWarm.payload),
  '冷/热响应逐字节一致（面板不会因为缓存看到两种形状）',
  JSON.stringify(profileCold.payload).slice(0, 200) + ' vs ' + JSON.stringify(profileWarm.payload).slice(0, 200)
)
check(
  profileWarm.payload && profileWarm.payload.profile && profileWarm.payload.profile.exists === false,
  '画像文件不存在时能缓存「空画像」（不每次重起子进程）',
  JSON.stringify(profileWarm.payload && profileWarm.payload.profile)
)
check(
  profileWarm.payload && Array.isArray(profileWarm.payload.tasks) && profileWarm.payload.tasks.length === 10,
  '缓存命中仍带回 10 个任务',
  String(profileWarm.payload && profileWarm.payload.tasks && profileWarm.payload.tasks.length)
)

// 外部改文件（老师换电脑/手改/别的会话写）→ mtime 变，必须立刻看见
writeFileSync(profileFile, JSON.stringify({ version: 1, school: { name: '试验中学' }, learned: [], history: [] }), 'utf8')
const afterCreate = await get('/profile')
check(
  afterCreate.payload && afterCreate.payload.profile && afterCreate.payload.profile.exists === true &&
    /试验中学/.test(afterCreate.payload.profile.text || ''),
  '画像文件后来才出现 → 下一次读就看得到（缓存自动失效）',
  JSON.stringify(afterCreate.payload && afterCreate.payload.profile && afterCreate.payload.profile.text)
)

writeFileSync(profileFile, JSON.stringify({ version: 1, school: { name: '改过名的中学' }, learned: [], history: [] }), 'utf8')
const afterEdit = await get('/profile')
check(
  afterEdit.payload && /改过名的中学/.test(afterEdit.payload.profile.text || ''),
  '文件被外部改（mtime 变）→ 读到新内容',
  JSON.stringify(afterEdit.payload && afterEdit.payload.profile && afterEdit.payload.profile.text)
)

// 同一份内容再读：这次应命中缓存，且内容仍是新的
const afterEditWarm = await get('/profile')
check(
  JSON.stringify(afterEdit.payload) === JSON.stringify(afterEditWarm.payload),
  '外部改完后的第二次读走缓存且内容一致'
)

// 面板保存画像（POST /profile）→ 立刻失效，下一次读能看到
const saved = await post('/profile', { patch: { paper: { key: '8K' } }, note: '性能自检', task: 'exam' })
check(saved.code === 200 && saved.payload && saved.payload.ok === true, 'POST /profile 保存成功 → 200', JSON.stringify(saved.payload).slice(0, 200))
const afterSave = await get('/profile')
check(
  afterSave.payload && afterSave.payload.profile && afterSave.payload.profile.profile &&
    afterSave.payload.profile.profile.paper && afterSave.payload.profile.profile.paper.key === '8K',
  'POST /profile 后画像缓存立即失效（下一次读是新的）',
  JSON.stringify(afterSave.payload && afterSave.payload.profile && afterSave.payload.profile.profile)
)

// /status：60 秒缓存，?refresh=1 绕过
const statusCached = await get('/status')
const statusForced = await get('/status', '?refresh=1')
check(statusCached.code === 200 && statusForced.code === 200, 'GET /status 与 ?refresh=1 都 200', statusCached.code + '/' + statusForced.code)
check(
  statusCached.payload && 'python' in statusCached.payload && 'canConvertPdf' in statusCached.payload,
  '/status 结构不变（python / canConvertPdf 还在）',
  JSON.stringify(statusCached.payload).slice(0, 160)
)
check(
  JSON.stringify(statusCached.payload) === JSON.stringify(statusForced.payload),
  '/status 缓存与强刷返回同一份体检',
  JSON.stringify(statusForced.payload).slice(0, 160)
)

// /launch：缓存画像不能影响投递，投递的引导语要带上画像里的学校名
const launched = await post('/launch', { task: 'exam', sessionId: 'perf-session' })
check(launched.code === 200 && launched.payload && launched.payload.ok === true, 'POST /launch → 200 且投递成功', JSON.stringify(launched.payload).slice(0, 160))
const lastGuide = guides[guides.length - 1]
const guideText = lastGuide && Array.isArray(lastGuide.content) && lastGuide.content[0] ? String(lastGuide.content[0].text || '') : ''
check(guideText.length > 100, '引导语真的投进了会话', guideText.slice(0, 80))
check(/改过名的中学/.test(guideText), '引导语带上了画像里的学校名（缓存读的是同一份画像）', guideText.slice(0, 200))

// 读坏了要退回 spawn，不能 500：把画像文件写成坏 JSON
writeFileSync(profileFile, '{ 这不是 JSON', 'utf8')
const brokenRead = await get('/profile')
check(brokenRead.code === 200 && brokenRead.payload && brokenRead.payload.ok === true, '画像文件坏掉 → 仍 200（退回引擎，不 500）', String(brokenRead.code) + ' ' + JSON.stringify(brokenRead.payload).slice(0, 160))
const brokenLaunch = await post('/launch', { task: 'plan', sessionId: 'perf-session' })
check(brokenLaunch.code === 200 && brokenLaunch.payload && brokenLaunch.payload.ok === true, '画像文件坏掉 → /launch 仍 200', String(brokenLaunch.code) + ' ' + JSON.stringify(brokenLaunch.payload).slice(0, 160))

// 收尾：所有接口都不许回 500
const codes = [profileCold, profileWarm, afterCreate, afterEdit, afterEditWarm, saved, afterSave, statusCached, statusForced, launched, brokenRead, brokenLaunch].map((item) => item.code)
check(!codes.includes(500), '全程没有一次 500', codes.join(','))

const summary = {
  label,
  at: new Date().toISOString(),
  status: { cold: ms(statusTimes[0]), warm: ms(median(statusTimes.slice(1))) },
  profile: { cold: ms(profileTimes[0]), warm: ms(median(profileTimes.slice(1))) },
  launch: { cold: ms(launchTimes[0]), warm: ms(median(launchTimes.slice(1))) },
  engine: engineBaseline.length ? { cold: ms(engineBaseline[0]), warm: ms(median(engineBaseline.slice(1))) } : null,
  probe: probeBaseline.length ? { cold: ms(probeBaseline[0]), warm: ms(median(probeBaseline.slice(1))) } : null,
  failures
}
writeFileSync(join(perfDir, label + '.json'), JSON.stringify(summary, null, 2), 'utf8')

// 有 before.json 就顺手打一张改前/改后对比表（发布说明直接抄这张表）
const beforeFile = join(perfDir, 'before.json')
if (label !== 'before' && existsSync(beforeFile)) {
  let before
  try { before = JSON.parse(readFileSync(beforeFile, 'utf8')) } catch { before = undefined }
  if (before) {
    const line = (name, key) => {
      const was = before[key]
      const now = summary[key]
      if (!was || !now) return
      const ratio = now.warm > 0.05 ? was.warm / now.warm : Infinity
      const speedup = Number.isFinite(ratio) ? Math.round(ratio) + '×' : '瞬时'
      console.log('  ' + name.padEnd(26) + ('改前 ' + ms(was.warm) + 'ms').padStart(14) + ('改后 ' + ms(now.warm) + 'ms').padStart(14) + speedup.padStart(10))
    }
    console.log('')
    console.log('改前 / 改后（热路径中位数）')
    line('GET /status', 'status')
    line('GET /profile', 'profile')
    line('POST /launch', 'launch')
  }
}
console.log('')
console.log('  数据已存：.perf/' + label + '.json')
console.log(failures === 0 ? '性能自测通过 ✅' : '性能自测有 ' + failures + ' 项没过 ❌')
process.exitCode = failures === 0 ? 0 : 1
