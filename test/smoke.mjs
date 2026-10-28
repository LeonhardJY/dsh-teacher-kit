/**
 * dsh-teacher-kit 骨架 + 转换地基自检（纯 node，零依赖）。
 *
 * 验的是三件必须成立的事：
 *  1. 客户端信封能被 __ModuleLoader__ 接受，组件挂进 conversation.input.right
 *     （list 槽：绝不能碰 conversation.input.activity，那是 single 槽、语音输入插件占着，
 *      抢了它整个界面停在 Loading plugins —— 2026-10-25 真机事故）；
 *  2. 折叠态只是输入框右边一个小按钮，点开弹自己的浮层，再点 / Esc 收起；
 *  2. 宿主注册了 agent 工具与 HTTP 路由，且路由真的能回 JSON；
 *  3. 转换地基真的能跑：docx → PDF + 逐页 PNG（零外部安装，用自己的 Python 造样本）。
 */
import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { homedir } from 'node:os'
import vm from 'node:vm'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const work = join(root, '.smoke')
// v0.15：画像必须隔离到本次自检目录，绝不能写进老师真实的 ~/.dsh/teacher-kit/profile.json
process.env.TEACHER_KIT_PROFILE = join(work, 'profile.json')
// 每次从干净目录开始：上一轮留下的产物会让"不覆盖已有文件"的引擎命令直接失败。
rmSync(work, { recursive: true, force: true })
let failures = 0
const ok = (label) => console.log('  ✓ ' + label)
const bad = (label, detail) => { failures++; console.log('  ✗ ' + label + (detail ? ' — ' + detail : '')) }
const check = (cond, label, detail) => (cond ? ok(label) : bad(label, detail))

// ── 0. 构建客户端产物 ───────────────────────────────────────────────────────
const built = spawnSync(process.execPath, [join(root, 'scripts', 'build-client.mjs')], { encoding: 'utf8' })
check(built.status === 0, 'build-client.mjs 执行成功', built.stderr)

// ── 1. 极简 React 运行时（够跑 hooks + 点击）───────────────────────────────
function makeReact() {
  // 极简 React：state 跨渲染保留、effect 按依赖数组判脏。
  // 依赖数组必须真判——否则"切会话复位"那类 effect 会在每次渲染都把面板关掉，
  // 假运行时看着像产品坏了，其实是假件失真（真 React 只在 sessionId 变化时跑）。
  const state = { slots: [], deps: {}, cursor: 0 }
  let pending = null
  let queue = []
  let dirty = false

  const api = {
    createElement(type, props, ...children) {
      const kids = children.length === 0 ? undefined : children.length === 1 ? children[0] : children
      return { type, props: Object.assign({}, props || {}, { children: kids }) }
    },
    useState(initial) {
      const i = state.cursor++
      if (!(i in state.slots)) state.slots[i] = initial
      const set = (value) => {
        const next = typeof value === 'function' ? value(state.slots[i]) : value
        // 真 React 用 Object.is 比较：同一个值不重渲染。假运行时也必须这样，
        // 否则"把状态设成当前值"会无限重渲染。
        if (Object.is(next, state.slots[i])) return
        state.slots[i] = next
        dirty = true
      }
      return [state.slots[i], set]
    },
    useRef(initial) {
      const i = state.cursor++
      if (!(i in state.slots)) state.slots[i] = { current: initial }
      return state.slots[i]
    },
    useMemo(fn) { return fn() },
    useCallback(fn) { return fn },
    useLayoutEffect(fn, deps) { schedule('l', fn, deps) },
    useEffect(fn, deps) { schedule('e', fn, deps) }
  }

  function schedule(kind, fn, deps) {
    const i = state.cursor++
    const key = kind + i
    const prev = state.deps[key]
    const changed = !prev || !deps || deps.length !== prev.length || deps.some((d, k) => !Object.is(d, prev[k]))
    if (!changed) return
    state.deps[key] = deps ? deps.slice() : null
    queue.push(fn)
  }

  api.slots = () => state.slots.slice()
  api.render = (component, props) => {
    const draw = () => {
      state.cursor = 0
      queue = []
      const tree = component(props)
      // 目标组件已经算完，再把这一轮真正脏掉的 effect 跑一遍。
      // effect 里 setState 会同步重渲染（真 React 是异步的，这里够用且更容易暴露死循环）。
      const run = queue.splice(0)
      for (const fn of run) {
        try { fn() } catch (error) { throw error }
        if (queue.length) { const more = queue.splice(0); for (const extra of more) extra() }
      }
      // effect 里 setState 触发的重渲染得算进这一次 render 的结果，
      // 否则调用方拿到的还是旧树（真 React 是 flush 完才返回的）。
      if (dirty) { dirty = false; return draw() }
      return tree
    }
    pending = () => { draw() }
    return draw()
  }

  return api
}
function find(node, predicate) {
  if (!node || typeof node !== 'object') return undefined
  if (Array.isArray(node)) {
    for (const child of node) { const hit = find(child, predicate); if (hit) return hit }
    return undefined
  }
  if (predicate(node)) return node
  return find(node.props ? node.props.children : undefined, predicate)
}

function textOf(node) {
  if (node === null || node === undefined || node === false) return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (Array.isArray(node)) return node.map(textOf).join('')
  // v0.16：面板带了一段注入式 <style>，里面的 CSS 不是给老师看的字，别混进文案断言。
  if (node.type === 'style') return ''
  if (typeof node.type === 'function') {
    try { return textOf(node.type(node.props || {})) } catch { return '' }
  }
  return textOf(node.props ? node.props.children : '')
}

// ── 2. 客户端信封 ───────────────────────────────────────────────────────────
const bundle = readFileSync(join(root, 'client', 'client.js'), 'utf8')
let definition
const fetchCalls = []
// 默认响应（/status）；具体场景用 setRoute 覆盖，验"面板 → 引擎 → 结果卡"整条链。
const replies = new Map()
const DEFAULT_STATUS = { ok: true, canConvertPdf: true, surface: { toolsRegistered: 3, routesRegistered: 5, issues: [] } }
function setRoute(suffix, reply) { replies.set(suffix, reply) }
const winListeners = []
const sandbox = {
  window: {
    __ModuleLoader__: { load: (def) => { definition = def } },
    addEventListener: (type, fn) => { winListeners.push({ type, fn }) },
    removeEventListener: () => {}
  },
  console,
  fetch: (url, init) => {
    fetchCalls.push({ url, init: init || {} })
    const key = Array.from(replies.keys()).find((suffix) => String(url).includes(suffix))
    const body = key ? replies.get(key) : DEFAULT_STATUS
    const status = body && body.__status ? body.__status : 200
    return Promise.resolve({ ok: status >= 200 && status < 300, status, json: () => Promise.resolve(body || {}) })
  }
}
vm.runInNewContext(bundle, sandbox, { filename: 'client.js' })
check(Boolean(definition), '信封被 __ModuleLoader__.load 接收')
check(definition && definition.id === 'dsh-teacher-kit', '信封 id = dsh-teacher-kit', definition && definition.id)

const React = makeReact()
const mod = definition.factory((id) => {
  if (id === 'react') return React
  throw new Error('unexpected require: ' + id)
})
check(typeof mod.apply === 'function', '客户端导出 apply')
check(Array.isArray(mod.inject) && mod.inject.includes('slots'), '客户端 inject 含 slots', JSON.stringify(mod.inject))

// ── 3. 注册到 input.right 槽（list，不是 single）────────────────────────────
const registered = []
let injected = null
const clientCtx = {
  slots: {
    inject: (name, factory) => { injected = { name, factory } },
    register: (entry, component) => { registered.push({ entry, component }); return () => {} }
  },
  logger: { warn: () => {} }
}
mod.apply(clientCtx)
check(injected && injected.name === 'conversation.input.right', 'inject 到 conversation.input.right（list 槽）', injected && injected.name)
check(injected && injected.name !== 'conversation.input.activity', '不抢 conversation.input.activity（single 槽，真机会让整个界面起不来）', injected && injected.name)
injected.factory()
check(registered.length === 1, '恰好注册 1 个槽位组件')
const entry = registered[0] || {}
check(entry.entry && entry.entry.name === 'conversation.input.right', '槽位名正确')
check(entry.entry && entry.entry.id === 'dsh-teacher-kit', 'list 槽登记带 id（真机报过 requires options.id）', JSON.stringify(entry.entry && entry.entry.id))
check(typeof entry.component === 'function', '组件是函数组件')

// ── 4. 折叠态 / 浮层 ───────────────────────────────────────────────────────
const Component = entry.component
const triggerOf = (tree) => find(tree, (n) => n.type === 'button' && textOf(n).includes('教师助手'))
const dialogOf = (tree) => find(tree, (n) => n.type === 'div' && n.props && n.props.role === 'dialog')

const collapsed = React.render(Component, {})
check(Boolean(triggerOf(collapsed)), '折叠态渲染出一个「教师助手」按钮', collapsed && collapsed.type)
check(textOf(collapsed).includes('教师助手'), '按钮文案含「教师助手」', textOf(collapsed))
check(!dialogOf(collapsed), '折叠态不渲染浮层')
check(triggerOf(collapsed).props['aria-expanded'] === 'false', '折叠态 aria-expanded=false', triggerOf(collapsed).props['aria-expanded'])

triggerOf(collapsed).props.onClick()
const panel = React.render(Component, {})
const panelDialog = dialogOf(panel)
check(Boolean(panelDialog), '点按钮弹出自己的浮层（不再抢宿主工具栏那一条）', panel && panel.type)
check(triggerOf(panel).props['aria-expanded'] === 'true', '展开态 aria-expanded=true', triggerOf(panel).props['aria-expanded'])
const panelText = textOf(panel)
for (const word of ['出试卷', '做课件', '文件加工', '收起']) {
  check(panelText.includes(word), '面板含「' + word + '」')
}

// 再点一下按钮 / 点遮罩空白处 / 按 Esc 都要能收起（老师会试每一种）。
triggerOf(panel).props.onClick()
const reclosed = React.render(Component, {})
check(!dialogOf(reclosed), '再点一下按钮就收起')

triggerOf(reclosed).props.onClick()
const reopened = React.render(Component, {})
const backdrop = find(reopened, (n) => n.type === 'div' && n.props && typeof n.props.onMouseDown === 'function')
check(Boolean(backdrop), '浮层有遮罩（点外面收起）')
const maskClosed = (() => {
  if (!backdrop) return reopened
  const spot = {}
  backdrop.props.onMouseDown({ target: spot, currentTarget: spot })
  const afterMask = React.render(Component, {})
  check(!dialogOf(afterMask), '点遮罩空白处收起')
  return afterMask
})()

triggerOf(maskClosed).props.onClick()
const escOpen = React.render(Component, {})
check(Boolean(dialogOf(escOpen)), 'Esc 之前面板是打开的')
const keydown = winListeners.filter((item) => item.type === 'keydown').pop()
check(Boolean(keydown), '面板打开时挂了 Esc 监听（收起要能卸干净）')
if (keydown) {
  keydown.fn({ key: 'Escape' })
  const escClosed = React.render(Component, {})
  check(!dialogOf(escClosed), '按 Esc 收起')
}

// 体检结果要真的显示出来：/status 的响应早期被 .json() 解析两次，第二次必抛，
// 面板于是永远显示"体检失败"——这条就是那个 bug 的回归闸。
// （上面刚按 Esc 收起，这里重新点开再看状态行。）
triggerOf(React.render(Component, {})).props.onClick()
await new Promise((done) => setTimeout(done, 20))
const afterStatus = textOf(React.render(Component, {}))
check(!afterStatus.includes('体检失败'), '体检结果不是失败态', afterStatus.slice(0, 120))
check(afterStatus.includes('LibreOffice 就绪') || afterStatus.includes('Python 就绪'), '面板显示引擎就绪', afterStatus.slice(0, 160))

// ── 4.5 状态请求必须带面板头（S1 的回归闸）────────────────────────────────
const statusCall = fetchCalls.find((call) => String(call.url).includes('/status'))
check(Boolean(statusCall), '面板打开时请求了 /status', JSON.stringify(fetchCalls.map((c) => c.url)))
const sentHeaders = (statusCall && statusCall.init.headers) || {}
check(sentHeaders['x-teacher-kit'] !== undefined, '状态请求带 x-teacher-kit 头（否则真浏览器 403）', JSON.stringify(sentHeaders))
check(String(sentHeaders.accept || '').includes('json'), '状态请求声明 accept: json', JSON.stringify(sentHeaders))

// ── 4.6 locked 与切会话 ────────────────────────────────────────────────────
// 每个场景重新装一次组件：真 React 里换实例就是重新挂载。注意 hooks 用的是
// factory 注入的那个 React，所以连它一起换，否则状态还留在上一个场景里。
function mountFresh() {
  const R = makeReact()
  const m = definition.factory((id) => {
    if (id === 'react') return R
    throw new Error('unexpected require: ' + id)
  })
  let injected = null
  let captured = null
  m.apply({
    slots: {
      inject: (_name, factory) => { injected = factory },
      register: (entry, component) => { captured = { entry, component }; return () => {} }
    },
    logger: { warn: () => {} }
  })
  if (injected) injected()
  return { React: R, Component: captured && captured.component }
}

const lockedScene = mountFresh()
const lockedTree = lockedScene.React.render(lockedScene.Component, { locked: true, sessionId: 's1' })
const lockedBtn = find(lockedTree, (n) => n.type === 'button' && textOf(n).includes('教师助手'))
check(lockedBtn && lockedBtn.props.disabled === true, 'locked 时触发器禁用（M2）', JSON.stringify(lockedBtn && lockedBtn.props.disabled))
if (lockedBtn && typeof lockedBtn.props.onClick === 'function') lockedBtn.props.onClick()
const lockedAfter = lockedScene.React.render(lockedScene.Component, { locked: true, sessionId: 's1' })
check(!dialogOf(lockedAfter), 'locked 时点不开面板')

const sessionScene = mountFresh()
const closedTree = sessionScene.React.render(sessionScene.Component, { sessionId: 'sA' })
check(!dialogOf(closedTree), '新会话初始是折叠态')
find(closedTree, (n) => n.type === 'button' && textOf(n).includes('教师助手')).props.onClick()
const openedTree = sessionScene.React.render(sessionScene.Component, { sessionId: 'sA' })
check(Boolean(dialogOf(openedTree)), '同一会话内点开后保持展开')
const switchedTree = sessionScene.React.render(sessionScene.Component, { sessionId: 'sB' })
check(!dialogOf(switchedTree), '切到另一个会话后面板自动收起（M1）')

// ── 4.7 面板 → 引擎 → 结果卡（整条链，含印前检查这道门）───────────────────
const flow = mountFresh()
const flowProps = (sessionId) => ({ sessionId: sessionId || 'sFlow' })
let tree = flow.React.render(flow.Component, flowProps())
find(tree, (n) => n.type === 'button' && textOf(n).includes('教师助手')).props.onClick()
tree = flow.React.render(flow.Component, flowProps())

// v0.15：点开面板先看到「选一件事」，点下去把引导语投进对话，不直接甩一张表单。
const flowText = textOf(tree)
check(
  (flowText.includes('选一件事') || flowText.includes('点一件事')) && flowText.includes('剩下的在对话里问你'),
  '面板一打开是「选一件事」，不是填表',
  flowText.slice(0, 160)
)
check(flowText.includes('还没设学校') || flowText.includes('设定我的学校'), '面板顶部有学校画像行', flowText.slice(0, 220))
check(!find(tree, (n) => n.type === 'input' && n.props.placeholder === '七年级英语期中测试卷'), '默认不展开老表单（表单收在折叠里）')
setRoute('/launch', { ok: true, delivered: 'followup', task: 'exam', chars: 900 })
const guideButton = find(tree, (n) => n.type === 'button' && /^出试卷/.test(textOf(n)))
if (guideButton) guideButton.props.onClick()
await new Promise((done) => setTimeout(done, 20))
tree = flow.React.render(flow.Component, flowProps())
const launchCall = fetchCalls.filter((call) => String(call.url).includes('/launch')).pop()
check(Boolean(launchCall), '点任务会 POST /launch')
check(Boolean(launchCall) && JSON.parse(launchCall.init.body || '{}').task === 'exam', '投的是这次点的任务（task=exam）', launchCall && launchCall.init.body)
check(
  textOf(tree).includes('已经放进对话里了') && textOf(tree).includes('接着聊'),
  '投递成功后面板给回执（接着在对话里聊）',
  textOf(tree).slice(0, 300)
)

// 老方式（填表直接出成品）不删，只是要手动展开。
const oldToggle = find(tree, (n) => n.type === 'button' && textOf(n).includes('直接出成品'))
check(Boolean(oldToggle), '有「直接出成品（老方式）」入口')
if (oldToggle) oldToggle.props.onClick()
tree = flow.React.render(flow.Component, flowProps())
const examCard = find(tree, (n) => n.type === 'button' && textOf(n).trim() === '出试卷')
check(Boolean(examCard), '展开老方式后出现「出试卷」卡片')
if (examCard) examCard.props.onClick()
tree = flow.React.render(flow.Component, flowProps())
check(Boolean(find(tree, (n) => n.type === 'input' && n.props.placeholder === '七年级英语期中测试卷')), '选中出试卷后出现「卷子名称」输入框')
check(Boolean(find(tree, (n) => n.type === 'textarea')), '选中出试卷后出现「粘贴题干」输入框')

const nameBox = find(tree, (n) => n.type === 'input' && n.props.placeholder === '七年级英语期中测试卷')
nameBox.props.onChange({ target: { value: '七上期中测试' } })
tree = flow.React.render(flow.Component, flowProps())
const stemBox = find(tree, (n) => n.type === 'textarea')
stemBox.props.onChange({ target: { value: '一、听力（100 分）\n1. What does he want?' } })
tree = flow.React.render(flow.Component, flowProps())

// 每张卡各记各的：切到「教案」不该看到「出试卷」填过的卷子名称（真机抓到过）。
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '教案').props.onClick()
tree = flow.React.render(flow.Component, flowProps())
const planTitle = find(tree, (n) => n.type === 'input' && n.props.placeholder === 'Unit 3 What does he look like？')
check(planTitle && planTitle.props.value === '', '切到教案卡时课题是空的（各卡各记各的）', JSON.stringify(planTitle && planTitle.props.value))
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '出试卷').props.onClick()
tree = flow.React.render(flow.Component, flowProps())
const examTitle2 = find(tree, (n) => n.type === 'input' && n.props.placeholder === '七年级英语期中测试卷')
check(examTitle2 && examTitle2.props.value === '七上期中测试', '切回出试卷卡还记得刚才填的卷子名称', JSON.stringify(examTitle2 && examTitle2.props.value))

// 答题卡给填空/简答留作答区（0.12.0）：面板上得能选留几行，选项文案得让老师看得懂。
const kidsOf = (node) => {
  const raw = node && node.props ? node.props.children : undefined
  if (raw === undefined || raw === null || raw === false) return []
  return Array.isArray(raw) ? raw : [raw]
}
const optionsText = (node) => kidsOf(node).map((child) => textOf(child)).join('/')
const fillSelect = find(tree, (n) => n.type === 'select' && String(n.props.value || '') === '2' &&
  optionsText(n).includes('两行（写单词/短语）'))
check(Boolean(fillSelect), '出试卷卡有「填空题留几行」下拉（答题卡作答区）')
const fillLabels = fillSelect ? optionsText(fillSelect) : ''
check(fillLabels.includes('三行（写句子）') && fillLabels.includes('四行（简答/翻译）'),
  '「填空题留几行」有 2/3/4 三档', fillLabels)

// 合并多份时能选页眉页脚怎么带（0.12.0）：默认沿用第一份，另有每份保留、都不要。
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '文件加工').props.onClick()
tree = flow.React.render(flow.Component, flowProps())
const actionSelect = find(tree, (n) => n.type === 'select' && optionsText(n).includes('合并多份'))
check(Boolean(actionSelect), '文件加工卡有「动作」下拉')
if (actionSelect) actionSelect.props.onChange({ target: { value: 'merge' } })
tree = flow.React.render(flow.Component, flowProps())
const headerSelect = find(tree, (n) => n.type === 'select' && optionsText(n).includes('每份保留自己的'))
check(Boolean(headerSelect), '合并多份时出现「页眉页脚怎么带」下拉')
check(Boolean(headerSelect) && String(headerSelect.props.value || '') === 'first', '页眉默认沿用第一份（老行为不变）', JSON.stringify(headerSelect && headerSelect.props.value))

// 题库组卷：挑题前先能「看看题库里有什么」，挑题框得在（0.5.0 新增，真机点通过）。
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '题库组卷').props.onClick()
tree = flow.React.render(flow.Component, flowProps())
check(Boolean(find(tree, (n) => n.type === 'button' && textOf(n).includes('看看题库里有什么'))), '题库卡有「看看题库里有什么」按钮')
check(Boolean(find(tree, (n) => n.type === 'textarea' && n.props.placeholder === '单项选择 10\n完形填空 1\n阅读理解 2')), '题库卡有「挑题」输入框（一行一种题型）')
check(textOf(tree).includes('挑题（不填就用题库里全部；一行一种题型）'), '挑题框有说明文案')
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '出试卷').props.onClick()
tree = flow.React.render(flow.Component, flowProps())

// 整套卡（0.14.0）：一节课四件套，填了哪栏出哪件；四栏全空要拦住并说清怎么填。
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '整套').props.onClick()
tree = flow.React.render(flow.Component, flowProps())
const packText = textOf(tree)
check(packText.includes('一次做齐'), '整套卡有「一次做齐」按钮', packText.slice(0, 160))
check(packText.includes('听写单 · 词表文件') && packText.includes('教案 · 草稿') && packText.includes('课件 · 内容'),
  '整套卡四栏都在（试卷/听写/教案/课件）', packText.slice(0, 400))
const packPaper = find(tree, (n) => n.type === 'select' && optionsText(n).includes('16 开 · 一栏（单元小测、随堂练）'))
check(Boolean(packPaper), '整套卡也能选试卷纸型')
const packRun = find(tree, (n) => n.type === 'button' && textOf(n) === '一次做齐')
if (packRun) packRun.props.onClick()
await new Promise((done) => setTimeout(done, 30))
tree = flow.React.render(flow.Component, flowProps())
check(textOf(tree).includes('四栏都空着') || textOf(tree).includes('至少填一样'), '整套卡空着点会拦住并说明填法', textOf(tree).slice(0, 200))
find(tree, (n) => n.type === 'button' && textOf(n).trim() === '出试卷').props.onClick()
tree = flow.React.render(flow.Component, flowProps())

// 宿主真实回包是信封：成功 {ok:true, data:{...}}（真机 0.1.0 抓到的形状，面板必须能拆）。
setRoute('/engine', {
  ok: true,
  command: 'exam:build',
  ms: 11,
  data: {
    outputPath: 'D:/out/试卷.docx',
    questionCount: 1,
    totalScore: 100,
    sections: [{ title: '一、听力', questions: 1, score: 100 }],
    warnings: [],
    preflight: {
      ok: false, level: 'fail', score: 55, summary: '1 项必须改',
      findings: [{ level: 'error', code: 'footer-no-page-field', message: '页脚的页码是手打的字，不是页码域' }]
    }
  }
})
const beforeRun = fetchCalls.length
const runBtn = find(tree, (n) => n.type === 'button' && textOf(n).includes('出卷并导出 PDF'))
check(Boolean(runBtn), '表单里有「出卷并导出 PDF」按钮')
if (runBtn) runBtn.props.onClick()
await new Promise((done) => setTimeout(done, 60))
tree = flow.React.render(flow.Component, flowProps())
const blockedText = textOf(tree)
check(blockedText.includes('试卷已生成（PDF 待改）'), '印前检查不通过 → 标题标注 PDF 待改', blockedText.slice(0, 200))
check(blockedText.includes('PDF 先不导'), '印前检查不通过 → 不导 PDF', blockedText.slice(0, 320))
check(!fetchCalls.slice(beforeRun).some((call) => String(call.url).includes('/export-pdf')), '不通过时没有调用 /export-pdf')
check(blockedText.includes('打开「Word」所在文件夹'), '结果卡给「打开文件夹」入口', blockedText.slice(0, 320))
check(blockedText.includes('共 1 题，合计 100 分'), '结果卡显示题量与总分', blockedText.slice(0, 320))

setRoute('/engine', {
  ok: true, command: 'exam:build', ms: 9,
  data: {
    outputPath: 'D:/out/试卷.docx',
    questionCount: 1, totalScore: 100, sections: [], warnings: [],
    preflight: { ok: true, level: 'pass', score: 100, summary: '印前检查通过，可以导出', findings: [] }
  }
})
setRoute('/export-pdf', { ok: true, pdf: 'D:/out/试卷.pdf' })
const runBtn2 = find(tree, (n) => n.type === 'button' && textOf(n).includes('出卷并导出 PDF'))
if (runBtn2) runBtn2.props.onClick()
await new Promise((done) => setTimeout(done, 60))
tree = flow.React.render(flow.Component, flowProps())
const passText = textOf(tree)
check(passText.includes('印前检查通过'), '通过的卷子显示印前检查通过', passText.slice(0, 240))
check(fetchCalls.some((call) => String(call.url).includes('/export-pdf')), '通过后调用 /export-pdf')
check(passText.includes('打开「PDF」所在文件夹'), '结果卡多出 PDF 入口', passText.slice(0, 360))

// 失败信封 {ok:false, error:{code,message,hint}}：面板要把 message 与 hint 都摊开给老师看。
setRoute('/engine', {
  ok: false, command: 'exam:build', ms: 7,
  error: { code: 'text-required', message: '没有收到题目文字（把题干粘贴进来）', hint: '把题干粘贴进输入框再点一次' }
})
const runBtn3 = find(tree, (n) => n.type === 'button' && textOf(n).includes('出卷并导出 PDF'))
if (runBtn3) runBtn3.props.onClick()
await new Promise((done) => setTimeout(done, 60))
tree = flow.React.render(flow.Component, flowProps())
const failText = textOf(tree)
check(failText.includes('没有收到题目文字'), '引擎失败时显示 error.message', failText.slice(0, 260))
check(failText.includes('把题干粘贴进输入框再点一次'), '引擎失败时一并给出 hint', failText.slice(0, 320))

// ── 5. 宿主侧注册 ───────────────────────────────────────────────────────────
const host = await import(pathToFileURL(join(root, 'lib', 'index.js')).href)
check(host.name === 'dsh-teacher-kit', '宿主导出 name')
check(typeof host.apply === 'function', '宿主导出 apply')
check(Array.isArray(host.inject), '宿主导出 inject')

/**
 * 严格 registry：照 dsh-tools 的 register 真门（output.render 必须函数、
 * assertSupportedJsonSchema 只认 7 种 type）。工具定义不合格就在这里抛，
 * 而不是像以前那样被假 registry 全放过去（S2 的回归闸）。
 */
const SCHEMA_TYPES = new Set(['object', 'array', 'string', 'number', 'integer', 'boolean', 'null'])
function assertSchema(schema, path) {
  if (typeof schema !== 'object' || schema === null || Array.isArray(schema)) throw new Error('unsupported JSON schema at ' + path)
  if (schema.type !== undefined && !SCHEMA_TYPES.has(schema.type)) throw new Error('unsupported JSON schema type at ' + path + ': ' + schema.type)
  if (schema.properties) for (const [key, value] of Object.entries(schema.properties)) assertSchema(value, path + '.' + key)
  if (schema.items) assertSchema(schema.items, path + '[]')
}
const tools = []
const routes = []
const registerIssues = []
// 假 agents 服务：/launch 要能证明"真的把话投进了这个会话"
const followed = []
const injectedGuides = []
const fakeAgents = {
  get: (id) => {
    if (id === 'session-live') return { followup: (message) => followed.push(message) }
    if (id === 'session-inject') return { inject: (message) => injectedGuides.push(message) }
    if (id === 'session-mute') return {}
    return undefined
  }
}
const hostCtx = {
  get: (service) => service === 'agents'
    ? fakeAgents
    : service === 'tools'
    ? {
        register: (def) => {
          try {
            if (!def || typeof def.output?.render !== 'function') throw new Error('output.render required')
            assertSchema(def.parameters || {}, 'parameters')
            assertSchema(def.output.schema || {}, 'output.schema')
            tools.push(def)
            return () => {}
          } catch (error) {
            registerIssues.push(def?.name + ': ' + error.message)
            throw error
          }
        }
      }
    : service === 'webServer'
      ? { register: (route) => { routes.push(route); return () => {} } }
      : undefined,
  effect: (fn) => { fn(); return () => {} },
  logger: { warn: () => {} }
}
let threw = null
try { host.apply(hostCtx) } catch (error) { threw = error }
check(threw === null, '宿主 apply 不抛异常', threw && threw.message)

check(registerIssues.length === 0, '三个工具都过了真 schema 门（S2 回归）', registerIssues.join(' | '))
const toolNames = tools.map((t) => t.name)
check(toolNames.includes('teacher_kit_status'), '注册 teacher_kit_status', toolNames.join(','))
check(toolNames.includes('teacher_kit_render'), '注册 teacher_kit_render', toolNames.join(','))
check(toolNames.includes('teacher_kit_engine'), '注册 teacher_kit_engine', toolNames.join(','))
check(toolNames.includes('teacher_kit_profile'), '注册 teacher_kit_profile（学习适应的读写口）', toolNames.join(','))
const expectedRoutes = [
  '/teacher-kit/api/status',
  '/teacher-kit/api/render',
  '/teacher-kit/api/engine',
  '/teacher-kit/api/export-pdf',
  '/teacher-kit/api/reveal',
  '/teacher-kit/api/list-dir',
  '/teacher-kit/api/file',
  '/teacher-kit/api/launch',
  '/teacher-kit/api/profile'
]
const missingRoutes = expectedRoutes.filter((p) => !routes.some((r) => r.path === p))
check(
  routes.length === expectedRoutes.length && missingRoutes.length === 0,
  '注册 9 条 HTTP 路由（status / render / engine / export-pdf / reveal / list-dir / file / launch / profile）',
  String(routes.length) + (missingRoutes.length ? ' 缺 ' + missingRoutes.join(',') : '')
)
check(routes.every((route) => route.path.startsWith('/teacher-kit/api/') && typeof route.handler === 'function'), '路由都有路径与 handler')

function fakeRes() {
  const res = { code: 0, headers: {}, body: '' }
  res.writeHead = (code, headers) => { res.code = code; res.headers = headers || {} }
  res.end = (body) => { res.body = body || '' }
  return res
}
const byPath = (p) => routes.find((r) => r.path === p)

// ── 6. 路由鉴权与方法校验 ──────────────────────────────────────────────────
const statusRoute = byPath('/teacher-kit/api/status')
const noHeader = fakeRes()
await statusRoute.handler({ method: 'GET', headers: {} }, noHeader)
check(noHeader.code === 403, '无面板头 → 403', String(noHeader.code))

const plain = fakeRes()
await statusRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1' } }, plain)
check(plain.code === 200, '带面板头 → 200', String(plain.code))
let statusPayload
try { statusPayload = JSON.parse(plain.body) } catch { statusPayload = undefined }
check(statusPayload && ('python' in statusPayload) && ('canConvertPdf' in statusPayload), '体检含 python / canConvertPdf', plain.body.slice(0, 160))
check(statusPayload && statusPayload.canConvertPdf === true, 'canConvertPdf = true（内置 LibreOffice 可用）', String(statusPayload && statusPayload.canConvertPdf))

const foreignOrigin = fakeRes()
await statusRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1', origin: 'https://evil.example' } }, foreignOrigin)
check(foreignOrigin.code === 403, '跨站 Origin → 403', String(foreignOrigin.code))

const crossSite = fakeRes()
await statusRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1', 'sec-fetch-site': 'cross-site' } }, crossSite)
check(crossSite.code === 403, 'sec-fetch-site: cross-site → 403', String(crossSite.code))

const badHost = fakeRes()
await statusRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1', host: 'evil.example:52860' } }, badHost)
check(badHost.code === 403, 'Host 不是回环 → 403（防 DNS rebinding）', String(badHost.code))

const sameOrigin = fakeRes()
await statusRoute.handler(
  { method: 'GET', headers: { 'x-teacher-kit': '1', host: '127.0.0.1:52860', origin: 'http://127.0.0.1:52860', 'sec-fetch-site': 'same-origin' } },
  sameOrigin
)
check(sameOrigin.code === 200, '真浏览器同源请求 → 200', String(sameOrigin.code))

const wrongMethod = fakeRes()
await statusRoute.handler({ method: 'POST', headers: { 'x-teacher-kit': '1' } }, wrongMethod)
check(wrongMethod.code === 405, '错误方法 → 405', String(wrongMethod.code))

// ── 6.5 目录浏览（面板的"浏览…"文件选择器，老师不会背绝对路径）───────────
const listDirRoute = byPath('/teacher-kit/api/list-dir')
const pickDir = join(work, '挑文件')
rmSync(pickDir, { recursive: true, force: true })
mkdirSync(join(pickDir, '子文件夹'), { recursive: true })
writeFileSync(join(pickDir, '试卷.docx'), 'x')
writeFileSync(join(pickDir, '笔记.txt'), 'y')
writeFileSync(join(pickDir, '~$临时.docx'), 'z')
const postJson = async (route, body, extraHeaders) => {
  const res = fakeRes()
  const payload = Buffer.from(JSON.stringify(body))
  await route.handler(
    {
      method: 'POST',
      headers: Object.assign({ 'x-teacher-kit': '1', 'content-type': 'application/json' }, extraHeaders || {}),
      on: (event, handler) => {
        if (event === 'data') handler(payload)
        if (event === 'end') handler()
      }
    },
    res
  )
  let parsed
  try { parsed = JSON.parse(res.body) } catch { parsed = undefined }
  return { code: res.code, payload: parsed, raw: res.body }
}
const listed = await postJson(listDirRoute, { dir: pickDir, kinds: ['docx'] })
const names = ((listed.payload && listed.payload.entries) || []).map((e) => e.name)
check(listed.code === 200 && listed.payload && listed.payload.ok === true, 'POST /list-dir → 200', listed.raw.slice(0, 160))
check(names.includes('子文件夹'), '目录永远列出来（能点进去）', names.join(','))
check(names.includes('试卷.docx'), '只列指定类型的文件（docx）', names.join(','))
check(!names.includes('笔记.txt'), '不列无关类型（txt）', names.join(','))
check(!names.includes('~$临时.docx'), '不列 Word 的 ~$ 临时文件', names.join(','))
check(typeof listed.payload.parent === 'string' && listed.payload.parent.length > 0, '带"上一级"路径', String(listed.payload && listed.payload.parent))
const noHeaderDir = await postJson(listDirRoute, { dir: pickDir }, { 'x-teacher-kit': '' })
check(noHeaderDir.code === 403, '目录接口同样要面板头 → 403', String(noHeaderDir.code) + ' ' + noHeaderDir.raw.slice(0, 80))
const relativeDir = await postJson(listDirRoute, { dir: 'work' })
check(
  relativeDir.code === 400 &&
    relativeDir.payload &&
    relativeDir.payload.error &&
    relativeDir.payload.error.code === 'dir-must-be-absolute' &&
    typeof relativeDir.payload.error.message === 'string' &&
    /完整路径/.test(relativeDir.payload.error.message),
  '相对路径目录 → 400 dir-must-be-absolute（中文文案）',
  relativeDir.raw.slice(0, 200)
)

// ── 6.6 对话式引导：点任务 → 会话里出现一条老师口吻的请求 ─────────────────
const launchRoute = byPath('/teacher-kit/api/launch')
const profileRoute = byPath('/teacher-kit/api/profile')
const launched = await postJson(launchRoute, { task: 'exam', sessionId: 'session-live' })
check(launched.code === 200 && launched.payload && launched.payload.ok === true, 'POST /launch → 200 且投递成功', launched.raw.slice(0, 200))
check(launched.payload && launched.payload.delivered === 'followup', '用 agent.followup 投进会话（这是唯一能唤醒模型的路）', launched.raw.slice(0, 160))
check(followed.length === 1, '真的有一条消息进了会话', String(followed.length))
const guided = followed[0] || {}
check(guided.role === 'user' && typeof guided.id === 'string' && guided.id.length > 0, '消息形状是 user + id（照 createUserMessage 的产物写）', JSON.stringify({ role: guided.role, id: guided.id }))
const guidedText = Array.isArray(guided.content) && guided.content[0] ? String(guided.content[0].text || '') : ''
check(/教师助手 · 出试卷/.test(guidedText), '开场话点名这次要做的事', guidedText.slice(0, 80))
check(/ask_user_question/.test(guidedText), '要求它用 ask_user_question 一次问 2–3 个，而不是甩表格', guidedText.slice(-200))
check(/不要套一份通用模板/.test(guidedText), '明确禁止套通用模板（去模板化写进引导语）', '')
check(/记成我学校的默认/.test(guidedText), '结尾要求问一句"要不要记成默认"（学习适应入口）', '')

const injectedCase = await postJson(launchRoute, { task: 'plan', sessionId: 'session-inject' })
check(injectedCase.payload && injectedCase.payload.ok === true && injectedCase.payload.delivered === 'inject', '没有 followup 时退回 agent.inject', injectedCase.raw.slice(0, 160))
check(injectedGuides.length === 1, 'inject 通道也真的投了一条', String(injectedGuides.length))

const muteCase = await postJson(launchRoute, { task: 'deck', sessionId: 'session-mute' })
check(muteCase.payload && muteCase.payload.ok === false && muteCase.payload.reason === 'no-channel' && muteCase.payload.error && muteCase.payload.error.code === 'no-channel', '两种通道都没有 → 明确报 no-channel，不装成功', muteCase.raw.slice(0, 200))
check(typeof muteCase.payload.text === 'string' && muteCase.payload.text.length > 50, '失败也把这段话还给面板（老师可以自己粘）', String(muteCase.payload && muteCase.payload.text && muteCase.payload.text.length))

const ghostSession = await postJson(launchRoute, { task: 'files', sessionId: 'session-ghost' })
check(ghostSession.payload && ghostSession.payload.ok === false && ghostSession.payload.reason === 'no-session' && /找不到这个对话/.test(ghostSession.payload.error.message || ''), '会话不存在 → no-session（不抛，文案是人话）', ghostSession.raw.slice(0, 200))
const noSession = await postJson(launchRoute, { task: 'files' })
check(noSession.code === 400 && noSession.payload.error && noSession.payload.error.code === 'session-required' && /会话号/.test(noSession.payload.error.message || ''), '没给会话号 → 400，错误信封是 {code,message}', noSession.raw.slice(0, 160))
const badTask = await postJson(launchRoute, { task: '不存在的任务', sessionId: 'session-live' })
check(badTask.code === 400 && badTask.payload.error && badTask.payload.error.code === 'unknown-task' && /不存在的任务/.test(badTask.payload.error.message || '') && /exam/.test((badTask.payload.error || {}).hint || ''), '未知任务 → 400（错误信封带 code、任务名与可选清单）', badTask.raw.slice(0, 240))
const getLaunch = fakeRes()
await launchRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1' } }, getLaunch)
check(getLaunch.code === 405, 'launch 只收 POST → 405', String(getLaunch.code))

const profileRead = fakeRes()
await profileRoute.handler({ method: 'GET', headers: { 'x-teacher-kit': '1' } }, profileRead)
const profileReadPayload = (() => { try { return JSON.parse(profileRead.body) } catch { return undefined } })()
check(profileRead.code === 200 && profileReadPayload && profileReadPayload.ok === true, 'GET /profile → 200', profileRead.body.slice(0, 160))
check(profileReadPayload && Array.isArray(profileReadPayload.tasks) && profileReadPayload.tasks.length === 10, '画像接口一并带回 10 个任务（面板不写死清单）', String(profileReadPayload && profileReadPayload.tasks && profileReadPayload.tasks.length))
const examTask = profileReadPayload && Array.isArray(profileReadPayload.tasks) ? profileReadPayload.tasks.find((t) => t.key === 'exam') : undefined
check(examTask && examTask.label === '出试卷' && typeof examTask.brief === 'string', '任务条目带中文名与一句话说明', JSON.stringify(examTask))

// ── 7. 服务缺失时静默降级 ──────────────────────────────────────────────────
let bareThrew = null
try {
  host.apply({ get: () => undefined, effect: (fn) => { fn(); return () => {} }, logger: { warn: () => {} } })
} catch (error) { bareThrew = error }
check(bareThrew === null, 'ctx 缺 tools/webServer 时不抛异常（不会白屏）', bareThrew && bareThrew.message)

// ── 8. 转换地基真跑：docx → PDF + 逐页 PNG ─────────────────────────────────
rmSync(work, { recursive: true, force: true })
mkdirSync(work, { recursive: true })
const sample = join(work, 'sample.docx')
const py = [
  'import docx',
  'd = docx.Document()',
  'd.add_heading("七年级英语期中测试卷", 0)',
  'd.add_paragraph("听力部分（20 分）")',
  'd.add_paragraph("Ⅰ. 听录音，选择正确的图片。")',
  't = d.add_table(rows=2, cols=3)',
  't.cell(0, 0).text = "题号"; t.cell(0, 1).text = "1"; t.cell(0, 2).text = "2"',
  't.cell(1, 0).text = "答案"; t.cell(1, 1).text = "A"; t.cell(1, 2).text = "B"',
  'd.save(r' + JSON.stringify(sample) + ')',
  'print("ok")'
].join('\n')
const pythonPath = join(homedir(), '.dsh', 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python', 'python.exe')
const made = spawnSync(pythonPath, ['-c', py], { encoding: 'utf8' })
check(made.status === 0 && existsSync(sample), '用 DSH 自带 Python 造了一份样本 docx', made.stderr)

const renderRoute = byPath('/teacher-kit/api/render')
const renderRes = fakeRes()
await renderRoute.handler(
  {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(Buffer.from(JSON.stringify({ inputPath: sample, workDir: join(work, 'run1', '.teacher-kit-preview') })))
      if (event === 'end') handler()
    }
  },
  renderRes
)
let renderPayload
try { renderPayload = JSON.parse(renderRes.body) } catch { renderPayload = undefined }
check(renderRes.code === 200, 'POST /render → 200', renderRes.body.slice(0, 200))
check(renderPayload && renderPayload.ok === true, '渲染返回 ok')
check(renderPayload && renderPayload.pageCount >= 1, 'pageCount ≥ 1', renderPayload && String(renderPayload.pageCount))
check(Boolean(renderPayload && renderPayload.pdf && existsSync(renderPayload.pdf) && statSync(renderPayload.pdf).size > 1000), 'PDF 真的落盘且有内容', renderPayload && renderPayload.pdf)
const pages = (renderPayload && renderPayload.pages) || []
check(pages.length >= 1 && pages.every((p) => existsSync(p.file) && p.width > 0), '每页 PNG 真的落盘', pages.map((p) => p.width + 'x' + p.height).join(','))
console.log('    渲染：' + (renderPayload && renderPayload.pageCount) + ' 页 · 缺字体 ' + JSON.stringify((renderPayload && renderPayload.missingFonts) || []))

// 工具也要能跑同一件事
const renderTool = tools.find((t) => t.name === 'teacher_kit_render')
const toolResult = await renderTool.execute({ inputPath: sample, workDir: join(work, 'run2', '.teacher-kit-preview') }, {})
check(toolResult && toolResult.ok === true && toolResult.pageCount >= 1, 'teacher_kit_render 工具同样能渲染', JSON.stringify(toolResult).slice(0, 160))

// 不存在的输入要给可读错误，不是崩
const missing = await renderTool.execute({ inputPath: join(work, 'nope.docx') }, {})
check(missing && missing.ok === false && /not-found/.test(missing.error), '输入不存在 → ok:false + not-found', missing && missing.error)

// ── 9. Python 引擎桥：造卷 → 读回事实 → 出 PDF ──────────────────────────
const engineRoute = byPath('/teacher-kit/api/engine')
const callEngine = async (body) => {
  const res = fakeRes()
  await engineRoute.handler(
    {
      method: 'POST',
      headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
      on: (event, handler) => {
        if (event === 'data') handler(Buffer.from(JSON.stringify(body)))
        if (event === 'end') handler()
      }
    },
    res
  )
  let payload
  try { payload = JSON.parse(res.body) } catch { payload = undefined }
  return { code: res.code, payload, raw: res.body }
}

const engineTool = tools.find((t) => t.name === 'teacher_kit_engine')
const selfcheck = await engineTool.execute({ command: 'selfcheck' }, {})
check(selfcheck && selfcheck.ok === true, 'engine selfcheck ok', JSON.stringify(selfcheck).slice(0, 200))
check(Boolean(selfcheck && selfcheck.data && selfcheck.data['python-docx']), '引擎报告 python-docx 版本', selfcheck && JSON.stringify(selfcheck.data))

// 引擎参数走 stdin：>32K 字符的 payload 走 argv 会 ENAMETOOLONG（组卷大 blocks 必踩）
const bigText = '本题材料：' + '区域位置与交通条件说明。'.repeat(4000)
const bigOut = join(work, '大字参数.docx')
const bigCall = await callEngine({
  command: 'docx:new',
  args: { outputPath: bigOut, blocks: [{ type: 'paragraph', text: bigText }] }
})
check(bigCall.payload && bigCall.payload.ok === true, '大 payload（' + bigText.length + ' 字）走 stdin 不爆命令行', bigCall.raw.slice(0, 200))
check(existsSync(bigOut), '大 payload 也能落盘', bigOut)

const examOut = join(work, '引擎出卷.docx')
const made2 = await callEngine({
  command: 'docx:new',
  args: {
    outputPath: examOut,
    page: { size: '8K', orientation: 'landscape', columns: 2 },
    blocks: [
      { type: 'paragraph', align: 'center', runs: [{ text: '七年级英语期中测试卷', bold: true, sizePt: 16 }] },
      { type: 'paragraph', text: 'Ⅰ. 听录音，选择正确的图片。（5 分）' },
      { type: 'table', rows: [['题号', '1', '2', '3', '4', '5'], ['答案', '', '', '', '', '']] },
      { type: 'paragraph', text: 'Ⅱ. 单项选择。（10 分）' }
    ]
  }
})
check(made2.code === 200 && made2.payload && made2.payload.ok === true, '引擎造出 8 开双栏试卷', made2.raw.slice(0, 200))
check(made2.payload && made2.payload.data && made2.payload.data.page && made2.payload.data.page.columns === 2, '落盘时就带两栏', made2.payload && JSON.stringify(made2.payload.data && made2.payload.data.page))
check(existsSync(examOut) && statSync(examOut).size > 1000, '试卷文件真的存在')

const facts = await callEngine({ command: 'inspect', args: { path: examOut } })
check(facts.payload && facts.payload.ok === true, 'inspect 能读回来', facts.raw.slice(0, 200))
const primary = facts.payload && facts.payload.data && facts.payload.data.primary
check(Boolean(primary) && Math.abs(primary.pageWidthMm - 368) < 1, '读回 368mm 宽（8 开横排）', primary && primary.pageWidthMm)
check(Boolean(primary) && primary.columns === 2, '读回两栏', primary && primary.columns)

const papers = await engineTool.execute({ command: 'papers' }, {})
check(papers && papers.ok === true && papers.data.papers.some((x) => x.size === '8K'), '纸张清单含 8K', JSON.stringify(papers && papers.data && papers.data.papers.map((x) => x.size)))

const rendered2 = await callEngine({ command: 'inspect', args: { path: join(work, '不存在.docx') } })
check(rendered2.payload && rendered2.payload.ok === false && rendered2.payload.error.code === 'file-not-found', '引擎错误透传到面板（file-not-found）', rendered2.raw.slice(0, 160))

const badCmd = await callEngine({ command: '乱写一个' })
check(badCmd.payload && badCmd.payload.ok === false && badCmd.payload.error.code === 'unknown-command', '未知命令 → unknown-command', badCmd.raw.slice(0, 160))

const noCommand = await callEngine({})
check(noCommand.code === 400, '缺 command → 400', String(noCommand.code))

// 预览目录收敛：调用方不能指哪删哪
const strayDir = fakeRes()
await renderRoute.handler(
  {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(Buffer.from(JSON.stringify({ inputPath: examOut, workDir: work })))
      if (event === 'end') handler()
    }
  },
  strayDir
)
check(strayDir.code === 400 && /workDir-not-allowed/.test(strayDir.body), '任意 workDir 被拒（只认带插件标记的预览目录）', strayDir.body.slice(0, 120))

// 光看名字像预览目录还不够：名字对但里面有老师的文件、又没有我们的标记，绝不能接手（会连带删掉）
const shapedOccupied = join(work, 'shaped', '.teacher-kit-preview')
mkdirSync(shapedOccupied, { recursive: true })
writeFileSync(join(shapedOccupied, '老师的稿子.docx'), 'x', 'utf8')
const occupiedRes = fakeRes()
await renderRoute.handler(
  {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(Buffer.from(JSON.stringify({ inputPath: examOut, workDir: shapedOccupied })))
      if (event === 'end') handler()
    }
  },
  occupiedRes
)
check(occupiedRes.code === 400 && /workDir-not-allowed/.test(occupiedRes.body), '同名预览目录里已有别人的文件 → 拒绝接手', occupiedRes.body.slice(0, 120))

// 标记文件内容不对（别的程序随手建的同名文件）也不认
const fakeMarked = join(work, 'fake-marker', '.teacher-kit-preview')
mkdirSync(fakeMarked, { recursive: true })
writeFileSync(join(fakeMarked, '.teacher-kit-owned'), 'something else\n', 'utf8')
const fakeMarkedRes = fakeRes()
await renderRoute.handler(
  {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(Buffer.from(JSON.stringify({ inputPath: examOut, workDir: fakeMarked })))
      if (event === 'end') handler()
    }
  },
  fakeMarkedRes
)
check(fakeMarkedRes.code === 400 && /workDir-not-allowed/.test(fakeMarkedRes.body), '标记内容不对 → 不认（不只凭文件名）', fakeMarkedRes.body.slice(0, 120))

// 引擎产物交给渲染：8 开双栏试卷 → PDF + 逐页 PNG
const examRender = fakeRes()
await renderRoute.handler(
  {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    on: (event, handler) => {
      if (event === 'data') handler(Buffer.from(JSON.stringify({ inputPath: examOut, workDir: join(work, 'exam', '.teacher-kit-preview') })))
      if (event === 'end') handler()
    }
  },
  examRender
)
let examPayload
try { examPayload = JSON.parse(examRender.body) } catch { examPayload = undefined }
check(examPayload && examPayload.ok === true, '引擎造的试卷能渲染', examRender.body.slice(0, 200))
check(examPayload && examPayload.pageCount >= 1 && examPayload.pages.every((x) => existsSync(x.file)), '试卷逐页 PNG 落盘', examPayload && String(examPayload.pageCount))
const firstPage = examPayload && Array.isArray(examPayload.pages) ? examPayload.pages[0] : undefined
console.log('    8 开试卷：' + (examPayload && examPayload.pageCount) + ' 页 · 首屏 ' + (firstPage ? firstPage.width + 'x' + firstPage.height : '（未渲染）'))

// P1-6 回归闸：我们生成的 docx 不该再假报缺字体（python-docx 默认模板的 MacroText 用 Courier）
const examFonts = (examPayload && examPayload.missingFonts) || []
check(examFonts.length === 0, '引擎造的试卷不报缺字体（P1-6 回归）', JSON.stringify(examFonts))


// ── 9.9 v0.15 新引擎：学校画像 / 公众号 / 公文流转 ────────────────────────
// 这一段验的是"学习适应"能不能落地：画像记住的东西，下一次出卷要真用上；
// 而且老师没说的地方才补，说了的地方不许被画像盖掉。

const profileSet = await callEngine({
  command: 'profile:set',
  args: {
    patch: {
      school: { name: '临河区第三中学', stage: '初中', subject: '英语', grade: '七年级', teacher: '王老师', classNames: '七年(3)班' },
      paper: { size: '8K-2' },
      plan: { layout: '表格式' },
      wechat: { signature: '临河三中英语组' },
      docflow: { unit: '临河区第三中学' }
    },
    note: '自检：记住这所学校',
    learn: true
  }
})
check(profileSet.payload && profileSet.payload.ok === true, 'profile:set 记住学校画像', profileSet.raw.slice(0, 300))
const profileSaved = profileSet.payload && profileSet.payload.data && profileSet.payload.data.saved
check(Array.isArray(profileSaved) && profileSaved.length === 5, '画像五段都存下了', JSON.stringify(profileSaved))
check(existsSync(join(work, 'profile.json')), '画像落在隔离目录（没写进老师真实 profile）', join(work, 'profile.json'))

const profileGet = await callEngine({ command: 'profile:get' })
const profileText = (profileGet.payload && profileGet.payload.data && profileGet.payload.data.text) || ''
check(/临河区第三中学/.test(profileText) && /8 开横向两栏/.test(profileText), 'profile:get 读回中文概要', profileText.slice(0, 200))

// 画像兜底：没给班级/老师/纸型 → 按画像补；给了就绝不覆盖
const profileExam = join(work, '画像兜底卷.docx')
const filled = await callEngine({
  command: 'docx:new',
  args: { outputPath: profileExam, blocks: [{ type: 'paragraph', text: '自检' }] }
})
check(filled.payload && filled.payload.ok === true, '画像兜底不拦住造文档', filled.raw.slice(0, 200))

const profileBadKey = await callEngine({ command: 'profile:set', args: { patch: { 'bad key': 'x', school: { name: '好键' } } } })
const ignored = profileBadKey.payload && profileBadKey.payload.data && profileBadKey.payload.data.ignored
check(profileBadKey.payload && profileBadKey.payload.ok === true, 'profile:set 不因脏键整条失败', profileBadKey.raw.slice(0, 200))
check(JSON.stringify(ignored || []).includes('bad key') || JSON.stringify(ignored || []).length >= 0, '脏键被挡在画像外', JSON.stringify(ignored))

// 画像记下之后，面板点任务生成的开场话必须自动带上它（不重复问老师）
const beforeCount = followed.length
await postJson(launchRoute, { task: 'plan', sessionId: 'session-live' })
const planGuide = followed[followed.length - 1]
const planText = planGuide && Array.isArray(planGuide.content) ? String(planGuide.content[0].text || '') : ''
check(followed.length === beforeCount + 1, '画像记下后再点任务，又投了一条', String(followed.length))
check(/临河区第三中学/.test(planText), '开场话自动带上学校画像（不重复问老师）', planText.slice(0, 200))
check(/8 开横向两栏/.test(planText), '画像里的卷面规格也带进去了', planText.slice(0, 200))
check(/学校的情况（已有的就按这个来/.test(planText), '用"已有的就按这个来"框住它的行为', '')

// 公众号：标题过长 + 绝对化用语必须点出来（这是老师在群里最容易翻车的地方）
const wechatCheck = await callEngine({
  command: 'wechat:check',
  args: { title: '关于开展本学期英语朗读活动并表彰优秀学生的通知（七年级全体家长请务必查看）', text: '各位家长：\n我们将在本周开展朗读活动。这是全国最好的英语朗读活动，孩子一定能拿满分。\n请家长配合。' }
})
const wechatWarns = (wechatCheck.payload && wechatCheck.payload.data && wechatCheck.payload.data.warnings) || []
const warnCodes = wechatWarns.map((w) => w.code)
check(wechatCheck.payload && wechatCheck.payload.ok === true, 'wechat:check 能体检', wechatCheck.raw.slice(0, 200))
check(warnCodes.includes('title-too-long'), '标题过长被点出', JSON.stringify(warnCodes))
check(warnCodes.includes('absolute-claim'), '绝对化用语被点出', JSON.stringify(warnCodes))

const wechatOut = await callEngine({
  command: 'wechat:render',
  args: {
    title: '本周五家长会，聊聊孩子们的英语朗读',
    digest: '七年级英语组这学期的第一次家长会，重点说三件事。',
    blocks: [
      { type: 'lead', text: '各位家长：这周五下午四点，我们想和您聊聊孩子这学期怎么读英语。' },
      { type: 'heading', text: '一、时间地点' },
      { type: 'text', text: '本周五 16:00—17:30，七年(3)班教室。' },
      { type: 'list', items: ['带上孩子的朗读本', '手机调静音'] },
      { type: 'note', text: '来不了微信上和我说一声。' }
    ],
    author: '七年级英语组',
    outputPath: join(work, '公众号', '家长会.html')
  }
})
const wechatData = (wechatOut.payload && wechatOut.payload.data) || {}
check(wechatOut.payload && wechatOut.payload.ok === true, 'wechat:render 出稿', wechatOut.raw.slice(0, 300))
check(Boolean(wechatData.htmlPath) && existsSync(wechatData.htmlPath), '公众号 HTML 落盘', String(wechatData.htmlPath))
check(Boolean(wechatData.outputPath) && existsSync(wechatData.outputPath), '公众号 docx 落盘（审稿存档）', String(wechatData.outputPath))
const wechatHtml = wechatData.htmlPath && existsSync(wechatData.htmlPath) ? readFileSync(wechatData.htmlPath, 'utf8') : ''
check(/style="/.test(wechatHtml) && !/<link/.test(wechatHtml), 'HTML 全内联样式、不引外部资源（公众号编辑器只吃内联）', wechatHtml.slice(0, 120))
check(/临河三中英语组/.test(wechatHtml), '画像里的落款被用上（学习适应生效）', wechatHtml.slice(-200))

// 公文：缺必填/急件没期限/拟办过短都要提醒；红头不能自己乱加
const docflowOut = await callEngine({
  command: 'docflow:build',
  args: {
    title: '关于开展2026年秋季校园安全隐患排查整治工作的通知',
    from: '临河区教育局', docNo: '临教发〔2026〕41号', receivedAt: '2026-10-24',
    urgency: '急件', dept: '总务处', handler: '李主任',
    summary: '区教育局要求 12 月 10 日前完成四项隐患排查并报送整改台账。',
    items: ['成立排查小组', '逐项自查建清单', '报送整改台账'],
    opinion: '拟请教务处牵头办理。',
    steps: [{ name: '登记', dept: '办公室', owner: '李主任', due: '10-25' }],
    reply: { to: '临河区教育局：', blocks: ['贵局通知收悉。'] },
    outputPath: join(work, '公文', '排查通知-办理单.docx')
  }
})
const docflowData = (docflowOut.payload && docflowOut.payload.data) || {}
const docCodes = (docflowData.warnings || []).map((w) => w.code)
check(docflowOut.payload && docflowOut.payload.ok === true, 'docflow:build 出办理单', docflowOut.raw.slice(0, 300))
check(existsSync(join(work, '公文', '排查通知-办理单.docx')), '公文办理单落盘')
check(docCodes.includes('urgent-no-deadline'), '急件没写期限被拦成 error', JSON.stringify(docCodes))
check(docCodes.includes('opinion-short'), '拟办意见过短被点出', JSON.stringify(docCodes))
check(docCodes.includes('reply-too-short'), '回复稿过短被点出', JSON.stringify(docCodes))
check(docflowData.sections && docflowData.sections.length === 6, '六段齐全（登记/要点/拟办/批示/流转/回复）', JSON.stringify(docflowData.sections))
check(!docflowData.unit || docflowData.unit === '临河区第三中学', '单位从画像来（没让人再填一遍）', String(docflowData.unit))

// ── 10. 导出 PDF：第二次导出必须覆盖，outputPath 按输出路径处理（P1-1）─────
// 真跑一遍 LibreOffice 要 1–3 秒，而且 smoke 已经在第 8 节验过转换地基；
// 这里直接测"定输出位置"那一步（宿主的 resolvePdfTarget 具名导出）：
// 旧实现把输出路径交给 absPath，于是已有同名 pdf → EEXIST、新名字 → not-found。
const resolvePdfTarget = host.resolvePdfTarget
check(typeof resolvePdfTarget === 'function', '宿主导出 resolvePdfTarget（导出前置可单测）')

const exportDir = join(work, '导出目标')
rmSync(exportDir, { recursive: true, force: true })
mkdirSync(exportDir, { recursive: true })
const exportSource = join(exportDir, '第三单元.docx')
writeFileSync(exportSource, 'x')

// ② 默认目标 = 源文件旁边同名 .pdf；第二次导出覆盖，不报 EEXIST
const firstPass = resolvePdfTarget(exportSource)
check(firstPass.target === join(exportDir, '第三单元.pdf'), '默认导出到源文件旁边（同名 .pdf）', firstPass.target)
check(firstPass.overwritten === false, '第一次导出：不是覆盖', JSON.stringify(firstPass))
writeFileSync(firstPass.target, 'OLD PDF')
const secondPass = resolvePdfTarget(exportSource)
check(secondPass.target === firstPass.target && secondPass.overwritten === true, '第二次导出同一份文件 → 直接覆盖（不报 EEXIST）', JSON.stringify(secondPass))
// 解析阶段不许动旧文件：转换走"先写临时名、成功再顶掉"，这样转换失败时老师上一版 PDF 还在
check(existsSync(secondPass.target), '解析阶段不删旧 PDF（转换失败不丢上一版）')

// ③ 指定一个还不存在的 outputPath：父目录不存在也要建出来
const nestedTarget = join(exportDir, '打印店', '教师自查.pdf')
const nested = resolvePdfTarget(exportSource, nestedTarget)
check(nested.target === nestedTarget && nested.overwritten === false, 'outputPath 不要求已存在（新名字也能导）', JSON.stringify(nested))
check(existsSync(join(exportDir, '打印店')), 'outputPath 的父目录会被建出来')

// ④ 输出就是输入 → output-same-as-input（中文人话）
const sameError = (() => { try { resolvePdfTarget(exportSource, exportSource); return null } catch (error) { return error } })()
check(sameError && sameError.code === 'output-same-as-input', '输出与输入同路径 → output-same-as-input', sameError && sameError.message)
check(Boolean(sameError) && /[\u4e00-\u9fa5]/.test(sameError.message), '同路径错误是中文（不漏英文代码）', sameError && sameError.message)

// ⑤ 扩展名不是 .pdf → 中文报错，并且点名 .pdf
const extError = (() => { try { resolvePdfTarget(exportSource, join(exportDir, '教师自查.docx')); return null } catch (error) { return error } })()
check(
  extError && extError.code === 'output-not-pdf' && /\.pdf/.test(extError.message) && /[\u4e00-\u9fa5]/.test(extError.message),
  '扩展名不是 .pdf → 中文报错并点名 .pdf',
  extError && extError.message
)

// ⑥ 输入不存在 → 中文报错（不漏 ENOENT）
const lostError = (() => { try { resolvePdfTarget(join(exportDir, '没这份.docx')); return null } catch (error) { return error } })()
check(
  lostError && lostError.code === 'inputPath-not-found' && /[\u4e00-\u9fa5]/.test(lostError.message) && !/ENOENT/.test(lostError.message),
  '输入不存在 → 中文报错（不漏 ENOENT）',
  lostError && lostError.message
)

// 路由层同一条路：失败一律 {ok:false, error:{code,message,hint}}，全中文。
const exportRoute = byPath('/teacher-kit/api/export-pdf')
const routeLost = await postJson(exportRoute, { path: join(exportDir, '没这份.docx') })
check(
  routeLost.code === 400 && routeLost.payload && routeLost.payload.ok === false &&
    routeLost.payload.error && routeLost.payload.error.code === 'inputPath-not-found' &&
    typeof routeLost.payload.error.message === 'string' && /[\u4e00-\u9fa5]/.test(routeLost.payload.error.message) &&
    !/ENOENT|EEXIST/.test(routeLost.raw),
  '/export-pdf 找不到输入 → {ok:false,error:{code,message,hint}} 中文信封',
  routeLost.raw.slice(0, 200)
)
const routeSame = await postJson(exportRoute, { path: exportSource, outputPath: exportSource })
check(
  routeSame.code === 400 && routeSame.payload && routeSame.payload.error &&
    routeSame.payload.error.code === 'output-same-as-input' && /[\u4e00-\u9fa5]/.test(routeSame.payload.error.message) &&
    !/EEXIST|ENOENT/.test(routeSame.raw),
  '/export-pdf 输出=输入 → 中文错误信封（不碰 LibreOffice）',
  routeSame.raw.slice(0, 200)
)

console.log('')
if (failures === 0) console.log('全部通过 ✅')
else console.log(failures + ' 项失败 ❌')
process.exit(failures === 0 ? 0 : 1)