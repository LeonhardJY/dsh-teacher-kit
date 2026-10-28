// 教师全流程自测（发布前排雷用）：把 7 张卡背后的引擎命令全跑一遍，走真实 HTTP 接口。
// 用法：先启动一个实例，再 node scripts/flow-check.mjs <端口>
//   端口 = 实例启动日志里那句 `dsh web: http://127.0.0.1:<端口>/?token=…`
//   产物落在 .verify/flow/，最后一行打印「全流程自测通过 ✅ N 项」。
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'

const port = process.argv[2] || '50325'
const base = 'http://127.0.0.1:' + port + '/teacher-kit/api'
const root = process.cwd()
const tmp = path.join(root, '.verify', 'flow')
fs.rmSync(tmp, { recursive: true, force: true })
fs.mkdirSync(tmp, { recursive: true })

// 自检会写学校画像（profile:set learn:true）。老师机器上跑一次就把真实画像换成「临河三中」了，
// 所以开跑前留一份，结束（含异常退出）原样还原（审查 C-18）。
function resolveProfilePath() {
  if (process.env.TEACHER_KIT_PROFILE) return path.resolve(process.env.TEACHER_KIT_PROFILE)
  const dshHome = path.join(os.homedir(), '.dsh', 'teacher-kit')
  if (fs.existsSync(dshHome)) return path.join(dshHome, 'profile.json')
  return path.join(os.homedir(), 'Documents', '教师助手', 'profile.json')
}
const profilePath = resolveProfilePath()
const profileBackup = fs.existsSync(profilePath) ? fs.readFileSync(profilePath) : null
let profileRestored = false
function restoreProfile() {
  if (profileRestored) return
  profileRestored = true
  try {
    if (profileBackup) {
      fs.mkdirSync(path.dirname(profilePath), { recursive: true })
      fs.writeFileSync(profilePath, profileBackup)
    } else if (fs.existsSync(profilePath)) {
      fs.rmSync(profilePath, { force: true })
    }
  } catch (e) {
    console.log('  !!   画像没还原：' + e.message + '（自检写的学校在 ' + profilePath + '，手动删掉重设）')
  }
}
process.on('exit', restoreProfile)
process.on('SIGINT', function () { restoreProfile(); process.exit(130) })
console.log('画像文件：' + profilePath + (process.env.TEACHER_KIT_PROFILE ? '（环境变量指定）' : '（跑完自动还原）'))

let pass = 0
let fail = 0
function check(cond, label, detail) {
  if (cond) { pass += 1; console.log('  ok   ' + label) }
  else { fail += 1; console.log('  FAIL ' + label + (detail ? '  << ' + String(detail).slice(0, 200) : '')) }
}

async function call(pathname, body) {
  const t0 = Date.now()
  const res = await fetch(base + pathname, {
    method: 'POST',
    headers: { 'x-teacher-kit': '1', 'content-type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body)
  })
  const text = await res.text()
  let parsed = null
  try { parsed = JSON.parse(text) } catch (e) { parsed = { raw: text.slice(0, 200) } }
  return { status: res.status, body: parsed, ms: Date.now() - t0 }
}

async function callGet(pathname) {
  const t0 = Date.now()
  const res = await fetch(base + pathname, { headers: { 'x-teacher-kit': '1' } })
  const text = await res.text()
  let parsed = null
  try { parsed = JSON.parse(text) } catch (e) { parsed = { raw: text.slice(0, 200) } }
  return { status: res.status, body: parsed, ms: Date.now() - t0 }
}

async function engine(command, args) {
  const out = await call('/engine', { command, args })
  const body = out.body || {}
  const data = body.data && typeof body.data === 'object' ? body.data : body
  return { ok: body.ok === true, data, error: body.error, status: out.status, ms: out.ms }
}

function timing(name, r) { console.log('  ·    ' + name + '  ' + r.ms + 'ms') }

const examText = fs.readFileSync(path.join(root, '.verify', 'exam_text.txt'), 'utf8')
console.log('教师全流程自测 @' + base)

// 1. 出试卷（8 开两栏 + 答题卡 + 答案页）
const examOut = path.join(tmp, 'exam.docx')
let r = await engine('exam:build', { text: examText, outputPath: examOut, page: { size: '8K', orientation: 'landscape', columns: 2 }, options: { minutes: 90, answerCard: 'card', answerSheet: 'end' } })
timing('exam:build', r)
check(r.ok, '出试卷 接口返回 ok', r.error)
check(r.ok && r.data.questionCount >= 30, '出卷认出 ≥30 题', r.data && r.data.questionCount)
check(r.ok && r.data.preflight && r.data.preflight.level !== 'fail', '印前检查不拦', r.data && r.data.preflight && r.data.preflight.summary)
check(fs.existsSync(examOut), '试卷 docx 落盘')

// 2. A/B 两套
r = await engine('exam:build-ab', { text: examText, outputPath: path.join(tmp, 'ab.docx'), page: { size: '16K', orientation: 'portrait', columns: 1 } })
timing('exam:build-ab', r)
check(r.ok && r.data.variants && r.data.variants.length === 2, 'A/B 两套都出来', r.error)
check(r.ok && r.data.variants.every((v) => fs.existsSync(v.outputPath)), 'A/B 两个文件都落盘')

// 3. 学生卷 + 教师卷
r = await engine('exam:build-teacher', { text: examText, outputPath: path.join(tmp, 'teacher.docx'), page: { size: 'A4', orientation: 'portrait', columns: 1 } })
timing('exam:build-teacher', r)
check(r.ok && r.data.teacherPath && fs.existsSync(r.data.teacherPath), '教师卷落盘', r.error)

// 4. 题库组卷（直接拿刚才的卷子当"题库"? 不行，用 xlsx 样本）
const bank = path.join(root, '.verify', '题库.xlsx')
if (fs.existsSync(bank)) {
  r = await engine('bank:inspect', { path: bank })
  timing('bank:inspect', r)
  check(r.ok && r.data.count > 0, '题库读出题量', r.error)
  r = await engine('bank:build', { path: bank, outputPath: path.join(tmp, 'bank.docx'), page: { size: '16K', orientation: 'portrait', columns: 1 }, options: { pick: { 单项选择: 2 } } })
  timing('bank:build', r)
  check(r.ok && r.data.questionCount === 2, '按题型挑 2 题', r.data && r.data.questionCount)
}

// 5. 听写单（含音频）
const words = 'Unit 3 单词听写\n1. tall 高的\n2. short 矮的\n3. thin 瘦的'
r = await engine('handout:build', { text: words, outputPath: path.join(tmp, 'dict.docx'), mode: 'en2cn', columns: 2, answer: true })
timing('handout:build', r)
check(r.ok && r.data.count === 3, '听写单 3 条', r.data && r.data.count)
r = await engine('handout:audio', { text: words, outputPath: path.join(tmp, 'dict.wav'), repeat: 1, pauseSec: 4 })
timing('handout:audio', r)
check(r.ok && fs.existsSync(path.join(tmp, 'dict.wav')), '听写音频落盘', r.error)

// 5.5 听力录音稿 → 音频（中英混排各用对应语音）
const script = ['What does the boy want to buy?', '他想要买一个书包。', 'Where are they talking?'].join(String.fromCharCode(10))
r = await engine('audio:script', { text: script, outputPath: path.join(tmp, 'listening.wav'), pauseSec: 2, gapSec: 1 })
timing('audio:script', r)
check(r.ok && fs.existsSync(path.join(tmp, 'listening.wav')), '听力录音稿念成音频', r.error)
check(r.ok && r.data.seconds > 3 && r.data.lines === 3, '听力音频 3 句、有声音', r.data && [r.data.seconds, r.data.lines])

// 6. 教案
const plan = '课题：Unit 3 Section A 2d\n课型：听说课\n课时：1 课时\n教学目标：能用 What does he look like? 描述外貌\n重点：描述外貌的形容词\n教学过程：\n导入：看图说话\n新授：听录音跟读\n操练：两人对话\n小结：总结本课句型\n板书设计：What does he look like? — He is tall.'
r = await engine('plan:build', { text: plan, outputPath: path.join(tmp, 'plan.docx') })
timing('plan:build', r)
check(r.ok && r.data.fieldCount >= 3, '教案认出字段', r.data && r.data.fieldCount)

// 7. 课件（pptx + html）
const deck = '【Unit 3 Weather】\n一、Warm up\n讲：先问学生今天天气怎么样\n二、New words\nrain 雨\nsnow 雪\nwindy 有风的\n三、Practice\nWhat is the weather like?'
r = await engine('lesson:deck', { text: deck, outputPath: path.join(tmp, 'deck.pptx') })
timing('lesson:deck', r)
check(r.ok && r.data.slides >= 3, 'pptx 课件页数', r.data && r.data.slides)
r = await engine('lesson:html', { text: deck, outputPath: path.join(tmp, 'deck.html') })
timing('lesson:html', r)
check(r.ok && fs.existsSync(path.join(tmp, 'deck.html')), 'html 课件落盘', r.error)

// 8. 文件加工
const dirty = path.join(tmp, 'dirty.docx')
fs.copyFileSync(examOut, dirty)
r = await engine('docx:clean', { path: dirty, options: { font: '宋体' } })
timing('docx:clean', r)
check(r.ok, '净化不报错', r.error)
r = await engine('docx:merge', { inputs: [examOut, path.join(tmp, 'teacher.docx')], outputPath: path.join(tmp, 'merged.docx') })
timing('docx:merge', r)
check(r.ok && fs.existsSync(path.join(tmp, 'merged.docx')), '合并落盘', r.error)
r = await engine('docx:extract', { path: path.join(tmp, 'teacher.docx') })
timing('docx:extract', r)
check(r.ok && String(r.data.text || '').length > 200, '提取出文字', r.data && String(r.data.text || '').length)

// 9. 印前检查 + 模板反推
r = await engine('preflight', { path: examOut })
timing('preflight', r)
check(
  r.ok && Array.isArray(r.data.findings) && ['pass', 'warn', 'fail'].indexOf(r.data.level) >= 0,
  '检查给出 findings 与结论等级',
  r.data && JSON.stringify({ level: r.data.level, n: (r.data.findings || []).length })
)
r = await engine('template:make', { source: examOut, outputPath: path.join(tmp, 'tpl.docx') })
timing('template:make', r)
check(r.ok && fs.existsSync(path.join(tmp, 'tpl.docx')), '模板反推落盘', r.error)

// 10. 导出 PDF + 逐页预览
r = await call('/export-pdf', { path: examOut })
timing('/export-pdf', r)
const pdf = r.body && r.body.pdf
check(r.body && r.body.ok === true && pdf && fs.existsSync(pdf), '导出 PDF', r.body && r.body.error)
r = await call('/render', { inputPath: examOut, pdfPath: pdf })
timing('/render', r)
check(r.body && r.body.ok === true && r.body.pageCount > 0, '逐页预览出图', r.body && r.body.error)

// 10.5 v0.15：学校画像 / 公众号 / 公文流转（新增的三个引擎命令）
const profileHome = path.join(tmp, 'tk-home')
fs.mkdirSync(profileHome, { recursive: true })
r = await engine('profile:get', {})
timing('profile:get', r)
check(r.ok && typeof r.data.text === 'string', '画像能读（空画像也给一段话）', r.error)
r = await engine('profile:set', {
  patch: {
    school: { name: '临河区第三中学', stage: '初中', subject: '英语', grade: '七年级' },
    paper: { size: '8K', orientation: 'landscape', columns: 2, footer: '使用班级：{{className}}　出卷老师：{{teacherName}}' },
    plan: { layout: 'form' },
    wechat: { tone: '正式' },
    docflow: { unit: '临河区第三中学', dept: '教务处' }
  },
  learn: true,
  task: 'profile'
})
timing('profile:set', r)
check(r.ok && r.data.profile && r.data.profile.school.name === '临河区第三中学', '画像记下学校', r.error)
r = await engine('profile:get', {})
check(r.ok && /临河区第三中学/.test(r.data.text) && /8 开横向两栏/.test(r.data.text), '画像概要是人话（校名 + 纸型）', r.data && r.data.text && r.data.text.slice(0, 120))

const article = { title: '临河三中开展英语书写比赛', digest: '以赛促学，规范书写', blocks: [
  { type: 'lead', text: '10 月 20 日，我校七年级英语书写比赛在报告厅举行。' },
  { type: 'heading', text: '赛前：人人练字' },
  { type: 'text', text: '各班利用课前五分钟练习字母与句子书写，教师逐本圈点。' },
  { type: 'list', items: ['书写规范', '卷面整洁', '格式正确'] },
  { type: 'note', text: '活动由英语组承办。' }
] }
r = await engine('wechat:check', article)
timing('wechat:check', r)
check(r.ok && r.data.stats && r.data.stats.titleChars > 0, '公众号体检出统计', r.error)
r = await engine('wechat:render', { title: article.title, digest: article.digest, blocks: article.blocks, outputPath: path.join(tmp, '推文.docx'), signature: '临河三中英语组' })
timing('wechat:render', r)
check(r.ok && fs.existsSync(r.data.htmlPath) && fs.existsSync(r.data.outputPath), '公众号出 html + docx', r.error)
if (r.ok) {
  const html = fs.readFileSync(r.data.htmlPath, 'utf8')
  check(!/<link|<script/i.test(html), '公众号 HTML 不带外链（能直接粘进编辑器）')
  check(/临河三中英语组/.test(html), '落款进 HTML')
}

r = await engine('docflow:build', {
  title: '关于做好 2025 年秋季学期期中考试工作的通知',
  from: '临河区教育局',
  docNo: '临教发〔2025〕42 号',
  date: '2025-10-20',
  urgency: '急件',
  deadline: '2025-10-25',
  points: ['各校于 11 月 3 日至 7 日组织期中考试', '考试科目由学校按课程方案确定', '成绩只用于教学诊断，不得排名'],
  opinion: '拟请教务处牵头，各年级组配合，按通知要求落实考场与监考安排。',
  tasks: ['教务处：制定考务方案', '年级组：上报考场与监考名单'],
  reply: {
    to: '临河区教育局',
    blocks: ['贵局《关于做好 2025 年秋季学期期中考试工作的通知》收悉。', '我校已按通知要求成立考务小组，制定考务方案，落实考场与监考安排，考试科目按课程方案确定，成绩只用于教学诊断。', '特此报告。'],
    signature: '临河区第三中学',
    date: '2025年10月22日'
  }
})
timing('docflow:build', r)
check(r.ok && fs.existsSync(r.data.outputPath), '公文办理单落盘', r.error)
check(r.ok && Array.isArray(r.data.sections) && r.data.sections.length === 6, '公文六段齐（登记/要点/拟办/批示/流转/回复）', r.data && r.data.sections)

// 10.6 v0.15 新路由：/profile 与 /launch
r = await callGet('/profile')
check(r.body && r.body.ok === true && Array.isArray(r.body.tasks) && r.body.tasks.length === 10, 'GET /profile 给十个任务', r.body && r.body.tasks && r.body.tasks.length)
check(r.body && r.body.profile && r.body.profile.profile && r.body.profile.profile.school && r.body.profile.profile.school.name === '临河区第三中学', 'GET /profile 带画像', r.body && r.body.profile && r.body.profile.profile && r.body.profile.profile.school)
r = await call('/launch', { task: 'exam' })
check(r.status === 400 && r.body && !r.body.ok && r.body.error && r.body.error.code === 'session-required', '/launch 缺会话号给 400', r.body && r.body.error)
r = await call('/launch', { task: '没有这个任务', sessionId: 'x' })
check(r.status === 400 && r.body && !r.body.ok && r.body.error && r.body.error.code === 'unknown-task', '/launch 未知任务给 400', r.body && r.body.error)

// 11. 错误路径：故意输错，必须给中文话
r = await engine('exam:build', { text: '   ' })
check(!r.ok && (r.error && r.error.message), '空题干报中文错', r.error)
r = await engine('docx:clean', { path: 'D:/nope/没有这个文件.docx' })
check(!r.ok && r.error && /打不开|找不到|没有|不存在/.test(r.error.message + (r.error.hint || '')), '路径不存在给人话', r.error)
r = await engine('lesson:html', { text: '' })
check(!r.ok, '空课件内容被拦', r.error)

restoreProfile()
console.log('')
console.log(fail === 0 ? '全流程自测通过 ✅  ' + pass + ' 项' : '失败 ' + fail + ' 项（通过 ' + pass + '）')
process.exit(fail === 0 ? 0 : 1)