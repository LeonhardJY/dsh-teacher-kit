/**
 * 把教师助手装进某个 DSH profile（本地安装，不经过市场）。
 *
 *   node scripts/install-local.mjs --profile desktop
 *   node scripts/install-local.mjs --profile web --dry-run
 *   node scripts/install-local.mjs --profile desktop --uninstall
 *
 * 做两件事：把包拷进 <profile>/node_modules/dsh-teacher-kit；在 <profile>/cordis.patch.yml
 * 里加一段带标记的 insert 条目（卸载时按标记删掉，不碰老师自己的配置）。
 */
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const BEGIN = '# >>> dsh-teacher-kit（install-local.mjs 添加，勿手改此行）'
const END = '# <<< dsh-teacher-kit'

const args = process.argv.slice(2)
const flag = (name, fallback) => {
  const index = args.indexOf(name)
  if (index === -1) return fallback
  const value = args[index + 1]
  return value && !value.startsWith('--') ? value : true
}
const profile = String(flag('--profile', 'desktop') === true ? 'desktop' : flag('--profile', 'desktop'))
const dryRun = args.includes('--dry-run')
const uninstall = args.includes('--uninstall')

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const profileDir = join(homedir(), '.dsh', 'profiles', profile)
const target = join(profileDir, 'node_modules', 'dsh-teacher-kit')
const patchFile = join(profileDir, 'cordis.patch.yml')

if (!existsSync(profileDir)) {
  console.error('找不到这个 profile：' + profileDir)
  console.error('先启动一次 DSH（桌面版或 dsh web），让它把 profile 建出来；可用 --profile web / desktop / headless。')
  process.exit(1)
}

const ENTRY = BEGIN + '\n- insert:\n    - id: teacher-kit\n      name: ' + JSON.stringify('dsh-teacher-kit') + '\n' + END + '\n'

// 这个文件是「顶层 YAML 数组」。列表为空时写成 `[]`（flow 序列），
// 而 flow 序列后面**不能再接 block 项**，否则 YAML 直接解析失败、DSH 起不来。
// 所以插入前必须先判断「列表是不是空的」，空的要把 `[]` 那一行摘掉再写 block 项。
const stripEntry = (text) => {
  const start = text.indexOf(BEGIN)
  if (start === -1) return { text, removed: false }
  const end = text.indexOf(END)
  const after = end === -1 ? '' : text.slice(end + END.length)
  return { text: (text.slice(0, start) + after).replace(/\n{3,}/g, '\n\n'), removed: true }
}
// 市场装法写进 patch 的条目没有我们的注释标记（形状是 `- insert: … id: teacher-kit …`），
// 只按标记去重会留下两条 id: teacher-kit（实测写出过）。这里再按 id 扫一遍顶层列表项。
const dropEntriesById = (text) => {
  const lines = text.split(/\r?\n/)
  const head = []
  const blocks = []
  let current = null
  for (const line of lines) {
    if (/^-\s/.test(line) || line.trim() === '-') {
      current = [line]
      blocks.push(current)
    } else if (current) {
      current.push(line)
    } else {
      head.push(line)
    }
  }
  const kept = blocks.filter((block) => !block.some((line) => /^\s*-\s*id:\s*["']?teacher-kit["']?/.test(line)))
  if (kept.length === blocks.length) return { text, removed: 0 }
  const joined = head.concat(...kept).join('\n').replace(/\n{3,}/g, '\n\n').replace(/\s+$/, '')
  return { text: joined, removed: blocks.length - kept.length }
}
const itemLines = (text) =>
  text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line !== '' && !line.startsWith('#') && line !== '[]')
const isEmptyList = (text) => itemLines(text).length === 0
const dropEmptyFlowList = (text) =>
  text
    .split(/\r?\n/)
    .filter((line) => line.trim() !== '[]')
    .join('\n')
    .replace(/\s+$/, '')
const commentLines = (text) =>
  text
    .split(/\r?\n/)
    .filter((line) => line.trim().startsWith('#'))
    .join('\n')
// 把「文件现状」整理成「能安静地追加 block 项」：空列表去掉 []，其余原样。
const prepareForAppend = (text) => (isEmptyList(text) ? dropEmptyFlowList(text) : text.replace(/\s+$/, ''))

if (uninstall) {
  if (dryRun) {
    console.log('会删除：' + target)
    console.log('会从 ' + patchFile + ' 里删掉启用记录')
    process.exit(0)
  }
  if (existsSync(target)) rmSync(target, { recursive: true, force: true })
  if (existsSync(patchFile)) {
    const { text, removed } = stripEntry(readFileSync(patchFile, 'utf8'))
    if (removed) {
      // 摘掉我们的段以后列表空了，要写回 []，否则文件是「只有注释、没有数组」→ 同样解析失败。
      const restored = isEmptyList(text) ? commentLines(text) + (commentLines(text) ? '\n\n' : '') + '[]\n' : text.replace(/\s+$/, '') + '\n'
      writeFileSync(patchFile, restored, 'utf8')
    }
  }
  console.log('已卸载。重启 DSH 生效。')
  process.exit(0)
}

const SOURCE_DIRS = ['lib', 'client', 'engine', 'scripts']
const SOURCE_FILES = ['package.json', 'cordis.patch.yml', 'README.md', 'RELEASE.md']

console.log('来源：' + root)
console.log('目标：' + target)
if (dryRun) {
  console.log('（--dry-run，只打印，不动文件）')
  process.exit(0)
}

mkdirSync(dirname(target), { recursive: true })
if (existsSync(target)) rmSync(target, { recursive: true, force: true })
for (const name of SOURCE_DIRS) {
  const from = join(root, name)
  if (existsSync(from)) cpSync(from, join(target, name), { recursive: true })
}
for (const name of SOURCE_FILES) {
  const from = join(root, name)
  if (existsSync(from)) cpSync(from, join(target, name))
}

let patch = existsSync(patchFile) ? readFileSync(patchFile, 'utf8') : ''
if (patch.includes(BEGIN)) {
  // 已经装过：先摘掉旧段再写回去，能顺手把上次写坏的 [] 修掉。
  patch = stripEntry(patch).text
}
const dedup = dropEntriesById(patch)
if (dedup.removed) {
  console.log('清掉了 ' + dedup.removed + ' 条同类旧启用记录（市场装法留下的，没标记）')
  patch = dedup.text
}
const body = prepareForAppend(patch)
patch = (body ? body + '\n\n' : '') + ENTRY
// 写前自检：block 项存在时绝不能再留着空的 flow 序列 []（这条曾让 DSH 直接起不来）。
const broken = patch.split(/\r?\n/).some((line) => line.trim() === '[]')
if (broken && itemLines(patch).length > 0) {
  console.error('内部错误：生成的 patch 里同时有空数组 [] 和列表项，写出去会让 DSH 起不来。已中止，没动你的配置。')
  process.exit(1)
}
writeFileSync(patchFile, patch, 'utf8')
console.log('已写入启用记录：' + patchFile)

console.log('')
console.log('装好了。现在**重启 DSH**（关窗口再打开），输入框右下角（模型选择器左边）应出现「教师助手」。')
console.log('回滚运行：node scripts/install-local.mjs --profile ' + profile + ' --uninstall')
