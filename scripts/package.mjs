/**
 * 打发布包：dist/dsh-teacher-kit-<version>.zip
 *
 * 只放运行所需的文件，不带 tests / reviews / .verify / smoke.log / node_modules。
 * Windows 自带 bsdtar（C:\Windows\System32\tar.exe），用 `tar -a -cf` 出 zip，不引额外依赖；
 * 其它平台退回系统 tar 的 --format=zip。
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
const distDir = join(root, 'dist')
const zip = join(distDir, 'dsh-teacher-kit-' + pkg.version + '.zip')

// 打包前先把三处版本与产物新鲜度对齐：以前出现过「包是 0.15.0、引擎自报 0.14.0、
// client.js 比 factory.js 旧 11 分钟」而脚本照样打印 ✅ 的事故。
const assertConsistent = () => {
  const problems = []
  const hostSrc = readFileSync(join(root, 'lib', 'index.js'), 'utf8')
  const hostVersion = (hostSrc.match(/export const VERSION = '([^']+)'/) || [])[1]
  if (hostVersion !== pkg.version) {
    problems.push('lib/index.js 的 VERSION=' + hostVersion + '，package.json 是 ' + pkg.version)
  }
  const pyInit = readFileSync(join(root, 'engine', 'teacher_kit', '__init__.py'), 'utf8')
  const pyVersion = (pyInit.match(/__version__ = "([^"]+)"/) || [])[1]
  if (pyVersion !== pkg.version) {
    problems.push('engine/teacher_kit/__init__.py 的 __version__=' + pyVersion + '，package.json 是 ' + pkg.version)
  }
  const srcBundle = join(root, 'src', 'client', 'factory.js')
  const outBundle = join(root, 'client', 'client.js')
  if (!existsSync(outBundle)) {
    problems.push('client/client.js 不存在：先跑 node scripts/build-client.mjs')
  } else if (existsSync(srcBundle) && statSync(outBundle).mtimeMs < statSync(srcBundle).mtimeMs) {
    problems.push('client/client.js 比 src/client/factory.js 旧：先跑 node scripts/build-client.mjs')
  }
  if (problems.length) {
    problems.forEach((line) => console.error('✗ ' + line))
    console.error('打包已中止（没生成新的 zip）。')
    process.exit(1)
  }
}
assertConsistent()

const FILES = [
  'lib',
  'client',
  'engine',
  'cordis.patch.yml',
  'package.json',
  'README.md',
  'RELEASE.md',
  'scripts/build-client.mjs',
  'scripts/install-local.mjs',
  'scripts/package.mjs'
]

const dropCaches = (dir) => {
  if (!existsSync(dir)) return 0
  let removed = 0
  for (const name of readdirSync(dir)) {
    const full = join(dir, name)
    if (!statSync(full).isDirectory()) continue
    if (name === '__pycache__') {
      rmSync(full, { recursive: true, force: true })
      removed += 1
    } else {
      removed += dropCaches(full)
    }
  }
  return removed
}
const caches = dropCaches(join(root, 'engine'))
if (caches) console.log('清掉 __pycache__ ' + caches + ' 个')

mkdirSync(distDir, { recursive: true })
rmSync(zip, { force: true })

const winTar = 'C:\\Windows\\System32\\tar.exe'
const tarBin = existsSync(winTar) ? winTar : 'tar'
const EXCLUDES = ['--exclude=engine/tests', '--exclude=__pycache__', '--exclude=*.pyc']
const tarArgs = existsSync(winTar)
  ? ['-a', '-c', '-f', zip, ...EXCLUDES, ...FILES]
  : ['--format=zip', '-c', '-f', zip, ...EXCLUDES, ...FILES]

const tar = spawnSync(tarBin, tarArgs, { cwd: root, encoding: 'utf8' })
if (tar.status !== 0) {
  console.error('打包失败：' + (tar.stderr || tar.error?.message || String(tar.status)))
  process.exit(1)
}

const listed = spawnSync(tarBin, ['-tf', zip], { encoding: 'utf8' })
const entries = (listed.stdout || '').split(/\r?\n/).map((line) => line.trim().replace(/\\/g, '/')).filter(Boolean)
const tops = Array.from(new Set(entries.map((line) => line.split('/')[0]))).sort()
console.log('已打包：' + zip)
console.log('文件大小：' + Math.round(statSync(zip).size / 1024) + 'KB · 条目 ' + entries.length + ' 个')
console.log('顶层：' + tops.join(' , '))

const leaks = entries.filter((line) => /(__pycache__|\\.verify|reviews\/|(^|\/)test\/|smoke\\.log|node_modules)/.test(line))
if (leaks.length) {
  console.error('包里有不该发的文件：' + leaks.slice(0, 10).join(', '))
  process.exit(1)
}
const missing = ['lib/index.js', 'client/client.js', 'engine/teacher_kit/commands.py', 'cordis.patch.yml', 'package.json', 'README.md']
  .filter((need) => !entries.some((line) => line === need || line.endsWith('/' + need)))
if (missing.length) {
  console.error('包里缺文件：' + missing.join(', '))
  process.exit(1)
}
console.log('检查：无 __pycache__ / tests / reviews / node_modules，关键文件齐全 ✅')
