<#
  dsh-teacher-kit 离线安装脚本（纯 PowerShell，不需要 Node）
  双击 一键安装.cmd / 一键卸载.cmd 即可；也可手动：
    powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -ProfileName desktop
    powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -ProfileRoot D:\tmp\profiles -DryRun
#>
param(
  [string]$ProfileName = '',
  [string]$ProfileRoot = '',
  [switch]$Uninstall,
  [switch]$DryRun,
  [switch]$NoPause
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$NL = [string][char]10

$BEGIN = '# >>> dsh-teacher-kit（一键安装添加，勿手改此行）'
$END   = '# <<< dsh-teacher-kit'
$ENTRY = $BEGIN + $NL + '- insert:' + $NL + '    - id: teacher-kit' + $NL + '      name: "dsh-teacher-kit"' + $NL + $END

$SOURCE_DIRS  = @('lib', 'client', 'engine')
$SOURCE_FILES = @('package.json', 'cordis.patch.yml', 'README.md', 'RELEASE.md', 'LICENSE')

function Say($text) { Write-Host $text }

function Fail($text) {
  Write-Host ''
  Write-Host $text -ForegroundColor Red
  if (-not $NoPause) { Write-Host ''; Read-Host '按回车关闭窗口' | Out-Null }
  exit 1
}

function Read-Text($path) {
  if (-not (Test-Path -LiteralPath $path)) { return '' }
  return [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
}

function Write-Text($path, $text) {
  # 必须不带 BOM：DSH 读 YAML 对 BOM 敏感
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($path, $text, $utf8)
}

function Squeeze-BlankLines($text) {
  $out = $text
  while ($out -match '(\n)\n\n') { $out = $out -replace '(\n)\n\n', $NL + $NL }
  return $out
}

# 旧装法（install-local.mjs / 早期版本）留下的注释标记：一律按「含 dsh-teacher-kit 的注释行」清掉，
# 不做 BEGIN..END 的区间删除 —— 标记可能不成对，区间删除会连用户的其他配置一起吃掉。
function Remove-TeacherKitComments($text) {
  $out = New-Object System.Collections.ArrayList
  foreach ($line in [regex]::Split($text, '\r?\n')) {
    if ($line -match '^\s*#.*dsh-teacher-kit') { continue }
    [void]$out.Add($line)
  }
  return (Squeeze-BlankLines (($out -join $NL).TrimEnd()))
}

# 市场装法写进 patch 的条目没有注释标记，只按标记去重会留下两条 id: teacher-kit
function Drop-EntriesById($text) {
  $lines = [regex]::Split($text, '\r?\n')
  $head = New-Object System.Collections.ArrayList
  $blocks = New-Object System.Collections.ArrayList
  $current = $null
  foreach ($line in $lines) {
    if ($line -match '^-\s' -or $line.Trim() -eq '-') {
      $current = New-Object System.Collections.ArrayList
      [void]$current.Add($line)
      [void]$blocks.Add($current)
    } elseif ($null -ne $current) {
      [void]$current.Add($line)
    } else {
      [void]$head.Add($line)
    }
  }
  $kept = New-Object System.Collections.ArrayList
  $removed = 0
  foreach ($block in $blocks) {
    $isOurs = $false
    foreach ($line in $block) {
      if ($line -match '^\s*-\s*id:\s*["'']?teacher-kit["'']?\s*$') { $isOurs = $true }
    }
    if ($isOurs) { $removed++ } else { [void]$kept.Add($block) }
  }
  if ($removed -eq 0) { return @{ Text = $text; Removed = 0 } }
  $all = New-Object System.Collections.ArrayList
  foreach ($line in $head) { [void]$all.Add($line) }
  foreach ($block in $kept) { foreach ($line in $block) { [void]$all.Add($line) } }
  return @{ Text = (Squeeze-BlankLines (($all -join $NL).TrimEnd())); Removed = $removed }
}

function Get-ItemLines($text) {
  $out = New-Object System.Collections.ArrayList
  foreach ($line in [regex]::Split($text, '\r?\n')) {
    $t = $line.Trim()
    if ($t -eq '') { continue }
    if ($t.StartsWith('#')) { continue }
    if ($t -eq '[]') { continue }
    [void]$out.Add($t)
  }
  return $out
}

function Test-IsEmptyList($text) { return ((Get-ItemLines $text).Count -eq 0) }

function Drop-EmptyFlowList($text) {
  $out = New-Object System.Collections.ArrayList
  foreach ($line in [regex]::Split($text, '\r?\n')) {
    if ($line.Trim() -eq '[]') { continue }
    [void]$out.Add($line)
  }
  return ($out -join $NL).TrimEnd()
}

function Get-CommentLines($text) {
  $out = New-Object System.Collections.ArrayList
  foreach ($line in [regex]::Split($text, '\r?\n')) {
    if ($line.Trim().StartsWith('#')) { [void]$out.Add($line) }
  }
  return ($out -join $NL)
}

function Prepare-ForAppend($text) {
  if (Test-IsEmptyList $text) { return Drop-EmptyFlowList $text }
  return $text.TrimEnd()
}

try {
  $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
  $source = Join-Path $scriptDir 'plugin'

  if (-not (Test-Path -LiteralPath (Join-Path $source 'package.json'))) {
    Fail ('安装包不完整：找不到 ' + (Join-Path $source 'package.json') + $NL + '请先把整个 zip 解压，再在解压后的文件夹里双击运行。')
  }

  $root = $ProfileRoot
  if ($root -eq '') { $root = Join-Path $env:USERPROFILE '.dsh\profiles' }
  if (-not (Test-Path -LiteralPath $root)) {
    Fail ('没找到 DSH 的配置目录：' + $root + $NL + '请先启动一次 DSH（桌面版或 dsh web），让它把目录建出来，再运行本安装包。')
  }

  $targets = New-Object System.Collections.ArrayList
  if ($ProfileName -ne '') {
    $one = Join-Path $root $ProfileName
    if (-not (Test-Path -LiteralPath $one)) {
      $names = (Get-ChildItem -LiteralPath $root -Directory | Select-Object -ExpandProperty Name) -join '、'
      Fail ('找不到这个 profile：' + $one + $NL + '已有的：' + $names)
    }
    [void]$targets.Add($one)
  } else {
    foreach ($name in @('desktop', 'web', 'headless')) {
      $candidate = Join-Path $root $name
      if ((Test-Path -LiteralPath $candidate) -and (Test-Path -LiteralPath (Join-Path $candidate 'cordis.patch.yml'))) {
        [void]$targets.Add($candidate)
      }
    }
    if ($targets.Count -eq 0) {
      foreach ($dir in (Get-ChildItem -LiteralPath $root -Directory)) {
        if (Test-Path -LiteralPath (Join-Path $dir.FullName 'cordis.patch.yml')) { [void]$targets.Add($dir.FullName) }
      }
    }
  }
  if ($targets.Count -eq 0) { Fail ('在 ' + $root + ' 下没找到任何 DSH profile。请先启动一次 DSH 再试。') }

  Say ''
  Say '教师助手 · 离线安装'
  Say '-------------------------------'
  Say ('安装包：' + $scriptDir)
  Say ('profiles：' + $root)
  if ($DryRun) { Say '（演练模式，不写任何文件）' }
  Say ''

  foreach ($profileDir in $targets) {
    $name = Split-Path -Leaf $profileDir
    $target = Join-Path $profileDir 'node_modules\dsh-teacher-kit'
    $patchFile = Join-Path $profileDir 'cordis.patch.yml'

    if ($Uninstall) {
      Say ('[' + $name + '] 卸载')
      if ($DryRun) {
        Say ('  会删除 ' + $target)
      } else {
        if (Test-Path -LiteralPath $target) {
          Remove-Item -LiteralPath $target -Recurse -Force
          Say '  已删除插件文件'
        }
        if (Test-Path -LiteralPath $patchFile) {
          $patch = Read-Text $patchFile
          $dedup = Drop-EntriesById (Remove-TeacherKitComments $patch)
          if ($dedup.Removed -gt 0 -or $patch -match 'dsh-teacher-kit') {
            $body = $dedup.Text
            if (Test-IsEmptyList $body) {
              $comments = Get-CommentLines $body
              if ($comments -ne '') { $body = $comments + $NL + $NL + '[]' } else { $body = '[]' }
            }
            Write-Text $patchFile $body
            Say '  已从 cordis.patch.yml 摘掉启用记录'
          } else {
            Say '  cordis.patch.yml 里没有本插件的记录，跳过'
          }
        }
      }
      Say ''
      continue
    }

    Say ('[' + $name + '] 安装')
    if ($DryRun) {
      Say ('  会写入 ' + $target)
    } else {
      $parent = Join-Path $profileDir 'node_modules'
      if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
      if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
      New-Item -ItemType Directory -Path $target -Force | Out-Null
      foreach ($dirName in $SOURCE_DIRS) {
        $from = Join-Path $source $dirName
        if (Test-Path -LiteralPath $from) { Copy-Item -LiteralPath $from -Destination (Join-Path $target $dirName) -Recurse -Force }
      }
      foreach ($fileName in $SOURCE_FILES) {
        $from = Join-Path $source $fileName
        if (Test-Path -LiteralPath $from) { Copy-Item -LiteralPath $from -Destination (Join-Path $target $fileName) -Force }
      }
      Say '  插件文件已就位'

      $patch = Read-Text $patchFile
      $dedup = Drop-EntriesById (Remove-TeacherKitComments $patch)
      if ($dedup.Removed -gt 0) { Say ('  清掉了 ' + $dedup.Removed + ' 条同类旧记录（插件市场装法留下的）') }
      $body = Prepare-ForAppend $dedup.Text
      if ($body -ne '') { $newPatch = $body + $NL + $NL + $ENTRY + $NL } else { $newPatch = $ENTRY + $NL }

      # 写前自检：有列表项时绝不能再留空的 flow 序列 []，否则 YAML 解析失败、DSH 起不来
      $hasItems = (Get-ItemLines $newPatch).Count -gt 0
      $hasEmptyFlow = $false
      foreach ($line in [regex]::Split($newPatch, '\r?\n')) { if ($line.Trim() -eq '[]') { $hasEmptyFlow = $true } }
      if ($hasItems -and $hasEmptyFlow) {
        Fail '内部自检失败：生成的 cordis.patch.yml 同时含空数组 [] 和列表项，写出去会让 DSH 起不来。已中止，没有改动你的配置。'
      }
      Write-Text $patchFile $newPatch
      Say '  已在 cordis.patch.yml 里启用'
    }
    Say ''
  }

  if ($Uninstall) {
    Say '卸载完成。重启 DSH 生效。'
  } else {
    Say '装好了。'
    Say ''
    Say '  现在请【完全退出 DSH 再重新打开】（关掉所有窗口）。'
    Say '  打开任意对话，输入框下面应该出现「教师助手」按钮。'
    Say ''
    Say '  如果没出现：设置 → 插件，确认列表里有 teacher-kit 且被勾选。'
  }
  if ((-not $DryRun) -and (-not $NoPause)) { Write-Host ''; Read-Host '按回车关闭窗口' | Out-Null }
} catch {
  Fail ('出错了：' + $_.Exception.Message + $NL + $_.ScriptStackTrace)
}