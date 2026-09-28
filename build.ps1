param([string]$OutputName = "ExcelTranslator.exe")
$ErrorActionPreference = 'Stop'
$framework = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
if (!(Test-Path -LiteralPath $compiler)) { throw '需要 Windows .NET Framework 4.x 的 C# 编译器。' }

function Find-Interop([string]$name, [string]$file) {
    $roots = @((Join-Path $env:WINDIR ('assembly/GAC_MSIL/' + $name)), (Join-Path $env:WINDIR ('Microsoft.NET/assembly/GAC_MSIL/' + $name)))
    foreach ($root in $roots) {
        if (Test-Path -LiteralPath $root) {
            $found = Get-ChildItem -LiteralPath $root -Recurse -Filter $file | Select-Object -First 1
            if ($found) { return $found.FullName }
        }
    }
    throw ('找不到 Office 互操作程序集：' + $name + '。请安装桌面版 Excel 的 .NET 可编程性支持。')
}

$excel = Find-Interop 'Microsoft.Office.Interop.Excel' 'Microsoft.Office.Interop.Excel.dll'
$office = Find-Interop 'office' 'office.dll'
$arguments = @('/nologo', '/target:winexe', '/platform:x64', '/optimize+', '/warn:4', '/codepage:65001',
    ('/out:' + (Join-Path $PSScriptRoot $OutputName)),
    ('/win32manifest:' + (Join-Path $PSScriptRoot 'app.manifest')),
    ('/link:' + $excel), ('/link:' + $office))
foreach ($reference in @('System.dll', 'System.Core.dll', 'System.Security.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll', 'System.Web.Extensions.dll', 'Microsoft.CSharp.dll')) {
    $arguments += '/reference:' + $reference
}
foreach ($name in @("UIAutomationClient.dll", "UIAutomationTypes.dll", "WindowsBase.dll")) { $arguments += "/reference:" + (Join-Path $framework ("WPF/" + $name)) }
$arguments += @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.cs' | ForEach-Object { $_.FullName })
& $compiler $arguments
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
Write-Output ("Built " + $OutputName + " (Office interop types embedded).")
