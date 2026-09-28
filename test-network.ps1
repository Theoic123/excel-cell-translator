# Run with Windows PowerShell 5.1 (powershell.exe), which uses .NET Framework.
# Sends only fixed test phrases to the configured provider. Does not access Excel.
$ErrorActionPreference = 'Stop'
$assembly = [System.Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot 'ExcelTranslator.exe'))
$service = $assembly.GetType('ExcelCellTranslator.TranslationService', $true)
$translate = $service.GetMethod('TranslateAsync', [System.Reflection.BindingFlags]'Public, Static')
$report = @()
foreach ($source in @('Hello world', 'Region_Score_Limit', 'model_score', ([string][char]0x4F60 + [char]0x597D + [char]0x4E16 + [char]0x754C))) {
    $task = $translate.Invoke($null, @($source, [System.Threading.CancellationToken]::None))
    $task.GetAwaiter().GetResult() | Out-Null
    $result = $task.GetType().GetProperty('Result').GetValue($task, $null)
    if ([string]::IsNullOrWhiteSpace($result.Text)) { throw 'Empty translation returned.' }
    if ($source -match '_' -and $result.Text -notmatch '[\u3400-\u9fff]') { throw 'Identifier did not produce a Chinese translation.' }
    $report += ($source + ' => ' + $result.Text + ' [' + $result.Direction + ']')
}
$report | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'network-test.txt') -Encoding UTF8
$report
