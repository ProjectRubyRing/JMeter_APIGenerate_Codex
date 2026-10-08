#Requires -Version 5.1
<#
.SYNOPSIS
Generate and run a JMeter 5.6.3 REST API plan from XLSX or UTF-8 CSV.
.EXAMPLE
.\Invoke-JMeterApi.ps1 -InputPath ..\templates\api-template.xlsx -JMeterHome C:\Tools\apache-jmeter-5.6.3
.EXAMPLE
.\Invoke-JMeterApi.ps1 -InputPath ..\templates\api-template.csv -GenerateOnly
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')]
    [string]$InputPath,
    [Parameter(ParameterSetName = 'Run')]
    [string]$JMeterHome,
    [Parameter(ParameterSetName = 'Run')]
    [string]$OutputDirectory,
    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(1, 10000)]
    [int]$Threads = 1,
    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(1, 1000000)]
    [int]$Loops = 1,
    [Parameter(ParameterSetName = 'Run')]
    [ValidateRange(1, 2147483647)]
    [int]$MaxRunSeconds = 300,
    [Parameter(ParameterSetName = 'Run')]
    [switch]$GenerateOnly,
    [Parameter(ParameterSetName = 'Run')]
    [switch]$HtmlReport,
    [Parameter(ParameterSetName = 'Run')]
    [string]$JavaPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Templates')]
    [switch]$CreateTemplates,
    [Parameter(ParameterSetName = 'Templates')]
    [string]$TemplateDirectory = '.\templates',
    [string]$PythonPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NativeDirectoryPath {
    param([string]$DirectoryPath)
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DirectoryPath)
    if ($resolved -eq [System.IO.Path]::GetPathRoot($resolved)) {
        return $resolved
    }
    # PowerShell 5.1 can merge the next native argument when a quoted path ends in a backslash.
    return $resolved.TrimEnd([char[]]@('\', '/'))
}

try {
    $pythonPrefix = @()
    if ($PythonPath) {
        $pythonCommand = (Get-Command -Name $PythonPath -CommandType Application -ErrorAction Stop).Source
    }
    elseif (Get-Command py.exe -CommandType Application -ErrorAction SilentlyContinue) {
        $pythonCommand = (Get-Command py.exe -CommandType Application).Source
        $pythonPrefix = @('-3')
    }
    elseif (Get-Command python.exe -CommandType Application -ErrorAction SilentlyContinue) {
        $pythonCommand = (Get-Command python.exe -CommandType Application).Source
    }
    else {
        throw 'Python 3.9+ is required. Install Python or specify -PythonPath.'
    }

    $entryPoint = Join-Path (Split-Path -Parent $PSScriptRoot) 'jmeter_api.py'
    $pythonArgs = @($entryPoint)
    if ($CreateTemplates) {
        $pythonArgs += @('--create-templates', '--template-dir', (Get-NativeDirectoryPath $TemplateDirectory))
    }
    else {
        $pythonArgs += @('--input', $InputPath, '--threads', "$Threads", '--loops', "$Loops",
            '--max-run-seconds', "$MaxRunSeconds")
        if ($JMeterHome) { $pythonArgs += @('--jmeter-home', (Get-NativeDirectoryPath $JMeterHome)) }
        if ($OutputDirectory) { $pythonArgs += @('--output', (Get-NativeDirectoryPath $OutputDirectory)) }
        if ($JavaPath) { $pythonArgs += @('--java', $JavaPath) }
        if ($GenerateOnly) { $pythonArgs += '--generate-only' }
        if ($HtmlReport) { $pythonArgs += '--html-report' }
    }
    & $pythonCommand @pythonPrefix @pythonArgs
    exit $LASTEXITCODE
}
catch {
    Write-Error -Message $_.Exception.Message -ErrorAction Continue
    exit 2
}
