# For more info consult:
# https://github.com/AsBuiltReport/AsBuiltReport.NetApp.ONTAP?tab=readme-ov-file
#
# ===============================
# Hard dependency validation
# ===============================
$ModuleName = 'AsBuiltReport.NetApp.ONTAP'

try {
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Install-Module -Name $ModuleName -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
    }

    Import-Module $ModuleName -ErrorAction Stop
}
catch {
    throw "Required module '$ModuleName' is missing and could not be installed. Script execution aborted."
}

# ===============================
# Script paths
# ===============================
$ScriptDir = $PSScriptRoot
$jsonfile  = Join-Path $ScriptDir 'AsBuiltReport.NetApp.ONTAP.json'

if (-not (Test-Path $jsonfile)) {
    throw "Config file not found: $jsonfile"
}

# ===============================
# User input
# ===============================
$cluster_ip = Read-Host "IP do cluster"
$Creds     = Get-Credential

# ===============================
# Report generation
# ===============================
New-AsBuiltReport `
    -Report NetApp.ONTAP `
    -Target $cluster_ip `
    -Credential $Creds `
    -Format Html,Word `
    -OutputFolderPath $ScriptDir `
    -ReportConfigFilePath $jsonfile `
    -EnableHealthCheck `
    -Verbose
