<#
.SYNOPSIS
    Cria um External Virtual Switch no Hyper-V 2025 seguindo as boas praticas Microsoft.

.DESCRIPTION
    Script interativo que guia o administrador na criacao de um External vSwitch com suporte a:

        - SET  (Switch Embedded Teaming)  — recomendado pela Microsoft para redundancia de NICs
        - SR-IOV                          — offload de I/O para maior desempenho por VM
        - AllowManagementOS               — controlo do trafego de gestao do host
        - VMQ  (Virtual Machine Queues)   — distribuicao de carga de rede por fila
        - QoS / Bandwidth Management      — garantia de largura de banda minima por fluxo
        - Logging                         — registo de todas as accoes em ficheiro de log

    Suporte a -WhatIf para simulacao sem alteracoes e -Verbose para diagnostico detalhado.

    Requisitos:
        - Windows Server 2016 ou superior (recomendado 2025)
        - Hyper-V instalado e modulo PowerShell disponivel
        - Execucao como Administrador
        - PowerShell 5.1 ou superior

.PARAMETER LogPath
    Caminho para o ficheiro de log. Por omissao: C:\Logs\HyperV_Switch_<timestamp>.log

.EXAMPLE
    .\HyperV_Create_Switch.ps1
    Execucao interativa completa com assistente passo a passo.

.EXAMPLE
    .\HyperV_Create_Switch.ps1 -Verbose
    Execucao interativa com saida de diagnostico detalhada.

.EXAMPLE
    .\HyperV_Create_Switch.ps1 -WhatIf
    Simula toda a configuracao sem aplicar qualquer alteracao.

.EXAMPLE
    .\HyperV_Create_Switch.ps1 -LogPath 'D:\Logs\switch.log'
    Execucao com destino de log personalizado.

.NOTES
    Autor        : Infraestrutura
    Versao       : 2.0.0
    Data         : 2025-03-18
    Referencia   : https://learn.microsoft.com/en-us/windows-server/virtualization/hyper-v/plan/plan-hyper-v-networking-in-windows-server
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules Hyper-V

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter()]
    [string] $LogPath = ("C:\Logs\HyperV_Switch_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Logging ---

function Initialize-Log {
    <#
    .SYNOPSIS
        Cria o directorio e ficheiro de log se nao existirem.
    .PARAMETER Path
        Caminho completo para o ficheiro de log.
    #>
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $logDir = Split-Path -Path $Path -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    $header = @"
================================================================================
  Hyper-V External Switch - Log de Criacao
  Data/Hora : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
  Servidor  : $env:COMPUTERNAME
  Usuario   : $env:USERDOMAIN\$env:USERNAME
================================================================================
"@
    Add-Content -Path $Path -Value $header
}

function Write-Log {
    <#
    .SYNOPSIS
        Escreve uma mensagem no log e opcionalmente no ecra.
    .PARAMETER Message
        Mensagem a registar.
    .PARAMETER Level
        Nivel da mensagem: INFO, WARN, ERROR.
    .PARAMETER NoConsole
        Se definido, nao escreve no ecra.
    #>
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [Parameter()]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string] $Level = 'INFO',

        [Parameter()]
        [switch] $NoConsole
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine   = '[{0}] [{1}] {2}' -f $timestamp, $Level, $Message

    Add-Content -Path $LogPath -Value $logLine -ErrorAction SilentlyContinue

    if (-not $NoConsole) {
        $color = switch ($Level) {
            'WARN'  { 'Yellow' }
            'ERROR' { 'Red'    }
            default { 'Gray'   }
        }
        Write-Verbose $logLine
    }
}

#endregion

#region --- Interface de utilizador ---

function Show-Header {
    <#
    .SYNOPSIS
        Apresenta o cabecalho principal da aplicacao.
    #>
    Clear-Host
    Write-Host ''
    Write-Host '  ================================================================' -ForegroundColor Cyan
    Write-Host '   Hyper-V 2025  -  Criacao de External vSwitch                 ' -ForegroundColor Cyan
    Write-Host '   Microsoft Best Practices  |  SET  |  SR-IOV  |  VMQ  |  QoS  ' -ForegroundColor DarkCyan
    Write-Host '  ================================================================' -ForegroundColor Cyan
    Write-Host ''
}

function Show-Step {
    <#
    .SYNOPSIS
        Apresenta o titulo de um passo no assistente.
    .PARAMETER Number
        Numero do passo.
    .PARAMETER Title
        Titulo do passo.
    #>
    param(
        [Parameter(Mandatory)] [int]    $Number,
        [Parameter(Mandatory)] [string] $Title
    )
    Write-Host ''
    Write-Host ('  [ Passo {0} ]  {1}' -f $Number, $Title) -ForegroundColor Yellow
    Write-Host ('  ' + ('-' * 60))     -ForegroundColor DarkGray
}

function Show-Info {
    <#
    .SYNOPSIS
        Apresenta uma caixa de informacao contextual ao utilizador.
    .PARAMETER Lines
        Array de linhas de texto a apresentar.
    #>
    param(
        [Parameter(Mandatory)]
        [string[]] $Lines
    )
    Write-Host ''
    Write-Host '  [i] ' -NoNewline -ForegroundColor Cyan
    Write-Host $Lines[0] -ForegroundColor Gray
    foreach ($line in ($Lines | Select-Object -Skip 1)) {
        Write-Host ('      ' + $line) -ForegroundColor DarkGray
    }
    Write-Host ''
}

#endregion

#region --- Deteccao de hardware e capacidades ---

function Get-PhysicalNetAdapter {
    <#
    .SYNOPSIS
        Devolve os adaptadores de rede fisicos com informacao de capacidades.
    .OUTPUTS
        [PSCustomObject[]] Adaptadores com propriedades enriquecidas.
    #>
    [OutputType([PSCustomObject[]])]
    param()

    $excludePattern = 'Hyper-V|Loopback|Virtual|WAN Miniport|Bluetooth|TAP|Tunnel|6to4|ISATAP'

    $rawAdapters = Get-NetAdapter |
        Where-Object { $_.InterfaceDescription -notmatch $excludePattern } |
        Sort-Object -Property Name

    $enriched = foreach ($nic in $rawAdapters) {
        # Verificar suporte a VMQ
        $vmqSupport = $false
        try {
            $vmqInfo    = Get-NetAdapterVmq -Name $nic.Name -ErrorAction SilentlyContinue
            $vmqSupport = ($null -ne $vmqInfo -and $vmqInfo.Enabled)
        } catch { }

        # Verificar suporte a SR-IOV
        $sriovSupport = $false
        try {
            $sriovInfo    = Get-NetAdapterSriov -Name $nic.Name -ErrorAction SilentlyContinue
            $sriovSupport = ($null -ne $sriovInfo -and $sriovInfo.SriovSupport -eq 'Supported')
        } catch { }

        # Velocidade legivel
        $speedText = if ($nic.LinkSpeed) {
            switch -Regex ($nic.LinkSpeed) {
                '(\d+) Gbps' { $matches[0] }
                '(\d+) Mbps' { $matches[0] }
                default      { $nic.LinkSpeed }
            }
        } else { 'N/D' }

        [PSCustomObject]@{
            Name                  = $nic.Name
            InterfaceDescription  = $nic.InterfaceDescription
            Status                = $nic.Status
            LinkSpeed             = $speedText
            MacAddress            = $nic.MacAddress
            VMQSupport            = $vmqSupport
            SRIOVSupport          = $sriovSupport
            IfIndex               = $nic.ifIndex
        }
    }

    # Preferir adaptadores ativos; fallback para todos
    $active = $enriched | Where-Object { $_.Status -eq 'Up' }
    if (($active | Measure-Object).Count -gt 0) {
        return $active
    }

    Write-Warning 'Nenhum adaptador ativo encontrado. A mostrar todos os adaptadores.'
    return $enriched
}

#endregion

#region --- Apresentacao de adaptadores ---

function Show-AdapterTable {
    <#
    .SYNOPSIS
        Apresenta os adaptadores disponiveis em tabela numerada com capacidades.
    .PARAMETER Adapters
        Array de PSCustomObject com informacao dos adaptadores.
    #>
    param(
        [Parameter(Mandatory)]
        [PSCustomObject[]] $Adapters
    )

    $sep = '  ' + ('-' * 80)

    Write-Host '  Interfaces de rede disponiveis:' -ForegroundColor White
    Write-Host $sep -ForegroundColor DarkGray
    Write-Host ('  {0,-5} {1,-20} {2,-10} {3,-10} {4,-6} {5,-6} {6}' -f `
        'No.', 'Nome', 'Estado', 'Velocidade', 'VMQ', 'SR-IOV', 'Descricao') -ForegroundColor Yellow
    Write-Host $sep -ForegroundColor DarkGray

    for ($i = 0; $i -lt $Adapters.Count; $i++) {
        $a           = $Adapters[$i]
        $statusText  = if ($a.Status -eq 'Up') { 'Ligado' } else { 'Desligado' }
        $statusColor = if ($a.Status -eq 'Up') { 'Green'  } else { 'Red'       }
        $vmq         = if ($a.VMQSupport)   { 'Sim' } else { 'Nao' }
        $sriov       = if ($a.SRIOVSupport) { 'Sim' } else { 'Nao' }
        $vmqColor    = if ($a.VMQSupport)   { 'Green' } else { 'DarkGray' }
        $sriovColor  = if ($a.SRIOVSupport) { 'Green' } else { 'DarkGray' }

        Write-Host ('  [{0}]  ' -f ($i + 1))           -NoNewline -ForegroundColor Cyan
        Write-Host ('{0,-20} '  -f $a.Name)             -NoNewline -ForegroundColor White
        Write-Host ('{0,-10} '  -f $statusText)          -NoNewline -ForegroundColor $statusColor
        Write-Host ('{0,-10} '  -f $a.LinkSpeed)         -NoNewline -ForegroundColor White
        Write-Host ('{0,-6} '   -f $vmq)                 -NoNewline -ForegroundColor $vmqColor
        Write-Host ('{0,-6} '   -f $sriov)               -NoNewline -ForegroundColor $sriovColor
        Write-Host ('{0}'        -f $a.InterfaceDescription)         -ForegroundColor DarkGray
    }

    Write-Host $sep -ForegroundColor DarkGray
}

#endregion

#region --- Selecao de adaptadores (SET suporta multiplos) ---

function Read-AdapterSelection {
    <#
    .SYNOPSIS
        Pede ao utilizador que selecione um ou mais adaptadores para o switch.
        Seleccionar dois ou mais activa automaticamente o SET (Switch Embedded Teaming).
    .PARAMETER Adapters
        Array de adaptadores disponiveis.
    .OUTPUTS
        [PSCustomObject[]] Adaptadores selecionados.
    #>
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject[]] $Adapters
    )

    Show-Info -Lines @(
        'Para SET (Switch Embedded Teaming) seleciona dois ou mais adaptadores.',
        'Exemplo com um NIC  : 1',
        'Exemplo com SET     : 1,2   ou   1 2'
    )

    $selected = $null

    do {
        $userInput = Read-Host -Prompt ('  Numero(s) do(s) adaptador(es) [1-{0}]' -f $Adapters.Count)

        # Aceita separador virgula ou espaco
        $parts = $userInput -split '[,\s]+' | Where-Object { $_ -match '^\d+$' }

        if ($parts.Count -eq 0) {
            Write-Warning 'Entrada invalida. Introduz um ou mais numeros separados por virgula.'
            continue
        }

        $indices = $parts | ForEach-Object { [int]$_ - 1 }
        $outOfRange = $indices | Where-Object { $_ -lt 0 -or $_ -ge $Adapters.Count }

        if (($outOfRange | Measure-Object).Count -gt 0) {
            Write-Warning ('Numero(s) fora do intervalo. Usa valores entre 1 e {0}.' -f $Adapters.Count)
            continue
        }

        $duplicates = $indices | Group-Object | Where-Object { $_.Count -gt 1 }
        if (($duplicates | Measure-Object).Count -gt 0) {
            Write-Warning 'Nao podes selecionar o mesmo adaptador mais do que uma vez.'
            continue
        }

        $selected = $indices | ForEach-Object { $Adapters[$_] }

    } while ($null -eq $selected)

    return $selected
}

#endregion

#region --- Configuracao de opcoes do switch ---

function Read-SwitchName {
    <#
    .SYNOPSIS
        Pede o nome do switch com sugestao automatica baseada nos NICs e modo de teaming.
    .PARAMETER AdapterNames
        Nomes dos adaptadores selecionados.
    .PARAMETER UseSet
        Se verdadeiro, o nome sugere SET.
    .OUTPUTS
        [string] Nome do switch validado.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] [string[]] $AdapterNames,
        [Parameter(Mandatory)] [bool]     $UseSet
    )

    $prefix  = if ($UseSet) { 'vSwitch-SET' } else { 'vSwitch' }
    $suffix  = $AdapterNames[0]
    $default = '{0}-{1}' -f $prefix, $suffix

    Write-Host ('  Nome sugerido : ') -NoNewline -ForegroundColor Gray
    Write-Host $default               -ForegroundColor Cyan
    Write-Host ''

    $nameInput = Read-Host -Prompt '  Nome do switch [Enter para aceitar o sugerido]'

    if ([string]::IsNullOrWhiteSpace($nameInput)) {
        return $default
    }

    return $nameInput.Trim()
}

function Read-AllowManagementOS {
    <#
    .SYNOPSIS
        Pergunta se o SO host deve partilhar a interface de rede com as VMs.
    .OUTPUTS
        [bool] Verdadeiro para AllowManagementOS = $true.
    #>
    [OutputType([bool])]
    param()

    Show-Info -Lines @(
        'AllowManagementOS: define se o host tem acesso a rede pelo mesmo switch.',
        'Sim (recomendado para NIC unico) : o host partilha a interface com as VMs.',
        'Nao (recomendado com SET/2+ NICs): dedicar o switch exclusivamente as VMs',
        '     e usar uma NIC separada para gestao do host.'
    )

    $resp = Read-Host -Prompt '  Permitir acesso do SO host ao switch? [S/N]'
    return $resp -match '^[Ss]$'
}

function Read-EnableSriov {
    <#
    .SYNOPSIS
        Pergunta se deve ser activado SR-IOV no switch (requer hardware compativel).
    .PARAMETER AnyNicSupports
        True se pelo menos um NIC selecionado suporta SR-IOV.
    .OUTPUTS
        [bool] Verdadeiro para EnableIov = $true.
    #>
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)] [bool] $AnyNicSupports
    )

    if (-not $AnyNicSupports) {
        Write-Host ''
        Write-Host '  [!] Nenhum dos NICs selecionados reporta suporte a SR-IOV. Opcao ignorada.' -ForegroundColor DarkYellow
        Write-Log -Message 'SR-IOV ignorado: nenhum NIC selecionado reporta suporte.' -Level 'WARN' -NoConsole
        return $false
    }

    Show-Info -Lines @(
        'SR-IOV (Single Root I/O Virtualization): permite que as VMs comuniquem',
        'diretamente com o hardware de rede, reduzindo latencia e carga no host.',
        'Requer: NIC compativel, BIOS/UEFI com SR-IOV activo, VM Generation 2.',
        'Nota: nao compativel com algumas extensoes de switch de terceiros.'
    )

    $resp = Read-Host -Prompt '  Activar SR-IOV? [S/N]'
    return $resp -match '^[Ss]$'
}

function Read-BandwidthMode {
    <#
    .SYNOPSIS
        Pergunta o modo de gestao de largura de banda do switch.
    .OUTPUTS
        [string] Modo: None, Default ou Weight.
    #>
    [OutputType([string])]
    param()

    Show-Info -Lines @(
        'Bandwidth Management (QoS): controla a largura de banda minima garantida.',
        '[1] None    - Sem gestao de largura de banda (simples, sem QoS).',
        '[2] Default - Reserva minima por peso relativo (recomendado pela Microsoft).',
        '[3] Absolute- Reserva em Mbps absolutos por adaptador virtual.'
    )

    $mode = $null
    do {
        $resp = Read-Host -Prompt '  Modo de Bandwidth Management [1/2/3]'
        $mode = switch ($resp) {
            '1' { 'None'     }
            '2' { 'Default'  }
            '3' { 'Absolute' }
            default { $null  }
        }
        if ($null -eq $mode) {
            Write-Warning 'Opcao invalida. Escolhe 1, 2 ou 3.'
        }
    } while ($null -eq $mode)

    return $mode
}

#endregion

#region --- Criacao do switch ---

function New-HyperVExternalSwitch {
    <#
    .SYNOPSIS
        Cria o External Virtual Switch com todas as opcoes configuradas.
    .PARAMETER SwitchName
        Nome do switch.
    .PARAMETER Adapters
        Adaptadores selecionados (um ou mais).
    .PARAMETER AllowManagementOS
        Se o host partilha a interface de rede.
    .PARAMETER EnableIov
        Activar SR-IOV.
    .PARAMETER BandwidthMode
        Modo de gestao de largura de banda.
    #>
    param(
        [Parameter(Mandatory)] [string]          $SwitchName,
        [Parameter(Mandatory)] [PSCustomObject[]] $Adapters,
        [Parameter(Mandatory)] [bool]            $AllowManagementOS,
        [Parameter(Mandatory)] [bool]            $EnableIov,
        [Parameter(Mandatory)] [string]          $BandwidthMode
    )

    $useSet       = ($Adapters.Count -gt 1)
    $adapterNames = $Adapters.Name

    # --- Verificar switch existente ---
    $existingSwitch = Get-VMSwitch -Name $SwitchName -ErrorAction SilentlyContinue
    if ($null -ne $existingSwitch) {
        Write-Warning ("Ja existe um switch com o nome '{0}'." -f $SwitchName)
        $overwrite = Read-Host -Prompt '  Deseja substituir? [S/N]'
        if ($overwrite -notmatch '^[Ss]$') {
            Write-Warning 'Operacao cancelada pelo utilizador.'
            Write-Log -Message 'Operacao cancelada: switch ja existe e utilizador optou por nao substituir.' -Level 'WARN' -NoConsole
            exit 0
        }
        Write-Verbose ('A remover switch existente: {0}' -f $SwitchName)
        Write-Log     ('A remover switch existente: {0}' -f $SwitchName) -NoConsole
        Remove-VMSwitch -Name $SwitchName -Force
    }

    # --- Apresentar resumo ---
    Show-Header

    $sep = '  ' + ('-' * 60)
    Write-Host '  Resumo da configuracao:' -ForegroundColor White
    Write-Host $sep -ForegroundColor DarkGray
    Write-Host ('  {0,-32} {1}' -f 'Nome do Switch:',          $SwitchName)                    -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'Tipo:',                    'External')                     -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'Interface(s):',            ($adapterNames -join ', '))     -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'SET Teaming:',             $(if ($useSet) {'Sim'} else {'Nao'}))  -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'AllowManagementOS:',       $(if ($AllowManagementOS) {'Sim'} else {'Nao'})) -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'SR-IOV:',                  $(if ($EnableIov) {'Sim'} else {'Nao'}))        -ForegroundColor White
    Write-Host ('  {0,-32} {1}' -f 'Bandwidth Management:',    $BandwidthMode)                 -ForegroundColor White
    Write-Host $sep -ForegroundColor DarkGray
    Write-Host ''

    $confirm = Read-Host -Prompt '  Confirmas a criacao do switch? [S/N]'
    if ($confirm -notmatch '^[Ss]$') {
        Write-Warning 'Operacao cancelada pelo utilizador.'
        Write-Log -Message 'Operacao cancelada na confirmacao final.' -Level 'WARN' -NoConsole
        exit 0
    }

    # --- Construir parametros do New-VMSwitch ---
    Write-Host ''
    Write-Host '  A criar o switch...' -ForegroundColor Gray
    Write-Log ('Criar switch: Name={0} | Adapters={1} | SET={2} | AllowMgmtOS={3} | SR-IOV={4} | BWMode={5}' -f `
        $SwitchName, ($adapterNames -join ','), $useSet, $AllowManagementOS, $EnableIov, $BandwidthMode) -NoConsole

    $switchParams = @{
        Name              = $SwitchName
        NetAdapterName    = $adapterNames
        AllowManagementOS = $AllowManagementOS
        ErrorAction       = 'Stop'
    }

    # SET requer EnableEmbeddedTeaming
    if ($useSet) {
        $switchParams['EnableEmbeddedTeaming'] = $true
        Write-Verbose 'SET (Switch Embedded Teaming) activado.'
    }

    # SR-IOV
    if ($EnableIov) {
        $switchParams['EnableIov'] = $true
        Write-Verbose 'SR-IOV activado.'
    }

    # Bandwidth Mode
    if ($BandwidthMode -ne 'None') {
        $switchParams['MinimumBandwidthMode'] = $BandwidthMode
        Write-Verbose ('Bandwidth mode: {0}' -f $BandwidthMode)
    }

    if ($PSCmdlet.ShouldProcess(
            ("Interface(s): '{0}'" -f ($adapterNames -join "', '")),
            ("Criar External vSwitch '{0}'" -f $SwitchName))) {

        New-VMSwitch @switchParams | Out-Null

        # --- Configurar SET Load Balancing (se SET activo) ---
        if ($useSet) {
            # Microsoft recomenda HyperVPort como algoritmo de load balancing para SET
            Set-VMSwitchTeam -Name $SwitchName -LoadBalancingAlgorithm HyperVPort
            Write-Verbose 'SET Load Balancing configurado: HyperVPort (recomendado Microsoft).'
            Write-Log -Message ('SET Load Balancing: HyperVPort') -NoConsole
        }

        Write-Log ('Switch criado com sucesso: {0}' -f $SwitchName) -NoConsole

        # --- Resultado ---
        Write-Host ''
        Write-Host '  ================================================================' -ForegroundColor Green
        Write-Host '   Switch criado com sucesso!                                    '  -ForegroundColor Green
        Write-Host '  ================================================================' -ForegroundColor Green
        Write-Host ''

        $createdSwitch = Get-VMSwitch -Name $SwitchName

        Write-Host ('  {0,-32} {1}' -f '  Nome:',             $createdSwitch.Name)                  -ForegroundColor White
        Write-Host ('  {0,-32} {1}' -f '  Tipo:',             $createdSwitch.SwitchType)             -ForegroundColor White
        Write-Host ('  {0,-32} {1}' -f '  AllowManagementOS:', $createdSwitch.AllowManagementOS)     -ForegroundColor White
        Write-Host ('  {0,-32} {1}' -f '  SR-IOV:',            $createdSwitch.IovEnabled)            -ForegroundColor White
        Write-Host ('  {0,-32} {1}' -f '  Bandwidth Mode:',    $createdSwitch.BandwidthReservationMode) -ForegroundColor White

        if ($useSet) {
            $team = Get-VMSwitchTeam -Name $SwitchName
            Write-Host ('  {0,-32} {1}' -f '  SET NICs:',      ($team.NetAdapterInterfaceDescription -join ', ')) -ForegroundColor White
            Write-Host ('  {0,-32} {1}' -f '  SET LB Mode:',   $team.LoadBalancingAlgorithm)                      -ForegroundColor White
        }

        Write-Host ''
        Write-Host '  Todos os vSwitches configurados neste servidor:' -ForegroundColor Yellow
        Write-Host ('  ' + '-' * 60) -ForegroundColor DarkGray

        Get-VMSwitch |
            Select-Object -Property Name, SwitchType, AllowManagementOS, IovEnabled, BandwidthReservationMode |
            Format-Table -AutoSize
    }
}

#endregion

#region --- Execucao principal ---

try {
    Initialize-Log -Path $LogPath
    Write-Log -Message ('Inicio da execucao. WhatIf={0} | Verbose={1}' -f `
        $WhatIfPreference, ($VerbosePreference -ne 'SilentlyContinue')) -NoConsole

    # Verificar Hyper-V instalado
    $hypervFeature = Get-WindowsFeature -Name 'Hyper-V' -ErrorAction SilentlyContinue
    if ($null -ne $hypervFeature -and $hypervFeature.InstallState -ne 'Installed') {
        throw 'O Hyper-V nao esta instalado. Executa: Install-WindowsFeature -Name Hyper-V -IncludeManagementTools -Restart'
    }

    # ------------------------------------------------------------------
    # PASSO 1 - Listar interfaces
    # ------------------------------------------------------------------
    Show-Header
    Show-Step -Number 1 -Title 'Selecao de Interface(s) de Rede'
    Write-Verbose 'A detectar adaptadores de rede fisicos e capacidades...'

    [PSCustomObject[]] $adapters = Get-PhysicalNetAdapter

    if (($adapters | Measure-Object).Count -eq 0) {
        throw 'Nenhuma interface de rede disponivel foi encontrada neste servidor.'
    }

    Show-AdapterTable -Adapters $adapters

    # ------------------------------------------------------------------
    # PASSO 2 - Selecao de NIC(s)
    # ------------------------------------------------------------------
    [PSCustomObject[]] $selectedAdapters = Read-AdapterSelection -Adapters $adapters

    $useSet = ($selectedAdapters.Count -gt 1)

    Write-Host ''
    if ($useSet) {
        Write-Host '  Modo: ' -NoNewline -ForegroundColor Gray
        Write-Host 'SET (Switch Embedded Teaming) com ' -NoNewline -ForegroundColor Cyan
        Write-Host ('{0} NICs' -f $selectedAdapters.Count) -ForegroundColor Green
    } else {
        Write-Host '  Modo: ' -NoNewline -ForegroundColor Gray
        Write-Host 'NIC unico' -ForegroundColor Cyan
    }
    foreach ($nic in $selectedAdapters) {
        Write-Host ('    - {0,-20} {1}' -f $nic.Name, $nic.InterfaceDescription) -ForegroundColor White
    }

    Write-Log ('NICs selecionados: {0}' -f ($selectedAdapters.Name -join ', ')) -NoConsole

    # ------------------------------------------------------------------
    # PASSO 3 - Nome do switch
    # ------------------------------------------------------------------
    Show-Step -Number 3 -Title 'Nome do Virtual Switch'
    $switchName = Read-SwitchName -AdapterNames $selectedAdapters.Name -UseSet $useSet

    Write-Log ('Nome do switch: {0}' -f $switchName) -NoConsole

    # ------------------------------------------------------------------
    # PASSO 4 - AllowManagementOS
    # ------------------------------------------------------------------
    Show-Step -Number 4 -Title 'Acesso do SO Host (AllowManagementOS)'
    $allowMgmtOS = Read-AllowManagementOS

    Write-Log ('AllowManagementOS: {0}' -f $allowMgmtOS) -NoConsole

    # ------------------------------------------------------------------
    # PASSO 5 - SR-IOV
    # ------------------------------------------------------------------
    Show-Step -Number 5 -Title 'SR-IOV (Single Root I/O Virtualization)'
    $anySriovSupport = ($selectedAdapters | Where-Object { $_.SRIOVSupport } | Measure-Object).Count -gt 0
    $enableSriov     = Read-EnableSriov -AnyNicSupports $anySriovSupport

    Write-Log ('SR-IOV: {0}' -f $enableSriov) -NoConsole

    # ------------------------------------------------------------------
    # PASSO 6 - Bandwidth Management
    # ------------------------------------------------------------------
    Show-Step -Number 6 -Title 'Bandwidth Management (QoS)'
    $bandwidthMode = Read-BandwidthMode

    Write-Log ('Bandwidth Mode: {0}' -f $bandwidthMode) -NoConsole

    # ------------------------------------------------------------------
    # PASSO 7 - Criacao do switch
    # ------------------------------------------------------------------
    Show-Step -Number 7 -Title 'Confirmacao e Criacao'

    New-HyperVExternalSwitch `
        -SwitchName        $switchName `
        -Adapters          $selectedAdapters `
        -AllowManagementOS $allowMgmtOS `
        -EnableIov         $enableSriov `
        -BandwidthMode     $bandwidthMode

    Write-Host ('  Log guardado em: {0}' -f $LogPath) -ForegroundColor DarkGray
    Write-Log -Message 'Execucao concluida com sucesso.' -NoConsole
}
catch {
    Write-Host ''
    Write-Error -Message ('Erro critico: {0}' -f $_.Exception.Message)
    Write-Log   -Message ('Erro critico: {0}' -f $_.Exception.Message) -Level 'ERROR' -NoConsole
    Write-Host ''
    Write-Host '  Causas possiveis:' -ForegroundColor Yellow
    Write-Host '  - Interface ja em uso noutro vSwitch'                       -ForegroundColor DarkGray
    Write-Host '  - SR-IOV requer reinicio apos activacao no BIOS/UEFI'       -ForegroundColor DarkGray
    Write-Host '  - SET nao e compativel com NIC Teaming tradicional activo'  -ForegroundColor DarkGray
    Write-Host '  - Servico vmms (Hyper-V) nao esta em execucao'              -ForegroundColor DarkGray
    Write-Host ''
    Write-Host ('  Log disponivel em: {0}' -f $LogPath) -ForegroundColor DarkGray
    exit 1
}
finally {
    Write-Host ''
    $null = Read-Host -Prompt '  Prima Enter para sair'
}

#endregion
