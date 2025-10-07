
# Executa como Administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funï¿½ï¿½o de LOG com timestamp
function Log($msg) {
    $time = (Get-Date).ToString("HH:mm:ss")
    $outputBox.AppendText("[$time] $msg`r`n")
    [System.Windows.Forms.Application]::DoEvents()
}

# Checa conexï¿½o com a internet (validando DNS do Google)
function Test-Internet {
    try {
        $test = Test-NetConnection -ComputerName "8.8.8.8" -Port 53 -InformationLevel Quiet
        return $test
    } catch {
        return $false
    }
}

function Disable-SysMain {
    try {
        Log "Desativando SysMain..."
        Stop-Service -Name "SysMain" -Force -ErrorAction Stop
        Set-Service -Name "SysMain" -StartupType Disabled -ErrorAction Stop
        Log "SysMain desativado com sucesso."
    } catch {
        Log "Erro ao desativar SysMain: $_"
    }
}

function Run-Cleano {
    try {
        Log "Executando Cleano..."
        $folders = @(
            "$env:TEMP",
            "$env:APPDATA\Microsoft\Windows\Recent",
            "$env:APPDATA\Microsoft\Windows\NetHood",
            "C:\Windows\Temp",
            "C:\Windows\Prefetch"
        )
        foreach ($folder in $folders) {
            if (Test-Path $folder) {
                $count = (Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue).Count
                if ($count -gt 0) {
                    Get-ChildItem -Path $folder -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                    Log "$count arquivo(s) limpos em $folder."
                } else {
                    Log "Nada para limpar em $folder."
                }
            }
        }
        if (Get-Command Clear-RecycleBin -ErrorAction SilentlyContinue) {
            Clear-RecycleBin -Force -ErrorAction SilentlyContinue
            Log "Lixeira limpa."
        }
        try { RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8; Log "Cache do IE limpo." } catch {}
        if (Test-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU") {
            Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU" -Recurse -Force -ErrorAction SilentlyContinue
            Log "Histï¿½rico 'Executar' limpo."
        }
        Log "Cleano finalizado com sucesso."
    } catch {
        Log "Erro ao executar Cleano: $_"
    }
}

function Set-HighPerformancePower {
    try {
        Log "Aplicando plano de energia de alto desempenho..."
        powercfg -setactive SCHEME_MIN
        powercfg -change -monitor-timeout-ac 0
        powercfg -change -standby-timeout-ac 0
        powercfg -change -disk-timeout-ac 0
        Log "Energia otimizada para alto desempenho."
    } catch {
        Log "Erro ao aplicar energia alto desempenho: $_"
    }
}

function Open-VisualSettings {
    Log "Abrindo ajustes de desempenho visual..."
    Start-Process "SystemPropertiesPerformance.exe"
    Log "Tela de desempenho aberta. Aplique manualmente 'Melhor Desempenho'."
}

function Activate-WindowsOffice {
    try {
        Log "Ativando o Windows e Office..."
        Start-Process powershell -WindowStyle Hidden -ArgumentList '-NoProfile -ExecutionPolicy Bypass -Command "irm https://get.activated.win | iex"' -Verb RunAs
        Log "Ativador executado. Aguarde confirmaï¿½ï¿½o do sistema."
    } catch {
        Log "Erro no ativador: $_"
    }
}

function Apply-Wallpaper {
    try {
        Log "Aplicando wallpaper da loja..."
        $url = "https://lh3.googleusercontent.com/d/1O98wBeT6O7QBO5O_pdFCSxniEEwPwDm1"
        $path = "$env:APPDATA\wallpaper_loja.jpg"
        Invoke-WebRequest -Uri $url -OutFile $path -UseBasicParsing -ErrorAction Stop
        Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
        [Wallpaper]::SystemParametersInfo(20, 0, $path, 3)
        Log "Wallpaper aplicado com sucesso."
    } catch {
        Log "Erro ao aplicar wallpaper: $_"
    }
}

# Instalar apps via Winget (inclui Anydesk)
function Install-AppsViaWinget {
    $apps = @(
        @{ Name = "Google Chrome"; Id = "Google.Chrome" },
        @{ Name = "Mozilla Firefox"; Id = "Mozilla.Firefox" },
        @{ Name = "Foxit PDF Reader"; Id = "Foxit.FoxitReader" },
        @{ Name = "WinRAR"; Id = "RARLab.WinRAR" },
        @{ Name = "VLC Media Player"; Id = "VideoLAN.VLC" },
        @{ Name = "AnyDesk"; Id = "AnyDeskSoftwareGmbH.AnyDesk" }
    )

    foreach ($app in $apps) {
        Log "Instalando $($app.Name)..."
        try {
            winget install --id $($app.Id) --silent --accept-package-agreements --accept-source-agreements
            Log "$($app.Name) instalado!"
        } catch {
            Log "Erro ao instalar $($app.Name): $_"
        }
    }

    Log "Instalando .NET Framework 4.8..."
    try {
        winget install --id "Microsoft.DotNet.Framework.DeveloperPack_4" --silent --accept-package-agreements --accept-source-agreements
        Log ".NET Framework 4.8 instalado!"
    } catch {
        Log "Erro ao instalar .NET Framework 4.8: $_"
    }
}

# DOWNLOAD: OfficeSetup.exe na ï¿½rea de Trabalho ï¿½ apenas download, oferece abrir instalador no fim
function Download-OfficeSetup {
    try {
        Log "Baixando instalador do Office 365 64bits PT-BR para a ï¿½rea de Trabalho..."
        $officeUrl = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=pt-br&version=O16GA"
        $desktop = [Environment]::GetFolderPath("Desktop")
        $file = Join-Path $desktop "OfficeSetup.exe"
        Invoke-WebRequest -Uri $officeUrl -OutFile $file -UseBasicParsing -ErrorAction Stop
        Log "Download concluï¿½do: $file"

        # Oferece ao usuï¿½rio instalar jï¿½ na sequï¿½ncia
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Download do Office concluï¿½do!\nDeseja iniciar a instalaï¿½ï¿½o agora?\n(Clique 'Nï¿½o' para instalar depois pela ï¿½rea de Trabalho)",
            "Download do Office",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Log "Iniciando o instalador do Office..."
            Start-Process -FilePath $file -Verb RunAs
        } else {
            Log "Instalador do Office Nï¿½O foi iniciado. Estï¿½ disponï¿½vel na ï¿½rea de Trabalho."
        }
    } catch {
        Log "Erro ao baixar o Office: $_"
    }
}

# --- Interface Grï¿½fica (WinForms) ---
$form = New-Object Windows.Forms.Form
$form.BackColor = [System.Drawing.Color]::FromArgb(245,245,250)
$form.Font = New-Object Drawing.Font("Segoe UI", 9)
$form.FormBorderStyle = "FixedDialog"
$form.Text = "OitoBits Informatica - ToolKit"
$form.Size = New-Object Drawing.Size(540,850)
$form.StartPosition = "CenterScreen"

# Logotipo
$logo = New-Object Windows.Forms.PictureBox
$logo.ImageLocation = "https://lh3.googleusercontent.com/d/1NbOK-zyWIIZ5H9wobBirfO8PJAgOH5Sc"
$logo.SizeMode = "StretchImage"
$logo.Size = New-Object Drawing.Size(420, 70)
$logo.Location = New-Object Drawing.Point(20,15)
$form.Controls.Add($logo)

# Array para checkboxes
$allCheckboxes = @()

$chkSelecionarTudo = New-Object Windows.Forms.CheckBox
$chkSelecionarTudo.Text = "Selecionar Todas"
$chkSelecionarTudo.Location = New-Object Drawing.Point(20,90)
$chkSelecionarTudo.AutoSize = $true
$chkSelecionarTudo.Checked = $true
$form.Controls.Add($chkSelecionarTudo)

# Seï¿½ï¿½o Otimizar
$lblManutencao = New-Object Windows.Forms.Label
$lblManutencao.Text = "Otimizar"
$lblManutencao.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblManutencao.Location = New-Object Drawing.Point(20,115)
$form.Controls.Add($lblManutencao)

$chkSysMain = New-Object Windows.Forms.CheckBox
$chkSysMain.Text = "Desativar SysMain"
$chkSysMain.Location = New-Object Drawing.Point(40,135)
$chkSysMain.AutoSize = $true
$chkSysMain.Checked = $true
$form.Controls.Add($chkSysMain)
$allCheckboxes += $chkSysMain

$chkCleano = New-Object Windows.Forms.CheckBox
$chkCleano.Text = "Cleano (Limpeza Geral do Sistema)"
$chkCleano.Location = New-Object Drawing.Point(40,165)
$chkCleano.AutoSize = $true
$chkCleano.Checked = $true
$form.Controls.Add($chkCleano)
$allCheckboxes += $chkCleano

# Seï¿½ï¿½o Performance
$lblPerformance = New-Object Windows.Forms.Label
$lblPerformance.Text = "Performance"
$lblPerformance.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblPerformance.Location = New-Object Drawing.Point(20,195)
$form.Controls.Add($lblPerformance)

$chkPower = New-Object Windows.Forms.CheckBox
$chkPower.Text = "Energia Alto Desempenho"
$chkPower.Location = New-Object Drawing.Point(40,215)
$chkPower.AutoSize = $true
$chkPower.Checked = $true
$form.Controls.Add($chkPower)
$allCheckboxes += $chkPower

# Seï¿½ï¿½o Personalizar
$lblVisual = New-Object Windows.Forms.Label
$lblVisual.Text = "Personalizar"
$lblVisual.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblVisual.Location = New-Object Drawing.Point(20,245)
$form.Controls.Add($lblVisual)

$chkVisuals = New-Object Windows.Forms.CheckBox
$chkVisuals.Text = "Abrir ajustes de desempenho visual"
$chkVisuals.Location = New-Object Drawing.Point(40,265)
$chkVisuals.AutoSize = $true
$chkVisuals.Checked = $true
$form.Controls.Add($chkVisuals)
$allCheckboxes += $chkVisuals

$chkWallpaper = New-Object Windows.Forms.CheckBox
$chkWallpaper.Text = "Wallpaper OitoBits Informatica"
$chkWallpaper.Location = New-Object Drawing.Point(40,295)
$chkWallpaper.AutoSize = $true
$chkWallpaper.Checked = $true
$form.Controls.Add($chkWallpaper)
$allCheckboxes += $chkWallpaper

# Seï¿½ï¿½o Licenciamento
$lblLicenciamento = New-Object Windows.Forms.Label
$lblLicenciamento.Text = "Licenciamento"
$lblLicenciamento.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblLicenciamento.Location = New-Object Drawing.Point(20,325)
$form.Controls.Add($lblLicenciamento)

$chkAtivador = New-Object Windows.Forms.CheckBox
$chkAtivador.Text = "Ativar o Windows e Office"
$chkAtivador.Location = New-Object Drawing.Point(40,345)
$chkAtivador.AutoSize = $true
$chkAtivador.Checked = $true
$form.Controls.Add($chkAtivador)
$allCheckboxes += $chkAtivador

# Seï¿½ï¿½o Programas (Instalaï¿½ï¿½o via Winget)
$lblProgramas = New-Object Windows.Forms.Label
$lblProgramas.Text = "Instalar"
$lblProgramas.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
$lblProgramas.Location = New-Object Drawing.Point(20,370)
$form.Controls.Add($lblProgramas)

$chkWinget = New-Object Windows.Forms.CheckBox
$chkWinget.Text = "Instalar Programas (Chrome, Firefox, Foxit, WinRAR, VLC, AnyDesk, .NET 4.8)"
$chkWinget.Location = New-Object Drawing.Point(40,390)
$chkWinget.AutoSize = $true
$chkWinget.Checked = $false
$form.Controls.Add($chkWinget)
$allCheckboxes += $chkWinget

# Seï¿½ï¿½o Office (caixa sï¿½ para download na ï¿½rea de trabalho)
$chkDownloadOffice = New-Object Windows.Forms.CheckBox
$chkDownloadOffice.Text = "Baixar instalador do Office 365 (64bits PT-BR) para a area de Trabalho"
$chkDownloadOffice.Location = New-Object Drawing.Point(40,420)
$chkDownloadOffice.AutoSize = $true
$chkDownloadOffice.Checked = $false
$form.Controls.Add($chkDownloadOffice)
$allCheckboxes += $chkDownloadOffice

# Checkbox Selecionar Tudo - sincronizaï¿½ï¿½o Dinï¿½mica
$chkSelecionarTudo.Add_CheckedChanged({
    foreach ($chk in $allCheckboxes) { $chk.Checked = $chkSelecionarTudo.Checked }
})

# Caixa de saï¿½da/output
$outputBox = New-Object Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Size = New-Object Drawing.Size(480,140)
$outputBox.Location = New-Object Drawing.Point(20,460)
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

# Aviso
$lblAviso = New-Object Windows.Forms.Label
$lblAviso.Text = "Reinicie o sistema quando finalizar."
$lblAviso.Location = New-Object Drawing.Point(20,610)
$lblAviso.Size = New-Object Drawing.Size(480,40)
$form.Controls.Add($lblAviso)

# Botï¿½o Aplicar
$btnAplicar = New-Object Windows.Forms.Button
$btnAplicar.Text = "Aplicar"
$btnAplicar.Location = New-Object Drawing.Point(200,710)
$btnAplicar.Add_Click({
    # CHECAGEM DE INTERNET ANTES DE QUALQUER COISA
    if (-not (Test-Internet)) {
        Log "Sem conexao com a internet. Operacoes canceladas."
        [System.Windows.Forms.MessageBox]::Show($form, "Sem conexao com a internet. Verifique sua rede e tente novamente.", "OitoBits Informatica - ToolKit")
        $form.Close()
        return
    }
    if ($chkSysMain.Checked) { Disable-SysMain }
    if ($chkCleano.Checked) { Run-Cleano }
    if ($chkVisuals.Checked) { Open-VisualSettings }
    if ($chkAtivador.Checked) { Activate-WindowsOffice }
    if ($chkPower.Checked) { Set-HighPerformancePower }
    if ($chkWallpaper.Checked) { Apply-Wallpaper }
    if ($chkWinget.Checked) { Install-AppsViaWinget }
    if ($chkDownloadOffice.Checked) { Download-OfficeSetup }
    Log "Tarefas finalizadas!"
    Start-Sleep -Seconds 2.5
    [System.Windows.Forms.MessageBox]::Show($form, "Sucesso!","OitoBits Informatica - ToolKit")
    $form.Close()
})
$form.Controls.Add($btnAplicar)

$form.Topmost = $true
[void]$form.ShowDialog()

