# Windows Downloader
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Name Window -Namespace Console -MemberDefinition '[DllImport("Kernel32.dll")]public static extern IntPtr GetConsoleWindow();[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
[void][Console.Window]::ShowWindow([Console.Window]::GetConsoleWindow(), 0)
[System.Windows.Forms.Application]::EnableVisualStyles()
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$wc = New-Object net.webclient
$mb = [System.Windows.Forms.MessageBox]::Show
$app = 'Windows Downloader'
$n = [Environment]::NewLine
$hidden = New-Object System.Windows.Forms.Form -Property @{
    TopMost = $true
    WindowState = "Minimized"
}

if ($PSUICulture -eq "ru-RU") {
    $strings = @("Активация",
    "Скачивание",
    "Другие функции",
    "Информация",
    "Активация Windows 10 или 11 (в том числе LTSC) по HWID",
    "Конвертировать пробную версию Windows 10 LTSC (Evaluation) в полноценную LTSC",
    "Активация Windows Server 2022, Windows Server Standard, Windows Server Datacenter, 2019, 2016, 2012, 2012 R2, 1803, 1709",
    "Через KMS, будет активирован как %p% 2021 (более старые версии обновятся)",
    ("Активация Office %v% путём добавления файла sppc.dll$n" + "%info%" + "Активация сработает и для Office %otherv%. Office потом автоматически конвертируется в %v%."),
    "Office 365 — всегда самая актуальная версия Office, лучше выбрать этот вариант.$n",
    "И всё-таки, я бы порекомендовал выбрать Office 365.$n",
    "Активация Office 2013 с помощью добавления файла sppc.dll",
    ("Разрешить создание автономного аккаунта Minecraft в Prism Launcher без добавления аккаунта Microsoft.$n" + "Не запускайте, если вы уже добавили аккаунт! Это действие удалит все аккаунты в лаунчере!"),
    "Премиум-аккаунт в TL, вы сможете отключить добавление рекламных серверов в его настройках",
    "Автономный аккаунт в Prism Launcher разблокирован!",
    "%p% не найден!",
    "Premium аккаунт в TL активирован! Теперь зайдите в настройки TL для отключения рекламных серверов!",
    "Ни один аккаунт не доблавен в TL!",
    "ISO образ последней версии Windows %v% с официального сайта Microsoft",
    "Онлайн-установщик Office с официального сайта Microsoft.",
    "ISO архив Office с официального сайта Microsoft. Запустите в нём setup.exe",
    " Следуйте инструкциям $app после запуска установщика.",
    "Не рекомендуется, устаревшая версия. ",
    "инструмент для записи ISO образов на флешку",
    "Не удалось получить ссылку для загрузки! Попробовать снова?",
    "Скачанному файлу нужно будет дописать формат .iso!",
    "рекомендуется",
    'Для онлайн-установки нужно обойти гео-ограничения. Для этого: запустите exe-файл, дождитесь ошибки "Сбой установки", нажмите "Да" в этом окне. После этого перезапустите файл установщика!',
    "Установить ",
    "Для LTSC-версий Windows без установленного Microsoft Store",
    "Удалить системные шпионские программы",
    "Удалить CompatTelRunner.exe и wsqmcons.exe",
    "Обойти гео-ограничения веб-сервисов",
    ("Обойти гео-ограничения через редактирование файла hosts$n" + "Будут работать ChatGPT, Gemini, NotebookLM, Copilot, Spotify, Codeium, GitHub Copilot, Claude, Notion, Canva, TikTok$n" + "Без VPN и других сторонних приложений"),
    "Инструмент удаления Office",
    "Активировать!",
    "https://drive.massgrave.dev/ru-ru_windows_10_enterprise_ltsc_2021_x64_dvd_5044a1e7.iso",
    "https://drive.massgrave.dev/ru-ru_windows_11_enterprise_ltsc_2024_x64_dvd_f9af5773.iso",
    "https://oemsoc.download.prss.microsoft.com/dbazure/X23-81967_26100.1742.240906-0331.ge_release_svc_refresh_SERVER_OEMRET_x64FRE_ru-ru.iso_0400d135-3d94-49a2-8627-8f1a8cb316bf?t=27afd6c5-3c63-4984-8139-b9c239276cb4&P1=102817441539&P2=601&P3=2&P4=K6P6PaBziMqVvDg7AgCqTBprjEMuo%2bmjluaix%2b9TaUldONUCc3PtGs30Rvmn3IKMuSZ7kcmGydK%2bmz38quTSTCyGmjPdKm6bLG%2f2m13pTKsdD1zp%2flccTbTkwvIN%2fdhU8qzwet9V56is8W7o7IykKbczeFlJ1yQV7xq6OCpOzudqomW5fUsUO0%2fRx%2b78zkGgyrHlxIQlX9bAC5Fr069%2byhr5OiXWk9R%2fzEj93%2bEfBrZMTFz1M%2fzf6UKw6tYjOjdSJkNKk%2bhjnAyC%2bcqCj2OKrw6yhEJ6vtXbNJomDZzfUBqMM%2f1uoRabPzPv5Adp3XEJ5DIzdBU%2foyhPbj0qcCzfPg%3d%3d",
    "https://drive.massgrave.dev/ru-ru_windows_server_2022_updated_nov_2024_x64_dvd_4e34897c.iso",
    "https://drive.massgrave.dev/ru-ru_windows_server_2019_x64_dvd_e02b76ba.iso",
    "https://drive.massgrave.dev/ru_windows_server_2016_vl_x64_dvd_11636694.iso",
    "https://drive.massgrave.dev/ru_windows_server_2012_r2_vl_with_update_x64_dvd_6052827.iso",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=ru-ru&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/ProPlus2024Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=ru-ru&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/ProPlus2021Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=ru-ru&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/ProPlus2019Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=ru-ru&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/ProPlusRetail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=ru-ru&version=O15GA",
    "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/ru-ru/ProfessionalRetail.img",
    "Нужно, чтобы приложение Техническая поддержка (Get Help) было установлено! Выберите Да и дождитесь завершения процесса",
    "https://drive.massgrave.dev/ru_windows_8.1_pro_vl_with_update_x64_dvd_6050899.iso",
    "Теперь снова запустите установщик Office!",
    "Microsoft Store устанавливается! Подождите пару минут, и он должен появиться в Пуске!",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/VisioPro2024Retail.img",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/ru-ru/ProjectPro2024Retail.img",
    "Удаление CompatTelRunner.exe и wsqmcons.exe",
    "Редактирую hosts файл",
    "https://drive.massgrave.dev/ru_windows_7_ultimate_with_sp1_x64_dvd_u_677391.iso",
    "https://drive.massgrave.dev/ru_windows_8_x64_dvd_915419.iso")
} else {
    $strings = @("Activation",
    "Download",
    "Other functions",
    "Info",
    "Windows 10 or 11 (including LTSC) activation by HWID",
    "Convert evaluation version of Windows 10 LTSC to full LTSC",
    "Activation of Windows Server 2022, Windows Server Standard, Windows Server Datacenter, 2019, 2016, 2012, 2012 R2, 1803, 1709",
    "Via KMS, will be activated as %p% 2021 (older versions will be updated)",
    ("Office %v% activation via sppc.dll file$n" + "%info%" + "The activation will also work for Office %otherv%. Office will then be automatically converted to %v%."),
    "Office 365 is always the latest version of Office, it is better to choose this option.$n",
    "Anyway, I would recommend selecting Office 365.$n",
    "Office 2013 activation via sppc.dll file.",
    ("Allow creation of an offline Minecraft account in Prism Launcher without Microsoft account.$n" + "Do not start if you have already added an account! This action will delete all accounts in the launcher!"),
    "Premium account in TL, you will be able to disable adding advertised servers in its settings",
    "Offline account in Prism Launcher unlocked!",
    "%p% not found!",
    "Premium account in TL is activated! Now open its settings to disable advertised servers!",
    "No account has been added to TL!",
    "ISO image of the latest version of Windows %v% from the official Microsoft website",
    "Office online installer from the official Microsoft website.",
    "ISO archive of Office from the official Microsoft website. Run setup.exe in it.",
    " Follow the instructions of $app after starting the installer.",
    "Not recommended, outdated version. ",
    "tool for writing ISO images to a flash drive",
    "Failed to get a download link! Try again?",
    "Downloaded file needs to be appended with .iso!",
    "recommended",
    'To install online, you need to bypass geo-restrictions. To do this: run the exe-file, wait for the "Command not supported" error, click "Yes" in this window. After this, restart the installer file!',
    "Install ",
    "For LTSC versions of Windows without installed Microsoft Store!",
    "Delete system spy programs",
    "Delete CompatTelRunner.exe and wsqmcons.exe",
    "Bypass web-apps geo-blocking",
    ("Bypass geo-restrictions via editing hosts file$n" + "The following services will work: ChatGPT, Gemini, NotebookLM, Copilot, Spotify, Codeium, GitHub Copilot, Claude, Notion, Canva, TikTok$n" + "Without VPN and other apps"),
    "Office uninstall tool",
    "Activate!",
    "https://drive.massgrave.dev/en-us_windows_10_iot_enterprise_ltsc_2021_x64_dvd_257ad90f.iso",
    "https://drive.massgrave.dev/en-us_windows_11_iot_enterprise_ltsc_2024_x64_dvd_f6b14814.iso",
    "https://oemsoc.download.prss.microsoft.com/dbazure/X23-81958_26100.1742.240906-0331.ge_release_svc_refresh_SERVER_OEMRET_x64FRE_en-us.iso_909fa35d-ba98-407d-9fef-8df76f75e133?t=34b8db0f-439b-497c-86ce-ec7ceb898bb7&P1=102816956391&P2=601&P3=2&P4=pG1WoVpBKlyWcmfj%2bt1gYgkTsP4At28ch8mG7vIQm%2fT4elz5v2ZQ3eKAN8%2fFjb1yaa4npBaABURtnI8YmrDv8p0VJmYpLCIUQ0FHEFR4IFiPgtvzwAAI8oNdiEl%2b2uM7MN8Gaju8BvIVgHRl%2fRxq0HFgrFoEGmvHZU4jY0RFsYAaHliUinDUzdVfT0IPwyWqNUJXZTSfguyphv8XZx8OQsBy3zwBp7tNHsKl36ZO2JdZK%2fyPY7QTpAr5ccazUPEa40ALhYRBJXxlQb1F0OeO7kHhW7DKK5D4Wpt5WbpjFn8MqcZBX3%2fQI6WAwzDSKIck7jYL7bYdl2ufoMRrFZrxxw%3d%3d",
    "https://drive.massgrave.dev/en-us_windows_server_2022_updated_nov_2024_x64_dvd_4e34897c.iso",
    "https://drive.massgrave.dev/en-us_windows_server_2019_x64_dvd_f9475476.iso",
    "https://drive.massgrave.dev/en_windows_server_2016_vl_x64_dvd_11636701.iso",
    "https://drive.massgrave.dev/en_windows_server_2012_r2_vl_with_update_x64_dvd_6052766.iso",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=en-us&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlusRetail.img",
    "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O15GA",
    "https://officecdn.microsoft.com/db/39168d7e-077b-48e7-872c-b232c3e72675/media/en-us/ProfessionalRetail.img",
    "You will need Get Help app installed! When it opens, select Yes and wait for it to complete!",
    "https://drive.massgrave.dev/en-gb_windows_8.1_pro_vl_with_update_x64_dvd_6050881.iso",
    "Now start Office setup again!",
    "Microsoft Office is installing! Wait a few minutes, it will appear in Start!",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/VisioPro2024Retail.img",
    "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProjectPro2024Retail.img",
    "Deleting CompatTelRunner.exe and wsqmcons.exe",
    "Editing hosts file",
    "https://drive.massgrave.dev/en_windows_7_ultimate_with_sp1_x64_dvd_u_677391.iso",
    "https://drive.massgrave.dev/en_windows_8_x64_dvd_915419.iso")
}
$gstrings = @("irm https://raw.githubusercontent.com/ImMALWARE/$app/main/Activators", " | iex", "$env:ProgramFiles\Microsoft Office\root\vfs\System")
$activators = @{"Win10" = "HWID.ps1"; "ConvertEvaluationToFull" = "LTSCEvaluationToFull.ps1"; "WinServer" = "ServerKMS.ps1"; "OfficeVisio" = "VisioProject.ps1"; "OfficeProject" = "VisioProject.ps1"; "MobaXterm" = "MXT.ps1"; "Office 365" = "Osppcs.ps1"; "Office 2024" = "Osppcs.ps1"; "Office 2021" = "Osppcs.ps1"; "Office 2019" = "Osppcs.ps1"; "Office 2016" = "Osppcs.ps1"; "Office 2013" = "Osppcs2013.ps1"}
$paths = @{"Office 365" = $gstrings[2]; "Office 2024" = $gstrings[2]; "Office 2021" = $gstrings[2]; "Office 2019" = $gstrings[2]; "Office 2016" = $gstrings[2]; "Office 2013" = "$env:ProgramFiles\Microsoft Office 15\root\vfs\System"; "Prism Launcher" = "$env:appdata\PrismLauncher"; "TL" = "$env:appdata\.minecraft\TlauncherProfiles.json"; "MobaXterm" = "${env:ProgramFiles(x86)}\Mobatek\MobaXterm\version.dat"}

function check_path ($path, $prod) {
    if (Test-Path $path) {
        return $true
    } else {
        $null = $mb.Invoke($strings[15].Replace('%p%', $prod))
        return $false
    }
}

$form = New-Object System.Windows.Forms.Form -Property @{
    StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    MaximizeBox = $false
    Text = $app
    ClientSize = [System.Drawing.Size]::new(627, 234)
    Font = [System.Drawing.Font]::new('Segoe UI', 8.5)
}

$tabs = New-Object System.Windows.Forms.TabControl -Property @{
    Dock = [System.Windows.Forms.DockStyle]::Fill
    Location = [System.Drawing.Point]::new(0, 0)
    SelectedIndex = 0
    Size = [System.Drawing.Size]::new(627, 234)
    SizeMode = [System.Windows.Forms.TabSizeMode]::FillToRight
}

$ActTab = New-Object System.Windows.Forms.TabPage -Property @{
    Location = [System.Drawing.Point]::new(4, 24)
    Padding = [System.Windows.Forms.Padding]::new(3)
    Size = [System.Drawing.Size]::new(619, 206)
    Text = $strings[0]
}

$DlTab = New-Object System.Windows.Forms.TabPage -Property @{
    Location = [System.Drawing.Point]::new(4, 24)
    Padding = [System.Windows.Forms.Padding]::new(3)
    Size = [System.Drawing.Size]::new(619, 206)
    Text = $strings[1]
}

$FunctionsTab = New-Object System.Windows.Forms.TabPage -Property @{
    Location = [System.Drawing.Point]::new(4, 24)
    Size = [System.Drawing.Size]::new(619, 206)
    Text = $strings[2]
}

@($ActTab, $DlTab, $FunctionsTab) | ForEach-Object { $tabs.TabPages.Add($_) }

$tooltip = New-Object System.Windows.Forms.ToolTip -Property @{
    AutoPopDelay = 5000
    InitialDelay = 5
    ReshowDelay = 100
    ShowAlways = $true
}

# Activation tab

$ActWin10 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Checked = $true
    Location = [System.Drawing.Point]::new(6, 6)
    Name = "Win10"
    Size = [System.Drawing.Size]::new(143, 19)
    TabStop = $true
    Text = "Windows 10/11 (HWID)"
}
$tooltip.SetToolTip($ActWin10, $strings[4])

$ConvertEvaluationToFull = New-Object System.Windows.Forms.RadioButton -Property @{ 
    AutoSize = $true
    Location = [System.Drawing.Point]::new(6, 31)
    Name = "ConvertEvaluationToFull"
    Size = [System.Drawing.Size]::new(130, 19)
    Text = "LTSC Evaluation -> LTSC"
}
$tooltip.SetToolTip($ConvertEvaluationToFull, $strings[5])

$ActWinServer = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(6, 56)
    Name = "WinServer"
    Size = [System.Drawing.Size]::new(193, 19)
    Text = "Windows Server 2022/2019/2016"
}
$tooltip.SetToolTip($ActWinServer, $strings[6])

$ActVisio = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(6, 106)
    Name = "OfficeVisio"
    Size = [System.Drawing.Size]::new(54, 19)
    Text = "Visio 2016/2019/2021"
}
$tooltip.SetToolTip($ActVisio, $strings[7].Replace("%p%", "Visio"))

$ActProject = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(6, 131)
    Name = "OfficeProject"
    Size = [System.Drawing.Size]::new(64, 19)
    Text = "Project 2016/2019/2021"
}
$tooltip.SetToolTip($ActProject, $strings[7].Replace("%p%", "Project"))

$ActOffice365 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 6)
    Name = "Office 365"
    Size = [System.Drawing.Size]::new(79, 19)
    Text = "Office 365"
}
$tooltip.SetToolTip($ActOffice365, $strings[8].Replace("%v%", "365").Replace("%info%", $strings[9]).Replace("%otherv%", "2016, 2019, 2021, 2024"))

$ActOffice2024 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 31)
    Name = "Office 2024"
    Size = [System.Drawing.Size]::new(83, 19)
    Text = "Office 2024"
}
$tooltip.SetToolTip($ActOffice2024, $strings[8].Replace("%v%", "2024").Replace("%info%", $strings[10]).Replace("%otherv%", "2016, 2019, 2021"))

$ActOffice2021 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 56)
    Name = "Office 2021"
    Size = [System.Drawing.Size]::new(83, 19)
    Text = "Office 2021"
}
$tooltip.SetToolTip($ActOffice2021, $strings[8].Replace("%v%", "2021").Replace("%info%", "").Replace("%otherv%", "2016, 2019, 2024"))

$ActOffice2019 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 81)
    Name = "Office 2019"
    Size = [System.Drawing.Size]::new(84, 19)
    Text = "Office 2019"
}
$tooltip.SetToolTip($ActOffice2019, $strings[8].Replace("%v%", "2019").Replace("%info%", "").Replace("%otherv%", "2016, 2021, 2024"))

$ActOffice2016 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 106)
    Name = "Office 2016"
    Size = [System.Drawing.Size]::new(84, 19)
    Text = "Office 2016"
}
$tooltip.SetToolTip($ActOffice2016, $strings[8].Replace("%v%", "2016").Replace("%info%", "").Replace("%otherv%", "2019, 2021, 2024"))

$ActOffice2013 = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(214, 131)
    Name = "Office 2013"
    Size = [System.Drawing.Size]::new(83, 19)
    Text = "Office 2013"
}
$tooltip.SetToolTip($ActOffice2013, $strings[11])

$ActPrismLauncher = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(373, 6)
    Name = "Prism Launcher"
    Size = [System.Drawing.Size]::new(201, 19)
    Text = "Prism Launcher"
}
$tooltip.SetToolTip($ActPrismLauncher, $strings[12])

$ActTL = New-Object System.Windows.Forms.RadioButton -Property @{
    AutoSize = $true
    Location = [System.Drawing.Point]::new(373, 31)
    Name = "TL"
    Size = [System.Drawing.Size]::new(81, 19)
    Text = "TL"
}
$tooltip.SetToolTip($