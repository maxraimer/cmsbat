# Перевірка, чи запущено скрипт від адміна
function TestAdminRights {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $isAdmin
}

if (-not (TestAdminRights)) {
    Write-Host ">> This script must be run as admin."
    Pause
    Exit
}



# Шляхи до папок
$pathCMS = "C:\Program Files (x86)\CMS\"
$pathPolyvision = "C:\Program Files (x86)\Polyvision\CMS\"

# Перевіряємо наявність програм
$isCMSInstalled = $false
$isPolyvisionInstalled = $false

# Актуальний шлях до актуальної програми
$actualPath;



# Перевіряємо наявність ЦМСа
if (Test-Path $pathCMS -PathType Container) {
    Write-Host ">> Found CMS"
    $isCMSInstalled = $true;
}

# Перевіряємо наявність Полівіжн/ЦМСа
if (Test-Path $pathPolyvision -PathType Container) {
    Write-Host ">> Found Polyvision/CMS"
    $isPolyvisionInstalled = $true;
} 



# Перевіряємо, яка з програм встановлена
if ($isCMSInstalled -and -not $isPolyvisionInstalled) {
    # Якщо встановлений тільки ЦМС, встановлюємо його шлях як актуальний
    $actualPath = $pathCMS
    
} elseif (-not $isCMSInstalled -and $isPolyvisionInstalled) {
    # Якщо встановлений тільки Полівіжн/ЦМС, встановлюємо його шлях як актуальний
    $actualPath = $pathPolyvision
} elseif (-not $isCMSInstalled -and -not $isPolyvisionInstalled) {
    # Якщо не встановлено нічого - виводимо помилку
    Write-Host ">> No CMS or Polyvision/CMS found! Exiting the script."
    Pause
    Exit
} elseif ($isCMSInstalled -and $isPolyvisionInstalled) {
    # Якщо встановлені обидві - виводимо помилку
    Write-Host ">> Both CMS and Polyvision/CMS found! Remove one of them and rerun the script. Exiting the script."
    Pause
    Exit
}



# Перевірка існування bat-файлу перед його створенням
$batFilePath = $actualPath + 'CMS.bat'
if (Test-Path $batFilePath -PathType Leaf) {
    Write-Host ">> Bat-file is already exist. Exiting the script."
    Pause
    Exit
}
 
# Актуальний код для батніка
$code = 'cmd /min /C "set __COMPAT_LAYER=RUNASINVOKER && start "" "' + $actualPath + 'CMS.exe""'



# Створюємо файл
Set-Content -Path $batFilePath -Value $code
Write-Host ">> File created"



# Створення ярлика
function CreateShortcut {
    param (
        [string]$TargetPath,
        [string]$ShortcutPath,
        [string]$IconPath
    )

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.IconLocation = $IconPath
    $Shortcut.Save()
}

# Шлях для збереження ярлика на робочому столі
# $shortcutPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "newCMS.lnk") 
# $shortcutPath = [System.IO.Path]::Combine("$env:USERPROFILE\Desktop", "newCMS.lnk")
$shortcutPath = "C:\Users\kassir\Desktop\newCMS.lnk"


# Шлях до іконки
$iconPath = (Get-Item -Path ".\cms.ico").FullName

# Створення ярлика
CreateShortcut -TargetPath $batFilePath -ShortcutPath $shortcutPath -IconPath $iconPath

Write-Host ">> Shortcut created on the Desktop."



Pause
Exit