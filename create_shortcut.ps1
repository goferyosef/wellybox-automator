$ws = New-Object -ComObject WScript.Shell
$s = $ws.CreateShortcut('C:\Users\gofer\Desktop\WellyBox Downloader.lnk')
$s.TargetPath = 'C:\Users\gofer\AppData\Local\Python\pythoncore-3.14-64\python.exe'
$s.Arguments = '"C:\Users\gofer\Desktop\wellybox-automator\wellybox_app.py"'
$s.WorkingDirectory = 'C:\Users\gofer\Desktop\wellybox-automator'
$s.IconLocation = 'C:\Users\gofer\Desktop\wellybox-automator\wellybox_icon.ico'
$s.Save()
Write-Host 'Shortcut created.'
