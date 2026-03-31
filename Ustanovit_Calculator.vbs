Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Калькулятор Столешниц.lnk")
oShellLink.TargetPath = "https://ваш-логин.github.io/название-репозитория/"
oShellLink.WindowStyle = 1
oShellLink.Description = "Расчет стоимости изделий из камня Poverkhnost"
oShellLink.Save
MsgBox "Иконка калькулятора успешно добавлена на ваш рабочий стол!", 64, "Poverkhnost"
