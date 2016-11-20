on error resume next
mDrive = "j:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "9-12 SMART Board Team Content"
mDrive = "k:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "6-8 SMART Board Team Content"
mDrive = "l:\"
Set oShell = CreateObject("Shell.Application")
oShell.NameSpace(mDrive).Self.Name = "K-5 SMART Board Team Content"