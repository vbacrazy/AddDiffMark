'On Error Resume Next

Dim str_to

Const wdStartupPath = 8
Const wdUserTemplatesPath = 2

Set objWord = CreateObject("Word.Application")
Set objOptions = objWord.Options

str_to = objOptions.DefaultFilePath(wdStartupPath)

objWord.Quit

Set sa = CreateObject("Shell.Application")
sa.Open str_to

