Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

Do
    ' Read all lines from texts.txt
    Set file = fso.OpenTextFile("texts.txt", 1)
    lines = Split(file.ReadAll, vbNewLine)
    file.Close

    ' Pick a random line
    Randomize
    randomIndex = Int(Rnd * (UBound(lines) + 1))

    ' Popup settings:
    ' 16 = Error icon
    ' 4096 = Always on top
    ' 5 = OK + Retry buttons
    result = sh.Popup(lines(randomIndex), 0, "Error", 16 + 4096 + 5)

    ' If user clicked Retry (value = 4), restart the script
    If result = 4 Then
        sh.Run WScript.ScriptFullName
        WScript.Quit
    End If

Loop
