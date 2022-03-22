Set WshShell = CreateObject("WScript.Shell")
Set ScriptFso = CreateObject("Scripting.FileSystemObject")
PriviledgesRequired
ProcessTemplate "shuffler-lib-template.wsc", "shuffler-lib.wsc", Array("___CLSID___"),_
Array(Left(CreateObject("Scriptlet.TypeLib").Guid, 38)), "RegSvr32 scrobj.dll /n /i:___OUTPUTPATH___ /s"

Private Sub PriviledgesRequired
    On Error Resume Next
    Err.Clear
    WshShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    If Err.Number <> 0 Then
        WScript.Echo "Elevated priviledges required."
        WScript.Quit(2)
    End If
End Sub

Private Sub ProcessTemplate(TemplateName, OutputName, arrTextTemplate, arrReplacement, CommandTemplate)
    OutputPath = BuildPath(OutputName)
    With ScriptFso.OpenTextFile(BuildPath(TemplateName), 1)
        TemplateText = .ReadAll()
        .Close()
    End With
    For i = 0 To UBound(arrTextTemplate)
        TemplateText = Replace(TemplateText, arrTextTemplate(i), arrReplacement(i))
    Next
    With ScriptFso.OpenTextFile(OutputPath, 2, True)
        .Write(TemplateText)
        .Close()
    End With
    If IsNull(CommandTemplate) Then Exit Sub
    With WshShell.Exec(Replace(CommandTemplate, "___OUTPUTPATH___", OutputPath))
        Do : Loop Until .Status = 1
    End With
End Sub

Private Function BuildPath(FileName)
    With ScriptFso
        BuildPath = .BuildPath(.GetParentFolderName(WScript.ScriptFullName), FileName)
    End With
End Function