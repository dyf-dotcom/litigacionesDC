Dim icoUrl, icoDir, icoPath, acceso, url

Dim wsh : Set wsh = CreateObject("WScript.Shell")

url     = "https://dyf-dotcom.github.io/litigacionesDC/app/login.html"
icoUrl  = "https://raw.githubusercontent.com/dyf-dotcom/litigacionesDC/main/app/logo_dc.ico"
icoDir  = wsh.ExpandEnvironmentStrings("%APPDATA%") & "\DCLitigaciones"
icoPath = icoDir & "\logo_dc.ico"
acceso  = wsh.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop\DC-Litigaciones.lnk"

' Crear carpeta si no existe
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(icoDir) Then fso.CreateFolder(icoDir)

' Descargar el ícono
On Error Resume Next
Dim http
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", icoUrl, False
http.Send

If Err.Number = 0 And http.Status = 200 Then
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binario
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile icoPath, 2 ' sobreescribir
    stream.Close
End If
On Error GoTo 0

' Crear acceso directo
Dim shell, shortcut
Set shell    = wsh
Set shortcut = shell.CreateShortcut(acceso)
shortcut.TargetPath  = url
shortcut.Description = "DC-Litigaciones — Defensoría de Casación MPBA"
If fso.FileExists(icoPath) Then shortcut.IconLocation = icoPath & ",0"
shortcut.Save

MsgBox "Acceso directo creado en el Escritorio.", vbInformation, "DC-Litigaciones"
