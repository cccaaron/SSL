Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objIE = CreateObject("InternetExplorer.Application")

With objIE
    .Navigate "about:blank"
    .Visible = True
    .Width = 500 : .Height = 600
    .Toolbar = False : .Statusbar = False
End With

Do While objIE.Busy: WScript.Sleep 100: Loop

objIE.Document.Title = "PFX Generator - Paste Certificate"
objIE.Document.Body.InnerHTML = _
    "<div style='font-family:sans-serif;font-size:12px;padding:15px;'>" & _
    "<h3>Step 2: Create PFX</h3>" & _
    "Paste your Certificate (CRT) content here:<br>" & _
    "<textarea id='crtData' style='width:100%;height:300px;font-family:monospace;' " & _
    "placeholder='-----BEGIN CERTIFICATE----- ...'></textarea><br><br>" & _
    "Enter a password for the PFX file:<br>" & _
    "<input type='password' id='pwd' style='width:100%'><br><br>" & _
    "<button id='btn' style='width:100%;height:30px;background-color:#008CBA;color:white;border:none;'>Generate PFX</button>" & _
    "<input type='hidden' id='clicked' value='0'></div>"

Set btn = objIE.Document.GetElementById("btn")
Set clicked = objIE.Document.GetElementById("clicked")
btn.onclick = GetRef("OnBtnClick")
Sub OnBtnClick: clicked.Value = "1": End Sub

Do While clicked.Value = "0": WScript.Sleep 100: Loop

strCRT = objIE.Document.GetElementById("crtData").Value
strPwd = objIE.Document.GetElementById("pwd").Value
objIE.Quit

If Trim(strCRT) <> "" Then
    Set objFile = objFSO.CreateTextFile("certificate.crt", True)
    objFile.Write strCRT
    objFile.Close
    
    Set objFilePwd = objFSO.CreateTextFile("pwd.txt", True)
    objFilePwd.Write strPwd
    objFilePwd.Close
    
    WScript.Quit(0)
Else
    WScript.Quit(1)
End If