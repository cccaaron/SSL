Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objIE = CreateObject("InternetExplorer.Application")

With objIE
    .Navigate "about:blank"
    .Visible = True
    .Width = 500 : .Height = 550
    .Toolbar = False : .Statusbar = False
End With

Do While objIE.Busy: WScript.Sleep 100: Loop

objIE.Document.Title = "CSR Renewal Wizard - Step 1"
objIE.Document.Body.InnerHTML = _
    "<div style='font-family:sans-serif;font-size:12px;padding:15px;'>" & _
    "<h3>Step 1: Paste Old CSR (Optional)</h3>" & _
    "<p>Paste the old CSR below to automatically populate the fields for the new CSR, then click Next.<br>" & _
    "If you don't have one, just leave it empty and click Next to start fresh.</p>" & _
    "<textarea id='csrBox' style='width:100%;height:300px;font-family:monospace;font-size:11px;' placeholder='-----BEGIN CERTIFICATE REQUEST-----\n...\n-----END CERTIFICATE REQUEST-----'></textarea><br><br>" & _
    "<button id='btnNext' style='width:100%;height:35px;font-weight:bold;'>Next &rarr;</button>" & _
    "<input type='hidden' id='clicked' value='0'></div>"

Set btnNext = objIE.Document.GetElementById("btnNext")
Set clicked = objIE.Document.GetElementById("clicked")

btnNext.onclick = GetRef("OnNextClick")
Sub OnNextClick: clicked.Value = "1": End Sub

' המתנה ללחיצה על Next
Do While clicked.Value = "0": WScript.Sleep 100: Loop

Dim csrText
csrText = Trim(objIE.Document.GetElementById("csrBox").Value)

objIE.Quit

' יצירת הקובץ רק אם הוכנס ערך
If csrText <> "" Then
    Set objFile = objFSO.CreateTextFile("old_CSR.txt", True)
    objFile.Write csrText
    objFile.Close
End If