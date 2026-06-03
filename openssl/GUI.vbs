Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objIE = CreateObject("InternetExplorer.Application")

' הגדרות חלון
With objIE
    .Navigate "about:blank"
    .Visible = True
    .Width = 450 : .Height = 650
    .Toolbar = False : .Statusbar = False
End With

Do While objIE.Busy: WScript.Sleep 100: Loop

objIE.Document.Title = "CSR Data Entry"
objIE.Document.Body.InnerHTML = _
    "<div style='font-family:sans-serif;font-size:12px;padding:15px;'>" & _
    "<h3>Enter CSR Details</h3>" & _
    "Country:<br><input type='text' id='C' value='IL' style='width:100%'><br>" & _
    "State:<br><input type='text' id='ST' style='width:100%'><br>" & _
    "Locality:<br><input type='text' id='L' style='width:100%'><br>" & _
    "Organization:<br><input type='text' id='O' style='width:100%'><br>" & _
    "Org Unit:<br><input type='text' id='OU' style='width:100%'><br>" & _
    "Common Name:<br><input type='text' id='CN' style='width:100%'><br>" & _
    "SAN (comma separated):<br><input type='text' id='SAN' style='width:100%'><br><br>" & _
    "Key Size:<br><select id='bits' style='width:100%'><option value='2048'>2048</option><option value='4096'>4096</option></select><br><br>" & _
    "<button id='btn' style='width:100%;height:30px;'>Create Config & Return to Batch</button>" & _
    "<input type='hidden' id='clicked' value='0'></div>"

Set btn = objIE.Document.GetElementById("btn")
Set clicked = objIE.Document.GetElementById("clicked")
btn.onclick = GetRef("OnBtnClick")
Sub OnBtnClick: clicked.Value = "1": End Sub

Do While clicked.Value = "0": WScript.Sleep 100: Loop

C = objIE.Document.GetElementById("C").Value
ST = objIE.Document.GetElementById("ST").Value
L = objIE.Document.GetElementById("L").Value
O = objIE.Document.GetElementById("O").Value
OU = objIE.Document.GetElementById("OU").Value
CN = objIE.Document.GetElementById("CN").Value
SAN = objIE.Document.GetElementById("SAN").Value
Bits = objIE.Document.GetElementById("bits").Value

objIE.Quit

strConf = "[req]" & vbCrLf & _
          "distinguished_name = req_distinguished_name" & vbCrLf & _
          "req_extensions = v3_req" & vbCrLf & _
          "prompt = no" & vbCrLf & vbCrLf & _
          "[req_distinguished_name]" & vbCrLf & _
          "C = " & C & vbCrLf & _
          "ST = " & ST & vbCrLf & _
          "L = " & L & vbCrLf & _
          "O = " & O & vbCrLf & _
          "OU = " & OU & vbCrLf & _
          "CN = " & CN & vbCrLf & vbCrLf & _
          "[v3_req]" & vbCrLf & _
          "subjectAltName = @alt_names" & vbCrLf & vbCrLf & _
          "[alt_names]" & vbCrLf & _
          "DNS.1 = " & CN

If SAN <> "" Then
    arrSAN = Split(SAN, ",")
    For i = 0 To UBound(arrSAN)
        strConf = strConf & vbCrLf & "DNS." & (i + 2) & " = " & Trim(arrSAN(i))
    Next
End If

Set objFile = objFSO.CreateTextFile("temp_openssl.conf", True)
objFile.Write strConf
objFile.Close

WScript.Quit(CInt(Bits))