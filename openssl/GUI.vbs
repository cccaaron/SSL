Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim C, ST, L, O, OU, CN, Bits, SAN
C    = "IL"
ST   = ""
L    = ""
O    = ""
OU   = ""
CN   = ""
Bits = "2048"
SAN  = ""

If WScript.Arguments.Count >= 1 Then C    = Trim(WScript.Arguments(0))
If WScript.Arguments.Count >= 2 Then ST   = Trim(WScript.Arguments(1))
If WScript.Arguments.Count >= 3 Then L    = Trim(WScript.Arguments(2))
If WScript.Arguments.Count >= 4 Then O    = Trim(WScript.Arguments(3))
If WScript.Arguments.Count >= 5 Then OU   = Trim(WScript.Arguments(4))
If WScript.Arguments.Count >= 6 Then CN   = Trim(WScript.Arguments(5))
If WScript.Arguments.Count >= 7 Then Bits = Trim(WScript.Arguments(6))
If WScript.Arguments.Count >= 8 Then SAN  = Trim(WScript.Arguments(7))

If Bits <> "4096" Then Bits = "2048"

Set objIE = CreateObject("InternetExplorer.Application")
With objIE
    .Navigate "about:blank"
    .Visible = True
    .Width = 450 : .Height = 650
    .Toolbar = False : .Statusbar = False
End With

Do While objIE.Busy: WScript.Sleep 100: Loop

Dim sel2048, sel4096
If Bits = "4096" Then sel4096 = "selected" Else sel2048 = "selected"

objIE.Document.Title = "CSR Data Entry"
objIE.Document.Body.InnerHTML = _
    "<div style='font-family:sans-serif;font-size:12px;padding:15px;'>" & _
    "<h3>Enter CSR Details</h3>" & _
    "Country:<br><input type='text' id='C' value='" & C & "' style='width:100%'><br>" & _
    "State:<br><input type='text' id='ST' value='" & ST & "' style='width:100%'><br>" & _
    "Locality:<br><input type='text' id='L' value='" & L & "' style='width:100%'><br>" & _
    "Organization:<br><input type='text' id='O' value='" & O & "' style='width:100%'><br>" & _
    "Org Unit:<br><input type='text' id='OU' value='" & OU & "' style='width:100%'><br>" & _
    "Common Name:<br><input type='text' id='CN' value='" & CN & "' style='width:100%'><br>" & _
    "SAN (comma separated):<br><input type='text' id='SAN' value='" & SAN & "' style='width:100%'><br><br>" & _
    "Key Size:<br><select id='bits' style='width:100%'><option value='2048' " & sel2048 & ">2048</option><option value='4096' " & sel4096 & ">4096</option></select><br><br>" & _
    "<button id='btn' style='width:100%;height:30px;'>Create Config & Return to Batch</button>" & _
    "<input type='hidden' id='clicked' value='0'></div>"

Set btn = objIE.Document.GetElementById("btn")
Set clicked = objIE.Document.GetElementById("clicked")
btn.onclick = GetRef("OnBtnClick")
Sub OnBtnClick: clicked.Value = "1": End Sub

Do While clicked.Value = "0": WScript.Sleep 100: Loop

C    = objIE.Document.GetElementById("C").Value
ST   = objIE.Document.GetElementById("ST").Value
L    = objIE.Document.GetElementById("L").Value
O    = objIE.Document.GetElementById("O").Value
OU   = objIE.Document.GetElementById("OU").Value
CN   = objIE.Document.GetElementById("CN").Value
SAN  = objIE.Document.GetElementById("SAN").Value
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
