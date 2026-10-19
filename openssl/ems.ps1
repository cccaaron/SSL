param (
    [string]$AttachmentPath,
	[string]$NewSubject
)

try {    
    $pwdurl = (Invoke-RestMethod -Uri "https://pwpush.com/p.json" -Method Post -Body @{ "password[payload]" = (Get-Content -Path "pwd.txt" -Raw); "password[expire_after_days]" = 14; "password[expire_after_views]" = 1 }).html_url
    
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)

    if ($NewSubject) {
        $Mail.Subject = $NewSubject
    }
    $Mail.SentOnBehalfOfName = "support@ccc.co.il"
	
	if ($AttachmentPath) {
        $FullAttachPath = [System.IO.Path]::GetFullPath($AttachmentPath.Trim('"') + "\")
        $Mail.Attachments.Add("$FullAttachPath\certificate_PBE-SHA1-3DES.pfx")
        $Mail.Attachments.Add("$FullAttachPath\certificate.pfx")
    }
	
    $Mail.HTMLBody = @"
<html>
<head>
<style>
    body { font-family: Arial, sans-serif; font-size: 14px; color: #333333; }
    p.tight { margin: 0 0 5px 0; padding: 0; line-height: 1.3; }
</style>
</head>
<body lang=EN-US link="#0563C1" vlink="#954F72" style='tab-interval:36.0pt; word-wrap:break-word'>

<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>שלום רב,</span></p>
<br>
<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>מצ&quot;ב קבצי </span><span dir=LTR>PFX</span> <span lang=HE>להתקנה בשרת</span></p>

<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>שרתי </span><span dir=LTR>Windows Server 2019</span> <span lang=HE>ומעלה: יש להשתמש בקובץ </span><span dir=LTR>certificate.pfx</span></p>

<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>שרתי </span><span dir=LTR>Windows Server 2016</span> <span lang=HE>ומטה: יש להשתמש בקובץ </span><span dir=LTR>certificate_PBE-SHA1-3DES.pfx</span></p>
<br>
<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>סיסמה - <a href="$pwdurl" target="_blank">לצפיה חד פעמית</a></span></p>
<br>
<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<span lang=HE>שמחנו לעמוד לשירותך.</span></p>
<br>
<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<b><span lang=HE style='color:blue;'>בברכה, </span></b></p>

<p class="tight" dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>
<b><span lang=HE style='color:blue;'>צוות ענן טריפל סי</span></b></p>

</body>
</html>
"@
	
    $OutputPath = Join-Path (Get-Location) "sendmail.msg"
    $Mail.SaveAs($OutputPath, 3)
 
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null }
    if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
}