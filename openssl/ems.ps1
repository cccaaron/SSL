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
    body { font-family: Arial, sans-serif; font-size: 14px; color: #333333; direction: rtl; text-align: right; }
    .line { margin-bottom: 5px; }
</style>
</head>
<body lang=HE>
<div class="line">שלום רב,</div><br>
<div class="line">מצ&quot;ב קבצי <span dir=LTR>PFX</span> להתקנה בשרת</div>
<div class="line">שרתי <span dir=LTR>Windows Server 2019</span> ומעלה: יש להשתמש בקובץ <span dir=LTR>certificate.pfx</span></div>
<div class="line">שרתי <span dir=LTR>Windows Server 2016</span> ומטה: יש להשתמש בקובץ <span dir=LTR>certificate_PBE-SHA1-3DES.pfx</span></div><br>
<div class="line">סיסמה - <a href="$pwdurl" target="_blank">לצפיה חד פעמית</a></div><br>
<div class="line">שמחנו לעמוד לשירותך.</div><br>
<div class="line"><b style='color:blue;'>בברכה,</b></div>
<div class="line"><b style='color:blue;'>צוות ענן טריפל סי</b></div>
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
