param (
    [Parameter(Mandatory=$true)]
    [string]$MsgPath,
    [Parameter(Mandatory=$false)]
    [string]$AttachmentPath,
    [Parameter(Mandatory=$false)]
    [string]$NewSubject
)

# הגדרת קידוד למסוף
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ניקוי גרשיים מיותרים והמרה לנתיב מלא
$FullMsgPath = [System.IO.Path]::GetFullPath($MsgPath.Trim('"'))

try {
    $Outlook = New-Object -ComObject Outlook.Application
    
    if (-not (Test-Path $FullMsgPath)) {
        throw "קובץ ה-MSG לא נמצא בנתיב: $FullMsgPath"
    }

    $Mail = $Outlook.CreateItemFromTemplate($FullMsgPath)

    if ($NewSubject) {
        $Mail.Subject = $NewSubject
    }

    if ($AttachmentPath) {
        $FullAttachPath = [System.IO.Path]::GetFullPath($AttachmentPath.Trim('"') + "\")
        $Mail.Attachments.Add("$FullAttachPath\certificate_PBE-SHA1-3DES.pfx")
        $Mail.Attachments.Add("$FullAttachPath\certificate.pfx")    
    }

    $Mail.SaveAs($FullMsgPath)
    Write-Host "Success: Mail saved at $FullMsgPath" -ForegroundColor Green
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Mail) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null }
    if ($Outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null }
}