####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

<#

```
  Usage: Shared-ListCsv.ps1
```

- すべての共有メールボックスのメンバーリストを出力する。
- 「`shared-list-all_yyyyMMdd_HHmmss.csv`」が作成される。各列は次の通り
  - Shared.Mailaddr	... 共有メールボックスのメールアドレス
  - Shared.DisplayName ... 共有メールボックスの表示名
  - User ... 利用者のMS365アカウント
  - AccessRights ... 利用者のアクセス権

#>

$outputFile = ("shared-list-all_" + (GetTimestamp) + ".out.csv")

try {
    Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox |
        Select-Object Identity,DisplayName,PrimarySmtpAddress |
        ForEach-Object {
            $smbox = $_
            Write-Host $smbox.PrimarySmtpAddress
            Get-MailboxPermission $_.Identity -ResultSize Unlimited |
                Where-Object {($_.User -Notlike "*S-1-5-21*") -And ($_.User -Notlike "*\*")} |
                Select-Object `
                    @{Label = "Shared.Mailaddr"; Expression = { $smbox.PrimarySmtpAddress }},
                    @{Label = "Shared.DisplayName"; Expression = { $smbox.DisplayName }},
                    User,AccessRights
        } |
        Export-Csv -Encoding UTF8 $outputFile
}
catch {
    Write-Host "ERROR: ALl-Group-Member-List: $($_.Exception.Message)"
}

####:FIN:START
EndTS
####:FIN:END
