####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 送信可能メールアドレスの追加

foreach ($file in $Args) {
    $groupName = GetMailAddr($file)
    Write-Host "DistributionGroup: $groupName"

    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupSenderGroup $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
