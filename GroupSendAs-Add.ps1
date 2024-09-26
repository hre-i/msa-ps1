####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    Write-Host "DistributionGroup: $groupName"

    #### 配布リストアドレスから送信可能
    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupSendAs $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
