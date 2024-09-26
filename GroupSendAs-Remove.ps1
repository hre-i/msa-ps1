####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 配布リストアドレスから送信可能

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    Write-Host "DistributionGroup: $groupName"

    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        RemoveGroupSendAs $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
