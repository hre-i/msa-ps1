####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 配布リストメンバー追加

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    Write-Host "DistributionGroup: $groupName"

    #### 外部アドレスの登録
    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupExtEmailAddress $_
    }

    #### 配布先アドレス追加
    Get-Content $file -Encoding UTF8 | ForEach-Object {
        AddGroupMember $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
