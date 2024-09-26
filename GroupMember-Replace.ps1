####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 配布リストメンバーを入れ替える

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    Write-Host "DistributionGroup: $groupName"

    #### 現メンバーリストを取得
    $oldMemberList = GetGroupMemberList $groupName

    #### 新メンバ
    $newMemberList = Get-Content $file -Encoding UTF8 | ForEach-Object { $_.Trim() }

    $addMemberList = $newMemberList | Where-Object { $_ -notin $oldMemberList }
    $delMemberList = $oldMemberList | Where-Object { $_ -notin $newMemberList }

    #### 外部アドレスの登録
    $addMemberList | ForEach-Object { AddGroupExtEmailAddress $_ }

    #### メンバー削除
    $delMemberList | ForEach-Object { RemoveGroupMember $groupName $_ }
    #### メンバー追加
    $addMemberList | ForEach-Object { AddGroupMember $groupName $_ }
}

####:FIN:START
EndTS
####:FIN:END
