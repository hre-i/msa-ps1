####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 配布リストを作成する

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    try {
        Write-Host "Create: $groupName"
        New-DistributionGroup -DisplayName $groupName `
            -Name $groupName `
            -PrimarySmtpAddress $groupName `
            -RequireSenderAuthenticationEnabled $false

            # 別途所有者を登録する場合（指定しない場合は、PowerShell実行ユーザーが自動的に指定されます）
            # Set-DistributionGroup -Identity $_.PrimarySmtpAddress -ManagedBy  $_.Admin
    }
    catch {
        Write-Host "ERROR: Message:$($_.Exception.Message)"
    }

    #### 外部アドレスの登録
    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupExtEmailAddress $_
    }

    #### 配布先アドレス追加
    Get-Content -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupMember $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
