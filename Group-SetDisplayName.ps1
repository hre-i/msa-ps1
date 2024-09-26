####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

#### 配布リストの表示名を設定する

foreach ($file in $Args) {
    Import-CSV -Path $file -Encoding UTF8 | ForEach-Object {
        try {
            Write-Host "Set DisplayName:$($_.displayName) Group:$($_.groupName)"
            Set-DistributionGroup $_.groupName -DisplayName $_.displayName
        } catch {
            Write-Host "ERROR: $($_.Exception.Message)"
        }
    }
}

####:FIN:START
EndTS
####:FIN:END
