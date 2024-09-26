####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($file in $Args) {
    $groupName = GetMailaddr($file)
    Write-Host "DistributionGroup: $groupName"

    Get-Content -Path $file -Encoding UTF8 | Foreach-Object {
        RemoveGroupMember $groupName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
