####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($groupName in $Args) {
    Write-Host "DistributionGroup SenderMemberOnly: $groupName"
    AddGroupSenderGroup $groupName $groupName
}

####:FIN:START
EndTS
####:FIN:END
