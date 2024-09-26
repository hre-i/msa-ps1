####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($groupName in $Args) {
    if (Test-Path -Path $groupName -PathType Leaf) {
        $groupName = GetMailaddr($groupName)
    }
    Write-Host $groupName

    $outputFile = ($groupName + ".member-out.txt")
    Write-Host "> $outputFile"
    GetGroupMemberList $groupName | Out-File $outputFile
}

####:FIN:START
EndTS
####:FIN:END
