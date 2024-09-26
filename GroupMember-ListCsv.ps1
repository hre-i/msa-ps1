####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($groupName in $Args) {
    if (Test-Path -Path $groupName -PathType Leaf) {
        $groupName = GetMailaddr($groupName)
    }
    Write-Host $groupName

    try {
        $group = Get-DistributionGroup $groupName -RecipientTypeDetails MailUniversalDistributionGroup |
                Select-Object DisplayName,PrimarySmtpAddress

        $outputFile = ($groupName + ".out.csv")
        $memberList = Get-DistributionGroupMember $groupName -ResultSize Unlimited
        if ($null -eq $memberList) {
            Write-Host "No Member"
        } else {
            GetGroupMemberListObject $group | Export-Csv -Encoding UTF8 $outputFile
            ConvertCsvToXlsx $outputFile
        }
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

####:FIN:START
EndTS
####:FIN:END
