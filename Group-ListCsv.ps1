####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

####
#### すべての配布リストの配布先リストを取得する
####

$outputFile = ("group-list-all_" + (GetTimestamp) + ".out.csv")

Get-DistributionGroup -ResultSize Unlimited -RecipientTypeDetails MailUniversalDistributionGroup |
    Select-Object DisplayName,PrimarySmtpAddress |
    ForEach-Object {
        $group = $_
        Write-Host $group.PrimarySmtpAddress

        GetGroupMemberListObject $group
    } | Export-Csv -Encoding UTF8 $outputFile

ConvertCsvToXlsx $outputFile

####:FIN:START
EndTS
####:FIN:END
