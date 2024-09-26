####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($groupName in $Args) {
    $outputFile = ($groupName + ".sender.out.csv")
    Write-Host "Output $outputFile"

    GetGroupSenderList $groupName |
        Export-Csv -Encoding UTF8 $outputFile

    ConvertCsvToXlsx $outputFile
}

####:FIN:START
EndTS
####:FIN:END
