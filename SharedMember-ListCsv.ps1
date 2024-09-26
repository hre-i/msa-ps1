####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($sharedName in $Args) {
    if (Test-Path -Path $sharedName -PathType Leaf) {
        $sharedName = GetMailaddr($sharedName)
    }
    $outputFile = ($sharedName + ".out.csv")
    try {
        Get-Mailbox $sharedName -RecipientTypeDetails SharedMailbox |
            Select-Object Identity,DisplayName,PrimarySmtpAddress |
            ForEach-Object {
                $smbox = $_
                Write-Host $smbox.PrimarySmtpAddress
                Get-MailboxPermission $_.Identity -ResultSize Unlimited |
                    Where { ($_.User -Notlike "*S-1-5-21*") -And ($_.User -Notlike "*\*") } |
                    Select-Object `
                        @{ Label = "Shared.Mailaddr"; Expression = { $smbox.PrimarySmtpAddress } },
                        @{ Label = "Shared.DisplayName"; Expression = { $smbox.DisplayName } },
                        User,AccessRights  |
                    Export-Csv -Encoding UTF8 $outputFile
                ConvertCsvToXlsx $outputFile
            }
    }
    catch {
        Write-Host "ERROR: $( $_.Exception.Message )"
    }
}

####:FIN:START
EndTS
####:FIN:END
