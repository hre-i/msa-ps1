####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($sharedName in $Args) {
    if (Test-Path -Path $sharedName -PathType Leaf) {
        $sharedName = GetMailaddr($sharedName)
    }
    $outputFile = ($sharedName + ".member-out.csv")
    try {
        $output = Get-MailboxPermission -Identity $sharedName |
                Where { ($_.User -Notlike "*S-1-5-21*") -And ($_.User -Notlike "*\*") } |
                Select-Object User |
                Foreach-Object { $_.User } |
                Out-File $outputFile
    }
    catch {
        Write-Host "【ERROR】Shared:$( $sharedName ) $( $_.Exception.Message )"
    }
}

####:FIN:START
EndTS
####:FIN:END
