####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($file in $Args) {
    $sharedName = GetMailAddr($file)
    Get-Content -Path $file -Encoding UTF8 | Foreach-object {
        AddSharedMember $sharedName $_
    }
}

####:FIN:START
EndTS
####:FIN:END
