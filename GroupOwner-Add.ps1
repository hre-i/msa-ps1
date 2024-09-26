####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($file in $Args) {
    Import-CSV -Path $file -Encoding UTF8 | ForEach-Object {
        AddGroupOwner $_.groupName $_.owner
    }
}

####:FIN:START
EndTS
####:FIN:END
