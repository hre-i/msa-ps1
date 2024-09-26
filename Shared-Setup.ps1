####:INIT:START
. "$env:NUWS\msa-ps1\lib\MSALib.ps1"
StartTS((Get-Item $PSCommandPath).BaseName)
####:INIT:END

foreach ($shared in $Args) {
    SetSharedConfig $shared
}

####:FIN:START
EndTS
####:FIN:END
