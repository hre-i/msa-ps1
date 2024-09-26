. "$env:NUWS\msa-ps1\lib\MSALib.ps1"

# 現在のプロセスに限り、一時的に PowerShell 実行ポリシーを緩和します。
#Set-ExecutionPolicy Unrestricted -Scope Process -Force

# PSGallery への接続を明示的に許可します
Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

# TLS バージョンを明示的に指定します
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#テナントに接続（全体管理者のIDでサインインします）
Connect-ExchangeOnline
Connect-MgGraph -Scopes "User.ReadWrite.All"
