#
# 日時文字列取得
#
function GetTimestamp() {
    return (get-date).ToUniversalTime().AddHours(9).ToString("yyyyMMdd_HHmmss")
}

#
# Transcript 出力開始
#
function StartTS($name) {
    $Transcript = Join-Path $env:NUWS_DATA "log" ($name + "_" + (GetTimestamp) + ".txt")
    Start-Transcript -Path $Transcript
}

#
# Transcript 出力終了
#
function EndTS() {
    Stop-Transcript
}

function DisconnectMSA() {
    Disconnect-ExchangeOnline
    Disconnect-MgGraph
}

#
# 管理下のドメインリスト
#
$ManagedDomainList = @(
	"@8mqvz6.onmicrosoft.com",
	"@niigata-u.ac.jp",
	"@ad.sc.niigata-u.ac.jp",
	"@adm.niigata-u.ac.jp",
	"@agr.niigata-u.ac.jp",
	"@bio.sc.niigata-u.ac.jp",
	"@bri.niigata-u.ac.jp",
	"@cais.niigata-u.ac.jp",
	"@cc.niigata-u.ac.jp",
	"@ccr.niigata-u.ac.jp",
	"@chem.sc.niigata-u.ac.jp",
	"@clg.niigata-u.ac.jp",
	"@create.niigata-u.ac.jp",
	"@dent.niigata-u.ac.jp",
	"@econ.niigata-u.ac.jp",
	"@ed.niigata-u.ac.jp",
	"@emeritus.niigata-u.ac.jp",
	"@eng.niigata-u.ac.jp",
	"@env.sc.niigata-u.ac.jp",
	"@fuchu.ngt.niigata-u.ac.jp",
	"@fusho.ngt.niigata-u.ac.jp",
	"@ge.niigata-u.ac.jp",
	"@geo.sc.niigata-u.ac.jp",
	"@gs.niigata-u.ac.jp",
	"@human.niigata-u.ac.jp",
	"@ie.niigata-u.ac.jp",
	"@isc.niigata-u.ac.jp",
	"@jura.niigata-u.ac.jp",
	"@lib.niigata-u.ac.jp",
	"@m.sc.niigata-u.ac.jp",
	"@med.niigata-u.ac.jp",
	"@mot.niigata-u.ac.jp",
	"@nagaoka.ed.niigata-u.ac.jp",
	"@nuh.niigata-u.ac.jp",
	"@phys.sc.niigata-u.ac.jp",
	"@sake.nu.niigata-u.ac.jp",
	"@sc.niigata-u.ac.jp",
	"@smh.ngt.niigata-u.ac.jp",
	"@test.niigata-u.ac.jp"
)

#
# csv ファイルを xlsx ファイルに変換する
#
function ConvertCsvToXlsx([string]$file) {
    $jar = "$env:NUWS/msa-ps1/lib/nuxlsx.jar"
    if (Test-Path $jar) {
        java -jar $jar to-xlsx -f $file
    }
}

#
# メールアドレスが管理下の度名の場合は true
#
function IsManagedDomain([string]$emailAddr) {
    $emailAddr = $emailAddr.Trim()
    forEach ($d in $ManagedDomainList) {
        if ($emailAddr.EndsWith($d)) {
            return $True
        }
    }
    $False
}

#
# ファイル名からメールアドレス部分を取得
#
function GetMailAddr([string]$path) {
    return (Get-Item $path).BaseName -replace "\.[-_A-Za-z0-9]+$",""
}

#
# 外部メールアドレスの登録
#
function AddGroupExtEmailAddress([string]$emailAddr) {
    $emailAddr = $emailAddr.Trim()
    if (-not (IsManagedDomain($emailAddr))) {
        $name = $emailAddr -replace "@", "=" -replace "\.", "_"
        try {
            Write-Host "Add ExternalEmailAddress:$emailAddr / $name"
            New-MailContact -Name $name -ExternalEmailAddress $emailAddr -Alias $name
            #### 連絡先非表示の設定
            Set-MailContact -Identity $name -HiddenFromAddressListsEnabled $true
        }
        catch {
            Write-Host "ERROR: $($_.Exception.Message)"
        }
    }
}

#### 配布リスト

#
# 配布先アドレスリストの取得
#
function GetGroupMemberListObject($group) {
    try {
        Get-DistributionGroupMember $group.PrimarySmtpAddress -ResultSize Unlimited |
            Select-Object `
                @{ Label = "Group.Mailaddr"; Expression = { $group.PrimarySmtpAddress } },
                @{ Label = "Group.DisplayName"; Expression = { $group.DisplayName } },
                @{ Label = "Member.Mailaddr"; Expression = { $_.PrimarySmtpAddress } },
                @{ Label = "Member.DisplayName"; Expression = { $_.DisplayName } }
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布先アドレスリストの取得
#
function GetGroupMemberList([string]$groupName) {
    return Get-DistributionGroupMember $groupName -ResultSize Unlimited |
        Foreach-Object { $_.PrimarySmtpAddress }
}

#
# 配布リストのメンバー登録
#
function AddGroupMember([string]$groupName, [string]$member) {
    $member = $member.Trim()
    try {
        Write-Host "Add $member to $groupName"
        Add-DistributionGroupMember -Identity $groupName -Member $member
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リストのメンバー削除
#
function RemoveGroupMember([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Remove $member from $groupName"
        Remove-DistributionGroupMember -Identity $groupName -Member $member -Confirm:$false
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 送信許可アドレスリストの取得
#
function GetGroupSenderList([string]$groupName) {
    $group = Get-DistributionGroup $groupName -ResultSize Unlimited
    $group.AcceptMessagesOnlyFrom | ForEach-Object {
        $member = Get-User $_ 2> $null
        if ($member -ne $null) {
            $member | Select-Object  `
                @{Label = "Group.Mailaddr";     Expression = { $group.PrimarySmtpAddress } },
                @{Label = "Group.DisplayName";  Expression = { $group.DisplayName } },
                @{Label = "Sender";             Expression = { $member.UserPrincipalName } },
                @{Label = "Sender.DisplayName"; Expression = { $member.DisplayName } }
        } else {
            $member = (Get-MailContact $_)
            if ($member -ne $null) {
                $member | Select-Object `
                    @{Label = "Group.Mailaddr";     Expression = { $group.PrimarySmtpAddress } },
                    @{Label = "Group.DisplayName";  Expression = { $group.DisplayName } },
                    @{Label = "Sender";             Expression = { $member.PrimarySmtpAddress } },
                    @{Label = "Sender.DisplayName"; Expression = { $member.DisplayName } }
            }
        }
    }
}

#
# 配布リストに送信可能なグループを追加
#
function AddGroupSenderGroup([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Add Sender:$member (AS GROUP) of $groupName"
        Set-DistributionGroup -Identity $groupName -AcceptMessagesOnlyFromDLMembers @{Add = $member}
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リストに送信可能なメンバーを追加
#
function AddGroupSenderMember([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Add Sender:$member (AS USER) of $groupName"
        Set-DistributionGroup -Identity $groupName -AcceptMessagesOnlyFrom @{Add = $member}
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リストに送信可能なメンバーを削除
#
function RemoveGroupSender([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Remove Sender:$member (AS USER) of $groupName"
        Set-DistributionGroup -Identity $groupName -AcceptMessagesOnlyFromSendersOrMembers @{Remove = $member}
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
    try {
        Write-Host "Remove Sender:$member (AS GROUP) of $groupName"
        Set-DistributionGroup -Identity $groupName -AcceptMessagesOnlyFromDLMembers @{Remove = $member}
    } catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リストの SendAs 権限を持つメンバーのリストを取得する
#
function GetGroupSendAsMemberList([string]$groupName) {
    $groupName = $groupName.Trim()
    try {
        Get-RecipientPermission -Identity $groupName |
            Where { $_.AccessRights.contains("SendAs") }
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リスト送信者の SendAs 権限を削除する
#
function AddGroupSendAs([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Add SendAs $member of $groupName"
        Add-RecipientPermission -Identity $groupName -AccessRights SendAs -Trustee $member -Confirm: $False
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リスト送信者の SendAs 権限を削除する
#
function RemoveGroupSendAs([string]$groupName, [string]$member) {
    $groupName = $groupName.Trim()
    $member = $member.Trim()
    try {
        Write-Host "Remove SendAs $member of $groupName"
        Remove-RecipientPermission -Identity $groupName -AccessRights SendAs -Trustee $member -Confirm: $False
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 配布リスト所有者の追加
#
function AddGroupOwner([string]$groupName, [string]$admin) {
    $groupName = $groupName.Trim()
    try {
        Write-Host "Add GroupOwner:$admin of $groupName"
        Set-DistributionGroup -Identity $groupName -ManagedBy @{Add = $admin}
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}


#### 共有メールボックス

#
# 共有メールボックスにユーザの追加
#
function AddSharedMember([string]$shared, [string]$member) {
    $shared = $shared.Trim()
    $member = $member.Trim()
    try {
        Add-MailboxPermission   -Identity $shared -User    $member -AccessRights FullAccess -Confirm: $False
        Add-RecipientPermission -Identity $shared -Trustee $member -AccessRights SendAs     -Confirm: $False
    }
    catch {
        Write-Host "ERROR: $($_.Exception.Message)"
    }
}

#
# 共有メールボックスの設定を変更
#
function SetSharedConfig([string]$shared) {
    $shared = $shared.Trim()
    try {
        Set-MailboxRegionalConfiguration -Identity $shared `
            -Language "ja-JP" `
            -DateFormat "yyyy/MM/dd" `
            -TimeFormat "H:mm" `
            -TimeZone "Tokyo Standard Time" `
            -LocalizeDefaultFolderName
    }
    catch {
        Write-Host "ERROR: $( $_.Exception.Message )"
    }
}
