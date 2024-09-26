msa-ps1
=======

Microsoft365 用スクリプト


準備
----

- 最新の PowerShell をインストールする
- jdk21 をインストールする
- ディレクトリ nuws を作成し、そのパスを環境変数 NUWS に設定する
- ${NUWS} 以下に次のようなディレクトリを作成する
  - ${NUWS}/msa-ps1
  - ${NUWS}/log
  - ${NUWS}/bin
- 環境変数 PATH に ${NUWS}/msa-ps1 を追加する
- このディレクトリのファイルをすべて、${NUWS}/msa-ps1 以下にコピーする
- pwsh を起動し、必要なモジュールをインストールする。
  ```
  Install-Module -Name ExchangeOnlineManagement
  Install-Module Microsoft.Graph
  ```


使用方法
--------

- pwsh を起動し、次のコマンドを実行
  ```
  Connect.ps1
  ```
  ブラウザ経由で認証を求められるので、管理者権限でサインインする。
- pwsh を終了しないで、作業する。


配布グループ
------------

引数に関する共通仕様
- GROUP-ADDRESS ... 配布リストのプライマリメールアドレス
- GROUP-ADDRESS.EXT.txt ... 配布リストのメンバーリストを変更するためのデータファイル。
- - GROUP-ADDRESS は操作対象の配布リストのアドレスと同じにすること。
  - 1行-1メールアドレスのテキストファイル。
  - EXT は、`\.[-_A-Za-z0-9]+` にマッチする文字列を使うこと。
    Usage に書かれているものと同じでなくてもかまわない。

### Group-ListCsv.ps1

```
  Usage: Group-ListCsv.ps1
```

- すべての配布リストの配布先アドレスを出力する。
- 「group-list-all.yyyyMMdd_HHmmss.csv」が作成される。各列は次の通り
  - Group.Mailaddr	... 配布リストのメールアドレス
  - Group.DisplayName ... 配布リストの表示名
  - Member.Mailaddr ... 配布先メールアドレス
  - Member.DisplayName ... 配布先メールアドレスの表示名

### Group-Create.ps1

```
  Usage: Group-Create.ps1 GROUP-ADDRESS.member.txt
```

- 配布リストを新規作成する
- 配布リストのメールアドレス、名前、表示名はすべて GROUP-ADDRESS になる
- GROUP-ADDRESS.member.txt は配布先リスト

### GroupMember-List.ps1

```
  Usage: GroupMember-List.ps1 GROUP-ADDRESS ...
```

- 配布リストの配布先アドレスリストを取得する
- GROUP-ADDRESS は空白で区切って複数指定可能
- 「配布リストのメールアドレス.member.txt」がカレントディレクトリに出力される

### GroupMember-ListCsv.ps1

```
  Usage: GroupMember-ListCsv.ps1 GROUP-ADDRESS ...
```

- 配布先アドレスを出力する。
- `GROUP-ADDRESS.out.csv` という名称のファイルが作成される。各列は次の通り。
  - Group.Mailaddr	... 配布リストのメールアドレス
  - Group.DisplayName ... 配布リストの表示名
  - Member.Mailaddr ... 配布先メールアドレス
  - Member.DisplayName ... 配布先メールアドレスの表示名

### GroupMember-Add.ps1

```
  Usage: GroupMember-Add.ps1 GROUP-ADDRESS.add.txt ...
```

- GROUP-ADDRESS.add.txt は追加するメールアドレスのリスト
- 外部のメールアドレスも登録する。既に登録してある場合はエラーメッセージが表示されるが無視してよい。

### GroupMember-Remove.ps1

```
  Usage: GroupMember-Remove.ps1 GROUP-ADDRESS.remove.txt ...
```

- GROUP-ADDRESS.remove.txt は削除するメールアドレスのリスト

### GroupMember-Replace.ps1

```
  Usage: GroupMember-Replace.ps1 GROUP-ADDRESS.member.txt
```

- 登録されているメンバーリストを入れ替える。
- 現在の配布先を取得して、GROUP-ADDRESS.member.txt にないものを削除してから、
  GROUP-ADDRESS.member.txt のみにあるものを追加する。

### GroupSendAs-Add.ps1

```
  Usage: GroupSendAs-Add.ps1 GROUP-ADDRESS.add-send-as.txt ...
```

- GROUP-ADDRESS.add-send-as.txt は SendAs 権限を追加するメールアドレスのリスト
- 外部のメールアドレスは登録できない

### GroupSendAs-Remove.ps1

```
  Usage: GroupSendAs-Remove.ps1 GROUP-ADDRESS.remove-send-as.txt ...
```

- GROUP-ADDRESS.remove-send-as.txt は SendAs 権限を削除するメールアドレスのリスト

### Group-SetSendMemberOnly.ps1

```
  Usage: Group-SetSendMemberOnly.ps1 GROUP-ADDRESS
```

- 配布リストにメールを送信できるメールアドレスに配布リスト自身のアドレスを追加する

### GroupSender-AddMember.ps1

```
  Usage: GroupSender-AddMember.ps1 GROUP-ADDRESS.sender-add.txt
```

- 配布リストにメールを送信できるメールアドレスを追加する
- メールアドレスが個人アドレスの時にこちらを使う

### GroupSender-AddGroup.ps1

```
  Usage: GroupSender-AddGroup.ps1 GROUP-ADDRESS.sender-add.txt
```

- 配布リストにメールを送信できるメールアドレスを追加する
- メールアドレスが配布リストのアドレスの時にこちらを使う


### GroupSender-Remove.ps1

```
  Usage: GroupSender-Remove.ps1 GROUP-ADDRESS.sender-remove.txt
```

- 配布リストにメールを送信できるメールアドレスから削除する

### Group-SetDisplayName.ps1

```
  Usage: Group-SetDisplayName.ps1 GroupName_Displayname.csv
```

- 配布リストの DisplayName を一括して変更する
- GroupName_DisplayName.csv のカラム名:
  - groupName ... 配布リストのアドレス
  - displayName ... 表示名

  - 例）
    ```
    groupName,displayName
    test.group01@niigata-u.ac.jp,テスト用配布リスト01
    test.group02@niigata-u.ac.jp,テスト用配布リスト02
    ```

### GroupOwner-Add.ps1

```
  Usage: GroupOwner-Add.ps1 GroupName_Owner.csv
```

- 配布リストの所有者に追加する
- GroupName_Owner.csv のカラム名:
  - groupName ... 配布リストのアドレス
  - owner ... 追加する管理者のUPNまたはメールアドレス

  - 例）
    ```
    groupName,owner
    test.group01@niigata-u.ac.jp,niigata.taro@niigata-u.ac.jp
    test.group02@niigata-u.ac.jp,nagaoka.jiro@niigata-u.ac.jp
    ```


共有メールボックス
------------------

引数に関する共通仕様
- SHARED-ADDRESS ... 共有メールボックスのプライマリメールアドレス
- SHARED-ADDRESS.EXT.txt ... 共有メールボックスののメンバーリストを変更するためのデータファイル。
  1行-1メールアドレス/UPNのテキストファイル。
  EXT は、'`.[-_A-Za-z0-9]+` にマッチする文字列を使うこと。

### Shared-Setup.ps1

```
  Usage: Shared-Setup.ps1 SHARED-ADDRESS
```

- 共有メールボックスを日本語用に設定する。


### Shared-ListCsv.ps1

```
  Usage: Shared-ListCsv.ps1
```

- すべての共有メールボックスのメンバーリストを出力する。
- `shared-list-all.yyyyMMdd_HHmmss.csv` がカレントディレクトリに作成される。各列は次の通り
  - Shared.Mailaddr	... 共有メールボックスのメールアドレス
  - Shared.DisplayName ... 共有メールボックスの表示名
  - User ... 利用者のMS365アカウント
  - AccessRights ... 利用者のアクセス権

### SharedMember-ListCsv.ps1

```
  Usage: SharedMember-ListCsv.ps1 SHARED-ADDRESS
```

- 共有メールボックスのメンバーリストを出力する。
- `SHARED-ADDRESS.member.csv` がカレントディレクトリに作成される。各列は次の通り
  - Shared.Mailaddr	... 共有メールボックスのメールアドレス
  - Shared.DisplayName ... 共有メールボックスの表示名
  - User ... 利用者のMS365アカウント
  - AccessRights ... 利用者のアクセス権

### SharedMember-Add.ps1

```
  Usage: SharedMember-Add.ps1 SHARED-ADDRESS.member-add.txt
```

- 作成済の共有メールボックスに利用者を追加する
