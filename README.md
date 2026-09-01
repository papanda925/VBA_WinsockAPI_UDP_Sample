# Excel VBA Winsock UDP Sample

Excel VBAからWindowsのWinsock APIを直接呼び出し、ローカルPC内でUDP通信を試すための学習用サンプルです。

追加のActiveXコントロールや外部ライブラリは使用しません。VBA標準モジュール1ファイルで、UDPの受信側と送信側の基本的な処理順序を確認できます。

> [!IMPORTANT]
> UDPは接続確認や再送を行わないため、送信成功が相手の受信成功を保証しません。このリポジトリは学習用です。認証や暗号化はなく、インターネットへ公開する用途には適していません。

## このサンプルで確認できること

- `WSAStartup` / `WSACleanup`によるWinsockの初期化と終了
- `socket` / `bind` / `recvfrom`によるUDP受信
- `socket` / `sendto`によるUDP送信
- `closesocket`によるソケットの解放
- `PtrSafe`と`LongPtr`を使った32ビット／64ビットOffice対応
- 別のExcelプロセスを受信側として起動する方法

## 動作環境

- Windows
- Excel 2010以降（VBA 7）
- 32ビット版または64ビット版Office
- マクロを保存できるExcelブック（`.xlsm`または`.xlsb`）

macOS版ExcelではWindows DLLを呼び出せないため動作しません。

## ファイル

| ファイル | 内容 |
| --- | --- |
| `VBA_WinsockAPI_UDP_Sample.bas` | UDP受信、UDP送信、動作確認用マクロを含む標準モジュール |
| `SECURITY.md` | 脆弱性情報を非公開で連絡する場合の案内 |

## 使い方

### 1. 標準モジュールを取り込む

1. マクロ有効ブックを作成し、一度保存します。
2. `Alt` + `F11`でVisual Basic Editorを開きます。
3. メニューの「ファイル」→「ファイルのインポート」を選びます。
4. `VBA_WinsockAPI_UDP_Sample.bas`を選択します。
5. 「デバッグ」→「VBAProjectのコンパイル」を実行します。

### 2. 受信側を起動する

`MainForMultiProcess`を実行します。

同じブックが読み取り専用で別のExcelプロセスに開かれ、1秒後に`UDPRecvFrom`が開始します。既定では`127.0.0.1:60051`でデータグラムを待ちます。

### 3. 送信側から送る

最初のExcelへ戻り、次のいずれかを実行します。

| マクロ | 送信内容 | 受信側の動作 |
| --- | --- | --- |
| `testHELLO` | `HELLO` | 送信元のIPアドレスとポート番号を表示 |
| `testElse` | `else message` | 受信文字列を表示 |
| `testQUIT` | `QUIT` | 受信ループを終了 |

`testQUIT`は受信ループだけを終了します。別プロセスのExcel自体は自動終了しないため、内容を確認してから手動で閉じてください。

## 通信の流れ

### 受信側

```text
WSAStartup → socket → bind → recvfrom → closesocket → WSACleanup
```

### 送信側

```text
WSAStartup → socket → sendto → closesocket → WSACleanup
```

## 接続先を変更する

モジュール先頭付近の次の定数を変更します。

```vb
Private Const DEFAULT_SERVER_IP As String = "127.0.0.1"
Private Const DEFAULT_SERVER_PORT As Long = 60051
```

初めて試す場合は、外部から接続できないループバックアドレス`127.0.0.1`のまま使用してください。

## TCPサンプルとの違い

| 項目 | UDP | TCP |
| --- | --- | --- |
| 接続処理 | なし | `connect` / `accept`あり |
| 到達・順序の保証 | なし | あり |
| メッセージの境界 | データグラム単位 | バイトストリーム |
| このサンプルの送受信API | `sendto` / `recvfrom` | `send` / `recv` |

TCP版は[`VBA_WinsockAPI_TCP_Sample`](https://github.com/papanda925/VBA_WinsockAPI_TCP_Sample)にあります。

## 制限事項

- IPv4のみ対応
- 受信バッファは2,048文字の固定長
- 文字列はASCII範囲での利用を想定
- `recvfrom`は同期・ブロッキング処理
- 通信の認証・暗号化・再送制御はなし

`DoEvents`はブロッキング中の`recvfrom`を非同期化するものではありません。このサンプルでは、Excelの操作用プロセスと受信用プロセスを分けて影響を限定しています。

## トラブルシューティング

### 送信しても何も表示されない

受信側のExcelプロセスが起動していること、送信側と受信側でIPアドレスとポート番号が一致していることを確認してください。UDPでは受信側が起動する前に送ったデータは保持されません。

### `Address already in use`に相当するエラーになる

同じポートを使う受信処理が既に起動している可能性があります。別プロセスのExcelを閉じるか、`DEFAULT_SERVER_PORT`を未使用のポートへ変更してください。

### 64ビットOfficeでAPI宣言エラーになる

最新版の`.bas`を取り込み直し、「デバッグ」→「VBAProjectのコンパイル」を実行してください。ソケットはWindowsのハンドルであるため、変数とAPIの戻り値に`LongPtr`を使用しています。

## 参考資料

- [64-bit Visual Basic for Applications overview](https://learn.microsoft.com/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview)
- [socket function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-socket)
- [sendto function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-sendto)
- [recvfrom function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-recvfrom)
- [closesocket function](https://learn.microsoft.com/windows/win32/api/winsock2/nf-winsock2-closesocket)

## Repository scope

This is an educational Excel VBA sample for direct UDP communication through the Windows Winsock API.
