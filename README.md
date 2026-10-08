# REST API → JMeter 5.6.3

ExcelまたはCSVにAPI定義を記入すると、JMXを生成してJMeterを非GUIで実行し、JTL・結果ログ・日本語のExcel利用ガイドを保存します。GETとPOST（JSON、フォーム、テキスト）に対応します。

**既定は有効なAPI行ごとに1リクエストです。** ユーザー数1・ループ1で、同梱テンプレートはGETの1行だけ有効です。自動リダイレクト、リトライ、埋め込みリソース取得は無効です。

## 配布ファイル

| ファイル | 用途 |
| --- | --- |
| [scripts/jmeter-api.sh](scripts/jmeter-api.sh) | RHEL 9.8用Bash起動スクリプト |
| [scripts/Invoke-JMeterApi.ps1](scripts/Invoke-JMeterApi.ps1) | Windows 11用PowerShell起動スクリプト |
| [templates/api-template.xlsx](templates/api-template.xlsx) | Excel入力テンプレート（APIs・入力ガイドの2シート） |
| [templates/api-template.csv](templates/api-template.csv) | CSV入力テンプレート（UTF-8 BOM付き） |
| [docs/JMeter-User-Guide.xlsx](docs/JMeter-User-Guide.xlsx) | セットアップ、入力、実行、JMeter操作、トラブル対応のExcelガイド |
| [examples/test-plan.jmx](examples/test-plan.jmx) | テンプレートから生成したJMXの例 |
| [examples/runtime.properties](examples/runtime.properties) | JMXの手動実行に併用する設定 |
| [examples/mock_api.py](examples/mock_api.py) | ローカル動作確認用API |
| [docs/VALIDATION.md](docs/VALIDATION.md) | 実施した検証と未検証範囲 |

両OSの処理は `jmeter_api.py` と `restjmeter/` に共通化しています。**プロジェクト全体を配置してください。** Excel本体、pipパッケージ、JMeterプラグインは必要ありません。

## 前提条件

- Apache JMeter **5.6.3**のバイナリ配布一式。
- Python **3.9以降**。
- Java **8以降**。以下のセットアップではJava 17を使用します。
- RHEL 9.8: Bash。Windows 11: Windows PowerShell 5.1またはPowerShell 7。

JMeter 5.6.xのJava要件と非GUI実行方法は[Apache公式ガイド](https://jmeter.apache.org/usermanual/get-started.html)を参照してください。[公式配布元](https://archive.apache.org/dist/jmeter/binaries/)の `apache-jmeter-5.6.3.zip` と対応する `.sha512` を取得し、SHA-512を照合してから展開します。本ツールは依存ソフトの自動インストールを行いません。

RHELでは必要に応じて管理者が次を導入します。

```bash
sudo dnf install python3 java-17-openjdk-headless unzip
export JMETER_HOME=/opt/apache-jmeter-5.6.3
python3 --version
java -version
```

WindowsではPythonとJavaを導入し、JMeterを展開します。

```powershell
$env:JMETER_HOME = 'C:\Tools\apache-jmeter-5.6.3'
py -3 --version
java -version
```

`JMETER_HOME` は `bin` の親ディレクトリです。Javaは `--java` / `-JavaPath`、`JAVA_HOME/bin/java`、`PATH` の順で選択します。

## 最初の実行

テンプレートの接続先は `http://127.0.0.1:8080/health` です。実際のAPIに書き換えるか、別ターミナルで同梱のAPIを起動してください。

```powershell
# Windows（RHELでは python3 examples/mock_api.py --port 8080）
py -3 examples/mock_api.py --port 8080
```

RHEL 9.8:

```bash
bash scripts/jmeter-api.sh --input templates/api-template.xlsx
bash scripts/jmeter-api.sh --input templates/api-template.csv --html-report
echo $?
```

Windows 11:

```powershell
.\scripts\Invoke-JMeterApi.ps1 -InputPath .\templates\api-template.xlsx
.\scripts\Invoke-JMeterApi.ps1 -InputPath .\templates\api-template.csv -HtmlReport
$LASTEXITCODE
```

PowerShellの実行ポリシーでスクリプトが禁止されている場合は、組織の設定に従ってください。グループポリシーの制限がなくローカル作成スクリプトの実行が認められる環境では、次のように対象プロセスに `RemoteSigned` を指定できます。永続設定の変更は行いません。[Microsoftの説明](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies)

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\scripts\Invoke-JMeterApi.ps1 -InputPath .\templates\api-template.xlsx
```

`--generate-only` / `-GenerateOnly` を付けるとAPIを呼ばず、JMX・設定・Excelガイドだけを生成します。この場合はPythonだけで動作します。

```bash
bash scripts/jmeter-api.sh --input templates/api-template.xlsx --generate-only
```

```powershell
.\scripts\Invoke-JMeterApi.ps1 -InputPath .\templates\api-template.xlsx -GenerateOnly
```

## 入力形式

Excelは通常の `.xlsx` 形式の **APIs** シートを読み取ります。CSVはUTF-8のカンマ区切りです（BOM有無可）。ExcelからCSVを保存する場合は「CSV UTF-8」を選んでください。

1行目が英語の列名、2行目以降がAPI定義です。列の順序変更と任意列の省略は可能ですが、未知の列名はエラーになります。

| 列 | 必須 | 既定値・記入例 |
| --- | --- | --- |
| enabled | | 空欄は `true`。`false` / `0` は実行しない |
| name | 必須 | 有効行で一意の識別名 |
| url | 必須 | `https://example.invalid/api/items?existing=1` |
| method | 必須 | `GET` / `POST` |
| query_params | | `{"q":"日本語","page":"1"}` |
| headers | | `{"Accept":"application/json","X-Example":"value"}` |
| body_type | | `none` / `json` / `form` / `raw`。空欄は `none` |
| body | | JSON本文、フォーム項目のJSON、またはテキスト本文 |
| connect_timeout_ms | | `10000`。正の整数、ミリ秒 |
| response_timeout_ms | | `30000`。正の整数、ミリ秒 |
| expected_status | | `200-299`。`200,201`、`200,404` 等も指定可能 |

- クエリ・ヘッダ・フォームの名前と値は **JSON文字列** で指定します。数値も `"1"` のように囲みます。
- 重複するクエリ／フォーム項目は `[{"name":"tag","value":"a"},{"name":"tag","value":"b"}]` で指定します。
- POSTのクエリはURLに付加し、本文とは混在させません。既存URLのクエリとパーセントエンコードを保持します。
- JSONとraw本文の先頭末尾空白、日本語、CRLF/LFを保持します。本文はUTF-8です。
- JSON／フォーム／rawのContent-Typeは未指定なら補完します。明示するcharsetもUTF-8を使用してください。
- GETの本文、ヘッダの重複や改行、Content-Length / Transfer-Encoding の手動指定、URL内の認証情報・フラグメントは拒否します。
- Excelの数式・結合セルは使用できません。数式は値へ変換してください。古い `.xls`、`.xlsm`、multipartは非対応です。
- `${...}` は本文／ヘッダにそのまま送信され、JMeterの式として実行されません。入力からの環境変数展開は行いません。

ファイルは10 MiB以内、APIは10000行以内。Excelは展開後50 MiB以内、64列以内、1セル32767 UTF-16文字以内です。

## オプション

| Bash / Python | PowerShell | 意味 |
| --- | --- | --- |
| `--input FILE` | `-InputPath FILE` | 入力ファイル |
| `--jmeter-home DIR` | `-JMeterHome DIR` | JMeter配置先。省略時はJMETER_HOME |
| `--output DIR` | `-OutputDirectory DIR` | 実行ディレクトリの親。既定 `output` |
| `--threads N` | `-Threads N` | ユーザー数。既定1、最大10000 |
| `--loops N` | `-Loops N` | 各APIの繰返し。既定1、最大1000000 |
| `--max-run-seconds N` | `-MaxRunSeconds N` | 各JMeterプロセスの上限。既定300秒 |
| `--generate-only` | `-GenerateOnly` | 生成のみ |
| `--html-report` | `-HtmlReport` | 実行後のHTMLレポート生成 |
| `--java PATH` | `-JavaPath PATH` | Java実行ファイル |
| `PYTHON_BIN=/path/python3` | `-PythonPath PATH` | Python実行ファイル |
| `--create-templates --template-dir DIR` | `-CreateTemplates -TemplateDirectory DIR` | テンプレート・Excelガイドを再生成。既存ファイルは上書きしない |

総リクエスト数は **有効API数 × ユーザー数 × ループ数** です。各ユーザーが行順に実行し、複数ユーザーでは並列になります。Cookie・認証フロー・レスポンス値の後続APIへの引継ぎは行いません。

Javaを直接起動してパスや引数を安全に渡します。JMeter起動用bat/shに依存する `HEAP` / `JVM_ARGS` は読み込みません。ヒープは128～512 MiBです。JVM自身の `JAVA_TOOL_OPTIONS` などと、JMeterホームの独自設定は影響します。

HTTPSはJMeter標準実装を使用します。JMeter 5.6.3のTrustManagerは証明書エラーを無視するため、本ツールのAPIテスト成功を証明書の有効性検証としては扱えません。[5.6.3の実装](https://github.com/apache/jmeter/blob/rel/v5.6.3/src/core/src/main/java/org/apache/jmeter/util/CustomX509TrustManager.java)

## 出力

実行ごとに `output/run-<UTC日時>-<識別子>/` を新規作成します。

| ファイル | 内容 |
| --- | --- |
| test-plan.jmx | HTTPリクエスト、期待ステータス判定、JTL保存、結果ログ出力を含むプラン |
| runtime.properties | 再試行無効化、JTL CSV形式などの設定 |
| JMeter-User-Guide.xlsx | 日本語のExcel利用ガイド（9シート） |
| results.jtl | JMeterのUTF-8 CSV計測結果 |
| jmeter.log | 診断ログとリクエストごとの `API_RESULT` |
| console.log | JMeterの標準出力・標準エラー |
| execution.log | スクリプトの開始・終了・サンプルごとの結果 |
| summary.json | 件数、失敗数、平均時間、状態、終了コード |
| html-report/index.html | HTMLレポート指定時のダッシュボード |
| report.log / report-console.log | HTMLレポート生成時のログ |

JTLには本文・リクエスト／レスポンスヘッダを保存しません。JMXには入力値が保存され、JMeter診断ログにも接続先情報が含まれる場合があるため、入力と出力は適切なアクセス権で管理してください。RHELは `umask 077`、Windowsは指定フォルダーのACLを継承します。

生成済みJMXを手動実行する場合は、新しい空のディレクトリにJMXと設定を置いて、次のように実行します。

```bash
"$JMETER_HOME/bin/jmeter" -n -t test-plan.jmx -q runtime.properties -j jmeter.log
```

JMX自身が `results.jtl` を保存します。保存先は `-Jrest.results.file=別名.jtl` で変更できます。同じ保存先の `-l` を併用しないでください。スクリプト独自の件数照合・終了コード・summaryは、同梱スクリプト経由で実行した場合に適用されます。

| 終了コード | 意味 |
| --- | --- |
| 0 | 生成成功、または全API成功・件数一致 |
| 2 | 入力、依存ソフト、パス／権限のエラー |
| 3 | JMeter実行、JTL解析、件数不一致、HTML生成のエラー |
| 4 | HTTP期待ステータス不一致／通信エラー |
| 5 | 全体タイムアウト |
| 130 | 中断をPythonが処理した場合 |

JMeterの終了コードが0でも、APIが失敗した場合は4を返します。全体タイムアウトはAPI実行とHTML生成のそれぞれに適用されます。中断・タイムアウトでも送信済みのPOSTは取り消されません。

## 検証

```bash
python3 -m unittest discover -s tests -v
JMETER_TEST_HOME=/opt/apache-jmeter-5.6.3 python3 -m unittest discover -s tests -v
```

```powershell
$env:JMETER_TEST_HOME = 'C:\Tools\apache-jmeter-5.6.3'
py -3 -m unittest discover -s tests -v
```

`JMETER_TEST_HOME` 未指定時は実JMeterテストが明示的にskipされます。テストは127.0.0.1に起動したAPIを使い、実際に届いた本文・ヘッダ・クエリ・件数を照合します。WindowsのPowerShellテストにはローカルスクリプトを実行できるポリシーが必要です。検証記録とRHEL実機での受入れ手順は [docs/VALIDATION.md](docs/VALIDATION.md) に記載します。
