"""Reproducible Japanese Excel templates and user guide."""
import csv
from pathlib import Path

from .inputs import COLUMNS
from .workbook import InputError, write_xlsx

EXAMPLES = [
    ['true', 'GET_health', 'http://127.0.0.1:8080/health', 'GET', '{"check":"日本語"}',
     '{"Accept":"application/json"}', 'none', '', '10000', '30000', '200'],
    ['false', 'POST_json', 'http://127.0.0.1:8080/echo?existing=1', 'POST', '{"mode":"test"}',
     '{"Accept":"application/json","X-Example":"sample"}', 'json', '{"message":"日本語","count":1}',
     '10000', '30000', '200-299'],
    ['false', 'POST_form', 'http://127.0.0.1:8080/echo', 'POST', '{"source":"sample"}', '{}', 'form',
     '[{"name":"item","value":"A"},{"name":"item","value":"B"},{"name":"text","value":"日本語 & +"}]',
     '10000', '30000', '200'],
    ['false', 'POST_raw', 'http://127.0.0.1:8080/echo', 'POST', '{}', '{"Content-Type":"text/plain; charset=UTF-8"}',
     'raw', '1行目\n2行目', '10000', '30000', '200'],
]

FIELD_GUIDE = [
    ['列名', '必須／既定値', '記入方法・制約'],
    ['enabled', '省略／空欄: true', 'true / false または 1 / 0。false の行は実行しません。テンプレートは GET の1行だけ true。'],
    ['name', '必須', '有効行内で一意の名前（1～120文字）。結果のラベルになります。改行・制御文字・${ は不可。秘密情報を含めないでください。'],
    ['url', '必須', 'http:// または https:// で始まる絶対URL。既存クエリを保持し query_params を追加。日本語や空白はUTF-8でURLエンコード。認証情報付きURLと #fragment は不可。'],
    ['method', '必須', 'GET または POST。大文字／小文字は問いません。GET に本文は指定できません。'],
    ['query_params', '空欄: なし', 'JSONオブジェクト {"q":"日本語","page":"1"}。値は文字列。重複キーは [{"name":"tag","value":"a"},{"name":"tag","value":"b"}]。POSTでもURLに付加します。'],
    ['headers', '空欄: なし', 'JSONオブジェクト {"Accept":"application/json","X-Example":"value"}。名前／値は文字列。重複名、改行、非ASCII値、Content-Length、Transfer-Encoding は不可。'],
    ['body_type', '空欄: none', 'none / json / form / raw。POSTで本文を指定する場合は明示。multipart・ファイルアップロードには対応しません。'],
    ['body', 'noneでは空欄', 'json: JSON本文をそのまま入力。form: query_paramsと同じJSON形式で指定するとURLエンコード。raw: テキストをそのまま送信。日本語・改行・先頭末尾空白を保持します。'],
    ['connect_timeout_ms', '空欄: 10000', '接続タイムアウト（ミリ秒）。1～2147483647 の整数。0（無制限）は不可。'],
    ['response_timeout_ms', '空欄: 30000', '応答読み取りのタイムアウト（ミリ秒）。1～2147483647 の整数。全実行時間の上限は別途 --max-run-seconds。'],
    ['expected_status', '空欄: 200-299', '成功と判定するHTTPコード。200、200,201、200-299、200,404 の形式。指定の404は成功として扱えます。通信エラーは失敗です。'],
    ['Content-Type', '本文形式から補完', 'json: application/json; charset=UTF-8、form: application/x-www-form-urlencoded; charset=UTF-8、raw: text/plain; charset=UTF-8。明示指定する場合も文字コードはUTF-8。'],
    ['Excel形式', '.xlsx / APIs シート', '1行目は英語列名、2行目以降はAPI。列順の変更は可。結合セル・数式は不可（値を貼り付け）。通常の「Excelブック (*.xlsx)」で保存。古い .xls / マクロ .xlsm は非対応。'],
    ['CSV形式', 'UTF-8、BOM有無可', 'カンマ区切り。Excelでは「CSV UTF-8」で保存。JSONのダブルクォートはCSV規則で二重化し、改行を含むセル全体を引用符で囲みます。テンプレートを表計算ソフトで編集すると自動処理されます。'],
    ['共通制限', 'ファイル10 MiB以内', 'API行数10000以内。XLSX展開後50 MiB以内、64列以内。Excelセルは32767 UTF-16文字以内。未知の列名はエラー。空行は無視します。'],
    ['JMeter式', '入力は文字列として送信', '${__groovy(...)} や ${variable} は実行せず本文／ヘッダのリテラルとして送ります。CSVやExcelから任意のJMeterスクリプトは実行できません。'],
]


def guide_sheets():
    overview = [
        ['項目', '説明'],
        ['目的', 'Excel/CSVに記入したREST API定義からJMeter 5.6.3用JMXを生成し、CLIモードで実行してJTLとログを保存します。'],
        ['対応対象', 'RHEL 9.8: Bash + Python 3.9以降。Windows 11: Windows PowerShell 5.1 / PowerShell 7 + Python 3.9以降。JMeterは5.6.3固定。Javaは8以降、手順では17を使用。'],
        ['追加ライブラリ', 'Python標準ライブラリのみ。pip、Excelのインストール、COM、追加のJMeterプラグインは不要です。'],
        ['デフォルト件数', 'ユーザー数1 × ループ1 × 有効API行数。テンプレートは有効APIが1行なので1件です。複数行を有効にすると各行を1回ずつ実行します。'],
        ['順序', 'ユーザーごとに入力行順に実行します。複数ユーザーでは並行実行されます。失敗後も残りのAPI行を実行して全件の結果を記録します。'],
        ['追加通信', '自動リダイレクト・自動リトライ・埋め込みリソース取得を無効化。HTTP 3xxもその応答だけを検査します。再実行するとPOSTも再送されます。'],
        ['認証・セッション', 'ヘッダに必要なAuthorization等を入力してください。Cookieの自動引継ぎ、認証フロー、レスポンス値の後続行への引継ぎは実装していません。'],
        ['成果物の管理', 'JMXには入力したURL・本文・ヘッダ値が保存されます。JMeterの診断ログにも接続先情報が含まれる場合があります。入力と出力は機密ファイルとして管理してください。'],
        ['OS検証範囲', '実際に実施した検証環境・件数は同梱の docs/VALIDATION.md を参照してください。RHEL 9.8実機での確認はその環境で検証手順を実行してください。'],
    ]
    setup = [
        ['環境／手順', 'コマンド・説明'],
        ['RHEL 9.8 前提', 'bash、python3（3.9以降）、Java 17、unzip。必要に応じ管理者が sudo dnf install python3 java-17-openjdk-headless unzip を実行。'],
        ['JMeter取得', 'Apache公式の apache-jmeter-5.6.3.zip を取得し、同じ公式配布元の .sha512 と照合して展開します。展開先を JMETER_HOME に指定してください。自動ダウンロード／インストールは行いません。'],
        ['RHEL 配置例', 'プロジェクト全体を /opt/jmeter-api 等に配置し、JMeterは /opt/apache-jmeter-5.6.3 に展開。スクリプトだけをコピーせず、restjmeter/ と jmeter_api.py も含めて配置してください。'],
        ['RHEL 環境設定', 'export JMETER_HOME=/opt/apache-jmeter-5.6.3\npython3 --version\njava -version\nbash scripts/jmeter-api.sh --help'],
        ['RHEL 権限・改行', 'bash scripts/jmeter-api.sh で実行すれば実行権限の付与は不要。シェルはLF形式を保持。出力ディレクトリは書込み可能な場所を指定します。umask 077で作成。'],
        ['Windows 11 前提', 'Python 3.9以降（py.exe または python.exe）とJava 17を導入。JMeterのZIPを C:\\Tools\\apache-jmeter-5.6.3 等に展開します。'],
        ['Windows 環境設定', "$env:JMETER_HOME = 'C:\\Tools\\apache-jmeter-5.6.3'\npy -3 --version\njava -version"],
        ['PowerShell 実行ポリシー', '組織の実行ポリシーに従って実行してください。ダウンロードしたファイルがブロックされている場合は内容を確認し、許可される場合に対象ps1のみ Unblock-File -LiteralPath を使用。スクリプト自身はポリシーを変更しません。'],
        ['ローカル作成スクリプトの実行', 'グループポリシーによる制限がなく、ローカルスクリプトの実行が認められる環境では powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File .\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.xlsx を使用できます。この指定は起動するプロセスだけに適用されます。'],
        ['任意の実行ファイル', 'RHEL: PYTHON_BIN=/path/to/python3 bash scripts/jmeter-api.sh ...。Windows: -PythonPath C:\\Python312\\python.exe。Javaは --java / -JavaPath、JAVA_HOME、PATH の順で選択します。'],
        ['Javaランチャー', 'Javaを直接起動するため、JMeterのbat/sh専用変数 HEAP・JVM_ARGS は使用しません。ヒープは128～512 MiB。JAVA_TOOL_OPTIONS等のJava自身の変数はJVMの仕様に従います。'],
        ['既存JMeter設定', 'JMeterホームの設定を読み込み、runtime.propertiesでリトライとJTL出力形式などを指定。プロキシ・追加プラグイン・独自設定のある環境では影響を確認してください。'],
    ]
    commands = [
        ['目的', 'RHEL 9.8', 'Windows 11 PowerShell'],
        ['ローカル確認用API起動（別ターミナル）', 'python3 examples/mock_api.py --port 8080', 'py -3 examples/mock_api.py --port 8080'],
        ['Excelから生成して1件実行', 'bash scripts/jmeter-api.sh --input templates/api-template.xlsx --jmeter-home /opt/apache-jmeter-5.6.3', ".\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.xlsx -JMeterHome C:\\Tools\\apache-jmeter-5.6.3"],
        ['CSVから実行', 'bash scripts/jmeter-api.sh --input templates/api-template.csv', '.\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.csv'],
        ['送信せずJMXだけ生成', 'bash scripts/jmeter-api.sh --input templates/api-template.xlsx --generate-only', '.\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.xlsx -GenerateOnly'],
        ['HTMLレポートも生成', 'bash scripts/jmeter-api.sh --input templates/api-template.csv --html-report', '.\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.csv -HtmlReport'],
        ['ユーザー数と繰返しを変更', 'bash scripts/jmeter-api.sh --input templates/api-template.csv --threads 2 --loops 3', '.\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.csv -Threads 2 -Loops 3'],
        ['出力先と全体タイムアウト', 'bash scripts/jmeter-api.sh --input templates/api-template.csv --output "./run results" --max-run-seconds 600', '.\\scripts\\Invoke-JMeterApi.ps1 -InputPath .\\templates\\api-template.csv -OutputDirectory ".\\run results" -MaxRunSeconds 600'],
        ['テンプレート再生成', 'bash scripts/jmeter-api.sh --create-templates --template-dir ./new-templates', '.\\scripts\\Invoke-JMeterApi.ps1 -CreateTemplates -TemplateDirectory .\\new-templates'],
        ['終了コード確認', 'echo $?', '$LASTEXITCODE'],
    ]
    outputs = [
        ['ファイル／終了コード', '内容'],
        ['output/run-<UTC日時>-<識別子>/', '実行ごとに新規ディレクトリを作成。前回出力を上書きしません。--outputはこの親ディレクトリです。'],
        ['test-plan.jmx', 'JMeter GUIで開けるテストプラン。指定したユーザー数／ループ数とAPI定義、ステータス検証、JTL保存リスナー、結果ログリスナーを含みます。'],
        ['runtime.properties', 'JMeterの起動時に -q で読み込む設定。生成したJMXを手動実行する場合も併用します。'],
        ['JMeter-User-Guide.xlsx', 'この日本語利用ガイド。テンプレート作成時・JMX生成時に自動生成。'],
        ['results.jtl', 'UTF-8 CSV形式のJMeter計測結果。各リクエストの時刻、経過時間、HTTPコード、成功判定、送受信バイト数など。レスポンス本文／送信本文／ヘッダは保存しません。'],
        ['jmeter.log', 'JMeter診断ログ。API_RESULT行にラベル・応答コード・時間・成功判定を記録します。'],
        ['console.log', 'JMeterプロセスの標準出力／標準エラー。実行中も参照できます。'],
        ['execution.log', 'スクリプトの開始・終了と、JTLを読み取った結果のSAMPLE行。応答コード・時間・成功判定を保存。'],
        ['summary.json', '総件数、成功・失敗件数、予定件数一致、平均時間、終了状態、終了コード。API行ごとの予定件数も照合します。'],
        ['html-report/index.html', '--html-report / -HtmlReport 指定時のHTMLダッシュボード。関連ログは report.log / report-console.log。1件でも作成できますが統計評価には十分な件数が必要です。'],
        ['終了コード 0', '生成成功、または全APIが期待ステータスで完了し、予定のサンプル件数と一致。'],
        ['終了コード 2', '入力・前提条件・ファイル操作エラー。'],
        ['終了コード 3', 'JMeter起動／実行異常、JTL不正／欠落、予定件数不一致、HTMLレポート生成異常。'],
        ['終了コード 4', 'APIの通信エラー、または期待ステータス以外の応答。JMeter自身の終了値が0でもこのスクリプトは4を返します。'],
        ['終了コード 5 / 130', '5:全体タイムアウト。130:Ctrl+Cで中断（Pythonが処理できた場合）。起動したJavaを終了してログとsummaryを残します。タイムアウトはJMeter実行とHTML生成の各プロセスに適用。'],
    ]
    jmeter = [
        ['操作', '方法'],
        ['GUIを起動', 'RHEL: $JMETER_HOME/bin/jmeter、Windows: & "$env:JMETER_HOME\\bin\\jmeter.bat"。RHELでGUIを使う場合は別途デスクトップ環境とGUI対応Javaが必要です。'],
        ['JMXを開く', 'File > Open から test-plan.jmx を選択。Thread Groupでユーザー数とループ数、各HTTP RequestでURL・メソッド・Body Dataを確認します。'],
        ['クエリと本文', 'クエリはHTTP RequestのPathにある完全URLへ含めています。POST本文はBody Dataにあります。Header Managerでヘッダを、Response Assertionで期待ステータスを確認します。'],
        ['CLI実行を推奨', '通常の実行は同梱スクリプトを使用。GUIはプランの確認・編集に使用します。スクリプトを再実行すると入力ファイルから新たなJMXを生成します。'],
        ['生成済みJMXの手動実行（RHEL）', '新しい空の実行用ディレクトリにJMXとruntime.propertiesをコピーして移動し、 "$JMETER_HOME/bin/jmeter" -n -t test-plan.jmx -q runtime.properties -j jmeter.log を実行。'],
        ['生成済みJMXの手動実行（Windows）', '新しい空の実行用ディレクトリで & "$env:JMETER_HOME\\bin\\jmeter.bat" -n -t test-plan.jmx -q runtime.properties -j jmeter.log を実行。'],
        ['手動実行の保存先', 'JMX内のSimple Data Writerが実行ディレクトリのresults.jtlを保存。-Jrest.results.file=別名.jtl で変更できます。同じファイルへの追記や二重保存を避けるため、新規ディレクトリを使い、同じ保存先の -l は併用しないでください。'],
        ['手動実行時の違い', 'JMXはJTLとJMeterログを出力します。スクリプト独自の終了コード判定・execution.log・summary.json・全体タイムアウトはスクリプト経由で実行した場合に有効です。'],
        ['JTLを確認', 'GUIのリスナー（Summary Report / View Results Tree）のBrowseから results.jtl を選択。CSV計測値を確認できます。本文は保存していないため後から本文表示はできません。'],
        ['HTMLだけ再生成', '既存results.jtlから jmeter -g results.jtl -o 新しい空の出力フォルダー -j report.log を実行。'],
        ['本文中の${...}', 'JMeterの関数実行を防ぐため、その文字列を含む本文／ヘッダはBase64から文字列へ復元する固定式としてJMXに記録します。送信内容は元の文字列です。'],
    ]
    troubleshooting = [
        ['症状', '確認事項'],
        ['入力エラー', '行番号・列名を確認。JSONはダブルクォートで文字列を囲む。CSVはUTF-8。ExcelはAPIsシートの1行目が列名、数式は値へ変換。'],
        ['JMeterが見つからない／版が違う', '--jmeter-home / -JMeterHome / JMETER_HOME がbinの親を指すこと。lib/ext/ApacheJMeter_core.jar のImplementation-Versionが5.6.3であること。'],
        ['Python / Javaが起動しない', '--help、python3 --version / py -3 --version、java -version を確認。空白を含むパスは引用。JAVA_HOMEの誤りはPATHで隠さず修正。'],
        ['接続不可', 'URL・ポート・DNS・ネットワーク経路・ファイアウォール・プロキシ・サービス起動状態を確認。テンプレートの127.0.0.1:8080は確認用APIを起動してから使用します。'],
        ['HTTP 3xx', '自動追跡しません。目的のURLに変更するか、3xx自体を検査する場合はexpected_statusに指定。'],
        ['HTTP 4xx / 5xx', 'jmeter.logとresults.jtlを確認し、認証・ヘッダ・クエリ・本文・サーバー側の状態を見直します。レスポンス本文は保存していません。'],
        ['HTTPS', 'JMeter 5.6.3の標準実装はサーバー証明書のエラーを無視するTrustManagerを使用します。本ツールもこの仕様に従うため、APIテストの成功で証明書の有効性は判定できません。証明書検証が必要な用途では別途検証が必要です。'],
        ['タイムアウト', '接続／応答タイムアウトと --max-run-seconds / -MaxRunSeconds を区別。件数を増やしたら全体上限も見直します。5で終了した場合、処理済みPOSTが存在し得るためサーバー側を確認してから再実行。'],
        ['テンプレート作成先に既存ファイル', '既存ファイルは上書きしません。別の --template-dir / -TemplateDirectory を指定。'],
        ['長いパス・権限', 'Windowsの長いパス制限とフォルダーACLを確認。出力先は短い書込み可能なパスにしてください。RHELは出力先の所有者・権限を確認。'],
    ]
    verification = [
        ['確認', '手順／合格条件'],
        ['単体テスト', 'RHEL: python3 -m unittest discover -s tests -v\nWindows: py -3 -m unittest discover -s tests -v\n実JMeterテストはJMETER_TEST_HOMEがない場合は明示的にskip。'],
        ['JMeter統合テスト', 'RHEL: JMETER_TEST_HOME=/opt/apache-jmeter-5.6.3 python3 -m unittest discover -s tests -v\nWindows: $env:JMETER_TEST_HOME="C:\\Tools\\apache-jmeter-5.6.3"; py -3 -m unittest discover -s tests -v'],
        ['実環境での受入れ', '別ターミナルでexamples/mock_api.pyを起動し、各OSのスクリプトでExcelとCSVを実行。結果がpassed、サンプル1、失敗0であること。POST例を有効にして送信形式と件数も確認。'],
        ['確認済み環境', '同梱 docs/VALIDATION.md が実施済み検証の記録です。特定のOS・ネットワーク・認証設定を含めた無条件の動作保証ではありません。'],
    ]
    sources = [
        ['一次資料', 'URL'],
        ['Apache JMeter 5.6.3 配布元', 'https://archive.apache.org/dist/jmeter/binaries/'],
        ['JMeter CLI・Java要件', 'https://jmeter.apache.org/usermanual/get-started.html'],
        ['HTTP Request / Listener / Assertion', 'https://jmeter.apache.org/usermanual/component_reference.html'],
        ['HTMLダッシュボード', 'https://jmeter.apache.org/usermanual/generating-dashboard.html'],
        ['5.6.3 固定版プロパティ', 'https://github.com/apache/jmeter/blob/rel/v5.6.3/bin/jmeter.properties'],
        ['PowerShell実行ポリシー', 'https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies'],
        ['JMeter 5.6.3の証明書検証仕様', 'https://github.com/apache/jmeter/blob/rel/v5.6.3/src/core/src/main/java/org/apache/jmeter/util/CustomX509TrustManager.java'],
    ]
    return [('概要', overview, [30, 110]), ('セットアップ', setup, [32, 120]),
            ('入力仕様', FIELD_GUIDE, [27, 30, 115]), ('実行例', commands, [34, 100, 110]),
            ('出力と終了コード', outputs, [40, 115]), ('JMeter操作', jmeter, [40, 115]),
            ('トラブルシュート', troubleshooting, [35, 110]), ('検証手順', verification, [32, 120]),
            ('参照資料', sources, [48, 110])]


def write_guide(path):
    write_xlsx(path, guide_sheets())


def create_templates(directory):
    directory = Path(directory)
    filenames = ['api-template.csv', 'api-template.xlsx', 'JMeter-User-Guide.xlsx']
    if any((directory / name).exists() for name in filenames):
        raise InputError('Template destination already contains an output file; choose a new directory.')
    directory.mkdir(parents=True, exist_ok=True)
    with (directory / filenames[0]).open('x', encoding='utf-8-sig', newline='') as stream:
        csv.writer(stream).writerows([COLUMNS] + EXAMPLES)
    write_xlsx(directory / filenames[1], [('APIs', [COLUMNS] + EXAMPLES,
               [12, 22, 60, 12, 48, 58, 14, 72, 22, 23, 24]), ('入力ガイド', FIELD_GUIDE, [27, 30, 115])])
    write_guide(directory / filenames[2])
