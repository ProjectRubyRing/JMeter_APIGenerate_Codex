# 検証記録

検証日: 2026-10-09（日本時間）。実行した確認と、対象OSのうち実機で確認できなかった範囲を記録します。

## 環境と結果

| 対象 | 確認内容 | 結果 |
| --- | --- | --- |
| Windows 11ネイティブ | OS build 26200、Windows PowerShell 5.1.26100.9549 | 起動スクリプトの実行確認済み |
| Python 3.12.10 | 全27テスト | 26成功、1 skip |
| Python 3.9.13（Windows検証用ランタイム） | 全27テスト | 26成功、1 skip |
| Java 17.0.11 | Apache JMeter 5.6.3の実起動 | 成功 |
| JMeter配布物 | Apache公式ZIPのSHA-512照合、core JARの5.6.3版確認 | 一致 |
| Bashスクリプト | ShellCheck 0.11.0 `--shell=bash` | 指摘なし |
| PowerShellスクリプト | PowerShell 5.1のParserによる構文検査 | エラーなし |
| Pythonソース | `compileall`、Python 3.9構文互換性のAST検査 | 成功 |
| XLSX | 別実装のopenpyxlでテンプレート・9シートの利用ガイドを読み込み | 成功 |
| XLSX再保存 | openpyxlで真偽値・数値セルへ変更し再保存、入力結果を照合 | 一致 |
| 空白・改行 | `git diff --no-index --check`、CRLF/LF・UTF-8確認 | 不正な空白なし |
| RHEL 9.8実機 | ネイティブBashからJMeterまでの一連の実行 | **未実施** |
| PowerShell 7 | 実プロセスでの実行 | 未実施 |

skipした1件は `test_posix_bash_default_one_request` です。Windows上ではPOSIXシェルの実行テストを明示的にskipします。RHELではこのテストが有効になり、Windows用2テストがskipされます。

Python 3.9.13はPython APIの互換性確認に使用したWindows版です。RHEL 9.8のPython RPM、OpenJDK、OS・ネットワークの組合せを実行した結果ではありません。

この作業環境ではWSLの列挙が `E_ACCESSDENIED`、Docker Engineが利用不可でした。Git Bashもsignal pipe作成に失敗したため、`bash -n` とLinuxでの実行は成功扱いにしていません。Bashの構文・問題検出はWindowsで動くShellCheckを使って確認しました。

開始時にフォルダーは空でGitリポジトリではなかったため、通常の `git status` / `git diff --check` は利用できません。新規テキストファイルを空ファイルと比較する `git diff --no-index --check` で確認しました。CRLFを許容する `cr-at-eol` を指定し、末尾空白や空行は検査しています。

## 実通信で検証した内容

外部サービスには送信せず、テスト内で127.0.0.1の空きポートにHTTPサーバーを起動し、サーバーが受け取った内容とJTL・終了コードを照合しました。

- CSVからGETを1件だけ送信し、追加のリクエストが発生しないこと。
- 既存クエリと追加入力、空値、`+`、`&`、日本語、ヘッダの保持。
- ExcelからPOSTのJSON・フォーム・rawを送信し、URLのクエリと本文を分離すること。
- 同名フォーム項目、本文のCRLF/LF、先頭末尾空白、日本語、引用符、バックスラッシュの保持。
- 本文・ヘッダの `${__groovy(...)}` が実行されず、そのまま文字列として届くこと。
- 無効行を送信しないこと。
- HTTP 404の期待成功／失敗の判定、302を追跡しないこと、読み取りタイムアウト。
- JMeter自身が0で終了しても、APIが失敗した場合はスクリプトが4を返すこと。
- 2ユーザー×3ループで6件になること。
- 全体タイムアウトで5を返すこと、子プロセスを終了して回収すること。
- PowerShellからExcel／CSVで実行でき、成功0・失敗4を返すこと。
- 空白・日本語・`&`を含む入出力パス、およびPowerShellの空白を含む末尾 `\` 付き出力先。
- JMX、CSV形式JTL、`API_RESULT`を含むJMeterログ、実行ログ、summary、HTMLダッシュボードの生成。

入力エラー・テンプレート同等性・Excel sharedStrings・疎なセル・数式拒否・不正JTL・API別件数不一致も単体テストで検証しました。

## 実行した主なコマンド

```powershell
$env:JMETER_TEST_HOME = '<検証用のApache JMeter 5.6.3配置先>'
python -m unittest discover -s tests -v
<Python3.9実行ファイル> -m unittest discover -s tests -v
python -m compileall -q jmeter_api.py restjmeter tests examples/mock_api.py
shellcheck.exe --shell=bash scripts/jmeter-api.sh
```

Windowsの初期実行ポリシーは全スコープUndefinedで、実効値はRestrictedでした。ローカルで作成したスクリプトの検証には、子PowerShellだけを `-ExecutionPolicy RemoteSigned` で起動しました。MachinePolicy / UserPolicyはUndefined、永続的な実行ポリシーの変更は行っていません。

PowerShellの末尾区切り文字の修正後は、全テストをPython 3.12で再実行し、Python 3.9でも対象のPowerShell成功テストを再実行して成功を確認しました。

## RHEL 9.8での受入れ確認

前提ソフトを導入し、JMeter 5.6.3を展開してからプロジェクトのルートで実行してください。

```bash
export JMETER_HOME=/opt/apache-jmeter-5.6.3
python3 --version
java -version

# ローカルAPIはテストが自動起動します。
JMETER_TEST_HOME="$JMETER_HOME" python3 -m unittest discover -s tests -v

# 送信せず、シェル起動・Excel読込み・成果物作成を確認します。
bash scripts/jmeter-api.sh --input templates/api-template.xlsx --generate-only
```

運用手順の確認は、別ターミナルで `python3 examples/mock_api.py --port 8080` を起動し、以下を実行します。

```bash
bash scripts/jmeter-api.sh --input templates/api-template.xlsx
bash scripts/jmeter-api.sh --input templates/api-template.csv --html-report
```

各実行が終了コード0となり、`summary.json` が `status: passed`、`total_samples: 1`、`failed_samples: 0`、`sample_counts_match: true` であることを確認してください。

## 残る検証範囲

- RHEL 9.8実機での上記受入れ確認。
- Excelデスクトップアプリでの目視確認（OOXML構造・別ライブラリとの相互読込みは確認済み）。
- 実際の接続先のHTTPS、プロキシ、認証、クライアント証明書、ネットワーク制限。
- 大量リクエスト・長時間実行の性能評価。本ツールの既定目的は少数件のAPI動作確認です。

JMeter 5.6.3標準のHTTPS実装は証明書エラーを無視するTrustManagerを使用します。APIの成功判定は証明書の有効性検証を意味しません。[Apache 5.6.3ソース](https://github.com/apache/jmeter/blob/rel/v5.6.3/src/core/src/main/java/org/apache/jmeter/util/CustomX509TrustManager.java)
