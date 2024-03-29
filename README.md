# 日報分析ツール

日報分析ツールは、PythonとTkinterを使用して構築されたデスクトップアプリケーションです。<br>
このツールを使用すると、指定されたフォルダ内のExcelファイルからデータを抽出し、それをテキストファイルに保存することができます。<br>
ユーザーはフォルダを選択し、ツールがExcelファイルからデータを抽出して指定されたフォーマットで保存します。

## 機能

- フォルダ選択: GUIを使用してフォルダを選択します。
- Excelファイルの読み込み: 選択されたフォルダ内のExcelファイルからデータを読み込みます。
- データの抽出: Excelファイルから特定の列のデータを抽出します。
- テキストファイルへの保存: 抽出されたデータをテキストファイルに保存します。

## 使用方法

1. アプリケーションを実行します。
2. アプリケーションウィンドウが表示され、"フォルダを選択"ボタンが表示されます。
3. ボタンをクリックしてフォルダ選択ダイアログを開きます。
4. フォルダを選択し、OKボタンをクリックします。
5. アプリケーションは選択されたフォルダ内のExcelファイルからデータを抽出します。
6. 抽出されたデータはテキストファイルに保存されます。
7. ユーザーには成功またはエラーメッセージが表示されます。

## インストール

1. リポジトリをクローンするか、ダウンロードします。
2. Pythonおよび、必要なPythonパッケージ（`openpyxl`）をインストールします。
3. `日報分析3.py`ファイルを実行します。

## 必要なもの

- Python 3.x
- Tkinter
- openpyxl

pythonのインストールはこちら⇒[Python公式サイト（https://www.python.org/downloads/windows/）](https://www.python.org/downloads/windows/)<br>
インストールが完了したら、次のコマンドを実行してバージョンを確認できます。

    python3 -VV

モジュールのインストールには、コマンドプロンプトで以下を実施してください。<br>

    pip install openpyxl

「pip」が使えなければ「pip3」、それでもだめなら「py -m pip」に置き換えて実行してみてください。<br>
それでもだめなら、python.exeの場所を確認するために、コマンドプロンプトで下記を実行して下さい。

    py --list-paths

上記の結果をもとに、環境変数にパスを追加します（コントロールパネルにて「環境変数」で検索）。<br>
例えば、上記コマンドの結果が、「C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python37_64\python.exe」であれば、<br>
環境変数のPathに、「C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python37_64\Scripts」を追加してください。

## 貢献方法

貢献は歓迎します！問題を開いたり、プルリクエストを送信したりしてください。

## ライセンス

このプロジェクトはMITライセンスの下で提供されています。詳細については、[LICENSE](LICENSE)ファイルを参照してください。
