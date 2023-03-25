# VBEThemeColorToolInstaller

License: The MIT license
Copyright (c) 2023 ことりちゅん(KotorinChunChun)
動作確認環境: Windows 10 Pro／Excel for Microsoft 365 MSO （32 ビット／64ビット）



# これは？

[furyutei氏が開発したVBEThemeColorTool](https://github.com/furyutei/VBEThemeColorTool) をタスクスケジューラに登録して実行し続けるインストーラです。

パソコンの使用中にOfficeが更新されてVBE7.dllが差し替わったことにより、突然VBEの色が変化するのが嫌でタスクスケジューラから1分毎に更新し続けるようにしました。

ついでに必要なファイル一式をインストールする機能を搭載させました。

# 使い方

VBEThemeColorTool.vbsをメモ帳で開いて、ソースコード序盤の「本スクリプトでカスタマイズが必要な箇所」のテーマ設定を自分のものに変えてください。
既定値は作者が愛用するテーマの「kc.xml」です。↓イメージ

![image](https://user-images.githubusercontent.com/55196383/227728255-9003d3fd-4399-4924-9576-91945c22e56a.png)

以下のように保存してからvbsを実行することでインストール/アンインストールが始まります。
- VBEThemeColorTool.vbs     ちゅんが新たに開発したインストーラ兼タスクスケジューラ呼び出し用のvbsファイルです。
- vbetctool.exe             風柳氏が開発したVBE7.dllの書き換えプログラム本体です。
- VBEThemeColorEditor.exe   例のxml作成用ツールです。VBE7.dllの書き換えには使いません。
- Themes\*.xml              VBEThemeColorEditor.exeで作成した変更テーマのxmlを保存するところです。

※残念ながら再配布できないため、VBEThemeColorEditor.exe と vbetctool.exe はそれぞれから入手してください。

# インストール時の処理内容

1. `C:\Program Files\VBEThemeColorTool` に一式をコピーします。
（Excelが32bitでインストールされている場合は、x86にインストールされます）
2. スタートメニューにショートカットを作成します。
- アンインストールのためのVBEThemeColorTool.vbs
- xml作成のためのVBEThemeColorEditor.exe
3. タスクスケジューラにVBEThemeColorToolを登録し1回実行します。

# アンインストールを選んだときの処理内容

1. `C:\Program Files\VBEThemeColorTool` フォルダを削除します。※自作テーマも消えます。注意してください。
2. スタートメニューから2つを削除します。
3. タスクスケジューラからVBEThemeColorToolを削除します。

# 色変更を止めたい場合

スタートメニューからVBEThemeColorTool.vbsを実行し、インストールを選択すると通常の色に戻すスケジュールに書き換わります。

手動でタスクスケジューラを起動しVBEThemeColorToolを削除すると1分毎に走るタスクも止められます。

