# VBEThemeColorToolInstaller

License: The MIT license

Copyright (c) 2023 ことりちゅん(KotorinChunChun)

動作確認環境: Windows 10 Pro／Excel for Microsoft 365 MSO （32 ビット／64ビット）



## これは？

[furyutei氏が開発したVBEThemeColorTool](https://github.com/furyutei/VBEThemeColorTool) を、タスクスケジューラに登録して実行し続けるスクリプト兼インストーラです。

パソコンの使用中にOfficeが更新されてVBE7.dllが差し替わったことにより、突然VBEの色が変化するのが嫌でタスクスケジューラから1分毎に更新し続けるようにしました。
（多少不便でも良いなら大本のbatをスタートアップに登録したり、手動で実行することで事足ります。）

ついでに必要なファイル一式をインストールする機能を搭載させました。

## 使い方

VBEThemeColorTool.vbsをメモ帳で開いて、ソースコード序盤の「本スクリプトでカスタマイズが必要な箇所」のテーマ設定を自分のものに変えてください。
既定値は作者が愛用するテーマの「kc.xml」です。↓イメージ

![image](https://user-images.githubusercontent.com/55196383/227728255-9003d3fd-4399-4924-9576-91945c22e56a.png)

以下のように保存してからvbsを実行することでインストール/アンインストールが始まります。
![image](https://user-images.githubusercontent.com/55196383/227729248-9ebbed79-c608-451f-b1d1-bbbea69320ad.png)

※次の2つのファイルはそれぞれから入手して格納してください。

- [VBEThemeColorEditor.exeはここから](https://github.com/gallaux/VBEThemeColorEditor)
- [vbetctool.exeはここから](https://github.com/furyutei/VBEThemeColorTool/raw/master/dist/vbetctool.exe)

## インストール時の処理内容

1. `C:\Program Files\VBEThemeColorTool` に一式をコピーします。
（Excelが32bitでインストールされている場合は、x86にインストールされます）
2. スタートメニューにショートカットを作成します。
- アンインストールのためのVBEThemeColorTool.vbs
- xml作成のためのVBEThemeColorEditor.exe
3. タスクスケジューラにVBEThemeColorToolを登録し1回実行します。

![image](https://user-images.githubusercontent.com/55196383/227728762-3f9fbed5-8587-44fd-b854-b80be1401f4b.png)

![image](https://user-images.githubusercontent.com/55196383/227728791-df3f947a-4bc7-47a8-b4a1-cccea1c7bf7d.png)

![image](https://user-images.githubusercontent.com/55196383/227728820-745daebb-4549-46e0-aa7a-46430df1be83.png)


## アンインストールを選んだときの処理内容

1. `C:\Program Files\VBEThemeColorTool` フォルダを削除します。※自作テーマも消えます。注意してください。
2. スタートメニューから2つを削除します。
3. タスクスケジューラからVBEThemeColorToolを削除します。



## 色変更を止めたい場合の手操作

スタートメニューからVBEThemeColorTool.vbsを実行し、インストールを選択すると通常の色に戻すスケジュールに書き換わります。

手動でタスクスケジューラを起動しVBEThemeColorToolを削除すると1分毎に走るタスクも止められます。



##  作者情報

作者：ことりちゅん

Twitter：[@KotorinChunChun](https://twitter.com/KotorinChunChun)

ブログ：<[えくせるちゅんちゅん](https://www.excel-chunchun.com/)>

[GitHubダウンロード](https://github.com/KotorinChunChun/VBEThemeColorToolInstaller/archive/master.zip)

[GitHubリポジトリを閲覧](https://github.com/KotorinChunChun/VBEThemeColorToolInstaller)



##  更新履歴

| 日付     | 概要         |
| -------- | ------------ |
| 2023/3/26 | 初回リリース |
