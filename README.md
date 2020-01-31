# Text2UtageExcel

### 何するのか
　テキストからUnity用ビジュアルノベルツール「宴」向けのExcelシートを作成します。

### なにが嬉しいのか
　「宴」 https://madnesslabo.net/utage/   はとても優れたUnityAssetですが「エクセルでシナリオ編集」という特徴があります。
　これはコンシューマーゲーやスマートフォン開発においては慣れ親しまれた手法ですが、膨大な文章量を扱うノベルゲー制作のおいては（歴史的背景もあり）テキストファイルでのシナリオ用データ制作が好まれています。
　テキストファイルから「宴」用のExcelシートが作れればそれなりに嬉しい人がいます。僕がそうです。

### どうやって使うの
　Text2UtageExcelConverterのScriptableObjectファイルを作成します。
（SampleData/ConverterTest.assetを参考）

　ScriptableObjectには追加するExcelファイルとテキストファイルを貼り付けておきます。
　メニューのTools > UtageからText2UtageExcelConverterを開き、作成したScriptableObjectを付けてConvertを押してください。

### テキスト仕様
　ゆるーく吉里吉里KAG風です。

[bg arg1=町]
@bg arg1=町

などで命令文認識されます。
全角のテキストは文章として認識されます。
[l]クリック待ち、[p]改行として扱えます。

### 今後の展望
　つうか全角で判断してるけど英語の場合はどうすんのって気がしてきた。
　現状は命令変換などはしていませんが、TyranoScriptとの互換性をつければTyranoBuilder https://b.tyrano.jp/ というこれまた優れたツールが使えるようになるわけでイカスかもしれません。

### ライセンス
本プログラムの利用には利用および改変ほか制限はありません。

NPOIはApache Licenseです
 Apache License Version 2.0
 Editor\NPOI\NpoiLicense.txt

