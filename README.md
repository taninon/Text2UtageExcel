# Text2UtageExcel

### 何するのか
　テキストからUnity用ビジュアルノベルツール「宴」向けのExcelシートを作成します。

### なにが嬉しいのか
　「宴」 https://madnesslabo.net/utage/   はとても優れたUnityAssetですが「エクセルでシナリオ編集」という特徴があります。
 
 　コンシューマーゲーやスマートフォン開発においては慣れ親しまれた手法ですが、膨大な文章量を扱うノベルゲー制作のおいては（歴史的背景もあり）テキストファイルでのシナリオ用データ制作が好まれています。
　テキストファイルから「宴」用のExcelシートが作れればそれなりに嬉しい人がいます。僕がそうです。

　あとまあバージョン管理の観点からもテキストで扱ったほうが有利だと思います。


### どうやって使うの
　Text2UtageExcelConverterのScriptableObjectファイルを作成します。
（SampleData/ConverterTest.assetを参考）

　ScriptableObjectには追加するExcelファイルとテキストファイルを貼り付けておきます。
　メニューのTools > UtageからText2UtageExcelConverterを開き、作成したScriptableObjectを付けてConvertを押してください。
 
　設定などは主にこのScriptableObjectから行います。
 
　「テキストに更新がない場合は変換しない」チェックボックスがつきました。MD5を記録しておいて変更がなかったら変換しないです。早い。みんな喜ぶ。

### スクリプト仕様
　基本的にはコンバーターですので宴のコマンドを参照してください。

スクリプトマニュアル
https://madnesslabo.net/utage/?page_id=252

@命令 Arg1=
という書式で書くことができます。

例:
@Wait Arg6=0.5

### ちょっと扱いが異なるもの

#### 宴Ver3からwaitType列ができました。

@Bg arg1=bg001 arg6=0.5 waitType=PageWait

のように指定できるようになっています。

#### TweenのArg3

複数の指定を書くには{}でまとめて,で区切ります。
@Tween Arg1=BG Arg2=MoveBy Arg3={time=0.4,y=-450}

（苦肉の策）



### ライセンス
本プログラムの利用には利用および改変ほか制限はありません。

NPOIはApache Licenseです
 Apache License Version 2.0
 Editor\NPOI\NpoiLicense.txt

