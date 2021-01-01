# OldMacroParseTest
## Overview
This is a program to parse macro in Microsft Office file.  
This program runs at the command prompt.
This is for older formats prior to 2003.  
This is a test code, not a finished product.  
Please use it as a reference for you to create an analysis program.  

## Build
It is made by Microsoft Visual Studio 2017, Visual C++.  
Please use same or upper version if you will build this solution.  


## 概要
2003以前の旧型式のMicrosoft Officeファイル（.doc、.xls）のストレージを開き、ストリームの内容をダンプするプログラム。  
コマンドプロンプトで動作。  
コマンドプロンプトには、ストリーム名称とサイズを表示。  
プログラムの実行ディレクトリに雑にダンプファイルが出力される。  
今のところ、同じストリーム名称の場合、ダンプデータが後勝ちで上書きされてしまう。  
ファイル存在チェックし、既に同名ファイルがある場合は(n)などを付ける、せめて解析ファイル名からディレクトリを作成してそこにダンプするなど、色々苦情は来そうだけど、あくまで自身の確認のためのプログラムなので、そこまで頑張って作っていない。  
時間があれば直すかもしれないけど、基本メンテは期待しない方向で。  
不満なら、ソースをある程度参考にして自分で作り直したほうがいいとは思うので、あくまでサンプル用ジャンクコードということで。  
ダンプする時に、マクロコードがある部分はLZNT1で解凍して別途出力している。  
ただし、一部形式のデータが判定にヒットせず、解凍処理をスルーしている。  
現在、判定漏れの理由は調べているところ。  
.xlsのドキュメント等を参考にすればいいとは思うが、かなりの量な上に英文のため現在のところ見つかっておらず。  
<BR>
ファイルフォーマットは以下を参照。  
**[MS-XLS]: Excel Binary File Format (.xls) Structure（英文）**  
https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06?redirectedfrom=MSDN  

## 構築
Microsoft Visual Studio 2017の Visual C++ で作成。  
今のところ、.xlsと.docでは動作することを確認。  
.pptは2021/1/1時点では未確認。  
ただし、.docで2007以降の形式（があるらしい）には非対応。  
2007以降の形式は、もう拡張子を.zipにして解凍したほうが早い気がするので頑張ってないです。  
