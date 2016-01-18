text2html
=====================================================================

## 概要 ##

text2htmlはwiki表記のファイルを、EXEに対してドラッグすることでHTMLファイルを出力します

## 起動方法 ## 

	1) 同梱されている「サンプル」を作りたいHTMLファイルのテンプレートに置き換える
	2) wiki形式で記述したファイルを text2html.exeに対してdrag-and-dropする
	3) drag-and-dropしたファイルと同じディレクトリにHTMLファイルが作成される

## サポートしているwikiの書き方 ## 

	pukiwikiに最低限の準拠
	
	○タイトル表示（拡張部分）
	
	#title(タイトル文字)
	
	○画像の表示
	
	&ref(画像)
	&ref(画像,説明文])
	#ref(画像)
	#ref(画像,説明文])
	
	説明文はalt属性値としてテキストとしてレンダリングされます
	
	○リンク
	
	[[文字>アドレス]]
	
	アドレスは相対パスも指定可能です
	
	○見出し
	
	* 見出し1
	** 見出し2
	*** 見出し3
	**** 見出し4
	
	○箇条書き（リスト）
	
	- リスト1
	-- リスト2
	--- リスト3
	
	○表
	
	| 説明1 | 説明2 |
	|説明|説明|
	|説明|説明|
	
	○罫線
	
	----
	
	○強調・斜体
	
	''インライン要素''
	
	○整形済みテキスト
	
	{{{
	テキスト
	}}}
	
	○その他
	
	コメント行
	
	// コメント

## License
Copyright 2016 nehori.

Licensed under the Apache License, Version 2.0 (the "License");
You may not use this file except in compliance with the License.
You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
