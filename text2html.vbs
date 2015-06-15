'**
'* Copyright (c) 2015 Kazutaka Yasuda
'* Released under the MIT license
'* http://opensource.org/licenses/mit-license.php
'*
'**********************************************************
'  更新
'     v2.5a 2011. 4.20    <!--DATECREATED-->
'     v2.4a 2011.11.01    <!--DESCRIPTION-->機能追加
'     v2.3a 2011.10.26    <!--BASENAME-->機能追加
'     v2.2a 2011. 4.20    <!--SUB_BODY-->機能暫定追加
'     v2.1a 2011. 3.14    目次追加処理追加
'     v2.0a 2011. 3.13    読み込み文字コード自動判別追加
'     v1.6a 2011. 2.27    文字カラー処理更新
'     v1.5a 2010. 7.31    <!--break-->処理更新
'     v1.4a 2010. 7.28    整形済みテキスト<pre>追加
'     v1.3a 2010. 6. 8    取り消し線(Strikethrough)追加
'**********************************************************

' 以後の暗黙の変数定義を許可しない
Option Explicit

' 下記のファイルは用意しておくこと
Const HEADER   = "/mdl/header.mdl"
Const FOODER   = "/mdl/fooder.mdl"
Const SUB_BODY = "/mdl/sub_body.mdl"
' EXE変換用
Const MYFILE = "text2html.vbs"
' 基本は 0
Const VBS_CHARSET_OBSOLETE = 0

Dim objArgs, FilePath(), BaseName()
If init(objArgs, FilePath, BaseName) = 1 Then
        Dim i
        For i = 0 To objArgs.Count - 1
            Call main(objArgs(i), _
                      BaseName(i), _
                      FilePath(i) & "\" & BaseName(i) & ".html", _
                      localfilepath() & HEADER, _
                      localfilepath() & FOODER)
        Next
End If

'**********************************************************
'  初期設定
'**********************************************************
Function init(ByRef objArgs, ByRef FilePath, ByRef BaseName)
        Dim strPath
        Set objArgs = WScript.Arguments

        Dim objFS, objFile
        Set objFS = CreateObject("Scripting.FileSystemObject")

        If objArgs.Count = 0 Then
           MsgBox "ファイルをドラッグしてください"
           init = 0
        Else
           Dim i
           i = 0
           ' 配列の再定義
           ReDim FilePath(objArgs.Count), BaseName(objArgs.Count)
           For Each strPath In objArgs
               Set objFile = objFS.GetFile(strPath)
               BaseName(i) = objFS.GetBaseName(strPath)
               FilePath(i) = objFile.ParentFolder
               i = i + 1
           Next

           Set objFS = Nothing
           init = 1
        End If
End Function

'**********************************************************
'  local file path 取得
'**********************************************************
Function localfilepath()
        Dim Wk_ScriptName, Wk_ScriptFullName
'        Wk_ScriptName     = WScript.ScriptName
        Wk_ScriptName     = MYFILE
        Wk_ScriptFullName = ModulePath()
        localfilepath = Left(Wk_ScriptFullName, Len(Wk_ScriptFullName) _
                        - Len(Wk_ScriptName))
End Function

' makeexeを使うためWScript.ScriptFullNameをラップする
Function ModulePath()
        Dim T, fHandle
        ModulePath = WScript.ScriptFullName
        T = Lcase(ModulePath)
        T = Left(T, len(T) - 4)
        If Right(T, 4) <> ".tmp" then exit function
        On Error resume Next
        Set fHandle = CreateObject("Scripting.FileSystemObject").OpenTextFile(T)
        T = fHandle.ReadLine
        fHandle.Close
        If Err.Number = 0 Then ModulePath = T
End Function

'**********************************************************
'  main関数(Subプロシージャは呼び出し元に値を返さない)
'**********************************************************
Function main(FileIn, BaseName, FileOut, FileHeader, FileFooder)
        Dim text_body, text_header, text_fooder
        Dim title, description, toc
        If VBS_CHARSET_OBSOLETE Then
            text_body   = loadFile(FileIn).ReadAll
            text_header = loadFile(FileHeader).ReadAll
            text_fooder = loadFile(FileFooder).ReadAll
        Else
            text_body   = loadFileUTF8(FileIn)
            text_header = loadFileUTF8(FileHeader)
            text_fooder = loadFileUTF8(FileFooder)
        End If
        main = text_to_html(text_body, title, description, toc)
        main = writeFileUTF8(FileOut, _
                         last_convert(BaseName, text_header, title, description, toc) & _
                         last_convert(BaseName, main, title, description, toc) & _
                         last_convert(BaseName, text_fooder, title, description, toc))
        main = 1
End Function

'**********************************************************
'  変換処理
'**********************************************************
Function text_to_html(str, ByRef title, ByRef  description, ByRef toc)
        Dim objRE
        Set objRE = new RegExp
        objRE.IgnoreCase = True  ' 大文字・小文字を区別しない
        objRE.Global = True      ' 全体を対象とする
        objRE.MultiLine = True   ' 複数行として検索する(5.5以降)

        ' 改行コードの統一
        str = Replace(str, vbCrLf, vbLf)
        str = Replace(str, vbCr, vbLf)

        Call gettitle(str, objRE, title) ' タイトルページの作成(仕様)
        Call getdescription(str, objRE, description) ' og:descriptionの作成(仕様)
        Call commentdel(str, objRE)
        Call horizontal_rule(str, objRE)
        Call headding(str, objRE)
        Call make_table(str, objRE)
        Call trac_preformat(str, objRE)
        Call color(str, objRE)
        Call preformat(str, objRE)
        Call unordered_list(str, objRE)
        Call paragraph(str, objRE)
        Call inline(str, objRE)
        Call link(str, objRE)
        Call img_link(str, objRE)
        Call line_break(str, objRE)  ' 最後に処理

        text_to_html = str
        Set objRE = Nothing
End Function

'**********************************************************
'入力ストリームの生成・設定（テキスト、UTF-8）
'**********************************************************
Function loadFileUTF8(FILE01)

        Dim inStream
        Set inStream = CreateObject("ADODB.Stream")
        inStream.Type = 2             '1:バイナリデータ 2:テキストデータ
        inStream.Charset = "_autodetect_all"    '入力ファイルの文字コード設定（自動判別）
        inStream.Open
        inStream.LoadFromFile FILE01  '入力ファイルを読み込む

        If Err.Number > 0 Then
            MsgBox "Can't open " & FILE01 & " for reading."
            inStream.Close
        End If

        loadFileUTF8 = inStream.ReadText
        inStream.Close
        Set inStream = Nothing

End Function

'**********************************************************
'ファイル書き込み用（テキスト、UTF-8）
'**********************************************************
Function writeFileUTF8(FILE02, src)

        Dim outStream  '出力ストリームの生成・設定
        Set outStream = CreateObject("ADODB.Stream")
        outStream.Type = 2
        outStream.Charset = "UTF-8"  '出力ファイルの文字コード設定
        outStream.Open

        outstream.WriteText src
        outstream.Savetofile FILE02, 2 '1 作成する 2 上書きする
        If Err.Number > 0 then
            MsgBox "Can't open " & FILE02 & " for writing."
            outstream.Close
        End If
        outstream.Close
        Set outstream = Nothing
        writeFileUTF8 = 1

End Function

'**********************************************************
'  ファイル読み込み用
'**********************************************************
Function loadFile(FILE01)
        Dim objFS
        Dim objFile
        'ファイル のオブジェクトを作る
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Err.Number = 0 Then
            If objFS.FileExists(FILE01) Then
               Set objFile = objFS.OpenTextFile(FILE01)
               If Err.Number = 0 Then
               Set loadFile = objFile
               Else
                  MsgBox "ファイルオープンエラー: " & Err.Description
               End If
            Else
               MsgBox "「" & FILE01 & "」が見つかりません"
            End If
        Else
            MsgBox "エラー: " & Err.Description
        End If
End Function

'**********************************************************
'  ファイル書き込み用
'**********************************************************
Function writeFile(FILE02, src)
        Dim objFS
        Dim objFile
        'ファイル のオブジェクトを作る
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Err.Number = 0 Then
            Set objFile = objFS.OpenTextFile(FILE02, 2, True)
            If Err.Number = 0 Then
               objFile.WriteLine src
               objFile.Close
            Else
               MsgBox "ファイルオープンエラー: " & Err.Description
            End If
        Else
            MsgBox "エラー: " & Err.Description
        End If
        writeFile =1
End Function

'**********************************************************
' 水平線
'**********************************************************
Sub horizontal_rule(ByRef str, objRE)
        objRE.Pattern = "^----"
        str = objRE.Replace(str, "<hr>" & vbLf)
End Sub

'**********************************************************
' コメント行の対応
'**********************************************************
Sub commentdel(ByRef str, objRE)
        objRE.Pattern = "^//(.*)"
        str = objRE.Replace(str, "")
End Sub

'**********************************************************
' インライン要素の変換
'**********************************************************
Function inline(ByRef str, objRE)
        ' Italic
        objRE.Pattern = "'''([^']+?)'''"
        str = objRE.Replace(str, "<em>$1</em>")
        ' Bold
        objRE.Pattern = "''([^']+?)''"
        str = objRE.Replace(str, "<strong>$1</strong>")
        ' Strikethrough
        objRE.Pattern = "%%([^']+?)%%"
        str = objRE.Replace(str, "<s>$1</s>")
        ' Under line
        objRE.Pattern = "%%%([^']+?)%%%"
        str = objRE.Replace(str, "<u>$1</u>")
End Function

'**********************************************************
' 目次追加 <!--TOC-->を目次に変更
'**********************************************************
Function insert_toc(ByRef str, objRE, ByRef toc)

        Dim count, K
        Dim Match, Matches
        objRE.pattern = "<h([234])>(?!<a id=""i\d+""> </a>)(.+?)</h\1>$"
        count = 0
        Set Matches = objRE.Execute(str)
        toc = ""
        For Each Match in Matches
            Dim arrayString1
            arrayString1 = Split(str, Match.Value, 2)
            str = arrayString1(0) & _
                  objRE.Replace(Match.Value, "<h$1><a id=""i" & count  & """> </a>$2</h$1>") & _
                  arrayString1(1)
            For K = 1 To objRE.Replace(Match.Value, "$1")
                toc = toc & "-" ' 順不同リストを作成する
            Next
            toc = toc & " <a href=""#i" & count & """>" & objRE.Replace(Match.Value, "$2") & "</a>" & vbLf
            count = count + 1
        Next
        Call unordered_list(toc, objRE)
        objRE.pattern = "<ul>" & vbLf
        toc = objRE.Replace(toc, "<ul class=""toc"">")

End Function

'**********************************************************
' 最終調整
'**********************************************************
Function last_convert(basename, str, title, description, toc)
        Dim objRE
        Set objRE = new RegExp
        objRE.IgnoreCase = True  ' 大文字・小文字を区別しない
        objRE.Global = True      ' 全体を対象とする
        objRE.MultiLine = True   ' 複数行として検索する(5.5以降)

        ' タイトル（pukiwiki用プラグイン title.inc.phpに準拠）
        objRE.pattern = "<!--TITLE-->"
        str = objRE.Replace(str, title)

        ' 見出し（独自実装）
        objRE.pattern = "<!--TOC-->"
        str = objRE.Replace(str, toc)

        ' 更新日時（独自実装）
        objRE.pattern = "<!--UPDATE-->"
        str = objRE.Replace(str, Date)

        ' 作成日時（独自実装）
        objRE.pattern = "<!--DATECREATED-->"
        str = objRE.Replace(str, "")

        ' ファイル名（独自実装）
        objRE.pattern = "<!--BASENAME-->"
        str = objRE.Replace(str, basename)

        ' ページ説明（独自実装）
        objRE.pattern = "<!--DESCRIPTION-->"
        str = objRE.Replace(str, description)

        ' サブBODY
        If SUB_BODY = "" Then
        Else
           Dim text_sub_body
           text_sub_body = loadFile(localfilepath() & SUB_BODY).ReadAll
           objRE.pattern = "<!--SUB_BODY-->"

           str = objRE.Replace(str, text_sub_body)
        End If

        last_convert = str
        Set objRE = Nothing
End Function

'**********************************************************
' <og:description>にセットする文字の変換
'**********************************************************
Sub getdescription(ByRef str, objRE, ByRef header)
        objRE.pattern = "#description\((.*?)\)(;|)"
        Dim Match, Matches
        Set Matches = objRE.Execute(str)
        For Each Match in Matches
            header = Mid(Match.Value, 14, Len(Match.Value) - 14)
        Next
        str = objRE.Replace(str, "")
End Sub

'**********************************************************
' <title>タグにセットする文字の変換
'**********************************************************
Sub gettitle(ByRef str, objRE, ByRef header)
        objRE.pattern = "#title\((.*?)\)(;|)"
        Dim Match, Matches
        Set Matches = objRE.Execute(str)
        For Each Match in Matches
            header = Mid(Match.Value, 8, Len(Match.Value) - 8)
        Next
        str = objRE.Replace(str, "")
End Sub

'**********************************************************
' 表
'**********************************************************
Sub make_table(ByRef str, objRE)
        objRE.pattern = "^\|(.+)"
        Dim Match, Matches
        Set Matches = objRE.Execute(str)
        objRE.Global = False ' 部分対象
        For Each Match in Matches
            objRE.pattern = "^\|(.+)"
            str = objRE.Replace(str, _
                      "<table>" & vbLf & "<tr><td>" & _
                      Join(Split(Mid(Match.Value, 2, _
                      Len(Match.Value) - 2), "|"), "</td><td>") & _
                      "</td></tr>" & vbLf & "</table>")
        Next
        objRE.Global = True  ' 全体を対象
        objRE.pattern = "</table>" & vbLf & "<table>" & vbLf
        str = objRE.Replace(str, "")
        objRE.pattern = "<td>\s+(.+?)\s*</td>"
        str = objRE.Replace(str, "<th>$1</th>")

End Sub


'**********************************************************
' 整形済みテキスト(trac版→pukiwiki版に変換)
'**********************************************************
Sub trac_preformat(ByRef str, objRE)
        objRE.pattern = "^{{{"
        Dim Match, Matches
        Set Matches = objRE.Execute(str)
        For Each Match in Matches
            Dim arrayString1, arrayString2
            arrayString1 = Split(str, "{{{", 2)
            arrayString2 = Split(arrayString1(1), vbLf & "}}}", 2)
            arrayString2(0) = Replace(arrayString2(0), vbLf, vbLf & " ")
            str = arrayString1(0) & arrayString2(0) & arrayString2(1)
        Next
End Sub

'**********************************************************
' 整形済みテキスト(pukiwiki版)
'**********************************************************
Sub preformat(ByRef str, objRE)
        objRE.pattern = "^ (.*)"
        str = objRE.Replace(str, "<pre>$1</pre>")
        objRE.pattern = "</pre>" & vbLf & "<pre>"
        str = objRE.Replace(str, "<!--break-->")
End Sub

'**********************************************************
' 段落
'**********************************************************
Sub paragraph(ByRef str, objRE)
        ' 行頭がインライン要素タグ
        objRE.pattern = "^(<(?:[ib]|em|strong)>.*)"
        str = objRE.Replace(str, "<p>$1</p>")
        objRE.pattern = "^(<(?:a |span).*)"
        str = objRE.Replace(str, "<p>$1</p>")
        ' 行頭が"<"と改行以外
        objRE.pattern = "^([^<" & vbLf & "].*)"
        str = objRE.Replace(str, "<p>$1</p>")
        objRE.pattern = "</p>" & vbLf & "<p>"
        str = objRE.Replace(str, "<br>")
End Sub

'**********************************************************
' 順不同リスト
'**********************************************************
Function unordered_list(ByRef str, objRE)
        ' 4段
        objRE.pattern = "^----(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 3段
        objRE.pattern = "^---(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 2段
        objRE.pattern = "^--(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 1段
        objRE.pattern = "^-(.*)"
        str = objRE.Replace(str, "<ul>" & vbLf & "<li>$1</li>" & _
                            vbLf & "</ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"& vbLf
        str = objRE.Replace(str, "")
End Function

'**********************************************************
' 見出し
'**********************************************************
Sub headding(ByRef str, objRE)
        objRE.pattern = "^\*\*\*\*(.+)"
        str = objRE.Replace(str, "<h4>$1</h4>")
        objRE.pattern = "^\*\*\*(.+)"
        str = objRE.Replace(str, "<h3>$1</h3>")
        objRE.pattern = "^\*\*(.+)"
        str = objRE.Replace(str, "<h2>$1</h2>")
        objRE.pattern = "^\*(.+)"
        str = objRE.Replace(str, "<h1>$1</h1>")
End Sub

'**********************************************************
' 改行の後処理
'**********************************************************
Sub line_break(ByRef str, objRE)
        objRE.pattern = "<!--break-->"
        str = objRE.Replace(str, vbLf)
        objRE.pattern = "<pre>"
        str = objRE.Replace(str, "<pre>" & vbLf)
        objRE.pattern = "</pre>"
        str = objRE.Replace(str, vbLf & "</pre>")
End Sub

'**********************************************************
' リンク作成(>のみ)
'**********************************************************
Sub link(ByRef str, objRE)
        ' ハイパーリンク
        ' 「:」の対応は「http://」で区切られる
        objRE.pattern = "\[\[(.*?)(>)(.*?)\]\]"
        str = objRE.Replace(str, "<a href=""$3"">$1</a>")
End Sub

'**********************************************************
' 画像の貼り付け
'**********************************************************
Sub img_link(ByRef str, objRE)
        ' 画像貼り付け
        objRE.pattern = "(&|#)ref\((.*?).(jpg|gif|png)(,|)(.*?)\)(;|)"
        str = objRE.Replace(str, "<img src=""$2.$3"" alt=""$5"">")
End Sub

'**********************************************************
' 文字色
'**********************************************************
Sub color(ByRef str, objRE)
        ' 文字色
        objRE.pattern = "(&|#|$)color\((.*?)\){(.*?)}(;|)"
        str = objRE.Replace(str, "<span style=""color:$2;"">$3</span>")
End Sub
