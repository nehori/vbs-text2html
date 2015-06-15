'**
'* Copyright (c) 2015 Kazutaka Yasuda
'* Released under the MIT license
'* http://opensource.org/licenses/mit-license.php
'*
'**********************************************************
'  �X�V
'     v2.5a 2011. 4.20    <!--DATECREATED-->
'     v2.4a 2011.11.01    <!--DESCRIPTION-->�@�\�ǉ�
'     v2.3a 2011.10.26    <!--BASENAME-->�@�\�ǉ�
'     v2.2a 2011. 4.20    <!--SUB_BODY-->�@�\�b��ǉ�
'     v2.1a 2011. 3.14    �ڎ��ǉ������ǉ�
'     v2.0a 2011. 3.13    �ǂݍ��ݕ����R�[�h�������ʒǉ�
'     v1.6a 2011. 2.27    �����J���[�����X�V
'     v1.5a 2010. 7.31    <!--break-->�����X�V
'     v1.4a 2010. 7.28    ���`�ς݃e�L�X�g<pre>�ǉ�
'     v1.3a 2010. 6. 8    ��������(Strikethrough)�ǉ�
'**********************************************************

' �Ȍ�̈Öق̕ϐ���`�������Ȃ�
Option Explicit

' ���L�̃t�@�C���͗p�ӂ��Ă�������
Const HEADER   = "/mdl/header.mdl"
Const FOODER   = "/mdl/fooder.mdl"
Const SUB_BODY = "/mdl/sub_body.mdl"
' EXE�ϊ��p
Const MYFILE = "text2html.vbs"
' ��{�� 0
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
'  �����ݒ�
'**********************************************************
Function init(ByRef objArgs, ByRef FilePath, ByRef BaseName)
        Dim strPath
        Set objArgs = WScript.Arguments

        Dim objFS, objFile
        Set objFS = CreateObject("Scripting.FileSystemObject")

        If objArgs.Count = 0 Then
           MsgBox "�t�@�C�����h���b�O���Ă�������"
           init = 0
        Else
           Dim i
           i = 0
           ' �z��̍Ē�`
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
'  local file path �擾
'**********************************************************
Function localfilepath()
        Dim Wk_ScriptName, Wk_ScriptFullName
'        Wk_ScriptName     = WScript.ScriptName
        Wk_ScriptName     = MYFILE
        Wk_ScriptFullName = ModulePath()
        localfilepath = Left(Wk_ScriptFullName, Len(Wk_ScriptFullName) _
                        - Len(Wk_ScriptName))
End Function

' makeexe���g������WScript.ScriptFullName�����b�v����
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
'  main�֐�(Sub�v���V�[�W���͌Ăяo�����ɒl��Ԃ��Ȃ�)
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
'  �ϊ�����
'**********************************************************
Function text_to_html(str, ByRef title, ByRef  description, ByRef toc)
        Dim objRE
        Set objRE = new RegExp
        objRE.IgnoreCase = True  ' �啶���E����������ʂ��Ȃ�
        objRE.Global = True      ' �S�̂�ΏۂƂ���
        objRE.MultiLine = True   ' �����s�Ƃ��Č�������(5.5�ȍ~)

        ' ���s�R�[�h�̓���
        str = Replace(str, vbCrLf, vbLf)
        str = Replace(str, vbCr, vbLf)

        Call gettitle(str, objRE, title) ' �^�C�g���y�[�W�̍쐬(�d�l)
        Call getdescription(str, objRE, description) ' og:description�̍쐬(�d�l)
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
        Call line_break(str, objRE)  ' �Ō�ɏ���

        text_to_html = str
        Set objRE = Nothing
End Function

'**********************************************************
'���̓X�g���[���̐����E�ݒ�i�e�L�X�g�AUTF-8�j
'**********************************************************
Function loadFileUTF8(FILE01)

        Dim inStream
        Set inStream = CreateObject("ADODB.Stream")
        inStream.Type = 2             '1:�o�C�i���f�[�^ 2:�e�L�X�g�f�[�^
        inStream.Charset = "_autodetect_all"    '���̓t�@�C���̕����R�[�h�ݒ�i�������ʁj
        inStream.Open
        inStream.LoadFromFile FILE01  '���̓t�@�C����ǂݍ���

        If Err.Number > 0 Then
            MsgBox "Can't open " & FILE01 & " for reading."
            inStream.Close
        End If

        loadFileUTF8 = inStream.ReadText
        inStream.Close
        Set inStream = Nothing

End Function

'**********************************************************
'�t�@�C���������ݗp�i�e�L�X�g�AUTF-8�j
'**********************************************************
Function writeFileUTF8(FILE02, src)

        Dim outStream  '�o�̓X�g���[���̐����E�ݒ�
        Set outStream = CreateObject("ADODB.Stream")
        outStream.Type = 2
        outStream.Charset = "UTF-8"  '�o�̓t�@�C���̕����R�[�h�ݒ�
        outStream.Open

        outstream.WriteText src
        outstream.Savetofile FILE02, 2 '1 �쐬���� 2 �㏑������
        If Err.Number > 0 then
            MsgBox "Can't open " & FILE02 & " for writing."
            outstream.Close
        End If
        outstream.Close
        Set outstream = Nothing
        writeFileUTF8 = 1

End Function

'**********************************************************
'  �t�@�C���ǂݍ��ݗp
'**********************************************************
Function loadFile(FILE01)
        Dim objFS
        Dim objFile
        '�t�@�C�� �̃I�u�W�F�N�g�����
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Err.Number = 0 Then
            If objFS.FileExists(FILE01) Then
               Set objFile = objFS.OpenTextFile(FILE01)
               If Err.Number = 0 Then
               Set loadFile = objFile
               Else
                  MsgBox "�t�@�C���I�[�v���G���[: " & Err.Description
               End If
            Else
               MsgBox "�u" & FILE01 & "�v��������܂���"
            End If
        Else
            MsgBox "�G���[: " & Err.Description
        End If
End Function

'**********************************************************
'  �t�@�C���������ݗp
'**********************************************************
Function writeFile(FILE02, src)
        Dim objFS
        Dim objFile
        '�t�@�C�� �̃I�u�W�F�N�g�����
        Set objFS = CreateObject("Scripting.FileSystemObject")
        If Err.Number = 0 Then
            Set objFile = objFS.OpenTextFile(FILE02, 2, True)
            If Err.Number = 0 Then
               objFile.WriteLine src
               objFile.Close
            Else
               MsgBox "�t�@�C���I�[�v���G���[: " & Err.Description
            End If
        Else
            MsgBox "�G���[: " & Err.Description
        End If
        writeFile =1
End Function

'**********************************************************
' ������
'**********************************************************
Sub horizontal_rule(ByRef str, objRE)
        objRE.Pattern = "^----"
        str = objRE.Replace(str, "<hr>" & vbLf)
End Sub

'**********************************************************
' �R�����g�s�̑Ή�
'**********************************************************
Sub commentdel(ByRef str, objRE)
        objRE.Pattern = "^//(.*)"
        str = objRE.Replace(str, "")
End Sub

'**********************************************************
' �C�����C���v�f�̕ϊ�
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
' �ڎ��ǉ� <!--TOC-->��ڎ��ɕύX
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
                toc = toc & "-" ' ���s�����X�g���쐬����
            Next
            toc = toc & " <a href=""#i" & count & """>" & objRE.Replace(Match.Value, "$2") & "</a>" & vbLf
            count = count + 1
        Next
        Call unordered_list(toc, objRE)
        objRE.pattern = "<ul>" & vbLf
        toc = objRE.Replace(toc, "<ul class=""toc"">")

End Function

'**********************************************************
' �ŏI����
'**********************************************************
Function last_convert(basename, str, title, description, toc)
        Dim objRE
        Set objRE = new RegExp
        objRE.IgnoreCase = True  ' �啶���E����������ʂ��Ȃ�
        objRE.Global = True      ' �S�̂�ΏۂƂ���
        objRE.MultiLine = True   ' �����s�Ƃ��Č�������(5.5�ȍ~)

        ' �^�C�g���ipukiwiki�p�v���O�C�� title.inc.php�ɏ����j
        objRE.pattern = "<!--TITLE-->"
        str = objRE.Replace(str, title)

        ' ���o���i�Ǝ������j
        objRE.pattern = "<!--TOC-->"
        str = objRE.Replace(str, toc)

        ' �X�V�����i�Ǝ������j
        objRE.pattern = "<!--UPDATE-->"
        str = objRE.Replace(str, Date)

        ' �쐬�����i�Ǝ������j
        objRE.pattern = "<!--DATECREATED-->"
        str = objRE.Replace(str, "")

        ' �t�@�C�����i�Ǝ������j
        objRE.pattern = "<!--BASENAME-->"
        str = objRE.Replace(str, basename)

        ' �y�[�W�����i�Ǝ������j
        objRE.pattern = "<!--DESCRIPTION-->"
        str = objRE.Replace(str, description)

        ' �T�uBODY
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
' <og:description>�ɃZ�b�g���镶���̕ϊ�
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
' <title>�^�O�ɃZ�b�g���镶���̕ϊ�
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
' �\
'**********************************************************
Sub make_table(ByRef str, objRE)
        objRE.pattern = "^\|(.+)"
        Dim Match, Matches
        Set Matches = objRE.Execute(str)
        objRE.Global = False ' �����Ώ�
        For Each Match in Matches
            objRE.pattern = "^\|(.+)"
            str = objRE.Replace(str, _
                      "<table>" & vbLf & "<tr><td>" & _
                      Join(Split(Mid(Match.Value, 2, _
                      Len(Match.Value) - 2), "|"), "</td><td>") & _
                      "</td></tr>" & vbLf & "</table>")
        Next
        objRE.Global = True  ' �S�̂�Ώ�
        objRE.pattern = "</table>" & vbLf & "<table>" & vbLf
        str = objRE.Replace(str, "")
        objRE.pattern = "<td>\s+(.+?)\s*</td>"
        str = objRE.Replace(str, "<th>$1</th>")

End Sub


'**********************************************************
' ���`�ς݃e�L�X�g(trac�Ł�pukiwiki�łɕϊ�)
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
' ���`�ς݃e�L�X�g(pukiwiki��)
'**********************************************************
Sub preformat(ByRef str, objRE)
        objRE.pattern = "^ (.*)"
        str = objRE.Replace(str, "<pre>$1</pre>")
        objRE.pattern = "</pre>" & vbLf & "<pre>"
        str = objRE.Replace(str, "<!--break-->")
End Sub

'**********************************************************
' �i��
'**********************************************************
Sub paragraph(ByRef str, objRE)
        ' �s�����C�����C���v�f�^�O
        objRE.pattern = "^(<(?:[ib]|em|strong)>.*)"
        str = objRE.Replace(str, "<p>$1</p>")
        objRE.pattern = "^(<(?:a |span).*)"
        str = objRE.Replace(str, "<p>$1</p>")
        ' �s����"<"�Ɖ��s�ȊO
        objRE.pattern = "^([^<" & vbLf & "].*)"
        str = objRE.Replace(str, "<p>$1</p>")
        objRE.pattern = "</p>" & vbLf & "<p>"
        str = objRE.Replace(str, "<br>")
End Sub

'**********************************************************
' ���s�����X�g
'**********************************************************
Function unordered_list(ByRef str, objRE)
        ' 4�i
        objRE.pattern = "^----(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 3�i
        objRE.pattern = "^---(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 2�i
        objRE.pattern = "^--(.*)"
        str = objRE.Replace(str, "<ul><li>$1</li></ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"
        str = objRE.Replace(str, "")
        objRE.pattern = "^-(.*)" & vbLf & "<ul>"
        str = objRE.Replace(str, "-$1<ul>")
        ' 1�i
        objRE.pattern = "^-(.*)"
        str = objRE.Replace(str, "<ul>" & vbLf & "<li>$1</li>" & _
                            vbLf & "</ul>")
        objRE.pattern = "</ul>" & vbLf & "<ul>"& vbLf
        str = objRE.Replace(str, "")
End Function

'**********************************************************
' ���o��
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
' ���s�̌㏈��
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
' �����N�쐬(>�̂�)
'**********************************************************
Sub link(ByRef str, objRE)
        ' �n�C�p�[�����N
        ' �u:�v�̑Ή��́uhttp://�v�ŋ�؂���
        objRE.pattern = "\[\[(.*?)(>)(.*?)\]\]"
        str = objRE.Replace(str, "<a href=""$3"">$1</a>")
End Sub

'**********************************************************
' �摜�̓\��t��
'**********************************************************
Sub img_link(ByRef str, objRE)
        ' �摜�\��t��
        objRE.pattern = "(&|#)ref\((.*?).(jpg|gif|png)(,|)(.*?)\)(;|)"
        str = objRE.Replace(str, "<img src=""$2.$3"" alt=""$5"">")
End Sub

'**********************************************************
' �����F
'**********************************************************
Sub color(ByRef str, objRE)
        ' �����F
        objRE.pattern = "(&|#|$)color\((.*?)\){(.*?)}(;|)"
        str = objRE.Replace(str, "<span style=""color:$2;"">$3</span>")
End Sub
