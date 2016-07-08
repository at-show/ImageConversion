Option Explicit

Dim fso, folder, file, subFolder
Dim objWshShell

Set objWshShell = CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(objWshShell.CurrentDirectory)
filelist(folder)
Set objWshShell = Nothing
Set fso = Nothing




Sub filelist(folder)
  For Each file In folder.Files
    Dim pos
    pos = InStrRev(file.Name, ".")
    If LCase(Mid(file.Name, pos + 1)) = "rar" Or _
       LCase(Mid(file.Name, pos + 1)) = "zip" Or _
       LCase(Mid(file.Name, pos + 1)) = "lzh" Then

       If InStr(file.Name, "[resize]") = 0 And _
          InStr(file.Name, "[inzip]") = 0 Then
         ' リサイズ済み、inzip以外のファイルを対象とする
         Call henkan(folder.Path, file.Name)
         WScript.Sleep(3000)
       End If
    End If
  Next

  For Each subFolder In folder.SubFolders
    filelist(subFolder)
  Next
End Sub


'# 画像変換処理
'# 圧縮ファイルを受け取り以下の処理を行います。
'# １：tempフォルダの作成
'# ２：解凍
'# ３：画像変換
'# ４：圧縮
'# ５：元となった圧縮ファイルの削除
Sub henkan(dir, name)
  Dim exe, filepath, resizefile, skipfile, folder
  exe = "C:\Program Files (x86)\ESTsoft\AlZip\ALZipCon.exe"
  filepath = dir & "\" & name
  resizefile = dir & "\[resize]" & Left(name,InstrRev(name,".") - 1) & ".zip"
  skipfile = dir & "\[inzip]" & Left(name,InstrRev(name,".") - 1) & ".zip"
  WScript.Echo "[" & Now() & "] " & filepath

  objWshShell.CurrentDirectory = dir

  ' tempフォルダが存在すれば削除
  If fso.FolderExists(dir & "\$$temp$$") = True Then
    WScript.Echo "  →temp delete : " & dir & "\$$temp$$"
    Call fso.DeleteFolder(dir & "\$$temp$$", True)
  End If

  ' tempフォルダ作成＆解凍
  fso.CreateFolder("$$temp$$")
  Call objWshShell.Run("""" & exe & """ -x -xf """ & filepath & """ $$temp$$", 0, True)
  Set folder = fso.GetFolder(dir & "\$$temp$$")
  If SearchZip(folder) = False Then
    ' 圧縮ファイルが無い場合のみ実行
    ' 読み取り専用解除
    Free(folder)

    ' tempフォルダへ移動
    objWshShell.CurrentDirectory = dir & "\$$temp$$"

    ' 画像変換
    PdfConvert(folder)
    PicConvert(folder)
    PicDel(folder)

    ' 圧縮
    Call objWshShell.Run("""" & exe & """ -a -nq * """ & resizefile & """", 0, True)

    ' ファイル存在チェック
    If fso.FileExists(resizefile) = True Then
      ' 変換に成功していたら元ファイルを削除
      Call fso.DeleteFile(filepath, True)
    Else
      WScript.Echo "  →Not Exists : " & resizefile
    End If

  Else
    ' 圧縮ファイルが存在した場合はリネームします。
    WScript.Echo "  →skip : " & filepath
    Call fso.MoveFile(filepath, skipfile)
  End If

  ' tempフォルダから移動
  objWshShell.CurrentDirectory = dir

  ' tempフォルダ削除
  On Error Resume Next
  fso.DeleteFolder dir & "\$$temp$$", True
  If Err.Number <> 0 Then
    WScript.Echo "  →Temp Dir Delete Error : " & dir & "\$$temp$$"
  End If
  On Error GoTo 0

End Sub


'# 圧縮ファイル検索処理
'# 圧縮ファイルが含まれているかどうかチェックします。
Function SearchZip(folder)
  SearchZip = False

  For Each file In folder.Files
    Dim pos
    pos = InStrRev(file.Name, ".")
    If LCase(Mid(file.Name, pos + 1)) = "rar" Or _
       LCase(Mid(file.Name, pos + 1)) = "zip" Or _
       LCase(Mid(file.Name, pos + 1)) = "lzh" Then

       SearchZip = True
       WScript.Echo "  →in zip : " & folder.Path & "\" & file.Name
       Exit Function
    End If
  Next

  '# 再起呼び出し
  For Each subFolder In folder.SubFolders
    If SearchZip(subFolder) = True Then
      SearchZip = True
      Exit Function
    End If
  Next
End Function


'# 画像削除処理
'# [.jpg]以外の画像を削除します。
Sub PicDel(folder)
  For Each file In folder.Files
    Dim pos
    pos = InStrRev(file.Name, ".")
    If LCase(Mid(file.Name, pos + 1)) <> "jpg" Then

       WScript.Echo "  →Del : " & folder.Path & "\" & file.Name
       Call fso.DeleteFile(folder.Path & "\" & file.Name, True)
    End If
  Next

  '# 再起呼び出し
  For Each subFolder In folder.SubFolders
    PicDel(subFolder)
  Next
End Sub


'# 読み取り専用解除処理
Sub Free(folder)
  For Each file In folder.Files
    If file.Attributes And 1 Then
      WScript.Echo "  →Free : " & folder.Path & "\" & file.Name
      file.Attributes = file.Attributes And &HFE
    End If
  Next

  '# 再起呼び出し
  For Each subFolder In folder.SubFolders
    Free(subFolder)
  Next
End Sub

'# PDF変換処理
Sub PdfConvert(folder)
  Dim resize, conv, pos, objExec
  conv = "L:\soft\xpdfbin-win-3.03\bin64\pdftoppm.exe"

  '# PDFの場合ppmに変換する
  For Each file In folder.Files
    pos = InStrRev(file.Name, ".")

    If LCase(Mid(file.Name, pos + 1)) = "pdf" Then
      WScript.Echo "  →pdf to ppm : " & conv & " """ & folder.Path & "\" & file.Name & """ ""out"""
      Call objWshShell.Run(conv & " """ & folder.Path & "\" & file.Name & """ ""out""", 0, True)
    End If
  Next

  '# 再起呼び出し
  For Each subFolder In folder.SubFolders
    PdfConvert(subFolder)
  Next
End Sub

'# 画像変換処理
Sub PicConvert(folder)
  Dim resize, conv, pos, objExec
  resize = "C:\Program Files\ImageMagick-6.8.1-Q16\mogrify.exe"

  '# 画像ファイルを変換する
  For Each file In folder.Files
    pos = InStrRev(file.Name, ".")

    If LCase(Mid(file.Name, pos + 1)) = "png" Or _
       LCase(Mid(file.Name, pos + 1)) = "ppm" Or _
       LCase(Mid(file.Name, pos + 1)) = "bmp" Or _
       LCase(Mid(file.Name, pos + 1)) = "jpg" Or _
       LCase(Mid(file.Name, pos + 1)) = "jpeg" Or _
       LCase(Mid(file.Name, pos + 1)) = "gif" Then

       Set objExec = objWshShell.Exec("""" & resize & """ -quality 75 -resize 2000x1024 -format jpg """ & folder.Path & "\" & file.Name & """")
       Do While objExec.Status = 0
         WScript.Sleep 10
       Loop
       
    End If
  Next

  '# 再起呼び出し
  For Each subFolder In folder.SubFolders
    PicConvert(subFolder)
  Next
End Sub
