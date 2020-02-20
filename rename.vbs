Option Explicit
 
Dim objFileSys
Dim objFolder
Dim objFile
Dim objOutputTextStream
Dim strFilePathFrom
Dim strFilePathTo
Dim i
 
'ファイルシステムを扱うオブジェクトを作成
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'ログ出力用 TextStream オブジェクトを作成
'第2引数は 1 ：読み取り、2 ：上書き、3 ：追記。
Set objOutputTextStream = objFileSys.OpenTextFile("log.txt", 2, True)
 
'c:\temp フォルダのオブジェクトを取得
Set objFolder = objFileSys.GetFolder("C:\Users\Owner\Pictures\images\")
 
'FolderオブジェクトのFilesプロパティからFileオブジェクトを取得
i = 1
For Each objFile In objFolder.Files
    'ファイル名を取得し、ログファイルに出力
    objOutputTextStream.WriteLine objFile.Name
    
    'コピー元のファイルのパスを指定
    strFilePathFrom = "C:\Users\Owner\Pictures\images\" & objFile.Name
    strFilePathTo   = "C:\Users\Owner\Pictures\images\test\" & CStr(i) & ".png"
 
    'エラー発生時にも処理を続行するよう設定
    On Error Resume Next
 
    'コピー先に同名のファイルが無い場合のみ上書き
    Call objFileSys.CopyFile(strFilePathFrom, strFilePathTo, true)
 
    'エラーになった場合の処理
    If Err.Number <> 0 Then

    'エラー情報をクリアする。
    Err.Clear
    End If
 
    '「On Error Resume Next」を解除
    On Error Goto 0

    i = i + 1
Next
 
'TextStream は Close を忘れずに
objOutputTextStream.Close
Set objOutputTextStream = Nothing
Set objFolder  = Nothing
Set objFileSys = Nothing

 