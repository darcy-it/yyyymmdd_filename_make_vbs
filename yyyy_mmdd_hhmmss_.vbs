' 現在の日時を取得
Dim currentDate
currentDate = Now()

' 日時を yyyy_mmdd_hhmmss 形式にフォーマット
Dim fileName
fileName = Year(currentDate) & "_" & Right("0" & Month(currentDate), 2) & Right("0" & Day(currentDate), 2) & "_" & Right("0" & Hour(currentDate), 2) & Right("0" & Minute(currentDate), 2) & Right("0" & Second(currentDate), 2) & "_.txt"

' デスクトップのパスを取得
Dim desktopPath
desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

' ファイルシステムオブジェクトの作成
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' デスクトップに一時的なテキストファイルを作成
Dim tempFilePath
tempFilePath = desktopPath & "\" & fileName
Dim outputFile
' Set outputFile = fso.CreateTextFile(tempFilePath, True)
Set outputFile = fso.CreateTextFile(desktopPath & "\" & fileName, True)


' ファイルに内容を書き込み
' outputFile.WriteLine("ここにテキストファイルの内容を記述します。")

' ファイルを閉じる
outputFile.Close

' 現在のディレクトリのパスを取得
Dim currentDirectory
currentDirectory = fso.GetAbsolutePathName(".")

' デスクトップから現在のディレクトリにファイルをコピー
fso.CopyFile tempFilePath, currentDirectory & "\", True

' デスクトップの一時ファイルを削除
' これによりデスクトップにファイルが残らないようにする
fso.DeleteFile(tempFilePath)

' オブジェクトの解放
Set outputFile = Nothing
Set fso = Nothing
