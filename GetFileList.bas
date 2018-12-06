Attribute VB_Name = "GetFileList"
Option Explicit

'2018-12-07
'指定したフォルダ内のファイル名一覧を取得するコード
'再帰的に下の階層のフォルダも探索してリストアップする

'これを実行する
Sub ListUp()

    Dim TargetFolder As String
    TargetFolder = "D:\Test"
    MsgBox IIf(GetFileListInFolder(TargetFolder, 1) = True, "成功", "失敗")
    
End Sub

'フォルダ内のファイル名を再帰的に取得してシートに書き込む
Function GetFileListInFolder(TargetFolder As String, Optional rRow As Integer) As Boolean
    
    Dim FSO As Object
    Dim Folder As Object
    Dim File As Object
  
    On Error GoTo errHnd
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.Getfolder(TargetFolder)
    Set File = Folder.Files

    For Each File In Folder.Files
        
        rRow = rRow + 1
        Cells(rRow, 1) = rRow - 1               'ファイル番号
        Cells(rRow, 2) = File.Name              'ファイル名
        Cells(rRow, 3) = File.Size              'ファイルサイズ
        Cells(rRow, 4) = File.DateLastModified  '最終更新日時
        Cells(rRow, 5) = File.Path              'フルパス
        
    Next
  
    'サブフォルダを見に行く
    For Each Folder In Folder.subfolders
        Call GetFileListInFolder(Folder.Path, rRow)
    Next
  
    Set FSO = Nothing
    Set Folder = Nothing
    Set File = Nothing
    GetFileListInFolder = True
    
    Exit Function
    
errHnd:

    Debug.Print Err.Number, Err.Description
    Set FSO = Nothing
    Set Folder = Nothing
    Set File = Nothing
    GetFileListInFolder = False
    
End Function
