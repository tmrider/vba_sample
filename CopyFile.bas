Attribute VB_Name = "CopyFile"
Option Explicit

' 2018-12-07
' 既存のファイルのコピーを作成するコード

' FileSystemObjectを使う例
Sub Test()

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' CopyFile( sorce, destination, [overwrite = true] )
    ' overwrite は保存先に既存同名ファイルがある場合に上書きするかどうか
    ' コピー先は フォルダ名 + \ まで指定するとその場所へ同名で保存
    ' ファイル名まで指定するとその名前で保存
    Call FSO.CopyFile("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\")
    Set FSO = Nothing

End Sub


' FileCopyステートメントを使う例
Sub Test2()
    
    ' FileCopy( sorce, destination )
    Call FileCopy("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\NewText.txt")

End Sub
