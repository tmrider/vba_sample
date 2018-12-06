Attribute VB_Name = "EditOtherExcelFile"
Option Explicit

' 2018-12-07
' 他のエクセルファイルを編集するコードサンプル
Sub Main()

    Dim TargetBookPath As String
        TargetBookPath = "D:\Test\ForEdit.xlsx"
        
    If Dir(TargetBookPath) = "" Then
        Debug.Print "target book is nothing"
        Exit Sub
    End If
    
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(TargetBookPath)
    
    ' 開いたエクセルファイルを編集する
    TargetBook.Worksheets(1).Cells(1, 1) = "Hello World !!"

    ' Trueは上書き保存。第2引数でファイル名を指定すると別名保存。
    ' Falseは保存せずに閉じる
    TargetBook.Close True
    Set TargetBook = Nothing
    
End Sub
