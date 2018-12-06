Attribute VB_Name = "ReadOtherExcelFile"
Option Explicit

' 2018-12-07
' 他のエクセルファイルからデータを読み出すコード

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    
    '読み取ったデータの書き込み先
    Set WriteSheet = ThisWorkbook.Worksheets("data")
    
    Dim FilePath As String
    Dim i As Integer
    For i = 1 To 3
    
        'データを読み取りたいファイルのパス
        FilePath = ThisWorkbook.Worksheets("path").Cells(i, 1)
    
        If Dir(FilePath) = "" Then
            Debug.Print FilePath + " is nothing"
            Exit Sub
        End If
    
        '対象ファイルオープン
        Set ReadBook = Workbooks.Open(FilePath)
        
        '対象シート取得
        Set ReadSheet = ReadBook.Worksheets("Sheet1")
        
        If ReadSheet Is Nothing Then
            Debug.Print "readsheet is null"
            Exit Sub
        End If
        
        'データ取得
        With WriteSheet
            .Cells(i * 3 + 1, 1) = ReadSheet.Cells(2, 2)
            .Cells(i * 3 + 2, 1) = ReadSheet.Cells(3, 2)
            .Cells(i * 3 + 3, 1) = ReadSheet.Cells(4, 2)
        End With
        
        '解放
        'Falseで保存しないで閉じる
        ReadBook.Close False
        Set ReadSheet = Nothing
        Set ReadBook = Nothing
    
    Next

End Sub
