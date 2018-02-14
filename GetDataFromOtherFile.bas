Attribute VB_Name = "GetDataFromOtherFile"
Option Explicit

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    
    Set WriteSheet = ThisWorkbook.Worksheets("data")
    
    Dim i As Integer
    For i = 1 To 3
    
        '対象ファイルオープン
        Set ReadBook = Workbooks.Open(Sheets("path").Cells(i, 1))
        
        If ReadBook Is Nothing Then
            Debug.Print "readbook is null"
            Exit Sub
        End If
        
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
        ReadBook.Close False
        Set ReadSheet = Nothing
        Set ReadBook = Nothing
    
    Next


End Sub
