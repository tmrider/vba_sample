Attribute VB_Name = "GetDataFromOtherFile"
Option Explicit

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    
    Set WriteSheet = ThisWorkbook.Worksheets("data")
    
    Dim i As Integer
    For i = 1 To 3
    
        '�Ώۃt�@�C���I�[�v��
        Set ReadBook = Workbooks.Open(Sheets("path").Cells(i, 1))
        
        If ReadBook Is Nothing Then
            Debug.Print "readbook is null"
            Exit Sub
        End If
        
        '�ΏۃV�[�g�擾
        Set ReadSheet = ReadBook.Worksheets("Sheet1")
        
        If ReadSheet Is Nothing Then
            Debug.Print "readsheet is null"
            Exit Sub
        End If
        
        '�f�[�^�擾
        With WriteSheet
            .Cells(i * 3 + 1, 1) = ReadSheet.Cells(2, 2)
            .Cells(i * 3 + 2, 1) = ReadSheet.Cells(3, 2)
            .Cells(i * 3 + 3, 1) = ReadSheet.Cells(4, 2)
        End With
        
        '���
        ReadBook.Close False
        Set ReadSheet = Nothing
        Set ReadBook = Nothing
    
    Next


End Sub
