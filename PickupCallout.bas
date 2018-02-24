Attribute VB_Name = "PickupCallout"
Option Explicit

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    Dim CurrShape As Shape
    Dim i As Integer
    Dim j As Integer

    Set WriteSheet = ThisWorkbook.Worksheets("data")
    

    For i = 1 To 3
    
        '�Ώۃt�@�C���I�[�v��
        Set ReadBook = Workbooks.Open(Sheets("path").Cells(i, 1))
        
        If ReadBook Is Nothing Then
            Debug.Print "readbook is null"
            Exit Sub
        End If
        
        WriteSheet.Cells(i, 1) = ReadBook.Name
        
        j = 2

        '�Ώۃu�b�N�̊e�V�[�g�ɑ΂��Ă̏���
        For Each ReadSheet In ReadBook.Worksheets
            
            For Each CurrShape In ReadSheet.Shapes
                If CurrShape.Type = msoAutoShape Then
                        
                    If InStr(CurrShape.Name, "Callout") > 0 Then
                        WriteSheet.Cells(i, j) = ReadSheet.Name
                        j = j + 1
                        Exit For
                    End If
                        
                End If
            Next

        Next
        
        '���
        ReadBook.Close False
        Set ReadBook = Nothing
    
    Next


End Sub
