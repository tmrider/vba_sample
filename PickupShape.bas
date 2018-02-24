Attribute VB_Name = "PickupShape"
Option Explicit

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    Dim Curr As Long
    Dim CurrShape As Shape
    
    Curr = 1
    Set WriteSheet = ThisWorkbook.Worksheets("data")
    
    Dim i As Integer
    For i = 1 To 3
    
        '�Ώۃt�@�C���I�[�v��
        Set ReadBook = Workbooks.Open(Sheets("path").Cells(i, 1))
        
        If ReadBook Is Nothing Then
            Debug.Print "readbook is null"
            Exit Sub
        End If
        
        '�Ώۃu�b�N�̊e�V�[�g�ɑ΂��Ă̏���
        For Each ReadSheet In ReadBook.Worksheets
            With WriteSheet
                
                '�V�F�C�v�̐��𒲂ׂ�
'                .Cells(Curr, 1) = ReadBook.Name
'                .Cells(Curr, 2) = ReadSheet.Name
'                .Cells(Curr, 3) = ReadSheet.Shapes.Count
'                Curr = Curr + 1
        
                '�V�F�C�v�̃^�C�v�ƒ��̃e�L�X�g�������o��
                If ReadSheet.Shapes.Count > 0 Then
                    For Each CurrShape In ReadSheet.Shapes
                        If CurrShape.TextFrame2.HasText Then
                            .Cells(Curr, 1) = ReadBook.Name
                            .Cells(Curr, 2) = ReadSheet.Name
                            .Cells(Curr, 3) = Left(CurrShape.Name, Len(CurrShape.Name) - 2)
                            .Cells(Curr, 4) = CurrShape.TextFrame2.TextRange.Text
                            .Cells(Curr, 5) = CurrShape.Type
                            Curr = Curr + 1
                        End If
                    Next
                End If
                
            End With
        Next
        
        '���
        ReadBook.Close False
        Set ReadBook = Nothing
    
    Next


End Sub
