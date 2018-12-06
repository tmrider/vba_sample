Attribute VB_Name = "ReadOtherExcelFile"
Option Explicit

' 2018-12-07
' ���̃G�N�Z���t�@�C������f�[�^��ǂݏo���R�[�h

Sub Main()

    Dim WriteSheet As Worksheet
    Dim ReadBook As Workbook
    Dim ReadSheet As Worksheet
    
    '�ǂݎ�����f�[�^�̏������ݐ�
    Set WriteSheet = ThisWorkbook.Worksheets("data")
    
    Dim FilePath As String
    Dim i As Integer
    For i = 1 To 3
    
        '�f�[�^��ǂݎ�肽���t�@�C���̃p�X
        FilePath = ThisWorkbook.Worksheets("path").Cells(i, 1)
    
        If Dir(FilePath) = "" Then
            Debug.Print FilePath + " is nothing"
            Exit Sub
        End If
    
        '�Ώۃt�@�C���I�[�v��
        Set ReadBook = Workbooks.Open(FilePath)
        
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
        'False�ŕۑ����Ȃ��ŕ���
        ReadBook.Close False
        Set ReadSheet = Nothing
        Set ReadBook = Nothing
    
    Next

End Sub
