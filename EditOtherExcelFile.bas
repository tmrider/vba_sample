Attribute VB_Name = "EditOtherExcelFile"
Option Explicit

' 2018-12-07
' ���̃G�N�Z���t�@�C����ҏW����R�[�h�T���v��
Sub Main()

    Dim TargetBookPath As String
        TargetBookPath = "D:\Test\ForEdit.xlsx"
        
    If Dir(TargetBookPath) = "" Then
        Debug.Print "target book is nothing"
        Exit Sub
    End If
    
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(TargetBookPath)
    
    ' �J�����G�N�Z���t�@�C����ҏW����
    TargetBook.Worksheets(1).Cells(1, 1) = "Hello World !!"

    ' True�͏㏑���ۑ��B��2�����Ńt�@�C�������w�肷��ƕʖ��ۑ��B
    ' False�͕ۑ������ɕ���
    TargetBook.Close True
    Set TargetBook = Nothing
    
End Sub
