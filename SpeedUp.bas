Attribute VB_Name = "SpeedUp"
'��������\���̤���I�������m�ۤSleep�ȂǐF�X

Option Explicit

' Sleep�֐����g����
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'�\���́i���[�U�[��`�^�j
Type Person
    Name As String
    Age As Long
End Type


Sub Main()

On Error GoTo CATCH
    
    '����������
    Application.ScreenUpdating = False '�`���~
    Application.EnableEvents = False '�C�x���g�}��
    Application.Calculation = xlCalculationManual '�蓮�v�Z

    '�ŏI�s�����m���āA���̃T�C�Y�̃������𓮓I�m��
    Dim Arr() As Person
    Dim LastRowNum As Long
    LastRowNum = ThisWorkbook.Sheets("Sheet1").Range("A100").End(xlUp).Row
    ReDim Arr(LastRowNum)
    'MsgBox LBound(Arr) & vbCrLf & UBound(Arr)

    '�V�[�g�̃f�[�^���������Ɋi�[
    Dim i As Long
    For i = 2 To UBound(Arr)
        Arr(i).Name = ThisWorkbook.Sheets("Sheet1").Cells(i, 1)
        Arr(i).Age = ThisWorkbook.Sheets("Sheet1").Cells(i, 2)
        Debug.Print Arr(i).Name & Arr(i).Age
        'DoEvents   'OS�ɏ�����n����
    Next

    'MsgBox "�I��"
    GoTo FINAL
CATCH:
    MsgBox "�G���[�I��"
FINAL:
    '�������̉���
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

