Attribute VB_Name = "CopyFile"
Option Explicit

' 2018-12-07
' �����̃t�@�C���̃R�s�[���쐬����R�[�h

' FileSystemObject���g����
Sub Test()

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' CopyFile( sorce, destination, [overwrite = true] )
    ' overwrite �͕ۑ���Ɋ��������t�@�C��������ꍇ�ɏ㏑�����邩�ǂ���
    ' �R�s�[��� �t�H���_�� + \ �܂Ŏw�肷��Ƃ��̏ꏊ�֓����ŕۑ�
    ' �t�@�C�����܂Ŏw�肷��Ƃ��̖��O�ŕۑ�
    Call FSO.CopyFile("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\")
    Set FSO = Nothing

End Sub


' FileCopy�X�e�[�g�����g���g����
Sub Test2()
    
    ' FileCopy( sorce, destination )
    Call FileCopy("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\NewText.txt")

End Sub
