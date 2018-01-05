Attribute VB_Name = "FileControlSample"
Option Explicit

'��������s����
Sub ListUp()

    Dim TargetFolder As String
    TargetFolder = "D:\Test"
    MsgBox IIf(GetFileListInFolder(TargetFolder, 1) = True, "����", "���s")
    
End Sub

'�t�H���_���̃t�@�C�������ċA�I�Ɏ擾���ăV�[�g�ɏ�������
Function GetFileListInFolder(TargetFolder As String, Optional rRow As Integer) As Boolean
    
    Dim FSO As Object
    Dim Folder As Object
    Dim File As Object
  
    On Error GoTo errHnd
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.Getfolder(TargetFolder)
    Set File = Folder.Files

    For Each File In Folder.Files
        
        rRow = rRow + 1
        Cells(rRow, 1) = rRow - 1               '�t�@�C���ԍ�
        Cells(rRow, 2) = File.Name              '�t�@�C����
        Cells(rRow, 3) = File.Size              '�t�@�C���T�C�Y
        Cells(rRow, 4) = File.DateLastModified  '�ŏI�X�V����
        Cells(rRow, 5) = File.Path              '�t���p�X
        
    Next
  
    For Each Folder In Folder.subfolders
        Call GetFileListInFolder(Folder.Path, rRow) '���T�u�t�H���_�����ɍs���܂�
    Next
  
    Set FSO = Nothing
    Set Folder = Nothing
    Set File = Nothing
    GetFileListInFolder = True
    
    Exit Function
    
errHnd:

    Debug.Print Err.Number, Err.Description
    Set FSO = Nothing
    Set Folder = Nothing
    Set File = Nothing
    GetFileListInFolder = False
    
End Function


' FileSystemObject���g����
Sub FileCopySample1()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' �R�s�[��� �t�H���_�� + \ �܂Ŏw��
    Call FSO.CopyFile("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\")
    Set FSO = Nothing
End Sub

' FileCopy�X�e�[�g�����g���g����
Sub FileCopySample2()
    ' �R�s�[��Ƀt�@�C�������w�肵�Ȃ���΂Ȃ�Ȃ�
    Call FileCopy("D:\Test\Sub2\Text3.txt", "D:\Test\ForCopy\NewText.txt")
End Sub

