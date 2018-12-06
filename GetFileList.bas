Attribute VB_Name = "GetFileList"
Option Explicit

'2018-12-07
'�w�肵���t�H���_���̃t�@�C�����ꗗ���擾����R�[�h
'�ċA�I�ɉ��̊K�w�̃t�H���_���T�����ă��X�g�A�b�v����

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
  
    '�T�u�t�H���_�����ɍs��
    For Each Folder In Folder.subfolders
        Call GetFileListInFolder(Folder.Path, rRow)
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
