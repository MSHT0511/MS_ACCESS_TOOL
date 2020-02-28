Attribute VB_Name = "_4Developer-Helper-OutputVbaCodes"
Option Compare Database

Public Function OutputCodes()

  Dim vbcmp As Object
  Dim strFileName As String
  Dim strExt As String
  Set dbs = CurrentDb

  savepath = CurrentProject.path & "\_VBA_" & mId(dbs.Name, InStrRev(dbs.Name, "\") + 1) & "\"
  
    If Dir(savepath, vbDirectory) = "" Then
        MkDir savepath
    End If
  
  For Each vbcmp In VBE.ActiveVBProject.VBComponents
    With vbcmp
      
      '// �o�͐�t�@�C���p�X
      strFileName = savepath & .Name
      
      '�g���q��ݒ�
      Select Case .Type
        Case 1    '�W�����W���[���̏ꍇ
          strExt = ".bas"
        Case 2    '�N���X���W���[���̏ꍇ
          strExt = ".cls"
        Case 100  '�t�H�[��/���|�[�g�̃��W���[���̏ꍇ
          strExt = ".cls"
      End Select
      '���W���[�����G�N�X�|�[�g
      .Export strFileName & strExt
    End With
  Next vbcmp
  
  Set dbs = Nothing
MsgBox "�v���O�����R�[�h�o�͊���"
End Function
