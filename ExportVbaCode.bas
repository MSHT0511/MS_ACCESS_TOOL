Attribute VB_Name = "ExportVbaCode"
Option Compare Database
'// TODO ADD COMMENT
Public Function ExportVbaProgramCode()

  Dim vbcmp As Object
  Dim strFileName As String
  Dim strExt As String
  Set dbs = CurrentDb

  savepath = CurrentProject.Path & "\_VBA_" & Mid(dbs.Name, InStrRev(dbs.Name, "\") + 1) & "\"
  
    If Dir(savepath, vbDirectory) = "" Then
        MkDir savepath
    End If
  
  For Each vbcmp In VBE.ActiveVBProject.VBComponents
    With vbcmp
      
      '// 出力先ファイルパス
      strFileName = savepath & .Name
      
      '拡張子を設定
      Select Case .Type
        Case 1    '標準モジュールの場合
          strExt = ".bas"
        Case 2    'クラスモジュールの場合
          strExt = ".cls"
        Case 100  'フォーム/レポートのモジュールの場合
          strExt = ".cls"
      End Select
      'モジュールをエクスポート
      .Export strFileName & strExt
    End With
  Next vbcmp
  
  Set dbs = Nothing
  MsgBox "プログラムコード出力完了"
End Function
