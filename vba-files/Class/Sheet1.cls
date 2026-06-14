VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ==========================================
' 筋トレ記録 エントリーポイント
' ==========================================

Sub 筋トレ記録()
  Call WorkoutLog.AddWorkoutEntry
End Sub

Sub 今日のトレーニング確認()
  Call WorkoutLog.ShowTodaysWorkouts
End Sub

Sub 筋トレ統計()
  Call WorkoutLog.ShowWorkoutStats
End Sub

Sub ログシート初期化()
  Call WorkoutLog.InitWorkoutSheet
End Sub

' ==========================================

Sub trimFormat()
' �Ώۃt�H���_�̃p�X���擾
  Dim folderPath as String 
  folderPath = getFolderPath()
  if folderPath = "" then exit Sub
  Dim confimReturn as Integer
  confimReturn = openCofirmMsgBox("�Ώۃt�@�C���p�X [" & folderPath & "]��OK�H") 
  select case confimReturn
    case vbNo
      exit Sub
  end select
' �Ώۃt�H���_�̃p�X���g�p���āA�z����Excel�t�@�C���̍ċA�I�ɒT��
' ��������Excel�t�@�C���ɑ΂��ăT�u�v���V�[�W���uexecute�v�̏��������s
  Call getFilesRecursive(folderPath)
  
End Sub

Sub openMsgBox(msg as String)
  MsgBox msg
End Sub

Function openCofirmMsgBox(msg as String) as Integer
  openCofirmMsgBox = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "�m�F")
End Function

Function getFolderPath() as String
  with Application.FileDialog(4)
    if .show = false then exit Function
    getFolderPath = .SelectedItems(1)
  end with
End Function

Sub getFilesRecursive(path As String)
  Dim fso As FileSystemObject: Set fso = New FileSystemObject
  Dim objFolder As folder
  Dim objFile As file
    
  ' GetFolder(�t�H���_��).SubFolders�Ńt�H���_�z���̃t�H���_�ꗗ���擾
  For Each objFolder In fso.GetFolder(path).SubFolders
    Call getFilesRecursive(objFolder.path)
  Next
    
  ' GetFolder(�t�H���_��).Files�Ńt�H���_�z���̃t�@�C���ꗗ���擾
  For Each objFile In fso.GetFolder(path).Files
    Dim fileExtension as String
    fileExtension = fso.GetExtensionName(objFile.path)
    if fileExtension = "xlsx" Or fileExtension = "xlsm" then
      Call execute(objFile)
    end if
  Next
End Sub

Sub execute(f As file)
  Debug.Print f.Name
  dim wb as Workbook
  ' IgnoreReadOnlyRecommended=True �͓ǂݎ�萄���̃_�C�A���O��\��
  Set wb = Workbooks.open(Filename:=f.path, IgnoreReadOnlyRecommended:=True)

  call doFormat(wb.Worksheets)

  Call wb.save
  call wb.close
End Sub

Sub doFormat(ws as Sheets)
  For Each Sheet In ws
    '�V�[�g��I��
    Sheet.Select
    '�A�N�e�B�u�Z����A1��
    Sheet.Range("A1").Activate
    '�{����100%
    Sheet.ActiveWindow.Zoom = 100
    '�X�N���[�����ō���
    ActiveWindow.ScrollColumn = 1
    '�X�N���[�����ŏ��
    ActiveWindow.ScrollRow = 1
  Next
  '���[�̃V�[�g���A�N�e�B�u�ɂ���
  ws(1).Activate

End Sub