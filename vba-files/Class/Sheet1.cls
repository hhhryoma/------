VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub trimFormat()
' 対象フォルダのパスを取得
  Dim folderPath as String 
  folderPath = getFolderPath()
  if folderPath = "" then exit Sub
  Dim confimReturn as Integer
  confimReturn = openCofirmMsgBox("対象ファイルパス [" & folderPath & "]でOK？") 
  select case confimReturn
    case vbNo
      exit Sub
  end select
' 対象フォルダのパスを使用して、配下のExcelファイルの再帰的に探索
' 見つかったExcelファイルに対してサブプロシージャ「execute」の処理を実行
  Call getFilesRecursive(folderPath)
  
End Sub

Sub openMsgBox(msg as String)
  MsgBox msg
End Sub

Function openCofirmMsgBox(msg as String) as Integer
  openCofirmMsgBox = MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2, "確認")
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
    
  ' GetFolder(フォルダ名).SubFoldersでフォルダ配下のフォルダ一覧を取得
  For Each objFolder In fso.GetFolder(path).SubFolders
    Call getFilesRecursive(objFolder.path)
  Next
    
  ' GetFolder(フォルダ名).Filesでフォルダ配下のファイル一覧を取得
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
  ' IgnoreReadOnlyRecommended=True は読み取り推奨のダイアログ非表示
  Set wb = Workbooks.open(Filename:=f.path, IgnoreReadOnlyRecommended:=True)

  call doFormat(wb.Worksheets)

  Call wb.save
  call wb.close
End Sub

Sub doFormat(ws as Sheets)
  For Each Sheet In ws
    'シートを選択
    Sheet.Select
    'アクティブセルをA1へ
    Sheet.Range("A1").Activate
    '倍率を100%
    Sheet.ActiveWindow.Zoom = 100
    'スクロールを最左へ
    ActiveWindow.ScrollColumn = 1
    'スクロールを最上へ
    ActiveWindow.ScrollRow = 1
  Next
  '左端のシートをアクティブにする
  ws(1).Activate

End Sub