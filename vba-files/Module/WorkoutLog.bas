Attribute VB_Name = "WorkoutLog"

Private Const SHEET_NAME As String = "筋トレログ"
Private Const COL_DATE As Integer = 1
Private Const COL_EXERCISE As Integer = 2
Private Const COL_SETS As Integer = 3
Private Const COL_REPS As Integer = 4
Private Const COL_WEIGHT As Integer = 5
Private Const COL_NOTES As Integer = 6

' ==========================================
' 筋トレログシートを初期化（なければ作成）
' ==========================================
Sub InitWorkoutSheet()
  Dim ws As Worksheet
  Set ws = GetOrCreateSheet(SHEET_NAME)

  If ws.Range("A1").Value <> "日付" Then
    ws.Range("A1").Value = "日付"
    ws.Range("B1").Value = "種目"
    ws.Range("C1").Value = "セット数"
    ws.Range("D1").Value = "回数"
    ws.Range("E1").Value = "重量(kg)"
    ws.Range("F1").Value = "メモ"

    With ws.Range("A1:F1")
      .Font.Bold = True
      .Interior.Color = RGB(70, 130, 180)
      .Font.Color = RGB(255, 255, 255)
      .HorizontalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 35

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
  End If

  ws.Activate
End Sub

' ==========================================
' 筋トレを記録する
' ==========================================
Sub AddWorkoutEntry()
  Call InitWorkoutSheet
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

  Dim exerciseName As String
  exerciseName = InputBox("種目名を入力してください" & vbCrLf & _
    "(例: ベンチプレス、スクワット、デッドリフト)", "筋トレ記録 - 種目")
  If exerciseName = "" Then Exit Sub

  Dim setsInput As String
  setsInput = InputBox("セット数を入力してください", "筋トレ記録 - セット数")
  If setsInput = "" Then Exit Sub
  If Not IsNumeric(setsInput) Then
    MsgBox "セット数は数値で入力してください。", vbExclamation, "入力エラー"
    Exit Sub
  End If

  Dim repsInput As String
  repsInput = InputBox("回数を入力してください" & vbCrLf & _
    "(複数セットで異なる場合は例: 10,10,8)", "筋トレ記録 - 回数")
  If repsInput = "" Then Exit Sub

  Dim weightInput As String
  weightInput = InputBox("重量(kg)を入力してください" & vbCrLf & _
    "(自重の場合は 0 を入力)", "筋トレ記録 - 重量")
  If weightInput = "" Then Exit Sub
  If Not IsNumeric(weightInput) Then
    MsgBox "重量は数値で入力してください。", vbExclamation, "入力エラー"
    Exit Sub
  End If

  Dim notesInput As String
  notesInput = InputBox("メモ (省略可能)" & vbCrLf & _
    "(フォーム崩れ注意、前回より増量 など)", "筋トレ記録 - メモ")

  Dim nextRow As Long
  nextRow = ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row + 1
  If nextRow < 2 Then nextRow = 2

  ws.Cells(nextRow, COL_DATE).Value = Now()
  ws.Cells(nextRow, COL_DATE).NumberFormat = "yyyy/mm/dd hh:mm"
  ws.Cells(nextRow, COL_EXERCISE).Value = exerciseName
  ws.Cells(nextRow, COL_SETS).Value = CInt(setsInput)
  ws.Cells(nextRow, COL_REPS).Value = repsInput
  ws.Cells(nextRow, COL_WEIGHT).Value = CDbl(weightInput)
  ws.Cells(nextRow, COL_NOTES).Value = notesInput

  ws.Cells(nextRow, COL_SETS).HorizontalAlignment = xlCenter
  ws.Cells(nextRow, COL_REPS).HorizontalAlignment = xlCenter
  ws.Cells(nextRow, COL_WEIGHT).HorizontalAlignment = xlCenter

  If nextRow Mod 2 = 0 Then
    ws.Range(ws.Cells(nextRow, COL_DATE), ws.Cells(nextRow, COL_NOTES)).Interior.Color = RGB(235, 245, 255)
  End If

  Dim weightLabel As String
  If CDbl(weightInput) = 0 Then
    weightLabel = "自重"
  Else
    weightLabel = weightInput & "kg"
  End If

  MsgBox "記録しました！" & vbCrLf & vbCrLf & _
    "種目: " & exerciseName & vbCrLf & _
    "セット数: " & setsInput & vbCrLf & _
    "回数: " & repsInput & vbCrLf & _
    "重量: " & weightLabel, vbInformation, "記録完了"

  ws.Activate
  ws.Cells(nextRow, COL_DATE).Select
End Sub

' ==========================================
' 今日のトレーニングを一覧表示
' ==========================================
Sub ShowTodaysWorkouts()
  Call InitWorkoutSheet
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

  Dim lastRow As Long
  lastRow = ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row

  If lastRow < 2 Then
    MsgBox "まだ記録がありません。" & vbCrLf & "「筋トレを記録する」から記録を始めましょう！", _
      vbInformation, "今日のトレーニング"
    Exit Sub
  End If

  Dim today As Date
  today = Date

  Dim result As String
  result = "【今日のトレーニング】 " & Format(today, "yyyy/mm/dd") & vbCrLf
  result = result & String(45, "-") & vbCrLf

  Dim count As Integer
  count = 0
  Dim i As Long

  For i = 2 To lastRow
    Dim cellVal As Variant
    cellVal = ws.Cells(i, COL_DATE).Value
    If IsDate(cellVal) Then
      If Int(CDbl(CDate(cellVal))) = Int(CDbl(CDate(today))) Then
        Dim weightVal As Double
        weightVal = CDbl(ws.Cells(i, COL_WEIGHT).Value)

        result = result & "■ " & ws.Cells(i, COL_EXERCISE).Value & vbCrLf
        result = result & "   " & ws.Cells(i, COL_SETS).Value & "セット × " & ws.Cells(i, COL_REPS).Value & "回"
        If weightVal > 0 Then
          result = result & "  @" & weightVal & "kg"
        Else
          result = result & "  (自重)"
        End If
        If ws.Cells(i, COL_NOTES).Value <> "" Then
          result = result & "  ※" & ws.Cells(i, COL_NOTES).Value
        End If
        result = result & vbCrLf
        count = count + 1
      End If
    End If
  Next i

  If count = 0 Then
    result = result & "(今日の記録はまだありません)" & vbCrLf
  End If

  result = result & String(45, "-") & vbCrLf
  result = result & "合計: " & count & " 種目"

  MsgBox result, vbInformation, "今日のトレーニング"
End Sub

' ==========================================
' 種目別 最大重量などの統計を表示
' ==========================================
Sub ShowWorkoutStats()
  Call InitWorkoutSheet
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Worksheets(SHEET_NAME)

  Dim lastRow As Long
  lastRow = ws.Cells(ws.Rows.Count, COL_DATE).End(xlUp).Row

  If lastRow < 2 Then
    MsgBox "まだ記録がありません。", vbInformation, "統計"
    Exit Sub
  End If

  Dim totalEntries As Long
  totalEntries = lastRow - 1

  Dim firstDate As Date
  firstDate = CDate(Int(CDbl(ws.Cells(2, COL_DATE).Value)))

  Dim exMaxWeight As Object
  Dim exCount As Object
  Set exMaxWeight = CreateObject("Scripting.Dictionary")
  Set exCount = CreateObject("Scripting.Dictionary")

  Dim i As Long
  For i = 2 To lastRow
    Dim exName As String
    Dim exWeight As Double
    exName = ws.Cells(i, COL_EXERCISE).Value
    exWeight = CDbl(ws.Cells(i, COL_WEIGHT).Value)

    If exMaxWeight.Exists(exName) Then
      If exMaxWeight(exName) < exWeight Then
        exMaxWeight(exName) = exWeight
      End If
      exCount(exName) = exCount(exName) + 1
    Else
      exMaxWeight.Add exName, exWeight
      exCount.Add exName, 1
    End If
  Next i

  Dim result As String
  result = "【筋トレ統計】" & vbCrLf & vbCrLf
  result = result & "開始日　 : " & Format(firstDate, "yyyy/mm/dd") & vbCrLf
  result = result & "総記録数 : " & totalEntries & " セット" & vbCrLf
  result = result & "種目数　 : " & exMaxWeight.Count & " 種目" & vbCrLf
  result = result & vbCrLf
  result = result & "【種目別 最大重量 / 実施回数】" & vbCrLf
  result = result & String(40, "-") & vbCrLf

  Dim key As Variant
  For Each key In exMaxWeight.Keys
    Dim maxW As Double
    maxW = exMaxWeight(key)
    Dim cnt As Integer
    cnt = exCount(key)

    If maxW > 0 Then
      result = result & key & ": 最大 " & maxW & "kg  (" & cnt & "セット実施)" & vbCrLf
    Else
      result = result & key & ": 自重  (" & cnt & "セット実施)" & vbCrLf
    End If
  Next key

  MsgBox result, vbInformation, "筋トレ統計"
End Sub

' ==========================================
' プライベートヘルパー: シートを取得または作成
' ==========================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
  Dim ws As Worksheet
  For Each ws In ThisWorkbook.Worksheets
    If ws.Name = sheetName Then
      Set GetOrCreateSheet = ws
      Exit Function
    End If
  Next ws
  Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
  ws.Name = sheetName
  Set GetOrCreateSheet = ws
End Function
