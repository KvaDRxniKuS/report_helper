Option Explicit
Public Const TL_FILTER_LIST As String = ""

Private Function GetTLFilterList() As String
    If Trim(TL_FILTER_LIST) <> "" Then
        GetTLFilterList = TL_FILTER_LIST
        Exit Function
    End If
    On Error Resume Next
    Dim ws As Worksheet, wb As Workbook, lastR As Long, r As Long, s As String
    Set ws = ThisWorkbook.Worksheets("Фильтр")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("TL")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("TL_FILTER")
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        s = ""
        For r = 1 To lastR
            If Trim(ws.Cells(r, 1).Value) <> "" Then
                If s = "" Then s = Trim(ws.Cells(r, 1).Value) Else s = s & vbCrLf & Trim(ws.Cells(r, 1).Value)
            End If
        Next r
        If s <> "" Then GetTLFilterList = s: Exit Function
    End If
    Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        Set ws = Nothing
        Set ws = wb.Worksheets("Фильтр")
        If ws Is Nothing Then Set ws = wb.Worksheets("TL")
        If ws Is Nothing Then Set ws = wb.Worksheets("TL_FILTER")
        If Not ws Is Nothing Then
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            s = ""
            For r = 1 To lastR
                If Trim(ws.Cells(r, 1).Value) <> "" Then
                    If s = "" Then s = Trim(ws.Cells(r, 1).Value) Else s = s & vbCrLf & Trim(ws.Cells(r, 1).Value)
                End If
            Next r
            If s <> "" Then GetTLFilterList = s: Exit Function
        End If
    End If
    On Error GoTo 0
    GetTLFilterList = ""
End Function

Sub UniversalReport()
    Dim ws As Worksheet, header As String, c As Long, lastCol As Long
    Set ws = ActiveSheet
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol = 1 And Trim(ws.Cells(1, 1).Value) = "" Then lastCol = ws.UsedRange.Columns.Count
    header = ""
    For c = 1 To lastCol
        header = header & " " & LCase(Trim(CStr(ws.Cells(1, c).Value)))
    Next c
    ' Проверяем первые 3 строки на всякий случай если шапка не в первой
    Dim rr As Long
    For rr = 2 To 3
        For c = 1 To lastCol
            header = header & " " & LCase(Trim(CStr(ws.Cells(rr, c).Value)))
        Next c
    Next rr
    
    If InStr(header, "результат вызова") > 0 Then
        Private_CreateCallReport
        Exit Sub
    End If
    If InStr(header, "суммарное время работы в системе") > 0 Or InStr(header, "перерыв") > 0 Or InStr(header, "суммарное") > 0 Then
        Private_ObrabotkaTablitsyTL_Multi
        Exit Sub
    End If
    If InStr(header, "уникальные звонки") > 0 Or InStr(header, "уникальн") > 0 Or InStr(header, "уник") > 0 Then
        Private_ObrabotkaUnikalnyh
        Exit Sub
    End If
    ' УК отчет часто не содержит слово "уникальные" в исходнике, но содержит логин+фио+тл
    If InStr(header, "тл") > 0 And (InStr(header, "логин") > 0 Or InStr(header, "фио") > 0) Then
        ' Если есть результат вызова - это АМД, уже проверили выше, значит это УК или Часы
        ' Часы уже проверили по суммарное/перерыв, значит осталось УК
        Private_ObrabotkaUnikalnyh
        Exit Sub
    End If
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If Not wb Is Nothing Then
        Dim wsOtchet As Worksheet
        Set wsOtchet = wb.Worksheets("отчет")
        If Not wsOtchet Is Nothing Then
            wsOtchet.Activate
            Private_CreateCallReport
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Dim ans As Long
    ans = MsgBox("Не удалось определить тип отчета." & vbCrLf & "1=АМД (Результат вызова)" & vbCrLf & "2=Часы" & vbCrLf & "3=УК (уникальные)" & vbCrLf & vbCrLf & "Нажми Да=АМД, Нет=Часы, Отмена=УК", vbYesNoCancel + vbQuestion, "Выбор отчета")
    If ans = vbYes Then Private_CreateCallReport
    If ans = vbNo Then Private_ObrabotkaTablitsyTL_Multi
    If ans = vbCancel Then Private_ObrabotkaUnikalnyh
End Sub

' Отдельные точки входа чтобы не зависеть от автоопределения (скрытые, вызываются из UniversalReport)
Private Sub CreateAMDReport()
    Private_CreateCallReport
End Sub

Private Sub CreateHoursReport()
    Private_ObrabotkaTablitsyTL_Multi
End Sub

Private Sub CreateUKReport()
    Private_ObrabotkaUnikalnyh
End Sub

Public Function CopySheetImage(Optional ByVal unused As Variant)
    Dim ws As Worksheet, wb As Workbook, btnName As String
    Set wb = ActiveWorkbook
    On Error Resume Next
    btnName = Application.Caller
    Set ws = wb.ActiveSheet.Buttons(btnName).Parent
    If ws Is Nothing Then Set ws = wb.ActiveSheet.Shapes(btnName).TopLeftCell.Worksheet
    If ws Is Nothing Then Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    
    Dim lastRow As Long, lastCol As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Function
    lastRow = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lastCol = lastCell.Column
    
    Dim isAMD As Boolean
    isAMD = False
    If LCase(Trim(CStr(ws.Cells(3, 1).Value))) Like "*результат*" Then isAMD = True
    If InStr(LCase(ws.Name), "сводная") > 0 Then isAMD = True
    
    If isAMD Then
        On Error Resume Next
        ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol)).CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        If Err.Number <> 0 Then
            Err.Clear
            ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, lastCol)).CopyPicture Appearance:=xlScreen, Format:=xlPicture
        End If
        On Error GoTo 0
    Else
        Call Private_SafeCopyPicture(ws, lastRow, lastCol, 1)
    End If
End Function

Private Sub Private_CreateCallReport()
    Dim wb As Workbook
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim lastRow As Long, lastColumn As Long
    Dim cache As PivotCache, pivot As PivotTable
    Dim resultField As PivotField
    Dim baseName As String, sheetName As String, suffix As Long
    Dim oldScreenUpdating As Boolean
    Dim grandRow As Long, grandCol As Long, aoRow As Long, rr As Long, cc As Long
    Dim pvtRange As Range, headerRow As Long
    Dim lastR As Long, lastC As Long
    Dim v As String, aoName As String
    Dim rBot As Long, cPct As Long
    Dim totVal As Double, aoVal As Double, rowTot As Double, grandTotalVal As Double
    Dim pf As PivotField, pfResult As PivotField, pfTM As PivotField
    
    oldScreenUpdating = Application.ScreenUpdating
    On Error GoTo Fin
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo Fin
    If UCase$(wb.Name) = "PERSONAL.XLSB" Then GoTo Fin
    On Error Resume Next
    Set wsSource = wb.Worksheets("отчет")
    If wsSource Is Nothing Then Set wsSource = ActiveSheet
    On Error GoTo Fin
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastColumn = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then GoTo Fin
    Dim srcRange As Range
    Set srcRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastColumn))
    baseName = "Сводная_" & Format(Now, "yyyymmdd_hhnnss")
    sheetName = baseName
    Do While Private_SheetExists(wb, sheetName)
        suffix = suffix + 1
        sheetName = baseName & "_" & CStr(suffix)
    Loop
    Set wsResult = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsResult.Name = sheetName
    Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRange.Address(ReferenceStyle:=xlR1C1, External:=True))
    Set pivot = cache.CreatePivotTable(TableDestination:=wsResult.Range("A3"))
    
    For Each pf In pivot.PivotFields
        If LCase(Trim(pf.Name)) Like "*результат*вызова*" Then Set pfResult = pf
        If LCase(Trim(pf.Name)) Like "*проект*тм*" Then Set pfTM = pf
    Next pf
    If pfResult Is Nothing Then
        On Error Resume Next
        Set pfResult = pivot.PivotFields("Результат вызова")
        On Error GoTo Fin
        If pfResult Is Nothing Then
            MsgBox "В листе 'отчет' не найдено поле 'Результат вызова'. Проверь что это АМД отчет.", vbCritical
            GoTo Fin
        End If
    End If
    
    With pivot
        .AddDataField pfResult, "Количество", xlCount
        Set resultField = pfResult
        resultField.Orientation = xlRowField
        resultField.Position = 1
        resultField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        If Not pfTM Is Nothing Then
            pfTM.Orientation = xlColumnField
            pfTM.Position = 1
        Else
            On Error Resume Next
            .PivotFields("ПРОЕКТ ТМ").Orientation = xlColumnField
            .PivotFields("ПРОЕКТ ТМ").Position = 1
            On Error GoTo Fin
        End If
        .RowAxisLayout xlTabularRow
        .RowGrand = True
        .ColumnGrand = True
        .RefreshTable
    End With
    
    Set pvtRange = pivot.TableRange2
    lastR = pvtRange.Row + pvtRange.Rows.Count - 1
    lastC = pvtRange.Column + pvtRange.Columns.Count - 1
    grandRow = lastR
    grandCol = lastC
    headerRow = pvtRange.Row
    
    For rr = lastR To pvtRange.Row Step -1
        v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
        If InStr(v, "общий итог") > 0 Then
            grandRow = rr
            Exit For
        End If
    Next rr
    
    For rr = pvtRange.Row To pvtRange.Row + 3
        For cc = lastC To pvtRange.Column Step -1
            v = LCase(Trim(CStr(wsResult.Cells(rr, cc).Value)))
            If InStr(v, "общий итог") > 0 Then
                grandCol = cc
                headerRow = rr
                Exit For
            End If
        Next cc
        If grandCol <> 0 And grandCol <> lastC Then Exit For
        If grandCol = lastC Then
            If InStr(LCase(Trim(CStr(wsResult.Cells(rr, lastC).Value))), "общий итог") > 0 Then Exit For
        End If
    Next rr
    If grandCol = 0 Then grandCol = lastC
    
    ' Поиск ТОЛЬКО "Вызов завершен при обнаружении автоответчика"
    For rr = pvtRange.Row + 1 To grandRow - 1
        v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
        If v = "результат вызова" Then GoTo N1
        If InStr(v, "общий итог") > 0 Then GoTo N1
        If v = "вызов завершен при обнаружении автоответчика" Then aoRow = rr: Exit For
N1:
    Next rr
    If aoRow = 0 Then
        For rr = pvtRange.Row + 1 To grandRow - 1
            v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
            If InStr(v, "вызов завершен при обнаружении автоответчика") > 0 Then aoRow = rr: Exit For
        Next rr
    End If
    If aoRow = 0 Then
        For rr = pvtRange.Row + 1 To grandRow - 1
            v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
            If InStr(v, "завершен при обнаружении автоответчика") > 0 Then aoRow = rr: Exit For
        Next rr
    End If
    If aoRow = 0 Then
        For rr = pvtRange.Row + 1 To grandRow - 1
            v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
            If InStr(v, "завершен") > 0 And InStr(v, "обнаружении") > 0 And InStr(v, "автоответ") > 0 Then aoRow = rr: Exit For
        Next rr
    End If
    If aoRow = 0 Then
        For rr = pvtRange.Row + 1 To grandRow - 1
            v = LCase(Trim(CStr(wsResult.Cells(rr, pvtRange.Column).Value)))
            If InStr(v, "завершен при обнаружении") > 0 Then aoRow = rr: Exit For
        Next rr
    End If
    If aoRow = 0 Then
        MsgBox "Не найдена строка 'Вызов завершен при обнаружении автоответчика'.", vbExclamation
        aoRow = pvtRange.Row + 1
    End If
    
    aoName = Trim(CStr(wsResult.Cells(aoRow, pvtRange.Column).Value))
    If aoName = "" Or LCase(aoName) = "результат вызова" Then aoName = "Вызов завершен при обнаружении автоответчика"
    
    pivot.ManualUpdate = True
    
    rBot = grandRow + 1
    cPct = grandCol + 1
    If rBot <= pvtRange.Row + pvtRange.Rows.Count - 1 Then rBot = pvtRange.Row + pvtRange.Rows.Count
    If cPct <= pvtRange.Column + pvtRange.Columns.Count - 1 Then cPct = pvtRange.Column + pvtRange.Columns.Count
    
    With wsResult
        .Columns("A:Z").AutoFit
        
        On Error Resume Next
        .Cells(rBot, pvtRange.Column).Value = aoName
        .Cells(rBot, pvtRange.Column).Font.Bold = True
        On Error GoTo 0
        
        For cc = pvtRange.Column + 1 To grandCol
            On Error Resume Next
            totVal = 0: aoVal = 0
            If IsNumeric(wsResult.Cells(grandRow, cc).Value) Then totVal = CDbl(wsResult.Cells(grandRow, cc).Value)
            If IsNumeric(wsResult.Cells(aoRow, cc).Value) Then aoVal = CDbl(wsResult.Cells(aoRow, cc).Value)
            If totVal <> 0 Then
                wsResult.Cells(rBot, cc).FormulaR1C1 = "=IF(R" & grandRow & "C=0,"""",R" & aoRow & "C/R" & grandRow & "C)"
                If Err.Number <> 0 Then
                    Err.Clear
                    wsResult.Cells(rBot, cc).Value = aoVal / totVal
                End If
                wsResult.Cells(rBot, cc).NumberFormat = "0.00%"
            End If
            On Error GoTo 0
        Next cc
        On Error Resume Next
        .Cells(rBot, grandCol).Font.Bold = True
        .Cells(headerRow, cPct).Value = "% от общего"
        .Cells(headerRow, cPct).Font.Bold = True
        On Error GoTo 0
        
        grandTotalVal = 0
        If IsNumeric(wsResult.Cells(grandRow, grandCol).Value) Then grandTotalVal = CDbl(wsResult.Cells(grandRow, grandCol).Value)
        
        For rr = pvtRange.Row + 1 To grandRow
            On Error Resume Next
            rowTot = 0
            If IsNumeric(wsResult.Cells(rr, grandCol).Value) Then rowTot = CDbl(wsResult.Cells(rr, grandCol).Value)
            If grandTotalVal <> 0 Then
                wsResult.Cells(rr, cPct).FormulaR1C1 = "=IF(R" & grandRow & "C" & grandCol & "=0,"""",RC" & grandCol & "/R" & grandRow & "C" & grandCol & ")"
                If Err.Number <> 0 Then
                    Err.Clear
                    wsResult.Cells(rr, cPct).Value = rowTot / grandTotalVal
                End If
                wsResult.Cells(rr, cPct).NumberFormat = "0.00%"
            End If
            On Error GoTo 0
        Next rr
        On Error Resume Next
        .Cells(grandRow, cPct).Font.Bold = True
        .Columns(cPct).AutoFit
        On Error GoTo 0
    End With
    
    pivot.ManualUpdate = False
    
    Dim btn As Button
    Dim btnLeft As Double
    On Error Resume Next
    wsResult.Buttons.Delete
    On Error GoTo 0
    btnLeft = wsResult.Cells(1, cPct + 2).Left + 15
    If btnLeft < 10 Then btnLeft = 10
    Set btn = wsResult.Buttons.Add(Left:=btnLeft, Top:=5, Width:=180, Height:=30)
    btn.Caption = "скопировать изображение"
    btn.OnAction = "CopySheetImage"
    btn.Font.Bold = True
    btn.Visible = True
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    wb.Activate
    wsResult.Activate
    wsResult.Range("A3").Select
    Exit Sub
Fin:
    Application.ScreenUpdating = oldScreenUpdating
    Application.Calculation = xlCalculationAutomatic
    If Err.Number <> 0 Then MsgBox "Ошибка АМД: " & Err.Description, vbCritical
End Sub

Private Sub Private_ObrabotkaTablitsyTL_Multi()
    Dim wb As Workbook, wsSource As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim colFIO As Long, colBreak As Long, colTotal As Long, colTL As Long
    Dim c As Long, r As Long, h As String
    Dim dict As Object, tl As Variant, ordered As Collection, k As Variant
    Dim oldScreenUpdating As Boolean, firstSheet As Boolean, wsNew As Worksheet, safeName As String, destRow As Long, dataLastRow As Long
    Dim btn As Button, btnLeft As Double
    Dim tlName As String, filterCol As Collection, userInput As String
    Set wb = ActiveWorkbook
    Set wsSource = ActiveSheet
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    wsSource.Cells.EntireColumn.Hidden = False
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    colFIO = 0: colBreak = 0: colTotal = 0: colTL = 0
    For c = 1 To lastCol
        h = LCase(Trim(CStr(wsSource.Cells(1, c).Value)))
        If h = "фио" Then colFIO = c
        If InStr(h, "перерыв") > 0 Then colBreak = c
        If InStr(h, "суммарное") > 0 And InStr(h, "системе") > 0 Then colTotal = c
        If h = "тл" Then colTL = c
    Next c
    If colFIO = 0 Or colBreak = 0 Or colTotal = 0 Or colTL = 0 Then GoTo Fin2
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    For r = 2 To lastRow
        tlName = Trim(CStr(wsSource.Cells(r, colTL).Value))
        If tlName <> "" Then If Not dict.Exists(tlName) Then dict.Add tlName, True
    Next r
    Set filterCol = Private_ParseTLList(GetTLFilterList())
    If filterCol.Count = 0 Then
        userInput = InputBox("Вставь список ТЛ для часов через , ; | или с новой строки." & vbCrLf & "Пусто = все ТЛ", "Фильтр по ТЛ")
        If Trim(userInput) <> "" Then Set filterCol = Private_ParseTLList(userInput)
    End If
    Set ordered = New Collection
    If filterCol.Count > 0 Then
        For Each k In dict.Keys
            If Private_TLMatches(k, filterCol) Then ordered.Add k
        Next k
    Else
        For Each k In dict.Keys
            ordered.Add k
        Next k
    End If
    If ordered.Count = 0 Then
        For Each k In dict.Keys
            ordered.Add k
        Next k
    End If
    firstSheet = True
    For Each tl In ordered
        safeName = Private_SafeSheetName(CStr(tl))
        On Error Resume Next
        Application.DisplayAlerts = False
        wb.Worksheets(safeName).Delete
        Application.DisplayAlerts = True
        On Error GoTo Fin2
        Set wsNew = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsNew.Name = safeName
        wsNew.Cells(1, 1).Value = wsSource.Cells(1, colFIO).Value
        wsNew.Cells(1, 2).Value = wsSource.Cells(1, colBreak).Value
        wsNew.Cells(1, 3).Value = wsSource.Cells(1, colTotal).Value
        wsNew.Cells(1, 4).Value = wsSource.Cells(1, colTL).Value
        wsNew.Cells(1, 5).Value = "Чистое время работы"
        wsNew.Range("A1:E1").Font.Bold = True
        destRow = 2
        For r = 2 To lastRow
            If LCase(Trim(CStr(wsSource.Cells(r, colTL).Value))) = LCase(Trim(CStr(tl))) Then
                wsNew.Cells(destRow, 1).Value = wsSource.Cells(r, colFIO).Value
                wsNew.Cells(destRow, 2).Value = wsSource.Cells(r, colBreak).Value
                wsNew.Cells(destRow, 3).Value = wsSource.Cells(r, colTotal).Value
                wsNew.Cells(destRow, 4).Value = wsSource.Cells(r, colTL).Value
                wsNew.Cells(destRow, 5).FormulaR1C1 = "=RC3-RC2"
                wsNew.Cells(destRow, 5).NumberFormat = "[h]:mm:ss"
                destRow = destRow + 1
            End If
        Next r
        dataLastRow = destRow - 1
        If dataLastRow >= 2 Then
            wsNew.Cells(dataLastRow + 2, 1).Value = "Итого часов:"
            wsNew.Cells(dataLastRow + 2, 1).Font.Bold = True
            wsNew.Cells(dataLastRow + 2, 5).Formula = "=SUM(E2:E" & dataLastRow & ")"
            wsNew.Cells(dataLastRow + 2, 5).NumberFormat = "[h]:mm:ss"
            wsNew.Cells(dataLastRow + 2, 5).Font.Bold = True
        End If
        wsNew.Columns("A:E").AutoFit
        wsNew.Columns(2).Hidden = True
        wsNew.Columns(3).Hidden = True
        wsNew.Columns(4).Hidden = True
        On Error Resume Next
        wsNew.Buttons.Delete
        On Error GoTo 0
        btnLeft = wsNew.Cells(1, 7).Left + 15
        Set btn = wsNew.Buttons.Add(Left:=btnLeft, Top:=5, Width:=180, Height:=30)
        btn.Caption = "скопировать изображение"
        btn.OnAction = "CopySheetImage"
        btn.Font.Bold = True
        btn.Visible = True
        If firstSheet Then
            wsNew.Move Before:=wb.Sheets(1)
            firstSheet = False
        End If
    Next tl
    Application.ScreenUpdating = True
    wb.Sheets(1).Activate
    wb.Sheets(1).Range("A1").Select
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
Fin2:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = True
End Sub

Private Sub Private_ObrabotkaUnikalnyh()
    Dim wb As Workbook, wsSource As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim colTL As Long, colFIO As Long, colLogin As Long
    Dim c As Long, r As Long, h As String
    Dim dict As Object, tl As Variant, ordered As Collection, k As Variant
    Dim firstSheet As Boolean, wsNew As Worksheet, safeName As String, destRow As Long, dataLastRow As Long
    Dim oldScreenUpdating As Boolean
    Dim btn As Button, btnLeft As Double
    Dim numericColsSource As Collection
    Dim tlName2 As String
    Dim filterCol As Collection, userInput As String
    Dim colMap As Object, destC As Long, srcCol As Variant, numericDestCols As Collection, newCol As Long
    Dim d As Long, formula As String, nd As Variant
    
    Set wb = ActiveWorkbook
    Set wsSource = ActiveSheet
    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    wsSource.Cells.EntireColumn.Hidden = False
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    For c = lastCol To 1 Step -1
        h = LCase(Trim(CStr(wsSource.Cells(1, c).Value)))
        If InStr(h, "ао") > 0 Then wsSource.Columns(c).Delete
    Next c
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    colTL = 0: colFIO = 0: colLogin = 0
    Set numericColsSource = New Collection
    For c = 1 To lastCol
        h = LCase(Trim(CStr(wsSource.Cells(1, c).Value)))
        If h = "тл" Then colTL = c
        If InStr(h, "фио") > 0 And InStr(h, "оператора") > 0 Then colFIO = c
        If InStr(h, "логин") > 0 And InStr(h, "оператора") > 0 Then colLogin = c
    Next c
    For c = 1 To lastCol
        If c <> colTL And c <> colFIO And c <> colLogin Then
            If Trim(wsSource.Cells(1, c).Value) <> "" Then numericColsSource.Add c
        End If
    Next c
    If colTL = 0 Then GoTo Fin3
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    For r = 2 To lastRow
        tlName2 = Trim(CStr(wsSource.Cells(r, colTL).Value))
        If tlName2 <> "" Then If Not dict.Exists(tlName2) Then dict.Add tlName2, True
    Next r
    Set filterCol = Private_ParseTLList(GetTLFilterList())
    If filterCol.Count = 0 Then
        userInput = InputBox("Вставь список ТЛ для УК через , ; | или с новой строки." & vbCrLf & "Пусто = все ТЛ", "Фильтр по ТЛ для УК", "")
        If Trim(userInput) <> "" Then Set filterCol = Private_ParseTLList(userInput)
    End If
    Set ordered = New Collection
    If filterCol.Count > 0 Then
        For Each k In dict.Keys
            If Private_TLMatches(k, filterCol) Then ordered.Add k
        Next k
    Else
        For Each k In dict.Keys
            ordered.Add k
        Next k
    End If
    If ordered.Count = 0 Then
        For Each k In dict.Keys
            ordered.Add k
        Next k
    End If
    firstSheet = True
    For Each tl In ordered
        safeName = Private_SafeSheetName(CStr(tl) & "_УК")
        On Error Resume Next
        Application.DisplayAlerts = False
        wb.Worksheets(safeName).Delete
        Application.DisplayAlerts = True
        On Error GoTo Fin3
        Set wsNew = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsNew.Name = safeName
        Set colMap = CreateObject("Scripting.Dictionary")
        destC = 1
        If colLogin > 0 Then
            wsNew.Cells(1, destC).Value = wsSource.Cells(1, colLogin).Value
            colMap.Add "LOGIN", destC
            destC = destC + 1
        End If
        If colFIO > 0 Then
            wsNew.Cells(1, destC).Value = wsSource.Cells(1, colFIO).Value
            colMap.Add "FIO", destC
            destC = destC + 1
        End If
        Set numericDestCols = New Collection
        For Each srcCol In numericColsSource
            wsNew.Cells(1, destC).Value = wsSource.Cells(1, srcCol).Value
            numericDestCols.Add destC
            destC = destC + 1
        Next srcCol
        newCol = destC
        wsNew.Cells(1, newCol).Value = "Итого уникальных"
        wsNew.Range(wsNew.Cells(1, 1), wsNew.Cells(1, newCol)).Font.Bold = True
        destRow = 2
        For r = 2 To lastRow
            If LCase(Trim(CStr(wsSource.Cells(r, colTL).Value))) = LCase(Trim(CStr(tl))) Then
                d = 1
                If colLogin > 0 Then
                    wsNew.Cells(destRow, d).Value = wsSource.Cells(r, colLogin).Value
                    d = d + 1
                End If
                If colFIO > 0 Then
                    wsNew.Cells(destRow, d).Value = wsSource.Cells(r, colFIO).Value
                    d = d + 1
                End If
                For Each srcCol In numericColsSource
                    wsNew.Cells(destRow, d).Value = wsSource.Cells(r, srcCol).Value
                    d = d + 1
                Next srcCol
                formula = ""
                For Each nd In numericDestCols
                    If formula = "" Then formula = "=RC" & nd Else formula = formula & "+RC" & nd
                Next nd
                If formula <> "" Then wsNew.Cells(destRow, newCol).FormulaR1C1 = formula
                destRow = destRow + 1
            End If
        Next r
        dataLastRow = destRow - 1
        If dataLastRow >= 2 Then
            wsNew.Cells(dataLastRow + 2, 1).Value = "Итого:"
            wsNew.Cells(dataLastRow + 2, 1).Font.Bold = True
            wsNew.Cells(dataLastRow + 2, newCol).Formula = "=SUM(" & wsNew.Cells(2, newCol).Address(False, False) & ":" & wsNew.Cells(dataLastRow, newCol).Address(False, False) & ")"
            wsNew.Cells(dataLastRow + 2, newCol).Font.Bold = True
        End If
        wsNew.Columns.AutoFit
        ' Скрываем все кроме ФИО и Итого уникальных
        Dim hideC As Long, fioCol As Long
        fioCol = 0
        If colMap.Exists("FIO") Then fioCol = colMap("FIO")
        For hideC = 1 To newCol - 1
            If hideC <> fioCol Then
                wsNew.Columns(hideC).Hidden = True
            Else
                wsNew.Columns(hideC).Hidden = False
            End If
        Next hideC
        wsNew.Columns(newCol).Hidden = False
        On Error Resume Next
        wsNew.Buttons.Delete
        On Error GoTo 0
        btnLeft = wsNew.Cells(1, newCol + 2).Left + 15
        Set btn = wsNew.Buttons.Add(Left:=btnLeft, Top:=5, Width:=180, Height:=30)
        btn.Caption = "скопировать изображение"
        btn.OnAction = "CopySheetImage"
        btn.Font.Bold = True
        btn.Visible = True
        If firstSheet Then
            wsNew.Move Before:=wb.Sheets(1)
            firstSheet = False
        End If
    Next tl
    Application.ScreenUpdating = True
    wb.Sheets(1).Activate
    wb.Sheets(1).Range("A1").Select
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub
Fin3:
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = True
End Sub

Private Sub Private_SafeCopyPicture(ByVal wsSource As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, Optional ByVal startRow As Long = 1)
    Dim wb As Workbook, wsTmp As Worksheet, c As Long, destC As Long
    Set wb = wsSource.Parent
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("TMP_COPY").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsTmp = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsTmp.Name = "TMP_COPY"
    destC = 1
    Application.ScreenUpdating = False
    For c = 1 To lastCol
        If Not wsSource.Columns(c).Hidden Then
            wsTmp.Columns(destC).ColumnWidth = wsSource.Columns(c).ColumnWidth
            wsTmp.Range(wsTmp.Cells(1, destC), wsTmp.Cells(lastRow - startRow + 1, destC)).Value = wsSource.Range(wsSource.Cells(startRow, c), wsSource.Cells(lastRow, c)).Value
            On Error Resume Next
            wsSource.Range(wsSource.Cells(startRow, c), wsSource.Cells(lastRow, c)).Copy
            wsTmp.Cells(1, destC).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
            On Error GoTo 0
            destC = destC + 1
        End If
    Next c
    If destC > 1 Then
        Dim oldScreenUpdating As Boolean
        oldScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = True
        wb.Activate
        wsTmp.Activate
        wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(lastRow - startRow + 1, destC - 1)).Select
        Application.CutCopyMode = False
        DoEvents
        On Error Resume Next
        wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(lastRow - startRow + 1, destC - 1)).CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        If Err.Number <> 0 Then
            Err.Clear
            wsTmp.Range(wsTmp.Cells(1, 1), wsTmp.Cells(lastRow - startRow + 1, destC - 1)).CopyPicture Appearance:=xlScreen, Format:=xlPicture
        End If
        On Error GoTo 0
        DoEvents
        Application.ScreenUpdating = oldScreenUpdating
    End If
    Application.DisplayAlerts = False
    wsTmp.Delete
    Application.DisplayAlerts = True
    On Error Resume Next
    wb.Activate
    wsSource.Activate
    wsSource.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function Private_SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(sheetName)
    On Error GoTo 0
    Private_SheetExists = Not sh Is Nothing
End Function

Private Function Private_SafeSheetName(ByVal s As String) As String
    Dim bad As Variant
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    Dim b As Variant
    For Each b In bad
        s = Replace(s, b, "_")
    Next b
    s = Trim(s)
    If Len(s) = 0 Then s = "Лист"
    If Len(s) > 31 Then s = Left(s, 31)
    Private_SafeSheetName = s
End Function

Private Function Private_ParseTLList(ByVal s As String) As Collection
    Dim col As New Collection
    Dim tmp As String
    tmp = s
    tmp = Replace(tmp, ";", ",")
    tmp = Replace(tmp, "|", ",")
    tmp = Replace(tmp, vbCrLf, ",")
    tmp = Replace(tmp, vbCr, ",")
    tmp = Replace(tmp, vbLf, ",")
    tmp = Replace(tmp, vbTab, ",")
    Dim parts() As String
    parts = Split(tmp, ",")
    Dim p As Variant
    For Each p In parts
        Dim t As String
        t = Trim(CStr(p))
        If t <> "" Then col.Add t
    Next p
    Set Private_ParseTLList = col
End Function

Private Function Private_TLMatches(ByVal tlName As String, ByVal filterCol As Collection) As Boolean
    Dim f As Variant
    For Each f In filterCol
        If InStr(LCase(tlName), LCase(f)) > 0 Or InStr(LCase(f), LCase(tlName)) > 0 Then
            Private_TLMatches = True
            Exit Function
        End If
    Next f
End Function
