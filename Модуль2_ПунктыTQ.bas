Attribute VB_Name = "Модуль2_ПунктыTQ"
' ============================================================================
' МОДУЛЬ: Модуль2_ПунктыTQ
' НАЗНАЧЕНИЕ: Заполнение дат и статусов из файла пунктов TQ.
' ЛОГИКА РАБОТЫ:
'   1. Поиск и открытие файла "*пункты TQ*.xlsb" (лист "Ведомость элементов ТСБ и М (2)").
'   2. Построение словаря по составному ключу A_B_C_D_F_G.
'   3. Обновление приёмника: Столбцы J (формат даты) и L (текстовый формат) заполняются ТОЛЬКО если ячейки пустые.
'   4. Ручные записи в J и L сохраняются и никогда не перезаписываются.
'   5. В лог выводится точное количество обновлённых строк по каждому столбцу.
' ============================================================================

Option Explicit

Public Sub Модуль2_Обновление_ПунктыTQ()
    Const PATH_SRC3 As String = "\\vls.lan\ULVZG-DFS\Велесстрой СМУ\ПТО\ИД\ТСБ и МОТ\!Ведомость элементов\"
    Const SHEET_REC As String = "Ведомость элементов ТСБ и МОТ"
    Const SHEET_SRC As String = "Ведомость элементов ТСБ и М (2)"
    Const REC_START_ROW As Long = 3
    Const SRC_START_ROW As Long = 3
    
    LogMsg "=== ЗАПУСК МОДУЛЯ 2 ==="
    ОбновитьСтатус 5, "Поиск файла пунктов TQ..."
    
    Dim fileSrc As String
    fileSrc = НайтиСамыйСвежийФайл(PATH_SRC3, "*пункты TQ*.xlsb")
    
    If fileSrc = "" Then
        LogMsg "Файл пунктов TQ не найден", "WARN"
        If MsgBox("Файл 'Ведомость элементов МК ТСБ и МОТ - пункты TQ.xlsb' не найден." & vbNewLine & "Продолжить без него?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
        ОбновитьСтатус 100, "[OK] Модуль 2 пропущен."
        Exit Sub
    End If
    LogMsg "Найден файл TQ: " & CreateObject("Scripting.FileSystemObject").GetFileName(fileSrc)
    
    ' Безопасное открытие
    ОбновитьСтатус 10, "Открытие источника TQ..."
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim wbSrc As Workbook, fname As String, isUserOpen As Boolean
    fname = fso.GetFileName(fileSrc)
    Set wbSrc = Nothing: isUserOpen = False
    
    On Error Resume Next
    Set wbSrc = Workbooks(fname)
    If Not wbSrc Is Nothing Then isUserOpen = True
    On Error GoTo 0
    
    If Not isUserOpen Then
        On Error Resume Next
        Set wbSrc = Workbooks.Open(fileSrc, ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo 0
        If wbSrc Is Nothing Then
            LogMsg "ОШИБКА: Не удалось открыть файл TQ", "ERROR"
            MsgBox "Не удалось открыть файл TQ.", vbCritical
            Exit Sub
        End If
    Else
        LogMsg "Файл TQ уже открыт пользователем. Использую текущий экземпляр."
    End If
    
    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets(SHEET_SRC)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        If Not isUserOpen Then wbSrc.Close False
        LogMsg "Лист '" & SHEET_SRC & "' не найден в файле TQ", "WARN"
        MsgBox "Лист '" & SHEET_SRC & "' не найден.", vbCritical
        Exit Sub
    End If
    
    ' Загрузка ключей и данных J/L
    ОбновитьСтатус 20, "Загрузка ключей источника TQ..."
    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    Dim dictSrc As Object: Set dictSrc = CreateObject("Scripting.Dictionary")
    Dim r As Long
    
    If lastRowSrc >= SRC_START_ROW Then
        For r = SRC_START_ROW To lastRowSrc
            Dim kSrc As String
            kSrc = БезопасныйКлюч(wsSrc.Cells(r, "A").Value, wsSrc.Cells(r, "B").Value, wsSrc.Cells(r, "C").Value, _
                                  wsSrc.Cells(r, "D").Value, wsSrc.Cells(r, "F").Value, wsSrc.Cells(r, "G").Value)
            
            If Not dictSrc.Exists(kSrc) Then
                dictSrc(kSrc) = Array(wsSrc.Cells(r, "J").Value, wsSrc.Cells(r, "L").Value)
            End If
            If r Mod 500 = 0 Then ОбновитьСтатус 20 + Int(r / lastRowSrc * 25), "Загрузка ключей источника TQ..."
        Next r
    End If
    LogMsg "Источник TQ загружен: " & dictSrc.Count & " уникальных ключей"
    If Not isUserOpen Then wbSrc.Close False
    
    ' Обновление приёмника (J и L независимо)
    ОбновитьСтатус 50, "Обновление столбцов J и L в приёмнике..."
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets(SHEET_REC)
    Dim lastRowRec As Long: lastRowRec = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row
    If lastRowRec < REC_START_ROW Then Exit Sub
    
    Dim i As Long, updatedJ As Long, updatedL As Long
    Dim stepSize As Long: stepSize = Application.Max(1, lastRowRec \ 20)
    
    For i = REC_START_ROW To lastRowRec
        Dim kRec As String
        kRec = БезопасныйКлюч(wsRec.Cells(i, "A").Value, wsRec.Cells(i, "B").Value, wsRec.Cells(i, "C").Value, _
                              wsRec.Cells(i, "D").Value, wsRec.Cells(i, "F").Value, wsRec.Cells(i, "G").Value)
        
        If dictSrc.Exists(kRec) Then
            Dim srcVals As Variant: srcVals = dictSrc(kRec)
            
            If Trim(CStr(wsRec.Cells(i, "J").Value)) = "" And Not IsNull(srcVals(0)) And srcVals(0) <> "" Then
                wsRec.Cells(i, "J").Value = srcVals(0)
                wsRec.Cells(i, "J").NumberFormat = "dd.mm.yyyy"
                updatedJ = updatedJ + 1
            End If
            
            If Trim(CStr(wsRec.Cells(i, "L").Value)) = "" And Not IsNull(srcVals(1)) And srcVals(1) <> "" Then
                wsRec.Cells(i, "L").Value = srcVals(1)
                wsRec.Cells(i, "L").NumberFormat = "@"
                updatedL = updatedL + 1
            End If
        End If
        
        If i Mod stepSize = 0 Then ОбновитьСтатус 50 + Int(i / lastRowRec * 45), "Обновление столбцов J и L..."
    Next i
    
    LogMsg "Модуль 2 завершён. Обновлено J: " & updatedJ & " строк | Обновлено L: " & updatedL & " строк"
    ОбновитьСтатус 100, "[OK] Модуль 2 завершён. Обновлено J: " & updatedJ & ", L: " & updatedL
End Sub

