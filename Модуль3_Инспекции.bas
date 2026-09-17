Attribute VB_Name = "Модуль3_Инспекции"
' ============================================================================
' МОДУЛЬ: Модуль3_Инспекции
' НАЗНАЧЕНИЕ: Синхронизация инспекций, финальное форматирование и сортировка.
' ЛОГИКА РАБОТЫ:
'   1. Поиск и открытие файла "*Инспекции на*.xlsx" (лист "Инспекции").
'   2. Словарь строится по ключу столбца D источника.
'   3. Очистка столбцов N:P в приёмнике для гарантированной загрузки актуальных статусов.
'   4. Обновление M:S при точном совпадении: Приёмник.L = Источник.D.
'   5. Применение тонких границ и выравнивания по центру для рабочего диапазона A:X.
'   6. Финальная сортировка таблицы по возрастанию: A > B > C > D > I.
' ============================================================================

Option Explicit

Public Sub Модуль3_Обновление_Инспекции()
    Const PATH_SRC4 As String = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\Выгрузки\RFI\"
    Const SHEET_REC As String = "Ведомость элементов ТСБ и МОТ"
    Const SHEET_SRC As String = "Инспекции"
    Const REC_START_ROW As Long = 3
    Const SRC_START_ROW As Long = 2
    
    LogMsg "=== ЗАПУСК МОДУЛЯ 3 ==="
    ОбновитьСтатус 5, "Поиск файла инспекций..."
    
    ' ?? Поиск самого свежего файла
    Dim fileSrc As String
    fileSrc = НайтиСамыйСвежийФайл(PATH_SRC4, "*Инспекции на*.xlsx")
    
    If fileSrc = "" Then
        LogMsg "Файл инспекций не найден", "WARN"
        If MsgBox("?? Файл 'Инспекции на *.xlsx' не найден." & vbNewLine & "Продолжить без него?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
        ОбновитьСтатус 100, "[OK] Модуль 3 пропущен."
        Exit Sub
    End If
    LogMsg "Найден файл инспекций: " & CreateObject("Scripting.FileSystemObject").GetFileName(fileSrc)
    
    ' ?? Безопасное открытие источника
    ОбновитьСтатус 10, "Открытие файла инспекций..."
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
            LogMsg "ОШИБКА: Не удалось открыть файл инспекций", "ERROR"
            MsgBox "Не удалось открыть файл инспекций.", vbCritical
            Exit Sub
        End If
    Else
        LogMsg "Файл инспекций уже открыт пользователем. Использую текущий экземпляр."
    End If
    
    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = wbSrc.Sheets(SHEET_SRC)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        If Not isUserOpen Then wbSrc.Close False
        LogMsg "Лист '" & SHEET_SRC & "' не найден", "WARN"
        MsgBox "Лист '" & SHEET_SRC & "' не найден.", vbCritical
        Exit Sub
    End If
    
    ' ?? Загрузка ключей (D) и данных (C, AF, AG, AH, AI, AK, AL) в словарь
    ОбновитьСтатус 15, "Загрузка ключей инспекций..."
    Dim lastRowSrc As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "D").End(xlUp).Row
    Dim dictSrc As Object: Set dictSrc = CreateObject("Scripting.Dictionary")
    Dim r As Long
    
    If lastRowSrc >= SRC_START_ROW Then
        For r = SRC_START_ROW To lastRowSrc
            Dim kSrc As String
            kSrc = Trim(CStr(wsSrc.Cells(r, "D").Value))
            If kSrc <> "" Then
                If Not dictSrc.Exists(kSrc) Then
                    dictSrc(kSrc) = Array(wsSrc.Cells(r, "C").Value, _
                                          wsSrc.Cells(r, "AF").Value, _
                                          wsSrc.Cells(r, "AG").Value, _
                                          wsSrc.Cells(r, "AH").Value, _
                                          wsSrc.Cells(r, "AI").Value, _
                                          wsSrc.Cells(r, "AK").Value, _
                                          wsSrc.Cells(r, "AL").Value)
                End If
            End If
            If r Mod 500 = 0 Then ОбновитьСтатус 15 + Int(r / lastRowSrc * 15), "Загрузка инспекций..."
        Next r
    End If
    LogMsg "Инспекции загружены: " & dictSrc.Count & " уникальных ключей (D)"
    If Not isUserOpen Then wbSrc.Close False
    
    ' ?? Очистка N, O, P перед обновлением
    ОбновитьСтатус 30, "Очистка столбцов N, O, P..."
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets(SHEET_REC)
    Dim lastRowRec As Long: lastRowRec = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row
    If lastRowRec >= REC_START_ROW Then
        wsRec.Range(wsRec.Cells(REC_START_ROW, "N"), wsRec.Cells(lastRowRec, "P")).ClearContents
    End If
    
    ' ?? Синхронизация данных приёмника (L=D -> M:S)
    ОбновитьСтатус 45, "Сопоставление L=D и обновление M:S..."
    If lastRowRec < REC_START_ROW Then lastRowRec = REC_START_ROW
    
    Dim i As Long, updatedCount As Long: updatedCount = 0
    Dim stepSize As Long: stepSize = Application.Max(1, lastRowRec \ 20)
    
    For i = REC_START_ROW To lastRowRec
        Dim kRec As String
        kRec = Trim(CStr(wsRec.Cells(i, "L").Value))
        
        If kRec <> "" And dictSrc.Exists(kRec) Then
            Dim srcVals As Variant: srcVals = dictSrc(kRec)
            With wsRec
                .Cells(i, "M").Value = srcVals(0) ' C -> M
                .Cells(i, "N").Value = srcVals(1) ' AF -> N
                .Cells(i, "O").Value = srcVals(2) ' AG -> O
                .Cells(i, "P").Value = srcVals(3) ' AH -> P
                .Cells(i, "Q").Value = srcVals(4) ' AI -> Q
                .Cells(i, "R").Value = srcVals(5) ' AK -> R
                .Cells(i, "S").Value = srcVals(6) ' AL -> S
            End With
            updatedCount = updatedCount + 1
        End If
        
        If i Mod stepSize = 0 Then ОбновитьСтатус 45 + Int(i / lastRowRec * 45), "Синхронизация инспекций..."
    Next i
    LogMsg "Обновлено строк инспекций: " & updatedCount
    
    ' ?? Отрисовка границ + выравнивание по центру A:X
    ОбновитьСтатус 90, "Применение границ и выравнивания A:X..."
    Dim finalRow As Long: finalRow = wsRec.Cells(wsRec.Rows.Count, "A").End(xlUp).Row
    If finalRow >= REC_START_ROW Then
        Dim rngTarget As Range
        Set rngTarget = wsRec.Range(wsRec.Cells(REC_START_ROW, "A"), wsRec.Cells(finalRow, "X"))
        With rngTarget
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .weight = xlThin
            End With
        End With
    End If
    
    ' СОРТИРОВКА: A -> B -> C -> D -> I (по возрастанию)
    ОбновитьСтатус 95, "Сортировка таблицы..."
    If finalRow >= REC_START_ROW Then
        Dim sortRange As Range
        Set sortRange = wsRec.Range(wsRec.Cells(REC_START_ROW, "A"), wsRec.Cells(finalRow, "X"))
        
        ' Очищаем предыдущие правила сортировки, чтобы не накапливались
        wsRec.Sort.SortFields.Clear
        
        ' Добавляем ключи строго в вашем порядке приоритета
        wsRec.Sort.SortFields.Add key:=wsRec.Range(wsRec.Cells(REC_START_ROW, "A"), wsRec.Cells(finalRow, "A")), Order:=xlAscending
        wsRec.Sort.SortFields.Add key:=wsRec.Range(wsRec.Cells(REC_START_ROW, "B"), wsRec.Cells(finalRow, "B")), Order:=xlAscending
        wsRec.Sort.SortFields.Add key:=wsRec.Range(wsRec.Cells(REC_START_ROW, "C"), wsRec.Cells(finalRow, "C")), Order:=xlAscending
        wsRec.Sort.SortFields.Add key:=wsRec.Range(wsRec.Cells(REC_START_ROW, "D"), wsRec.Cells(finalRow, "D")), Order:=xlAscending
        wsRec.Sort.SortFields.Add key:=wsRec.Range(wsRec.Cells(REC_START_ROW, "I"), wsRec.Cells(finalRow, "I")), Order:=xlAscending
        
        With wsRec.Sort
            .SetRange sortRange
            .Header = xlNo ' Заголовки (строки 1-2) не входят в диапазон сортировки
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        LogMsg "Сортировка применена: A > B > C > D > I"
    End If
    
    ОбновитьСтатус 100, "[OK] Модуль 3 завершён."
    LogMsg "=== МОДУЛЬ 3 ЗАВЕРШЁН УСПЕШНО ===" & vbCrLf
End Sub
