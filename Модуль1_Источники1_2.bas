Attribute VB_Name = "Модуль1_Источники1_2"
' ============================================================================
' МОДУЛЬ: Модуль1_Источники1_2
' НАЗНАЧЕНИЕ: Синхронизация данных из мастер-файлов МОТ и ТСБ.
' ЛОГИКА РАБОТЫ:
'   1. Поиск и безопасное открытие последних версий файлов "*МОТ - Монтаж МК..." и "*ТСБ - Монтаж МК...".
'   2. Сканирование приёмника: сохранение существующих позиций (ключ B_C_D_F + индекс I) и значений столбца G.
'   3. Чтение источников: генерация уникальных индексов позиций, сравнение с приёмником.
'   4. Вставка: Недостающие или изменённые (по столбцу G) строки добавляются в конец таблицы (зелёная заливка RGB 102,255,102).
'   5. Проверка: Строки приёмника, отсутствующие в источниках или с изменённым G, помечаются оранжевым (RGB 255,153,51).
'   6. Файлы, открытые пользователем вручную, макрос не закрывает. Закрываются только те, что открыл скрипт.
' ============================================================================

Option Explicit

Public Sub Модуль1_Обновление_Источники1_2()
    Const PATH_SRC As String = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\!Ведомость элементов\"
    Const SHEET_REC As String = "Ведомость элементов ТСБ и МОТ"
    Const REC_START_ROW As Long = 3
    Const SRC_START_ROW As Long = 25
    
    LogMsg "=== ЗАПУСК МОДУЛЯ 1 ==="
    ОбновитьСтатус 5, "Поиск файлов источников..."
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileМОТ As String, fileТСБ As String
    fileМОТ = НайтиСамыйСвежийФайл(PATH_SRC, "*МОТ - Монтаж МК - Мастер файл*.xlsb")
    fileТСБ = НайтиСамыйСвежийФайл(PATH_SRC, "*ТСБ - Монтаж МК - Мастер файл*.xlsb")
    
    If fileМОТ <> "" Then LogMsg "Найден МОТ: " & fso.GetFileName(fileМОТ) Else LogMsg "МОТ не найден", "WARN"
    If fileТСБ <> "" Then LogMsg "Найден ТСБ: " & fso.GetFileName(fileТСБ) Else LogMsg "ТСБ не найден", "WARN"
    
    Dim missingFile As String
    If fileМОТ = "" Then missingFile = "МОТ-Мастер"
    If fileТСБ = "" Then missingFile = IIf(missingFile = "", "ТСБ-Мастер", "МОТ и ТСБ")
    If missingFile <> "" Then
        If MsgBox("Файл(ы) не найден: " & missingFile & vbNewLine & "Продолжить без него?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    
    ' 1. АНАЛИЗ ПРИЁМНИКА
    ОбновитьСтатус 10, "Анализ существующих позиций..."
    LogMsg ChrW(8635) & "Чтение приёмника..."
    Dim wsRec As Worksheet: Set wsRec = ThisWorkbook.Sheets(SHEET_REC)
    Dim lastRowRec As Long: lastRowRec = wsRec.Cells(wsRec.Rows.Count, "B").End(xlUp).Row
    If lastRowRec < REC_START_ROW Then LogMsg "Приёмник пуст. Выход.", "WARN": Exit Sub
    
    Dim dictRecG As Object: Set dictRecG = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = REC_START_ROW To lastRowRec
        Dim baseKeyRec As String
        baseKeyRec = БезопасныйКлюч(wsRec.Cells(i, "B").Value, wsRec.Cells(i, "C").Value, wsRec.Cells(i, "D").Value, wsRec.Cells(i, "F").Value)
        Dim idxRec As Variant: idxRec = wsRec.Cells(i, "I").Value
        If IsNumeric(idxRec) And idxRec <> "" Then
            dictRecG(baseKeyRec & "|" & CLng(idxRec)) = wsRec.Cells(i, "G").Value
        End If
        If i Mod 500 = 0 Then ОбновитьСтатус 10 + Int(i / lastRowRec * 15), "Анализ приёмника..."
    Next i
    LogMsg "Приёмник: " & dictRecG.Count & " позиций (сохранены значения G для сверки)"
    
    ' 2. ЧТЕНИЕ ИСТОЧНИКОВ
    ОбновитьСтатус 25, "Чтение источников (МОТ/ТСБ)..."
    Dim arrFiles As Variant: arrFiles = Array(fileМОТ, fileТСБ)
    Dim arrSheets As Variant: arrSheets = Array("TQ МОТ", "TQ ТСБ")
    Dim dictSrcData As Object: Set dictSrcData = CreateObject("Scripting.Dictionary")
    Dim dictSrcCounter As Object: Set dictSrcCounter = CreateObject("Scripting.Dictionary")
    Dim dictOpenedByMacro As Object: Set dictOpenedByMacro = CreateObject("Scripting.Dictionary")
    
    Dim f As Integer
    For f = 0 To 1
        If Len(Dir(arrFiles(f))) = 0 Then GoTo NextSource
        ОбновитьСтатус 25 + (f * 25), "Подключение: " & ChrW(9203) & " " & arrSheets(f)
        LogMsg "Обработка: " & ChrW(8987) & arrSheets(f)
        
        Dim wbSrc As Workbook, fname As String, isUserOpen As Boolean
        fname = fso.GetFileName(arrFiles(f))
        Set wbSrc = Nothing: isUserOpen = False
        
        On Error Resume Next
        Set wbSrc = Workbooks(fname)
        If Not wbSrc Is Nothing Then isUserOpen = True
        On Error GoTo 0
        
        If Not isUserOpen Then
            LogMsg "Файл не открыт. Открываю в ReadOnly..."
            On Error Resume Next
            Set wbSrc = Workbooks.Open(arrFiles(f), ReadOnly:=True, UpdateLinks:=False)
            On Error GoTo 0
            If wbSrc Is Nothing Then
                LogMsg "ОШИБКА: Не удалось открыть " & fname, "ERROR"
                GoTo NextSource
            End If
            dictOpenedByMacro(arrFiles(f)) = True
            LogMsg "Файл успешно открыт макросом."
        Else
            LogMsg "Файл уже открыт пользователем. Использую текущий экземпляр (не буду закрывать)."
        End If
        
        Dim wsSrc As Worksheet
        On Error Resume Next
        Set wsSrc = wbSrc.Sheets(arrSheets(f))
        On Error GoTo 0
        If wsSrc Is Nothing Then
            LogMsg "Лист '" & arrSheets(f) & "' не найден в файле.", "WARN"
            GoTo NextSource
        End If
        
        Dim lastRowSrc As Long: lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
        If lastRowSrc >= SRC_START_ROW Then
            Dim r As Long
            For r = SRC_START_ROW To lastRowSrc
                If Trim(CStr(wsSrc.Cells(r, "B").Value)) = "" Then GoTo SkipRow
                
                Dim baseKeySrc As String
                baseKeySrc = БезопасныйКлюч(wsSrc.Cells(r, "C").Value, arrSheets(f), wsSrc.Cells(r, "E").Value, wsSrc.Cells(r, "I").Value)
                
                Dim seqIdx As Long
                seqIdx = dictSrcCounter(baseKeySrc) + 1
                dictSrcCounter(baseKeySrc) = seqIdx
                
                Dim fullKeySrc As String: fullKeySrc = baseKeySrc & "|" & seqIdx
                ' Array: 0=B->A, 1=C->B, 2=Sheet->C, 3=E->D, 4=I->F, 5=J->G, 6=SeqIdx->I (V исключён)
                dictSrcData(fullKeySrc) = Array(wsSrc.Cells(r, "B").Value, wsSrc.Cells(r, "C").Value, arrSheets(f), _
                                                wsSrc.Cells(r, "E").Value, wsSrc.Cells(r, "I").Value, _
                                                wsSrc.Cells(r, "J").Value, seqIdx)
SkipRow:
                If r Mod 500 = 0 Then ОбновитьСтатус 25 + (f * 25) + Int(r / lastRowSrc * 20), "Чтение источников..."
            Next r
            LogMsg arrSheets(f) & ": обработано строк данных: " & (lastRowSrc - SRC_START_ROW + 1) & " (позиций: " & dictSrcCounter.Count & ")"
        Else
            LogMsg arrSheets(f) & ": данных нет (строки < " & SRC_START_ROW & ").", "WARN"
        End If
NextSource:
    Next f
    
    ' 3. ВСТАВКА НЕДОСТАЮЩИХ ИЛИ ИЗМЕНЁННЫХ ПОЗИЦИЙ
    ОбновитьСтатус 75, "Восстановление недостающих позиций..."
    LogMsg "Поиск недостающих/изменённых записей..."
    Dim newRowsColl As Collection: Set newRowsColl = New Collection
    Dim key As Variant
    
    For Each key In dictSrcData.Keys
        Dim shouldInsert As Boolean: shouldInsert = False
        If Not dictRecG.Exists(key) Then
            shouldInsert = True
        Else
            Dim srcG As String, recG As String
            srcG = Trim(CStr(dictSrcData(key)(5)))
            recG = Trim(CStr(dictRecG(key)))
            If srcG <> recG Then shouldInsert = True
        End If
        
        If shouldInsert Then newRowsColl.Add dictSrcData(key)
    Next key
    LogMsg "Найдено позиций для восстановления/корректировки: " & newRowsColl.Count
    
    Dim nextRow As Long: nextRow = lastRowRec + 1
    Dim nd As Variant
    If newRowsColl.Count > 0 Then
        For Each nd In newRowsColl
            With wsRec
                .Cells(nextRow, "A").Value = nd(0)
                .Cells(nextRow, "B").Value = nd(1)
                .Cells(nextRow, "C").Value = nd(2)
                .Cells(nextRow, "D").Value = nd(3)
                .Cells(nextRow, "F").Value = nd(4)
                .Cells(nextRow, "G").Value = nd(5)
                .Cells(nextRow, "I").Value = nd(6)
                .Range(.Cells(nextRow, "A"), .Cells(nextRow, "X")).Interior.Color = RGB(102, 255, 102)
                nextRow = nextRow + 1
            End With
        Next nd
    End If
    
    ' 4. ОРАНЖЕВАЯ ПОМЕТКА
    ОбновитьСтатус 90, "Проверка изменённых элементов..."
    Dim origLastRow As Long: origLastRow = lastRowRec
    Dim orangeCount As Long: orangeCount = 0
    
    For i = REC_START_ROW To origLastRow
        Dim baseKey As String
        baseKey = БезопасныйКлюч(wsRec.Cells(i, "B").Value, wsRec.Cells(i, "C").Value, wsRec.Cells(i, "D").Value, wsRec.Cells(i, "F").Value)
        Dim idxVal As Variant: idxVal = wsRec.Cells(i, "I").Value
        
        Dim shouldHighlight As Boolean: shouldHighlight = False
        If IsNumeric(idxVal) And idxVal <> "" Then
            Dim fullKey As String: fullKey = baseKey & "|" & CLng(idxVal)
            If Not dictSrcData.Exists(fullKey) Then
                shouldHighlight = True
            Else
                Dim srcG2 As String, recG2 As String
                srcG2 = Trim(CStr(dictSrcData(fullKey)(5)))
                recG2 = Trim(CStr(wsRec.Cells(i, "G").Value))
                If srcG2 <> recG2 Then shouldHighlight = True
            End If
        Else
            shouldHighlight = True
        End If
        
        If shouldHighlight Then
            wsRec.Range(wsRec.Cells(i, "A"), wsRec.Cells(i, "X")).Interior.Color = RGB(255, 153, 51)
            orangeCount = orangeCount + 1
        End If
        
        If i Mod 500 = 0 Then ОбновитьСтатус 90 + Int(i / origLastRow * 8), "Проверка изменённых..."
    Next i
    LogMsg "Помечено как изменённых/отсутствующих (оранжевым): " & orangeCount
    
    ' 5. ЗАКРЫТИЕ ТОЛЬКО ФАЙЛОВ, ОТКРЫТЫХ МАКРОСОМ
    Dim openedPath As Variant
    For Each openedPath In dictOpenedByMacro.Keys
        On Error Resume Next
        Workbooks(fso.GetFileName(openedPath)).Close False
        On Error GoTo 0
        LogMsg "Закрыт файл (открыт макросом): " & fso.GetFileName(openedPath)
    Next openedPath
    
    ОбновитьСтатус 100, "[OK] Модуль 1 завершён."
    LogMsg "=== МОДУЛЬ 1 ЗАВЕРШЁН УСПЕШНО ===" & vbCrLf
End Sub

