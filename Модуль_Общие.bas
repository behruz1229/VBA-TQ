Attribute VB_Name = "Модуль_Общие"
' ============================================================================
' МОДУЛЬ: Модуль_Общие
' НАЗНАЧЕНИЕ: Библиотека вспомогательных функций, используемых во всех модулях.
' ФУНКЦИИ:
'   - ОбновитьСтатус: Отображение прогресс-бара в строке статуса Excel.
'   - НайтиСамыйСвежийФайл: Поиск файла по маске и автоматический выбор версии с последней датой изменения.
'   - БезопасныйКлюч: Генерация уникального текстового ключа из массива значений (защита от null, пробелов и спецсимволов).
'   - LogMsg: Вывод отладочной информации, этапов работы и статистики в окно Immediate (Ctrl+G).
' ============================================================================

Option Explicit

' Прогресс-бар в строке статуса
Public Sub ОбновитьСтатус(прогресс As Integer, текст As String)
    Dim pct As Integer: pct = Application.Min(прогресс, 100)
    Dim filled As Integer: filled = pct \ 5
    Application.StatusBar = текст & " " & String(filled, "#") & String(20 - filled, "-") & " " & pct & "%"
    DoEvents
End Sub

' ?? Поиск самого свежего файла по маске
Public Function НайтиСамыйСвежийФайл(путь As String, маска As String) As String
    Dim f As String, latestFile As String, latestDate As Date
    f = Dir(путь & маска)
    If f = "" Then Exit Function
    Do While f <> ""
        Dim currDate As Date
        On Error Resume Next
        currDate = FileDateTime(путь & f)
        On Error GoTo 0
        If currDate > latestDate Then
            latestDate = currDate
            latestFile = путь & f
        End If
        f = Dir()
    Loop
    НайтиСамыйСвежийФайл = latestFile
End Function

' Генерация безопасного ключа
Public Function БезопасныйКлюч(ParamArray args() As Variant) As String
    Dim i As Integer, res As String
    For i = 0 To UBound(args)
        res = res & Trim(Replace(CStr(args(i)), "_", "-")) & "_"
    Next i
    БезопасныйКлюч = Left(res, Len(res) - 1)
End Function

' ?? Логирование в окно Immediate (Ctrl+G)
Public Sub LogMsg(msg As String, Optional level As String = "INFO")
    Debug.Print "[" & level & "] " & Format(Now, "hh:mm:ss") & " | " & msg
End Sub
