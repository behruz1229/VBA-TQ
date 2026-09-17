Attribute VB_Name = "Модуль_Запуск"
' ============================================================================
' МОДУЛЬ: Модуль_Запуск
' НАЗНАЧЕНИЕ: Главный управляющий модуль и настройка интерфейса.
' ФУНКЦИИ:
'   - СоздатьКнопкуМеню / УдалитьКнопкуМеню: Добавление/удаление кнопки "Обновить таблицу" в контекстное меню ячеек (ПКМ).
'   - ОбновитьТаблицу: Точка входа макроса. Сбрасывает фильтры, отключает обновление экрана/автопересчёт,
'                       последовательно вызывает Модуль 1, 2 и 3. Выводит итоговое окно при успехе или ошибку при сбое.
'                       Автоматически восстанавливает настройки Excel после завершения.
' ============================================================================

Option Explicit

' Создание кнопки в контекстном меню
Public Sub СоздатьКнопкуМеню()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Обновить таблицу").Delete
    On Error GoTo 0
    
    Dim btn As Object
    Set btn = Application.CommandBars("Cell").Controls.Add(Type:=1, Before:=1)
    With btn
        .Caption = "Обновить таблицу"
        .FaceId = 2608
        .OnAction = "'" & ThisWorkbook.Name & "'!ОбновитьТаблицу"
        .BeginGroup = True
    End With
'    MsgBox " Кнопка 'Обновить таблицу' добавлена в начало контекстного меню.", vbInformation
End Sub

Public Sub УдалитьКнопкуМеню()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Обновить таблицу").Delete
    On Error GoTo 0
'    MsgBox "Кнопка удалена.", vbInformation
End Sub

' ГЛАВНАЯ ТОЧКА ВХОДА
Public Sub ОбновитьТаблицу()
    Dim success As Boolean: success = False
    On Error GoTo ErrHandler
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
        .StatusBar = ChrW(8987) & " Инициализация..."
    End With

    ' Сброс фильтров в приёмнике
    Dim wsRec As Worksheet
    On Error Resume Next
    Set wsRec = ThisWorkbook.Sheets("Ведомость элементов ТСБ и МОТ")
    On Error GoTo ErrHandler
    
    If wsRec Is Nothing Then
        MsgBox "Лист 'Ведомость элементов ТСБ и МОТ' не найден!", vbCritical
        GoTo CleanExit
    End If
    
    If wsRec.FilterMode Then wsRec.ShowAllData
    If wsRec.AutoFilterMode Then wsRec.AutoFilter.ShowAllData

    ' Пошаговый вызов модулей
    Модуль1_Обновление_Источники1_2
    Модуль2_Обновление_ПунктыTQ
    Модуль3_Обновление_Инспекции
    
    success = True
    GoTo CleanExit

ErrHandler:
    MsgBox "Ошибка выполнения: " & Err.Description & vbCrLf & "Код: " & Err.Number, vbCritical

CleanExit:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
    
    ' ? Показываем окно только при успешном завершении
    If success Then
        MsgBox "Обновление таблицы успешно завершено!" & vbCrLf & _
               "Все данные синхронизированы, границы и выравнивание применены.", _
               vbInformation, "Обновление завершено"
    End If
End Sub
