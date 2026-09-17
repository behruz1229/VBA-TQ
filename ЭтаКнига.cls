VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Автоматически создаёт кнопку при открытии файла
    Call Модуль_Запуск.СоздатьКнопкуМеню
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Опционально: удаляет кнопку перед закрытием, чтобы не загрязнять контекстное меню
    Call Модуль_Запуск.УдалитьКнопкуМеню
End Sub
