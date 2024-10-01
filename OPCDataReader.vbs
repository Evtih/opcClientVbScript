Option Explicit

' Объявление глобальных переменных
Dim OPCServer, OPCGroup, OPCItems, OPCItem

Sub Main()
    On Error Resume Next
    
    ' Создание объекта OPC сервера
    Set OPCServer = CreateObject("OPC.Automation.1")
    CheckError "Создание объекта OPC сервера"
    
    ' Подключение к серверу
    OPCServer.Connect "OPCServer.WinCC.1", "192.168.0.102"
    CheckError "Подключение к серверу"
    
    ' Добавление группы
    Set OPCGroup = OPCServer.OPCGroups.Add("Group1")
    CheckError "Добавление группы"
    
    ' Настройка группы
    OPCGroup.UpdateRate = 1000 ' Обновление каждую секунду
    OPCGroup.IsActive = True
    
    ' Получение объекта OPCItems
    Set OPCItems = OPCGroup.OPCItems
    
    ' Добавление тега
    Set OPCItem = OPCItems.AddItem("FIR1_121_PV_act", 1)
    CheckError "Добавление тега"
    
    ' Чтение значения тега
    Dim Value, Quality, TimeStamp
    OPCItem.Read 1, Value, Quality, TimeStamp
    CheckError "Чтение значения тега"
    
    ' Вывод прочитанного значения в форматированном виде
    If Err.Number = 0 Then
        WScript.Echo "Тег             | Значение | Качество | Временная метка"
        WScript.Echo "FIR1_121_PV_act | " & PadRight(Value, 8) & " | " & PadRight(Quality, 8) & " | " & TimeStamp
    End If
    
    ' Отключение от сервера
    OPCServer.Disconnect
    Set OPCServer = Nothing
End Sub

' Функция для проверки и вывода ошибок
Sub CheckError(operation)
    If Err.Number <> 0 Then
        WScript.Echo "Ошибка при выполнении операции '" & operation & "':"
        WScript.Echo "Описание: " & Err.Description
        WScript.Echo "Код ошибки: " & Err.Number
        WScript.Quit
    End If
End Sub

' Функция для выравнивания текста по правому краю
Function PadRight(value, length)
    PadRight = Left(value & Space(length), length)
End Function

' Запуск основной процедуры
Main