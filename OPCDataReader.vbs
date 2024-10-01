Option Explicit

' ���������� ���������� ����������
Dim OPCServer, OPCGroup, OPCItems, OPCItem

Sub Main()
    On Error Resume Next
    
    ' �������� ������� OPC �������
    Set OPCServer = CreateObject("OPC.Automation.1")
    CheckError "�������� ������� OPC �������"
    
    ' ����������� � �������
    OPCServer.Connect "OPCServer.WinCC.1", "192.168.0.102"
    CheckError "����������� � �������"
    
    ' ���������� ������
    Set OPCGroup = OPCServer.OPCGroups.Add("Group1")
    CheckError "���������� ������"
    
    ' ��������� ������
    OPCGroup.UpdateRate = 1000 ' ���������� ������ �������
    OPCGroup.IsActive = True
    
    ' ��������� ������� OPCItems
    Set OPCItems = OPCGroup.OPCItems
    
    ' ���������� ����
    Set OPCItem = OPCItems.AddItem("FIR1_121_PV_act", 1)
    CheckError "���������� ����"
    
    ' ������ �������� ����
    Dim Value, Quality, TimeStamp
    OPCItem.Read 1, Value, Quality, TimeStamp
    CheckError "������ �������� ����"
    
    ' ����� ������������ �������� � ��������������� ����
    If Err.Number = 0 Then
        WScript.Echo "���             | �������� | �������� | ��������� �����"
        WScript.Echo "FIR1_121_PV_act | " & PadRight(Value, 8) & " | " & PadRight(Quality, 8) & " | " & TimeStamp
    End If
    
    ' ���������� �� �������
    OPCServer.Disconnect
    Set OPCServer = Nothing
End Sub

' ������� ��� �������� � ������ ������
Sub CheckError(operation)
    If Err.Number <> 0 Then
        WScript.Echo "������ ��� ���������� �������� '" & operation & "':"
        WScript.Echo "��������: " & Err.Description
        WScript.Echo "��� ������: " & Err.Number
        WScript.Quit
    End If
End Sub

' ������� ��� ������������ ������ �� ������� ����
Function PadRight(value, length)
    PadRight = Left(value & Space(length), length)
End Function

' ������ �������� ���������
Main