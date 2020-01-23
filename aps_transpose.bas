Attribute VB_Name = "aps_transpose"
Option Explicit



Sub aps_transpose()
Dim i As Integer, k As Integer, shl As Integer
Dim FinalColumn As Long, FinalRow As Long
Dim arshl As Range
Dim wsSh As Worksheet
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
ActiveSheet.Copy Before:=ActiveSheet
FinalRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row ' ��������� ���������� ������
 
For i = FinalRow To 2 Step -1 '�������� ���� �� ������� � ���� �����
FinalColumn = ActiveSheet.Cells(i, Columns.Count).End(xlToLeft).Column '��������� ������ ���������� ������
shl = Application.RoundUp(((FinalColumn - 3) / 4), 0) '����������� ������� �� ������� ����� ������ 4(-3 ���� ������ ��������)
If shl > 1 Then '������� ��� ������ �������� ��� �������
Range(Rows(i + 1), Rows(i + shl - 1)).Insert Shift:=xlDown '������� ��������� ��� ������ i + ����������� ������� -1
 
 If shl > 1 Then '������� ��� ������� ������ ���� ������ ������
 For k = 1 To (shl - 1) '���� ������� ����������� � ������ �� ���������� �������
 Cells(i + k, 2) = "���." & k
 Set arshl = Range(Cells(i, 4), Cells(i, 7)) '����� �������� � 4 �� 7 �������
 arshl.Offset(k, 0).Value = arshl.Offset(0, 4 * k).Value '���������� �������� ��������� arshl
 Next k
 End If
End If
Next i

FinalRow = ActiveSheet.Cells(Rows.Count, 10).End(xlUp).Row
FinalColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(1, 8), Cells(FinalRow, FinalColumn)).ClearContents
On Error Resume Next
Set wsSh = Sheets("���������")
If wsSh Is Nothing Then
ActiveSheet.Name = "���������"
Else
Application.DisplayAlerts = False
Sheets("���������").Delete
Application.DisplayAlerts = True
ActiveSheet.Name = "���������"
End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

