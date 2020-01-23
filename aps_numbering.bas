Attribute VB_Name = "aps_numbering"
Option Explicit

Sub aps_numbering()
Attribute aps_numbering.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Val As Integer, FirstRow As Integer, FirstColumn As Long
Dim FinalColumn As Long, FinalRow As Long
Dim wsSh As Worksheet
Dim sVal As String, pVal As String, psVal '�������� �������, �������� � �������� ������ �������
Dim sstr As Integer '����������� �������� � �������
Dim USRange As Range
Dim i As Integer
Dim j As Integer
ActiveSheet.Copy Before:=ActiveSheet '�������� ���� ����� ��������
    
    On Error GoTo Errors1
    
    Set USRange = Application.InputBox _
    (prompt:="������� �������� ���������� �����", Type:=8) '������ ���������
    If USRange Is Nothing Then 'Exit Sub ����� ��� ���������� ���������
  
 'If IsNumeric(USRange(1, 1).Value2) Then '����� ���� ������ �������� ��������� �� ��������
 'FirstValue = USRange(1, 1).Value2 '������ ��������
 
 'Else
Errors1:       MsgBox ("�������� ��������")
        Application.DisplayAlerts = False '��������� ������ ��������� ��� "������" �������� �����
        ActiveSheet.Delete '"����" ������� ����
        Application.DisplayAlerts = True '���������� ��������� �������
        Exit Sub
    End If


FinalRow = USRange.Rows.Count ' ��������� ���������� ������
FinalColumn = USRange.Columns.Count '��������� ������ ���������� ������

Application.Calculation = xlCalculationManual '�������� �������� �������
Application.ScreenUpdating = False '���������� ���������� ������


sVal = 0
For i = 1 To FinalRow
    For j = 1 To FinalColumn
        psVal = USRange(i, j).Value2
        sstr = suf(USRange(i, j)) '������� ������ ��������� �������
        If sstr < 0 Then
            GoTo line1
            ElseIf sstr = 0 And sVal = 0 Then
            sVal = 1 '����� ��� ���������� ��������
            ElseIf sstr > 0 And sVal = 0 Then
            sVal = Right(USRange(i, j).Value2, sstr) '��������� ��������� ��������
        End If
        pVal = Left(psVal, (Len(psVal) - sstr))
        USRange(i, j).Value2 = pVal + sVal '������������ ��������
        sVal = sVal + 1
line1:
    Next j
Next i

On Error Resume Next '��� ������ ������ ������
Set wsSh = Sheets("���������") '������ ���������� ������ �� ������ - ���� "���������"
If wsSh Is Nothing Then '�������� ������� ����� "���������"
ActiveSheet.Name = "���������" '���� �������� ����� ��� "���������"
Else
Application.DisplayAlerts = False '���������� ����������
Sheets("���������").Delete
Application.DisplayAlerts = True '�������� ����������
ActiveSheet.Name = "���������" '���� �������� ����� ��� "���������"
End If

Application.Calculation = xlCalculationAutomatic '�������� �������� �������������
Application.ScreenUpdating = True '��������� ���������� ������
End Sub

Private Function suf(US As Range) '������� �� ����������� ������� ��������� �������
Dim USval As String
Dim i As Integer
i = 1
USval = US.Value2

Do While Not IsNumeric(USval) And Len(USval) > 0 '���� ���� �������� �� �������� ��� ����� ������ �� ����
 USval = Right(USval, Len(USval) - i) '������� ���� ������ � �����
 Loop '��������� ����, ����� �� ����� � �������� ������� (0 ���� ��� ���)
 
 Select Case US '����� ��������� ������
 Case "<>"
 suf = -1
 Case "."
 suf = -1
 US.Value2 = "<>"
 US.Interior.Color = vbYellow
 Case Else
 suf = Len(USval)
 End Select
 
 
End Function

'Private Function pref(USRange As Range)


'End Function
