Attribute VB_Name = "Module2"
Option Explicit

Dim wbThis As Workbook
Dim wsSource As Worksheet

Dim rgSourcePharmacyTitle As Range
Dim rgSourceDrugTitle As Range
Dim rgSourcePackFirst As Range

Dim firstPharmacy As Long
Dim lastPharmacy As Long
Dim numPharmacy As Long

Dim firstDrug As Long
Dim lastDrug As Long
Dim numDrug As Long


Sub InitData()
  
  Set wbThis = Workbooks(ThisWorkbook.Name)
  Set wsSource = wbThis.Worksheets("������")
  
  ' ��������� ������ � "������"
  Set rgSourcePharmacyTitle = wsSource.Range("C2")
  Set rgSourceDrugTitle = wsSource.Range("A3")
  Set rgSourcePackFirst = wsSource.Range("D4")
  
  ' �������� ��� ���� ����� ��������� �� �������� � "������"
  firstPharmacy = rgSourcePharmacyTitle.Column
  lastPharmacy = wsSource.Cells(rgSourcePharmacyTitle.Row, Columns.Count).End(xlToLeft).Column
  numPharmacy = lastPharmacy - firstPharmacy
 
  ' �������� ��� ���� ��������� ��������� �� �������� � "������"
  firstDrug = rgSourceDrugTitle.Row
  lastDrug = wsSource.Cells(Rows.Count, rgSourceDrugTitle.Column).End(xlUp).Row
  numDrug = lastDrug - firstDrug

End Sub

Sub Solution1()

  Dim wsTarget As Worksheet
  
  Dim rgTargetPharmacyTitle As Range
  Dim rgTargetDrugTitle As Range
  Dim rgTargetPackTitle As Range
 
  Dim counterPharmacy As Long
  Dim counterDrug As Long
  Dim iterTarget As Long
  
  Call InitData
 
  ' ��������� ������ � "��������"
  Set wsTarget = wbThis.Worksheets("��������")
  Set rgTargetPharmacyTitle = wsTarget.Range("A1")
  Set rgTargetDrugTitle = wsTarget.Range("B1")
  Set rgTargetPackTitle = wsTarget.Range("C1")
 
  For counterPharmacy = 1 To numPharmacy
    For counterDrug = 1 To numDrug
      iterTarget = (counterPharmacy - 1) * numDrug + counterDrug
      rgTargetPharmacyTitle.Offset(iterTarget, 0).Value = rgSourcePharmacyTitle.Offset(0, counterPharmacy).Value
      rgTargetDrugTitle.Offset(iterTarget, 0).Value = rgSourceDrugTitle.Offset(counterDrug, 0).Value
      rgTargetPackTitle.Offset(iterTarget, 0).Value = rgSourcePackFirst.Offset(counterDrug - 1, counterPharmacy - 1).Value
    Next
  Next
 
End Sub

Sub Solution2()
  
  Dim wsTarget As Worksheet
  
  Dim rgTargetPharmacyTitle As Range
  Dim rgTargetDrugTitle As Range
  Dim rgTargetPackTitle As Range
  
  Dim counterPharmacy As Long
  Dim counterDrug As Long
  
  Dim strCellFrom As String
  Dim strCellTo As String
  Dim strFormula As String
  
  
  Call InitData
 
  Set wsTarget = wbThis.Worksheets("�������")
  
  '''''''''''''''''''''''''''''''''''''''''''''
  ' ��������� ������ ��� ������� � "�������": '
  ' ����� ���������� �� �������               '
  '''''''''''''''''''''''''''''''''''''''''''''
  
  Set rgTargetPharmacyTitle = wsTarget.Range("B11")
  Set rgTargetPackTitle = wsTarget.Range("C11")
  
  rgTargetPharmacyTitle.Value = "��� ������"
  rgTargetPackTitle.Value = "����� ����������"
  
  rgTargetPharmacyTitle.Columns.AutoFit
  rgTargetPackTitle.Columns.AutoFit
  
  For counterPharmacy = 1 To numPharmacy
    ' ��� ������
    rgTargetPharmacyTitle.Offset(counterPharmacy, 0).Value = rgSourcePharmacyTitle.Offset(0, counterPharmacy)
    '����� ����������
    strCellFrom = wsSource.Name & "!" & rgSourcePackFirst.Offset(0, counterPharmacy - 1).Address
    strCellTo = wsSource.Name & "!" & rgSourcePackFirst.Offset(numDrug - 1, counterPharmacy - 1).Address
    strFormula = "=SUM(" & strCellFrom & ":" & strCellTo & ")"
    rgTargetPackTitle.Offset(counterPharmacy, 0).Value = strFormula
  Next
  
  '''''''''''''''''''''''''''''''''''''''''''''
  ' ��������� ������ ��� ������� � "�������": '
  ' ����� ���������� �� ����������            '
  '''''''''''''''''''''''''''''''''''''''''''''
  
  Set rgTargetDrugTitle = wsTarget.Range("E11")
  Set rgTargetPackTitle = wsTarget.Range("F11")
  
  rgTargetDrugTitle.Value = "��� ���������"
  rgTargetPackTitle.Value = "����� ����������"
  
  rgTargetDrugTitle.Columns.AutoFit
  rgTargetPackTitle.Columns.AutoFit
  
   
  For counterDrug = 1 To numDrug
    ' ��� ���������
    rgTargetDrugTitle.Offset(counterDrug, 0).Value = rgSourceDrugTitle.Offset(counterDrug, 0)
    ' ����� ����������
    strCellFrom = wsSource.Name & "!" & rgSourcePackFirst.Offset(counterDrug - 1, 0).Address
    strCellTo = wsSource.Name & "!" & rgSourcePackFirst.Offset(counterDrug - 1, numPharmacy - 1).Address
    strFormula = "=SUM(" & strCellFrom & ":" & strCellTo & ")"
    rgTargetPackTitle.Offset(counterDrug, 0).Value = strFormula
  Next

End Sub
