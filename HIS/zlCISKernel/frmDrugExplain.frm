VERSION 5.00
Begin VB.Form frmDrugExplain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ҩƷ˵����"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13410
   Icon            =   "frmDrugExplain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   13410
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "frmDrugExplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngҩƷID As Long
Private mobjForm As Object

Public Sub ShowMe(ByVal lngҩƷID As Long, objParent As Object)
    mlngҩƷID = lngҩƷID
    Me.Show 0, objParent
End Sub

Private Sub Form_Load()
    Dim strUnitName As String, rsTmp As Recordset, strSQL As String
    Dim str��λ�� As String
    
    If gobjDrugExplain Is Nothing Then Exit Sub: Unload Me
    On Error GoTo errH
    strSQL = "Select A.��λ��,B.���� from ҩƷ��� A,�շ���ĿĿ¼ B where A.ҩƷID=B.ID And B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҩƷID)
    If rsTmp.RecordCount = 0 Then Exit Sub: Unload Me
    
    Me.Caption = "ҩƷ˵����-" & rsTmp!����
    str��λ�� = rsTmp!��λ�� & ""
    
    Set rsTmp = zlDatabase.OpenSQLRecord("Select Item,Text From Table(Cast(zltools.f_Reg_Info(0) As zlTools.t_Reg_Rowset)) Where Item='��λ����'", Me.Caption)
    If rsTmp.RecordCount = 0 Then Exit Sub: Unload Me
    strUnitName = rsTmp!Text
    Me.AutoRedraw = True
    Set mobjForm = gobjDrugExplain.GetControlHWND
    SetParent mobjForm.hwnd, Me.hwnd
    mobjForm.Show
    gobjDrugExplain.InitDataByADODB gcnOracle, strUnitName
    gobjDrugExplain.LoadContent str��λ��
    mobjForm.Move 0, 0, Me.Width - 240, Me.Height - 570
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    mobjForm.Move 0, 0, Me.Width - 240, Me.Height - 570
End Sub
