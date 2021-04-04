VERSION 5.00
Begin VB.Form frmDrugExplain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "药品说明书"
   ClientHeight    =   10110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13410
   Icon            =   "frmDrugExplain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   13410
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmDrugExplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng药品ID As Long
Private mobjForm As Object

Public Sub ShowMe(ByVal lng药品ID As Long, objParent As Object)
    mlng药品ID = lng药品ID
    Me.Show 0, objParent
End Sub

Private Sub Form_Load()
    Dim strUnitName As String, rsTmp As Recordset, strSQL As String
    Dim str本位码 As String
    
    If gobjDrugExplain Is Nothing Then Exit Sub: Unload Me
    On Error GoTo errH
    strSQL = "Select A.本位码,B.名称 from 药品规格 A,收费项目目录 B where A.药品ID=B.ID And B.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng药品ID)
    If rsTmp.RecordCount = 0 Then Exit Sub: Unload Me
    
    Me.Caption = "药品说明书-" & rsTmp!名称
    str本位码 = rsTmp!本位码 & ""
    
    Set rsTmp = zlDatabase.OpenSQLRecord("Select Item,Text From Table(Cast(zltools.f_Reg_Info(0) As zlTools.t_Reg_Rowset)) Where Item='单位名称'", Me.Caption)
    If rsTmp.RecordCount = 0 Then Exit Sub: Unload Me
    strUnitName = rsTmp!Text
    Me.AutoRedraw = True
    Set mobjForm = gobjDrugExplain.GetControlHWND
    SetParent mobjForm.hwnd, Me.hwnd
    mobjForm.Show
    gobjDrugExplain.InitDataByADODB gcnOracle, strUnitName
    gobjDrugExplain.LoadContent str本位码
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
