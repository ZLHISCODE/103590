VERSION 5.00
Begin VB.Form frmLabItemSignificance 
   BorderStyle     =   0  'None
   Caption         =   "临床意义"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtSignificance 
      Height          =   2835
      Left            =   60
      MaxLength       =   4000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4140
   End
End
Attribute VB_Name = "frmLabItemSignificance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngItemID As Long          '当前显示的项目id

Public Function zlEditSave() As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    strSQL = "Zl_临床意义_Edit(" & mlngItemID & ",'" & DelInvalidChar(txtSignificance, "'") & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function zlEditStart() As Boolean
    Me.Tag = "编辑": Call Form_Resize
    zlEditStart = True: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlRefresh(lngItemID As Long) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    mlngItemID = lngItemID
    strSQL = "Select A.临床意义 From 检验报告项目 B,检验项目 A Where A.诊治项目id=B.报告项目id And B.诊疗项目id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngItemID)
    txtSignificance.Text = ""
    Do Until rsTmp.EOF
        txtSignificance.Text = Trim("" & rsTmp!临床意义)
        rsTmp.MoveNext
    Loop
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.txtSignificance
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub txtSignificance_KeyPress(KeyAscii As Integer)
    If Me.Tag = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
