VERSION 5.00
Begin VB.Form frmLabSampleCheckFind 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "查找"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboFind 
      Height          =   300
      ItemData        =   "frmLabSampleCheckFind.frx":0000
      Left            =   1530
      List            =   "frmLabSampleCheckFind.frx":0002
      TabIndex        =   3
      Top             =   705
      Width           =   3800
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4275
      TabIndex        =   2
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一条(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2700
      TabIndex        =   1
      Top             =   1650
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   -30
      TabIndex        =   0
      Top             =   1485
      Width           =   5925
   End
   Begin VB.Label Label1 
      Caption         =   "查找内容"
      Height          =   180
      Left            =   675
      TabIndex        =   6
      Top             =   750
      Width           =   915
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望查找项目的编码、名称。如存在多条，可依序""查找下一条""，直到找到你希望查找的项目。"
      Height          =   420
      Left            =   1035
      TabIndex        =   5
      Top             =   90
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   615
      TabIndex        =   4
      Top             =   1215
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   105
      Picture         =   "frmLabSampleCheckFind.frx":0004
      Top             =   30
      Width           =   840
   End
End
Attribute VB_Name = "frmLabSampleCheckFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFind As New ADODB.Recordset
Private mrsFindRecord As New ADODB.Recordset
Private strCurSql As String
Private intCount As Integer
Private mstrSql As String                 '查找的SQL
Private mstrFind As String                '上次查找条件
Private mway As String                    '进入查找窗体的方式
Public Event Finded(ByVal blnFind As Boolean, ByVal strVale As String)

Public Function ShowFind(ByVal strSQL As String)
    '功能： 通用查找窗体，不处理查找到结果后的定位操作，定位操作需调用者在Finded事件中处理。
    
    If strSQL = "" Then Exit Function
    mstrSql = strSQL
    mway = "SQL"
    Me.Show vbModal
    
End Function

Public Function ShowFindRecordset(ByVal rsTmp As Recordset)
    '功能： 通用查找窗体，不处理查找到结果后的定位操作，定位操作需调用者在Finded事件中处理。
    Set mrsFindRecord = rsTmp
    mway = "记录集"
    Me.Show vbModal
    
End Function


Private Sub cboFind_Click()
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboFind_GotFocus()
    Me.cboFind.SelStart = 0: Me.cboFind.SelLength = 100
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
'    If InStr(gSysParameter.InvaidWord, Chr(KeyAscii)) > 0 Then KeyAscii = 0
'    If KeyAscii = vbKeyReturn Then Call ComPressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboFind.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer
    Dim strFind As String
    Dim strFilter As String
    Dim strReturn As String

    If Trim(Me.cboFind.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation
        Me.cboFind.SetFocus: Exit Sub
    End If
    strFind = ""
    For intCount = 0 To Me.cboFind.ListCount
        strFind = strFind & ";" & Me.cboFind.List(intCount)
    Next
    If InStr(1, strFind, ";" & Trim(Me.cboFind.Text)) = 0 Then
        Me.cboFind.AddItem Trim(Me.cboFind.Text), 0
    End If
    
    
    strFind = Trim(Me.cboFind.Text)
   
        
        
    Err = 0: On Error GoTo ErrHand
    Select Case mway
    Case "SQL"
        If mstrFind <> strFind Then
            strCurSql = ""
            mstrFind = strFind
        End If
        strReturn = ""
        With rsFind
            
            If strCurSql <> mstrSql Or .State <> adStateOpen Then
                If .State = adStateOpen Then .Close
                Set rsFind = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, CStr("%" & strFind & "%"))
                If rsFind.EOF Then
                    MsgBox "不存在查找的内容！", vbExclamation
                    rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                    Me.cboFind.SetFocus: Exit Sub
                End If
                strCurSql = mstrSql
            Else
                rsFind.MoveNext
                If rsFind.EOF Then
                    MsgBox "已查找到最后一条项目！", vbExclamation
                    rsFind.Close: Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                    Me.cboFind.SetFocus: Exit Sub
                End If
            End If
            Me.lblNote.Caption = "(共查找到" & rsFind.RecordCount & "条，当前为第" & rsFind.AbsolutePosition & "条)"
            For i = 0 To rsFind.Fields.Count - 1
                strReturn = strReturn & "," & rsFind.Fields(i).Value
            Next
        End With
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        RaiseEvent Finded(True, strReturn)
    Case "记录集"
        If mstrFind <> strFind Then
            mstrFind = strFind
            strReturn = ""
            strFilter = "仪器类型 like '" & strFind & "*' or 仪器名称 like '" & strFind & "*'"
            mrsFindRecord.filter = strFilter
        End If
        If mrsFindRecord.RecordCount = 0 Then
            MsgBox "不存在查找的内容！", vbExclamation
            Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboFind.SetFocus: Exit Sub
        ElseIf mrsFindRecord.RecordCount = 1 Then
            If mrsFindRecord.EOF Then
                MsgBox "已查找到最后一条项目！", vbExclamation
                Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                mstrFind = "": Me.cboFind.SetFocus: Exit Sub
            Else
                For i = 0 To mrsFindRecord.Fields.Count - 1
                    strReturn = strReturn & "," & mrsFindRecord.Fields(i).Value
                Next
                Me.lblNote.Caption = "(共查找到" & mrsFindRecord.RecordCount & "条，当前为第" & mrsFindRecord.AbsolutePosition & "条)"
            End If
            mrsFindRecord.MoveNext
        Else
            If mrsFindRecord.EOF Then
                MsgBox "已查找到最后一条项目！", vbExclamation
                Me.cboFind.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
                mstrFind = "": Me.cboFind.SetFocus: Exit Sub
            End If
            For i = 0 To mrsFindRecord.Fields.Count - 1
                strReturn = strReturn & "," & mrsFindRecord.Fields(i).Value
            Next
            Me.lblNote.Caption = "(共查找到" & mrsFindRecord.RecordCount & "条，当前为第" & mrsFindRecord.AbsolutePosition & "条)"
            mrsFindRecord.MoveNext
        End If
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        RaiseEvent Finded(True, strReturn)
    End Select
    Exit Sub
ErrHand:
'    If ComErrCenter() = 1 Then
'    Resume
'    End If
End Sub

Private Sub Form_Activate()
    Me.cboFind.SetFocus
End Sub

Private Sub Form_Load()
    strCurSql = ""
    Me.lblNote.Caption = ""
    mstrFind = ""
End Sub


Private Sub rsRecordset()

End Sub



