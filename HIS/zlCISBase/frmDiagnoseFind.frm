VERSION 5.00
Begin VB.Form frmDiagnoseFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmDiagnoseFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4210
      TabIndex        =   3
      Top             =   1935
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一条(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   2
      Top             =   1935
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   30
      TabIndex        =   4
      Top             =   1785
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1860
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3435
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望查找的疾病诊断的编码、名称、别名或者其简码。如存在多条，可依序查找下一条，直到找到你希望查找的项目。"
      Height          =   525
      Left            =   810
      TabIndex        =   6
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   810
      TabIndex        =   5
      Top             =   1455
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmDiagnoseFind.frx":058A
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "查找内容(&F)"
      Height          =   180
      Left            =   810
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmDiagnoseFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer
Dim mblnShowStop As Boolean         '是否显示停用项目

Public Sub ShowFind(ByVal blnShowStop As Boolean)
    mblnShowStop = blnShowStop
    
    frmDiagnoseFind.Show vbModal, frmDiagnoses
End Sub
Private Sub cboSource_GotFocus()
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim lng分类id As Long, lng诊断ID As Long
    Dim strTemp As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboSource.ListCount
        strTemp = strTemp & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    gstrSql = "select distinct R.分类ID,R.诊断ID" & _
            " from 疾病诊断目录 L,疾病诊断属类 R,疾病诊断别名 N" & _
            " where L.类别=[1] and R.诊断ID=L.ID and L.ID=N.诊断ID " & _
            "       and (L.编码 like [2] " & _
            "            or N.名称 like [3] " & _
            "            or N.简码 like [3]) " & IIf(mblnShowStop = False, " and (L.撤档时间 Is Null Or L.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd'))", "")
    Err = 0: On Error GoTo ErrHand
    
    If strFind <> gstrSql Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        If rsFind.EOF Then
            MsgBox "不存在查找的诊断参考！", vbExclamation, gstrSysName
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
        strFind = gstrSql
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "已查找到最后一条诊断参考！", vbExclamation, gstrSysName
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
    End If
    Me.lblNote.Caption = "(共查找到" & rsFind.RecordCount & "条，当前为第" & rsFind.AbsolutePosition & "条)"
    lng分类id = rsFind!分类id
    lng诊断ID = rsFind!诊断ID

    Call frmDiagnoses.zlLocateItem(lng分类id, lng诊断ID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    If Val(frmDiagnoses.tvwClass.Tag) = 0 Then
        Me.Tag = 1: Me.Caption = "西医诊断参考查找..."
    Else
        Me.Tag = 2: Me.Caption = "中医诊断参考查找..."
    End If
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strFind = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
