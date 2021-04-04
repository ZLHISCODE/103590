VERSION 5.00
Begin VB.Form frmDiagnoseFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
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
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4210
      TabIndex        =   3
      Top             =   1935
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "������һ��(&N)"
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
      Caption         =   "    ����ϣ�����ҵļ�����ϵı��롢���ơ�������������롣����ڶ����������������һ����ֱ���ҵ���ϣ�����ҵ���Ŀ��"
      Height          =   525
      Left            =   810
      TabIndex        =   6
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(�����ҵ�10������ǰΪ��1��)"
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
      Caption         =   "��������(&F)"
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
Dim mblnShowStop As Boolean         '�Ƿ���ʾͣ����Ŀ

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
    Dim lng����id As Long, lng���ID As Long
    Dim strTemp As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation, gstrSysName
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboSource.ListCount
        strTemp = strTemp & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    gstrSql = "select distinct R.����ID,R.���ID" & _
            " from �������Ŀ¼ L,����������� R,������ϱ��� N" & _
            " where L.���=[1] and R.���ID=L.ID and L.ID=N.���ID " & _
            "       and (L.���� like [2] " & _
            "            or N.���� like [3] " & _
            "            or N.���� like [3]) " & IIf(mblnShowStop = False, " and (L.����ʱ�� Is Null Or L.����ʱ�� >= To_Date('3000-01-01', 'yyyy-mm-dd'))", "")
    Err = 0: On Error GoTo ErrHand
    
    If strFind <> gstrSql Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        If rsFind.EOF Then
            MsgBox "�����ڲ��ҵ���ϲο���", vbExclamation, gstrSysName
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
        strFind = gstrSql
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "�Ѳ��ҵ����һ����ϲο���", vbExclamation, gstrSysName
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
    End If
    Me.lblNote.Caption = "(�����ҵ�" & rsFind.RecordCount & "������ǰΪ��" & rsFind.AbsolutePosition & "��)"
    lng����id = rsFind!����id
    lng���ID = rsFind!���ID

    Call frmDiagnoses.zlLocateItem(lng����id, lng���ID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    If Val(frmDiagnoses.tvwClass.Tag) = 0 Then
        Me.Tag = 1: Me.Caption = "��ҽ��ϲο�����..."
    Else
        Me.Tag = 2: Me.Caption = "��ҽ��ϲο�����..."
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
