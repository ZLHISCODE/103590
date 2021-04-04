VERSION 5.00
Begin VB.Form frmCureFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmCureFind.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����ҵĲο�Ŀ¼�ı��롢���ơ�������������롣����ڶ����������������һ����ֱ���ҵ���ϣ�����ҵ���Ŀ��"
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
Attribute VB_Name = "frmCureFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer

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
    Dim lng����ID As Long, lngĿ¼ID As Long
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
    gstrSql = "select distinct L.����ID,N.�ο�Ŀ¼ID" & _
            " from ���Ʋο�Ŀ¼ L,���Ʋο����� N" & _
            " where L.����=[1] and L.ID=N.�ο�Ŀ¼ID " & _
            "       and (L.���� like [2] " & _
            "            or N.���� like [3] " & _
            "            or N.���� like [3])"
    Err = 0: On Error GoTo ErrHand

    If strFind <> gstrSql Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        
        If rsFind.EOF Then
            MsgBox "�����ڲ��ҵĲο���", vbExclamation, gstrSysName
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblnote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
        strFind = gstrSql
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "�Ѳ��ҵ����һ���ο���", vbExclamation, gstrSysName
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblnote.Caption = ""
            Me.cboSource.SetFocus: Exit Sub
        End If
    End If
    Me.lblnote.Caption = "(�����ҵ�" & rsFind.RecordCount & "������ǰΪ��" & rsFind.AbsolutePosition & "��)"
    lng����ID = rsFind!����ID
    lngĿ¼ID = rsFind!�ο�Ŀ¼ID

    Call frmCureRefers.zlLocateItem(lng����ID, lngĿ¼ID)
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    Me.Tag = Val(frmCureRefers.tvwClass.Tag) + 1
    Me.Caption = "�ο�����..."
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
