VERSION 5.00
Begin VB.Form frmMediFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   2490
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   ControlBox      =   0   'False
   Icon            =   "frmMediFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk��� 
      Alignment       =   1  'Right Justify
      Caption         =   "���ҹ��ҩƷ"
      Height          =   210
      Left            =   810
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4210
      TabIndex        =   6
      Top             =   2085
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "������һ��(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   5
      Top             =   2085
      Width           =   1530
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   30
      TabIndex        =   4
      Top             =   1935
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1905
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3405
   End
   Begin VB.Label lblComment 
      Caption         =   "    ����ϣ�����ҵ�ҩƷ���롢���ơ�������������롣����ڶ����������������һ����ֱ���ҵ���ϣ�����ҵ�ҩƷ��"
      Height          =   525
      Left            =   855
      TabIndex        =   7
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(�����ҵ�10������ǰΪ��1��)"
      Height          =   180
      Left            =   855
      TabIndex        =   3
      Top             =   1635
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmMediFind.frx":058A
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "��������(&F)"
      Height          =   180
      Left            =   855
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmMediFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFind As New ADODB.Recordset
Dim strFind As String
Dim intCount As Integer
Private mbln��ʾͣ��ҩƷ As Boolean
Private mblnSelfMedi As Boolean  '�Թ�ҩ true-�Թ�ҩ false-�����Թ�ҩ

Private Sub cboSource_Click()
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub cboSource_GotFocus()
    Me.cboSource.SelStart = 0: Me.cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub chk���_Click()
    If Trim(Me.cboSource.Text) <> "" Then
        Me.cmdFind.Enabled = True
    Else
        Me.cmdFind.Enabled = False
    End If
End Sub

Private Sub chk���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    Dim lng����id As Long, lngҩ��id As Long, lngҩƷID As Long
    Dim strTemp As String
    
    If Trim(Me.cboSource.Text) = "" Then
        MsgBox "��������ҵ�����", vbExclamation, gstrSysName
        On Error Resume Next
        Me.Show
        Me.cboSource.SetFocus: Exit Sub
    End If
    strTemp = ""
    For intCount = 0 To Me.cboSource.ListCount
         strTemp = strTemp & ";" & Me.cboSource.List(intCount)
    Next
    If InStr(1, strTemp, ";" & Trim(Me.cboSource.Text)) = 0 Then
        Me.cboSource.AddItem Trim(Me.cboSource.Text), 0
    End If
    
    If Me.chk���.Value = 0 Then
        If mblnSelfMedi = True Then '�Թ�ҩ
            gstrSql = "Select Distinct i.����id, i.Id As ҩ��id, 0 As ҩƷid" & vbNewLine & _
                    "From ������ĿĿ¼ I, ������Ŀ���� N, ҩƷ���� A" & vbNewLine & _
                    "Where i.Id = n.������Ŀid And i.Id = a.ҩ��id And a.�ٴ��Թ�ҩ = 1 And i.��� = [1] And" & vbNewLine & _
                    "      (i.���� Like [2] Or n.���� Like [2] Or n.���� Like [2])"

        Else
            gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,0 AS ҩƷID" & _
                    " FROM ������ĿĿ¼ I,������Ŀ���� N" & _
                    " WHERE I.ID=N.������ĿID " & _
                    " AND I.���=[1] " & _
                    " AND (I.���� LIKE [2] " & _
                    "     OR N.���� LIKE [3] " & _
                    "     OR N.���� LIKE [3])"
        End If
    Else
        If mblnSelfMedi = True Then '�Թ�ҩ
            gstrSql = "Select Distinct i.����id, i.Id As ҩ��id, d.Id As ҩƷid" & vbNewLine & _
                    "From ������ĿĿ¼ I, ҩƷ��� T, ҩƷ���� A, �շ���ĿĿ¼ D, �շ���Ŀ���� N" & vbNewLine & _
                    "Where i.Id = t.ҩ��id And i.Id = a.ҩ��id And t.ҩƷid = d.Id And t.ҩƷid = n.�շ�ϸĿid And i.��� = [1] And" & vbNewLine & _
                    "      (d.���� Like [2] Or n.���� Like[2] Or n.���� Like[2]) And a.�ٴ��Թ�ҩ = 1"

        Else
            gstrSql = "SELECT DISTINCT I.����ID,I.ID AS ҩ��ID,D.ID AS ҩƷID  " & _
                     " FROM ������ĿĿ¼ I,ҩƷ��� T,�շ���ĿĿ¼ D,�շ���Ŀ���� N " & _
                     " WHERE I.ID=T.ҩ��ID And T.ҩƷID=D.ID AND T.ҩƷID=N.�շ�ϸĿID  " & _
                     " AND I.���=[1] " & _
                     " AND (D.���� LIKE [2] " & _
                     "     OR N.���� LIKE [3] " & _
                     "     OR N.���� LIKE [3])"
        End If
    End If
    If Not mbln��ʾͣ��ҩƷ Then
        gstrSql = gstrSql & " And (I.����ʱ�� Is NULL Or to_Char(I.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
        If Me.chk���.Value = 1 Then gstrSql = gstrSql & " And (D.����ʱ�� Is NULL Or to_Char(D.����ʱ��,'yyyy-MM-dd')='3000-01-01')"
    End If
    
    Err = 0: On Error GoTo errHand
 
    If strFind <> chk���.Value & ";" & Trim(Me.cboSource.Text) Or rsFind.State <> adStateOpen Then
        Set rsFind = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, gstrMatch & Trim(Me.cboSource.Text) & "%", gstrMatch & Trim(Me.cboSource.Text) & "%")
        
        If rsFind.EOF Then
            MsgBox "�����ڲ��ҵ�ҩƷ��", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            rsFind.Close: Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus
            Exit Sub
        End If
        strFind = chk���.Value & ";" & Trim(Me.cboSource.Text)
    Else
        rsFind.MoveNext
        If rsFind.EOF Then
            MsgBox "�Ѳ��ҵ����һ��ҩƷ��", vbExclamation, gstrSysName
            On Error Resume Next
            Me.Show
            rsFind.Close: Me.cboSource.Text = "": Me.cmdFind.Enabled = False: Me.lblNote.Caption = ""
            Me.cboSource.SetFocus
            Exit Sub
        End If
    End If
    Me.lblNote.Caption = "(�����ҵ�" & rsFind.RecordCount & "������ǰΪ��" & rsFind.AbsolutePosition & "��)"
    lng����id = rsFind!����ID
    lngҩ��id = rsFind!ҩ��ID
    lngҩƷID = rsFind!ҩƷid
           
    
    Me.Hide
    Call frmMediLists.zlLocateItem(lng����id, lngҩ��id, lngҩƷID)
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub Form_Activate()
    Select Case Val(frmMediLists.tvwClass.Tag)
    Case 0
        Me.Tag = 5: Me.Caption = "����ҩ����..."
    Case 1
        Me.Tag = 6: Me.Caption = "�г�ҩ����..."
    Case 2
        Me.Tag = 7: Me.Caption = "�в�ҩ����..."
    End Select
    Me.cboSource.SetFocus
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    strFind = ""
    Me.lblNote.Caption = ""
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnͣ�� As Boolean, ByVal blnSelMedi As Boolean)
    mbln��ʾͣ��ҩƷ = blnͣ��
    mblnSelfMedi = blnSelMedi
    Me.Show , frmParent
End Sub

Public Sub FindNext()
    Call cmdFind_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
