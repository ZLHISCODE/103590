VERSION 5.00
Begin VB.Form Frm������ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ҩƷ"
   ClientHeight    =   2580
   ClientLeft      =   3135
   ClientTop       =   4320
   ClientWidth     =   6525
   Icon            =   "Frm������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   0
      ScaleHeight     =   2925
      ScaleWidth      =   6495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton CmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   600
         Picture         =   "Frm������.frx":020A
         TabIndex        =   18
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "��"
         Height          =   240
         Left            =   6050
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1695
         Width           =   255
      End
      Begin VB.TextBox TxtSelect���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1665
         Width           =   5120
      End
      Begin VB.CommandButton Cmd���� 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3780
         Picture         =   "Frm������.frx":0354
         TabIndex        =   15
         Top             =   2160
         Width           =   1100
      End
      Begin VB.CommandButton Cmd���� 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5190
         Picture         =   "Frm������.frx":049E
         TabIndex        =   17
         Top             =   2160
         Width           =   1100
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   2
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox Txtͨ������ 
         Height          =   300
         Left            =   4440
         MaxLength       =   40
         TabIndex        =   1
         Top             =   390
         Width           =   1875
      End
      Begin VB.TextBox TxtҩƷ���� 
         Height          =   300
         Left            =   1200
         TabIndex        =   0
         Top             =   390
         Width           =   1875
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1260
         Width           =   1875
      End
      Begin VB.TextBox Txt���� 
         Height          =   300
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1290
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblָ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ָ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3990
         TabIndex        =   14
         Top             =   1350
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Lbl��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   13
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʒ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3990
         TabIndex        =   11
         Top             =   900
         Width           =   360
      End
      Begin VB.Label LblҩƷ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   720
         TabIndex        =   10
         Top             =   450
         Width           =   360
      End
      Begin VB.Label Lblͨ������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ͨ������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   9
         Top             =   450
         Width           =   720
      End
   End
End
Attribute VB_Name = "Frm������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strTmp As String
Public StrBit As Byte '�ó�����ҵ�ƥ�䷽ʽ
Dim rsTmp As ADODB.Recordset
Private mfrmMain As Form    '������

Private Type Type_SQLCondition
    strͨ���� As String
    str���� As String
    str���� As String
    str���� As String
    str��� As String
    str���� As String
End Type

Private SQLCondition As Type_SQLCondition

Public Function GetSearch(ByVal FrmMain As Form, _
    ByRef strͨ���� As String, _
    ByRef str���� As String, _
    ByRef str���� As String, _
    ByRef str���� As String, _
    ByRef str��� As String, _
    ByRef str���� As String) As String
    strTmp = ""
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = strTmp
    
    strͨ���� = SQLCondition.strͨ����
    str���� = SQLCondition.str����
    str���� = SQLCondition.str����
    str���� = SQLCondition.str����
    str��� = SQLCondition.str���
    str���� = SQLCondition.str����
    
End Function
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub CmdSelect_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(TxtSelect����.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
    'Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩƷ������", gstrNodeNo)
    
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ������", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        TxtSelect����.SetFocus
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    TxtSelect����.Tag = 1
    TxtSelect����.Text = rsProvider!����
    Cmd����.SetFocus
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd����_Click()
    If LTrim(Txtͨ������) = "" And LTrim(TxtҩƷ����) = "" And LTrim(Txt����) = "" & _
        LTrim(Txt����) = "" And LTrim(txt���) = "" And LTrim(TxtSelect����) = "" Then MsgBox "����������һ����Ϣ��", vbInformation, gstrSysName
    strTmp = ""
    If LTrim(Txtͨ������) <> "" Then strTmp = strTmp & " And A.���� like [1] "
    If LTrim(TxtҩƷ����) <> "" Then strTmp = strTmp & " And A.���� like [2] "
    If LTrim(Txt����) <> "" Then strTmp = strTmp & " And B.���� like [3] "
    If LTrim(Txt����) <> "" Then strTmp = strTmp & " And B.���� like [4] "
    If LTrim(txt���) <> "" Then strTmp = strTmp & " And upper(A.���) like [5] "
    If LTrim(TxtSelect����) <> "" Then strTmp = strTmp & " And upper(A.����) like [6] "
    
    SQLCondition.strͨ���� = IIf(StrBit = "0", "%", "") & LTrim(Txtͨ������) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(TxtҩƷ����)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(Txt����)) & "%"
    SQLCondition.str��� = IIf(StrBit = "0", "%", "") & UCase(LTrim(txt���)) & "%"
    SQLCondition.str���� = IIf(StrBit = "0", "%", "") & UCase(LTrim(TxtSelect����)) & "%"
    
    Unload Me
End Sub

Private Sub Cmd����_Click()
    strTmp = ""
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    StrBit = GetSetting(appName:="ZLSOFT", Section:="����ģ��\����", Key:="����ƥ��", Default:="0")
End Sub

Private Sub Pic����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub TxtSelect����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(TxtSelect����.hWnd)
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(TxtSelect����) = "" Then Exit Sub
        TxtSelect���� = UCase(TxtSelect����)
        
        gstrSQL = "Select ���� as id ,����,���� From ҩƷ������ Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) Order By ����"
'        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ҩƷ������]", _
'                        IIf(gstrMatchMethod = "0", "%", "") & TxtSelect���� & "%", _
'                        TxtSelect���� & "%", gstrNodeNo)
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ������", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & TxtSelect���� & "%", TxtSelect���� & "%", gstrNodeNo)
        
        If blnCancel Then TxtSelect����.SetFocus: Exit Sub
        
        With rsTemp
            If rsTemp.State = 0 Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                TxtSelect����.SelStart = 0
                TxtSelect����.SelLength = Len(TxtSelect����)
                KeyCode = 0
                Exit Sub
            End If
            
            TxtSelect���� = IIf(IsNull(!����), "", !����)
            TxtSelect����.Tag = 1
            Cmd����.SetFocus
            
        End With
        
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub TxtSelect����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub Txtͨ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub

Private Sub TxtҩƷ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey (vbKeyTab)
End Sub
