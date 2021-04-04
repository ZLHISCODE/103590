VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2100
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmCheckUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picVisible 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   15
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   1950
      TabIndex        =   5
      Top             =   1050
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   248446979
      CurrentDate     =   43074
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   8
      Top             =   1380
      Width           =   5025
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3015
      TabIndex        =   7
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton CMDȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1905
      TabIndex        =   6
      Top             =   1650
      Width           =   1100
   End
   Begin VB.TextBox TXT���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   2115
   End
   Begin VB.TextBox txt�û� 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   2115
   End
   Begin VB.Image ImgAudit 
      Height          =   810
      Left            =   315
      Picture         =   "frmCheckUser.frx":000C
      Stretch         =   -1  'True
      Top             =   210
      Width           =   720
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "�˶�ʱ��"
      Height          =   180
      Left            =   1125
      TabIndex        =   4
      Top             =   1095
      Width           =   720
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   330
      Picture         =   "frmCheckUser.frx":0396
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1500
      TabIndex        =   2
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmCheckUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintTimes As Integer
Dim mstrCaption As String
Dim mstrUser As String

'ʱ��ؼ�����
Private mblnShowTime As Boolean
Private mstrShowTitle As String
Private mstrExeTitle As String
Private mstrExeTime As String
Private mstrCurDate As String
Private mstrFormat As String
Private mstrModeName As String

Public Function IsValidUser(ByVal strModeName As String, ByVal strTitle As String, Optional ByVal blnShowTime As Boolean = False, Optional ByVal strShowTitle As String = "", Optional ByVal strExeTitle As String = "", _
    Optional ByVal strExeTime As String, Optional ByVal strCurDate As String = "", Optional ByVal strFormat As String = "yyyy-MM-dd HH:mm") As String
    '�������ݸ�ʽ����¼�û���;�������Լ�����
    '������
    '         strModeName--����ģ������
    '         blnShowTime---�Ƿ���ʾ����ѡ�񣬸ò���ΪTRUEʱ����Ĳ�������Ч
    '         strShowTitle-���ڱ��⣻����Ҫ�޶���������ڲ���С��ĳ��ʱ��㣬�ɴ���strExeTime��strExeTitle��
    '         strCurDate-ȱʡ��ʾ�����ڣ�Ϊ����Ĭ��Ϊ��ǰʱ�䡣
    '         strFormat--������ʾ�ĸ�ʽ
    mstrUser = ""
    mintTimes = 1
    mstrModeName = strModeName
    mstrCaption = strTitle
    mblnShowTime = blnShowTime
    mstrShowTitle = strShowTitle
    mstrExeTitle = strExeTitle
    mstrExeTime = strExeTime
    mstrCurDate = strCurDate
    mstrFormat = strFormat
    If mstrFormat = "" Then mstrFormat = "yyyy-MM-dd HH:mm"
    Me.Show 1
    IsValidUser = mstrUser
End Function

Private Sub CMDȷ��_Click()
    Dim strSQL As String
    Dim strNote As String
    Dim strUserName As String
    Dim strServerName As String
    Dim strPassword As String
    Dim rsUser As New ADODB.Recordset
    Dim strCurDate As String
    On Error GoTo InputError
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(UCase(txt�û�.Text))
    strPassword = Trim(TXT����.Text)
    strServerName = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", "")
    
    '��Ч�ַ���Ч��
    If Len(strUserName) = 0 Then
        strNote = "�������û���"
        If txt�û�.Enabled And txt�û�.Visible Then txt�û�.SetFocus
        GoTo InputError
    End If
    
    If Len(strPassword) = 0 Then
        strNote = "����������"
        If TXT����.Enabled And TXT����.Visible Then TXT����.SetFocus
        GoTo InputError
    End If
    
    If mblnShowTime = True Then
        If IsDate(mstrExeTime) Then
            If Format(dtpDate.Value, mstrFormat) < Format(mstrExeTime, mstrFormat) Then
                strNote = mstrShowTitle & "����С��" & mstrExeTitle & "��" & Format(mstrExeTime, mstrFormat) & "��"
                If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
                GoTo InputError
            End If
        End If
        strCurDate = Format(gobjDatabase.Currentdate, mstrFormat)
        If Format(dtpDate.Value, mstrFormat) > strCurDate Then
            MsgBox mstrShowTitle & "���ܴ��ڵ�ǰʱ�䡾" & strCurDate & "��", vbInformation, gstrSysName
            If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
            GoTo InputError
        End If
        strCurDate = Format(dtpDate.Value, mstrFormat)
    End If
    
    SetConState False
    mintTimes = mintTimes + 1
     '�û���¼��֤
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, strUserName, strPassword, strNote) = False Then
        TXT����.Text = ""
        If TXT����.Enabled Then TXT����.SetFocus
        SetConState
        GoTo InputError
    End If
        
    '�޸�ע���
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & mstrModeName & "\" & mstrCaption, "�û���", strUserName)
    strSQL = " Select A.���� From ��Ա�� A,�ϻ���Ա�� B Where A.ID=B.��ԱID And B.�û���=[1] "
    Set rsUser = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��¼�û�����", strUserName)
    mstrUser = strUserName & "'" & rsUser!����
    If mblnShowTime = True Then mstrUser = mstrUser & "'" & strCurDate
    
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ������˳���", vbExclamation, gstrSysName
        CMD����_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If

End Sub

Private Sub CMD����_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Trim(txt�û�.Text) = "" Then
        CMDȷ��.Default = False
        txt�û�.SetFocus
    Else
        If TXT����.Enabled Then
            TXT����.SetFocus
        Else
            CMDȷ��.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl.name = "dtpDate" And KeyCode = vbKeyReturn Then picVisible.SetFocus: Exit Sub
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.name = "TXT����" Then
            Call CMDȷ��_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim lngTop As Long
    Me.Caption = mstrCaption
    txt�û�.Text = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & mstrModeName & "\" & mstrCaption, "�û���", "")
    
    If mblnShowTime = False Then
        lblDate.Visible = False
        dtpDate.Visible = False
        dtpDate.Enabled = False
        lngTop = TXT����.Top + TXT����.Height + 150
    Else
        lblDate.Visible = True
        dtpDate.Visible = True
        dtpDate.Enabled = True
        dtpDate.CustomFormat = mstrFormat
        lblDate.Caption = mstrShowTitle
        lngTop = dtpDate.Top + dtpDate.Height + 150
        If IsDate(mstrCurDate) Then
            dtpDate.Value = Format(mstrCurDate, mstrFormat)
        Else
            dtpDate.Value = Format(gobjDatabase.Currentdate, mstrFormat)
        End If
    End If
    Frame1.Top = lngTop
    CMDȷ��.Top = Frame1.Top + Frame1.Height + 150
    CMD����.Top = CMDȷ��.Top
    
    Me.Height = CMD����.Top + CMD����.Height + 535
    
    '�������������ȡ��������
    HookDefend TXT����.hWnd
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ActiveControl.name = "dtpDate" Then picVisible.SetFocus
End Sub

Private Sub txt�û�_Change()
    CMDȷ��.Default = False
End Sub

Private Sub TXT����_GotFocus()
    GetFocus TXT����
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD����.Enabled = BlnState
    CMDȷ��.Enabled = BlnState
End Sub
