VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRunOption 
   BackColor       =   &H80000005&
   Caption         =   "ϵͳ����ѡ��"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmRunOption.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   9870
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   495
      Begin VB.Image imgMain 
         Height          =   480
         Left            =   0
         Picture         =   "FrmRunOption.frx":04F9
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H80000005&
      Height          =   7080
      Left            =   930
      TabIndex        =   14
      Top             =   570
      Width           =   8490
      Begin VB.TextBox txtMsgToDays 
         Height          =   300
         Left            =   2100
         MaxLength       =   18
         TabIndex        =   35
         Tag             =   "26"
         Top             =   6334
         Width           =   1395
      End
      Begin VB.TextBox txtAuditLogDays 
         Height          =   300
         Left            =   2115
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "25"
         Top             =   2941
         Width           =   1755
      End
      Begin VB.CheckBox chkShutDown 
         BackColor       =   &H80000005&
         Caption         =   "����ر������ĵ���̨"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Tag             =   "24"
         Top             =   5475
         Width           =   4455
      End
      Begin VB.CheckBox chkLenCtrl 
         BackColor       =   &H80000005&
         Caption         =   "�������볤�ȿ���"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Tag             =   "20"
         Top             =   4329
         Width           =   1750
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "21"
         Text            =   "3"
         Top             =   4321
         Width           =   300
      End
      Begin VB.TextBox txtLen 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   2910
         MaxLength       =   2
         TabIndex        =   26
         Tag             =   "22"
         Text            =   "12"
         Top             =   4321
         Width           =   300
      End
      Begin VB.TextBox txtExcelRPTPath 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1830
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "6"
         Top             =   3605
         Width           =   3195
      End
      Begin VB.CommandButton CmdExcelRPTPath 
         Caption         =   "��"
         Height          =   300
         Left            =   5040
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3605
         Width           =   285
      End
      Begin VB.CheckBox chkRunLog 
         BackColor       =   &H80000005&
         Caption         =   "������־��¼(&S)"
         Height          =   210
         Left            =   240
         TabIndex        =   0
         Tag             =   "1"
         Top             =   315
         Width           =   1695
      End
      Begin VB.TextBox txtRunLogDays 
         Height          =   300
         Left            =   2325
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "2"
         Top             =   617
         Width           =   1755
      End
      Begin VB.CheckBox chkErrLog 
         BackColor       =   &H80000005&
         Caption         =   "��¼���д���(&A)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Tag             =   "3"
         Top             =   1341
         Width           =   1695
      End
      Begin VB.TextBox txtErrLogDays 
         Height          =   300
         Left            =   2325
         MaxLength       =   18
         TabIndex        =   5
         Tag             =   "4"
         Top             =   1613
         Width           =   1755
      End
      Begin VB.TextBox txtNoticeDays 
         Height          =   300
         Left            =   1965
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "5"
         Top             =   2277
         Width           =   1755
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   1
         Left            =   3210
         TabIndex        =   25
         Top             =   4321
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(1)"
         BuddyDispid     =   196616
         BuddyIndex      =   1
         OrigLeft        =   3240
         OrigTop         =   3855
         OrigRight       =   3495
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown udLen 
         Height          =   270
         Index           =   0
         Left            =   2340
         TabIndex        =   30
         Top             =   4321
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtLen(0)"
         BuddyDispid     =   196616
         BuddyIndex      =   0
         OrigLeft        =   2370
         OrigTop         =   3855
         OrigRight       =   2625
         OrigBottom      =   4110
         Max             =   16
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.CheckBox chkSpecial 
         BackColor       =   &H80000005&
         Caption         =   "���Ӷȿ���(���ٰ���һ�����֡���ĸ���������)"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Tag             =   "23"
         Top             =   4736
         Width           =   4455
      End
      Begin VB.Label lblMsgToDays 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ƽ̨��Ϣ��������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   37
         Top             =   6394
         Width           =   1800
      End
      Begin VB.Label lblMsgToDaysNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(���ڼ���ƽ̨��ҵ����Ϣ�ı�������������ʱϵͳ�����Զ�ɾ����Ϊ0���Զ�����7��)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   36
         Top             =   6735
         Width           =   6840
      End
      Begin VB.Label lblAuditLogDaysNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(������־����ܱ��������������ʱϵͳ�����Զ�ɾ�������ٱ���90�죬����Ϊ0ʱ��ʾ���ñ���)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   34
         Top             =   3333
         Width           =   7830
      End
      Begin VB.Label lblAuditLogDays 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������־�����������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   3001
         Width           =   1800
      End
      Begin VB.Label lblShutDown 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRunOption.frx":227B
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   480
         TabIndex        =   32
         Top             =   5882
         Width           =   5400
      End
      Begin VB.Label lblSpecialNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(���������ASCIIֵ��32-126֮����ַ���ɡ������е������ַ����ܰ���˫���š�@��\���ո�)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   29
         Top             =   5143
         Width           =   7650
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "-->"
         Height          =   135
         Left            =   2640
         TabIndex        =   27
         Top             =   4389
         Width           =   375
      End
      Begin VB.Label lblExcelRPTPathNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(��ѡ��������ϵ�ApplyĿ¼��Ϊ����·��)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   22
         Top             =   3997
         Width           =   3510
      End
      Begin VB.Label lblExcelRPTPath 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EXCEL������·��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   3665
         Width           =   1530
      End
      Begin VB.Label lblRunLogNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�Ƿ��Զ���¼�û���ʹ��ϵͳ�����)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   19
         Top             =   315
         Width           =   3060
      End
      Begin VB.Label lblRunLogDays 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��־��ౣ������(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   675
         Width           =   1710
      End
      Begin VB.Label lblRunLogDaysNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(ʹ����־��ౣ�������������ʱϵͳ���Զ�ɾ����ʱ�ļ�¼)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   840
         TabIndex        =   18
         Top             =   1005
         Width           =   5040
      End
      Begin VB.Label lblErrLogNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�Ƿ��¼ʹ�ù����з����ĸ��ִ���)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1950
         TabIndex        =   17
         Top             =   1341
         Width           =   3060
      End
      Begin VB.Label lblErrLogDays 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ౣ������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   1680
         Width           =   1710
      End
      Begin VB.Label lblErrLogDaysNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(�����¼��ౣ�������������ʱϵͳ���Զ�ɾ����ʱ�ļ�¼)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   840
         TabIndex        =   16
         Top             =   2010
         Width           =   5040
      End
      Begin VB.Label lblNoticeDays 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ�����������(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   2337
         Width           =   1710
      End
      Begin VB.Label lblNoticeDaysNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(��Ϣ����ܱ��������������ʱϵͳ�����Զ�ɾ��������Ϊ0ʱ��ʾ���ñ���)"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   2669
         Width           =   6210
      End
   End
   Begin VB.CommandButton cmdRestore 
      Cancel          =   -1  'True
      Caption         =   "��ԭ(&R)"
      Height          =   350
      Left            =   2190
      TabIndex        =   13
      Top             =   7845
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   900
      TabIndex        =   12
      Top             =   7845
      Width           =   1100
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ����ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   21
      Top             =   150
      Width           =   1440
   End
End
Attribute VB_Name = "FrmRunOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private recOption As New ADODB.Recordset

Private Sub chkLenCtrl_Click()
    Dim blnEnabled  As Boolean
    blnEnabled = (chkLenCtrl.value = 1)
    txtLen(0).Enabled = blnEnabled
    txtLen(1).Enabled = blnEnabled
    udLen(0).Enabled = blnEnabled
    udLen(1).Enabled = blnEnabled
    cmdSave.Enabled = True
End Sub

Private Sub chkShutDown_Click()
    cmdSave.Enabled = True
End Sub

Private Sub chkSpecial_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cmdExcelRPTPath_Click()
    Dim strPath As String
    strPath = OpenFolder(Me, "Excel������·����")
    If strPath = "" Then Exit Sub
    txtExcelRPTPath = strPath
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    If txtErrLogDays.Enabled = True And Val(txtErrLogDays.Text) > 10 ^ 8 Then
        MsgBox "������־�����Ŀ��̫��", vbInformation, gstrSysName
        txtErrLogDays.SetFocus
        Exit Sub
    End If
    If txtRunLogDays.Enabled = True And Val(txtRunLogDays.Text) > 10 ^ 8 Then
        MsgBox "������־�����Ŀ��̫��", vbInformation, gstrSysName
        txtRunLogDays.SetFocus
        Exit Sub
    End If
    If txtNoticeDays.Enabled = True And Val(txtNoticeDays.Text) > 10 ^ 8 Then
        MsgBox "��Ϣ�����Ŀ��̫��", vbInformation, gstrSysName
        txtNoticeDays.SetFocus
        Exit Sub
    End If
    If txtAuditLogDays.Enabled = True And Val(txtAuditLogDays.Text) > 10 ^ 8 Then
        MsgBox "������־��������̫��", vbInformation, gstrSysName
        txtAuditLogDays.SetFocus
        Exit Sub
    End If
    If StrIsValid(txtExcelRPTPath.Text, 50) = False Then
        txtExcelRPTPath.SetFocus
        Exit Sub
    End If
    If SaveCons = False Then Exit Sub
End Sub

Private Sub chkRunLog_Click()
    cmdSave.Enabled = True
    txtRunLogDays.Enabled = chkRunLog.value = 1
End Sub

Private Sub chkErrLog_Click()
    cmdSave.Enabled = True
    txtErrLogDays.Enabled = chkErrLog.value = 1
End Sub

Private Sub cmdRestore_Click()
    Call InitCons
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ActiveControl Is txtErrLogDays Or ActiveControl Is txtNoticeDays Or ActiveControl Is txtRunLogDays Or ActiveControl Is txtAuditLogDays Then
        If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call InitCons
End Sub

Private Sub InitCons()
    Dim ConThis As Control
    '--��ʼ�����ؼ���ֵ--
    Set recOption = gclsBase.OpenSQLRecord(gcnOracle, "Select Nvl(����ֵ, ȱʡֵ) Option_Value, ������, ������ From ZLTOOLS.zlOptions", "RunOption")
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            recOption.Filter = "������=" & Val(ConThis.Tag)
            With recOption
                If Val(ConThis.Tag) = 6 Then
                    ConThis.Enabled = Not (.EOF)
                    CmdExcelRPTPath.Enabled = Not (.EOF)
                End If
                
                Select Case TypeName(ConThis)
                Case "CheckBox"
                    If .EOF Then
                        ConThis.value = 0
                    Else
                        ConThis.value = IIf(IsNull(!Option_Value), 0, !Option_Value)
                    End If
                Case "TextBox"
                    If .EOF Then
                        ConThis.Text = ""
                    Else
                        ConThis.Text = IIf(IsNull(!Option_Value), "", !Option_Value)
                    End If
                End Select
            End With
        End If
    Next
    txtRunLogDays.Enabled = chkRunLog.value = 1
    txtErrLogDays.Enabled = chkErrLog.value = 1
    
    cmdSave.Enabled = False
End Sub

Private Function SaveCons() As Boolean
    Dim ConThis As Control, strValue As String
    Dim strNote As String, strName As String
    '--������ؼ���ֵ--
    
    SaveCons = False
    On Error Resume Next
    err = 0
    Set recOption = gclsBase.OpenSQLRecord(gcnOracle, "Select NVL(����ֵ,ȱʡֵ) Option_Value, ������, ������ From ZLTOOLS.zlOptions", "RunOption")
    gcnOracle.BeginTrans
    For Each ConThis In Controls
        If Val(ConThis.Tag) <> 0 Then
            recOption.Filter = "������=" & Val(ConThis.Tag)
            If Not recOption.EOF Then
                Select Case TypeName(ConThis)
                Case "CheckBox"
                    strValue = ConThis.value
                    If Nvl(recOption!Option_Value, 0) <> strValue Then strNote = strNote & "," & IIf(strValue = 1, "����", "ͣ��") & "�ˡ�" & Split(ConThis.Caption, "(")(0) & "��"
                Case "TextBox"
                    strValue = IIf(ConThis.Enabled = True, ConThis.Text, "")
                    If Nvl(recOption!Option_Value) <> strValue And ConThis.Enabled = True Then
                        strName = recOption!������
                        strNote = strNote & ",������" & strName & "���ɡ�" & Nvl(recOption!Option_Value) & "���޸�Ϊ�ˡ�" & strValue & "��"
                    End If
                End Select
                gcnOracle.Execute "Update ZLTOOLS.ZlOptions Set ����ֵ='" & strValue & "' Where ������=" & Val(ConThis.Tag)
            End If
        End If
    Next
    
    If err <> 0 Then
        MsgBox "�������в���ʱ����������", vbInformation, gstrSysName
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    gcnOracle.CommitTrans
    MsgBox "���в����޸ĳɹ���", vbInformation, gstrSysName
    '������Ҫ������־
    Call SaveAuditLog(2, "����", "���в����޸ĳɹ�" & strNote)
    cmdSave.Enabled = False
    SaveCons = True
End Function

Private Sub SelLen(ByVal ConObj As TextBox)
    With ConObj
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtLen_Change(Index As Integer)
    cmdSave.Enabled = True
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
End Sub

Private Sub txtLen_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtLen_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtLen(Index).Text) < udLen(Index).Min Then
        txtLen(Index).Text = udLen(Index).Min
    ElseIf Val(txtLen(Index).Text) > udLen(Index).Max Then
        txtLen(Index).Text = udLen(Index).Max
    End If
    If Val(txtLen(0).Text) > Val(txtLen(1).Text) Then
        If Index = 0 Then
            txtLen(1).Text = txtLen(0).Text
        Else
            txtLen(0).Text = txtLen(1).Text
        End If
    End If
    If Val(txtLen(1 - Index).Text) < udLen(1 - Index).Min Then
        txtLen(1 - Index).Text = udLen(1 - Index).Min
    ElseIf Val(txtLen(1 - Index).Text) > udLen(1 - Index).Max Then
        txtLen(1 - Index).Text = udLen(1 - Index).Max
    End If
End Sub

Private Sub txtAuditLogDays_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtAuditLogDays_GotFocus()
    SelLen txtAuditLogDays
End Sub

Private Sub txtAuditLogDays_LostFocus()
    '��Ϊ�䶯��־���ٱ���90�죬�ʵ��ı����е�ֵС��90ʱ��0���⣩�����Զ�����ֵ��Ϊ90
    txtAuditLogDays.Text = Val(txtAuditLogDays.Text)
    If txtAuditLogDays.Text < 90 And txtAuditLogDays.Text <> 0 Then
        txtAuditLogDays.Text = 90
    End If
End Sub

Private Sub txtExcelRPTPath_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtExcelRPTPath_GotFocus()
    SelAll txtExcelRPTPath
End Sub

Private Sub txtErrLogDays_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtErrLogDays_GotFocus()
    SelLen txtErrLogDays
End Sub

Private Sub txtMsgToDays_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtMsgToDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNoticeDays_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtNoticeDays_GotFocus()
    SelLen txtNoticeDays
End Sub

Private Sub txtRunLogDays_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtRunLogDays_GotFocus()
    SelLen txtRunLogDays
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

