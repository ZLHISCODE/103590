VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAppReplant 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ��ϵͳ��ֲ"
   ClientHeight    =   4416
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7140
   Icon            =   "frmAppReplant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   0
      Left            =   1305
      TabIndex        =   4
      Top             =   -120
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   465
         Width           =   5800
      End
      Begin VB.Frame fraSys 
         Height          =   1005
         Left            =   765
         TabIndex        =   26
         Top             =   2340
         Width           =   3930
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "�汾�ţ�"
            Height          =   180
            Left            =   210
            TabIndex        =   28
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblSysName 
            AutoSize        =   -1  'True
            Caption         =   "ϵͳ����"
            Height          =   180
            Left            =   210
            TabIndex        =   27
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdSetupFile 
         Caption         =   "ѡ��(&S)��"
         Height          =   350
         Left            =   765
         TabIndex        =   5
         Top             =   1980
         Width           =   1260
      End
      Begin VB.Label lblSetupFile 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   765
         TabIndex        =   18
         Top             =   1650
         Width           =   3930
      End
      Begin VB.Label lbliniFile 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�ð�װ�����ļ�"
         Height          =   180
         Left            =   765
         TabIndex        =   17
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblNote 
         Caption         =   "    Ӧ��ϵͳ����ֲ�����������ļ�����֮��صķ����������ű��ļ�������ȷָ����װ�����ļ���"
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   720
         Width           =   5250
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "��һ�� ָ����װ�����ļ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   225
         Width           =   2595
      End
   End
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   1
      Left            =   1305
      TabIndex        =   9
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraOwner 
         Caption         =   "ֲ��ϵͳ������"
         Height          =   2010
         Left            =   600
         TabIndex        =   19
         Top             =   1320
         Width           =   4530
         Begin VB.ComboBox cboOwnerUsr 
            Height          =   300
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   360
            Width           =   2160
         End
         Begin VB.CheckBox chkDBA 
            Caption         =   "����DBA��ɫ"
            Height          =   255
            Left            =   405
            TabIndex        =   30
            Top             =   1545
            Width           =   1320
         End
         Begin VB.TextBox txtOwnerPwd 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   825
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   750
            Width           =   2160
         End
         Begin VB.Label lblDBA 
            AutoSize        =   -1  'True
            Caption         =   "���Ը��ݹ���ϰ�߾����Ƿ�����DBA��ɫ��"
            Height          =   180
            Left            =   405
            TabIndex        =   29
            Top             =   1290
            Width           =   3330
         End
         Begin VB.Label lblNewUser 
            AutoSize        =   -1  'True
            Caption         =   "�û�"
            Height          =   180
            Left            =   405
            TabIndex        =   22
            Top             =   420
            Width           =   360
         End
         Begin VB.Label lblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   405
            TabIndex        =   21
            Top             =   810
            Width           =   360
         End
      End
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "�ڶ��� ָ��ֲ��ϵͳ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Top             =   225
         Width           =   2820
      End
      Begin VB.Label lblNote 
         Caption         =   "    ֲ��ϵͳ��Ȼ��һ�����ݿ��û���Ϊ�����ߣ�ͬʱ�����֪��ֲ��ϵͳ�����ߵ����룬�Ա���ֲ��ϵͳ����ȷ�ԡ�"
         Height          =   585
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   720
         Width           =   5250
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   1545
      TabIndex        =   25
      Top             =   3645
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbState 
      Height          =   150
      Left            =   3180
      TabIndex        =   24
      Top             =   4185
      Visible         =   0   'False
      Width           =   3210
      _ExtentX        =   5652
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4044
      Left            =   0
      ScaleHeight     =   3996
      ScaleWidth      =   1284
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.Image imgSetup 
         Height          =   3315
         Left            =   60
         Picture         =   "frmAppReplant.frx":058A
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4695
      TabIndex        =   3
      Top             =   3645
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   3645
      Width           =   1100
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5790
      TabIndex        =   0
      Top             =   3645
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   23
      Top             =   4044
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmAppReplant.frx":5B70
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8975
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1185
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "11:37"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraSetup 
      Height          =   3675
      Index           =   2
      Left            =   1305
      TabIndex        =   13
      Top             =   -120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.Frame fraStep 
         Height          =   120
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   465
         Width           =   5800
      End
      Begin VB.Label lblRegAudit 
         AutoSize        =   -1  'True
         Caption         =   "    ���ڻ����߱���ϵͳӦ����Ȩ����Ȼ���Լ���װ�أ����޷�����ʹ�á�"
         Height          =   360
         Left            =   225
         TabIndex        =   33
         Top             =   1335
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNextDo 
         AutoSize        =   -1  'True
         Caption         =   "    ���""���""��ʼ�Զ�װ��ϵͳ������""ȡ��""��ֹϵͳװ�أ���""��һ��""���µ���Ӧ��ϵͳװ�����á�"
         Height          =   360
         Left            =   225
         TabIndex        =   32
         Top             =   2025
         Width           =   5580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         Caption         =   "������ ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   165
         TabIndex        =   16
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "    �Ѿ�����˶Ը�ϵͳ��ֲ��ȫ�����á�"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   15
         Top             =   720
         Width           =   3420
      End
   End
End
Attribute VB_Name = "frmAppReplant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strIniPath      As String                 '��װ�����ļ�Ŀ¼
Dim intDefSysCode   As String                 'ϵͳ���
Dim strDefSysName   As String                 'ϵͳ����
Dim strDefVersion   As String                 '�汾��
Dim strDefSpace   As String                   '��ռ䶨�崮
Dim strDefUser      As String                 '�µ�ȱʡ�û���
Dim strDefData      As String                 '�û���ѡ������

Dim mstrExtSysCode  As String                  'Ҫ������չ����ϵͳ�ı��
Dim mstrExtVersion  As String                  'Ҫ������չ����ϵͳ�İ汾

Dim strTbsPath As String                        'ȱʡ��ռ�·�����ƣ�������ʷ��ռ����

Dim objText As TextStream

Dim mbln���� As Boolean    '���ΰ�װ�Ƿ����������װ�װ
Dim mlng���� As Long       '���׺�
Dim mlst��׼ As ListItem   '�����Ҫ��װ�����ף������ṩ��׼�������ݵ�ϵͳ

Dim intStep As Integer

Dim mcnOwner As New ADODB.Connection
Dim mlngEnjoy As Long
Dim strSQL As String, strTemp As String
Dim intCount As Integer, intItems As Integer
        
Dim aryRow() As String
Dim aryVal() As String

Private Sub cboOwnerUsr_Click()
    Dim rsTemp As New ADODB.Recordset
    
On Error GoTo errHandle
    txtOwnerPwd.Text = ""
    
    If mstrExtSysCode = "" Then
        '����չϵͳ
        With rsTemp
            If .State = adStateOpen Then .Close
            strSQL = "select ���,����" & _
                    " from zlSystems" & _
                    " where ������='" & cboOwnerUsr.Text & "'" & _
                    " start with ����� is null" & _
                    " connect by prior ���=�����" & _
                    " order by level"
            .Open strSQL, gcnOracle, adOpenKeyset
            mlngEnjoy = 0
            If .EOF Or .BOF Then Exit Sub
            .MoveLast
            mlngEnjoy = .Fields(0).value
            strSQL = "����������" & .Fields(1).value & "�������ߣ�" & _
                vbCr & "ѡ����û���ʾ����ϵͳ�ǹ���װ��"
            MsgBox strSQL, vbExclamation, gstrSysName
        End With
    Else
        '��չϵͳ��ֻ��ʹ�ù���ʽ
        mlngEnjoy = cboOwnerUsr.ItemData(cboOwnerUsr.ListIndex)
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("��װδ��ɣ����ȡ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdSetupFile_Click()
    With frmMDIMain.dlgMain
        .FileName = ""
        .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
        .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblSetupFile.Caption = .FileName
        End If
    End With
    If ChkSetupFile(True) = False Then
        lblSetupFile.Caption = ""
        cmdSetupFile.SetFocus
    End If

End Sub

Private Sub cmdNext_Click()
    Dim strErr As String
    
    If fraSetup(0).Visible Then
        '------------------------------------------------------------
        '��һ����
        '------------------------------------------------------------
        If Trim(lblSetupFile.Caption) = "" Then
            MsgBox "δ��ȷѡ���������װ�����ļ������ܼ���", vbExclamation, gstrSysName
            cmdSetupFile.SetFocus
            Exit Sub
        End If
        
        '------------------------------
        fraSetup(0).Visible = False
        fraSetup(1).Visible = True
        cmdPrevious.Enabled = True
    
    ElseIf fraSetup(1).Visible Then
        '------------------------------------------------------------
        '�ڶ�����
        '------------------------------------------------------------
        Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(cboOwnerUsr.Text), Trim(txtOwnerPwd.Text), False, MSODBC, "", False)
        If mcnOwner.State = adStateClosed Then
            Set mcnOwner = gobjRegister.GetConnection(gstrServer, Trim(cboOwnerUsr.Text), Trim(txtOwnerPwd.Text), True, MSODBC, "", False)
            If mcnOwner.State = adStateClosed Then
                MsgBox "������������󣬲��ܼ���" & vbNewLine & strErr, vbExclamation, gstrSysName
                txtOwnerPwd.SetFocus
                Exit Sub
            End If
        End If
        Call SetSQLTrace(gstrServer, Trim(cboOwnerUsr.Text), mcnOwner)
        
        On Error Resume Next
        Dim strErrInfo As String
        MousePointer = 11
        cmdSetupFile.Enabled = False
        cmdCancel.Enabled = False
        cmdNext.Enabled = False
        stbThis.Panels(2).Text = "���ֲ��ϵͳ"
        strErrInfo = CheckTable(strIniPath & "zlTable.sql")
        MousePointer = 0
        cmdSetupFile.Enabled = True
        cmdCancel.Enabled = True
        cmdNext.Enabled = True
        stbThis.Panels(2).Text = ""
                
        If strErrInfo <> "" Then
            If InStr(strErrInfo, "�Ƿ������") > 0 Then
                If MsgBox(strErrInfo, vbQuestion Or vbYesNo Or vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox strErrInfo, vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        '------------------------------
        fraSetup(1).Visible = False
        fraSetup(2).Visible = True
        cmdNext.Caption = "���(&F)"
    ElseIf fraSetup(2).Visible Then
        '------------------------------------------------------------
        '��������
        '------------------------------------------------------------
        If mlngEnjoy = 0 Then
            Set gcnTools = GetConnection("ZLTOOLS")
            If gcnTools Is Nothing Then Exit Sub
        End If
        
        strSQL = "    �Ѿ���������е���ֲ���ã�ϵͳ�������Զ���ֲ���̡�" & vbCr & vbCr _
                & "    ��ֲ���̿������нϳ�ʱ�䣬�벻Ҫ����ǿ���жϣ�����" & vbCr _
                & "�����ܲ�������������Ӱ��ϵͳ���С�" & vbCr & vbCr _
                & "   ������ֲ��"
        If MsgBox(strSQL, vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        
        cmdCancel.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        fraSetup(2).Enabled = False
        
        If SysInstall() Then
            MsgBox "��ֲ�ɹ������������Ӧ�ó�����ֲ������ʹ�ø�ϵͳ��", vbInformation, gstrSysName
            
            
            
        Else
            gcnOracle.Execute "delete zlSystems where ���=" & intDefSysCode * 100 + mlng����
            MsgBox "��ֲʧ�ܣ����鰲װ�ļ�����ȷ�ԡ�", vbInformation, gstrSysName
        End If
        cmdNext.Enabled = True
        Unload Me
    End If

End Sub

Private Sub cmdPrevious_Click()
    If fraSetup(2).Visible Then
        fraSetup(2).Visible = False
        fraSetup(1).Visible = True
        cmdNext.Caption = "��һ��(&N)"
    ElseIf fraSetup(1).Visible Then
        fraSetup(1).Visible = False
        fraSetup(0).Visible = True
        cmdPrevious.Enabled = False
    End If

End Sub

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    Dim objItem As ListItem
    With imgSetup
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    pgbState.Left = stbThis.Panels(2).Left + TextWidth("���ڴ������ݱ�")
    pgbState.Width = stbThis.Panels(3).Left - pgbState.Left - 100
    pgbState.Top = stbThis.Top + stbThis.Height / 3
    
    '������ֵ�ǰĿ¼���ڰ�װ��ֲ�ļ�����ֱ����д
    If Dir(App.Path & "\zlSetup.ini") <> "" Then
        lblSetupFile.Caption = App.Path & "\zlSetup.ini"
        If ChkSetupFile() = False Then
            lblSetupFile.Caption = ""
        End If
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdNext.Enabled = False Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Function ChkSetupFile(Optional blnMsg As Boolean) As Boolean
    
    '-------------------------------------
    '�����Ͱ�װ�����ļ�����ȷ��
    '-------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim varVersion As Variant, varExtVersin As Variant
    Dim i As Long
    
    strIniPath = Mid(lblSetupFile.Caption, 1, Len(lblSetupFile.Caption) - 11)
    '����ļ�ƥ���Լ��
    strTemp = ""
    If Dir(strIniPath & "zlSequence.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & strIniPath & "zlSequence.sql"
    If Dir(strIniPath & "zlTable.sql") = "" Then strTemp = strTemp & vbCr & "���ݱ��ļ�" & strIniPath & "zlTable.sql"
    If Dir(strIniPath & "zlConstraint.sql") = "" Then strTemp = strTemp & vbCr & "Լ���ļ�" & strIniPath & "zlConstraint.sql"
    If Dir(strIniPath & "zlIndex.sql") = "" Then strTemp = strTemp & vbCr & "�����ļ�" & strIniPath & "zlIndex.sql"
    If Dir(strIniPath & "zlView.sql") = "" Then strTemp = strTemp & vbCr & "��ͼ�ļ�" & strIniPath & "zlView.sql"
    If Dir(strIniPath & "zlProgram.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & strIniPath & "zlProgram.sql"
    If Dir(strIniPath & "zlManData.sql") = "" Then strTemp = strTemp & vbCr & "���������ļ�" & strIniPath & "zlManData.sql"
    If Dir(strIniPath & "zlAppData.sql") = "" Then strTemp = strTemp & vbCr & "Ӧ�������ļ�" & strIniPath & "zlAppData.sql"
    If strTemp <> "" Then
        If blnMsg Then MsgBox "���·�������װ������ļ���ʧ�����ܼ�����������" & strTemp, vbExclamation, gstrSysName
        Exit Function
    End If
    
    '��װ�����ļ�����
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(lblSetupFile.Caption)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        intDefSysCode = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        strDefSysName = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        strDefVersion = Trim(Mid(strTemp, 6))
    
        '�ж��Ƿ�Ӧ�ðѱ��ΰ�װ��Ϊ���װ�װ
        Dim lngTemp As Long
        Dim lngMax As Long        '�������׺�
        Dim blnHase  As Boolean   '�Ƿ���ͬϵͳ����
        Dim lstTemp As ListItem

        
        mbln���� = False
        mlng���� = 0
        For Each lstTemp In frmAppStart.lvwSys.ListItems
            lngTemp = Mid(lstTemp.Key, 2)
            If lngTemp \ 100 = intDefSysCode Then
                'ϵͳ��ͬ
                blnHase = True
                If lngMax < lngTemp Mod 100 Then
                    lngMax = lngTemp Mod 100 '�����������׺�
                End If
                
                If strDefVersion = lstTemp.SubItems(1) Then
                    '�汾Ҳ��ͬ���ǾͿ�����
                    mbln���� = True
                    Set mlst��׼ = lstTemp
                End If
            End If
        Next
        If blnHase = True Then
            '��ͬϵͳ�İ�װ
            If mbln���� = False Then
                If blnMsg Then MsgBox "��ǰ���ݿ���Ҳ����ͬ���͵�ϵͳ���ڣ������ڰ汾������������ֲ��", vbInformation, gstrSysName
                Exit Function
            Else
                If blnMsg = False Then
                    Exit Function
                Else
                    If lngMax >= 99 Then
                        MsgBox "��ǰ���ݿ���Ҳ����ͬ���͵�ϵͳ���ڣ��������㹻�࣬������ֲ��", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If MsgBox("��ǰ���ݿ�������" & strDefSysName & "ϵͳ���ڣ����Ƿ�Ҫ����ֲһ���µģ�", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                    mlng���� = lngMax + 1
                End If
            End If
        End If
    Else
        err.Raise 10
    End If
    Caption = "Ӧ��ϵͳ��װ" & " - " & strDefSysName & " V" & strDefVersion
    lblSysName.Caption = "ϵͳ����" & strDefSysName
    lblVersion.Caption = "�汾�ţ�" & strDefVersion
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[��ռ�]" Then
        strDefSpace = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�û���]" Then
        strDefUser = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[������]" Then
        strDefData = Trim(Mid(strTemp, 6))
    Else
        err.Raise 10
    End If
    
    mstrExtSysCode = ""
    mstrExtVersion = ""
    If Not objText.AtEndOfStream Then
        '������չϵͳ������
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 5) = "[��ϵͳ]" Then
            mstrExtSysCode = Trim(Mid(strTemp, 6))
            
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 5) = "[���汾]" Then
                mstrExtVersion = Trim(Mid(strTemp, 6))
            Else
                mstrExtSysCode = ""
            End If
        End If
    End If
    
    If err <> 0 Then
        If blnMsg Then MsgBox "��װ�����ļ���ʧ����ȷ", vbExclamation, gstrSysName
        Exit Function
    End If
    objText.Close
    
    '�����û��嵥
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    cboOwnerUsr.Clear
    If mstrExtSysCode = "" Then
        '����չϵͳ
        With rsTemp
            gstrSQL = "select username " & _
                    " from dba_users U" & _
                    " where username not in ('SYS','SYSTEM','ZLTOOLS')" & _
                    "       and exists (select 1 from dba_Tables T where T.owner=U.username)" & _
                    "       and username not in (select ������ from zlsystems where FLOOR(���/100)=" & intDefSysCode & ")"
            .Open gstrSQL, gcnOracle, adOpenKeyset
            Do While Not .EOF
                cboOwnerUsr.AddItem .Fields(0).value
                .MoveNext
            Loop
            If cboOwnerUsr.ListCount > 1 Then cboOwnerUsr.ListIndex = 0
        End With
    Else
        '����չϵͳ���Ǳ���Ҫ�����������ж�
        '1)ϵͳ�����
        '2)û����������ͬϵͳ��չ
        '3)�汾���ܵ���Ҫ��
        gstrSQL = "select A.������,A.�汾��,A.��� from zlsystems A " & _
                  "  Where floor(A.��� / 100) = " & mstrExtSysCode & _
                  "        and not exists (select B.��� from zlsystems B where B.�����=A.��� and floor(B.���/100)=" & intDefSysCode & ")"
        
        If Not rsTemp Is Nothing Then
            If rsTemp.State = 1 Then rsTemp.Close
        End If
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        varExtVersin = Split(mstrExtVersion, ".")
        Do Until rsTemp.EOF
            '�жϰ汾
            varVersion = Split(rsTemp("�汾��"), ".")
            For i = LBound(varExtVersin) To UBound(varExtVersin)
                If varExtVersin(i) > varVersion(i) Then
                    '�����ݿ��ж����İ汾����
                    Exit For
                End If
            Next
            If i > UBound(varExtVersin) Then
                '��������
                cboOwnerUsr.AddItem rsTemp("������")
                cboOwnerUsr.ItemData(cboOwnerUsr.NewIndex) = rsTemp("���")
            End If
            rsTemp.MoveNext
        Loop
        
    End If
    
    For intCount = 0 To cboOwnerUsr.ListCount - 1
        If cboOwnerUsr.List(intCount) = UCase(strDefUser) Then
            cboOwnerUsr.ListIndex = intCount
            Exit For
        End If
    Next
    If cboOwnerUsr.ListCount = 0 Then
        If blnMsg Then MsgBox "û�к��ʵĿ���ֲ���û���", vbInformation, gstrSysName
        Exit Function
    End If
    
    If cboOwnerUsr.ListIndex < 0 Then cboOwnerUsr.ListIndex = 0
    
    
    '˳���ע���ļ�Ҳһ�������
    Call ChkRegFile
    
    ChkSetupFile = True
End Function

Private Sub ChkRegFile()
    '�ж�ϵͳ��Ȩ
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo errHand
    gstrSQL = "Select Count(*) From Zlregfunc f, Zlreginfo r, zlRegCheck t Where r.��Ŀ = '��Ȩ֤��' And f.ϵͳ = " & intDefSysCode
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.Fields(0).value > 0 Then
        Me.lblRegAudit.Caption = "    �Ѿ��߱���ϵͳӦ����Ȩ��������װ�غ�������Ȩʹ�á�"
        Exit Sub
    End If
errHand:
    Me.lblRegAudit.Caption = "    ���ڻ����߱���ϵͳӦ����Ȩ����Ȼ���Լ���װ�أ����޷�������Ȩʹ�ã�"
End Sub

Private Function CheckTable(FileName As String) As String
    '--------------------------------------------
    '���ܣ�������ݱ�ͬʱ�ж����ݱ�����Ƿ���ȷ
    '--------------------------------------------
    Dim arySql() As String, strObjName As String, strTables As String
    Dim rsTemp As New ADODB.Recordset
    
    CheckTable = ""
    pgbState.value = 0
    pgbState.Visible = True
    With rsTemp
        .Filter = 0
        If gblnDBA Then
            strSQL = "select TABLE_NAME from DBA_TABLES where OWNER='" & cboOwnerUsr.Text & "'"
        Else
            strSQL = "select TABLE_NAME from USER_TABLES"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        
        err = 0
        On Error Resume Next
        Set objText = gobjFile.OpenTextFile(FileName)
        If err.Number <> 0 Then
            CheckTable = "���ڰ�װ�ű�����ȷ���޷���ֲ֤��ϵͳ����ȷ�ԡ����ܼ�����"
            .Filter = 0
            Exit Function
        End If
        intCount = GetFileLineCount(objText)
        objText.Close
        Set objText = gobjFile.OpenTextFile(FileName)
        
        strTables = ""
        On Error GoTo 0
        strSQL = ""
        Do While Not objText.AtEndOfStream
            strTemp = Trim(objText.ReadLine)
            If Left(strTemp, 2) <> "--" Then
                If Right(strTemp, 1) = ";" Then
                    strSQL = strSQL & vbCrLf & Left(strTemp, Len(strTemp) - 1)
                Else
                    strSQL = strSQL & vbCrLf & strTemp
                End If
                If Left(strSQL, 2) = vbCrLf Then
                    If Len(strSQL) = 2 Then
                        strSQL = ""
                    Else
                        strSQL = Mid(strSQL, 3)
                    End If
                End If
            End If
            If (Right(strTemp, 1) = ";" Or objText.AtEndOfStream) And Len(strSQL) <> 0 Then
                strSQL = UCase(Replace(Replace(Trim(strSQL), vbTab, " "), vbCrLf, " "))
                arySql = Split(strSQL, " TABLE ")
                If InStr(1, arySql(1), " ") > 0 And InStr(1, arySql(1), " ") < InStr(1, arySql(1), "(") Then
                    strObjName = Trim(Left(arySql(1), InStr(1, arySql(1), " ")))
                Else
                    strObjName = Trim(Left(arySql(1), InStr(1, arySql(1), "(") - 1))
                End If
                .Filter = "TABLE_NAME='" & strObjName & "'"
                If .EOF Then
                    strTables = strTables & vbCr & "    " & strObjName
                    If UBound(Split(strTables, vbCr)) > 16 Then Exit Do
                End If
                strSQL = ""
            End If
            
            pgbState.value = objText.Line / intCount * 100
        Loop
        .Filter = 0
    End With
    pgbState.value = 0
    pgbState.Visible = False
    If strTables <> "" Then
        CheckTable = "    ���ڸ��û������и�ϵͳҪ����������ݱ�" & _
              vbCr & "�����ж�����ȷ��ϵͳ�����ߣ��Ƿ������" & _
              vbCr & "    ȱ�ٵ����ݱ������" & strTables
    End If
End Function

Private Function SysInstall() As Boolean
    '----------------------------------
    '���ܣ����ϵͳ�İ�װ����
    '---------��װ�㷨-----------------
    '    ������ϵͳ���ݱ�ռ�
    '    If not �����Ѿ���װ��ϵͳ Then
    '        ������ϵͳ������
    '        �ɹ��������������Ҫ�Ĺ������ݶ���Ȩ��
    '    End If
    '    ������ϵͳ���ݶ���
    '    �������ݼ���ѡ���ݰ�װ
    '----------------------------------
    Dim strTmpSpace As String
    Dim rsTemp As New ADODB.Recordset, cnCtxsys As New ADODB.Connection
    
    
    err = 0
    On Error GoTo errHand
    gcnOracle.Execute "Grant Select on sys.v_$session to Public"
    gcnOracle.Execute "Grant Select on sys.v_$parameter to Public"
        
    With rsTemp
        If .State = adStateOpen Then .Close
        strSQL = "SELECT TEMPORARY_TABLESPACE FROM DBA_USERS WHERE USERNAME='ZLTOOLS'"
        .Open strSQL, gcnOracle, adOpenKeyset
        If .EOF Or .BOF Then SysInstall = False: Exit Function
        strTmpSpace = .Fields(0).value
    End With
    
    'SYS����ϵͳ��Ȩ
    gstrSQL = "Grant Connect,Resource," & IIf(chkDBA.value = 1, "DBA,", "") & _
            " Create Table,UNLIMITED TABLESPACE,Create Role,Create User,Drop User,Create Public Synonym,Drop Public Synonym" & _
            " to " & cboOwnerUsr.Text & " With Admin Option"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Select on sys.dba_role_privs to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Select on sys.dba_roles to " & cboOwnerUsr.Text
    gcnOracle.Execute gstrSQL
    gstrSQL = "Grant Execute on sys.dbms_sql to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL
    ' 2007-8-07 ������������Ȩ
    gstrSQL = "Grant Select on sys.gv_$session to " & cboOwnerUsr.Text & " With Grant Option"
    gcnOracle.Execute gstrSQL

    
    On Error Resume Next '����ȫ�ļ����Ĳ������п���û�и��û������԰Ѵ�������
    gstrSQL = "Grant CTXAPP to " & cboOwnerUsr.Text & " With Admin Option"
    gcnOracle.Execute gstrSQL
    gcnOracle.Execute "alter user ctxsys identified by ctxsys"
    cnCtxsys.Open "Driver={Microsoft ODBC for Oracle};Server=" & gstrServer, "ctxsys", "ctxsys"
    cnCtxsys.Execute "Grant Execute on ctx_ddl to " & cboOwnerUsr.Text & " With Grant Option" 'Ϊ���ڹ�����ִ�а�����
    
    On Error GoTo errHand
    '����������Ѿ���װ��ϵͳ
    If mlngEnjoy = 0 Then
        '������ϵͳ������
        stbThis.Panels(2).Text = "��ֲ��������" & cboOwnerUsr.Text & "��Ȩ"
        '�ɹ��������������Ҫ�Ĺ������ݶ���Ȩ��
        With rsTemp
            If .State = adStateOpen Then .Close
            strSQL = "select OBJECT_NAME,OBJECT_TYPE from user_objects where OBJECT_TYPE in('FUNCTION','PROCEDURE','SEQUENCE','TABLE','VIEW')  And Instr(OBJECT_NAME,'BIN$')<=0"
            .Open strSQL, gcnTools, adOpenKeyset
            Do While Not .EOF
                pgbState.value = .AbsolutePosition / .RecordCount * 100
                Select Case !Object_Type
                Case "FUNCTION", "PROCEDURE"
                    gcnTools.Execute "grant execute on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "VIEW"
                    gcnTools.Execute "grant select on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "SEQUENCE"
                    gcnTools.Execute "grant select,alter on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                Case "TABLE"
                    gcnTools.Execute "grant select,insert,update,delete on " & !Object_Name & " to " & cboOwnerUsr.Text & " With GRANT Option"
                End Select
                DoEvents
                .MoveNext
            Loop
        End With
        pgbState.value = 0
        pgbState.Visible = False
    End If
    
    '��д��װϵͳ�嵥
    strSQL = "insert into zlSystems(���,�����,����,������,��װ����,������װ,�汾��)" & _
            " values(" & intDefSysCode * 100 + mlng����
    If mlngEnjoy <> 0 Then
        strSQL = strSQL & "," & mlngEnjoy
    Else
        strSQL = strSQL & ",null"
    End If
    strSQL = strSQL & ",'" & strDefSysName & "'"
    strSQL = strSQL & ",'" & Trim(cboOwnerUsr.Text) & "'"
    strSQL = strSQL & ",sysdate,0,'" & strDefVersion & "')"
    gcnOracle.Execute strSQL
    
    '������ϵͳ���ݶ���(�ڶ����Ѵ�������mcnOwner)
    
    '��������
    stbThis.Panels(2).Text = "�������ݰ�װ"
    If mbln���� = False Then
        Call RunSetupFile(mcnOwner, strIniPath & "zlManData.sql", ";", True)
    Else
        'ͨ�����ݿ��п����õ�
        If CopyManageData(mcnOwner) = False Then GoTo errHand
    End If
    
    '��װ����
    stbThis.Panels(2).Text = "�̶�����װ"
    
    If mbln���� = False Then
        If RunSetupFile(mcnOwner, strIniPath & "zlReport.sql", ";", False) = 3 Then GoTo errHand
    Else
        'ͨ�����ݿ��п����õ�
        If CopyReport(mcnOwner, Mid(mlst��׼.Key, 2), intDefSysCode * 100 + mlng����) = False Then GoTo errHand
    End If
    
    '������װ���µ�������ʵ����ֵ��ƥ��
    stbThis.Panels(2).Text = "���м��"
    DoEvents
    Call ChkSequence
    
    '��д��װ��¼Ϊ������װ
    strSQL = "update zlSystems set ������װ=1 where ���=" & intDefSysCode * 100 + mlng����
    gcnOracle.Execute strSQL
    strSQL = "insert into zlSysFiles(ϵͳ,����,�ļ���,����,������)" & _
            " values (" & intDefSysCode * 100 + mlng���� & ",1,'" & lblSetupFile.Caption & "',sysdate,user)"
    gcnOracle.Execute strSQL
     
    '���˺������ʷ���ݿռ���ж�
    gstrSQL = "Select ���� from zltools.zlBakTables where rownum<=1 and  ϵͳ=" & intDefSysCode * 100 + mlng����
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If Not rsTemp.EOF Then
        '��Ҫ����Ƿ����H��
        gstrSQL = "Select 1 From User_Tables where table_name='H" & Nvl(rsTemp!����) & "'"
        OpenRecordset rsTemp, gstrSQL, Me.Caption, , , mcnOwner
        If Not rsTemp.EOF Then
            MsgBox "��ϵͳ�����ݽṹ̫�ɣ����ֹ��������ٽ�����ֲ!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        
        Dim strMsg As String, lngCount As Long
        lngCount = 0
ResumeGo:
        lngCount = lngCount + 1
        strMsg = "����ֲ��ϵͳ���������ʷ���ݿռ�,�Ƿ񴴽���ʷ���ݿռ䣿" & vbCrLf & _
             "ѡ���ǡ����´���һ����ʷ���ݿռ䡣" & vbCrLf & _
             "ѡ�񡾷񡿣���ֲһ���Ѿ����ڵ���ʷ���ݿռ䡣"
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
        'ע��*****,���ڵ���gobjRegister.GetPassword������Ҫ�󱾴���ı���̶�Ϊ"Ӧ��ϵͳ��ֲ",�����ܷ���ת���������
            '����һ����ʷ���ݿռ�
            If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, cboOwnerUsr.Text, gobjRegister.GetPassword, intDefSysCode * 100 + mlng����, 0, 0) = False Then
                '��ѡ������
                If lngCount > 3 Then
                        MsgBox "��ϵͳ����ֲʧ��!", vbInformation + vbDefaultButton1, gstrSysName
                        GoTo errHand:
                Else
                    GoTo ResumeGo:
                End If
            End If
        Else
            '����һ����ʷ���ݿռ�
            If frmHistorySpaceSet.ShowInstall(Me, mcnOwner, cboOwnerUsr.Text, gobjRegister.GetPassword, intDefSysCode * 100 + mlng����, 2, 0) = False Then
                '��ѡ������
                If lngCount > 3 Then
                        MsgBox "��ϵͳ����ֲʧ��!", vbInformation + vbDefaultButton1, gstrSysName
                        GoTo errHand:
                Else
                    GoTo ResumeGo:
                End If
            End If
        End If
    End If
    
   
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    
    SysInstall = True
    Exit Function

errHand:
    If mcnOwner.State = adStateOpen Then mcnOwner.Close
    pgbState.Visible = False
    SysInstall = False
    MsgBox err.Description, vbExclamation, "����"
End Function

Private Function CopyManageData(ByVal cnExecuter As ADODB.Connection) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Dim lngNewSystem As Long
    Dim lngOldSystem As Long
    Dim strOldOwner As String
    
    pgbState.value = 0
    pgbState.Visible = True
    
    lngNewSystem = intDefSysCode * 100 + mlng����
    lngOldSystem = Mid(mlst��׼.Key, 2)
    
    strOldOwner = GetOwnerName(lngOldSystem, gcnOracle)
    
    On Error GoTo errHandle
    'zlComponent����
    gstrSQL = "insert into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ) " & _
                "select ����,����,���汾,�ΰ汾,���汾," & lngNewSystem & " from zlComponent where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 5
    
    'zlPrograms����
    gstrSQL = "insert into zlPrograms(���,����,˵��,����,ϵͳ) " & _
                "select ���,����,˵��,����," & lngNewSystem & " from zlPrograms where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 15
    
    'zlProgFuncs����
    gstrSQL = "insert into zlProgFuncs(���,����,ϵͳ) " & _
                "select ���,����," & lngNewSystem & " from zlProgFuncs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 35
    
    'zlProgPrivs����
    gstrSQL = "insert into zlProgPrivs(���,����,������,����,Ȩ��,ϵͳ) " & _
                "select ���,����,decode(������,'" & strOldOwner & "',user,������),����,Ȩ��," & lngNewSystem & " from zlProgPrivs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 65
    
    'zlMenus����
    '������Ч�˵�
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            gstrSQL = "select 1 from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            .Open gstrSQL, cnExecuter
            If .EOF Then Exit Do
            strSQL = "delete from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            cnExecuter.Execute gstrSQL
        Loop
    End With
    CopyMenu gcnOracle, lngOldSystem, lngNewSystem
    pgbState.value = 85
    
    'zlBaseCode����
    gstrSQL = "insert into zlBaseCode(����,�̶�,˵��,����,ϵͳ) " & _
                "select ����,�̶�,˵��,����," & lngNewSystem & " from zlBaseCode where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 90
    
    'zlDataMove����
    gstrSQL = "insert into zlDataMove(���,����,˵��,�����ֶ�,ת������,�ϴ�����,ϵͳ,״̬) " & _
                "select ���,����,˵��,�����ֶ�,ת������,�ϴ�����," & lngNewSystem & ",״̬ from zlDataMove where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 95
    
    'zlAutoJobs����
    gstrSQL = "insert into zlAutoJobs(����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��,ϵͳ) " & _
                "select ����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��," & lngNewSystem & " from zlAutoJobs where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 97
    
    'zlParameters����
    gstrSQL = "Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��) " & _
            " Select zlParameters_ID.Nextval," & lngNewSystem & ",ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ϵͳ=" & lngOldSystem
    cnExecuter.Execute gstrSQL
    pgbState.value = 99
    
    pgbState.value = 0
    pgbState.Visible = True
    CopyManageData = True
    Exit Function
errHandle:
    If MsgBox("�������д����Ƿ������" & vbCrLf & "    " & err.Description, vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
    pgbState.value = 0
    pgbState.Visible = True
    
End Function

Private Function RunSetupFile(cnThisDB As ADODB.Connection, FileName As String, Optional DeLimiter As String = ";", Optional ResumeNext As Boolean) As Byte
    '----------------------------------------------
    '���ܣ�ִ�а�װ�ű��ļ�
    '������
    '   cnThisDB:   ָ�������ݿ�����
    '   FileName:   �ű��ļ�
    '   Delimiter:  �ű��ļ����ָ�����
    '   ResumeNext: �Ƿ�������
    '���أ�1-ִ�гɹ���2-���ڴ��󵫼���ִ����ϣ�3-������ж�
    '----------------------------------------------
    Dim lngLines As Long
    err = 0
    On Error Resume Next
    Set objText = gobjFile.OpenTextFile(FileName)
    If err <> 0 Then
        MsgBox "�޷��򿪽ű��ļ�" & FileName & ",ִ���ж�" & vbNewLine & err.Description, vbExclamation, gstrSysName
        RunSetupFile = 3
        Exit Function
    End If
    
    lngLines = GetFileLineCount(objText)
    objText.Close
    Set objText = gobjFile.OpenTextFile(FileName)
        
    pgbState.value = 0
    pgbState.Visible = True
    
    RunSetupFile = 1
    err = 0
    On Error GoTo 0
    strSQL = ""
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        If Left(strTemp, 2) <> "--" Then
            If Right(strTemp, 1) = DeLimiter Then
                strSQL = strSQL & vbCrLf & Left(strTemp, Len(strTemp) - 1)
            Else
                strSQL = strSQL & vbCrLf & strTemp
            End If
            If Left(strSQL, 2) = vbCrLf Then
                If Len(strSQL) = 2 Then
                    strSQL = ""
                Else
                    strSQL = Mid(strSQL, 3)
                End If
            End If
        End If
        If (Right(strTemp, 1) = DeLimiter Or objText.AtEndOfStream) And Len(strSQL) <> 0 Then
            err = 0
            On Error Resume Next
            cnThisDB.Execute strSQL
            If err <> 0 Then
                If ResumeNext Then
                    RunSetupFile = 2
                Else
                    MsgBox "�����ļ�" & FileName & "�д������������ִ���жϣ�" & vbCr & strSQL, vbExclamation, gstrSysName
                    RunSetupFile = 3
                    Exit Function
                End If
            End If
            err = 0
            On Error GoTo 0
            strSQL = ""
        End If
        pgbState.value = objText.Line / lngLines * 100
        DoEvents
    Loop
    pgbState.value = 0
    pgbState.Visible = False

End Function

Private Sub ChkSequence()
    '----------------------------------------------
    '���ܣ��������еĵ�ǰ����
    '----------------------------------------------
    Dim rsLst As ADODB.Recordset
    
    pgbState.value = 0
    pgbState.Visible = True
    Set rsLst = GetSequence("", mcnOwner)
    With rsLst
        Do While Not .EOF
            DoEvents
            pgbState.value = .AbsolutePosition / .RecordCount * 100
            Call AdjustNameSequece(!Owner & "." & !Table_Name, mcnOwner, !Column_Name)
            .MoveNext
        Loop
        
        Call Adjust����ID(mcnOwner)
    End With
    pgbState.value = 0
    pgbState.Visible = False
End Sub
