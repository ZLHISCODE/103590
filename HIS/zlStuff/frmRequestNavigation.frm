VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequestNavigation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������Զ�������"
   ClientHeight    =   5070
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmRequestNavigation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7890
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   45
      TabIndex        =   29
      Top             =   4590
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Height          =   4485
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   -15
      Width           =   1485
      Begin VB.Image imgSetup 
         Height          =   4335
         Left            =   60
         Picture         =   "frmRequestNavigation.frx":1582
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ��(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3675
      TabIndex        =   1
      Top             =   4605
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   2
      Top             =   4605
      Width           =   1230
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   4605
      Width           =   1230
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":6B68
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStep 
      Height          =   4605
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   -120
      Width           =   6435
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   6255
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Frame Frame1 
         Caption         =   "���췽ʽ"
         Height          =   2865
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   1530
         Width           =   6105
         Begin VB.OptionButton optMode 
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ����쵥����"
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   30
            Top             =   2364
            Width           =   3516
         End
         Begin VB.OptionButton optMode 
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ����쵥"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   14
            Top             =   1827
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "�����ϵĴ�������"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   13
            Top             =   1293
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "�����ϵĴ�������"
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   12
            Top             =   759
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "����ָ��ʱ�䷶Χ�ڲ��ϵ�������"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   3045
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ����쵥���ܣ������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   4
            Left            =   396
            TabIndex        =   31
            Top             =   2556
            Width           =   4140
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����ָ��ʱ�䷶Χ�ڵ����쵥��δ�������������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   3
            Left            =   396
            TabIndex        =   18
            Top             =   2027
            Width           =   4680
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ��ǰ�ⷿ�Ĳ��ϴ�����ʼ�ձ��������ޱ�׼�������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   2
            Left            =   396
            TabIndex        =   17
            Top             =   1498
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ��ǰ�ⷿ�Ĳ��ϴ�����ʼ�ձ��������ޱ�׼�������������쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   396
            TabIndex        =   16
            Top             =   969
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ָ����ʱ�䷶Χ���Բ��ϵ�������Ϊ���ݣ��������ε����쵥"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   0
            Left            =   396
            TabIndex        =   15
            Top             =   440
            Width           =   5400
         End
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��һ���������������쵥�ķ�ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��׼�����ĸ��ⷿ������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1650
         TabIndex        =   7
         Top             =   900
         Width           =   2730
      End
      Begin VB.Label lbl�ⷿ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   360
      End
   End
   Begin VB.Frame fraStep 
      Height          =   4605
      Index           =   1
      Left            =   1470
      TabIndex        =   19
      Top             =   -120
      Width           =   6435
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   285
         Left            =   4170
         TabIndex        =   23
         Top             =   1350
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   52101123
         CurrentDate     =   38096
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   285
         Left            =   4170
         TabIndex        =   25
         Top             =   1980
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   52101123
         CurrentDate     =   38096
      End
      Begin MSComctlLib.TreeView tvw���� 
         Height          =   3465
         Left            =   90
         TabIndex        =   28
         Top             =   1050
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.Label lbl������������ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   27
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lbl����ѡ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lbl����ʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&E)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   24
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lbl��ʼʱ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��(&S)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   22
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ڶ�������ָ�����������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   240
         Width           =   4500
      End
   End
End
Attribute VB_Name = "frmRequestNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ģʽ
    ����������
    ��������
    ��������
    �������쵥δ����
    �������쵥����
End Enum
Private mblnOk As Boolean
Private mlngStockID As Long                 '����ⷿID
Private mbln��ȷ���� As Boolean             '����ʱ�Ƿ���ȷ����
Private mintCheck As Integer                '��������
Private mblnFirst  As Boolean
Private mintUnit As Integer
Private Const mlngModule = 1722

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private Const mconIntColҩ��  As Integer = 2
Private Const mconIntCol���   As Integer = 3
Private Const mconIntCol���    As Integer = 4
Private Const mconIntCol��������  As Integer = 5
Private Const mconIntCol���Ч��  As Integer = 6
Private Const mconIntCol��������  As Integer = 7
Private Const mconIntColָ������� As Integer = 8
Private Const mconIntColʵ�ʽ�� As Integer = 9
Private Const mconIntColʵ�ʲ�� As Integer = 10
Private Const mconIntCol����ϵ�� As Integer = 11
Private Const mconIntCol���� As Integer = 12
Private Const mconIntCol���� As Integer = 13
Private Const mconIntCol��׼�ĺ� As Integer = 14
Private Const mconIntCol��λ As Integer = 15
Private Const mconIntCol���� As Integer = 16
Private Const mconIntColЧ�� As Integer = 17
Private Const mconIntCol���ʧЧ�� As Integer = 18
Private Const mconIntCol��д���� As Integer = 21
Private Const mconIntColʵ������ As Integer = 22
Private Const mconIntCol�ɹ��� As Integer = 24
Private Const mconIntCol�ɹ���� As Integer = 25
Private Const mconIntCol�ۼ� As Integer = 26
Private Const mconIntCol�ۼ۽�� As Integer = 27
Private Const mconintCol��� As Integer = 28
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdNext_Click()
    
    If fraStep(0).Visible Then
        fraStep(1).Visible = True
        fraStep(0).Visible = False
        fraStep(1).ZOrder
        cmdPrevious.Enabled = True
        cmdNext.Caption = "���(&F)"
        
        Call ResizeStuff
    Else
        'ȷ����Ӧ����:
        Dim i As Long
        Dim str����ID As String
        str����ID = ""
        For i = 1 To tvw����.Nodes.count
            If tvw����.Nodes(i).Key <> "Root" And _
                tvw����.Nodes(i).Checked Then
                str����ID = str����ID & "," & Mid(tvw����.Nodes(i).Key, 2)
            End If
        Next
        
        If str����ID <> "" Then
            str����ID = Mid(str����ID, 2)
        End If
    
        If Not CheckData(str����ID) Then Exit Sub
        
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub cmdPrevious_Click()
    If fraStep(1).Visible Then
        fraStep(1).Visible = False
        fraStep(0).Visible = True
        fraStep(0).ZOrder
        cmdPrevious.Enabled = False
        cmdNext.Caption = "��һ��(&N)"
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    '----ȱʡѡ�����м���----
    If Not mblnFirst Then Exit Sub
    fraStep(0).ZOrder
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    
    On Error GoTo ErrHandle
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    
    mblnFirst = True
    
    '----��ȡ���Ŀⷿ----
    Set rsTemp = ReturnSQL(mlngStockID, "��ȡ��������Ŀⷿ", False, , 1722)
    
    If rsTemp.EOF Then
        MsgBox "û���κοⷿ�������죬����[������������]���������������ã�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    With cbo�ⷿ
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            rsTemp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    '----�ж��Ƿ���Ҫ��ȷ����----
    mbln��ȷ���� = IS��������
    
    
   gstrSQL = "" & _
        "   Select Level as ��,ID,�ϼ�ID,���� From ���Ʒ���Ŀ¼ where ����=7" & _
        "   Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        "   Order by ��"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "���ķ��಻������", vbInformation, gstrSysName
        Exit Sub
    End If

    Dim objNode As Node
    Set objNode = tvw����.Nodes.Add(, , "Root", "�������ķ���", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!�� = 1 Then
            Set objNode = tvw����.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        Else
            Set objNode = tvw����.Nodes.Add("_" & rsTemp!�ϼ�ID, 4, "_" & rsTemp!Id, rsTemp!����, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw����.Nodes("Root").Selected = True
    tvw����.Nodes("Root").Expanded = True
    '----����ȱʡ��ʱ�䷶Χ��һ���£�----
    Me.dtp��ʼʱ��.Value = Format(DateAdd("d", -7, sys.Currentdate()), "yyyy-MM-dd") & " 00:00:00"
    Me.dtp����ʱ��.Value = Format(sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub ResizeStuff()
    Dim blnEnable As Boolean
    '�ж��Ƿ������û�������������
    blnEnable = (optMode(�������쵥δ����) Or optMode(����������) Or optMode(�������쵥����))
    lbl������������.Visible = blnEnable
    lbl��ʼʱ��.Visible = blnEnable
    lbl����ʱ��.Visible = blnEnable
    dtp��ʼʱ��.Visible = blnEnable
    dtp����ʱ��.Visible = blnEnable
    
    If blnEnable Then
        tvw����.Width = lbl��ʼʱ��.Left - 200 - tvw����.Left
    Else
        tvw����.Width = fraStep(1).Width - 200 - tvw����.Left
    End If
End Sub

Public Function ShowNavigation(ByVal frmParent As Object, ByVal lngStockID As Long) As Boolean
    On Error Resume Next
    mlngStockID = lngStockID
    mblnOk = False
    Me.Show 1, frmParent
    ShowNavigation = mblnOk
End Function

Private Function CheckData(Optional str����id_IN As String = "") As Boolean
    Dim lngTargetID As Long             'Ŀ��ⷿ��ID
    Dim rsCheck As New ADODB.Recordset
    Dim str����IN As String
    
    '����Ƿ���ڷ��������ļ�¼��ʼ��ֻ�����������бȽϣ�����ֵ�ʱ���ٰ��Ƿ���ȷ��������������Σ�
    On Error GoTo ErrHand
    CheckData = False
    
    gstrSQL = ""
    lngTargetID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    
    If optMode(����������) Then
        '�����ȷ���Σ���ҩƷ�����û�м�¼���������ݣ�����ȡ����
     gstrSQL = "" & _
                 " Select sum(Nvl(A.ʵ������,0)) ��������,max(Nvl(B.��������,0)) ��������,max(Nvl(B.ʵ������,0)) ʵ������,max(Nvl(B.ʵ�ʽ��,0)) ʵ�ʽ��,max(Nvl(B.ʵ�ʲ��,0)) ʵ�ʲ��, " & _
                 "        D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ� �ۼ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "        D.��װ��λ,D.����ϵ��,F.���㵥λ �ۼ۵�λ " & _
                 " From ҩƷ�շ���¼ A,�������� D,�շ���ĿĿ¼ F,����ִ�п��� Z,������ĿĿ¼ N, " & _
                 "      (Select �շ�ϸĿID,�ּ� From �շѼ�Ŀ Where SysDate Between ִ������ And Nvl(��ֹ����,Sysdate)" & _
                 GetPriceClassString("") & ") P, " & _
                 "      (Select ҩƷid ����ID,sum(��������) as ��������,sum(ʵ������) ʵ������,sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ��" & _
                 "       From ҩƷ��� Where �ⷿID=[4] And ����=1 Group by ҩƷid) B " & IIf(str����id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.���� IN (20,21,24,25,26) And A.���ϵ��=-1 And A.������� Between [2]" & _
                 " And [3]" & _
                 " And A.ҩƷID=D.����ID and D.����ID=N.id  And Nvl(A.ҩƷid,0)=B.����ID(+) And A.�ⷿID=[1] " & _
                 " And A.ҩƷID=D.����ID  And D.����ID=P.�շ�ϸĿID" & IIf(str����id_IN = "", "", " And N.����id+0=Q.Column_Value") & _
                 " And D.����ID=F.ID And D.����ID=Z.������ĿID And Z.ִ�п���ID+0=[1]" & _
                 " Having Sum(Nvl(A.ʵ������,0))>0 " & _
                 " Group By D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "       D.��װ��λ,D.����ϵ��,F.���㵥λ "
                 
    ElseIf optMode(��������) Then
       gstrSQL = "Select Nvl(A.����,0)-Sum(Nvl(B.��������,0)) ��������,Sum(Nvl(K.��������,0)) ��������,Sum(Nvl(K.ʵ������,0)) ʵ������,Sum(Nvl(K.ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(K.ʵ�ʲ��,0)) ʵ�ʲ��,  " & _
                "         D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ� �ۼ�,F.���,F.����,D.���Ч��,D.ָ�������,  " & _
                "         D.��װ��λ,D.����ϵ��,F.���㵥λ �ۼ۵�λ  " & _
                "  From (Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� Where �ⷿID=[1] And Nvl(����,0)>0) A, " & _
                "       �������� D,�շ���ĿĿ¼ F,����ִ�п��� Z,������ĿĿ¼ N,  " & _
                "       (Select �շ�ϸĿID,�ּ� From �շѼ�Ŀ Where SysDate Between ִ������ And Nvl(��ֹ����,Sysdate)" & _
                GetPriceClassString("") & ") P,  " & _
                "       (Select ҩƷid ����ID,Sum(��������) as ��������,sum(ʵ������) ʵ������, sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ�� " & _
                "        From ҩƷ��� Where �ⷿID=[1] And ����=1 Group by ҩƷid) B,  " & _
                 "      (Select ҩƷid ����ID,Sum(��������) as ��������,sum(ʵ������) ʵ������, sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ�� " & _
                 "      From ҩƷ��� Where �ⷿID=[4] And ����=1 Group by ҩƷid ) K " & IIf(str����id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                "  Where A.����ID=D.����ID  and A.����id=b.����id(+) And A.����id=K.����ID(+) And D.����ID=P.�շ�ϸĿID" & IIf(str����id_IN = "", "", " And N.����id+0=Q.Column_Value") & _
                "  And D.����ID=F.ID and D.����ID=N.id  And D.����ID=Z.������ĿID And Z.ִ�п���ID+0=[1]" & _
                "  Having Nvl(A.����,0)-Sum(Nvl(B.��������,0))>0 " & _
                "  Group By Nvl(A.����,0),D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ�,F.���,F.����,D.���Ч��,D.ָ�������,  " & _
                "        D.��װ��λ,D.����ϵ��,F.���㵥λ "
    ElseIf optMode(��������) Then
       gstrSQL = "Select Nvl(A.����,0)-Sum(Nvl(B.��������,0)) ��������,Sum(Nvl(K.��������,0)) ��������,Sum(Nvl(K.ʵ������,0)) ʵ������,Sum(Nvl(K.ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(K.ʵ�ʲ��,0)) ʵ�ʲ��,  " & _
                "         D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ� �ۼ�,F.���,F.����,D.���Ч��,D.ָ�������,  " & _
                "         D.��װ��λ,D.����ϵ��,F.���㵥λ �ۼ۵�λ  " & _
                "  From (Select �ⷿid, ����id, ����, ����, �̵�����, �ⷿ��λ From ���ϴ����޶� Where �ⷿID=[1] And Nvl(����,0)>0) A, " & _
                "       �������� D,�շ���ĿĿ¼ F,����ִ�п��� Z,������ĿĿ¼ N,  " & _
                "       (Select �շ�ϸĿID,�ּ� From �շѼ�Ŀ Where SysDate Between ִ������ And Nvl(��ֹ����,Sysdate)" & _
                GetPriceClassString("") & ") P,  " & _
                "       (Select ҩƷid ����ID,Sum(��������) as ��������,sum(ʵ������) ʵ������, sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ�� " & _
                "        From ҩƷ��� Where �ⷿID=[1] And ����=1 Group by ҩƷid) B,  " & _
                "      (Select ҩƷid ����ID,Sum(��������) as ��������,sum(ʵ������) ʵ������, sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ�� " & _
                "      From ҩƷ��� Where �ⷿID=[4] And ����=1 Group by ҩƷid ) K " & IIf(str����id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                "  Where A.����ID=D.����ID and A.����id=b.����id(+)  And A.����ID=K.����ID(+)  And D.����ID=P.�շ�ϸĿID" & IIf(str����id_IN = "", "", " And N.����id+0=Q.Column_Value") & _
                "  And D.����ID=F.ID and D.����ID=N.id  And D.����ID=Z.������ĿID And Z.ִ�п���ID+0=[1]" & _
                "  Having Nvl(A.����,0)-Sum(Nvl(B.��������,0))>0 " & _
                "  Group By Nvl(A.����,0),D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ�,F.���,F.����,D.���Ч��,D.ָ�������,  " & _
                "        D.��װ��λ,D.����ϵ��,F.���㵥λ "
                
    ElseIf optMode(�������쵥δ����) Then   '�������쵥δ��������������And Nvl(A.��ҩ��ʽ,0)=1 ������Ϊ���ʱ����ɾ�����쵥�������ƿⵥ����˵ģ���־�Ѿ�û���ˣ�
        gstrSQL = "select sum(A.��д����-A.ʵ������) ��������,max(Nvl(B.��������,0)) ��������,max(Nvl(B.ʵ������,0)) ʵ������,max(Nvl(B.ʵ�ʽ��,0)) ʵ�ʽ��,max(Nvl(B.ʵ�ʲ��,0)) ʵ�ʲ��, " & _
                 "        D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ� �ۼ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "        D.��װ��λ,D.����ϵ��,F.���㵥λ �ۼ۵�λ " & _
                 " from ҩƷ�շ���¼ A,�������� D,�շ���ĿĿ¼ F,����ִ�п��� Z,������ĿĿ¼ N, " & _
                 "      (Select �շ�ϸĿID,�ּ� From �շѼ�Ŀ Where SysDate Between ִ������ And Nvl(��ֹ����,Sysdate)" & _
                 GetPriceClassString("") & ") P, " & _
                 "      (Select ҩƷid ����ID,sum(��������) as ��������,sum(ʵ������) as ʵ������, sum(ʵ�ʽ��) as ʵ�ʽ��,sum(ʵ�ʲ��) as ʵ�ʲ��" & _
                 "      From ҩƷ��� Where �ⷿID=[4] And ����=1 Group by ҩƷid) B " & IIf(str����id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.����=19 And A.������� Between [2]" & _
                 " And [3]" & _
                 " And A.�ⷿID=[4] And A.�Է�����ID=[1]" & _
                 " And A.ҩƷID=D.����ID  and A.���ϵ��<>1 and D.����ID=N.id  And A.ҩƷid =b.����id(+)" & _
                 " And D.����ID=P.�շ�ϸĿID" & IIf(str����id_IN = "", "", " And N.����id+0=Q.Column_Value") & _
                 " And D.����ID=F.ID And D.����ID=Z.������ĿID And Z.ִ�п���ID+0=[1]" & _
                 " having sum(A.��д����-A.ʵ������)>0 " & _
                 " Group By D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "       D.��װ��λ,D.����ϵ��,F.���㵥λ "
    
    ElseIf optMode(�������쵥����) Then
        gstrSQL = "select sum(Nvl(A.ʵ������,0)) ��������,max(Nvl(B.��������,0)) ��������,max(Nvl(B.ʵ������,0)) ʵ������,max(Nvl(B.ʵ�ʽ��,0)) ʵ�ʽ��,max(Nvl(B.ʵ�ʲ��,0)) ʵ�ʲ��, " & _
                 "        D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ� �ۼ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "        D.��װ��λ,D.����ϵ��,F.���㵥λ �ۼ۵�λ " & _
                 " from ҩƷ�շ���¼ A,�������� D,�շ���ĿĿ¼ F,(Select Distinct ������Ŀid, ִ�п���id From ����ִ�п���) Z,������ĿĿ¼ N, " & _
                 "      (Select �շ�ϸĿID,�ּ� From �շѼ�Ŀ Where SysDate Between ִ������ And Nvl(��ֹ����,Sysdate) " & GetPriceClassString("") & ") P, " & _
                 "      (Select ҩƷid ����ID,Sum(��������) as ��������,sum(ʵ������) ʵ������, sum(ʵ�ʽ��) ʵ�ʽ��,sum(ʵ�ʲ��) ʵ�ʲ�� " & _
                 "       From ҩƷ��� Where �ⷿID=[4] And ����=1 Group by ҩƷid) B " & IIf(str����id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.����=19 And A.���ϵ��=1 And A.������� Between [2]" & _
                 " And [3]" & _
                 " And A.ҩƷID=D.����ID and D.����ID=N.id  And Nvl(A.ҩƷid,0)=B.����ID(+) And A.�ⷿID=[1] " & _
                 " And A.ҩƷID=D.����ID  And D.����ID=P.�շ�ϸĿID" & IIf(str����id_IN = "", "", " And N.����id+0=Q.Column_Value") & _
                 " And D.����ID=F.ID And D.����ID=Z.������ĿID And Z.ִ�п���ID+0=[1]" & _
                 " Having Sum(Nvl(A.ʵ������,0))>0 " & _
                 " Group By D.����ID,F.����,F.����,F.�Ƿ���,D.�ⷿ����,D.���÷���,P.�ּ�,F.���,F.����,D.���Ч��,D.ָ�������, " & _
                 "       D.��װ��λ,D.����ϵ��,F.���㵥λ "

    End If
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���ڷ��������ļ�¼", mlngStockID, CDate(Format(dtp��ʼʱ��.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp����ʱ��.Value, "yyyy-MM-dd HH:mm:ss")), lngTargetID, str����id_IN)
    
    If rsCheck.RecordCount = 0 Then
        MsgBox "û�ҵ����������ļ�¼��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call WriteResult(rsCheck)
    
    Dim intCount As Integer
    With frmRequestStuffCard
        For intCount = 0 To .cboStock.ListCount - 1
            If .cboStock.ItemData(intCount) = lngTargetID Then
                .cboStock.ListIndex = intCount: Exit For
            End If
        Next
    End With
    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub WriteResult(ByVal rsCheck As ADODB.Recordset)
    Dim strUnit As String
    Dim lngTargetID As Long
    Dim bln��ʾ As Boolean, bln�ⷿ As Boolean
    Dim bln���� As Boolean, bln��ҩ As Boolean       'bln����-����ϵͳ����������顱���û������������Ƿ�����޿������ģ�bln��ҩ-��ǰ�����Ƿ���ʱ�ۻ���������
    Dim dbl�������� As Double, dbl��д���� As Double
    Dim rsStock As New ADODB.Recordset  'ҩƷ���
    Dim rsTemp  As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    lngTargetID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
    Call GetPara(lngTargetID)
    bln�ⷿ = CheckStock(lngTargetID)
    Dim blnData As Boolean
    '׼���������ݣ�ȫ�������۵�λΪ׼��������SetColValue������ת���������ϵ��Ϊ��ǰ��λ��ϵ����
    With rsCheck
        Do While Not .EOF
            If mbln��ȷ���� Then
                dbl�������� = zlStr.NVL(!��������, 0)
                gstrSQL = " Select Nvl(��������,0) ��������,Nvl(ʵ������,0) ʵ������,Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ��," & _
                          "     Nvl(����,0) ����,Ч��,���Ч��,�ϴ����� ����,�ϴβ��� ����,��׼�ĺ� " & _
                          " From ҩƷ��� Where �ⷿID=[1] And ҩƷID=[2]  And ����=1"
                If gSystem_Para.P156_�����㷨 = 0 Then
                    gstrSQL = gstrSQL & " Order by Nvl(����,0)"
                Else
                    gstrSQL = gstrSQL & " Order by Ч��,Nvl(����,0)"
                End If
                          
                Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�����ĵĿ��", lngTargetID, Val(zlStr.NVL(rsCheck!����ID)))
                          
                blnData = False
                If rsStock.RecordCount <> 0 Then
                    '�п������ġ�������ʱ�����İ��˲���
                    Do While Not rsStock.EOF
                        If dbl�������� >= rsStock!�������� Then
                            dbl��д���� = rsStock!��������
                        Else
                            dbl��д���� = dbl��������
                        End If
                        If rsStock!�������� < 0 Then
                            dbl��д���� = 0
                        End If
                     
                         blnData = SetColValue(!����ID, "[" & !���� & "]" & !����, IIf(IsNull(!���), "", !���), IIf(IsNull(rsStock!����), "", rsStock!����), _
                            IIf(mintUnit = 0, !�ۼ۵�λ, !��װ��λ), _
                            !�ۼ�, IIf(IsNull(rsStock!����), "", rsStock!����), _
                            IIf(IsNull(rsStock!Ч��), "", rsStock!Ч��), IIf(IsNull(!���Ч��), 0, !���Ч��), _
                            IIf(zlStr.NVL(rsStock!���Ч��) = "", "", Format(rsStock!���Ч��, "yyyy-mm-dd")), _
                            !�ⷿ����, zlStr.NVL(!��������, 0), _
                            IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !ָ�������, _
                            IIf(mintUnit = 0, 1, !����ϵ��), _
                            rsStock!����, dbl��д����, !���÷���, !�Ƿ���, IIf(IsNull(rsStock!��׼�ĺ�), "", rsStock!��׼�ĺ�))
                        
                        dbl�������� = dbl�������� - dbl��д����
                        If dbl�������� = 0 Then Exit Do
                        rsStock.MoveNext
                    Loop
                    With frmRequestStuffCard.mshBill
                          If dbl�������� <> 0 And blnData And optMode(�������쵥����).Value = False Then
                            'δ�����������ȫ���������һ�еĲ�����
                            .TextMatrix(.Rows - 2, mconIntCol��д����) = Format(Val(.TextMatrix(.Rows - 2, mconIntCol��д����)) + dbl�������� / IIf(Val(.TextMatrix(.Rows - 2, mconIntCol����ϵ��)) = 0, 1, Val(.TextMatrix(.Rows - 2, mconIntCol����ϵ��))), mFMT.FM_����)
                          End If
                    End With
        
                Else
                    '����������ʱ�����Ե����İ��˲���
                    '�������Ϊ�����ֹ��������ִ���������
                    If mintCheck <> 2 Then
                        gstrSQL = " Select Nvl(A.�ⷿ����,0) �ⷿ����,Nvl(A.���÷���,0) ���÷���,Nvl(B.�Ƿ���,0) ʱ�� " & _
                                  " From �������� A,�շ���ĿĿ¼ B" & _
                                  " Where A.����ID = B.ID And A.����ID =[1] "
                        
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ķ��ڳ���ⷿ�Ƿ������ʱ�۵�����", Val(zlStr.NVL(!����ID)))
                                  
                        bln��ҩ = (rsTemp!ʱ�� = 1) Or IIf(bln�ⷿ, (rsTemp!�ⷿ���� = 1), (rsTemp!���÷��� = 1))
                        If Not bln��ҩ Then
                            If Not bln��ʾ Then
                                If mintCheck = 1 Then
                                    bln���� = (MsgBox("�޿�������Ƿ�������죿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
                                Else
                                    bln���� = True
                                End If
                                bln��ʾ = True
                            End If
                            If bln���� Then
                                'Ϊ�޿�����Ĳ��������¼
                                Call SetColValue(!����ID, "[" & !���� & "]" & !����, IIf(IsNull(!���), "", !���), "", _
                                    IIf(mintUnit = 0, !�ۼ۵�λ, !��װ��λ), _
                                    !�ۼ�, "", "", IIf(IsNull(!���Ч��), 0, !���Ч��), "", !�ⷿ����, IIf(IsNull(!��������), 0, !��������), _
                                    IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !ָ�������, _
                                    IIf(mintUnit = 0, 1, !����ϵ��), _
                                    0, !��������, !���÷���, !�Ƿ���, "")
                            End If
                        End If
                    End If
                End If
            Else
                '���ݴ����¼����������
                Call SetColValue(!����ID, "[" & !���� & "]" & !����, IIf(IsNull(!���), "", !���), IIf(IsNull(!����), "", !����), _
                    IIf(mintUnit = 0, !�ۼ۵�λ, !��װ��λ), _
                    !�ۼ�, "", "", IIf(IsNull(!���Ч��), 0, !���Ч��), "", !�ⷿ����, IIf(IsNull(!��������), 0, !��������), _
                    IIf(IsNull(!ʵ�ʽ��), 0, !ʵ�ʽ��), IIf(IsNull(!ʵ�ʲ��), 0, !ʵ�ʲ��), !ָ�������, _
                    IIf(mintUnit = 0, 1, !����ϵ��), _
                    0, !��������, !���÷���, !�Ƿ���, "")
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'������Ŀ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal lng����ID As Long, ByVal strҩ�� As String, ByVal str��� As String, _
    ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal str���ʧЧ�� As String, ByVal int�������� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal dbl���� As Double, ByVal int���÷��� As Integer, ByVal int�Ƿ��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
      
    Dim intCount As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    
    Dim numʵ������ As Double
    Dim rsTemp As New ADODB.Recordset
        
 
    On Error GoTo ErrHandle
       
    SetColValue = False
    
    '�����������Ϊ�����˳�
    If IIf(dbl���� >= num��������, num��������, dbl����) = 0 And mbln��ȷ���� And (int�Ƿ��� = 1 Or lng���� <> 0) Then Exit Function
    
    '�����ʱ�۲�������ȷ����,��Ҫ���¼����ۼ�;�����Դ������ۼ�Ϊ׼
    If int�Ƿ��� = 1 Then
        'ȡʵ������
        gstrSQL = " Select nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ������,0) ʵ������ From ҩƷ��� " & _
                " Where ����=1 And ҩƷID=[1] And �ⷿID=[2] And Nvl(����,0)=[3]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʱ�۲���,��Ҫ���¼����ۼ�", lng����ID, cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), lng����)
        If Not rsTemp.EOF Then
            If rsTemp!ʵ������ > 0 Then
                num�ۼ� = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������
            End If
        End If
    End If
    
    With frmRequestStuffCard.mshBill
        intRow = .Rows - 1
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, 1) = intRow
        .TextMatrix(intRow, mconIntColҩ��) = strҩ��
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
        
        .TextMatrix(intRow, mconIntCol�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mconIntCol��������) = int��������
        .TextMatrix(intRow, mconIntCol��������) = Format(num�������� / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mconIntCol���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntColָ�������) = numָ�������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        '�����ʱ�۲��ϻ��������,���ܳ�����ǰ�������
        If (int�Ƿ��� = 1 Or lng���� <> 0) And mbln��ȷ���� Then
            .TextMatrix(intRow, mconIntCol��д����) = Format(IIf(dbl���� >= num��������, num��������, dbl����) / num����ϵ��, mFMT.FM_����)
            .TextMatrix(intRow, mconIntColʵ������) = Format(IIf(dbl���� >= num��������, num��������, dbl����) / num����ϵ��, mFMT.FM_����)
        Else
            .TextMatrix(intRow, mconIntCol��д����) = Format(dbl���� / num����ϵ��, mFMT.FM_����)
            .TextMatrix(intRow, mconIntColʵ������) = Format(dbl���� / num����ϵ��, mFMT.FM_����)
        End If
        
        If .TextMatrix(intRow, mconIntCol�ۼ�) <> "" Then
            .TextMatrix(intRow, mconIntCol�ۼ۽��) = Format(.TextMatrix(intRow, mconIntCol�ۼ�) * .TextMatrix(intRow, mconIntCol��д����), mFMT.FM_���)
        End If
        
        
        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
        'cboStock.ItemData(cboStock.ListIndex)
        Call ��֤�����ۼ���(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), lng����ID, lng����, _
            num����ϵ�� + 0, numʵ�ʲ��, numʵ�ʽ��, _
            numָ������� / 100, Val(.TextMatrix(intRow, mconIntCol��д����)), Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), _
            dbl���, dbl����, dbl�ɱ����)
            
        .TextMatrix(intRow, mconintCol���) = Format(dbl���, mFMT.FM_���)
        .TextMatrix(intRow, mconIntCol�ɹ���) = Format(dbl����, mFMT.FM_�ɱ���)
        .TextMatrix(intRow, mconIntCol�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
'
'
'        If .TextMatrix(intRow, mconIntColʵ�ʽ��) = 0 Then
'            .TextMatrix(intRow, mconintCol���) = Format(.TextMatrix(intRow, mconIntCol�ۼ۽��) * .TextMatrix(intRow, mconIntColָ�������) / 100, mFMT.FM_���)
'        Else
'            .TextMatrix(intRow, mconintCol���) = Format(.TextMatrix(intRow, mconIntCol�ۼ۽��) * (.TextMatrix(intRow, mconIntColʵ�ʲ��) / .TextMatrix(intRow, mconIntColʵ�ʽ��)), mFMT.FM_���)
'        End If
'        .TextMatrix(intRow, mconIntCol�ɹ���) = Format((.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconintCol���)) / IIf(Val(.TextMatrix(intRow, mconIntCol��д����)) = 0, 1, Val(.TextMatrix(intRow, mconIntCol��д����))), mFMT.FM_�ɱ���)
'        .TextMatrix(intRow, mconIntCol�ɹ����) = Format(.TextMatrix(intRow, mconIntCol�ɹ���) * .TextMatrix(intRow, mconIntCol��д����), mFMT.FM_���)
'
        .Rows = .Rows + 1
    End With
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lng�ⷿID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '���ָ���ⷿ�����Ŀ⡢���ϲ��Ż����Ƽ���(����Ŀⷿ�϶������Ŀ⡢���ϲ��Ż��Ƽ����е�һ��)
    
    On Error GoTo ErrHandle
    gstrSQL = " Select ����ID From ��������˵�� " & _
              " Where (�������� like '���ϲ���' Or �������� like '%�Ƽ���') And ����id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��ǲ��Ƿ��ϲ��Ż��Ƽ���", lng�ⷿID)
              
    If rsCheck.EOF Then
        CheckStock = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPara(ByVal lng�ⷿID As Long)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    '��ȡ������Ĳ�������ֵ��0-�����;1-��飬��������;2-�����ֹ��
    gstrSQL = " Select Nvl(��鷽ʽ,0) Value From ���ϳ����� Where �ⷿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ĳ���", lng�ⷿID)
    
    If Not rsTemp.EOF Then
        mintCheck = rsTemp!Value
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '���Ƿ������ֵܽӵ��Ƿ�Ҳȫ��TRUE�����ǣ������丸�ڵ�ҲΪTRUE�����򣬲���
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If tvw����.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw����.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If tvw����.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck
        End If
    End If
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw����.Nodes.count
        If tvw����.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function

Private Sub tvw����_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub



