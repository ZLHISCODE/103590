VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmStuffRxSend 
   Caption         =   "���ĵ��ݷ���"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11595
   Icon            =   "frmStuffRxSend.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2295
      ScaleWidth      =   3255
      TabIndex        =   16
      Top             =   1800
      Width           =   3255
      Begin VB.Frame fraLine 
         Height          =   3375
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   45
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7695
      Left            =   120
      ScaleHeight     =   7695
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3975
         ScaleWidth      =   3255
         TabIndex        =   14
         Top             =   2160
         Width           =   3255
         Begin XtremeSuiteControls.TabControl tbcList 
            Height          =   975
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1455
            _Version        =   589884
            _ExtentX        =   2566
            _ExtentY        =   1720
            _StockProps     =   64
            Enabled         =   -1  'True
         End
      End
      Begin VB.PictureBox picConMain 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   120
         Width           =   3495
         Begin VB.CheckBox ChkShowPro 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʾ���й��̵���"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton cmdIC 
            Caption         =   "����"
            Height          =   300
            Left            =   2760
            TabIndex        =   6
            Top             =   1200
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   120
            Width           =   2415
         End
         Begin VB.TextBox txtPati 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   960
            TabIndex        =   4
            Top             =   1200
            Width           =   1245
         End
         Begin VB.CommandButton cmdFind 
            Height          =   300
            Left            =   3000
            Picture         =   "frmStuffRxSend.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "������λ(F2)"
            Top             =   1200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox ChkShowReturn 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʾ��ҩ��������"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker DtpEndTime 
            Height          =   315
            Left            =   960
            TabIndex        =   7
            Top             =   840
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   276299779
            CurrentDate     =   36985
         End
         Begin MSComCtl2.DTPicker DtpBeginTime 
            Height          =   315
            Left            =   960
            TabIndex        =   8
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   276299779
            CurrentDate     =   36985
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "ʱ�䷶Χ"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   180
            Width           =   720
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   11
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   120
            TabIndex        =   10
            Top             =   900
            Width           =   720
         End
         Begin VB.Image imgFilter 
            Height          =   240
            Left            =   2400
            Picture         =   "frmStuffRxSend.frx":699C
            Top             =   1200
            Width           =   240
         End
         Begin VB.Label lblPati 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���￨��"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Tag             =   "0"
            ToolTipText     =   "���˶�λ(F3)"
            Top             =   1245
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   19
      Top             =   8085
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffRxSend.frx":D1EE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   6480
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStuffRxSend.frx":DA82
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   7920
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmStuffRxSend.frx":13B84
      Left            =   3960
      Top             =   960
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStuffRxSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�ļ����ֲ˵�����
Private Const mconMenu_FilePopup = 1
Private Const mconMenu_File_PrintSet = 11
Private Const mconMenu_File_Preview = 12
Private Const mconMenu_File_Print = 13
Private Const mconMenu_File_Excel = 14
Private Const mconMenu_File_Recipe_BillPrintSend = 15
Private Const mconMenu_File_Recipe_BillPrintReturn = 16
Private Const mconMenu_File_Parameter = 17
Private Const mconMenu_File_Exit = 18

'�༭���ֲ˵�����
Private Const mconMenu_EditPopup = 2
Private Const mconMenu_Edit_Recipe_Send = 21
Private Const mconMenu_Edit_Recipe_Return = 22
Private Const mconMenu_Edit_Recipe_Batch = 23
Private Const mconMenu_Edit_Recipe_SendByBill = 24
Private Const mconMenu_Edit_Recipe_SendOther = 25
Private Const mconMenu_File_Recipe_ReturnByBill = 26
Private Const mconMenu_Edit_Recipe_Flag = 27

'�鿴���ֲ˵�
Private Const mconMenu_ViewPopup = 3
Private Const mconMenu_View_ToolBar = 31
Private Const mconMenu_View_ToolBar_Button = 311
Private Const mconMenu_View_ToolBar_Text = 312
Private Const mconMenu_View_ToolBar_Size = 313
Private Const mconMenu_View_StatusBar = 32
Private Const mconMenu_View_FontSize = 33
Private Const mconMenu_View_FontSize_1 = 331
Private Const mconMenu_View_FontSize_2 = 332
Private Const mconMenu_View_FontSize_3 = 333
Private Const mconMenu_View_Refresh = 34

'�������ֲ˵�
Private Const mconMenu_HelpPopup = 4
Private Const mconMenu_Help_Help = 41
Private Const mconMenu_Help_Web = 42
Private Const mconMenu_Help_Web_Home = 421
Private Const mconMenu_Help_Web_Mail = 422
Private Const mconMenu_Help_About = 43

'�Ҽ������˵�
Private Const mconMenu_InputPopup = 5
Private Const mconMenu_Input_Recipe_NO = 51
Private Const mconMenu_Input_Recipe_OPNO = 52
Private Const mconMenu_Input_Recipe_Name = 53
Private Const mconMenu_Input_Recipe_IDCard = 54
Private Const mconMenu_Input_Recipe_ICCard = 55
Private Const mconMenu_Input_Recipe_MINo = 56
Private Const mconMenu_Input_Recipe_HosNumber = 57

'�б��ҳ����
Private Const mconTab_Recipe_Send = 0
Private Const mconTab_Recipe_Return = 1

'�Ӵ��嶨��
Private mfrmList As New frmStuffRxList
Private mfrmDetail As New frmStuffRxDetail

Private mstrCardType As String
Private mintCardCount As Integer
Private mint���￨���� As Integer
Private mlng�ⷿid As Long  '��ǰ���ϲ���
Private mlngModule As Long      'ģ���
Private mstrPrivs As String     'Ȩ���ַ���
Private mlngIC����id As Long
Private mblnCard As Boolean
Private mintMoneyDigit As Integer
Private mint������� As Integer

'��ǰ���ҵ���������
Private mint����ģʽ As Integer
Private mstrContent As String

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private Enum mFindType
    ���ݺ� = 0
    ����� = 1
    ���� = 2
    ���֤ = 3
    IC�� = 4
    ҽ���� = 5
    סԺ�� = 6
End Enum

'��������ṹ
Private Type TYPE_Para
    ���ĵ�λ As Integer
    �������� As String
    �շѵ��� As Integer
    intFont  As Integer
    intTool As Integer
    blnTool As Boolean
End Type

Private Enum mTimeRange
    ���� = 0
    ������ = 1
    ������ = 2
    ָ��ʱ�䷶Χ = 3
End Enum

Private T_Para As TYPE_Para

Private Sub InitComandBars()
    '��ʼ���˵�������ȫ���˵����������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPane As Pane
    Dim blnGroup As Boolean
    Dim intCount As Integer
    Dim strCardName As String
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgList.Icons
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsMain.ActiveMenuBar.Title = "�˵�"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    '�ļ����ֲ˵�
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "�����&Excel��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintSend, "��ӡ�����嵥(&W)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_BillPrintReturn, "��ӡ����֪ͨ��(&R)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&T)")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    '�༭���ֲ˵�
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Send, "����(&B)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Return, "����(&B)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Batch, "����������(&B)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendByBill, "��Ʊ�ݺŷ���(&I)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_SendOther, "�������ⷿ�Ĵ���(&F)")
'        '�жϵ�ǰ�Ƿ���Ȩ����ʾ
'        If InStr(1, mstrPrivs, "�������ⷿ�Ĵ���") > 0 Then
'            cbrControlMain.Visible = True
'        Else
'            cbrControlMain.Visible = False
'        End If
'
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Recipe_ReturnByBill, "�����ݺ�����(&I)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Flag, "ֹͣ���ϱ��(&S)")
'        '����Ȩ���ж�ֹͣ�����Ƿ���ʾ
'        cbrControlMain.Visible = (mPrives.blnֹͣ��ҩ = True Or mPrives.bln�ָ���ҩ = True)
'        blnGroup = (mPrives.blnֹͣ��ҩ = True Or mPrives.bln�ָ���ҩ = True)
'        cbrControlMain.BeginGroup = blnGroup
    End With
    
    '�鿴���ֲ˵�
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "������(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
        cbrControl.Checked = T_Para.blnTool
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
        If T_Para.intTool = 0 Or T_Para.intTool = 1 Then cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        If T_Para.intTool = 0 Or T_Para.intTool = 2 Then cbrControl.Checked = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)")
        cbrControlMain.Checked = True
        
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "����(&F)")
        cbrControlMain.BeginGroup = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "С����(&S)", -1, False)
        If T_Para.intFont = 0 Then cbrControl.Checked = True
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "������(&M)", -1, False)
        If T_Para.intFont = 1 Then cbrControl.Checked = True
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "������(&B)", -1, False)
        If T_Para.intFont = 2 Then cbrControl.Checked = True
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)")
        cbrControlMain.BeginGroup = True
    End With
    
    '�������ֲ˵�
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��")
        cbrControlMain.BeginGroup = True
    End With
    
    '�����
    With Me.cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), mconMenu_File_Print
        .Add FCONTROL, Asc("C"), mconMenu_Edit_Recipe_Send
        .Add FCONTROL, Asc("H"), mconMenu_Edit_Recipe_Return

        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help

    End With

    '���õ����˵�
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_InputPopup, "¼��(&I)", -1, False)
    cbrMenuBar.Id = mconMenu_InputPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_NO, "���ݺ�(&0)")
        cbrControlMain.Parameter = "��|���ݺ�|0||||||"
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_OPNO, "�����(&1)")
        cbrControlMain.Parameter = "��|�����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_Name, "����(&2)")
        cbrControlMain.Parameter = "��|����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_IDCard, "���֤(&3)")
        cbrControlMain.Parameter = "��|���֤|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_ICCard, "IC��(&4)")
        cbrControlMain.Parameter = "IC|IC����|1|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_MINo, "ҽ����(&5)")
        cbrControlMain.Parameter = "ҽ|ҽ����|0|||||"
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber, "סԺ��(&6)")
        cbrControlMain.Parameter = "ס|סԺ��|0|||||"
        
        '��̬ȡ����ҽ�ƿ�����Ҫ�����ѿ���
        If mstrCardType <> "" Then
            mintCardCount = UBound(Split(mstrCardType, ";")) + 1
            For intCount = 0 To UBound(Split(mstrCardType, ";"))
                'ȡ���ѿ�����
                strCardName = Split(Split(mstrCardType, ";")(intCount), "|")(1)
                
                Set cbrControlMain = .Add(xtpControlButton, mconMenu_Input_Recipe_HosNumber + intCount + 1, strCardName & "(&" & intCount + 7 & ")")
                
                '���濨��Ϣ
                cbrControlMain.Parameter = Split(mstrCardType, ";")(intCount)
                
                If intCount = 0 Then
                    cbrControlMain.BeginGroup = True
                End If
                
                If Split(cbrControlMain.Parameter, "|")(gCardFormat.����) = "��" Then
                    mint���￨���� = Val(Split(cbrControlMain.Parameter, "|")(gCardFormat.���ų���))
                End If
            Next
        End If
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Send, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_Recipe_Return, "����")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "����")
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub SetComandBars()
    '������Ĳ˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    On Error GoTo ErrHandle
    
    '��������
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_File_Parameter, , True)
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (InStr(1, mstrPrivs, "��������") > 0)
    
    '����
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (InStr(1, mstrPrivs, "�������Ϸ���") > 0)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (InStr(1, mstrPrivs, "�������Ϸ���") > 0)
    
    '����
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (InStr(1, mstrPrivs, "������������") > 0)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (InStr(1, mstrPrivs, "������������") > 0)
    
    'ֹͣ���ϲ���
    Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Flag, , True)
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = ((InStr(1, mstrPrivs, "ֹͣ����") > 0) Or (InStr(1, mstrPrivs, "�ָ�����") > 0))
    
    '���ð�ť����״̬
    If Me.tbcList.Selected.Index = 0 Then
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
        Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
        Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    Else
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
        Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Send, , True)
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        
        Set cbrMenu = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
        Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Return, , True)
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetInputState(ByVal intType As Integer)
    Dim cbrControl As CommandBarControl
    
    Set cbrControl = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Input_Recipe_NO + intType, , True)
    If Not cbrControl Is Nothing Then
        SetInputPopupCheck cbrControl
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    
    Select Case Control.Id
        '��ӡ����
        Case mconMenu_File_PrintSet
            zlPrintSet
        'Ԥ��
        Case mconMenu_File_Preview
            zlSubPrint 2
        '��ӡ
        Case mconMenu_File_Print
            zlSubPrint 1
        '�����Excel
        Case mconMenu_File_Excel
            zlSubPrint 3
        '��ӡ�ѷ����嵥
        Case mconMenu_File_Recipe_BillPrintSend
        '��ӡ����֪ͨ��
        Case mconMenu_File_Recipe_BillPrintReturn
        '��������
        Case mconMenu_File_Parameter
            Call ReSetPara
        '����,����
        Case mconMenu_Edit_Recipe_Send, mconMenu_Edit_Recipe_Return
            Call Stuff_Work(Me.tbcList.Selected.Index)
        '��������
        Case mconMenu_Edit_Recipe_Batch
            Call ShowSendBatch(False)
        '��Ʊ�ݺŷ���
        Case mconMenu_Edit_Recipe_SendByBill
            Call ShowSendBatch(True)
        '�������ⷿ����
        Case mconMenu_Edit_Recipe_SendOther
            Call ShowSendOther
        'ֹͣ���ϱ��
        Case mconMenu_Edit_Recipe_Flag
            Call ShowFlag
        '�����ݺ�����
        Case mconMenu_File_Recipe_ReturnByBill
            Call ShowReturn
        '���ù�����-�ı�+ͼ����ʾ
        Case mconMenu_View_ToolBar_Button
            T_Para.blnTool = Not Control.Checked
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        '���ù�����-�ı���ʾ
        Case mconMenu_View_ToolBar_Text
             If T_Para.intTool = 0 Then
                T_Para.intTool = 2
            ElseIf T_Para.intTool = 2 Then
                T_Para.intTool = 0
            ElseIf T_Para.intTool = 3 Then
                T_Para.intTool = 1
            ElseIf T_Para.intTool = 1 Then
                T_Para.intTool = 3
            End If
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        '���ù�����-��ͼ��
        Case mconMenu_View_ToolBar_Size
            If T_Para.intTool = 0 Then
                T_Para.intTool = 1
            ElseIf T_Para.intTool = 1 Then
                T_Para.intTool = 0
            ElseIf T_Para.intTool = 3 Then
                T_Para.intTool = 2
            ElseIf T_Para.intTool = 2 Then
                T_Para.intTool = 3
            End If
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        '����״̬��
        Case mconMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        '����С����
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3
            T_Para.intFont = Control.Parameter
            Call SetFontSize(Control.Parameter)
        'ˢ��
        Case mconMenu_View_Refresh
            Call RefreshData
        '����
        Case mconMenu_Help_Help                         '����
            Call ShowHelp(App.ProductName, Me.hwnd, "frmStuffRxSend")
        Case mconMenu_Help_Web                          'WEB�ϵ�����
            Case mconMenu_Help_Web_Home                     '������ҳ
                Call zlHomePage(Me.hwnd)
            Case mconMenu_Help_Web_Mail                     '���ͷ���
                Call zlMailTo(Me.hwnd)
        Case mconMenu_Help_About                        '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        '�˳�
        Case mconMenu_File_Exit
            Unload Me
        Case Else
            If Control.Id > 401 And Control.Id < 499 Then
                'ִ���Զ��屨��
'                Call BillPrint_Custom(Control)
            End If
            
            '�����˵�
            If Control.Id >= mconMenu_Input_Recipe_NO And Control.Id <= mconMenu_Input_Recipe_NO + 6 + mintCardCount Then
                Call SetInputPopupCheck(Control)
                mint����ģʽ = Control.Id - mconMenu_Input_Recipe_NO
                mstrContent = ""
            '������Ŀ�����˵�
            End If
    End Select
End Sub

Private Sub ReSetPara()
    frmStuffParaSet.ShowSetPara Me, mlngModule, mstrPrivs
    
    GetParaMe
    Call GetPara(mlngModule)
    Call RefreshData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '����
            Control.Checked = Val(Control.Parameter) = T_Para.intFont
    End Select
End Sub

Private Sub zlSubPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, ObjAppRow As zlTabAppRow

    Set objPrint.Body = mfrmDetail.VSFDetail
    objPrint.Title.Text = tbcList.Selected.Caption & "���"
    Set ObjAppRow = New zlTabAppRow
    Call ObjAppRow.Add("��ӡ��:" & gstrUserName)
    Call ObjAppRow.Add("��ӡʱ��:" & Format(sys.Currentdate, "yyyy��MM��DD��"))
    Call objPrint.BelowAppRows.Add(ObjAppRow)
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub ShowSendBatch(ByVal blnƱ�ݺŷ��� As Boolean)
'�����������Ͻӿڣ�blnƱ�ݺŷ���=false
'���ð�Ʊ�ݺŷ��Ͻӿڣ�blnƱ�ݺŷ���=true
    With Frm�����ŷ���
        .In_���� = 0
        .In_����IN = Me.txtPati.Text
        .In_���ϲ���id = mlng�ⷿid
        .In_����� = GetCheckPara(mlng�ⷿid)
        .In_����δ���Ϸ��� = 1
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        .��Ʊ�ݺŷ��� = blnƱ�ݺŷ���
        .Show 1, Me
    End With
    
    Call RefreshData
End Sub

Private Sub ShowSendOther()
'���ô����Ͻӿ�
    With frm������
        .In_���� = 0
        .In_����IN = Me.txtPati.Text
        .In_���ϲ���id = mlng�ⷿid
        .In_����� = GetCheckPara(mlng�ⷿid)
        .In_����δ���Ϸ��� = 1
        .In_Ȩ�� = mstrPrivs
        .mstr������ = gstrUserName
        .Show 1, Me
    End With
    
    Call RefreshData
End Sub


Private Sub ShowReturn()
'���ð����ݺ�����
    If Frm����������.ShowCard(Me, mlng�ⷿid, mstrPrivs) = False Then Exit Sub
    Call RefreshData
End Sub
Private Sub ShowFlag()
 '-----------------------------------------------------------------------------------------------------------
    '����:ֹͣ���ϱ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    'ֹͣ����
    '��ҩ��ʽ=-1
    Dim frmFlag As New Frm���ٷ�ҩ������־
    frmFlag.In_����� = GetCheckPara(mlng�ⷿid)
    
    '��ֹͣ���ϵķ��ϲ���id���и�ֵ
    frmFlag.In_�ⷿid = mlng�ⷿid
    
    frmFlag.gstrParentName = Replace(Me.Name, "_New", "")
    frmFlag.Show 1, Me
    
    Call RefreshData
End Sub

Private Sub SetFontSize(ByVal intType As Integer)
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case intType
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    mfrmDetail.SetFontSize intFont
    mfrmList.SetFontSize intFont
    Me.FontSize = intFont
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub ChkShowPro_Click()
    Call RefreshData
End Sub

Private Sub ChkShowReturn_Click()
    Call RefreshData
End Sub

Private Sub cmdIC_Click()
    Dim strOutXML As String
    Dim strText As String
    
    If Val(lblPati.Tag) = mFindType.IC�� Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPati.Text = mobjICCard.Read_Card()
            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
        End If
    Else
        If Not gobjSquareCard Is Nothing Then
            Call gobjSquareCard.zlReadCard(Me, mlngModule, Val(Split(txtPati.Tag, "|")(gCardFormat.�����ID)), True, "", strText, strOutXML)
            txtPati.Text = strText
            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
        End If
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnFirst As Boolean
    Dim strInput As String
    
    If KeyCode = vbKeyF3 Then
        If imgFilter.BorderStyle = 0 Then
            If txtPati.Text = "" Then
                txtPati.SetFocus
            Else
                Call txtPati_Validate(False)
                Call zlControl.TxtSelAll(txtPati)
                If Val(lblPati.Tag) = mFindType.IC�� Then
                    strInput = mlngIC����id
                ElseIf Val(lblPati.Tag) <= 6 Then
                    strInput = txtPati.Text
                Else
                    '���ѿ����ʱ����Ϊ��ID+����
                    strInput = Split(txtPati.Tag, "|")(gCardFormat.�����ID) & "|" & txtPati.Text
                End If
                Call mfrmList.FindSpecialRow(Val(lblPati.Tag), strInput)
            End If
        Else
            RefreshData
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim dteTime As Date
    
    dteTime = sys.Currentdate
    mstrPrivs = gstrPrivs
    mlngModule = glngModul
    Set mfrmDetail = New frmStuffRxDetail
    Set mfrmList = New frmStuffRxList
    
    DtpBeginTime.Value = Format(dteTime, "yyyy-MM-dd 00:00:00")
    DtpEndTime.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    
    '��ʼ��IC��ID������
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    mintMoneyDigit = GetDigit
    
    '���ز˵�
    Call InitComandBars

    '��ʼ�����ؼ�
    Call InitPanes
    
    '��ʼ��ҳ��ؼ�
    Call InitTabControl
    
    '����ʱ�䷶Χ
    Call LoadTime
    
    '���ز���
    Call GetPara(mlngModule)
    
    '���ر��ز���
    Call GetParaMe
    
    Call SetInputState(mint����ģʽ)
    
    '���ð�ť״̬
    Call SetComandBars
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        '�ָ�������ʷ״̬
        Call RestoreWinState(Me, App.ProductName)
        Call SetFontSize(T_Para.intFont)
        
        '�ָ�����
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    End If
End Sub

Private Sub GetParaMe()
'��ȡ���ز���
    On Error GoTo ErrHandle
    T_Para.intFont = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "��������", 0)
    T_Para.intTool = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "������", 0)
    T_Para.blnTool = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "��������ʾ", True)
    mlng�ⷿid = zlDatabase.GetPara("���Ͽ���", glngSys, mlngModule, "0")
    
    '���ز��ŷ������
    Call Get�������(mlng�ⷿid)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitPanes()
    Dim lngHeight As Long
    
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane
    Dim objPaneList As Pane
    Dim objPaneDetail As Pane
    
    lngHeight = 145
    
    If cboTime.ListIndex <> 3 Then
        lngHeight = lngHeight - 55
    End If
    
    Set objPaneCon = Me.dkpMain.CreatePane(1, 230, lngHeight, DockLeftOf, Nothing)
    objPaneCon.Title = "��������"
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable

    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then objPaneCon.Hidden = False
End Sub

Private Sub InitTabControl()
    '��ʼ����ҳ�ؼ�
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(0, "������ϸ�嵥", mfrmDetail.hwnd, 0).Tag = "������ϸ�嵥_"
        .Item(0).Selected = True
    End With
    
    With Me.tbcList
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_Recipe_Send, "����", mfrmList.hwnd, 0).Tag = "������_"
        .InsertItem(mconTab_Recipe_Return, "����", mfrmList.hwnd, 0).Tag = "����_"
        .Item(mconTab_Recipe_Return).Selected = True
        .Item(mconTab_Recipe_Send).Selected = True
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmStuffRxList = Nothing
    Set frmStuffRxDetail = Nothing
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "��������", T_Para.intFont)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "��������ʾ", T_Para.blnTool)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name, "������", T_Para.intTool)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgFilter_Click()
    imgFilter.BorderStyle = Abs(imgFilter.BorderStyle - 1)
    
    If imgFilter.BorderStyle = 0 Then
        imgFilter.ToolTipText = "����л�������ģʽ"
    Else
        imgFilter.ToolTipText = "����л�����λģʽ"
    End If
    
    '������涨λ��ʽ
    Call SaveSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "���ĵ��ݷ���", "���涨λ", imgFilter.BorderStyle)
    
    '����ˢ��
    mlngIC����id = 0
    
    txtPati.Text = ""
    Call RefreshData
End Sub

Private Sub lblPati_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 1 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_InputPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLine
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLine.Left
        .Width = picDetail.Width - fraLine.Width
        .Height = picDetail.Height - 50
    End With
    
    err.Clear
End Sub

Private Sub picCondition_Resize()
    On Error Resume Next
    
    With picConMain
        .Top = 0
        .Left = 0
        .Width = picCondition.Width
    End With
    
    With picList
        .Top = picConMain.Top + picConMain.Height
        .Left = 0
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top
    End With
    
    err.Clear
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    
    With tbcList
        .Move 0, 0, picList.Width, picList.Height - 50
    End With
    
    err.Clear
End Sub

Public Sub RefreshSendData(ByVal int���� As Integer, ByVal strNo As String, ByVal lng�ⷿID As Long, int��¼״̬, ByVal int�ɲ��� As Integer)
    Dim rsTemp As Recordset
    
    If Me.tbcList.Selected.Index = 0 Then
        Set rsTemp = Stuff_RxRefSendDetail(int����, strNo, lng�ⷿID)
    Else
        Set rsTemp = Stuff_RxRefReturnDetail(int����, strNo, lng�ⷿID, int��¼״̬)
    End If
    
    Call mfrmDetail.WriteSendList(Me.tbcList.Selected.Index, rsTemp, int�ɲ���)
End Sub

Private Sub picConMain_Resize()
    On Error Resume Next
    
    With cboTime
        .Width = picCondition.Width - .Left - 50
    End With

    If cboTime.ListIndex <> 3 Then
        lblTimeBegin.Visible = False
        DtpBeginTime.Visible = False
        lblTimeEnd.Visible = False
        DtpEndTime.Visible = False
        
        With lblPati
            .Top = lblTime.Top + lblTime.Height + 180
        End With
        
        With txtPati
            .Top = cboTime.Top + cboTime.Height + 60
        End With
    Else
        lblTimeBegin.Visible = True
        DtpBeginTime.Visible = True
        lblTimeEnd.Visible = True
        DtpEndTime.Visible = True
        
        With lblTimeBegin
            .Top = lblTime.Top + lblTime.Height + 180
        End With
        
        With DtpBeginTime
            .Top = cboTime.Top + cboTime.Height + 60
            .Width = cboTime.Width
        End With
        
        With lblTimeEnd
            .Top = lblTimeBegin.Top + lblTimeBegin.Height + 180
        End With
        
        With DtpEndTime
            .Top = DtpBeginTime.Top + DtpBeginTime.Height + 60
            .Width = cboTime.Width
        End With
        
        With lblPati
            .Top = lblTimeEnd.Top + lblTimeEnd.Height + 180
        End With
        
        With txtPati
            .Top = DtpEndTime.Top + DtpEndTime.Height + 60
        End With
    End If
    
    With cmdIC
        .Visible = (Val(Split(txtPati.Tag, "|")(gCardFormat.ˢ����־)) = 1)
        .Top = txtPati.Top
        .Left = picCondition.Width - .Width - 80
    End With

    With imgFilter
        .Top = txtPati.Top
        .Left = IIf(Val(Split(txtPati.Tag, "|")(gCardFormat.ˢ����־)) = 1, cmdIC.Left, picCondition.Width) - imgFilter.Width - 120
    End With
    
    With cmdFind
        .Top = cmdIC.Top
        .Left = imgFilter.Left + 120
    End With

    With txtPati
        .Width = imgFilter.Left - .Left - 200
    End With
    

    With ChkShowReturn
        .Left = lblPati.Left
        .Top = lblPati.Top + 350
    End With
    
    With ChkShowPro
        .Left = ChkShowReturn.Left
        .Top = ChkShowReturn.Top
    End With
    
    With picConMain
        .Height = ChkShowReturn.Top + ChkShowReturn.Height + 50
    End With
    
    err.Clear
End Sub

Private Sub cboTime_Click()
    With cboTime
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call picConMain_Resize
                Call picCondition_Resize
            End If
            .Tag = .ListIndex
        End If
        
        If .ListIndex < mTimeRange.ָ��ʱ�䷶Χ Then
            Select Case .ListIndex
                Case mTimeRange.����
                    Me.DtpBeginTime.Value = Format(sys.Currentdate, "yyyy-MM-dd 00:00:00")
                Case mTimeRange.������
                    Me.DtpBeginTime.Value = Format(sys.Currentdate - 1, "yyyy-MM-dd 00:00:00")
                Case mTimeRange.������
                    Me.DtpBeginTime.Value = Format(sys.Currentdate - 2, "yyyy-MM-dd 00:00:00")
            End Select
        
            Call RefreshData
        End If
    End With
End Sub

Private Sub LoadTime()
    With cboTime
        .Clear
        .AddItem "0-����"
        .AddItem "1-������"
        .AddItem "2-������"
        .AddItem "3-ָ��ʱ�䷶Χ"
        
        .ListIndex = 0
        .Tag = 0
    End With
End Sub


Private Sub RefreshData()
    Call ChangeSend(Me.tbcList.Selected.Index)
End Sub

Private Sub tbcList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    Call ChangeSend(Item.Index)
End Sub
Private Sub ChangeSend(ByVal intType As Integer)
    Dim rsTemp As Recordset
    
    If intType = 0 Then
        Me.ChkShowPro.Visible = False
        Me.ChkShowReturn.Visible = True
        
        'ˢ�´���ҩ�б�
        Set rsTemp = Stuff_RxRefSendNO(mlng�ⷿid, DtpBeginTime.Value, DtpEndTime.Value, mint����ģʽ, mstrContent, (Me.ChkShowReturn.Value = 1), (Me.imgFilter.BorderStyle = 1), mint�������)
    Else
        Me.ChkShowPro.Visible = True
        Me.ChkShowReturn.Visible = False
        
        'ˢ���ѷ�ҩ�б�
        Set rsTemp = Stuff_RxRefReturnNO(mlng�ⷿid, DtpBeginTime.Value, DtpEndTime.Value, mint����ģʽ, mstrContent, (ChkShowPro.Value = 1), (Me.imgFilter.BorderStyle = 1), mint�������)
    End If
    
    stbThis.Panels(2) = ""
    If Not rsTemp Is Nothing Then
        Call mfrmList.RefreshList(rsTemp, Me.tbcList.Selected.Index)
        
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        If Not rsTemp.EOF Then
            stbThis.Panels(2) = "����" & rsTemp.RecordCount & "�Ŵ�����" & "�ϼƽ��" & GetSumMoney(rsTemp) & "Ԫ"
        End If
    End If
       
    Call SetComandBars
End Sub

Private Function GetSumMoney(ByVal rsRecipt As ADODB.Recordset) As String
    Dim rsTemp As ADODB.Recordset
    Dim dblSum As Double
    
    Set rsTemp = rsRecipt.Clone
    
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            dblSum = dblSum + Val(.Fields("���").Value)
            .MoveNext
        Loop
    End With
    
    GetSumMoney = zlStr.FormatEx(dblSum, mintMoneyDigit)
End Function

Private Sub Stuff_Work(ByVal intType As Integer)
'�ù������ڴ����ϣ�����
'������intTypeҵ�����ͣ�0-���ϣ�1-����
    Dim rsTemp As Recordset
    Dim str��ҩ���� As String
    Dim strNo As String
    Dim int���� As Integer

    '��ȡ���������ݼ�
    Set rsTemp = mfrmDetail.GetWorkRs(intType, str��ҩ����)
    If rsTemp.RecordCount <> 0 Then
        rsTemp.MoveFirst
    Else
        Exit Sub
    End If
    If Not rsTemp.EOF Then
        strNo = rsTemp!NO
        int���� = rsTemp!����
        '���ù��̽��з��ϣ����ϲ���
        If Not Stuff_RxWork(intType, mstrPrivs, rsTemp, mlng�ⷿid, int����, strNo, str��ҩ����) Then Exit Sub
        
        
        '��ӡ����
        If intType = 0 Then
            If MsgBox("����Ҫ��ӡ���ϵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1727", Me, "NO=" & strNo, "����=" & int����, "�ⷿ==" & mlng�ⷿid, 1)
            End If
        Else
        End If
        'ˢ�½�������
        Call RefreshData
    End If
End Sub


Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim blnDoIt As Boolean
    Dim strInput As String
    Dim strCondition As String
    Dim i As Integer
    Dim blnˢ�� As Boolean
    Dim blnSta As Boolean
    Dim lng����id As Long
    Dim rsData As Recordset
    Dim str���� As String
    Dim str����id As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If lblPati.Tag = mFindType.IC�� Then
            If Not gobjSquareCard Is Nothing Then Call gobjSquareCard.zlGetPatiID("IC��", Trim(txtPati.Text), True, mlngIC����id)
            If txtPati.Text <> "" Then blnDoIt = True
        Else
            If Trim(txtPati.Text) <> "" Then blnDoIt = True
        End If
        
        If Val(lblPati.Tag) > mFindType.סԺ�� And Len(txtPati.Text) = txtPati.MaxLength - 1 And KeyAscii <> 8 Then blnˢ�� = True
    ElseIf KeyAscii <> 13 Then
        mblnCard = False
        If lblPati.Tag = mFindType.���� Then
            '�������
            mblnCard = zlCommFun.InputIsCard(txtPati, KeyAscii, glngSys)
        End If
        
        If mblnCard Then
            If lblPati.Tag = mFindType.���� Then
                If Len(txtPati.Text) = mint���￨���� - 1 And KeyAscii <> 8 And txtPati.SelLength <> Len(txtPati.Text) Then
                    txtPati.Text = txtPati.Text & Chr(KeyAscii)
                    txtPati.SelStart = Len(txtPati.Text)
                    KeyAscii = 0: blnDoIt = True
                End If
            End If
        Else
            Select Case lblPati.Tag
'                Case mFindType.���￨
'                    If InStr(":��;��?��''||" & Chr(22), Chr(KeyAscii)) > 0 Then
'                        KeyAscii = 0
'                    Else
'                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                    End If
                Case mFindType.�����
                    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
                Case mFindType.���ݺ�
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))

                    If Not (txtPati.Text = "" Or txtPati.SelLength = Len(txtPati.Text)) _
                        And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
                Case mFindType.����
                    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    End If
                Case Else
                    blnˢ�� = True
                    '�����������ѿ�
                    If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
                        KeyAscii = 0
                    Else
                        KeyAscii = Asc(UCase(Chr(KeyAscii)))
                    End If
                    
                    If Len(txtPati.Text) = txtPati.MaxLength - 1 And KeyAscii <> 8 Then
'                        txtPati.Text = Mid(txtPati.Text, 1, Len(txtPati.Text) - 1) & Chr(KeyAscii)
                        txtPati.Text = txtPati.Text & Chr(KeyAscii)
                        txtPati.SelStart = Len(txtPati.Text)
                        KeyAscii = 0
                        blnDoIt = True
                    End If
            End Select
        End If
    End If
    
    If blnˢ�� Then
        If Val(lblPati.Tag) <= 6 Then
            strInput = txtPati.Text
        Else
            '���ѿ����ʱ����Ϊ��ID+����
            strInput = Split(txtPati.Tag, "|")(gCardFormat.�����ID) & "|" & txtPati.Text
        End If
        
        mstrContent = zlfuncCard_GetPatiID(Val(Split(strInput, "|")(0)), Split(strInput, "|")(1))
    End If
    
    If blnDoIt Then
        DoEvents
        KeyAscii = 0
               
        If imgFilter.BorderStyle = 0 Then
            Call Form_KeyDown(vbKeyF3, 0)
        Else
            If Val(lblPati.Tag) = mFindType.���ݺ� Then
                If IsNumeric(txtPati.Text) Then
                    txtPati.Text = UCase(zlCommFun.GetFullNO(txtPati.Text, 13))
                End If
            End If
            
            mstrContent = txtPati.Text
            DoEvents
            Call RefreshData
        End If
        
        Call zlControl.TxtSelAll(txtPati)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
End Sub

Private Sub txtPati_GotFocus()
    If Not mobjIDCard Is Nothing And txtPati.Text = "" Then
        Call mobjIDCard.SetEnabled(True)
    End If
    
    If Not mobjICCard Is Nothing And txtPati.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
    
    txtPati.Tag = ""
    mstrContent = ""
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Sub SetInputPopupCheck(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_InputPopup)
    If Not cbrMenuBar Is Nothing Then
        For Each cbrControl In cbrMenuBar.CommandBar.Controls
            cbrControl.Checked = (cbrControl.Id = Control.Id)
        Next
        
        lblPati.Caption = Split(Control.Caption, "(")(0) & "��"
        lblPati.Tag = Val(Control.Id - mconMenu_Input_Recipe_NO)
        
'        mParams.int����ģʽ = Val(lblPati.Tag)
'        mintOld����ģʽ = mParams.int����ģʽ
        
        zlfuncCard_SetText txtPati, Control.Parameter
        
        picConMain_Resize
    End If
End Sub

Private Sub txtPati_Validate(Cancel As Boolean)
    If Val(lblPati.Tag) = mFindType.���ݺ� Then
        If IsNumeric(txtPati.Text) Then
            txtPati.Text = zlCommFun.GetFullNO(txtPati.Text, 13)
        End If
    End If
End Sub

Private Sub Get�������(ByVal lng�ⷿID As Long)
    Dim RecTestPeople As Recordset
    
    On Error GoTo ErrHandle
    If lng�ⷿID <> 0 Then
        gstrSQL = "Select nvl(�������,1) ������� From ��������˵�� Where ����ID+0=[1]"
        Set RecTestPeople = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ŷ������", lng�ⷿID)
        
        If Not RecTestPeople.EOF Then
            mint������� = RecTestPeople!�������
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
