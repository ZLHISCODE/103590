VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmStuffPriceList 
   Caption         =   "���ĵ��۹���"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   Icon            =   "frmStuffPriceList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4080
      ScaleHeight     =   255
      ScaleWidth      =   1935
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblExecute 
         AutoSize        =   -1  'True
         Caption         =   "����Ч"
         Height          =   180
         Left            =   1320
         TabIndex        =   12
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblNotExecute 
         AutoSize        =   -1  'True
         Caption         =   "δ��Ч"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   7560
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7680
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15240
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   885
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
      _cx             =   8705
      _cy             =   1561
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   975
      Left            =   3120
      TabIndex        =   4
      Top             =   4320
      Width           =   3015
      _cx             =   5318
      _cy             =   1720
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCost 
      Height          =   975
      Left            =   6840
      TabIndex        =   5
      Top             =   4440
      Width           =   3135
      _cx             =   5530
      _cy             =   1720
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1080
      MousePointer    =   7  'Size N S
      ScaleHeight     =   255
      ScaleWidth      =   7455
      TabIndex        =   6
      Top             =   3360
      Width           =   7455
      Begin VB.Label lblScope 
         Caption         =   "���ڷ�Χ��2012��11��1����2012��11��31��"
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3615
      End
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Left            =   0
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStuffPriceList.frx":6852
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmStuffPriceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconMenu_FilePopup As Long = 1 '�ļ�
Private Const mconMenu_ReportPopup As Long = 2 '����
Private Const mconMenu_EditPopup As Long = 3 '�༭
Private Const mconMenu_ViewPopup As Long = 4 '�鿴
Private Const mconMenu_HelpPopup As Long = 5 '����

'�ļ�
Private Const mconMenu_File_PrintSet = 100           '*��ӡ����(&S)��
Private Const mconMenu_File_Preview = 101            '*Ԥ��(&V)
Private Const mconMenu_File_Print = 102              '*��ӡ(&P)
Private Const mconMenu_File_BillPrint = 103 '���ݴ�ӡ��&B��
Private Const mconMenu_File_BillPreview = 104 '����Ԥ����&L��
Private Const mconMenu_File_Excel = 105              '�����&Excel��
Private Const mconMenu_File_Parameter = 106 '��������(&R)
Private Const mconMenu_File_Exit = 107 '�˳�(&E)
'�༭
Private Const mconMenu_Edit_Add = 200 '����(&A)
Private Const mconMenu_Edit_Update = 201 '�޸�(&U)
Private Const mconMenu_Edit_Delete = 202 'ɾ��(&D)
Private Const mconMenu_Edit_BatchPrice = 203 '����ִ�е���(&B)
'�鿴
Private Const mconMenu_View_Filter = 300 '����(&F)
Private Const mconMenu_View_Refresh = 301 'ˢ��(&R)
'����
Private Const mconMenu_Help_Title = 400 '��������(&H)
Private Const mconMenu_Help_Web = 401 'web������
Private Const mconMenu_Help_web_WebHome = 402 '������ҳ(&H)
Private Const mconMenu_Help_web_WebForum = 403 '������̳(&F)
Private Const mconMenu_Help_web_WebMail = 404 '���ͷ���(&K)
Private Const mconMenu_Help_About = 405 '����(&A)
Private Const mstrCaption As String = "���ĵ��۹���"

Private mstrPrivs As String

Private mlngForeColor As Long '��¼��ǰ��Ԫ��ǰ��ɫ

Private mintUnit As Integer     '��λ:0-ɢװ��λ,1-��װ��λ

Private Type Type_Condition '����ʱ���õ�����
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    dateִ��ʱ�俪ʼ As Date
    dateִ��ʱ����� As Date
End Type
Private mSQLCondition As Type_Condition

Private mFMT As g_FmtString

Private mstrResult As String '���˽��
'��������
Private mdaStart As Date
Private mdaEnd As Date
Private mdaVerifyStart As Date
Private mdaVerifyEnd As Date
Private mblnLoad As Boolean     '�Ƿ������� ture-�������
Private mintFindDay As Integer  '��ѯ������Χ

'����ȫ�ֱ���
Private Const mconlngRowHeight As Long = 300 '����и����и�

'���ۻ��ܱ�
Private Enum menuListCol
    ���ۺ� = 1
    ��������
    ������
    ��������
    ִ������
    ˵��
    ������
End Enum
'�ۼ۵��۱�
Private Enum menuPriceCol
    NO = 1
    ԭ��id
    ������Ϣ
    ���
    ��λ
    ��λϵ��
    ԭ��
    �ּ�
    ִ������
    ������
    ������
End Enum
'�ɱ��۵���
Private Enum menuCostCol
    NO = 1
    ������Ϣ
    �ⷿ
    ���
    ����
    ����
    ��λ
    ԭ��
    �ּ�
    Ч��
    ִ������
    ������
    ������
End Enum


Private Sub initCommandBars()
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "����������Ϣ��ҵ�������ι�˾" '��˾����
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        Set .Icons = imgList.Icons '���ù�����ͼ��ؼ�
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '����仯ʱ�������ʾ����˵�Ҳ������
        .ActiveMenuBar.Title = "�˵�"
    End With
End Sub

Private Sub initMenu()
'�����˵�
    Dim cbrMenuPopup As CommandBarPopup
    Dim cbrMenuControl As CommandBarControl

    With cbsMain
        '�ļ�
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)")
        cbrMenuPopup.Id = mconMenu_FilePopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Preview, "��ӡԤ��(&V)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_BillPrint, "���ݴ�ӡ(&B)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_BillPreview, "����Ԥ��(&L)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Excel, "�����Excel...")
            cbrMenuControl.BeginGroup = True
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&R)")
            cbrMenuControl.BeginGroup = True
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&E)")
        End With
        '����
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ReportPopup, "����(&R)")
        cbrMenuPopup.Id = mconMenu_ReportPopup
        cbrMenuPopup.Visible = False

        '�༭
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)")
        cbrMenuPopup.Id = mconMenu_EditPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Add, "����(&A)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Update, "�޸�(&U)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_Delete, "ɾ��(&D)")
'            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Edit_BatchPrice, "����ִ�е���(&B)")
'            cbrMenuControl.BeginGroup = True
        End With
        '�鿴
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)")
        cbrMenuPopup.Id = mconMenu_ViewPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_View_Filter, "����(&F)")
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)")
        End With
        '����
        Set cbrMenuPopup = .ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)")
        cbrMenuPopup.Id = mconMenu_HelpPopup
        With cbrMenuPopup.CommandBar.Controls
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Help_Title, "��������(&H)")
            Set cbrMenuControl = .Add(xtpControlPopup, mconMenu_Help_Web, "web������")
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebHome, "������ҳ(&H)", -1, False
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebForum, "������̳(&F)", -1, False
            cbrMenuControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_web_WebMail, "���ͷ���(&K)", -1, False
            Set cbrMenuControl = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)")
            cbrMenuControl.BeginGroup = True
        End With
    End With
End Sub

Private Sub InitTool()
    '����������
    Dim cbrToolBar As CommandBar
    Dim cbrMenuPopup As CommandBarPopup
    Dim cbrMenuControl As CommandBarControl

    Set cbrToolBar = cbsMain.Add("������", xtpBarTop)
    With cbrToolBar
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Add, "����")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Update, "�޸�")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Edit_Delete, "ɾ��")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_View_Filter, "����")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_Help_Title, "����")
        cbrMenuControl.BeginGroup = True
        Set cbrMenuControl = cbrToolBar.Controls.Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
    End With

    For Each cbrMenuControl In cbrToolBar.Controls  '�ù������а�ťͬʱ��ʾͼ�������
        cbrMenuControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitTabControl()
    '��ʼ��TabControl�ؼ�
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfList.Top + vsfList.Height + 400
    With TabCtlDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 1, "�ۼ۵���", vsfPrice.hwnd, 0
        .InsertItem 2, "�ɱ��۵���", vsfCost.hwnd, 0
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picSplit.Top - picSplit.Height - stbThis.Height
        .Item(0).Selected = True
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intMethod As Integer
    Dim blnPrivs As Boolean

    Select Case Control.Id
    Case mconMenu_Edit_Add '����
        frmStuffPriceCard.ShowMe Me, 0, "", 0
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_Edit_Update '�޸�
        If vsfList.Rows = 1 Then Exit Sub
        blnPrivs = CheckPrivs(intMethod)
        If blnPrivs = True Then
            frmStuffPriceCard.ShowMe Me, 1, vsfList.TextMatrix(vsfList.Row, menuListCol.���ۺ�), intMethod
            Call getListInfo
            vsfList.SetFocus
        Else
            MsgBox "����Ա���߱�" & IIf(intMethod = 0, "�ۼ۵���", IIf(intMethod = 1, "�ɱ��۵���", "�ۼ۳ɱ���һ�����")) & "Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
    Case mconMenu_Edit_Delete 'ɾ��
        If vsfList.Rows = 1 Then Exit Sub
        
        blnPrivs = CheckPrivs(intMethod)
        If blnPrivs = True Then
            Call deleteNotExecutePirce
        Else
            MsgBox "����Ա���߱�" & IIf(intMethod = 0, "�ۼ۵���", IIf(intMethod = 1, "�ɱ��۵���", "�ۼ۳ɱ���һ�����")) & "Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
    Case mconMenu_File_Exit '�˳�
        Unload Me
    Case mconMenu_View_Refresh 'ˢ��
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_View_Filter '����
        frmStuffPriceSearch.ShowMe Me, mstrResult, mSQLCondition.date����ʱ�俪ʼ, mSQLCondition.date����ʱ�����, mSQLCondition.dateִ��ʱ�俪ʼ, mSQLCondition.dateִ��ʱ�����
        Call getListInfo
        vsfList.SetFocus
    Case mconMenu_File_Parameter '��������
        frmParaset.���ò��� 1726, mstrPrivs, Me, mstrCaption
        mintUnit = IIf(Val(zldatabase.GetPara("���ĵ�λ", glngSys, 1726, "0")) = 1, 1, 0)
        Call initJinDu
        Call getListInfo
    Case mconMenu_File_PrintSet '��ӡ����
        Call zlPrintSet
    Case mconMenu_File_Preview '��ӡԤ��
        Call printbill(2)
    Case mconMenu_File_Print '��ӡ
        Call printbill(1)
    Case mconMenu_File_BillPrint '���ݴ�ӡ
        Call danJuprint(2)
    Case mconMenu_File_BillPreview '����Ԥ��
        Call danJuprint(1)
    Case mconMenu_File_Excel '�����Excel
        Call billExcel
    Case mconMenu_Help_About    '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case mconMenu_Help_Title '��������
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case mconMenu_Help_web_WebHome '������ҳ
        Call zlHomePage(Me.hwnd)
    Case mconMenu_Help_web_WebForum '������̳
        Call zlWebForum(Me.hwnd)
    Case mconMenu_Help_web_WebMail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case Else '����
        Call vsfPrint_Custom(Control)
    End Select
End Sub

Private Function CheckPrivs(ByRef intMethod As Integer) As Boolean
    '���ܣ��ж��Ƿ���ж�Ӧ������Ȩ��
    '����ֵ��ture-���ж�Ӧ����Ȩ�ޣ�false-�����ж�Ӧ����Ȩ��
    '���Σ����ز������� 0-�ۼ۵��ۣ�1-�ɱ��۵��� 2-һ�����
    With vsfList
        If .TextMatrix(vsfList.Row, menuListCol.��������) = "���ۼ۵���" Then
            intMethod = 0
            If InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") > 0 Then CheckPrivs = True
        ElseIf .TextMatrix(vsfList.Row, menuListCol.��������) = "���ɱ��۵���" Then
            intMethod = 1
            If InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") > 0 Then CheckPrivs = True
        ElseIf .TextMatrix(vsfList.Row, menuListCol.��������) = "�ۼ۳ɱ���һ�����" Then
            intMethod = 2
            If InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") > 0 And InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") > 0 Then CheckPrivs = True
        End If
    End With
End Function

Private Sub vsfPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ӡ�Զ��屨��NO=���ۻ��ܵ���
    Dim strNo As String
    
    With vsfList
        If .Rows < 2 Then Exit Sub
        strNo = .TextMatrix(.Row, menuListCol.���ۺ�)
    End With
    
    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "NO=" & strNo)
End Sub


Private Sub danJuprint(ByVal intType As Integer)
    '���ݴ�ӡ
    Dim strTemp As String

    On Error GoTo ErrHandle
    With vsfList
        If .TextMatrix(.Row, menuListCol.���ۺ�) = "" Then Exit Sub
        strTemp = .TextMatrix(.Row, menuListCol.���ۺ�)
    End With
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1726_1", Me, "���ۺ�=" & strTemp, "���㵥λ=" & mintUnit, intType)

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Freview()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

 Private Sub printbill(ByVal intPrintMode As Byte)
    '-------------------------------------------------------------------------------------
    '����:��ӡ
    '����:intPrintMode-1-��ӡ,2-Ԥ��,3-Excel
    '-------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.NO)) = "" Then Exit Sub
    objPrint.Title.Text = "���ĵ���֪ͨ��"
    
    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & vsfList.TextMatrix(vsfList.Row, menuListCol.˵��)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(vsfList.TextMatrix(vsfList.Row, menuListCol.ִ������), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, menuListCol.������)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.vsfList
    objPrint.PageFooter = 2
     
    If intPrintMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, intPrintMode
    End If
    Set objPrint = Nothing
End Sub

Private Sub billExcel()
    '�����Excel
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfPrice Then
        vsfPrice.Redraw = flexRDNone
        subExcel 3
        vsfPrice.Redraw = flexRDDirect
        vsfPrice.Col = 0
        vsfPrice.ColSel = vsfPrice.Cols - 1
    End If
End Sub
Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String

    If Format(mdaStart, "yyyy-mm-dd") = "1901-01-01" And Format(mdaVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(mdaVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdaVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(mdaVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(mdaStart, "yyyy��MM��dd��") & "��" & Format(mdaEnd, "yyyy��MM��dd��") & "  ������� " & Format(mdaVerifyStart, "yyyy��MM��dd��") & "��" & Format(mdaVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(mdaStart, "yyyy��MM��dd��") & "��" & Format(mdaEnd, "yyyy��MM��dd��")
    End If

    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True

    objPrint.Title.Text = "���ĵ��۹���"

    objRow.Add "ʱ�䣺" & strRange
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡ����:" & Format(sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow

    If vsfList Is ActiveControl Then
        Set objPrint.Body = vsfList
    ElseIf vsfPrice Is ActiveControl Then
        Set objPrint.Body = vsfPrice
    ElseIf vsfCost Is ActiveControl Then
        Set objPrint.Body = vsfCost
    End If

    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub subExcel(bytMode As Byte)
'����:���������EXCEL
'����:bytMode3 �����EXCEL

    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String


    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True

    objPrint.Title.Text = "���ĵ��۹���"

    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "���ۺ�." & Trim(vsfList.TextMatrix(vsfList.Row, menuListCol.���ۺ�))
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, menuListCol.������) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, menuListCol.��������)

    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = vsfPrice
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub cmdView_Click()
    Dim intMethod As Integer

    If vsfList.Row <= 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, menuListCol.��������) = "���ۼ۵���" Then
        intMethod = 0
    ElseIf vsfList.TextMatrix(vsfList.Row, menuListCol.��������) = "���ɱ��۵���" Then
        intMethod = 1
    ElseIf vsfList.TextMatrix(vsfList.Row, menuListCol.��������) = "�ۼ۳ɱ���һ�����" Then
        intMethod = 2
    End If
    frmStuffPriceCard.ShowMe Me, 2, vsfList.TextMatrix(vsfList.Row, menuListCol.���ۺ�), intMethod
End Sub

Private Sub Form_Load()

    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    mstrPrivs = ";" & GetPrivFunc(glngSys, 1726) & ";"

    Call initJinDu
    Call initCommandBars
    Call initMenu
    Call InitTool
    Call InitTabControl
    Call initVsflexgrid
    Call SetMenuEnable '��Ȩ�������Ʋ˵�
    '����Զ��屨��
    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)

    Call RestoreWinState(Me, App.ProductName, mstrCaption)

    Call getListInfo
    stbThis.Panels(2).Picture = picColor
    mblnLoad = True
End Sub

Private Sub initJinDu()
    '���ܣ���ʼ��������λ����ľ���
    '�ж��Ƿ���ҩ�ⵥλ��ʾ
    '��ȡ���õĵ�λ
    Dim intUnitTemp As Integer

    mintUnit = Val(zldatabase.GetPara("���ĵ�λ", glngSys, 1726, 1))
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    If Me.ScaleHeight / 2 < 2000 Then Exit Sub
    vsfList.Move 0, 900, Me.ScaleWidth, Me.ScaleHeight / 2 - 2000
    picSplit.Left = 50
    picSplit.Top = vsfList.Top + vsfList.Height + 50
    picSplit.Width = Me.ScaleWidth
    cmdView.Move Me.ScaleWidth - cmdView.Width - 500, picSplit.Top + 50

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - picSplit.Top - picSplit.Height - stbThis.Height
    End With
    vsfPrice.Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height
    vsfCost.Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub

Private Sub initVsflexgrid()
    With vsfList
        .Editable = flexEDNone
        .Cols = menuListCol.������
        .Rows = 1
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExSortShowAndMove '������ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        '�����п�
        .ColWidth(menuListCol.���ۺ�) = 1500
        .ColWidth(menuListCol.��������) = 2000
        .ColWidth(menuListCol.������) = 1500
        .ColWidth(menuListCol.��������) = 2000
        .ColWidth(menuListCol.ִ������) = 2000
        .ColWidth(menuListCol.˵��) = 2000
        '���뷽ʽ
        .ColAlignment(menuListCol.���ۺ�) = flexAlignLeftCenter
        .ColAlignment(menuListCol.��������) = flexAlignLeftCenter
        .ColAlignment(menuListCol.������) = flexAlignLeftCenter
        .ColAlignment(menuListCol.��������) = flexAlignLeftCenter
        .ColAlignment(menuListCol.ִ������) = flexAlignLeftCenter
        .ColAlignment(menuListCol.˵��) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, menuListCol.���ۺ�) = "���ۺ�"
        .TextMatrix(0, menuListCol.��������) = "��������"
        .TextMatrix(0, menuListCol.������) = "������"
        .TextMatrix(0, menuListCol.��������) = "��������"
        .TextMatrix(0, menuListCol.ִ������) = "ִ������"
        .TextMatrix(0, menuListCol.˵��) = "˵��"
    End With

    With vsfPrice
        .Editable = flexEDNone
        .Cols = menuPriceCol.������
        .Rows = 1
        .ColHidden(0) = True
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .Cols - 1) = 50 '����Ӵ�
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExSortShowAndMove '������ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        '�����п�
        .ColWidth(menuPriceCol.ԭ��id) = 0
        .ColWidth(menuPriceCol.NO) = 1000
        .ColWidth(menuPriceCol.������Ϣ) = 3500
        .ColWidth(menuPriceCol.���) = 1500
        .ColWidth(menuPriceCol.��λ) = 800
        .ColWidth(menuPriceCol.��λϵ��) = 0
        .ColWidth(menuPriceCol.ԭ��) = 1000
        .ColWidth(menuPriceCol.�ּ�) = 1000
        .ColWidth(menuPriceCol.ִ������) = 0
        .ColWidth(menuPriceCol.������) = 1000
        '���뷽ʽ
        .ColAlignment(menuPriceCol.NO) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.������Ϣ) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.���) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.ԭ��) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.�ּ�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ִ������) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.������) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, menuPriceCol.ԭ��id) = "ԭ��id"
        .TextMatrix(0, menuPriceCol.NO) = "NO"
        .TextMatrix(0, menuPriceCol.������Ϣ) = "������Ϣ"
        .TextMatrix(0, menuPriceCol.���) = "���"
        .TextMatrix(0, menuPriceCol.��λ) = "��λ"
        .TextMatrix(0, menuPriceCol.��λϵ��) = "��λϵ��"
        .TextMatrix(0, menuPriceCol.ԭ��) = "ԭ��"
        .TextMatrix(0, menuPriceCol.�ּ�) = "�ּ�"
        .TextMatrix(0, menuPriceCol.ִ������) = "ִ������"
        .TextMatrix(0, menuPriceCol.������) = "������"
    End With

    With vsfCost
        .Editable = flexEDNone
        .Cols = menuCostCol.������
        .Rows = 1
        .ColHidden(0) = True
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .Cols - 1) = 50
        .RowHeight(0) = mconlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExSortShowAndMove '������ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        '�����п�
        .ColWidth(menuCostCol.NO) = 1000
        .ColWidth(menuCostCol.������Ϣ) = 3500
        .ColWidth(menuCostCol.�ⷿ) = 800
        .ColWidth(menuCostCol.���) = 1500
        .ColWidth(menuCostCol.����) = 1000
        .ColWidth(menuCostCol.����) = 1500
        .ColWidth(menuCostCol.��λ) = 800
        .ColWidth(menuCostCol.ԭ��) = 1000
        .ColWidth(menuCostCol.�ּ�) = 1000
        .ColWidth(menuCostCol.Ч��) = 1500
        .ColWidth(menuCostCol.ִ������) = 0
        .ColWidth(menuCostCol.������) = 1000
        '���뷽ʽ
        .ColAlignment(menuCostCol.NO) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.������Ϣ) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.�ⷿ) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.���) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.����) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.����) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuCostCol.ԭ��) = flexAlignRightCenter
        .ColAlignment(menuCostCol.�ּ�) = flexAlignRightCenter
        .ColAlignment(menuCostCol.Ч��) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.ִ������) = flexAlignLeftCenter
        .ColAlignment(menuCostCol.������) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

        .TextMatrix(0, menuCostCol.NO) = "NO"
        .TextMatrix(0, menuCostCol.������Ϣ) = "������Ϣ"
        .TextMatrix(0, menuCostCol.�ⷿ) = "�ⷿ"
        .TextMatrix(0, menuCostCol.���) = "���"
        .TextMatrix(0, menuCostCol.����) = "����"
        .TextMatrix(0, menuCostCol.����) = "����"
        .TextMatrix(0, menuCostCol.��λ) = "��λ"
        .TextMatrix(0, menuCostCol.ԭ��) = "ԭ�ɱ���"
        .TextMatrix(0, menuCostCol.�ּ�) = "�ֳɱ���"
        .TextMatrix(0, menuCostCol.Ч��) = "Ч��"
        .TextMatrix(0, menuCostCol.ִ������) = "ִ������"
        .TextMatrix(0, menuCostCol.������) = "������"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrCaption)
    mblnLoad = False
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfList.Height + y <= 800 Then Exit Sub
    If TabCtlDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    cmdView.Move Me.ScaleWidth - cmdView.Width - 500, picSplit.Top + 50
    vsfList.Move 0, 900, Me.ScaleWidth, vsfList.Height + y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - y
    End With
End Sub

Private Sub vsfCost_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfCost
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 300
    End With
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    '�ƶ���һ���ı�ǵ���ǰ�У�
    With vsfList
        .Cell(flexcpText, 0, 0, .Rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 8
        End If
    End With
End Sub

Private Sub SetMenuEnable()
    '�ж�Ȩ�޶Բ˵���Ӱ��
    Dim cbrMenuControl As CommandBarControl
    Dim cbrMenuPop As CommandBarControl

    '���������˵�
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Add, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Add, , True)
    If InStr(1, gstrPrivs, "�Ǽ�") = 0 Or (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If

    '�����޸Ĳ˵�
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
    If InStr(1, gstrPrivs, "�޸�") = 0 Or (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If

    '����ɾ���˵�
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
    Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
    If InStr(1, gstrPrivs, "ɾ��") = 0 Or (InStr(1, ";" & gstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & gstrPrivs & ";", ";�ۼ۵���;") = 0) Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
        If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
    End If
    Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_File_Parameter, , True)
    If InStr(1, gstrPrivs, "��������") = 0 Then
        If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
    End If
End Sub

Private Sub getListInfo()
    '��ȡ���ۻ�����Ϣ
    Dim rsTemp As ADODB.Recordset
    Dim strClass As String '��������
    Dim i As Integer
    Dim dateCurrentDate As Date

    On Error GoTo ErrHandle

    dateCurrentDate = sys.Currentdate
    mintFindDay = Val(zldatabase.GetPara("��ѯ����", glngSys, 1726, 1))
    mdaStart = Format(DateAdd("d", -mintFindDay, dateCurrentDate), "yyyy-MM-dd")
    mdaEnd = CDate(Format(dateCurrentDate, "yyyy-MM-dd") & " 23:59:59")
    
    mdaVerifyStart = "1901-01-01"
    mdaVerifyEnd = "1901-01-01"
    If mSQLCondition.date����ʱ�俪ʼ = "0:00:00" Then
        lblScope.Caption = "���ڷ�Χ��" & Format(mdaStart, "yyyy-mm-dd") & "��" & Format(mdaEnd, "yyyy-mm-dd")
    Else
        lblScope.Caption = "���ڷ�Χ��" & Format(mSQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") & "��" & Format(mSQLCondition.date����ʱ�����, "yyyy-mm-dd")
    End If

    vsfList.Rows = 1
    vsfPrice.Rows = 1
    vsfCost.Rows = 1
    gstrSQL = "select a.���ۺ�, a.����, a.ִ������, a.��������, a.������, a.˵�� from ���ۻ��ܼ�¼ a"

    '�����ڿմ����ǹ���
    If mstrResult <> "" Then
        gstrSQL = gstrSQL + " where " + mstrResult + " and a.����=1 order by a.���ۺ� desc"

    Else 'Ĭ��ֻ��ѯ����������һ���ܵĵ�����Ϣ
        gstrSQL = gstrSQL + " where " + " a.�������� between to_date('" & mdaStart & "', 'yyyy-mm-dd hh24:mi:ss') and to_date('" & mdaEnd & "', 'yyyy-mm-dd hh24:mi:ss') and a.����=1 order by a.���ۺ� desc"
    End If

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ���ۻ��ܼ�¼", mSQLCondition.date����ʱ�俪ʼ, mSQLCondition.date����ʱ�����, mSQLCondition.dateִ��ʱ�俪ʼ, mSQLCondition.dateִ��ʱ�����)

    If rsTemp.RecordCount = 0 Then Exit Sub
    rsTemp.MoveFirst
    For i = 0 To rsTemp.RecordCount - 1
        With vsfList
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = mconlngRowHeight
            .TextMatrix(.Rows - 1, menuListCol.���ۺ�) = rsTemp!���ۺ�
            If rsTemp!���� = 0 Then
                strClass = "���ۼ۵���"
            ElseIf rsTemp!���� = 1 Then
                strClass = "���ɱ��۵���"
            ElseIf rsTemp!���� = 2 Then
                strClass = "�ۼ۳ɱ���һ�����"
            End If
            .TextMatrix(.Rows - 1, menuListCol.��������) = strClass
            .TextMatrix(.Rows - 1, menuListCol.������) = rsTemp!������
            .TextMatrix(.Rows - 1, menuListCol.��������) = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, menuListCol.ִ������) = Format(rsTemp!ִ������, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, menuListCol.˵��) = IIf(IsNull(rsTemp!˵��), "", rsTemp!˵��)

            If rsTemp!ִ������ > dateCurrentDate Then 'δִ�е��ú�ɫ��ʾ
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            End If
            rsTemp.MoveNext
        End With
    Next

    If vsfList.TextMatrix(1, menuListCol.���ۺ�) <> "" Then
        vsfList.Row = 1
        vsfList.Col = 1
        Call getPriceInfo
        Call getCostInfo
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub getCostInfo()
    '��ȡ�ɱ��۵�����Ϣ
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim db��װϵ�� As Double
    Dim strUnit As String

    On Error GoTo ErrHandle
    gstrSQL = " Select B.NO, I.ID As ҩƷid, '[' || I.���� || ']' || I.���� ||  ' ' || I.���� As ҩƷ, P.���� As �ⷿ,A.����,A.Ч��,A.����,i.���, " & _
            " I.���㵥λ As ��λ, s.��װ��λ,s.����ϵ�� , A.ԭ�ɱ���,A.�³ɱ��� As �ɱ���, A.ִ������, B.ժҪ " & _
            " From ҩƷ�շ���¼ B, �շ���ĿĿ¼ I, �������� S, ���ű� P, �ɱ��۵�����Ϣ A " & _
            " Where A.�շ�id = B.ID(+) And A.ҩƷid = I.ID And " & _
            " I.ID = S.����id And A.�ⷿid = P.ID(+) And a.���ۻ��ܺ�=[1] " & _
            " Order By ҩƷ, ִ������ Desc, NO Desc"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ�ɱ��۵���", vsfList.TextMatrix(vsfList.Row, menuListCol.���ۺ�))
    vsfCost.Rows = 1
    If rsTemp.RecordCount = 0 Then Exit Sub

    With vsfCost
        For i = 0 To rsTemp.RecordCount - 1
            .Rows = .Rows + 1
            Select Case mintUnit
                Case 0  'ɢװ
                    db��װϵ�� = 1
                    strUnit = rsTemp!��λ
                Case 1  '��װ
                    db��װϵ�� = rsTemp!����ϵ��
                    strUnit = rsTemp!��װ��λ
            End Select
            .RowHeight(.Rows - 1) = mconlngRowHeight
            .TextMatrix(.Rows - 1, menuCostCol.NO) = IIf(IsNull(rsTemp!NO), "", rsTemp!NO)
            .TextMatrix(.Rows - 1, menuCostCol.������Ϣ) = rsTemp!ҩƷ
            .TextMatrix(.Rows - 1, menuCostCol.�ⷿ) = IIf(IsNull(rsTemp!�ⷿ), "", rsTemp!�ⷿ)
            .TextMatrix(.Rows - 1, menuCostCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.Rows - 1, menuCostCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, menuCostCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, menuCostCol.��λ) = strUnit
            .TextMatrix(.Rows - 1, menuCostCol.ԭ��) = Format(IIf(IsNull(rsTemp!ԭ�ɱ���), 0, rsTemp!ԭ�ɱ���) * db��װϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(.Rows - 1, menuCostCol.�ּ�) = Format(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * db��װϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(.Rows - 1, menuCostCol.Ч��) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, menuCostCol.ִ������) = Format(IIf(IsNull(rsTemp!ִ������), "", rsTemp!ִ������), "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, menuCostCol.������) = vsfList.TextMatrix(vsfList.Row, menuListCol.������)
            rsTemp.MoveNext
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub getPriceInfo()
    '��ȡ�ۼ۵�����Ϣ
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim db��װϵ�� As Double
    Dim strUnit As String

    On Error GoTo ErrHandle

    gstrSQL = "Select p.Id, i.���,Decode(i.�Ƿ���, 1, 'ʱ��', '����') As ҩ������, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & _
                       " Nvl(s.ָ�������, 0) As ָ������, Nvl(i.���ηѱ�, 0) As ���ηѱ�," & _
                       " Decode(Sign(p.ִ������ - Sysdate), 1, 1, Decode(Sign(p.��ֹ���� - Sysdate), -1, -1, 0)) As ִ�����," & _
                       " '[' || i.���� || ']' || i.���� || ' '  || i.���� As ҩƷ, i.���㵥λ As ��λ,s.��װ��λ,Nvl(s.����ϵ��, 1) ����ϵ��," & _
                       " p.ԭ��, p.�ּ� , u.���� As ������Ŀ, p.����˵��, To_Char(p.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������,p.������," & _
                       " i.Id ҩƷid, p.No ����no" & _
                " From �շѼ�Ŀ P, ������Ŀ U, �շ���ĿĿ¼ I, �������� S" & _
                " Where p.�շ�ϸĿid = i.Id And p.������Ŀid = u.Id And i.Id = s.����id And p.���ۻ��ܺ� = [1] " & _
                GetPriceClassString("P") & _
                " Order By i.����, p.ִ������ Desc"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, mstrCaption, vsfList.TextMatrix(vsfList.Row, menuListCol.���ۺ�))
    vsfPrice.Rows = 1
    If rsTemp.RecordCount = 0 Then Exit Sub

    With vsfPrice
        For i = 0 To rsTemp.RecordCount - 1
            .Rows = .Rows + 1
            .RowHeight(.Rows - 1) = mconlngRowHeight

            Select Case mintUnit
                Case 0  'ɢװ
                    db��װϵ�� = 1
                    strUnit = rsTemp!��λ
                Case 1  '��װ
                    db��װϵ�� = rsTemp!����ϵ��
                    strUnit = rsTemp!��װ��λ
            End Select

            .TextMatrix(.Rows - 1, menuPriceCol.ԭ��id) = rsTemp!Id
            .TextMatrix(.Rows - 1, menuPriceCol.NO) = rsTemp!����no
            .TextMatrix(.Rows - 1, menuPriceCol.������Ϣ) = rsTemp!ҩƷ
            .TextMatrix(.Rows - 1, menuPriceCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.Rows - 1, menuPriceCol.��λ) = strUnit
            .TextMatrix(.Rows - 1, menuPriceCol.��λϵ��) = IIf(mintUnit = 0, 1, rsTemp!����ϵ��)
            .TextMatrix(.Rows - 1, menuPriceCol.ԭ��) = Format(IIf(IsNull(rsTemp!ԭ��), 0, rsTemp!ԭ��) * db��װϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(.Rows - 1, menuPriceCol.�ּ�) = Format(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * db��װϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(.Rows - 1, menuPriceCol.ִ������) = Format(rsTemp!ִ������, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(.Rows - 1, menuPriceCol.������) = rsTemp!������
            rsTemp.MoveNext
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_DblClick()
    Call cmdView_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim cbrMenuControl As CommandBarControl
    Dim cbrMenuPop As CommandBarControl

    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect

        If .TextMatrix(.Row, menuListCol.��������) = "���ۼ۵���" Then
            TabCtlDetails.Item(1).Visible = False
            TabCtlDetails.Item(0).Visible = True
            TabCtlDetails.Item(0).Selected = True
        ElseIf .TextMatrix(.Row, menuListCol.��������) = "���ɱ��۵���" Then
            TabCtlDetails.Item(1).Visible = True
            TabCtlDetails.Item(0).Visible = False
            TabCtlDetails.Item(1).Selected = True
        Else
            TabCtlDetails.Item(1).Visible = True
            TabCtlDetails.Item(0).Visible = True
            TabCtlDetails.Item(0).Selected = True
        End If
        If .TextMatrix(.Row, menuListCol.���ۺ�) <> "" And .Row > 0 Then
            'ִ�����ڴ���ϵͳ��ǰ���ڲ����޸ĵ��۵�
            Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
            Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Update, , True)
            If CDate(.TextMatrix(.Row, menuListCol.ִ������)) < sys.Currentdate() Then
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
            Else
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = True
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = True
            End If

            'ִ�����ڴ���ϵͳ��ǰ���ڲ���ɾ�����۵�
            Set cbrMenuPop = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
            Set cbrMenuControl = Me.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Delete, , True)
            If CDate(.TextMatrix(.Row, menuListCol.ִ������)) < sys.Currentdate() Then
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = False
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = False
            Else
                If Not cbrMenuPop Is Nothing Then cbrMenuPop.Enabled = True
                If Not cbrMenuControl Is Nothing Then cbrMenuControl.Enabled = True
            End If
            Call SetMenuEnable

            Call getPriceInfo
            Call getCostInfo
        End If
        If mblnLoad = True Then
'            vsfList.SetFocus
        End If
    End With
End Sub

Private Sub deleteNotExecutePirce()
    '���δִ�м۸�
    Dim rsTemp As ADODB.Recordset
    Dim arrSQL As Variant
    Dim i As Integer

    On Error GoTo ErrHandle
    arrSQL = Array()
    With vsfList
        If .TextMatrix(.Row, menuListCol.���ۺ�) <> "" Then
            If MsgBox("ȷ��ɾ���������۵��ݣ�", vbYesNo, gstrSysName) = vbNo Then Exit Sub

            gstrSQL = "select �շ�ϸĿid as id from �շѼ�Ŀ where ���ۻ��ܺ�=[1]" & GetPriceClassString("") & _
                        " union " & _
                      " select ҩƷid as id from �ɱ��۵�����Ϣ where ���ۻ��ܺ�=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����۸�", .TextMatrix(.Row, menuListCol.���ۺ�))
            If rsTemp.RecordCount = 0 Then
                MsgBox "�õ�����Ϣ�Ѿ�����ɾ����", vbInformation, gstrSysName
                Exit Sub
            Else
                rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                    gstrSQL = "Zl_ɾ��δִ�м۸�_Delete(" & rsTemp!Id & "," & 0 & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    rsTemp.MoveNext
                Loop
                
                gcnOracle.BeginTrans
                For i = 0 To UBound(arrSQL)
                    Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), "ɾ�����ۼ�¼")
                Next
                gcnOracle.CommitTrans
                
            End If
        End If
    End With
    'ɾ����ˢ�½���
    Call getListInfo

    Exit Sub
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfPrice_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfPrice
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 300
    End With
End Sub

Private Sub vsfPrice_GotFocus()
    Call SetGridFocus(vsfPrice, True)
End Sub
Private Sub vsfPrice_LostFocus()
    Call SetGridFocus(vsfPrice, False)
End Sub

Private Sub vsfcost_GotFocus()
    Call SetGridFocus(vsfCost, True)
End Sub
Private Sub vsfcost_LostFocus()
    Call SetGridFocus(vsfCost, False)
End Sub
