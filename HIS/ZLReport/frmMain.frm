VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "�������"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRPT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1680
      ScaleHeight     =   1095
      ScaleWidth      =   1935
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1935
      Begin XtremeSuiteControls.TabControl tbcRPT 
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   1296
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   7080
      ScaleHeight     =   4455
      ScaleWidth      =   3015
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
      Begin VB.PictureBox picGroup_S 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   120
         MousePointer    =   7  'Size N S
         ScaleHeight     =   60
         ScaleWidth      =   2535
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3120
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfGroup 
         DragIcon        =   "frmMain.frx":058A
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989546359
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
         ExplorerBar     =   0
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
         VirtualData     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfGroupDetail 
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989543184
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
         ExplorerBar     =   0
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
         VirtualData     =   -1  'True
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
      Begin VB.Label lblGroupDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������Ա"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   900
      End
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3840
      ScaleHeight     =   1095
      ScaleWidth      =   3015
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfReport 
         DragIcon        =   "frmMain.frx":0CF4
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2535
         _cx             =   1989546359
         _cy             =   1989543184
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
         ExplorerBar     =   0
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
         VirtualData     =   -1  'True
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
   End
   Begin VB.PictureBox picClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   240
      ScaleHeight     =   3735
      ScaleWidth      =   1215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1215
      Begin XtremeReportControl.ReportControl rptClass 
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   855
         _Version        =   589884
         _ExtentX        =   1508
         _ExtentY        =   1508
         _StockProps     =   0
         MultipleSelection=   0   'False
         ShowHeader      =   0   'False
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3840
      ScaleHeight     =   285
      ScaleWidth      =   1905
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1905
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Enter�������ң�F3������������"
         Top             =   0
         Width           =   1875
      End
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   6570
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":145E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
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
   Begin MSComDlg.CommonDialog cdg 
      Left            =   1320
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1920
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CF2
            Key             =   "rpt"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":208C
            Key             =   "rpt_ena"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2426
            Key             =   "rpt_dis"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27C0
            Key             =   "grp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B5A
            Key             =   "grp_ena"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EF4
            Key             =   "grp_dis"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":328E
      Left            =   480
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enuMenus
    �ļ� = 1
        �������� = 181
        �������� = 121
        ���뱨�� = 122
        ����ȫ�� = 123
        ����ȫ�� = 124
        �˳� = 2613
    �༭ = 2
        ���������� = 3051
        �޸ı����� = 3053
        ɾ�������� = 3054
        ���������� = 6861
        �޸ı����� = 6862
'        ���������� = 6881
        ɾ�������� = 6863
        �������� = 3001
        �޸ı��� = 3003
'        ���뱨�� = 3052
'        �Ƴ����� = 3036
        ɾ������ = 3004
        ��Ʊ��� = 4113
        ������ = 3551
        ִ�б��� = 3010
        �������� = 8106
        ����ͣ�� = 8099
    ���� = 5
        ������ = 741
            ������̨�˵� = 100501
            ��ģ���ڲ˵� = 100502
        ȡ������ = 742
            �ӵ���̨�˵� = 100511
            ��ģ���ڲ˵� = 100512
        ���ܼ�� = 100521
        �����ʷ����Դ = 100522
        ����������־ = 100523
    �鿴 = 7
        ������ = 701
            ��׼��ť = 702
            �ı���ǩ = 703
            ��ͼ�� = 704
        ״̬�� = 711
        �����С = 721
            С���� = 722
            ������ = 723
        ���� = 721
        ˢ�� = 791
        ��ʾ���з����¼� = 751
        ����ʾͣ��״̬ = 7510
        ��ʾ�������� = 752
        ��ʾ�ӱ��� = 753
    ���� = 9
        �������� = 901
        WEB�ϵ����� = 911
            ������ҳ = 912
            ������̳ = 913
            ���ͷ��� = 914
        ���� = 991
    ���� = 10
        ѡ��ϵͳ��ǩ = 1001
        ѡ��ϵͳ�ؼ� = 1002
        ���ұ����ǩ = 1003
        ���ұ���ؼ� = 1004
        TabRPT_1 = 1011
        TabRPT_2 = 1012
End Enum

Private Const MSTR_REPORT_COLS = _
    "���,,3,2000|ID,,0,0,n|����,,3,2500|˵��,,3,3000|����ID,,0,0,n|�޸�ʱ��,,3,2000,DT|����ʱ��,,3,2000,DT|ϵͳ,,0,0|" & _
    "���ִ��ʱ��,,3,2000,DT|���ִ����,,3,1000|����,,3,1000|����,,3,1000|�������,,3,1500|���ܼ����,,3,2000|" & _
    "����������,,3,2000|������������,,3,2000|����ID,,0,0,n|ͣ��,,0,0,n"
Private Const MSTR_GROUP_COLS = _
    "���,,3,2000|����,,3,2500|˵��,,3,6000|�������,,3,1500|ID,,0,0,n|����ʱ��,,3,2000,DT|����ID,,0,0,n|����ID,,0,0,n|" & _
    "ͣ��,,0,0,n"
Private Const MSTR_GROUPDETAIL_COLS = _
    "���,,3,2000|ID,,0,0,n|����,,3,2500|˵��,,3,3000|����ID,,0,0,n|�޸�ʱ��,,3,2000,DT|����ʱ��,,3,2000,DT|ϵͳ,,0,0|" & _
    "���ִ��ʱ��,,3,2000,DT|���ִ����,,3,1000|����,,3,1000|����,,3,1000|������������,,3,2000|ͣ��,,0,0,n"

Private WithEvents mobjClass As clsReportControlEx
Attribute mobjClass.VB_VarHelpID = -1
Private WithEvents mobjReport As clsVSFlexGridEx
Attribute mobjReport.VB_VarHelpID = -1
Private WithEvents mobjGroup As clsVSFlexGridEx
Attribute mobjGroup.VB_VarHelpID = -1
Private WithEvents mobjSub As clsVSFlexGridEx
Attribute mobjSub.VB_VarHelpID = -1

Private mbytFontSize As Byte                                            '1-�����壻0-С����
Private mbytReportGroup As Byte                                         '1-��ʾ��������0-��ʾ�ӱ���
Private mblnDisplayChild As Boolean                                     'True-��ʾ�����ӽ�����Ŀ��False-��ʾ��ǰ������Ŀ
Private mblnDisable As Boolean                                          'True-����ͣ��
Private mblnMemory As Boolean                                           '���Ի�����
Private mblnAllowPublish As Boolean                                     'True-��������ģ�飻False-������
Private mblnReportControlFocus As Boolean                               'ReportControl��������Ӧ���������
Private mcbpReport As XtremeCommandBars.CommandBarPopup                 '����ĵ����˵�
Private mobjFindPreviors As Object                                      '�ϴβ��ҵĶ���
Private mlngFindRow As Long                                             '�ϴβ��Ҷ�λ���к�
Private mintFind As Integer
Private mblnEnter As Boolean

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As Object
    Dim i As Integer
    Dim lngID As Long
    Dim blnTemp As Boolean
    Dim strPath As String, strTemp As String
    
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.id
    Case enuMenus.ִ�б���
        Call GetVsfControl(lngID, blnTemp)
        If lngID > 0 Then
            If blnTemp Then
                '������
                ''��������ӱ����Ȩ��
                For i = 1 To vsfGroupDetail.Rows - 1
                    If mdlPublic.CheckReportPriv(lngID, True) = False Then
                        MsgBox mdlPublic.FormatString("��û��Ȩ�޲�ѯ����[1]����ĳЩ����Դ�Ķ���" _
                                            , Val(vsfGroupDetail.TextMatrix(i, vsfGroupDetail.ColIndex("ID")))) _
                            , vbInformation, App.Title
                        Exit Sub
                    End If
                Next
            Else
                '����
                If mdlPublic.CheckReportPriv(lngID) = False Then
                    MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���", vbInformation, App.Title
                    Exit Sub
                End If
            End If
            
            'ִ��
            If blnTemp Then
                '������
                Set gobjReport = Nothing
                glngGroup = lngID
            Else
                '����
                If mdlPublic.CheckPass(lngID) = False Then
                    MsgBox "�������ݴ��󣬲���ִ�иñ���", vbInformation, App.Title
                    Exit Sub
                End If
                
                glngGroup = 0
                Set gobjReport = Nothing
                Set gobjReport = mdlPublic.ReadReport(lngID)
            End If
            
            'ʹ��ȱʡ����
            garrPars = Array()
            If Not mdlPublic.ShowReport(Me) Then MsgBox "�����ʧ�ܣ�", vbInformation, App.Title
        End If
    Case enuMenus.��������
        If frmReportPara.ShowMe(Me) Then
            '���²���
            Call mdlPublic.InitPar
        End If
    Case enuMenus.���ܼ��
        Call CheckSQLPlanEx
    Case enuMenus.��������, enuMenus.����ȫ��
        Call Export(Control.id)
    Case enuMenus.���뱨��, enuMenus.����ȫ��
        Call Import(Control.id)
    Case enuMenus.�˳�
        Unload Me
    Case enuMenus.����������, enuMenus.����������, enuMenus.��������
        mblnReportControlFocus = enuMenus.���������� = Control.id
        Call NewEx
    Case enuMenus.�޸ı�����, enuMenus.�޸ı�����, enuMenus.�޸ı���
        mblnReportControlFocus = enuMenus.�޸ı����� = Control.id
        Call Modify
    Case enuMenus.ɾ��������, enuMenus.ɾ��������, enuMenus.ɾ������
        mblnReportControlFocus = enuMenus.ɾ�������� = Control.id
        Call Delete(Control.id)
    Case enuMenus.��Ʊ���
        Call Design
    Case enuMenus.��������
        Call StateSwitch(Control.id, True)
    Case enuMenus.����ͣ��
        Call StateSwitch(Control.id)
    Case enuMenus.�����ʷ����Դ
        frmClearHistory.Show vbModal, Me
    Case enuMenus.������
        Call Guide
    Case enuMenus.������̨�˵�
        Call ReportGrantToNavigator
    Case enuMenus.��ģ���ڲ˵�
        Call ReportGrantToModule
    Case enuMenus.�ӵ���̨�˵�
        Call ReportRevokeFromNavigator
    Case enuMenus.��ģ���ڲ˵�
        Call ReportRevokeFromModule
    Case enuMenus.����
        If txtFind.Visible And txtFind.Enabled Then
            txtFind.SetFocus
        End If
    Case enuMenus.���ұ���ؼ�
        Call Find(txtFind.Text)     '������һ��ƥ����
    Case enuMenus.����������־
        Call ShowRunLog
    Case enuMenus.��׼��ť
        cbsMain(Val("2-������")).Visible = Not cbsMain(Val("2-������")).Visible
        cbsMain.RecalcLayout
    Case enuMenus.�ı���ǩ
        For Each objControl In cbsMain(Val("2-������")).Controls
            If UCase(TypeName(objControl)) = UCase("ICommandBarButton") _
                Or UCase(TypeName(objControl)) = UCase("ICommandBarPopup") Then
                objControl.Style = IIF(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            End If
        Next
        cbsMain.RecalcLayout
    Case enuMenus.��ͼ��
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
    Case enuMenus.С����
        If mbytFontSize <> 0 Then Call SetControlFontSize(0)
        mbytFontSize = 0
    Case enuMenus.������
        If mbytFontSize <> 1 Then Call SetControlFontSize(1)
    Case enuMenus.״̬��
        staMain.Visible = Not Control.Checked
        cbsMain.RecalcLayout
    Case enuMenus.ˢ��
        rptClass.Tag = ""
        Call RefreshEx
    Case enuMenus.��ʾ���з����¼�
        mblnDisplayChild = Not mblnDisplayChild
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.����ʾͣ��״̬
        mblnDisable = Not mblnDisable
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.��ʾ��������
        mbytReportGroup = 0
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.��ʾ�ӱ���
        mbytReportGroup = 1
        rptClass.Tag = ""
        Call rptClass_SelectionChanged
    Case enuMenus.��������
        Call mdlPublic.ShowHelpRpt(Me.hwnd, "main", 0)
    Case enuMenus.������ҳ
        Call mdlPublic.zlHomePage(Me.hwnd)
    Case enuMenus.������̳
        Call mdlPublic.zlWebForum(Me.hwnd)
    Case enuMenus.���ͷ���
        Call mdlPublic.zlMailTo(Me.hwnd)
    Case enuMenus.����
        Call mdlPublic.ShowAbout(Me)
    Case enuMenus.ѡ��ϵͳ�ؼ�
        Call SelectedSysComboBox(Control)
    Case enuMenus.TabRPT_1, enuMenus.TabRPT_2
        tbcRPT.Item(Control.id - enuMenus.TabRPT_1).Selected = True
    End Select
    mblnReportControlFocus = False
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If staMain.Visible Then
        Bottom = staMain.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnPublication As Boolean
    
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.id
    Case enuMenus.ִ�б���
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        Select Case UCase(Me.ActiveControl.name)
        Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
            If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
                Control.Enabled = vsfReport.Row > 0
            Else
                Control.Enabled = vsfGroup.Row > 0 Or vsfGroupDetail.Row > 0
            End If
        Case Else
            Control.Enabled = False
        End Select
    Case enuMenus.����������
        Control.Enabled = mblnReportControlFocus
    Case enuMenus.�޸ı�����, enuMenus.ɾ��������
        Control.Enabled = mblnReportControlFocus And rptClass.SelectedRows.count > 0
        If Control.Enabled Then
            Control.Enabled = Nvl(rptClass.FocusedRow.Record(mobjClass.GetColIndex("����")).Value) <> "����"
        End If
    Case enuMenus.����������
        Control.Enabled = tbcRPT.Selected.Index = Val("1-������ҳ��") And glngSys = 0
    Case enuMenus.ɾ��������
        If Not Me.ActiveControl Is Nothing Then
            Control.Enabled = UCase(Me.ActiveControl.name) = "VSFGROUP" And glngSys = 0
            If Control.Enabled Then
                Control.Enabled = Me.ActiveControl.Rows > 1
            End If
        End If
    Case enuMenus.�޸ı�����
        If Not Me.ActiveControl Is Nothing Then
            Control.Enabled = UCase(Me.ActiveControl.name) = "VSFGROUP"
            If Control.Enabled Then
                Control.Enabled = Me.ActiveControl.Rows > 1
            End If
        End If
    Case enuMenus.��������
        Control.Enabled = glngSys = 0
    Case enuMenus.�޸ı���
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0
            ElseIf UCase(Me.ActiveControl.name) = "VSFGROUPDETAIL" Then
                Control.Enabled = vsfGroupDetail.Row > 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.ɾ������
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0 And glngSys = 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.��Ʊ���
        If Not Me.ActiveControl Is Nothing Then
            If UCase(Me.ActiveControl.name) = "VSFREPORT" Then
                Control.Enabled = vsfReport.Row > 0
            ElseIf UCase(Me.ActiveControl.name) = "VSFGROUPDETAIL" Then
                Control.Enabled = vsfGroupDetail.Row > 0
            Else
                Control.Enabled = False
            End If
        End If
    Case enuMenus.��������, enuMenus.����ͣ��
        If Not Me.ActiveControl Is Nothing Then
            Select Case UCase(ActiveControl.name)
            Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
                blnPublication = ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("����ʱ��")) <> "" _
                                And glngSys = 0
                If blnPublication Then
                    If Control.id = enuMenus.�������� Then
                        blnPublication = Val(ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("ͣ��"))) = 1
                    Else
                        blnPublication = Val(ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("ͣ��"))) <> 1
                    End If
                End If
            Case Else
                blnPublication = False
            End Select
            Control.Enabled = blnPublication
        End If
    Case enuMenus.���ܼ��
        Control.Enabled = tbcRPT.Selected.Index = Val("0-����ҳ��")
    Case enuMenus.��׼��ť
        Control.Checked = cbsMain(2).Visible
    Case enuMenus.�ı���ǩ
        Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption _
                        Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
    Case enuMenus.��ͼ��
        Control.Checked = cbsMain.Options.LargeIcons
    Case enuMenus.С����
        Control.IconId = IIF(mbytFontSize = 0, 90004, 90003)
    Case enuMenus.������
        Control.IconId = IIF(mbytFontSize = 1, 90004, 90003)
    Case enuMenus.״̬��
        Control.Checked = staMain.Visible
    Case enuMenus.��ʾ���з����¼�
        Control.Checked = mblnDisplayChild
    Case enuMenus.����ʾͣ��״̬
        Control.Checked = mblnDisable
    Case enuMenus.��ʾ��������
        Control.IconId = IIF(mbytReportGroup = 0, 90004, 90003)
    Case enuMenus.��ʾ�ӱ���
        Control.IconId = IIF(mbytReportGroup = 1, 90004, 90003)
    Case enuMenus.����������־
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        Select Case UCase(Me.ActiveControl.name)
        Case "VSFREPORT", "VSFGROUPDETAIL"
            If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
                Control.Enabled = vsfReport.Row > 0
            Else
                Control.Enabled = vsfGroupDetail.Row > 0
            End If
        Case Else
            Control.Enabled = False
        End Select
    Case enuMenus.������̨�˵�, enuMenus.��ģ���ڲ˵�, enuMenus.�ӵ���̨�˵�, enuMenus.��ģ���ڲ˵�
        If Me.ActiveControl Is Nothing Then
            Control.Enabled = False
            Exit Sub
        End If
        If Me.ActiveControl.name = "" Then
            Control.Enabled = False
            Exit Sub
        End If
        
        If glngSys = 0 Then
            Select Case UCase(Me.ActiveControl.name)
            Case "VSFREPORT", "VSFGROUP", "VSFGROUPDETAIL"
                If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
                    Control.Enabled = vsfReport.Row > 0
                Else
                    Control.Enabled = vsfGroup.Row > 0 Or vsfGroupDetail.Row > 0
                End If
                If Control.Enabled Then
                    If Control.id = enuMenus.�ӵ���̨�˵� Or Control.id = enuMenus.��ģ���ڲ˵� Then
                        Control.Enabled = ActiveControl.TextMatrix(ActiveControl.Row, ActiveControl.ColIndex("����ʱ��")) <> ""
                        If Control.id = enuMenus.��ģ���ڲ˵� Then
                            '�����鲻��������ģ�飬Ҳ��û��ȡ������
                            Control.Enabled = Control.Enabled And UCase(Me.ActiveControl.name) <> "VSFGROUP"
                        End If
                    ElseIf Control.id = enuMenus.��ģ���ڲ˵� Then
                        '�����鲻��������ģ��
                        Control.Enabled = UCase(Me.ActiveControl.name) <> "VSFGROUP"
                    End If
                End If
            Case Else
                Control.Enabled = False
            End Select
        Else
            Control.Enabled = False
        End If
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case 1
        Item.handle = picClass.hwnd
    Case 2
        Item.handle = picRPT.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnEnter Then
        Call SetControlFontSize(mbytFontSize)       '�����С
        mblnEnter = False
    End If
End Sub

Private Sub Form_Load()
    Dim objControl As Object
    Dim strPane As String, strRegPath As String
    Dim lngTmp As Long
    
    mblnEnter = False
    mblnReportControlFocus = False
    strRegPath = mdlPublic.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.name)

    '��ȡ����ֵ
    mblnMemory = mdlPublic.GetMemoryParam()
    mblnDisplayChild = Val(GetSetting("ZLSOFT", strRegPath, "��ʾ���з����¼�")) = 1
    mblnDisable = Val(GetSetting("ZLSOFT", strRegPath, "����ʾͣ��״̬")) = 1
    mbytReportGroup = Val(GetSetting("ZLSOFT", strRegPath, "��ʾ�������"))
    mbytFontSize = Val(GetSetting("ZLSOFT", strRegPath, "�����С"))
    strPane = GetSetting("ZLSOFT", strRegPath, "����")
    
    Call InitOther
    Call InitCommandBars
    Call InitDockPane
    Call InitTabControl
    Call InitReportControl
    Call InitVSF
    
    Call FillData(Val("5-cboSystem"))
    Call FillData(Val("1-rptClass"), True)
    If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
        Call FillData(Val("2-vsfReport"), True)
    Else
        Call FillData(Val("3-vsfGroup"), True)
        Call FillData(Val("4-vsfGroupDetial"), True)
    End If
    
    '�ָ��ϴν���
    If mblnMemory Then
        mdlPublic.RestoreWinState Me, App.ProductName

        'DockingPane
        If strPane <> "" Then
            On Error Resume Next
            dkpMain.LoadStateFromString strPane
            If Err.Number <> 0 Then
                MsgBox Err.Description, vbCritical, App.Title
            End If
            On Error GoTo 0
        End If
    Else
        Me.WindowState = vbMaximized
    End If
    
    Call VisibleToolButton                      '����Button״̬
    
    mblnEnter = True
End Sub

Private Sub InitCommandBars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbmTmp As CommandBarControlCustom
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        Set .Icons = mdlPublic.GetPubIcons
        .EnableCustomization False
        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
    picGroup_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picGroup.BackColor = picGroup_S.BackColor
    lblGroupDetail.BackColor = picGroup_S.BackColor
    
    '�ļ�
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�ļ�, "�ļ�(&F)", -1, False)
    With cbpTmp
        .id = enuMenus.�ļ�
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.���뱨��, "���뱨��")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����ȫ��, "����ȫ��")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����ȫ��, "����ȫ��")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�"): cbcTmp.BeginGroup = True
    End With
    
    '�༭
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�༭, "�༭(&E)", -1, False)
    With cbpTmp
        .id = enuMenus.�༭
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����������, "�����������(&N)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�޸ı�����, "�޸ı������(&M)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ɾ��������, "ɾ���������(&D)")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����������, "����������(&W)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�޸ı�����, "�޸ı�����(&M)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ɾ��������, "ɾ��������(&D)")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�޸ı���, "�޸ı���")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ɾ������, "ɾ������")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��Ʊ���, "��Ʊ���"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.������, "������(&G)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ִ�б���, "ִ�б���")
        
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "����(&S)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����ͣ��, "ͣ��(&T)")
    End With
    
    '����
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.����, "����(&T)", -1, False)
    With cbpTmp
        .id = enuMenus.����
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.������, "������(&B)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̨�˵�, "������̨�˵�(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.ȡ������, "ȡ������(&U)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ӵ���̨�˵�, "�ӵ���̨�˵�(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.���ܼ��, "���ܼ��(&V)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�����ʷ����Դ, "�����ʷ����Դ(&C)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����������־, "����������־(&L)")
    End With
    
    '�鿴
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�鿴, "�鿴(&V)", -1, False)
    With cbpTmp
        .id = enuMenus.�鿴
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.������, "������(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��׼��ť, "��׼��ť(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ı���ǩ, "�ı���ǩ(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ͼ��, "��ͼ��(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.״̬��, "״̬��(&S)")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.�����С, "�����С(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.С����, "С����(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������, "������(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����, "����"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ʾ���з����¼�, "��ʾ���з����¼�(&A)"): cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����ʾͣ��״̬, "����ʾͣ��״̬(&P)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ʾ��������, "ֻ��ʾ��������(&R)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ʾ�ӱ���, "ֻ��ʾ�ӱ���(&S)")
    End With
    
    '����
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.����, "����(&H)", -1, False)
    With cbpTmp
        .id = enuMenus.����
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB�ϵ�����, "&WEB�ϵ�����")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������ҳ, "������ҳ(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̳, "������̳(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.���ͷ���, "���ͷ���(&K)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����, "����(&A)"): cbcTmp.BeginGroup = True
    End With
    
    '���幤����
    Set cbrTmp = cbsMain.Add("������", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap

        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.����������, "������")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�޸ı�����, "�޸���")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ɾ��������, "ɾ����")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.����������, "������")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�޸ı�����, "�޸���")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ɾ��������, "ɾ����")

        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��������, "����")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�޸ı���, "�޸�")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ɾ������, "ɾ��")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��Ʊ���, "���"): cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.������, "��")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ִ�б���, "ִ��")
        
        Set cbpTmp = .Controls.Add(xtpControlPopup, enuMenus.������, "������"): cbpTmp.BeginGroup = True
        cbpTmp.IconId = enuMenus.������
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̨�˵�, "������̨�˵�(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
        Set cbpTmp = .Controls.Add(xtpControlPopup, enuMenus.ȡ������, "ȡ������")
        cbpTmp.IconId = enuMenus.ȡ������
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ӵ���̨�˵�, "�ӵ���̨�˵�(&1)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��"): cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��������, "����")
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�"): cbcTmp.BeginGroup = True
        
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuMenus.ѡ��ϵͳ��ǩ, "ϵͳ"): cbcTmp.BeginGroup = True
        cbcTmp.Flags = xtpFlagRightAlign
        Set cbcTmp = .Controls.Add(xtpControlComboBox, enuMenus.ѡ��ϵͳ�ؼ�, "")
        cbcTmp.Flags = xtpFlagRightAlign
        
        Set cbcTmp = .Controls.Add(xtpControlLabel, enuMenus.���ұ����ǩ, "����")
        cbcTmp.Flags = xtpFlagRightAlign
        Set cbmTmp = .Controls.Add(xtpControlCustom, enuMenus.���ұ���ؼ�, "")
        cbmTmp.handle = picFind.hwnd: cbmTmp.Flags = xtpFlagRightAlign
    End With
    
    '�˵���Ŀ����
    With cbsMain.KeyBindings
        'alt
        .Add 16, vbKeyI, enuMenus.���뱨��
        .Add 16, vbKeyO, enuMenus.��������
        .Add 16, vbKeyF1, enuMenus.����ȫ��
        .Add 16, vbKeyF2, enuMenus.����ȫ��
        .Add 16, vbKey1, enuMenus.TabRPT_1
        .Add 16, vbKey2, enuMenus.TabRPT_2
        'ctrl
        .Add 8, vbKeyX, enuMenus.�˳�
        .Add 8, vbKeyW, enuMenus.��������
        .Add 8, vbKeyM, enuMenus.�޸ı���
        .Add 8, vbKeyF, enuMenus.����
        .Add 8, vbKeyE, enuMenus.��Ʊ���
        'none
        .Add 0, vbKeyDelete, enuMenus.ɾ������
        .Add 0, vbKeyF8, enuMenus.ִ�б���
        .Add 0, vbKeyF12, enuMenus.��������
        .Add 0, vbKeyF1, enuMenus.��������
        .Add 0, vbKeyF3, enuMenus.���ұ���ؼ�
        .Add 0, vbKeyF5, enuMenus.ˢ��
    End With
    
    '��ͼ�꣬���ı��İ�ť���
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub InitDockPane()
    Dim panLeft As Pane, panRight As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeVisio
        
        Set panLeft = .CreatePane(1, 100, 0, DockLeftOf)
        With panLeft
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "�������"
            .MaxTrackSize.Width = 400
            .MinTrackSize.Width = 50
        End With
        
        Set panRight = .CreatePane(2, ScaleX(Me.Width, vbTwips, vbPixels) * 0.8, 0, DockRightOf)
        With panRight
            .Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .MinTrackSize.Width = 100
        End With
    End With
End Sub

Private Sub InitTabControl()
    With tbcRPT.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    
    With tbcRPT
        .InsertItem 0, "����(&1)", picReport.hwnd, 0
        .InsertItem 1, "������(&2)", picGroup.hwnd, 0
    End With
End Sub

Private Sub InitOther()
    On Error Resume Next
    
    With txtFind
        .Top = 0
        .Left = 0
        .MaxLength = 20
    End With
    
    With picFind
        .Width = txtFind.Width
        .Height = txtFind.Height
    End With
    
End Sub

Private Sub InitReportControl()
    '��ʼ��rptClass
    
    rptClass.ShowHeader = False
    rptClass.Icons = cbsMain.Icons
        
    If mobjClass Is Nothing Then
        Set mobjClass = New clsReportControlEx
    End If
    
    With mobjClass
        .AppTemplate atTree, rptClass, , "ID|�ϼ�ID|˵��", "ID|�ϼ�ID|����", Val("100-ͼ������")
        .Init Me
    End With
End Sub

Private Sub InitVSF()
    Set mobjReport = New clsVSFlexGridEx
    Set mobjGroup = New clsVSFlexGridEx
    Set mobjSub = New clsVSFlexGridEx
    
    With mobjReport
        .AppTemplate EM_Display, vsfReport, MSTR_REPORT_COLS, "", True
        .Init True
    End With
    
    With mobjGroup
        .AppTemplate EM_Display, vsfGroup, MSTR_GROUP_COLS, "", True
        .Init True
    End With
    
    With mobjSub
        .AppTemplate EM_Display, vsfGroupDetail, MSTR_GROUPDETAIL_COLS, "", True
        .Init True
    End With
End Sub

Private Sub FillData(ByVal bytType As Byte, Optional ByVal blnColumn As Boolean = False)
'���ܣ�Ϊ�ؼ���������
'������
'  blnColumn��True-��ͷ�����嶼�������ݣ�False-ֻ�����������
    
    Dim objCBS_ComBox As CommandBarComboBox
    Dim rsData As ADODB.Recordset
    Dim strSQL As String
    Dim lngClassID As Long, lngID As Long
    Dim intTab As Integer
    
    Set objCBS_ComBox = cbsMain.FindControl(, enuMenus.ѡ��ϵͳ�ؼ�, , True)
    
    Select Case bytType
    Case Val("1-rptClass")
        strSQL = _
            "Select * " & vbCr & _
            "From (" & vbCr & _
            "    Select ID, Nvl(�ϼ�id, 0) �ϼ�id, ����, ˵��" & vbCr & _
            "    From zlRPTClasses" & vbCr & _
            "    Union All " & vbCr & _
            "    Select 0, Null, '����', null From Dual" & vbCr & _
            ")" & vbCr & _
            "Start With �ϼ�ID Is Null Connect By Prior ID  = �ϼ�ID"
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��������Ϣ")
        If blnColumn Then
            Call mobjClass.Repaint(rsData, True)
        Else
            Call mobjClass.Repaint(rsData)
        End If
        If rptClass.Rows.count > 0 Then
            rptClass.FocusedRow = rptClass.Rows(0)
        End If
        rsData.Close
        
    Case Val("2-vsfReport")
        'ϵͳ��
        lngID = objCBS_ComBox.ItemData(objCBS_ComBox.ListIndex)
        '�������
        lngClassID = Val(rptClass.FocusedRow.Record.Item(mobjClass.GetColIndex("ID")).Value)
        
        strSQL = _
            "Select A.ID,A.���,A.����,A.˵��,A.����ID,A.�޸�ʱ��,A.����ʱ��,A.ϵͳ,A.���ִ��ʱ��, a.����ID, " & vbNewLine & _
            "    Decode(Nvl(A.Ʊ��, 0), 1, 'Ʊ��', '����') ����, " & vbNewLine & _
            "    Decode(Nvl(A.ϵͳ, 0), 0, '����', 'ϵͳ') ����, " & vbNewLine & _
            "    A.ִ����Ա ���ִ����, zlSpellCode(A.����) ����, b.���� �������, c.����������, d.������������, " & vbNewLine & _
            "    A.�Ƿ�ͣ�� ͣ�� " & vbNewLine & _
            "From zlReports A, zlRPTClasses B," & vbNewLine & _
            "   (Select c1.����id, f_List2Str(Cast(Collect(c2.����) as t_StrList)) ����������" & vbNewLine & _
            "    From zlRPTSubs C1, ZlRPTGroups C2" & vbNewLine & _
            "    Where c1.��id = c2.ID And c2.ϵͳ Is Null" & vbNewLine & _
            "    Group By c1.����id" & vbNewLine & _
            "    ) C," & vbNewLine & _
            "   (Select d1.����id, f_list2str(Cast(Collect(d2.����) As t_Strlist)) ������������" & vbNewLine & _
            "    From zlRPTDatas D1, zlConnections D2" & vbNewLine & _
            "    Where d1.�������ӱ�� = d2.���" & vbNewLine & _
            "    Group By d1.����id) D" & vbNewLine
        
        strSQL = strSQL & _
            "Where a.����id = b.id(+) And a.id = c.����id(+) And a.id = d.����id(+)" & vbNewLine & _
            IIF(lngID <= 0 _
                    , "    And a.ϵͳ Is Null " _
                    , "    And a.ϵͳ = [1] ") & vbNewLine & _
            IIF(mbytReportGroup = 1 _
                    , "    And Exists(Select 1 From zlRPTSubs Where ����id = a.Id) " _
                    , "    And Not Exists(Select 1 From zlRPTSubs Where ����id = a.Id) ") & vbNewLine & _
            IIF(mblnDisplayChild _
                    , IIF(lngClassID > 0 _
                            , " And b.Id In (Select ID From ZLRPTClasses Start With Id = [2] Connect By Prior ID = �ϼ�id) " _
                            , "") _
                    , IIF(lngClassID > 0 _
                            , " And b.Id = [2] " _
                            , " And Nvl(a.����Id, 0) = 0 ")) & _
            IIF(mblnDisable, " And a.�Ƿ�ͣ�� = 1 ", " ") & vbNewLine & _
            "Order by A.���"
        
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������Ϣ" _
                    , lngID, lngClassID)
                    
        mobjReport.Recordset = rsData
        If blnColumn Then
            Call mobjReport.Repaint(RT_ColsAndRows)
        Else
            Call mobjReport.Repaint(RT_Rows)
        End If
        rsData.Close
        
        If mbytReportGroup = Val("0-��ʾ��������") Then
            mobjReport.ColsHide = "���ܼ����|����������"
        Else
            mobjReport.ColsHide = "���ܼ����"
        End If
        If mblnDisplayChild = False Or lngID > 0 Then
            mobjReport.ColsHide = mobjReport.ColsHide & "|�������"
        End If
        mobjReport.SetColsHide
        
    Case Val("3-vsfGroup")
        'ϵͳ��
        lngID = objCBS_ComBox.ItemData(objCBS_ComBox.ListIndex)
        '�������
        lngClassID = rptClass.FocusedRow.Record.Item(mobjClass.GetColIndex("ID")).Value
        '��ǰҳ��
        intTab = tbcRPT.Selected.Index
        
        strSQL = _
            "Select a.���, a.���� ����, a.˵��, a.����ʱ��, a.ID, a.����id, a.����id, b.���� �������, a.�Ƿ�ͣ�� ͣ�� " & vbNewLine & _
            "From zlRPTGroups A, zlRPTClasses B " & vbNewLine & _
            "Where a.����id = b.Id(+) " & _
            IIF(lngID <= 0, " And a.ϵͳ Is Null", " And a.ϵͳ = [1]") & vbNewLine & _
            IIF(mblnDisplayChild = True And intTab = 1 _
                    , IIF(lngClassID > 0 _
                            , "    And a.����id in (Select Id From ZLRPTClasses Start With Id = [2] Connect By Prior ID = �ϼ�id)" _
                            , "") _
                    , IIF(lngClassID > 0, "    And a.����id = [2] ", " And Nvl(a.����id, 0) = 0 ")) & vbNewLine & _
            IIF(mblnDisable, " And a.�Ƿ�ͣ�� = 1 ", " ") & vbNewLine & _
            "Order By a.��� "
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��������Ϣ" _
                   , lngID, lngClassID)
        mobjGroup.Recordset = rsData
        If blnColumn Then
            Call mobjGroup.Repaint(RT_ColsAndRows)
        Else
            Call mobjGroup.Repaint(RT_Rows)
        End If
        rsData.Close
        
        If mblnDisplayChild And lngID <= 0 Then
            mobjGroup.ColsHide = ""
        Else
            mobjGroup.ColsHide = "�������"
        End If
        mobjGroup.SetColsHide
        
    Case Val("4-vsfGroupDetail")
        '������ID
        If vsfGroup.Row >= 1 Then
            lngID = Val(vsfGroup.TextMatrix(vsfGroup.Row, vsfGroup.ColIndex("ID")))
        End If
        
        strSQL = _
            "Select a.Id, b.��Id, a.���, a.����, a.˵��, a.����id, a.�޸�ʱ��, a.����ʱ��, a.ϵͳ, a.���ִ��ʱ��," & vbNewLine & _
            "    Decode(Nvl(A.Ʊ��, 0), 1, 'Ʊ��', '����') ����, " & vbNewLine & _
            "    Decode(Nvl(A.ϵͳ, 0), 0, '����', 'ϵͳ') ����, " & vbNewLine & _
            "    a.ִ����Ա ���ִ����, zlSpellCode(a.����) ����, d.������������, a.�Ƿ�ͣ�� ͣ�� " & vbNewLine & _
            "From ZLReports A, ZLRPTSubs B," & vbNewLine
'            "    (Select C1.����id, f_List2str(Cast(Collect(C2.����) As t_Strlist)) ����������" & vbNewLine & _
'            "     From zlRPTSubs C1, zlRPTGroups C2" & vbNewLine & _
'            "     Where C1.��id = C2.Id And C2.ϵͳ Is Null" & vbNewLine & _
'            "     Group By C1.����id) C," & vbNewLine &
        strSQL = strSQL & _
            "    (Select D1.����id, f_List2str(Cast(Collect(D2.����) As t_Strlist)) ������������" & vbNewLine & _
            "     From zlRPTDatas D1, Zlconnections D2" & vbNewLine & _
            "     Where D1.�������ӱ�� = D2.���" & vbNewLine & _
            "     Group By D1.����id) D" & vbNewLine & _
            "Where a.Id = b.����id And a.Id = d.����id(+)" & vbNewLine & _
            IIF(mblnDisable, " And a.�Ƿ�ͣ�� = 1 ", " ") & vbNewLine & _
            "    And b.��id = [1] " & vbNewLine & _
            "Order By a.��� "
        Set rsData = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��������ӱ���Ϣ" _
                   , lngID)
        mobjSub.Recordset = rsData
        If blnColumn Then
            Call mobjSub.Repaint(RT_ColsAndRows)
        Else
            Call mobjSub.Repaint(RT_Rows)
        End If
        rsData.Close
        
    Case Val("5-cboSystem")
        If Not objCBS_ComBox Is Nothing Then
            objCBS_ComBox.Clear
            
            strSQL = _
                "Select 0 ���, '����ϵͳ����' ���� From Dual Union All " & _
                "Select ���, ����||'��'||���||'��' From zlSystems Order By ���"
            Set rsData = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��װϵͳ��Ϣ")
            With rsData
                Do While .EOF = False
                    objCBS_ComBox.AddItem rsData!����
                    objCBS_ComBox.ItemData(objCBS_ComBox.ListCount) = rsData!���
                    .MoveNext
                Loop
                .Close
            End With
        
            If objCBS_ComBox.ListCount > 0 Then
                objCBS_ComBox.ListIndex = 1
                glngSys = objCBS_ComBox.ItemData(1)
            End If
            objCBS_ComBox.Width = 160
        End If
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    
    If Width < 8000 Then Width = 8000
    If Height < 5000 Then Height = 5000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String, strPane As String
    
    mdlPublic.SaveWinState Me, App.ProductName
    
    strRegPath = mdlPublic.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.name)
    If glngSys <= 0 Then
        strPane = dkpMain.SaveStateToString
        Call SaveSetting("ZLSOFT", strRegPath, "����", strPane)
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "��ʾ���з����¼�", IIF(mblnDisplayChild, "1", "0"))
    Call SaveSetting("ZLSOFT", strRegPath, "��ʾ�������", mbytReportGroup)
    Call SaveSetting("ZLSOFT", strRegPath, "�����С", mbytFontSize)
    Call SaveSetting("ZLSOFT", strRegPath, "����ʾͣ��״̬", IIF(mblnDisable, "1", "0"))
End Sub

Private Sub mobjGroup_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("����ʱ��")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("ͣ��")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("����ʱ��") > intCol Then
        intCol = vsfVar.ColIndex("����ʱ��")
    End If
    
    If Col = intCol Then
        lngIcon = Val("4-������")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("����ʱ��")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("ͣ��"))) = 1 Then
                lngIcon = Val("6-����ͣ��")
            Else
                lngIcon = Val("5-��������")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub mobjReport_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("����ʱ��")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("ͣ��")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("����ʱ��") > intCol Then
        intCol = vsfVar.ColIndex("����ʱ��")
    End If

    If Col = intCol Then
        lngIcon = Val("1-����")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("����ʱ��")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("ͣ��"))) = 1 Then
                lngIcon = Val("3-����ͣ��")
            Else
                lngIcon = Val("2-��������")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub mobjSub_EventFillData(ByVal vsfVar As VSFlex8Ctl.VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
    Dim intCol As Integer
    Dim lngIcon As Long
    
    intCol = vsfVar.ColIndex("����ʱ��")
    If intCol < 0 Then Exit Sub
    intCol = vsfVar.ColIndex("ͣ��")
    If intCol < 0 Then Exit Sub
    
    If vsfVar.ColIndex("����ʱ��") > intCol Then
        intCol = vsfVar.ColIndex("����ʱ��")
    End If

    If Col >= intCol Then
        lngIcon = Val("1-����")
        If vsfVar.TextMatrix(Row, vsfVar.ColIndex("����ʱ��")) <> "" And glngSys = 0 Then
            If Val(vsfVar.TextMatrix(Row, vsfVar.ColIndex("ͣ��"))) = 1 Then
                lngIcon = Val("3-����ͣ��")
            Else
                lngIcon = Val("2-��������")
            End If
        End If
        
        If lngIcon = 0 Then
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = Nothing
        Else
            Set vsfVar.Cell(flexcpPicture, Row, vsfVar.ColIndex("���")) = imgList.ListImages(lngIcon).Picture
        End If
    End If
End Sub

Private Sub picClass_Resize()
    On Error Resume Next
    
    With rptClass
        .Left = 0
        .Top = 0
        .Width = picClass.ScaleWidth
        .Height = picClass.ScaleHeight
    End With
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    
    With txtFind
        .Left = 0
        .Top = 0
        .Width = picFind.ScaleWidth
        If .Height > picFind.Height Then
            picFind.Height = .Height
        End If
    End With
End Sub

Private Sub picGroup_Resize()
    On Error Resume Next
    
    With picGroup_S
        .Left = 0
        .Width = picGroup.ScaleWidth
        If .Top > picGroup.ScaleHeight Then
            .Top = picGroup.ScaleHeight - 1500
        End If
    End With
    
    With vsfReport
        .Left = 0
        .Top = 0
        .Width = picReport.ScaleWidth
        .Height = picReport.ScaleHeight
    End With
    
    With vsfGroup
        .Left = 0
        .Top = 0
        .Width = picGroup.ScaleWidth
        .Height = picGroup_S.Top
    End With
    
    With lblGroupDetail
        .Top = picGroup_S.Top + picGroup_S.Height + 60
        .Left = 60
    End With
    
    With vsfGroupDetail
        .Left = 0
        .Top = lblGroupDetail.Top + lblGroupDetail.Height + 60
        .Width = picGroup.ScaleWidth
        .Height = picGroup.ScaleHeight - vsfGroup.Height - lblGroupDetail.Height - 60 * 2
    End With
End Sub

Private Sub picGroup_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '�϶�ʱ�ı���ɫ
    If Button = vbLeftButton Then picGroup_S.BackColor = &H80000010
End Sub

Private Sub picGroup_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        With picGroup_S
            If .Top + Y < picGroup.ScaleHeight * 0.3 Then
                .Top = picGroup.ScaleHeight * 0.3
                Exit Sub
            End If
            If .Top + Y > picGroup.ScaleHeight * 0.8 Then
                .Top = picGroup.ScaleHeight * 0.8
                Exit Sub
            End If
            .Move .Left, .Top + Y
        End With
    End If
End Sub

Private Sub picGroup_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picGroup_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    Call picGroup_Resize
End Sub

Private Sub picReport_Resize()
    On Error Resume Next
    
    With vsfReport
        .Left = 0
        .Top = 0
        .Width = picReport.ScaleWidth
        .Height = picReport.ScaleHeight
    End With
End Sub

Private Sub picRPT_Resize()
    On Error Resume Next
    
    With tbcRPT
        .Left = 0
        .Top = 0
        .Width = picRPT.ScaleWidth
        .Height = picRPT.ScaleHeight
    End With
End Sub

Private Sub rptClass_DragDrop(Source As Control, X As Single, Y As Single)
    Dim lngSelRow As Long, l As Long, lngClassID As Long, lngID As Long, lngTemp As Long
    Dim strSQL As String
    Dim objInfo As XtremeReportControl.ReportHitTestInfo
    
    On Error GoTo hErr
    
    Set objInfo = rptClass.HitTest(Me.ScaleX(X, vbTwips, vbPixels) _
                                 , Me.ScaleY(Y, vbTwips, vbPixels))
    If objInfo Is Nothing Then Exit Sub
    If objInfo.Row Is Nothing Then Exit Sub
    
    Select Case UCase(Source.name)
    Case "VSFREPORT", "VSFGROUP"
        lngSelRow = 0
        For l = 1 To Source.Rows - 1
            If Source.SelectedRow(lngSelRow) = l Then
                '��ӱ���������
                
                '������ID
                lngID = Val(Source.TextMatrix(l, Source.ColIndex("ID")))
                lngClassID = Val(objInfo.Row.Record(mobjClass.GetColIndex("ID")).Value)
                lngTemp = Val(Source.TextMatrix(l, Source.ColIndex("����ID")))
                If lngTemp <> 0 And lngTemp = lngClassID Then
                    MsgBox "�ܾ�ͬһ������϶���", vbInformation, App.Title
                    Exit Sub
                End If
            
                '�޸�
                If UCase(Source.name) = "VSFREPORT" Then
                    strSQL = _
                        "Update zlReports " & vbCrLf & _
                        "Set ����ID = " & IIF(lngClassID <= 0, "Null", lngClassID) & vbCrLf & _
                        "Where ID = " & lngID
                Else
                    strSQL = _
                        "Update zlRPTGroups " & vbCrLf & _
                        "Set ����ID = " & IIF(lngClassID <= 0, "Null", lngClassID) & vbCrLf & _
                        "Where ID = " & lngID
                End If
                gcnOracle.Execute strSQL
                
                lngSelRow = lngSelRow + 1
            End If
        Next
        
        rptClass.Tag = ""
        Call RefreshEx
    End Select
    
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub rptClass_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mblnReportControlFocus = True
    Call Modify
End Sub

Private Sub rptClass_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If rptClass.Visible Then rptClass.SetFocus
    Call PopupMenuEx(Val("3-������˵�"))
End Sub

Private Sub RefreshEx(Optional ByVal bytType As Byte = 0)
'���ܣ�
'������
'  bytType��0-ˢ�°�ť������1-�����㴥��

    Dim lngID As Long
    
    If Me.Visible = False Then Exit Sub
    
    If bytType = 1 Then
        mblnReportControlFocus = glngSys <= 0
    Else
        mblnReportControlFocus = False
    End If
    
    lngID = mobjClass.GetColIndex("ID")
    If rptClass.Tag <> rptClass.FocusedRow.Record.Item(lngID).Value Then
        If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
            Call FillData(Val("2-vsfReport"), False)
        Else
            Call FillData(Val("3-vsfGroup"), False)
            Call FillData(Val("4-vsfGroupDetail"), False)
        End If
    End If
    rptClass.Tag = rptClass.FocusedRow.Record.Item(lngID).Value
    Call UpdateStatusBar(rptClass)
    
    If mblnReportControlFocus Then
        Call VisibleToolButton(2)
    Else
        If tbcRPT.Selected.Index = 0 Then
            Call VisibleToolButton(0)
            vsfReport.SetFocus
        Else
            Call VisibleToolButton(1)
            vsfGroup.SetFocus
        End If
    End If
End Sub

Private Sub rptClass_SelectionChanged()
    Call RefreshEx(1)
End Sub

Private Sub SetControlFontSize(ByVal bytSize As Byte)
'���ܣ����ô���ؼ��������С
'������
'  bytSize��0-С���壻1-������

    mbytFontSize = bytSize
    Call mdlPublic.SetPublicFontSize(Me, bytSize)
    picFind.Height = txtFind.Height
    
    If bytSize = 1 Then
        mobjReport.HeightColumn = 450
        mobjReport.HeightRow = 350
        mobjGroup.HeightColumn = 450
        mobjGroup.HeightRow = 350
        mobjSub.HeightColumn = 450
        mobjSub.HeightRow = 350
    Else
        mobjReport.HeightColumn = 350
        mobjReport.HeightRow = 250
        mobjGroup.HeightColumn = 350
        mobjGroup.HeightRow = 250
        mobjSub.HeightColumn = 350
        mobjSub.HeightRow = 250
    End If
    '�ػ�߶�
    mobjReport.RepaintRowHeight
    mobjGroup.RepaintRowHeight
    mobjSub.RepaintRowHeight
End Sub

Private Sub tbcRPT_GotFocus()
    mblnReportControlFocus = False
End Sub

Private Sub tbcRPT_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible = False Then Exit Sub
    
    Set mobjFindPreviors = Nothing
    rptClass.Tag = ""
    Call rptClass_SelectionChanged
    
    mblnReportControlFocus = False
    If Item.Index = Val("0-����ҳ��") Then
        vsfReport.SetFocus
        Call VisibleToolButton
    Else
        vsfGroup.SetFocus
        Call VisibleToolButton(1)
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
    mblnReportControlFocus = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        '����
        Call Find(txtFind.Text, True)
    End Select
End Sub

Private Sub vsfGroup_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngID As Long
    
    If Me.Visible = False Then Exit Sub
    
    If OldRow <> NewRow Then
        lngID = Val(vsfGroup.TextMatrix(NewRow, vsfGroup.ColIndex("ID")))
        Call FillData(Val("4-vsfGroupDetail"), False)
        Call UpdateStatusBar(vsfGroup)
    End If
End Sub

Private Sub CheckSQLPlanEx()
'���ܣ���鵱ǰ�б��еı���ִ�мƻ��Ƿ������������
    Dim i As Long
    Dim objReport As Report, objData As RPTData
    Dim strSQLCheck As String, strErr As String, strFields As String
    Dim strMsg As String, objPar As RPTPar, strSQL As String
    Dim lngCount As Long

    If MsgBox("��ǰĿ¼һ��" & vsfReport.Rows - 1 & "�ű�����������Щ����(������)����Դ�е�SQL����ִ�мƻ���" & _
              "Ȼ����ִ�мƻ��Ƿ�������������" & vbCrLf & _
              "    1.�������ͱ��ȫ��ɨ��;" & vbCrLf & _
              "    2.�������ͱ������ȫɨ�����Ծʽ����ɨ��;" & vbCrLf & _
              "    3.��������û������Ǵ���������������������ҽ����¼_IX_������ĿID��;" & vbCrLf & _
              "    ���д����ָzlBakTables ZlBigTables�ж���ı�;" & vbCrLf & _
              "    ���ͱ���ָ�ռ�ͳ����Ϣ���¼����ȱʡ��3ǧ��1����֮��ı� (����ƽ����ִ�мƻ��鿴�п����¶���);" & vbCrLf & vbCrLf & _
              "�˹��̿��ܻỨ�Ѽ����ӵ�ʱ�䣬��ȷ��Ҫ������" _
        , vbQuestion + vbOKCancel + vbDefaultButton1, "���ܼ��") = vbCancel Then
        Exit Sub
    End If
    
    For i = 1 To vsfReport.Rows - 1
        Set objReport = ReadReport(Val(vsfReport.TextMatrix(i, vsfReport.ColIndex("ID"))), , True)
        strMsg = ""
        For Each objData In objReport.Datas
            With objData
                If .�������ӱ�� > 0 Then GoTo makContinue
                
                '�ȼ������Դ��SQL
                strSQLCheck = ""
                strFields = ""
                strSQL = RemoveNote(.SQL)
                strSQL = TrimChar(strSQL)
                strSQL = Replace(strSQL, "[ϵͳ]", glngSys)
                If GetParCount(strSQL) = 0 Then
                    strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .�������ӱ��)
                Else
                    strFields = mdlPublic.CheckSQL(strSQL, strErr, ReplaceParSysNo(.Pars, glngSys) _
                        , strSQLCheck, , objReport.Datas, .�������ӱ��)
                End If
                If strFields <> "" Then
                    If strSQLCheck <> "" Then
                        If mdlPublic.CheckSQLPlan(strSQLCheck, , .�������ӱ��) = True Then
                            strMsg = strMsg & "," & .����
                        End If
                    End If
                End If
                '�ټ�������ϸ�ͷ���SQL
                For Each objPar In .Pars
                    '�ų��Ѿ�������
                    If objPar.����SQL <> "" And InStr(strMsg, "(" & objPar.���� & ")[����]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.����SQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[ϵͳ]", glngSys)
                        Call mdlPublic.CheckParsRela(strSQL, objReport.Datas, objPar.����, True)
                        strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , objReport.Datas, .�������ӱ��)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If mdlPublic.CheckSQLPlan(strSQLCheck, , .�������ӱ��) = True Then
                                    strMsg = strMsg & "," & .���� & "(" & objPar.���� & ")[����]"
                                End If
                            End If
                        End If
                    End If
                    
                    If objPar.��ϸSQL <> "" And InStr(strMsg, "(" & objPar.���� & ")[��ϸ]") = 0 Then
                        strSQLCheck = ""
                        strFields = ""
                        strSQL = RemoveNote(objPar.��ϸSQL)
                        strSQL = TrimChar(strSQL)
                        strSQL = Replace(strSQL, "[ϵͳ]", glngSys)
                        Call mdlPublic.CheckParsRela(strSQL, objReport.Datas, objPar.����, True)
                        strFields = mdlPublic.CheckSQL(strSQL, strErr, , strSQLCheck, , , objData.�������ӱ��)
                        If strFields <> "" Then
                            If strSQLCheck <> "" Then
                                If mdlPublic.CheckSQLPlan(strSQLCheck, , objData.�������ӱ��) = True Then
                                    strMsg = strMsg & "," & .���� & "(" & objPar.���� & ")[��ϸ]"
                                End If
                            End If
                        End If
                    End If
                Next
            End With
makContinue:
        Next
        
        strMsg = Mid(strMsg, 2)
        If strMsg <> "" Then
            vsfReport.TextMatrix(i, vsfReport.ColIndex("���ܼ����")) = strMsg
            lngCount = lngCount + 1
        End If
        
        ShowFlash "���ڼ�鱨������ԴSQL���ڵ���������,���Ժ� ...", i / (vsfReport.Rows - 1)
    Next
    
    vsfReport.ColHidden(vsfReport.ColIndex("���ܼ����")) = False
    ShowFlash
    
    If lngCount > 0 Then
        MsgBox "����" & lngCount & "�ű���(������)������Դ���ܴ����������⣬���""������������Դ""�е���Ϣ��" & vbCrLf & _
               "���ڱ�����ƽ���鿴��ϸ��ִ�мƻ���������SQL�����Ż���" _
            , vbInformation, "���ܼ����"
    End If
End Sub

Private Function GetVsfControl(ByRef lngID As Long, ByRef blnIsGroup As Boolean _
    , Optional ByRef vsfActive As VSFlexGrid _
    , Optional ByRef strIDs As String) As Boolean
    
    Dim l As Long, lngSelRow As Long
    
    If Me.ActiveControl Is Nothing Then Exit Function
    If Me.ActiveControl.name = "" Then Exit Function
    
    lngID = 0
    blnIsGroup = False
    Set vsfActive = Nothing
    
    Select Case UCase(Me.ActiveControl.name)
    Case "VSFREPORT"
        Set vsfActive = vsfReport
    Case "VSFGROUP"
        Set vsfActive = vsfGroup: blnIsGroup = True
    Case "VSFGROUPDETAIL"
        Set vsfActive = vsfGroupDetail
    Case Else
        Exit Function
    End Select
    
    If Not vsfActive Is Nothing Then
        If vsfActive.Row > 0 Then
            lngID = Val(vsfActive.TextMatrix(vsfActive.Row, vsfActive.ColIndex("ID")))
        End If
    End If
    
    '��ѡ
    lngSelRow = 0: strIDs = ""
    If vsfActive.SelectedRows > 0 Then
        For l = 1 To vsfActive.Rows
            If vsfActive.SelectedRow(lngSelRow) = l Then
                strIDs = strIDs & "," & vsfActive.TextMatrix(l, vsfActive.ColIndex("ID"))
                lngSelRow = lngSelRow + 1
            End If
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End If
    
    GetVsfControl = True
End Function

Private Sub Import(ByVal lngMenuID As Long)
'���ܣ����뱨��

    Dim arrFile As Variant
    Dim i As Long, lngCurGroup As Long, lngGroupID As Long, lngID As Long
    Dim rsFiles As ADODB.Recordset, rsGroups As ADODB.Recordset
    Dim strRegPath As String, strPath As String, strFile As String, strSQL As String
    Dim strName As String, strCode As String
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim objFSO As New FileSystemObject, objFile As File, objFolder As Folder
    Dim arrTmp As Variant
    
    On Error GoTo hErr
    
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        'Ĭ��VSF�ؼ�
        If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
            Set vsfTemp = vsfReport
            blnGroup = False
        Else
            Set vsfTemp = vsfGroupDetail
            blnGroup = True
        End If
    End If
    
    If UCase(vsfTemp.name) = "VSFGROUPDETAIL" Then
        '�ӱ���
        Set vsfTemp = vsfGroup
        blnGroup = True
        lngGroupID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        '�鱨��
        lngID = 0
        lngGroupID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    Else
        '����
        lngGroupID = 0
        lngID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("ID")))
    End If
    
    strRegPath = "����ģ��\" & App.ProductName & "\Path"
    
    If lngMenuID = enuMenus.���뱨�� Then
        '���뱨��
        cdg.DialogTitle = "ѡ���뱨��"
        cdg.Filter = "�Զ��屨���ļ�|*.ZLR"
        cdg.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
        cdg.InitDir = GetSetting("ZLSOFT", strRegPath, "Import", GetSetting("ZLSOFT", strRegPath, "Export", App.Path))
        cdg.FileName = ""
        cdg.MaxFileSize = 32767
        cdg.CancelError = True
        
        On Error Resume Next
        cdg.ShowOpen
        If Err.Number = 0 Then
            On Error GoTo hErr
            
            If cdg.FileTitle = "" Then
                'ѡ�����ļ�����
                Call SaveSetting("ZLSOFT", strRegPath, "Import", Left(cdg.FileName, InStr(cdg.FileName, Chr(0)) - 1))
                arrFile = Split(cdg.FileName, Chr(0))
                For i = 1 To UBound(arrFile)
                    strFile = strFile & "|" & arrFile(0) & "\" & arrFile(i)
                Next
                strFile = Mid(strFile, 2)
            Else
                'ѡ�񵥸��ļ�����
                Call SaveSetting("ZLSOFT", strRegPath, "Import", Left(cdg.FileName, InStrRev(cdg.FileName, "\")))
                strFile = cdg.FileName
            End If
            If strFile = "" Then Exit Sub
            
            arrFile = Split(strFile, "|")
            
            Set rsFiles = CopyNewRec(Nothing, , True _
                            , Array("FilePath", adVarChar, 1000, Empty _
                                  , "FileName", adVarChar, 200, Empty _
                                  , "��ID", adBigInt, Empty, Empty _
                                  , "ͬ��ID", adBigInt, Empty, Empty _
                                  , "��������", adInteger, Empty, Empty _
                                  , "��������", adInteger, Empty, Empty _
                                  , "ErrType", adInteger, Empty, Empty _
                                  , "ImportResult", adInteger, Empty, Empty _
                                  , "ImportInfo", adVarChar, 200, Empty) _
                            )
            For i = LBound(arrFile) To UBound(arrFile)
                rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������", "��������" _
                                   , "ErrType", "ImportResult", "ImportInfo") _
                             , Array(arrFile(i), gobjFile.GetFileName(arrFile(i)), 0, 0, 0, 0, 0, 0, "")
            Next
            
            '����
            Call ImportReportBeach(glngSys, lngGroupID, lngID, rsFiles, vsfTemp)
        End If
        Err.Clear: On Error GoTo hErr
    Else
        '����ȫ��
        strPath = BrowseForFolder(Me.hwnd, "ѡ����Ҫ���뱨������Ŀ¼", strPath)
        If strPath <> "" Then
            If MsgBox("�Ƿ��롰" & strPath & "���ļ��м����ļ����µ����б���", vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                Exit Sub
            End If
            
            lngCurGroup = lngGroupID
            
            'FilePath=����ȫ·����FileName=�����ļ�������ID=����Ҫ����ı�����ID
            'ͬ��ID=�뽫Ҫ����ı���ͬ���ı���ı���ID���̶�����ͨ������ƥ�䣬�ǹ̶�ͨ������ƥ��
            '��������=0-�����룬1-��������,2-���ǵ���;��������=0-���帲�ǣ�1-������Դ����
            'ErrType=0-�޴���,1-�����ͬ����һ��������2-�����ͬ����һ�𸲸ǣ�3-ϵͳ����ֻ�ܸ��ǣ�������ͬ������
            '                            4-���ݴ�������,5-�汾��������,6-���Ʊ�Ŵ�������
            'ImportResult=-1-�Ѿ��ɹ����뵫�Ǳ��������δͨ����0-������,1-����ɹ�,2-����ʧ��
            'ImportInfo=����ɹ�����󷵻صı�����Ϣ
            Set rsFiles = CopyNewRec(Nothing, , True _
                                , Array("FilePath", adVarChar, 1000, Empty _
                                      , "FileName", adVarChar, 200, Empty _
                                      , "��ID", adBigInt, Empty, Empty _
                                      , "ͬ��ID", adBigInt, Empty, Empty _
                                      , "��������", adInteger, Empty, Empty _
                                      , "��������", adInteger, Empty, Empty _
                                      , "ErrType", adInteger, Empty, Empty _
                                      , "ImportResult", adInteger, Empty, Empty _
                                      , "ImportInfo", adVarChar, 200, Empty) _
                                )
            
            With rsFiles
                '�Ѽ����뵽���б����еĵı���,����ǰ�ļ����µı���
                For Each objFile In objFSO.GetFolder(strPath).Files
                    If UCase(objFile.name) Like "*.ZLR" Then
                        rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������" _
                                           , "��������", "ErrType", "ImportResult", "ImportInfo") _
                            , Array(objFile.Path, objFile.name, 0, 0, 0, 0, 0, 0, "")
                    End If
                Next
                '����Ҫ�����Զ��屨��ķ���
                '�̶��������ڱ���Ψһ�ԣ��Ѿ�ȷ������
                If glngSys = 0 Then
                    strSQL = "Select ID,���,���� From zlRPTGroups Where ϵͳ Is Null"
                    Set rsGroups = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption))
                End If
                
                '�Ѽ���ǰ�ļ��µ��Ӽ��ļ���
                For Each objFolder In objFSO.GetFolder(strPath).SubFolders
                    strFile = ""
                    For Each objFile In objFolder.Files
                        If UCase(objFile.name) Like "*.ZLR" Then
                            strFile = strFile & "|" & objFile.name
                        End If
                    Next
                    
                    If strFile <> "" Then
                        arrTmp = Split(Mid(strFile, 2), "|")
                        lngGroupID = 0
                        '���Զ�������Ҫ���ҷ��飬�̶��������ϵͳ�ű���ȷ������
                        If glngSys = 0 Then
                            Call SplitNameCode(objFolder.name, strName, strCode)
                            rsGroups.Filter = "���='" & strCode & "'"                          '���Ψһ��
                            If rsGroups.EOF Then rsGroups.Filter = "����='" & strName & "'"     '�����ӷ���û�б���
                            If Not rsGroups.EOF Then
                                lngGroupID = Nvl(rsGroups!id, 0)
                            Else
                                '���ɳ����Եı�����
                                '���������ƹ淶�����������µı�������
                                lngGroupID = GetNextID("zlRPTGroups")
                                If TLen(strName) > 30 Then strName = ConvertSBC(MidB(strName, 1, 30))
                                If strCode <> "" Then
                                    If TLen(strCode) > 20 Then strCode = ConvertSBC(MidB(strCode, 1, 20))
                                    If CheckExist("zlRPTGroups", "���", strCode) Then
                                        strCode = GetNextNO(True)
                                    End If
                                Else
                                    strCode = GetNextNO(True)
                                End If
                                strSQL = "Insert Into zlRPTGroups(ID,���,����,˵��) Values(" & _
                                                lngGroupID & "," & _
                                                "'" & strCode & "'," & _
                                                "'" & strName & "',Null)"
                                On Error Resume Next
                                gcnOracle.Execute strSQL
                                If Err.Number <> 0 Then
                                    lngGroupID = 0  '���ɱ�����ʧ�ܣ����Զ����÷����µı����뵽��������
                                Else '���ɷ���ɹ������뵽����Ϣ������
                                    rsGroups.AddNew Array("ID", "���", "����"), Array(lngGroupID, strCode, strName)
                                End If
                                On Error GoTo hErr
                            End If
                        End If
                        
                        For i = LBound(arrTmp) To UBound(arrTmp)
                            rsFiles.AddNew Array("FilePath", "FileName", "��ID", "ͬ��ID", "��������" _
                                               , "��������", "ErrType", "ImportResult", "ImportInfo") _
                                    , Array(objFolder.Path & "\" & arrTmp(i), arrTmp(i), lngGroupID, 0, 0, 0, 0, 0, "")
                        
                        Next
                    End If
                Next
                
                .Filter = "": .Sort = "��ID"
                If .RecordCount = 0 Then
                    MsgBox "��ǰ·����δ�ҵ��κοɵ���ı���", vbInformation, App.Title
                    Exit Sub
                End If
                
                Call ImportReportBeach(glngSys, lngCurGroup, lngID, rsFiles, vsfTemp, True)
            End With
        End If
    End If
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    Call mdlPublic.ErrCenter
End Sub

Private Sub Export(ByVal lngMenuID As Long)
'���ܣ���������

    Dim strPath As String, strRegPath As String, strChoose As String
    Dim strCode As String, strName As String, strFile As String, strPathTmp As String
    Dim strSQL As String
    Dim blnGroup As Boolean, blnDo As Boolean
    Dim lngID As Long, lngCount As Long, l As Long, lngSelRow As Long, lngExp As Long
    Dim vsfTemp As VSFlexGrid
    Dim rsReports As ADODB.Recordset
    Dim objFile As New FileSystemObject
    
    On Error GoTo hErr
    
    strRegPath = mdlPublic.FormatString("����ģ��\[1]\Path", App.ProductName)
    strPath = GetSetting("ZLSOFT" _
            , strRegPath _
            , "Export" _
            , GetSetting("ZLSOFT", strRegPath, "Import", App.Path))

    If lngMenuID = enuMenus.�������� Then
        '���
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            MsgBox "��ѡ�д������Ķ���������ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "��ѡ�д������Ķ���������ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        If UCase(vsfTemp.name) = "VSFGROUP" Then
            Set vsfTemp = vsfGroupDetail
        End If
        
        If vsfTemp.SelectedRows > 1 Then
            strChoose = frmMsgBox.ShowMsgBox(App.Title _
                        , "��ѡ�񱨱�����ʽ��" & _
                          "^������ǰ�嵥�е����б���ʱ���ļ��Զ�����[���]���ơ�������" & _
                          "^�������Ŀ¼�д�����ͬ���Ƶı����ļ����ļ����ݽ������ǡ�" _
                        , "���б���(&Y),!ѡ�б���(&N),?ȡ��(&C)" _
                        , Me)
        Else
            strChoose = frmMsgBox.ShowMsgBox(App.Title _
                        , "��ѡ�񱨱�����ʽ��" & _
                          "^������ǰ�嵥�е����б���ʱ���ļ��Զ�����[���]���ơ�������" & _
                          "^�������Ŀ¼�д�����ͬ���Ƶı����ļ����ļ����ݽ������ǡ�" _
                        , "���б���(&Y),!��ǰ����(&N),?ȡ��(&C)" _
                        , Me)
        End If
        If strChoose = "" Or strChoose = "ȡ��" Then Exit Sub
        
        If strChoose = "��ǰ����" Then
            'ȱʡ�Ա����������ļ���
            strCode = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("���"))
            strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����"))
            
            strFile = "[" & strCode & "]" & strName & ".ZLR"
            strFile = Replace(strFile, "\", "��")
            strFile = Replace(strFile, "/", "�M")
            strFile = Replace(strFile, ":", "��")
            strFile = Replace(strFile, "*", "�~")
            strFile = Replace(strFile, "?", "��")
            strFile = Replace(strFile, """", "")
            strFile = Replace(strFile, "<", "��")
            strFile = Replace(strFile, ">", "��")
            strFile = Replace(strFile, "|", "�O")

            cdg.DialogTitle = "���������ļ�"
            cdg.Filter = "�Զ��屨���ļ�|*.ZLR"
            cdg.Flags = &H200000 Or &H4 Or &H2 Or &H800 Or &H4000
            cdg.InitDir = strPath
            cdg.FileName = strFile
            cdg.CancelError = True

            On Error Resume Next
            Call cdg.ShowSave
            If Err.Number = 0 Then
                Call SaveSetting("ZLSOFT" _
                        , strRegPath _
                        , "Export" _
                        , Left(cdg.FileName, InStrRev(cdg.FileName, "\")))
                Call mdlPublic.ExportReport(CLng(lngID), cdg.FileName)
            End If
            On Error GoTo 0
        Else
            strFile = BrowseForFolder(Me.hwnd, "ѡ�񱨱���Ŀ¼", strPath)
            If strFile <> "" Then
                strPath = strFile
                Call SaveSetting("ZLSOFT", strRegPath, "Export", strPath)
                
                lngCount = IIF(strChoose = "ѡ�б���", vsfTemp.SelectedRows, vsfTemp.Rows - 1)
                If MsgBox("���ι����� " & lngCount & " �ű��� " & strPath & "��Ҫ������" _
                        , vbQuestion + vbYesNo + vbDefaultButton2 _
                        , App.Title) = vbNo Then
                    Exit Sub
                End If
                
                lngSelRow = 0
                For l = 1 To vsfTemp.Rows - 1
                    lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
                    strCode = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("���"))
                    strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
                    strFile = "[" & strCode & "]" & strName & ".ZLR"
                    
                    blnDo = False
                    If strChoose = "ѡ�б���" Then
                        If vsfTemp.SelectedRow(lngSelRow) = l Then
                            blnDo = True
                            lngSelRow = lngSelRow + 1
                        End If
                    Else
                        blnDo = True
                    End If
                    
                    If blnDo And lngID > 0 Then
                        Call ShowFlash("���ڵ���:" & strFile & ".ZLR", l / lngCount, Me, True)
                        If mdlPublic.ExportReport(lngID, strPath & "\" & strFile) = False Then
                            Call ShowFlash
                            If MsgBox("��������ʱ���ִ���Ҫ����������һ�ű�����", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                        End If
                    End If
                Next
                Call ShowFlash
            End If
        End If
    Else
        '��ǰϵͳȫ������
        strPath = BrowseForFolder(Me.hwnd, "ѡ�񱨱���Ŀ¼", strPath)
        If strPath <> "" Then
            Call SaveSetting("ZLSOFT", strRegPath, "Export", strPath)
            strSQL = _
                "Select A.Id, A.���, A.����, C.Id ��id, C.��� ����, C.���� ���� " & vbNewLine & _
                "From zlReports A, zlRPTSubs B, zlRPTGroups C " & vbNewLine & _
                "Where A.Id = B.����id(+) And B.��id = C.Id(+) And " & vbNewLine & _
                IIF(glngSys = 0, " A.ϵͳ Is Null ", " A.ϵͳ=[1] ") & vbNewLine & _
                "Order By C.���,A.��� "
            Set rsReports = OpenSQLRecord(strSQL, Me.Caption, glngSys)
            lngCount = rsReports.RecordCount
            
            If lngCount = 0 Then
                MsgBox "Ŀǰ�ޱ���ɵ�����", vbInformation, App.Title
                Exit Sub
            End If
            
            If MsgBox("���ι����� " & lngCount & " �ű��� " & strPath & "��Ҫ������" _
                , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
            
            lngExp = 0
            rsReports.MoveFirst
            For l = 1 To rsReports.RecordCount
                lngExp = lngExp + 1
                Call ShowFlash("���ڵ�����" & rsReports!���� & ".ZLR", lngExp / lngCount, Me, True)
                
                If Nvl(rsReports!��ID, 0) = 0 Then
                    strPathTmp = strPath
                Else
                    strPathTmp = strPath & "\[" & rsReports!���� & "]" & rsReports!����
                    If Not objFile.FolderExists(strPathTmp) Then
                        Call objFile.CreateFolder(strPathTmp)
                    End If
                End If
                strFile = "[" & rsReports!��� & "]" & rsReports!���� & ".ZLR"
                
                If Not ExportReport(rsReports!id, strPathTmp & "\" & strFile) Then
                    Call ShowFlash
                    If MsgBox("��������ʱ���ִ���Ҫ����������һ�ű�����" _
                        , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
                End If
                rsReports.MoveNext
            Next
            rsReports.Close

            Call ShowFlash
        End If
    End If
    
    Exit Sub
    
hErr:
    Call mdlPublic.ErrCenter
End Sub

Private Function ImportReportBeach(ByVal lngSys As Long, ByVal lngGroup As Long, ByVal lngCurPRTID As Long _
    , ByVal rsFiles As ADODB.Recordset, vsfImp As VSFlexGrid _
    , Optional ByVal blnALLImp As Boolean) As Boolean
'���ܣ��������뱨�����Ե���1�������
'������
'    lngSys = ��ǰѡ���ϵͳ
'    lngGroup = ��ǰѡ��ļ�¼��
'    rsFiles = ��Ҫ����ı����ļ�
'    lngCurPRTID = ��ǰѡ��ı���ID
'    blnALLImp=�Ƿ���ȫ�����룬�ǹ̶�����ȫ������ʱ��Ҳ��Ҫ��ȡ���б���
'    vsfImp=
'���أ��Ƿ�ɹ�����

    Dim rsReports As New ADODB.Recordset, strSQL As String
    Dim arrTmp As Variant, strInfo As String
    Dim strFilter As String
    Dim intErrType As Integer, intImpType As Integer, lngImpGroup As Long, lngRPTID As Long
    Dim strMsg As String, strOption As String, strReturn As String
    Dim i As Long, lngCount As Long, lngGroupID As Long, lngClassID As Long
    Dim blnSingle  As Boolean, strFileName As String
    Dim strCurRPT As String, strSameRPT As String
    
    On Error GoTo hErr
    
    '�̶������Լ�����ʾ�������µķǹ̶���������б������ʱ����Ҫ��ȡ���б���
    If lngSys <> 0 Or mbytReportGroup <> 0 And lngGroup = 0 And lngSys = 0 Or blnALLImp Then
        '��ѯ���еı���
        strSQL = _
            "Select A.ID,A.���,A.����,A.˵��,Nvl(B.��id,0) ��id " & vbNewLine & _
            "From zlReports A, zlRPTSubs B " & vbNewLine & _
            "Where " & IIF(lngSys = 0, " A.ϵͳ Is Null ", " A.ϵͳ=[1] ") & vbNewLine & _
            "    And A.ID=B.����ID(+)" & vbNewLine & _
            "Order by A.���"
    Else
        '�ǹ̶������ȡ
        If lngGroup <> 0 Then
            strSQL = _
                "Select Id,���,����,[2] ��id " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where Id In (Select ����id From Zlrptsubs Where ��id = [2]) " & vbNewLine & _
                "Order By ���"
        Else
            strSQL = _
                "Select ID,���,����,0 ��id " & vbNewLine & _
                "From zlReports " & vbNewLine & _
                "Where " & IIF(lngSys = 0, " ϵͳ Is Null ", " ϵͳ=[1] ") & vbNewLine & _
                "    And ID Not In (Select ����ID From zlRPTSubs) " & vbNewLine & _
                "Order by ���"
        End If
    End If
    Set rsReports = CopyNewRec(OpenSQLRecord(strSQL, Me.Caption, lngSys, lngGroup))
    
    If lngCurPRTID <> 0 Then
        rsReports.Filter = "ID=" & lngCurPRTID
        If rsReports.EOF Then
            MsgBox "��ǰѡ�б����Ѿ������ڣ���ˢ�º������", vbInformation, App.Title
            Exit Function
        Else
            strCurRPT = "[" & rsReports!��� & "]" & rsReports!����
        End If
    End If
    
    '��ȡ��ǰ�������ID
    lngClassID = 0
    If Not rptClass.FocusedRow Is Nothing Then
        lngClassID = Val(rptClass.FocusedRow.Record(mobjClass.GetColIndex("ID")).Value)
    End If
    If lngClassID < 0 Then lngClassID = 0
    
    With rsFiles
        '��ͬ���ļ����뵽ͬһ����ʱ��ͬ���ļ����
        '����������£�[GROUP_001]סԺ��������ASD��סԺ��������[GROUP_001]סԺ��������
        '                        ���������ļ��ı�����Ե��뵽[GROUP_001]סԺ������������
        '��ͬ�ļ����ı���������ͬһ������
        '��鵼���ļ����Լ�ȷ���������ͣ���������Լ����ǵı���ID��
        .Filter = "": .Sort = "FilePath Desc"
        blnSingle = rsFiles.RecordCount = 1 '�Ƿ񵥸�������
        If blnSingle Then strFileName = rsFiles!FileName
        Do While Not .EOF
            intErrType = 0: intImpType = 0: lngImpGroup = 0: lngRPTID = 0
            arrTmp = Split(GetReportInfo(!FilePath & ""), ";") '��ȡ�ļ���Ϣ
            If UBound(arrTmp) <> 2 Then
                intErrType = 4 '�ļ����
            ElseIf Val(arrTmp(2)) <> 9 Then
                intErrType = 5  '�汾���
                If blnSingle Then strFileName = strFileName & "(ԭʼ���ƣ�[" & arrTmp(0) & "]" & arrTmp(1) & ")"
            Else
                If blnSingle Then strFileName = strFileName & "(ԭʼ���ƣ�[" & arrTmp(0) & "]" & arrTmp(1) & ")"
                If lngSys = 0 Then '��ϵͳ����Ҫ�����ı����в��ܴ�����ͬ����
                    '�ǹ̶�����ȫ�������Ѿ�ȷ������Ҫ����ķ���
                    rsReports.Filter = "����='" & arrTmp(1) & "' And ���='" & arrTmp(0) & "' And ID>0 " & IIF(blnALLImp, " And ��ID=" & !��ID, "")
                    If rsReports.EOF Then rsReports.Filter = "����='" & arrTmp(1) & "'  And ID>0 " & IIF(blnALLImp, " And ��ID=" & !��ID, "")
                Else 'ϵͳ����ͨ�����ֱ�Ӳ���
                    rsReports.Filter = "����='" & arrTmp(1) & "' And ���='" & arrTmp(0) & "' And ID>0"
                    If rsReports.EOF Then rsReports.Filter = "���='" & arrTmp(0) & "' And ID>0"
                End If
                'ȷ��������ķ��飬������ڵ�ͬ���ģ����Ȳ���û�з���ı���
                rsReports.Sort = "ID Desc,��ID"
                If Not rsReports.EOF Then
                    lngRPTID = rsReports!id: lngImpGroup = rsReports!��ID
                    If lngRPTID = 0 Then
                        intErrType = 1 '�ñ����Ѿ����������
                    ElseIf lngRPTID < 0 Then
                        intErrType = 2 '�ñ����Ѿ�����Ǹ���
                    Else
                        intImpType = 2
                        '������Ʋ�ƥ��
                        If (CStr(arrTmp(0)) <> rsReports!��� & "" Or CStr(arrTmp(1)) <> rsReports!����) Then intErrType = 6
                        rsReports.Update "Id", lngRPTID * -1 '����Ѿ�����
                        If blnSingle Then strSameRPT = "[" & rsReports!��� & "]" & rsReports!����
                    End If
                Else
                    If lngSys <> 0 Then
                        intErrType = 3  'ϵͳ�̶�������븲��ͬ������
                    Else
                        intImpType = 1  '��ϵͳ����û��ͬ��������������
                        If lngSys = 0 And blnALLImp Then
                            lngImpGroup = !��ID         '����ȡԭ���ķ���
                        Else
                            lngImpGroup = lngGroup      '���뵽����ָ���ķ���
                        End If
                        '�ñ�����������������뻺�棬��ֹ�������
                        If !��ID = 0 Then
                            rsReports.AddNew Array("Id", "���", "����", "��iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), lngImpGroup)
                        Else
                            rsReports.AddNew Array("Id", "���", "����", "��iD"), Array(lngRPTID, arrTmp(0), arrTmp(1), !��ID)
                        End If
                    End If
                End If
            End If
            If lngSys = 0 And blnALLImp Then lngImpGroup = !��ID '�ǹ̶�������ȡԭ���ķ���
            .Update Array("��ID", "ͬ��ID", "��������", "ErrType") _
                  , Array(lngImpGroup, lngRPTID, intImpType, intErrType)
            .MoveNext
        Loop
        
        If blnSingle Then
            '���������ļ�
            .Filter = ""
            Select Case !ErrType
            Case 4
                MsgBox "����" & strFileName & "���������ݴ���������޷����룡", vbInformation, App.Title
                Exit Function
            Case 5
                MsgBox "����" & strFileName & "�����ڰ汾���Զ��޷����룡", vbInformation, App.Title
                Exit Function
            Case 3
                If lngCurPRTID <> 0 Then '����״̬��Ĭ�ϸ��ǵ�ǰ�ı���
                    .Update Array("��ID", "ͬ��ID", "��������", "ErrType"), Array(lngGroup, lngCurPRTID, 2, 6)
                Else
                    MsgBox "��ѡ����Ҫ���ǵı���������", vbInformation, App.Title
                    Exit Function
                End If
            End Select
            
            Select Case !��������
            Case 1
                strReturn = frmMsgBox.ShowMsgBox(App.Title, "�Ƿ��������뱨��""" & strFileName & """��", "��������(&N),!?ȡ��(&C)", Me)
            Case 2
                If lngSys = 0 And lngGroup = 0 Then '����ϵͳ�����Ϊ����ı���,��ʱ���Դ�����������ѡ��
                    If lngCurPRTID = !ͬ��ID Then
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """��Ż�����" & vbNewLine & "��Ҫ���ǵĵ�ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """��ź�����" & vbNewLine & "�뵱ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),��������(&N),!?ȡ��(&C)", Me)
                    ElseIf lngCurPRTID = 0 Then
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """���ڲ���ƥ��ı���""" & strSameRPT & """," & vbNewLine & "���Ƕ��߱�Ż����Ʋ��������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """���ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "����ƥ��(&O),��������(&N),!?ȡ��(&C)", Me)
                    Else
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """�ı�Ż�����" & vbNewLine & "�벿��ƥ�䱨��""" & strSameRPT & """" & vbNewLine & "�Լ���ǰѡ�񱨱�""" & strCurRPT & """�����������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """��Ż�����" & vbNewLine & "�뵱ǰѡ��""" & strCurRPT & """�������" & vbNewLine & "���Ǵ��ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),����ƥ��(&O),��������(&N),!?ȡ��(&C)", Me)
                    End If
                Else
                   If lngCurPRTID = !ͬ��ID Then
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """��Ż�����" & vbNewLine & "��Ҫ���ǵĵ�ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """��ź�����" & vbNewLine & "�뵱ǰѡ�񱨱�""" & strCurRPT & """���������ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),!?ȡ��(&C)", Me)
                    ElseIf lngCurPRTID = 0 Then
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """���ڲ���ƥ��ı���""" & strSameRPT & """," & vbNewLine & "���Ƕ��߱�Ż����Ʋ��������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """����" & vbNewLine & "���������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "����ƥ��(&O),!?ȡ��(&C)", Me)
                    Else
                        strMsg = IIF(!ErrType = 6, "����""" & strFileName & """�ı�Ż�����" & vbNewLine & "�벿��ƥ�䱨��""" & strSameRPT & """" & vbNewLine & " �Լ���ǰѡ�񱨱�""" & strCurRPT & """�����������ѡ��ȷ�ϣ�", _
                                    "����""" & strFileName & """��Ż�����" & vbNewLine & "�뵱ǰѡ��""" & strCurRPT & """�������" & vbNewLine & "���Ǵ��ڱ��������ƾ�����ı���""" & strSameRPT & """����ѡ��ȷ�ϣ�") & vbNewLine & "^^ע�⣺���Ҫ���Ǳ������ȶ�Ҫ���Ǳ�����б��ݡ�"
                        strReturn = frmMsgBox.ShowMsgBox(App.Title, strMsg, "���ǵ�ǰ(&S),����ƥ��(&O),!?ȡ��(&C)", Me)
                    End If
                End If
            End Select
            
            If strReturn = "" Then
                Exit Function
            ElseIf strReturn = "��������" Then
                .Update Array("��ID", "ͬ��ID", "��������", "ErrType") _
                      , Array(lngGroup, 0, 1, 0)
            Else
                If strReturn = "���ǵ�ǰ" Then
                    .Update Array("��ID", "ͬ��ID", "��������", "ErrType") _
                          , Array(lngGroup, lngCurPRTID, 2, 0)
                Else
                    .Update Array("��������", "ErrType") _
                          , Array(2, 0)
                End If
                strMsg = frmMsgBox.ShowMsgBox(App.Title _
                            , "�Ƿ�ֻ��������Դ��" & vbNewLine & _
                              "ֻ��������Դ���Ա������б���ĸ�ʽ������ϸ���������ѯϵͳ����Ա��" _
                            , "������Դ(&D),!?���嵼��(&F)" _
                            , Me)
                If strMsg = "������Դ" Then
                    .Update "��������", 1
                End If
            End If
        Else
            '��������ļ�
            If MsgBox("��ǰ������ű���ϵͳ���Զ�Ѱ�ұ��������ƥ��ı�����и��ǡ���ȷ���Ƿ������", vbInformation + vbYesNo, App.Title) = vbNo Then
                Exit Function
            End If
            
            '���ܵ����������Ϣ����
            .Filter = "ErrType>0 And ErrType<6": .Sort = "ErrType": intImpType = 0
            Do While Not .EOF
                If intImpType <> Val(!ErrType & "") Then
                    If intImpType <> 0 Then
                        strMsg = strMsg & vbNewLine
                    End If
                    intImpType = Val(!ErrType & ""): lngCount = 0
                    Select Case intImpType
                    Case 1
                        strMsg = strMsg & vbNewLine & "���±������ڴ�����ͬ���ݵı�����޷��������룺"
                    Case 2
                        strMsg = strMsg & vbNewLine & "���±������ڴ�����ͬ���ݵı�����޷����ǵ��룺"
                    Case 3
                        strMsg = strMsg & vbNewLine & "���±�������û�п��Ը��ǵı�����޷����룺"
                    Case 4
                        strMsg = strMsg & vbNewLine & "���±����������ݴ���������޷����룺"
                    Case 5
                        strMsg = strMsg & vbNewLine & "���±������ڰ汾���Զ��޷����룺"
                    End Select
                End If
                If lngCount < 4 Then
                    strMsg = strMsg & vbNewLine & !FileName
                ElseIf lngCount = 4 Then
                    strMsg = strMsg & vbNewLine & "... ..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            
            .Filter = "��������<>0"
            If .RecordCount = 0 Then 'û�е��뱨��
                MsgBox "û�п��Ե���ı���" & Mid(strMsg, 1, Len(strMsg) - 2) & "��", vbInformation, App.Title
                Exit Function
            End If
            
            '�ļ����Լ����벻ƥ����ʾ
            .Filter = "ErrType=6"
            If Not .EOF Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��Ż������븲�ǵı����������ѡ��ȷ�ϣ�"
                Do While Not .EOF
                    If lngCount < 5 Then
                        strMsg = strMsg & vbNewLine & CStr(lngCount + 1) & "." & !FileName
                    ElseIf lngCount = 5 Then
                        strMsg = strMsg & vbNewLine & "..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
                .Filter = "ErrType=0" '�����ڿ���ֱ�ӵ���ģ�����ʾ�Ƿ����
                If .RecordCount = 0 Then
                    strReturn = frmMsgBox.ShowMsgBox(App.Title _
                                    , Mid(strMsg, 1, Len(strMsg) - Len(vbNewLine)) _
                                    , "���帲��(&A),����Դ����(&D),!?ȡ��(&C)" _
                                    , Me)
                    If strReturn = "" Then Exit Function
                End If
            End If
            
            .Filter = "��������=2 And ErrType=0": .Sort = "ErrType" '���ڸ��Ǳ�������ʾѡ�����帲�ǣ���������Դ����
            If Not .EOF Then
                strMsg = strMsg & vbNewLine & "���±����Ḳ��ԭ�б�����ѡ��ȷ�ϣ�"
                strOption = "���帲��(&A),����Դ����(&D),!?ȡ��(&C)"
                lngCount = 0
            End If

            Do While Not .EOF
                If lngCount < 5 Then
                    strMsg = strMsg & vbNewLine & CStr(lngCount + 1) & "." & !FileName
                ElseIf lngCount = 5 Then
                    strMsg = strMsg & vbNewLine & "..."
                End If
                lngCount = lngCount + 1: .MoveNext
                If .EOF Then strMsg = strMsg & vbNewLine
            Loop
            
            .Filter = "��������=1" '��������
            If .RecordCount <> 0 And strReturn = "" And strOption = "" Then '���б�������
                strReturn = frmMsgBox.ShowMsgBox(App.Title _
                                , Mid(strMsg, Len(vbNewLine) + 1) & "��ȷ���Ƿ��룿" _
                                , "����(&N),!?ȡ��(&C)" _
                                , Me)
                If strReturn = "" Then Exit Function
            End If
            
            'ѡ�񸲸�����
            If strReturn = "" And strOption <> "" Then '���ڸ���,�Ҳ�����ErrType=6������
                strReturn = frmMsgBox.ShowMsgBox(App.Title, Mid(strMsg, Len(vbNewLine) + 1, Len(strMsg) - Len(vbNewLine) * 2), strOption, Me)
                If strReturn = "" Then Exit Function
            End If
        End If
        
        If strReturn = "����Դ����" Then
            .Filter = "��������=2"
            Do While Not .EOF
                .Update "��������", 1
                .MoveNext
            Loop
        End If
        
        Screen.MousePointer = vbHourglass
        
        .Filter = "��������<>0"
        .Sort = "��������"
        lngCount = .RecordCount
        Do While Not .EOF
            If Not blnSingle Then
                Call ShowFlash("���ڵ���:" & !FileName, i / lngCount, Me, True)
            Else
                Call ShowFlash("���ڵ���:" & !FileName, , Me, True)
            End If
            Me.Refresh
            DoEvents
            
            '��ʽ�����ļ�
            strInfo = ImportReport(!FilePath & "", Val(!ͬ��ID & ""), Val(!�������� & "") = 1 _
                                    , Val(!��ID & ""), lngClassID)
            .Update Array("ImportResult", "ImportInfo"), Array(IIF(strInfo <> "", 1, 2), strInfo)
            
            '�������Ȩ�޼��
            If strInfo <> "" Then
                arrTmp = Split(strInfo, "|")
                If Not mdlPublic.CheckReportPriv(CLng(arrTmp(0))) Then
                    .Update Array("ImportResult", "ͬ��ID"), Array(-1, Val(arrTmp(0)))
                Else
                    .Update "ͬ��ID", Val(arrTmp(0))
                End If
            End If
            
            i = i + 1
            .MoveNext
        Loop
        Call ShowFlash
        
        '���������ʾ
        strMsg = ""
        If Not blnSingle Then
            .Filter = "ImportResult=1 Or ImportResult=-1"
            If .RecordCount = 0 Then
                strMsg = "���б����Ϊ����ɹ���"
            Else
                strMsg = "�ɹ������� " & .RecordCount & " �ű���"
            End If
            
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±���ı����ļ����ݿ����ѱ��Ƿ��޸ģ�"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=-1 And ��������=1"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��û��Ȩ�޲�ѯ���µ��뱨����ȫ���򲿷����ݶ���"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=-1 And ��������=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "��û��Ȩ�޲�ѯ���µ��뱨����ȫ���򲿷����ݶ���,��ʹ�øñ���֮ǰ,���ֹ��Ա������ݽ��е�����"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=1 And ��������=2"
            If .RecordCount <> 0 And lngSys <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±���ɹ�������Ӧ����,�������Ҫ������Ȩ��������ʹ����Щ����"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
            
            .Filter = "ImportResult=2"
            If .RecordCount <> 0 Then
                lngCount = 0: strMsg = strMsg & vbNewLine & "���±�����ʧ�ܣ�"
                Do While Not .EOF
                    If lngCount < 4 Then
                        strMsg = strMsg & vbNewLine & !FileName
                    ElseIf lngCount = 4 Then
                        strMsg = strMsg & vbNewLine & "... ..."
                    End If
                    lngCount = lngCount + 1: .MoveNext
                    If .EOF Then strMsg = strMsg & vbNewLine
                Loop
            End If
        Else
            .Filter = ""
            Select Case !ImportResult
            Case -1
                strMsg = "��û��Ȩ�޲�ѯ����" & strFileName & "����ȫ���򲿷����ݶ���" & _
                         IIF(!�������� = 2, "���������Ҫ�ֹ��Ա������ݽ��е�����������Ȩ��������ʹ�øñ���", "��")
            Case 1
                strMsg = "����" & strFileName & "������ɹ�" & _
                         IIF(!�������� = 2, "���������Ҫ������Ȩ��������ʹ�øñ���", "��")
            Case 2
                strMsg = "����" & strFileName & "��" & _
                         IIF(!�������� = 2, "����ʧ�ܡ������ļ����ݿ����ѱ��Ƿ��޸ģ�", "��������ʧ�ܣ�")
            End Select
        End If
        
        Screen.MousePointer = vbDefault
        MsgBox strMsg, vbInformation, App.Title
    End With
    
    Exit Function
    
hErr:
    Call ShowFlash
End Function

Private Sub Delete(ByVal lngMenuID As Long)
    Dim rsCheck As New ADODB.Recordset, rsGetGroups As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, lngID As Long, lngSelRow As Long, lngCount As Long
    Dim strSQL As String, strIDs As String, strRec As String
    Dim vsfTemp As VSFlexGrid
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim arrItem As Variant
    Dim colSQL As New Collection
    
    If mblnReportControlFocus = False Then
        If GetVsfControl(lngID, blnGroup, vsfTemp, strIDs) = False Then
            MsgBox "��ѡ�д������Ķ������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "��ѡ�д������Ķ������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    Select Case lngMenuID
    Case enuMenus.ɾ��������
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "��ѡ��һ�������࣡", vbInformation, App.Title
            Exit Sub
        End If
        
        strRec = rptClass.FocusedRow.Record(mobjClass.GetColIndex("����")).Value
        
        If MsgBox(mdlPublic.FormatString("��ȷ��ɾ����[1]��������ࣿ" & vbCrLf & _
                                         "ע�⣺�������������齫�޷��࣬��������������Ȼ���ڡ�" _
                                , strRec) _
            , vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        
        'ɾ��
        With rptClass
            lngID = Val(.FocusedRow.Record(mobjClass.GetColIndex("ID")).Value)
            
            On Error GoTo hErr
            
            strSQL = _
                "Update zlReports Set ����id = Null " & vbNewLine & _
                "Where ����id In (Select ID From zlRPTClasses Start With ID = " & lngID & " Connect By Prior ID = �ϼ�id)"
            Call AddArray(colSQL, strSQL)
            
            strSQL = _
                "Update zlRPTGroups Set ����id = Null " & vbNewLine & _
                "Where ����id In (Select ID From zlRPTClasses Start With ID = " & lngID & " Connect By Prior ID = �ϼ�id)"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete zlRPTClasses Where ID = " & lngID
            Call AddArray(colSQL, strSQL)
            
            'ִ��DML
            gcnOracle.BeginTrans: blnTrans = True
            For lngRow = 1 To colSQL.count
                gcnOracle.Execute colSQL(lngRow)
            Next
            gcnOracle.CommitTrans: blnTrans = False
        End With
        
        'ˢ��
        Call FillData(Val("1-������"))
        
    Case enuMenus.ɾ��������
        If blnGroup = False Then
            MsgBox "��ѡ�񱨱��飡", vbInformation, App.Title
            Exit Sub
        End If
        
        '����Ƿ��ѷ���
        strRec = "": lngSelRow = 0: lngCount = 0
        For lngRow = 1 To vsfTemp.Rows - 1
            If lngCount <= 4 Then
                If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                    If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����ʱ��")) <> "" Then
                        strRec = strRec & vbCrLf & CStr(lngCount + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����"))
                        lngCount = lngCount + 1
                    End If
                    lngSelRow = lngSelRow + 1
                End If
            Else
                strRec = strRec & vbCrLf & "..."
                Exit For
            End If
        Next
        If strRec <> "" Then
            MsgBox "���б����Ѿ�����������ȡ����������ɾ����" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        strRec = GetSelectedReport(vsfTemp, "����")
        If MsgBox("��ȷ��Ҫɾ�����б�������" & strRec _
            , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        
        On Error GoTo hErr: blnTrans = True
        gcnOracle.BeginTrans
        gcnOracle.Execute "Delete zlRPTSubs Where ��ID=" & lngID
        gcnOracle.Execute "Delete zlRPTGroups Where ID=" & lngID
        gcnOracle.CommitTrans: blnTrans = False
        
    Case enuMenus.ɾ������
        '����Ƿ�Ϊ�������Ա
        lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.���� " & vbNewLine & _
            "From zlReports A, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = d.Column_Value " & vbNewLine & _
            "    And Exists(Select 1 From zlRPTSubs Where ����id = a.Id) " & vbNewLine & _
            "Order By a.���� "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!����
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If strRec <> "" Then
            MsgBox "���Ȱ����б���ӱ��������Ƴ�����ɾ����" & strRec _
                , vbInformation, App.Title
            Exit Sub
        End If
        
        '����Ƿ��ѷ���
        strRec = "": lngSelRow = 0: lngCount = 0
        For lngRow = 1 To vsfTemp.Rows - 1
            If lngCount <= 4 Then
                If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                    If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����ʱ��")) <> "" Then
                        strRec = strRec & vbCrLf & CStr(lngCount + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����"))
                        lngCount = lngCount + 1
                    End If
                    lngSelRow = lngSelRow + 1
                End If
            Else
                If lngCount > 4 Then
                    strRec = strRec & vbCrLf & "..."
                End If
                Exit For
            End If
        Next
        If strRec <> "" Then
            MsgBox "���б����Ѿ�����������ȡ����������ɾ����" & strRec, vbInformation, App.Title
            Exit Sub
        End If

        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.���� " & vbNewLine & _
            "From zlReports A, zlRPTPuts B, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = b.����Id And a.Id = d.Column_Value " & vbNewLine & _
            "Order By a.���� "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!����
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If strRec <> "" Then
            MsgBox "���б����Ѿ�����������ȡ����������ɾ����" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        '����Ƿ������������й���
        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(A, 10)*/ a.Id ����ID, a.���� " & vbNewLine & _
            "From zlReports A, Zlrptrelation B, Table(Cast(f_Str2List([1]) as t_StrList)) C " & vbNewLine & _
            "Where a.id = b.����id and a.id = c.Column_Value " & vbNewLine & _
            "Union all " & vbNewLine & _
            "Select /*+ cardinality(A, 10)*/ a.Id ����ID, a.���� " & vbNewLine & _
            "From zlReports A, Zlrptrelation B, Table(Cast(f_Str2List([1]) as t_StrList)) C " & vbNewLine & _
            "Where a.id = b.��������id and a.id = c.Column_Value "
        strSQL = "Select Distinct ����ID, ���� From (" & strSQL & ")"
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!����
                strRec = strRec & GetRelationList(rsCheck!����ID)
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        If strRec <> "" Then
            MsgBox "���б�����ڹ���������ȡ����������ɾ����" & strRec, vbInformation, App.Title
            Exit Sub
        End If
        
        '��ȡ��ɾ����������
        strRec = "": lngRow = 0
        strSQL = _
            "Select /*+ cardinality(D, 10)*/ a.���� " & vbNewLine & _
            "From zlReports A, Table(Cast(f_Str2List([1]) as t_StrList)) D " & vbNewLine & _
            "Where a.Id = d.Column_Value " & vbNewLine & _
            "Order By a.���� "
        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, strIDs)
        Do While rsCheck.EOF = False
            If lngRow <= 4 Then
                strRec = strRec & vbCrLf & CStr(lngRow + 1) & "." & rsCheck!����
            Else
                strRec = strRec & vbCrLf & "..."
                Exit Do
            End If
            lngRow = lngRow + 1
            
            rsCheck.MoveNext
        Loop
        rsCheck.Close
        
        If MsgBox("ȷ��Ҫɾ�����б�����" & strRec _
                , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Exit Sub
        End If
        
        On Error GoTo hErr
        blnTrans = True
        gcnOracle.BeginTrans
        
        arrItem = Split(strIDs, ",")
        For lngRow = LBound(arrItem) To UBound(arrItem)
            lngID = arrItem(lngRow)
            If lngID <> 0 Then
                gcnOracle.Execute "Delete From zlReports Where ID=" & CStr(lngID)
            End If
        Next
        
        gcnOracle.CommitTrans
        blnTrans = False
        On Error GoTo 0
    End Select
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Call ErrCenter
End Sub

Private Sub SplitNameCode(ByVal strInput As String, ByRef strName As String, ByRef strCode As String)
'����:�ָ��������
'������strInput=������ַ����������ʽΪ[����]����,���Զ��ָ����Ĭ��Ϊֻ��ȡ������
'���أ�strName=����
'           strCode=����
    Dim arrTmp As Variant
    Dim strTmp As Variant
    If InStr(strInput, "\") > 0 Then
        strTmp = StrReverse(strInput)
        strInput = StrReverse(Mid(strTmp, 1, InStr(strTmp, "\") - 1))
    End If
    
    If strInput Like "[[]?*[]]?*" Then '���Ϲ淶���ļ���
        arrTmp = Split(strInput, "]")
        strName = arrTmp(1)
        strCode = Mid(arrTmp(0), 2)
    Else
        strName = strInput
        strCode = ""
    End If
End Sub

Private Sub Modify()
    Dim lngID As Long, lngProgID As Long, lngGroupID As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim strCode As String, strName As String, strDescription As String
    Dim bytMode As Byte
    
    '���
    If mblnReportControlFocus Then
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "��ѡ��һ�������࣡", vbInformation, App.Title
            Exit Sub
        End If
    Else
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        
        lngProgID = Val(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����ID")))
        strCode = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("���"))
        strDescription = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("˵��"))
    End If
        
    If mblnReportControlFocus Then
        '������
        bytMode = Val("2-������")
        lngProgID = 0
        strCode = ""
        With rptClass.FocusedRow
            lngGroupID = Val(Nvl(.Record(mobjClass.GetColIndex("�ϼ�ID")).Value, 0))
            lngID = Val(Nvl(.Record(mobjClass.GetColIndex("ID")).Value, 0))
            strName = .Record(mobjClass.GetColIndex("����")).Value
            strDescription = Nvl(.Record(mobjClass.GetColIndex("˵��")).Value)
        End With
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����"))
        bytMode = Val("1-������")
    Else
        If UCase(vsfTemp.name) = "VSFGROUPDETAIL" Or mbytReportGroup = 1 Then
            bytMode = Val("3-�ӱ���")
        Else
            bytMode = 0
        End If
        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����"))
    End If
    
    If mblnReportControlFocus Then
    Else
        If blnGroup Then
            lngGroupID = lngID
            lngID = 0
        Else
            lngGroupID = 0
        End If
    End If
    
    '�޸ı���
    If frmReportEdit.ShowMe(Me, glngSys, bytMode, lngProgID, lngGroupID, lngID, strName, strCode, strDescription) Then
        If mblnReportControlFocus Then
            'ˢ�·���ؼ�
            Call FillData(1, False)
        End If
        
        'ˢ��
        rptClass.Tag = ""
        Call RefreshEx
    End If
    Unload frmReportEdit
    Exit Sub
    
hErr:
    Call ErrCenter
    Call SaveErrLog
    Unload frmReportEdit
End Sub

Private Sub Design()
    Dim lngID As Long, lngindex As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup = True Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If

    If CheckPass(lngID) = False Then
        MsgBox "�������ݴ��󣬲�����Ƹñ���", vbInformation, App.Title
        Exit Sub
    End If
    If CheckReportPriv(lngID) = False Then
        MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ���������ƻ�����������", vbInformation, App.Title
    End If
    
    frmDesign.lngRPTID = lngID
    
    On Error Resume Next
    frmDesign.Show vbModal, Me
    On Error GoTo hErr
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    Exit Sub
    
hErr:
    Call ErrCenter
End Sub

Private Sub vsfGroup_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton(1)
End Sub

Private Sub vsfGroup_DblClick()
    mblnReportControlFocus = False
    Call Modify
End Sub

Private Sub vsfGroup_GotFocus()
    mblnReportControlFocus = False
    Call VisibleToolButton(1)
End Sub

Private Sub vsfGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And vsfGroup.Rows > 1 Then
        Call vsfGroup.Drag
    End If
End Sub

Private Sub vsfGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfGroup.Visible And vsfGroup.Enabled Then vsfGroup.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("2-������˵�"))
    End If
End Sub

Private Sub vsfGroupDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    Call UpdateStatusBar(vsfGroupDetail)
End Sub

Private Sub vsfGroupDetail_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfGroupDetail_DblClick()
    Call Design
End Sub

Private Sub vsfGroupDetail_GotFocus()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfGroupDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfGroupDetail.Visible And vsfGroupDetail.Enabled Then vsfGroupDetail.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("1-����˵�"))
    End If
End Sub

Private Sub vsfReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    If NewRow <> OldRow Then Call UpdateStatusBar(vsfReport)
End Sub

Private Sub vsfReport_Click()
    mblnReportControlFocus = False
    Call VisibleToolButton
End Sub

Private Sub vsfReport_DblClick()
    Dim cbcTemp As CommandBarControl
    
    Set cbcTemp = cbsMain.FindControl(, enuMenus.��Ʊ���, , True)
    If Not cbcTemp Is Nothing Then
        cbcTemp.Execute
    End If
End Sub

Private Sub SetToolBarButtonIcon(ByVal lngMenuID As Long, ByVal lngIcon As Long)
    Dim cbcControl As CommandBarControl
    
    Set cbcControl = cbsMain.FindControl(, lngMenuID, , True)
    If Not cbcControl Is Nothing Then
        If cbcControl.IconId <> lngIcon Then cbcControl.IconId = lngIcon
    End If
End Sub

Private Sub vsfReport_GotFocus()
    mblnReportControlFocus = False
End Sub

Private Sub vsfReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mbytReportGroup = 0 And vsfReport.Rows > 1 Then
            Call vsfReport.Drag
        End If
    End If
End Sub

Private Sub vsfReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If vsfReport.Visible And vsfReport.Enabled Then vsfReport.SetFocus
        mblnReportControlFocus = False
        Call PopupMenuEx(Val("1-����˵�"))
    End If
End Sub

Private Sub PopupMenuEx(ByVal bytType As Byte)
    Dim cbrTmp As XtremeCommandBars.CommandBar
    Dim cbbTmp As XtremeCommandBars.CommandBarButton
    Dim cbpTmp As XtremeCommandBars.CommandBarPopup
    
    Select Case bytType
    Case Val("1-����˵����ӱ���˵�")
        Set cbrTmp = cbsMain.Add("����", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.��������, "��������")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.�޸ı���, "�޸ı���")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.ɾ������, "ɾ������")
            
            Set cbbTmp = .Add(xtpControlButton, enuMenus.��Ʊ���, "��Ʊ���"): cbbTmp.BeginGroup = True
            Set cbbTmp = .Add(xtpControlButton, enuMenus.ִ�б���, "ִ�б���")
            
            If glngSys = 0 Then
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.������, "������"): cbpTmp.BeginGroup = True
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̨�˵�, "������̨�˵�(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.ȡ������, "ȡ������")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ӵ���̨�˵�, "�ӵ���̨�˵�(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
                
                Set cbbTmp = .Add(xtpControlButton, enuMenus.��������, "����(&S)"): cbbTmp.BeginGroup = True
                Set cbbTmp = .Add(xtpControlButton, enuMenus.����ͣ��, "ͣ��(&T)")
            End If
        End With
    Case Val("2-������˵�")
        Set cbrTmp = cbsMain.Add("������", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.����������, "����������(&N)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.�޸ı�����, "�޸ı�����(&M)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.ɾ��������, "ɾ��������(&D)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.ִ�б���, "ִ�б�����"): cbbTmp.BeginGroup = True
            
            If glngSys = 0 Then
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.������, "������"): cbpTmp.BeginGroup = True
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̨�˵�, "������̨�˵�(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
                Set cbpTmp = .Add(xtpControlPopup, enuMenus.ȡ������, "ȡ������")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ӵ���̨�˵�, "�ӵ���̨�˵�(&1)")
                    Set cbbTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ģ���ڲ˵�, "��ģ���ڲ˵�(&2)")
                    
                Set cbbTmp = .Add(xtpControlButton, enuMenus.��������, "����(&S)"): cbbTmp.BeginGroup = True
                Set cbbTmp = .Add(xtpControlButton, enuMenus.����ͣ��, "ͣ��(&T)")
            End If
        End With
    Case Val("3-������˵�")
        Set cbrTmp = cbsMain.Add("������", xtpBarPopup)
        With cbrTmp.Controls
            Set cbbTmp = .Add(xtpControlButton, enuMenus.����������, "�����������(&N)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.�޸ı�����, "�޸ı������(&M)")
            Set cbbTmp = .Add(xtpControlButton, enuMenus.ɾ��������, "ɾ���������(&D)")
        End With
    End Select
    
    If Not cbrTmp Is Nothing Then
        Call cbrTmp.ShowPopup
    End If
End Sub

Private Sub NewEx()
    Dim lngProgID As Long, lngGroupID As Long, lngID As Long, l As Long
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim strCode As String
    Dim bytMode As Byte
    
    '���
    If mblnReportControlFocus Then
        If rptClass.SelectedRows.count <= 0 Then
            MsgBox "��ѡ��һ�������࣡", vbInformation, App.Title
            Exit Sub
        End If
    Else
        If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
            'ȱʡ�ؼ�
            On Error Resume Next
            vsfReport.SetFocus
            If Err.Number = 0 Then
                Set vsfTemp = Me.vsfReport
            Else
                MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
                Exit Sub
            End If
            On Error GoTo 0
        End If
    End If

    If mblnReportControlFocus Then
        '������
        bytMode = Val("2-������")
        With rptClass.FocusedRow
            lngGroupID = Val(Nvl(.Record(mobjClass.GetColIndex("�ϼ�ID")).Value, 0))
        End With
    ElseIf UCase(vsfTemp.name) = "VSFGROUPDETAIL" Then
        bytMode = Val("0-����")
        lngProgID = Val(vsfGroup.TextMatrix(vsfGroup.Row, vsfGroup.ColIndex("����ID")))
        If lngProgID <= 0 Then
            MsgBox "�����顰����ID�������쳣������ϵ����", vbInformation, App.Title
            Exit Sub
        End If
    ElseIf UCase(vsfTemp.name) = "VSFGROUP" Then
        bytMode = Val("1-������")
    Else
        bytMode = Val("0-����")
    End If
    
    If frmReportEdit.ShowMe(Me, glngSys, bytMode, lngProgID, lngGroupID, , , strCode) Then
        If mblnReportControlFocus Then
            'ˢ�·���ؼ�
            Call FillData(1, False)
        Else
            If (UCase(vsfTemp.name) = "VSFREPORT" Or UCase(vsfTemp.name) = "VSFGROUPDETAIL") Then
                'ˢ��
                rptClass.Tag = ""
                Call RefreshEx
                
                '��λ
                For l = 1 To vsfTemp.Rows - 1
                    If UCase(strCode) = UCase(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("���"))) Then
                        '���
                        vsfTemp.Row = l
                        
                        If MsgBox("��Ҫ������Ʊ�����", vbQuestion + vbDefaultButton1 + vbYesNo) = vbYes Then
                            Call Design
                        End If
                        
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub Remove()
    Dim lngID As Long, lngGroupID As Long, lngSelRow As Long, lngRow As Long
    Dim rsGroup As ADODB.Recordset
    Dim colSQL As New Collection
    Dim strRec As String, strSQL As String
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ���ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "��ѡ���ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    strRec = "": lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        If lngSelRow <= 4 Then
            If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                strRec = strRec & vbCrLf & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����"))
                lngSelRow = lngSelRow + 1
            End If
        Else
            strRec = strRec & "|..."
            Exit For
        End If
    Next
    
    If MsgBox("��ȷ��Ҫ�ӱ��������Ƴ����±�����" & strRec _
        , vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        '��ȡ�ӱ�����
        If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
            lngID = Val(vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("ID")))
            Set rsGroup = GetReportGroups(lngID)
            Do While rsGroup.EOF = False
                lngGroupID = Nvl(rsGroup!id, 0)
                
                strSQL = "Update zlRPTSubs Set ���=���-1 " & _
                         "Where ���>(Select ��� From zlRPTSubs Where ����ID=" & lngID & " And ��ID=" & lngGroupID & ")" & _
                         "    And ��ID=" & lngGroupID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Delete From zlRPTSubs Where ����ID=" & lngID & " And ��ID=" & lngGroupID
                Call AddArray(colSQL, strSQL)
                
                rsGroup.MoveNext
            Loop
            rsGroup.Close
            
            lngSelRow = lngSelRow + 1
        End If
    Next
    
    'ִ��DML
    gcnOracle.BeginTrans: blnTrans = True
    For lngRow = 1 To colSQL.count
        gcnOracle.Execute colSQL(lngRow)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
End Sub

Private Function GetReportGroups(ByVal lngID As Long) As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    
    strSQL = _
        "Select a.Id, a.���, a.���� " & vbNewLine & _
        "From zlRPTGroups A, zlRPTSubs B " & vbNewLine & _
        "Where a.Id = b.��id And ϵͳ Is Null And b.����id = [1] " & vbNewLine & _
        "Order By a.����"
    Set GetReportGroups = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��������Ϣ", lngID)
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Guide()
    Dim objReport As Report, objItem As Object
    Dim objCBS_ComBox As CommandBarComboBox
    Dim objControl As Object
    Dim lngNextID As Long
    Dim strSQL As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    Set objReport = New Report
    With objReport
        '��ֽ��15-ȱʡΪ�Զ�ѡ��
        .��ֽ = 15
        'ȱʡʹ�õ�ǰ��ӡ��
        If Printers.count > 0 Then .��ӡ�� = Printer.DeviceName
        'ȱʡΪA4����,Ϊ����
        .Fmts.Add 1, "��ʽ1", INIT_WIDTH, INIT_HEIGHT, 9, 1, False, 0, "_1"
    End With
    
    frmGuide.blnNew = True
    Set frmGuide.objReport = objReport
    Set frmGuide.mobjFmt = objReport.Fmts(1)
    frmGuide.Show vbModal, Me
    
    If gblnOK Then
        Set objControl = cbsMain.FindControl(, enuMenus.ѡ��ϵͳ�ؼ�, , True)
        If Not objControl Is Nothing Then
            '�ָ���ϵͳ����ѡ��
            objControl.ListIndex = 1
            
            'ˢ�½���
            Call SelectedSysComboBox(objControl)
        End If
        
        '���ɱ���
        With frmGuide
            Set objReport.Items = .objGuide.Items       '����Ԫ�ض��󼯺�
            Set objReport.Datas = .objGuide.Datas       '��������Դ���󼯺�
            Set objReport.Fmts = .objGuide.Fmts         '�����ʽ���󼯺�
            
            lngNextID = GetNextID("zlReports")
            strSQL = "Insert Into zlReports(ID,���,����,˵��,ϵͳ,����) " & vbCrLf & _
                     "Values (" & _
                        lngNextID & _
                        ",'" & .txtNO.Text & "'" & _
                        ",'" & .txtTitle.Text & "'" & _
                        ",'" & .txtNote.Text & "'" & _
                        "," & IIF(glngSys = 0, "NULL", glngSys) & _
                        "," & AdjustStr(GetPass(.txtNO, .txtTitle)) & ")"
                        
            On Error GoTo hErr
            
            gcnOracle.BeginTrans: blnTrans = True
            gcnOracle.Execute strSQL
            gcnOracle.CommitTrans: blnTrans = False
            
            '��������
            If Not SaveReport(lngNextID, objReport, staMain.Panels(2)) Then
                gcnOracle.BeginTrans: blnTrans = True
                gcnOracle.Execute "Delete From zlReports Where ID=" & lngNextID
                gcnOracle.CommitTrans: blnTrans = False
                
                MsgBox "�����ɱ���ʱ�����������,�����Ըò�����", vbInformation, App.Title
                Unload frmGuide
                Exit Sub
            End If

        End With
        Unload frmGuide
        
        'ˢ��
        rptClass.Tag = ""
        Call RefreshEx
    End If
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Unload frmGuide
End Sub

Private Sub SelectedSysComboBox(ByVal objControl As XtremeCommandBars.CommandBarComboBox)
    Dim lngID As Long

    If objControl Is Nothing Then
        glngSys = 0
        GoSub proPub
        Exit Sub
    End If
    
    '���½���
    If objControl.ListIndex > Val("1-ϵͳ����") Then
        If dkpMain.Panes(Val("1-������")).Closed = False Then dkpMain.Panes(Val("1-������")).Close
        rptClass.FocusedRow = rptClass.Rows(0)
    Else
        dkpMain.ShowPane Val("1-������")
    End If
    
    '���±���
    glngSys = objControl.ItemData(objControl.ListIndex)
    GoSub proPub
    
    '���½���
    rptClass.Tag = ""
    Call rptClass_SelectionChanged
    Exit Sub
    
proPub:
    mblnAllowPublish = False
    Select Case glngSys
    Case Val("0-����"), Val("200-����"), Val("500-�ɱ�"), Val("700-����")
        mblnAllowPublish = True
    Case Else
        If Val(GetSysVersion(lngID)) >= 10 Then
            mblnAllowPublish = True
        End If
    End Select
    Return
End Sub

Private Sub ShowRunLog()
    Dim lngID As Long
    Dim strName As String
    Dim blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��Ҫ�鿴��־�Ķ��������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "��ѡ��Ҫ�鿴��־�Ķ��������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    strName = Trim(vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����")))
    
    '�鿴����������־��¼
    If lngID > 0 Then
        Call frmReportRunLog.ShowMe(Me, lngID, "����" & strName & "����������־")
    End If
End Sub

Private Sub VisibleToolButton(Optional ByVal bytMode As Byte = 0)
'���ܣ����¹��������������޸ġ�ɾ������ť��ʾ
'���ܣ�
'  bytMode��0-����1-�����飻2-������

    Dim objTemp As Object
    
    Select Case bytMode
    Case 1
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.����������, enuMenus.�޸ı�����, enuMenus.ɾ�������� _
                , enuMenus.��������, enuMenus.�޸ı���, enuMenus.ɾ������
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    Case 2
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.����������, enuMenus.�޸ı�����, enuMenus.ɾ�������� _
                , enuMenus.��������, enuMenus.�޸ı���, enuMenus.ɾ������
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    Case Else
        For Each objTemp In cbsMain.Item(2).Controls
            Select Case objTemp.id
            Case enuMenus.����������, enuMenus.�޸ı�����, enuMenus.ɾ�������� _
                , enuMenus.����������, enuMenus.�޸ı�����, enuMenus.ɾ��������
                objTemp.Visible = False
            Case Else
                objTemp.Visible = True
            End Select
        Next
    End Select
End Sub

Private Function GetReportObjects(ByVal lngID As Long, Optional ByVal AllConnect As Boolean = False) As String
'���ܣ���ȡָ����������ݿ���ʶ���
'������
'  lngID������ID
'  AllConnect��True-�����������ӣ�False-��ǰ��������
'���أ���������
    
    Dim strObjects As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim l As Long, k As Long
    Dim arrTmp As Variant, arrItem As Variant
    
    On Error GoTo hErr

    '�������������Դ���ʶ���
    strObjects = ""
    strSQL = "Select ���� From zlRPTDatas " & _
             "Where ���� is Not NULL And ����ID=[1] " & _
             IIF(AllConnect, "", "And nvl(�������ӱ��,0) <= 0")
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡָ����������ݿ���ʶ�������", lngID)
    Do While rsTmp.EOF = False
        arrTmp = Split(rsTmp!����, ",")
        For l = LBound(arrTmp) To UBound(arrTmp)
            If InStr(strObjects & ",", "," & arrTmp(l) & ",") = 0 Then
                strObjects = strObjects & "," & arrTmp(l)
            End If
        Next
        rsTmp.MoveNext
    Loop
    rsTmp.Close

    '�����ñ���Ĳ�������Դ���ʶ���
    strSQL = "Select B.���� From zlRPTDatas A,zlRPTPars B " & _
             "Where A.ID=B.ԴID And B.���� is Not NULL And A.����ID=[1] " & _
             IIF(AllConnect, "", "And nvl(�������ӱ��,0) <= 0")
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡָ���������в��������ݿ���ʶ�������", lngID)
    Do While rsTmp.EOF = False
        arrTmp = Split(rsTmp!����, "|")
        For l = LBound(arrTmp) To UBound(arrTmp)
            arrItem = Split(arrTmp(l), ",")
            For k = LBound(arrItem) To UBound(arrItem)
                If InStr(strObjects & ",", "," & arrItem(k) & ",") = 0 Then
                    strObjects = strObjects & "," & arrItem(k)
                End If
            Next
        Next
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    GetReportObjects = strObjects
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ReportGrantToNavigator()
'���ܣ�������ǰ����(��)������̨,���ܲ��ǵ�һ��
'˵������֧�ַ������б���֧�ַ�����ѡ��ȫѡ����

    Dim lngID  As Long, lngReportID As Long, lngGroupID As Long, lngProgID As Long, lngTemp As Long
    Dim l As Long, k As Long, lngSelRow As Long, lngMenuID As Long, lngMenuPID As Long
    Dim blnGroup As Boolean, blnResult As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim rsReport As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, strName As String, strObject As String, strTmp As String
    Dim strDesc As String, strOwner As String
    Dim objNode As Node
    Dim arrTmp As Variant
    Dim colSQL As Collection
    
    mblnReportControlFocus = False
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    'ѡ��һ���˵�λ��
    Set rsTmp = GetMainTreeMenu
    If rsTmp Is Nothing Then
        MsgBox "��ȡ�˵���ϵʱ�����쳣�������������жϣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    Load frmSelTree
    If blnGroup Then
        frmSelTree.Caption = "���������鵽����̨ - �˵�λ��ѡ��"
    Else
        frmSelTree.Caption = "������������̨ - �˵�λ��ѡ��"
    End If
    
    With frmSelTree.tvw
        .Nodes.Clear
        Do While rsTmp.EOF = False
            If rsTmp!Flag = 0 Then
                Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!����, "Root")
                objNode.Tag = "��ѡ��ϵͳ��һ������Ĳ˵�λ�ã�"
            Else
                If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                    Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, rsTmp!����, "Path")
                Else
                    Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, rsTmp!���� _
                                        , IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                    objNode.ForeColor = vbBlue
                    objNode.Tag = "�����ѷ����ı���,ѡ��һ���˵�λ�ã�"

                    '���ܷ�������ͬλ��
                    If blnGroup Then
                        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����"))
                        If objNode.Text = strName Then
                            objNode.Parent.Tag = "ͬһ��������鲻�ܷ�������ͬ��λ��,��ѡ�������˵�λ�ã�"
                        End If
                    Else
                        strName = vsfTemp.TextMatrix(vsfTemp.Row, vsfTemp.ColIndex("����"))
                        If objNode.Text = strName Then
                            objNode.Parent.Tag = "ͬһ��������鲻�ܷ�������ͬ��λ��,��ѡ�������˵�λ�ã�"
                        End If
                    End If
                End If
            End If
            objNode.Expanded = True
            
            rsTmp.MoveNext
        Loop
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
        rsTmp.Close
    End With
    
    frmSelTree.Show vbModal, Me
    If Not gblnOK Then Exit Sub
    lngMenuPID = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2))      '�˵����ϼ�ID
    Unload frmSelTree
    
    'ѡ���1..N����򱨱���
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        Set colSQL = New Collection
        lngMenuID = GetNextID("zlMenus")    '�˵�ID
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        If blnGroup Then
            '������
            
            '���
            lngGroupID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
            strSQL = _
                "Select a.Id, a.���� " & vbCrLf & _
                "From zlReports A, zlRPTSubs B " & vbCrLf & _
                "Where a.Id = b.����id And b.��id = [1] "
            Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ������ı���ID", lngGroupID)
            Do While rsTmp.EOF = False
                lngReportID = Nvl(rsTmp!id, 0)
                strName = Nvl(rsTmp!����)
                GoSub proVerify
                If blnResult = False Then
                    Exit Sub
                End If
                rsTmp.MoveNext
            Loop
            rsTmp.Close
            
            '������֯
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
            strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("˵��"))
            lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����ID")))
            If lngProgID = 0 Then
                '�״η���
                lngProgID = GetNewProgID()          '����ID
                
                strSQL = _
                    "Update zlRPTSubs A Set ����=(Select ���� From zlReports Where ID=A.����ID) Where ��ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Update zlRPTGroups Set ����ID=" & lngProgID & ",����ʱ��=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlPrograms(���,����,˵��,ϵͳ,����) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    "(" & lngProgID & _
                    ",'" & strName & "'" & _
                    "," & IIF(strDesc = "", "Null", "'" & strDesc & "'") & _
                    "," & IIF(glngSys <= 0, "Null", glngSys) & _
                    ",'zl9Report' ) "
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlProgFuncs(ϵͳ,���,����,˵��)" & vbCrLf & _
                    "Select " & _
                    IIF(glngSys <= 0, "Null", glngSys) & _
                    "," & lngProgID & ",����,˵�� From zlReports " & vbCrLf & _
                    "Where ID In (Select ����ID From zlRPTSubs Where ��ID=" & lngID & ")"
                Call AddArray(colSQL, strSQL)
                
                'ѭ����ȡ���ӱ����Ȩ��
                strSQL = "Select A.����ID,B.���� From zlRPTSubs A,zlReports B Where A.��ID=[1] And A.����ID=B.ID"
                Set rsReport = OpenSQLRecord(strSQL, "��ȡ�������ӱ�����Ϣ", lngID)
                Do While rsReport.EOF = False
                    strObject = GetReportObjects(rsReport!����ID)
                    If strObject <> "" Then
                        '����п��ܲ��������ݿ�
                        strObject = Mid(strObject, 2)
                        arrTmp = Split(strObject, ",")
                        For k = LBound(arrTmp) To UBound(arrTmp)
                            strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                            If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                                strTmp = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                                strSQL = GetInsertProgPrivs(glngSys, lngProgID, rsReport!����, strTmp, strOwner, "SELECT")
                                Call AddArray(colSQL, strSQL)
                            End If
                        Next
                    End If
    
                    rsReport.MoveNext
                Loop
                rsReport.Close
            Else
                '���·���
                strSQL = "Update zlRPTGroups Set ����ʱ�� = Sysdate Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
        Else
            '���������ӱ���
            
            '���
            lngReportID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
            GoSub proVerify
            If blnResult = False Then
                Exit Sub
            End If
            
            strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
            strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("˵��"))
            lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����ID")))
            If lngProgID = 0 Then
                '�״η���
                
                '�����ɳ���ID
                strObject = GetReportObjects(lngID)
                If strObject <> "" Then strObject = Mid(strObject, 2)
                lngProgID = GetNewProgID()
                
                strSQL = _
                    "Update zlReports " & vbCrLf & _
                    "Set ���� = '����', ����ID = " & lngProgID & ", ����ʱ�� = Sysdate " & vbCrLf & _
                    "Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlPrograms(���,����,˵��,ϵͳ,����) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    " (" & lngProgID & _
                    ",'" & strName & "'" & _
                    ",'" & strDesc & "'" & _
                    ", " & IIF(glngSys <= 0, "Null", glngSys) & _
                    ",'zl9Report' )"
                Call AddArray(colSQL, strSQL)
                
                strSQL = _
                    "Insert Into zlProgFuncs(ϵͳ,���,����) " & vbCrLf & _
                    "Values" & vbCrLf & _
                    "(" & IIF(glngSys <= 0, "Null", glngSys) & _
                    "," & lngProgID & _
                    ",'����')"
                Call AddArray(colSQL, strSQL)
                
                If strObject <> "" Then
                    '�ñ���п��ܲ��������ݿ�
                    arrTmp = Split(strObject, ",")
                    For k = LBound(arrTmp) To UBound(arrTmp)
                        strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                        If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                            strTmp = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                            strSQL = GetInsertProgPrivs(glngSys, lngProgID, "����", strTmp, strOwner, "SELECT")
                            Call AddArray(colSQL, strSQL)
                        End If
                    Next
                End If
               
            Else
                '�ظ�����
                strSQL = "Update zlReports Set ����ʱ�� = Sysdate Where ID = " & lngID
                Call AddArray(colSQL, strSQL)
            End If
        End If
        
        '���ɲ˵�����
        strSQL = _
            "Insert Into zlMenus(���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) " & vbCrLf & _
            "Values('ȱʡ'," & lngMenuID & "," & lngMenuPID & ",'" & strName & "',Null" & _
            "," & IIF(strDesc = "", "Null", "'" & strDesc & "'") & _
            "," & IIF(glngSys <= 0, "Null", glngSys) & _
            "," & lngProgID & ",'" & strName & "',105)"
        Call AddArray(colSQL, strSQL)
        
        'ִ��DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            'Debug.Print colSQL(k)
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1

makContinue:
    Next
    
    Set grsReport = Nothing
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx

    Exit Sub
    
proVerify:
    blnResult = False
    
    '��֤����
    If CheckPass(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("��[1]��������֤����ͨ�����ܾ�������", strName) _
            , vbInformation, App.Title
        Return
    End If
    
    'Ȩ��
    If CheckReportPriv(lngReportID) = False Then
        MsgBox mdlPublic.FormatString("��û�С�[1]������������Դ�漰���ݿ����Ĳ�ѯȨ�ޣ����飡", strName) _
            , vbInformation, App.Title
        Return
    End If
    
    blnResult = True
    Return
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then Resume
    End If
    Call SaveErrLog
    
    Set grsReport = Nothing
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
End Sub

Private Function GetMainTreeMenu(Optional ByVal lngProgID As Long) As ADODB.Recordset
'���ܣ���ȡ����������̨�������β˵���ϵ
'������lngProgID=�Ƿ�ֻ��ʾָ������ID�ı���
'˵�����˵���ϵ�а����Զ��屨�����Ĳ˵���(�����),��־Ϊ"FLAG=999"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngSys As Long
    
    On Error GoTo errH
    
    lngSys = glngSys
    If lngSys = 0 Then
        'ֻ��ʾ�û��������ݱ���
        strSQL = _
            "Select Distinct * From (" & _
            " Select ��� as SCOL,0 as Flag,-��� as ID,-NULL as �ϼ�ID,'['||���||']'||���� as ����,-NULL as ģ�� " & _
            " From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(�ϼ�ID,-ϵͳ) as �ϼ�ID,����,ģ�� " & _
            " From zlMenus Where ���='ȱʡ' And ģ�� is NULL" & _
            " Start With �ϼ�ID is NULL And ���='ȱʡ' Connect by Prior ID=�ϼ�ID And ���='ȱʡ'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            "     And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            "     And A.ϵͳ is NULL And B.ϵͳ is Null And C.ϵͳ is Null" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlReports C" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            "     And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            "     And A.ϵͳ is NULL And B.ϵͳ is Null And C.ϵͳ is Null" & _
            ") Order by SCOL,FLAG,ID"
    Else
        'ֻ��ʾ�̶����ݱ���(����Ȩ����)
        strSQL = _
            "Select Distinct * From (" & _
            " Select ��� as SCOL,0 as Flag,-��� as ID,-NULL as �ϼ�ID,'['||���||']'||���� as ����,-NULL as ģ�� " & _
            " From zlSystems Union ALL" & _
            " Select 99999 as SCOL,Level as FLAG,ID,Nvl(�ϼ�ID,-ϵͳ) as �ϼ�ID,����,ģ�� " & _
            " From zlMenus Where ���='ȱʡ' And ģ�� is NULL" & _
            " Start With �ϼ�ID is NULL And ���='ȱʡ' Connect by Prior ID=�ϼ�ID And ���='ȱʡ'" & _
            " Union ALL" & _
            " Select 99999 as SCOL,999 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlRPTGroups C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            "     And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            "     And A.ϵͳ=B.ϵͳ And A.ϵͳ=C.ϵͳ And Trunc(B.ϵͳ/100)=D.ϵͳ And B.���=D.���" & _
            " Union ALL" & _
            " Select 99999 as SCOL,888 as FLAG,A.ID,A.�ϼ�ID,A.����,A.ģ��" & _
            " From zlMenus A,zlPrograms B,zlReports C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & _
            " Where A.ģ��=B.��� And A.���='ȱʡ' And C.����ID=A.ģ�� " & _
            "     And Upper(B.����)='ZL9REPORT'" & IIF(lngProgID = 0, "", " And B.���=[1]") & _
            "     And A.ϵͳ=B.ϵͳ And A.ϵͳ=C.ϵͳ And Trunc(B.ϵͳ/100)=D.ϵͳ And B.���=D.���" & _
            ") Order by SCOL,FLAG,ID"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngProgID)
    Set GetMainTreeMenu = rsTmp
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewProgID() As Long
'���ܣ���ȡ��һ�����õ��Զ��屨������,���ڷ���
'˵��������Ŵ�100000��ʼ,���Զ���ȱ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Decode(Sign(Max(���)-99999),1,Max(���),99999) as ID From zlPrograms"
    Set rsTmp = mdlPublic.OpenSQLRecord(strSQL, "��ȡ��ģ��ID���Զ��屨��")
    GetNewProgID = IIF(IsNull(rsTmp!id), 100000, rsTmp!id + 1)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReportRevokeFromNavigator()
'���ܣ�ȡ����ǰ����򱨱����ڵ���̨�ϵķ���
'˵����
'  1.����������ڶ��λ�ã�����ʹ����ѡ��ȡ��������һ��λ�ã�ɾ��zlMenus��Ӧλ������
'  2.���ֻ��һ������λ�ã���zlReport�еĳ���ID=NULL��ɾ��zlPrograms�еķ���ģ��
'������
''''''  blnRevokeByProgram��True-ȡ��ģ�飻False-ȡ������̨

    Dim rsTmp As ADODB.Recordset
    Dim objNode As Node
    Dim lngID As Long, lngProgID As Long, lngMenuID As Long, lngSelRow As Long
    Dim l As Long, k As Long
    Dim blnTrans As Boolean, blnGroup As Boolean
    Dim colSQL As Collection
    Dim vsfTemp As VSFlexGrid
    Dim strSQL As String, strName As String, strMsg As String, strTemp As String
    
    mblnReportControlFocus = False
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Then
        MsgBox "��ѡ��һ���������������顢�ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    If UCase(vsfTemp.name) = UCase("vsfGroup") Then
        strMsg = GetSelectedReport(vsfTemp, "����")
    Else
        strMsg = GetSelectedReport(vsfTemp, "����")
    End If
    If MsgBox("��������±���ӵ���̨�˵���ȡ�������������û�������ʹ�øñ���Ҫ������" & _
        strMsg, vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    'ѡ���1..N����򱨱���
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        '��������λ��
        Set colSQL = New Collection
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        lngProgID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����ID")))
        If lngProgID <= 0 Then
            If blnGroup Then
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
                strTemp = mdlPublic.FormatString("��[1]��������û�з���������̨�˵���", strName)
            Else
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
                strTemp = mdlPublic.FormatString("��[1]������û�з���������̨�˵���", strName)
            End If
            If lngSelRow < vsfTemp.SelectedRows - 1 Then
                strTemp = strTemp & "Ҫ������ȡ��������ʣ�µı�����"
                If MsgBox(strTemp, vbInformation + vbDefaultButton1 + vbYesNo) = vbYes Then
                    lngSelRow = lngSelRow + 1
                    GoTo makContinue
                Else
                    Exit Sub
                End If
            Else
                MsgBox strTemp, vbInformation, App.Title
                Exit Sub
            End If
        Else
            If blnGroup Then
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
            Else
                strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
            End If
        End If
        
        '��ȡģ����Ϣ
        Set rsTmp = GetMainTreeMenu(lngProgID)
        rsTmp.Filter = "ģ��=" & lngProgID
        If rsTmp Is Nothing Then
            MsgBox "��ȡ�˵���ϵʱ�����쳣����ȡ���������жϣ�", vbInformation, App.Title
            Exit Sub
        ElseIf rsTmp.EOF Or rsTmp.RecordCount = 1 Then
            If rsTmp.EOF Then
                '����λ���쳣
                strTemp = mdlPublic.FormatString("��[1]������ķ������ڲ�����״̬�����������ݲ���ȷ����", strName)
                MsgBox strTemp, vbInformation, App.Title
            End If
            
            If blnGroup Then
                '������
                strSQL = "Update zlRPTGroups Set ����ID=NULL,����ʱ��=NULL,�Ƿ�ͣ��=NULL Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Update zlRPTSubs A Set ����=Null Where ��ID=" & lngID
                Call AddArray(colSQL, strSQL)
            Else
                '����
                strSQL = _
                    "Update zlReports Set ����=NULL,����ID=NULL,�Ƿ�ͣ��=NULL,����ʱ��=NULL " & vbCrLf & _
                    "Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
            strSQL = "Delete From zlMenus Where ģ��=" & lngProgID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlPrograms Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
        
            strSQL = "Delete From zlRoleGrant Where ���=" & lngProgID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
            
        Else
            '�������λ��
            rsTmp.Filter = ""
            
            '�����˵����ṹ
            Load frmSelTree
            frmSelTree.Caption = "ȡ������ - ����̨�˵�λ��"
            
            With frmSelTree.tvw
                .Nodes.Clear
                Do While rsTmp.EOF = False
                    If rsTmp!Flag = 0 Then
                        Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!����, "Root")
                        objNode.Tag = "���ڱ�ϵͳ��ѡ��һ��Ҫȡ�������ı�����飡"
                    Else
                        If rsTmp!Flag <> 999 And rsTmp!Flag <> 888 Then
                            Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, rsTmp!����, "Path")
                            objNode.Tag = "���ڲ˵���ѡ��һ��Ҫȡ�������ı�����飡"
                        Else
                            Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, rsTmp!����, IIF(rsTmp!Flag = 999, "GroupNode", "ReportNode"))
                            objNode.ForeColor = vbBlue
                            If .SelectedItem Is Nothing Then
                                objNode.Selected = True
                            ElseIf .SelectedItem.Index = 1 Then
                                objNode.Selected = True
                            End If
                        End If
                    End If
                    objNode.Expanded = True

                    '����б���(��)��·��
                    If rsTmp!Flag = 999 Or rsTmp!Flag = 888 Then
                        objNode.SelectedImage = objNode.Image
                        Do While Not objNode.Parent Is Nothing
                            Set objNode = objNode.Parent
                            objNode.SelectedImage = objNode.Image
                        Loop
                    End If
    
                    rsTmp.MoveNext
                Loop
    
                'ɾ��û�б���(��)��·��
                For k = .Nodes.count To 1 Step -1
                    If .Nodes(k).SelectedImage = "" Then
                        .Nodes.Remove k
                    End If
                Next
            End With
            
            frmSelTree.Show vbModal, Me
            If Not gblnOK Then Exit Sub
            lngMenuID = CLng(Mid(frmSelTree.tvw.SelectedItem.Key, 2))     '����˵�ID
            Unload frmSelTree
            
            If blnGroup Then
                strSQL = "Update zlRPTGroups Set ����ʱ��=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
                
                strSQL = "Update zlRPTSubs A Set ����=Null Where ��ID=" & lngID
                Call AddArray(colSQL, strSQL)
            Else
                strSQL = "Update zlReports Set ����ʱ��=Sysdate Where ID=" & lngID
                Call AddArray(colSQL, strSQL)
            End If
            
            strSQL = "Delete From zlMenus Where ID=" & lngMenuID & " And Nvl(ϵͳ,0)=" & glngSys
            Call AddArray(colSQL, strSQL)
        End If
        
        'ִ��DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
'            Debug.Print colSQL(k)
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1
        
makContinue:
    Next
    
    Set grsReport = Nothing
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub

hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ReportGrantToModule()
'���ܣ�������ǰ����ģ��,���ܲ��ǵ�һ��
'˵������֧�ֱ����鷢����ģ��

    Dim l As Long, j As Long, k As Long
    Dim lngSelRow As Long, lngID As Long, lngSys As Long, lngProgID As Long
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim rsModule As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strFunc As String, strIDs As String, strSQL As String, strObject As String
    Dim strDesc As String, strOwner As String, strName As String
    Dim objNode As Node
    Dim colSQL As Collection
    Dim arrTmp As Variant, arrItems As Variant
    
    mblnReportControlFocus = False
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    lngSelRow = 0
    strIDs = ""
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue1
        
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        If CheckPass(lngID) = False Then
            MsgBox "�������ݴ��󣬲��ܷ����ñ���", vbInformation, App.Title
            Exit Sub
        End If
        If Not CheckReportPriv(lngID) Then
            MsgBox "��û��Ȩ�޲�ѯ�ñ���ĳЩ����Դ�еĶ��󣬲������ܼ�����", vbInformation, App.Title
            Exit Sub
        End If
        
        strIDs = strIDs & "," & CStr(lngID)
        lngSelRow = lngSelRow + 1
        
makContinue1:
    Next
    If strIDs = "" Then Exit Sub
    
    strIDs = Mid(strIDs, 2)
    Set rsModule = GetModuleTreeMenu(strIDs)
    If rsModule Is Nothing Then
        MsgBox "��ȡģ��˵���ϵʱ����������󣬱������жϣ�", vbInformation, App.Title
        Exit Sub
    End If
    
    '1.ѡ��һ���˵�ģ��λ��
    Load frmSelTree
    frmSelTree.Caption = "��������ģ�� - ģ��λ��ѡ��"
    With frmSelTree.tvw
        .Nodes.Clear
        Do While rsModule.EOF = False
            If IsNull(rsModule!�ϼ�ID) Then
                Set objNode = .Nodes.Add(, , "_" & rsModule!id, rsModule!����)
            Else
                Set objNode = .Nodes.Add("_" & rsModule!�ϼ�ID, 4, "_" & rsModule!id, rsModule!����)
            End If
            
            If Left(rsModule!id, 1) = "S" Then     'System
                objNode.Image = "Root"
                objNode.Tag = "��ѡ��ϵͳ�в˵��µ�ģ��λ�á�"
            ElseIf Left(rsModule!id, 1) = "T" Then 'MenuTree
                objNode.Image = "Path"
                objNode.Tag = "��ѡ��ϵͳ�в˵��µ�ģ��λ�á�"
            ElseIf Left(rsModule!id, 1) = "M" Then 'Module
                objNode.Image = "App"
            ElseIf Left(rsModule!id, 1) = "R" Then 'Report
                objNode.Image = "ReportNode"
                objNode.ForeColor = vbBlue
                objNode.Tag = "�����ѷ����ı���,ѡ�������˵��µ�ģ��λ�á�"
                objNode.Parent.Tag = "�������ظ�������ͬһ��ģ��,��ѡ������ģ�顣"
            End If
            objNode.Expanded = True
            
            '������¼�ģ��Ĳ˵�(��SQL����)
            If Left(rsModule!id, 1) = "M" Then
                If objNode.Parent.SelectedImage = "" Then
                    Do While Not objNode.Parent Is Nothing
                        Set objNode = objNode.Parent
                        objNode.SelectedImage = objNode.Image
                    Loop
                End If
            End If
            
            rsModule.MoveNext
        Loop
        
        'ɾ�����¼�ģ��Ŀղ˵�
        For k = .Nodes.count To 1 Step -1
            If .Nodes(k).SelectedImage = "" And Mid(.Nodes(k).Key, 2, 1) = "T" Then
                .Nodes.Remove k
            End If
        Next
        
        If .Nodes.count > 0 Then .Nodes(1).Selected = True
    End With
    
    frmSelTree.Show vbModal, Me
    If Not gblnOK Then
        Exit Sub
    End If
    rsModule.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
    If rsModule.EOF Then Exit Sub
    
    '�����ظ����
    lngSys = rsModule!ϵͳ
    lngProgID = rsModule!����id
    Unload frmSelTree
    rsModule.Close
    
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue2

        strFunc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strSQL = _
            "Select ���� From zlRPTPuts Where ����ID=[1] And ϵͳ=[2] And ����ID=[3] " & vbCrLf & _
            "Union ALL " & vbCrLf & _
            "Select ���� From zlProgFuncs Where ϵͳ=[2] And ���=[3] And ����=[4] "
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngID, lngSys, lngProgID, strFunc)
        If Not rsTmp.EOF Then
            MsgBox "������λ�û򷢲������ظ������ݿ��е����ݿ��ܲ���ȷ��", vbInformation, App.Title
            Exit Sub
        End If
        rsTmp.Close

        lngSelRow = lngSelRow + 1

makContinue2:
    Next

    '2.��ȨȨ�޷���

    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue3

        strObject = ""
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strFunc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
        strDesc = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("˵��"))

        strObject = GetReportObjects(lngID)
        If strObject <> "" Then strObject = Mid(strObject, 2)
        
        '3.��д����Ȩ��
        Set colSQL = New Collection
        
        strSQL = "Update zlReports Set ����ʱ��=Sysdate Where ID=" & lngID
        Call AddArray(colSQL, strSQL)
        
        strSQL = _
            "Insert Into zlRPTPuts(����ID,ϵͳ,����ID,����) " & vbCrLf & _
            "Values(" & lngID & _
            "," & lngSys & _
            "," & lngProgID & _
            ",'" & strFunc & "')"
        Call AddArray(colSQL, strSQL)
        
        strSQL = _
            "Insert Into zlProgFuncs(ϵͳ,���,����,˵��) " & vbCrLf & _
            "Values(" & lngSys & _
            "," & lngProgID & _
            ",'" & strFunc & "'" & _
            ",'" & strDesc & "')"
        Call AddArray(colSQL, strSQL)
        
        If strObject <> "" Then
            '�ñ���п��ܲ��������ݿ�
            arrTmp = Split(strObject, ",")
            For k = 0 To UBound(arrTmp)
                strOwner = Left(arrTmp(k), InStr(arrTmp(k), ".") - 1)
                If strOwner <> "SYS" And strOwner <> "ZLTOOLS" And strOwner <> "SYSTEM" Then
                    strName = Mid(arrTmp(k), InStr(arrTmp(k), ".") + 1)
                    strSQL = GetInsertProgPrivs(lngSys, lngProgID, strFunc, strName, strOwner, "SELECT")
                    Call AddArray(colSQL, strSQL)
                End If
            Next
        End If
        
        'ִ��DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False

        lngSelRow = lngSelRow + 1

makContinue3:
    Next
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetModuleTreeMenu(ByVal strIDs As String) As ADODB.Recordset
'���ܣ���ȡ������ģ��ı������β˵���ϵ
'������
'  strIDs��Ҫ������ȡ�������ı���ID����ʽ��id1[,id2[...]]

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '���˵���ʾģ��ķ�ʽ
    '-------------------------------------------------------------------------------------------------------------
    '1.ϵͳ + 2.�м�˵� + 3.ģ��˵�(��Ȩģ��) + 4.��������(��������Ȩģ����)
    'ע��ͬһģ������ظ�λ�ڲ�ͬ�˵�,����ʾ(��������ı���)
    '�ſ��������Զ��屨��ģ��(����='zl9Report')
    
    'ֻ��ʾģ��ķ�ʽ
    strSQL = _
        " Select '1' as Sort1,To_Char(���) as Sort2, 'S'||��� as ID,Null as �ϼ�ID,��� as ϵͳ" & vbCrLf & _
        "     ,-Null as ����ID,Null as ����,'['||���||']'||���� as ����" & vbCrLf & _
        " From zlSystems" & vbCrLf & _
        " Union ALL " & vbCrLf & _
        " Select '3' as Sort1,To_Char(B.���) as Sort2, 'M'||B.���||'_'||B.ϵͳ as ID,'S'||B.ϵͳ as �ϼ�ID,B.ϵͳ" & vbCrLf & _
        "     ,B.��� as ����ID,Null as ����,'['||B.���||']'||B.����" & vbCrLf & _
        " From zlPrograms B,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) C" & vbCrLf & _
        " Where Upper(B.����)<>Upper('zl9Report') And Trunc(B.ϵͳ/100)=C.ϵͳ And B.���=C.���" & vbCrLf & _
        " Union All " & vbCrLf & _
        " Select '4' as Sort1,C.��� as Sort2," & vbCrLf & _
        "   'R'||Rownum as ID,'M'||B.����ID||'_'||B.ϵͳ as �ϼ�ID,B.ϵͳ,B.����ID,B.����,'['||C.���||']'||C.���� as ����" & vbCrLf & _
        " From zlPrograms A,zlRPTPuts B,zlReports C,(Select ϵͳ,��� From zlRegFunc Group By ϵͳ,���) D" & vbCrLf & _
        "    , Table(Cast(f_Num2List([1]) as T_NumList)) E" & vbCrLf & _
        " Where A.ϵͳ=B.ϵͳ And A.���=B.����ID And Upper(A.����)<>Upper('zl9Report')" & vbCrLf & _
        "   And Trunc(A.ϵͳ/100)=D.ϵͳ And A.���=D.���" & vbCrLf & _
        "   And B.����ID=C.ID And c.Id = e.Column_Value " & vbCrLf & _
        " Order by Sort1,Sort2"
    
    '�̶������������¡��ɱ�������ϵͳ��ģ�飬����ϵͳ��10�汾����
    strSQL = "Select A.* From (" & strSQL & ") A,zlSystems B" & _
        " Where A.ϵͳ=B.��� And (To_Number(Substr(B.�汾��,1,Instr(B.�汾��,'.')-1))>=10 Or Trunc(���/100) IN(2,5,7))" & _
        " Order by Sort1,Sort2"
    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, strIDs)
    Set GetModuleTreeMenu = rsTmp
    Exit Function
    
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetSelectedReport(ByVal vsfVar As VSFlexGrid, ByVal strColName As String) As String
    Dim strResult As String
    Dim l As Long, lngSelRow As Long
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For l = 1 To vsfVar.Rows - 1
        If vsfVar.SelectedRow(lngSelRow) = l Then
            strResult = strResult & vbCrLf & CStr(lngSelRow + 1) & "." & vsfVar.TextMatrix(l, vsfVar.ColIndex(strColName))
            lngSelRow = lngSelRow + 1
        End If
        If lngSelRow >= 5 Then
            '�����ʾ5����Ϣ
            strResult = strResult & vbCrLf & "..."
            Exit For
        End If
    Next
    GetSelectedReport = strResult
    Exit Function
    
hErr:
    Call ErrCenter
End Function

Private Sub ReportRevokeFromModule()
'���ܣ�ȡ����ǰ������ģ���ϵ�һ������
'˵����
'  1.�������λ�ô���1������ʹ����ѡ��ȡ��������һ��λ��
'  2.���ֻ��һ������λ�ã���ֱ����ʾ����

    Dim rsTmp As ADODB.Recordset
    Dim strFunc As String, strSQL As String, strMsg As String, strName As String
    Dim strTmp As String
    Dim objNode As Node
    Dim lngID As Long, lngSys As Long, lngProgID As Long, lngSelRow As Long
    Dim l As Long, k As Long
    Dim blnTrans As Boolean, blnGroup As Boolean
    Dim vsfTemp As VSFlexGrid
    Dim colSQL As Collection
    
    mblnReportControlFocus = False
    
    '���
    If GetVsfControl(lngID, blnGroup, vsfTemp) = False Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    If vsfTemp.Row <= 0 Or blnGroup Then
        MsgBox "��ѡ��һ�����������ӱ���", vbInformation, App.Title
        Exit Sub
    End If
    
    If UCase(vsfTemp.name) = UCase("vsfGroup") Then
        strMsg = GetSelectedReport(vsfTemp, "����")
    Else
        strMsg = GetSelectedReport(vsfTemp, "����")
    End If
    If MsgBox("��������±���Ӹ�ģ��˵���ȡ�������������û�������ʹ�øñ���Ҫ������" & _
        strMsg, vbInformation + vbDefaultButton2 + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    lngSelRow = 0
    For l = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) <> l Then GoTo makContinue
        
        lngID = Val(vsfTemp.TextMatrix(l, vsfTemp.ColIndex("ID")))
        strName = vsfTemp.TextMatrix(l, vsfTemp.ColIndex("����"))
        strFunc = strName
        Set colSQL = New Collection
        
        '1.������ǰ����λ��
        strSQL = "Select ϵͳ,����ID,���� From zlRPTPuts Where ����ID=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngID)
        If rsTmp.EOF Then
            rsTmp.Close
            strTmp = mdlPublic.FormatString("��ǰ����[1]��û�з�����ģ����", strName)
            If vsfTemp.SelectedRows - 1 <= lngSelRow Then
                strTmp = strTmp & "��"
                MsgBox strTmp, vbInformation, App.Title
                rptClass.Tag = ""
                Call RefreshEx
                Exit Sub
            Else
                strTmp = strTmp & "��Ҫ������ȡ��������ʣ�µı�����"
                If MsgBox(strTmp, vbInformation + vbDefaultButton1 + vbYesNo, App.Title) = vbYes Then
                    lngSelRow = lngSelRow + 1
                    GoTo makContinue
                Else
                    rptClass.Tag = ""
                    Call RefreshEx
                    Exit Sub
                End If
            End If
        ElseIf rsTmp.RecordCount = 1 Then
            'ֻ��һ������λ��
            lngSys = rsTmp!ϵͳ
            lngProgID = rsTmp!����id
            rsTmp.Close
            
            strSQL = "Update zlReports Set ����ʱ��=NULL, �Ƿ�ͣ��=NULL Where ����ID Is Null And ID=" & lngID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRPTPuts Where ����ID=" & lngID & " And ϵͳ=" & lngSys & " And ����ID=" & lngProgID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRoleGrant Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
        
        Else
            '���ж������λ�ã�ѡ����ȡ��
            Set rsTmp = GetModuleTreeMenu(lngID)
            If rsTmp Is Nothing Then
                MsgBox "��ȡģ��˵���ϵʱ����������󣬱���ȡ���������жϣ�", vbInformation, App.Title
                Exit Sub
            End If
            
            Load frmSelTree
            frmSelTree.Caption = "ȡ������ - ģ��˵�λ��"
            With frmSelTree.tvw
                .Nodes.Clear
                Do While rsTmp.EOF = False
                    If IsNull(rsTmp!�ϼ�ID) Then
                        Set objNode = .Nodes.Add(, , "_" & rsTmp!id, rsTmp!����)
                    Else
                        Set objNode = .Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!id, rsTmp!����)
                    End If
                    If Left(rsTmp!id, 1) = "S" Then     'System
                        objNode.Image = "Root"
                        objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                    ElseIf Left(rsTmp!id, 1) = "T" Then 'MenuTree
                        objNode.Image = "Path"
                        objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                    ElseIf Left(rsTmp!id, 1) = "M" Then 'Module
                        objNode.Image = "App"
                        objNode.Tag = "��ѡ��Ҫȡ�������ı���"
                    ElseIf Left(rsTmp!id, 1) = "R" Then 'Report
                        objNode.Image = "ReportNode"
                        objNode.ForeColor = vbBlue
                    End If
                    objNode.Expanded = True
                    
                    '����з���������ϼ�
                    If Left(rsTmp!id, 1) = "R" Then
                        objNode.SelectedImage = objNode.Image
                        If objNode.Parent.SelectedImage = "" Then
                            Do While Not objNode.Parent Is Nothing
                                Set objNode = objNode.Parent
                                objNode.SelectedImage = objNode.Image
                            Loop
                        End If
                    End If
                    
                    rsTmp.MoveNext
                Loop
                
                'ɾ���޷��������·��
                For k = .Nodes.count To 1 Step -1
                    If .Nodes(k).SelectedImage = "" Then
                        .Nodes.Remove k
                    End If
                Next
                
                If .Nodes.count > 0 Then .Nodes(1).Selected = True
            End With
            
            frmSelTree.Show vbModal, Me
            If Not gblnOK Then
                Exit Sub
            End If
            rsTmp.Filter = "ID='" & Mid(frmSelTree.tvw.SelectedItem.Key, 2) & "'"
            If rsTmp.EOF Then
                MsgBox "��ȡģ��˵������쳣������ȡ���������жϣ�", vbInformation, App.Title
                Exit Sub
            End If
            
            lngSys = rsTmp!ϵͳ
            lngProgID = rsTmp!����id
            Unload frmSelTree
            
            strSQL = "Delete From zlRPTPuts Where ����ID=" & lngID & " And ϵͳ=" & lngSys & " And ����ID=" & lngProgID
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgPrivs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlProgFuncs Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
            
            strSQL = "Delete From zlRoleGrant Where ϵͳ=" & lngSys & " And ���=" & lngProgID & " And ����='" & strFunc & "'"
            Call AddArray(colSQL, strSQL)
        End If
        
        'ִ��DML
        gcnOracle.BeginTrans: blnTrans = True
        For k = 1 To colSQL.count
            gcnOracle.Execute colSQL(k)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        lngSelRow = lngSelRow + 1
        
makContinue:
    Next
    
    Set grsReport = Nothing '�������
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub

hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Find(ByVal strText As String, Optional ByVal blnFirst As Boolean = False)
'���ܣ�����ƥ��������λ
'������
'  strText��������ƥ����ı�
'  blnFirst��True�״β��ң�False�ٴβ���

    Dim lngRow As Long, l As Long
    Dim arrOrder As Variant
    Dim i As Integer, j As Integer, intIndex As Integer
    Dim blnStart As Boolean
    Dim colOrder As Collection
    Dim strName As String
    
    If Trim(strText) = "" Then Exit Sub
    
    '׼�������ҵĶ�����к�
    If blnFirst Or mobjFindPreviors Is Nothing Then
        '�״�
        lngRow = 0
        blnFirst = True
    Else
        '�ٴ�
        lngRow = mlngFindRow
    End If
    
    Set colOrder = New Collection
    If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
        If blnFirst Then
            strName = "VSFREPORT"
        Else
            strName = UCase(mobjFindPreviors.name)
        End If
        '������ҵĶ���
        colOrder.Add Me.vsfReport
        colOrder.Add Me.rptClass
    Else
        If blnFirst Then
            strName = "VSFGROUP"
        Else
            strName = UCase(mobjFindPreviors.name)
        End If
        '������ҵĶ���
        colOrder.Add Me.vsfGroup
        colOrder.Add Me.vsfGroupDetail
        colOrder.Add Me.rptClass
    End If
    
    '����
    blnStart = False
    For i = 1 To colOrder.count
        With colOrder(i)
            If .Visible = False Then GoTo makContinue
            
            If UCase(.name) = strName Or blnStart = True Then
                '��ʼ����
                If UCase(.name) = "RPTCLASS" Then
                    '�������
                    For l = lngRow + IIF(blnStart, 0, 1) To .Rows.count - 1
                        For j = 0 To .Columns.count - 1
                            If UCase(.Rows(l).Record(j).Value) Like "*" & UCase(strText) & "*" _
                                And .Columns(j).Visible Then
                                '�ҵ�
                                .Rows(l).Selected = True
                                .SetFocus
                                lngRow = l
                                GoTo makFinish
                            End If
                        Next
                    Next
                Else
                    '�������������顢�ӱ���
                    For l = lngRow + 1 To .Rows - 1
                        For j = 0 To .Cols - 1
                            If UCase(.TextMatrix(l, j)) Like "*" & UCase(strText) & "*" _
                                And .ColWidth(j) > 0 Then
                                '�ҵ�
                                .TopRow = l
                                .LeftCol = j
                                .Row = l
                                .Col = j
                                .SetFocus
                                lngRow = l
                                GoTo makFinish
                            End If
                        Next
                    Next
                End If
            
                blnStart = True
                lngRow = 0
            End If
        End With
        
makContinue:
    Next
    
    Set mobjFindPreviors = Nothing
    mlngFindRow = 0
    
    If mintFind <= 0 Then
        If MsgBox("δ���ҵ�ƥ������ݣ��Ƿ��ͷ��ʼ���ң�", vbInformation + vbDefaultButton1 + vbYesNo, App.Title) = vbYes Then
            mintFind = mintFind + 1
            Call Find(strText, True)
        Else
            mintFind = 0
        End If
    Else
        mintFind = 0
    End If
    
    Exit Sub
    
makFinish:
    '������Ϣ�´�ʹ��
    Set mobjFindPreviors = colOrder(i)
    mlngFindRow = lngRow
    mintFind = 0
End Sub

Private Sub UpdateStatusBar(ByVal objFocus As Object)
'���ܣ�����״̬������ʾ��Ϣ
'������
'  objFocus���������

    Dim strMsg As String
    Dim lngID As Long

    With objFocus
        Select Case UCase(objFocus.name)
        Case "VSFGROUP"
            If mblnReportControlFocus Then Exit Sub
            
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strMsg = mdlPublic.FormatString("��[1]��[2]�������� [3] �ű���" _
                        , .TextMatrix(.Row, .ColIndex("���")) _
                        , .TextMatrix(.Row, .ColIndex("����")) _
                        , vsfGroupDetail.Rows - 1)
            If .TextMatrix(.Row, .ColIndex("����ʱ��")) <> "" Then
                strMsg = strMsg & "�� ����λ�ã�" & GetMenuPath(lngID, True)
            End If
        Case "RPTCLASS"
            If tbcRPT.Selected.Index = Val("0-����ҳ��") Then
                strMsg = mdlPublic.FormatString("��[1]���������� [2] �ű���" _
                            , .FocusedRow.Record(mobjClass.GetColIndex("����")).Value _
                            , vsfReport.Rows - 1)
            Else
                strMsg = mdlPublic.FormatString("��[1]���������� [2] �ݱ�����" _
                            , .FocusedRow.Record(mobjClass.GetColIndex("����")).Value _
                            , vsfGroup.Rows - 1)
            End If
        Case Else
            If mblnReportControlFocus Then Exit Sub
            
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
            strMsg = mdlPublic.FormatString("��[1]��[2]" _
                        , .TextMatrix(.Row, .ColIndex("���")) _
                        , .TextMatrix(.Row, .ColIndex("����")))
            If .TextMatrix(.Row, .ColIndex("����ʱ��")) <> "" Then
                strMsg = strMsg & "�� ����λ�ã�" & GetMenuPath(lngID, False)
            End If
            If .TextMatrix(.Row, .ColIndex("˵��")) <> "" Then
                strMsg = strMsg & "�� ˵����" & .TextMatrix(.Row, .ColIndex("˵��"))
            End If
        End Select
    End With
    
    Me.staMain.Panels(2).Text = strMsg
End Sub

Private Sub StateSwitch(ByVal lngID As Long, Optional ByVal blnEnabled As Boolean = False)
'���ܣ��������á�ͣ�õ��л�
'������
'  lngID���˵�ID
'  blnEnabled��True���ã�Falseͣ��

    Dim lngRow As Long, lngSelRow As Long, lngReportID As Long
    Dim vsfTemp As VSFlexGrid
    Dim blnGroup As Boolean, blnTrans As Boolean
    Dim strIDs As String, strRec As String, strNonRec As String, strName As String
    Dim strSQL  As String, strTmp As String
    Dim colSQL As New Collection
 
    If mblnReportControlFocus = False Then
        If GetVsfControl(lngID, blnGroup, vsfTemp, strIDs) = False Then
            MsgBox "��ѡ�ж������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
        If vsfTemp.Row <= 0 Then
            MsgBox "��ѡ�ж������������顢�ӱ���", vbInformation, App.Title
            Exit Sub
        End If
    End If
    
    '���
    strName = IIF(blnGroup, "����", "����")
    For lngRow = 1 To vsfTemp.Rows - 1
        If lngSelRow <= 5 Then
            If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
                If vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("����ʱ��")) = "" Then
                    If lngSelRow >= 5 Then
                        strNonRec = strNonRec & vbCrLf & "..."
                    Else
                        strNonRec = strNonRec & vbCrLf & CStr(lngSelRow + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex(strName))
                    End If
                Else
                    If lngSelRow >= 5 Then
                        strRec = strRec & vbCrLf & "..."
                    Else
                        strRec = strRec & vbCrLf & CStr(lngSelRow + 1) & "." & vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex(strName))
                    End If
                End If
                lngSelRow = lngSelRow + 1
            End If
        Else
            Exit For
        End If
    Next
    If strNonRec <> "" Then
        MsgBox "��ȷ�����±���" & IIF(blnGroup, "��", "") & "�ѷ�����" & strNonRec, vbInformation, App.Title
        Exit Sub
    End If
    
    On Error GoTo hErr
    
    '����
    strTmp = IIF(blnEnabled, "����", "ͣ��")
    strNonRec = IIF(blnGroup, "��", "")
    If MsgBox("��ȷ��Ҫ��" & strTmp & "�����б���" & strNonRec & "��" & strRec, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    lngSelRow = 0
    For lngRow = 1 To vsfTemp.Rows - 1
        If vsfTemp.SelectedRow(lngSelRow) = lngRow Then
            lngReportID = Val(vsfTemp.TextMatrix(lngRow, vsfTemp.ColIndex("ID")))
            If blnGroup Then
                '������
                strSQL = "Update zlRPTGroups " & vbCrLf & _
                         "Set �Ƿ�ͣ�� = " & IIF(blnEnabled, "Null", "1") & vbCrLf & _
                         "Where Not ����ʱ�� Is Null And ID = " & lngReportID & " "
            Else
                '����
                strSQL = "Update zlReports " & vbCrLf & _
                         "Set �Ƿ�ͣ�� = " & IIF(blnEnabled, "Null", "1") & vbCrLf & _
                         "Where Not ����ʱ�� Is Null And ID = " & lngReportID & " "
            End If
            Call AddArray(colSQL, strSQL)
            
            lngSelRow = lngSelRow + 1
        End If
    Next
    
    'ִ��DML
    gcnOracle.BeginTrans: blnTrans = True
    For lngRow = 1 To colSQL.count
        gcnOracle.Execute colSQL(lngRow)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Screen.MousePointer = vbDefault
    
    'ˢ��
    rptClass.Tag = ""
    Call RefreshEx
    
    Exit Sub
    
hErr:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
    Call ErrCenter
End Sub

Private Function GetRelationList(ByVal lngReportID As Long) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo hErr
    
    strSQL = _
        "Select a.����id, b.����, a.��������id, c.���� �������� " & vbNewLine & _
        "From Zlrptrelation A, zlReports B, zlReports C " & vbNewLine & _
        "Where a.����id = b.Id(+) And a.��������id = c.Id(+) And a.����id = [1] " & vbNewLine & _
        "Union All " & vbNewLine & _
        "Select a.����id, b.����, a.��������id, c.���� �������� " & vbNewLine & _
        "From Zlrptrelation A, zlReports B, zlReports C " & vbNewLine & _
        "Where a.����id = b.Id(+) And a.��������id = c.Id(+) And a.��������id = [1] "
    strSQL = "Select Distinct ����id, ����, ��������id, �������� From (" & strSQL & ")"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "", lngReportID)
    Do While rsTemp.EOF = False
        If i <= 4 Then
            If rsTemp!����ID = lngReportID Then
                GetRelationList = GetRelationList & vbCrLf & String(4, " ") & Chr(97 + i) & ") " & rsTemp!�������� & "��������"
            Else
                GetRelationList = GetRelationList & vbCrLf & String(4, " ") & Chr(97 + i) & ") " & rsTemp!���� & "��������"
            End If
        Else
            GetRelationList = GetRelationList & "..."
            Exit Do
        End If
        
        i = i + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

