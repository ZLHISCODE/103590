VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManagePrice 
   AutoRedraw      =   -1  'True
   Caption         =   "���ﻮ�۹���"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManagePrice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7410
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1695
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4170
      Width           =   45
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   15
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9675
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4140
      Width           =   9675
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Price"
               Description     =   "����"
               Object.ToolTipText     =   "���뻮�۴���"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Del"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰѡ��Ļ��۵�"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5844
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManagePrice.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11986
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
            AutoSize        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   1665
      Left            =   7470
      TabIndex        =   2
      Top             =   4185
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":115E
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   4185
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   2937
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":1478
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3405
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6006
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManagePrice.frx":1792
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   60
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":1AAC
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":1CC6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":1EE0
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":20FA
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2314
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2A8E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2CA8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":2EC2
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":30DC
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":32F6
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3510
            Key             =   "Modi"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":372A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3944
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3B5E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3D78
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":3F92
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":470C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4926
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4B40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4D5A
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":4F74
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManagePrice.frx":518E
            Key             =   "Modi"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Price 
         Caption         =   "���ﻮ��(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Price_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Modi 
         Caption         =   "�޸ĵ���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "����ʱ��(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "ɾ������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "��ӡ����֪ͨ��(&P)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵ���(&V)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmManagePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mblnNOMoved  As Boolean 'ɸѡ���Ϊ�շѵ���ʱ��ǰ�����Ƿ��ں󱸱���.
Private mrsDetail As ADODB.Recordset
Private mrsMoney As ADODB.Recordset

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    Operator As String
    PatientID As Long '����� ���ھ�ȷ����
    PatientName As String '���� ����ģ������
    ChargeKind As String
    DeptID As Long
    str�շ���� As String
    int�����־ As Integer  '1-����;2-סԺ;3-�����סԺ 126174
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mbln�շ� As Boolean

'��Ϣ��ض������
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEdit_Adjust_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Ե�����", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 2
    frmCharge.mstrInNO = strNo
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Modi_Click()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����޸ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrModiNO = ""
    
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mstrInNO = strNo
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK And gstrModiNO <> "" Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,�޸ĺ�ĵ��ݺ�Ϊ:[" & gstrModiNO & "],Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_Print_Click()
    Dim strNo As String
    
    If mbln�շ� Then Exit Sub   '�չ����˾�û�л��۵���
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo <> "" Then
        If MsgBox("ȷʵҪ��ӡ��ǰ���ݵĻ���֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & strNo, 2)
        End If
    Else
        MsgBox "��ǰû�е��ݿ��Դ�ӡ��", vbInformation, gstrSysName
    End If
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim blnPre As Boolean, intFrom As Integer
    
    blnPre = gblnҩ����λ
    intFrom = gint������Դ
        
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 1
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
    
    '������ҩƷ��λ����,����ˢ��
    If gblnҩ����λ <> blnPre Or gint������Դ <> intFrom Then
        If SQLCondition.Default Then SQLCondition.int�����־ = gint������Դ
        ShowBills mstrFilter
    End If
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo <> "" Then
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                    "NO=" & .TextMatrix(.Row, GetColNum("���ݺ�")), _
                    "������=" & .TextMatrix(.Row, GetColNum("ҽ��")))
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmPriceFilter.mstrPrivs = mstrPrivs
    '������Դ
    If gint������Դ = 1 Then
        frmPriceFilter.opt����(0).Value = True
    ElseIf gint������Դ = 2 Then
        frmPriceFilter.opt����(1).Value = True
    End If
    
    frmPriceFilter.Show 1, Me
    If gblnOK Then
        
        With frmPriceFilter
            mbln�շ� = .chk�շ�.Value = 1
            mstrFilter = .mstrFilter
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.DeptID = 0
            If .cbo����.ListIndex <> -1 Then
                SQLCondition.DeptID = .cbo����.ItemData(.cbo����.ListIndex)
            End If
            SQLCondition.PatientID = .mlngPrePatient
            SQLCondition.PatientName = UCase(.txt����.Text)
            SQLCondition.Operator = zlStr.NeedName(.cbo����Ա.Text)
            SQLCondition.ChargeKind = zlStr.NeedName(.cbo�ѱ�.Text)
            SQLCondition.str�շ���� = "," & .mstr�շ���� & ","
            SQLCondition.int�����־ = IIf(.opt����(0).Value, 0, IIf(.opt����(1).Value, 1, 2)) + 1
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim strNo As String
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If mshList.Row = 0 Or strNo = "" Then Exit Sub
    stbThis.Panels(2) = "�� " & mrsList.RecordCount & " �ŵ���"
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(strNo)
    Call ShowMoney(strNo)
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Public Function CheckBillDel(ByVal strPrivs As String, strNo As String, int��¼���� As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��黮�۵��Ƿ�����ɾ��
    '���:
    '����:
    '����:����ɾ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-13 12:35:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnҩƷ As Boolean
    
     If InStr(1, mstrPrivs, ";ҩƷ����ɾ��;") > 0 And InStr(1, mstrPrivs, ";���ƻ���ɾ��;") > 0 Then
        CheckBillDel = True: Exit Function
     End If
     '45774
     If InStr(1, mstrPrivs, ";ҩƷ����ɾ��;") = 0 And InStr(1, mstrPrivs, ";���ƻ���ɾ��;") = 0 Then
        MsgBox "�㲻����ɾ�����۵���Ȩ��,�������Ա��ϵ!", vbInformation, gstrSysName
        Exit Function
      End If
     blnҩƷ = InStr(1, mstrPrivs, ";ҩƷ����ɾ��;") > 0
     
    On Error GoTo errH
    
    strSQL = "Select Nvl(Count(ID),0) as ��Ŀ" & _
        " From ������ü�¼ " & _
        " Where NO=[1] And ��¼����=[2] And ��¼״̬ =0    " & _
        IIf(blnҩƷ = False, "  And �շ���� IN ('5','6','7')", "And Not �շ����  in ('5','6','7')")
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNo, int��¼����)
    
    If Val(Nvl(rsTmp!��Ŀ)) = 0 Then
        CheckBillDel = True
        Exit Function
    End If
    If blnҩƷ Then
        MsgBox "ע��:" & vbCrLf & "    ���۵��а�����������Ŀ,�㲻�߱�ɾ��������ĿȨ��,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    Else
        MsgBox "ע��:" & vbCrLf & "    ���۵��а�����ҩƷ��Ŀ,�㲻�߱�ɾ��ҩƷ��ĿȨ��,����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub mnuEdit_Del_Click()
    Dim strNo As String, strSQL As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ���ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ɾ��Ȩ��
    If Not BillOperCheck(3, mshList.TextMatrix(mshList.Row, GetColNum("������")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))), "ɾ��", , , 1) Then Exit Sub
    
    If HaveExecute(1, strNo, 1) Then
        MsgBox "�õ����а�����ִ�е�����,������ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckBillDel(mstrPrivs, strNo, 1) = False Then Exit Sub
    
    If MsgBox("ȷʵҪ������""" & strNo & """ɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & strNo & "')"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Price_Click()
    On Error Resume Next
    Err.Clear
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 0
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, strDate As Date
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    strDate = mshList.TextMatrix(mshList.Row, GetColNum(IIf(mbln�շ�, "�շ�ʱ��", "����ʱ��")))
    '��ʾ��������
    frmCharge.mlngModul = mlngModul
    frmCharge.mstrPrivs = mstrPrivs
    frmCharge.mbytInFun = 1
    frmCharge.mbytInState = 1
    frmCharge.mstrInNO = strNo
    frmCharge.mstrTime = strDate
    Set frmCharge.mobjMsgModule = mobjMsgModule
    frmCharge.Show 1, Me
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshMoney_GotFocus()
    Call SetActiveList(mshMoney)
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        picVsc.Top = picVsc.Top + Y
        picVsc.Height = picVsc.Height - Y
        mshMoney.Top = mshMoney.Top + Y
        mshMoney.Height = mshMoney.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDetail.Width + X < 1000 Or mshMoney.Width - X < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        mshDetail.Width = mshDetail.Width + X
        mshMoney.Left = mshMoney.Left + X
        mshMoney.Width = mshMoney.Width - X
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Price"
            mnuEdit_Price_Click
        Case "Modi"
            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '��ͷ
    If Not mbln�շ� Then
        objOut.Title.Text = "���ﻮ�۵����嵥"
    Else
        objOut.Title.Text = "�����շѵ����嵥"
    End If
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmPriceFilter
        objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " �� " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Adjust.Enabled = blnUsed And Not mbln�շ�
    mnuEdit_Modi.Enabled = blnUsed And Not mbln�շ�
    tbr.Buttons("Modi").Enabled = blnUsed And Not mbln�շ�
    
    mnuEdit_Del.Enabled = blnUsed And Not mbln�շ�
    mnuEdit_Print.Enabled = blnUsed And Not mbln�շ�
    mnuEdit_View.Enabled = blnUsed And Not mbln�շ�
    tbr.Buttons("Del").Enabled = blnUsed And Not mbln�շ�
    tbr.Buttons("View").Enabled = blnUsed And Not mbln�շ�
    
    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strSQL As String
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
        
    If gintDelPrice > 0 And InStr(mstrPrivs, "ɾ��") > 0 Then
        If MsgBox("ϵͳ׼��������ۺ󳬹� " & gintDelPrice & " ��δ�շ�δ��ҩ�Ļ��۵�,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call zlCommFun.ShowFlash("����������۵�,���Ժ� ...", Me)
            DoEvents
            
            strSQL = "zl_���ﻮ�ۼ�¼_Clear(" & gintDelPrice & ")"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
            
            Call zlCommFun.StopFlash
        End If
    End If
    Call RestoreWinState(Me, App.ProductName)
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    'Ȩ������
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    If InStr(mstrPrivs, ";����;") = 0 Then
        mnuEdit_Price.Visible = False
        mnuEdit_Print.Visible = False
        mnuEdit_Price_.Visible = False
        tbr.Buttons("Price").Visible = False
    End If
    
    If InStr(mstrPrivs, ";�޸�;") = 0 Then
        mnuEdit_Modi.Visible = False
        tbr.Buttons("Modi").Visible = False
    End If
    If InStr(mstrPrivs, ";����;") = 0 Then
        mnuEdit_Adjust.Visible = False
    End If
    If InStr(mstrPrivs, ";�޸�;") = 0 And InStr(mstrPrivs, ";����;") = 0 Then
        mnuEdit_Adjust_.Visible = False
    End If
    
    If InStr(mstrPrivs, ";ɾ��;") = 0 Then
        mnuEdit_Del.Visible = False
        mnuEdit_Del_.Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Del_").Visible = False
    End If
    
    'ȱʡ��������(������)
    mbln�շ� = False
    mstrFilter = " And �Ǽ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And ������||''=[5]"
    frmPriceFilter.mblnDateMoved = False
    SQLCondition.Default = True
    SQLCondition.int�����־ = gint������Դ
    
    Call SetHeader
    Call SetDetail
    Call SetMoney
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
    
    '��ʼ����Ϣ�������ģ��
    Call zlMsgModuleInit
        
    Exit Sub
errH:
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Call zlCommFun.ShowFlash("����������۵�,���Ժ� ...", Me)
        DoEvents
        Resume
    End If
    Call SaveErrLog
    Unload Me
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim sngVsc As Single, sngHsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    sngHsc = mshMoney.Width / (mshMoney.Width + mshDetail.Width)
    
    If mblnMax Then
        sngVsc = 0.3: sngHsc = 0.2
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    mshList.Left = Me.ScaleLeft
    mshList.Top = Me.ScaleTop + cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = 0
    picHsc.Width = mshList.Width
    
    mshDetail.Left = 0
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - picHsc.Height - mshList.Height
    mshDetail.Width = (Me.ScaleWidth - picVsc.Width) * (1 - sngHsc)
    
    picVsc.Top = mshDetail.Top
    picVsc.Left = mshDetail.Left + mshDetail.Width
    picVsc.Height = mshDetail.Height
    
    mshMoney.Top = mshDetail.Top
    mshMoney.Left = picVsc.Left + picVsc.Width
    mshMoney.Height = mshDetail.Height
    mshMoney.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mstrFilter = ""
    Unload frmPriceFilter
    Unload frmPriceGo
    
    Call SaveWinState(Me, App.ProductName)
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    '��ж��Ϣ����
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    frmPriceGo.mstrPrivs = mstrPrivs
    frmPriceGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmPriceGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmPriceGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
            End If
            If .cbo����Ա.ListIndex > 0 Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("������")) = zlStr.NeedName(.cbo����Ա.Text)
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
                        
            Call mshList_EnterCell
            mlngGo = i + 1
            
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("���ݺ�")) = "" Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If Not mbln�շ� Then
        strHead = "���ݺ�,1,850|��������,1,850|ҽ��,1,800|����,1,800|�Ա�,1,500|����,1,500|Ӧ�ս��,7,850|ʵ�ս��,7,850|������,1,800|����ʱ��,1,1850"
    Else
        strHead = "���ݺ�,1,850|��������,1,850|ҽ��,1,800|����,1,800|�Ա�,1,500|����,1,500|Ӧ�ս��,7,850|ʵ�ս��,7,850|������,1,800|�շ�ʱ��,1,1850|�շ���,1,800"
    End If
    
    With mshList
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        i = GetColNum("ҽ��")
        If InStr(mstrPrivs, "ҽ����ѯ") = 0 Then
            mshList.ColWidth(i) = 0
        ElseIf mshList.ColWidth(i) = 0 Then
            mshList.ColWidth(i) = 800
        End If
        
        .RowHeight(0) = 320
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .COLS - 1
        Call mshList_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, strSQL As String, strTable As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        
        If frmPriceFilter.mblnDateMoved Then
            '���ַ�ʽ�������ֿ�д��Ч����һ����,���õ������������
            strTable = zlGetFullFieldsTable("������ü�¼", 2, "", True, "")
        Else
            strTable = "������ü�¼"
        End If
        strIF = " Where ��¼����=1 And " & IIf(mbln�շ�, " ��¼״̬ IN(1,3)", "��¼״̬=0") & _
                " And ������ is Not NULL And ����Ա���� IS " & IIf(mbln�շ�, "NOT", "") & " NULL " & strIF
         
        Select Case SQLCondition.int�����־
        Case 1 '����
            strIF = strIF & " And  �����־ in (1,4)"
        Case 2 'סԺ
            strIF = strIF & " And  �����־ =2"
        Case Else   '����
        End Select
         
        strSQL = "Select * From " & strTable & " A " & strIF
        
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False Then
            If gblnUserIsClinic Then '113577�����ƿ�������
                strSQL = strSQL & " And Not Exists(" & _
                        "Select 1 From (Select NO From " & strTable & " C " & strIF & _
                        " And Not Exists(Select 1 From ������Ա D Where C.��������ID+0=D.����ID And D.��ԱID=[9]) Group by NO) E" & _
                        " Where A.NO=E.NO)"
            Else '����ִ�п���
                strSQL = strSQL & " And Not Exists(" & _
                        "Select 1 From (Select NO From " & strTable & " C " & strIF & _
                        " And Not Exists(Select 1 From ������Ա D Where C.ִ�в���ID+0=D.����ID And D.��ԱID=[9]) Group by NO) E" & _
                        " Where A.NO=E.NO)"
            End If
        End If
        
        strSQL = _
            "Select A.NO as ���ݺ�,B.���� as ��������,A.������ as ҽ��,Ltrim(A.����) as ����,A.�Ա�,A.����," & _
            " To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��," & _
            " To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��,A.������," & _
            " To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as " & IIf(mbln�շ�, "�շ�ʱ��,A.����Ա���� as �շ���", "����ʱ��") & _
            " From (" & strSQL & ") A,���ű� B" & _
            " Where A.��������ID = B.ID" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
            " Group by A.NO,B.����,A.������,A.����,A.�Ա�,A.����,A.������," & IIf(mbln�շ�, "A.����Ա����,", "") & "A.�Ǽ�ʱ��" & _
            " Order by " & IIf(mbln�շ�, "�շ�ʱ��", "����ʱ��") & " Desc,���ݺ� Desc"
        
        With SQLCondition
            If .Default Then .Operator = UserInfo.����
            Set mrsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .DateB, .DateE, .NOB, .NOE, .Operator, .PatientName, .ChargeKind, .DeptID, _
                              UserInfo.ID, .str�շ����, .PatientID)
        End With
    End If
    
    mshList.Clear:   mshList.Rows = 2
    mshDetail.Clear: mshDetail.Rows = 2
    mshMoney.Clear:  mshMoney.Rows = 2
    
    If Not mbln�շ� Then
        mshList.ForeColor = vbBlack:    mshDetail.ForeColor = vbBlack:   mshMoney.ForeColor = vbBlack
    Else
        mshList.ForeColor = &H808080:   mshDetail.ForeColor = &H808080:  mshMoney.ForeColor = &H808080
    End If
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        stbThis.Panels(2) = "�� " & mrsList.RecordCount & " �ŵ���"
        Call SetMenu(True)
    End If
    Call SetHeader
    Call SetDetail
    Call SetMoney
    
    mnuEdit_Del.Enabled = Not mrsList.EOF And Not mbln�շ�
    tbr.Buttons("Del").Enabled = Not mrsList.EOF And Not mbln�շ�
    mnuEdit_Modi.Enabled = Not mrsList.EOF And Not mbln�շ�
    tbr.Buttons("Modi").Enabled = Not mrsList.EOF And Not mbln�շ�
    mnuEdit_Adjust.Enabled = Not mrsList.EOF And Not mbln�շ�
    
    If Not blnSort Then Call zlCommFun.StopFlash
    
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetail(Optional strNo As String, Optional blnSort As Boolean)
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        '���֮ǰ��ɸѡѡ��,�����嵥����������շѼ�¼,��Ҫ���
        If frmPriceFilter.mblnDateMoved Then
            mblnNOMoved = zlDatabase.NOMoved("������ü�¼", strNo, , "1")
        Else
            mblnNOMoved = False
        End If
        strSQL = _
        " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
                IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
        "       To_Char(Avg(Nvl(A.����,1)*A.����)" & _
                IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
        "       A.�ѱ�,To_Char(Sum(A.��׼����)" & _
                IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
        "       To_Char(Sum(A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
        "       To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
        "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����" & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
        "       And A.��¼����=1 and A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.����,", "") & " B.���,A.���㵥λ,A.�ѱ�," & _
        "       D.����,Nvl(A.��������,B.��������),X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
        " Order by Nvl(A.�۸񸸺�,A.���)"
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    mshDetail.Clear
    mshDetail.Rows = 2

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    Call SetDetail
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "���,1,750|����,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1000|��λ,4,500|����,7,850|�ѱ�,1,750|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,1,850|����,1,850"
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        'Call mshDetail_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowMoney(Optional strNo As String, Optional blnSort As Boolean)
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        strSQL = _
            "Select " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����") & " as ��Ŀ," & _
            " To_Char(Sum(A.ʵ�ս��),'9999999" & gstrDec & "') as ��� " & _
            " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,������Ŀ B " & _
            " Where A.������ĿID=B.ID AND A.��¼����=1" & _
            " And A.��¼״̬ IN(0,1,3) And A.NO=[1]" & _
            " Group by " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����")
        Set mrsMoney = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    mshMoney.Clear
    mshMoney.Rows = 2
    
    If Not mrsMoney.EOF Then Set mshMoney.DataSource = mrsMoney
    Call SetMoney
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMoney()
    Dim strHead As String
    Dim i As Long
    
    strHead = "��Ŀ,1,850|���,7,850"
    With mshMoney
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshMoney, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        'Call mshMoney_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, True)
    End If
End Sub

Private Sub mshMoney_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshMoney.MouseRow = 0 Then
        mshMoney.MousePointer = 99
    Else
        mshMoney.MousePointer = 0
    End If
End Sub

Private Sub mshMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshMoney.MouseCol
    
    If Button = 1 And mshMoney.MousePointer = 99 Then
        If mshMoney.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshMoney.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsMoney Is Nothing Then Exit Sub
        
        Set mshMoney.DataSource = Nothing

        mrsMoney.Sort = mshMoney.TextMatrix(0, lngCol) & IIf(mshMoney.ColData(lngCol) = 0, "", " DESC")
        mshMoney.ColData(lngCol) = (mshMoney.ColData(lngCol) + 1) Mod 2
        
        Call ShowMoney(, True)
    End If
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &H8000000D
        mshDetail.BackColorSel = &H8000000C
        mshMoney.BackColorSel = &H8000000C
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000D
        mshMoney.BackColorSel = &H8000000C
    ElseIf obj Is mshMoney Then
        mshList.BackColorSel = &H8000000C
        mshDetail.BackColorSel = &H8000000C
        mshMoney.BackColorSel = &H8000000D
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '���:lngModule -ģ���
    '     strPivs-Ȩ�޴�
    '����:objMsgModule-������Ϣ����
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModuleInit = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModuleUnload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ж��Ϣģ��
    '���:objMsgModule-��Ϣ����
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModuleUnload = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


