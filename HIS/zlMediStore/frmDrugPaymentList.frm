VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDrugPaymentList 
   Caption         =   "ҩƷ��������"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   Icon            =   "frmDrugPaymentList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   2760
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3690
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbrTool 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11775
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Payment"
                     Text            =   "���"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Imprest"
                     Text            =   "Ԥ���"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Strike"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugPaymentList.frx":014A
         Begin VB.Timer LimitTime 
            Enabled         =   0   'False
            Interval        =   8000
            Left            =   6660
            Top             =   180
         End
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4620
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPaymentList.frx":0464
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmDrugPaymentList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":215C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
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
            Picture         =   "frmDrugPaymentList.frx":237C
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":259C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":27BC
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":29D8
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":2BF8
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":2E18
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":3034
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":3250
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":346A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":35C4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPaymentList.frx":37E4
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshAddition 
      Height          =   2595
      Left            =   6810
      TabIndex        =   7
      Top             =   3000
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorFixed  =   -2147483644
      GridColorFixed  =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgVsc_s 
      Height          =   2325
      Left            =   6480
      MousePointer    =   9  'Size W E
      Top             =   3210
      Width           =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Begin VB.Menu mnuEditAddPayment 
            Caption         =   "���(&P)"
         End
         Begin VB.Menu mnuEditAddImprest 
            Caption         =   "Ԥ���(&I)"
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&C)"
      End
      Begin VB.Menu mnuEditStrike 
         Caption         =   "����(&K)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
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
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmDrugPaymentList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mlastRow As Long                '�ϴε������
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mintPreDetailCol As Integer     'ǰһ�ε������������
Private mintDetailsort As Integer       'ǰһ�ε����������
Private mblnStartup As Boolean


Private Sub GetList(ByVal StrFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    
    On Error GoTo errHandle
    mlastRow = 0
    Call zlcommfun.ShowFlash("��������ҩƷ�����¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    mshList.Redraw = False
    
    gstrSQL = "SELECT a.no,b.id, b.����,nvl(Ԥ����,0) as Ԥ���� , ltrim(to_char(SUM (a.���),'9999999999999990.00')) AS ���, a.������ AS ������," _
            & "TO_CHAR (min(a.��������), 'yyyy-MM-dd HH24:MI:SS') AS ��������, a.�����," _
            & "TO_CHAR (min(a.�������), 'yyyy-MM-dd HH24:MI:SS') AS �������, a.��¼״̬, a.ժҪ " _
            & "FROM ҩƷ�����¼ a, ҩƷ��Ӧ�� b " _
            & "Where a.��λid = b.id " _
            & StrFind _
            & " GROUP BY a.no,b.id,b.����,nvl(Ԥ����,0),a.������,a.�����,a.��¼״̬, a.ժҪ " _
            & " ORDER BY a.no desc "
    
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsList = zldatabase.OpenSQLRecord(gstrSQL, "GetList")
    Call SQLTest
   
    Set mshList.Recordset = rsList
    With mshList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            
            .TopRow = 1
            .rows = .rows - 99
            
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    SetListColWidth
    
    mshlist_EnterCell    '�г�������
    
    SetStrikeColor
    mshList.Redraw = True
    Call zlcommfun.StopFlash
    Screen.MousePointer = vbDefault
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim IntCol As Integer
    
    With mshList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = .TextMatrix(intRow, .Cols - 2)
            If intStatus = 3 Then
                .Row = intRow
                For IntCol = 0 To .Cols - 1
                    .Col = IntCol
                    .CellForeColor = &H80000001
                Next
            End If
            If intStatus = 2 Then
                .Row = intRow
                For IntCol = 0 To .Cols - 1
                    .Col = IntCol
                    .CellForeColor = &HFF
                Next
            End If
        Next
    End With
                
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim IntCol As Integer
    
    With mshList
        .ColAlignment(4) = flexAlignRightCenter
        
        If mblnStartup = False Then
            For IntCol = 1 To .Cols - 1
                .ColWidth(IntCol) = 1500
            Next
        End If
        .ColWidth(1) = 0
        .ColWidth(1) = 0
        .ColWidth(3) = 0
        .ColWidth(.Cols - 2) = 0
    End With
End Sub


Private Sub SetDetailColWidth()
    Dim IntCol As Integer
    
    With mshDetail
        .ColAlignment(1) = flexAlignLeftCenter      '��ⵥ��
        .ColAlignment(2) = flexAlignLeftCenter      '��Ʊ��
        .ColAlignment(3) = flexAlignRightCenter     '��Ʊ���
        .ColAlignment(8) = flexAlignCenterCenter    '��λ
        .ColAlignment(9) = flexAlignRightCenter     '����
        .ColAlignment(10) = flexAlignRightCenter    '�ɹ���
        .ColAlignment(11) = flexAlignRightCenter    '�ɹ����
           
        If mblnStartup = False Then
            For IntCol = 0 To .Cols - 1
                .ColWidth(IntCol) = 1000
            Next
            .ColWidth(4) = 2000
        End If
            
    End With
End Sub


'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    '�⹺�������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա�����

    If InStr(1, gstrprivs, "��������") = 0 Then
         mnuFileParameter.Visible = False
         mnuFileLine3.Visible = False                '��Ӧ�ķָ���
    End If
     
    If InStr(1, gstrprivs, "�Ǽ�") = 0 Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If InStr(1, gstrprivs, "�޸�") = 0 Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If InStr(1, gstrprivs, "ɾ��") = 0 Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If InStr(1, gstrprivs, "���") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If InStr(1, gstrprivs, "����") = 0 Then
        mnuEditStrike.Visible = False
        tlbTool.Buttons("Strike").Visible = False
        
        If mnuEditVerify.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    If InStr(1, gstrprivs, "���ݴ�ӡ") = 0 Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '�ָ�����
    Dim strStart As String
    Dim strEnd As String
    Dim StrFind As String
    
    mblnStartup = False
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    strStart = Format(DateAdd("m", -1, zldatabase.Currentdate), "yyyy-MM-dd")
    strEnd = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    StrFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between To_Date('" & strStart & " 00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & strEnd & " 23:59:59','YYYY-MM-DD HH24:MI:SS')"
    mstrFind = StrFind
    
    lblRange.Caption = "��ѯ��Χ:" & Format(DateAdd("m", -1, zldatabase.Currentdate), "yyyy��MM��dd��") & "��" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    
    GetList (mstrFind)  '�г�����ͷ
    mblnStartup = True
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If
    
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
        .Height = 720
    End With
    
    With picSeparate_s
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
        
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = imgVsc_s.Left   '- 10
    End With
    
    With mshAddition
        .Top = mshDetail.Top
        .Left = imgVsc_s.Left + imgVsc_s.Width  '+ 10
        .Width = mshList.Width - .Left
        .Height = mshDetail.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub imgVsc_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With imgVsc_s
            If .Left + x < 2000 Then Exit Sub
            If .Left + x > ScaleWidth - 2000 Then Exit Sub
            .Left = .Left + x
        End With
        
        Me.mshDetail.Width = Me.mshDetail.Width + x
        Me.mshAddition.Left = Me.mshAddition.Left + x
        Me.mshAddition.Width = Me.mshAddition.Width - x
    End If
End Sub


Private Sub mnuEditAddImprest_Click()
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    StrNo = ""
    frmDrugImprestCard.ShowCard Me, StrNo, 1, , BlnSuccess
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditAddPayment_Click()
    Dim BlnSuccess  As Boolean
    
    FrmDrugPaymentCard.ShowCard Me, "", 1, 1, BlnSuccess
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    With mshList
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 1 Then
            frmDrugImprestCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
        Else
            FrmDrugPaymentCard.ShowCard Me, StrNo, 3, .TextMatrix(.Row, .Cols - 2), BlnSuccess
        End If
        
    End With
    If BlnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With mshList
        
        If .TextMatrix(.Row, 3) = 1 Then
            strTitle = "ҩƷԤ���"
        Else
            strTitle = "ҩƷ���"
        End If
        On Error GoTo errHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & strBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_ҩƷ�������_delete('" & strBillNo & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
            Call SQLTest
            intRecord = intRecord - 1
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With mshDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                SetEnable
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            mshlist_EnterCell
        End If
    End With
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim StrNo As String
    With mshList
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 1 Then
            frmDrugImprestCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        Else
            FrmDrugPaymentCard.ShowCard Me, StrNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
    
End Sub

Private Sub mnuEditStrike_Click()
    '����
    With mshList
        If MsgBox("��ȷʵҪ�������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If StrikeSave = True Then
                mnuViewRefresh_Click
            End If
        End If
    End With
End Sub

Private Function StrikeSave() As Boolean
    
    StrikeSave = False
    With mshList
        On Error GoTo errHandle
        gstrSQL = "zl_ҩƷ�������_STRIKE('" & .TextMatrix(.Row, 0) & "'," _
            & .TextMatrix(.Row, 4) & "," & .TextMatrix(.Row, 1) & ",'" & UserInfo.�û����� & "')"
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        Call SQLTest
        
    End With
    StrikeSave = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

    'MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
End Function

Private Sub mnuEditModify_Click()
    '�޸�
    Dim StrNo As String
    Dim BlnSuccess As Boolean
    
    BlnSuccess = False
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        StrNo = .TextMatrix(.Row, 0)
        If .TextMatrix(.Row, 3) = 0 Then
            FrmDrugPaymentCard.ShowCard Me, StrNo, 2, mshList.TextMatrix(mshList.Row, mshList.Cols - 2), BlnSuccess
        Else
            frmDrugImprestCard.ShowCard Me, StrNo, 2, mshList.TextMatrix(mshList.Row, .Cols - 2), BlnSuccess
        End If
        If BlnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If IIf(IsNull(.TextMatrix(.Row, 3)), 0, .TextMatrix(.Row, 3)) = 1 Then
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_2", "zl8_bill_1320_2"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), 1
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), 1
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With mshList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If IIf(IsNull(.TextMatrix(.Row, 3)), 0, .TextMatrix(.Row, 3)) = 1 Then
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_2", "zl8_bill_1320_2"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), 2
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), 2
        End If
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    mshList.Redraw = False
    subPrint 3
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
    
End Sub

Private Sub mnuFileParameter_Click()
    '��������
    frm��������.���ò��� Me, Me.Caption
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    mshList.Redraw = False
    subPrint 2
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    mshList.Redraw = False
    subPrint 1
    mshList.Redraw = True
    mshList.Col = 0
    mshList.ColSel = mshList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    
    Dim strStart As Date
    Dim strEnd As Date
    Dim strVerifyStart As Date
    Dim strVerifyEnd As Date
    Dim StrFind As String
    
    
    StrFind = FrmDrugPaymentSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd)
    
    If StrFind <> "" Then
        mstrFind = StrFind
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub mshDetail_Click()
    With mshDetail
         If .Row < 1 Or .TextMatrix(.Row, 0) = "" Then Exit Sub
         If .MouseRow = 0 Then
            DetailSort          '������
            Exit Sub
         End If
    End With
End Sub

Private Sub mshList_Click()
    With mshList
         If .Row < 1 Then Exit Sub
         If .MouseRow = 0 Then
            ListSort
            Exit Sub
         End If
    End With
End Sub

Private Sub mshlist_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If mshList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub mshlist_EnterCell()
    Dim rsDetail As New Recordset
    Dim strUnitName As String                       '��λ����:�����ﵥλ��סԺ��λ��
    Dim str��װϵ�� As String
    Dim intUnit As Integer
    
    On Error GoTo errHandle
    If mlastRow = mshList.Row Then Exit Sub
    mlastRow = mshList.Row
    SetEnable
    
    intUnit = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Caption, "ҩƷ��λ", "0")
    If glngSys \ 100 = 8 Then
        strUnitName = Choose(intUnit + 1, "c.ҩ�ⵥλ", "c.�ۼ۵�λ")
        str��װϵ�� = Choose(intUnit + 1, "c.ҩ���װ", "1")
    Else
        strUnitName = Choose(intUnit + 1, "c.ҩ�ⵥλ", "c.���ﵥλ", "c.סԺ��λ", "c.�ۼ۵�λ")
        str��װϵ�� = Choose(intUnit + 1, "c.ҩ���װ", "c.�����װ", "c.סԺ��װ", "1")
    End If
    
    If mshList.Row >= 1 And LTrim(mshList.TextMatrix(mshList.Row, 0)) <> "" And IIf(mshList.TextMatrix(mshList.Row, 3) = "", 0, mshList.TextMatrix(mshList.Row, 3)) <> 1 And mshList.TextMatrix(mshList.Row, mshList.Cols - 2) <> "2" Then
        mshList.Col = 0
        mshList.ColSel = mshList.Cols - 1
        
        mshDetail.Redraw = False
        gstrSQL = "SELECT distinct b.������� AS �������, b.no, a.��Ʊ��, to_char(a.��Ʊ���,'99999999999999990.00') as ��Ʊ��� ," _
            & "'[' || c.���� || ']' || DECODE (d.����, NULL, e.ͨ������, d.����) AS ҩƷ��Ϣ," _
            & "c.���, b.����, b.����," & strUnitName & " AS ��λ, to_char(b.ʵ������/" & str��װϵ�� & ",'99999999999999990.00000')  AS ����, to_char(b.�ɱ���*" & str��װϵ�� & ",'99999999999999990.00000') as �ɹ��� ,to_char(b.�ɱ����,'99999999999999990.00') as �ɹ���� " _
            & " FROM ҩƷӦ����¼ a, " _
               & "(SELECT * From ҩƷ�շ���¼ WHERE ���� = 1 and ��ҩ��λid=" & mshList.TextMatrix(mshList.Row, 1) & " ) b," _
                & "ҩƷĿ¼ c," _
                & "ҩƷ���� d," _
                & "ҩƷ��Ϣ e " _
           & " Where a.�շ�id = b.id " _
           & "   AND b.ҩƷid = c.ҩƷid " _
           & "   AND c.ҩƷid = d.ҩƷid (+) " _
           & "   AND c.ҩ��id = e.ҩ��id " _
           & "   AND a.������� = (SELECT DISTINCT ������� FROM ҩƷ�����¼ WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  ������� is not null and ��¼״̬<>2 ) " _
          & "  ORDER BY a.��Ʊ��, b.no "
        
        With rsDetail
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
            Call SQLTest
            
            Set mshDetail.Recordset = rsDetail
            .Close
        End With
        With mshDetail
            If .rows = 1 Then
                .rows = .rows + 100
                .Row = 1
                .Redraw = True
                
                .TopRow = 1
                .rows = .rows - 99
            End If
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        
        mshDetail.Redraw = True
    ElseIf LTrim(mshList.TextMatrix(mshList.Row, 0)) = "" Or mshList.TextMatrix(mshList.Row, 3) = "1" Or mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "2" Then
        With mshDetail
            .Cols = 12
            .rows = 2
            .Clear
    
            .TextMatrix(0, 0) = "�������"
            .TextMatrix(0, 1) = "��ⵥ��"
            .TextMatrix(0, 2) = "��Ʊ��"
            .TextMatrix(0, 3) = "��Ʊ���"
            .TextMatrix(0, 4) = "ҩƷ��Ϣ"
            .TextMatrix(0, 5) = "���"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "����"
            .TextMatrix(0, 8) = "��λ"
            .TextMatrix(0, 9) = "����"
            .TextMatrix(0, 10) = "�ɹ���"
            .TextMatrix(0, 11) = "�ɹ����"
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
        End With
    End If
    SetDetailColWidth
    
    If mshList.TextMatrix(mshList.Row, 3) = "1" Or mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "2" Then
        gstrSQL = "SELECT DECODE (Ԥ����, 1, '��', 0, '��') AS Ԥ����," _
            & " TO_CHAR (���, '99999999999999990.00') AS ���, ���㷽ʽ, �������, " _
            & " DECODE (Ԥ����, 1, no, '') AS ���Ԥ����� " _
            & " From ҩƷ�����¼ " _
            & "WHERE no = '" & mshList.TextMatrix(mshList.Row, 0) _
            & "'  and ��¼״̬=" & mshList.TextMatrix(mshList.Row, mshList.Cols - 2) _
            & " ORDER BY Ԥ����,��� "
    Else
        gstrSQL = "SELECT DECODE (Ԥ����, 1, '��', 0, '��') AS Ԥ����," _
            & " TO_CHAR (���, '99999999999999990.00') AS ���, ���㷽ʽ, �������, " _
            & " DECODE (Ԥ����, 1, no, '') AS ���Ԥ����� " _
            & " From ҩƷ�����¼ " _
            & "WHERE ������� = (SELECT DISTINCT ������� FROM ҩƷ�����¼ WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  ������� is not null and ��¼״̬<>2 ) " _
            & "  and ��¼״̬=" & IIf(mshList.TextMatrix(mshList.Row, mshList.Cols - 2) = "", 1, mshList.TextMatrix(mshList.Row, mshList.Cols - 2)) _
            & "  and Ԥ����=0 "
        gstrSQL = gstrSQL & _
             " union " _
             & " SELECT DECODE (Ԥ����, 1, '��', 0, '��') AS Ԥ����," _
            & " TO_CHAR (���, '99999999999999990.00') AS ���, ���㷽ʽ, �������, " _
            & " DECODE (Ԥ����, 1, no, '') AS ���Ԥ����� " _
            & " From ҩƷ�����¼ " _
            & "WHERE ������� = (SELECT DISTINCT ������� FROM ҩƷ�����¼ WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  ������� is not null and ��¼״̬<>2 ) " _
            & "  and Ԥ����=1 "
        gstrSQL = gstrSQL & _
            " union " _
            & " SELECT DECODE (Ԥ����, 1, '��', 0, '��') AS Ԥ����," _
            & " TO_CHAR (���, '99999999999999990.00') AS ���, ���㷽ʽ, �������, " _
            & " DECODE (Ԥ����, 1, no, '') AS ���Ԥ����� " _
            & " From ҩƷ�����¼ " _
            & "WHERE ������� = (SELECT DISTINCT ������� FROM ҩƷ�����¼ WHERE no='" & mshList.TextMatrix(mshList.Row, 0) & "' and  ������� is not null and ��¼״̬<>2 ) " _
            & "  and (��¼״̬=2) " _
            & "  and nvl(Ԥ����,0)=0 "
            
        gstrSQL = " select * from (" & gstrSQL & ") order by Ԥ���� "
            
        '& "  and (��¼״̬=1 or ��¼״̬=3) "
            
    End If
    
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsDetail = zldatabase.OpenSQLRecord(gstrSQL, "cmd����_Click")
    Call SQLTest
    
    Set mshAddition.Recordset = rsDetail
    rsDetail.Close
    With mshAddition
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = True
            
            .TopRow = 1
            .rows = .rows - 99
        End If
        
        If mblnStartup = False Then
            .ColWidth(0) = 800
            .ColWidth(1) = 1000
            .ColWidth(2) = 800
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000
        End If
        
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        .ColAlignment(1) = flexAlignRightCenter
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub mshlist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With mshList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Height = picSeparate_s.Top - .Top
    End With
    
    With mshDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        mshAddition.Top = .Top
        mshAddition.Height = .Height
    End With
    
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddPayment_Click
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Strike"
            mnuEditStrike_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        
    End Select
    
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    With mshList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
        
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditStrike.Visible = True Then
                mnuEditStrike.Enabled = False
                tlbTool.Buttons("Strike").Enabled = False
            End If
             
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
         Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .Cols - 4) = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = True
                    tlbTool.Buttons("Strike").Enabled = True
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
            Else   '2,3 ������
                If .TextMatrix(.Row, .Cols - 2) = 3 Then
                    .ToolTipText = "�������ݵ�ԭ����"
                ElseIf .TextMatrix(.Row, .Cols - 2) = 2 Then
                    .ToolTipText = "��������"
                End If
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditStrike.Visible = True Then
                    mnuEditStrike.Enabled = False
                    tlbTool.Buttons("Strike").Enabled = False
                End If
                 
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    
        
    objPrint.Title.Text = "ҩƷ�������"
    Set objPrint.Body = mshList
    
        
    'objRow.Add "��ӡ��:" & UserInfo.�û�����
    'objRow.Add "��ӡ����:" & Format(ZlDatabase.Currentdate, "yyyy-MM-dd")
    'objPrint.UnderAppRows.Add objRow
    
    objRow.Add "��ӡ�ˣ�" & UserInfo.�û�����
    objRow.Add "��ӡʱ�䣺" & Format(zldatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
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

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
            mnuEditAddPayment_Click
        Case "Imprest"
            mnuEditAddImprest_Click
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'�Ե���ͷ������
Private Sub ListSort()
    Dim IntCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshList
        If .rows > 1 Then
            .Redraw = False
            IntCol = .MouseCol
            .Col = IntCol
            .ColSel = IntCol
            intTemp = .TextMatrix(.Row, 0)
                    
            If IntCol = 4 Then
                If IntCol = mintPreCol And mintsort = flexSortNumericDescending Then
                    .Sort = flexSortNumericAscending
                    mintsort = flexSortNumericAscending
                Else
                    .Sort = flexSortNumericDescending
                    mintsort = flexSortNumericDescending
                End If
            Else
                If IntCol = mintPreCol And mintsort = flexSortStringNoCaseDescending Then
                    .Sort = flexSortStringNoCaseAscending
                    mintsort = flexSortStringNoCaseAscending
                Else
                    .Sort = flexSortStringNoCaseDescending
                    mintsort = flexSortStringNoCaseDescending
                End If
            End If
                
            mintPreCol = IntCol
            .Row = FindRow(mshList, intTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'�Ե���ͷ������
Private Sub DetailSort()
    Dim IntCol As Integer
    Dim intRow As Integer
    Dim intTemp As String
    
    With mshDetail
        If .rows > 1 Then
            .Redraw = False
            IntCol = .MouseCol
            .Col = IntCol
            .ColSel = IntCol
            intTemp = .TextMatrix(.Row, 1)
            
            Select Case IntCol
                Case 3, 9, 10, 11
                    If IntCol = mintPreDetailCol And mintDetailsort = flexSortNumericDescending Then
                       .Sort = flexSortNumericAscending
                       mintDetailsort = flexSortNumericAscending
                    Else
                       .Sort = flexSortNumericDescending
                       mintDetailsort = flexSortNumericDescending
                    End If
                    
                Case Else
                    If IntCol = mintPreDetailCol And mintDetailsort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       mintDetailsort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       mintDetailsort = flexSortStringNoCaseDescending
                    End If
            End Select
                
            mintPreDetailCol = IntCol
            .Row = FindRow(mshDetail, intTemp, 1)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

'Ѱ����ĳһ����ȵ���
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal IntCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, IntCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, IntCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

