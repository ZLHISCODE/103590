VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����ҵ������ 
   Caption         =   "����ҵ������"
   ClientHeight    =   5940
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "frm����ҵ������.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9345
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImgColor 
      Left            =   660
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":1F14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":2448
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":2662
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":2F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":3256
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":3470
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":368A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgBlack 
      Left            =   90
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":38A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":3ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":3CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":3FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":420C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":4AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":4E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":501A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ҵ������.frx":5234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1296
      BandCount       =   1
      _CBWidth        =   9345
      _CBHeight       =   735
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   675
      Width1          =   615
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   675
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   1191
         ButtonWidth     =   1455
         ButtonHeight    =   1191
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImgBlack"
         HotImageList    =   "ImgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "��ͥ����"
               Key             =   "Home"
               Object.ToolTipText     =   "�����ͥ����"
               Object.Tag             =   "��ͥ����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�涨��"
               Key             =   "Spec"
               Object.ToolTipText     =   "��������涨��"
               Object.Tag             =   "�涨��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "�����ؼ�"
               Key             =   "Especial"
               Object.ToolTipText     =   "���������ؼ�"
               Object.Tag             =   "�����ؼ�"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����תԺ"
               Key             =   "Switch"
               Object.ToolTipText     =   "����תԺ��ת�����"
               Object.Tag             =   "����תԺ"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh����嵥 
      Height          =   2475
      Left            =   150
      TabIndex        =   2
      Top             =   900
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   4366
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   7513801
      BackColorSel    =   14783374
      BackColorBkg    =   16777215
      GridColorFixed  =   0
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5580
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   635
      SimpleText      =   $"frm����ҵ������.frx":544E
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm����ҵ������.frx":5495
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11430
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
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFileSplitSet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSplitPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplitReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuRequest 
      Caption         =   "ҵ������(&R)"
      Begin VB.Menu mnuRequestHome 
         Caption         =   "��ͥ����(&H)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestSpec 
         Caption         =   "�������ⲡ(&S)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestEspecial 
         Caption         =   "�����ؼ�(&E)"
         Shortcut        =   {DEL}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRequestSwitch 
         Caption         =   "תԺ����(&W)"
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
         Begin VB.Menu mnuViewTool_1 
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_1 
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
Attribute VB_Name = "frm����ҵ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����
Private str��ʼ���� As String, str�������� As String, str��˱�־ As String
Private rsVerify As New ADODB.Recordset

Private Sub Form_Load()
    Call initGird
End Sub

Private Sub initGird()
    Dim intCol As Integer, intCols As Integer
    Dim arrColumn
    Const strColumns As String = "����|�Ա�|���֤��|��������|ת��ҽԺ|��������|��˱�־"
    
    arrColumn = Split(strColumns, "|")
    intCols = UBound(arrColumn)
    With msh����嵥
        .Rows = 2
        .Cols = intCols + 1
        
        For intCol = 0 To intCols
            .TextMatrix(0, intCol) = arrColumn(intCol)
            .ColAlignmentFixed(intCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Me.msh����嵥
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuRequestSwitch_Click()
    Dim blnReturn As Boolean
    blnReturn = frmתԺ����.ShowME(1, Me)
End Sub

Private Sub mnuViewFind_Click()
    Dim blnReturn As Boolean
    blnReturn = frm����תԺ_����.ShowME(str��ʼ����, str��������, str��˱�־)
    Call RefreshData
End Sub

Private Sub mnuViewRefresh_Click()
    Call RefreshData
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = mnuViewStatus.Checked Xor True
    stbThis.Visible = stbThis.Visible Xor True
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = mnuViewToolButton.Checked Xor True
    cbrThis.Visible = cbrThis.Visible Xor True
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim tbrbutton As Button
    mnuViewToolText.Checked = mnuViewToolText.Checked Xor True
    For Each tbrbutton In tbrThis.Buttons
        tbrbutton.Caption = IIf(mnuViewToolText.Checked, tbrbutton.Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
    '���ܣ�������б�
    '������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh����嵥.Row
    
    '��ͷ
    objOut.Title.Text = "תԺ�����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zldatabase.Currentdate, "yyyy��MM��DD��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = msh����嵥
    
    '���
    msh����嵥.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh����嵥.Redraw = True
    
    msh����嵥.Row = intRow
    msh����嵥.Col = 0: msh����嵥.ColSel = msh����嵥.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Print"
        Call mnuFilePrint_Click
    Case "Preview"
        Call mnuFilePreview_Click
    Case "Switch"
        Call mnuRequestSwitch_Click
    Case "Filter"
        Call mnuViewFind_Click
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Exit"
        Call mnuFileQuit_Click
    End Select
End Sub

Private Sub RefreshData()
    Dim intCol As Integer, intCols As Integer
    Dim strColumns As String, strData As String
    Dim strFields As String, strValues As String
    Dim arrColumn
    
    '��ȡ��������תԺ������
    If str��ʼ���� = "" Then Exit Sub
    If str��˱�־ = "" Then Exit Sub
    
    If Not ҽ����ʼ��_������ Then Exit Sub
    
    DoEvents
    Call zlcommfun.ShowFlash("���ڴ�������ȡתԺ������Ϣ,���Ժ� ...", Me)
    DoEvents
    
'    1   hospital_idҽ�ƻ�������   20  ��
'    2   from_date  ��ѯ��ʼ����       ��  ��ʽ��YYYY-MM-DD
'    3   to_date    ��ѯ��ֹ����       ��  ��ʽ��YYYY-MM-DD
'    4   audit_flag ��˱�־       4   ��  "0"��δ���    "1"�����ͨ��    "2"�����δͨ��    "all"��ȫ��
    Call DebugTool("׼�����û�ȡ������Ϣ����")
    gstrField_������ = "hospital_id||from_date||to_date||audit_flag"
    gstrValue_������ = gCominfo_������.ҽԺ���� & "||" & str��ʼ���� & "||" & str�������� & "||" & str��˱�־
    If Not ���ýӿ�_׼��_������(Function_������.תԺ����_��ѯ�����Ϣ) Then GoTo StopFlash
    If Not ���ýӿ�_д��ڲ���_������(1) Then GoTo StopFlash
    If Not ���ýӿ�_ִ��_������ Then GoTo StopFlash
    If Not ���ýӿ�_ָ����¼��_������("ToanotherHosInfo") Then GoTo StopFlash
    
    '��ʼ����¼��
    Call DebugTool("׼����ʼ���ڲ���¼��")
    strColumns = "name|sex|idcard|disease|to_hospital_name|input_date|audit_flag"
    arrColumn = Split(strColumns, "|")
    intCols = UBound(arrColumn)
    For intCol = 0 To intCols
        strFields = strFields & "|" & arrColumn(intCol) & "," & adLongVarChar & ",50"
    Next
    strFields = Mid(strFields, 2)
    Call Record_Init(rsVerify, strFields)
    
'    1   name           ����    10
'    2   sex            �Ա�    2
'    3   birthday       ��������        ��ʽ��YYYY-MM-DD
'    4   idcard         ���֤����  20
'    5   insr_code      ���պ�  30
'    6   corp_name      ��λ����    50
'    7   pers_name      ��Ա���    20
'    8   official_name  ����Ա����  20
'    9   indi_id        ���˱��    8
'    10  serial_apply   �������к�  12
'    11  busi_type      ҵ������    2   "16"�������ؼ�
'    12  apply_type     ��������    1   "0"����ͨ����    "1"��׷������
'    13  apply_content  ��������    1   "1"�������ؼ�
'    14  icd            ��������    20
'    15  disease        ��������    50
'    16  disease_deac   ����ժҪ    500
'    17  oper_desc      ��Ҫ���Ʒ���    500
'    18  doctor_name    ����ҽʦ    10
'    19  apply_opinion  ��������    500
'    20  intend_fee     Ԥ�Ʒ���    10
'    21  apply_date     ������Ч����        ��ʽ��YYYY-MM-DD
'    22  admit_date     �����Ч����        ��ʽ��YYYY-MM-DD
'    23  audit_date     ��������        ��ʽ��YYYY-MM-DD
'    24  audit_flag     ������־    10
'    25  input_man      ¼��������  10
'    26  input_date     ¼������        ��ʽ��YYYY-MM-DD
'    27  note           ��ע    500
    '���ӿڷ��ص������ӳ���¼����
    If ���ýӿ�_��¼��_������ Then
        Call DebugTool("���ؼ�¼������" & CZ_GetRowCount(glngInterface_������))
        Call ���ýӿ�_�ƶ���¼��_������(MoveFirst)
        Do While True
            strValues = ""
            For intCol = 0 To intCols
                'todo �˴�����ʽ���룬��Ҫȡ��ע��
                Call ���ýӿ�_��ȡ����_������(arrColumn(intCol), strData)
                strValues = strValues & "|" & strData
            Next
            strValues = Mid(strValues, 2)
            Call Record_Add(rsVerify, strColumns, strValues)
            
            Call DebugTool("�Ѽ���һ�м�¼")
            If Not ���ýӿ�_�ƶ���¼��_������(MoveNext) Then Exit Do
        Loop
    End If
    
    '������
    If rsVerify.RecordCount = 0 Then
        Call DebugTool("�ڲ���¼���ļ�¼��Ϊ��")
        With msh����嵥
            .Clear
            .Rows = 2
            For intCol = 0 To .Cols - 1
                .TextMatrix(1, intCol) = ""
            Next
        End With
    Else
        Call DebugTool("��ʾ�ڲ���¼���е�����")
        Set msh����嵥.DataSource = rsVerify
    End If
    
    '������ͷ
    Call DebugTool("�������ñ�ͷ")
    strColumns = "����|�Ա�|���֤��|��������|ת��ҽԺ|��������|��˱�־"
    arrColumn = Split(strColumns, "|")
    With msh����嵥
        For intCol = 0 To .Cols - 1
            .TextMatrix(0, intCol) = arrColumn(intCol)
            .ColAlignmentFixed(intCol) = 4
        Next
    End With
    Call zlControl.MshSetColWidth(msh����嵥, Me)
StopFlash:
    Call zlcommfun.StopFlash
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intTYPE As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intTYPE = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intTYPE
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intTYPE, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
