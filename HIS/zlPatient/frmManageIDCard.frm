VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmManageIDCard 
   AutoRedraw      =   -1  'True
   Caption         =   "���﷢������"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8850
   Icon            =   "frmManageIDCard.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4710
      Left            =   45
      TabIndex        =   3
      Top             =   780
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   8308
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
      MouseIcon       =   "frmManageIDCard.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8850
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   8730
         _ExtentX        =   15399
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
            NumButtons      =   12
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
               Key             =   "IDCard"
               Description     =   "����"
               Object.ToolTipText     =   "���뷢������"
               Object.Tag             =   "����"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˿�"
               Key             =   "Del"
               Description     =   "�˿�"
               Object.ToolTipText     =   "�Ե�ǰѡ�м�¼�˿�"
               Object.Tag             =   "�˿�"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "�����������¶�ȡ�б�"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�ڵ�ǰ�б������������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      TabIndex        =   0
      Top             =   5490
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageIDCard.frx":075C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10530
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":0FF0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":120A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1424
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":163E
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1858
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":1FD2
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":21EC
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2406
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2620
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":283A
            Key             =   "Report"
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2A54
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2C6E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":2E88
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":30A2
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":32BC
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3A36
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3C50
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":3E6A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":4084
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageIDCard.frx":429E
            Key             =   "Report"
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
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileWorkReport 
         Caption         =   "��ӡ�ɿ���(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_IDCard 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "�˿�(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "����(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPass 
         Caption         =   "�޸�����(&P)"
         Shortcut        =   {F4}
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
      Begin VB.Menu mnuView_5 
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
      Begin VB.Menu mnuViewReFlash 
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
Attribute VB_Name = "frmManageIDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''Option Explicit 'Ҫ���������
''''Private mrsList As ADODB.Recordset  '�����б�
''''Private mstrFilter As String
''''Private mblnCancel As Boolean
''''Private mblnGo As Boolean, mlngGo As Long
''''Private mlngCurRow As Long, mlngTopRow As Long
''''Private mstrPrivs As String
''''Private mlngModul As Long
''''Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
''''Private mcllFilterA As Collection
'''''by lesfeng 2010-1-11 �����Ż�
''''Private Sub InitFilter()
''''    '-----------------------------------------------------------------------------------------------------------
''''    '����:��ʼ����������
''''    '���:
''''    '����:
''''    '����:
''''    '����:lesfeng
''''    '����:2010-01-11 16:10:40
''''    '-----------------------------------------------------------------------------------------------------------
''''    Set mcllFilterA = New Collection
''''    mcllFilterA.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "�Ǽ�ʱ��"
''''    mcllFilterA.Add Array("", ""), "���ݺ�"
''''    mcllFilterA.Add Array("", ""), "Ʊ�ݺ�"
''''    mcllFilterA.Add "", "סԺ��"
''''    mcllFilterA.Add "", "����"
''''    mcllFilterA.Add "", "��¼״̬"
''''    mcllFilterA.Add "", "���ӱ�־"
''''    mcllFilterA.Add "", "�տ���"
''''    mstrFilter = ""
''''End Sub
''''
''''Private Sub Form_Activate()
''''    Call InitLocPar(mlngModul)
''''End Sub
''''
''''Private Sub mnuEditPass_Click()
''''    frmModiPass.Show 1, Me
''''End Sub
''''
''''Private Sub mnuFileLocalSet_Click()
''''    Call frmLocalSet.zlSetPara(Me, mstrPrivs, mlngModul)
''''    If glng�ſ�ID > 0 Then
''''        If Not ExistBill(glng�ſ�ID, 5) Then
''''            zldatabase.SetPara "���þ��￨����", 0, glngSys, mlngModul
''''            glng�ſ�ID = 0
''''        End If
''''    End If
''''End Sub
''''
''''Private Sub mnuFileWorkReport_Click()
''''    Call frmWorkTime.ShowMe(Me, 5)
''''End Sub
''''
''''Private Sub mnuReportItem_Click(Index As Integer)
''''    Dim strNO As String, strTmp As String
''''
''''    strNO = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
''''    If strNO <> "" Then
''''        With mshList
''''            If glngSys Like "8??" Then
''''                strTmp = "�ͻ�ID"
''''            Else
''''                strTmp = "����ID"
''''            End If
''''            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
''''                    "NO=" & strNO, "���￨��=" & .TextMatrix(.Row, GetColNum("����")), _
''''                    strTmp & "=" & .TextMatrix(.Row, GetColNum(strTmp)))
''''        End With
''''    Else
''''        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
''''    End If
''''End Sub
''''
''''Private Sub mnuViewFilter_Click()
''''    frmIDCardFilter.Show 1, Me
''''    If gblnOK Then
''''        mstrFilter = frmIDCardFilter.mstrFilter
''''        'by lesfeng 2010-03-08 �����Ż�
''''        Set mcllFilterA = frmIDCardFilter.mcllFilter
''''        mblnCancel = (frmIDCardFilter.chkCancel.Value = Checked)
''''        mnuViewReFlash_Click
''''    End If
''''End Sub
''''
''''Private Sub mshList_DblClick()
''''    If mshList.MouseRow = 0 Then Exit Sub
''''    If mnuEdit_View.Enabled Then mnuEdit_View_Click
''''End Sub
''''
''''Private Sub mshList_EnterCell()
''''    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, 0) = "" Then Exit Sub
''''    mlngGo = mshList.Row
''''    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
''''
''''    If frmIDCardFilter.mblnDateMoved Then
''''        mblnNOMoved = zldatabase.NOMoved("סԺ���ü�¼", mshList.TextMatrix(mshList.Row, 0), , "5", Me.Caption)
''''    Else
''''        mblnNOMoved = False
''''    End If
''''
''''End Sub
''''
''''Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
''''    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
''''End Sub
''''
''''Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then PopupMenu mnuEdit, 2
''''End Sub
''''
''''Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''    Select Case KeyCode
''''        Case vbKeyF3
''''            'ʼ�մӵ�ǰ�п�ʼ
''''            If mnuViewGo.Enabled Then Call SeekBill(False)
''''        Case vbKeyReturn
''''            If mnuEdit_View.Enabled Then mnuEdit_View_Click
''''        Case vbKeyEscape
''''            mblnGo = False
''''    End Select
''''End Sub
''''
''''Private Sub mnuEdit_Del_Click()
''''    If mshList.TextMatrix(mshList.Row, 0) = "" Then
''''        MsgBox "��ǰû�м�¼�����˿���", vbExclamation, gstrSysName
''''        Exit Sub
''''    End If
''''
''''    '����Ȩ��
''''    If Not BillOperCheck(8, mshList.TextMatrix(mshList.Row, GetColNum("������")), _
''''        CDate(mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))), "�˿�") Then Exit Sub
''''
''''    On Error Resume Next
''''    Err.Clear
''''
''''    '�Ƿ���ת������ݱ���
''''    If mblnNOMoved Then
''''        If Not ReturnMovedExes(mshList.TextMatrix(mshList.Row, 0), 5, Me.Caption) Then Exit Sub
''''        mblnNOMoved = False  '��ʱ��ת���������ݱ�
''''    End If
''''
''''    frmIDCard.mbytInState = 2
''''    frmIDCard.mstrInNO = mshList.TextMatrix(mshList.Row, 0)
''''    frmIDCard.Show 1, Me
''''    If gblnOK Then Call mnuViewReFlash_Click
''''End Sub
''''
''''Private Sub mnuHelpTitle_Click()
''''ShowHelp App.ProductName, Me.hwnd, Me.Name
''''End Sub
''''
''''Private Sub mnuEdit_IDCard_Click()
''''    On Error Resume Next
''''    Err.Clear
''''
''''    frmIDCard.mbytInState = 0
''''    frmIDCard.Show 1, Me
''''    If gblnOK Then mnuViewReFlash_Click
''''End Sub
''''
''''Private Sub mnuEdit_View_Click()
''''    If mshList.TextMatrix(mshList.Row, 0) = "" Then
''''        MsgBox "��ǰû�м�¼���Բ��ģ�", vbExclamation, gstrSysName
''''        Exit Sub
''''    End If
''''
''''    On Error Resume Next
''''    Err.Clear
''''    '��ʾ��������
''''    frmIDCard.mbytInState = 1
''''    If mblnCancel Then frmIDCard.mblnViewCancel = True
''''    frmIDCard.mstrInNO = mshList.TextMatrix(mshList.Row, 0)
''''    frmIDCard.mblnNOMoved = mblnNOMoved
''''    frmIDCard.Show 1, Me
''''End Sub
''''
''''Private Sub mnuFile_Quit_Click()
''''    Unload Me
''''End Sub
''''
''''Private Sub mnuHelpAbout_Click()
''''    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
''''End Sub
''''
''''Private Sub mnuViewReFlash_Click()
''''    ShowBills mstrFilter
''''End Sub
''''
''''Private Sub mnuViewStatus_Click()
''''    mnuViewStatus.Checked = Not mnuViewStatus.Checked
''''    stbThis.Visible = Not stbThis.Visible
''''    Form_Resize
''''End Sub
''''
''''Private Sub mnuViewToolButton_Click()
''''    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
''''    cbr.Visible = Not cbr.Visible
''''    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
''''    Form_Resize
''''End Sub
''''
''''Private Sub mnuViewToolText_Click()
''''    Dim i As Integer
''''    mnuViewToolText.Checked = Not mnuViewToolText.Checked
''''    For i = 1 To tbr.Buttons.Count
''''        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
''''    Next
''''    cbr.Bands(1).MinHeight = tbr.ButtonHeight
''''    Form_Resize
''''End Sub
''''
''''Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
''''    Select Case Button.Key
''''        Case "Quit"
''''            mnuFile_Quit_Click
''''        Case "Go" '��λ
''''            mnuViewGo_Click
''''        Case "Filter" '����
''''            mnuViewFilter_Click
''''        Case "View"
''''            mnuEdit_View_Click
''''        Case "IDCard"
''''            mnuEdit_IDCard_Click
''''        Case "Del"
''''            mnuEdit_Del_Click
''''        Case "Print"
''''            mnuFile_Print_Click
''''        Case "Preview"
''''            mnuFile_PreView_Click
''''        Case "Help"
''''            mnuHelpTitle_Click
''''    End Select
''''End Sub
''''
''''Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If Button = 2 Then PopupMenu mnuViewTool, 2
''''End Sub
''''
''''Private Sub mnuFile_Excel_Click()
''''    Call OutputList(3)
''''End Sub
''''
''''Private Sub mnuFile_PreView_Click()
''''    Call OutputList(2)
''''End Sub
''''
''''Private Sub mnuFile_Print_Click()
''''    Call OutputList(1)
''''End Sub
''''
''''Private Sub mnuFile_PrintSet_Click()
''''    Call zlPrintSet
''''End Sub
''''
''''Private Sub OutputList(bytStyle As Byte)
'''''���ܣ�������б�
'''''������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
''''    Dim objOut As New zlPrint1Grd
''''    Dim objRow As New zlTabAppRow
''''    Dim bytR As Byte, intRow As Integer
''''
''''    intRow = mshList.Row
''''
''''    '��ͷ
''''    If glngSys Like "8??" Then
''''        objOut.Title.Text = "��Ա�����ŵ��嵥"
''''    Else
''''        objOut.Title.Text = "���￨���ŵ��嵥"
''''    End If
''''    objOut.Title.Font.Name = "����_GB2312"
''''    objOut.Title.Font.Size = 18
''''    objOut.Title.Font.Bold = True
''''
''''    '����
''''    With frmIDCardFilter
''''        If IsNull(.dtpEnd.Value) Then
''''            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd")
''''        Else
''''            objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, "yyyy-MM-dd HH:MM") & " �� " & Format(.dtpEnd.Value, "yyyy-MM-dd HH:MM")
''''        End If
''''        objRow.Add "���ʣ�" & IIf(.chkCancel.Value = 1, "�˿���¼", "������¼")
''''        objOut.UnderAppRows.Add objRow
''''    End With
''''
''''    Set objRow = New zlTabAppRow
''''    objRow.Add "��ӡ�ˣ�" & UserInfo.����
''''    objRow.Add "��ӡ���ڣ�" & Format(zldatabase.Currentdate(), "yyyy��MM��dd��")
''''    objOut.BelowAppRows.Add objRow
''''
''''    '����
''''    mshList.Redraw = False
''''    Set objOut.Body = mshList
''''
''''    '���
''''    If bytStyle = 1 Then
''''        bytR = zlPrintAsk(objOut)
''''        Me.Refresh
''''        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
''''    Else
''''        zlPrintOrView1Grd objOut, bytStyle
''''    End If
''''
''''    mshList.Row = intRow
''''    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
''''    mshList.Redraw = True
''''End Sub
''''
''''Private Sub mnuHelpWebHome_Click()
''''    zlHomePage hwnd
''''End Sub
''''
''''Private Sub mnuHelpWebMail_Click()
''''    zlMailTo hwnd
''''End Sub
''''
''''Private Sub SetHeader()
''''    Dim i As Integer
''''    With mshList
''''        .Redraw = False
''''        .Cols = 12
''''        .TextMatrix(0, 0) = "���ݺ�"
''''        If mblnCancel Then
''''            .TextMatrix(0, 1) = "�˿�ʱ��"
''''        Else
''''            .TextMatrix(0, 1) = "����ʱ��"
''''        End If
''''        .TextMatrix(0, 2) = "����"
''''        .TextMatrix(0, 3) = "����"
''''        If glngSys Like "8??" Then
''''            .TextMatrix(0, 4) = "�ͻ�ID"
''''        Else
''''            .TextMatrix(0, 4) = "����ID"
''''        End If
''''        .TextMatrix(0, 5) = "��ʶ��"
''''        .TextMatrix(0, 6) = "����"
''''        .TextMatrix(0, 7) = "�Ա�"
''''        .TextMatrix(0, 8) = "����"
''''        .TextMatrix(0, 9) = "���"
''''        .TextMatrix(0, 10) = "����"
''''        If mblnCancel Then
''''            .TextMatrix(0, 11) = "�˿���"
''''        Else
''''            .TextMatrix(0, 11) = "������"
''''        End If
''''
''''        .ColAlignment(0) = 4
''''        .ColAlignment(1) = 4
''''        .ColAlignment(2) = 1
''''        .ColAlignment(3) = 4
''''        .ColAlignment(4) = 1
''''        .ColAlignment(5) = 1
''''        .ColAlignment(6) = 1
''''        .ColAlignment(7) = 4
''''        .ColAlignment(8) = 4
''''        .ColAlignment(9) = 7
''''        .ColAlignment(10) = 4
''''        .ColAlignment(11) = 1
''''
''''        If Not Visible Then
''''            .ColWidth(0) = 850
''''            .ColWidth(1) = 1000
''''            .ColWidth(2) = 850
''''            .ColWidth(3) = 500
''''            .ColWidth(4) = 750
''''            If glngSys Like "8??" Then
''''                .ColWidth(5) = 0
''''            Else
''''                .ColWidth(5) = 750
''''            End If
''''            .ColWidth(6) = 800
''''            .ColWidth(7) = 500
''''            .ColWidth(8) = 500
''''            .ColWidth(9) = 850
''''            .ColWidth(10) = 500
''''            .ColWidth(11) = 800
''''        End If
''''
''''        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
''''
''''        .RowHeight(0) = 320
''''        For i = 0 To .Cols - 1
''''            .ColAlignmentFixed(i) = 4
''''        Next
''''        '�ָ��ϴ���
''''        If mlngCurRow = 0 Then mlngCurRow = 1
''''        If mlngTopRow = 0 Then mlngTopRow = 1
''''        If mlngCurRow <= .Rows - 1 Then
''''            .Row = mlngCurRow
''''        Else
''''            .Row = .Rows - 1
''''        End If
''''        If mlngTopRow <= .Rows - 1 Then
''''            .TopRow = mlngTopRow
''''        Else
''''            .TopRow = .Row
''''        End If
''''
''''         .Col = 0: .ColSel = .Cols - 1
''''        Call mshList_EnterCell
''''
''''        .Redraw = True
''''    End With
''''End Sub
''''
''''Private Sub ShowBills(Optional ByVal strIF As String, Optional blnSort As Boolean)
'''''����:��������ȡ�����б�(���˹���)
'''''����:strIF=��"AND"��ʼ��������
''''    Dim strCard As String, i As Long
''''
''''    On Error GoTo errH
''''
''''    If Not blnSort Then
''''        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
''''        DoEvents
''''        Me.Refresh
''''
''''        strIF = " Where ��¼����=5 " & strIF
''''        'by lesfeng 2010-03-08 �����Ż�
''''        If frmIDCardFilter.mblnDateMoved Then
''''            strIF = "" & _
''''            " Select NO,�Ǽ�ʱ��,ʵ��Ʊ��,���ӱ�־,����id,��ʶ��,����,�Ա�,����,ʵ�ս��,���ʷ���,����Ա���� " & _
''''            " From סԺ���ü�¼ " & strIF & _
''''            " UNION ALL " & _
''''            " Select NO,�Ǽ�ʱ��,ʵ��Ʊ��,���ӱ�־,����id,��ʶ��,����,�Ա�,����,ʵ�ս��,���ʷ���,����Ա���� " & _
''''            " From HסԺ���ü�¼ " & strIF
''''        Else
''''            strIF = "Select NO,�Ǽ�ʱ��,ʵ��Ʊ��,���ӱ�־,����id,��ʶ��,����,�Ա�,����,ʵ�ս��,���ʷ���,����Ա���� From סԺ���ü�¼ " & strIF
''''        End If
''''
''''        strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,A.ʵ��Ʊ��,LPAD('*',Length(A.ʵ��Ʊ��),'*')) as ����,"
''''        gstrSQL = _
''''        " Select A.NO as ���ݺ�,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as " & IIf(mblnCancel, "�˿�", "����") & "ʱ��," & strCard & _
''''        "           Decode(A.���ӱ�־,1,'����',2,'����','����') as ����,A.����ID,A.��ʶ�� as סԺ��,A.����,A.�Ա�,A.����," & _
''''        "           To_Char(" & IIf(mblnCancel, " - ", "") & "Sum(A.ʵ�ս��),'99990.00') as ���," & _
''''        "           Decode(Nvl(A.���ʷ���,0),0,NULL,'��') as ����," & _
''''        "           A.����Ա���� as " & IIf(mblnCancel, "�˿���", "������") & " " & _
''''        " From (" & strIF & ") A " & _
''''        " Group by A.NO,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD'),A.ʵ��Ʊ��,Decode(A.���ӱ�־,1,'����',2,'����','����')," & _
''''        "           A.����ID,A.��ʶ��,A.����,A.�Ա�,A.����,Decode(Nvl(A.���ʷ���,0),0,NULL,'��'),A.����Ա����" & _
''''        " Order by " & IIf(mblnCancel, "�˿�", "����") & "ʱ�� Desc,���ݺ� Desc"
''''        Set mrsList = New ADODB.Recordset
''''
''''        Set mrsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CDate(mcllFilterA("�Ǽ�ʱ��")(0)), CDate(mcllFilterA("�Ǽ�ʱ��")(1)), _
''''        CStr(Val(mcllFilterA("���ݺ�")(0))), CStr(Val(mcllFilterA("���ݺ�")(1))), _
''''        CStr(Val(mcllFilterA("Ʊ�ݺ�")(0))), CStr(Val(mcllFilterA("Ʊ�ݺ�")(1))), CLng(Val(mcllFilterA("סԺ��"))), _
''''        CStr(mcllFilterA("����")), CLng(Val(mcllFilterA("��¼״̬"))), CLng(Val(mcllFilterA("���ӱ�־"))), CStr(mcllFilterA("�տ���")))
''''
''''    End If
''''
''''    mshList.Clear
''''    mshList.Rows = 2
''''
''''    mshList.ForeColor = IIf(mblnCancel, &HC0, ForeColor)
''''
''''    If mrsList.EOF Then
''''        Call SetHeader
''''        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
''''        Call SetMenu(False)
''''    Else
''''        Set mshList.DataSource = mrsList
''''        Call SetHeader
''''        stbThis.Panels(2) = "�� " & mrsList.RecordCount & " �ŵ���"
''''        Call SetMenu(True)
''''    End If
''''
''''    mnuEdit_Del.Enabled = Not mblnCancel And Not mrsList.EOF
''''    tbr.Buttons("Del").Enabled = Not mblnCancel And Not mrsList.EOF
''''
''''    If Not blnSort Then Call zlCommFun.StopFlash
''''
''''    Me.Refresh
''''    Exit Sub
''''errH:
''''    If errCenter() = 1 Then Resume
''''    Call SaveErrLog
''''End Sub
''''
''''Private Sub SetMenu(blnUsed As Boolean)
'''''���ܣ��������޼�¼���ò˵�����״̬
''''    mnuFile_Print.Enabled = blnUsed
''''    mnuFile_Preview.Enabled = blnUsed
''''    mnuFile_Excel.Enabled = blnUsed
''''    tbr.Buttons("Print").Enabled = blnUsed
''''    tbr.Buttons("Preview").Enabled = blnUsed
''''
''''    mnuEdit_Del.Enabled = blnUsed
''''    mnuEdit_View.Enabled = blnUsed
''''    tbr.Buttons("Del").Enabled = blnUsed
''''    tbr.Buttons("View").Enabled = blnUsed
''''
''''    mnuViewGo.Enabled = blnUsed
''''    tbr.Buttons("Go").Enabled = blnUsed
''''End Sub
''''
''''Private Sub Form_Load()
''''    Dim curDate As Date
''''    'by lesfeng 2010-03-08 �����Ż�
''''    Call InitFilter
''''
''''    mstrPrivs = gstrPrivs
''''    mlngModul = glngModul
''''    Call zldatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
''''
''''    If glngSys Like "8??" Then Caption = "��Ա�����Ź���"
''''
''''    Call RestoreWinState(Me, App.ProductName)
''''
''''    If glng�ſ�ID > 0 Then
''''        If Not ExistBill(glng�ſ�ID, 5) Then
''''            zldatabase.SetPara "���þ��￨����", 0, glngSys, mlngModul
''''            glng�ſ�ID = 0
''''        End If
''''    End If
''''
''''    'Ȩ������(��Ϊ���ɼ�)
''''    If InStr(mstrPrivs, "��������") = 0 Then
''''        mnuEdit_IDCard.Visible = False
''''        mnuEdit_Del.Visible = False
''''        tbr.Buttons("IDCard").Visible = False
''''        tbr.Buttons("Del").Visible = False
''''    End If
''''    If InStr(mstrPrivs, "�޸�����") = 0 Then
''''        mnuEdit_1.Visible = False
''''        mnuEditPass.Visible = False
''''    End If
''''
''''    'ȱʡ��������
''''    curDate = zldatabase.Currentdate
''''    'by lesfeng 2010-03-08 �����Ż�
''''    mstrFilter = ""
''''    mstrFilter = mstrFilter & " And (�Ǽ�ʱ��  Between [1] And [2]) "
''''    mstrFilter = mstrFilter & " And ��¼״̬=[9]"
''''    mstrFilter = mstrFilter & " And ����Ա����=[11]"
''''
''''    mcllFilterA.Remove "�Ǽ�ʱ��"
''''    mcllFilterA.Add Array(Format(DateAdd("d", -7, curDate), "yyyy-mm-dd") & " 00:00:00", Format(curDate, "yyyy-mm-dd") & " 23:59:59"), "�Ǽ�ʱ��"
''''    mcllFilterA.Remove "��¼״̬"
''''    mcllFilterA.Add "1", "��¼״̬"
''''    mcllFilterA.Remove "�տ���"
''''    mcllFilterA.Add Trim(UserInfo.����), "�տ���"
''''
''''    mblnCancel = False
''''
''''    Call SetHeader
''''    Call SetMenu(False)
''''
''''    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
''''End Sub
''''
''''Private Sub Form_Resize()
''''    Dim cbrH As Long '������ռ�ø߶�
''''    Dim staH As Long '״̬��ռ�ø߶�
''''
''''    On Error Resume Next
''''
''''    If WindowState = 1 Then Exit Sub
''''
''''    mshList.MousePointer = 0
''''
''''    '����ؼ���Ⱥ͸߶�
''''    cbrH = IIf(cbr.Visible, cbr.Height, 0)
''''    staH = IIf(stbThis.Visible, stbThis.Height, 0)
''''    With mshList
''''        .Left = Me.ScaleLeft
''''        .Top = Me.ScaleTop + cbrH
''''        .Width = Me.ScaleWidth
''''        .Height = Me.ScaleHeight - cbrH - staH
''''    End With
''''End Sub
''''
''''Private Sub Form_Unload(Cancel As Integer)
''''    mstrFilter = ""
''''    Unload frmIDCardFilter
''''    Unload frmIDCardFind
''''    Call SaveWinState(Me, App.ProductName)
''''End Sub
''''
''''Private Sub mnuViewGo_Click()
''''    If Not mblnCancel Then
''''        frmIDCardFind.lbl����Ա.Caption = "������"
''''    Else
''''        frmIDCardFind.lbl����Ա.Caption = "�˿���"
''''    End If
''''    frmIDCardFind.Show 1, Me
''''    If gblnOK Then Call SeekBill(frmIDCardFind.optHead)
''''End Sub
''''
''''Private Sub SeekBill(blnHead As Boolean)
''''    Dim i As Long
''''    Dim blnFill As Boolean
''''
''''    Screen.MousePointer = 11
''''    mblnGo = True
''''    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
''''    Me.Refresh
''''
''''    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
''''        DoEvents
''''
''''        '�Ƚ�����
''''        blnFill = True
''''        With frmIDCardFind
''''            If .txtNO.Text <> "" Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 0) = .txtNO.Text
''''            End If
''''            If .txtCard.Text <> "" Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 2) = .txtCard.Text
''''            End If
''''            If .cbo����Ա.ListIndex > 0 Then
''''                blnFill = blnFill And mshList.TextMatrix(i, 11) = NeedName(.cbo����Ա.Text)
''''            End If
''''            If .txt����.Text <> "" Then
''''                blnFill = blnFill And UCase(mshList.TextMatrix(i, 6)) Like "*" & UCase(.txt����.Text) & "*"
''''            End If
''''            If IsNumeric(.txtסԺ��.Text) Then
''''                blnFill = blnFill And Val(mshList.TextMatrix(i, 5)) = Val(.txtסԺ��.Text)
''''            End If
''''        End With
''''
''''        '�������˳�
''''        If blnFill Then
''''            mlngGo = i + 1
''''            mshList.Row = i: mshList.TopRow = i
''''            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
''''            stbThis.Panels(2).Text = "�ҵ�һ����¼"
''''            Screen.MousePointer = 0: Exit Sub
''''        End If
''''
''''        '��ESCȡ��
''''        If mblnGo = False Then
''''            stbThis.Panels(2).Text = "�û�ȡ����λ����"
''''            Screen.MousePointer = 0: Exit Sub
''''        End If
''''    Next
''''    mlngGo = 1
''''    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
''''    Screen.MousePointer = 0
''''End Sub
''''
''''Private Function GetColNum(strHead As String) As Integer
''''    Dim i As Integer
''''    For i = 0 To mshList.Cols - 1
''''        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
''''    Next
''''End Function
''''
''''Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    If mshList.MouseRow = 0 Then
''''        mshList.MousePointer = 99
''''    Else
''''        mshList.MousePointer = 0
''''    End If
''''End Sub
''''
''''Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''    Dim lngCol As Long
''''
''''    lngCol = mshList.MouseCol
''''
''''    If Button = 1 And mshList.MousePointer = 99 Then
''''        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
''''        If mshList.TextMatrix(1, GetColNum("���ݺ�")) = "" Then Exit Sub
''''
''''        Set mshList.DataSource = Nothing
''''        If mshList.TextMatrix(0, lngCol) = "�ͻ�ID" Then
''''            mrsList.Sort = "����ID" & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
''''        Else
''''            mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
''''        End If
''''        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
''''
''''        Call ShowBills(, True)
''''    End If
''''End Sub
''''
''''Private Sub mnuHelpWebForum_Click()
''''    '-----------------------------------------------------------------------------
''''    '����:���ӵ�������̳
''''    '�޸���:���˺�
''''    '�޸�����:2006-12-11
''''    '-----------------------------------------------------------------------------
''''    Call zlWebForum(Me.hwnd)
''''End Sub
''''
