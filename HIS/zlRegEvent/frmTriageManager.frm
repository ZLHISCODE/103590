VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmTriageManager 
   BorderStyle     =   0  'None
   Caption         =   "����������"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   600
      ScaleHeight     =   3030
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   600
      Width           =   3735
      Begin MSComctlLib.ListView lvwMain 
         Height          =   1770
         Left            =   -105
         TabIndex        =   6
         Tag             =   "�ɱ仯��"
         Top             =   1455
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvwYY 
         Height          =   1770
         Left            =   1695
         TabIndex        =   7
         Tag             =   "�ɱ仯��"
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4875
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picHZPati 
      BorderStyle     =   0  'None
      Height          =   2430
      Left            =   330
      ScaleHeight     =   2430
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   4065
      Width           =   2895
      Begin MSComctlLib.ListView lvwHZPati 
         Height          =   1770
         Left            =   60
         TabIndex        =   4
         Tag             =   "�ɱ仯��"
         Top             =   30
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   4815
      ScaleHeight     =   5250
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   585
      Width           =   4485
      Begin MSComctlLib.ListView lvwRoom 
         Height          =   4230
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��������"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "״̬"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "����"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "��������"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "����"
            Object.Width           =   2293
         EndProperty
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTittl 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2385
         _Version        =   589884
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "��ǰ������״̬"
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
   End
   Begin MSComctlLib.ImageList img161 
      Left            =   885
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483637
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0000
            Key             =   "ry"
            Object.Tag             =   "ry"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":059A
            Key             =   "yf"
            Object.Tag             =   "yf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0934
            Key             =   "zz"
            Object.Tag             =   "zz"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0ECE
            Key             =   "yz"
            Object.Tag             =   "yz"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1468
            Key             =   "bm"
            Object.Tag             =   "bm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1A02
            Key             =   "WomanStop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1F9C
            Key             =   "ManStop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":2536
            Key             =   "WomanSign_in"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":2AD0
            Key             =   "ManSign_in"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":306A
            Key             =   "rySign_in"
            Object.Tag             =   "rySign_in"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":3604
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":3958
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTriageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModul As Long, mintFindKeys As Integer
Private mbytViewScrop(0 To 3) As Byte  '0-��ʾ�ѷ��ﲡ��;1-��ʾ�ѽ��ﲡ��;2-��ʾ����ɲ���;3-��ʾ�����ﲡ��
Private mbyt��������ʽ As Byte  '���ﲡ�˵�����ʽ,0-���ұ���,����,���ݺ�;1-���ұ���,����,�Һ�ʱ��;
Private mstr�������   As String               '�Զ��ŷָ��ķ������id��,�ձ�ʾ���п���,0��ʾû��ѡ���κο���
Private mfrmMain As Form     '���õĸ�����
Private mlngOutModeMC As Long    '����ҽ�����õ����ʽҽ������
Private Const STR_COMP = "|',~" '�ָ��ַ���
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    DeptID As Long
    Patient As String
    Operator As String
    ����� As String
    ���￨�� As String
    ҽ���� As String
End Type
Private mSQLCondition As Type_SQLCondition
'-----------------------------------------------------------------------------------
'��Ϣ��ر���
Private mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private mcllFilter As Collection

Private mint��Ч���� As Integer
Private mlngPre����ID As Long   '�ϴβ���ID
Private mbytIDKind As Byte  '0-�����;1-����;2-�Һŵ�;3-���￨��;4-ҽ����
Private mlngDefaultCardID As Long 'Ĭ�Ͽ����

Private Const conPane_PatiList = 1
Private Const conPane_Room = 2
Private Const conPane_PatiHZ = 3

Private Enum midx
    idx_�ŶӶ��� = 0
    idx_ԤԼ���� = 1
End Enum

Public Event zlPopuMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event zlShowInfor(strShowInfor As String)
Public Event zlQueueAsk(intType As Integer, strNO As String, lng����ID As Long, Cancel As Boolean)
'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����;7-����
'strNO:-���ݺ�
' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥;7����

Private mcbsThis As Object
Private mobjPublicPatient As Object
Private Enum EnmCol
    Enm�Һŵ� = 0
    Enm����״̬ = 1
    Enm���� = 2
    Enm���� = 3
    Enm�Һ���Ŀ = 4
    Enm����� = 5
    Enm���� = 6
    Enm�Ա� = 7
    Enm���� = 8
    Enm���� = 9
    Enmҽ�� = 10
    Enm����ʱ�� = 11
    Enm�Һ�ʱ�� = 12
    Enm���� = 13
    Enmҽ���� = 14
    EnmժҪ = 15
    EnmԤԼ = 16
    Enm���� = 17
    Enm������ = 18
    Enm�������� = 19
    Enm����ʱ�� = 20
End Enum
'-----------------------------------------------------------------------------------
Private mstrRegistIdsed As String '�Ѿ�ˢ�µĹҺ�ID,����ö��ŷ���,��Ϣ����ʱ��Ч


Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case conPane_PatiList
        Item.Handle = picMain.Hwnd
    Case conPane_Room
        Item.Handle = picDept.Hwnd
    Case conPane_PatiHZ
        Item.Handle = picHZPati.Hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If mfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.ActiveIDKindKey
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panLeft As Pane
    
    Set panLeft = dkpMan.CreatePane(conPane_PatiList, 200, 580, DockLeftOf, Nothing)
    panLeft.Title = "�����б�": panLeft.Tag = conPane_PatiList
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Handle = picMain.Hwnd
    
    Set panThis = dkpMan.CreatePane(conPane_Room, 250, 580, DockRightOf, panLeft)
    panThis.Title = "�������"
    panThis.Tag = conPane_Room
    panThis.Handle = picDept.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(conPane_PatiHZ, 200, 580, DockBottomOf, panLeft)
    panThis.Title = "���ﲡ���б�": panThis.Tag = panThis
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picHZPati.Hwnd
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Sub

 

Private Sub picMain_Resize()
    Err = 0: On Error Resume Next
    With picMain
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub picHZPati_Resize()
    Err = 0: On Error Resume Next
    With picHZPati
        lvwHZPati.Left = .ScaleLeft
        lvwHZPati.Width = .ScaleWidth
        lvwHZPati.Top = .ScaleTop
        lvwHZPati.Height = .ScaleHeight
    End With
End Sub
Private Sub picDept_Resize()
    Err = 0: On Error Resume Next
    With picDept
        lvwRoom.Left = .ScaleLeft
        lvwRoom.Width = .ScaleWidth
        stcTittl.Top = .ScaleTop: stcTittl.Left = .ScaleLeft
        stcTittl.Width = .ScaleWidth
        lvwRoom.Top = stcTittl.Top + stcTittl.Height
        lvwRoom.Height = .ScaleHeight - lvwRoom.Top
    End With
End Sub
Public Sub zlInitVar(ByVal frmMain As Form, ByVal byt��������ʽ As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ó��ñ���
    '��Σ�frmMain-���õĸ�����
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-01 17:21:15
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mbyt��������ʽ = byt��������ʽ
    Set mfrmMain = frmMain
End Sub
Public Sub zlExcuteReport(ByVal lngSys As Long, ByVal strReportNO As String)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ִ�б���
    '���ƣ����˺�
    '���ڣ�2010-06-01 15:53:17
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If Not lvwMain.SelectedItem Is Nothing Then
         With lvwMain.SelectedItem
             Call ReportOpen(gcnOracle, lngSys, strReportNO, Me, _
                 "NO=" & .Text, "�����=" & .SubItems(EnmCol.Enm�����), _
                 "ҽ��=" & .SubItems(EnmCol.Enmҽ��), "ִ�п���=" & .ListSubItems(2).Tag)
         End With
     Else
         Call ReportOpen(gcnOracle, lngSys, strReportNO, Me)
     End If
End Sub

Public Sub zlExcǩ��(ByVal blnȡ��ǩ�� As Boolean, Optional ByVal blnClick As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ǩ��
    '����:���˺�
    '����:2010-12-08 10:56:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLvw As ListView, blnHZ As Boolean '�Ƿ����
    Dim blnTriage As Boolean, lng����ID  As Long, lngExeState As Long
    Dim lngID As Long, strTittle As String
    Dim bln�ѷ��� As Boolean, strNO As String, bln�Ƿ�ԤԼ As Boolean
    Dim str����ʱ�� As String
    Dim strDoc As String, strRoom As String
    Dim bln����̨ǩ���Ŷ� As Boolean
    
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        Set objLvw = lvwMain
        If objLvw.SelectedItem Is Nothing Then Exit Sub
        If Val(Split(objLvw.SelectedItem.Tag, "|")(8)) = 0 Then
            'ת�ﲡ�˶����ѷ��ﲡ��
            bln�ѷ��� = True
        End If
        bln����̨ǩ���Ŷ� = objLvw.SelectedItem.ListSubItems(4).Tag = 1
    End If
    
    If tbPage.Item(midx.idx_ԤԼ����).Selected Then
        Set objLvw = LvwYY
        If objLvw.SelectedItem Is Nothing Then Exit Sub
        bln�Ƿ�ԤԼ = True
        bln����̨ǩ���Ŷ� = objLvw.SelectedItem.ListSubItems(4).Tag = 1
    End If
    
    lng����ID = Val(Split(objLvw.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(objLvw.SelectedItem.Tag, "|")(6)) 'ִ��״̬
    
    blnTriage = (lngExeState = 0)
    lngID = Val(objLvw.SelectedItem.ListSubItems(1).Tag)
    strDoc = lvwMain.SelectedItem.SubItems(EnmCol.Enmҽ��)
    strNO = Trim(objLvw.SelectedItem.Text)
    str����ʱ�� = objLvw.SelectedItem.SubItems(EnmCol.Enm����ʱ��)
    Err = 0: On Error GoTo Errhand:
    If lngID = 0 Then Exit Sub
    
    If blnTriage Then
        If objLvw.SelectedItem.SubItems(EnmCol.Enm����) <> "" Then bln�ѷ��� = True
    End If
    
    '125454:���ϴ���2018/5/18���ֶ�ǩ������Ҫ��ʾ���Զ�ǩ����ʱ����ֱ��ȷ��ǩ����ֱ�ӷ���false������ʾ
    '95637:���ϴ�,2016/7/18,ǩ�����鵱ǰ�ű��Ƿ����Ŷ��У����ߵ����������ű����Ŷ���
    If Not blnȡ��ǩ�� Then
        If Checkǩ��(bln�Ƿ�ԤԼ, lng����ID, lngID, str����ʱ��, blnClick, bln����̨ǩ���Ŷ�) = False Then Exit Sub
    End If
    
    If Not bln�ѷ��� And Not blnȡ��ǩ�� Then
        'δ���ﲿ��
        zlExecuteTriage mfrmMain, True: Exit Sub
        objLvw.SelectedItem.ListSubItems(3).Tag = IIf(blnȡ��ǩ��, 0, 1)
        objLvw.SelectedItem.Icon = "rySign_in"
        objLvw.SelectedItem.SmallIcon = "rySign_in"
        Exit Sub
    End If
    
    strRoom = objLvw.SelectedItem.SubItems(EnmCol.Enm����)
    
    If ExcPlugInFun(IIf(blnȡ��ǩ��, 14, 4), lngID, strDoc, strRoom) = False Then Exit Sub
    
    
    'intType:---��������_IN:0-ǩ��;1-ȡ��ǩ��/ȡ��ҽ����ǻ���;2-ҽ����ǻ���/ȡ������̨����ǩ��;3-����̨����ǩ��
    If zlǩ����ȡ��(Not blnȡ��ǩ��, lngID, strNO) = False Then
        strTittle = IIf(blnȡ��ǩ��, "ȡ��ǩ��ʧ��!", "ǩ��ʧ��!")
        RaiseEvent zlShowInfor(strTittle)
        ShowMsgbox strTittle
        Exit Sub
    End If
    If blnȡ��ǩ�� = False Then '����:38165
        Call zlPrintBill(lngID)
        '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
        Call zlPrintBarcode
    End If
        
    strTittle = IIf(blnȡ��ǩ��, "ȡ��ǩ���ɹ�!", "ǩ���ɹ�!")
    RaiseEvent zlShowInfor(strTittle)
    ShowMsgbox strTittle
    objLvw.SelectedItem.ListSubItems(3).Tag = IIf(blnȡ��ǩ��, 0, 1)
    If blnȡ��ǩ�� Then
        If bln�ѷ��� Then
            objLvw.SelectedItem.Icon = "yf"
            objLvw.SelectedItem.SmallIcon = "yf"
        Else
            objLvw.SelectedItem.Icon = "ry"
            objLvw.SelectedItem.SmallIcon = "ry"
        End If
    Else
        objLvw.SelectedItem.Icon = "rySign_in"
        objLvw.SelectedItem.SmallIcon = "rySign_in"
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function zlǩ����ȡ��(blnǩ�� As Boolean, ByVal lng�Һ�ID As Long, ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:intǩ������ 0-����ǩ����1-����ǩ��;2-ת��ǩ��;3-����ǩ��;4-��������ǩ��;5-����ҵ������ǩ��
    '����:
    '����:
    '����:���˺�
    '����:2011-01-16 13:56:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String
    On Error GoTo errHandle
    If blnǩ�� Then
        'Zl_���˹Һż�¼_ǩ��
        '  Id_In     ���˹Һż�¼.ID%Type,
        '  ��ǩ��_In Integer:=0
        strSQL = "Zl_���˹Һż�¼_ǩ��(" & lng�Һ�ID & "," & 0 & ",'" & zl_GetԤԼ��ʽByID(lng�Һ�ID) & "')" '�����:48350
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����,7-����
        RaiseEvent zlQueueAsk(1, strNO, mlngPre����ID, blnCancel)
        If blnCancel Then Exit Function
        zlǩ����ȡ�� = True
        Exit Function
    End If
    'Zl_���˹Һż�¼_ȡ��ǩ��
    '  Id_In           ���˹Һż�¼.ID%Type,
    '  ����ǩ����־_In Integer:=0
    strSQL = "Zl_���˹Һż�¼_ȡ��ǩ��(" & lng�Һ�ID & "," & 0 & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    zlǩ����ȡ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Sub zlExcuteFunction()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ִ����ع���
    '���ƣ����˺�
    '���ڣ�2010-05-31 16:42:44
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnTriage As Boolean, lng����ID  As Long, lngExeState As Long
    Dim objLvw As ListView
        
    'ListSubItem(3).tag:A.��¼��־:0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
    'ListSubItem(4).tag:A.��¼��־:0-��ʾ�����÷���̨ǩ���Ŷ�,1-��ʾ���÷���̨ǩ���Ŷ�;
    
    If Val(zlDatabase.GetPara("��Һ�ģʽ", glngSys)) = 1 Then Exit Sub
    If Not (Me.ActiveControl Is lvwMain Or Me.ActiveControl Is LvwYY Or Me.ActiveControl Is lvwHZPati) Then Exit Sub
    Set objLvw = Me.ActiveControl
    If objLvw.SelectedItem Is Nothing Then
        blnTriage = False: lng����ID = 0
    Else
        lng����ID = Val(Split(objLvw.SelectedItem.Tag, "|")(0))
        '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬
        lngExeState = Val(Split(objLvw.SelectedItem.Tag, "|")(6))
        blnTriage = (lngExeState = 0) And lng����ID <> 0
    End If
    If blnTriage Then
        If Val(objLvw.SelectedItem.ListSubItems(4).Tag) = 1 Then
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            If Val(objLvw.SelectedItem.ListSubItems(3).Tag) = 0 Then
                GoTo EdPati:
            End If
        End If
        zlExecuteTriage mfrmMain: Exit Sub
    End If
EdPati:
    
    If lng����ID = 0 Then
         zlExcuteEditPati mfrmMain: Exit Sub
    End If
End Sub

Public Sub zlSubPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    On Error GoTo errHandle
    Dim objPrint As New zlPrintLvw
    If Me.ActiveControl Is lvwHZPati Then
        objPrint.Title.Text = "���ﲡ�����"
        Set objPrint.Body.objData = lvwHZPati
    
    ElseIf tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        objPrint.Title.Text = Me.Caption
        Set objPrint.Body.objData = lvwMain
    Else
        objPrint.Title.Text = "ԤԼ�������"
        Set objPrint.Body.objData = LvwYY
    End If
    objPrint.BelowAppItems.Add "��ӡ�ˣ�" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlExecuteTriage(ByVal frmMain As Object, _
    Optional blnǩ�� As Boolean = False, _
    Optional blnAppointment As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ִ�з���
    '���: blnAppointment-�Ƿ��ԤԼ���˽��з���
    '���ƣ����˺�
    '���ڣ�2010-05-31 14:54:48
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strRoom As String, strDate As String
    Dim strDoctor As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, blnCancel As Boolean, lngID As Long
    Dim cllPro As Collection
    Dim lng��¼���� As Long, blnԤԼ As Boolean, lng����ID As Long
    Dim strNO As String
     
    On Error GoTo errH
    
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        'ûѡ����˳�
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
    End If
    
    lng��¼���� = IIf(tbPage.Item(midx.idx_�ŶӶ���).Selected, 1, 2)
    
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        strNO = lvwMain.SelectedItem.Text
        strDoctor = lvwMain.SelectedItem.SubItems(EnmCol.Enmҽ��)
        lngID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
        lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
    Else
        strNO = LvwYY.SelectedItem.Text
        strDoctor = LvwYY.SelectedItem.SubItems(EnmCol.Enmҽ��)
        lngID = Val(LvwYY.SelectedItem.ListSubItems(1).Tag)
        lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        blnԤԼ = True
    End If
    
    ReadRoom
    strRoom = STR_COMP
    If frmDistRoom.cmb.ListCount > 0 Then
        frmDistRoom.cmb.ListIndex = 0
        For i = 0 To frmDistRoom.cmb.ListCount - 1
            If frmDistRoom.cmb.List(i) Like "*-" & Trim(strDoctor) Then
                frmDistRoom.cmb.ListIndex = i
                Exit For
            End If
        Next
    End If
    frmDistRoom.ShowMe strRoom, Me, blnǩ��, lngID
    If strRoom = STR_COMP Then
        RaiseEvent zlShowInfor("�û�ȡ��!") 'ѡ����"ȡ��"��
        Exit Sub
    End If
    'NO_IN       ���˹Һż�¼.NO%TYPE:=NULL,
    '����ID_IN   ���˹Һż�¼.����id%TYPE:=NULL,
    '����_IN     ���˹Һż�¼.����%TYPE:=NULL
    '
    Set cllPro = New Collection
    strDoctor = Trim(Split(strRoom, STR_COMP)(1))
    strDate = "To_date('" & Split(strRoom, STR_COMP)(2) & "','yyyy-mm-dd hh24:mi:ss')"
    strRoom = Split(strRoom, STR_COMP)(0)
    
    '111121:������2017/7/17,�ظ�����"Zl_���˹Һż�¼_ǩ��"
    '�����:48350
    strSQL = "ZL_���˹Һż�¼_�������� ('" & strNO & "'," & mlngPre����ID & ",'" & strRoom & "','" & strDoctor & "'," & strDate & ",'1','" & zl_GetԤԼ��ʽByNo(strNO) & "')"
    zlAddArray cllPro, strSQL
    
    If blnǩ�� Then
        'Zl_���˹Һż�¼_ǩ��
        '  Id_In     ���˹Һż�¼.ID%Type,
        '  ��ǩ��_In Integer:=0
        '�����:48350
        strSQL = "Zl_���˹Һż�¼_ǩ��(" & lngID & "," & 0 & ",'" & zl_GetԤԼ��ʽByID(lngID) & "')"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    If blnǩ�� = False Then
      '���ﴥ��������Ϣ
      Call SendMsgModule(strNO)
    End If
    
    RaiseEvent zlQueueAsk(1, strNO, mlngPre����ID, blnCancel)
    'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    Err = 0: On Error GoTo errH:
     '��ʾ�����к�
     strSQL = " Select A.�ŶӺ���,B.���� From �ŶӽкŶ��� A,���˹Һż�¼ B Where a.ҵ��ID= B.ID and  A.ҵ��id = [1] And A.ҵ������ = 0 and b.��¼����=[2] and b.��¼״̬=1  "
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID, lng��¼����)
     strDate = ""
     If Not rsTemp.EOF Then
        strDate = Nvl(rsTemp!����) & " ���ŶӺ���Ϊ:" & Nvl(rsTemp!�ŶӺ���)
     End If
     '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    Call zlPrintBarcode
    Call zlRefreshData
    If blnǩ�� Then  '����:38165
        Call zlPrintBill(lngID)
    End If
     RaiseEvent zlShowInfor(strDate) '��ʾ�ű�
     Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlExcuteChangeNum(ByVal frmMain As Form)
    '���˻���
    Dim blnCancel As Boolean, strNO As String, lngID As Long
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    strNO = Trim(lvwMain.SelectedItem.Text)
    lngID = Split(lvwMain.SelectedItem.Tag, "|")(4)
    
    If InStr(lvwMain.SelectedItem.Tag, "|") > 0 Then
        If gbytRegistMode = 0 Then
            If frmChangeNum.ShowMe(lngID, Me) Then
                RaiseEvent zlQueueAsk(2, strNO, mlngPre����ID, blnCancel)
                'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
                ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
                
                '���Ŵ���������Ϣ
                Call SendMsgModule(strNO)
                Call zlRefreshData
            End If
        Else
            If Sys.Currentdate < gdatRegistTime Then
                If frmChangeNum.ShowMe(lngID, Me) Then
                    RaiseEvent zlQueueAsk(2, strNO, mlngPre����ID, blnCancel)
                    'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
                    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
                    
                    '���Ŵ���������Ϣ
                    Call SendMsgModule(strNO)
                    Call zlRefreshData
                End If
            Else
                If frmChangeNumNew.ShowMe(lngID, Me) Then
                    RaiseEvent zlQueueAsk(2, strNO, mlngPre����ID, blnCancel)
                    'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
                    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
                    
                    '���Ŵ���������Ϣ
                    Call SendMsgModule(strNO)
                    Call zlRefreshData
                End If
            End If
        End If
    End If
End Sub
Public Sub zlExcuteEditPati(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��༭���˵�����Ϣ
    '���ƣ����˺�
    '���ڣ�2010-05-31 15:41:06
    '˵����
    '------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errH
    Dim lng����ID As Long, lng���� As Long, bln��Ժ As Boolean
    Dim i As Long
    Dim strNO As String
    Dim str��֤���� As String
    Dim str���￨�� As String
    
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        'ûѡ����˳�
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
    End If
    
   If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        lng����ID = CLng(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lng���� = CLng(Split(lvwMain.SelectedItem.Tag, "|")(1))
        strNO = lvwMain.SelectedItem.Text
        str���￨�� = Split(lvwMain.SelectedItem.Tag, "|")(2)
        str��֤���� = Split(lvwMain.SelectedItem.Tag, "|")(3)
    Else
        lng����ID = CLng(Split(LvwYY.SelectedItem.Tag, "|")(0))
        lng���� = CLng(Split(LvwYY.SelectedItem.Tag, "|")(1))
        strNO = LvwYY.SelectedItem.Text
        str���￨�� = Split(LvwYY.SelectedItem.Tag, "|")(2)
        str��֤���� = Split(LvwYY.SelectedItem.Tag, "|")(3)
    End If


    With frmDistRoomPatiEdit
        .mstrNo = strNO
        .mlng����ID = lng����ID
        .mlng���� = lng����
        .mstrPrivs = mstrPrivs
        .mlngModul = mlngModul
        '���ʽҽ��û������
        If lng���� = 0 And mlngOutModeMC > 0 Then
            .mlngOutModeMC = mlngOutModeMC
        Else
            .mlngOutModeMC = 0
        End If
        .m���￨�� = str���￨��
        .m��֤���� = str��֤����
        .InitData
        .Init����ҩ��
        '79912:���ϴ�,2014/11/20,��Ժ���˲����ڷ���̨�޸���Ϣ
        Call LoadPatientInfo(lng����ID, bln��Ժ)
        If bln��Ժ Then
            MsgBox "�ò�������Ժ,������������Ժ�����޸���Ϣ��", vbInformation, gstrSysName
            Exit Sub
        End If
        If lng����ID <= 0 Then
            If Not .GetRegBillID() Then
                MsgBox "�޷���ȡ�Һ�ID", vbInformation, gstrSysName
                 Unload frmDistRoomPatiEdit
                Exit Sub
            End If
        End If
        '67070:������,2013-11-04,��ȡ����������Ϣ
        .UCPatiVitalSigns.LoadPatiVitalSigns .mlng����ID, .mlng�Һ�ID
        Call .SetPatiBaseInforEnabled
        If mfrmMain Is Nothing Then
            .Show 1, Me
        Else
            .Show 1, mfrmMain
        End If
    End With
    
    '����ˢ��
    If gblnOk Then zlRefreshData (True)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlExcutePatiLeave(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����˲�����
    '���ƣ����˺�
    '���ڣ�2010-05-31 15:46:16
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Call Set���˹Һ�״̬(-1)
End Sub
Public Sub zlExcutePatiWait(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����˴���
    '���ƣ����˺�
    '���ڣ�2010-05-31 15:50:34
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
     Call Set���˹Һ�״̬(0)
End Sub

Public Sub zlExcutePatiCancelOver(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ȡ����ɾ���
    '���ƣ����˺�
    '���ڣ�2010-05-31 15:56:58
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strMsgbox As String, strSQL As String, lng����ID As Long, lngExeState As Long
    Dim blnCancel As Boolean, lngID As Long
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "��ɾ���") = 0 Then Exit Sub
    lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    If lngExeState <> 1 Then Exit Sub
    If lng����ID = 0 Then MsgBox "�����ڵĲ��ˣ�", vbInformation, gstrSysName: Exit Sub
    
    strMsgbox = "��������£�Ӧ����ҽ���ڱ�Ҫʱ�����Ƿ�ȡ��������ɣ�" & vbCrLf & _
                "�������л�ʿ�ڷ���ֱ̨�ӱ�ǵĲ��˽�����ɣ������ܽ��иò�����" & vbCrLf & vbCrLf & _
                "���Ҫȡ�������"
    If MsgBox(strMsgbox, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If ExcPlugInFun(13, Val(lvwMain.SelectedItem.ListSubItems(1).Tag)) = False Then Exit Sub


    Err = 0: On Error GoTo errHandle
    gcnOracle.BeginTrans
    strSQL = "zl_���˽������_Cancel(" & Split(lvwMain.SelectedItem.Tag, "|")(0) & ",'" & lvwMain.SelectedItem.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    'intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����,7-����
   RaiseEvent zlQueueAsk(6, Trim(lvwMain.SelectedItem.Text), mlngPre����ID, blnCancel)
    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    If blnCancel = True Then
        gcnOracle.RollbackTrans: Exit Sub
    End If
    gcnOracle.CommitTrans
    
    Exit Sub
errHandle:
     gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Set���˹Һ�״̬(ByVal lngState As Long)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ò��˹Һ�״̬
    '��Σ�lngState : -1- ���˲�����
    '                         0-���˴���
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-03 15:24:48
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnCancel As Boolean
    blnCancel = False

    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If ExcPlugInFun(IIf(lngState = -1, 3, 6), Val(lvwMain.SelectedItem.ListSubItems(1).Tag)) = False Then Exit Sub

    On Error GoTo errH
    gcnOracle.BeginTrans
    strSQL = "Zl_���˹Һż�¼_״̬ ('" & lvwMain.SelectedItem.Text & "'," & lngState & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    RaiseEvent zlQueueAsk(IIf(lngState = -1, 3, 4), Trim(lvwMain.SelectedItem.Text), mlngPre����ID, blnCancel)
    'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
    ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
    If blnCancel = True Then
        gcnOracle.RollbackTrans: Exit Sub
    End If
    
    gcnOracle.CommitTrans
    MsgBox "�����ɹ�!", vbInformation, gstrSysName

    Call zlRefreshData
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ReadRoom()
    On Error GoTo ErrHead
    Dim rsTmp As ADODB.Recordset
    Dim objListItem As ListItem
    Dim i As Long, lngSel As Long
    Dim strNO As String, strSQL As String
    Dim lng��¼���� As Long
    Dim strTmp As String
    Dim blnBusy As Boolean
    Dim lngת�����ID As Long
    'û��ѡ����˳�
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        'ûѡ����˳�
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        strNO = lvwMain.SelectedItem.Text
        
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
        strNO = LvwYY.SelectedItem.Text
    End If
    
    lng��¼���� = IIf(tbPage.Item(midx.idx_�ŶӶ���).Selected, 1, 2)
    
    frmDistRoom.lvwMain.ListItems.Clear
    frmDistRoom.cmb.Clear
   
    
    
    '������ҽ��
    '95637�����ϴ���2016/7/17��ת��ǩ��,��ת����һ�ȡҽ����Ϣ
    strSQL = _
        " Select c.���,c.����,Nvl(d.ת�����ID,0) as ת�����ID From ��Ա����˵�� a, ������Ա b ,��Ա�� c,���˹Һż�¼ d" & vbCrLf & _
        " Where b.��Աid=c.id And b.��Աid=a.��Աid And a.��Ա����='ҽ��' " & vbCrLf & _
        " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
        " And b.����id=Nvl(d.ת�����ID,d.ִ�в���ID) And d.��¼����=[2] and d.��¼״̬=1 and  d.NO=[1] And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng��¼����)
    frmDistRoom.cmb.AddItem "��"
    If rsTmp.RecordCount > 0 Then
        lngת�����ID = Val(Nvl(rsTmp!ת�����ID))
        For i = 1 To rsTmp.RecordCount
            frmDistRoom.cmb.AddItem zlCommFun.Nvl(rsTmp!���) & "-" & zlCommFun.Nvl(rsTmp!����)
            rsTmp.MoveNext
        Next
    End If

    '����Ϊ��ǰҽ��
    If gbytRegistMode = 0 Then
        strSQL = "Select A.ҽ������,B.�Ǽ�ʱ��,sysdate ��ǰʱ�� From �ҺŰ��� A,���˹Һż�¼ B Where A.����=B.�ű� And B.NO=[1] and b.��¼����=[2] and b.��¼״̬=1"
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select A.ҽ������,B.�Ǽ�ʱ��,sysdate ��ǰʱ�� From �ҺŰ��� A,���˹Һż�¼ B Where A.����=B.�ű� And B.NO=[1] and b.��¼����=[2] and b.��¼״̬=1"
        Else
            strSQL = "Select A.ҽ������,B.�Ǽ�ʱ��,sysdate ��ǰʱ�� From �ٴ������¼ A,���˹Һż�¼ B Where A.ID=B.�����¼ID And B.NO=[1] and b.��¼����=[2] and b.��¼״̬=1"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng��¼����)
    If rsTmp.RecordCount > 0 Then
        For i = 0 To frmDistRoom.cmb.ListCount - 1
            If frmDistRoom.cmb.List(i) Like "*-" & zlCommFun.Nvl(rsTmp!ҽ������) Then
                frmDistRoom.cmb.ListIndex = i
                Exit For
            End If
        Next
        frmDistRoom.dtpBegin.MaxDate = rsTmp!��ǰʱ��
        frmDistRoom.dtpBegin.MinDate = rsTmp!�Ǽ�ʱ��
        frmDistRoom.dtpBegin.Value = rsTmp!��ǰʱ��
    End If

    '�������еĸú�������ҹ�ѡ��
    '79694:���ϴ�,2014/11/25,���ݲ�����ȡ��������
    blnBusy = Val(zlDatabase.GetPara("����æʱ�������", glngSys, mlngModul, 0)) = 1
    If lngת�����ID <> 0 Then
        '95637:���ϴ�,2016/7/18,������ת�ֻ����ת�����ȥȷ������
        If gbytRegistMode = 0 Then
            strSQL = _
                " Select Distinct b.����, b.����, b.λ��" & vbNewLine & _
                " From �ҺŰ������� a, �������� b, �ҺŰ��� c" & vbNewLine & _
                " Where a.�������� = b.���� And a.�ű�id = c.Id And c.ID IN (Select ID From �ҺŰ��� Where ����ID=[3]) " & _
                IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & _
                " Order By B.���� "
        Else
            If Sys.Currentdate < gdatRegistTime Then
                strSQL = _
                    " Select Distinct b.����, b.����, b.λ��" & vbNewLine & _
                    " From �ҺŰ������� a, �������� b, �ҺŰ��� c" & vbNewLine & _
                    " Where a.�������� = b.���� And a.�ű�id = c.Id And c.ID IN (Select ID From �ҺŰ��� Where ����ID=[3]) " & _
                    IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & _
                    " Order By B.����"
            Else
                strSQL = _
                    " Select Distinct b.����, b.����, b.λ��" & vbNewLine & _
                    " From �����������ÿ��� a, �������� b" & vbNewLine & _
                    " Where a.����id = b.id And a.����ID=[3] " & _
                    IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & _
                    " Order By B.����"
            End If
        End If
    Else
        If gbytRegistMode = 0 Then
            strSQL = _
                " Select b.����, b.����, b.λ��" & vbNewLine & _
                " From �ҺŰ������� a, �������� b, �ҺŰ��� c, ���˹Һż�¼ d" & vbNewLine & _
                " Where a.�������� = b.���� And a.�ű�id = c.Id And c.���� = d.�ű� And d.No = [1] " & _
                IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & " and d.��¼����=[2] and d.��¼״̬=1" & _
                    " Order By B.����"
        Else
            If Sys.Currentdate < gdatRegistTime Then
                strSQL = _
                    " Select b.����, b.����, b.λ��" & vbNewLine & _
                    " From �ҺŰ������� a, �������� b, �ҺŰ��� c, ���˹Һż�¼ d" & vbNewLine & _
                    " Where a.�������� = b.���� And a.�ű�id = c.Id And c.���� = d.�ű� And d.No = [1] " & _
                    IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & " and d.��¼����=[2] and d.��¼״̬=1" & _
                    " Order By B.����"
            Else
                strSQL = "Select ���﷽ʽ From �ٴ������¼ A,���˹Һż�¼ B Where B.NO = [1] And B.��¼���� = [2] And A.ID = B.�����¼ID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng��¼����)
                If rsTmp.EOF Then
                    strSQL = _
                        " Select b.����, b.����, b.λ��" & vbNewLine & _
                        " From �����������ÿ��� a, �������� b, ���˹Һż�¼ d" & vbNewLine & _
                        " Where a.����id = b.id And a.����id = d.ִ�в���ID And d.No = [1] " & _
                        IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & " and d.��¼����=[2] and d.��¼״̬=1"
                Else
                    If Val(Nvl(rsTmp!���﷽ʽ)) = 0 Then
                        strSQL = _
                            " Select b.����, b.����, b.λ��" & vbNewLine & _
                            " From �����������ÿ��� a, �������� b, ���˹Һż�¼ d" & vbNewLine & _
                            " Where a.����id = b.id And a.����id = d.ִ�в���ID And d.No = [1] " & _
                            IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & " and d.��¼����=[2] and d.��¼״̬=1"
                    Else
                        strSQL = _
                            " Select b.����, b.����, b.λ��" & vbNewLine & _
                            " From �ٴ��������Ҽ�¼ a, �������� b, ���˹Һż�¼ d" & vbNewLine & _
                            " Where a.����id = b.id And a.��¼id = d.�����¼id And d.No = [1] " & _
                            IIf(blnBusy, " ", " And b.ȱʡ��־=0 ") & " and d.��¼����=[2] and d.��¼״̬=1"
                    End If
                End If
            End If
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng��¼����, lngת�����ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set objListItem = frmDistRoom.lvwMain.ListItems.Add(, , zlCommFun.Nvl(rsTmp!����))
            objListItem.SubItems(1) = zlCommFun.Nvl(rsTmp!λ��)
            
            If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
                strTmp = Me.lvwMain.SelectedItem.SubItems(EnmCol.Enm����)
            Else
                strTmp = Me.LvwYY.SelectedItem.SubItems(EnmCol.Enm����)
            End If
            If rsTmp!���� = strTmp Then
                objListItem.Selected = True
                objListItem.EnsureVisible
                lngSel = i
            End If
            rsTmp.MoveNext
        Next
        If lngSel = 0 Then
            frmDistRoom.lvwMain.ListItems(1).Selected = True
            frmDistRoom.lvwMain.ListItems(1).EnsureVisible
        End If
    End If

    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatientInfo(ByVal lng����ID As Long, Optional ByRef bln��Ժ As Boolean = False)
    On Error GoTo errH
    '����:��ȡ������Ϣ
    Dim str���� As String, strSQL As String
    Dim i As Integer
    Dim lngTmp As Long
    Dim rsTmp As ADODB.Recordset
    Dim strNO As String
    Dim strName As String
    Dim strSex  As String, strAge As String
    Dim lng��¼���� As String
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        strNO = lvwMain.SelectedItem.Text
        strName = lvwMain.SelectedItem.SubItems(EnmCol.Enm����)
        strSex = lvwMain.SelectedItem.SubItems(EnmCol.Enm�Ա�)
        strAge = lvwMain.SelectedItem.SubItems(EnmCol.Enm����)
    Else
        strNO = LvwYY.SelectedItem.Text
        strName = LvwYY.SelectedItem.SubItems(EnmCol.Enm����)
        strSex = LvwYY.SelectedItem.SubItems(EnmCol.Enm�Ա�)
        strAge = LvwYY.SelectedItem.SubItems(EnmCol.Enm����)
    End If
    
    lng��¼���� = IIf(tbPage.Item(midx.idx_�ŶӶ���).Selected, 1, 2)
    
    With frmDistRoomPatiEdit
        .txtPatient.MaxLength = GetColumnLength("������Ϣ", "����")
        .txt����.MaxLength = GetColumnLength("������Ϣ", "����")
        .txt�����.MaxLength = GetColumnLength("������Ϣ", "�����")
        .padd��ͥ��ַ.MaxLength = GetColumnLength("������Ϣ", "��ͥ��ַ")
        .padd���ڵ�ַ.MaxLength = GetColumnLength("������Ϣ", "���ڵ�ַ")
        .mstr���� = "": .mstr�Ա� = "": .mstr���� = "": .mstr�������� = ""
        .mblnҽ��ҵ�� = False
        If lng����ID <= 0 Then
            .mbytType = 1  '����һ���µĲ�����Ϣ
            .txt�����.Text = zlDatabase.GetNextNo(3)
            .txtPatient.Text = strName
            For i = 0 To .cbo�Ա�.ListCount - 1
                If .cbo�Ա�.List(i) Like "*" & Trim(strSex) Then
                    .cbo�Ա�.ListIndex = i
                    Exit For
                End If
            Next
            Call LoadOldData(strAge, .txt����, .cbo���䵥λ)
            Exit Sub
        End If
        '79912:���ϴ�,2014/11/20,��Ժ���˲����ڷ���̨�޸���Ϣ
        strSQL = _
            "Select A.*,D.���� as �Һ�����,D.�Ա� as �Һ��Ա�,D.���� as �Һ�����,Decode(B.����ID,NULL,0,1) As ����,C.ҽ�����,To_Char(C.����ʱ��,'YYYY-MM-DD HH24:MI:SS') As  ����ʱ��,D.ID As �Һ�ID " & vbCrLf & _
            " From ������Ϣ A,���ﲡ����¼ B,����ǼǼ�¼ C,���˹Һż�¼ D" & vbCrLf & _
            " Where A.����ID=B.����ID(+) And A.����ID=[1] " & vbCrLf & _
            "��And D.NO=[2] And D.�Ǽ�ʱ��=C.����ʱ��(+) And D.����ID=C.����ID(+) and d.��¼����=[3] and d.��¼״̬=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strNO, lng��¼����)
        
        If rsTmp.RecordCount > 0 Then
            If Val(Nvl(rsTmp!��ǰ����id)) <> 0 Then bln��Ժ = True: Exit Sub
            .ClearFace
            '��������Ϣ�Ĳ���,ҲȱʡΪ���޸���Ϣ,��������
            '�����в����Ĳ���,ֻ��Ϊ�޸Ĳ�����Ϣ
            If rsTmp!���� = 1 Then
                .mbytType = 3 'ֻ���²�����Ϣ
            Else
                .mbytType = 2 '�޲���,����һ���µĲ�����Ϣ
            End If
            .mstr�������� = Format(rsTmp!��������, "YYYY-MM-DD")
            If lng����ID = 0 Then
                .mblnҽ��ҵ�� = False
            Else
                 .mblnҽ��ҵ�� = zlExistOperationData(lng����ID, strNO, Val(Nvl(rsTmp!�Һ�ID)))
            End If
            .mstr�������� = Format(rsTmp!��������, "YYYY-MM-DD")
            .mstr����_���� = Nvl(rsTmp!����)
            .mstr����_�Ա� = Nvl(rsTmp!�Ա�)
            .mstr����_���� = Nvl(rsTmp!����)
            .mstr���� = Nvl(rsTmp!�Һ�����)
            .mstr�Ա� = Nvl(rsTmp!�Һ��Ա�)
            .mstr���� = Nvl(rsTmp!�Һ�����)
            
            If IsNull(rsTmp!�����) Then
                .txt�����.Text = zlDatabase.GetNextNo(3)     '�в�����һ���������
            Else
                .txt�����.Text = rsTmp!�����
            End If
            If .mblnҽ��ҵ�� And InStr(.mstrPrivsPubPatient, ";������Ϣ����;") = 0 Then
                .txtPatient.Text = zlCommFun.Nvl(rsTmp!�Һ�����)
                .cbo�Ա�.ListIndex = cbo.FindIndex(.cbo�Ա�, zlCommFun.Nvl(rsTmp!�Һ��Ա�), True)
                Call LoadOldData("" & rsTmp!�Һ�����, .txt����, .cbo���䵥λ)
            Else
                .txtPatient.Text = zlCommFun.Nvl(rsTmp!����)
                .cbo�Ա�.ListIndex = cbo.FindIndex(.cbo�Ա�, zlCommFun.Nvl(rsTmp!�Ա�), True)
                Call LoadOldData("" & rsTmp!����, .txt����, .cbo���䵥λ)
            End If
            '74428�����ϴ���2014-7-8������������ʾ��ɫ����
            Call SetPatiColor(.txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), .ForeColor, vbRed))
            .mblnChange = False
            .txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "____-__-__", rsTmp!��������), "YYYY-MM-DD")
            .mblnChange = True

            If Not IsNull(rsTmp!��������) Then
                If .mblnҽ��ҵ�� = False Then
                    .txt����.Text = ReCalcOld(CDate(.txt��������.Text), .cbo���䵥λ, lng����ID) '���ݳ���������������
                    If CDate(.txt��������.Text) - CDate(rsTmp!��������) <> 0 Then .txt����ʱ��.Text = Format(rsTmp!��������, "HH:MM")
                End If
            Else
                .txt����ʱ��.Text = "__:__"
                .mblnChange = False
                  If .mblnҽ��ҵ�� = False Then
                    .txt��������.Text = ReCalcBirth(.txt����.Text, .cbo���䵥λ.Text)
                End If
                .mblnChange = True
            End If


            If .mlngOutModeMC > 0 Then
                If .mlngOutModeMC = 920 Then
                    .txtPatiMCNO(0).MaxLength = 12
                Else
                    .txtPatiMCNO(0).MaxLength = 30
                End If
                .txtPatiMCNO(0).ToolTipText = "��󳤶�" & .txtPatiMCNO(0).MaxLength & "λ"
                .txtPatiMCNO(1).MaxLength = .txtPatiMCNO(0).MaxLength

                .txtPatiMCNO(0).Text = "" & rsTmp!ҽ����    '�Զ��ضϳ�����󳤶ȵ��ַ�
                .txtPatiMCNO(0).Tag = .txtPatiMCNO(0).Text
                .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text

                If Not IsNull(rsTmp!ҽ�����) Then
                    For i = 0 To .cboҽ�����.ListCount - 1
                        lngTmp = InStr(1, .cboҽ�����.List(i), "-")
                        If lngTmp > 1 Then
                            If Mid(.cboҽ�����.List(i), 1, lngTmp - 1) = rsTmp!ҽ����� Then
                                .cboҽ�����.ListIndex = i: Exit For
                            End If
                        End If
                    Next
                    .cboҽ�����.Tag = "" & rsTmp!����ʱ��
                End If
            ElseIf .mlng���� > 0 Then
                 
                .mstrҽ���� = "" & rsTmp!ҽ����
                 
            End If
            

            .cbo�ѱ�.ListIndex = cbo.FindIndex(.cbo�ѱ�, zlCommFun.Nvl(rsTmp!�ѱ�), True)
            .cbo���ʽ.ListIndex = cbo.FindIndex(.cbo���ʽ, zlCommFun.Nvl(rsTmp!ҽ�Ƹ��ʽ), True)
            .cbo����.ListIndex = cbo.FindIndex(.cbo����, zlCommFun.Nvl(rsTmp!����), True)
            .cbo����.ListIndex = cbo.FindIndex(.cbo����, zlCommFun.Nvl(rsTmp!����), True)
            .cbo����.ListIndex = cbo.FindIndex(.cbo����, zlCommFun.Nvl(rsTmp!����״��), True)
            .cboְҵ.ListIndex = cbo.FindIndex(.cboְҵ, zlCommFun.Nvl(rsTmp!ְҵ), True)
            .txt���֤��.Text = zlCommFun.Nvl(rsTmp!���֤��)
            .txt��λ����.Text = zlCommFun.Nvl(rsTmp!������λ)
            .txt��λ����.Tag = zlCommFun.Nvl(rsTmp!��ͬ��λID)
            .txt��λ�绰.Text = zlCommFun.Nvl(rsTmp!��λ�绰)
            .txt��λ�ʱ�.Text = zlCommFun.Nvl(rsTmp!��λ�ʱ�)
            .cbo��ͥ��ַ.Text = zlCommFun.Nvl(rsTmp!��ͥ��ַ)
            '89242:���ϴ�,2015/12/7,ʹ�ýṹ����ַ
            Call zlReadAddrInfo(.padd��ͥ��ַ, Val(Nvl(rsTmp!����ID)), Val(Nvl(rsTmp!��ҳID)), 3, .cbo��ͥ��ַ.Text)
            If .padd��ͥ��ַ.Value = "" Then Call zlLoadDefaultAddr(.padd��ͥ��ַ)
            .txt��ͥ�绰.Text = zlCommFun.Nvl(rsTmp!��ͥ�绰)
            .txt��ͥ�ʱ�.Text = zlCommFun.Nvl(rsTmp!��ͥ��ַ�ʱ�)
            .txt���ڵ�ַ.Text = zlCommFun.Nvl(rsTmp!���ڵ�ַ)
            Call zlReadAddrInfo(.padd���ڵ�ַ, Val(Nvl(rsTmp!����ID)), Val(Nvl(rsTmp!��ҳID)), 4, .txt���ڵ�ַ.Text)
            If .padd���ڵ�ַ.Value = "" Then Call zlLoadDefaultAddr(.padd���ڵ�ַ)
            .txt�����ʱ�.Text = zlCommFun.Nvl(rsTmp!���ڵ�ַ�ʱ�)
            .txtEdit(0).Text = zlCommFun.Nvl(rsTmp!�໤��)
            .mlng�Һ�ID = Val(Nvl(rsTmp!�Һ�ID))
            '����ҩ��
            str���� = Get����ҩ��(rsTmp!����ID)
            If str���� <> "" Then
                If UBound(Split(str����, ";")) + 1 > .msh����.Rows - 1 Then .msh����.Rows = UBound(Split(str����, ";")) + 2
                For i = 0 To UBound(Split(str����, ";"))
                    .msh����.RowData(i + 1) = Val(Split(Split(str����, ";")(i), "|")(0))
                    .msh����.TextMatrix(i + 1, 0) = Split(Split(str����, ";")(i), "|")(1)
                    .msh����.TextMatrix(i + 1, 1) = Split(Split(str����, ";")(i), "|")(2)
                Next
            End If
        Else
            .mbytType = 1 '����һ���µĲ�����Ϣ
            .txt�����.Text = zlDatabase.GetNextNo(3)

            .txtPatient.Text = strName
            For i = 0 To .cbo�Ա�.ListCount - 1
                If .cbo�Ա�.List(i) Like "*" & Trim(strSex) Then
                    .cbo�Ա�.ListIndex = i
                    Exit For
                End If
            Next
            Call LoadOldData(strAge, .txt����, .cbo���䵥λ)
        End If
        strSQL = "Select ����ID,��Ϣ��,��Ϣֵ From ������Ϣ�ӱ� Where ����ID=[1] And ����ID=[2] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .mlng����ID, .mlng�Һ�ID)
        If rsTmp.EOF Then Exit Sub
        'idx_�໤�� = 0
        'idx_��� = 1
        'idx_���� = 2
        'idx_���� = 3
        rsTmp.Filter = "��Ϣ��='���'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(1).Text = Nvl(rsTmp!��Ϣֵ)
        End If
        rsTmp.Filter = "��Ϣ��='����'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(2).Text = Nvl(rsTmp!��Ϣֵ)
        End If
        rsTmp.Filter = "��Ϣ��='����'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(3).Text = Nvl(rsTmp!��Ϣֵ)
        End If
        
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub InitVariateAndPara()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ������
    '���ƣ����˺�
    '���ڣ�2010-05-31 16:58:25
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim arrTmp As Variant, i As Long
    '��������¼����ID���г�ʼ��
    mlngPre����ID = 0: mlngOutModeMC = 0
    arrTmp = Split(GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", ""), ",")
    For i = 0 To UBound(arrTmp)
        If IsNumeric(arrTmp(i)) Then
            If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For
        End If
    Next
    mlngDefaultCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, 0))
    
    frmDistRoomPatiEdit.mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '���˵�ַ�ṹ��¼��
    frmDistRoomPatiEdit.mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '�����ַ�ṹ��¼��
End Sub
Private Sub Form_Load()
    Dim strText As String, i As Integer
    On Error GoTo errHandle
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call InitPancel
    Call InitPage
    lvwMain.View = lvwReport
    lvwMain.ColumnHeaders.Clear
    'ҽ����,ժҪ:21101
    '74898:���ϴ�,2015/4/9,��ǲ��˵ĺ���״̬
    zlControl.LvwSelectColumns lvwMain, "�Һŵ�,1100,0,2;����״̬,900,0,1;����,1200,0,1;����,600,0,1;�Һ���Ŀ,1350,0,1;�����,800,0,1;����,960,0,1;�Ա�,400,0,1;����,600,0,1;����,1600,0,1;ҽ��,960,0,1;����ʱ��,2000,0,1;�Һ�ʱ��,2000,0,1;���,600,0,1;ҽ����,600,0,1;ժҪ,2000,0,1;ԤԼ,800,0,1;����,400,0,1;������,960,0,1;��������,1600,0,1;����ʱ��,2000,0,1", True
    zlControl.LvwSelectColumns LvwYY, "ԤԼ��,1100,0,2;����״̬,900,0,1;����,1200,0,1;����,600,0,1;�Һ���Ŀ,1350,0,1;�����,800,0,1;����,960,0,1;�Ա�,400,0,1;����,600,0,1;����,1600,0,1;ҽ��,960,0,1;ԤԼʱ��,2000,0,1;�Һ�ʱ��,2000,0,1;���,600,0,1;ҽ����,600,0,1;ժҪ,2000,0,1", True
    zlControl.LvwSelectColumns lvwHZPati, "�Һŵ�,1100,0,2;����״̬,900,0,1;����,1200,0,1;����,600,0,1;�Һ���Ŀ,1350,0,1;�����,800,0,1;����,960,0,1;�Ա�,400,0,1;����,600,0,1;����,1600,0,1;ҽ��,960,0,1;����ʱ��,2000,0,1;�Һ�ʱ��,2000,0,1;���,600,0,1;ҽ����,600,0,1;ժҪ,2000,0,1;ԤԼ,800,0,1", True
    Call InitVariateAndPara
'    '����ʱˢ��һ��
'    Call zlRefreshData     ����108110,��ε���ˢ�·����б�
    strText = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\zl9RegEvent\" & Me.Name & "\ListView", lvwMain.Name & "����")
    For i = 1 To lvwMain.ColumnHeaders.Count
        '���������У��򲻻ָ����Ի�
        If InStr(strText, lvwMain.ColumnHeaders(i).Text) = 0 Then lvwMain.Tag = "": Exit For
        '����������У�Ҳ���ָ����Ի�
        strText = Replace(strText, lvwMain.ColumnHeaders(i).Text, "")
    Next
    strText = Replace(strText, ",", "")
    If strText <> "" Then lvwMain.Tag = ""
    Call RestoreWinState(Me, App.ProductName)
    lvwMain.Tag = "�ɱ仯��"
    
    If CreatePlugInOK(glngModul) Then
        gblnPlugin = True
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function CheckMCOutMode(ByVal strMCCode As String) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String

    strSQL = "Select 1 From ������� Where ���=1 And ���=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlSaveDockPanceToReg Me, dkpMan, "����"
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.index - 1
    End If
End Sub

Private Sub lvwYY_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvwYY.Sorted = True
    If LvwYY.SortKey = ColumnHeader.index - 1 Then
        If LvwYY.SortOrder = lvwAscending Then
            LvwYY.SortOrder = lvwDescending
        Else
            LvwYY.SortOrder = lvwAscending
        End If
    Else
        LvwYY.SortKey = ColumnHeader.index - 1
    End If
End Sub
Private Sub lvwMain_DblClick()
    If Not lvwMain.SelectedItem Is Nothing Then
    
        Call zlExcuteFunction
    End If
End Sub

Private Sub lvwYY_DblClick()
    If Not LvwYY.SelectedItem Is Nothing Then
         Call zlExcuteFunction
    End If
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng����ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem


    '�д����˳�
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    lvwMain.Tag = Item.Text

    '�����Ƿ��Ѿ���������(���ڲ���id)��ִ��״̬�������Ƿ�ɷ�����š�������������ɽ����ϵ�в���
    lng����ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre����ID = lng����ID

    RaiseEvent zlShowInfor("���ݺ�:" & Item.Text & _
        "  ����:" & Item.SubItems(EnmCol.Enm����) & _
        "  ����:" & IIf(Item.SubItems(EnmCol.Enm����) = "", "δ����", Item.SubItems(EnmCol.Enm����)) & _
        "  ҽ��:" & IIf(Item.SubItems(EnmCol.Enmҽ��) = "", "δָ��", Item.SubItems(EnmCol.Enmҽ��)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwYY_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng����ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem
    '�д����˳�
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    LvwYY.Tag = Item.Text

    '�����Ƿ��Ѿ���������(���ڲ���id)��ִ��״̬�������Ƿ�ɷ�����š�������������ɽ����ϵ�в���
    lng����ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre����ID = lng����ID

    RaiseEvent zlShowInfor("ԤԼ����:" & Item.Text & _
        "  ����:" & Item.SubItems(EnmCol.Enm����) & _
        "  ����:" & IIf(Item.SubItems(EnmCol.Enm����) = "", "δ����", Item.SubItems(EnmCol.Enm����)) & _
        "  ҽ��:" & IIf(Item.SubItems(EnmCol.Enmҽ��) = "", "δָ��", Item.SubItems(EnmCol.Enmҽ��)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bln����̨ǩ���Ŷ� As Boolean
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    bln����̨ǩ���Ŷ� = Val(lvwMain.SelectedItem.ListSubItems(4).Tag) = 1
    If Button = 1 And IsNumeric(Trim(lvwMain.SelectedItem.SubItems(EnmCol.Enm�����))) Then
        If bln����̨ǩ���Ŷ� Then
            '    'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            If Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 0 And Val(Split(lvwMain.SelectedItem.Tag, "|")(6)) = 0 Then
                Exit Sub
            End If
        End If
        Set lvwMain.DragIcon = lvwMain.SelectedItem.CreateDragImage
        lvwMain.Drag 1
    End If
End Sub
Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        Call ReadRoom
        RaiseEvent zlPopuMenu(Button, Shift, X, Y)
    End If
End Sub

Private Sub lvwRoom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRoom.Sorted = True
    If lvwRoom.SortKey = ColumnHeader.index - 1 Then
        If lvwRoom.SortOrder = lvwAscending Then
            lvwRoom.SortOrder = lvwDescending
        Else
            lvwRoom.SortOrder = lvwAscending
        End If
    Else
        lvwRoom.SortKey = ColumnHeader.index - 1
    End If
End Sub

Private Sub lvwRoom_DragDrop(Source As Control, X As Single, Y As Single)
    On Error GoTo errH
    Dim strRoom As String, strDate As String
    Dim strDoctor As String, strSQL As String
    Dim i As Long
    Dim blnCancel As Boolean
    If Source Is lvwMain And Not lvwRoom.DropHighlight Is Nothing Then
        Set lvwRoom.SelectedItem = lvwRoom.DropHighlight
        Set lvwRoom.DropHighlight = Nothing
        ReadRoom
        strRoom = STR_COMP
        If frmDistRoom.cmb.ListCount > 0 Then
            frmDistRoom.cmb.ListIndex = 0
            For i = 0 To frmDistRoom.cmb.ListCount - 1
                If frmDistRoom.cmb.List(i) Like "*-" & lvwMain.SelectedItem.SubItems(EnmCol.Enmҽ��) Then
                    frmDistRoom.cmb.ListIndex = i
                    Exit For
                End If
            Next
        End If
        If frmDistRoom.lvwMain.ListItems.Count > 0 Then
            For i = 1 To frmDistRoom.lvwMain.ListItems.Count
                If frmDistRoom.lvwMain.ListItems(i).Text = lvwRoom.SelectedItem.Text Then
                    frmDistRoom.lvwMain.ListItems(i).Selected = True
                    frmDistRoom.lvwMain.ListItems(i).EnsureVisible
                    Exit For
                End If
            Next
        End If
        frmDistRoom.ShowMe strRoom, Me
        If strRoom = STR_COMP Then
            RaiseEvent zlShowInfor("�û�ȡ����"): Exit Sub   'ѡ����"ȡ��"��
        End If
        'NO_IN       ���˹Һż�¼.NO%TYPE:=NULL,
        '����ID_IN   ���˹Һż�¼.����id%TYPE:=NULL,
        '����_IN     ���˹Һż�¼.����%TYPE:=NULL
        '
        strDoctor = Trim(Split(strRoom, STR_COMP)(1))
        strDate = "To_date('" & Split(strRoom, STR_COMP)(2) & "','yyyy-mm-dd hh24:mi:ss')"
        strRoom = Split(strRoom, STR_COMP)(0)
        '�����:48350
        strSQL = "ZL_���˹Һż�¼_�������� ('" & lvwMain.SelectedItem.Text & "'," & mlngPre����ID & ",'" & strRoom & "','" & strDoctor & "'," & strDate & ",'','" & zl_GetԤԼ��ʽByNo(lvwMain.SelectedItem.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        RaiseEvent zlQueueAsk(2, Trim(lvwMain.SelectedItem.Text), mlngPre����ID, blnCancel)
        'intType:intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-�ָ�����
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        
        '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
        Call zlPrintBarcode

        Call zlRefreshData
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwRoom_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    On Error GoTo errH
    Dim objOver As ListItem

    'ûѡ����˳�
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Source Is lvwMain Then
        Set objOver = lvwRoom.HitTest(X, Y)
        If Not objOver Is Nothing Then
            If objOver.ForeColor <> RGB(255, 0, 0) And Trim(lvwMain.SelectedItem.SubItems(EnmCol.Enm�����)) <> "" And objOver.SubItems(5) Like "*" & lvwMain.SelectedItem.SubItems(EnmCol.Enm����) & "*" Then
                Set lvwRoom.DropHighlight = objOver
            Else
                Set lvwRoom.DropHighlight = Nothing
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlRefreshData(Optional blnFilter As Boolean = False, _
    Optional strFindValue As String = "", Optional bytReadType As Byte = 0, Optional objCard As Card, Optional ByVal blnAutoǩ�� As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ˢ������
    '��Σ�blnFilter-�Ƿ����
    '          bytReadType-��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
    '���ƣ����˺�
    '���ڣ�2010-06-02 09:43:08
    '------------------------------------------------------------------------------------------------------------------------
    Call ShowBills(blnFilter, strFindValue, bytReadType, objCard, blnAutoǩ��)
End Sub

Private Sub ShowBills(blnFilter As Boolean, Optional strFindValue As String = "", _
    Optional bytReadType As Byte = 0, Optional objCard As Card, Optional ByVal blnAutoǩ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '��Σ�blnFilter-�Ƿ����
    '          bytReadType-��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
    '     blbAppointment-�Ƿ�ˢ��ԤԼ����
    '����:���˺�
    '����:2011-11-21 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String, bytType As Byte
    mstrRegistIdsed = ""
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    
    If GetFilterCons(strFindValue, objCard, bytReadType, strValue, bytType) = False Then Screen.MousePointer = 0: Exit Sub
        
    '����ԤԼ����
    Call ShowBillsAppointment(blnFilter, strValue, bytType, objCard, blnAutoǩ��)
    
    '���عҺ�����
    Call ShowBillRegister(blnFilter, strValue, bytType, objCard, blnAutoǩ��)
    
    '���ػ�������
    Call ShowBillRegisterHZ(blnFilter, strValue, bytType, objCard)
    
    '��������
    Call LoadRooms
    
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlSetViewScrop(ByVal index As Integer, ByVal bytValue As Byte, Optional blnRefrashData As Boolean = False)
    '������ص���ʾ����
    'mbytViewScrop(0 To 3) As Byte  '0-��ʾ�ѷ��ﲡ��;1-��ʾ�ѽ��ﲡ��;2-��ʾ����ɲ���;3-��ʾ�����ﲡ��
    mbytViewScrop(index) = bytValue
     
    If blnRefrashData Then Call zlRefreshData
End Sub

'���ù�������
Public Sub zlSetFilterCons(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '����:���˺�
    '����:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcllFilter = ArrFilter
End Sub

Public Sub zlSetobjMsgModule(ByVal objMsgModule As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '����:���˺�
    '����:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjMsgModule = objMsgModule
End Sub
 

'------------------------------------------------------------------------------------------------------
'�����������:
'    �������;��Ч����
Public Property Get zl�������() As String
    zl������� = mstr�������
End Property

Public Property Let zl�������(ByVal vNewValue As String)
    mstr������� = vNewValue
End Property

Public Property Get zl��Ч����() As Integer
    zl��Ч���� = mint��Ч����
End Property

Public Property Let zl��Ч����(ByVal vNewValue As Integer)
    mint��Ч���� = vNewValue
End Property

Public Property Get zlintFindKeys() As Integer
    zlintFindKeys = mintFindKeys
End Property
Public Property Let zlintFindKeys(ByVal vNewValue As Integer)
    mintFindKeys = vNewValue
End Property

Public Property Get zlIsHaveData() As Boolean
    If Me.ActiveControl Is Me.lvwHZPati Then
        zlIsHaveData = lvwHZPati.ListItems.Count <> 0
    ElseIf tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        zlIsHaveData = lvwMain.ListItems.Count <> 0
    Else
        zlIsHaveData = LvwYY.ListItems.Count <> 0
    End If
End Property
Public Property Get zlIsTriage() As Boolean
    Dim lng����ID As Long, lngExeState As Long
    
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        '�Ƿ��ܷ���
        If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIsTriage = False
        Else
            lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            'ListSubItems(4).Tag:0-�����÷���̨ǩ���Ŷ�;1-���÷���̨ǩ���Ŷ�
            If Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1 Then
                '�Ѿ�ǩ��
                 zlIsTriage = (lngExeState = 0)
            Else 'δǩ��
                 zlIsTriage = (lngExeState = 0) And Not Val(lvwMain.SelectedItem.ListSubItems(4).Tag) = 1
            End If
        End If
    Else
        '�Ƿ��ܷ���
        If LvwYY.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIsTriage = False
        Else
            lng����ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            'ListSubItems(4).Tag:0-�����÷���̨ǩ���Ŷ�;1-���÷���̨ǩ���Ŷ�
            If Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 1 Then
                '�Ѿ�ǩ��
                 zlIsTriage = (lngExeState = 0)
            Else 'δǩ��
                 zlIsTriage = (lngExeState = 0) And Not Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            End If
        End If
    End If
End Property
Public Property Get zlIsPatiLeave() As Boolean
    '������������
    Dim lng����ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
        lngExeState = 0
    Else
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    End If
    zlIsPatiLeave = (lngExeState = 0)
End Property

Public Property Get zlIsPatiWait() As Boolean
    '�����Ƿ��������
    Dim lng����ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
        lngExeState = 0
    Else
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    End If
     zlIsPatiWait = (lngExeState = -1)
End Property
Public Property Get zlIsPatiFinish() As Boolean
    '�����Ƿ�������ɾ���
    Dim lng����ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Or tbPage.Item(midx.idx_ԤԼ����).Selected Then
        zlIsPatiFinish = False
    Else
        lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
         zlIsPatiFinish = (lngExeState = 0 Or lngExeState = 2)
    End If
End Property

Public Property Get zlIsPatiReDo() As Boolean
    '�Ƿ������˻ָ�����
    Dim lng����ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Or tbPage.Item(midx.idx_ԤԼ����).Selected Then
        zlIsPatiReDo = False
    Else
        lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
         zlIsPatiReDo = (lngExeState = 1)
    End If
End Property

Public Property Get zlIs����ǩ��(Optional ByRef bytQueue As Byte) As Boolean
    '�Ƿ�������ǩ��
    'mbln����̨ǩ���Ŷ�: 0-�Һ������Ŷ�,1-����̨ǩ���Ŷ�
    'bytQueue0-����ǩ����1-����ǩ��
    Dim lng����ID As Long, lngExeState As Long, lngTrunState As Long
    If Me.ActiveControl Is Me.lvwHZPati Then
        '���ﲡ��
        zlIs����ǩ�� = False
    '63789,������,2014-01-09,����ԤԼ����ǩ��
    ElseIf tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        If lvwMain.SelectedItem Is Nothing Then
            zlIs����ǩ�� = False
        Else
            lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            lngTrunState = Val(Split(lvwMain.SelectedItem.Tag, "|")(8))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            
            '95637:���ϴ�,2016/7/17,֧�ַ���̨ǩ���Ŷ�ģʽ�Ļ��ţ�ת��ǩ���Լ�����ǩ��
            zlIs����ǩ�� = (lngExeState = 0 Or lngExeState = 2 And lngTrunState = 0) And (Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 0 Or Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1)
            bytQueue = Val(lvwMain.SelectedItem.ListSubItems(3).Tag)
            '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬ & "|" & !ǩ������ & "|" & !ת��״̬
        End If
    ElseIf tbPage.Item(midx.idx_ԤԼ����).Selected Then
        If LvwYY.SelectedItem Is Nothing Then
            zlIs����ǩ�� = False
        Else
            lng����ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            'ListSubItems(4).Tag:0-�����÷���̨ǩ���Ŷ�;1-���÷���̨ǩ���Ŷ�
            zlIs����ǩ�� = (lngExeState = 0) And Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 0 And Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬
        End If
    End If
End Property

Public Property Get zlIs����ȡ��ǩ��() As Boolean
    '�Ƿ�������ǩ��
    Dim lng����ID As Long, lngExeState As Long, lngTrunState As Long
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIs����ȡ��ǩ�� = False
        Else
            lng����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            lngTrunState = Val(Split(lvwMain.SelectedItem.Tag, "|")(8))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            'δ�����ת�������
            zlIs����ȡ��ǩ�� = (lngExeState = 0 Or lngExeState = 2 And lngTrunState = 0) And Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1
            '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬
        End If
    ElseIf tbPage.Item(midx.idx_ԤԼ����).Selected Then
        If LvwYY.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIs����ȡ��ǩ�� = False
        Else
            lng����ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            'ListSubItems(4).Tag:0-�����÷���̨ǩ���Ŷ�;1-���÷���̨ǩ���Ŷ�
            zlIs����ȡ��ǩ�� = (lngExeState = 0) And Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 1 And Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬
        End If
    End If
End Property
Public Property Get zlIsRegistData() As Boolean
    If Me.ActiveControl Is Me.lvwHZPati Then
        zlIsRegistData = Not Me.lvwHZPati.SelectedItem Is Nothing
        Exit Property
    End If
    If Me.ActiveControl Is Me.LvwYY Then
        zlIsRegistData = Not Me.LvwYY.SelectedItem Is Nothing
        Exit Property
    End If
    zlIsRegistData = Not Me.lvwMain.SelectedItem Is Nothing
End Property

Public Property Get zlIs�������(Optional ByRef bytQueue As Byte) As Boolean
    '�Ƿ��������
    Dim lng����ID As Long, lngExeState As Long
    If Me.ActiveControl Is Me.lvwHZPati And Not Me.lvwHZPati.SelectedItem Is Nothing Then
        '���ﲡ��
        'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
        zlIs������� = Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 2 Or Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3
        bytQueue = IIf(Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3, 1, 0)
    Else
        zlIs������� = False
    End If
End Property
Public Property Get zlIs����ȡ������() As Boolean
    '�Ƿ�������ǩ��
    Dim lng����ID As Long, lngExeState As Long
    If Not Me.ActiveControl Is Me.lvwHZPati Or lvwHZPati.SelectedItem Is Nothing Then
        zlIs����ȡ������ = False
    Else
        lng����ID = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(6))
        'ListSubItem(3).tag:A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
        zlIs����ȡ������ = Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3
        '!����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !ID & "|" & !�ű� & "|" & !ִ��״̬
    End If
End Property

Public Property Get zlGet����ID() As Long
    If lvwMain.SelectedItem Is Nothing Then zlGet����ID = 0: Exit Property
    zlGet����ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
End Property
 Public Property Get zlGet�Һ�NO() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet�Һ�NO = "": Exit Property
    zlGet�Һ�NO = lvwMain.SelectedItem.Text
End Property
 Public Property Get zlGet�Һ�ҽ��() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet�Һ�ҽ�� = "": Exit Property
    zlGet�Һ�ҽ�� = lvwMain.SelectedItem.SubItems(EnmCol.Enmҽ��)
End Property
 Public Property Get zlGet�Һ�����() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet�Һ����� = "": Exit Property
    zlGet�Һ����� = lvwMain.SelectedItem.SubItems(EnmCol.Enm����)
End Property
 Public Property Get zlGet�Һ�ִ��״̬() As Integer
    If lvwMain.SelectedItem Is Nothing Then zlGet�Һ�ִ��״̬ = 0: Exit Property
    zlGet�Һ�ִ��״̬ = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
End Property

Public Property Get zlGet�Һ�ID() As Long
    If lvwMain.SelectedItem Is Nothing Then zlGet�Һ�ID = 0: Exit Property
    zlGet�Һ�ID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
End Property
 
Private Sub lvwHZPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwHZPati.Sorted = True
    If lvwHZPati.SortKey = ColumnHeader.index - 1 Then
        If lvwHZPati.SortOrder = lvwAscending Then
            lvwHZPati.SortOrder = lvwDescending
        Else
            lvwHZPati.SortOrder = lvwAscending
        End If
    Else
        lvwHZPati.SortKey = ColumnHeader.index - 1
    End If
End Sub
Private Sub lvwHZPati_DblClick()
    If Not lvwHZPati.SelectedItem Is Nothing Then
        Call zlExcuteFunction
    End If
End Sub

Private Sub lvwHZPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng����ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem


    '�д����˳�
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    lvwHZPati.Tag = Item.Text

    '�����Ƿ��Ѿ���������(���ڲ���id)��ִ��״̬�������Ƿ�ɷ�����š�������������ɽ����ϵ�в���
    lng����ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre����ID = lng����ID

    RaiseEvent zlShowInfor("���ݺ�:" & Item.Text & _
        "  ����:" & Item.SubItems(EnmCol.Enm����) & _
        "  ����:" & IIf(Item.SubItems(EnmCol.Enm����) = "", "δ����", Item.SubItems(EnmCol.Enm����)) & _
        "  ҽ��:" & IIf(Item.SubItems(EnmCol.Enmҽ��) = "", "δָ��", Item.SubItems(EnmCol.Enmҽ��)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwHZPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        Call ReadRoom
        RaiseEvent zlPopuMenu(Button, Shift, X, Y)
    End If
End Sub
Public Sub zlExc����(ByVal blnȡ������ As Boolean, Optional ByVal blnClick As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ǩ��
    '����:���˺�
    '����:2010-12-08 10:56:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTriage As Boolean, lng����ID  As Long, lngExeState As Long
    Dim lngID As Long, strTittle As String, strSQL As String
    Dim strNO As String, bln����̨ǩ���Ŷ� As Boolean
    
    If Not Me.ActiveControl Is lvwHZPati Then
        Exit Sub
    End If
    If lvwHZPati.SelectedItem Is Nothing Then Exit Sub
    bln����̨ǩ���Ŷ� = lvwHZPati.SelectedItem.ListSubItems(4).Tag = 1
    lng����ID = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(6))
    blnTriage = (lngExeState = 0)
    lngID = Val(lvwHZPati.SelectedItem.ListSubItems(1).Tag)
    strNO = Trim(lvwHZPati.SelectedItem.Text)
    Err = 0: On Error GoTo Errhand:
    If lngID = 0 Then Exit Sub
    If ExcPlugInFun(IIf(blnȡ������, 15, 5), lngID) = False Then Exit Sub
    
    If Not blnȡ������ Then '����ǩ��
        '95637:���ϴ�,2016/7/18,ǩ�����鵱ǰ�ű��Ƿ����Ŷ��У����ߵ����������ű����Ŷ���
        If Checkǩ��(False, lng����ID, lngID, , blnClick, bln����̨ǩ���Ŷ�) = False Then Exit Sub
        If frmDistRoomHz.ShowMe(mfrmMain, mlngModul, mstrPrivs, strNO) = False Then Exit Sub
        strTittle = IIf(blnȡ������, "ȡ������ɹ�!", "���˻���ɹ�!")
        ShowMsgbox strTittle
        RaiseEvent zlShowInfor(strTittle)
        Call ShowBills(False, "")
        Exit Sub
    End If
    'Zl_���˹Һż�¼_ȡ������
    strSQL = "Zl_���˹Һż�¼_ȡ������("
    '  Id_In     ���˹Һż�¼.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  �����_In Integer:=0
    strSQL = strSQL & "0)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strTittle = IIf(blnȡ������, "ȡ������ɹ�!", "���˻���ɹ�!")
    ShowMsgbox strTittle
    RaiseEvent zlShowInfor(strTittle)
    Call ShowBills(False, "")
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub zlPrintBill(ByVal lngID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡָ�����Ŷӵ�
    '����:���˺�
    '����:2011-05-24 15:57:41
    '����:38165
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
     blnPrint = True
     If InStr(1, mstrPrivs, ";�����Ŷӵ�;") = 0 Then Exit Sub
     
     Select Case Val(zlDatabase.GetPara("�Ŷӵ���ӡ", glngSys, mlngModul))
     Case 0
         blnPrint = False
     Case 2
         If MsgBox("���Ƿ�Ҫ��ӡ�Ŷӵ���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnPrint = False
     End Select
     If blnPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113", Me, "�Һ�ID=" & lngID, 2)
End Sub
Public Sub zlRePrintBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش��Ŷӵ�
    '����:���˺�
    '����:2011-05-24 16:36:20
    '����:38165
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, strSQL As String, rsTemp As ADODB.Recordset
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    lngID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
    strNO = Trim(lvwMain.SelectedItem.Text)
    If lngID = 0 Then Exit Sub
    strSQL = "Select  1 From �ŶӽкŶ��� Where ҵ������=0 and ҵ��ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If Not rsTemp.EOF Then
        Call zlPrintBill(lngID)
    Else
        MsgBox "�ò���δ�����ŶӶ��У����ܴ�ӡ�Ŷӵ�!", vbInformation + vbOKOnly, gstrSysName
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҳ
    '����:��⸣
    '����:2013-05-02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    tbPage.RemoveAll
    
    Set ObjItem = tbPage.InsertItem(midx.idx_�ŶӶ���, "�ҺŲ���", Me.lvwMain.Hwnd, 0)
    ObjItem.Tag = midx.idx_�ŶӶ���
    Set ObjItem = tbPage.InsertItem(midx.idx_ԤԼ����, "ԤԼ����", Me.LvwYY.Hwnd, 0)
    ObjItem.Tag = midx.idx_ԤԼ����
     With tbPage
         
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    InitPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ShowBillsAppointment(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card, Optional ByVal blnAutoǩ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '��Σ�strValue:��������
    '      bytType:0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
    '����:���˺�
    '����:2011-11-21 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String, strFilter As String
    Dim objList As ListItem
    Dim i As Long, j As Long, blnUnOutJoin As Boolean
    Dim strNO As String, strTmp As String
    Dim str����� As String, str���� As String, str���￨�� As String, strҽ���� As String, str�Һŵ���ʼ�� As String
    Dim str����ID As String
    Dim lng��ǰ����Сʱ As Long, str�ű� As String, lngPatiID As Long
    Dim bln����̨ǩ���Ŷ� As Boolean
    On Error GoTo errHandle
    'ԤԼ�Һż�¼��ˢ��
    If LvwYY.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = LvwYY.SelectedItem.Text
    End If

    LockWindowUpdate LvwYY.Hwnd
    LvwYY.ListItems.Clear
    LvwYY.Sorted = False

    If blnFilter Then
        strFilter = mcllFilter("����")
    Else
         '�����:51223
        lng��ǰ����Сʱ = CLng(zlDatabase.GetPara("��ǰNСʱ����", glngSys, mlngModul, 0))
        strFilter = " And A.����ʱ�� Between Trunc(sysdate)-" & mint��Ч���� & " And sysdate + 1/24 * " & lng��ǰ����Сʱ   'gbytNODay :27600
    End If
    
    str���￨�� = CStr(mcllFilter("���￨��"))
    str���� = CStr(mcllFilter("��������"))
    str����� = CStr(mcllFilter("�����"))
    strҽ���� = CStr(mcllFilter("ҽ����"))
    str�Һŵ���ʼ�� = CStr(mcllFilter("�Һ�NO")(0))
    str����ID = Val(mcllFilter("����ID"))
    If str���� <> "" Or str���￨�� <> "" Or str����� <> "" Or strҽ���� <> "" Or Nvl(str����ID) <> 0 Then blnUnOutJoin = True
    If str���� <> "" Then str���� = str���� & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
            Case 0  '�����в���:��ȱʡ������������
            Case 1  '����ID
                str����ID = Val(strValue)
                strFilter = strFilter & " And A.����ID=[12]"
                blnUnOutJoin = True
            Case 2 '�����
                str����� = strValue
                strFilter = strFilter & " And A.����� = [11]"
                blnUnOutJoin = True
            Case 3  '������ģ������
                str���� = strValue
                strFilter = strFilter & " And A.���� Like [8]"
            Case 4 '�Һŵ�
                str�Һŵ���ʼ�� = strValue
                strFilter = strFilter & " And A.NO=[3]"
            Case 5 'ҽ����
                strҽ���� = strValue
                strFilter = strFilter & " And B.ҽ����=[13]"
                blnUnOutJoin = True
        End Select
    End If
    
    'A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
    strFilter = strFilter & IIf(mstr������� <> "", " And Instr(','||[10]||',',','||A.ִ�в���id||',')>0", "") & _
            " And (Nvl(A.ִ��״̬,0) = 0 And A.���� Is Null" & _
            IIf(mbytViewScrop(0) = 1, " Or nvl(A.ִ��״̬,0) = 0 And A.���� Is Not Null", "") & _
            IIf(mbytViewScrop(1) = 1, " Or A.ִ��״̬ = 2", "") & _
            IIf(mbytViewScrop(2) = 1, " Or A.ִ��״̬ = 1", "") & _
            IIf(mbytViewScrop(3) = 1, " Or A.ִ��״̬ = -1", "") & _
            " ) "
    '����:43012
    'mbyt��������ʽ��:Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)
    
    
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
            "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
            "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
            "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
            "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
            "      NVL(B.����, 0) ����,A.ִ����," & _
            "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־" & vbCrLf & _
            "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e " & vbCrLf & _
            " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & " and ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
            "           And a.ִ�в���id=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
            "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=2 and a.��¼״̬=1" & vbNewLine & _
            "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
                "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
                "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
                "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
                "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
                "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
                "      NVL(B.����, 0) ����,A.ִ����," & _
                "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־" & vbCrLf & _
                "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e " & vbCrLf & _
                " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & " and ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
                "           And a.ִ�в���id=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
                "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=2 and a.��¼״̬=1" & vbNewLine & _
                "  "
        Else
            strSQL = _
                "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
                "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
                "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
                "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
                "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
                "      NVL(B.����, 0) ����,A.ִ����," & _
                "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־" & vbCrLf & _
                "  From ���˹Һż�¼ a,������Ϣ b,�ٴ������Դ c,�ٴ������¼ c1,�շ���ĿĿ¼ d,���ű� e " & vbCrLf & _
                " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & " And ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
                "           And a.ִ�в���id=e.id And a.�����¼id=c1.id And c1.��Դid=c.id And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
                "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=2 and a.��¼״̬=1" & vbNewLine & _
                "  "
        End If
    End If
     '50427
     Select Case mbyt��������ʽ
     Case 0  '���ұ���,����,NO
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����,a.NO "
     Case 1 '���ұ���,����,�Һ�ʱ��
        strSQL = strSQL & vbCrLf & _
        " Order By e.����,c.����, Decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)"
     Case 2 '���ұ���,����,����ʱ��
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����, A.����ʱ��,A.�Ǽ�ʱ�� " '����ţ�51665
     End Select
     
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("�Һ�ʱ��")(0)), CDate(mcllFilter("�Һ�ʱ��")(1)), _
        str�Һŵ���ʼ��, CStr(mcllFilter("�Һ�NO")(1)), _
        CStr(mcllFilter("��Ʊ��")(0)), CStr(mcllFilter("��Ʊ��")(1)), _
        Val(mcllFilter("����")), _
        str����, CStr(mcllFilter("�Һ�Ա")), mstr�������, _
        str�����, str����ID, strҽ����)

    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            str�ű� = Nvl(!�ű�): lngPatiID = !����ID
        End If
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            Set objList = LvwYY.ListItems.Add(, , !NO, "ry", "ry")
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!ִ�в�������)
            objList.SubItems(EnmCol.Enm�Һ���Ŀ) = zlCommFun.Nvl(!�Һ���Ŀ)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm�����) = IIf(!����� = 0, "", CStr(!�����))
            objList.SubItems(EnmCol.Enm�Ա�) = zlCommFun.Nvl(!�Ա�)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enmҽ��) = zlCommFun.Nvl(!ִ����)
            objList.SubItems(EnmCol.Enm����ʱ��) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS") '51774
            objList.SubItems(EnmCol.Enm�Һ�ʱ��) = Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm����) = "" & !����
            objList.SubItems(EnmCol.Enmҽ����) = Nvl(!ҽ����)
            objList.SubItems(EnmCol.EnmժҪ) = Nvl(!ժҪ)
            objList.SubItems(EnmCol.Enm����״̬) = IIf(Nvl(!ִ��״̬, 0) = 1, "�����", IIf(Nvl(!ִ��״̬, 0) = 2, "�ѽ���", IIf(Nvl(!ִ��״̬, 0) = -1, "������", _
                                                       IIf(zlCommFun.Nvl(!����) <> "", "�ѷ���", "������"))))
            '95637�����ϴ���2016/7/17��ԤԼǩ������������ǩ��
            objList.Tag = !����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !id & "|" & !�ű� & "|" & !ִ��״̬ & "|" & Nvl(!��¼��־, 0)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !ִ�в���id
            objList.ListSubItems(3).Tag = Nvl(!��¼��־)
            bln����̨ǩ���Ŷ� = Val(zlDatabase.GetPara("����̨ǩ���Ŷ�", glngSys, mlngModul, 0, , , , Val(!ִ�в���id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln����̨ǩ���Ŷ�, 1, 0)
              
            If str�ű� <> Nvl(!�ű�) Or lngPatiID <> !����ID Then blnAutoǩ�� = False
            '0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
            Select Case Nvl(!ִ��״̬, 0)
            Case 0
                If Not (IsNull(!����) Or !����ID = 0) Then
                    objList.Icon = "yf": objList.SmallIcon = "yf"
                    If Val(Nvl(!��¼��־)) = 1 Or bln����̨ǩ���Ŷ� = False Then objList.ForeColor = &H8000000C
                ElseIf zlCommFun.Nvl(!����) = "ר��" Then
                    objList.ForeColor = RGB(0, 0, 255)      '��ɫ
                End If
                'A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
                If Val(Nvl(!��¼��־)) = 1 Then
                    objList.Icon = "rySign_in": objList.SmallIcon = "rySign_in"
                End If
            Case 2
                objList.Icon = "zz": objList.SmallIcon = "zz"
                objList.ForeColor = RGB(255, 192, 0)        '��ɫ
            Case 1
                objList.Icon = "yz": objList.SmallIcon = "yz"
                objList.ForeColor = RGB(255, 0, 0)          '��ɫ
            Case -1
                objList.ForeColor = &HC000&                 '��ɫ
            End Select
            For j = 1 To LvwYY.ColumnHeaders.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
        '95637:���ϴ�,2016/7/18,���ֻ��һ�����͵ĹҺŵ���ֱ��ǩ��
        If tbPage.Item(midx.idx_ԤԼ����).Selected And blnAutoǩ�� Then
            Call ShowBillsAppointment(blnFilter, strValue, bytType, objCard)         'ˢ���б���˳�
            If zlIs����ǩ��() Then Screen.MousePointer = 0: Call zlExcǩ��(False)
            Screen.MousePointer = 0
            Exit Sub
        End If
    End With

    If Me.LvwYY.ListItems.Count > 0 Then
        LvwYY.ListItems(1).Selected = True
        For i = 1 To LvwYY.ListItems.Count
            If LvwYY.ListItems(i).Text = strNO Then
                LvwYY.ListItems(i).Selected = True
                LvwYY.Drag 0
                LvwYY.Drag 2
                Exit For
            End If
        Next
        LvwYY.SelectedItem.EnsureVisible
        'lvwYY_ItemClick LvwYY.SelectedItem
    Else
        lvwRoom.ListItems.Clear
    End If
    LockWindowUpdate 0
    LvwYY.Refresh
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Public Property Get zlGetRegistIDsed() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ϵĹҺ��б�
    '����:�Һ��б�,����ö��ŷ���
    '����:���˺�
    '����:2014-03-11 16:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetRegistIDsed = mstrRegistIdsed
End Property
 

 Public Sub SendMsgModule(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ϣ���ʹ���
    '���: strNO-�Һŵ���
    '����:���˺�
    '����:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    '0-�Һ�,1-ԤԼ,2-����,3-ȡ��ԤԼ ,4-�˺� ԤԼ������ģʽ:0-�Һ�,��ʱԤԼҪ�շ�,1-ԤԼ,���շ�
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub



    strSQL = "" & _
    " Select A.id ,A.����,nvl(A.�����,B.�����) as �����,A.����Id,b.���֤��,A.NO,A.ִ�в���ID,C.���� as ִ�в�������,A.����,A.ִ����  " & _
    " From ���˹Һż�¼ A,������Ϣ B,���ű� C  " & _
    " where A.No=[1] and a.��¼״̬ =1 And a.��¼����=1 and a.����ID=b.����id and a.ִ�в���id=c.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    '3.1.1.  ZLHIS_REGIST_002 -�������֪ͨ
    '�ڵ�����    ����    ����    �ظ�    ����    ȱʡֵ  ֵ������
    '<patient_info>
    '    <patient_id>����ID</patient_id>
    '    <patient_name>��������</patient_name>
    '    <identity_card>���֤��</identity_card>
    '    <out_number>�����</out_number>
    '</patient_info>
    '<register_info>
    '    <register_id>�Һ�id</register_id>
    '    <register_no>�Һŵ���</register_no>
    '    <register_dept_id>�Һſ���id</register_dept_id>
    '    <register_dept_title>�Һſ���</register_dept_title>
    '    <register_room>�Һ�����</register_room>
    '    <register_doctor>�Һ�ҽ��</register_doctor>
    '</register_info>
    zlXML.ClearXmlText
 
    Call zlXML.AppendNode("patient_info")
        Call zlXML.appendData("patient_id", Val(Nvl(rsTemp!����ID)))
        Call zlXML.appendData("patient_name", Nvl(rsTemp!����))
        Call zlXML.appendData("identity_card", Nvl(rsTemp!���֤��))
        Call zlXML.appendData("out_number", Nvl(rsTemp!�����))
    Call zlXML.AppendNode("patient_info", True)
    
    Call zlXML.AppendNode("triage_info")
        Call zlXML.appendData("register_id", Val(Nvl(rsTemp!id)))
        Call zlXML.appendData("register_no", strNO)
        Call zlXML.appendData("register_dept_id", Val(Nvl(rsTemp!ִ�в���id)))
        Call zlXML.appendData("register_dept_title", Nvl(rsTemp!ִ�в�������))
        Call zlXML.appendData("register_doctor", Nvl(rsTemp!ִ����))
        Call zlXML.appendData("triage_room", Nvl(rsTemp!����))
    Call zlXML.AppendNode("triage_info", True)
    Call mobjMsgModule.CommitMessage("ZLHIS_REGIST_002", zlXML.XmlText)
    zlXML.ClearXmlText
 End Sub
 
 Public Sub zlModiyPatiBaseInfo(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ��������˻�����Ϣ
    '��Σ�frmMain-������
    '���ƣ����ϴ�
    '���ڣ�2014-07-03
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim lng�Һ�ID As Long
    Dim strInfo As String
    On Error GoTo Errhand
    
    lng����ID = 0: lng�Һ�ID = 0
    
    If Me.ActiveControl Is lvwHZPati Then
        If Not lvwHZPati.SelectedItem Is Nothing Then
            lng����ID = CLng(Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0)))
            lng�Һ�ID = CLng(Val(Split(lvwHZPati.SelectedItem.Tag, "|")(4)))
        End If
    ElseIf tbPage.Item(midx.idx_�ŶӶ���).Selected And Not lvwMain.SelectedItem Is Nothing Then
        lng����ID = CLng(Val(Split(lvwMain.SelectedItem.Tag, "|")(0)))
        lng�Һ�ID = CLng(Val(Split(lvwMain.SelectedItem.Tag, "|")(4)))
    ElseIf tbPage.Item(midx.idx_ԤԼ����).Selected And Not LvwYY.SelectedItem Is Nothing Then
        lng����ID = CLng(Val(Split(LvwYY.SelectedItem.Tag, "|")(0)))
        lng�Һ�ID = CLng(Val(Split(LvwYY.SelectedItem.Tag, "|")(4)))
    End If
    
    If mobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If Not mobjPublicPatient Is Nothing Then
        If mobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) Then
            If mobjPublicPatient.ModipatiBaseInfo(Me, "�������", lng����ID, lng�Һ�ID, 1) Then
                '����ˢ��
                zlRefreshData (True)
            End If
            Exit Sub
        End If
    End If
    MsgBox "����������Ϣ��������(zlPublicPatient.clsPublicPatient)ʧ�ܣ�", vbExclamation, gstrSysName
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlPrintBarcode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ϊָ���Ĳ��˴�ӡ����
    '����:���ϴ�
    '����:2014/9/2 09:43
    '����:77412
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, lng����ID As Long, strNO As String
    Dim objLvw As ListView
    blnPrint = True
    If InStr(1, mstrPrivs, ";�����ӡ;") = 0 Then Exit Sub
    'ûѡ����˳�
    If tbPage.Item(midx.idx_�ŶӶ���).Selected Then
        Set objLvw = lvwMain
        If objLvw.SelectedItem Is Nothing Then Exit Sub
    End If
    If tbPage.Item(midx.idx_ԤԼ����).Selected Then
        Set objLvw = LvwYY
        If objLvw.SelectedItem Is Nothing Then Exit Sub
    End If
    lng����ID = CLng(Val(Split(objLvw.SelectedItem.Tag, "|")(0)))
    strNO = Trim(objLvw.SelectedItem.Text)
    
    Select Case Val(zlDatabase.GetPara("�����ӡ��ʽ", glngSys, mlngModul))
    Case 0
         blnPrint = False
    Case 2
         If MsgBox("���Ƿ�Ҫ��ӡ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnPrint = False
    End Select
    If blnPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113_1", Me, "����ID=" & lng����ID, "NO=" & strNO, "PrintEmpty=0", 2)
End Sub

Private Function Checkǩ��(ByVal blnԤԼ As Boolean, ByVal lng����ID As Long, ByVal lng�Һ�ID As Long, _
                Optional ByVal str����ʱ�� As String, Optional ByVal blnNeedMsg As Boolean, _
                Optional bln����̨ǩ���Ŷ� As Boolean) As Boolean
    '���ܣ���鵱ǰ�Һŵ��Ƿ�����ǩ��
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strMsg As String
    '����̨��ǩ���Ĳ��˶�Ӧ����Ҫ������Ĳ��ˣ����м�鵱����Ŷ���Ϣ
    On Error GoTo Errhand
    '�������ң�
        'û���Ŷ����ɵ���Ķ���
        '���С��ǰʱ�����ɵ���Ķ���
        'ԤԼ�ڷ����ʱ�����ɶ���
    If blnԤԼ Then
        If CDate(Format(str����ʱ��, "YYYY-MM-DD")) > zlDatabase.Currentdate Then
            strMsg = "ԤԼǩ��ֻ��Խ�����ǰ�ĵ��ݣ����Ҫ��ǰǩ�����뵽����ҺŹ�����ǰ����!"
            If blnNeedMsg Then
                MsgBox strMsg, vbInformation, gstrSysName
            ElseIf bln����̨ǩ���Ŷ� Then
                RaiseEvent zlShowInfor(strMsg)
            End If
            Exit Function
        End If
    End If
    strSQL = "Select ҵ��ID,�ŶӺ��� From �ŶӽкŶ��� Where ����ID= [1] And  Trunc(�Ŷ�ʱ��) < sysdate And ҵ������ = 0 And �Ŷ�״̬ IN (0,1,7)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鵱ǰ���˵��ŶӶ���", lng����ID)
    If rsTemp.RecordCount = 0 Then
    ElseIf rsTemp.RecordCount > 1 Or (rsTemp.RecordCount = 1 And rsTemp!ҵ��ID <> lng�Һ�ID) Then
        If bln����̨ǩ���Ŷ� Then '�Һ������Ŷ��»����ǩ����һ��������ǩ��
            If blnNeedMsg Then
                If MsgBox("�����������Һ���Ŀ�������Ŷ��У���ʱǩ����ȡ���Ŷӣ��Ƿ����ǩ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                strMsg = "�����������Һ���Ŀ�������Ŷ��У������Զ����ǩ��!"
                RaiseEvent zlShowInfor(strMsg)
                Exit Function
            End If
        End If
    ElseIf rsTemp!ҵ��ID = lng�Һ�ID Then
        strMsg = "�����������Ŷ��У���������ǩ��!"
        If blnNeedMsg Then
            MsgBox strMsg, vbInformation, gstrSysName
        ElseIf bln����̨ǩ���Ŷ� Then
            RaiseEvent zlShowInfor(strMsg)
        End If
        Exit Function
    End If
    Checkǩ�� = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFilterCons(ByVal strFindValue As String, ByVal objCard As Card, ByVal bytReadType As Byte, _
                               ByRef strFindValue_Out As String, ByRef bytType_Out As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��صĲ�������
    '���:bytReadType-��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
    '����:bytType_Out-0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-02-08 16:09:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����� As String, str���� As String, str����ID As String
    Dim strPassWord As String, strErrMsg As String, lng����ID As Long
    On Error GoTo errHandle
    
    strFindValue_Out = "": bytType_Out = 0
    If strFindValue = "" Then GetFilterCons = True: Exit Function
    
    If bytReadType = 1 Then
        '������ˢ��
        If objCard.���� = "����" Or objCard.���� Like "*��*��*" Then
             If gobjSquare.objSquareCard.zlGetPatiID(mlngDefaultCardID, strFindValue, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        Else
            If gobjSquare.objSquareCard.zlGetPatiID(objCard.�ӿ����, strFindValue, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
        End If
        
        strFindValue_Out = lng����ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function
    ElseIf bytReadType = 2 Or objCard.���� = "���֤��" Or objCard.���� = "�������֤" Then '��ȡ���֤
            If gobjSquare.objSquareCard.zlGetPatiID("���֤��", strFindValue, False, lng����ID, _
                strPassWord, strErrMsg) = False Then lng����ID = 0
        strFindValue_Out = lng����ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function

    ElseIf bytReadType = 3 Or objCard.���� = "IC����" Then '��ȡIC��
        If gobjSquare.objSquareCard.zlGetPatiID("IC����", strFindValue, False, lng����ID, _
            strPassWord, strErrMsg) = False Then lng����ID = 0
        strFindValue_Out = lng����ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function

    ElseIf (Left(strFindValue, 1) = "-" And IsNumeric(Mid(strFindValue, 2))) Then
        str����ID = Val(Mid(strFindValue, 2))
        strFindValue_Out = str����ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function
    ElseIf (Left(strFindValue, 1) = "*" And IsNumeric(Mid(strFindValue, 2))) Or objCard.���� = "�����" Then
        str����� = IIf(Left(strFindValue, 1) = "*", Val(Mid(strFindValue, 2)), Val(strFindValue))
        strFindValue_Out = str�����: bytType_Out = 2
        GetFilterCons = True
        Exit Function
    Else
       Select Case objCard.����
       Case "����"
            str���� = strFindValue & "%"
            strFindValue_Out = str����: bytType_Out = 3
            GetFilterCons = True
            Exit Function
       Case "�Һŵ�"
            strFindValue_Out = strFindValue: bytType_Out = 4
            GetFilterCons = True
            Exit Function
       Case "ҽ����"
            strFindValue_Out = strFindValue: bytType_Out = 5
            GetFilterCons = True
            Exit Function
       Case Else
            '��������,��ȡ��صĲ���ID
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
            If objCard.�ӿ���� <> 0 Then
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.�ӿ����, strFindValue, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strFindValue, False, lng����ID, _
                    strPassWord, strErrMsg) = False Then Exit Function
            End If
            strFindValue_Out = lng����ID: bytType_Out = 1
            GetFilterCons = True
            Exit Function
       End Select
    End If

    GetFilterCons = True
    Exit Function
errHandle:
  Screen.MousePointer = 0
  If ErrCenter = 1 Then
    Resume
End If
  Call SaveErrLog
End Function

Private Sub ShowBillRegister(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card, Optional ByVal blnAutoǩ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�Һ�����
    '��Σ�strValue:��������
    '      bytType:0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
    '����:���˺�
    '����:2018-2-8 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim objList As ListItem, blnUnOutJoin As Boolean
    Dim i As Long, j As Long, strFilter As String
    Dim str����� As String, str���� As String, str���￨�� As String, strҽ���� As String, str�Һŵ���ʼ�� As String
    Dim strNO As String, strTmp As String
    Dim str�ű� As String, lngPatiID As Long
    Dim lng��ǰ����Сʱ As Long, str����ID As String
    Dim bln����̨ǩ���Ŷ� As Boolean
    On Error GoTo errHandle

    '�Һż�¼��ˢ��
    If lvwMain.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = lvwMain.SelectedItem.Text
    End If
   
    LockWindowUpdate lvwMain.Hwnd
    lvwMain.ListItems.Clear
    lvwMain.Sorted = False
    
    If blnFilter Then
        strFilter = mcllFilter("����")
    Else
         '�����:51223
        lng��ǰ����Сʱ = CLng(zlDatabase.GetPara("��ǰNСʱ����", glngSys, mlngModul, 0))
        strFilter = " And A.����ʱ�� Between Trunc(sysdate)-" & mint��Ч���� & " And sysdate + 1/24 * " & lng��ǰ����Сʱ   'gbytNODay :27600
    End If
    
    str���￨�� = CStr(mcllFilter("���￨��"))
    str���� = CStr(mcllFilter("��������"))
    str����� = CStr(mcllFilter("�����"))
    strҽ���� = CStr(mcllFilter("ҽ����"))
    str�Һŵ���ʼ�� = CStr(mcllFilter("�Һ�NO")(0))
    str����ID = Val(mcllFilter("����ID"))
    If str���� <> "" Or str���￨�� <> "" Or str����� <> "" Or strҽ���� <> "" Or Val(str����ID) <> 0 Then blnUnOutJoin = True
    If str���� <> "" Then str���� = str���� & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
        Case 0  '�����в���:��ȱʡ������������
        Case 1  '����ID
            str����ID = Val(strValue)
            strFilter = strFilter & " And A.����ID=[12]"
            blnUnOutJoin = True
        Case 2 '�����
            str����� = strValue
            strFilter = strFilter & " And A.����� = [11]"
            blnUnOutJoin = True
        Case 3  '������ģ������
            str���� = strValue
            strFilter = strFilter & " And A.���� Like [8]"
        Case 4 '�Һŵ�
            str�Һŵ���ʼ�� = strValue
            strFilter = strFilter & " And A.NO=[3]"
        Case 5 'ҽ����
            strҽ���� = strValue
            strFilter = strFilter & " And B.ҽ����=[13]"
            blnUnOutJoin = True
        End Select
    End If
    
    'A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
    strFilter = strFilter & IIf(mstr������� <> "", " And Instr(','||[10]||',',','||A.ִ�в���id||',')>0", "") & _
            " And (Nvl(A.ִ��״̬,0) = 0 And A.���� Is Null" & _
            IIf(mbytViewScrop(0) = 1, " Or nvl(A.ִ��״̬,0) = 0 And A.���� Is Not Null", "") & _
            IIf(mbytViewScrop(1) = 1, " Or A.ִ��״̬ = 2", "") & _
            IIf(mbytViewScrop(2) = 1, " Or A.ִ��״̬ = 1", "") & _
            IIf(mbytViewScrop(3) = 1, " Or A.ִ��״̬ = -1", "") & _
            " ) "
    '����:43012
    'mbyt��������ʽ��:Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)
    
    '74898:���ϴ�,2015/4/9,��ǲ��˵ĺ���״̬
    '95637:���ϴ�,2016/7/17 ǩ������ 0 -����ǩ����1-����ǩ����2-ת��ǩ����4-����ǩ����5-����ǩ��
    '      ת�������ʾת����ң�ת��ҽ����ת������
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select decode(A.ת�����ID,Null,A.����,A.ת������) as ����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
            "      Nvl(A.ת�����ID,A.ִ�в���ID) as ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
            "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
            "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
            "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
            "      NVL(B.����, 0) ����,decode(A.ת�����ID,Null,A.ִ����,A.ת��ҽ��) as ִ����," & _
            "      f.����ҽ�� As ������, f.���� As ��������, f.����ʱ��, " & _
            "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־,f.�Ŷ�״̬" & vbCrLf & _
            "      ,Nvl(A.ת��״̬, 10) as ת��״̬ " & vbCrLf & _
            "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e,�ŶӽкŶ��� f " & vbCrLf & _
            " Where a.����id=b.����id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.ҵ��id(+) and ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
            "           And Nvl(A.ת�����ID,A.ִ�в���ID)=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
            "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1" & vbNewLine & _
            "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
                "Select decode(A.ת�����ID,Null,A.����,A.ת������) as ����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
                "      Nvl(A.ת�����ID,A.ִ�в���ID) as ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
                "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
                "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
                "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
                "      NVL(B.����, 0) ����,decode(A.ת�����ID,Null,A.ִ����,A.ת��ҽ��) as ִ����," & _
                "      f.����ҽ�� As ������, f.���� As ��������, f.����ʱ��, " & _
                "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־,f.�Ŷ�״̬" & vbCrLf & _
                "      ,Nvl(A.ת��״̬, 10) as ת��״̬ " & vbCrLf & _
                "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e,�ŶӽкŶ��� f " & vbCrLf & _
                " Where a.����id=b.����id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.ҵ��id(+) and ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
                "           And Nvl(A.ת�����ID,A.ִ�в���ID)=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
                "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1" & vbNewLine & _
                "  "
        Else
            strSQL = _
                "Select decode(A.ת�����ID,Null,A.����,A.ת������) as ����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
                "      Nvl(A.ת�����ID,A.ִ�в���ID) as ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
                "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
                "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
                "      decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
                "      NVL(B.����, 0) ����,decode(A.ת�����ID,Null,A.ִ����,A.ת��ҽ��) as ִ����," & _
                "      f.����ҽ�� As ������, f.���� As ��������, f.����ʱ��, " & _
                "      nvl(A.ִ��״̬,0) as ִ��״̬,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����,A.��¼��־,f.�Ŷ�״̬" & vbCrLf & _
                "      ,Nvl(A.ת��״̬, 10) as ת��״̬ " & vbCrLf & _
                "  From ���˹Һż�¼ a,������Ϣ b,�ٴ������Դ c,�ٴ������¼ c1,�շ���ĿĿ¼ d,���ű� e,�ŶӽкŶ��� f " & vbCrLf & _
                " Where a.����id=b.����id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.ҵ��id(+) and ((nvl(A.ִ��״̬,0)=2 and nvl(A.��¼��־,0) in (0,1))   or Nvl(A.ִ��״̬,0)<>2 ) " & _
                "           And Nvl(A.ת�����ID,A.ִ�в���ID)=e.id And a.�����¼id=c1.id And c1.��Դid=c.id And c.��Ŀid=d.ID" & vbCrLf & strFilter & _
                "           And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1" & vbNewLine & _
                "  "
        End If
    End If
     '50427
     Select Case mbyt��������ʽ
     Case 0  '���ұ���,����,NO
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����,a.NO "
     Case 1 '���ұ���,����,�Һ�ʱ��
        strSQL = strSQL & vbCrLf & _
        " Order By e.����,c.����, Decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)"
     Case 2 '���ұ���,����,����ʱ��
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����, A.����ʱ��,A.�Ǽ�ʱ�� " '����ţ�51665
     End Select
     
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("�Һ�ʱ��")(0)), CDate(mcllFilter("�Һ�ʱ��")(1)), _
        str�Һŵ���ʼ��, CStr(mcllFilter("�Һ�NO")(1)), _
        CStr(mcllFilter("��Ʊ��")(0)), CStr(mcllFilter("��Ʊ��")(1)), _
        Val(mcllFilter("����")), _
        str����, CStr(mcllFilter("�Һ�Ա")), mstr�������, _
        str�����, str����ID, strҽ����)

    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            str�ű� = Nvl(!�ű�): lngPatiID = !����ID
        End If
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            Set objList = lvwMain.ListItems.Add(, , !NO, "ry", "ry")
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!ִ�в�������)
            objList.SubItems(EnmCol.Enm�Һ���Ŀ) = zlCommFun.Nvl(!�Һ���Ŀ)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm�����) = IIf(!����� = 0, "", CStr(!�����))
            objList.SubItems(EnmCol.Enm�Ա�) = zlCommFun.Nvl(!�Ա�)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enmҽ��) = zlCommFun.Nvl(!ִ����)
            objList.SubItems(EnmCol.Enm����ʱ��) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS") '51774
            objList.SubItems(EnmCol.Enm�Һ�ʱ��) = Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm����) = "" & !����
            objList.SubItems(EnmCol.Enmҽ����) = Nvl(!ҽ����)
            objList.SubItems(EnmCol.EnmժҪ) = Nvl(!ժҪ)
            objList.SubItems(EnmCol.EnmԤԼ) = Nvl(!ԤԼ)
            objList.SubItems(EnmCol.Enm������) = Nvl(!������)
            objList.SubItems(EnmCol.Enm��������) = Nvl(!��������)
            objList.SubItems(EnmCol.Enm����ʱ��) = Nvl(!����ʱ��)
            objList.SubItems(EnmCol.Enm����״̬) = IIf(Nvl(!ִ��״̬, 0) = 1, "�����", IIf(Nvl(!ִ��״̬, 0) = 2, "�ѽ���", IIf(Nvl(!ִ��״̬, 0) = -1, "������", _
                                                       IIf(zlCommFun.Nvl(!����) <> "", "�ѷ���", "������"))))
            '74898:���ϴ�,2015/4/9,��ǲ��˵ĺ���״̬
            objList.SubItems(EnmCol.Enm����) = IIf(Nvl(!�Ŷ�״̬) = 1 Or Nvl(!�Ŷ�״̬) = 7, "��", "")
            objList.Tag = !����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !id & "|" & !�ű� & "|" & !ִ��״̬ & "|" & Nvl(!��¼��־, 0) & "|" & Nvl(!ת��״̬, 10)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !ִ�в���id
            objList.ListSubItems(3).Tag = Nvl(!��¼��־)
            bln����̨ǩ���Ŷ� = Val(zlDatabase.GetPara("����̨ǩ���Ŷ�", glngSys, mlngModul, 0, , , , Val(!ִ�в���id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln����̨ǩ���Ŷ�, 1, 0)
              
            If str�ű� <> Nvl(!�ű�) Or lngPatiID <> !����ID Then blnAutoǩ�� = False
            '0-�ȴ�����,1-��ɾ���,2-���ھ���,-1���Ϊ������
            Select Case Nvl(!ִ��״̬, 0)
            Case 0
                If Not (IsNull(!����) Or !����ID = 0) Then
                    objList.Icon = "yf": objList.SmallIcon = "yf"
                    If Val(Nvl(!��¼��־)) = 1 Or bln����̨ǩ���Ŷ� = False Then objList.ForeColor = &H8000000C
                ElseIf zlCommFun.Nvl(!����) = "ר��" Then
                    objList.ForeColor = RGB(0, 0, 255)      '��ɫ
                End If
                'A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
                If Val(Nvl(!��¼��־)) = 1 Then
                    objList.Icon = "rySign_in": objList.SmallIcon = "rySign_in"
                End If
            Case 2
                objList.Icon = "zz": objList.SmallIcon = "zz"
                objList.ForeColor = RGB(255, 192, 0)        '��ɫ
            Case 1
                objList.Icon = "yz": objList.SmallIcon = "yz"
                objList.ForeColor = RGB(255, 0, 0)          '��ɫ
            Case -1
                objList.ForeColor = &HC000&                 '��ɫ
            End Select
            
            For j = 1 To objList.ListSubItems.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
        '95637:���ϴ�,2016/7/18,���ֻ��һ�����͵ĹҺŵ���ֱ��ǩ��
        If tbPage.Item(midx.idx_�ŶӶ���).Selected And blnAutoǩ�� Then
            Call ShowBillRegister(blnFilter, strValue, bytType, objCard)         'ˢ���б���˳�
            If zlIs����ǩ��() Then Screen.MousePointer = 0: Call zlExcǩ��(False)
            Screen.MousePointer = 0
            Exit Sub
        End If
    End With
    
    If Me.lvwMain.ListItems.Count > 0 Then
        lvwMain.ListItems(1).Selected = True
        For i = 1 To lvwMain.ListItems.Count
            If lvwMain.ListItems(i).Text = strNO Then
                lvwMain.ListItems(i).Selected = True
                lvwMain.Drag 0
                lvwMain.Drag 2
                Exit For
            End If
        Next
        lvwMain.SelectedItem.EnsureVisible
        lvwMain_ItemClick lvwMain.SelectedItem
    Else
        lvwRoom.ListItems.Clear
    End If
    LockWindowUpdate 0
    lvwMain.Refresh
    
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowBillRegisterHZ(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card)
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '��Σ�strValue:��������
    '      bytType:0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
    '����:���˺�
    '����:2018-2-8 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIco As String, objList As ListItem
    Dim i As Long, j As Long, blnUnOutJoin As Boolean
    Dim str����� As String, str���� As String, str���￨�� As String, strҽ���� As String, str�Һŵ���ʼ�� As String
    Dim strNO As String, str����ID As String
    Dim strHzWhere As String '���ﲡ����Ϣ����
    Dim blnBusy As Boolean, strFilter As String
    Dim str�ű� As String, lngPatiID As Long
    Dim lng��ǰ����Сʱ As Long
    Dim bln����̨ǩ���Ŷ� As Boolean
    
    On Error GoTo errHandle
    '���ػ��ﲡ����Ϣ
    If lvwHZPati.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = lvwHZPati.SelectedItem.Text
    End If
    LockWindowUpdate lvwHZPati.Hwnd
    lvwHZPati.ListItems.Clear
    lvwHZPati.Sorted = False
    
    If blnFilter Then
        strFilter = mcllFilter("����")
    Else
         '�����:51223
        lng��ǰ����Сʱ = CLng(zlDatabase.GetPara("��ǰNСʱ����", glngSys, mlngModul, 0))
        strFilter = " And A.����ʱ�� Between Trunc(sysdate)-" & mint��Ч���� & " And sysdate + 1/24 * " & lng��ǰ����Сʱ   'gbytNODay :27600
    End If
    
    str���￨�� = CStr(mcllFilter("���￨��"))
    str���� = CStr(mcllFilter("��������"))
    str����� = CStr(mcllFilter("�����"))
    strҽ���� = CStr(mcllFilter("ҽ����"))
    str�Һŵ���ʼ�� = CStr(mcllFilter("�Һ�NO")(0))
    str����ID = Val(mcllFilter("����ID"))
    If str���� <> "" Or str���￨�� <> "" Or str����� <> "" Or strҽ���� <> "" Or Val(str����ID) <> 0 Then blnUnOutJoin = True
    If str���� <> "" Then str���� = str���� & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-�����в���;1-����ID;2-�����;3-������ģ������;4-�Һŵ�;5-ҽ����
        Case 0  '�����в���:��ȱʡ������������
        Case 1  '����ID
            str����ID = Val(strValue)
            strFilter = strFilter & " And A.����ID=[12]"
            blnUnOutJoin = True
        Case 2 '�����
            str����� = strValue
            strFilter = strFilter & " And A.����� = [11]"
            blnUnOutJoin = True
        Case 3  '������ģ������
            str���� = strValue
            strFilter = strFilter & " And A.���� Like [8]"
        Case 4 '�Һŵ�
            str�Һŵ���ʼ�� = strValue
            strFilter = strFilter & " And A.NO=[3]"
        Case 5 'ҽ����
            strҽ���� = strValue
            strFilter = strFilter & " And B.ҽ����=[13]"
            blnUnOutJoin = True
        End Select
    End If
    
    strHzWhere = strFilter & IIf(mstr������� <> "", " And Instr(','||[10]||',',','||A.ִ�в���id||',')>0", "")
    'A.��¼��־:��ʾ0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
            "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
            "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
            "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
            "      Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��,NVL(B.����, 0) ����, " & _
            "       A.ִ����,nvl(A.ִ��״̬,0) as ִ��״̬,A.��¼��־,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����" & vbCrLf & _
            "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e" & vbCrLf & _
            " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & _
            "           And (A.ִ��״̬=2 and nvl(A.��¼��־,0) in (2,3) )" & _
            "           And a.ִ�в���id=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strHzWhere & _
            " And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1 " & vbNewLine & _
           "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
             "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
             "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
             "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
             "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
             "      Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��,NVL(B.����, 0) ����, " & _
             "       A.ִ����,nvl(A.ִ��״̬,0) as ִ��״̬,A.��¼��־,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����" & vbCrLf & _
             "  From ���˹Һż�¼ a,������Ϣ b,�ҺŰ��� c,�շ���ĿĿ¼ d,���ű� e" & vbCrLf & _
             " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & _
             "           And (A.ִ��״̬=2 and nvl(A.��¼��־,0) in (2,3) )" & _
             "           And a.ִ�в���id=e.id And a.�ű�=c.���� And c.��Ŀid=d.ID" & vbCrLf & strHzWhere & _
             " And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1 " & vbNewLine & _
            "  "
        Else
            strSQL = _
                "Select A.����,A.ID,A.�ű�,C.����,D.���� as �Һ���Ŀ," & vbCrLf & _
                "      A.ִ�в���ID,E.���� as ִ�в�������,A.NO,NVL(A.����ID, 0) ����ID,A.����," & vbCrLf & _
                "      NVL(B.�����, 0) �����,B.���￨��,B.����֤��,A.�Ա�,A.����," & vbCrLf & _
                "      A.����ʱ�� as ����ʱ��," & vbCrLf & _
                "      Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��) as �Ǽ�ʱ��,NVL(B.����, 0) ����, " & _
                "       A.ִ����,nvl(A.ִ��״̬,0) as ִ��״̬,A.��¼��־,A.����,A.ժҪ,decode(A.ԤԼ,1,'��','') as ԤԼ,B.ҽ����" & vbCrLf & _
                "  From ���˹Һż�¼ a,������Ϣ b,�ٴ������Դ c,�ٴ������¼ c1,�շ���ĿĿ¼ d,���ű� e" & vbCrLf & _
                " Where a.����id=b.����id " & IIf(blnUnOutJoin, "", "(+)") & _
                "           And (A.ִ��״̬=2 and nvl(A.��¼��־,0) in (2,3) )" & _
                "           And a.ִ�в���id=e.id And a.�����¼id=c1.id And c1.��Դid = c.id And c.��Ŀid=d.ID" & vbCrLf & strHzWhere & _
                " And (E.վ��='" & gstrNodeNo & "' Or E.վ�� is Null) and a.��¼����=1 and a.��¼״̬=1 " & vbNewLine & _
               "  "
        End If
    End If
    '����:43012
    '50427
     Select Case mbyt��������ʽ
     Case 0  '���ұ���,����,NO
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����,a.NO "
     Case 1 '���ұ���,����,�Һ�ʱ��
        strSQL = strSQL & vbCrLf & _
        " Order By e.����,c.����, Decode(A.ԤԼ,1,nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)"
     Case 2 '���ұ���,����,����ʱ��
        strSQL = strSQL & vbCrLf & " Order By e.����,c.����, A.����ʱ��,A.�Ǽ�ʱ�� " '�����:51665
     End Select
     
    'mbyt��������ʽ��:Decode(A.ԤԼ,1, nvl(A.����ʱ��,A.�Ǽ�ʱ��),A.�Ǽ�ʱ��)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("�Һ�ʱ��")(0)), CDate(mcllFilter("�Һ�ʱ��")(1)), _
        str�Һŵ���ʼ��, CStr(mcllFilter("�Һ�NO")(1)), _
        CStr(mcllFilter("��Ʊ��")(0)), CStr(mcllFilter("��Ʊ��")(1)), _
        Val(mcllFilter("����")), _
        str����, CStr(mcllFilter("�Һ�Ա")), mstr�������, _
        str�����, str����ID, strҽ����)
   With rsTmp
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            '0��ʾ���ﲡ��,1-��ʾǩ���Ĳ���,2-��ʾ��Ҫ�������Ĳ���; 3-��ʾ�ѻ��ﵫ��δ���յĲ���;
            If Val(Nvl(rsTmp!��¼��־)) = 2 Then
                strIco = "ManStop"
                If Nvl(!�Ա�) Like "*Ů*" Then strIco = "WomanStop"
            Else
                strIco = "ManSign_in"
                If Nvl(!�Ա�) Like "*Ů*" Then strIco = "WomanSign_in"
            End If
            Set objList = lvwHZPati.ListItems.Add(, , !NO, strIco, strIco)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!ִ�в�������)
            objList.SubItems(EnmCol.Enm�Һ���Ŀ) = zlCommFun.Nvl(!�Һ���Ŀ)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm�����) = IIf(!����� = 0, "", CStr(!�����))
            objList.SubItems(EnmCol.Enm�Ա�) = zlCommFun.Nvl(!�Ա�)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enm����) = zlCommFun.Nvl(!����)
            objList.SubItems(EnmCol.Enmҽ��) = zlCommFun.Nvl(!ִ����)
            objList.SubItems(EnmCol.Enm����ʱ��) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS") ''51774
            objList.SubItems(EnmCol.Enm�Һ�ʱ��) = Format(!�Ǽ�ʱ��, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm����) = "" & !����
            objList.SubItems(EnmCol.Enmҽ����) = Nvl(!ҽ����)
            objList.SubItems(EnmCol.EnmժҪ) = Nvl(!ժҪ)
            objList.SubItems(EnmCol.EnmԤԼ) = Nvl(!ԤԼ)
            objList.SubItems(EnmCol.Enm����״̬) = IIf(Nvl(!ִ��״̬, 0) = 1, "�����", IIf(Nvl(!ִ��״̬, 0) = 2, "�ѽ���", IIf(Nvl(!ִ��״̬, 0) = -1, "������", _
                                                       IIf(zlCommFun.Nvl(!����) <> "", "�ѷ���", "������"))))
            objList.Tag = !����ID & "|" & !���� & "|" & !���￨�� & "|" & !����֤�� & "|" & !id & "|" & !�ű� & "|" & !ִ��״̬ & "|" & Nvl(!��¼��־, 0)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !ִ�в���id
            objList.ListSubItems(3).Tag = Nvl(!��¼��־)
            bln����̨ǩ���Ŷ� = Val(zlDatabase.GetPara("����̨ǩ���Ŷ�", glngSys, mlngModul, 0, , , , Val(!ִ�в���id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln����̨ǩ���Ŷ�, 1, 0)
            'objList.ForeColor = RGB(255, 192, 0)        '��ɫ
            For j = 1 To lvwHZPati.ColumnHeaders.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
    End With
    If mstrRegistIdsed <> "" Then mstrRegistIdsed = Mid(mstrRegistIdsed, 2)
    
    If Me.lvwHZPati.ListItems.Count > 0 Then
        lvwHZPati.ListItems(1).Selected = True
        For i = 1 To lvwHZPati.ListItems.Count
            If lvwHZPati.ListItems(i).Text = strNO Then
                lvwHZPati.ListItems(i).Selected = True
                lvwHZPati.Drag 0
                lvwHZPati.Drag 2
                Exit For
            End If
        Next
        lvwHZPati.SelectedItem.EnsureVisible
        lvwHZPati_ItemClick lvwHZPati.SelectedItem
    End If
    LockWindowUpdate 0
    lvwHZPati.Refresh
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRooms()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim objList As ListItem, i As Integer, j As Integer
    Dim strTmp As String
    Dim blnBusy As Boolean
    On Error GoTo errHandle

    '79694:���ϴ�,2014/11/25,���ݲ�����ȡ��������
    blnBusy = Val(zlDatabase.GetPara("����æʱ�������", glngSys, mlngModul, 0)) = 1
    '�������ˢ��
    If gbytRegistMode = 0 Then
        strSQL = "Select Distinct R.����, R.����, R.ȱʡ��־ As æ��״̬, T.����, T.����, T.��������" & vbNewLine & _
                "From �������� R, �ҺŰ������� S, �ҺŰ��� P," & vbNewLine & _
                "     (Select ����, Sum(Decode(ִ��״̬, Null, 1, 0, 1, 0)) As ����, Sum(Decode(ִ��״̬, 2, 1, 0)) As ����," & vbNewLine & _
                "              Sum(Decode(ִ��״̬, 1, Decode(Sign(Trunc(Sysdate) - ִ��ʱ��), 1, 0, 1), 0)) As ��������" & vbNewLine & _
                "       From ���˹Һż�¼" & vbNewLine & _
                "       Where ����ʱ�� > Sysdate - " & mint��Ч���� & " And ���� Is Not Null and ��¼����=1 and ��¼״̬=1  " & vbNewLine & _
                "       Group By ����) T" & vbNewLine & _
                "Where R.���� = S.�������� And S.�ű�id = P.ID And R.���� = T.����(+) " & _
                IIf(blnBusy, " ", " And R.ȱʡ��־=0 ") & _
                " And (R.վ��='" & gstrNodeNo & "' Or R.վ�� is Null)" & vbNewLine & _
                IIf(mstr������� <> "", " And Instr(','||[1]||',',','||P.����id||',')>0", "") & vbNewLine & _
                "Order By R.����"
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select Distinct R.����, R.����, R.ȱʡ��־ As æ��״̬, T.����, T.����, T.��������" & vbNewLine & _
                    "From �������� R, �ҺŰ������� S, �ҺŰ��� P," & vbNewLine & _
                    "     (Select ����, Sum(Decode(ִ��״̬, Null, 1, 0, 1, 0)) As ����, Sum(Decode(ִ��״̬, 2, 1, 0)) As ����," & vbNewLine & _
                    "              Sum(Decode(ִ��״̬, 1, Decode(Sign(Trunc(Sysdate) - ִ��ʱ��), 1, 0, 1), 0)) As ��������" & vbNewLine & _
                    "       From ���˹Һż�¼" & vbNewLine & _
                    "       Where ����ʱ�� > Sysdate - " & mint��Ч���� & " And ���� Is Not Null and ��¼����=1 and ��¼״̬=1  " & vbNewLine & _
                    "       Group By ����) T" & vbNewLine & _
                    "Where R.���� = S.�������� And S.�ű�id = P.ID And R.���� = T.����(+) " & _
                    IIf(blnBusy, " ", " And R.ȱʡ��־=0 ") & _
                    " And (R.վ��='" & gstrNodeNo & "' Or R.վ�� is Null)" & vbNewLine & _
                    IIf(mstr������� <> "", " And Instr(','||[1]||',',','||P.����id||',')>0", "") & vbNewLine & _
                    "Order By R.����"
        Else
            strSQL = "Select Distinct R.����, R.����, R.ȱʡ��־ As æ��״̬, T.����, T.����, T.��������" & vbNewLine & _
                    "From �������� R, �ٴ��������Ҽ�¼ S, �ٴ������¼ P," & vbNewLine & _
                    "     (Select ����, Sum(Decode(ִ��״̬, Null, 1, 0, 1, 0)) As ����, Sum(Decode(ִ��״̬, 2, 1, 0)) As ����," & vbNewLine & _
                    "              Sum(Decode(ִ��״̬, 1, Decode(Sign(Trunc(Sysdate) - ִ��ʱ��), 1, 0, 1), 0)) As ��������" & vbNewLine & _
                    "       From ���˹Һż�¼" & vbNewLine & _
                    "       Where ����ʱ�� > Sysdate - " & mint��Ч���� & " And ���� Is Not Null and ��¼����=1 and ��¼״̬=1  " & vbNewLine & _
                    "       Group By ����) T" & vbNewLine & _
                    "Where R.id = S.����id And S.��¼id = P.ID And R.���� = T.����(+) " & _
                    IIf(blnBusy, " ", " And R.ȱʡ��־=0 ") & _
                    " And (R.վ��='" & gstrNodeNo & "' Or R.վ�� is Null)" & vbNewLine & _
                    IIf(mstr������� <> "", " And Instr(','||[1]||',',','||P.����id||',')>0", "") & vbNewLine & _
                    "Order By R.����"
        End If
    End If
    'gbytNODay: ����:27600

    LockWindowUpdate lvwRoom.Hwnd
    lvwRoom.ListItems.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�������)
    With rsTmp
        If .RecordCount > 0 Then rsTmp.MoveFirst
        Do While Not .EOF
            Set objList = lvwRoom.ListItems.Add(, "K" & !����, !����, "bm", "bm")
            objList.SubItems(1) = IIf(!æ��״̬ <> 0, "æ", "")
            objList.SubItems(2) = Format(!����, "0;0; ; ")
            objList.SubItems(3) = Format(!����, "0;0; ; ")
            objList.SubItems(4) = Format(!��������, "0;0; ; ")
            objList.SubItems(5) = ""
            If !æ��״̬ <> 0 Then
                objList.ForeColor = RGB(255, 0, 0)
                For j = 1 To Me.lvwRoom.ColumnHeaders.Count - 1
                    objList.ListSubItems(j).ForeColor = objList.ForeColor
                Next
            End If
            If InStr(1, strTmp & ",", "," & !���� & "") = 0 Then
                strTmp = strTmp & "," & !���� & ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    strTmp = Mid(strTmp, 2)
    If gbytRegistMode = 0 Then
        strSQL = "Select Distinct S.��������,D.����" & _
                " From �ҺŰ������� S,�ҺŰ��� P,���ű� D" & _
                " Where S.�ű�id=P.ID And P.����id=D.ID And Instr(','||[1]||',',','||S.��������||',')>0" & _
                " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
                IIf(mstr������� <> "", " And Instr(','||[2]||',',','||P.����id||',')>0", "")
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select Distinct S.��������,D.����" & _
                    " From �ҺŰ������� S,�ҺŰ��� P,���ű� D" & _
                    " Where S.�ű�id=P.ID And P.����id=D.ID And Instr(','||[1]||',',','||S.��������||',')>0" & _
                    " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
                    IIf(mstr������� <> "", " And Instr(','||[2]||',',','||P.����id||',')>0", "")
        Else
            strSQL = "Select Distinct S.���� As ��������,D.����" & _
                    " From �������� S,�ٴ��������Ҽ�¼ S1,�ٴ������¼ P,�ٴ������Դ E,���ű� D" & _
                    " Where S1.��¼id=P.ID And P.��ԴID=E.ID And E.����id=D.ID And S.ID=S1.����ID And Instr(','||[1]||',',','||S.����||',')>0" & _
                    " And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
                    IIf(mstr������� <> "", " And Instr(','||[2]||',',','||E.����id||',')>0", "")
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp, mstr�������)

    '����ɷ��䵽�����ҵĿ���
    For Each objList In Me.lvwRoom.ListItems
        strTmp = ""
        rsTmp.Filter = "��������='" & objList.Text & "'"
        Do While Not rsTmp.EOF
            If InStr(1, strTmp & ";", ";" & rsTmp!���� & ";") = 0 Then strTmp = strTmp & ";" & rsTmp!����
            rsTmp.MoveNext
        Loop
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        objList.SubItems(5) = strTmp
    Next
    LockWindowUpdate 0
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

