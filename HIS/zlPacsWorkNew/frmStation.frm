VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.0#0"; "zlIDKind.ocx"
Begin VB.Form frmStation 
   Caption         =   "Ӱ��ҽ������վ"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11325
   Icon            =   "frmStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   6015
      ScaleWidth      =   4500
      TabIndex        =   2
      Top             =   720
      Width           =   4495
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   250
         Left            =   495
         TabIndex        =   4
         Top             =   60
         Width           =   1485
      End
      Begin VB.TextBox Txt������Ϣ 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3285
         Width           =   4205
      End
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7875
      Top             =   165
   End
   Begin zlIDKind.IDKind IDKind 
      Bindings        =   "frmStation.frx":1CFA
      Height          =   360
      Left            =   6975
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStation.frx":1D0E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7832
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   2850
      Top             =   60
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
            Picture         =   "frmStation.frx":25A2
            Key             =   "����"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":2B3C
            Key             =   "סԺ"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3416
            Key             =   "����"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3570
            Key             =   "Ӱ��"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":3CEA
            Key             =   "�ѽ�"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":4084
            Key             =   "��ɫͨ��"
            Object.Tag             =   "6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1980
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":41DE
            Key             =   "��ѡ����"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStation.frx":4778
            Key             =   "��ѡѡ��"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmStation.frx":4D12
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mcol
    Col���� = 0: Col��Դ: Col����: Col����: Col����: Col����: Col������: Col�Ա�: Col����: Col����: Col��λ: Colִ�м�: Col���ʱ��: Col����ʱ��: Col����ҽ��
    
    Col��� = 15: Col����: ColӤ��: Col�Ǽ���: Col������: Col�����: Col��ӡ��Ƭ: Col�������: Col��ɫͨ��: Col�����ӡ: Col������: Col������: Col��鼼ʦ: Col��ͼʱ��
    
    ColӰ����� = 29: Col����ID: Col��ҳID: Col�Һŵ�: Colҽ��ID: Col���ͺ�: Col���UID: Col���״̬: Colת�� '��29�п�ʼ����ʾ
End Enum

Private Enum FilterID
    ID_���� = 4001: ID_סԺ = 4002: ID_��� = 4003: ID_���� = 4004
    ID_���� = 4005: ID_�ѽ� = 4006: ID_δ�� = 4007: ID_�Ǽ� = 4008
    ID_���� = 4009: ID_���� = 4010: ID_��� = 4011: ID_��� = 4012
    ID_���˷�ʽ = 4013: ID_����ֵ = 4014: ID_��ʼ���� = 4015: ID_����סԺ = 4016
End Enum

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0��������￨ = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
End Enum

Private mlngCur����ID As Long                               '��ǰ����ID
Private mstrCur���� As String                               '��ǰ���� ����-����
Private mstrCanUse���� As String                            '��ǰ���ÿ���  ID_����-����
Private mstrCurFindtype As String                           '��λ��ʽ
Private mblnFinishCommit As Boolean                         '�ޱ��������,�Ƿ������ٴ�ȷ��
Private mblnCompleteCommit As Boolean                       '��˺������ٴ�ȷ��
Private mblnInitOK As Boolean                               '��ʼ�����
Private mblnIgnoreResult As Boolean                         '���������� '=true ����
Private mblnShowImgAtReport As Boolean                      '�򿪱���ʱ�򿪹�Ƭվ
Private mblnReportWithImage As Boolean                      '��ͼ�����д���棬��ͼ�񲻿�д����
Private mblnReportWithResult As Boolean                     '��Ӱ�����Ϊ����
Private mblnLocalizerBackward As Boolean                    '��λƬ����
Private mblnPacsReport As Boolean                           '�Ƿ�ʹ��PACS����༭����Fasleʱʹ�õ��Ӳ����༭��
Private mblnPrintCommit As Boolean                          '��ӡ��ֱ�����

Private mstrRoom As String                                  'ֻ����ִ�м��ڵĲ���
Private mstrPrivs As String, mlngModul As Long
Private mblnPatTrack As Boolean                             '�Ƿ�Խ����˽��и���
Private mblnֱ�Ӽ�� As Boolean                             '�ǼǺ�ֱ�ӽ�����
Private mblnNoShowCancel As Boolean                         '����ʾȡ���ļ��
Private mBeforeDays As Integer                              'Ĭ�ϲ�ѯ������
Private mblnMoved As Boolean                                '��ǰʱ������Ƿ�ת�ƹ�
Private mblnOpenReport As Boolean                           '��ʼ����Զ��򿪱���
Private mblnTechReptSame As Boolean                         'ֻ����д�Լ����ı���
Private mintResultInput As Integer                          '��ʾ���������Ժ�Ӱ������

Private mblnUse3D As Boolean                                '�Ƿ�������ά�ؽ�����
Private mstr3DExeDir As String                              '��ά�ؽ�����·��
Private mstr3DPara As String                                '��ά�ؽ�����
Private mstr3DFunctions As String                           '��ά�ؽ�����

'������������
Private mdatFBegin As Date
Private mdatFEnd As Date
Private mDatType As Integer                                 'ʱ���ѯ��ʽ 1=�����ʱ�䡢2=������ʱ��
Private mstrFNO As String
Private mlngF����ID As Long
Private mstrF��ʶ�� As String
Private mstrF���￨ As String
Private mstrF���� As String
Private mdblFChkNO As Double
Private mstr�걾��λ As String
Private mstr���ҽ�� As String
Private mstr���ҽ�� As String
Private mstr������� As String
Private mbln������� As Boolean
Private mstrӰ������ As String
Private mstr��鼼ʦ As String
Private mstr������ As String
Private mstrӰ����� As String
Private mstr������� As String
Private mstr������ As String
Private mstr���� As String
Private mlngRefreshInterval As Long                         '�����б��Զ�ˢ�¼��
Private Sub InitVslist()
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 38

        .ColWidth(Col����) = 200: .ColWidth(Col��Դ) = 200: .ColWidth(Col����) = 200: .ColWidth(Col����) = 200: .ColWidth(Col����) = 400
        .ColWidth(Col����) = 600: .ColWidth(Col������) = 600: .ColWidth(Col�Ա�) = 400: .ColWidth(Col����) = 400: .ColWidth(Col����) = 800
        .ColWidth(Col��λ) = 800: .ColWidth(Colִ�м�) = 600: .ColWidth(Col���ʱ��) = 1000: .ColWidth(Col����ʱ��) = 1000: .ColWidth(Col����ҽ��) = 600
        .ColWidth(Col���) = 400: .ColWidth(Col����) = 400: .ColWidth(ColӤ��) = 400: .ColWidth(Col�Ǽ���) = 600: .ColWidth(Col������) = 600
        .ColWidth(Col�����) = 600: .ColWidth(Col��ӡ��Ƭ) = 800: .ColWidth(Col�������) = 800: .ColWidth(Col��ɫͨ��) = 800: .ColWidth(Col�����ӡ) = 800
        .ColWidth(Col������) = 600: .ColWidth(Col������) = 600: .ColWidth(Col��鼼ʦ) = 800: .ColWidth(Col��ͼʱ��) = 1000
        
        .ColWidth(ColӰ�����) = 0: .ColWidth(Col����ID) = 0: .ColWidth(Col��ҳID) = 0: .ColWidth(Col�Һŵ�) = 0
        .ColWidth(Colҽ��ID) = 0: .ColWidth(Col���ͺ�) = 0: .ColWidth(Col���UID) = 0: .ColWidth(Col���״̬) = 0: .ColWidth(Colת��) = 0


        .TextMatrix(0, Col����) = 200: .TextMatrix(0, Col��Դ) = 200: .TextMatrix(0, Col����) = 200: .TextMatrix(0, Col����) = 200: .TextMatrix(0, Col����) = 400
        .TextMatrix(0, Col����) = 600: .TextMatrix(0, Col������) = 600: .TextMatrix(0, Col�Ա�) = 400: .TextMatrix(0, Col����) = 400: .TextMatrix(0, Col����) = 800
        .TextMatrix(0, Col��λ) = 800: .TextMatrix(0, Colִ�м�) = 600: .TextMatrix(0, Col���ʱ��) = 1000: .TextMatrix(0, Col����ʱ��) = 1000: .TextMatrix(0, Col����ҽ��) = 600
        .TextMatrix(0, Col���) = 400: .TextMatrix(0, Col����) = 400: .TextMatrix(0, ColӤ��) = 400: .TextMatrix(0, Col�Ǽ���) = 600: .TextMatrix(0, Col������) = 600
        .TextMatrix(0, Col�����) = 600: .TextMatrix(0, Col��ӡ��Ƭ) = 800: .TextMatrix(0, Col�������) = 800: .TextMatrix(0, Col��ɫͨ��) = 800: .TextMatrix(0, Col�����ӡ) = 800
        .TextMatrix(0, Col������) = 600: .TextMatrix(0, Col������) = 600: .TextMatrix(0, Col��鼼ʦ) = 800: .TextMatrix(0, Col��ͼʱ��) = 1000
        
        .TextMatrix(0, ColӰ�����) = 0: .TextMatrix(0, Col����ID) = 0: .TextMatrix(0, Col��ҳID) = 0: .TextMatrix(0, Col�Һŵ�) = 0
        .TextMatrix(0, Colҽ��ID) = 0: .TextMatrix(0, Col���ͺ�) = 0: .TextMatrix(0, Col���UID) = 0: .TextMatrix(0, Col���״̬) = 0: .TextMatrix(0, Colת��) = 0
        
        Dim i As Integer
        For i = 0 To .Cols
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        
        .Editable = flexEDNone
    End With
End Sub
Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim i As Integer
    Dim i3DFunc As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        '.SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    
'�˵�����
'Begin------------------------�ļ��˵�--------------------------------------Ĭ�Ͽɼ�
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)"): cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.IconId = 102
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�嵥��ӡ(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "Ӱ���豸����(&D)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_SendImg, "����ͼ��(&T)"): cbrControl.IconId = 3061
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------���˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "���(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "��ӡ���뵥��(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "���Ǽ�(&I)"): cbrControl.IconId = 211: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CopyCheck, "���ƵǼ�(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "ȡ���Ǽ�(&R)"): cbrControl.IconId = 742
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "�ٻ�ȡ��(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "�޸���Ϣ(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "��鱨��(&L)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ������(&D)"): cbrControl.IconId = 743
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "Ӱ���Ƭ(&S)"): cbrControl.IconId = 8111:  cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "��Ƭ�Ա�(&E)"): cbrControl.IconId = 8112
        
        '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "��ά�ؽ�"): cbrControl.ID = conMenu_Img_3D
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Delete, "Ӱ��ɾ��(&K)"): cbrControl.IconId = 8113
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Query, "Q/R��ȡͼ��(&Q)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "����Ӱ��(&C)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 505: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "ȡ������(&B)"): cbrControl.IconId = 506
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "�����(&X)"):  cbrControl.BeginGroup = True: cbrControl.ID = conMenu_Manage_Result
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "����(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "����(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "Ӱ������(&Y)"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׼�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�Ҽ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_GChannel, "��ɫͨ��(&G)"): cbrControl.ID = conMenu_Manage_GChannel
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelOk, "���(&J)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelCancel, "ȡ��(&Y)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "�ޱ������(&F)"): cbrControl.IconId = 216: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "�ޱ������(&U)"):  cbrControl.IconId = 3012
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "������(&E)"): cbrControl.IconId = 225
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "ȡ�����(&U)"): cbrControl.IconId = 219
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ChangeDevice, "�����豸"): cbrControl.IconId = 3203
    End With
    
    
'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '�����˵�
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_FindType, "���ҷ�ʽ(&G)"): cbrControl.ID = conMenu_View_FindType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "������"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_PatInfor, "������Ϣ(&P)"): cbrControl.IconId = 812
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "���ٹ���(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)")
    End With


'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������", -1, False)
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "WEB�ϵ�����(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------�����------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '��ӡ------------------Ctrl+P
        .Add 0, VK_F12, conMenu_File_Parameter      '��������--------------F12
        
        .Add 0, VK_F2, conMenu_Manage_Regist       '�Ǽ�-----------------F2
        .Add 0, VK_F7, conMenu_Manage_CopyCheck    '���ƵǼ�-------------F7
        .Add 0, VK_F4, conMenu_Manage_Receive       '����-----------------F4
        .Add 0, VK_F9, conMenu_Manage_ClearUp       '���ر���------------F9
        .Add 0, VK_F6, conMenu_Manage_Complete         '��˱���----------F6
        
        
        .Add 0, VK_F1, conMenu_Help_Help              '����-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           'ˢ��-------------F5
        .Add FCONTROL, Asc("F"), conMenu_View_FindType    '���ҷ�ʽ---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '����-------------F3
    End With
    
'---------------------�������Ͻǵ�ǰ����----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "������")
            cbrControl.ID = conMenu_View_Filter * 10#: cbrControl.Flags = xtpFlagRightAlign: cbrControl.Category = "Main"
    
'---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.IconId = 102: cbrControl.ToolTipText = "����Ԥ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ"): cbrControl.IconId = 103: cbrControl.ToolTipText = "�����ӡ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "�Ǽ�"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "����"): cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ��"): cbrControl.IconId = 743: cbrControl.ToolTipText = "ȡ������"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "��Ƭ"): cbrControl.ToolTipText = "Ӱ���Ƭ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "�Ա�"): cbrControl.IconId = 8112: cbrControl.ToolTipText = "��Ƭ�Ա�"
        '���������ά�ؽ����ܣ��򴴽���Ӧ�˵�
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "��ά"): cbrControl.ID = conMenu_Img_3D: cbrControl.ToolTipText = "��ά�ؽ�"
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "�ݻ��ؽ�")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "�����ڿ���")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "�����ؽ�")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "��ע����")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "���"):  cbrControl.BeginGroup = True: cbrControl.ID = conMenu_Manage_Result: cbrControl.IconId = 3506: cbrControl.ToolTipText = "�����������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "����(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "����(&Y)"): cbrPopControl.IconId = 3507
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "����"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061: cbrControl.ToolTipText = "Ӱ������"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "�׼�(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "�Ҽ�(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "���"): cbrControl.IconId = 225: cbrControl.ToolTipText = "����������"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
    End With
End Sub
Private Function InitDepts() As Boolean
'���ܣ���ʼ��סԺ�ٴ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String, str��Դ As String
    
    On Error GoTo errH
    
    str��Դ = "1,2,3"
    If InStr(mstrPrivs, "���п���") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')> 0 And B.�������� IN('���')" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=" & UserInfo.ID & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And instr([1],','||B.�������||',')>0  And B.�������� IN('���')" & _
            " Order by A.����"
    End If
   

    
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, "," & str��Դ & ",")
    
    If rsTmp.EOF Then
        MsgBoxD Me, "û�з���ҽ��������Ϣ,���ȵ����Ź��������á�", vbInformation, gstrSysName
        Exit Function
    Else
        str����IDs = GetUser����IDs
        Do Until rsTmp.EOF
            mstrCanUse���� = mstrCanUse���� & "|" & rsTmp!ID & "_" & rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.����ID Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� '��ȡĬ�Ͽ���
            If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur����ID = 0 Then mlngCur����ID = rsTmp!ID: mstrCur���� = rsTmp!���� & "-" & rsTmp!���� 'û��Ĭ�Ͽ���ȡ��һ������������
            rsTmp.MoveNext
        Loop
        mstrCanUse���� = Mid(mstrCanUse����, 2)
        If InStr(mstrPrivs, "���п���") > 0 And mlngCur����ID = 0 Then
            mlngCur����ID = Split(Split(mstrCanUse����, "|")(0), "_")(0)
            mstrCur���� = Split(Split(mstrCanUse����, "|")(0), "_")(1)
        End If
        
        If mlngCur����ID = 0 And InStr(mstrPrivs, "���п���") <= 0 Then 'û�����п��Ҳ���Ȩ��,���Ҳ����߿��Ҳ����ڼ�������
            MsgBoxD Me, "û�з�������������,����ʹ��ҽ������վ��", vbInformation, gstrSysName
            Exit Function
        End If
        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "����б�"
    Pane1.Handle = picList.Hwnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "�Ӵ���"
'    Pane2.Handle = PicWindow.Hwnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub
Private Sub InitLocalPars()
Dim TitleFont As New StdFont                                '�����б��ͷ����
Dim TextFont As New StdFont                                 '�����б���������
'��ʼ����ʱ���ز������Ը������ã�ע������Ϊ��,������أ��������õȵ���
    On Error GoTo err
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmdסԺ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", 1))
    mblncmd�ѽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", 0))
    mblncmdδ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", 0))
    mblncmd�Ǽ� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", 1))
    mblncmd���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", 1))
    mblncmd��� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", 1))
    mstrCurFindtype = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", "��ʶ��")
    
    mbln���� = (Val(zlDatabase.GetPara("ֻ��ʾ����סԺ��Ŀ", glngSys, mlngModul, "1")) = 1)
    mblnֱ�Ӽ�� = (Val(zlDatabase.GetPara("�Ǽ�ֱ�Ӽ��", glngSys, mlngModul, 0)) = 1)
    mblnOpenReport = (Val(zlDatabase.GetPara("��ʼ����Զ��򿪱���", glngSys, mlngModul, 0)) = 1)
    mblnShowImgAtReport = (Val(zlDatabase.GetPara("����ʱ��Ƭ", glngSys, mlngModul, 0)) = 1)
    mblnNoShowCancel = (Val(zlDatabase.GetPara("����ʾ��ȡ���ĵǼ�", glngSys, mlngModul, 0)) = 1)
    mblnPatTrack = (Val(zlDatabase.GetPara("���˸���", glngSys, mlngModul, 0)) = 1)
    mstrRoom = zlDatabase.GetPara("ִ�м䷶Χ", glngSys, mlngModul, "")
    If mstrRoom <> "" Then mstrRoom = "'," & Replace(mstrRoom, "|", ",") & ",'"
    
    '��ȡ�����ò����б������
    TitleFont.Name = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, "����")
    TitleFont.Size = Val(zlDatabase.GetPara("�����б��ͷ�ֺ�", glngSys, mlngModul, 9))
    TitleFont.Bold = zlDatabase.GetPara("�����б��ͷ����", glngSys, mlngModul, 0) = 1
    TitleFont.Italic = zlDatabase.GetPara("�����б��ͷб��", glngSys, mlngModul, 0) = 1
    
    TextFont.Name = zlDatabase.GetPara("�����б���������", glngSys, mlngModul, "����")
    TextFont.Size = Val(zlDatabase.GetPara("�����б������ֺ�", glngSys, mlngModul, 9))
    TextFont.Bold = zlDatabase.GetPara("�����б����ݴ���", glngSys, mlngModul, 0) = 1
    TextFont.Italic = zlDatabase.GetPara("�����б�����б��", glngSys, mlngModul, 0) = 1
    
    Set rptList.PaintManager.CaptionFont = TitleFont
    Set rptList.PaintManager.TextFont = TextFont
    
    '��ȡ��ά�ؽ�����
    mblnUse3D = Val(zlDatabase.GetPara("������ά�ؽ�", glngSys, mlngModul, 0))
    mstr3DExeDir = zlDatabase.GetPara("3D����·��", glngSys, mlngModul, "")
    mstr3DPara = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")
    mstr3DFunctions = zlDatabase.GetPara("3D����", glngSys, mlngModul, "")

    '����������ʼ
    '-----------------------------------------------------
    mDatType = 1
    mstrFNO = ""
    mlngF����ID = 0
    mstrF��ʶ�� = 0
    mstrF���￨ = ""
    mstrF���� = ""
    mdblFChkNO = 0
    mstr�걾��λ = ""
    mstr���ҽ�� = ""
    mstr���ҽ�� = ""
    mstr������� = ""
    mbln������� = False
    mstrӰ������ = ""
    mstr������� = ""
    mstr������ = ""
    mstr���� = ""
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub InitMvar()
'����:��ʼ��ģ�鼶����,���������ʱ����һ��
    
    mblnIgnoreResult = GetDeptPara(mlngCur����ID, "���Խ��������", "0") = "0"     '���Խ��������
    mintResultInput = GetDeptPara(mlngCur����ID, "��ʾ������", "1") = "1"           '��ʾ������
    mblnFinishCommit = GetDeptPara(mlngCur����ID, "�ޱ�����ɺ�ֱ�����", "0") = "0"      '�ޱ�����ɺ�ֱ�����
    mblnReportWithImage = GetDeptPara(mlngCur����ID, "��ͼ�����д����", "0") = "0"    '��ͼ�����д����
    mblnReportWithResult = GetDeptPara(mlngCur����ID, "��Ӱ�����Ϊ����", "0") = "0"   '��Ӱ�����Ϊ����
    mblnLocalizerBackward = GetDeptPara(mlngCur����ID, "��λƬ����", "0") = "0"  '��λƬ����
    mblnCompleteCommit = GetDeptPara(mlngCur����ID, "��˺�ֱ�����", "0") = "0"      '��˺�ֱ�����
    mBeforeDays = GetDeptPara(mlngCur����ID, "Ĭ�Ϲ�������", "2")                  'Ĭ�Ϲ�������
    mblnTechReptSame = GetDeptPara(mlngCur����ID, "ֻ����д�Լ����ı���", "0") = "0"    'ֻ����д�Լ����ı���
    mblnPacsReport = GetDeptPara(mlngCur����ID, "����༭��", "0") = "0"        '����༭��
    mblnPrintCommit = GetDeptPara(mlngCur����ID, "��ӡ��ֱ�����", "0") = "0"         '��ӡ��ֱ�����
    mlngRefreshInterval = GetDeptPara(mlngCur����ID, "�Զ�ˢ�¼��", "0")         '�Զ�ˢ�¼��
    If mlngRefreshInterval > 65 Then
        mlngRefreshInterval = 30
    End If
    If mlngRefreshInterval <> 0 Then
        TimerRefresh.Interval = mlngRefreshInterval * 1000
        TimerRefresh.Enabled = True
    Else
        TimerRefresh.Enabled = False
    End If
    
    mdatFEnd = CDate(0)
    mdatFBegin = CDate(Format(zlDatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
    mblnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With
    cbrdock.AddImageList img16 '��VB.ImageList��Tag��ID���й���
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("��Դ", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_סԺ, "סԺ")
            objControl.ToolTipText = "��ʾסԺ����"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ���ﲡ��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ��첡��"
        Set objControl = .Add(xtpControlButtonPopup, ID_����, " ��  ��")
            objControl.ToolTipText = "��ʾ�����ѽ�/δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_δ��, "δ��")
            cbrPopControl.ToolTipText = "��ʾ����δ�ɲ���"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_�ѽ�, "�ѽ�")
            cbrPopControl.ToolTipText = "��ʾ�����ѽɲ���"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("״̬", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_�Ǽ�, "�Ǽ�")
            objControl.ToolTipText = "��ʾ�ѵǼǲ���"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ�������"
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.ToolTipText = "��ʾ�ѱ��没��"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����˲���"
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.ToolTipText = "��ʾ����ɲ���"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbsMain.Add("����", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    Set objPopbar = objBar.Controls.Add(xtpControlPopup, ID_���˷�ʽ, "��ʶ��(&D)")
        objPopbar.ID = ID_���˷�ʽ
        objPopbar.Flags = xtpFlagRightAlign
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_����ֵ, "��ʶ��")
        objCusControl.Handle = txtFilter.Hwnd
        objCusControl.Flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_��ʼ����, "����")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_����סԺ, "����")
    objControl.ToolTipText = "ֻ��ʾ����סԺ����¼"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_����
        .Add FCONTROL, vbKey1, ID_סԺ
        .Add FCONTROL, vbKey2, ID_����
        .Add FCONTROL, vbKey3, ID_���
        .Add FCONTROL, vbKey4, ID_����
        .Add FCONTROL, vbKey5, ID_�Ǽ�
        .Add FCONTROL, vbKey6, ID_����
        .Add FCONTROL, vbKey7, ID_����
        .Add FCONTROL, vbKey8, ID_���
        .Add FCONTROL, vbKey9, ID_���
    End With
End Sub
Private Sub Menu_File_Excel_click(ByVal blnNoRecord As Boolean)
Dim bytMode As Byte
    If blnNoRecord Then Exit Sub
    On Error GoTo ErrHandle
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgList
    objPrint.Title.Text = "�����Ŀ�嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_File_BatPrint()
Dim cbrControl As CommandBarControl, strReturn As String, l As Integer
Dim objReportPrint As New zlRichEPR.cDockReport
    Set cbrControl = Me.cbrMain(2).FindControl(, conMenu_File_Print)
    If Not cbrControl Is Nothing Then
        cbrControl.ID = conMenu_File_BatPrint
    Else
        Exit Sub
    End If

    'ѡ����
    strReturn = frmDocPrintPatiList.Showfrm(rptList, Me)
    'ѭ������
    For l = 0 To UBound(Split(strReturn, "|"))
        objReportPrint.zlRefresh CLng(Split(strReturn, "|")(l)), mlngCur����ID
        Call objReportPrint.zlExecuteCommandBars(cbrControl)
        Call AfterPrinted(CLng(Split(strReturn, "|")(l)))
    Next
    cbrControl.ID = conMenu_File_Print
    Unload objReportPrint.zlGetForm
End Sub
Private Sub Menu_RichEPR(ByVal cbrID As Long)
    Dim cbrControl As CommandBarControl, i As Integer
    
    '����ҳ�治�ɼ�ʱ��ִ���κβ���
    If TabWindow.Selected.Tag <> "������д" Then
        For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
            If TabWindow(i).Tag = "������д" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.Tag <> "������д" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    'ˢ��Ƕ��ҳ������
    If mblnPacsReport = True Then
        Call mfrmPacsReport.zlRefresh(Nvl(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, 0), Nvl(rptList.FocusedRow.Record(mcol.���ͺ�).Value, 0), mlngCur����ID, mstrPrivs, mlngModul, Me, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
    Else
        Call mobjReport.zlRefresh(Nvl(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, 0), mlngCur����ID, True)
    End If
    
    '�жϰ���������
    Set cbrControl = Me.cbrMain.FindControl(, IIf(mblnPacsReport, conMenu_PacsReport_Open, cbrID))
    If cbrControl Is Nothing Then Exit Sub
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    Call cbrMain_Execute(cbrControl)
End Sub
Private Sub Menu_File_Parmeter_click()
    With frmTechnicSetup
        .mlngModul = mlngModul
        .mlng����ID = mlngCur����ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        If .mblnOK Then
            InitLocalPars
            Call RefreshRptlist
        End If
    End With
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '���ܣ����ð�������
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.Hwnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo Hwnd
End Sub

Private Sub Menu_Manage_ȡ������(ByVal intState As Integer)
'ȡ��������������ǣ�ÿ��ȡ��������ͼ��ȫ���������б���ɢ��N����ʱ��¼
Dim strFilter As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    '��ʾ����ѡ�񴰿�
    With rptList.FocusedRow
        gstrSQL = "select 0 as ѡ��,B.����UID as ID ,B.���к�,B.��������,SUM(1) AS ͼ���� from Ӱ�����¼ A ," & _
                "Ӱ�������� B, Ӱ����ͼ�� C Where a.���UID = B.���UID And B.����UID = C.����UID" & _
                " And a.ҽ��ID = [1] and A.���ͺ�= [2] group by B.����UID,B.���к�,B.��������"
        Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.ҽ��ID).Value), CLng(.Record(mcol.���ͺ�).Value))
        
        frmSelectMuli.ShowSelect rsTmp, "ID,3000,0,1;���к�,800,0,1;��������,2000,0,1;ͼ����,800,0,1", 0, 0, 14000, 10000, "ȡ������"
        
        If frmSelectMuli.mblnOK = True Then
            strFilter = frmSelectMuli.strFilter
            rsTmp.Filter = strFilter
            '�����ѡ�����У�����ÿһ�����е�ȡ��
            While Not rsTmp.EOF
                subCancelSeriesRelate CLng(.Record(mcol.ҽ��ID).Value), CLng(.Record(mcol.���ͺ�).Value), rsTmp!ID
                rsTmp.MoveNext
            Wend
            
            '����Ӱ����״̬�������ǰҽ���Ѿ�û��ͼ�񣬶��Ҽ�����Ϊ3�����޸�Ϊ2
            If intState = 3 Then
                gstrSQL = "Select ���uid From Ӱ�����¼ Where  ҽ��ID=[1] And ���ͺ�=[2]"
                Set rsTmp = OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.ҽ��ID).Value), CLng(.Record(mcol.���ͺ�).Value))
                If IsNull(rsTmp!���UID) Then
                    gstrSQL = "Zl_Ӱ����_State(" & CLng(.Record(mcol.ҽ��ID).Value) & "," & CLng(.Record(mcol.���ͺ�).Value) & ",2)"
                    zlDatabase.ExecuteProcedure gstrSQL, "ȡ������"
                End If
            End If
            
            mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
            Call RefreshRptlist '����ȡ��������ȷ����ˢ��
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�ޱ������()
Dim blnTran As Boolean, arrSQL() As Variant, l As Long
'ֻ�н����еı�����Բ����ò˵�,��Ϊ��ʱ��û��ǩ��
        On Error GoTo ErrHandle
        arrSQL = Array()
        With rptList.FocusedRow
            If .Record(mcol.����ID).Value <> 0 Then
                If MsgBoxD(Me, "�Ƿ��ޱ���ֱ�����,ֱ����ɽ�ɾ������д�ı���!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            If mblnFinishCommit And InStr(mstrPrivs, "������") > 0 Then '�ޱ�����ɺ������ٴ�ȷ�����,����Ҫ�м����ɵ�Ȩ��
                '�˹���,��״̬=6,���ұ���ID��Ϊ�ս�ɾ�����Ӳ�����¼
                If zlDatabase.GetPara(81, glngSys) = 1 And Not bln������Ժ(Nvl(.Record(mcol.����ID).Value), Nvl(.Record(mcol.��ҳID).Value)) And bln����δ�󻮼۵�(Nvl(.Record(mcol.ҽ��ID).Value)) Then 'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
                    MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵�������ɣ�", vbExclamation, gstrSysName
                Else
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_Ӱ����_STATE(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & ",6" & IIf(.Record(mcol.����ID).Value <> 0, "," & .Record(mcol.����ID).Value, "") & ")"
                End If
            Else
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_Ӱ����_STATE(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & ",5" & _
                            IIf(.Record(mcol.����ID).Value <> 0, "," & .Record(mcol.����ID).Value, "") & ")"
            End If
        End With
        
        gcnOracle.BeginTrans '--------------------------д������
        blnTran = True
        For l = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "д�����Ĳ�������")
        Next
        gcnOracle.CommitTrans
        
        If mblnPatTrack Then
            If mblnFinishCommit Then
                Call StateCheck(6)
            Else
                Call StateCheck(5)
            End If
        Else
            Call RefreshRptlist
        End If
        Exit Sub
ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Edit_�ޱ������()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If MsgBoxD(Me, "ȷ��Ҫ���˸�������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    With rptList.FocusedRow
            
            '�����ͼ������˵����Ѽ�顱��������˵����ѱ�����
            strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ���ͼ��", CLng(.Record(mcol.ҽ��ID).Value))
            
            gstrSQL = "ZL_Ӱ����_STATE(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & "," & IIf(IsNull(rsTemp!���UID) = True, 2, 3) & ")"
            ExecuteProc gstrSQL, Me.Caption
    End With
    If mblnPatTrack Then
        Call StateCheck(2)
    Else
        Call RefreshRptlist
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_����������(lngҽ��ID As Long, Optional blnRefresh As Boolean = True)
    Dim arrSQL() As Variant, l As Long, blnTran As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If InStr(mstrPrivs, "������") <= 0 Then Exit Sub
    
    strSQL = "Select a.���ͺ�,b.����ID,b.��ҳID From ����ҽ������ a,����ҽ����¼ b Where a.ҽ��id = [1] And a.ҽ��ID=b.Id"
    Set rsTemp = OpenSQLRecord(strSQL, "����������", lngҽ��ID)
    
    If rsTemp.EOF = True Then Exit Sub
    
    arrSQL = Array()
    If zlDatabase.GetPara(81, glngSys) = 1 And Not bln������Ժ(Nvl(rsTemp!����ID), Nvl(rsTemp!��ҳID, 0)) And bln����δ�󻮼۵�(Nvl(lngҽ��ID)) Then 'ִ�к��Զ���˻��۵���Ч�����Ҳ����ѳ�Ժ������δ��˵Ļ��۵�
        MsgBoxD Me, "�ò����ѳ�Ժ������δ��˵Ļ��۵���������ɣ�", vbExclamation, gstrSysName
    Else
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_Ӱ����_STATE(" & lngҽ��ID & "," & rsTemp!���ͺ� & ",6)"
    End If

    gcnOracle.BeginTrans '--------------------------д������
    blnTran = True
    For l = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(l)), "д�����Ĳ�������")
    Next
    gcnOracle.CommitTrans

    If blnRefresh Then Call StateCheck(6)
    Exit Sub

ErrHandle:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_ȡ��������()
    On Error GoTo ErrHandle
    With rptList.FocusedRow
            If .Record(mcol.ת��).Value = 1 Then MsgBox "�ò��˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������", vbInformation, gstrSysName: Exit Sub
            gstrSQL = "ZL_Ӱ����_STATE(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & ",5)"
            ExecuteProc gstrSQL, "ȡ��������"
    End With

    Call StateCheck(5)
    Exit Sub

ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�������(ByVal lngID As Long)
    Dim iresult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_Negative
            iresult = 1
        Case conMenu_Manage_Positive
            iresult = 0
    End Select
    With rptList.FocusedRow
        gstrSQL = "ZL_Ӱ����_���(" & .Record(mcol.ҽ��ID).Value & "," & iresult & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���������")
        .Record(mcol.����).Value = IIf(iresult = 1, "����", "")
        .Record(mcol.����).Icon = IIf(iresult = 1, Me.imgList.ListImages("����").Index - 1, -1)
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_��ɫͨ��(ByVal lngID As Long)
    Dim intResult As Integer

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    With rptList.FocusedRow
        gstrSQL = "Zl_��ɫͨ��_Update(" & .Record(mcol.ҽ��ID).Value & ",'" & intResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "Ӱ������")
        .Record(mcol.��ɫͨ��).Value = intResult
        .Record(mcol.����).Icon = IIf(intResult = 1, Me.imgList.ListImages("��ɫͨ��").Index - 1, -1)
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_Ӱ������(ByVal lngID As Long)
    Dim strResult As String

    On Error GoTo ErrHandle
    Select Case lngID
        Case conMenu_Manage_First
            strResult = "��"
        Case conMenu_Manage_Second
            strResult = "��"
    End Select
    With rptList.FocusedRow
        gstrSQL = "Zl_Ӱ������_Update(" & .Record(mcol.ҽ��ID).Value & ",'" & strResult & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "Ӱ������")
        .Record(mcol.����).Value = strResult
        .Selected = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�޸�(ByVal blnNoRecord As Boolean, ByVal intState As Integer)
    If blnNoRecord Then Exit Sub
    
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = rptList.FocusedRow.Record(mcol.���ͺ�).Value
        .mlngAdviceID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value
        .mintEditMode = IIf(intState > 1, 3, 1) '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        .RefreshPatiInfor False 'ˢ�²���
        .mblnOK = False
        .Show 1, Me
        If .mblnOK Then RefreshRptlist '�ɹ�����
    End With
End Sub
Private Sub Menu_Manage_���ƵǼ�()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        .CopyCheck rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value 'ˢ�²���
        .Show 1, Me
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2)
            Else
                Call RefreshRptlist
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_�Ǽ�()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .mblnOK = False
        .InitMvar
        .Show 1, Me
        If .mblnOK Then '�ɹ�����
            If mblnֱ�Ӽ�� Then
                Call StateCheck(2)
            Else
                Call RefreshRptlist
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_ȡ���Ǽ�()
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        If MsgBoxD(Me, "ȷ��Ҫȡ����ǰ������" & Chr(10) & Chr(13) & "����ȡ�������Ӧ��ҽ�����ܾ�ִ�У�", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "ZL_����ҽ��ִ��_�ܾ�ִ��(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����Ǽ�")
        Call RefreshRptlist
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_�ٻ�ȡ��()
'���ܣ��ٻر�ȡ���ĵǼ�
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    
    On Error GoTo errH
    
    With rptList.SelectedRows(0)
        If MsgBoxD(Me, "ȷʵҪ�ٻر�ȡ���Ǽǵ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        lngҽ��ID = .Record(mcol.ҽ��ID).Value
        lng���ͺ� = .Record(mcol.���ͺ�).Value
    End With
    
    gstrSQL = "ZL_����ҽ��ִ��_ȡ���ܾ�(" & lngҽ��ID & "," & lng���ͺ� & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call RefreshRptlist
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Menu_Manage_����()
Dim i As Long, cbrControl As CommandBarControl, blnFocusFind As Boolean

    blnFocusFind = (Me.ActiveControl.Name = "Txt��ʶ��")
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = rptList.FocusedRow.Record(mcol.���ͺ�).Value
        .mlngAdviceID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value
        .mintEditMode = 2 '0���Ǽǡ�1���ǼǺ��޸ġ�2��������3���������޸�
        .mlngCurDeptId = mlngCur����ID
        .InitMvar
        .RefreshPatiInfor True 'ˢ�²���
        .mblnOK = False
        .Show 1, Me
        If .mblnOK Then  '�ɹ�����
            Call StateCheck(2)
            If mblnOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '��ʼ����Զ��򿪱���
        End If
        If blnFocusFind Then Txt��ʶ��.SetFocus '�Զ���λ����λ��
    End With
End Sub
Private Sub Menu_Manage_ȡ������(ByVal intState As Integer)
Dim rsTemp As ADODB.Recordset, lngcurҽ��ID As Long
    If intState <= 1 Then Call Menu_Manage_ȡ���Ǽ�: Exit Sub '����������
    
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        '------------------------------------��ǩ������Ҫ�Ȼ���ǩ�����ٳ���
        lngcurҽ��ID = .Record(mcol.ҽ��ID).Value
        gstrSQL = "Select Distinct B.���ʱ�� From ����ҽ������ A, ���Ӳ�����¼ B Where A.����ID=B.Id And A.ҽ��ID=[1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�Ƿ�ǩ��", CLng(.Record(mcol.ҽ��ID).Value))
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!���ʱ��, "") <> "" Then 'ǩ������
                MsgBoxD Me, "��ǰ���˵ļ�鱨���Ѿ�ǩ��,����ȡ�����,���Ȼ���ǩ��!", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If

        If MsgBoxD(Me, "ȡ�����μ�齫ɾ����Ӧ�ļ��ͼ��ͼ�鱨�棬�Ƿ������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        If .Record(mcol.���UID).Value <> "" And InStr(mstrPrivs, "���ͼ��") <= 0 Then
            MsgBoxD Me, "��û��������ͼ��Ȩ��,�������ͼ��,���в���ȡ��������!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        gstrSQL = "ZL_Ӱ����_CANCEL(" & .Record(mcol.ҽ��ID).Value & "," & .Record(mcol.���ͺ�).Value & "," & Nvl(.Record(mcol.����ID).Value, 0) & ")"
        ExecuteProc gstrSQL, Me.Caption
        'ɾ��Ӱ���ļ���Ŀ¼
        RemoveCheckImages .Record(mcol.ҽ��ID).Value, .Record(mcol.���ͺ�).Value

    End With
    
    Call StateCheck(1)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_��Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("Ӱ����")
End Sub
Private Sub Menu_Manage_�Աȹ�Ƭ()
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("Ӱ��Ա�")
End Sub
            
Private Sub Menu_Manage_ͼ��ɾ��()
Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
    End If
    
    gstrSQL = "select ���UID from Ӱ�����¼ where ҽ��ID =[1] and  ���ͺ� = [2]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���UID", CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value), CLng(rptList.FocusedRow.Record(mcol.���ͺ�).Value))
    If rsTemp.EOF Then Exit Sub
    
    If MsgBoxD(Me, "�Ƿ�ȷ��Ҫɾ���ü�������Ӱ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    'ɾ��Ӱ���ļ���Ŀ¼
    RemoveCheckImages CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value), CLng(rptList.FocusedRow.Record(mcol.���ͺ�).Value)
    gstrSQL = "ZL_Ӱ����_PhotoDelete(" & CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.���ͺ�).Value) & ")"
    ExecuteProc gstrSQL, Me.Caption
    
    '����Ӱ����״̬�����������Ϊ3�����޸�Ϊ2
    If Val(Mid(rptList.FocusedRow.Record(mcol.���״̬).Value, 1, 1)) = 3 Then
        gstrSQL = "Zl_Ӱ����_State(" & CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.���ͺ�).Value) & ",2)"
        zlDatabase.ExecuteProcedure gstrSQL, "ɾ��ͼ��"
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
        
Private Sub Menu_Manage_��ȡͼ��()
Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
    End If
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
    
    'û��Ĭ���豸ʱ����
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "û������Ĭ��Ӱ�����豸���Ƿ��������ã�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmPACSImageDeviceSetup", "Ĭ��Ӱ���豸", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select �豸��,�豸��, IP��ַ,�˿ں�,����AE,�豸AE from Ӱ���豸Ŀ¼ where �豸�� = [1] "
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, Mid(strImageDeviceNumber, 2))
    
    '��Ĭ���豸��ɾ������������
    If rsTemp.EOF = True Then
        MsgBoxD Me, "Ĭ���豸�ѱ�ɾ�������������ã�", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP��ַ"), rsTemp("�˿ں�"), rsTemp("�豸��"), Nvl(rsTemp("����AE")), Nvl(rsTemp("�豸AE")), rptList.FocusedRow.Record(mcol.ҽ��ID).Value
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_����Ӱ��(ByVal intState As Integer)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With rptList.FocusedRow
        Call funRelateSeries(CLng(.Record(mcol.ҽ��ID).Value), CLng(.Record(mcol.���ͺ�).Value))
    End With
    
    '����Ӱ����״̬�����ԭ����״̬���ѱ��������޸ĳ��Ѽ�飬
    If intState < 3 Then
        '��������Ѿ���ͼ�����޸ĳ��Ѽ��
        strSQL = "Select ���UID From Ӱ�����¼ Where ҽ��ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ���ͼ��", CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value))
        
        If Not IsNull(rsTemp!���UID) Then
            gstrSQL = "Zl_Ӱ����_State(" & CLng(rptList.FocusedRow.Record(mcol.ҽ��ID).Value) & "," & CLng(rptList.FocusedRow.Record(mcol.���ͺ�).Value) & ",3)"
            zlDatabase.ExecuteProcedure gstrSQL, "����Ӱ��"
        End If
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshRptlist
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_View_Find_click()
    Txt��ʶ��.SetFocus
End Sub
Private Sub Menu_View_Find_Type_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    mstrCurFindtype = Split(control.Caption, "(")(0)
    cbrMain.RecalcLayout
    If mstrCurFindtype = "�ɣÿ�" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Else
            Txt��ʶ��.Text = mobjICCard.Read_Card(Me)
        End If
    End If
    Txt��ʶ��.SetFocus
End Sub
Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    If mlngCur����ID <> control.DescriptionText Then
        mlngCur����ID = control.DescriptionText
        mstrCur���� = Split(control.Caption, "(")(0)
        Call cbrMain.RecalcLayout
        Call InitMvar
        Call InitSubForm
        Call RefreshRptlist
    End If
End Sub
Private Sub Menu_View_������Ϣ(ByVal blnNoRecord As Boolean, ByVal intState As Integer)
    If blnNoRecord Then Exit Sub
    Call frmDegreeCard.ShowInfo(Me, rptList.FocusedRow.Record(mcol.����ID).Value)
End Sub
Private Sub Menu_View_Refresh_click()
   Call RefreshRptlist
End Sub

'
Private Sub Menu_Help_Web_Home_click()
    zlHomePage Hwnd
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub cboTimes_Click()
Dim lngҽ��ID As Long, strҽ������ As String, rsTemp As ADODB.Recordset, i As Integer, strSQLBak As String, rsAnnex As ADODB.Recordset
    If cboTimes.ListCount <= 1 Then Exit Sub
    
    If lbl������Ϣ.Caption = "" Then Exit Sub '��ʱ��û������Ϣ��ɣ���listindex��ֵ����
    On Error GoTo ErrHandle
    lngҽ��ID = cboTimes.ItemData(cboTimes.ListIndex)
    
    If lngҽ��ID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value Then '�����뵱ǰѡ��ҽ��ID��ͬʱ���ɱ���������
        Call rptList_SelectionChanged
        mblnIsHistory = False
        Exit Sub
    End If
    mblnIsHistory = True
    
    '��ȡ��������ҽ���������Ϣ
    gstrSQL = "Select /*+ Rule */ a.������Դ,A.����ID,a.Id ҽ��id,a.��ҳid ,a.���˿���id,a.�Һŵ�,a.ҽ������,c.���UID, " & _
                "b.���ͺ�,b.ִ��״̬,b.ִ�й���,c.���,c.����,c.����,d.��ǰ����id ����id ,d.����,d.�Ա�,d.����,0 as ת��," & _
                "Decode(a.������Դ, 1, d.�����, 2, d.סԺ��, 4, d.�����, Null) As ��ʶ��,d.��ǰ����,a.����ҽ��,f.���� as ���˿��� " & _
                " From ����ҽ����¼ a, ����ҽ������ b, Ӱ�����¼ c, ������Ϣ d, Ӱ������Ŀ e,���ű� f" & _
                " Where a.Id = [1] And a.���id Is Null " & _
                    " And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ��id(+) And b.���ͺ� = c.���ͺ�(+)" & _
                    " And a.����id = d.����id And a.������Ŀid = e.������Ŀid  And f.id = a.���˿���ID "
    strSQLBak = gstrSQL
    strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
    strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
    strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
    strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
    gstrSQL = gstrSQL & " Union ALL " & strSQLBak
    
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ���μ�¼", lngҽ��ID)
    If rsTemp.EOF Then
        Select Case TabWindow(TabWindow.Selected.Index).Tag
            Case "Ӱ��ͼ��"
                mfrmPACSImg.zlRefresh 0, 0, mstrPrivs, False
            Case "������д"
                If mblnPacsReport = True Then
                    mfrmPacsReport.zlRefresh 0, 0, 0, mstrPrivs, mlngModul, Me, False
                Else
                    mobjReport.zlRefresh 0, mlngCur����ID, False
                End If
            Case "�������"
                mobjExpense.zlRefresh mlngCur����ID, 0, 0, False
            Case "סԺҽ��"
                mobjInAdvice.zlRefresh 0, 0, 0, 0, 0, False, 0, 0
            Case "����ҽ��"
                mobjOutAdvice.zlRefresh 0, "", False, False, 0
            Case "סԺ����"
                mobjInEPRs.zlRefresh 0, 0, mlngCur����ID, False
            Case "���ﲡ��"
                mobjOutEPRs.zlRefresh 0, 0, mlngCur����ID, False
        End Select
        Txt������Ϣ = ""
        lbl������Ϣ.Caption = "��  ��:" & Space(12) & "��  ��:" & Space(13) & "��  ��:" & Space(10) & "��ʶ��:" & Space(12) & "��  ��:" & Space(10)
        lbl�����Ϣ.Caption = "����:" & Space(12) & "���˿���:" & Space(11) & "����ҽ��:" & Space(8) & "�����Ŀ:"
        lblCash.Visible = False
        Exit Sub
    End If
    
    Txt������Ϣ = ""
    If InStr(Nvl(rsTemp!ҽ������), ":") > 0 Then
        For i = 0 To UBound(Split(Split(rsTemp!ҽ������, ":")(1), "),"))
            If i = 0 Then
                Txt������Ϣ = "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(Split(rsTemp!ҽ������, ":")(1), "),")(i) & ")"
            Else
                Txt������Ϣ = Txt������Ϣ & vbCrLf & Space(2) & i + 1 & ":" & Split(Split(rsTemp!ҽ������, ":")(1), "),")(i) & ")"
            End If
        Next
        If Trim(Txt������Ϣ) <> "" Then Txt������Ϣ = Mid(Txt������Ϣ, 1, Len(Txt������Ϣ) - 1)
    Else
        Txt������Ϣ = Txt������Ϣ & "��鲿λ:" & Nvl(rsTemp!ҽ������)
    End If
    
    '��ʾ������Ϣ
    '��ȡ�������β���ҽ������
    gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
    If rsTemp!ת�� = 1 Then gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
    Set rsAnnex = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҽ������", lngҽ��ID)
    Do Until rsAnnex.EOF
        strҽ������ = strҽ������ & rsAnnex!��Ŀ & ":" & Nvl(rsAnnex!����) & vbCrLf
        rsAnnex.MoveNext
    Loop
    Txt������Ϣ = Txt������Ϣ & vbCrLf & vbCrLf & strҽ������
    lbl������Ϣ.Caption = "��  ��:" & Rpad(Nvl(rsTemp!����), 12, " ") & "��  ��:" & Rpad(Nvl(rsTemp!�Ա�), 13, " ") & _
                                  "��  ��:" & Rpad(Nvl(rsTemp!����), 10, " ") & "��ʶ��:" & Rpad(Nvl(rsTemp!��ʶ��), 12, " ") & _
                                  "��  ��:" & Rpad(Nvl(rsTemp!��ǰ����), 10, " ")
    lbl�����Ϣ.Caption = "����:" & Rpad(Nvl(rsTemp!����), 12, " ") & "���˿���:" & Rpad(Nvl(rsTemp!���˿���), 11, " ") & _
                                  "����ҽ��:" & Rpad(Nvl(rsTemp!����ҽ��), 8, " ") & "�����Ŀ:" & Split(rsTemp!ҽ������, ":")(0)
                                  
    lblCash.Caption = "��": lblCash.Visible = True
    
    mlngHOrderID = lngҽ��ID
    mlngHSendNo = Nvl(rsTemp!���ͺ�, 0)
    mstrHStudyUID = Nvl(rsTemp!���UID)
    mblnHMoved = IIf(rsTemp!ת�� = 1, True, False)
    
    If Nvl(rsTemp!������Դ, 3) <> 3 Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = False
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = True
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д" '�ѵǼ�״̬���ܲ鿴����ҳ
                    TabWindow(i).Visible = Nvl(rsTemp!ִ�й���, 0) > 1
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "���ﲡ��", "����ҽ��"
                    TabWindow(i).Visible = True
                Case "סԺ����", "סԺҽ��"
                    TabWindow(i).Visible = False
                Case "Ӱ��ͼ��"
                    TabWindow(i).Visible = True
                Case "������д"
                    TabWindow(i).Visible = Nvl(rsTemp!ִ�й���, 0) > 1
            End Select
        Next
    End If
    'ˢ�µ�ǰҳ��Ϣ
    Select Case TabWindow(TabWindow.Selected.Index).Tag
        Case "Ӱ��ͼ��"
            mfrmPACSImg.zlRefresh lngҽ��ID, Nvl(rsTemp!���ͺ�, 0), mstrPrivs, rsTemp!ת�� = 1
        Case "������д"
            If mblnPacsReport = True Then
                mfrmPacsReport.zlRefresh lngҽ��ID, Nvl(rsTemp!���ͺ�, 0), mlngCur����ID, mstrPrivs, mlngModul, Me, rsTemp!ת�� = 1
            Else
                mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False
            End If
        Case "�������"
            mobjExpense.zlRefresh mlngCur����ID, lngҽ��ID, Nvl(rsTemp!���ͺ�, 0), rsTemp!ת�� = 1
        Case "סԺҽ��"
            mobjInAdvice.zlRefresh Nvl(rsTemp!����ID, 0), Nvl(rsTemp!��ҳID, 0), Nvl(rsTemp!����ID, 0), Nvl(rsTemp!���˿���ID, 0), 0, rsTemp!ת�� = 1, lngҽ��ID, Nvl(rsTemp!ִ��״̬, 1)
        Case "����ҽ��"
            mobjOutAdvice.zlRefresh Nvl(rsTemp!����ID, 0), Nvl(rsTemp!�Һŵ�, ""), False, rsTemp!ת�� = 1, lngҽ��ID
        Case "סԺ����"
            mobjInEPRs.zlRefresh Nvl(rsTemp!����ID, 0), Nvl(rsTemp!��ҳID, 0), mlngCur����ID, False, rsTemp!ת�� = 1
        Case "���ﲡ��"
            mobjOutEPRs.zlRefresh Nvl(rsTemp!����ID, 0), Nvl(rsTemp!��ҳID, 0), mlngCur����ID, False, rsTemp!ת�� = 1
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.Hwnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
Select Case control.ID
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_סԺ
            mblncmdסԺ = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_�ѽ�
            mblncmd�ѽ� = Not control.Checked
            If mblncmd�ѽ� Then mblncmdδ�� = False
        Case ID_δ��
            mblncmdδ�� = Not control.Checked
            If mblncmdδ�� Then mblncmd�ѽ� = False
        Case ID_�Ǽ�
            mblncmd�Ǽ� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_����
            mblncmd���� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
        Case ID_���
            mblncmd��� = Not control.Checked
    End Select
cbrdock.RecalcLayout
Call RefreshRptlist
End Sub

Private Sub cbrdock_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    If CommandBar.Parent Is Nothing Then Exit Sub
    If CommandBar.Parent.ID = ID_���˷�ʽ Then
        With CommandBar.Controls
            If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 0, "��ʶ��(&1)"): objControl.Checked = True
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 1, "���￨(&2)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 2, "����(&3)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 3, "���ݺ�(&4)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 4, "����(&5)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 5, "���֤(&6)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 6, "�ɣÿ�(&7)")
                Set objControl = .Add(xtpControlButton, ID_���˷�ʽ * 100# + 7, "�����(&8)")
            End If
        End With
    End If

End Sub

Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    rptList.Top = lngTop
    rptList.Width = picList.Width
    rptList.Height = picList.Height - lngTop - Txt������Ϣ.Height - 100

    Txt������Ϣ.Top = rptList.Top + rptList.Height + 100
    Txt������Ϣ.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_סԺ
            control.Checked = mblncmdסԺ
            control.IconId = IIf(mblncmdסԺ, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd�ѽ� Xor mblncmdδ��
            control.Caption = IIf(mblncmd�ѽ� Xor mblncmdδ��, IIf(mblncmd�ѽ�, " �ѽɷ�", " δ�ɷ�"), " ��  ��")
        Case ID_�ѽ�
            control.Checked = mblncmd�ѽ�
            control.IconId = IIf(mblncmd�ѽ�, 90001, 90000)
        Case ID_δ��
            control.Checked = mblncmdδ��
            control.IconId = IIf(mblncmdδ��, 90001, 90000)
        Case ID_�Ǽ�
            control.Checked = mblncmd�Ǽ�
            control.IconId = IIf(mblncmd�Ǽ�, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_����
            control.Checked = mblncmd����
            control.IconId = IIf(mblncmd����, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
        Case ID_���
            control.Checked = mblncmd���
            control.IconId = IIf(mblncmd���, 90001, 90000)
    End Select
End Sub
Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnNoRecord As Boolean  '�Ƿ��е�ǰ��¼
    Dim intState As Integer
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    If Not blnNoRecord Then
        intState = Val(Mid(rptList.FocusedRow.Record(mcol.���״̬).Value, 1, 1))
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------�ļ�------------------
        Case conMenu_File_PrintSet '��ӡ����
            Call zlPrintSet
           
        Case conMenu_File_Excel '�嵥��ӡ
            Call Menu_File_Excel_click(blnNoRecord)
            
        Case conMenu_File_BatPrint '������ӡ
            Call Menu_File_BatPrint
            
        Case conMenu_File_Parameter '��������
            Call Menu_File_Parmeter_click
            
        Case conMenu_Cap_DevSet 'Ӱ���豸����
            frmPACSImageDeviceSetup.Show vbModal, Me
            
        Case conMenu_File_SendImg '����ͼ��
            frmPacsSendImage.ShowMe Me
            
        Case conMenu_File_Exit '�˳�
            Unload Me
            
'---------------------------���-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '��ӡ���Ƶ���
            Call FuncBillPrint(control)
            
        Case conMenu_Manage_Regist                          '�Ǽ�
            Call Menu_Manage_�Ǽ�
            
        Case conMenu_Manage_CopyCheck                       '���ƵǼ�
            Call Menu_Manage_���ƵǼ�
            
        Case conMenu_Manage_Receive                         '����
            Call Menu_Manage_����
            
        Case conMenu_Manage_Redo                            'ȡ���Ǽ�
            Call Menu_Manage_ȡ���Ǽ�
            
        Case conMenu_Manage_ReGet                           '�ٻ�ȡ��
            Call Menu_Manage_�ٻ�ȡ��
        
        Case conMenu_Manage_ThingModi                       '�޸ĵǼ�
            Call Menu_Manage_�޸�(blnNoRecord, intState)
            
        Case conMenu_Manage_Logout                          'ȡ������
            Call Menu_Manage_ȡ������(intState)
            
        Case conMenu_Img_Look                         '��Ƭ
            Call Menu_Manage_��Ƭ
        
        Case conMenu_Img_Contrast                     '�Աȹ�Ƭ
            Call Menu_Manage_�Աȹ�Ƭ
        
        Case conMenu_Img_3D_MMPR                    '��ά�ؽ���MMPR
            Call sub��ά�ؽ�("MMPR")
        Case conMenu_Img_3D_MPR                     '��ά�ؽ���MPR
            Call sub��ά�ؽ�("MPR")
        Case conMenu_Img_3D_PF                     '��ά�ؽ�,��ע����
            Call sub��ά�ؽ�("PF")
        Case conMenu_Img_3D_SA                     '��ά�ؽ��������ؽ�
            Call sub��ά�ؽ�("SA")
        Case conMenu_Img_3D_VA                     '��ά�ؽ����ݻ��ؽ�
            Call sub��ά�ؽ�("VA")
        Case conMenu_Img_3D_VE                     '��ά�ؽ��������ڿ���
            Call sub��ά�ؽ�("VE")
            
        Case conMenu_Img_Delete                       'ͼ��ɾ��
            Call Menu_Manage_ͼ��ɾ��
        
        Case conMenu_Img_Query                        '���豸��ȡͼ��
            Call Menu_Manage_��ȡͼ��
        
        Case conMenu_Manage_Transfer                        '����Ӱ��
            Call Menu_Manage_����Ӱ��(intState)
            
        Case conMenu_Manage_Cancel                          'ȡ������
            Call Menu_Manage_ȡ������(intState)
        
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '���������
            Call Menu_Manage_�������(control.ID)
            
        Case conMenu_Manage_First, conMenu_Manage_Second
            Call Menu_Manage_Ӱ������(control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_��ɫͨ��(control.ID)
            
        Case conMenu_Manage_ClearUp                           '�ޱ������
            Call Menu_Edit_�ޱ������
                    
        Case conMenu_Manage_Finish                          '�ޱ���ֱ�����
            Call Menu_Manage_�ޱ������
            
        Case conMenu_Manage_Complete                        '������
            If Not rptList.FocusedRow Is Nothing Then
                Call Menu_Manage_����������(rptList.FocusedRow.Record(mcol.ҽ��ID).Value)
            End If
        
        Case conMenu_Manage_Undone                          'ȡ��������
            Call Menu_Manage_ȡ��������
            
        Case conMenu_Manage_ChangeDevice                    '��������豸
            Call Menu_Manage_��������豸
            
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_FindType * 10# To conMenu_View_FindType * 10# + 6
            Call Menu_View_Find_Type_click(control)
        Case conMenu_View_PatInfor
            Call Menu_View_������Ϣ(blnNoRecord, intState)
        Case conMenu_View_Filter '����
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh 'ˢ��
            Call Menu_View_Refresh_click
            
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|"))
            Call Menu_Dept_Select(control)
        Case Else
            If Between(control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And control.Parameter <> "" Then
                 'ִ�з�������ǰģ��ı���
                If Not blnNoRecord Then
                    Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, _
                        "NO=" & rptList.FocusedRow.Record(mcol.NO).Value, "����=" & rptList.FocusedRow.Record(mcol.��¼����).Value, _
                        "ҽ��id=" & rptList.FocusedRow.Record(mcol.ҽ��ID).Value, 1)
                Else
                    Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, "", 1)
                End If
            Else
                If Not blnNoRecord Then
                    Select Case TabWindow.Selected.Tag
                        Case "������д"
                            'û���治�ܴ�ӡ��Ԥ��
                            If Nvl(rptList.FocusedRow.Record(mcol.����ID).Value, 0) = 0 And (control.ID = conMenu_File_Preview Or control.ID = conMenu_File_Print) Then
                                MsgBoxD Me, "��ǰ����û�м�鱨�棬���ܲ��������飡", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '���汻ĳ�˴򿪺��ٱ��������˱༭���޶�
                            If control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Modify Or control.ID = conMenu_PacsReport_Open Or control.ID = conMenu_Edit_Delete Then
                                If CheckConcurrentReport(rptList.FocusedRow.Record(mcol.ҽ��ID).Value) = False Then Exit Sub
                            End If
                            
                            '���� ֻ����д�Լ����ı���,'��������д���޶���ɾ��
                            If mblnTechReptSame = True _
                                And (control.ID = conMenu_Edit_Modify Or control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Delete) _
                                And Nvl(rptList.FocusedRow.Record(mcol.��鼼ʦ).Value) <> "" _
                                And Nvl(rptList.FocusedRow.Record(mcol.��鼼ʦ).Value) <> mstrUserNameHIS Then
                                MsgBoxD Me, "�㲻��������ߵļ�鼼ʦ���޷�������ݱ�", vbInformation, gstrSysName
                            Else
                                If mblnPacsReport = True Then
                                    If control.ID = conMenu_PacsReport_Open Then   '�򿪱��洰��
                                        Call Menu_Manage_PACS����
                                    Else
                                        mfrmPacsReport.zlExecuteCommandBars control
                                    End If
                                Else
                                    mobjReport.zlExecuteCommandBars control
                                End If
                            End If
                        Case "�������"
                            mobjExpense.zlExecuteCommandBars control
                        Case "סԺҽ��"
                            mobjInAdvice.zlExecuteCommandBars control
                        Case "����ҽ��"
                            mobjOutAdvice.zlExecuteCommandBars control
                        Case "סԺ����"
                            mobjInEPRs.zlExecuteCommandBars control
                        Case "���ﲡ��"
                            mobjOutEPRs.zlExecuteCommandBars control
                    End Select
                End If
            End If
    End Select
End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo ErrHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
        .mDept = mlngCur����ID '��ǰ����
        .Show 1, Me
        If Not .mblnOK Then Exit Sub 'û�з�������
        
        mdatFBegin = Format(.dtpBegin.Value, "yyyy-MM-dd HH:mm:00")
        If Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            mdatFEnd = CDate(0) '��ʾȡ��ǰʱ��
        Else
            mdatFEnd = Format(.dtpEnd.Value, "yyyy-MM-dd HH:mm:59")
        End If
        
        mblnMoved = MovedByDate(IIf(mdatFBegin = CDate(0), CDate(zlDatabase.Currentdate) - mBeforeDays, mdatFBegin))
        
        '�Ƿ񱾴�סԺ
        mbln���� = (.chk����סԺ.Value = 1)
        
        'ʱ����ҷ�ʽ��1�������ʱ�䡢2��������ʱ��
        If .optFindType(0).Value = True Then
            mDatType = 1
        Else
            mDatType = 2
        End If
        
        '���ݺ�
        If .txtNO.Text <> "" Then
            mstrFNO = .txtNO.Text
        Else
        mstrFNO = ""
        End If
        
        '���걾��λ
        If .cboPart.ListIndex <> 0 Then
            mstr�걾��λ = .cboPart.Text
        Else
            mstr�걾��λ = ""
        End If
        
        '���˿���
        If .cboDept.ListIndex <> 0 Then
            mlngF����ID = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            mlngF����ID = 0
        End If
        
        '���˱�ʶ
        If .Txt��ʶ��.Text <> "" Then
            mstrF��ʶ�� = Trim(.Txt��ʶ��.Text)
        Else
            mstrF��ʶ�� = 0
        End If
        '���￨
        If .txt���￨.Text <> "" Then
            mstrF���￨ = .txt���￨.Text
        Else
            mstrF���￨ = ""
        End If
        '����
        If .txt����.Text <> "" Then
            mstrF���� = .txt����.Text
        Else
            mstrF���� = ""
        End If
        '����
        If .txtChkNO.Text <> "" Then
            mdblFChkNO = Val(.txtChkNO.Text)
        Else
            mdblFChkNO = 0
        End If

        '���ҽ��
        If .cbodiagdoc.ListIndex <> 0 Then
            mstr���ҽ�� = NeedName(.cbodiagdoc.Text)
        Else
            mstr���ҽ�� = ""
        End If
        '���ҽ��
        If .cboAuditing.ListIndex <> 0 Then
            mstr���ҽ�� = NeedName(.cboAuditing.Text)
        Else
            mstr���ҽ�� = ""
        End If
        
        '������
        If .cboCheckStep.ListIndex <> 0 Then
            mstr������ = .cboCheckStep.Text
        Else
            mstr������ = ""
        End If
        
        'Ӱ�����
        If .cboModality.ListIndex <> 0 Then
            mstrӰ����� = Split(.cboModality.Text, "--")(1)
        Else
            mstrӰ����� = ""
        End If
        
        'Ӱ�����
        If Trim(.TxtӰ�����) <> "" Then
            mstr������� = Trim(.TxtӰ�����)
        Else
            mstr������� = ""
        End If
        
        If .chk�������.Value = 1 Then
            mbln������� = True
        Else
            mbln������� = False
        End If
        
        If .cbo����.ListIndex = 0 Then
            mstrӰ������ = ""
        Else
            mstrӰ������ = NeedName(.cbo����.Text)
        End If
        
        If .cbo��鼼ʦ.ListIndex = 0 Then
            mstr��鼼ʦ = ""
        Else
            mstr��鼼ʦ = NeedName(.cbo��鼼ʦ.Text)
        End If
        
        'PACS�������
        If Trim(.txtPacsRpt(0)) <> "" Then
            mstr������� = Trim(.txtPacsRpt(0))
        Else
            mstr������� = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            mstr������ = Trim(.txtPacsRpt(1))
        Else
            mstr������ = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            mstr���� = Trim(.txtPacsRpt(2))
        Else
            mstr���� = ""
        End If
        
        '����ˢ��
        Call RefreshRptlist
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_View_FindType
            With CommandBar.Controls
                If .Count = 0 Then '��̬�Ӳ˵�,��1λ
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10#, "��ʶ��(&1)"): objControl.Category = "Main": objControl.Checked = True
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 1, "���￨(&2)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 2, "����(&3)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 3, "���ݺ�(&4)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 4, "����(&5)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 5, "���֤(&6)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_View_FindType * 10# + 6, "�ɣÿ�(&7)"): objControl.Category = "Main"
                End If
            End With
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    For i = 0 To UBound(Split(mstrCanUse����, "|")) 'mstrCanUse����=id_����-����|id_����-����
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i, Split(Split(mstrCanUse����, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse����, "|")(i), "_")(0)
                        If mlngCur����ID = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.Tag
                Case "סԺҽ��"
                    mobjInAdvice.zlPopupCommandBars CommandBar
                Case "����ҽ��" '����
                    mobjOutAdvice.zlPopupCommandBars CommandBar
                Case "�������"
                    mobjExpense.zlPopupCommandBars CommandBar
            End Select
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim blnNoRecord As Boolean, intState As Integer, intStep As Integer
    If Not mblnInitOK Then Exit Sub
    
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    control.Style = cbrMain(2).Controls(3).Style
    
    If Not blnNoRecord Then
        intState = Nvl(rptList.FocusedRow.Record(mcol.���״̬).Value, 0)
        intStep = Nvl(rptList.FocusedRow.Record(mcol.ִ��״̬).Value, 0)
    End If
    
    Select Case control.ID
        Case conMenu_View_FindType
            control.Caption = "��" & mstrCurFindtype & "����(&G)"
        Case conMenu_View_FindType * 10# To conMenu_View_FindType * 10# + 6
            control.Checked = (InStr(control.Caption, mstrCurFindtype) > 0)
        Case conMenu_View_Filter * 10#
            control.Caption = "��ǰ����:" & mstrCur����
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse����, "|"))
            control.Checked = (control.DescriptionText = mlngCur����ID)
        Case conMenu_View_ToolBar_Button '������
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            control.Checked = Me.cbrMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            control.Checked = Me.stbThis.Visible
        Case conMenu_View_PatInfor '������Ϣ
            control.Enabled = Not blnNoRecord
        Case conMenu_View_Filter   '����
        
        Case conMenu_View_Refresh  'ˢ��
        
        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord
                
        Case conMenu_Manage_Regist   '���Ǽ�(&I)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            End If
        Case conMenu_Manage_CopyCheck '�ٴεǼ�
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Redo   'ȡ���Ǽ�(&R)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '�ٻ�ȡ��
            If Not blnNoRecord Then
                control.Enabled = intStep = 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '�޸���Ϣ(&M)
            If InStr(mstrPrivs, "���Ǽ�") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 3 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '��鱨��(&L)
            If InStr(mstrPrivs, "��鱨��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intStep <> 2
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   'ȡ������(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent.Type = xtpControlPopup Then
                If InStr(mstrPrivs, "ȡ������") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "ȡ������"
                    control.Caption = "ȡ������(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' �������е���ȡ��������ȡ���Ǽ�,ͬһ�������ȡ���ǼǺ�ȡ����鹦��
                control.Visible = IIf(intState <= 1, InStr(mstrPrivs, "���Ǽ�") > 0, InStr(mstrPrivs, "ȡ������") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intStep <> 2) '���ܾ��Ĳ��ܱ��ٴξܾ�
                control.ToolTipText = IIf(intState <= 1, "ȡ���Ǽ�", "ȡ������")
                control.Caption = "ȡ��"
            End If
        Case conMenu_Manage_Transfer   '����Ӱ��(&C)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Cancel   'ȡ������(&B)
            If InStr(mstrPrivs, "ͼ�����") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_First, conMenu_Manage_Second, conMenu_Manage_Quality
            If InStr(mstrPrivs, "Ӱ���ʿ�") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive '���������(&X)
            If (InStr(GetInsidePrivs(p���Ʊ������), "������д") <= 0 And InStr(GetInsidePrivs(p���Ʊ������), "�����޶�") <= 0) Or _
                mblnIgnoreResult Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '��ɫͨ�����/ȡ��
            If InStr(mstrPrivs, "��ɫͨ��") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '��2---5֮�����
            End If
        Case conMenu_Manage_Finish   '�ޱ������(&F)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 2 Or intState = 3
            End If
        Case conMenu_Manage_ClearUp   '�ޱ������(&U)
            If InStr(mstrPrivs, "�ޱ������") <= 0 Then
                control.Visible = False
            ElseIf intState = 5 Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.����ID).Value) = 0
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '������(&E)
            If InStr(mstrPrivs, "������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState = 4 Or intState = 5)
            End If
        Case conMenu_Manage_Undone   'ȡ�����(&U)
            If InStr(mstrPrivs, "ȡ��������") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_File_SendImg  '����ͼ��
            If InStr(mstrPrivs, "�ļ�����") <= 0 Then control.Visible = False
        Case conMenu_Img_Contrast, conMenu_Img_Look     'Ӱ��Ա�,Ӱ���Ƭ
            If blnNoRecord Then control.Enabled = False: Exit Sub
            If mblnIsHistory = True Then
                control.Enabled = mstrHStudyUID <> ""
            Else
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
            End If
            If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
'        Case conMenu_Img_Look      'Ӱ���Ƭ
'            If blnNoRecord Then control.Enabled = False: Exit Sub
'
'            If control.Parent.Type <> xtpControlPopup Then
'                control.Visible = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
'                control.Enabled = control.Visible
'            Else
'                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
'            End If
        Case conMenu_Img_3D     '��ά�ؽ�
            If InStr(mstrPrivs, "��ά�ؽ�����") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
                End If
            End If
        Case conMenu_Img_Delete '���ͼ��
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Img_Query ',��ȡͼ��
            If InStr(mstrPrivs, "���ͼ��") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState > 1
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ChangeDevice    '����Ӱ���豸
            If blnNoRecord = True Then
                control.Enabled = False
            Else
                If UCase(Nvl(rptList.FocusedRow.Record(mcol.Ӱ�����).Value)) = "CR" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.Ӱ�����).Value)) = "DR" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.Ӱ�����).Value)) = "DX" Or _
                    UCase(Nvl(rptList.FocusedRow.Record(mcol.Ӱ�����).Value)) = "RF" Then
                    control.Enabled = True
                Else
                    control.Enabled = False
                End If
            End If
        Case conMenu_File_PrintSet     '��ӡ����(&S)
        Case conMenu_File_Preview, conMenu_File_Print '����Ԥ��(&V) �����ӡ(&P)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_Excel         '�嵥��ӡ(&L)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_BatPrint    ' ������ӡ(&B)
            control.Enabled = rptList.Records.Count > 0
        Case conMenu_File_Parameter     '��������(&O)
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '����
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_File_Exit
        Case conMenu_View_ToolBar
        Case Else
            If control.Category <> "Main" Then
                Select Case TabWindow.Selected.Tag
                    Case "������д"
                        If mblnPacsReport = True Then
                            mfrmPacsReport.zlUpdateCommandBars control
                        Else
                            mobjReport.zlUpdateCommandBars control
                        End If
                    Case "�������"
                        mobjExpense.zlUpdateCommandBars control
                    Case "סԺҽ��"
                        mobjInAdvice.zlUpdateCommandBars control
                    Case "����ҽ��"
                        mobjOutAdvice.zlUpdateCommandBars control
                    Case "סԺ����"
                        mobjInEPRs.zlUpdateCommandBars control
                    Case "���ﲡ��"
                        mobjOutEPRs.zlUpdateCommandBars control
                End Select

                If Not blnNoRecord Then
                    'ɾ��ֻ�����ѱ���ͽ����п���
                    If control.ID = conMenu_Edit_Delete And rptList.FocusedRow.Record(mcol.���״̬).Value >= 4 Then
                        control.Enabled = False
                    End If
                    '��ǰ�鿴�������μ�¼��˵���������
                    If cboTimes.ListIndex <> -1 Then
                        If rptList.FocusedRow.Record(mcol.ҽ��ID).Value <> cboTimes.ItemData(cboTimes.ListIndex) Then control.Enabled = False
                    End If
                    '����ɳ�����,�Լ�ҽ���б���鿴��ӡ����Ƭ�˵����������
                    If rptList.FocusedRow.Record(mcol.���״̬).Value = 6 Then
                        Select Case control.ID
                            Case conMenu_Edit_MarkMap, conMenu_File_Open, conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                control.Enabled = True
                            Case Else
                                control.Enabled = False
                        End Select
                    End If
                End If
            End If
    End Select
End Sub

Private Sub chkSource_Click(Index As Integer)
    If Not mblnInitOK Then Exit Sub
    Call RefreshRptlist
End Sub



Private Sub Menu_Manage_PACS����()
    Dim i As Integer
    
    If Not rptList.FocusedRow Is Nothing Then
        If Not mfrmPacsReportDock Is Nothing Then
            '���жϵ�ǰ�����Ƿ�����Ҫ�򿪵Ĵ��壬������ǣ�����Ҵ�������
            If rptList.FocusedRow.Record(mcol.ҽ��ID).Value = mfrmPacsReportDock.mlngAdviceID Then
                '��ǰmfrmPacsReportDockָ��Ĵ��壬������Ҫ�򿪵Ĵ���
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        End If
        
        '���Ҵ�������,�ҵ���Ҫ�򿪵Ĵ��壬��ͨ��Zorder�Ѵ�����ʾ����ǰ��
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If rptList.FocusedRow.Record(mcol.ҽ��ID).Value = mobjPacsReportArry(i).mlngAdviceID Then
                    Set mfrmPacsReportDock = mobjPacsReportArry(i)
                    mfrmPacsReportDock.WindowState = 0  'normal
                    mfrmPacsReportDock.ZOrder
                    Exit Sub
                End If
            Next i
        End If
        
        'û���ҵ���Ҫ�򿪵Ĵ��壬�Ҵ��´���,����¼��ǰ����
        Set mfrmPacsReportDock = New frmReport
        mfrmPacsReportDock.zlEditReport rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mlngCur����ID, Me, mstrPrivs, mlngModul, rptList.FocusedRow.Record(mcol.ת��).Value = 1
        
        If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
            ReDim mobjPacsReportArry(1) As frmReport
        Else
            ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
        End If
        Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.Hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.Hwnd
    End If
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs           'Ȩ��
    mlngModul = glngModul           'ģ���
    mlngCur����ID = 0
    mstrCur���� = ""
    mstrCanUse���� = ""
    mstrCurFindtype = "���￨"
    mblnInitOK = False  '��ʼ����,��ʼ�����֮ǰ���������ݵ���ȡ
    Call InitLocalPars '����ע������
    If Not InitDepts Then Unload Me: Exit Sub '��ʼ��ҽ������
    Call InitMvar '��ʼ��ģ�鼶����
    '��ʼ�Ӵ���
    Set mfrmPACSImg = New frmPACSImg
    Set mfrmPacsReport = New frmReport  'PACS����
    Set mobjReport = New zlRichEPR.cDockReport
    Set mobjExpense = New zlCISKernel.clsDockExpense
    Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
    Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
    Set mobjInEPRs = New zlRichEPR.cDockInEPRs
    Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    
    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitRptList

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    'ȥ��PACS���洰��Ŀ��ƿ�
    FormSetCaption mfrmPacsReport, False, False
    mblnInitOK = True '��ʼ�����
    Call RefreshRptlist
    
    Call RestoreWinState(Me, App.ProductName)
    
    ClearCacheFolder App.Path & "\TmpImage\"    '����ʱĿ¼���ˣ�����ո�Ŀ¼
    
    mstrUserNameHIS = UserInfo.����
    Me.stbThis.Panels(3).Text = "����ҽ����" & mstrUserNameHIS
    ReDim mobjPacsReportArry(0) As frmReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "סԺ����", IIf(mblncmdסԺ, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���ﲡ��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��첡��", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�����ѽ�", IIf(mblncmd�ѽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����δ��", IIf(mblncmdδ��, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "�Ǽǲ���", IIf(mblncmd�Ǽ�, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��������", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "���没��", IIf(mblncmd����, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��˲���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ɲ���", IIf(mblncmd���, 1, 0)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��λ��ʽ", mstrCurFindtype
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    
    '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
    If mblnPacsReport = True Then    'ʹ��PACS����༭��
        Call mfrmPacsReport.PromptModify
    End If
    
    Unload mfrmPACSImg
    Unload mfrmPacsReport
    Unload mobjReport.zlGetForm
    Unload mobjExpense.zlGetForm
    Unload mobjInAdvice.zlGetForm
    Unload mobjOutAdvice.zlGetForm
    Unload mobjInEPRs.zlGetForm
    Unload mobjOutEPRs.zlGetForm
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mfrmPacsReport = Nothing
    Set mobjReport = Nothing
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjPacsCore = Nothing
    
    '�������ά�ؽ����ر���ά�ؽ��Ĵ���
    If mblnUse3D = True Then
        On Error Resume Next
        Call sub3DProcess("EXIT")
    End If
End Sub

Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Private Sub mfrmPacsReport_BeforeEdit()
Dim lngOrderID As Long

    On Error GoTo ErrHandle
    lngOrderID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value
    If CheckConcurrentReport(lngOrderID) Then '����Ƿ��������ڲ�������
        Call UpdateReporter(lngOrderID, UserInfo.����)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsReportDock_AfterOpen()
    Call AfterReportOpen
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If Txt��ʶ��.Text = "" And Me.ActiveControl Is Txt��ʶ�� Then
        IDKind.IDKind = IDKinds.C2���֤��
        mstrCurFindtype = "���֤"
        Txt��ʶ�� = strID
        Call Txt��ʶ��_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub mobjInAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lngҽ��ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select ҽ��ID FROM ����ҽ������ where ����ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡҽ��ID", CLng(����ID))
    If rsTemp.EOF Then Exit Sub
    
    lngҽ��ID = Nvl(rsTemp!ҽ��ID, 0)
    mobjReport.zlRefresh lngҽ��ID, mlngCur����ID, False '�Բ���Edit��ʽˢ�¶���
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '���ò��ı���
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal ҽ��ID As Long)
    '����100��ͼ������У�Ĭ��ÿ��5�Ŵ�һ��
    Call OpenViewer(mobjPacsCore, ҽ��ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    If mblnPacsReport = True Then
        mfrmPacsReport.RefPacsPic 'ˢ��ͼƬ
        If Not mfrmPacsReportDock Is Nothing Then
            mfrmPacsReportDock.RefPacsPic 'ˢ��ͼƬ
        End If
    Else
        mobjReport.RefPacsPic 'ˢ��ͼƬ
    End If
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub
Public Sub EditorClosed(ByVal lngOrderID As Long)
    Dim i As Integer
    Dim j As Integer
    
    Call UpdateReporter(lngOrderID, "")
    '����PACS����༭���Ĵ�������
    On Error Resume Next
    If mblnPacsReport = True Then
        '���Ҵ������飬�ҵ���Ӧ�Ĵ��ڲ�ɾ��
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If mobjPacsReportArry(i).mlngAdviceID = lngOrderID Then
                    '��������ɾ��
                    For j = i To UBound(mobjPacsReportArry)
                        Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                    Next j
                    ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
                    Exit For
                End If
            Next i
        End If
        
        If Not mfrmPacsReportDock Is Nothing Then
            If lngOrderID = mfrmPacsReportDock.mlngAdviceID Then
                '�رյ�ǰ���洰�ڣ�����ǰ�������óɿ�
                Set mfrmPacsReportDock = Nothing
            End If
        End If
    End If
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
    On Error GoTo ErrHandle
    gstrSQL = "ZL_Ӱ�񱨸���_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "��ձ��"
    Call RefreshRptlist
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
    Call AfterReportOpen
End Sub

Private Sub AfterReportOpen()
Dim lngOrderID As Long
    On Error GoTo ErrHandle
    lngOrderID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value
    
    Call UpdateReporter(lngOrderID, UserInfo.����)
    
    If mblnShowImgAtReport And Nvl(rptList.FocusedRow.Record(mcol.���UID).Value) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , mblnLocalizerBackward, intImageInverval)
    End If
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub
Public Sub AfterPrinted(lngOrderID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "ZL_Ӱ�񱨸��ӡ_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "���´�ӡ���"
    
    If Not mblnIgnoreResult And mintResultInput = 2 Then
        strSQL = "Select �������  From  ����ҽ������ Where ҽ��id= [1]"
        Set rsTemp = OpenSQLRecord(strSQL, "��ȡ�������", lngOrderID)
        
        If IsNull(rsTemp!�������) Then  '�ڱ���ʱ��ʾ���������
            Call PromptResult(lngOrderID, mlngModul, Me)
        End If
    End If
    
    If mblnPrintCommit = True Then
        Call Menu_Manage_����������(lngOrderID, False)
    End If
    
    Call RefreshRptlist
    Exit Sub
    
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID)
End Sub

Public Sub AfterReportSaved(lngOrderID As Long)
    Dim rsTemp As ADODB.Recordset, i As Integer, intState As Integer, lngSendId As Long
    If mblnPacsReport = True Then
'        mfrmPacsReport.zlRefresh 0, 0, 0
    Else
        mobjReport.zlRefresh 0, mlngCur����ID, False
    End If

    gstrSQL = "Select Distinct A.ҽ��id, B.ID,B.������,B.������,B.ǩ������, B.���ʱ��, B.���汾, C.���ͺ�,C.�������, D.���UID " & vbNewLine & _
                "From ����ҽ������ A, ���Ӳ�����¼ B, ����ҽ������ C,Ӱ�����¼ D " & vbNewLine & _
                "Where A.ҽ��id =[1] And A.����id = B.ID And A.ҽ��id = C.ҽ��id AND D.ҽ��id = C.ҽ��id"
    Set rsTemp = OpenSQLRecord(gstrSQL, "��ȡ�Ƿ�ǩ��", CLng(lngOrderID))
    If rsTemp.EOF Then Exit Sub
    lngSendId = rsTemp!���ͺ�
    
    If Nvl(rsTemp!���ʱ��, "") = "" And rsTemp!���汾 = 1 Then 'δǩ������ �����һ��ҽʦ��ǩ
        gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & "," & IIf(Nvl(rsTemp!���UID) = "", 2, 3) & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ����ʱ"
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & Nvl(rsTemp!������, rsTemp!������) & "','')"
        zlDatabase.ExecuteProcedure gstrSQL, "���汨����"
        intState = IIf(Nvl(rsTemp!���UID) = "", 2, 3)
    Else
        If rsTemp!ǩ������ < 2 Then '���һ��ǩ��Ϊҽʦ,�п��ܵ���� 1-ҽʦ��N��ǩ�� 2-���μ������һ����ǩ 3-�޶�ģʽ�±���(ǩ������=0)
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & ",4)"
            zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ����ʱ"
            
            intState = 4
        Else                        '���μ����ϼ���ǩ��
            gstrSQL = "Zl_Ӱ����_State(" & lngOrderID & "," & lngSendId & ",5)"
            zlDatabase.ExecuteProcedure gstrSQL, "��Ϊ���ʱ"

            intState = 5
            If mblnCompleteCommit Then
                intState = 6
                Call Menu_Manage_����������(lngOrderID, False)
            End If
        End If
        
        gstrSQL = "ZL_Ӱ�񱨸汣��_Update(" & lngOrderID & ",'" & IIf(rsTemp!ǩ������ = 1, Nvl(rsTemp!������), IIf(rsTemp!���汾 = 1, Nvl(rsTemp!������), "")) & "','" & IIf(rsTemp!ǩ������ = 1, "", Nvl(rsTemp!������)) & "')"
        zlDatabase.ExecuteProcedure gstrSQL, "���渴����" 'ǩ�����𣽣���ʾ��ҽ��ǩ��,�����ǵ�N�Σ���ʱ����������Ҫ���棬��������Ҫ���;������������˴��գ��������д�������������ֵ
    
        If Not mblnIgnoreResult And IsNull(rsTemp!�������) Then  '�ڱ���ʱ��ʾ���������
            If mblnReportWithResult Then '��Ӱ�����Ϊ����  -����ʾ�Զ����
                gstrSQL = "ZL_Ӱ����_���(" & lngOrderID & ",0)"
                zlDatabase.ExecuteProcedure gstrSQL, "���������"
            ElseIf mintResultInput = 1 Then
                Call PromptResult(lngOrderID, mlngModul, Me)
            End If
        End If
    End If

    Call StateCheck(intState)
End Sub

Private Sub StateCheck(ByVal intState As Integer)
Dim cbrControl As CommandBarControl
    Select Case intState '���ݲ�����״̬ȷ����״̬�����Ƿ�ѡ��
        Case 0, 1
            If Not mblncmd�Ǽ� Then Set cbrControl = Me.cbrdock.FindControl(, ID_�Ǽ�)
        Case 2, 3
            If Not mblncmd���� Then Set cbrControl = Me.cbrdock.FindControl(, ID_����)
        Case 4
            If Not mblncmd���� Then Set cbrControl = Me.cbrdock.FindControl(, ID_����)
        Case 5
            If Not mblncmd��� Then Set cbrControl = Me.cbrdock.FindControl(, ID_���)
        Case 6
            If Not mblncmd��� Then Set cbrControl = Me.cbrdock.FindControl(, ID_���)
    End Select
    If mblnPatTrack Then
        If Not cbrControl Is Nothing Then '����ѡ��,ѡ�д����б�ˢ��ͬʱʵ�ָ���
            cbrdock_Execute cbrControl
        Else
            Call RefreshRptlist
        End If
    Else '������ֻˢ���б�
        Call RefreshRptlist
    End If
End Sub
Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'���ܣ���ʾ��ǰִ��ҽ�����Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    With rptList.FocusedRow
        gstrSQL = "Select Distinct C.���,C.����,C.˵��" & _
            " From ����ҽ����¼ A,��������Ӧ�� B,�����ļ��б� C" & _
            " Where A.ID=[1] And A.���ID IS NULL" & _
            " And A.������ĿID=B.������ĿID" & _
            " And B.Ӧ�ó���=[2] And B.�����ļ�ID=C.ID And C.����=7" & _
            " Order by C.���"
        If .Record(mcol.ת��).Value = 1 Then
            gstrSQL = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.Record(mcol.ҽ��ID).Value), CLng(Decode(Nvl(.Record(mcol.��Դ).Value, "��"), "��", 1, "סԺ", 2, "��", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!���� & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
    On Error GoTo errH
    If objControl.Parameter = "" Then '��֣�ֱ�Ӱ�F10ʱ����һ���յ�Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.Parameter = "" Then Exit Sub
    
    With rptList.FocusedRow
        If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
            Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & .Record(mcol.NO).Value, "����=" & .Record(mcol.��¼����).Value, 1)
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshRptlist()
Dim i As Integer, lngcurҽ��ID As Long
    If rptList.Records.Count >= 1 Then
        lngcurҽ��ID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value
    End If
    Call LoadPatiList
    If lngcurҽ��ID = 0 Then
        If rptList.Records.Count >= 1 And rptList.Rows.Count >= 1 Then
            rptList.FocusedRow = rptList.Rows(0)
        End If
        Exit Sub
    End If
    
    
    '�м�¼ʱҪ���¶�λ��֮ǰ��¼
    For i = 0 To rptList.Records.Count - 1
        If lngcurҽ��ID = rptList.Rows(i).Record.Item(mcol.ҽ��ID).Value Then
            rptList.FocusedRow = rptList.Rows(i)
            Exit Sub
        End If
    Next
    'û�ܶ�λ֮ǰ�ļ�¼����λ����0��
    If rptList.Records.Count >= 1 Then
        rptList.FocusedRow = rptList.Rows(0)
    End If
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraInfo.Top = -75
    fraInfo.Left = fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    lblCash.Top = (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl������Ϣ.Width = lblCash.Left
    lbl�����Ϣ.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'���ܣ���ȡ��ǰҽ�����ҵ�ִ��ҽ��(����)�嵥
'blnLocate���Ƿ�ˢ�µ�ǰ״̬������
Dim strSQL As String, strSQLBak As String, i As Long, rsPatList As ADODB.Recordset
Dim str��Դ As String
Dim strFilter As String

    
    If Not mblnInitOK Then Exit Sub      '��ʼ��δ���
    
    
    On Error GoTo ErrHandle
        
    '������ԴȨ��:(1-����,2-סԺ,3-����,4-���)
    If mblncmd���� Then str��Դ = "1,"
    If mblncmdסԺ Then str��Դ = str��Դ & "2,"
    If mblncmd���� Then str��Դ = str��Դ & "3,"
    If mblncmd��� Then str��Դ = str��Դ & "4,"
    If InStr(Len(str��Դ), str��Դ, ",") > 0 Then str��Դ = Mid(str��Դ, 1, Len(str��Դ) - 1) 'ȥ�����Ķ���,���߶����ڲ���������
        
    '����ʱ��
    If mdatFEnd <> CDate(0) Then
        strFilter = " And " & IIf(mblncmd�Ǽ�, "A.����ʱ��", IIf(mDatType = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [1] and [2] "
    Else 'ȱʡ��ѯ����
        strFilter = " And " & IIf(mblncmd�Ǽ�, "A.����ʱ��", IIf(mDatType = 2, "A.����ʱ��", "A.�״�ʱ��")) & " Between [1] and Sysdate+1/(24*3600) "
    End If
    '���ݺ�
    If mstrFNO <> "" Then
        strFilter = strFilter & " And A.NO= [3] "
    End If

    '���˿���
    If mlngF����ID <> 0 Then
        strFilter = strFilter & " And B.���˿���ID+0= [4] "
    End If

    '���˱�ʶ

    If mstrF��ʶ�� <> 0 Then
        strFilter = strFilter & " And Decode(B.������Դ,2,D.סԺ��,D.�����)= [5] "
    End If

    If mstrF���￨ <> "" Then
        strFilter = strFilter & " And D.���￨�� = [6] "
    End If

    If mstrF���� <> "" Then
        strFilter = strFilter & " And Instr(D.���� , [7])>0 "
    End If

    If mstr�걾��λ <> "" Then
        strFilter = strFilter & " And instr(B.ҽ������,[12])>0"
    End If

    If mdblFChkNO <> 0 Then
        strFilter = strFilter & " And H.����=[11] "
    End If
    
    If mbln������� Then
        strFilter = strFilter & " And Nvl(a.�������, 0)=1"
    End If
    
    If mstr���ҽ�� <> "" Then
        strFilter = strFilter & " And h.������=[13] "
    End If
    
    If mstr���ҽ�� <> "" Then
        strFilter = strFilter & " And h.������=[14] "
    End If
    
    '������
    If mstr������ <> "" Then
        If mstr������ = "ȫ��" Then
        
        ElseIf mstr������ = "�ѵǼ�" Then
            strFilter = strFilter & " And (a.ִ�й��� =0 or a.ִ�й���=1 Or a.ִ�й��� Is Null) "
        ElseIf mstr������ = "�ѱ���" Then
            strFilter = strFilter & " And (a.ִ�й��� = 2 and h.������ is null) "
        ElseIf mstr������ = "�Ѽ��" Then
            strFilter = strFilter & " And (a.ִ�й��� = 3 and h.������ is null) "
        ElseIf mstr������ = "������" Then
            strFilter = strFilter & " And (not h.������� is null) "
        ElseIf mstr������ = "������" Then
            strFilter = strFilter & " And ((a.ִ�й��� =2 or a.ִ�й���=3) and not h.������ is null and h.������� is null) "
        ElseIf mstr������ = "�ѱ���" Then
            strFilter = strFilter & " And (a.ִ�й���=4 and h.������ is null) "
        ElseIf mstr������ = "�����" Then
            strFilter = strFilter & " And (a.ִ�й���=4 and not h.������ is null) "
        ElseIf mstr������ = "�����" Then
            strFilter = strFilter & " And a.ִ�й���=5 "
        ElseIf mstr������ = "�����" Then
            strFilter = strFilter & " And a.ִ�й���=6 "
        End If
    End If
    
    If mstr������� <> "" Then
        strFilter = strFilter & " And F.����ID IN(Select Distinct A.Id From ���Ӳ�����¼ A,���Ӳ������� B Where A.����ʱ��>[1] AND A.Id=B.�ļ�ID And instr(B.�����ı�,[15])>0)"
    End If
    
    If mstrӰ������ <> "" Then
        strFilter = strFilter & " And h.Ӱ������=[16]"
    End If
    
    If mstr��鼼ʦ <> "" Then
        strFilter = strFilter & " And h.��鼼ʦ=[17]"
    End If
    
    If mstrRoom <> "" Then
        If Not mblncmd�Ǽ� Then
            strFilter = strFilter & " And Instr([10],','|| A.ִ�м� || ',' )>0"
        Else
            strFilter = strFilter & " And Instr([10],','|| A.ִ�м� || ',' )>0"
        End If
    End If
    
    If mblnNoShowCancel Then '����ʾȡ���Ǽǵļ��
        strFilter = strFilter & " And A.ִ��״̬<>2 "
    End If

    'Ӱ�����
    If mstrӰ����� <> "" Then
        strFilter = strFilter & " And h.Ӱ�����=[18] "
    End If
    
    '����PACS�����������
    If mstr������� <> "" Or mstr������ <> "" Or mstr���� <> "" Then
        Dim strSubFilter As String
        If mstr������� <> "" Then
            strSubFilter = " (b.�����ı� ='�������' And Instr(c.�����ı�, [19]) > 0)"
        End If
        
        If mstr������ <> "" Then
            If strSubFilter = "" Then
                strSubFilter = " (b.�����ı� ='������' And Instr(c.�����ı�, [20]) > 0)"
            Else
                strSubFilter = strSubFilter & " or (b.�����ı� ='������' And Instr(c.�����ı�, [20]) > 0)"
            End If
        End If
        
        If mstr���� <> "" Then
            If strSubFilter = "" Then
                strSubFilter = " (b.�����ı� ='����' And Instr(c.�����ı�, [21]) > 0)"
            Else
                strSubFilter = strSubFilter & " or (b.�����ı� ='����' And Instr(c.�����ı�, [21]) > 0)"
            End If
        End If
        
        strSubFilter = " (" & strSubFilter & ")"
        
        
        strFilter = strFilter & " And F.����ID IN(Select Distinct a.Id From ���Ӳ�����¼ a, ���Ӳ������� b,���Ӳ������� c " _
            & " Where a.����ʱ�� > [1] And a.Id = b.�ļ�id And b.Id = C.��ID And b.�������� = 3 And c.�������� = 2 And c.��ֹ�� = 0 and " _
            & strSubFilter & ")"
    End If
    

    
    strSQL = "Select /*+ Rule*/" & vbNewLine & _
                "Distinct a.ҽ��id, a.���ͺ�, a.�״�ʱ�� As ���ʱ��, a.����ʱ�� As ����ʱ��, a.No, a.��¼����, a.ִ��״̬," & vbNewLine & _
                "         Nvl(a.ִ�й���, 0) As ���״̬, a.ִ�м�, a.������� As ����, b.������Ŀid, b.����id, b.��ҳid," & vbNewLine & _
                "         b.�Һŵ� As �Һŵ�, b.���˿���id, Decode(b.������Դ, 1, '��', 2, 'סԺ', 3, '��', 4, '��') As ��Դ, b.ҽ������," & vbNewLine & _
                "         b.�걾��λ, Nvl(b.������־, 0) ������־, Nvl(b.Ӥ��, 0) Ӥ��, c.���� As ����, d.����, d.�Ա�, d.����," & vbNewLine & _
                "         d.���֤��, Decode(b.������Դ, 1, d.�����, 2, d.סԺ��, 4, d.�����, Null) As ��ʶ��," & vbNewLine & _
                "         Nvl(d.�ѱ�, '��ͨ') As �ѱ�, d.��ǰ����id As ����id, d.���￨��, e.���� As ����, Nvl(f.����id, 0) As ����id," & vbNewLine & _
                "         Nvl(h.���, '') ���, Nvl(h.����, '') ����, Nvl(h.����, '') As ����, Nvl(h.���uid, '') As ���uid,H.Ӱ������,h.��鼼ʦ, " & vbNewLine & _
                "         H.�Ƿ��ӡ,H.�������,0 as ת��,h.Ӱ�����,H.��ɫͨ��,H.�����ӡ,H.������,H.������,a.������ as �Ǽ���,h.������,h.�����,d.��ǰ����,b.����ҽ��,h.�������� as ��ͼʱ��  " & vbNewLine & _
                " From ����ҽ������ a, ����ҽ����¼ b, ������ĿĿ¼ c, ������Ϣ d, ���ű� e, ����ҽ������ f, Ӱ�����¼ h,Ӱ������Ŀ G" & vbNewLine & _
                " Where a.ҽ��id = b.Id And b.������Ŀid = c.Id And b.����id = d.����id And b.���˿���id = e.Id And" & vbNewLine & _
                "      a.ҽ��id = h.ҽ��id(+) And a.���ͺ� = h.���ͺ�(+) And a.ҽ��id = f.ҽ��id(+) And B.������ĿID=G.������ĿID AND" & vbNewLine & _
                "      Instr([8],','||B.������Դ||',')> 0 And A.ִ�в���ID+0= [9] And" & _
                IIf(mbln����, " (B.������Դ=2 And b.��ҳID=d.סԺ���� Or Nvl(B.������Դ,0)<>2) and ", "") & _
                "      B.���ID is NULL " & strFilter
    '���������ת����Ҫ�����󱸱�
    If mblnMoved Then
        strSQLBak = strSQL
        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
        strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ������", "H����ҽ������")
        strSQLBak = Replace(strSQLBak, "���Ӳ�����¼", "H���Ӳ�����¼")
        strSQLBak = Replace(strSQLBak, "���Ӳ�������", "H���Ӳ�������")
        strSQLBak = Replace(strSQLBak, "0 as ת��", "1 as ת��")
        strSQL = strSQL & " Union ALL " & strSQLBak
    End If
    strSQL = "Select * From (" & strSQL & ") Order by ���״̬,���ʱ��,����ʱ��"
    
    Set rsPatList = OpenSQLRecord(strSQL, Me.Caption, CDate(Format(mdatFBegin, "yyyy-MM-dd HH:mm:00")), CDate(Format(mdatFEnd, "yyyy-MM-dd HH:mm:59")), _
                                mstrFNO, mlngF����ID, mstrF��ʶ��, mstrF���￨, mstrF����, "," & str��Դ & ",", mlngCur����ID, _
                                mstrRoom, mdblFChkNO, mstr�걾��λ, mstr���ҽ��, mstr���ҽ��, mstr�������, mstrӰ������, mstr��鼼ʦ, mstrӰ�����, mstr�������, mstr������, mstr����)
    strFilter = ""
    If mblncmd�Ǽ� Then strFilter = "���״̬=0 or ���״̬=1 or "
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "���״̬=2 or ���״̬=3 or ", "���״̬=2 or ���״̬=3 or ")
    If mblncmd���� Then strFilter = IIf(strFilter <> "", strFilter & "���״̬=4 or ", "���״̬=4 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "���״̬=5 or ", "���״̬=5 or ")
    If mblncmd��� Then strFilter = IIf(strFilter <> "", strFilter & "���״̬=6 or ", "���״̬=6 or ")
    
    If strFilter = "" Then
        strFilter = "���״̬<0"
    Else
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
    End If
    rsPatList.Filter = strFilter
    Call RefreshPatList(rsPatList)
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub RefreshPatList(ByVal rsPatList As ADODB.Recordset)
Dim rptRecord As ReportRecord, i As Long, j As Long, blnRowShow As Boolean
Dim rsTemp As New ADODB.Recordset, StrAdviceIds As String, strSQL As String
        
    On Error GoTo ErrHandle
    If Not mblnInitOK Then Exit Sub
    
    rptList.Records.DeleteAll
    If rsPatList.EOF Then
        rptList.Populate
        rptList_SelectionChanged
        Exit Sub
    End If
    
    If mblncmd�ѽ� Xor mblncmdδ�� Then '����ѡ��
        gstrSQL = "Select /*+ RULE */" & vbNewLine & _
                    "Distinct A.ҽ��id" & vbNewLine & _
                    "From ����ҽ������ A, ���˷��ü�¼ B, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) C" & vbNewLine & _
                    "Where A.ҽ��id = C.Column_Value And A.NO = B.NO And A.��¼����=B.��¼���� And B.��¼״̬ = 1"
        Do Until rsPatList.EOF
            StrAdviceIds = StrAdviceIds & "," & rsPatList!ҽ��ID
            If Len(StrAdviceIds) > 3880 Or rsPatList.AbsolutePosition = rsPatList.RecordCount Then 'VARCHAR2��󳤶�4000
                StrAdviceIds = Mid(StrAdviceIds, 2)
                strSQL = strSQL & " Union " & Replace(gstrSQL, "[1]", "'" & StrAdviceIds & "'")
                StrAdviceIds = ""
            End If
            rsPatList.MoveNext
        Loop
        strSQL = Mid(strSQL, 8)
        Call zlDatabase.OpenRecordset(rsTemp, strSQL, "��ȡ�Ƿ��շ�")
    End If
    
    rsPatList.MoveFirst
    Do Until rsPatList.EOF
        blnRowShow = False
        If mblncmd�ѽ� Xor mblncmdδ�� Then '����ѡ��
            If Not rsTemp Is Nothing Then
                rsTemp.Filter = ""
                rsTemp.Filter = "ҽ��ID=" & rsPatList!ҽ��ID
                If mblncmd�ѽ� Then '�������ѽ�
                    If Not rsTemp.EOF Then blnRowShow = True
                Else                '������δ��
                    If rsTemp.EOF Then blnRowShow = True
                End If
            End If
        Else
            blnRowShow = True
        End If
    
        If blnRowShow Then
            Set rptRecord = rptList.Records.Add
            For j = 0 To Me.rptList.Columns.Count + 1
                rptRecord.AddItem ""
                If rsPatList!���״̬ = 6 Then rptRecord.Item(j).BackColor = &HFF00&
                If rsPatList!ִ��״̬ = 2 Then rptRecord.Item(j).BackColor = &HFFFF&
            Next

            rptRecord.Item(mcol.����).Value = IIf(rsPatList("������־") = 0, "", "����")
                rptRecord.Item(mcol.����).Icon = IIf(rsPatList("������־") = 0, -1, Me.imgList.ListImages("����").Index - 1)
            rptRecord.Item(mcol.��Դ).Value = IIf(rsPatList("��Դ") = "סԺ", "סԺ", rsPatList!��Դ)
                rptRecord.Item(mcol.��Դ).Icon = IIf(rsPatList("��Դ") = "סԺ", Me.imgList.ListImages("סԺ").Index - 1, -1)
            rptRecord.Item(mcol.����).Value = Decode(rsPatList!����, 0, "", 1, "����", rsPatList!����)
                rptRecord.Item(mcol.����).Icon = IIf(Nvl(rsPatList!����, 0) = 0, -1, Me.imgList.ListImages("����").Index - 1)
            rptRecord.Item(mcol.����).Value = Nvl(rsPatList!Ӱ������)
            rptRecord.Item(mcol.��ɫͨ��).Value = Nvl(rsPatList!��ɫͨ��, 0)
            rptRecord.Item(mcol.����).Value = rsPatList("����")
                rptRecord.Item(mcol.����).Icon = IIf(rptRecord.Item(mcol.��ɫͨ��).Value = 1, Me.imgList.ListImages("��ɫͨ��").Index - 1, -1)
            rptRecord.Item(mcol.����).Value = Nvl(rsPatList!����)
                rptRecord.Item(mcol.����).Icon = IIf(Len(rsPatList("���UID")) > 0, Me.imgList.ListImages("Ӱ��").Index - 1, -1)
            rptRecord.Item(mcol.��ʶ��).Value = Nvl(rsPatList("��ʶ��"))
            rptRecord.Item(mcol.�Ա�).Value = Nvl(rsPatList("�Ա�"))
            rptRecord.Item(mcol.����).Value = Nvl(rsPatList("����"))
            rptRecord.Item(mcol.������).Value = IIf(rsPatList!ִ��״̬ = 2, "�Ѿܾ�", Decode(Nvl(rsPatList!���״̬, 0), 0, "�ѵǼ�", 1, "�ѵǼ�", 2, IIf(Nvl(rsPatList!�������) <> "", "������", IIf(Nvl(rsPatList!������) = "", "�ѱ���", "������")), 3, IIf(Nvl(rsPatList!�������) <> "", "������", IIf(Nvl(rsPatList!������) = "", "�Ѽ��", "������")), 4, IIf(Nvl(rsPatList!�������) <> "", "������", IIf(Nvl(rsPatList!������) <> "", "�����", "�ѱ���")), 5, "�����", "�����"))
             rptRecord.Item(mcol.����).Value = rsPatList("����")
            If InStr(Nvl(rsPatList!ҽ������), ":") > 0 Then '�µ�ģʽ������ҽ����������Ϣ�� ����,ִ�б��:��λ(����,����),��λ---
'                rptRecord.Item(mcol.����).Value = Split(Split(rsPatList!ҽ������, ":")(0), ",")(0)
                rptRecord.Item(mcol.��λ).Value = Split(rsPatList!ҽ������, ":")(1)
            Else
'                rptRecord.Item(mcol.����).Value = Nvl(rsPatList!ҽ������)
                rptRecord.Item(mcol.��λ).Value = Nvl(rsPatList!�걾��λ)
            End If
            rptRecord.Item(mcol.ִ�м�).Value = Nvl(rsPatList("ִ�м�"))
            rptRecord.Item(mcol.���ʱ��).Value = Nvl(rsPatList("���ʱ��"))
            rptRecord.Item(mcol.����ʱ��).Value = Nvl(rsPatList("����ʱ��"))
            rptRecord.Item(mcol.��ͼʱ��).Value = Nvl(rsPatList("��ͼʱ��"))
            
            rptRecord.Item(mcol.�ѱ�).Value = Nvl(rsPatList("�ѱ�"))
            rptRecord.Item(mcol.���˿���).Value = rsPatList("����")
            rptRecord.Item(mcol.���￨��).Value = Nvl(rsPatList!���￨��)
            rptRecord.Item(mcol.���֤��).Value = Nvl(rsPatList!���֤��)
            rptRecord.Item(mcol.IC��).Value = Nvl(rsPatList!���￨��)
            
            rptRecord.Item(mcol.����ID).Value = Nvl(rsPatList("����ID"), 0)
            rptRecord.Item(mcol.��ҳID).Value = Nvl(rsPatList!��ҳID, 0)
            rptRecord.Item(mcol.���˿���ID).Value = Nvl(rsPatList("���˿���ID"), 0)
            rptRecord.Item(mcol.����ID).Value = Nvl(rsPatList("����ID"), 0)
            rptRecord.Item(mcol.�Һŵ�).Value = Nvl(rsPatList("�Һŵ�"))
            rptRecord.Item(mcol.ҽ��ID).Value = rsPatList("ҽ��ID")
            rptRecord.Item(mcol.���ͺ�).Value = rsPatList("���ͺ�")
            rptRecord.Item(mcol.������ĿID).Value = rsPatList("������ĿID")
            rptRecord.Item(mcol.����).Value = Nvl(rsPatList("��ǰ����"))
            rptRecord.Item(mcol.����ҽ��).Value = Nvl(rsPatList("����ҽ��"))
            rptRecord.Item(mcol.��ӡ��Ƭ).Value = IIf(Nvl(rsPatList!�Ƿ��ӡ, 0) = 0, "δ��ӡ", "�Ѵ�ӡ")
            rptRecord.Item(mcol.�������).Value = Nvl(rsPatList!�������)
            rptRecord.Item(mcol.�����ӡ).Value = IIf(Nvl(rsPatList!�����ӡ, 0) = 0, "δ��ӡ", "�Ѵ�ӡ")
            rptRecord.Item(mcol.������).Value = Nvl(rsPatList!������)
            rptRecord.Item(mcol.������).Value = Nvl(rsPatList!������)
            
            
            rptRecord.Item(mcol.NO).Value = rsPatList("NO")
            rptRecord.Item(mcol.��¼����).Value = Nvl(rsPatList!��¼����, 0)
            rptRecord.Item(mcol.ҽ������).Value = Nvl(rsPatList("ҽ������"))
            rptRecord.Item(mcol.���UID).Value = rsPatList("���UID")
            rptRecord.Item(mcol.���״̬).Value = rsPatList("���״̬")
            rptRecord.Item(mcol.Ӥ��).Value = rsPatList!Ӥ��
            rptRecord.Item(mcol.����ID).Value = Nvl(rsPatList("����ID"))
            rptRecord.Item(mcol.ҽ������).Value = ""
            rptRecord.Item(mcol.ִ��״̬).Value = rsPatList!ִ��״̬
            rptRecord.Item(mcol.ת��).Value = rsPatList!ת��
            rptRecord.Item(mcol.���).Value = Nvl(rsPatList!���)
            rptRecord.Item(mcol.����).Value = Nvl(rsPatList!����)
            rptRecord.Item(mcol.��鼼ʦ).Value = Nvl(rsPatList!��鼼ʦ)
            rptRecord.Item(mcol.Ӱ�����).Value = Nvl(rsPatList!Ӱ�����)
            
            rptRecord.Item(mcol.�Ǽ���).Value = Nvl(rsPatList!�Ǽ���)
            rptRecord.Item(mcol.������).Value = Nvl(rsPatList!������)
            rptRecord.Item(mcol.�����).Value = Nvl(rsPatList!�����)
        End If
        rsPatList.MoveNext
    Loop
    rptList.Populate
'    If rptList.Records.Count > 0 Then
'            rptList.FocusedRow = rptList.Rows(0)
'    End If
    stbThis.Panels(2).Text = "�� " & rptList.Records.Count & " ����¼": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub


Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("�Ҽ��˵�", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup _
                And Menucontrol.ID <> conMenu_View_Filter * 10# And Menucontrol.ID <> conMenu_View_FindType) And Menucontrol.Type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim blnNoRecord As Boolean

    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If

    If Not blnNoRecord Then
            Select Case rptList.FocusedRow.Record(mcol.���״̬).Value
                Case 1, 0
                    Call Menu_Manage_����
                Case 2, 3               '˫������д����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                    Call Menu_RichEPR(conMenu_Edit_Modify)
                Case 4, 5               '˫���޶�����,�����ʱ�����趨�Ƿ�򿪹�Ƭվ
                    Call Menu_RichEPR(conMenu_Edit_Audit)
                Case 6                  '����
                    Call Menu_RichEPR(conMenu_File_Open)
            End Select
    End If
End Sub

Private Sub rptList_SelectionChanged()
Dim blnNoRecord As Boolean, rsTemp As ADODB.Recordset, rptRecord As ReportRecord, strҽ������ As String, i As Integer
Dim blnShowReport As Boolean, strTemp As String
    
    mblnIsHistory = False
    
    blnShowReport = True
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    
    If Not blnNoRecord Then
        '�ж� ��ͼ����д����
        If mblnReportWithImage = True Then
            If rptList.FocusedRow.Record(mcol.���UID).Value = "" Or IsNull(rptList.FocusedRow.Record(mcol.���UID).Value) Then blnShowReport = False
        End If
        
        With rptList.FocusedRow
            lbl������Ϣ.Caption = "" 'cbotime�л��õ�������������listindexʱ�������ǵ��cbotimes����
            gstrSQL = "Select A.ID ҽ��ID,A.����ʱ��  ����ʱ��,A.ҽ������ " & _
                       " From ����ҽ����¼ A,����ҽ������ B,Ӱ������Ŀ C" & _
                       " Where A.����id = [1] And A.���id Is Null And A.ִ�п���id+0 =[2] And B.ҽ��ID=A.ID " & _
                       "" & IIf(.Record(mcol.������).Value = "�Ѿܾ�", "", " And B.ִ��״̬<>2 ") & _
                       " AND A.������ĿID=C.������ĿID"
            strTemp = Replace(gstrSQL, "����ҽ����¼", "H����ҽ����¼")
            strTemp = Replace(strTemp, "����ҽ������", "H����ҽ������")
            gstrSQL = gstrSQL & " Union ALL " & strTemp
            gstrSQL = "select * from (" & gstrSQL & ") Order By ����ʱ�� Asc"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(.Record(mcol.����ID).Value), mlngCur����ID)
            cboTimes.Clear
            Do Until rsTemp.EOF
               cboTimes.AddItem "��" & rsTemp.AbsolutePosition & "��(" & Format(rsTemp!����ʱ��, "yyyy-mm-dd") & ")  " & Trim(rsTemp!ҽ������)
               cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!ҽ��ID
               If rsTemp!ҽ��ID = rptList.FocusedRow.Record(mcol.ҽ��ID).Value Then cboTimes.ListIndex = cboTimes.NewIndex
               rsTemp.MoveNext
            Loop
            
            '�ж�Ƕ��ʽ����༭���еı����Ƿ�û�б���
            If mblnPacsReport = True Then    'ʹ��PACS����༭��
                Call mfrmPacsReport.PromptModify
            End If
                
            '��ʾ������Ϣ
            lbl������Ϣ.Caption = "��  ��:" & Rpad(.Record(mcol.����).Value, 12, " ") & "��  ��:" & Rpad(.Record(mcol.�Ա�).Value, 13, " ") & _
                                  "��  ��:" & Rpad(.Record(mcol.����).Value, 10, " ") & "��ʶ��:" & Rpad(.Record(mcol.��ʶ��).Value, 12, " ") & _
                                  "��  ��:" & Rpad(.Record(mcol.����).Value & "", 10, " ")
            lbl�����Ϣ.Caption = "����:" & Rpad(Nvl(.Record(mcol.����).Value), 12, " ") & "���˿���:" & Rpad(Nvl(.Record(mcol.���˿���).Value), 11, " ") & _
                                  "����ҽ��:" & Rpad(Nvl(.Record(mcol.����ҽ��).Value), 8, " ") & "�����Ŀ:" & .Record(mcol.����).Value
            lblCash.Caption = "��": lblCash.Visible = False
            
            If Nvl(.Record(mcol.ҽ������).Value) = "" Then
                gstrSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order By ����"
                If .Record(mcol.ת��).Value = 1 Then
                    gstrSQL = Replace(gstrSQL, "����ҽ������", "H����ҽ������")
                End If
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���˸���", CLng(.Record(mcol.ҽ��ID).Value))
                Do Until rsTemp.EOF
                    strҽ������ = strҽ������ & rsTemp!��Ŀ & ":" & Nvl(rsTemp!����) & vbCrLf
                    rsTemp.MoveNext
                Loop
                rptList.FocusedRow.Record(mcol.ҽ������).Value = strҽ������
            End If
            Txt������Ϣ = ""
            If InStr(Nvl(.Record(mcol.ҽ������).Value), ":") > 0 Then
                For i = 0 To UBound(Split(Split(.Record(mcol.ҽ������).Value, ":")(1), "),"))
                    If i = 0 Then
                        Txt������Ϣ = "��鲿λ:" & vbCrLf & Space(2) & "1:" & Split(Split(.Record(mcol.ҽ������).Value, ":")(1), "),")(i) & ")"
                    Else
                        Txt������Ϣ = Txt������Ϣ & vbCrLf & Space(2) & i + 1 & ":" & Split(Split(.Record(mcol.ҽ������).Value, ":")(1), "),")(i) & ")"
                    End If
                Next
                If Trim(Txt������Ϣ) <> "" Then Txt������Ϣ = Mid(Txt������Ϣ, 1, Len(Txt������Ϣ) - 1)
            Else
                Txt������Ϣ = "��鲿λ:" & Nvl(.Record(mcol.ҽ������).Value)
            End If
            
            Txt������Ϣ = Txt������Ϣ & vbCrLf & vbCrLf & Nvl(.Record(mcol.ҽ������).Value)
            lblCash.Visible = CheckChargeState(.Record(mcol.ҽ��ID).Value) = 1
            
            '�м�¼ʱ���ݲ�ͬ����״̬�ṩѡ�
            If .Record(mcol.��Դ).Value = "סԺ" Then '���ݲ�����Դ���Ʋ�����ҽ��ѡ�
                For i = 0 To TabWindow.ItemCount - 1
                    Select Case TabWindow(i).Tag
                        Case "���ﲡ��", "����ҽ��"
                            TabWindow(i).Visible = False
                        Case "סԺ����", "סԺҽ��"
                            TabWindow(i).Visible = True
                        Case "Ӱ��ͼ��"
                            TabWindow(i).Visible = True
                        Case "������д"
                            TabWindow(i).Visible = .Record(mcol.���״̬).Value > 1 And blnShowReport
                    End Select
                Next
            Else
                For i = 0 To TabWindow.ItemCount - 1
                    Select Case TabWindow(i).Tag
                        Case "���ﲡ��", "����ҽ��"
                            TabWindow(i).Visible = True
                        Case "סԺ����", "סԺҽ��"
                            TabWindow(i).Visible = False
                        Case "Ӱ��ͼ��"
                            TabWindow(i).Visible = True
                        Case "������д"
                            TabWindow(i).Visible = .Record(mcol.���״̬).Value > 1 And blnShowReport
                    End Select
                Next
            End If
            
            If mstrFirstTab <> "" Then '��Ϊ�ձ�ʾ��������ҳ��ʾ
                For i = 0 To TabWindow.ItemCount - 1
                    If InStr(TabWindow.Item(i).Tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                        If TabWindow.Item(i).Selected Then
                            Call TabWindow_SelectedChanged(TabWindow.Item(i))
                        Else
                            TabWindow.Item(i).Selected = True
                        End If
                        Exit Sub
                    End If
                Next
                TabWindow(0).Selected = True 'ûѭ�����˴�����0��tab
            Else
                If TabWindow.Selected.Visible Then
                    Call TabWindow_SelectedChanged(TabWindow(TabWindow.Selected.Index))
                Else
                    Select Case TabWindow.Selected.Tag
                        Case "���ﲡ��"
                            For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                                If TabWindow(i).Tag = "סԺ����" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "����ҽ��"
                            For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                                If TabWindow(i).Tag = "סԺҽ��" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "סԺ����"
                            For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                                If TabWindow(i).Tag = "���ﲡ��" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                        Case "סԺҽ��"
                            For i = 0 To TabWindow.ItemCount - 1 'ѭ�����˲Ŵ���
                                If TabWindow(i).Tag = "����ҽ��" Then TabWindow(i).Selected = True: Exit Sub
                            Next
                    End Select
                    TabWindow(0).Selected = True 'ûѭ�����˴�����0��tab
                End If
            End If
        End With
        '��ʾ�ɴ�ӡ�����Ƶ���:֮���Լ�ʱ����,��Ϊ��ʹ��F2�ȼ�
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    Else
        If TabWindow(TabWindow.Selected.Index).Visible Then
            Call TabWindow_SelectedChanged(TabWindow(TabWindow.Selected.Index))
        Else
             TabWindow(0).Selected = True
        End If
        cboTimes.Clear
        Txt������Ϣ = ""
        
        lbl������Ϣ.Caption = "��  ��:" & Space(12) & "��  ��:" & Space(13) & "��  ��:" & Space(10) & "��ʶ��:" & Space(12) & _
                                  "��  ��:" & Space(10)
        lbl�����Ϣ.Caption = "����:" & Space(12) & "���˿���:" & Space(11) & "����ҽ��:" & Space(8) & "�����Ŀ:"
                                  
        lblCash.Visible = False
    End If
    
    On Error Resume Next
    If rptList.Visible = True Then rptList.SetFocus
    err.Clear
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim blnNoRecord As Boolean
Dim Menucontrol As CommandBarPopup, cbrControl As CommandBarButton, bcontrol As CommandBarControl, i As Integer
    
    If Not mblnInitOK Then Exit Sub
    
    On Error Resume Next
    '�ж��Ƿ��м�¼
    blnNoRecord = False
    If rptList.FocusedRow Is Nothing Then
        blnNoRecord = True
    ElseIf rptList.FocusedRow.GroupRow Then
        blnNoRecord = True
    End If
    
    '����ָ���Ĳ˵���ѭ��ʱѭ������,����ָ��ɾ��
    
    For Each Menucontrol In cbrMain.ActiveMenuBar.Controls 'ɾ����������˵�,���Ǳ����������˵�
        If Menucontrol.ID <> conMenu_ReportPopup Then
            For Each bcontrol In Menucontrol.CommandBar.Controls
                If bcontrol.Category <> "Main" Then bcontrol.Delete
            Next
            If Menucontrol.Category <> "Main" Then Menucontrol.Delete
        End If
    Next
    
    Set Menucontrol = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    If Not Menucontrol Is Nothing Then
        Set cbrControl = Menucontrol.CommandBar.Controls.Find(, conMenu_View_Append)
        If Not cbrControl Is Nothing Then cbrControl.Delete
    End If
    
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_NewItem)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_Delete)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    Set bcontrol = cbrMain(2).Controls.Find(, conMenu_Edit_Modify)
    If Not bcontrol Is Nothing Then bcontrol.Delete
    For Each bcontrol In cbrMain(2).Controls 'ɾ�������湤����
        If bcontrol.Category <> "Main" Then
            bcontrol.Delete
        End If
    Next
    
    err.Clear


    On Error GoTo ErrHandle
    If blnNoRecord Then
        Select Case Item.Tag
            Case "Ӱ��ͼ��"
                Call mfrmPACSImg.zlRefresh(0, 0, mstrPrivs, False)
            Case "������д"
                If mblnPacsReport = True Then    'ʹ��PACS����༭��
                    mfrmPacsReport.zlDefCommandBars Me.cbrMain
                    mfrmPacsReport.zlRefresh 0, 0, 0, mstrPrivs, mlngModul, Me, False
                Else
                    mobjReport.zlDefCommandBars Me.cbrMain
                    mobjReport.zlRefresh 0, mlngCur����ID, False
                End If
            Case "�������"
                mobjExpense.zlDefCommandBars Me, Me.cbrMain
                mobjExpense.zlRefresh 0, 0, 0
            Case "סԺҽ��"
                mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjInAdvice.zlRefresh 0, 0, 0, 0, 0, False, 0, 0
            Case "����ҽ��"
                mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjOutAdvice.zlRefresh 0, "", False
            Case "סԺ����"
                mobjInEPRs.zlDefCommandBars cbrMain
                mobjInEPRs.zlRefresh 0, 0, 0, False
            Case "���ﲡ��"
                mobjOutEPRs.zlDefCommandBars cbrMain
                mobjOutEPRs.zlRefresh 0, 0, 0, False
        End Select
        Exit Sub
    End If
    
    
    With rptList.FocusedRow

        If cboTimes.ListIndex <> -1 Then
            If .Record(mcol.ҽ��ID).Value <> cboTimes.ItemData(cboTimes.ListIndex) Then '��ǰҽ��ID�뵱ǰ���μ�¼��ҽ��ID��ͬʱ��CboTimes����
                Call cboTimes_Click
                Exit Sub
            End If
        End If
        
        Select Case Item.Tag
            Case "Ӱ��ͼ��"
                '���ܵ�ǰ��¼û��ͼ�񣬶���mfrmPACSImg����ǿ��ˢ��
                Call mfrmPACSImg.zlRefresh(.Record(mcol.ҽ��ID).Value, .Record(mcol.���ͺ�).Value, mstrPrivs, .Record(mcol.ת��).Value = 1, True)
                '���ˢ�³����м�¼����ˢ�²����б�
                If (IsNull(.Record(mcol.���UID).Value) Or .Record(mcol.���UID).Value = "") And mfrmPACSImg.lvwSeq.ListItems.Count > 0 Then
                    Call RefreshRptlist
                End If
            Case "������д"
                If mblnPacsReport = True Then
                    mfrmPacsReport.zlDefCommandBars Me.cbrMain
                    Call mfrmPacsReport.zlRefresh(.Record(mcol.ҽ��ID).Value, .Record(mcol.���ͺ�).Value, mlngCur����ID, mstrPrivs, mlngModul, Me, .Record(mcol.ת��).Value = 1)
                Else
                    mobjReport.zlDefCommandBars Me.cbrMain
                    Call mobjReport.zlRefresh(Nvl(.Record(mcol.ҽ��ID).Value, 0), mlngCur����ID, True)
                End If
            Case "�������"
                mobjExpense.zlDefCommandBars Me, Me.cbrMain
                mobjExpense.zlRefresh mlngCur����ID, .Record(mcol.ҽ��ID).Value, .Record(mcol.���ͺ�).Value, .Record(mcol.ת��).Value = 1
            Case "סԺҽ��"
                mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                mobjInAdvice.zlRefresh .Record(mcol.����ID).Value, Val(.Record(mcol.��ҳID).Value), _
                    .Record(mcol.����ID).Value, .Record(mcol.���˿���ID).Value, 0, .Record(mcol.ת��).Value = 1, _
                    Nvl(.Record(mcol.ҽ��ID).Value, 0), .Record(mcol.ִ��״̬).Value
            Case "����ҽ��"
                mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
                If .Record(mcol.�Һŵ�).Value = "" Then
                    mobjOutAdvice.zlRefresh 0, "", False
                Else
                    mobjOutAdvice.zlRefresh .Record(mcol.����ID).Value, .Record(mcol.�Һŵ�).Value, True, .Record(mcol.ת��).Value = 1, Nvl(.Record(mcol.ҽ��ID).Value, 0)
                End If
            Case "סԺ����"
                mobjInEPRs.zlDefCommandBars cbrMain
                mobjInEPRs.zlRefresh .Record(mcol.����ID).Value, Val(.Record(mcol.��ҳID).Value), .Record(mcol.���˿���ID).Value, False, .Record(mcol.ת��).Value = 1
            Case "���ﲡ��"
                mobjOutEPRs.zlDefCommandBars cbrMain
                mobjOutEPRs.zlRefresh .Record(mcol.����ID).Value, Val(.Record(mcol.��ҳID).Value), .Record(mcol.���˿���ID).Value, False, .Record(mcol.ת��).Value = 1
        End Select
    End With
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub TimerRefresh_Timer()
    'ˢ�²����б�
    Call Menu_View_Refresh_click
End Sub

Private Sub txt��ʶ��_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Txt��ʶ��.Text = "" And Me.ActiveControl Is Txt��ʶ��)
    If Txt��ʶ��.Text = "" Then Txt��ʶ��.Tag = ""
End Sub

Private Sub txt��ʶ��_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '���֤ʶ�����
    
    If Txt��ʶ��.Text <> "" Then Call zlControl.TxtSelAll(Txt��ʶ��)
    If mstrCurFindtype = "����" Then
        Call zlCommFun.OpenIme(True)
    End If
    If Not mobjIDCard Is Nothing And Txt��ʶ��.Text = "" Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub Txt��ʶ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txt��ʶ��_Validate(False)
        Call zlControl.TxtSelAll(Txt��ʶ��)
        Call SeekNextPati(Txt��ʶ��.Tag <> Txt��ʶ��.Text)
    End If
End Sub

Private Sub txt��ʶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrCurFindtype
            Case "��ʶ��"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "���￨"
                Dim blnCard As Boolean
    
                'ȥ���ſ��������������ַ�
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = InputIsCard(Me.Txt��ʶ��, KeyAscii)
                
                'ˢ����ɻ�ȷ������
                If blnCard And Len(Me.Txt��ʶ��.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.Txt��ʶ��.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.Txt��ʶ��.Text = Me.Txt��ʶ��.Text & Chr(KeyAscii)
                        Me.Txt��ʶ��.SelStart = Len(Me.Txt��ʶ��.Text)
                    End If
                    KeyAscii = 0
                    Me.Txt��ʶ��.Text = UCase(Me.Txt��ʶ��)
                    Me.Txt��ʶ��.SetFocus
                End If
            Case "���ݺ�"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (Txt��ʶ��.Text = "" Or Txt��ʶ��.SelLength = Len(Txt��ʶ��.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "����"
            
        End Select
    End If
End Sub

Private Sub Txt��ʶ��_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txt��ʶ��_Validate(Cancel As Boolean)
    If mstrCurFindtype = "���ݺ�" Then
        If IsNumeric(Txt��ʶ��.Text) Then
            Txt��ʶ��.Text = GetFullNO(Txt��ʶ��.Text, 0)
        End If
    End If
End Sub


Private Sub SeekNextPati(ByVal blnFirst As Boolean)
Dim blnOK As Boolean, l As Long, intB As Integer

    If rptList.FocusedRow Is Nothing Then 'û��¼
        Exit Sub
    ElseIf rptList.FocusedRow.GroupRow Then
        Exit Sub
    End If


    intB = 0
    If Not blnFirst Then
        intB = rptList.SelectedRows(l).Index + 1
        If intB > rptList.Records.Count Then intB = 0
    End If

    blnOK = False
    For l = intB To rptList.Rows.Count - 1 '�ڵ�ǰ״̬�в���
        Select Case mstrCurFindtype
            Case "��ʶ��"
                If Nvl(rptList.Rows(l).Record.Item(mcol.��ʶ��).Value, 0) = Txt��ʶ��.Text Then blnOK = True
            Case "���￨", "�ɣÿ�"
                If Nvl(rptList.Rows(l).Record.Item(mcol.���￨��).Value, "") = Txt��ʶ��.Text Then blnOK = True
            Case "���ݺ�"
                If Nvl(rptList.Rows(l).Record.Item(mcol.NO).Value, "") = Txt��ʶ��.Text Then blnOK = True
            Case "����"
                If Nvl(rptList.Rows(l).Record.Item(mcol.����).Value, "") = Txt��ʶ��.Text Then blnOK = True
            Case "����"
                If Nvl(rptList.Rows(l).Record.Item(mcol.����).Value, "") Like Txt��ʶ��.Text & "*" Then blnOK = True
                If zlCommFun.SpellCode(Nvl(rptList.Rows(l).Record.Item(mcol.����).Value, "")) Like UCase(Txt��ʶ��.Text) & "*" Then blnOK = True
            Case "���֤"
                If Nvl(rptList.Rows(l).Record.Item(mcol.���֤��).Value, "") = Txt��ʶ��.Text Then blnOK = True
        End Select
        
        If blnOK Then
            Txt��ʶ��.Tag = Txt��ʶ��.Text
            If rptList.FocusedRow.Index <> l Then     '�����ǵ�ǰѡ������ѡ��
                rptList.FocusedRow = rptList.Rows(l)
            End If
            rptList.SetFocus
            Exit Sub
        End If
    Next
End Sub

Private Sub Menu_Manage_��������豸()
    Dim strModality As String
    Dim rResult As VbMsgBoxResult
    Dim strSQL As String
    
    If rptList.FocusedRow Is Nothing Then 'û��¼
        Exit Sub
    ElseIf rptList.FocusedRow.GroupRow Then
        Exit Sub
    End If
    
    If UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.Ӱ�����).Value)) = "CR" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.Ӱ�����).Value)) = "DR" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.Ӱ�����).Value)) = "DX" Or _
       UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.Ӱ�����).Value)) = "RF" Then
       
       
       frmChangeDevice.ShowMe UCase(Nvl(rptList.FocusedRow.Record.Item(mcol.Ӱ�����).Value)), Me
       strModality = frmChangeDevice.strDeviceType
       
        If strModality <> "" Then
            strSQL = "Zl_Ӱ����_Ӱ�����(" & rptList.FocusedRow.Record(mcol.ҽ��ID).Value & "," & rptList.FocusedRow.Record(mcol.���ͺ�).Value & ",'" & strModality & "')"
            ExecuteProc strSQL, Me.Caption
        End If
        
        'ˢ�²����б�
        Call RefreshRptlist
    End If
End Sub

Private Sub sub3DProcess(strCommand As String)
    Dim str3DCommand As String
    Dim str3DImgDir As String

    str3DImgDir = App.Path & "\TmpImage\3D\"

    '��֯��ά�ؽ����
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & str3DImgDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub��ά�ؽ�(strCommand As String)

    If TabWindow.Selected.Tag <> "Ӱ��ͼ��" Then '��ˢ��ͼ������
        Call mfrmPACSImg.zlRefresh(rptList.FocusedRow.Record(mcol.ҽ��ID).Value, rptList.FocusedRow.Record(mcol.���ͺ�).Value, mstrPrivs, rptList.FocusedRow.Record(mcol.ת��).Value = 1)
    End If
     
'    Call sub3DProcess("IDLE")
    '��֯��ά�ؽ���Ҫ��ͼ��
    Call mfrmPACSImg.zlMenuClick("��ά�ؽ�")
    Call sub3DProcess(strCommand)
End Sub

