VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPartogram 
   Caption         =   "����ͼ����"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10725
   Icon            =   "frmPartogram.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10725
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   720
         ScaleHeight     =   300
         ScaleWidth      =   1005
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1000
         Begin VB.ComboBox cboBaby 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   0
            Width           =   1000
         End
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   5040
         TabIndex        =   3
         Top             =   1320
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1572864
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   2640
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1572865
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   5055
         TabIndex        =   2
         Top             =   0
         Width           =   5055
         Begin VB.PictureBox PicDraw 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2000
            Left            =   600
            ScaleHeight     =   1995
            ScaleWidth      =   1995
            TabIndex        =   7
            Top             =   480
            Width           =   2000
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6435
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPartogram.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16007
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPartogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************
'���˻�����Ϣ
'***************************************************************
Private Type type_Patient
    lng�ļ�ID As Long
    lng����ID As Long
    lng��ҳID As Long
    lng����ID As Long
    lng�༭ As Long
    lng���� As Long
End Type
Private T_Info As type_Patient

Private mblnChildForm As Boolean
Private mblnShowState As Boolean
Private mblnShowOk As Boolean
Private mstrPrivs As String
Private msinVStep As Single      '�������Ĳ���
Private msinHStep As Single      '�������Ĳ���
Private mblnInit As Boolean
Private mintPage As Integer
Private mbytFontSize  As Byte

'��ȡ��ǰѡ�е�Ӥ����
Public Property Get FileNumIndex() As Long
    FileNumIndex = T_Info.lng����
End Property

Public Property Let FileNumIndex(lngIndex As Long)
    T_Info.lng���� = lngIndex
End Property

Public Property Get FileID() As Long
    FileID = T_Info.lng�ļ�ID
End Property

Public Property Get PatiID() As Long
    PatiID = T_Info.lng����ID
End Property

Public Property Get PageID() As Long
    PageID = T_Info.lng��ҳID
End Property

Public Property Get PartogramParam() As String
    PartogramParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng�༭ & ";" & T_Info.lng����
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property


Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-19 15:16
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = mbytFontSize
    
    Me.FontSize = bytFontSize
    Me.FontName = "����"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    stbThis.Font.Size = bytFontSize
    cboBaby.FontSize = bytFontSize
    picTmp.Height = cboBaby.Height
    cbsThis.RecalcLayout
End Sub

Public Function ShowEdit(ByVal frmMain As Object, strParam As String, Optional ByVal bytMode As Byte = 1, Optional strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
'******************************************************************************************************************
'���ܣ���ɲ���ͼչʾ;����չʾ
'������frmMain��������;strParam:������Ϣ��ʽ��(�ļ�ID;����ID;��ҳID;����ID;�Ƿ�����༭);bytmode(�Ƿ���ģ̬������ʾ);strPrivs;Ȩ��
'���أ����ݳɹ�����True,���򷵻�false
'******************************************************************************************************************
    Dim blnShowSatate As Boolean
    Dim varParam() As String
    
    mstrPrivs = strPrivs
    mblnChildForm = False
    mblnShowOk = False
    blnShowSatate = mblnShowState
    mblnInit = False
    
    If UBound(Split(strParam & ";", ";")) < 5 Then
        MsgBox "���鴫��Ĳ�����ʽ!" & vbCrLf & _
            "��ʽΪ��[�ļ�ID;����ID;��ҳID;����ID;�Ƿ�����༭]", vbInformation, gstrSysName
        Exit Function
    End If
    
    varParam = Split(strParam, ";")
    If blnShowSatate Then
        If Val(varParam(0)) = T_Info.lng�ļ�ID Then
            Call ShowWindow(Me.Hwnd, SW_RESTORE)
            Call BringWindowToTop(Me.Hwnd)
            Exit Function
        End If
    End If
    
    T_Info.lng�ļ�ID = Val(varParam(0))
    T_Info.lng����ID = Val(varParam(1))
    T_Info.lng��ҳID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(3))
    
    If InStr(1, ";" & strPrivs & ";", ";����ͼ��ͼ;") = 0 Then
        T_Info.lng�༭ = 0
    Else
        T_Info.lng�༭ = 1
    End If
    If UBound(varParam) > 3 Then T_Info.lng�༭ = Val(varParam(4))
    T_Info.lng���� = 1
    If UBound(varParam) > 4 Then T_Info.lng���� = Val(varParam(5))
    If T_Info.lng���� < 0 Then T_Info.lng���� = 1
    '��������չʾ
    If Not InitBodyPartogram Then
        Unload Me
        Exit Function
    End If
    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = Not mblnChildForm
    mblnShowState = True
    If blnShowSatate = False Then
        Call SetFontSize(bytSize)
        Me.WindowState = 2
        Set mfrmPartogram = Me
        Hook mfrmPartogram.Hwnd
        If bytMode = 1 Then
            Me.Show 1, frmMain
        Else
            Me.Show , frmMain
        End If
        strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng�༭ & ";" & T_Info.lng����
        ShowEdit = mblnShowOk
    End If
End Function

Public Function zlRefresh(ByVal frmMain As Object, strParam As String, Optional strPrivs As String, Optional blnChildForm As Boolean = True) As Boolean
'******************************************************************************************************************
'���ܣ���ɲ���ͼ����ˢ��;�ؼ�����
'������frmMain��������;strParam:������Ϣ��ʽ��(�ļ�ID;����ID;��ҳID;����ID;�Ƿ�����༭);strPrivs;Ȩ��
'���أ����ݳɹ�����True,���򷵻�false
'******************************************************************************************************************
    Dim blnShowSatate As Boolean
    Dim varParam() As String
    
    mstrPrivs = strPrivs
    mblnChildForm = blnChildForm
    mblnShowState = False
    mblnShowOk = False
    mblnInit = False
    
    If UBound(Split(strParam & ";", ";")) < 5 Then
        MsgBox "���鴫��Ĳ�����ʽ!" & vbCrLf & _
            "��ʽΪ��[�ļ�ID;����ID;��ҳID;����ID;�Ƿ�����༭]", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ȥ�����������
    If mblnChildForm = True Then Call FormSetCaption(Me, False, False)

    stbThis.Visible = Not mblnChildForm
    cbsThis.ActiveMenuBar.Visible = Not mblnChildForm
    
    varParam = Split(strParam, ";")
    
    T_Info.lng�ļ�ID = Val(varParam(0))
    T_Info.lng����ID = Val(varParam(1))
    T_Info.lng��ҳID = Val(varParam(2))
    T_Info.lng����ID = Val(varParam(3))
    
    If InStr(1, ";" & strPrivs & ";", ";����ͼ��ͼ;") = 0 Then
        T_Info.lng�༭ = 0
    Else
        T_Info.lng�༭ = 1
    End If
    If UBound(varParam) > 3 Then T_Info.lng�༭ = Val(varParam(4))
    T_Info.lng���� = 1
    If UBound(varParam) > 4 Then T_Info.lng���� = Val(varParam(5))
    
    '��������չʾ
    If Not InitBodyPartogram Then
        Exit Function
    End If
    zlRefresh = True
End Function

Public Function InitBodyPartogram() As Boolean
 '***************************************************
 '���ܣ���ʼ�����ݣ��Լ���������չʾ
 '***************************************************
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long, i As Integer
    On Error GoTo errHand
    
    mintPage = 1
    gblnMoved = False
    gstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "�ж������Ƿ�ת��", T_Info.lng����ID, T_Info.lng��ҳID)
    gblnMoved = (NVL(rs!����ת��, 0) <> 0)
    
    gstrSQL = "select ��Ŀ���� from �����¼��Ŀ where ������Ŀ=1"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, "�����¼��Ŀ")
    With rs
        Do While Not .EOF
            If InStr(1, "[ZLSOFTLPF]��������[ZLSOFTLPF]��¶�ߵ�[ZLSOFTLPF]����[ZLSOFTLPF]����[ZLSOFTLPF]", "[ZLSOFTLPF]" & NVL(rs!��Ŀ����) & "[ZLSOFTLPF]") <> 0 Then lngCount = lngCount + 1
        .MoveNext
        Loop
    End With
    If lngCount < 4 Then
        MsgBox "���̶̹���Ŀ��ʧ�����ڻ�����Ŀ������м�飡" & vbCrLf & _
            "��Ŀ[����������¶�ߵ͡�����������]", vbInformation, gstrSysName
        Exit Function
    End If
    '��ȡ�ļ�����
    lngCount = GetFileCount(T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID)
    If T_Info.lng���� < 1 Or T_Info.lng���� > lngCount Then T_Info.lng���� = 1
    With cboBaby
        .Clear
        .Tag = 0
        For i = 1 To lngCount
            .AddItem i: .ItemData(.NewIndex) = i
            If i = T_Info.lng���� Then
                .ListIndex = i - 1
            End If
        Next i
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    
    InitBodyPartogram = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboBaby_Click()
'--------------------------------------------------
'��ʼ���в�������չʾ(����ģ�� mdlPrint)
'---------------------------------------------------
    Dim strParam As String
    If T_Info.lng�ļ�ID <= 0 Then Exit Sub
    If Val(cboBaby.Tag) = Val(cboBaby.ItemData(cboBaby.ListIndex)) Then Exit Sub
    T_Info.lng���� = Val(cboBaby.ItemData(cboBaby.ListIndex))
    cboBaby.Tag = T_Info.lng����
    strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng���� & ";" & mintPage
    Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
    mblnInit = True
End Sub

Private Sub Form_Load()
    If Not mblnChildForm Then
        Call RestoreWinState(Me, App.ProductName)
    End If
    Call InitMenuBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShowState = False
    If Not mblnChildForm Then
        If Not (mfrmPartogram Is Nothing) Then UnHook mfrmPartogram.Hwnd
        Call SaveWinState(Me, App.ProductName)
    End If
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom  As CommandBarControlCustom
    
    On Error GoTo errHand
    PicDraw.AutoRedraw = True
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&E)")
                
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
               
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        cbrControl.BeginGroup = True
    End With


    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "���ݱ༭(&E)")
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    With cbrMenuBar.CommandBar.Controls
                
'       Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
'
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
'       cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
                
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)..."):
        cbrControl.BeginGroup = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0

    With cbrToolBar.Controls
        .Add xtpControlLabel, conMenu_View_Option, "Ӥ��"
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Option, "")
        cbrCustom.Flags = xtpFlagAlignLeft
        picTmp.Visible = True
        cbrCustom.Handle = picTmp.Hwnd
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Billing, "����"): cbrControl.ToolTipText = "�������ݱ༭": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "��ҳ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��ҳ"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "��ҳ"):   cbrControl.ToolTipText = "��ҳ"
    End With

    '��λ������
    '------------------------------------------------------------------------------------------------------------------

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("E"), conMenu_Edit_Billing
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey As String
    Dim lngLoop As Long
    Dim strSQL() As String, strParam As String
    Dim blnTran As Boolean
    Dim lngIndex As Long
    Dim cbrControl As CommandBarControl
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
        
    Select Case Control.ID
        Case conMenu_File_PrintSet   '��ӡ����
            
            On Error Resume Next
            Call frmPrintSet.ShowMe(Me, 1)
            
        Case conMenu_File_Preview  '��ӡԤ��
            
            Call PrintData(1)
            
        Case conMenu_File_Print  '��ӡ
        
            Call PrintData(2)
            
            
        Case conMenu_View_StatusBar
        
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_Edit_Billing '���ݱ༭
            If frmPartogramEditor.ShowMe(Me, T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, 0, mstrPrivs, (T_Info.lng�༭ = 1), IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))) = True Then
                '��������ˢ��
                If Not InitBodyPartogram Then Exit Sub
                strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng���� & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
                mblnShowOk = Not mblnChildForm
            End If
        Case conMenu_Edit_PrevPage '��ҳ
            If mintPage > 1 Then
                mintPage = mintPage - 1
                strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng���� & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
            End If
        Case conMenu_Edit_NextPage '��ҳ
            If mintPage < mintMaxPage Then
                mintPage = mintPage + 1
                strParam = T_Info.lng�ļ�ID & ";" & T_Info.lng����ID & ";" & T_Info.lng��ҳID & ";" & T_Info.lng����ID & ";" & T_Info.lng���� & ";" & mintPage
                Call PreViewOrPrintPartogram(strParam, PicDraw, Me)
            End If
        Case conMenu_Help_Help
        
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_About
            
            Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            
            Call zlHomePage(Me.Hwnd)
            
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.Hwnd)
            
        Case conMenu_Help_Web_Mail
            
            Call zlMailTo(Me.Hwnd)
        
        Case conMenu_File_Exit
            Unload Me
            Exit Sub
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup
        End Select
    End If
    
    Err = 0
    On Error Resume Next
    
    Select Case Control.ID

    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Billing '���ݱ༭
        Control.Enabled = (T_Info.lng�༭ = 1)
        
    Case conMenu_View_StatusBar
    
        Control.Checked = Me.stbThis.Visible
    Case conMenu_Edit_PrevPage '��һҳ
        Control.Enabled = (mintPage > 1)
    Case conMenu_Edit_NextPage '��һҳ
        Control.Enabled = (mintPage < mintMaxPage)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    With picBack
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub picBack_Resize()
    With vsb
        .Left = picBack.Width - .Width
        .Top = 0
        .Height = IIf(picBack.Height - hsb.Height < 0, 0, picBack.Height - hsb.Height)
    End With
    
    With hsb
        .Left = 0
        .Top = picBack.Height - .Height
        .Width = IIf(picBack.Width - vsb.Width < 0, 0, picBack.Width - vsb.Width)
    End With
    
    With picMain
        .Left = 0
        .Top = 0
        .Height = IIf(hsb.Visible = True, hsb.Top, picBack.Height)
        .Width = IIf(vsb.Visible = True, vsb.Left, picBack.Width)
    End With
End Sub

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ� ���óɹ�����TRUE������FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (PicDraw.Width - picMain.Width) / 100
    msinVStep = (PicDraw.Height - picMain.Height) / 100

    
    hsb.Max = 0 - Int(0 - ((PicDraw.Width - picMain.Width) / 300)) - 1
    vsb.Max = 0 - Int(0 - ((PicDraw.Height - picMain.Height) / 300)) - 1
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    
    With vsb
        .Height = picMain.Height - IIf(hsb.Visible = True, hsb.Height, 0)
    End With
    
    With hsb
        .Width = picMain.Width - IIf(vsb.Visible = True, vsb.Width, 0)
    End With
    picMain.Move 0, 0, IIf(vsb.Visible = True, vsb.Left, picBack.Width), IIf(hsb.Visible = True, hsb.Top, picBack.Height)
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
    If hsb.Enabled Then
        hsb.Max = 100
        hsb.LargeChange = 100 / Int((Round((PicDraw.Width - picMain.Width) / picMain.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 100
        vsb.LargeChange = 100 / Int((Round((PicDraw.Height - picMain.Height) / picMain.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
    
End Function

Private Sub picMain_Paint()
    CalcScrollBarSize
End Sub

Private Sub picMain_Resize()
    picMain.BackColor = PicDraw.BackColor
    With PicDraw
        .Left = 0
        .Top = 0
'        .Width = IIf(.Width < picMain.Width, picMain.Width, .Width)
'        .Height = IIf(.Height < picMain.Height, picMain.Height, .Height)
    End With
    CalcScrollBarSize
End Sub

Private Sub vsb_Change()
    PicDraw.Top = -1 * vsb.Value * msinVStep
End Sub

Private Sub hsb_Change()
    PicDraw.Left = -1 * hsb.Value * msinHStep
End Sub

Public Sub PrintData(ByVal bytMode As Byte, Optional ByVal strPrintDevice As String = "")
    Dim bytOp As Byte
    Dim lngFileIndex As Long, lngFilePage As Long
    Dim blnPrint As Boolean
    
    If strPrintDevice = "" Then
        lngFileIndex = T_Info.lng����
        lngFilePage = mintPage
        '��ӡѡ���
        bytOp = frmPartogramPrintSet.PrintSet(Me, bytMode, cboBaby.ListCount, lngFileIndex, lngFilePage)
    Else
        lngFileIndex = -1
        lngFilePage = -1
    End If
    If bytOp = 0 Then Exit Sub
    '��ʼ����Ԥ����ӡ
    blnPrint = (bytOp = 2)
    
    Call ShowPrintPartogram(Me, T_Info.lng�ļ�ID, T_Info.lng����ID, T_Info.lng��ҳID, T_Info.lng����ID, lngFileIndex, lngFilePage, blnPrint, strPrintDevice)
End Sub
