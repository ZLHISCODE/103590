VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucPacsHelper 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   LockControls    =   -1  'True
   ScaleHeight     =   10020
   ScaleWidth      =   6885
   Begin VB.CommandButton cmdAttach 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdDel 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":047A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   375
      Left            =   6120
      Picture         =   "ucPacsHelper.ctx":08F4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
      Width           =   375
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox cmdMenu 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   6120
      ScaleHeight     =   405
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   3960
      Width           =   495
      Begin VB.Image imgMenu 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   120
         Picture         =   "ucPacsHelper.ctx":0D6E
         Top             =   30
         Width           =   270
      End
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   360
      ScaleHeight     =   9495
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   2535
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   238
         BackColor       =   12632256
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "picVideoContainer"
         Control2Name    =   "picHelperContainer"
      End
      Begin VB.PictureBox picVideoContainer 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2535
         ScaleWidth      =   5415
         TabIndex        =   4
         Top             =   0
         Width           =   5415
      End
      Begin VB.PictureBox picHelperContainer 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   6825
         Left            =   0
         ScaleHeight     =   6825
         ScaleWidth      =   5415
         TabIndex        =   2
         Top             =   2670
         Width           =   5415
         Begin zl9PACSWork.ucCacheImages ucCache 
            Height          =   4935
            Left            =   1800
            TabIndex        =   7
            Top             =   1320
            Visible         =   0   'False
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   8705
         End
         Begin zl9PACSWork.ucReportHistory ucHistory 
            Height          =   4695
            Left            =   1320
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   8281
         End
         Begin zl9PACSWork.ucBgImgViewer ucImages 
            Height          =   4575
            Left            =   840
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   8070
         End
         Begin XtremeSuiteControls.TabControl tabSelect 
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1815
            _Version        =   589884
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   64
         End
         Begin zl9PACSWork.ucReportSegment ucWord 
            Height          =   4335
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Visible         =   0   'False
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   7646
         End
         Begin VB.PictureBox picTemp 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   4215
            Left            =   120
            ScaleHeight     =   4215
            ScaleWidth      =   3015
            TabIndex        =   8
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   6120
      Top             =   2160
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "ucPacsHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const C_MODULE_NAME = "ucPacsHelper"


Private Const CON_TAB_TAG_ͼ�� As String = "ͼ��"
Private Const CON_TAB_TAG_����ͼ As String = "����ͼ"
Private Const CON_TAB_TAG_�ʾ� As String = "�ʾ�"
Private Const CON_TAB_TAG_��ʷ As String = "��ʷ"
Private Const CON_TAB_TAG_���� As String = "����"



Private Const conMenu_Helper_Refresh = 8140         'ˢ��


'ͼ�����
Private Const conMenu_Helper_AddToReport = 8141     '���뱨��ͼ
Private Const conMenu_Helper_ImageProcess = 8142    'ͼ����
Private Const conMenu_Helper_BigImageShow_Move = 8143    '�ƶ���ʾ��ͼ
Private Const conMenu_Helper_BigImageShow_Click = 8144    '������ʾ��ͼ
Private Const conMenu_Helper_BigImageShow_Delay = 8145    '�ӳٹرմ�ͼ

Private Const conMenu_Helper_SelAll = 8146          'ȫѡ
Private Const conMenu_Helper_DelOper = 8147         'ɾ������

Private Const conMenu_Helper_Import = 8148          '����
Private Const conMenu_Helper_Export = 8149          '����

Private Const conMenu_Helper_SendStudy = 8150       '���͵����
Private Const conMenu_Helper_SendCache = 8151       '���͵�����

Private Const conMenu_Helper_SplitPage = 8152
Private Const conMenu_Helper_ReDo = 8153        '���³���   ��ֻ��Դ���ʧ�ܵ�ͼ��
Private Const conMenu_Helper_ReDown = 8154      '��������
Private Const conMenu_Helper_ReUp = 8155        '�����ϴ�

Private Const conMenu_Helper_OpenImgPos = 8156     '��ͼ��λ��


'�����ʷ���
Private Const conMenu_Helper_ImgViewer = 8157           'Ӱ���Ƭ(&S)
Private Const conMenu_Helper_ImgContrast = 8158         '��Ƭ�Ա�(&E)
Private Const conMenu_Helper_ReportOpen = 8159          '�����
Private Const conMenu_Helper_Analysis = 8160      '�ۺϷ���

Private Const conMenu_Helper_ViewReportImage = 8161          '�鿴����ͼ
Private Const conMenu_Helper_ViewReportContext = 8162        '�鿴��������
Private Const conMenu_Helper_WriteReport = 8163         'д�뱨��
Private Const conMenu_Helper_LinkViewer = 8164        '�����鿴
Private Const conMenu_Helper_CloseViewer = 8165

Private Const conMenu_Helper_RelateCfg = 8166          '�������
Private Const conMenu_Helper_ThisTime = 8167            '�������
Private Const conMenu_Helper_OtherDept = 8168           '���Ƽ��
Private Const conMenu_Helper_AutoLine = 8169            '�Զ�����

Private Const conMenu_Helper_DateRange = 8170          '���ڷ�Χ
Private Const conMenu_Helper_OneMonth = 8171            'һ����
Private Const conMenu_Helper_TwoMonth = 8172            '������
Private Const conMenu_Helper_ThreeMonth = 8173          '������
Private Const conMenu_Helper_HalfYear = 8174            '����
Private Const conMenu_Helper_OneYear = 8175             'һ��
Private Const conMenu_Helper_TwoYear = 8176             '����
Private Const conMenu_Helper_ThreeYear = 8177           '����
Private Const conMenu_Helper_DateUn = 8178              '��������
Private Const conMenu_Helper_DateCus = 8179             '�Զ�����


'����ʾ����
Private Const conMenu_Helper_DirectWrite = 8180         'ֱ��д��
Private Const conMenu_Helper_EditWrite = 8181           '�༭д��

Private Const conMenu_Helper_FullSave = 8182            'ȫ��д��
Private Const conMenu_Helper_NewWord = 8183           '�����ʾ�
Private Const conMenu_Helper_ModWord = 8184             '�޸Ĵʾ�
Private Const conMenu_Helper_DelWord = 8185             'ɾ���ʾ�

Private Const conMenu_Helper_AutoHide = 8186            '�Զ�����
Private Const conMenu_Helper_DblWrite = 8187
Private Const conMenu_Helper_ExpandLevel = 8188        'չ���㼶
Private Const conMenu_Helper_OneLevel = 8189         'һ��
Private Const conMenu_Helper_TwoLevel = 8190           '����
Private Const conMenu_Helper_ThreeLevel = 8191         '����
Private Const conMenu_Helper_AllLevel = 8192           '����


Private Const conMenu_Helper_Log = 8300    '��־
 

Private Type DlgFileInfo
    iCount As Long
    sPath As String
    sFIle() As String
End Type


Private mObjNotify As IEventNotify
Private mobjEmbedVideo As Object
Private mobjLinkEditor As ucReportEditor
Private mobjSel As Object


Private WithEvents mobjImageProcessV2 As frmImageProcessV2
Attribute mobjImageProcessV2.VB_VarHelpID = -1

Private mblnIsEmbedVideoArea As Boolean     '�Ƿ�Ƕ����Ƶ�ɼ�����
Private mblnImgAscOrder As Boolean
Private mblnAllowEmbedVideo As Boolean  '�Ƿ�����Ƕ����Ƶ�ɼ�

Private mlngModuleNo As Long
Private mlngDeptID As Long
Private mstrPrivs As String
Private mstrGrantDeptIds As String
Private mstrParentName As String
 
Private mobjStudyInfo As clsStudyInfo
Private mlngReleationImgAdvice As Long

Private mlngFileID As Long
Private mblnBgImgTrans As Boolean         '��̨ͼ����
Private mblnIsTabIniting As Boolean
Private mblnIsValid As Boolean
Private mobjMainVideo As Object

Private mblnMoveBigImageShow As Boolean
Private mblnClickBigImageShow As Boolean
Private mblnDelayCloseImage As Boolean
Private mlngBigImageIndex As Long
Private mlngStartMoveTime As Long

Private mblnIgnoreResult As Boolean
    
Private mstrReportImageUids As String       '������Ϊ����ͼ��ͼ��UID
Private mblnAllowWrite As Boolean
Private mblnIsProcessing As Boolean
Private mlngReleationImgDays As Long
Private mlngImageDBClickOper As Long    'ͼ��˫��������ʽ

'�ʾ�����¼�
Public Event OnWordRequestState(ByRef lngOutlineType As TOutlineType, _
                            ByRef str�������� As String, ByRef str������� As String, ByRef str�������� As String)
    
Public Event OnWordSendContext(ByVal strFreeText As String, _
                            ByVal str�������� As String, ByVal str������� As String, ByVal str�������� As String)
                                       
Public Event OnTabChanged(ByVal strTabName As String)

Public Event OnLinkHistoryView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)
                            
Public Event OnDockHideClick()
Public Event OnDockAttachClick()



Property Get IsSyncWordFragment() As Boolean
    IsSyncWordFragment = ucWord.IsSyncWordFragment
End Property


Property Get ImgCount() As Long
    ImgCount = ucImages.ImgCount
End Property

Property Get EmbedVideo() As Object
    Set EmbedVideo = mobjEmbedVideo
End Property
                            
Property Get Processing() As Boolean
    Processing = mblnIsProcessing
End Property

Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property


Property Get HideButtonEnable() As Boolean
    HideButtonEnable = cmdHide.Enabled
End Property

Property Let HideButtonEnable(ByVal value As Boolean)
    cmdHide.Enabled = value
End Property



Property Get IsValid() As Boolean
    IsValid = mblnIsValid
End Property
 
 
Property Get SelTabName() As String
    SelTabName = tabSelect.Selected.tag
End Property


'���ӵı���༭��
Property Get LinkEditor() As Object
    Set LinkEditor = mobjLinkEditor
End Property

Property Set LinkEditor(value As Object)
    Set mobjLinkEditor = value
End Property


'����Ƶ����
Property Get MainVideoWindow() As Object
    Set MainVideoWindow = mobjMainVideo
End Property

Property Set MainVideoWindow(value As Object)
    Set mobjMainVideo = value
End Property


Property Get AllowLinkerViewer() As Boolean
    AllowLinkerViewer = ucHistory.AllowLinkViewer
End Property


Property Let AllowLinkerViewer(ByVal value As Boolean)
    ucHistory.AllowLinkViewer = value
End Property


Property Get TabEnable(ByVal strTabName As String) As Boolean
    Dim i As Long
    
    For i = 1 To tabSelect.ItemCount
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            TabEnable = tabSelect(i).Enabled
            Exit Property
        End If
    Next
End Property

Property Let TabEnable(ByVal strTabName As String, ByVal value As Boolean)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            tabSelect(i).Enabled = value
            
            If value = False And tabSelect(i).Selected Then
                For j = 0 To tabSelect.ItemCount - 1
                    If tabSelect(j).Enabled Then
                        tabSelect(j).Selected = True
                        Exit Property
                    End If
                Next
            End If
            
            Exit Property
        End If
    Next
End Property


Property Get TabVisible(ByVal strTabName As String) As Boolean
    Dim i As Long
    
    For i = 1 To tabSelect.ItemCount
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            TabVisible = tabSelect(i).Visible
            Exit Property
        End If
    Next
End Property

Property Let TabVisible(ByVal strTabName As String, ByVal value As Boolean)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            tabSelect(i).Visible = value
            
            If value = False And tabSelect(i).Selected Then
                For j = 0 To tabSelect.ItemCount - 1
                    If tabSelect(j).Visible Then
                        tabSelect(j).Selected = True
                        Exit Property
                    End If
                Next
            End If
            
            Exit Property
        End If
    Next
End Property

'�Ƿ�����д��
Property Get AllowWrite() As Boolean
    AllowWrite = mblnAllowWrite
End Property

Property Let AllowWrite(ByVal value As Boolean)
    mblnAllowWrite = value
    
    If mblnAllowWrite = False Then
        cmdAdd.Enabled = mblnAllowWrite
    Else
        cmdAdd.Enabled = IIf(tabSelect.Selected.tag <> CON_TAB_TAG_����, True, False)
    End If
    
    ucHistory.AllowWrite = value
End Property



'ҽ��ID
Property Get AdviceId() As Long
        AdviceId = mobjStudyInfo.lngAdviceId
End Property

'Property Let AdviceId(ByVal value As Long)
'    mlngAdviceID = value
'End Property



'�Ƿ�Ƕ����Ƶ�ɼ�
Property Get IsEmbedVideoArea() As Boolean
    IsEmbedVideoArea = mblnIsEmbedVideoArea
End Property


'�Ƿ�����Ƕ����Ƶ�ɼ�
Property Get AllowEmbedVideo() As Boolean
    AllowEmbedVideo = mblnAllowEmbedVideo
End Property

Property Let AllowEmbedVideo(ByVal value As Boolean)
    mblnAllowEmbedVideo = value
End Property


Property Get IsStudying() As Boolean
    IsStudying = IIf(mobjStudyInfo.intStep < 6 And mobjStudyInfo.intStep > 1, True, False)
End Property


Private Function HintError(objErr As ErrObject, ByVal strMethodName As String, _
    Optional ByVal blnIsDataErr As Boolean = True) As Long
    If mObjNotify Is Nothing Then Exit Function
    
    If blnIsDataErr Then
        HintError = mObjNotify.PrintErr(objErr, infDataErr, GetRootHwnd, C_MODULE_NAME, strMethodName)
    Else
        HintError = mObjNotify.PrintErr(objErr, infNormalErr, GetRootHwnd, C_MODULE_NAME, strMethodName)
    End If
End Function

Private Function HintMsg(ByVal strMsg As String, ByVal strMethodName As String, _
    Optional ByVal lngMsgType As Long = infHint) As Long
        HintMsg = mObjNotify.PrintInfo(strMsg, lngMsgType, GetRootHwnd, C_MODULE_NAME, strMethodName)
End Function


Public Sub HideEmbedVideo()
'����Ƕ��ʽ��Ƶ�ɼ�����
    mblnIsEmbedVideoArea = False
    
    picVideoContainer.Visible = mblnIsEmbedVideoArea
    ucSplitter1.Visible = mblnIsEmbedVideoArea
    
    '����Ѿ�������Ƶ�ɼ�,�������Ҫ����Ƶ�ɼ��ָ���������
    Call picBack_Resize
End Sub


Public Function ShowEmbedVideo(objCapLinker As Object, Optional ByVal blnIsForce As Boolean = False) As Boolean
'Ƕ����Ƶ
'blnIsForce:�Ƿ�ǿ��Ƕ����Ƶ�ɼ������ж���Ƶ���ڵĸ������Ƿ���ͬ,��Ҫ����Ӱ��ɼ��ͼ�鱨��ģ��ҳ֮ǰ����Ƶ�л�
    Dim objCapHelper As ICapHelper
    Dim blnAfterOrLock As Boolean
    
    ShowEmbedVideo = False
    mblnIsEmbedVideoArea = False
    blnAfterOrLock = False
    
     
    '���������Ƕ����Ƶ�ɼ������˳�
    If mblnAllowEmbedVideo = False Then
        Call HideEmbedVideo
        Exit Function
    End If
    
    If Not mobjEmbedVideo Is Nothing Then
        blnAfterOrLock = IIf(mobjEmbedVideo.isLock Or mobjEmbedVideo.IsAfter, True, False)
    End If
    
    Set objCapHelper = objCapLinker
    
'    If objCapHelper.IsAllowCapture = False And blnAfterOrLock = False Then
'        '��ǰ״̬������ɼ�����û�п�����̨�������ɼ�����£�����Ƕ��ʽ��Ƶ����
'        If picVideoContainer.Visible = False Then Exit Function
'        Call HideEmbedVideo
'        Exit Function
'    Else
        '����Ѿ���ʾ��Ƕ��ʽ��Ƶ���ڣ���ֱ���˳�
        If Not mobjEmbedVideo Is Nothing Then
            If picVideoContainer.Visible And (GetAncestor(GetAncestor(mobjEmbedVideo.VideoHwnd, GA_PARENT), GA_PARENT) = picVideoContainer.hwnd) Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False, True)
                
                mblnIsEmbedVideoArea = True
                ShowEmbedVideo = True
                Exit Function
            End If
        Else
            Set mobjEmbedVideo = New clsPacsCaptureV2
            Call mobjEmbedVideo.zlInitModule(gcnOracle, objCapLinker, glngSys, mlngModuleNo, mstrPrivs, mlngDeptID, hwnd, True)
        End If

    '�ж��Ƿ񵯳��˶�������Ƶ�ɼ����ڣ����������������Ƕ����Ƶ
    If mobjEmbedVideo.VideoDockState Then
        Call HideEmbedVideo
        Exit Function
    End If
    
    mblnIsEmbedVideoArea = True
    
    picVideoContainer.Visible = True
    ucSplitter1.Visible = True
    
    Call picBack_Resize
    Call picVideoContainer_Resize

    
    '�������������ͬ���ڣ�����Ҫ��������Ƕ��
    If blnIsForce = False And GetAncestor(mobjEmbedVideo.VideoHwnd, GA_ROOT) = GetAncestor(picVideoContainer.hwnd, GA_ROOT) Then
        If GetAncestor(mobjEmbedVideo.ContainerHwnd, GA_ROOT) = GetAncestor(picVideoContainer.hwnd, GA_ROOT) Then
            '�����Ƿ�ֻ��״̬
            If mobjEmbedVideo.ContainerHwnd = GetAncestor(mobjEmbedVideo.VideoHwnd, GA_PARENT) Then
               
            
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False, True)
                ShowEmbedVideo = True
                Exit Function
            End If
        End If
    End If
    
    SetParent mobjEmbedVideo.ContainerHwnd, picVideoContainer.hwnd
    
    ShowWindow mobjEmbedVideo.ContainerHwnd, 1
    
    Call mobjEmbedVideo.zlRefreshVideoWindow
    
    If mobjStudyInfo Is Nothing Then
        '��ʼ�������е���Ƕ����Ƶ��ʾʱ��mobjStudyInfoΪnothing
        Call mobjEmbedVideo.zlRestoreWindow(True, False)
    Else
        Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
    End If
    ShowEmbedVideo = True
End Function


Public Sub LocateTab(ByVal strTabName As String)
'��λָ����tabҳ
    Dim i As Long
    
    For i = 0 To tabSelect.ItemCount - 1
        If UCase(tabSelect(i).tag) = UCase(strTabName) Then
            If tabSelect(i).Visible Then tabSelect(i).Selected = True
            Exit Sub
        End If
    Next
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    FontSize = bytFontSize
    
    picBack.FontSize = bytFontSize
    picVideoContainer.FontSize = bytFontSize
    picHelperContainer.FontSize = bytFontSize
    picTemp.FontSize = bytFontSize
    
    Call ucWord.SetFontSize(bytFontSize)
    Call ucHistory.SetFontSize(bytFontSize)
    Call ucCache.SetFontSize(bytFontSize)
    
    Set tabSelect.PaintManager.Font = Font
    
    '����ı����Ҫʹ�ø�����ˢ�½�����ʾ
    tabSelect.PaintManager.Layout = xtpTabLayoutAutoSize
End Sub

Public Sub Init(objNotify As IEventNotify, ByVal lngMudleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, Optional ByVal blnIsForce As Boolean = False)
'��ʼ��
'lngMudleNo:ģ���
'lngDeptId����ǰ����ID
'strGrantDepts:��Ȩ����ID
'
    Set mObjNotify = objNotify
    
    mlngModuleNo = lngMudleNo
    mlngDeptID = lngDeptId
    mstrPrivs = strPrivs
    mstrParentName = Parent.Name
    
    Call InitTab
     
    Call InitPar
    
    Call ucImages.Init
    Call ucWord.Init(lngMudleNo, lngDeptId, blnIsForce)
    Call ucHistory.Init(lngMudleNo, lngDeptId, mstrPrivs, blnIsForce)
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeWhidbey

    mblnIsValid = True
End Sub



Public Function GetFileFormatId(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As Long
'��ȡ����Ӧ�����Ƶ��ݸ�ʽID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetFileFormatId = 0
    
    strSQL = "Select l.������Դ, a.�����ļ�id" & vbNewLine & _
            " From ����ҽ����¼ l, ��������Ӧ�� a" & vbNewLine & _
            " Where l.������Ŀid = a.������Ŀid(+) And a.Ӧ�ó���(+) = Decode(l.������Դ, 2, 2, 4 ,4, 1) And l.Id = [1]"
            
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ݸ�ʽ", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetFileFormatId = Val(nvl(rsData!�����ļ�id))
    
End Function


Public Sub zlRefresh(objStudyInfo As clsStudyInfo, ByVal lngFileId As Long, Optional ByVal blnIsForceRefresh As Boolean = False)
'ˢ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errhandle
    
    If mobjSel Is Nothing Then Exit Sub
    
    If Not objStudyInfo Is Nothing And Not mobjStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    mblnIsProcessing = True
    
    If Not mobjImageProcessV2 Is Nothing Then
        '�ж����ڴ����ͼ���Ƿ񱣴�
        Unload mobjImageProcessV2
    End If
    
    Set mobjStudyInfo = objStudyInfo
     
    mlngFileID = lngFileId
    mlngReleationImgAdvice = 0
    
    If mlngFileID = 0 Then mlngFileID = GetFileFormatId(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    Call RefreshHelperComponent(mobjStudyInfo.lngAdviceId, mobjStudyInfo.strStudyUID, mlngFileID, mobjStudyInfo.blnMoved)
    
    If mblnIsEmbedVideoArea Then
        If Not mobjEmbedVideo Is Nothing Then
            Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
        End If
    End If
    
    mblnIsProcessing = False
Exit Sub
errhandle:
    mblnIsProcessing = False
    If HintError(err, "zlRefresh", False) = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshHelperComponent(ByVal lngAdviceId As Long, ByVal strSutdyUid As String, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
'ˢ�°������
    If mobjSel Is Nothing Then Exit Sub
    Select Case mobjSel.hwnd
        Case ucImages.hwnd  'ͼ��
             ucImages.ClearAll
             
            '�жϼ���Ƿ��Ѿ���ͼ��
            If Len(strSutdyUid) > 0 Then Call LoadExamImages(lngAdviceId, blnMoved)
            
            If ucImages.ImgCount <= 0 Then
                If mlngReleationImgDays > 0 Then
                    '�ж��Ƿ���й���ͼ���ȡ
                    If mlngReleationImgAdvice = 0 Then
                        mlngReleationImgAdvice = GetReleationImageAdvice(lngAdviceId)
                    End If
                    
                    If mlngReleationImgAdvice > 0 Then
                        Call LoadExamImages(mlngReleationImgAdvice, False)
                    End If
                Else
                    mstrReportImageUids = ""
                    Call ucImages.ClearAll
                End If
            End If
            
        Case ucWord.hwnd    '�ʾ�
            Call ucWord.Refresh(lngAdviceId, lngFileId)

        Case ucHistory.hwnd '��ʷ
            ucHistory.AllowWrite = mblnAllowWrite And (IsStudying Or (CheckPopedom(mstrPrivs, "��¼����") And mobjStudyInfo.intStep > 5 And mobjStudyInfo.strReportDoctor = ""))
            Call ucHistory.Refresh(lngAdviceId)
            
        Case ucCache.hwnd   '����
            Call ucCache.Refresh
            
    End Select
End Sub

Private Function GetReleationImageAdvice(ByVal lngAdviceId As Long) As Long
'�򿪹������߼��ͼ��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsReturn As ADODB.Recordset
    Dim dtStudy As Date
    
    GetReleationImageAdvice = 0
    
    strSQL = "select a.ִ�в���id, a.����ʱ��, b.����id,c.����id,c.Ӱ����� " & _
            " from ����ҽ������ a, ����ҽ����¼ b,Ӱ�����¼ c" & _
            " where a.ҽ��id=b.id and a.ҽ��id=c.ҽ��id and a.���ͺ�=c.���ͺ� and a.ҽ��id=[1]"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������Ϣ", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '����ͬ�����£����Ӱ�������ͬ�ļ�飬��Ҫ���ͬһ���߲�ͬҽ����ͬ���Ĳ�ͬ��λ��ͬһ�豸�����������
    strSQL = "Select Distinct * from (" & vbCrLf & _
            " select b.ҽ��ID, b.����,b.����,b.�Ա�,b.����,b.Ӱ�����, a.ҽ������, b.���uid,b.λ��һ,b.λ�ö� " & _
            " from ����ҽ����¼ a, Ӱ�����¼ b" & vbCrLf & _
            " Where a.ID = b.ҽ��ID And a.����ID = [1] And b.ִ�п���id = [3]" & vbCrLf & _
            "       and b.�������� between [4] and [5] " & vbCrLf & _
            "       and b.���UID is not null and b.ҽ��ID<>[6] and b.Ӱ�����=[7] " & vbCrLf & _
            " Union All " & vbCrLf & _
            " select a.ҽ��ID, a.����,a.����,a.�Ա�,a.����,a.Ӱ�����, b.ҽ������,a.���uid,a.λ��һ,a.λ�ö� " & _
            " from Ӱ�����¼ a , ����ҽ����¼ b " & vbCrLf & _
            " Where a.ҽ��ID=b.Id and a.����id = [2] And a.ִ�п���id = [3]" & vbCrLf & _
            "   and a.�������� between [4] and [5] " & vbCrLf & vbCrLf & _
            " and a.���UID is not null and a.ҽ��ID<>[6] and a.Ӱ�����=[7] " & vbCrLf & _
            " )"
            
    dtStudy = CDate(Format(nvl(rsData!����ʱ��, 0), "yyyy-mm-dd 00:00"))
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����������", _
                            Val(nvl(rsData!����ID)), Val(nvl(rsData!����ID)), Val(nvl(rsData!ִ�в���ID)), _
                            CDate(Format(dtStudy - mlngReleationImgDays, "yyyy-mm-dd 00:00:00")), _
                            CDate(Format(dtStudy + mlngReleationImgDays, "yyyy-mm-dd 23:59:59")), _
                            lngAdviceId, nvl(rsData!Ӱ�����)) '����7��ʾ��ѯ��������Χ
                            
    If rsData.RecordCount <= 0 Then Exit Function
    
    If rsData.RecordCount = 1 Then
        If HintMsg("��ǰ�����δ�������ͼ���Ƿ�����µ���������ȡ?" & vbCrLf & _
                                "    ���ţ�" & nvl(rsData!����) & vbCrLf & _
                                "    ������" & nvl(rsData!����) & vbCrLf & _
                                "    �Ա�" & nvl(rsData!�Ա�) & vbCrLf & _
                                "    ���䣺" & nvl(rsData!����) & vbCrLf & _
                                "    " & nvl(rsData!ҽ������), "GetReleationImageAdvice", vbYesNo) = vbNo Then
            GetReleationImageAdvice = -1
            Exit Function
        End If
        
        GetReleationImageAdvice = Val(nvl(rsData!ҽ��ID))
    Else
        If HintMsg("��ǰ�����δ�������ͼ���Ƿ������� [" & rsData.RecordCount & "] ����ؼ������ȡͼ��", "GetReleationImageAdvice", vbYesNo) = vbNo Then
            GetReleationImageAdvice = -1
            Exit Function
        End If

        If FS.ShowRecSelect(mObjNotify.Owner, cmdMenu, rsData, rsReturn, True, "ҽ��ID,λ��һ,λ�ö�,���UID") Then
            GetReleationImageAdvice = Val(nvl(rsReturn!ҽ��ID))
        End If
    End If
    
End Function

Public Sub ClearReportImgState(Optional ByVal strImgKey As String = "")
'�������ͼ״̬
    If Len(strImgKey) > 0 Then
        Call ucImages.ImgDrawHint(strImgKey, "", "��")
    Else
        Call ucImages.ClearDrawHint("��")
    End If
End Sub
 

Public Sub SyncReportImgState(ByVal strImgs As String)
'ͬ������ͼ״̬
    Call ucImages.SyncDrawHint(strImgs, "��", "��")
    mstrReportImageUids = strImgs
End Sub

Public Function GetLayoutStr() As String
'���ظ�ʽ�ַ���[Key=picturebox1.width:20;picturebox1.height:30;]
    Dim strPros As String
    strPros = "[KEY=HELPER@" & _
                        GetProFmt("PICVIDEO.HEIGHT", picVideoContainer.Height) & _
                        GetProFmt("PICHELPER.HEIGHT", picHelperContainer.Height) & _
                        ";]"
                        
    GetLayoutStr = strPros & ucWord.GetLayoutStr & ucHistory.GetLayoutStr
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String

    If Len(strLayout) <= 0 Then Exit Sub
    
    Call ucWord.SetLayout(strLayout)
    Call ucHistory.SetLayout(strLayout)
    
    
    strPros = GetPros(strLayout, "HELPER")
    
    strPro = GetProValue(strPros, "PICVIDEO.HEIGHT")
    If Val(strPro) > 0 Then picVideoContainer.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "PICHELPER.HEIGHT")
    If Val(strPro) > 0 Then picHelperContainer.Height = Val(strPro)
End Sub

Private Sub LoadExamImages(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean)
'������ͼ��
    Dim objImgInf As New clsBgImgInfo
    Dim strStudyUID As String
    Dim rsData As ADODB.Recordset
    Dim strLocalPath As String
    Dim i As Long
    Dim strResult As String
    
    mstrReportImageUids = ""
    
    ucImages.ClearAll
    
    If lngAdviceId <= 0 Then
        Exit Sub
    End If
    
    objImgInf.PatientName = mobjStudyInfo.strPatientName
    objImgInf.ImgCommand = icDownload
    objImgInf.AdviceId = lngAdviceId
    objImgInf.Format = ifDcm
    
    
    strResult = ResetStorageDevice(lngAdviceId, objImgInf, blnMoved)
    
    If Len(strResult) > 0 Then
        HintMsg strResult, "ResetStorageDevice", vbOKOnly
        Exit Sub
    End If
    
    strStudyUID = objImgInf.StudyUID
    If Len(strStudyUID) <= 0 Then Exit Sub
    
    Set rsData = GetExamImgData(strStudyUID, blnMoved)
    
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount <= 0 Then Exit Sub
    
    '��������ͼ��
    Call ucImages.EraseImgData
    
    i = 1
    While Not rsData.EOF
        objImgInf.Key = nvl(rsData!ͼ��UID)
        
        '���汨��ͼ��UID
        If mobjStudyInfo.lngAdviceId = lngAdviceId Then
            If nvl(rsData!����ͼ) <> "" Then
                mstrReportImageUids = mstrReportImageUids & ";" & objImgInf.Key & ";"
                
                If Val(rsData!����ͼ) = 2 Then
                    objImgInf.DrawHint = "" '"��"
                Else
                    '1��0��״̬
                    objImgInf.DrawHint = "��"
                End If
            Else
                objImgInf.DrawHint = ""
            End If
        End If
        
        objImgInf.FtpFile = nvl(rsData!ͼ��UID) & IIf(nvl(rsData!ͼ������) = "REPIMG", ".jpg", "")
        objImgInf.Filename = nvl(rsData!ͼ��UID) & IIf(nvl(rsData!ͼ������) = "REPIMG", ".jpg", "")
        objImgInf.AdviceDes = IIf(nvl(rsData!ͼ������) = "REPIMG", "REPIMG", "")
        
        objImgInf.IsBackGround = mblnBgImgTrans
        objImgInf.JpgConvert = False
        
        If Val(nvl(rsData!��̬ͼ)) = ImgTag Then objImgInf.Format = ifDcm
        If Val(nvl(rsData!��̬ͼ)) = VIDEOTAG Then objImgInf.Format = ifAvi
        If Val(nvl(rsData!��̬ͼ)) = AUDIOTAG Then objImgInf.Format = ifWav
        If Val(nvl(rsData!��̬ͼ)) = BMPTAG Then objImgInf.Format = ifBmp
        
        objImgInf.SeriesNoTag = nvl(rsData!���к�, "*")
        objImgInf.ImageOrder = nvl(rsData!ͼ���, 0)
        
        Call ucImages.ConstructionImgData(objImgInf.CopyNew())
        
        i = i + 1
        
        Call rsData.MoveNext
    Wend
    
    Call ucImages.Refresh
End Sub

Public Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function


Private Function GetExamImgData(ByVal strStudyUID As String, ByVal blnMoved As Boolean) As ADODB.Recordset
'��ȡ���ͼ������
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim blnReadAllImg As Boolean
    
On Error GoTo errhandle
    Set GetExamImgData = Nothing
    
    blnReadAllImg = True
    If mlngModuleNo = G_LNG_PACSSTATION_MODULE Then
            'Ӱ��ҽ���ı���ͼֻ�ܸ���Ӱ�����¼�еı���ͼ�ֶν��м���
        strSQL = "  Select 1 as ���к�, Replace(Trim(B.Column_Value),'.jpg','') as ͼ��UID, rownum as ͼ���, 'REPIMG' as ͼ������, 2 as ����ͼ, 5 as ��̬ͼ, " & _
                        " null as ��������, null as �ɼ�ʱ��, null as ¼�Ƴ��� " & _
                        " From Ӱ�����¼ A, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) B " & _
                        " Where ���UID = [1]"
                            
        If blnMoved Then
            strSQL = Replace(strSQL, "Ӱ�����¼", "Ӱ�����¼")
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ͼ", strStudyUID)
        If rsData.RecordCount > 0 Then
            blnReadAllImg = False
        Else
            '���û�б���ͼ���򲻽���ͼ�����
            Exit Function
        End If
    End If
    
    If blnReadAllImg Then
        strSQL = "Select B.���к�,A.ͼ��UID, A.ͼ���,A.ͼ������, A.����ͼ, A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            " From Ӱ����ͼ�� A,Ӱ�������� B" & _
            " Where A.����UID=B.����UID And B.���UID=[1]"
     
        If mobjStudyInfo.blnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    
        If mblnImgAscOrder Then
            strSQL = strSQL & " order by B.���к�, A.�ɼ�ʱ��, ͼ���"
        Else
            strSQL = strSQL & " order by B.���к� Desc, A.�ɼ�ʱ�� Desc, ͼ��� Desc"
        End If
        
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ͼ������", strStudyUID)
    End If

    Set GetExamImgData = rsData
Exit Function
errhandle:
    If HintError(err, "GetExamImgData") Then Resume
End Function


Public Sub SyncCaptureImage(objImgInfo As clsBgImgInfo, Optional ByVal blnIsProxyTrans As Boolean = False)
'ͬ���ɼ�ͼ��
    Dim blnSyncStudyUID As Boolean
    mblnIsProcessing = True
On Error GoTo errhandle
    blnSyncStudyUID = IIf(ucImages.ImgCount <= 0 Or mobjStudyInfo.strStudyUID = "", True, False)
    
    If blnIsProxyTrans = False Then
        Call ucImages.AddImg(objImgInfo)
    Else
        Call ucImages.ProxyTransfer(objImgInfo)
    End If
    
    If blnSyncStudyUID Then
        mobjStudyInfo.strStudyUID = objImgInfo.StudyUID
    End If
    
    mblnIsProcessing = False
    
Exit Sub
errhandle:
    mblnIsProcessing = False
    err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
End Sub

Public Sub SyncAfterCapture(objImg As Object, strAfterTag As String)
'ͬ����̨�ɼ�
    Call ucCache.SyncAfterShow(objImg, strAfterTag)
End Sub


Public Sub SyncAfterTag(strAfterTag As String)
'ͬ����̨���
    Call ucCache.Refresh
End Sub


Public Sub SyncOutline(ByVal strOutlineName As String)
'ͬ�����
     Call ucWord.SyncOutline(strOutlineName)
End Sub


Private Sub InitTab()
On Error GoTo errH
    Dim i As Integer
    Dim iCount As Integer
    Dim strName() As String
     
    mblnIsTabIniting = True
    
    If tabSelect.ItemCount >= 1 Then
        mblnIsTabIniting = False
        Exit Sub
    End If
   
    With tabSelect
    
    
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ShowIcons = False ' True
        .PaintManager.Layout = xtpTabLayoutAutoSize ' xtpTabLayoutFixed ' ' xtpTabLayoutAutoSize
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionLeft
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.HeaderMargin.Left = 110 '135
        .PaintManager.HeaderMargin.Top = 0
        .PaintManager.HeaderMargin.Right = 2
        .PaintManager.HeaderMargin.Bottom = 0
        .PaintManager.ButtonMargin.Left = 0
        .PaintManager.ButtonMargin.Top = 0
        .PaintManager.ButtonMargin.Right = 2
        .PaintManager.ButtonMargin.Bottom = 3
 
        
        .RemoveAll
        
        
        If mlngModuleNo = G_LNG_PACSSTATION_MODULE Then
            .InsertItem 1, CON_TAB_TAG_����ͼ, picTemp.hwnd, 0
        Else
            .InsertItem 1, CON_TAB_TAG_ͼ��, picTemp.hwnd, 0
        End If
        
        .Item(0).tag = CON_TAB_TAG_ͼ��
         
        .InsertItem 2, CON_TAB_TAG_�ʾ�, picTemp.hwnd, 0
        .Item(1).tag = CON_TAB_TAG_�ʾ�

        .InsertItem 3, CON_TAB_TAG_��ʷ, picTemp.hwnd, 0
        .Item(2).tag = CON_TAB_TAG_��ʷ

        If mlngModuleNo = G_LNG_VIDEOSTATION_MODULE Then
            .InsertItem 4, CON_TAB_TAG_����, picTemp.hwnd, 0
            .Item(3).tag = CON_TAB_TAG_����
        End If

        .Item(0).Selected = True
        Set mobjSel = ucImages
        
        'Ĭ����ʾΪͼ��ɼ�����
        SetParent mobjSel.hwnd, picTemp.hwnd
        mobjSel.Visible = True
    End With
    
 
    
    mblnIsTabIniting = False
    Exit Sub
errH:
    mblnIsTabIniting = False
End Sub


Private Sub cmdAdd_Click()
On Error GoTo errhandle
    Call WriteData
Exit Sub
errhandle:
    HintError err, "cmdAdd_Click"
End Sub

Private Sub cmdAttach_Click()
On Error GoTo errhandle
    RaiseEvent OnDockAttachClick
Exit Sub
errhandle:
    HintError err, "cmdAttach_Click"
End Sub

Private Sub cmdDel_Click()
On Error GoTo errhandle
    Call DelData
Exit Sub
errhandle:
    HintError err, "cmdDel_Click"
End Sub
 

Private Sub cmdHide_Click()
On Error GoTo errhandle
    RaiseEvent OnDockHideClick
Exit Sub
errhandle:
    HintError err, "cmdHide_Click"
End Sub

Private Sub cmdMenu_Click()
On Error GoTo errhandle
    Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
End Sub


Private Sub cmdMenu_Resize()
On Error Resume Next
    imgMenu.Left = (cmdMenu.Width - imgMenu.Width) / 2
End Sub

Private Sub CmdRefresh_Click()
On Error GoTo errhandle
    Call RefreshData
Exit Sub
errhandle:
    HintError err, "cmdRefresh_Click"
End Sub

Private Sub imgMenu_Click()
    Call cmdMenu_Click
End Sub

Private Sub mobjImageProcessV2_OnSaveImage(ByVal emImageType As TImageType, dcmImage As DicomObjects.DicomImage)
'���洦����ͼ��
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    Dim objResult As clsBgImgInfo
    Dim strReportImgFile As String
    
On Error GoTo errhandle
    Select Case emImageType
        Case mtStudyImage   '���浽���ͼ
            strLineDeviceNo = GetDeptPara(mlngDeptID, "�洢�豸��")
            strBackDeviceNo = GetDeptPara(mlngDeptID, "�����豸��")
            
            Set objResult = SaveDicomImageToStudy(dcmImage, strLineDeviceNo, strBackDeviceNo)
            If Not objResult Is Nothing Then
                If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> dcmImage.StudyUID Then
                    mobjStudyInfo.strStudyUID = dcmImage.StudyUID
                End If
            End If
        Case mtReportImage  '���浽����ͼ
            strLineDeviceNo = GetDeptPara(mlngDeptID, "�洢�豸��")
            strBackDeviceNo = GetDeptPara(mlngDeptID, "�����豸��")
            
            Set objResult = SaveDicomImageToStudy(dcmImage, strLineDeviceNo, strBackDeviceNo)
            If Not objResult Is Nothing Then
                If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> dcmImage.StudyUID Then
                    mobjStudyInfo.strStudyUID = dcmImage.StudyUID
                End If
                
                strReportImgFile = objResult.FilePath & objResult.Filename & ".jpg"
                
                If FileExists(strReportImgFile) = False Then
                    Call dcmImage.FileExport(strReportImgFile, "BMP")
                End If
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_ADDIMG, , mobjStudyInfo.lngAdviceId, strReportImgFile)
            End If
    End Select
    
Exit Sub
errhandle:
    HintError err, "mobjImageProcessV2_OnSaveImage"
End Sub

Private Sub mobjImageProcessV2_OnUnload()
'    mlngBigImageIndex = 0
    Set mobjImageProcessV2 = Nothing
End Sub

Private Sub picBack_Resize()
On Error Resume Next
    If mblnIsEmbedVideoArea Then
        Call ucSplitter1.RePaint
    Else
        picHelperContainer.Move 0, 0, picBack.ScaleWidth, picBack.ScaleHeight
    End If
End Sub
 
Private Sub picHelperContainer_Resize()
On Error Resume Next
    tabSelect.Left = 0
    tabSelect.Top = 0
    tabSelect.Width = picHelperContainer.ScaleWidth
    tabSelect.Height = picHelperContainer.ScaleHeight
    
    cmdMenu.Left = 0
    cmdMenu.Top = picHelperContainer.Top
    cmdMenu.Width = ScaleWidth - ucImages.Width
    
    cmdRefresh.Left = 0
    cmdRefresh.Top = cmdMenu.Top + cmdMenu.Height
    cmdRefresh.Width = cmdMenu.Width
    
    cmdDel.Left = 0
    cmdDel.Top = cmdRefresh.Top + cmdRefresh.Height
    cmdDel.Width = cmdMenu.Width
    
    cmdAdd.Left = 0
    cmdAdd.Top = cmdDel.Top + cmdDel.Height
    cmdAdd.Width = cmdMenu.Width
    
    cmdHide.Left = 0
    cmdHide.Top = cmdAdd.Top + cmdAdd.Height
    cmdHide.Width = cmdMenu.Width
    
    cmdAttach.Left = 0
    cmdAttach.Top = cmdHide.Top + cmdHide.Height
    cmdAttach.Width = cmdMenu.Width
    
    
    imgMenu.Left = (cmdMenu.Width - imgMenu.Width) / 2
End Sub
 
Private Sub picTemp_Resize()
On Error Resume Next
    
    ucWord.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucImages.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucHistory.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
    ucCache.Move 0, 0, picTemp.ScaleWidth, picTemp.ScaleHeight
End Sub

 

 

Private Sub picVideoContainer_Resize()
On Error GoTo errhandle
    If mblnIsEmbedVideoArea = False Then Exit Sub
    If mobjEmbedVideo Is Nothing Then Exit Sub
    
    Call MoveWindow(mobjEmbedVideo.ContainerHwnd, 0, 0, _
            picVideoContainer.ScaleX(picVideoContainer.Width, vbTwips, vbPixels), _
            picVideoContainer.ScaleY(picVideoContainer.Height, vbTwips, vbPixels), 0)
    
    '��ʾ����
    ShowWindow mobjEmbedVideo.ContainerHwnd, 1
      
Exit Sub
errhandle:

End Sub

Private Sub tabSelect_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo errhandle

    If mblnIsTabIniting Then Exit Sub
    
    Select Case Item.tag
        Case CON_TAB_TAG_ͼ��
            Set mobjSel = ucImages
            
        Case CON_TAB_TAG_�ʾ�
            Set mobjSel = ucWord

        Case CON_TAB_TAG_��ʷ
            Set mobjSel = ucHistory
            
        Case CON_TAB_TAG_����
            Set mobjSel = ucCache
            
    End Select
    
    If ucWord.hwnd <> mobjSel.hwnd Then ucWord.Visible = False
    If ucImages.hwnd <> mobjSel.hwnd Then ucImages.Visible = False
    If ucHistory.hwnd <> mobjSel.hwnd Then ucHistory.Visible = False
    If ucCache.hwnd <> mobjSel.hwnd Then ucCache.Visible = False
    
    cmdDel.Enabled = IIf(Item.tag <> CON_TAB_TAG_��ʷ, True, False)
    cmdAdd.Enabled = mblnAllowWrite And IIf(Item.tag <> CON_TAB_TAG_����, True, False)
    
    SetParent mobjSel.hwnd, picTemp.hwnd
    mobjSel.Visible = True
    
    Call RefreshHelperComponent(mobjStudyInfo.lngAdviceId, mobjStudyInfo.strStudyUID, mlngFileID, mobjStudyInfo.blnMoved)
    
    RaiseEvent OnTabChanged(Item.tag)
Exit Sub
errhandle:
    HintError err, "tabSelect_SelectedChanged", False
End Sub

Public Sub FreeVideo()
On Error GoTo errhandle:
    Set mobjSel = Nothing
    
    If Not mobjEmbedVideo Is Nothing Then
        mobjEmbedVideo.zlNotifyQuit
        
        ShowWindow mobjEmbedVideo.ContainerHwnd, 0
        SetParent mobjEmbedVideo.ContainerHwnd, 0
    End If
    
    Set mobjEmbedVideo = Nothing
    
Exit Sub
errhandle:

End Sub
 
Private Sub ucCache_OnDblClick()
On Error GoTo errhandle
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    Call OpenImageProcess(True, True)
Exit Sub
errhandle:
    HintError err, "ucCache_OnDblClick", False
End Sub

Private Sub ucCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
'��ʾ�����Ҽ��˵�
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucCache_OnMouseUp", False
End Sub

Private Sub ucHistory_OnLinkView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)
On Error GoTo errhandle
    RaiseEvent OnLinkHistoryView(lngAdviceId, blnMoved, blnIsDBClick)
Exit Sub
errhandle:
    HintError err, "ucCache_OnMouseUp", False
End Sub

Private Sub ucHistory_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��ʾ��ʷ�Ҽ��˵�
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucHistory_OnMouseUp", False
End Sub

Private Sub ucHistory_OnSend()
    Dim i As Long
    Dim strtext As String
    Dim arySelIndex() As Long
    Dim objImg As DicomImage
    
On Error GoTo errhandle
    If mobjLinkEditor Is Nothing Then
        HintMsg "��δ����༭״̬������д�롣", "ucHistory_OnSend", vbOKOnly
        Exit Sub
    End If
    
    strtext = ucHistory.SelReportText
    If Len(strtext) > 0 Then
        Call mobjLinkEditor.InputWord(strtext, "", "", "")
        Exit Sub
    End If
    
    arySelIndex = ucHistory.GetSelects
    If UBound(arySelIndex) <= 0 Then Exit Sub
    
    For i = 1 To UBound(arySelIndex)
        Set objImg = ucHistory.GetImage(arySelIndex(i))
        If Not objImg Is Nothing Then
            Call mobjLinkEditor.AddRepImage(objImg, mlngReleationImgAdvice)
        End If
    Next
Exit Sub
errhandle:
    HintError err, "ucHistory_OnSend", False
End Sub

Private Sub ucImages_OnClick(ByVal lngImgIndex As Long)
    '��ͼƬ���͵���Ƶ�ɼ��н��д���
    Dim objImg As DicomImage
    Dim objImgTag As clsImageTagInf
    Dim objBgImgInfo As clsBgImgInfo
    
    Set objImg = ucImages.GetImage(lngImgIndex, objBgImgInfo)
    If objImg Is Nothing Then Exit Sub
        
    If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsRedo Or objBgImgInfo.LoadState = lsSent Then
        If Not mobjEmbedVideo Is Nothing Then
            If mobjEmbedVideo.VideoDockState Then Exit Sub
            
            If mblnIsEmbedVideoArea Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
            Else
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
            
            Exit Sub
        End If
        
        If Not mobjMainVideo Is Nothing Then
            If mobjMainVideo.VideoDockState Then Exit Sub
            Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
        End If
        
        Exit Sub
    End If
    
    '����tag
    Set objImgTag = New clsImageTagInf
    objImgTag.videoFile = objBgImgInfo.FilePath & objBgImgInfo.Filename
    objImgTag.EncoderName = ""
    objImgTag.CaptureTime = 0
    objImgTag.RecordTimeLen = 0
    objImgTag.FilePath = objBgImgInfo.FilePath
    
    Select Case objBgImgInfo.Format
        Case ifAvi
            objImgTag.tag = VIDEOTAG
        Case ifWav
            objImgTag.tag = AUDIOTAG
        Case ifDcm
            objImgTag.tag = ImgTag
    End Select
    
    Set objImg.tag = objImgTag
    
    
    If mblnIsEmbedVideoArea And Not mobjEmbedVideo Is Nothing Then
        If mobjEmbedVideo.VideoDockState Then Exit Sub
        '�����Ƶ���ں͵�ǰ�ؼ����ڲ���ͬһ�����ڣ�����ʾѡ��ͼ��
        If GetAncestor(mobjEmbedVideo.VideoHwnd, GA_ROOT) <> GetAncestor(hwnd, GA_ROOT) Then Exit Sub
        Call mobjEmbedVideo.zlPreviewThumbnail(objImg)
        Exit Sub
    End If
    
    
    If Not mobjMainVideo Is Nothing Then
        '��Ҫ�ж���Ƶ�ɼ���pacshelper�Ƿ���ͬһ��������
        If GetAncestor(mobjMainVideo.VideoHwnd, GA_ROOT) <> GetAncestor(hwnd, GA_ROOT) Then Exit Sub
        If mobjMainVideo.VideoDockState Then Exit Sub
        
        Call mobjMainVideo.zlPreviewThumbnail(objImg)
    End If
    

    If mblnClickBigImageShow Then
        If lngImgIndex <> mlngBigImageIndex Then
            If objBgImgInfo.Format <> ifAvi And objBgImgInfo.Format <> ifWav Then
                '����ͼ����ʾ
                Call ShowImageProcess(lngImgIndex, ptPreview)
            End If
        End If

        mlngBigImageIndex = lngImgIndex
    End If
End Sub

 

Private Sub ucImages_OnCmdEvent(ByVal strCmd As String)
    If strCmd = "REFRESH" Then Call LoadExamImages(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
End Sub

Private Sub ucImages_OnDBClick(ByVal lngImgIndex As Long)
    '�Ƿ�˫����ӱ���ͼ,����Ҫ�ж��Ƿ�Ϊ�ɼ�ͼ��
    
    '��ͼƬ���͵���Ƶ�ɼ��н��д���
    Dim objImg As DicomImage
    Dim objImgTag As clsImageTagInf
    Dim objBgImgInfo As clsBgImgInfo
    Dim strImgFailedFile As String
    
    Set objImg = ucImages.GetImage(lngImgIndex, objBgImgInfo)
    If objImg Is Nothing Then Exit Sub
        
    If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsRedo Or objBgImgInfo.LoadState = lsSent Then Exit Sub
    
    '����tag
    Set objImgTag = New clsImageTagInf
    objImgTag.videoFile = objBgImgInfo.FilePath & objBgImgInfo.Filename
    objImgTag.EncoderName = ""
    objImgTag.CaptureTime = 0
    objImgTag.RecordTimeLen = 0
    objImgTag.FilePath = objBgImgInfo.FilePath
    
    Select Case objBgImgInfo.Format
        Case ifAvi
            objImgTag.tag = VIDEOTAG
        Case ifWav
            objImgTag.tag = AUDIOTAG
        Case ifDcm
            objImgTag.tag = ImgTag
    End Select
    
    Set objImg.tag = objImgTag
    
    If objImgTag.tag <> ImgTag Then
        '����ý��
        If Not mobjMainVideo Is Nothing Then
            Call mobjMainVideo.playVideo(objImgTag.videoFile)
        End If
    Else
        strImgFailedFile = GetImgCmdFailed(objBgImgInfo)
        If objBgImgInfo.LoadState = lsError Or objBgImgInfo.LoadState = lsSent Or objBgImgInfo.LoadState = lsRedo _
            Or FileExists(strImgFailedFile) Then
            HintMsg "��ǰͼ��״̬��������", "WriteData", vbOKOnly
            Exit Sub
        End If
            
            
        '������Ҫ��ӱ���ͼ
        If mlngImageDBClickOper = 0 And mblnAllowWrite Then
            Call SendImageToReport
        Else
            Call OpenImageProcess
        End If
    End If
End Sub


Private Sub InitPar()
'��ȡ���ز���
    Dim strPrivatePath As String
    
    strPrivatePath = GetPrivateRegPath("ucPacsHelper")
    
    mblnAllowEmbedVideo = Val(GetDeptPara(mlngDeptID, "��ʾ��Ƶ�ɼ�", "0")) = 1
    
    mblnMoveBigImageShow = Val(GetSetting("ZLSOFT", strPrivatePath, "�ƶ���ʾ��ͼ", 0)) = 1  'Ӧ�õ���Ϊ�û�����
    mblnClickBigImageShow = Val(GetSetting("ZLSOFT", strPrivatePath, "������ʾ��ͼ", 0)) = 1
    mblnDelayCloseImage = Val(GetSetting("ZLSOFT", strPrivatePath, "�ӳٹرմ�ͼ", 0)) = 1
    
    ucHistory.IsOtherDept = Val(GetSetting("ZLSOFT", strPrivatePath, "���Ƽ��", "0")) = 1
    ucHistory.IsThisTime = Val(GetSetting("ZLSOFT", strPrivatePath, "�������", "0")) = 1
    ucHistory.IsAutoLine = Val(GetSetting("ZLSOFT", strPrivatePath, "�Զ�����", "0")) = 1
    
    ucWord.AutoHide = Val(GetSetting("ZLSOFT", strPrivatePath, "�Զ�����", "0")) = 1
    ucWord.ExpandLevel = Val(GetSetting("ZLSOFT", strPrivatePath, "չ���㼶", "1"))
    ucWord.DblWrite = Val(GetSetting("ZLSOFT", strPrivatePath, "����ʾ�˫������", "0")) = 0
    
    ucImages.PageRecordCount = Val(GetSetting("ZLSOFT", strPrivatePath, "����ͼ����(" & mstrParentName & ")", "8"))
    
    mblnIgnoreResult = GetDeptPara(mlngDeptID, "���Խ��������", 0) = "1" '        '���Խ��������
    mlngImageDBClickOper = GetDeptPara(mlngDeptID, "����ͼ˫������", 0)
    
    mlngReleationImgDays = Val(GetDeptPara(mlngDeptID, "�Զ�����ʷͼ������", 0))
End Sub


Private Sub SavePar()
'���汾�ز���
    Dim strPrivatePath As String
    
    strPrivatePath = GetPrivateRegPath("ucPacsHelper")
    
    Call SaveSetting("ZLSOFT", strPrivatePath, "�ƶ���ʾ��ͼ", IIf(mblnMoveBigImageShow, 1, 0))
    Call SaveSetting("ZLSOFT", strPrivatePath, "������ʾ��ͼ", IIf(mblnClickBigImageShow, 1, 0))
    Call SaveSetting("ZLSOFT", strPrivatePath, "�ӳٹرմ�ͼ", IIf(mblnDelayCloseImage, 1, 0))
    
    SaveSetting "ZLSOFT", strPrivatePath, "���Ƽ��", IIf(ucHistory.IsOtherDept, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "�������", IIf(ucHistory.IsThisTime, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "�Զ�����", IIf(ucHistory.IsAutoLine, 1, 0)
    
    
    SaveSetting "ZLSOFT", strPrivatePath, "�Զ�����", IIf(ucWord.AutoHide, 1, 0)
    SaveSetting "ZLSOFT", strPrivatePath, "չ���㼶", ucWord.ExpandLevel
    SaveSetting "ZLSOFT", strPrivatePath, "����ʾ�˫������", IIf(ucWord.DblWrite, 0, 1)
    
    SaveSetting "ZLSOFT", strPrivatePath, "����ͼ����(" & mstrParentName & ")", ucImages.PageRecordCount
End Sub


Public Sub RefreshData(Optional ByVal strHelperName As String = "")
    Dim strSelName As String
    
    strSelName = tabSelect.Selected.tag
    
    If Len(strHelperName) > 0 Then strSelName = strHelperName
    
    Select Case strSelName
        Case CON_TAB_TAG_ͼ��
            If Len(mobjStudyInfo.strStudyUID) > 0 Then Call LoadExamImages(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
            
            If ucImages.ImgCount <= 0 Then
                If mlngReleationImgAdvice <> 0 Then
                    Call LoadExamImages(mlngReleationImgAdvice, False)
                Else
                    mstrReportImageUids = ""
                    ucImages.ClearAll
                End If
            End If
    
        Case CON_TAB_TAG_����
            Call ucCache.Refresh
    
        Case CON_TAB_TAG_��ʷ
            'ˢ�¼����ʷ
            Call ucHistory.Refresh(mobjStudyInfo.lngAdviceId, True)
    
        Case CON_TAB_TAG_�ʾ�
            'ˢ�´ʾ�
            Call ucWord.Refresh(mobjStudyInfo.lngAdviceId, mlngFileID, , True)
    End Select
End Sub

Private Sub WriteData()
    Dim objImgInfo As clsBgImgInfo
    Dim strImgFailedFile As String
    Dim blnAllowAdditionInput As Boolean
    
    
    blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "��¼����") And mobjStudyInfo.intStep > 5)
    
    If Not (mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)) Then
        HintMsg "��ǰ���治��д�롣", "WriteData", vbOKOnly
        Exit Sub
    End If
    
    Select Case tabSelect.Selected.tag
        Case CON_TAB_TAG_ͼ��
            If ucImages.SelImgIndex <= 0 Then
                HintMsg "��ѡ����Ҫд���ͼ��", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            '�����ý��ͼ������ֱ�Ӽ��뱨��ͼ��ֻ�ܽ���ý�岥��
            Call ucImages.GetImage(ucImages.SelImgIndex, objImgInfo)
            
            If objImgInfo.Format = ifAvi And objImgInfo.Format = ifWav Then
                HintMsg "��ǰ��ʽ���ݲ�����д�롣", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            strImgFailedFile = GetImgCmdFailed(objImgInfo)
            If objImgInfo.LoadState = lsError Or objImgInfo.LoadState = lsSent Or objImgInfo.LoadState = lsRedo _
                Or FileExists(strImgFailedFile) Then
                HintMsg "��ǰͼ��״̬������д�롣", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call SendImageToReport
        
        Case CON_TAB_TAG_��ʷ
            If ucHistory.IsReportEnable(ucHistory.SelAdviceId) = False Then
                HintMsg "��ǰ״̬������д��,��ȷ����ʷ�����Ƿ�Ϊ�ա�", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucHistory.WriteReport
            
        Case CON_TAB_TAG_�ʾ�
            If ucWord.SelNodeType <> 2 Then
                HintMsg "��ѡ����Ҫд��Ĵʾ�ڵ㡣", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucWord.DirectWrite
            
'            Case CON_TAB_TAG_����
    End Select
End Sub

Private Sub DelData()
    Dim blnIsCancel As Boolean
    Dim aryIndex() As Long
    Dim objImgInfo As clsBgImgInfo
    
    Select Case tabSelect.Selected.tag
        Case CON_TAB_TAG_ͼ��
            If IsStudying = False Then
                HintMsg "��ǰ״̬������ɾ����", "DelData", vbOKOnly
                Exit Sub
            End If
            
            aryIndex = ucImages.GetSelects
            If UBound(aryIndex) <= 0 Then
                HintMsg "��ѡ����Ҫɾ�������ݡ�", "DelData", vbOKOnly
                Exit Sub
            End If
            
            Call ucImages.GetImage(aryIndex(1), objImgInfo)
            
            If objImgInfo.AdviceDes = "REPIMG" Then
                HintMsg "��ǰͼ���������ɾ����", "DelData", vbOKOnly
                Exit Sub
            End If
        
            'ɾ�����ͼ��
            Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel, hwnd)
            If blnIsCancel Then Exit Sub
            
            If DeleteStudyImage Then
                If ucImages.ImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1��ʾɾ�����һ��ͼ��
                Else
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
                End If
            End If
                
'        Case CON_TAB_TAG_��ʷ
            
        Case CON_TAB_TAG_�ʾ�
            If ucWord.SelNodeType <> 2 Then
                HintMsg "��ǰ�ڵ㲻����ɾ����", "WriteData", vbOKOnly
                Exit Sub
            End If
            
            Call ucWord.WordDelete
            
        Case CON_TAB_TAG_����
            Call DeleteCacheImage
            
    End Select
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnIsCancel As Boolean
    Dim lngImgCount As Long
    
On Error GoTo errhandle
    
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 0, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
    
    Select Case Control.ID
        Case conMenu_Helper_DelOper   'ɾ������
            If tabSelect.Selected.tag = CON_TAB_TAG_ͼ�� Then
                'ɾ�����ͼ��
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 0, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, blnIsCancel, hwnd)
                If blnIsCancel Then Exit Sub
                
                If DeleteStudyImage Then
                    If ucImages.ImgCount <= 0 Then
                        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1��ʾɾ�����һ��ͼ��
                    Else
                        Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
                    End If
                End If
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_���� Then
                'ɾ������ͼ��
                Call DeleteCacheImage
            End If
            
        Case conMenu_Helper_Refresh    'ˢ�²���
            Call RefreshData
            
        Case conMenu_Helper_SelAll 'ȫѡ
            If tabSelect.Selected.tag = CON_TAB_TAG_ͼ�� Then
                Call ucImages.SelectedAll
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_���� Then
                Call ucCache.SelectedAll
            End If
            
        Case conMenu_Helper_AddToReport '���뱨��ͼ
            If Len(Control.Parameter) <= 0 Then
                Call SendImageToReport
            Else
                '����ý��
                If Not mobjMainVideo Is Nothing Then
                    Call mobjMainVideo.playVideo(Control.Parameter)
                End If
            End If
            
        Case conMenu_Helper_ImageProcess 'ͼ��-ͼ����
            If tabSelect.Selected.tag = CON_TAB_TAG_ͼ�� Then
                If Len(Control.Parameter) <= 0 Then
                    Call OpenImageProcess
                Else
                    Call OpenImageProcess(True)
                End If
            ElseIf tabSelect.Selected.tag = CON_TAB_TAG_���� Then
                Call OpenImageProcess(True, True)
            End If
            
        Case conMenu_Helper_Log '������־�鿴
            If FileExists(Control.Parameter) Then
                Call OpenFilePos(Control.Parameter, True, "notepad")
            Else
                HintMsg "��־��δ���ɡ�", "cbrMain_Execute", infHint
            End If
            
        Case conMenu_Helper_Import 'ͼ��-����ͼ��
            lngImgCount = ucImages.ImgCount
            
            Call ImportImageFile
            
            If ucImages.ImgCount > 0 Then
                If lngImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, -1, , hwnd)    '-1��ʾ�״βɼ�
                ElseIf ucImages.ImgCount > lngImgCount Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, , , hwnd)
                End If
            End If
            
        Case conMenu_Helper_Export 'ͼ��-����ͼ��
            Call ExportImageFile
            
        Case conMenu_Helper_OpenImgPos '��ͼ��λ��
            If tabSelect.Selected.tag = CON_TAB_TAG_ͼ�� Then
                Call OpenImagePos
            End If
            
            If tabSelect.Selected.tag = CON_TAB_TAG_���� Then
                Call OpenCachePos
            End If
            
        Case conMenu_Helper_BigImageShow_Move 'ͼ��-�ƶ���ʾ��ͼ
            Control.Checked = Not Control.Checked
            mblnMoveBigImageShow = Control.Checked
            
            
        Case conMenu_Helper_BigImageShow_Click  '������ʾ��ͼ
            Control.Checked = Not Control.Checked
            mblnClickBigImageShow = Control.Checked
            
        Case conMenu_Helper_BigImageShow_Delay  '�ӳٹرմ�ͼ
            Control.Checked = Not Control.Checked
            mblnDelayCloseImage = Control.Checked
        
'        Case conMenu_Helper_ReDo    'ͼ��-���³���
'            Call ucImages.Redo
             
        Case conMenu_Helper_ReDown  'ͼ��-��������
            Call ucImages.ReDown
            
        Case conMenu_Helper_ReUp    'ͼ��-�����ϴ�
            If HintMsg("�����ϴ���ʹFTP�洢���ݱ��滻���Ƿ������", "cbrMain_Execute", vbYesNo) = vbNo Then Exit Sub
            Call ucImages.ReUp
            
        Case conMenu_Helper_SendCache 'ͼ��-���͵�����
            Call SendCache
            
            If ucImages.ImgCount <= 0 Then
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, -1, hwnd)    '-1��ʾɾ�����һ��ͼ��
            Else
                Call mObjNotify.Broadcast(BM_IMAGE_EVENT_DEL, 1, mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo, , hwnd)
            End If
        
        Case conMenu_Helper_SendStudy '����-���͵����
            lngImgCount = ucImages.ImgCount
            Call SendStudy
            
            If ucImages.ImgCount > 0 Then
                If lngImgCount <= 0 Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, -1, , hwnd)    '-1��ʾ�״βɼ�
                ElseIf ucImages.ImgCount > lngImgCount Then
                    Call mObjNotify.Broadcast(BM_IMAGE_EVENT_FIRST, , mobjStudyInfo.lngAdviceId, , , hwnd)
                End If
            End If
            
        Case conMenu_Helper_ImgViewer '��ʷ-��Ƭ
            Call mObjNotify.SendRequest(WM_IMG_OPENVIEW, , ucHistory.SelAdviceId)
            
        Case conMenu_Helper_ImgContrast '��ʷ-�Ա�
            Call mObjNotify.SendRequest(WM_IMG_CONTRASTVIEW, , ucHistory.SelAdviceId)
            
        Case conMenu_Helper_ReportOpen '��ʷ-����鿴
            Call mObjNotify.SendRequest(WM_REPORT_VIEW, , ucHistory.SelAdviceId, ucHistory.SelMoved)
        
        Case conMenu_Helper_ViewReportImage '�鿴����ͼ
            Call ucHistory.ViewReportImage
            
        Case conMenu_Helper_ViewReportContext '�鿴�����ı�
            Call ucHistory.ViewReportContext
            
        Case conMenu_Helper_WriteReport 'д�뱨��
            Call ucHistory.WriteReport
            
        Case conMenu_Helper_LinkViewer
            ucHistory.LinkViewed = Not ucHistory.LinkViewed
            
        Case conMenu_Helper_CloseViewer
            Call ucHistory.CloseLinkViewer
        
        Case conMenu_Helper_HalfYear, conMenu_Helper_TwoYear, conMenu_Helper_TwoMonth, _
            conMenu_Helper_ThreeYear, conMenu_Helper_ThreeMonth, conMenu_Helper_OneYear, _
            conMenu_Helper_OneMonth, conMenu_Helper_DateCus, conMenu_Helper_DateUn, _
            conMenu_Helper_DateCus '��ʷ-�Զ�����
             
            
            If Control.ID = conMenu_Helper_DateCus Then
                Call ucHistory.SetDateRange("�Զ���")
                Call ucHistory.ShowDateConfig
            Else
                Call ucHistory.SetDateRange(Control.Caption)
            End If
            
            Call ucHistory.Refresh(mobjStudyInfo.lngAdviceId, True)
            
        Case conMenu_Helper_ThisTime    '��ʷ-�Ƿ񱾴����
            Control.Checked = Not Control.Checked
            ucHistory.IsThisTime = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
        
        Case conMenu_Helper_OtherDept   '��ʷ-�Ƿ����Ƽ��
            Control.Checked = Not Control.Checked
            ucHistory.IsOtherDept = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
        
        Case conMenu_Helper_AutoLine  '��ʷ-�Ƿ��Զ�����
            Control.Checked = Not Control.Checked
            ucHistory.IsAutoLine = Control.Checked
            
            ucHistory.Refresh mobjStudyInfo.lngAdviceId, True
            
        Case conMenu_Helper_DirectWrite '�ʾ�-ֱ��д��
            Call ucWord.DirectWrite
            
        Case conMenu_Helper_EditWrite   '�ʾ�-�༭д��
            Call ucWord.EditWrite
            
        Case conMenu_Helper_FullSave    '�ʾ�-ȫ�״���
            Call ucWord.FullSave
            
        Case conMenu_Helper_NewWord     '�ʾ�-�����ʾ�
            Call ucWord.WordNew
            
        Case conMenu_Helper_ModWord     '�ʾ�-�޸Ĵʾ�
            Call ucWord.WordModify
            
        Case conMenu_Helper_DelWord     '�ʾ�-ɾ���ʾ�
            Call ucWord.WordDelete
            
        Case conMenu_Helper_AutoHide    '�ʾ�-�Զ�����
            Control.Checked = Not Control.Checked
            ucWord.AutoHide = Control.Checked
            
        Case conMenu_Helper_DblWrite    '˫��д��
            Control.Checked = Not Control.Checked
            ucWord.DblWrite = Control.Checked
            
        Case conMenu_Helper_AllLevel    '�ʾ�-չ������
            Control.Checked = True
            ucWord.ExpandLevel = 0
            
        Case conMenu_Helper_OneLevel    '�ʾ�-չ��һ��
            Control.Checked = True
            ucWord.ExpandLevel = 1
            
        Case conMenu_Helper_TwoLevel    '�ʾ�-չ������
            Control.Checked = True
            ucWord.ExpandLevel = 2
            
        Case conMenu_Helper_ThreeLevel  '�ʾ�-չ������
            Control.Checked = True
            ucWord.ExpandLevel = 3
    End Select
    
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 1, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
Exit Sub
errhandle:
    HintError err, "cbrMain_Execute", False
End Sub

Private Sub SendImageToReport()
'����ͼ�񵽱���
    Dim i As Long
    Dim arySelIndex() As Long
    Dim objImg As DicomImage
    Dim strSQL As String
    
    
     arySelIndex = ucImages.GetSelects()
     
     If UBound(arySelIndex) <= 0 Then
        HintMsg "��ѡ����Ҫ���͵������ͼ��", "SendImageToReport", vbOKOnly
        Exit Sub
     End If
     
     For i = 1 To UBound(arySelIndex)
        Set objImg = ucImages.GetImage(arySelIndex(i))
        If Not objImg Is Nothing Then
        
            '�������ݿ�
            strSQL = "Zl_Ӱ����_���ñ���ͼ('" & objImg.InstanceUID & "',1)"
            Call zlDatabase.ExecuteProcedure(strSQL, "Ԥ�ñ���ͼ")
                
            If Not mobjLinkEditor Is Nothing Then
                mobjLinkEditor.AddRepImage objImg, mlngReleationImgAdvice
            End If
            
            '���ñ���ͼ���
            Call ucImages.ImgDrawHint(objImg.InstanceUID, "��")
        End If
     Next
End Sub

Private Sub SendStudy()
'����ͼ�񵽼��
    Dim i As Long
    Dim arySelIndex() As Long
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    Dim objDcmImg As DicomImage
    
    arySelIndex = ucCache.GetSelects
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "��ѡ����Ҫ���͵ļ��ͼ��", "SendStudy", vbOKOnly
    End If
    
    strLineDeviceNo = GetDeptPara(mlngDeptID, "�洢�豸��")
    strBackDeviceNo = GetDeptPara(mlngDeptID, "�����豸��")
     
    For i = UBound(arySelIndex) To 1 Step -1
        Set objDcmImg = ucCache.GetImage(arySelIndex(i))
        If Not objDcmImg Is Nothing Then
            If Len(objDcmImg.InstanceUID) > 0 Then
                If Not SaveDicomImageToStudy(objDcmImg, strLineDeviceNo, strBackDeviceNo) Is Nothing Then
                    If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> objDcmImg.StudyUID Then mobjStudyInfo.strStudyUID = objDcmImg.StudyUID
                End If
            End If
        End If
    Next
    
    Call ucCache.DeleteCacheImg(-1)
End Sub

Private Sub SendCache()
'���͵�����
    Dim i As Long
    Dim arySelIndex() As Long
    Dim strCacheTag As String
    Dim strCachePath As String
    Dim objImg As DicomImage
    Dim objImgInfo As clsBgImgInfo
    
    
    arySelIndex = ucImages.GetSelects()
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "��ѡ����Ҫ���͵����ػ���ļ��ͼ��", "SendCache", vbOKOnly
        Exit Sub
    End If
    
    If HintMsg("ͼ���͵������ɾ������Ӧ�����ݣ��Ƿ������", "SendCache", vbYesNo) = vbNo Then Exit Sub
    
    strCacheTag = mobjStudyInfo.strPatientName & "(��ʱ)"  '  Format(Now, "hhmmss")
    strCachePath = GetCachePath(Format(Now, "YYYYMMDD"), strCacheTag)
    
    If DirExists(strCachePath) = False Then Call MkLocalDir(strCachePath)
    
    '�����ļ�������Ŀ¼
    For i = 1 To UBound(arySelIndex)
        Set objImg = ucImages.GetImage(arySelIndex(i), objImgInfo)
        If Not objImg Is Nothing Then
            FileCopy objImgInfo.FilePath & objImgInfo.Filename, strCachePath & objImgInfo.Filename
        End If
    Next
    
    'ɾ����Ӧ��ͼ������
    Call DeleteStudyImage(True)
    
    HintMsg "ͼ���ѷ��͵����Ϊ [" & strCacheTag & "] �Ļ���Ŀ¼�С�", "SendCache", vbOKOnly
End Sub


Private Sub OpenImageProcess(Optional ByVal blnForceRead As Boolean = False, Optional ByVal blnIsCache As Boolean = False)
    Dim arySelIndex() As Long
    
    If blnIsCache Then
        If ucCache.ImgCount <= 0 Then
            HintMsg "��ѡ����Ҫ�鿴��ͼ��", "ImageProcess", vbOKOnly
            Exit Sub
        End If
        
        arySelIndex = ucCache.GetSelects

    Else
        If ucImages.ImgCount <= 0 Then
            HintMsg "��ѡ����Ҫ�����ͼ��", "ImageProcess", vbOKOnly
            Exit Sub
        End If
        
        arySelIndex = ucImages.GetSelects
    End If
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "��δѡ�������ڵ�ǰ������ͼ��", "ImageProcess", vbOKOnly
        Exit Sub
    End If
     
    Call ShowImageProcess(arySelIndex(1), ptProcess, blnForceRead, blnIsCache)
End Sub

Private Function GetDlgSelectFileInfo(strFileName As String) As DlgFileInfo
'------------------------------------------------
'���ܣ����ļ���ת��Ϊȫ·������
'������strFileName--�ļ�����ͨ�����ļ��ؼ�����á�
'���أ�ȫ·������
'------------------------------------------------
    Dim sPath, tmpStr As String
    Dim sFIle() As String
    Dim iCount, i As Integer
    On Error GoTo errhandle
    sPath = CurDir()  '��õ�ǰ��·������Ϊ��CommonDialog�иı�·��ʱ��ı䵱ǰ��Path
    tmpStr = Right$(strFileName, Len(strFileName) - Len(sPath)) '���ļ����������
    
    If Left$(tmpStr, 1) = Chr$(0) Then
        'ѡ���˶���ļ�(����Ϊ��һ���ַ�Ϊ�ո�)
        For i = 1 To Len(tmpStr)
            If Mid$(tmpStr, i, 1) = Chr$(0) Then
                iCount = iCount + 1
                ReDim Preserve sFIle(iCount)
            Else
                sFIle(iCount) = sFIle(iCount) & Mid$(tmpStr, i, 1)
            End If
        Next i
    Else
        'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·����û��"\"��
        iCount = 1
        ReDim Preserve sFIle(iCount)
        If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
        sFIle(iCount) = tmpStr
    End If
    
    GetDlgSelectFileInfo.iCount = iCount
    
    ReDim GetDlgSelectFileInfo.sFIle(iCount)
    
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    GetDlgSelectFileInfo.sPath = sPath
    
    For i = 1 To iCount
        GetDlgSelectFileInfo.sFIle(i) = sFIle(i)
    Next i
    Exit Function
errhandle:
    HintError err, "GetDlgSelectFileInfo", False
End Function


Private Sub ImportImageFile()
'------------------------------------------------
'���ܣ����ⲿ�ļ�����������ͼ��
'��������
'���أ���
'------------------------------------------------
'TASK:����������������������������������������ʱ��֧��AVI���룬�����������ӡ�����������������������
On Error GoTo errH
    Dim DlgInfo As DlgFileInfo
    Dim i As Integer
    Dim objDcmImg As New DicomImage
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Dim dcmInfo As TDicomBaseInfo
    Dim strError As String
    
    Dim strLineDeviceNo As String
    Dim strBackDeviceNo As String
    
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    
    Dim objImgInfo As clsBgImgInfo
    
    'ѡ���ļ�
    With dlgOpen
        .CancelError = False
        .MaxFileSize = 32767 '���򿪵��ļ����ߴ�����Ϊ��󣬼�32K
        .flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
        .DialogTitle = "ѡ���ļ�"
        .Filter = "DICOM�ļ���*.dcm��(*.img)|*.dcm;*.img|ͼ���ļ� (*.BMP)(*.JPG)|*.BMP;*.JPG|�����ļ���*.*��|*.*"
        .ShowOpen
        If .Filename <> "" Then
            DlgInfo = GetDlgSelectFileInfo(.Filename)
        End If
        '�ڴ���*.pif�ļ����뽫Filename�����ÿգ�����ѡȡ���*.pif�ļ��󣬵�ǰ·����ı�
        .Filename = ""
    End With
    
    strLineDeviceNo = GetDeptPara(mlngDeptID, "�洢�豸��")
    strBackDeviceNo = GetDeptPara(mlngDeptID, "�����豸��")
    
    For i = 1 To DlgInfo.iCount
        
        Set objDcmImg = ReadDicomFile(DlgInfo.sPath & DlgInfo.sFIle(i), strError)
        
        If objDcmImg Is Nothing Then
            HintMsg "�ļ� " & DlgInfo.sPath & DlgInfo.sFIle(i) & " ��ȡ�쳣��" & IIf(strError = "", "", vbCrLf & strError), "ImportImageFile", infNormalErr
            Exit Sub
        End If
        
        If Len(objDcmImg.InstanceUID) <= 0 Then
            HintMsg "�ļ� " & DlgInfo.sPath & DlgInfo.sFIle(i) & " ��ȡ�쳣��δ����ʵ��UID��" & IIf(strError = "", "", vbCrLf & strError), "ImportImageFile", infNormalErr
            Exit Sub
        End If
        
        If Not SaveDicomImageToStudy(objDcmImg, strLineDeviceNo, strBackDeviceNo) Is Nothing Then
            If mobjStudyInfo.strStudyUID = "" Or mobjStudyInfo.strStudyUID <> objDcmImg.StudyUID Then mobjStudyInfo.strStudyUID = objDcmImg.StudyUID
        End If
    Next
    
    Exit Sub
errH:
    HintError err, "ImportImageFile", False
End Sub


Private Function SaveDicomImageToStudy(objDcmImg As DicomImage, _
    ByVal strLineDeviceNo As String, ByVal strBackDeviceNo As String) As clsBgImgInfo
 
    Dim dcmInfo As TDicomBaseInfo
    Dim strError As String
     
    Dim lineFtpInfo As TFtpDeviceInf
    Dim backFtpInfo As TFtpDeviceInf
    
    Dim objImgInfo As clsBgImgInfo
    Dim objResult As clsBgImgInfo
    
    Set SaveDicomImageToStudy = Nothing
    
    dcmInfo = GetDicomBaseInfo(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    Call WriteDicomPara(objDcmImg, dcmInfo)
    
    lineFtpInfo = GetLineFtpInfo(strLineDeviceNo, False, dcmInfo, strError)
    If Len(strError) > 0 Then
        err.Raise 0, "", strError
        Exit Function
    End If
    
    backFtpInfo = GetBackFtpInfo(strBackDeviceNo, dcmInfo, strError)
    If Len(strError) > 0 Then
        err.Raise 0, "", strError
        Exit Function
    End If
    
    If FileExists(GetStudyImgPath(dcmInfo) & dcmInfo.strInstanceUID) = False Then
        objDcmImg.WriteFile GetStudyImgPath(dcmInfo) & dcmInfo.strInstanceUID, True, "1.2.840.10008.1.2.1"
    End If
    
    Call SaveImageInfo(dcmInfo, lineFtpInfo)
    
    
    Set objImgInfo = GetBgImgInfo(dcmInfo, lineFtpInfo, backFtpInfo, True)
    
    objImgInfo.JpgConvert = True
    
    Set objResult = objImgInfo.CopyNew
    
    Call ucImages.AddImg(objImgInfo)
    
    Set SaveDicomImageToStudy = objResult
End Function

Private Sub ExportSingleImg(objImg As DicomImage, objImgInfo As clsBgImgInfo, _
    Optional ByVal blnUseFix As Boolean = False, Optional ByVal strExportFile As String = "", Optional ByVal strSuffix As String = "")
'��������ͼ��
    Dim strFileName As String
    Dim blnIsCopy As Boolean
    Dim strFileType As String
     
    
    blnIsCopy = False
    
    If blnUseFix = False Then
        '��ʹ��ǰ׺�ַ�
        strFileName = Replace(UCase(strExportFile), "." & strSuffix, "")
    Else
        strFileName = Replace(UCase(strExportFile), "." & strSuffix, "") & "_" & objImg.InstanceUID
    End If
        
    If Len(strFileName) <= 0 Then Exit Sub
    
    If Trim(strSuffix) = "" Then
        blnIsCopy = True
    Else
        If objImgInfo.Format = ifAvi Or objImgInfo.Format = ifWav Then
            blnIsCopy = True
        Else
            strFileName = strFileName & "." & strSuffix
        End If
    End If
    
    If blnIsCopy Then
        Call FileCopy(objImgInfo.FilePath & objImgInfo.Filename, strFileName)
        Exit Sub
    End If
    
    strFileType = UCase(Right(Trim(strFileName), 3))
    
    Select Case strFileType
        Case "AVI"
            objImg.WriteAVI strFileName, 1, objImg.FrameCount, 1, "", 100, False
        Case "DCM"
            objImg.WriteFile strFileName, True
        Case "BMP"
            objImg.FileExport strFileName, "BMP"
        Case "JPG"
            objImg.FileExport strFileName, "JPG"
    End Select
End Sub

Private Sub ExportImageFile()
'------------------------------------------------
'���ܣ����dcmView�е�ͼ��,֧�ֵĸ�ʽΪAVI,DCM,BMP,JPE
'��������
'���أ���
'------------------------------------------------
    Dim i As Long
    
    Dim arySelIndex() As Long
    Dim objDcmImg As DicomImage
    Dim objImgInfo As clsBgImgInfo
    
    Dim strExt As String
    
    arySelIndex = ucImages.GetSelects
    
    If UBound(arySelIndex) <= 0 Then
        HintMsg "��ѡ����Ҫ������ͼ��", "ExportImageFile", vbOKOnly
        Exit Sub
    End If
    
    dlgOpen.Filter = "ԭʼ��ʽ| |(*.dcm)|*.dcm|(*.bmp)|*.bmp|(*.jpg)|*.jpg|(*.avi)|*.avi|(*.mpeg)|*.mpeg"
    dlgOpen.FilterIndex = 1
    dlgOpen.Filename = mobjStudyInfo.strPatientName
    
    dlgOpen.ShowSave
    
    If Len(dlgOpen.Filename) <= 0 Then Exit Sub
    
    strExt = ""
    
    Select Case dlgOpen.FilterIndex
        Case 1
            strExt = ""
        Case 2
            strExt = "DCM"
        Case 3
            strExt = "BMP"
        Case 4
            strExt = "JPG"
        Case 5
            strExt = "AVI"
        Case 6
            strExt = "MPEG"
    End Select
        
    If UBound(arySelIndex) = 1 Then
        '��������ͼ��
        Set objDcmImg = ucImages.GetImage(arySelIndex(1), objImgInfo)

        Call ExportSingleImg(objDcmImg, objImgInfo, False, dlgOpen.Filename, strExt)
    Else
        '�������ͼ��
        For i = 1 To UBound(arySelIndex)
            Set objDcmImg = ucImages.GetImage(arySelIndex(i), objImgInfo)
            
            Call ExportSingleImg(objDcmImg, objImgInfo, True, dlgOpen.Filename, strExt)
        Next
    End If
    
    If HintMsg("�ļ�������ɣ��Ƿ�򿪵���Ŀ¼?", "ExportImageFile", vbYesNo) = vbYes Then
        Call OpenFilePos(dlgOpen.Filename)
    End If
End Sub

Private Sub OpenFilePos(ByVal strFile As String, Optional ByVal blnIsOpenFile As Boolean = False, Optional ByVal strOpenWay As String = "")
'���ļ�λ��
    If blnIsOpenFile = False Then
        ShellExecute 0, "open", Mid(strFile, 1, InStrRev(strFile, "\")), "", "", 1
    Else
        If Len(strOpenWay) <= 0 Then
            ShellExecute 0, "open", strFile, "", "", 1
        Else
            ShellExecute 0, "open", strOpenWay, strFile, "", 1
        End If
    End If
    
End Sub

Private Sub OpenImagePos()
'��ͼ��λ��
    Dim objImgInfo As clsBgImgInfo
    
    If ucImages.ImgCount <= 0 Then Exit Sub
    
    Call ucImages.GetImage(1, objImgInfo)
    
    If objImgInfo Is Nothing Then Exit Sub
    
    Call OpenFilePos(objImgInfo.FilePath & objImgInfo.Filename)
End Sub


Private Sub OpenCachePos()
'�򿪻���λ��
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    ucCache.OpenCachePath
End Sub

Private Sub DeleteCacheImage()
'ɾ�������ͼ��
    If ucCache.ImgCount <= 0 Then Exit Sub
    
    '�ж��Ƿ���ͼ��ѡ��
    If ucCache.IsSelected = False Then
        HintMsg "��ѡ����Ҫɾ���Ļ���ͼ��", "DeleteCacheImage", vbOKOnly
        Exit Sub
    End If
    
    If HintMsg("����ͼ��ɾ���󽫲��ָܻ����Ƿ������", "DeleteCacheImage", vbYesNo) = vbNo Then Exit Sub
    
    Call ucCache.DeleteCacheImg(-1)
End Sub

Private Sub SetRemoveTag(ByVal strFile As String)
On Error Resume Next
    Call SetFileHide(strFile)
    
    Name strFile As strFile & ".DEL"
End Sub

Private Function DeleteStudyImage(Optional ByVal blnIsSendCache As Boolean = False) As Boolean
'------------------------------------------------
'���ܣ�ɾ������ͼ�б�ѡ�е�ͼ���ȴ����ݿ���ɾ����Ȼ���FTP��ɾ����
'��������
'���أ��ޣ�ֱ��ɾ������ͼ�����һ��ͼ��
'------------------------------------------------
On Error GoTo errH
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    Dim i As Long
    Dim lngIndex  As Long
    Dim arySelIndex() As Long
    
    Dim objImgInfo As clsBgImgInfo
    Dim blnIsContainReport As Boolean
    
    Dim strucFtpTag As TFtpConTag

    DeleteStudyImage = False
    If ucImages.ImgCount <= 0 Then Exit Function
     
    blnIsContainReport = False
    arySelIndex = ucImages.GetSelects()
    
    If UBound(arySelIndex) <= 0 Then Exit Function
    
    If blnIsSendCache = False Then
        If HintMsg("ͼ��ɾ���󽫲��ָܻ����Ƿ������", "DeleteStudyImage", vbYesNo) = vbNo Then Exit Function
    End If
    
    
    '�õ���Ҫɾ����ͼ��uid�м���';'����
    For i = UBound(arySelIndex) To 1 Step -1
        lngIndex = arySelIndex(i)
        Call ucImages.GetImage(lngIndex, objImgInfo)

        If InStr(mstrReportImageUids, ";" & objImgInfo.Key & ";") <= 0 And InStr(objImgInfo.DrawHint, "��") <= 0 Then '����Ǳ���ͼ�������б���ͼ��ǣ������������ɾ��
            If strucFtpTag.Ip <> objImgInfo.FtpIp Then
                strucFtpTag = FtpTagInstance(objImgInfo.FtpIp, objImgInfo.FtpUser, objImgInfo.FtpPwd, objImgInfo.FtpVirtualPath)
            End If
            
            strSQL = "ZL_Ӱ��ͼ��_DELETE(" & objImgInfo.AdviceId & ",0,'" & objImgInfo.Key & "',Null)"
            zlDatabase.ExecuteProcedure strSQL, "ɾ�����ͼ��"
            
            'ɾ��ͼ���ļ�
            If FtpDelete(strucFtpTag, objImgInfo.Key, False, False) = frAbort Then Exit Function
            '�����ļ�����Ϊ����
            Call SetRemoveTag(objImgInfo.FilePath & objImgInfo.Filename)
            
            '�����dcm�ļ����ſ��ܴ��ڶ�Ӧ��jpg�ļ�
            If objImgInfo.Format = ifDcm Then
                If FtpDelete(strucFtpTag, objImgInfo.Key & ".jpg", True, False) = frAbort Then Exit Function
                Call SetRemoveTag(objImgInfo.FilePath & objImgInfo.Key & ".jpg")
            End If
            
            Call ucImages.DelImgView(lngIndex)
        Else
            blnIsContainReport = True
        End If
    Next
    
    Call ucImages.Selected(lngIndex)
     
    If blnIsContainReport Then
        HintMsg IIf(blnIsSendCache, "���͵������", "ɾ��") & "��ͼ���а�������ͼ������ͼ������ɾ����", "DeleteStudyImage", vbOKOnly
        Exit Function
    End If
    
    DeleteStudyImage = True
    
    '�ж��Ƿ񻹴��ڼ��ͼ��
    If ucImages.ImgCount <= 0 Then
 
'        strSQL = "Select * from ����ҽ������ where ҽ��ID=[1] and ���ͺ�=[2]"
'        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ҽ��״̬", mlngAdviceID, mlngSendNo)
'
'        '������״̬Ϊ�Ѽ�飬����ɾ������ͼ�����Ҫ��ͼ����л���
'        If Val(nvl(rsData!ִ�й���)) = 3 Then
'            '����Ӱ����״̬�����ɾ�����һ��ͼ����ԭ������Ϊ3�����޸�Ϊ2
'            strSQL = "Zl_Ӱ����_State(" & mlngAdviceID & "," & mlngSendNo & ",2,NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngDeptID & ")"
'            zlDatabase.ExecuteProcedure strSQL, "ɾ�����һ��ͼ��"
'
'            '����״̬�ı���Ϣ
''            Call mObjNotify.SendRequest(WM_LIST_SYNCROW, , mlngAdviceID, mlngSendNo)
'        End If
        
        '�ָ�ʵʱ��Ƶ��ʾ
        If Not mobjEmbedVideo Is Nothing Then
            If mobjEmbedVideo.VideoDockState Then Exit Function
            
            If mblnIsEmbedVideoArea Then
                Call mobjEmbedVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), False)
            Else
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
        Else
            If Not mobjMainVideo Is Nothing Then
                If mobjMainVideo.VideoDockState Then Exit Function
                Call mobjMainVideo.zlRestoreWindow(IIf(mobjStudyInfo.intStep > 1 And mobjStudyInfo.intStep < 5, False, True), True)
            End If
        End If
    End If
        
    Exit Function
errH:
    If HintError(err, "DeleteStudyImage", False) = 1 Then Resume
End Function


Private Sub ShowPopupMenu(ByVal strTabTag As String)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'intType: 1--����ͼ��2--����ͼ
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    Dim blnVisible As Boolean
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim strCaption As String
    Dim blnAllowAdditionInput As Boolean
    Dim objImgInfo As clsBgImgInfo
    Dim lngImgFmt As TImageFmt
    Dim blnResultOk As Boolean
    Dim strImgFailedFile As String
    Dim blnIsReportImg As Boolean
    
    '����Ҽ������˵�
    cbrMain.DeleteAll
    
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    
    blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "��¼����") And mobjStudyInfo.intStep > 5)
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Refresh, "ˢ��(&R)")
                
 
        Select Case strTabTag
            Case CON_TAB_TAG_ͼ�� '**********************************************************************************
                blnIsReportImg = False
                If ucImages.ImgCount <= 0 Then
                    blnVisible = False
                    blnResultOk = False
                Else
                    If ucImages.SelImgIndex > 0 Then
                        blnVisible = True
                        '�����ý��ͼ������ֱ�Ӽ��뱨��ͼ��ֻ�ܽ���ý�岥��
                        Call ucImages.GetImage(ucImages.SelImgIndex, objImgInfo)
                        
                        lngImgFmt = objImgInfo.Format
                        blnIsReportImg = IIf(objImgInfo.AdviceDes = "REPIMG", True, False)
                        blnResultOk = IIf(objImgInfo.LoadState <> lsError And objImgInfo.LoadState <> lsSent And objImgInfo.LoadState <> lsRedo, True, False)
                        strImgFailedFile = GetImgCmdFailed(objImgInfo)
                    Else
                        blnVisible = False
                    End If
                End If
            
                If lngImgFmt <> ifAvi And lngImgFmt <> ifWav Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AddToReport, "���뱨��ͼ(&i)")
                    cbrControl.Parameter = ""
                    cbrControl.Visible = blnVisible And blnResultOk
                    cbrControl.Enabled = (IsStudying Or blnAllowAdditionInput) 'mblnAllowWrite And
                    
'                    If Not mobjLinkEditor Is Nothing Then
'                        cbrControl.Enabled = cbrControl.Enabled And mobjLinkEditor.IsEditable
'                    End If
                Else
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AddToReport, "����(&P)")
                    cbrControl.Parameter = objImgInfo.FilePath & objImgInfo.Filename
                    cbrControl.Visible = blnVisible And blnResultOk
                End If
                
                cbrControl.BeginGroup = True
                
                    
'                If blnResultOk Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImageProcess, "ͼ����(&C)")
                        cbrControl.Visible = blnVisible
                        cbrControl.Parameter = strImgFailedFile
'                Else
                    If FileExists(strImgFailedFile) Then
                        Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Log, "������־(&L)")
                            cbrControl.Parameter = strImgFailedFile
                            cbrControl.Visible = blnVisible
                    End If
'                End If
                
                
                    
    '            Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SplitPage, "��ҳ����")
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SelAll, "ȫѡͼ��(&F)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = IIf(ucImages.ImgCount > 0, True, False)
                    
                If mlngModuleNo = G_LNG_VIDEOSTATION_MODULE Then
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SendCache, "���͵�����(&E)")
                        cbrControl.Visible = blnVisible And blnResultOk
                        
                    Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Import, "����ͼ��(&M)")
                        cbrControl.Visible = mlngModuleNo <> G_LNG_PACSSTATION_MODULE
                        cbrControl.Enabled = IsStudying
                End If
                 
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Export, "����ͼ��(&T)")
                    cbrControl.Visible = blnVisible
                      
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelOper, "ɾ��ͼ��(&D)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible And blnIsReportImg = False
                    cbrControl.Enabled = IsStudying
                    

                
'                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReDo, "����...")
'                    cbrControl.BeginGroup = True
'                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReDown, "��������(&N)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReUp, "�����ϴ�(&U)")
                    cbrControl.Visible = blnVisible
                     
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_OpenImgPos, "��ͼ��λ��(&O)...")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Move, "�ƶ���ʾ��ͼ(&B)")
                    cbrControl.BeginGroup = True
                    cbrControl.Checked = mblnMoveBigImageShow
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Click, "������ʾ��ͼ(&K)")
                    cbrControl.Checked = mblnClickBigImageShow
                    
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_BigImageShow_Delay, "�ӳٹرմ�ͼ(&A)")
                    cbrControl.Checked = mblnDelayCloseImage
                    

            Case CON_TAB_TAG_���� '**********************************************************************************
                blnVisible = IIf(ucCache.ImgCount > 0, True, False)
    '            Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SplitPage, "��ҳ����")
    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImageProcess, "ͼ��Ԥ��(&V)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SendStudy, "���͵����(&S)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = IsStudying
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelOper, "ɾ��(&D)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_SelAll, "ȫѡͼ��(&F)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = IIf(ucCache.ImgCount > 0, True, False)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_OpenImgPos, "��ͼ��λ��(&O)...")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
            
            Case CON_TAB_TAG_��ʷ '**********************************************************************************
                blnVisible = ucHistory.SelAdviceId > 0
                
                strCaption = IIf(mlngModuleNo <> G_LNG_PACSSTATION_MODULE, "Ӱ�����(&F)", "Ӱ���Ƭ(&F)")
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImgViewer, strCaption)
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsImageEnable(ucHistory.SelAdviceId) Or mlngReleationImgDays > 0
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ImgContrast, "Ӱ��Ա�(&C)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsImageEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ReportOpen, "����Ԥ��(&V)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
'                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_Analysis, "�ۺϷ���")
'                    cbrControl.Visible = blnVisible
'                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ViewReportImage, "�鿴����ͼ(&i)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ViewReportContext, "�鿴�����ı�(&T)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_WriteReport, "д�뱨��(&W)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucHistory.IsReportEnable(ucHistory.SelAdviceId) And mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_LinkViewer, "�����鿴(&K)")
                    cbrControl.Visible = blnVisible And ucHistory.AllowLinkViewer
                    cbrControl.Checked = ucHistory.LinkViewed
                    cbrControl.BeginGroup = True
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_CloseViewer, "�˳��鿴(&Q)")
                    cbrControl.Visible = blnVisible And ucHistory.AllowLinkViewer
                    
            
                'ʱ��.........................................................
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_DateRange, "����(&D)")
                    cbrControl.ToolTipText = "����"
                    cbrControl.BeginGroup = True
        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneMonth, "һ����(&1)")
                        objControl.Checked = IIf(ucHistory.DataRange = "һ����", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoMonth, "������(&2)")
                        objControl.Checked = IIf(ucHistory.DataRange = "������", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeMonth, "������(&3)")
                        objControl.Checked = IIf(ucHistory.DataRange = "������", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_HalfYear, "����(&4)")
                        objControl.Checked = IIf(ucHistory.DataRange = "����", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneYear, "һ��(&5)")
                        objControl.Checked = IIf(ucHistory.DataRange = "һ��", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoYear, "����(&6)")
                        objControl.Checked = IIf(ucHistory.DataRange = "����", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeYear, "����(&7)")
                        objControl.Checked = IIf(ucHistory.DataRange = "����", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_DateUn, "����(&F)")
                        objControl.Checked = IIf(ucHistory.DataRange = "����", True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_DateCus, "�Զ���(&C)")
                        objControl.Checked = IIf(ucHistory.DataRange = "�Զ���", True, False)
 
                
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_RelateCfg, "ѡ��(&O)")
                    cbrControl.ToolTipText = "ѡ��"
                
                    If mobjStudyInfo.lngPatientFrom = 2 Then 'ֻ��סԺ���ߣ�����Ҫ��ʾ������ز˵�
                        Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThisTime, "�������(&1)")
                            cbrPopControl.Checked = ucHistory.IsThisTime
                    End If
                    
                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OtherDept, "���Ƽ��(&2)")
                        cbrPopControl.Checked = ucHistory.IsOtherDept
                        
                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_AutoLine, "�Զ�����(&3)")
                        cbrPopControl.Checked = ucHistory.IsAutoLine
                        
            Case CON_TAB_TAG_�ʾ� '**********************************************************************************
                blnVisible = ucWord.NodeCount > 0
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DirectWrite, "ֱ��д��(&E)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_EditWrite, "�༭д��(&W)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_FullSave, "ȫ�״���(&S)")
                    cbrControl.BeginGroup = True
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_NewWord, "�����ʾ�(&N)")
                    cbrControl.Visible = blnVisible
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_ModWord, "�޸Ĵʾ�(&M)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucWord.SelNodeType = 2
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DelWord, "ɾ���ʾ�(&D)")
                    cbrControl.Visible = blnVisible
                    cbrControl.Enabled = ucWord.SelNodeType = 2
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_AutoHide, "�Զ�����(&H)")
                    cbrControl.BeginGroup = True
                    cbrControl.Checked = ucWord.AutoHide
                    
                Set cbrControl = .Add(xtpControlButton, conMenu_Helper_DblWrite, "˫��д��(&O)")
                    cbrControl.Checked = ucWord.DblWrite
                    
                Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Helper_ExpandLevel, "չ���㼶(&V)")
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_OneLevel, "һ��(&1)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 1, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_TwoLevel, "����(&2)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 2, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_ThreeLevel, "����(&3)")
                        objControl.Checked = IIf(ucWord.ExpandLevel = 3, True, False)
                        
                    Set objControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Helper_AllLevel, "����(&F)")
                        objControl.BeginGroup = True
                        objControl.Checked = IIf(ucWord.ExpandLevel = 0, True, False)
                
        End Select
    End With
    
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub ucImages_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim blnShowImg As Boolean
    Dim intCurImg As Integer
    Dim objDcmViewer As DicomViewer
    Dim objImgInfo As clsBgImgInfo
    
    If mblnMoveBigImageShow = False Then Exit Sub
 

    Set objDcmViewer = ucImages.Viewer

    '�ж��Ƿ���Ҫ��ʾͼ��
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= objDcmViewer.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= objDcmViewer.Height) Then
        blnShowImg = True
    End If

    If blnShowImg Then        '��ʾͼ��
        SetCapture objDcmViewer.hwnd    '�������
        
        If mlngStartMoveTime = 0 Then mlngStartMoveTime = GetTickCount
        
        '����ƶ���ͼ���ϵ�ʱ���ӡ500����󣬲ſ�ʼ��ʾ��ͼ
        If GetTickCount - mlngStartMoveTime < 500 Then
            mlngBigImageIndex = 0
            Exit Sub
        End If
        
        
        intCurImg = objDcmViewer.ImageIndex(X, Y)
        
        Call ucImages.GetImage(intCurImg, objImgInfo)

        If objImgInfo Is Nothing Then Exit Sub


        If intCurImg <> mlngBigImageIndex Then
            If objImgInfo.Format <> ifAvi And objImgInfo.Format <> ifWav Then
            '����ͼ����ʾ
                If objImgInfo.LoadState = lsError Or objImgInfo.LoadState = lsRedo Or objImgInfo.LoadState = lsSent Then
                    Call ShowImageProcess(intCurImg, ptPreview, True)
                Else
                    Call ShowImageProcess(intCurImg, ptPreview)
                End If
            End If
        End If

        mlngBigImageIndex = intCurImg
    Else
        ReleaseCapture
        
        If mblnDelayCloseImage = False Then
            CloseImageProcess
        End If
        
        mlngStartMoveTime = 0
    End If
     
Exit Sub
errhandle:
    Call HintError(err, "ucImages_OnMouseMove", False)
End Sub


Public Sub CloseImageProcess()
    If mobjImageProcessV2 Is Nothing Then Exit Sub
    
    If Not mobjImageProcessV2 Is Nothing Then
        If mobjImageProcessV2.WinType = 0 Then
            Unload mobjImageProcessV2
             
            Set mobjImageProcessV2 = Nothing
        End If
    End If
End Sub


Public Sub ShowImageProcess(ByVal lngImgIndex As Long, ByVal lngType As TImgProcessType, _
    Optional ByVal blnForaceRead As Boolean = False, Optional ByVal blnIsCache As Boolean = False)
    Dim i As Long
    Dim objDcmImg As DicomImage
    Dim objImgInfos() As Object
    Dim blnReportImgState As Boolean
    Dim blnAllowAdditionInput As Boolean
    
    
    If lngImgIndex <= 0 Then Exit Sub
    
    If mobjImageProcessV2 Is Nothing Then
        Set mobjImageProcessV2 = New frmImageProcessV2
    Else
        '����Ԥ��ͼ��ʱ���ƶ�����л�ͼ��ˢ��
        If mobjImageProcessV2.WinType <> 0 And lngType = 0 Then Exit Sub
    End If
           
    If blnIsCache Then
        Set objDcmImg = ucCache.GetImage(lngImgIndex)
        
        Call mobjImageProcessV2.SetButtonState(False, False)
        
        ReDim objImgInfos(0)
        
        Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, 0, objDcmImg, objImgInfos, ptPreview, 30, mblnAllowWrite)
    Else
        Set objDcmImg = ucImages.GetImage(lngImgIndex)
        
        ReDim objImgInfos(ucImages.ImgCount - 1)
        
        For i = 0 To ucImages.ImgCount - 1
            Set objImgInfos(i) = ucImages.ImageInfo(i)
        Next
         
        blnAllowAdditionInput = (CheckPopedom(mstrPrivs, "��¼����") And mobjStudyInfo.intStep > 5)
        blnReportImgState = mblnAllowWrite And (IsStudying Or blnAllowAdditionInput)
        
        If Not mobjLinkEditor Is Nothing Then
            blnReportImgState = blnReportImgState And mobjLinkEditor.IsEditable
        End If
        
        If blnForaceRead Then
            Call mobjImageProcessV2.SetButtonState(False, False)
        Else
            Call mobjImageProcessV2.SetButtonState(IsStudying, blnReportImgState)
        End If
        
        '��û���κδ����£�2����Զ��رմ�ͼԤ��
        If mblnDelayCloseImage Then
            Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, mobjStudyInfo.lngAdviceId, objDcmImg, objImgInfos, lngType, 2, mblnAllowWrite)
        Else
            Call mobjImageProcessV2.ZlShowMe(mObjNotify.Owner, mobjStudyInfo.lngAdviceId, objDcmImg, objImgInfos, lngType, 30, mblnAllowWrite)
        End If
    End If
     
    

End Sub


Private Sub ucImages_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
'��ʾͼ���Ҽ��˵�
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucImages_OnMouseUp", False
End Sub

Private Sub ucWord_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'��ʾ�ʾ��Ҽ��˵�
On Error GoTo errhandle
    If Button = 2 Then Call ShowPopupMenu(tabSelect.Selected.tag)
Exit Sub
errhandle:
    HintError err, "ucWord_OnMouseUp", False
End Sub
 

Private Sub ucWord_OnRequestState(lngOutlineType As TOutlineType, str�������� As String, str������� As String, str�������� As String)
'��ȡ����������Ϣ
    If mobjLinkEditor Is Nothing Then Exit Sub
    
    mobjLinkEditor.GetReportContext str��������, str�������, str��������
    
    lngOutlineType = mobjLinkEditor.CurOutlineType
End Sub

Private Sub ucWord_OnSendContext(ByVal strFreeText As String, ByVal str�������� As String, ByVal str������� As String, ByVal str�������� As String)
'���ʹʾ����ݵ�����
    If mobjLinkEditor Is Nothing Then Exit Sub
    
    mobjLinkEditor.InputWord strFreeText, str��������, str�������, str��������
    
End Sub

Private Sub UserControl_Initialize()
    mblnIsValid = False
    mblnIsEmbedVideoArea = True
    mblnAllowEmbedVideo = True
    mblnBgImgTrans = True
    mblnMoveBigImageShow = True
    mblnClickBigImageShow = False
    mblnDelayCloseImage = True
    mblnAllowWrite = True
    
    mlngReleationImgAdvice = 0
    
    Set mobjSel = Nothing
    
End Sub
 


Private Sub UserControl_Resize()
On Error Resume Next
    picBack.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub UserControl_Show()
'On Error Resume Next
'    Call InitTab
End Sub

Public Sub Destory()
    mlngReleationImgAdvice = 0
    
    Call FreeVideo
     
    
    ucSplitter1.Destory
    
    If Not mobjImageProcessV2 Is Nothing Then
        Unload mobjImageProcessV2
    End If
    
    Set mobjImageProcessV2 = Nothing
    
    ucImages.Visible = False
    ucWord.Visible = False
    ucCache.Visible = False
    ucHistory.Visible = False
    
    SetParent ucImages.hwnd, 0
    SetParent ucWord.hwnd, 0
    SetParent ucCache.hwnd, 0
    SetParent ucHistory.hwnd, 0
    
    ucImages.Destory
    ucWord.Destory
    ucCache.Destory
    ucHistory.Destory
    
    tabSelect.RemoveAll
    
    Set mobjSel = Nothing
    Set mobjLinkEditor = Nothing
    Set mobjMainVideo = Nothing
    Set mobjLinkEditor = Nothing
    Set mObjNotify = Nothing
    Set mobjStudyInfo = Nothing
End Sub

Private Sub UserControl_Terminate()
On Error GoTo errhandle
    mblnIsValid = False
    
    Call SavePar
    
    Call Destory
Exit Sub
errhandle:
    Debug.Print "ucPacsHelper_Terminate Err:" & err.Description
End Sub
