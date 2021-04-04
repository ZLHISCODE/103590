VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucImagePreview 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   3795
   ScaleWidth      =   7605
   ToolboxBitmap   =   "ucImagePreview.ctx":0000
   Begin VB.Timer tmrLoad 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   3360
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":0312
            Key             =   "avi"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":15FC64
            Key             =   "aviDownLoad"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":2BF5B6
            Key             =   "wav"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":41EF08
            Key             =   "wavDownLoad"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImagePreview.ctx":57E85A
            Key             =   "fileDisconet"
         EndProperty
      EndProperty
   End
   Begin zl9PacsControl.ucSplitPage ucPage 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   6210
      _ExtentX        =   10504
      _ExtentY        =   582
      PageCount       =   0
      PageRecord      =   9
      AutoRedrawStyle =   0   'False
   End
   Begin DicomObjects.DicomViewer dcmMiniImage 
      Height          =   3135
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _Version        =   262147
      _ExtentX        =   12938
      _ExtentY        =   5530
      _StockProps     =   35
      BackColor       =   4210752
      UseScrollBars   =   0   'False
      UseMouseWheel   =   -1  'True
   End
   Begin VB.Menu menuPopup 
      Caption         =   "�Ҽ��˵�"
      Begin VB.Menu mnuSplitPageTool 
         Caption         =   "��ҳ����(&P)"
      End
      Begin VB.Menu mnuReUpLoad 
         Caption         =   "�����ϴ�(&S)"
      End
   End
End
Attribute VB_Name = "ucImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"
Private Const M_STR_FAILD_TAG As String = "FAILD"
Private Const CON_INT_DICOMSELECTWIDTH As Integer = 18  'dicomͼ���ɫѡ�п����Ͻǻ�ɫС���Ⱥ͸߶�

Public Enum tQueryLevel
    slAdvice = 0    'ҽ��ID
    slStudy = 1     '���
    slSeries = 2    '����
    slImage = 3     'ͼ��
    slLocal = 4     '����
End Enum

Public Enum TMoveType
    mtLast = 0
    mtNext = 1
    mtFirst = 2
    mtEnd = 3
End Enum

Private mblnIsDock As Boolean  '�Ƿ�������ڣ����ڷ�ҳ�ؼ���ʾ
Private mobjFile As New FileSystemObject
Private mstrQueryValue As String         '���ҽ��ID
Private mblnMoved As Boolean             '�����Ƿ�ת��
Private mslQueryLevel As tQueryLevel      'ͼ����ʾ����
Private mblnQueryTmpRecord As Boolean

Private mblnOnlyLoadReportImage As Boolean     'ΪTrueʱ���� ����ͼ�� �ֶ��еı���ͼ,��֮�������б���ͼ
Private mblnIsShowCheckbox As Boolean   '�Ƿ���ʾ��ѡ��
Private mblnEnable As Boolean           '�Ƿ�ɽ��б༭
Private mlngBigImageWay As Long         '��ͼ��ʾ��ʽ��0-����ʾ��ͼ��1-����ƶ�ʱ��ʾ��ͼ��2-��굥��ʱ��ʾ��ͼ
Private mlngPreViewTime As Long         '�ƶ�Ԥ����ʱʱ��
Private mblnIsShowPopup As Boolean      '�Ƿ���ʾ�Ҽ��˵�
Private mtyFileLoadType As FileLoadType
Private mblnIsAutoHidePageControl As Boolean
Private mblnShowPageControl As Boolean
Private mlngFailedLoadCount As Long     'ʧ�ܼ��ش���
Private mblnIsLoadReportImage As Boolean '�Ǹ��ݱ���ͼ���ֶμ��صı���ͼ
Private mrsRecord As ADODB.Recordset
Private mlngMouseMoveZoom As Double     '�����ͼ�����ƶ�ʱ����ʾ��ͼ�ķŴ��������Ϊ0����ʾ��ͼ
Private mblnBigImageCtl As Boolean      '��ͼ��ʾ���ƣ�True--�����õķֱ��ʽ��д�С����

Private mblnDo As Boolean      '��ʱ������������ʱ���βɼ�����ͼ�����汨��ͼ����

Private WithEvents mobjImageProcess As clsImageProcess
Attribute mobjImageProcess.VB_VarHelpID = -1

Private mMultiCols As Long
Private mMultiRows As Long

Private mlngSelectIndex As Long
Private mobjFailedImgs As New Scripting.Dictionary    '����ʧ�ܵ�ͼ�񼯺�

Private mblnClickCheckState As Boolean
Private mintImage As Integer


Public Event OnSelChange(ByVal lngSelectedIndex As Long)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnClick(ByVal lngSelectedIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
Public Event OnSaveImage(ByVal dcmImage As DicomImage, ByVal lngImageType As Long)
Public Event OnReUpload()
Public Event AfterSaveStudy()

Private Sub DoOnSelChange(ByVal lngSelectedIndex As Long)
On Error Resume Next
    RaiseEvent OnSelChange(lngSelectedIndex)

    err.Clear
End Sub


Private Sub DoOnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
On Error Resume Next
    RaiseEvent OnCheckChange(lngSelectedIndex, blnSelected)
    
    err.Clear
End Sub

Private Sub DoOnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
On Error Resume Next

    RaiseEvent OnDbClick(lngSelectedIndex, blnContinue)

    err.Clear
End Sub

Private Sub DoOnClick(ByVal lngSelectedIndex As Long)
On Error Resume Next
    
    RaiseEvent OnClick(lngSelectedIndex)
    
    err.Clear
End Sub

'��ʱ���ԣ�������ʱ���βɼ�����ͼ�����汨��ͼ����
Property Get DoShield() As Boolean
    DoShield = mblnDo
End Property

Property Let DoShield(value As Boolean)
    mblnDo = value
End Property

'
Property Get AutoRedrawStyle() As Boolean
    AutoRedrawStyle = AutoRedraw
End Property

Property Let AutoRedrawStyle(value As Boolean)
    AutoRedraw = value
    
    ucPage.AutoRedrawStyle = value
End Property

'����ƶ���ͼ���ϵķŴ��������Ϊ0�򲻽��зŴ�
Property Get MouseMoveZoom() As Double
    MouseMoveZoom = mlngMouseMoveZoom
End Property

Property Let MouseMoveZoom(value As Double)
    mlngMouseMoveZoom = value
End Property

'��ͼ��ʾ�������õķֱ��ʽ��д�С����
Property Get BigImageCtl() As Boolean
    BigImageCtl = mblnBigImageCtl
End Property

Property Let BigImageCtl(value As Boolean)
    mblnBigImageCtl = value
End Property

Property Get OnlyLoadReportImage() As Boolean
    OnlyLoadReportImage = mblnOnlyLoadReportImage
End Property

Property Let OnlyLoadReportImage(value As Boolean)
    mblnOnlyLoadReportImage = value
End Property

'�Ƿ���ʾͼ��ѡ��
Property Get ShowCheckBox() As Boolean
    ShowCheckBox = mblnIsShowCheckbox
End Property

Property Let ShowCheckBox(value As Boolean)
    mblnIsShowCheckbox = value
End Property


'��ͼ��ʾ��ʽ
Property Get BigImageWay() As Long
    BigImageWay = mlngBigImageWay
End Property

Property Let BigImageWay(value As Long)
    mlngBigImageWay = value
End Property

'Ԥ����ʱʱ��
Property Get PreViewTime() As Long
    PreViewTime = mlngPreViewTime
End Property

Property Let PreViewTime(value As Long)
    mlngPreViewTime = value
End Property

'�Ƿ���ʾ�Ҽ��˵�
Property Get ShowPopup() As Boolean
    ShowPopup = mblnIsShowPopup
End Property

Property Let ShowPopup(value As Boolean)
    mblnIsShowPopup = value
End Property

'�Ƿ�ɽ��б༭
Property Get Enable() As Boolean
    Enable = mblnEnable
End Property

Property Let Enable(value As Boolean)
    mblnEnable = value
End Property

'��ȡͼ��������
Property Get ImgViewer() As Object
    Set ImgViewer = dcmMiniImage
End Property

'ͼ������
Property Get ImageTotal() As Long
    ImageTotal = ucPage.RecordCount
End Property

'��ȡ�ؼ����
Property Get Handle() As Long
    Handle = UserControl.hWnd
End Property

'�ļ����ط�ʽ
Property Get ImgLoadType() As FileLoadType
    ImgLoadType = mtyFileLoadType
End Property

Property Let ImgLoadType(value As FileLoadType)
    mtyFileLoadType = value
    mnuReUpLoad.Visible = value = Service
End Property

'�Զ����ط�ҳ�������ÿҳ����ʾ����С��ָ����ÿҳ��ʾ����ʱ
Property Get AutoHidePageControl() As Boolean
    AutoHidePageControl = mblnIsAutoHidePageControl
End Property


Property Let AutoHidePageControl(value As Boolean)
    mblnIsAutoHidePageControl = value
End Property


'��ѯ����ֵ
Property Get QueryValue() As String
    QueryValue = mstrQueryValue
End Property

Property Let QueryValue(value As String)
    mstrQueryValue = value
End Property


'��Ŀ�Ƿ�ѡ��
Property Get ImgChecked(Index As Long) As Boolean
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    ImgChecked = False
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            ImgChecked = Not objLabs(i).Transparent And objLabs(i).Visible
            Exit Property
        End If
    Next i
End Property

Property Let ImgChecked(Index As Long, value As Boolean)
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            objLabs(i).Transparent = Not value
            Call dcmMiniImage.Images(Index).Refresh(False)
            
            Exit Property
        End If
    Next i
End Property

Property Get CellSpacing() As Long
    CellSpacing = dcmMiniImage.CellSpacing
End Property

Property Let CellSpacing(value As Long)
    dcmMiniImage.CellSpacing = value
End Property


'ÿҳͼ����ʾ����
Property Get PageImgCount() As Long
    PageImgCount = ucPage.PageRecord
End Property

Property Let PageImgCount(value As Long)
    ucPage.PageRecord = value
End Property

'��ǰҳ��
Property Get PageNumber() As Long
    PageNumber = ucPage.PageNumber
End Property

'������ɫ
Property Get BackColor() As OLE_COLOR
    BackColor = dcmMiniImage.BackColour
End Property


Property Let BackColor(value As OLE_COLOR)
    dcmMiniImage.BackColour = value
End Property

'��ȡ��ǰѡ�е�����
Property Get SelectIndex() As Long
    SelectIndex = mlngSelectIndex
End Property


'��ȡ��ǰѡ�е�ͼ��
Property Get SelectImage() As DicomImage
    Set SelectImage = Nothing
    
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        Set SelectImage = dcmMiniImage.Images(mlngSelectIndex)
    End If
End Property


'��ȡ��ǰ��ʾ��ͼ������
Property Get CurImageCount() As Long
    CurImageCount = dcmMiniImage.Images.Count
End Property


Public Sub RedrawSelf()
'�ػ����
    Call dcmMiniImage.Refresh
    Call ucPage.RedrawSelf
End Sub

Public Sub RefreshFace(ByVal IsDock As Boolean)
'ˢ�½���ؼ�λ��
    mblnIsDock = IsDock
    Call UserControl_Resize
End Sub


Public Sub ShowPageConfig()
'��ʾ��ҳ����
    Call ShowPageControl
End Sub

Public Sub MovePage(ByVal lngMoveType As TMoveType)
'�ƶ�/��תͼ��ҳ��
    Select Case lngMoveType
        Case mtLast
            Call ucPage.LastPage
        Case mtNext
            Call ucPage.NextPage
        Case mtFirst
            Call ucPage.FirstPage
        Case mtEnd
            Call ucPage.EndPage
    End Select
End Sub

Public Sub RefreshImage(ByVal slQueryLevel As tQueryLevel, _
    ByVal strQueryValue As String, _
    ByVal blnMoved As Boolean, _
    Optional ByVal blnFoceRefresh As Boolean = False, _
    Optional ByVal blnTmpRecord As Boolean = False)
    
'ˢ��ͼ����ʾ
    Dim rsData As ADODB.Recordset
    Dim blnLoadResult As Boolean
    Dim i As Long
    
BUGEX "RefreshImage 1"
    mnuReUpLoad.Enabled = False
    If mstrQueryValue = strQueryValue And Not blnFoceRefresh And slQueryLevel <> slLocal Then
        mslQueryLevel = slQueryLevel
        Exit Sub
    End If
    
    mstrQueryValue = strQueryValue
    mslQueryLevel = slQueryLevel
    mblnQueryTmpRecord = blnTmpRecord
    mblnMoved = blnMoved
    
    If slQueryLevel = slLocal Then
        If mobjFile.FolderExists(strQueryValue) = False Then
            MsgBox "���ػ���Ŀ¼�����ڣ�", vbExclamation, CON_STR_HINT_TITLE
            Exit Sub
        End If
    End If
    
    ucPage.RecordCount = 0
    mlngSelectIndex = 0
    
BUGEX "RefreshImage 2"
    Call RefreshPageControl
    
BUGEX "RefreshImage 3"
        
    If strQueryValue = "" Then
        '���ͼ��
        Call ClearCurrentPageImage
        Exit Sub
    End If
    
BUGEX "RefreshImage 4"
    '���÷�ҳ���
    If slQueryLevel = slLocal Then
        Call ConfigPageControlWithLocal(strQueryValue)
    Else
        If mblnOnlyLoadReportImage Then
            Call ConfigRptPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
        Else
            Call ConfigPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
        End If
    End If
    
BUGEX "RefreshImage 5"
    
    '����ͼ����Ϣ
'    For i = 1 To dcmMiniImage.Images.Count
'        dcmMiniImage.Images(i).BorderColour = vbWhite
'        ImgChecked(i) = False
'    Next

    ChangeImgSelected dcmMiniImage, 1, True
    blnLoadResult = LoadImage(1, ucPage.PageRecord)
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
    
BUGEX "RefreshImage 6"
     '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
        
    Call UserControl_Resize
    
    If blnLoadResult = True Then Call dcmMiniImage.Refresh
    
BUGEX "RefreshImage End"
End Sub

Public Sub RefreshLabelTag()
    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
End Sub

Private Sub RefreshPageControl()
'ˢ�·�ҳ�����ʾ
On Error Resume Next
    If Not mblnIsAutoHidePageControl Then Exit Sub
    
    mblnShowPageControl = IIf(ucPage.RecordCount <= ucPage.PageRecord, False, True)
    ucPage.Visible = mblnShowPageControl
    
    Call UserControl_Resize
    
    err.Clear
End Sub

Public Sub ClearCurrentPageImage()
'���ͼ��
On Error GoTo errHandle
    mlngSelectIndex = 0
    
    dcmMiniImage.Images.Clear
    dcmMiniImage.MultiColumns = 1
    dcmMiniImage.MultiRows = 1
    
Exit Sub
errHandle:
    MsgboxEx hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE
    err.Clear
End Sub


Private Sub ConfigImgDisplayFormat(ByVal lngPageRecord As Long)
'����ͼ����ʾ��ʽ
    Dim iRows As Integer
    Dim iCols As Integer
    
    ResizeRegion lngPageRecord, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    
    mMultiCols = iCols
    mMultiRows = iRows

    dcmMiniImage.MultiColumns = iCols
    dcmMiniImage.MultiRows = iRows
End Sub

'�Ƿ����ʧ�ܵ�ͼ��
Public Function IsFailedImg(Index As Long) As Boolean
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    IsFailedImg = False
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_FAILD_TAG Then
            IsFailedImg = True
            Exit Function
        End If
    Next i
End Function

Private Function SyncDelImage(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As String
'ͬ��ɾ����ͼ��
    Dim i As Long
    Dim strImageInstanceUid As String
    Dim blnIsDel As Boolean
    
    SyncDelImage = ""
    blnIsDel = False
    For i = dcmViewer.Images.Count To 1 Step -1
        strImageInstanceUid = dcmViewer.Images(i).InstanceUID
        
        rsCurImageData.Filter = "ͼ��UID ='" & strImageInstanceUid & "'"
        
        If rsCurImageData.RecordCount <= 0 Then
            dcmViewer.Images.Remove (i)
            blnIsDel = True
        Else
            SyncDelImage = SyncDelImage & ";" & strImageInstanceUid & ";"
            
            '��������ͼ��ı���ͼ��ǣ���Ϊˢ����Ƶ�ɼ�����ʱ�����ͼ���Ѿ����ڣ��������¼���ͼ����Ϣ����������������
            dcmViewer.Images(i).Tag.ReportImage = NVL(rsCurImageData!����ͼ)
        End If
    Next i
    
    If blnIsDel = True Then dcmViewer.Refresh
    
    rsCurImageData.Filter = ""
End Function

Private Function SendDataToservice(ByVal strDataTag As String, ByVal intCommandIdentify As Integer, ByVal strDataFrom As String, fileMsg As TransferFileMsg)
    Dim objServiceHelper As New clsServiceHelper
    
    SendDataToservice = objServiceHelper.SendDataToservice(strDataTag, intCommandIdentify, strDataFrom, fileMsg)
    
    Set objServiceHelper = Nothing
End Function

Private Function LoadViewImageToFaceWithService(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
'ͨ��ZLPacsServerCenter�������Ԥ��ͼ�񵽽���
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim fileMsg As TransferFileMsg
    Dim blnIsSendOk As Boolean
    
    blnIsAddImage = False
    mlngSelectIndex = 0
    mlngFailedLoadCount = 0
    
    LoadViewImageToFaceWithService = False
    
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = SyncDelImage(rsCurImageData, dcmViewer)
        
    '����ͼ����ʾ��ʽ
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
        
    '��������ͼ�񻺴�Ŀ¼
    MkLocalDir GetResourceDir
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(NVL(rsCurImageData!ͼ��UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If NVL(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Trim(strTmpFile)
            
            blnIsSendOk = True
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ļ����ݷ�����������ʹ�÷����̨����
                With fileMsg
                    fileMsg.strAdviceId = Val(NVL(rsCurImageData("ҽ��ID")))
                    fileMsg.strName = NVL(rsCurImageData("����"))
                    fileMsg.strSex = NVL(rsCurImageData("�Ա�"))
                    fileMsg.strAge = NVL(rsCurImageData("����"))
                    
                    fileMsg.ftpInfo.strDeviceId = NVL(rsCurImageData("�豸��1"))
                    fileMsg.ftpInfo.strFtpDir = NVL(rsCurImageData("Root1"))
                    fileMsg.ftpInfo.strFTPIP = NVL(rsCurImageData("Host1"))
                    fileMsg.ftpInfo.strFTPPwd = NVL(rsCurImageData("Pwd1"))
                    fileMsg.ftpInfo.strFTPUser = NVL(rsCurImageData("User1"))
                    fileMsg.ftpInfo.strSDDir = NVL(rsCurImageData("����Ŀ¼1"))
                    fileMsg.ftpInfo.strSDPswd = NVL(rsCurImageData("����Ŀ¼����1"))
                    fileMsg.ftpInfo.strSDUser = NVL(rsCurImageData("����Ŀ¼�û���1"))
                    
                    fileMsg.bakFtpInfo.strDeviceId = NVL(rsCurImageData("�豸��2"))
                    fileMsg.bakFtpInfo.strFtpDir = NVL(rsCurImageData("Root2"))
                    fileMsg.bakFtpInfo.strFTPIP = NVL(rsCurImageData("Host2"))
                    fileMsg.bakFtpInfo.strFTPPwd = NVL(rsCurImageData("Pwd2"))
                    fileMsg.bakFtpInfo.strFTPUser = NVL(rsCurImageData("User2"))
                    fileMsg.bakFtpInfo.strSDDir = NVL(rsCurImageData("����Ŀ¼2"))
                    fileMsg.bakFtpInfo.strSDPswd = NVL(rsCurImageData("����Ŀ¼����2"))
                    fileMsg.bakFtpInfo.strSDUser = NVL(rsCurImageData("����Ŀ¼�û���2"))
                    
                    fileMsg.strLocalDir = strTmpFile
                    fileMsg.strFileName = NVL(rsCurImageData("ͼ��UID")) & IIf(mblnIsLoadReportImage, ".jpg", "")
                    fileMsg.strSubDir = NVL(rsCurImageData("URL"))
                    fileMsg.strMediaType = NVL(rsCurImageData!��̬ͼ, imgTag)
                End With
                
                If Not SendDataToservice("����ͼ", LoadCommand.COMMAND_RPTIMG_DOWNLOAD, "ͼ������", fileMsg) Then
                    blnIsSendOk = False
                End If
            End If
            
            If NVL(rsCurImageData!��̬ͼ, imgTag) <> VIDEOTAG And NVL(rsCurImageData("��̬ͼ"), imgTag) <> AUDIOTAG Then
                '����ͼ����
                Set dcmTag = New clsImageTagInf
                dcmTag.Tag = NVL(rsCurImageData!��̬ͼ, imgTag)
                
                If Dir(strTmpFile) = vbNullString Then
                    If Dir(strCachePath & "\fileDisconet.bmp") = vbNullString Then
                        Call SavePicture(imgList.ListImages("fileDisconet").Picture, strCachePath & "\fileDisconet.bmp")
                    End If
                    
                    Set curImage = dcmViewer.Images.AddNew
                    Call curImage.FileImport(strCachePath + "fileDisconet.bmp", "DIB/BMP")
                    curImage.InstanceUID = strImgInstanceUid
                    
                    Dim imgLoadInfo As New DicomLabel
                    Dim iCols As Long, iRows As Long
                    
                    iCols = dcmViewer.MultiColumns
                    iRows = dcmViewer.MultiRows
                    
                    If blnIsSendOk Then
                        imgLoadInfo.Text = "[" + NVL(rsCurImageData!�豸��1, NVL(rsCurImageData!�豸��2)) + "] �ļ�������..."
                    Else
                        imgLoadInfo.Text = "[" + NVL(rsCurImageData!�豸��1, NVL(rsCurImageData!�豸��2)) + "] �ļ���������ʧ��."
                    End If
                                        
                    imgLoadInfo.Width = dcmViewer.Width
                    imgLoadInfo.Height = 20
                    
                    imgLoadInfo.Left = 0
                    imgLoadInfo.Top = dcmViewer.Height / Screen.TwipsPerPixelY / iRows - imgLoadInfo.Height * 2

                    imgLoadInfo.AutoSize = True
                    imgLoadInfo.ShowTextBox = False
                    imgLoadInfo.Font.Size = 12
                    imgLoadInfo.Font.Bold = True
                    imgLoadInfo.ForeColour = vbRed
                    imgLoadInfo.Tag = M_STR_FAILD_TAG
                    
                    Call curImage.Labels.Add(imgLoadInfo)
                    
                    '��ʧ�ܵ�ͼ����뼯����
                    If mobjFailedImgs.Exists(strImgInstanceUid) Then Call mobjFailedImgs.Remove(strImgInstanceUid)
                    Call mobjFailedImgs.Add(strImgInstanceUid, strTmpFile)
                Else
                    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
                End If
                                    
                Set curImage.Tag = dcmTag
                
                With curImage
                    .BorderStyle = 6
                    .BorderWidth = 1
                    .BorderColour = vbWhite
                End With
            Else
                Set curImage = New DicomImage
                    
                If Dir(strTmpFile) = vbNullString Then
                    If NVL(rsCurImageData("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                        Call SavePicture(imgList.ListImages("avi").Picture, strTmpFile)
                    Else
                        Call SavePicture(imgList.ListImages("wav").Picture, strTmpFile)
                    End If
                End If

                Call curImage.FileImport(strTmpFile, "DIB/BMP")

                Set dcmTag = New clsImageTagInf

                dcmTag.Tag = NVL(rsCurImageData!��̬ͼ, VIDEOTAG)
                dcmTag.EncoderName = NVL(rsCurImageData("��������"), "")
                dcmTag.CaptureTime = NVL(rsCurImageData("�ɼ�ʱ��"))
                
                If NVL(rsCurImageData("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                    dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".avi"
                Else
                    dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".wav"
                End If
                
                dcmTag.RecordTimeLen = Val(NVL(rsCurImageData("¼�Ƴ���"), "0"))
                
                Set curImage.Tag = dcmTag
                
                curImage.InstanceUID = NVL(rsCurImageData("ͼ��UID"))
                curImage.SeriesUID = NVL(rsCurImageData("����UID"))
                curImage.StudyUID = NVL(rsCurImageData("���UID"))
                
                Call ShowAVInf(curImage, dcmTag)
                
                With curImage
                    .BorderStyle = 6
                    .BorderWidth = 1
                    .BorderColour = vbWhite
                End With
                
                Call dcmViewer.Images.Add(curImage)
            End If
            
            'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
            '���½�ú��DSAͼ����������ʾ
            '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
            '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
            If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                curImage.Attributes.Remove &H28, &H6100
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    Call UpdateSelectIndex(1)
    
    If Dir(strCachePath & "\fileDisconet.bmp") <> vbNullString Then objFile.DeleteFile (strCachePath & "\fileDisconet.bmp")
    If mobjFailedImgs.Count > 0 Then tmrLoad.Enabled = True '����Timer����ʼ����
    
    LoadViewImageToFaceWithService = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
End Function

Private Function ReadDicomFile(ByVal strFile As String, dcmImgs As DicomImages) As DicomImage
On Error Resume Next
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    If blnUseUrl Then
        'readurl��֧�ֿո�
        Set curImage = dcmImgs.ReadURL(strFile)
    Else
        Set curImage = dcmImgs.ReadFile(strFile)
    End If
    
    If err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    '2098����һ�����ļ�����dicom�ļ�����һ���Ǵ��ڹ�����ʴ���
    If InStr(err.Description, "sharing violation") > 0 Then
        err.Clear
        strFileTime = Format(Now, "YYMMDD") & GetTickCount
        
        Call FileCopy(strFile, strFile & "_copy_vdat_" & strFileTime)
    
        If blnUseUrl Then
            'readurl��֧�ֿո�
            Set curImage = dcmImgs.ReadURL(strFile & "_copy_vdat_" & strFileTime)
        Else
            Set curImage = dcmImgs.ReadFile(strFile & "_copy_vdat_" & strFileTime)
        End If
    
        If err.Number = 0 Then
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
            err.Clear
        Else
            Call Kill(strFile & "_copy_vdat_" & strFileTime)
        End If
    Else
        err.Clear
        Set curImage = dcmImgs.AddNew
        Call curImage.FileImport(strFile, "JPG")
        
        If err.Number <> 0 Then
            err.Clear
            'not a JPG file
            Call curImage.FileImport(strFile, "BMP")
        End If
        
        If err.Number <> 0 Then
            err.Clear
            'not a BMP file
            Call curImage.FileImport(strFile, "AVI")
        End If
        
        If err.Number <> 0 Then
            err.Clear
            'not a AVI file
            Call curImage.FileImport(strFile, "MPG")
        End If
    End If
    
    If err.Number = 0 Then
        Set ReadDicomFile = curImage
        Exit Function
    End If
    
    Set ReadDicomFile = Nothing
    
err.Clear
End Function

Private Function ReadViewImage(ByVal strFile As String, Optional ByRef dcmViewer As DicomViewer = Nothing) As DicomImage
On Error GoTo errHandle
    Dim dImgs As DicomImages
        
    '�������_copy_vdat_��˵������ʱ�ļ�
    If InStr(strFile, "_copy_vdat_") > 0 Then
        Set ReadViewImage = Nothing
        Call Kill(strFile)
        
        Exit Function
    End If
    
    If dcmViewer Is Nothing Then
        Set dImgs = New DicomImages
    Else
        Set dImgs = dcmViewer.Images
    End If
    
    Set ReadViewImage = ReadDicomFile(strFile, dImgs)
    
Exit Function
errHandle:
    Set ReadViewImage = Nothing
End Function


Private Function LoadViewImageToFaceWithNormal(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
'����Ԥ��ͼ�񵽽���
    Dim strTmpFile As String
    Dim strCachePath As String
    
    Dim curImage As DicomImage
    
    Dim objFile As New Scripting.FileSystemObject
    
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    
    Dim dcmTag As clsImageTagInf
    Dim strImgInstanceUid As String
    Dim strCurInstanceUids As String
    Dim blnIsAddImage As Boolean
    Dim objImgInfo As Object
    
BUGEX "LoadViewImageToFaceWithNormal 1"

    blnIsAddImage = False
    mlngSelectIndex = 0
    
    LoadViewImageToFaceWithNormal = False
        
BUGEX "LoadViewImageToFaceWithNormal 2"
    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    strCurInstanceUids = SyncDelImage(rsCurImageData, dcmViewer)
        
    '����ͼ����ʾ��ʽ
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
        
    '��������ͼ�񻺴�Ŀ¼
    strCachePath = GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(NVL(rsCurImageData("URL")))
    
BUGEX "LoadViewImageToFaceWithNormal 3"
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Trim(NVL(rsCurImageData!ͼ��UID))
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 And strImgInstanceUid <> "" Then
            
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If NVL(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = GetResourceDir & "Avi.bmp"
            ElseIf NVL(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & NVL(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            
            strTmpFile = Replace(Trim(strTmpFile), "/", "\")
            
            If Dir(strTmpFile) = vbNullString Then
                '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                '����FTP����
                If NVL(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                    If Inet1.FuncFtpConnect(NVL(rsCurImageData("Host1")), NVL(rsCurImageData("User1")), NVL(rsCurImageData("Pwd1"))) = 0 Then
                        If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.FuncFtpConnect(NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))) = 0 Then
                                MsgboxEx hWnd, "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                                Exit Function
                            End If
                        Else
                            MsgboxEx hWnd, "FTP�����������ӣ������������á�", vbOKOnly, CON_STR_HINT_TITLE
                            Exit Function
                        End If
                    End If
                End If
                
                If Inet1.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""), , hWnd) <> 0 Then
                    '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                    If NVL(rsCurImageData("�豸��2")) <> vbNullString Then
                        If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsCurImageData("Host2")), NVL(rsCurImageData("User2")), NVL(rsCurImageData("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(NVL(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")) & IIf(mblnIsLoadReportImage, ".jpg", ""), , hWnd)
                        End If
                    End If
                End If
    
    BUGEX "LoadViewImageToFaceWithNormal DCM TmpFile:" & strTmpFile
    
            If Dir(strTmpFile) <> vbNullString Then
                If NVL(rsCurImageData!��̬ͼ, imgTag) <> VIDEOTAG And NVL(rsCurImageData("��̬ͼ"), imgTag) <> AUDIOTAG Then
                    
    BUGEX "LoadViewImageToFaceWithNormal Dcm ReadURL"
                    
                    Set curImage = ReadViewImage(strTmpFile, dcmViewer)
                    
                    '����ͼ����
                    Set dcmTag = New clsImageTagInf
                    dcmTag.Tag = NVL(rsCurImageData!��̬ͼ, imgTag)
                    dcmTag.ReportImage = NVL(rsCurImageData!����ͼ)
                                       
                    Set curImage.Tag = dcmTag
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                Else
                    Set curImage = New DicomImage
                    
                    If Dir(strTmpFile) = vbNullString Then
                        If NVL(rsCurImageData("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                            Call SavePicture(imgList.ListImages("avi").Picture, strTmpFile)
                        Else
                            Call SavePicture(imgList.ListImages("wav").Picture, strTmpFile)
                        End If
                    End If
                    
                    Call curImage.FileImport(strTmpFile, "DIB/BMP")
                    Set dcmTag = New clsImageTagInf
                    
BUGEX "LoadViewImageToFaceWithNormal DCM Set Pro."

                    dcmTag.Tag = NVL(rsCurImageData!��̬ͼ, VIDEOTAG)
                    dcmTag.EncoderName = NVL(rsCurImageData("��������"), "")
                    dcmTag.CaptureTime = NVL(rsCurImageData("�ɼ�ʱ��"))
                    dcmTag.ReportImage = NVL(rsCurImageData!����ͼ)
                    
                    If NVL(rsCurImageData("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                        dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".avi"
                    Else
                        dcmTag.VideoFile = strCachePath & NVL(rsCurImageData("URL")) & ".wav"
                    End If
                    
                    dcmTag.RecordTimeLen = Val(NVL(rsCurImageData("¼�Ƴ���"), "0"))
                    
'                        '�������Ƶ¼���ļ������ڲ���ʱ��������
'                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
'                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
'                        End If
                    
                    Set curImage.Tag = dcmTag
                    
                    curImage.InstanceUID = NVL(rsCurImageData("ͼ��UID"))
                    curImage.SeriesUID = NVL(rsCurImageData("����UID"))
                    curImage.StudyUID = NVL(rsCurImageData("���UID"))
                    
                    Call ShowAVInf(curImage, dcmTag)
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
    BUGEX "LoadViewImageToFaceWithNormal DCM AddImage"
                    Call dcmViewer.Images.Add(curImage)
                End If
                
                
                'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
                '���½�ú��DSAͼ����������ʾ
                '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
                '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
                If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
                    curImage.Attributes.Remove &H28, &H6100
                End If
            End If
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    Call UpdateSelectIndex(1)
    
    Inet1.FuncFtpDisConnect
    Inet2.FuncFtpDisConnect
    
    LoadViewImageToFaceWithNormal = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
    
BUGEX "LoadViewImageToFaceWithNormal End"
End Function

Private Function LoadViewImageToFaceFromLocal(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
    Dim strTmpFile As String
    Dim curImage As DicomImage
    Dim dcmTag As clsImageTagInf
    
On Error GoTo ErrorHand

    If rsCurImageData.RecordCount <= 0 Then
        dcmViewer.Images.Clear
        
        dcmViewer.MultiColumns = 1
        dcmViewer.MultiRows = 1
        
        Exit Function
    End If
    
    '����ͼ����ʾ��ʽ
    If rsCurImageData.RecordCount <> dcmViewer.Images.Count Then
        Call ConfigImgDisplayFormat(IIf(rsCurImageData.RecordCount < ucPage.PageRecord, rsCurImageData.RecordCount, ucPage.PageRecord))
    End If
  
    Do While Not rsCurImageData.EOF
        strTmpFile = Trim(NVL(rsCurImageData!·��))
        
        Set curImage = ReadViewImage(strTmpFile, dcmMiniImage)
        
        '����ͼ����
        Set dcmTag = New clsImageTagInf
        dcmTag.Tag = imgTag
        dcmTag.FilePath = strTmpFile
                            
        Set curImage.Tag = dcmTag
        
        With curImage
            .BorderStyle = 6
            .BorderWidth = 1
            .BorderColour = vbWhite
        End With
        
        rsCurImageData.MoveNext
    Loop
    
     Call UpdateSelectIndex(1)
     
     LoadViewImageToFaceFromLocal = True
     Exit Function
ErrorHand:
    LoadViewImageToFaceFromLocal = False
    BUGEX "LoadViewImageToFaceFromLocal err = " & err.Description
End Function


Public Sub PlayMedia(ByVal lngMediaIndex As Long)
'����ָ����������ý��

End Sub

Private Sub ConfigPageControlWithLocal(ByVal strQueryPath As String)
    Dim objFile As File
    
    If mobjFile.FolderExists(strQueryPath) = False Then Exit Sub
    
    ucPage.RecordCount = mobjFile.GetFolder(strQueryPath).Files.Count
End Sub

Private Sub ConfigRptPageControl(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'���÷�ҳ�ؼ�
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    strSQL = "Select Count(B.Column_Value) ����ֵ From Ӱ�����¼ A, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) B Where ҽ��ID = [1]"
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
    If rsData.RecordCount > 0 Then lngRecordCount = NVL(rsData!����ֵ)
    
    If lngRecordCount <= 0 Then
        Select Case slQueryLevel
            Case slAdvice
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and nvl(a.��̬ͼ,0)=0 and c.ҽ��ID=[1]"
            Case slStudy
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and nvl(a.��̬ͼ,0)=0 and b.���UID=[1]"
            Case slSeries
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ����UID=[1]"
            Case slImage
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where nvl(��̬ͼ,0)=0 and ͼ��UID=[1]"
        End Select
        
        '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
        If blnTmpRecord Then
            strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
        Else
            If mblnMoved Then
                strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
                strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
            End If
        End If
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
        
        If rsData.RecordCount > 0 Then
            lngRecordCount = NVL(rsData!����ֵ)
        Else
            lngRecordCount = 0
        End If
    End If
    
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub

Private Sub ConfigPageControl(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'���÷�ҳ�ؼ�
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1]"
        Case slStudy
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
        Case slSeries
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ����UID=[1]"
        Case slImage
            strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ��  where  ͼ��UID=[1]"
    End Select
    
    '�����ѯ��ʱ��¼������Ҫ����ѯ���滻Ϊ��ʱ�洢���ݵı�
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
        End If
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = NVL(rsData!����ֵ)
    Else
        lngRecordCount = 0
    End If
    
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub

Private Function GetImageViewDataFromLocal(ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
    Dim objFile As File
    Dim strDatas() As String
    Dim lngTmpCount As Long
    Dim rsData As New ADODB.Recordset
    Dim lngStartRecord As Long, lngEndRecord As Long
    
    If mobjFile.FolderExists(mstrQueryValue) = False Then Exit Function
    If mobjFile.GetFolder(mstrQueryValue).Files.Count <= 0 Then Exit Function
    
    rsData.Fields.Append "·��", adVarChar, 4000
    rsData.Open
    
    For Each objFile In mobjFile.GetFolder(mstrQueryValue).Files
        lngTmpCount = lngTmpCount + 1
        ReDim Preserve strDatas(lngTmpCount - 1) As String
        strDatas(UBound(strDatas)) = objFile.Path
    Next
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    For lngTmpCount = lngStartRecord - 1 To lngEndRecord - 1
        If UBound(strDatas) >= lngTmpCount Then
            rsData.AddNew
            rsData!·�� = strDatas(lngTmpCount)
            rsData.Update
        Else
            Exit For
        End If
    Next
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Set GetImageViewDataFromLocal = rsData
End Function

Private Function GetImageRptData(ByVal lngOrderID As Long, ByVal lngCurPage As Long, ByVal lngPageRecord As Long) As ADODB.Recordset
'���ݱ���ͼ�� �ֶλ�ȡ���ͼ��
    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    strSQL = "Select rownum As ˳���,a.ҽ��id,a.����,a.�Ա�,a.����, rownum As ͼ���,Replace(Trim(D.Column_Value),'.jpg','') as ͼ��UID, A.���UID, " & _
            "'' As ����UID, 0 as ��̬ͼ,'' as ��������,'' as �ɼ�ʱ��, '' as ¼�Ƴ���, '' as ����ͼ," & _
            "B.FTP�û��� As User1,B.FTP���� As Pwd1,B.IP��ַ As Host1,'/'||B.FtpĿ¼||'/' As Root1, " & _
            "B.����Ŀ¼ as ����Ŀ¼1,B.����Ŀ¼�û��� as ����Ŀ¼�û���1,B.����Ŀ¼���� as ����Ŀ¼����1, " & _
            "Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/') ||A.���UID||'/'||Replace(Trim(D.Column_Value),'.jpg','') As URL,B.�豸�� as �豸��1, B.�豸�� as �豸��1, " & _
            "C.FTP�û��� As User2,C.FTP���� As Pwd2,C.IP��ַ As Host2,'/'||C.FtpĿ¼||'/' As Root2, " & _
            "C.����Ŀ¼ as ����Ŀ¼2,C.����Ŀ¼�û��� as ����Ŀ¼�û���2,C.����Ŀ¼���� as ����Ŀ¼����2,C.�豸�� as �豸��2, C.�豸�� as �豸��2 " & _
            "From Ӱ�����¼ A, Ӱ���豸Ŀ¼ B, Ӱ���豸Ŀ¼ C, Table(Cast(f_Str2list(A.����ͼ��,';') As zlTools.t_Strlist)) D " & _
            "Where A.λ��һ = B.�豸��(+) And A.λ�ö� = C.�豸��(+) And A.ҽ��id = [1]"
            
    If mblnMoved = True Then strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ˳���>=" & lngStartRecord & " and ˳���<=" & lngEndRecord
    
    Set GetImageRptData = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ͼ��", lngOrderID)
End Function

Private Function GetImageViewData(ByVal slQueryLevel As tQueryLevel, ByVal strSearchValue As String, _
    ByVal lngCurPage As Long, ByVal lngPageRecord As Long, ByVal blnTmpRecord As Boolean) As ADODB.Recordset
'��ȡԤ��ͼ������
'intSearchType:0-�����uid����,1-������UID����,2-��ͼ��UID����

    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    strSQL = "Select rownum as ˳���,[2] ҽ��id,c.����,c.�Ա�,c.����, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,D.����Ŀ¼ as ����Ŀ¼1,D.����Ŀ¼�û��� as ����Ŀ¼�û���1,D.����Ŀ¼���� as ����Ŀ¼����1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/') " & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, D.�豸�� As �豸��1," & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2,A.����ͼ," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,E.����Ŀ¼ as ����Ŀ¼2,E.����Ŀ¼�û��� as ����Ŀ¼�û���2,E.����Ŀ¼���� as ����Ŀ¼����2," & _
            "E.�豸�� as �豸��2, E.�豸�� As �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+)" & IIf(mblnOnlyLoadReportImage, " And nvl(A.��̬ͼ,0) = 0 ", "")
    
    If blnTmpRecord Then
        strSQL = Replace(strSQL, "Ӱ����", "Ӱ����ʱ")
    Else
        If mblnMoved Then
            strSQL = Replace(strSQL, "Ӱ����ͼ��", "HӰ����ͼ��")
            strSQL = Replace(strSQL, "Ӱ��������", "HӰ��������")
            strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        End If
    End If

    Select Case slQueryLevel
        Case slAdvice
            strSQL = "select * from (" & strSQL & " and C.ҽ��ID=[1])"
        Case slStudy
            strSQL = "select * from (" & strSQL & " and C.���UID=[1])"
        Case slSeries
            strSQL = "select * from (" & strSQL & " and B.����UID=[1])"
        Case slImage
            strSQL = "select * from (" & strSQL & " and A.ͼ��UID=[1])"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ˳���>=" & lngStartRecord & " and ˳���<=" & lngEndRecord
    
    Set GetImageViewData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ����Ϣ", strSearchValue, IIf(mblnQueryTmpRecord, "-1", mstrQueryValue))
End Function

Public Sub AddImage(Img As Object, Optional objImgTag As Object = Nothing)
'����ͼ��
    Dim i As Long
    
    If dcmMiniImage.Images.Count < ucPage.PageRecord Then
        Call ConfigImgDisplayFormat(dcmMiniImage.Images.Count + 1)
        
        Call dcmMiniImage.Images.Add(Img)
    Else
        '�ƶ�ͼ��
        For i = 2 To dcmMiniImage.Images.Count
            Call dcmMiniImage.Images.Move(i, i - 1)
            dcmMiniImage.Images(i - 1).BorderColour = vbWhite
        Next i
        
        Call dcmMiniImage.Images.Remove(dcmMiniImage.Images.Count)
        dcmMiniImage.Images.Add Img
    End If
    
    '����ѡ�еı߿���ɫ
    With dcmMiniImage.Images(dcmMiniImage.Images.Count)
        .BorderWidth = 1
        .BorderStyle = 6
        .BorderColour = vbRed
        
        If Not objImgTag Is Nothing Then
            Set .Tag = objImgTag
        End If
        
        If Not .Tag Is Nothing Then
            If UCase(TypeName(.Tag)) = UCase("clsImageTagInf") Then Call ShowAVInf(dcmMiniImage.Images(dcmMiniImage.Images.Count), .Tag)
        End If

    End With
    
    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
    
    Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    Call UpdateImageCount(1)
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
End Sub


Private Sub ShowAVInf(Img As DicomImage, objImgTag As Object)
'��ʾ����Ƶ��Ϣ
    If objImgTag.Tag = VIDEOTAG Or objImgTag.Tag = AUDIOTAG Then
        Call AddVideoLabelToDicomImage(Img, _
        IIf(objImgTag.Tag = VIDEOTAG, "¼��ʱ�䣺", "¼��ʱ�䣺") & objImgTag.CaptureTime, _
        IIf(objImgTag.Tag = VIDEOTAG, "¼�񳤶ȣ�", "¼�����ȣ�") & objImgTag.RecordTimeLen & " ��", _
        "�������ƣ�" & objImgTag.EncoderName)
    End If
End Sub

Public Sub DeleteImage(ByVal lngImgIndex As Long, Optional blMovePage As Boolean = True, Optional blMustMovePage As Boolean = False)
'ɾ��ͼ�� blMovePage:�Ƿ��ж��Զ���ҳ blMustMovePage�Ƿ�ǿ�Ʒ�ҳ
    Dim i As Long
    Dim lngCurPageCount As Long
        
    
    Call dcmMiniImage.Images.Remove(lngImgIndex)
    
    For i = lngImgIndex + 1 To dcmMiniImage.Images.Count
        Call dcmMiniImage.Move(i, i - 1)
    Next i

    Call dcmMiniImage.Refresh
    
    If lngImgIndex <= dcmMiniImage.Images.Count Then
        Call UpdateSelectIndex(lngImgIndex)
    Else
        Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    End If
    
    
    lngCurPageCount = ucPage.PageCount
    
    Call UpdateImageCount(-1)
        
    If lngCurPageCount > ucPage.PageCount Then
        If blMovePage Then
            Call ucPage.MovePage(ucPage.PageNumber)
            If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
        End If
    Else
        If blMovePage And blMustMovePage Then

            Call ucPage.MovePage(ucPage.PageNumber)
            If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
        End If
    End If
    
    For i = 1 To dcmMiniImage.Images.Count
        If i <> mlngSelectIndex Then dcmMiniImage.Images(i).BorderColour = vbWhite
    Next
    
    mnuReUpLoad.Enabled = dcmMiniImage.Images.Count > 0
End Sub

Private Sub UpdateSelectIndex(ByVal lngSelectIndex As Long)
'����ͼ���ѡ������
    Dim blnIsValidIndex As Boolean
    
    blnIsValidIndex = IIf(lngSelectIndex > 0 And lngSelectIndex <= dcmMiniImage.Images.Count, True, False)
    
    If Not blnIsValidIndex Then Exit Sub

    If blnIsValidIndex Then dcmMiniImage.Images(lngSelectIndex).BorderColour = vbRed
    If mlngSelectIndex = lngSelectIndex Then Exit Sub

    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        dcmMiniImage.Images(mlngSelectIndex).BorderColour = vbWhite
    End If

    mlngSelectIndex = lngSelectIndex
    
    'ִ�������ı��¼�
    Call DoOnSelChange(mlngSelectIndex)
End Sub


Private Sub UpdateImageCount(ByVal lngValue As Long)
    ucPage.RecordCount = ucPage.RecordCount + lngValue
    
    Call RefreshPageControl
End Sub


Public Function SelectedCount() As Long
'��ȡѡ���ͼ������
    Dim i As Long
    Dim j As Long
    Dim lngCount As Long
    Dim objLabs As DicomLabels
    
    
    lngCount = 0
    For i = 1 To dcmMiniImage.Images.Count
        Set objLabs = dcmMiniImage.Images(i).Labels
        
        For j = 1 To objLabs.Count
            If objLabs(j).Tag = M_STR_SELECT_TAG Then
                If Not objLabs(j).Transparent Then lngCount = lngCount + 1
                Exit For
            End If
        Next j
    Next i
    
    SelectedCount = lngCount
End Function



Private Sub dcmMiniImage_Click()
On Error GoTo errHandle
    Dim i As Integer
    
    If mlngSelectIndex <= 0 Or mlngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    
    Call DoOnClick(mlngSelectIndex)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_DblClick()
On Error GoTo errHandle
    Dim blnContinue As Boolean
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    If mlngSelectIndex <= 0 Then Exit Sub

    blnContinue = True
    
    If mlngBigImageWay = 1 Then  '�رմ�ͼ��ʾ
        ReleaseCapture      '�������
'        frmShowImg.HideMe
    End If
    
    Call DoOnDbClick(mlngSelectIndex, blnContinue)
    
    ImgChecked(mlngSelectIndex) = mblnClickCheckState
    
    If Not blnContinue Then Exit Sub

    
    If dcmMiniImage.MultiColumns = 1 And dcmMiniImage.MultiRows = 1 Then
        dcmMiniImage.MultiColumns = mMultiCols
        dcmMiniImage.MultiRows = mMultiRows
        dcmMiniImage.CurrentIndex = 1
    Else
        mMultiCols = dcmMiniImage.MultiColumns
        mMultiRows = dcmMiniImage.MultiRows
        
        dcmMiniImage.MultiColumns = 1
        dcmMiniImage.MultiRows = 1
        
        dcmMiniImage.CurrentIndex = mlngSelectIndex
    End If
    
    Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub


Private Sub dcmMiniImage_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    
    If KeyCode = 37 Then '<-
        lngSelectIndex = mlngSelectIndex - 1
        If lngSelectIndex <= 0 Then Exit Sub
    ElseIf KeyCode = 38 Then
        lngSelectIndex = mlngSelectIndex - dcmMiniImage.MultiColumns
        If lngSelectIndex <= 0 Then Exit Sub
    ElseIf KeyCode = 39 Then
        lngSelectIndex = mlngSelectIndex + 1
        If lngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    ElseIf KeyCode = 40 Then
        lngSelectIndex = mlngSelectIndex + dcmMiniImage.MultiColumns
        If lngSelectIndex > dcmMiniImage.Images.Count Then Exit Sub
    Else
        Exit Sub
    End If
    
    Call UpdateSelectIndex(lngSelectIndex)
    
    If mblnEnable Then
        For i = 1 To dcmMiniImage.Images.Count
            If i <> lngSelectIndex Then
                ImgChecked(i) = False
            Else
                ImgChecked(i) = True
            End If
        Next i
    End If
    
    Call DoOnClick(lngSelectIndex)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    Dim i As Long
    Dim objLabs As DicomLabels
    Dim lngImgIndex As Long
    Dim blnClickCheck As Boolean
    
    mblnClickCheckState = ImgChecked(mlngSelectIndex)

        lngImgIndex = dcmMiniImage.ImageIndex(X, Y)
        
        Call UpdateSelectIndex(lngImgIndex)
        
        If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
            
            If mblnEnable Then
                            
                '����ѡ���״̬
                blnClickCheck = False
                Set objLabs = dcmMiniImage.LabelHits(X, Y, False, True, True)
                For i = 1 To objLabs.Count
                    If objLabs(i).Tag = M_STR_SELECT_TAG And objLabs(i).Visible Then
                        '��objLabs(i).Visible=false��˵��ѡ�п��Ѿ������أ�����ѡ�д���
                        blnClickCheck = True

                        objLabs(i).Transparent = Not objLabs(i).Transparent
            
                        Call dcmMiniImage.Images(lngImgIndex).Refresh(False)
                        
                        '����ͼ��ѡ�¼�
                        Call DoOnCheckChange(mlngSelectIndex, Not objLabs(i).Transparent)
                        
                        Exit For
                    End If
                Next i
                
                            '��ȡ��ѡ��
                If Shift = 0 Then
                    If blnClickCheck = False Then
                        If Button = 2 Then
                            If Not ImgChecked(lngImgIndex) Then
                                ChangeImgSelected dcmMiniImage, lngImgIndex, False
                            End If
                        ElseIf Button = 1 Then
                            ChangeImgSelected dcmMiniImage, lngImgIndex, False
                            ImgChecked(lngImgIndex) = Not ImgChecked(lngImgIndex)
                        End If
                    End If
                Else
                    If blnClickCheck = False And Button = 1 Then ImgChecked(lngImgIndex) = Not ImgChecked(lngImgIndex)
                End If
                
            End If
        End If

    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub

Private Sub ChangeImgSelected(dcmImage As DicomViewer, lngImage As Long, blnChaBorderColor As Boolean)
    Dim i As Long
       
    For i = 1 To dcmImage.Images.Count
    
        If i <> lngImage Then ImgChecked(i) = False '�ı�Check���ѡ��
        
        If blnChaBorderColor Then dcmMiniImage.Images(i).BorderColour = vbWhite '�ı�߿���ɫ
    Next
End Sub

Private Sub dcmMiniImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)

On Error GoTo errHandle
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer
    
    If mlngBigImageWay <> 1 Then Exit Sub
    
    '�ж��Ƿ���Ҫ��ʾͼ��
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImage.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImage.Height) Then
        blnShowImg = True
    End If
    
    If blnShowImg Then        '��ʾͼ��
        SetCapture dcmMiniImage.hWnd    '�������
        
        intCurrImg = dcmMiniImage.ImageIndex(X, Y)
        
        
        If intCurrImg <> 0 And intCurrImg <> mintImage Then
            If dcmMiniImage.Images(intCurrImg).Tag.Tag <> VIDEOTAG And dcmMiniImage.Images(intCurrImg).Tag.Tag <> AUDIOTAG Then
            '����ͼ����ʾ
            
                If mobjImageProcess Is Nothing Then
                    Set mobjImageProcess = New clsImageProcess
                End If
                
                mobjImageProcess.ShowImageProcess mstrQueryValue, dcmMiniImage.Images(intCurrImg), ucPage.PageRecord * (ucPage.PageNumber - 1) + intCurrImg, Me, mblnMoved, mslQueryLevel, 1, mlngPreViewTime, mblnDo
    '            frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(intCurrImg)), Me, 1, 0, 0, BigImageCtl, mlngMouseMoveZoom
                
            End If
            
        End If
        mintImage = intCurrImg
    Else
        ReleaseCapture
        mintImage = 0
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Function GetBigImage(dcmImg As DicomImage) As DicomImage
    
    Set GetBigImage = dcmImg.SubImage(0, 0, dcmImg.SizeX, dcmImg.SizeY, 1, dcmImg.Frame)
     
    GetBigImage.Labels.Clear
    GetBigImage.BorderColour = vbWhite
End Function

Private Sub dcmMiniImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim curPointer As POINTAPI
    Dim i As Integer
    
    If mlngBigImageWay = 1 Then  '�رմ�ͼ��ʾ
        ReleaseCapture      '�������
'        frmShowImg.HideMe
    End If
    
    If Button = 2 And mblnIsShowPopup Then
        '��ʾ�Ҽ��˵�
        Call GetCursorPos(curPointer)
        
        Call ScreenToClient(hWnd, curPointer)  'ScreenToClient����ʹ�õĵ�λΪ����ֵ
        Call PopupMenu(menuPopup, 0, ScaleX(curPointer.X, vbPixels, vbTwips), ScaleY(curPointer.Y, vbPixels, vbTwips))
        
    Else
        '��ʾ��ͼ
        If mlngBigImageWay = 2 And Button = 1 Then
            
            If dcmMiniImage.Images.Count > 0 Then

                i = dcmMiniImage.ImageIndex(X, Y)
                If i = 0 Then i = 1
                
                If dcmMiniImage.Images(i).Tag.Tag <> VIDEOTAG And dcmMiniImage.Images(i).Tag.Tag <> AUDIOTAG Then
                '����ͼ����ʾ
                
                    If mobjImageProcess Is Nothing Then
                        Set mobjImageProcess = New clsImageProcess
                    End If
                    
                    mobjImageProcess.ShowImageProcess mstrQueryValue, dcmMiniImage.Images(i), ucPage.PageRecord * (ucPage.PageNumber - 1) + i, Me, mblnMoved, mslQueryLevel, 1, 0, mblnDo
        '            frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(intCurrImg)), Me, 1, 0, 0, BigImageCtl, mlngMouseMoveZoom
                End If
'                frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(i)), Me, 2, 0, 0, BigImageCtl
            End If
        End If
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
errHandle:
End Sub

Private Sub dcmMiniImage_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
    If Delta > 0 Then
        Call ucPage.LastPage
    Else
        Call ucPage.NextPage
    End If
    
    RaiseEvent OnMouseWheel(Shift, Delta, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub mnuReUpLoad_Click()
'�����ϴ�ѡ����ļ�
On Error GoTo errHandle
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        RaiseEvent OnReUpload
    End If
    
    Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Private Sub mnuSplitPageTool_Click()
'��ʾ��ҳ������
    Call ShowPageControl
End Sub

Private Sub mobjImageProcess_AfterSaveStady()
    RaiseEvent AfterSaveStudy
End Sub

Private Sub mobjImageProcess_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    RaiseEvent OnSaveImage(dcmImage, lngImageType)
End Sub

Private Sub mobjImageProcess_OnUnload()
    Set mobjImageProcess = Nothing
End Sub

Private Sub tmrLoad_Timer()
'ʹ�ú�̨��������ͼ��ʱ�������ӳ٣�����Timer�м���֮ǰδ���ص�ͼ��
    Dim i As Long, j As Long
    Dim strTmpFile As String
    Dim strTmpKey, dcmTag As Object
    Dim objTmpImg As DicomImage
    Dim iCols As Long, iRows As Long
    Dim strDevice As String
On Error GoTo errHandle
    
    If mobjFailedImgs Is Nothing Then Exit Sub
    If mobjFailedImgs.Count <= 0 Or mlngFailedLoadCount > 30 Then
        If mobjFailedImgs.Count > 0 Then
            iCols = dcmMiniImage.MultiColumns
            iRows = dcmMiniImage.MultiRows
            
            '�������δ���سɹ�����Ϊ����ʧ��
            For Each strTmpKey In mobjFailedImgs.Keys
                For i = 1 To dcmMiniImage.Images.Count
                    If strTmpKey = dcmMiniImage.Images(i).InstanceUID Then
                        For j = 1 To dcmMiniImage.Images(i).Labels.Count
                            If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_FAILD_TAG Then
                                strDevice = dcmMiniImage.Images(i).Labels(j).Text
                                
                                strDevice = Mid(strDevice, 1, InStr(strDevice, "]"))
                                dcmMiniImage.Images(i).Labels(j).Text = strDevice + "�ļ�����ʧ��."
                            End If
                        Next
                        
                        dcmMiniImage.Refresh
                        Exit For
                    End If
                Next
            Next
        End If
        
        tmrLoad.Enabled = False
        Exit Sub
    End If
    
    mlngFailedLoadCount = mlngFailedLoadCount + 1
    
    For Each strTmpKey In mobjFailedImgs.Keys
        strTmpFile = mobjFailedImgs(strTmpKey)
        
        '�����ص����أ����滻ԭ���ı��ͼƬ
        If Dir(strTmpFile) <> vbNullString Then
            For i = 1 To dcmMiniImage.Images.Count
                If strTmpKey = dcmMiniImage.Images(i).InstanceUID Then
                    Set dcmTag = dcmMiniImage.Images(i).Tag
                                        
                    Set objTmpImg = ReadViewImage(strTmpFile)
                    If err.Number <> 0 Then
                        err.Clear
                        Exit For
                    End If
                    
                    Call dcmMiniImage.Images.Remove(i)
                    
                    Set objTmpImg.Tag = dcmTag '����ͼ����

                    With objTmpImg
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                    
                    '����ѡ���
                    If mblnIsShowCheckbox Then Call DrawItemCheckBorder(objTmpImg)
                    '������ͼ���
                    Call DrawReportImgTag(objTmpImg)
                    
                    Call dcmMiniImage.Images.Add(objTmpImg)
                    
                    Call dcmMiniImage.Images.Move(dcmMiniImage.Images.Count, i)
                    Call mobjFailedImgs.Remove(strTmpKey)
                    
                    Exit For
                End If
            Next
        End If
    Next
    
    Exit Sub
errHandle:
End Sub


Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo errHandle
    Call LoadImage(lngPageIndex, lngPageCount)
    
    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
        
    Call UserControl_Resize
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

'��������ʧ��ͼ��
Public Sub ReLoadFailedImage()

    Call dcmMiniImage.Images.Clear
     
    Call LoadImage(1, ucPage.PageRecord)

    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
End Sub


Private Function LoadImage(ByVal lngPageIndex As Long, ByVal lngPageCount As Long, Optional ByVal blnGetPath As Boolean) As Boolean
    Dim rsData As ADODB.Recordset

On Error GoTo errHandle
    LoadImage = True
    
    If mstrQueryValue = "0" Then Exit Function
    
    If mslQueryLevel = slLocal Then
        Set rsData = GetImageViewDataFromLocal(lngPageIndex, lngPageCount)
    Else
        If mblnOnlyLoadReportImage Then
            '���� Ӱ�����¼.����ͼ�� �ֶ��е�ֵ���أ����Ϊ�գ� ���������б���ͼ��
            Set rsData = GetImageRptData(mstrQueryValue, lngPageIndex, lngPageCount)
            
            mblnIsLoadReportImage = rsData.RecordCount > 0
            
            If rsData.RecordCount <= 0 Then
                Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
            End If
        Else
            Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
        End If
    End If
    
    If blnGetPath Then
        Set mrsRecord = rsData
        Exit Function
    End If
    If rsData Is Nothing Then Exit Function
        
    If mslQueryLevel = slLocal Then
        LoadImage = LoadViewImageToFaceFromLocal(rsData, dcmMiniImage)
    Else
        If ImgLoadType = FileLoadType.Normal Then
            LoadImage = LoadViewImageToFaceWithNormal(rsData, dcmMiniImage)     'ʹ��ԭʼģʽ����
        Else
            LoadImage = LoadViewImageToFaceWithService(rsData, dcmMiniImage)    'ʹ��ZLPacsServerCenter����,��̨����
        End If
    End If
    
    Exit Function
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Function

Private Sub UserControl_Initialize()
    
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True
    
    mblnIsShowCheckbox = False
    mblnIsShowPopup = False
    mblnShowPageControl = False
    
    mlngBigImageWay = 0
    
    mstrQueryValue = ""
    mlngSelectIndex = 0
    
    mnuReUpLoad.Visible = False
    mnuReUpLoad.Enabled = False
    
    ucPage.PageRecord = 5
    mblnIsAutoHidePageControl = True
End Sub




Public Sub ClearChecked()
'���ѡ��
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ImgChecked(i) = False
    Next i
End Sub



Public Sub SelectedAll()
'ȫѡ
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ImgChecked(i) = True
    Next i
End Sub



Private Sub UserControl_Resize()
    Dim iCols As Integer, iRows As Integer
    Dim i As Long, j As Long
    Dim Img As DicomImage
    Dim sngW As Single '�ƿ�ռͼ�����
    
On Error Resume Next
    
    dcmMiniImage.Left = 0
    dcmMiniImage.Top = 0
    dcmMiniImage.Width = UserControl.ScaleWidth
    If mblnIsDock Then
        dcmMiniImage.Height = UserControl.ScaleHeight - IIf(mblnShowPageControl, ucPage.Height + 480, 420)
    Else
        dcmMiniImage.Height = UserControl.ScaleHeight - IIf(mblnShowPageControl, ucPage.Height + 60, 0)
    End If
    
    ucPage.Left = 0
    ucPage.Top = dcmMiniImage.Height + 30
    
    ResizeRegion dcmMiniImage.Images.Count, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    dcmMiniImage.MultiColumns = iCols
    dcmMiniImage.MultiRows = iRows
    
    '�ж��Ƿ�ƿ�ռ��ͼƬ����20%
    If dcmMiniImage.Images.Count > 0 Then
        Set Img = dcmMiniImage.Images(mlngSelectIndex)
        sngW = CON_INT_DICOMSELECTWIDTH / (Img.SizeX * Img.ActualZoom)
    End If

    If sngW > 0.2 Then
        'δ��ѡͼ���һƿ�ռ��ͼƬ����20%����Ҫ����ѡ�п�
        For i = 1 To dcmMiniImage.Images.Count
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_SELECT_TAG Then dcmMiniImage.Images(i).Labels(j).Visible = False
            Next
        Next
    Else
        '��ʾѡ�п�
        For i = 1 To dcmMiniImage.Images.Count
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_SELECT_TAG Then dcmMiniImage.Images(i).Labels(j).Visible = True
            Next
        Next
    End If

    '��ǰ���ܱ���
    For i = 1 To dcmMiniImage.Images.Count
        If mobjFailedImgs.Exists(dcmMiniImage.Images(i).InstanceUID) Then
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_FAILD_TAG Then
                    dcmMiniImage.Images(i).Labels(j).Left = 0
                    dcmMiniImage.Images(i).Labels(j).Top = dcmMiniImage.Height / Screen.TwipsPerPixelY / iRows - dcmMiniImage.Images(i).Labels(j).Height * 2
                End If
            Next
        End If
    Next

    err.Clear
End Sub



Private Function GetImageRow(ByVal lngImageIndex As Long) As Integer
'ȡ�õ�ǰ������
    GetImageRow = CInt(lngImageIndex / dcmMiniImage.MultiColumns) + 1
End Function

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub


Private Sub DrawItemCheckBorder(dcmImg As DicomImage)
    Dim lSelect As DicomLabel
    Dim lBorder As DicomLabel
    
    Set lBorder = New DicomLabel

    With lBorder
        .LabelType = 2            '�߿�
        .Width = 1000
        .Height = 1000
        .Left = 0
        .Top = 0
        .LineWidth = 2
    
    
        .ForeColour = vbYellow
        .BackColour = vbYellow
    
    
        .Transparent = True
        .ScaleWithCell = True
        .Tag = M_STR_BORDER_TAG
    
        .Visible = True
    End With
    
    dcmImg.Labels.Add lBorder
    


    Set lSelect = New DicomLabel
    
    With lSelect
        .LabelType = 2            '����
        .Width = CON_INT_DICOMSELECTWIDTH
        .Height = CON_INT_DICOMSELECTWIDTH
        .Left = 1
        .Top = 1
        .LineWidth = 2
        
        .ForeColour = vbYellow
        .BackColour = vbRed
        
                
        .Transparent = True
        .ScaleWithCell = False
        .ImageTied = False
    
        .Tag = M_STR_SELECT_TAG
        
        .Visible = True
    End With
    
    dcmImg.Labels.Add lSelect
    
    dcmImg.BorderStyle = vbRed
End Sub

Public Sub DrawReportImgTag(dcmImg As DicomImage)
    Dim lRpt As DicomLabel
    Dim i As Integer
    
    
    If dcmImg.Tag.ReportImage <> "" Then
        Set lRpt = New DicomLabel
                
        With lRpt
            .LabelType = doLabelText
            .Width = 300
            .Height = 80
            .ImageTied = False
            .Transparent = True
            .ScaleWithCell = True
            .ScaleFontSize = 40
            .Font.Name = "����"
            .Font.Size = 40
            .Font.Bold = True
            .ForeColour = vbWhite
            .BackColour = vbRed
            .Left = 350
            .Top = 20
            .Text = "����ͼ"
            .ShowTextBox = True
            .Shadow = doShadowBottomRight
            .Alignment = doAlignCentre
            .Visible = True
            .Tag = "����ͼ"
        End With
        
        dcmImg.Labels.Add lRpt
    Else
        For i = 1 To dcmImg.Labels.Count
            '����Ƴ���һ����ע����ע��������٣��ж��Ƿ��Ѿ����������б�ע������i��һ
            If i > dcmImg.Labels.Count Then Exit For
            If dcmImg.Labels(i).Tag = "����ͼ" Then
                Call dcmImg.Labels.Remove(i)
                i = i - 1
            End If
        Next i
    End If
    
    dcmImg.Refresh False
End Sub

Private Sub DrawImageLabels(dcmViewer As DicomViewer)
'����ͼ��ĸ��ֱ�ע
    Dim i As Long

    'ѭ��ÿһ��ͼ�񣬻���ע
    For i = 1 To dcmViewer.Images.Count
        '��ѡ���
        If mblnIsShowCheckbox Then
            Call DrawItemCheckBorder(dcmViewer.Images(i))
        End If
        '������ͼ���
        Call DrawReportImgTag(dcmViewer.Images(i))
    Next i
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    mblnIsShowCheckbox = PropBag.ReadProperty("ShowCheckbox", False)
    mblnIsShowPopup = PropBag.ReadProperty("ShowPopup", False)
    ucPage.PageRecord = PropBag.ReadProperty("PageImgCount", 5)
    AutoRedraw = PropBag.ReadProperty("AutoRedrawStyle", False)
    mlngMouseMoveZoom = PropBag.ReadProperty("MouseMoveZoom", 0)
    
    ucPage.AutoRedrawStyle = AutoRedraw
    
    err.Clear
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("CellSpacing", dcmMiniImage.CellSpacing, 3)
    Call PropBag.WriteProperty("BackColor", dcmMiniImage.BackColour, vbBlack)
    Call PropBag.WriteProperty("Enable", mblnEnable, True)
    Call PropBag.WriteProperty("ShowCheckbox", mblnIsShowCheckbox, False)
    Call PropBag.WriteProperty("ShowPopup", mblnIsShowPopup, False)
    Call PropBag.WriteProperty("PageImgCount", ucPage.PageRecord, 5)
    Call PropBag.WriteProperty("AutoRedrawStyle", AutoRedraw, False)
    Call PropBag.WriteProperty("MouseMoveZoom", mlngMouseMoveZoom, 0)
    
    err.Clear
End Sub

Private Sub ShowPageControl()
'��ʾ��ҳ������
On Error GoTo errHandle
    mblnShowPageControl = True
    ucPage.Visible = mblnShowPageControl
    
    Call UserControl_Resize
errHandle:
End Sub

'��ʱ��������ȡ����ͼ�񻺴�·��
Public Function GetPathString() As String
    Dim strTmpFile As String

    Call LoadImage(1, ucPage.RecordCount, True)
    
    If mrsRecord Is Nothing Then Exit Function
    If mrsRecord.RecordCount <= 0 Then Exit Function
    
    strTmpFile = ""
    mrsRecord.MoveFirst
    If mslQueryLevel = slLocal Then
        Do While Not mrsRecord.EOF
            strTmpFile = strTmpFile & "|" & Trim(NVL(mrsRecord!·��))
            mrsRecord.MoveNext
        Loop
    Else
        '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
        Do While Not mrsRecord.EOF
            If NVL(mrsRecord!��̬ͼ, imgTag) <> VIDEOTAG And NVL(mrsRecord!��̬ͼ, imgTag) <> AUDIOTAG Then
                strTmpFile = strTmpFile & "|" & GetCacheDir & NVL(mrsRecord("URL")) & IIf(mblnIsLoadReportImage, ".jpg", "")
            End If
            mrsRecord.MoveNext
        Loop
    End If
    
    Set mrsRecord = Nothing
    GetPathString = strTmpFile
End Function

Public Sub AfterSaveStudy(dcmImage As DicomImage)
    If Not mobjImageProcess Is Nothing Then
        mobjImageProcess.AfterSaveStudy dcmImage
    End If
End Sub


