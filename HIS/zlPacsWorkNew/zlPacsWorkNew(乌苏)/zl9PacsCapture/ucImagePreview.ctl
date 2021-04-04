VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.UserControl ucImagePreview 
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   3795
   ScaleWidth      =   7605
   ToolboxBitmap   =   "ucImagePreview.ctx":0000
   Begin zl9PacsCapture.ucSplitPage ucPage 
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
   End
End
Attribute VB_Name = "ucImagePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"

Public Enum TQueryLevel
    slStudy = 0     '���
    slSeries = 1    '����
    slImage = 2     'ͼ��
End Enum


Public Enum TMoveType
    mtLast = 0
    mtNext = 1
    mtFirst = 2
    mtEnd = 3
End Enum


Private mstrQueryValue As String         '���ҽ��ID
Private mblnMoved As Boolean             '�����Ƿ�ת��
Private mslQueryLevel As TQueryLevel      'ͼ����ʾ����
Private mblnQueryTmpRecord As Boolean

Private mblnIsShowCheckbox As Boolean   '�Ƿ���ʾ��ѡ��
Private mblnEnable As Boolean           '�Ƿ�ɽ��б༭
Private mlngMouseMoveZoom As Double     '�����ͼ�����ƶ�ʱ����ʾ��ͼ�ķŴ��������Ϊ0����ʾ��ͼ
Private mlngBigImageWay As Long         '��ͼ��ʾ��ʽ��0-����ʾ��ͼ��1-����ƶ�ʱ��ʾ��ͼ��2-��굥��ʱ��ʾ��ͼ
Private mblnIsShowPopup As Boolean      '�Ƿ���ʾ�Ҽ��˵�
Private mblnIsAutoHidePageControl As Boolean
Private mblnShowPageControl As Boolean

Private mMultiCols As Long
Private mMultiRows As Long

Private mlngSelectIndex As Long


Public Event OnSelChange(ByVal lngSelectedIndex As Long)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnClick(ByVal lngSelectedIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)

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

'
Property Get AutoRedrawStyle() As Boolean
    AutoRedrawStyle = AutoRedraw
End Property

Property Let AutoRedrawStyle(value As Boolean)
    AutoRedraw = value
    
    ucPage.AutoRedrawStyle = value
End Property

'�Ƿ���ʾͼ��ѡ��
Property Get ShowCheckBox() As Long
    ShowCheckBox = mblnIsShowCheckbox
End Property

Property Let ShowCheckBox(value As Long)
    mblnIsShowCheckbox = value
End Property

'����ƶ���ͼ���ϵķŴ��������Ϊ0�򲻽��зŴ�
Property Get MouseMoveZoom() As Double
    MouseMoveZoom = mlngMouseMoveZoom
End Property

Property Let MouseMoveZoom(value As Double)
    mlngMouseMoveZoom = value
End Property


'��ͼ��ʾ��ʽ
Property Get BigImageWay() As Long
    BigImageWay = mlngBigImageWay
End Property

Property Let BigImageWay(value As Long)
    mlngBigImageWay = value
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
            ImgChecked = Not objLabs(i).Transparent
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


Public Sub ShowPageConfig()
'��ʾ��ҳ����
    Call mnuSplitPageTool_Click
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

Public Sub RefreshImage(ByVal slQueryLevel As TQueryLevel, _
    ByVal strQueryValue As String, _
    ByVal blnMoved As Boolean, _
    Optional ByVal blnFoceRefresh As Boolean = False, _
    Optional ByVal blnTmpRecord As Boolean = False)
    
'ˢ��ͼ����ʾ
    Dim rsData As ADODB.Recordset
    Dim blnLoadResult As Boolean
    
BUGEX "RefreshImage 1"
    If mstrQueryValue = strQueryValue And Not blnFoceRefresh Then Exit Sub
    
    mstrQueryValue = strQueryValue
    mslQueryLevel = slQueryLevel
    mblnQueryTmpRecord = blnTmpRecord
    mblnMoved = blnMoved
    
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
    Call ConfigPageControl(slQueryLevel, strQueryValue, blnTmpRecord)
    
BUGEX "RefreshImage 5"
    '��ȡͼ������
    Set rsData = GetImageViewData(slQueryLevel, strQueryValue, 1, ucPage.PageRecord, blnTmpRecord)
    
    '����ͼ����Ϣ
    blnLoadResult = LoadViewImageToFace(rsData, dcmMiniImage)
    
BUGEX "RefreshImage 6"
    If mblnIsShowCheckbox Then
        '����ѡ���
        Call DrawImageSelectBorder(dcmMiniImage)
    End If
    
    If blnLoadResult = True Then Call dcmMiniImage.Refresh
    
BUGEX "RefreshImage End"
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
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
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
        End If
    Next i
    
    If blnIsDel = True Then dcmViewer.Refresh
    
    rsCurImageData.Filter = ""
End Function



Private Function LoadViewImageToFace(rsCurImageData As ADODB.Recordset, dcmViewer As DicomViewer) As Boolean
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
    
BUGEX "LoadViewImageToFace 1"

    blnIsAddImage = False
    mlngSelectIndex = 0
    
    LoadViewImageToFace = False
        
BUGEX "LoadViewImageToFace 2"
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
    strCachePath = zlCL_GetCacheDir
    MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsCurImageData("URL")))
    
    If rsCurImageData.RecordCount > 0 Then
        If gblnUseActivexLoad Then
            ReDim Preserve gobjGetImage(UBound(gobjGetImage) + 1) As Object
            Set gobjGetImage(UBound(gobjGetImage)) = DynamicCreate("zlGetImageEx.clsImageTransfer", "zlGetImageEx.exe")
            
            If Not gobjGetImage(UBound(gobjGetImage)) Is Nothing Then
                Call gobjGetImage(UBound(gobjGetImage)).RegEventObj(Me)
                Call gobjGetImage(UBound(gobjGetImage)).zlInitModule(False, UBound(gobjGetImage))
            End If
        End If
    End If
    
BUGEX "LoadViewImageToFace 3"
    Do While Not rsCurImageData.EOF
        'ѭ������ͼ��DicomViewer��
        strImgInstanceUid = Nvl(rsCurImageData!ͼ��UID)
        
        If InStr(strCurInstanceUids, strImgInstanceUid) <= 0 Then
            
            blnIsAddImage = True
            
            '��������Ƶ����ʾ�ļ������Ϊ����Ƶ�ļ�ʱ���ù��̽����ӷ�������ֱ�����������ļ�
            If Nvl(rsCurImageData!��̬ͼ, imgTag) = VIDEOTAG Then
                strTmpFile = zlCL_GetResourceDir & "Avi.bmp"
            ElseIf Nvl(rsCurImageData!��̬ͼ, imgTag) = AUDIOTAG Then
                strTmpFile = zlCL_GetResourceDir & "wav.bmp"
            Else
                strTmpFile = strCachePath & Nvl(rsCurImageData("URL"))
            End If
            
            If gblnUseActivexLoad Then
                If Dir(strTmpFile) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    '����ͼ�����ع��ߣ�zlGetImage.exe��ǰ�����ж��ܷ���������FTP
                    '����FTP����
                    If Nvl(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                            If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                    MsgboxCus "FTP�����������ӣ������������á�", vbOKOnly, G_STR_HINT_TITLE
                                    Exit Function
                                End If
                            Else
                                MsgboxCus "FTP�����������ӣ������������á�", vbOKOnly, G_STR_HINT_TITLE
                                Exit Function
                            End If
                        End If
                    End If
                End If
                
                If Not gobjGetImage(UBound(gobjGetImage)) Is Nothing Then
                    Set objImgInfo = gobjGetImage(UBound(gobjGetImage)).ImgInfo
                    
                    With objImgInfo
                        .MediaType = Val(Nvl(rsCurImageData("��̬ͼ")))
                        .EncoderName = Nvl(rsCurImageData("��������"))
                        .CaptureTime = Nvl(rsCurImageData("�ɼ�ʱ��"))
                        .SubDir = Nvl(rsCurImageData("URL"))
                        .RecordTimeLen = Val(Nvl(rsCurImageData("¼�Ƴ���")))
                        .InstanceUID = Nvl(rsCurImageData("ͼ��UID"))
                        .SeriesUID = Nvl(rsCurImageData("����UID"))
                        .StudyUID = Nvl(rsCurImageData("���UID"))
                        .TmpFilePath = strTmpFile
                        .IsLoadSingleFile = True
                        .IsUpLoad = False
                        .DestMainDir = App.Path & "\TmpImage\"
                    End With
                
                    If Dir(strTmpFile) = vbNullString Then
                        If rsCurImageData("�豸��1") <> "" Then
                            With objImgInfo
                                .FTPDir = Nvl(rsCurImageData("Root1"))
                                .IP = Nvl(rsCurImageData("Host1"))
                                .FTPPswd = Nvl(rsCurImageData("Pwd1"))
                                .FTPUser = Nvl(rsCurImageData("User1"))
                                .SDDir = Nvl(rsCurImageData("����Ŀ¼1"))
                                .SDPswd = Nvl(rsCurImageData("����Ŀ¼����1"))
                                .SDUser = Nvl(rsCurImageData("����Ŀ¼�û���1"))
                            End With
                        End If
                        
                        If rsCurImageData("�豸��2") <> "" Then
                            With objImgInfo
                                .BakFTPDir = Nvl(rsCurImageData("Root2"))
                                .BakIP = Nvl(rsCurImageData("Host2"))
                                .BakFTPPswd = Nvl(rsCurImageData("Pwd2"))
                                .BakFTPUser = Nvl(rsCurImageData("User2"))
                                .BakSDDir = Nvl(rsCurImageData("����Ŀ¼2"))
                                .BakSDPswd = Nvl(rsCurImageData("����Ŀ¼����2"))
                                .BakSDUser = Nvl(rsCurImageData("����Ŀ¼�û���2"))
                            End With
                        End If
                        
                        Call gobjGetImage(UBound(gobjGetImage)).MsgInQueue(objImgInfo)
                        
                    Else
                        Call OnComplete(objImgInfo)
                    End If
                End If
            Else
                If Dir(strTmpFile) = vbNullString Then
                    '���ػ���ͼ�񲻴��ڣ����ȡFTPͼ��
                    
                    '����FTP����
                    If Nvl(rsCurImageData("�豸��1")) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(Nvl(rsCurImageData("Host1")), Nvl(rsCurImageData("User1")), Nvl(rsCurImageData("Pwd1"))) = 0 Then
                            If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                                If Inet2.FuncFtpConnect(Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))) = 0 Then
                                    MsgboxCus "FTP�����������ӣ������������á�", vbOKOnly, G_STR_HINT_TITLE
                                    Exit Function
                                End If
                            Else
                                MsgboxCus "FTP�����������ӣ������������á�", vbOKOnly, G_STR_HINT_TITLE
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root1")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL"))) <> 0 Then
                        '���豸��1��ȡͼ��ʧ�ܣ�����豸��2��ȡͼ��
                        If Nvl(rsCurImageData("�豸��2")) <> vbNullString Then
                            If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect Nvl(rsCurImageData("Host2")), Nvl(rsCurImageData("User2")), Nvl(rsCurImageData("Pwd2"))
                            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsCurImageData("Root2")) & rsCurImageData("URL")), strTmpFile, objFile.GetFileName(rsCurImageData("URL")))
                        End If
                    End If
                End If
        
        BUGEX "LoadViewImageToFace DCM TmpFile:" & strTmpFile
        
                If Dir(strTmpFile) <> vbNullString Then
                    If Nvl(rsCurImageData!��̬ͼ, imgTag) <> VIDEOTAG And Nvl(rsCurImageData("��̬ͼ"), imgTag) <> AUDIOTAG Then
                        
        BUGEX "LoadViewImageToFace Dcm ReadFile"
                        Set curImage = dcmViewer.Images.ReadFile(strTmpFile)
                        
                        '����ͼ����
                        Set dcmTag = New clsImageTagInf
                        dcmTag.Tag = Nvl(rsCurImageData!��̬ͼ, imgTag)
                                            
                        Set curImage.Tag = dcmTag
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With
                    Else
                        Set curImage = New DicomImage
                        
                        On Error GoTo continue
        BUGEX "LoadViewImageToFace DCM ImportFile"
                            Call curImage.FileImport(strTmpFile, "DIB/BMP")
continue:
                            Set dcmTag = New clsImageTagInf
                            
        BUGEX "LoadViewImageToFace DCM Set Pro."
        
                            dcmTag.Tag = Nvl(rsCurImageData!��̬ͼ, VIDEOTAG)
                            dcmTag.EncoderName = Nvl(rsCurImageData("��������"), "")
                            dcmTag.CaptureTime = Nvl(rsCurImageData("�ɼ�ʱ��"))
                            
                            If Nvl(rsCurImageData("��̬ͼ"), VIDEOTAG) = VIDEOTAG Then
                                dcmTag.VideoFile = strCachePath & Nvl(rsCurImageData("URL")) & ".avi"
                            Else
                                dcmTag.VideoFile = strCachePath & Nvl(rsCurImageData("URL")) & ".wav"
                            End If
                            
                            dcmTag.RecordTimeLen = Val(Nvl(rsCurImageData("¼�Ƴ���"), "0"))
                            
        '                        '�������Ƶ¼���ļ������ڲ���ʱ��������
        '                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
        '                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
        '                        End If
                            
                            Set curImage.Tag = dcmTag
                            
                            curImage.InstanceUID = Nvl(rsCurImageData("ͼ��UID"))
                            curImage.SeriesUID = Nvl(rsCurImageData("����UID"))
                            curImage.StudyUID = Nvl(rsCurImageData("���UID"))
                            
                        
                        Call ShowAVInf(curImage, dcmTag)
                        
                        With curImage
                            .BorderStyle = 6
                            .BorderWidth = 1
                            .BorderColour = vbWhite
                        End With
                        
        BUGEX "LoadViewImageToFace DCM AddImage"
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
        End If
        
        rsCurImageData.MoveNext
    Loop
    
    If rsCurImageData.RecordCount > 0 Then
        If gblnUseActivexLoad Then
            If Not gobjGetImage(UBound(gobjGetImage)) Is Nothing Then
                Call gobjGetImage(UBound(gobjGetImage)).zlLoadImage
            End If
        End If
    End If
    
    Call UpdateSelectIndex(1)
    
    If Not gblnUseActivexLoad Then
        Inet1.FuncFtpDisConnect
        Inet2.FuncFtpDisConnect
    End If
    
    LoadViewImageToFace = IIf(Trim(strCurInstanceUids) <> "" And blnIsAddImage = True, True, False)
    
BUGEX "LoadViewImageToFace End"
End Function

Public Sub OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadId As Long)
    Call frmWork_Video.OnState(blnLoadFinish, blnUpLoad, lngThreadId)
End Sub

Public Sub OnComplete(ByVal objImgInfo As Object)
    Dim curImage As DicomImage
    Dim dcmTag As clsImageTagInf
    Dim strCachePath As String
    
    strCachePath = zlCL_GetCacheDir
BUGEX "LoadViewImageToFace DCM TmpFile:" & objImgInfo.TmpFilePath
    If Dir(objImgInfo.TmpFilePath) <> vbNullString Then
BUGEX " " & Nvl(objImgInfo.TmpFilePath, imgTag)
        If Nvl(objImgInfo.MediaType, imgTag) <> VIDEOTAG And Nvl(objImgInfo.MediaType, imgTag) <> AUDIOTAG Then
            
BUGEX "LoadViewImageToFace Dcm ReadFile"
            Set curImage = dcmMiniImage.Images.ReadFile(objImgInfo.TmpFilePath)
            
            '����ͼ����
            Set dcmTag = New clsImageTagInf
            dcmTag.Tag = Nvl(objImgInfo.MediaType, imgTag)
                                
            Set curImage.Tag = dcmTag
            
            With curImage
                .BorderStyle = 6
                .BorderWidth = 1
                .BorderColour = vbWhite
            End With
        Else
            Set curImage = New DicomImage
            
            On Error GoTo continue
BUGEX "LoadViewImageToFace DCM ImportFile"
                Call curImage.FileImport(objImgInfo.TmpFilePath, "DIB/BMP")
continue:
                Set dcmTag = New clsImageTagInf
                
BUGEX "LoadViewImageToFace DCM Set Pro."

                dcmTag.Tag = Nvl(objImgInfo.MediaType, VIDEOTAG)
                dcmTag.EncoderName = Nvl(objImgInfo.EncoderName, "")
                dcmTag.CaptureTime = Nvl(objImgInfo.CaptureTime)
                
                If Nvl(objImgInfo.MediaType, VIDEOTAG) = VIDEOTAG Then
                    dcmTag.VideoFile = strCachePath & Nvl(objImgInfo.SubDir) & ".avi"
                Else
                    dcmTag.VideoFile = strCachePath & Nvl(objImgInfo.SubDir) & ".wav"
                End If
                
                dcmTag.RecordTimeLen = Val(Nvl(objImgInfo.RecordTimeLen, "0"))
                
'                        '�������Ƶ¼���ļ������ڲ���ʱ��������
'                        If Trim(dcmTag.VideoFile) <> "" And Dir(dcmTag.VideoFile) <> "" Then
'                            Name dcmTag.VideoFile As dcmTag.VideoFile & ".avi"
'                        End If
                
                Set curImage.Tag = dcmTag
                
                curImage.InstanceUID = Nvl(objImgInfo.InstanceUID)
                curImage.SeriesUID = Nvl(objImgInfo.SeriesUID)
                curImage.StudyUID = Nvl(objImgInfo.StudyUID)
                
            
            Call ShowAVInf(curImage, dcmTag)
            
            With curImage
                .BorderStyle = 6
                .BorderWidth = 1
                .BorderColour = vbWhite
            End With
            
BUGEX "LoadViewImageToFace DCM AddImage"
            Call dcmMiniImage.Images.Add(curImage)
        End If
        
        'ȡ���Զ���Ӱ,��ΪDicomObjects�ؼ�����Դ����Ӱ��BUG�����ڣ�0028��6100��ʱ�����Զ���ͼ����м�Ӱ��
        '���½�ú��DSAͼ����������ʾ
        '��Ȼ����ͼ���mask=0 ,����ȡ����Ӱ������ÿ��ͼ����ӵ��µ�Dicomimages֮���Զ��ֽ�mask���ó�1�ˣ�
        '�����ڳ������޷��ܺõĿ��ƣ����ֱ��ȥ����0028��6100��������ԡ�
        If Not IsNull(curImage.Attributes(&H28, &H6100).value) Then
            curImage.Attributes.Remove &H28, &H6100
        End If
    End If
End Sub

Public Sub PlayMedia(ByVal lngMediaIndex As Long)
'����ָ����������ý��

End Sub


Private Sub ConfigPageControl(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, ByVal blnTmpRecord As Boolean)
'���÷�ҳ�ؼ�
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRecordCount As Long
    
    
    Select Case slQueryLevel
        Case slStudy
        
            If IsNumeric(strSearchValue) Then
BUGEX "ConfigPageControl:IsNumeric----->"
                strSQL = "select sum(����ֵ) ����ֵ from ( " & " select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.���UID=to_char([1]) " & vbCrLf & _
                                              " union all " & vbCrLf & _
                                              " select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=[1] ) "
            Else
BUGEX "ConfigPageControl:IsNotNumeric----->"
                strSQL = "select count(1)  as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b where a.����UID=b.����UID and b.���UID=[1]"
            End If
            
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
    
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ѯͼ������", strSearchValue)
    
    If rsData.RecordCount > 0 Then
        lngRecordCount = Nvl(rsData!����ֵ)
    Else
        lngRecordCount = 0
    End If

    
'    ucPage.PageRecord = mlngPageRecord
    ucPage.RecordCount = lngRecordCount
    
    Call RefreshPageControl
End Sub



Private Function GetImageViewData(ByVal slQueryLevel As TQueryLevel, ByVal strSearchValue As String, _
    ByVal lngCurPage As Long, ByVal lngPageRecord As Long, ByVal blnTmpRecord As Boolean) As ADODB.Recordset
'��ȡԤ��ͼ������
'intSearchType:0-�����uid����,1-������UID����,2-��ͼ��UID����

    Dim strSQL As String
    Dim lngStartRecord As Long
    Dim lngEndRecord As Long
    
    If gblnUseActivexLoad Then
        strSQL = "Select rownum as ˳���, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1,D.����Ŀ¼ as ����Ŀ¼1,D.����Ŀ¼�û��� as ����Ŀ¼�û���1,D.����Ŀ¼���� as ����Ŀ¼����1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2,E.����Ŀ¼ as ����Ŀ¼2,E.����Ŀ¼�û��� as ����Ŀ¼�û���2,E.����Ŀ¼���� as ����Ŀ¼����2," & _
            "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
    Else
        strSQL = "Select rownum as ˳���, A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1," & _
            "D.IP��ַ As Host1,'/'||D.FtpĿ¼||'/' As Root1," & _
            "Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
            "||C.���UID||'/'||A.ͼ��UID As URL,d.�豸�� as �豸��1, " & _
            "E.FTP�û��� As User2,E.FTP���� As Pwd2," & _
            "E.IP��ַ As Host2,'/'||E.FtpĿ¼||'/' As Root2," & _
            "e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� " & _
            "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
            "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) "
    End If
    
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
        Case slStudy
            If IsNumeric(strSearchValue) Then
                strSQL = "select * from ( " & strSQL & "and c.���UID=to_char([1]) " & vbCrLf & _
                                          " union all " & vbCrLf & _
                                          strSQL & " and c.ҽ��ID=[1] ) "
            Else
                strSQL = "select * from (" & strSQL & " and C.���UID=[1])"
            End If
        Case slSeries
            strSQL = "select * from (" & strSQL & " and B.����UID=[1])"
        Case slImage
            strSQL = "select * from (" & strSQL & " and A.ͼ��UID=[1])"
    End Select
    
    lngStartRecord = (lngCurPage - 1) * lngPageRecord + 1
    lngEndRecord = lngCurPage * lngPageRecord
    
    strSQL = "select * from (" & strSQL & " order by ����UID, ͼ���) where ˳���>=" & lngStartRecord & " and ˳���<=" & lngEndRecord
    
    Set GetImageViewData = zlCL_GetDBObj.OpenSQLRecord(strSQL, "��ѯͼ����Ϣ", strSearchValue)
End Function

Public Sub AddImage(img As Object, Optional objImgTag As Object = Nothing)
'����ͼ��
    Dim i As Long
    
    If dcmMiniImage.Images.Count < ucPage.PageRecord Then
        Call ConfigImgDisplayFormat(dcmMiniImage.Images.Count + 1)
        
        Call dcmMiniImage.Images.Add(img)
    Else
        '�ƶ�ͼ��
        For i = 2 To dcmMiniImage.Images.Count
            Call dcmMiniImage.Images.Move(i, i - 1)
            dcmMiniImage.Images(i - 1).BorderColour = vbWhite
        Next i
        
        Call dcmMiniImage.Images.Remove(dcmMiniImage.Images.Count)
        dcmMiniImage.Images.Add img
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
    
    Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    Call UpdateImageCount(1)
End Sub


Private Sub ShowAVInf(img As DicomImage, objImgTag As clsImageTagInf)
'��ʾ����Ƶ��Ϣ
    If objImgTag.Tag = VIDEOTAG Or objImgTag.Tag = AUDIOTAG Then
        Call AddVideoLabelToDicomImage(img, _
        IIf(objImgTag.Tag = VIDEOTAG, "¼��ʱ�䣺", "¼��ʱ�䣺") & objImgTag.CaptureTime, _
        IIf(objImgTag.Tag = VIDEOTAG, "¼�񳤶ȣ�", "¼�����ȣ�") & objImgTag.RecordTimeLen & " ��", _
        "�������ƣ�" & objImgTag.EncoderName)
    End If
End Sub

Public Sub DeleteImage(ByVal lngImgIndex As Long)
'ɾ��ͼ��
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
        '��ͼ��ҳ���ı�ʱ������ˢ�µ�ǰҳͼ����ʾ
        Call ucPage.MovePage(ucPage.PageNumber)
        If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
    End If
    
'    If dcmMiniImage.Images.Count <= 0 Then
'        Call ucPage.MovePage(ucPage.PageNumber)
'
'        If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
'    End If
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
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_DblClick()
On Error GoTo errHandle
    Dim blnContinue As Boolean
    
    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
    If mlngSelectIndex <= 0 Then Exit Sub

    blnContinue = True
    
    Call DoOnDbClick(mlngSelectIndex, blnContinue)
    
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
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
End Sub


Private Sub dcmMiniImage_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error Resume Next
    Dim i As Long
    Dim objLabs As DicomLabels
    Dim lngImgIndex As Long
    
    lngImgIndex = dcmMiniImage.ImageIndex(X, Y)
    
    Call UpdateSelectIndex(lngImgIndex)
    
    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        
        If mblnEnable Then
        
            '����ѡ���״̬
            Set objLabs = dcmMiniImage.LabelHits(X, Y, False, True, True)
            For i = 1 To objLabs.Count
                If objLabs(i).Tag = M_STR_SELECT_TAG Then
                    objLabs(i).Transparent = Not objLabs(i).Transparent
                    
                    '����ͼ��ѡ�¼�
                    Call DoOnCheckChange(mlngSelectIndex, Not objLabs(i).Transparent)
                    
                    Exit For
                End If
            Next i
            
        End If
    End If
    
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
End Sub


Private Sub dcmMiniImage_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer
    
    'û�зŴ�����ͼ���򲻽���ͼ������
    If mlngMouseMoveZoom = 0 Or mlngBigImageWay <> 1 Or dcmMiniImage.Images.Count <= 0 Then
        RaiseEvent OnMouseMove(Button, Shift, X, Y)
        Exit Sub
    End If
    
    '�ж��Ƿ���Ҫ��ʾͼ��
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImage.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImage.Height) Then
        blnShowImg = True
    End If
    
    If blnShowImg Then      '��ʾͼ��
        SetCapture dcmMiniImage.hWnd    '�������
        
        intCurrImg = dcmMiniImage.ImageIndex(X, Y)
        
        If intCurrImg <> 0 Then
            '����ͼ����ʾ
            frmShowImg.ShowMe dcmMiniImage.Images(intCurrImg), Me, 1, 0, 0, mlngMouseMoveZoom
        Else
            frmShowImg.HideMe
        End If
    Else        '�ر�ͼ����ʾ
        ReleaseCapture      '�������
        frmShowImg.HideMe
    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errHandle:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
End Sub

Private Sub dcmMiniImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim curPointer As PointAPI
    Dim i As Integer
    
    If Button = 2 And mblnIsShowPopup Then
        '��ʾ�Ҽ��˵�
        Call GetCursorPos(curPointer)
        
        Call ScreenToClient(hWnd, curPointer)  'ScreenToClient����ʹ�õĵ�λΪ����ֵ
        Call PopupMenu(menuPopup, 0, ScaleX(curPointer.X, vbPixels, vbTwips), ScaleY(curPointer.Y, vbPixels, vbTwips))
        
    Else
        '��ʾ��ͼ
        If mlngMouseMoveZoom <> 0 And mlngBigImageWay = 2 Then
            
            If dcmMiniImage.Images.Count > 0 Then

                i = dcmMiniImage.ImageIndex(X, Y)
                If i = 0 Then i = 1

                frmShowImg.ShowMe dcmMiniImage.Images(i), Me, 2, 0, 0
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
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
End Sub


Private Sub mnuSplitPageTool_Click()
'��ʾ��ҳ������
On Error GoTo errHandle
    mblnShowPageControl = True
    ucPage.Visible = mblnShowPageControl
    
    Call UserControl_Resize
errHandle:
End Sub

Private Sub ucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    
    Set rsData = GetImageViewData(mslQueryLevel, mstrQueryValue, lngPageIndex, lngPageCount, mblnQueryTmpRecord)
    Call LoadViewImageToFace(rsData, dcmMiniImage)
    
    
    If mblnIsShowCheckbox Then
        Call DrawImageSelectBorder(dcmMiniImage)
    End If
Exit Sub
errHandle:
    Call MsgboxCus(err.Description, vbOKOnly, G_STR_HINT_TITLE)
End Sub

Private Sub UserControl_Initialize()
    
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True
    
    mblnIsShowCheckbox = False
    mblnIsShowPopup = False
    mblnShowPageControl = False
    
    mlngMouseMoveZoom = 0
    mlngBigImageWay = 0
    
    mstrQueryValue = ""
    mlngSelectIndex = 0
    
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
On Error Resume Next
    
    dcmMiniImage.Left = 0
    dcmMiniImage.Top = 0
    dcmMiniImage.Width = UserControl.ScaleWidth
    dcmMiniImage.Height = UserControl.ScaleHeight - IIf(mblnShowPageControl, ucPage.Height + 60, 0)
    
    ucPage.Left = 0
    ucPage.Top = dcmMiniImage.Height + 30
    
'    Refresh

    err.Clear
End Sub



Private Function GetImageRow(ByVal lngImageIndex As Long) As Integer
'ȡ�õ�ǰ������
    GetImageRow = CInt(lngImageIndex / dcmMiniImage.MultiColumns) + 1
End Function

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub



Private Sub DrawImageSelectBorder(dcmViewer As DicomViewer)
'����ͼ��ѡ��߿�
    Dim i As Long
    
    Dim lSelect As DicomLabel
    Dim lBorder As DicomLabel

    
    'ѭ��ÿһ��ͼ�񣬻���ע
    For i = 1 To dcmViewer.Images.Count
        Call dcmViewer.Images(i).Labels.Clear
        
        Set lBorder = New DicomLabel

        lBorder.LabelType = 2            '�߿�
        lBorder.Width = 1000
        lBorder.Height = 1000
        lBorder.Left = 0
        lBorder.Top = 0
        lBorder.LineWidth = 2


        lBorder.ForeColour = vbYellow
        lBorder.BackColour = vbYellow


        lBorder.Transparent = True
        lBorder.ScaleWithCell = True
        lBorder.Tag = M_STR_BORDER_TAG

        lBorder.Visible = True
        dcmViewer.Images(i).Labels.Add lBorder
        

    
    
        Set lSelect = New DicomLabel
        
        lSelect.LabelType = 2            '����
        lSelect.Width = 18
        lSelect.Height = 18
        lSelect.Left = 1
        lSelect.Top = 1
        lSelect.LineWidth = 2
        
        lSelect.ForeColour = vbYellow
        lSelect.BackColour = vbRed
        
                
        lSelect.Transparent = True
        lSelect.ScaleWithCell = False
        lSelect.ImageTied = False

        lSelect.Tag = M_STR_SELECT_TAG
        
        lSelect.Visible = True
        dcmViewer.Images(i).Labels.Add lSelect
        
        dcmViewer.Images(1).BorderStyle = vbRed
    Next i
End Sub





Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    mblnIsShowCheckbox = PropBag.ReadProperty("ShowCheckbox", False)
    mblnIsShowPopup = PropBag.ReadProperty("ShowPopup", False)
    mlngMouseMoveZoom = PropBag.ReadProperty("MouseMoveZoom", 0)
    ucPage.PageRecord = PropBag.ReadProperty("PageImgCount", 5)
    AutoRedraw = PropBag.ReadProperty("AutoRedrawStyle", False)
    
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
    Call PropBag.WriteProperty("MouseMoveZoom", mlngMouseMoveZoom, 0)
    Call PropBag.WriteProperty("PageImgCount", ucPage.PageRecord, 5)
    Call PropBag.WriteProperty("AutoRedrawStyle", AutoRedraw, False)
    
    err.Clear
End Sub
