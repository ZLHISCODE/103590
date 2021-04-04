VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucImageThumbnail 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   ScaleHeight     =   3075
   ScaleWidth      =   7140
   Begin DicomObjects.DicomViewer dcmMiniImage 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _Version        =   262147
      _ExtentX        =   12515
      _ExtentY        =   5318
      _StockProps     =   35
      BackColor       =   4210752
      UseScrollBars   =   0   'False
      UseMouseWheel   =   -1  'True
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
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
            Picture         =   "ucImageThumbnail.ctx":0000
            Key             =   "avi"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImageThumbnail.ctx":15F952
            Key             =   "aviDownLoad"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImageThumbnail.ctx":2BF2A4
            Key             =   "wav"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImageThumbnail.ctx":41EBF6
            Key             =   "wavDownLoad"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucImageThumbnail.ctx":57E548
            Key             =   "fileDisconet"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuPopup 
      Caption         =   "�Ҽ��˵�"
      Begin VB.Menu mnuSplitPageTool 
         Caption         =   "��ҳ����(&P)"
      End
   End
End
Attribute VB_Name = "ucImageThumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"
Private Const M_STR_FAILD_TAG As String = "FAILD"
Private Const M_SRT_TITLE_TAG As String = "TITLE"

Private Const CON_INT_DICOMSELECTWIDTH As Integer = 18  'dicomͼ���ɫѡ�п����Ͻǻ�ɫС���Ⱥ͸߶�

Private mobjFile As New FileSystemObject

Private mblnIsShowCheckbox As Boolean   '�Ƿ���ʾ��ѡ��
Private mblnEnable As Boolean           '�Ƿ�ɽ��б༭

Private mblnShowPageControl As Boolean
'Private mblnIsShowPopup As Boolean      '�Ƿ���ʾ�Ҽ��˵�
Private mMultiCols As Long
Private mMultiRows As Long
Private mtImageType As TSorceType
Private mlngImgCount As Long             'ͼ������
Private mlngPageCount As Long            'ÿҳͼ����
Private mstrImagePath() As String          '��ʾͼ��
Private mlngSelectIndex As Long
Private mblnClickCheckState As Boolean
Private WithEvents mucPage As ucSplitPageNew         '�󶨷�ҳ�ؼ�
Attribute mucPage.VB_VarHelpID = -1

Private Enum TSorceType
    stImagePath = 0
    stImageObj = 1
End Enum

Public Event OnSelChange(ByVal lngOldIndex As Long, ByVal lngNewIndex As Long)
Public Event OnCheckChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)
Public Event OnClick(ByVal lngSelectedIndex As Long)
Public Event OnDbClick(ByVal lngSelectedIndex As Long, ByRef blnContinue As Boolean)
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
Public Event OnMouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
Public Event OnReUpload()
Public Event OmReMove()

Private Sub DoOnSelChange(ByVal lngOldIndex As Long, ByVal lngNewIndex As Long)

On Error Resume Next
    RaiseEvent OnSelChange(lngOldIndex, lngNewIndex)

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


'�Ƿ���ʾͼ��ѡ��
Property Get ShowCheckBox() As Boolean
    ShowCheckBox = mblnIsShowCheckbox
End Property

Property Let ShowCheckBox(value As Boolean)
    mblnIsShowCheckbox = value
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
    ImageTotal = mlngImgCount
End Property

Property Let ImageTotal(value As Long)
    mlngImgCount = value
End Property

''�󶨷�ҳ�ؼ�
'Property Get SplitPage() As Object
'    Set SplitPage = mucPage
'End Property
'
'Property Let SplitPage(value As Object)
'    Set mucPage = value
'
'    If Not mucPage Is Nothing Then
'        mucPage.RecordCount = mlngImgCount
'    End If
'
'End Property

''�Ƿ���ʾ�Ҽ��˵�
'Property Get ShowPopup() As Boolean
'    ShowPopup = mblnIsShowPopup
'End Property
'
'Property Let ShowPopup(value As Boolean)
'    mblnIsShowPopup = value
'End Property

'��ȡ�ؼ����
Property Get Handle() As Long
    Handle = UserControl.hWnd
End Property

''�Զ����ط�ҳ�������ÿҳ����ʾ����С��ָ����ÿҳ��ʾ����ʱ
'Property Get AutoHidePageControl() As Boolean
'    AutoHidePageControl = mblnIsAutoHidePageControl
'End Property
'
'
'Property Let AutoHidePageControl(value As Boolean)
'    mblnIsAutoHidePageControl = value
'End Property

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

Public Sub SplitPage(objPage As Object)
    Set mucPage = objPage
End Sub

Public Sub RedrawSelf()
'�ػ����
    Call dcmMiniImage.Refresh
End Sub

Public Sub RefreshFace(ByVal IsDock As Boolean)
'ˢ�½���ؼ�λ��
    Call UserControl_Resize
End Sub


Public Sub RefreshImage(strPath() As String)
'ˢ��ͼ����ʾ
    Dim blnLoadResult As Boolean
    Dim i As Long
    
BUGEX "RefreshImage 1"
    
    mstrImagePath = strPath
    '���ͼ��
    Call ClearCurrentPageImage
    
    mlngSelectIndex = 0
    
BUGEX "RefreshImage 2"
    
    Call RefreshDisplay(strPath())
    
    Call dcmMiniImage.Refresh
        
BUGEX "RefreshImage End"
End Sub

Private Sub RefreshDisplay(strPath() As String)
    Call LoadViewImage(strPath())

     '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
        
    Call UserControl_Resize

End Sub


Public Sub RefreshLabelTag()
    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
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

Private Function LoadViewImage(strPath() As String) As Boolean
    Dim strTmpFile As String
    Dim curImage As DicomImage
    Dim dcmTag As clsImageTagInf
    Dim arrImage() As String
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim lngPage As Long
    Dim lngCount As Long
    Dim strCachePath As String
    Dim i As Long
    
On Error GoTo ErrorHand
    
    If UBound(strPath) < 0 Then Exit Function
    
    Call ClearCurrentPageImage
    ConfigImgDisplayFormat UBound(strPath)
    
    mlngImgCount = UBound(strPath)
    
    For i = 1 To UBound(strPath)
        arrImage = Split(strPath(i), "|")
        
        If Len(arrImage(0)) > 0 Then
            If arrImage(1) <> VIDEOTAG And arrImage(1) <> AUDIOTAG Then
            '����ͼ����
                strCachePath = GetCacheDir
                If Dir(arrImage(0)) = vbNullString Then
                    If Dir(strCachePath & "\fileDisconet.bmp") = vbNullString Then
                        Call SavePicture(imgList.ListImages("fileDisconet").Picture, strCachePath & "\fileDisconet.bmp")
                    End If
                    
                    Set curImage = dcmMiniImage.Images.AddNew
                    Call curImage.FileImport(strCachePath + "fileDisconet.bmp", "DIB/BMP")
                    curImage.InstanceUID = arrImage(7)
                    
                    Dim imgLoadInfo As New DicomLabel
                    Dim iCols As Long, iRows As Long
                    
                    iCols = dcmMiniImage.MultiColumns
                    iRows = dcmMiniImage.MultiRows
                    
'                    If blnIsSendOk Then
'                        imgLoadInfo.Text = "[" + Nvl(rsCurImageData!�豸��1, Nvl(rsCurImageData!�豸��2)) + "] �ļ�������..."
'                    Else
'                        imgLoadInfo.Text = "[" + Nvl(rsCurImageData!�豸��1, Nvl(rsCurImageData!�豸��2)) + "] �ļ���������ʧ��."
'                    End If
                    imgLoadInfo.Text = "�ļ�������..."
                                        
                    imgLoadInfo.Width = dcmMiniImage.Width
                    imgLoadInfo.Height = 20
                    
                    imgLoadInfo.Left = 0
                    imgLoadInfo.Top = dcmMiniImage.Height / Screen.TwipsPerPixelY / iRows - imgLoadInfo.Height * 2

                    imgLoadInfo.AutoSize = True
                    imgLoadInfo.ShowTextBox = False
                    imgLoadInfo.Font.Size = 12
                    imgLoadInfo.Font.Bold = True
                    imgLoadInfo.ForeColour = vbRed
                    imgLoadInfo.Tag = M_STR_FAILD_TAG
                    
                    Call curImage.Labels.Add(imgLoadInfo)
                Else
                    Set curImage = ReadViewImage(arrImage(0), dcmMiniImage)
                End If
                
                
                If Not curImage Is Nothing Then
                    Set dcmTag = New clsImageTagInf
                    dcmTag.Tag = arrImage(1)
                                       
                    Set curImage.Tag = dcmTag
                    
                    With curImage
                        .BorderStyle = 6
                        .BorderWidth = 1
                        .BorderColour = vbWhite
                    End With
                End If
            Else
                Set curImage = New DicomImage
                    
                If Dir(arrImage(0)) = vbNullString Then
                    If arrImage(1) = VIDEOTAG Then
                        Call SavePicture(imgList.ListImages("avi").Picture, arrImage(0))
                    Else
                        Call SavePicture(imgList.ListImages("wav").Picture, arrImage(0))
                    End If
                End If
                
                Call curImage.FileImport(arrImage(0), "DIB/BMP")
                Set dcmTag = New clsImageTagInf
                
                dcmTag.Tag = arrImage(1)
                dcmTag.EncoderName = arrImage(3)
                dcmTag.CaptureTime = arrImage(4)
                dcmTag.ReportImage = arrImage(2)
                dcmTag.VideoFile = arrImage(5)
                dcmTag.RecordTimeLen = arrImage(6)
                
                Set curImage.Tag = dcmTag
                
                curImage.InstanceUID = arrImage(7)
                curImage.SeriesUID = arrImage(8)
                curImage.StudyUID = arrImage(8)
                
                Call ShowAVInf(curImage, dcmTag)
                
                With curImage
                    .BorderStyle = 6
                    .BorderWidth = 1
                    .BorderColour = vbWhite
                End With
                
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
    Next
    
'    If Mid(strPath, 1, 1) = "|" Then
'        strPath = Mid(strPath, 2)
'    End If
'
'    If Mid(strPath, Len(strPath) - 1, 1) = "|" Then
'        strPath = Mid(strPath, 1, Len(strPath) - 1)
'    End If
'
'    arrPath = Split(strPath, "|")
'
'    Call ClearCurrentPageImage
'
'    mlngImgCount = UBound(arrPath) + 1
'    ConfigImgDisplayFormat mlngImgCount
'
'    If Not mucPage Is Nothing Then
'        mucPage.RecordCount = mlngImgCount
'    End If
'
'    If mucPage Is Nothing Then
'        lngStart = 0
'        lngEnd = UBound(arrPath)
'    Else
'        lngPage = IIf(mucPage.PageNumber <= 0, 1, mucPage.PageNumber)
'        lngStart = (lngPage - 1) * mucPage.PageRecord
'        lngEnd = IIf(lngPage * mucPage.PageRecord - 1 < UBound(arrPath), lngPage * mucPage.PageRecord - 1, UBound(arrPath))
'    End If
'
'    For i = lngStart To lngEnd
'        strTmpFile = Trim(Nvl(arrPath(i)))
'
'        If Len(strTmpFile) > 0 Then
'            Set curImage = ReadViewImage(strTmpFile, dcmMiniImage)
'
'            '����ͼ����
'            Set dcmTag = New clsImageTagInf
'            dcmTag.Tag = imgTag
'            dcmTag.FilePath = strTmpFile
'
'            Set curImage.Tag = dcmTag
'
'            With curImage
'                .BorderStyle = 6
'                .BorderWidth = 1
'                .BorderColour = vbWhite
'            End With
'        End If
'    Next
     
'    UpdateSelectIndex 1
    LoadViewImage = True
    Exit Function
ErrorHand:
    LoadViewImage = False
    BUGEX "LoadViewImage err = " & err.Description
End Function

Public Sub AddImage(Img As Object, Optional objImgTag As Object = Nothing)
'����ͼ��
    Dim i As Long
    
    Call ConfigImgDisplayFormat(dcmMiniImage.Images.Count + 1)
    Call dcmMiniImage.Images.Add(Img)
    
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
    
    mlngImgCount = mlngImgCount + 1
    '����ͼ��ĸ��ֱ�ע
    Call DrawImageLabels(dcmMiniImage)
    
    Call UpdateSelectIndex(dcmMiniImage.Images.Count)
    
    If Not mucPage Is Nothing Then
        mucPage.RecordCount = mucPage.RecordCount + 1
    End If
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
    
    If Not mucPage Is Nothing Then
        lngCurPageCount = mucPage.PageCount
        
        mucPage.RecordCount = mucPage.RecordCount - 1
            
        If lngCurPageCount > mucPage.PageCount Then
            If blMovePage Then
                Call mucPage.MovePage(mucPage.PageNumber)
                If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
            End If
        Else
            If blMovePage And blMustMovePage Then
    
                Call mucPage.MovePage(mucPage.PageNumber)
                If dcmMiniImage.Images.Count > 0 Then Call UpdateSelectIndex(1)
            End If
        End If
        
        For i = 1 To dcmMiniImage.Images.Count
            If i <> mlngSelectIndex Then dcmMiniImage.Images(i).BorderColour = vbWhite
        Next
        
    End If
End Sub

Public Sub UpdateSelectIndex(ByVal lngSelectIndex As Long)
'����ͼ���ѡ������
    Dim blnIsValidIndex As Boolean
    Dim lngOldIndex As Long
    
    blnIsValidIndex = IIf(lngSelectIndex > 0 And lngSelectIndex <= dcmMiniImage.Images.Count, True, False)
    
    If Not blnIsValidIndex Then Exit Sub

    If blnIsValidIndex Then dcmMiniImage.Images(lngSelectIndex).BorderColour = vbRed
    If mlngSelectIndex = lngSelectIndex Then Exit Sub

    If mlngSelectIndex > 0 And mlngSelectIndex <= dcmMiniImage.Images.Count Then
        dcmMiniImage.Images(mlngSelectIndex).BorderColour = vbWhite
    End If
    
    lngOldIndex = mlngSelectIndex
    mlngSelectIndex = lngSelectIndex
    
    If Not mucPage Is Nothing Then
        mucPage.ItemIndex = (mucPage.PageNumber - 1) * mucPage.PageRecord + lngSelectIndex
    End If
    
    'ִ�������ı��¼�
    Call DoOnSelChange(lngOldIndex, mlngSelectIndex)
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
    
'    If dcmMiniImage.Images.Count <= 0 Then Exit Sub
'    If mlngSelectIndex <= 0 Then Exit Sub
'
    blnContinue = True
'
'    If mlngBigImageWay = 1 Then  '�رմ�ͼ��ʾ
'        ReleaseCapture      '�������
'        frmShowImg.HideMe
'    End If
'
    Call DoOnDbClick(mlngSelectIndex, blnContinue)
'
'    ImgChecked(mlngSelectIndex) = mblnClickCheckState
'
'    If Not blnContinue Then Exit Sub
'
'
'    If dcmMiniImage.MultiColumns = 1 And dcmMiniImage.MultiRows = 1 Then
'        dcmMiniImage.MultiColumns = mMultiCols
'        dcmMiniImage.MultiRows = mMultiRows
'        dcmMiniImage.CurrentIndex = 1
'    Else
'        mMultiCols = dcmMiniImage.MultiColumns
'        mMultiRows = dcmMiniImage.MultiRows
'
'        dcmMiniImage.MultiColumns = 1
'        dcmMiniImage.MultiRows = 1
'
'        dcmMiniImage.CurrentIndex = mlngSelectIndex
'    End If
    
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
    
'    'û�зŴ�����ͼ���򲻽���ͼ������
'    If mlngMouseMoveZoom = 0 Or mlngBigImageWay <> 1 Or dcmMiniImage.Images.Count <= 0 Then
'        RaiseEvent OnMouseMove(Button, Shift, X, Y)
'        Exit Sub
'    End If
'
'    '�ж��Ƿ���Ҫ��ʾͼ��
'    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImage.Width) And _
'       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImage.Height) Then
'        blnShowImg = True
'    End If
'
'    If blnShowImg Then      '��ʾͼ��
'        SetCapture dcmMiniImage.hWnd    '�������
'
'        intCurrImg = dcmMiniImage.ImageIndex(X, Y)
'
'        If intCurrImg <> 0 Then
'            '����ͼ����ʾ
'            frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(intCurrImg)), Me, 1, 0, 0, BigImageCtl, mlngMouseMoveZoom
'        Else
'            frmShowImg.HideMe
'        End If
'    Else        '�ر�ͼ����ʾ
'        ReleaseCapture      '�������
'        frmShowImg.HideMe
'    End If
    
    RaiseEvent OnMouseMove(Button, Shift, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

'Private Function GetBigImage(dcmImg As DicomImage) As DicomImage
'
'    Set GetBigImage = dcmImg.SubImage(0, 0, dcmImg.SizeX, dcmImg.SizeY, 1, dcmImg.Frame)
'
'    GetBigImage.Labels.Clear
'    GetBigImage.BorderColour = vbWhite
'End Function

Private Sub dcmMiniImage_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
'    Dim curPointer As POINTAPI
'    Dim i As Integer
'
'    If mlngBigImageWay = 1 Then  '�رմ�ͼ��ʾ
'        ReleaseCapture      '�������
'        frmShowImg.HideMe
'    End If
'
'    If Button = 2 And mblnIsShowPopup Then
'        '��ʾ�Ҽ��˵�
'        Call GetCursorPos(curPointer)
'
'        Call ScreenToClient(hWnd, curPointer)  'ScreenToClient����ʹ�õĵ�λΪ����ֵ
'        Call PopupMenu(menuPopup, 0, ScaleX(curPointer.X, vbPixels, vbTwips), ScaleY(curPointer.Y, vbPixels, vbTwips))
'    Else
'        '��ʾ��ͼ
'        If mlngMouseMoveZoom <> 0 And mlngBigImageWay = 2 Then
'
'            If dcmMiniImage.Images.Count > 0 Then
'
'                i = dcmMiniImage.ImageIndex(X, Y)
'                If i = 0 Then i = 1
'
'                frmShowImg.ShowMe GetBigImage(dcmMiniImage.Images(i)), Me, 2, 0, 0, BigImageCtl
'            End If
'        End If
'    End If
    
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
errHandle:
End Sub

Private Sub dcmMiniImage_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
On Error GoTo errHandle
'    If Delta > 0 Then
'        Call ucPage.LastPage
'    Else
'        Call ucPage.NextPage
'    End If
    
    Call MouseWheel(Delta)
    
    RaiseEvent OnMouseWheel(Shift, Delta, X, Y)
Exit Sub
errHandle:
    Call MsgboxEx(hWnd, err.Description, vbOKOnly, CON_STR_HINT_TITLE)
End Sub

Public Sub MouseWheel(ByVal Delta As Integer)
    If Not mucPage Is Nothing Then
        If Delta > 0 Then
            mucPage.MoveItem (mucPage.ItemIndex - 1)
        Else
            mucPage.MoveItem (mucPage.ItemIndex + 1)
        End If
    End If
End Sub

'���¼���ͼ��
Public Sub ReLoadFailedImage()

    Call RefreshDisplay(mstrImagePath)
End Sub



'Private Sub mucPage_OnItemChange(ByVal lngPageIndex As Long, ByVal lngPageRecord As Long)
'    Call UpdateSelectIndex(lngPageIndex)
'End Sub
'
'Private Sub mucPage_OnPageChange(ByVal lngPageIndex As Long, ByVal lngPageCount As Long)
'    Call RefreshDisplay(mstrImagePath)
'End Sub
'
'Private Sub mucPage_OnPageRecordChange(ByVal lngPageRecord As Long)
'    Call RefreshDisplay(mstrImagePath)
'End Sub

Private Sub UserControl_Initialize()
    
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True

    mblnIsShowCheckbox = False
    mblnShowPageControl = False
    mlngSelectIndex = 0
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
    dcmMiniImage.Height = UserControl.ScaleHeight
    
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
                If dcmMiniImage.Images(i).Labels(j).Tag = M_SRT_TITLE_TAG Then dcmMiniImage.Images(i).Labels(j).Left = 1
            Next
        Next
    Else
        '��ʾѡ�п�
        For i = 1 To dcmMiniImage.Images.Count
            For j = 1 To dcmMiniImage.Images(i).Labels.Count
                If dcmMiniImage.Images(i).Labels(j).Tag = M_STR_SELECT_TAG Then dcmMiniImage.Images(i).Labels(j).Visible = True
                If dcmMiniImage.Images(i).Labels(j).Tag = M_SRT_TITLE_TAG Then dcmMiniImage.Images(i).Labels(j).Left = CON_INT_DICOMSELECTWIDTH + 4
            Next
        Next
    End If
    
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

Public Sub DrawTitleBorder(dcmImg As DicomImage)
'    Dim iLeft As Integer
'    Dim iWidth As Integer
'    Dim iTop As Integer
'    Dim iHeight As Integer
'    Dim imgResult As New DicomImage
'    Dim iPlane As Integer
'    Dim lngWhiteX As Long
'    Dim lngWhiteY As Long
'    Dim dlMemoText As New DicomLabel
    
'    iLeft = IIf(mblnIsShowCheckbox, CON_INT_DICOMSELECTWIDTH + 4, 0)
'    iTop = 0
'    iWidth = dcmImg.SizeX
'    iHeight = dcmImg.SizeY + 20

    'ʹ��PrinterImage���������Խ�ͼ���ϵı�ǩ����עͬʱ���л���
'    Set imgResult = dcmImg.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight - 20)
'
'    '��ӱ�ע
'
'    dlMemoText.LabelType = doLabelText
'    dlMemoText.ImageTied = True
'    dlMemoText.Transparent = False
'    dlMemoText.AutoSize = False
'    dlMemoText.Left = 0
'    dlMemoText.Top = dcmImg.SizeY
'    dlMemoText.Width = iWidth
'    dlMemoText.Height = 20
'
'    dlMemoText.BackColour = vbWhite
'    dlMemoText.ForeColour = vbBlack
            
'    dlMemoText.Font.Name = Me.Font.Name
'    dlMemoText.Font.Italic = Me.Font.Italic
'    dlMemoText.Font.Strikethrough = Me.Font.Strikethrough
'    dlMemoText.Font.Underline = Me.Font.Underline
'    dlMemoText.Font.Size = Me.Font.Size
'    dlMemoText.Font.Bold = Me.Font.Bold
'    dlMemoText.FontName = Me.Font.Name
'    dlMemoText.FontSize = Me.Font.Size
'    dlMemoText.ShowTextBox = True
'
'    dlMemoText.Text = "1235465"
'
'    imgResult.Labels.Add dlMemoText
'
'    Set imgResult = imgResult.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight)

'    '����ͼ��
'    Me.DViewer.Images.Clear
'    Me.DViewer.Images.Add imgResult
    
'    Dim lTitle As DicomLabel
'
'    If Len(dcmImg.Tag.ImageTitle) > 0 Then
'        Set lTitle = New DicomLabel
'
'        With lTitle
'            .LabelType = 0           '����
'            .Width = IIf(mblnIsShowCheckbox, dcmImg.Width - CON_INT_DICOMSELECTWIDTH - 4, dcmImg.Width)
'            .Height = CON_INT_DICOMSELECTWIDTH
'            .Left = IIf(mblnIsShowCheckbox, CON_INT_DICOMSELECTWIDTH + 4, 1)
'            .Top = 1
'
'            .ForeColour = vbYellow
'            .BackColour = vbRed
'            .Text = dcmImg.Tag.ImageTitle
'            .Text = "����1(1/5)"
'            .Transparent = True
'            .ScaleWithCell = False
'            .ImageTied = False
'
'            .Visible = True
'            .Tag = M_SRT_TITLE_TAG
'        End With
'        dcmImg.Labels.Add lTitle
'        dcmImg.Refresh False
'    End If
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
        
        Call DrawTitleBorder(dcmViewer.Images(i))
    Next i
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    mblnIsShowCheckbox = PropBag.ReadProperty("ShowCheckbox", False)
'    mblnIsShowPopup = PropBag.ReadProperty("ShowPopup", False)
    err.Clear
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("CellSpacing", dcmMiniImage.CellSpacing, 3)
    Call PropBag.WriteProperty("BackColor", dcmMiniImage.BackColour, vbBlack)
    Call PropBag.WriteProperty("Enable", mblnEnable, True)
    Call PropBag.WriteProperty("ShowCheckbox", mblnIsShowCheckbox, False)
'    Call PropBag.WriteProperty("ShowPopup", mblnIsShowPopup, False)
    err.Clear
End Sub





