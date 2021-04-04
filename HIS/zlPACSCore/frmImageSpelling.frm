VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmImageSpelling 
   BackColor       =   &H8000000B&
   Caption         =   "ͼ��ƴ��"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   10005
   Icon            =   "frmImageSpelling.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicViewer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5292
      Left            =   600
      ScaleHeight     =   5265
      ScaleWidth      =   8955
      TabIndex        =   1
      Top             =   960
      Width           =   8988
      Begin DicomObjects.DicomViewer viewer 
         Height          =   2505
         Index           =   0
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   2790
         _Version        =   262147
         _ExtentX        =   4932
         _ExtentY        =   4424
         _StockProps     =   35
      End
   End
   Begin VB.ListBox lstSort 
      Height          =   240
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7350
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2302
            MinWidth        =   2293
            Text            =   "�������"
            TextSave        =   "�������"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10186
            Text            =   "˫����Ƭ����ͼ��ֱ�Ӱ�ͼ�����ƴ����"
            TextSave        =   "˫����Ƭ����ͼ��ֱ�Ӱ�ͼ�����ƴ����"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "���ű�����"
            TextSave        =   "���ű�����"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "��д"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   926
            MinWidth        =   706
            Text            =   "����"
            TextSave        =   "NUM"
            Object.ToolTipText     =   "����"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImgIcons 
      Left            =   3000
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmImageSpelling.frx":0CCA
   End
   Begin XtremeCommandBars.CommandBars CommBar_ImageSelling 
      Left            =   120
      Top             =   360
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmImageSpelling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intImageCount As Integer
Private intBaseX As Long                     '''��¼���ԭ����Xλ��
Private intBaseY As Long                     '''��¼���ԭ����Yλ��
Dim intSelectedViewer As Integer
Dim iMaxTag As Long
Dim iMaxViewer As Integer
Private mintMouseState As Integer        '''��¼����״̬��0���ޣ����Σ���2������;3-�ü�
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mblnMouseDown As Boolean        '����Ƿ񱻰���
Private SelectedImg As DicomImage               '''��¼��ǰ��ѡ�е�ͼ��
Private mRViewerWidth As Long                   '''��¼ƴ����ɺ�ͼ��Ŀ��
Private mRViewerHeight As Long                  '''��¼ƴ����ɺ�ͼ��ĸ߶�
Private mstrSeriesUID As String                 '''��¼ƴ����ɵ�ͼ����ʹ�õ�ԭͼ������UID������ͼ�񱣴�

Public f As frmViewer

''''''''''''''''�ü�''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutViewer                '�ü������ڵ�viewer���
Private mintCutOutImage                 '�ü������ڵ�ͼ�����
Private mintCutOutLabel                 '�ü������ڵı�ע���
Private mblnLabelMoving As Boolean      '�����ƶ��ü���
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function funCompleteSplling() As DicomImage
'------------------------------------------------
'���ܣ� �Ե�ǰ�ںõ�ͼ�����ƴ�ӣ����ɽ��ͼ��
'������ ��
'���أ� ƴ�ӵĽ��ͼ��
'˵���� ��ͼ��ֱ��� Monochrome1��Monochrome2ʱ��Window=False,���ĸ�Ϊ������һ���ͻᱻ��ɫ��
'       ��ͼ��ֱ��� Monochrome1��Monochrome2ʱ��Window=true������ͼ���ᱻ��ɫ���������ͼ����Signed�ģ���������ΪĿ��ͼ��
'       ��ͼ���а�����Monochrome1��Monochrome2���������RGBʱ��Window=False,����ʾ������Ҫ����Window=True��
'       ��ͼ���а�����Monochrome1��Monochrome2���������RGBʱ��Window=True���Ժڰ�ͼΪ�����õ��ڰ�ͼ���Բ�ɫͼΪ���õ���ɫͼ��
'------------------------------------------------
    If intImageCount < 1 Then
        Exit Function
    End If
    Dim i As Integer
    Dim dblLeft As Double, dblTop As Double
    Dim dblWidth As Double, dblHeight As Double
    Dim dblRight As Double, dblBottom As Double
    Dim NewImg As New DicomImage
    Dim sizex As Long   '��һ��ͼ���x��������
    Dim sizey As Long   '��һ��ͼ���y��������
    Dim iViewerIndex As Integer     '��ʾ��ǰʹ�õ����ĸ�viewer
    Dim ClipSizex As Integer    '�����е�ͼ���x������������
    Dim ClipSizey As Integer    '�����е�ͼ���y������������
    Dim view As DicomViewer
    Dim dblMaxZoom As Double
    Dim dcmGlobal As New DicomGlobal
    Dim blnWindow As Boolean    'ƴ�ӵ�ʱ��Window������ֵ
    Dim intMainImage As Integer     'ƴ�ӵ�ʱ��ʹ���ĸ�ͼ����Ϊ��ͼ������ʹ�ò�ɫͼ��Ϊ��ͼ��
    Dim MainImage As New DicomImage     'ƴ��ʱ�õ���ͼ��

    On Error GoTo err
    
    blnWindow = False       'Ĭ��ΪFalse
    intMainImage = 0        'Ĭ��Ϊ0
    
    '��ʼ����ͼ���λ�úʹ�С
    dblLeft = Viewer(intSelectedViewer).left
    dblTop = Viewer(intSelectedViewer).top
    dblRight = Viewer(intSelectedViewer).left + Viewer(intSelectedViewer).width
    dblBottom = Viewer(intSelectedViewer).top + Viewer(intSelectedViewer).height
    dblMaxZoom = Viewer(intSelectedViewer).Images(1).ActualZoom

    '������е�lstSort
    Me.lstSort.Clear
    
    '��ԭͼ��������򣬰���tag�Ĵ�С�����򣬴�С����.
    '��ȡ��ͼ���λ��:�󣬶����ң���,ȡ��ǰ����ͼ��
    'ѭ�����е�ͼ��
    For Each view In Viewer
        If view.Index <> 0 Then
            If view.left < dblLeft Then
                dblLeft = view.left
            End If
            If view.top < dblTop Then
                dblTop = view.top
            End If
            If view.left + view.width > dblRight Then
                dblRight = view.left + view.width
            End If
            If view.top + view.height > dblBottom Then
                dblBottom = view.top + view.height
            End If
            If dblMaxZoom < view.Images(1).ActualZoom Then
                dblMaxZoom = view.Images(1).ActualZoom
            End If
            
            '��ͼ���TAG��viewer ��index�ŵ�lstSort�У������������
            Me.lstSort.AddItem Format(view.Images(1).Tag, "0000")
            Me.lstSort.ItemData(Me.lstSort.NewIndex) = view.Index
            
            '��¼ͼ���(0028,0004) Photometric Interpretation
            If intMainImage = 0 And view.Images(1).Attributes(&H28, &H4).Exists And Not IsNull(view.Images(1).Attributes(&H28, &H4).Value) Then
                If UCase(view.Images(1).Attributes(&H28, &H4).Value) = "MONOCHROME1" Or UCase(view.Images(1).Attributes(&H28, &H4).Value) = "MONOCHROME2" Then
                    '��������
                Else
                    blnWindow = True
                    intMainImage = view.Index
                End If
            End If
        End If
    Next
    
    If intMainImage = 0 Then intMainImage = intSelectedViewer
    
    '������Viewer�Ŀ�Ⱥ͸߶�
    dblWidth = dblRight - dblLeft
    dblHeight = dblBottom - dblTop
    mRViewerWidth = dblWidth
    mRViewerHeight = dblHeight
    
    '����ͼ���λ�ô��ת��Ϊ����
    dblLeft = dblLeft / Screen.TwipsPerPixelX
    dblWidth = dblWidth / Screen.TwipsPerPixelX
    dblTop = dblTop / Screen.TwipsPerPixelY
    dblHeight = dblHeight / Screen.TwipsPerPixelY
    
    '������ڲ�ɫͼ���Ƚ���ͼ�񱣴��BMP��ʽ��,����JPGͼ��ƴ�Ӻ�����ͼ��
    If blnWindow = True Then
        Viewer(intMainImage).Images(1).FileExport "tmpBMPFile", "BMP"
        MainImage.FileImport "tmpBMPFile", "BMP"
        MainImage.StudyUID = Viewer(intMainImage).Images(1).StudyUID
        MainImage.SeriesUID = Viewer(intMainImage).Images(1).SeriesUID
        MainImage.PatientID = Viewer(intMainImage).Images(1).PatientID
        MainImage.Name = Viewer(intMainImage).Images(1).Name
    Else
        Set MainImage = Viewer(intMainImage).Images(1)
    End If

    '������ͼ��,ʹ��sizex��sizey��Ϊ���ڴ�����ͼ��ʱ��������ԭ����ͼ��
    sizex = MainImage.sizex
    sizey = MainImage.sizey
    
    Set NewImg = MainImage.SubImage(sizex, sizey, dblWidth / dblMaxZoom, dblHeight / dblMaxZoom, 1, MainImage.Frame)

    'ɾ��ͼ����ԭ�е��ڸ�Shutter��Ϣ
    NewImg.Attributes.Remove &H18, &H1600
    NewImg.Attributes.Remove &H18, &H1602
    NewImg.Attributes.Remove &H18, &H1604
    NewImg.Attributes.Remove &H18, &H1606
    NewImg.Attributes.Remove &H18, &H1608
    NewImg.Attributes.Remove &H18, &H1610
    NewImg.Attributes.Remove &H18, &H1612
    NewImg.Attributes.Remove &H18, &H1620
    NewImg.Attributes.Remove &H18, &H1622
    
    '��ԭ��ͼ��һ�������Ƶ���ͼ���С�
    For i = 0 To Me.lstSort.ListCount - 1
        iViewerIndex = Me.lstSort.ItemData(i)
        ClipSizex = (Viewer(iViewerIndex).width / Screen.TwipsPerPixelX) / dblMaxZoom
        ClipSizey = (Viewer(iViewerIndex).height / Screen.TwipsPerPixelY) / dblMaxZoom
                        
        NewImg.Blt Viewer(iViewerIndex).Images(1), Viewer(iViewerIndex).Images(1).ActualScrollX, _
                Viewer(iViewerIndex).Images(1).ActualScrollY, (Viewer(iViewerIndex).left / Screen.TwipsPerPixelX - dblLeft) / dblMaxZoom, _
                (Viewer(iViewerIndex).top / Screen.TwipsPerPixelY - dblTop) / dblMaxZoom, ClipSizex, ClipSizey, Viewer(iViewerIndex).Images(1).Frame, 1, Viewer(iViewerIndex).Images(1).ActualZoom / dblMaxZoom, blnWindow
    Next
    
    NewImg.width = MainImage.width
    NewImg.Level = MainImage.Level
    
    '������޸�����UID
    mstrSeriesUID = NewImg.SeriesUID
    NewImg.SeriesUID = dcmGlobal.NewUID

    Set funCompleteSplling = NewImg
    
    'ɾ������ƴ�ӵ�ͼ��
    If intImageCount >= 1 Then
        For Each view In Viewer
            If view.Index <> 0 Then Unload view
        Next
    End If
    intImageCount = 0
    intSelectedViewer = 0
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CommBar_ImageSelling_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        '���ƴ��
        Case ID_frmImageSpelling_CompleteSpelling
            subLoadImage funCompleteSplling
            Viewer(iMaxViewer).width = mRViewerWidth
            Viewer(iMaxViewer).height = mRViewerHeight
            Call subDrawFrame(Viewer(iMaxViewer), False, True)
        '����ͼ���˳�
        Case ID_frmImageSpelling_SavePhoto
            If intSelectedViewer = 0 Then Exit Sub
            
            '��ͼ�񱣴浽������
            If subSaveImage(Viewer(intSelectedViewer).Images(1), mstrSeriesUID) = True Then
                '�򿪲���ʾ���ͼ��
                Call subOpenCurrentImage(f, Viewer(intSelectedViewer).Images(1))
            End If
            '�˳�
            Unload Me
        'ɾ��ͼ��
        Case ID_frmImageSpelling_DelPhoto
            If intImageCount < 1 Then Exit Sub
            Unload Viewer(intSelectedViewer)
            intImageCount = intImageCount - 1
            intSelectedViewer = Viewer.UBound
        '�ƶ�
        Case ID_frmImageSpelling_Move
            subSetToolBarChecked control.Id
            mintMouseState = 0
        '����
        Case ID_frmImageSpelling_ZoomOut
            subSetToolBarChecked control.Id
            mintMouseState = 2
        '�ü�
        Case ID_frmImageSpelling_CutOut
            subSetToolBarChecked control.Id
            
            '�������״̬
            mintMouseState = 3
            
        '�˳�
        Case ID_frmImageSpelling_Quit
            Unload Me
    End Select
End Sub

Private Sub CommBar_ImageSelling_GetClientBordersWidth(left As Long, top As Long, Right As Long, Bottom As Long)
    If sbStatusBar.Visible Then
        Bottom = sbStatusBar.height
    End If
End Sub

Private Sub CommBar_ImageSelling_Resize()
    On Error Resume Next
    
    Dim left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.CommBar_ImageSelling.GetClientRect left, top, Right, Bottom
    If Right >= left And Bottom >= top Then
        picViewer.Move left, top, Right - left, Bottom - top
    Else
        picViewer.Move 0, 0, 0, 0
    End If
    
End Sub

Private Sub CommBar_ImageSelling_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.Id
        Case ID_frmImageSpelling_CompleteSpelling, ID_frmImageSpelling_SavePhoto, ID_frmImageSpelling_DelPhoto, _
             ID_frmImageSpelling_Move, ID_frmImageSpelling_ZoomOut, ID_frmImageSpelling_CutOut
            If Viewer.Count <= 1 Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
    End Select
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    '����������
    Call CreateBar
    '����״̬��ͼ��
    'Set sbStatusBar.Panels(1).Picture = f.ImgList32.ListImages(4).Picture
    intImageCount = 0
    iMaxViewer = 0
    iMaxTag = 0
End Sub

Public Sub subLoadImage(im As DicomImage)
'------------------------------------------------
'���ܣ� װ��ͼ�񣬰�ͼ����ؽ���ƴ�Ӵ���
'������ im---��Ҫ���ص�ͼ��
'���أ� ��
'------------------------------------------------
    
    If im Is Nothing Then Exit Sub
    
    On Error GoTo err
    
    If im.Attributes(&H28, &H4) = "PALETTE COLOR" Then
        MsgBox "��ɫͼ���ܽ���ͼ��ƴ�ӡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    intImageCount = intImageCount + 1
    iMaxViewer = iMaxViewer + 1
    load Viewer(iMaxViewer)
    Viewer(iMaxViewer).Visible = True
    Viewer(iMaxViewer).MultiColumns = 1
    Viewer(iMaxViewer).MultiRows = 1
    Viewer(iMaxViewer).UseScrollBars = False
    Viewer(iMaxViewer).Images.Add im
    Viewer(iMaxViewer).BackColour = vbBlack
    Viewer(iMaxViewer).width = im.sizex * im.ActualZoom * Screen.TwipsPerPixelX
    Viewer(iMaxViewer).height = im.sizey * im.ActualZoom * Screen.TwipsPerPixelY
    subDrawFrame Viewer(iMaxViewer), True, True
    
    'ɾ����ע
    Viewer(iMaxViewer).Images(1).Labels.Clear
    
    Viewer(iMaxViewer).Images(1).StretchToFit = True
    
    iMaxTag = iMaxTag + 1
    Viewer(iMaxViewer).Images(1).Tag = iMaxTag
    If intImageCount = 1 Then
        intSelectedViewer = iMaxViewer
        Viewer(iMaxViewer).Labels(1).ForeColour = vbRed
    End If
    Viewer(iMaxViewer).Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub subDrawFrame(v As DicomViewer, isNew As Boolean, isSelected As Boolean)
'------------------------------------------------
'���ܣ� ��ͼ���ѡ��߿�
'������ v ---ͼ�����ڵ�Viwer
'       isNew ---�Ƿ������ӵ�ͼ��,�����ӵ�ͼ����Ҫ�����ӱ߿�label
'       isSelected --- ͼ���Ƿ�ѡ�񣬱�ѡ��ʱ�߿���ɫ��ͬ
'���أ� ��
'------------------------------------------------
    Dim l As New DicomLabel
    
    On Error GoTo err
    
    If Not isNew Then Set l = v.Labels(1)
    
    l.LabelType = doLabelRectangle
    l.left = 0
    l.top = 0
    l.width = v.width / Screen.TwipsPerPixelX
    l.height = v.height / Screen.TwipsPerPixelY
    l.ForeColour = IIf(isSelected, vbRed, vbWhite)
    If isNew Then
        l.ImageTied = True
        v.Labels.Add l
    End If
    v.Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f.blnfis = False
End Sub

Private Sub picViewer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub viewer_DblClick(Index As Integer)
    Call sub�ü�
End Sub

Private Sub Viewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim i As Integer
    Dim v As DicomViewer
    Dim intImgIndex As Integer
    Dim ls As DicomLabels
    
    intBaseX = x
    intBaseY = y
    
    intImgIndex = Viewer(Index).ImageIndex(x, y)
    If Viewer(Index).Images.Count > 0 And intImgIndex <> 0 Then
        Set SelectedImg = Viewer(Index).Images(intImgIndex)
        If Button = 1 Then
            If mintMouseState = 0 Then      '�ƶ�ͼ��
                mblnMouseDown = True
            ElseIf mintMouseState = 1 Then       '����
                mblnMouseDown = True
            ElseIf mintMouseState = 2 Then       '����
                mblnMouseDown = True
            ElseIf mintMouseState = 3 Then       '�ü�
                '�ü�״̬�µ����down�������ֲ�����1�����ü��򣨼�¼��ǣ���2���ƶ��ü���(�н���) ��3��˫�����вü�
                If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then  '���ü���
                    '���ӿ�ѡ��ע
                    Viewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), 0, 0)
                    Set mdcmSelectLabel = Viewer(Index).Images(intImgIndex).Labels(Viewer(Index).Images(intImgIndex).Labels.Count)
                    mdcmSelectLabel.Tag = CUT_LABEL
                    mblnMouseDown = True
                    mintCutOutViewer = Index
                    mintCutOutImage = intImgIndex
                    mintCutOutLabel = Viewer(Index).Images(intImgIndex).Labels.Count
                    Viewer(Index).Refresh
                Else            '��ʼ�ƶ��ü���
                    Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
                    If ls.Count <> 0 And Screen.MousePointer <> vbArrow Then
                        '��ʼ�ƶ��ü���
                        If ls(1).Tag = CUT_LABEL And SelectedImg.Labels(SelectedImg.Labels.Count).Tag = CUT_LABEL Then
                            mblnLabelMoving = True
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Viewer(Index).ZOrder
    iMaxTag = iMaxTag + 1
    Viewer(Index).Images(1).Tag = iMaxTag
    intSelectedViewer = Index
    For Each v In Viewer
        If v.Index <> 0 Then
            subDrawFrame v, False, IIf((v.Index = Index), True, False)
        End If
    Next
End Sub

Private Sub Viewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    Dim i As Integer
    Dim v As DicomViewer
    Dim dblZoom As Double
    Dim dblZoomRatio As Double
    
    If SelectedImg Is Nothing Then Exit Sub
    
    If Button = 1 Then
        Select Case mintMouseState
            Case 0                      '�ƶ�ͼ��
                If mblnMouseDown = True Then
                    If Viewer(Index).left + (x - intBaseX) * Screen.TwipsPerPixelX > 24 And x <> intBaseX Then
                        Viewer(Index).left = Viewer(Index).left + (x - intBaseX) * Screen.TwipsPerPixelX
                    End If
                    If Viewer(Index).top + (y - intBaseY) * Screen.TwipsPerPixelX > 24 And y <> intBaseY Then
                        Viewer(Index).top = Viewer(Index).top + (y - intBaseY) * Screen.TwipsPerPixelX
                    End If
                End If
            Case 2                  '����
                If mblnMouseDown = True Then
                    '���ŵ�λ��0.01��
                    dblZoom = SelectedImg.ActualZoom * (1 + (intBaseY - y) * 0.001)
                    If dblZoom < 0.01 Then dblZoom = 0.01
                    If dblZoom > 64 Then dblZoom = 64
                    dblZoomRatio = dblZoom / SelectedImg.ActualZoom
                    Viewer(Index).width = Viewer(Index).width * dblZoomRatio
                    Viewer(Index).height = Viewer(Index).height * dblZoomRatio

                    Call subDrawFrame(Viewer(Index), False, True)

                    intBaseX = x
                    intBaseY = y
                End If
            Case 3                  '�ü�
                If mblnMouseDown = True Then
                    mdcmSelectLabel.width = Viewer(Index).ImageXPosition(x, y) - mdcmSelectLabel.left
                    mdcmSelectLabel.height = Viewer(Index).ImageYPosition(x, y) - mdcmSelectLabel.top
                    Viewer(Index).Refresh
                End If
        End Select
    End If
    
    '��������ü�
    If mintMouseState = 3 And mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        If Button = 1 Then          '��걻����
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(Viewer(Index), SelectedImg, x, y)
                Set lblCUT = SelectedImg.Labels(SelectedImg.Labels.Count)

                If (Screen.MousePointer = vbSizeWE And (SelectedImg.RotateState = doRotateNormal Or SelectedImg.RotateState = doRotate180)) _
                    Or (Screen.MousePointer = vbSizeNS And (SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight)) Then       '�����ƶ�

                    lngXOffset = (Viewer(Index).ImageXPosition(x, y) - Viewer(Index).ImageXPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.left - Viewer(Index).ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - Viewer(Index).ImageXPosition(x, y)) Then '�ұߵ��ƶ�
                            lblCUT.width = lblCUT.width + lngXOffset
                    Else    '������ƶ�
                            lblCUT.left = lblCUT.left + lngXOffset
                            lblCUT.width = lblCUT.width - lngXOffset
                    End If
                ElseIf (Screen.MousePointer = vbSizeNS And (SelectedImg.RotateState = doRotateNormal Or SelectedImg.RotateState = doRotate180)) _
                    Or (Screen.MousePointer = vbSizeWE And (SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight)) Then   '�����ƶ�

                    lngYOffset = (Viewer(Index).ImageYPosition(x, y) - Viewer(Index).ImageYPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.top - Viewer(Index).ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - Viewer(Index).ImageYPosition(x, y)) Then    '�����ߵ��ƶ�
                        lblCUT.height = lblCUT.height + lngYOffset

                    Else    '�������ƶ�
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                ElseIf Screen.MousePointer = vbSizePointer Then     '�����ƶ�

                    lngXOffset = (Viewer(Index).ImageXPosition(x, y) - Viewer(Index).ImageXPosition(intBaseX, intBaseY))
                    lngYOffset = (Viewer(Index).ImageYPosition(x, y) - Viewer(Index).ImageYPosition(intBaseX, intBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                intBaseX = x
                intBaseY = y
                Viewer(Index).Refresh
            End If
        ElseIf Button = 0 Then
            If ls.Count <> 0 Then
                If Abs(ls(1).left - Viewer(Index).ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - Viewer(Index).ImageXPosition(x, y)) < 4 Then
                    If SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight Then
                        Screen.MousePointer = vbSizeNS
                    Else
                        Screen.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - Viewer(Index).ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - Viewer(Index).ImageYPosition(x, y)) < 4 Then
                    If SelectedImg.RotateState = doRotateLeft Or SelectedImg.RotateState = doRotateRight Then
                        Screen.MousePointer = vbSizeWE
                    Else
                        Screen.MousePointer = vbSizeNS
                    End If
                Else
                    Screen.MousePointer = vbSizePointer
                End If
            Else
                Screen.MousePointer = vbArrow
            End If
        End If
    End If
End Sub
Private Sub Viewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = 1 Then
        If mintMouseState = 2 Then          '����
            'ȡ�����Ų���
            Call subSetToolBarChecked(ID_frmImageSpelling_Move)
            mintMouseState = 0
        ElseIf mintMouseState = 3 Then
            If mblnMouseDown Then           '�ü�
                '�����κβ���
                '����ü���Ϊ0 ����ȡɾ���ü�������ü��ı��
                If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                    'ɾ����ѡ�õ���ʱ��ע
                    SelectedImg.Labels.Remove SelectedImg.Labels.Count
                    Set mdcmSelectLabel = Nothing
                    Viewer(Index).Refresh
                    
                    mintCutOutViewer = 0
                    mintCutOutImage = 0
                    mintCutOutLabel = 0
                End If
            End If
        End If
    End If
    mblnMouseDown = False
    mblnLabelMoving = False
End Sub

Private Sub CreateBar()
    '------------------------------------------------
    '���ܣ�                                  �����˵�
    '������
    '���أ�                                  ��
    '------------------------------------------------
    Dim ToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.CommBar_ImageSelling.VisualTheme = xtpThemeOffice2003
    Me.CommBar_ImageSelling.Icons = ImgIcons.Icons
    Me.CommBar_ImageSelling.Item(1).Visible = False                                 '���ز˵���
    
    With Me.CommBar_ImageSelling.Options
        .ShowExpandButtonAlways = False     'ȥ����չ��ť
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    '������������
    Set ToolBar = Me.CommBar_ImageSelling.Add("��������", xtpBarBottom)
    ToolBar.Position = xtpBarTop
    ToolBar.ShowTextBelowIcons = True
    With ToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_CompleteSpelling, "ƴ��")
            cbrControl.IconId = 1010: cbrControl.ToolTipText = "����ͼ��ƴ��"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_SavePhoto, "�����˳�")
            cbrControl.IconId = 1009: cbrControl.ToolTipText = "����ƴ����ɵ�ͼ�񣬲��˳�ϵͳ"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_Move, "�ƶ�ͼ��"): cbrControl.BeginGroup = True
            cbrControl.IconId = 1007: cbrControl.ToolTipText = "�ƶ�ͼ��"
            cbrControl.Checked = True                                       'Ĭ��Ϊ�ƶ�ͼ��
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_ZoomOut, "����")
            cbrControl.IconId = 1005: cbrControl.ToolTipText = "ͼ������"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_CutOut, "�ü�")
            cbrControl.IconId = 1006: cbrControl.ToolTipText = "ͼ��ü�"
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_DelPhoto, "ɾ��ͼ��")
            cbrControl.IconId = 1002: cbrControl.ToolTipText = "ɾ��ͼ��": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ID_frmImageSpelling_Quit, "�˳�")
            cbrControl.IconId = 1003: cbrControl.ToolTipText = "ֱ���˳�ϵͳ"
    End With
End Sub

Private Sub sub�ü�()
'------------------------------------------------
'���ܣ� �ü�ͼ�񣬲ü���ǰѡ�е�ͼ��
'������ ��
'���أ� �ޣ�ֱ����ʾ�ü��Ľ��ͼ
'------------------------------------------------
    If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then Exit Sub
    If mintCutOutImage > Viewer(mintCutOutViewer).Images.Count Then Exit Sub
    If mintCutOutLabel <> Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then Exit Sub
    
    Dim Image As DicomImage
    Dim i As Integer
    Dim lblCUT As DicomLabel
    Dim sourceImage As DicomImage
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim lngNewLeft As Long
    Dim lngNewTop As Long
    
    On Error GoTo err
    
    Set sourceImage = Viewer(mintCutOutViewer).Images(mintCutOutImage)
    Set lblCUT = sourceImage.Labels(sourceImage.Labels.Count)
    
    If lblCUT.width < 0 Then
        lngNewLeft = (lblCUT.left + lblCUT.width) * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    Else
        lngNewLeft = lblCUT.left * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    End If
    If lblCUT.height < 0 Then
        lngNewTop = (lblCUT.top + lblCUT.height) * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    Else
        lngNewTop = lblCUT.top * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    End If
    lngNewWidth = Abs(sourceImage.Labels(sourceImage.Labels.Count).width) * sourceImage.ActualZoom * Screen.TwipsPerPixelX
    lngNewHeight = Abs(sourceImage.Labels(sourceImage.Labels.Count).height) * sourceImage.ActualZoom * Screen.TwipsPerPixelY
    
    Set Image = CutOutAImage(sourceImage)
    
    'ɾ����ѡ�õ���ʱ��ע
    sourceImage.Labels.Remove mintCutOutLabel
    Set mdcmSelectLabel = Nothing
    
    '�������ɵ�ͼ����ӵ�Viewer��
    If mintCutOutImage = 1 And Viewer(mintCutOutViewer).Images.Count = 1 Then
        Viewer(mintCutOutViewer).Images.Clear
        Viewer(mintCutOutViewer).Images.Add Image
    Else
        Viewer(mintCutOutViewer).Images.Remove mintCutOutImage
        Viewer(mintCutOutViewer).Images.Add Image
        Viewer(mintCutOutViewer).Images.Move Viewer(mintCutOutViewer).Images.Count, mintCutOutImage
    End If
    
    '����viewer�ĳ��ȺͿ��,����Viewer�ƶ���ԭ�����ü����λ��
    Viewer(mintCutOutViewer).left = Viewer(mintCutOutViewer).left + lngNewLeft
    Viewer(mintCutOutViewer).top = Viewer(mintCutOutViewer).top + lngNewTop
    Viewer(mintCutOutViewer).width = lngNewWidth
    Viewer(mintCutOutViewer).height = lngNewHeight
    

    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    Screen.MousePointer = vbArrow
    
    '��ɲü���ȡ�����ü��Ĺ��ܣ��ָ��˵���
    
    subSetToolBarChecked ID_frmImageSpelling_Move
    mintMouseState = 0
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subSetToolBarChecked(lngControlID As Long)
    On Error GoTo err
    
    '�Ȱ����а�ť���ó�False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_Move, , True).Checked = False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_ZoomOut, , True).Checked = False
    CommBar_ImageSelling.Item(2).FindControl(, ID_frmImageSpelling_CutOut, , True).Checked = False
    '���ö�Ӧ�İ�ťΪTrue
    CommBar_ImageSelling.Item(2).FindControl(, lngControlID, , True).Checked = True
    
    '����ü��Ŀ�
    '���ԭ���Ѿ��вü�������ɾ������ü���
    If mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        If mintCutOutViewer < Viewer.Count Then
            If mintCutOutImage <= Viewer(mintCutOutViewer).Images.Count Then
                If mintCutOutLabel = Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then
                    'ɾ����ѡ�õ���ʱ��ע
                    Viewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Remove mintCutOutLabel
                    Set mdcmSelectLabel = Nothing
                    Viewer(mintCutOutViewer).Refresh
                End If
            End If
        End If
    End If
    
    '��ʼ���ü�����
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

