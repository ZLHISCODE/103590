VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#58.0#0"; "zl9PacsControl.ocx"
Begin VB.Form frmReportImage 
   BorderStyle     =   0  'None
   Caption         =   "����ͼ��"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMiniViewer 
      Height          =   1845
      Left            =   4440
      ScaleHeight     =   1785
      ScaleWidth      =   3615
      TabIndex        =   16
      Top             =   4185
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   1695
         Left            =   45
         TabIndex        =   17
         Top             =   45
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   2990
         BackColor       =   4210752
      End
   End
   Begin VB.PictureBox picMenu 
      Height          =   540
      Left            =   2100
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   585
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin MSComctlLib.ImageList listCur 
      Bindings        =   "frmReportImage.frx":0000
      Left            =   1200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReportImage.frx":0014
            Key             =   "Pen"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMiniImageC 
      Height          =   1935
      Left            =   285
      ScaleHeight     =   1875
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   4170
      Width           =   3735
      Begin VB.VScrollBar vscrollMini 
         Height          =   1815
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin DicomObjects.DicomViewer dcmMiniImageC 
         Height          =   1695
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3135
         _Version        =   262147
         _ExtentX        =   5530
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picReportImage 
      Height          =   2055
      Left            =   3480
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
      Begin DicomObjects.DicomViewer dcmReportImage 
         Height          =   1695
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _Version        =   262147
         _ExtentX        =   3413
         _ExtentY        =   2990
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picMark 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      Begin VB.PictureBox picNumMark 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1030
         Left            =   300
         ScaleHeight     =   1035
         ScaleWidth      =   2040
         TabIndex        =   6
         Top             =   1300
         Width           =   2040
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":032E
            Height          =   510
            Index           =   1
            Left            =   490
            Picture         =   "frmReportImage.frx":0F70
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":1BB2
            Height          =   510
            Index           =   4
            Left            =   510
            Picture         =   "frmReportImage.frx":27F4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":3436
            Height          =   510
            Index           =   2
            Left            =   1000
            Picture         =   "frmReportImage.frx":4078
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":4CBA
            Height          =   510
            Index           =   5
            Left            =   1010
            Picture         =   "frmReportImage.frx":58FC
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":653E
            Height          =   510
            Index           =   3
            Left            =   1560
            Picture         =   "frmReportImage.frx":7180
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   0
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":7DC2
            Height          =   510
            Index           =   6
            Left            =   1510
            Picture         =   "frmReportImage.frx":8A04
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":9646
            Height          =   1020
            Index           =   0
            Left            =   0
            Picture         =   "frmReportImage.frx":A288
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "�Զ����"
            Top             =   0
            Value           =   1  'Checked
            Width           =   510
         End
      End
      Begin DicomObjects.DicomViewer dcmMark 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _Version        =   262147
         _ExtentX        =   2990
         _ExtentY        =   1720
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Private mlngAdviceID As Long    'ҽ��ID
Private mintEditType As Integer '����״̬ 0 ������1��д��2 �޶�
Private mlngReportID As Long    '��������ID
Private mlngFileID As Long      '���浥��ʽID
Private mlngShowBigImg As Long          '�Ƿ���ʾ��ͼ,0-����ʾ��1-����ƶ�ʱ��ʾ��2-��굥����ʾ��������
Private mdblBigImgZoom As Double        '�����ͼ�Ŵ���
Private mintImageDblClick As Integer    '����ͼ˫����Ĳ��� 0--ֱ��д�뱨�棻1--��ͼƬ�༭����
Private mblnEditable As Boolean         '�Ƿ���Ա༭����
Private mintMoustType As Integer        '��깤������
Private mblnUserInvoke As Boolean       '�Ƿ��û���������
Private mblnMoved As Boolean            '�Ƿ��Ѿ�ת��
Private mintCurImgIndex As Integer      '��ǰѡ�е�ͼ��
Private mintShowPhotoNumber As Integer  '��ǰ�����ܹ���ʾ��ͼ����������
Private mlngModule As Long

Public mSelMiniImg As DicomImage
Private mSelReportImg As DicomImage
Private mSelViewerIndex As Integer  '��ǰ��ѡ�еı���ͼ���ID����1��ʼ����
Private mselReportImgIndex As Integer   '��ǰ��ѡ�еı���ͼ��ID����1��ʼ����
Private mdblMarkZoom As Double          '��ǰ���ͼ��ʵ�����غͱ��֮������ű���
Private lngColor(10) As Long             '���ͼ��Բ�α��ʹ�õ�9����ɫ
Private mlngCY1 As Long                 '���ͼ�ĸ߶�
Private mlngMarkW As Long               '���ͼ�Ŀ��
Private mlngCY2 As Long                 '����ͼ�ĸ߶�
Private mlngRptImgW As Long             '����ͼ�Ŀ��
Private mlngCY3 As Long                 '����ͼͼ�ĸ߶�

Public pMarkModified As Boolean        '���ͼ�ı���иĶ�
Public pImageModified As Boolean       '��¼����ͼ���Ƿ��޸ģ����û���޸ģ��򱣴汨���ʱ���ٱ���ͼ��
Public pobjMarks As cPicMarks          '��ǰ���ͼ�ı�ע����
Public pMarkImageID As Long            '��ǰ���ͼ�����ݿ�����Ӳ������ݡ����е�ID
Public pTableID As String              '��ǰͼ�����ڱ���ID�����á�;���ָ���


Private mintShowMarkImage As Integer   '�Ƿ���ʾ���ͼ   0-���ر��ͼ  1-��ʾ���ͼ
Private mblnIsInitFace As Boolean        '�Ƿ��Ѿ����ش���

Private blnLoadImages As Boolean        '��¼����ˢ���Ƿ������ͼ��


Private mdcmGlobal As New DicomGlobal    '����UIDRoot=1

Private mblnUseActiveVideo As Boolean

Private Enum MarkType
    �Զ���� = 0: ���1: ���2: ���3: ���4: ���5: ���6
End Enum

Property Get ImageCount() As Long
    If mblnUseActiveVideo Then
'        ImageCount = mobjStudyImage.Images.CurImageCount
        ImageCount = ucMiniImageViewer.CurImageCount
    Else
        ImageCount = dcmMiniImageC.Images.Count
    End If
End Property

Property Get dcmImages() As Object
    If mblnUseActiveVideo Then
'        Set dcmImages = mobjStudyImage.Images.ImgViewer.Images
        Set dcmImages = ucMiniImageViewer.ImgViewer.Images
    Else
        Set dcmImages = dcmMiniImageC.Images
    End If
End Property

Public Sub MovePage(ByVal lngPageType As TMoveType)
'�ƶ�����ͼҳ��
    If mblnUseActiveVideo Then
        ucMiniImageViewer.MovePage (lngPageType)
    End If
End Sub


Public Sub zlRefresh(ByVal lngAdviceID As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        dblBigImgZoom As Double, intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, ByVal lngModule As Long)
    Dim i As Integer
    Dim intShowMarkImage As Integer
    
    mlngAdviceID = lngAdviceID
    mlngFileID = FileID
    mlngReportID = ReportID
    mlngShowBigImg = lngShowBigImg
    mdblBigImgZoom = dblBigImgZoom
    mintImageDblClick = intImageDblClick
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    mintShowPhotoNumber = intMinImageCount
    mlngModule = lngModule
    mblnSingleWindow = blnSingleWindow
    
    intShowMarkImage = DecideMarkImagesVisible    '�жϱ��ͼ�Ƿ�ɼ�
    
    
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.MouseMoveZoom = dblBigImgZoom
    ucMiniImageViewer.ShowPopup = False
    
    
    '�ж������ �������� ���� û�м��ع����� ���� ���ͼ״̬�Ѿ��ı䣬�����¼��س�ʼ������
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Then
        mintShowMarkImage = intShowMarkImage
        Call InitLoaclParas     '��ȡ��������
        Call InitFaceScheme     '��ʼ���������
    End If
    
    
    
    '���³�ʼ���ڲ�����
    pTableID = ""
    pMarkImageID = 0
    pImageModified = False
    pMarkModified = False
    dcmMark.Images.Clear
    
    If Not (pobjMarks Is Nothing) Then
        For i = 1 To pobjMarks.Count
            pobjMarks.Remove 1
        Next i
    End If
    
    
    '��Ǳ���ˢ�»�û�м���ͼ��
    blnLoadImages = False
    '������������ڱ���ʾ�ģ������ͼ��
    If blnFormIsSelected = True And Me.Visible Then
        '������Ҫ����ͼ��
        Call LoadImages
    Else
        Call ClearReportImages
    End If
    
    '���ý���ؼ��Ƿ���Ա༭
    picMark.Enabled = mblnEditable
    picReportImage.Enabled = mblnEditable
    picMiniImageC.Enabled = mblnEditable
    picMiniViewer.Enabled = mblnEditable
End Sub

Private Sub ClearReportImages()
    Dim i As Integer
    
    '��ʼ����������
    For i = 1 To dcmReportImage.Count - 1
        Unload dcmReportImage(i)
    Next i
    dcmMark.Images.Clear
End Sub

Public Sub RefPacsPic()
    '��ȡ����ʾ��ǰ��ѡ����ͼ��
    Call LoadMiniImages
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case control.ID
        Case comMenu_Cap_Process 'ͼ����
            Call OpenImageProcessWind
        Case conMenu_Cap_DevSet
            Call ucMiniImageViewer.ShowPageConfig
        Case conMenu_PacsReport_DelImage    'ɾ��ͼ��
            If dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex And mselReportImgIndex <> 0 Then
                dcmReportImage(mSelViewerIndex).Images.Remove mselReportImgIndex
                Call picReportImage_Resize
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveUp      'ǰ��ͼ��
            If mselReportImgIndex > 1 And dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex - 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_MoveDown    '����ͼ��
            If mselReportImgIndex > 0 And dcmReportImage(mSelViewerIndex).Images.Count > mselReportImgIndex Then
                dcmReportImage(mSelViewerIndex).Images.Move mselReportImgIndex, mselReportImgIndex + 1
                pImageModified = True
            End If
        Case conMenu_PacsReport_DelMarks    '�����ע
            If dcmMark.Images.Count > 0 Then
                dcmMark.Images(1).Labels.Clear
                dcmMark.Refresh
                For i = 1 To pobjMarks.Count
                    pobjMarks.Remove 1
                Next i
                pMarkModified = True
            End If
        Case conMenu_View_Refresh           'ˢ��
            '��ȡ����ʾ��ǰ��ѡ����ͼ��
            Call LoadMiniImages
        Case conMenu_PacsReport_DelMiniImage    'ɾ������ͼ
            
        Case conMenu_PacsReport_SelMiniImage    '��ȡ����ͼ
            Dim resImages As DicomImages
            
            Set resImages = frmSelectRepImage.ShowMe(Me, mlngAdviceID, mlngShowBigImg, mdblBigImgZoom)
            '�ѵ�ǰͼ����ӵ�ͼ�����
            If resImages.Count > 0 Then
                For i = 1 To resImages.Count
                    dcmReportImage(mSelViewerIndex).Images.Add resImages(i)
                    dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
                Next i
                dcmReportImage(mSelViewerIndex).CurrentIndex = 1
                Call picReportImage_Resize
                pImageModified = True
            End If
    End Select
End Sub

Private Sub chkMark_Click(Index As Integer)
    Dim i As Integer
    If mblnUserInvoke = False Then
        mblnUserInvoke = True
    Select Case Index
        Case 0
            mintMoustType = MarkType.�Զ����
        Case 1
            mintMoustType = MarkType.���1
        Case 2
            mintMoustType = MarkType.���2
        Case 3
            mintMoustType = MarkType.���3
        Case 4
            mintMoustType = MarkType.���4
        Case 5
            mintMoustType = MarkType.���5
        Case 6
            mintMoustType = MarkType.���6
    End Select
    For i = 0 To 6
        chkMark(i).value = 0
    Next i
    chkMark(Index).value = 1
    mblnUserInvoke = False
    End If
End Sub

Private Sub dcmMark_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lTemp As DicomLabel
    Dim strNum As Integer
    
    If Button = 1 And dcmMark.Images.Count > 0 And picMark.MousePointer = 99 Then
        '����ע
        '�������͵ı�ע��һ����ֱ���Զ���ţ���һ�����ֹ����
        pobjMarks.Add pobjMarks.Count + 1
        pobjMarks(pobjMarks.Count).Selected = False
        pobjMarks(pobjMarks.Count).���� = 6     'Բ�α��
        If mintMoustType = MarkType.�Զ���� Then
            pobjMarks(pobjMarks.Count).���� = pobjMarks.Count
        Else
            Select Case mintMoustType
                Case MarkType.���1
                    pobjMarks(pobjMarks.Count).���� = 1
                Case MarkType.���2
                    pobjMarks(pobjMarks.Count).���� = 2
                Case MarkType.���3
                    pobjMarks(pobjMarks.Count).���� = 3
                Case MarkType.���4
                    pobjMarks(pobjMarks.Count).���� = 4
                Case MarkType.���5
                    pobjMarks(pobjMarks.Count).���� = 5
                Case MarkType.���6
                    pobjMarks(pobjMarks.Count).���� = 6
            End Select
        End If
        '�㼯û������
        Set lTemp = New DicomLabel
        lTemp.Left = X
        lTemp.Top = Y
        lTemp.Width = 20
        lTemp.Height = 20
        lTemp.ImageTied = True
        lTemp.Rescale dcmMark.Images(1)
        pobjMarks(pobjMarks.Count).X1 = lTemp.Left / mdblMarkZoom
        pobjMarks(pobjMarks.Count).Y1 = lTemp.Top / mdblMarkZoom
        pobjMarks(pobjMarks.Count).X2 = pobjMarks(pobjMarks.Count).X1
        pobjMarks(pobjMarks.Count).Y2 = pobjMarks(pobjMarks.Count).Y1
        pobjMarks(pobjMarks.Count).���ɫ = lngColor(pobjMarks.Count Mod 9 + 1)
        pobjMarks(pobjMarks.Count).��䷽ʽ = -2
        '����ɫ���գ�����ɫ����
        pobjMarks(pobjMarks.Count).���� = 1
        pobjMarks(pobjMarks.Count).�߿� = 1
        Set pobjMarks(pobjMarks.Count).���� = New StdFont '  "����"
        drawPicMarks dcmMark.Images(1), pobjMarks
        dcmMark.Refresh
        
        pMarkModified = True
    End If
End Sub

Private Sub dcmMark_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If dcmMark.Images.Count = 1 Then
        '�������
        If dcmMark.ImageXPosition(X, Y) > 0 And dcmMark.ImageXPosition(X, Y) < dcmMark.Images(1).SizeX _
           And dcmMark.ImageYPosition(X, Y) > 0 And dcmMark.ImageYPosition(X, Y) < dcmMark.Images(1).SizeY Then
            picMark.MousePointer = 99
            picMark.MouseIcon = listCur.ListImages("Pen").Picture
        Else
            picMark.MousePointer = 0
            picMark.MouseIcon = Nothing
        End If
    End If
End Sub

Private Sub dcmMark_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 2 Then ShowPopupMark
End Sub

Private Sub OpenImageProcessWind()
    Call frmReportImageEdit.zlShowMe(mSelMiniImg, Me, mintCurImgIndex, mSelViewerIndex, mlngModule)
End Sub

Private Sub dcmMiniImageC_DblClick()
    
    If dcmMiniImageC.Images.Count > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '�жϵ�ǰ˫���Ĳ�������
        If mintImageDblClick = 0 Then   'ֱ��д�뱨��
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '���ý���ǰͼ����ӵ�ͼ������
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '�ȴ�ͼƬ�༭����
            '�ȹرմ�ͼ����
            ReleaseCapture      '�������
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
End Sub

Public Sub DcmAddImage(dcmImage As DicomImage, SelViewerIndex As Integer)
'�ѵ�ǰͼ����ӵ�ͼ�����
    If Not dcmImage Is Nothing Then
        dcmReportImage(SelViewerIndex).Images.Add dcmImage
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).Tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(SelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
    End If
End Sub

Private Sub dcmMiniImageC_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
    If dcmMiniImageC.Images.Count > 0 Then
        For i = 1 To dcmMiniImageC.Images.Count
            dcmMiniImageC.Images(i).BorderColour = vbWhite
            
        Next i
        
        i = dcmMiniImageC.ImageIndex(X, Y)
        If i = 0 Then
            Set mSelMiniImg = dcmMiniImageC.Images(1)
        Else
            Set mSelMiniImg = dcmMiniImageC.Images(i)
        End If
        
        mSelMiniImg.BorderColour = vbRed
        
        mintCurImgIndex = i
        
        '�ж��Ƿ���Ҫ��ʾ��ͼ
        If mlngShowBigImg = 2 Then
            frmShowImg.ShowMe mSelMiniImg, Me, 2, 0, 0
        End If
    End If
End Sub

Private Sub dcmMiniImageC_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim blnShowImg As Boolean
    Dim intCurrImg As Integer

    If dcmMiniImageC.Images.Count <= 0 Or mlngShowBigImg <> 1 Then Exit Sub
    
    '�ж��Ƿ���Ҫ��ʾͼ��
    If (0 <= X * Screen.TwipsPerPixelX) And (X * Screen.TwipsPerPixelX <= dcmMiniImageC.Width) And _
       (0 <= Y * Screen.TwipsPerPixelY) And (Y * Screen.TwipsPerPixelY <= dcmMiniImageC.Height) Then
        blnShowImg = True
    End If
    If blnShowImg Then      '��ʾͼ��
        SetCapture dcmMiniImageC.hWnd    '�������
        
        intCurrImg = dcmMiniImageC.ImageIndex(X, Y)
        If intCurrImg <> 0 Then
            '����ͼ����ʾ
            frmShowImg.ShowMe dcmMiniImageC.Images(intCurrImg), Me, 1, 0, 0, mdblBigImgZoom
        Else
            frmShowImg.HideMe
        End If
    Else        '�ر�ͼ����ʾ
        ReleaseCapture      '�������
        frmShowImg.HideMe
    End If
End Sub

Private Sub dcmMiniImageC_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngShowBigImg = 1 Then  '�رմ�ͼ��ʾ
        ReleaseCapture      '�������
        frmShowImg.HideMe
    End If
    
    If Button = 2 Then Call ShowPopupImage(True)
End Sub

Private Sub dcmReportImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim i As Integer
    
'    If dcmReportImage(Index).Images.Count = 0 Then Exit Sub
    
    mSelViewerIndex = Index
    mselReportImgIndex = dcmReportImage(Index).ImageIndex(X, Y)
    
    For i = 1 To dcmReportImage.Count - 1
        dcmReportImage(i).Labels(1).ForeColour = vbWhite
        dcmReportImage(i).Refresh
    Next i
    dcmReportImage(Index).Labels(1).ForeColour = vbRed
    dcmReportImage(Index).Refresh
    
    If mselReportImgIndex <> 0 Then
        For i = 1 To dcmReportImage(Index).Images.Count
            dcmReportImage(Index).Images(i).BorderColour = vbWhite
        Next i
        dcmReportImage(Index).Images(mselReportImgIndex).BorderColour = vbBlue
    End If
    
    
End Sub

Private Sub dcmReportImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If Button = 2 Then Call ShowPopupImage(False)
End Sub

Private Sub ShowPopupImage(ByVal blnIsDcmMiniImage As Boolean)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    If mblnUseActiveVideo Then
'        If mobjStudyImage.Images.CurImageCount < 1 Then Exit Sub
        If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
    Else
        '�������ͼû��ͼ�����ֹ�Ҽ�����
        If Me.dcmMiniImageC.Images.Count < 1 Then Exit Sub
    End If

    
    '����Ҽ������˵�
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If Not blnIsDcmMiniImage Then
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelImage, "ɾ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveUp, "ǰ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveDown, "����")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SelMiniImage, "��ȡ����ͼ")
         Else
            Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "ͼ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "��ҳ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
         End If
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub ShowPopupMark()
    '------------------------------------------------
'���ܣ���������Ҽ������˵�
'------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelMarks, "�����ע")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub Form_Activate()
    '������Ҫ����ͼ��
    
    'ע����Form��Activate��Paintʱ���б������LoadImages����
    '��Ϊ���ֻ��Activate�����е���LoadImages������������ɱ���ͼ�����ڵ�һʱ����ʾ�������������һ�±���ͼ�Ż���ʾ
    '���ֻ��Paint�����е���LoadImages���������ڸ÷�����ʹ����UnLoadж�ؿؼ����飬������ɡ����ܴӸ���������ж�ء��Ĵ���
    
    Call LoadImages
End Sub

Private Sub Form_Load()
    
    '��Ǳ���ˢ���Ѿ�����ͼ��
    blnLoadImages = True
    
    '��Ǵ����Ѿ��״μ���
    mblnIsInitFace = False
        
    mintMoustType = MarkType.�Զ����

    
    '����Ĭ����ɫ
    lngColor(1) = RGB(186, 186, 186)
    lngColor(2) = RGB(255, 215, 0)
    lngColor(3) = RGB(255, 0, 255)
    lngColor(4) = RGB(255, 0, 130)
    lngColor(5) = RGB(0, 255, 0)
    lngColor(6) = RGB(130, 255, 255)
    lngColor(7) = RGB(255, 255, 0)
    lngColor(8) = RGB(0, 0, 255)
    lngColor(9) = RGB(0, 160, 0)
    
    '����UIDRoot=1
    mdcmGlobal.RegString("UIDRoot") = "1"
    
End Sub

Public Sub MouseWheel(intDirection As Integer)
'���������ֵ��¼�
'������intDirection --- �����ֵķ���1--���ϣ�0--����
    
    On Error Resume Next
    
    If vscrollMini.Visible = False Then Exit Sub
    
    If intDirection = 1 Then '�Ϸ�һҳ
        If vscrollMini.value - 1 < 1 Then
            vscrollMini.value = 1
        Else
            vscrollMini.value = vscrollMini.value - 1
        End If
    Else        '�·�һҳ
        If vscrollMini.value + 1 > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
        Else
            vscrollMini.value = vscrollMini.value + 1
        End If
    End If
End Sub

Public Sub subDispScroll()
'------------------------------------------------
'���ܣ��Զ��ж��Ƿ���Ҫ��ʾ�����ع�����
'���أ��ޣ�ֱ����ʾ�����ع�������
'------------------------------------------------
    Dim ii As Integer
    
    If dcmMiniImageC.Images.Count > dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows Then       'ͼ������������ʾ������ʾ������
        '�ڷŹ�����λ�ã�����ʾ������
        vscrollMini.Move dcmMiniImageC.Width - vscrollMini.Width, dcmMiniImageC.Top, vscrollMini.Width, dcmMiniImageC.Height
        vscrollMini.Visible = True
        vscrollMini.ZOrder
        vscrollMini.Refresh
        
        ''''''''''''''''''[���ڹ�������Ҫ������ϸ����]'''''''''''''''''''''''''
        vscrollMini.Min = 1
        vscrollMini.Max = dcmMiniImageC.Images.Count - dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows + 1
        If vscrollMini.Max < 1 Then vscrollMini.Max = 1
        vscrollMini.LargeChange = dcmMiniImageC.MultiColumns * dcmMiniImageC.MultiRows
        If dcmMiniImageC.CurrentIndex > vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
            dcmMiniImageC.CurrentIndex = vscrollMini.Max
        Else
            vscrollMini.value = dcmMiniImageC.CurrentIndex
        End If
    Else                'ͼ�������ڿ���ʾ�������ع�����
'        ii = dcmMiniature.Images.Count - dcmMiniature.MultiColumns * dcmMiniature.MultiRows + 1
'        If dcmMiniature.Images.Count - dcmMiniature.CurrentIndex + 1 < dcmMiniature.MultiColumns * dcmMiniature.MultiRows Then
'            dcmMiniature.CurrentIndex = IIf(ii < 1, 1, ii)
'        End If
'        vscrollMini.Value = dcmMiniature.CurrentIndex
        vscrollMini.Visible = False
    End If
    
    If vscrollMini.Visible = True Then
        dcmMiniImageC.Width = dcmMiniImageC.Width - vscrollMini.Width - 20
        
        vscrollMini.Height = dcmMiniImageC.Height - 40
        vscrollMini.Left = dcmMiniImageC.Width - 20
    Else
        dcmMiniImageC.Width = dcmMiniImageC.Width
    End If
End Sub


Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage"
    End If
    
    ucMiniImageViewer.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "��������ͼ����", 5))
    
    '��ȡ���ͼ���򣬱���ͼ���� ������ͼ����ĸ߶�
    mlngCY1 = GetSetting("ZLSOFT", strRegPath, "CY1", 180)
    mlngMarkW = GetSetting("ZLSOFT", strRegPath, "MarkW", 300)
    mlngCY2 = GetSetting("ZLSOFT", strRegPath, "CY2", 400)
    mlngRptImgW = GetSetting("ZLSOFT", strRegPath, "RptImgW", 100)
    mlngCY3 = GetSetting("ZLSOFT", strRegPath, "CY3", 200)
End Sub

Private Sub Form_Paint()
    '������Ҫ����ͼ��
    Call LoadImages
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage"
    End If
    
    Call SaveSetting("ZLSOFT", strRegPath, "��������ͼ����", ucMiniImageViewer.PageImgCount)
    
    '������ͼ���򣬱���ͼ���������ͼ����ĸ߶�
    '285��Pane�ı���߶ȣ�ʹ���˱��⣬����Ҫ�ӻ�����߶�
    SaveSetting "ZLSOFT", strRegPath, "CY1", picMark.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "MarkW", picMark.Width
    SaveSetting "ZLSOFT", strRegPath, "CY2", picReportImage.Height + 285
    SaveSetting "ZLSOFT", strRegPath, "RptImgW", picReportImage.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", picMiniImageC.Height + 285
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX3", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", Me.Height
End Sub

Private Sub picMark_Resize()
    If picMark.Height = 0 Or picMark.Width = 0 Then Exit Sub
    
    On Error Resume Next
    
    '�жϿ�߱�
    If picMark.Width / picMark.Height > 2 Then  '���ֱ�Ƿ����ұ�
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = Abs(picMark.ScaleWidth - picNumMark.ScaleWidth - 50)
        dcmMark.Height = picMark.ScaleHeight
        
        picNumMark.Left = dcmMark.Width
        If picMark.Height > picNumMark.Height Then
            picNumMark.Top = (picMark.ScaleHeight - picNumMark.ScaleHeight) / 2
        Else
            picNumMark.Top = 0
        End If
    Else    '���ֱ�Ƿ�������
        dcmMark.Left = 0
        dcmMark.Top = 0
        dcmMark.Width = picMark.ScaleWidth
        dcmMark.Height = Abs(picMark.ScaleHeight - picNumMark.ScaleHeight - 50)
        
        If picMark.Width > picNumMark.Width Then
            picNumMark.Left = (picMark.ScaleWidth - picNumMark.ScaleWidth) / 2
        Else
            picNumMark.Left = 0
        End If
        picNumMark.Top = dcmMark.Height
    End If
End Sub

Private Sub picMiniImageC_Resize()
'    If picMiniImage.Width < 50 Or picMiniImage.Height < 50 Then Exit Sub
'    dcmMiniImage.Left = 0
'    dcmMiniImage.Top = 0
'    dcmMiniImage.Width = picMiniImage.Width - 50
'    dcmMiniImage.Height = picMiniImage.Height - 50

    Dim iRows As Integer
    Dim iCols As Integer
    
    On Error Resume Next
    
    dcmMiniImageC.Left = 0
    dcmMiniImageC.Top = 0
    dcmMiniImageC.Width = picMiniImageC.Width
    dcmMiniImageC.Height = picMiniImageC.Height
    
    '�Զ���ͼ��������
    '��������ͼ��ͼ�񲼾�
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, picMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
    
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows
    '���������
    'If vscrollMini.Visible = True Then
    dcmMiniImageC.Width = picMiniImageC.Width - vscrollMini.Width - 20
    
    vscrollMini.Height = dcmMiniImageC.Height - 40
    vscrollMini.Left = dcmMiniImageC.Width - 20
    'End If
End Sub


Private Sub InitFaceScheme()
    '��ʼ���沼��
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    
    With dkpMain
        .CloseAll
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    If mintShowMarkImage = 1 Then
        picMark.Visible = True
        dcmMark.Visible = True
        picNumMark.Visible = True
        
        Set Pane1 = dkpMain.CreatePane(1, mlngMarkW, mlngCY1, DockTopOf, Nothing)
        Pane1.Title = "���ͼ"
        Pane1.Handle = picMark.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        '���ݿ�߱ȣ��ڷű���ͼ��λ��
        If ((mlngCY1 = mlngCY2) And (mlngMarkW + mlngRptImgW > mlngCY1)) _
            Or (((mlngCY1 <> mlngCY2)) And (mlngMarkW + mlngRptImgW > mlngCY1 + mlngCY2)) Then
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockLeftOf, Pane1)
        Else
            Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockBottomOf, Pane1)
        End If
    Else
        picMark.Visible = False
        dcmMark.Visible = False
        picNumMark.Visible = False
        
        Set Pane2 = dkpMain.CreatePane(2, mlngRptImgW, mlngCY2, DockTopOf, Nothing)
    End If
    
'    If mobjStudyImage Is Nothing Then
'        Set mobjStudyImage = New clsStudyImages
'    End If
    
    mblnUseActiveVideo = False
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Or G_LNG_PATHSTATION_MODULE Then
        mblnUseActiveVideo = GetSetting("ZLSOFT", "����ģ��", "UseActiveVideo", "true")
        Call SaveSetting("ZLSOFT", "����ģ��", "UseActiveVideo", mblnUseActiveVideo)
    End If

    Pane2.Title = "����ͼ"
    Pane2.Handle = picReportImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    picMiniImageC.Visible = Not mblnUseActiveVideo
    picMiniViewer.Visible = mblnUseActiveVideo

    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Nothing)
    Pane3.Title = "����ͼ"
    Pane3.Handle = IIf(mblnUseActiveVideo, picMiniViewer.hWnd, picMiniImageC.hWnd)
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    mblnIsInitFace = True
End Sub


Private Function LoadMiniImages() As Boolean
    Dim lngMsgHwnd As Long
    
    
    If mblnUseActiveVideo Then
'        lngMsgHwnd = mobjStudyImage.hWnd
'
'        Call mobjStudyImage.RefreshImages(mlngAdviceID, mlngAdviceID, mblnMoved, True)

        Call ucMiniImageViewer.RefreshImage(0, mlngAdviceID, mblnMoved, True)
    Else
        Call GetRptImages(dcmMiniImageC, mlngAdviceID, mblnMoved)
    
        Call AdjustDicomViewerLayout
    End If

End Function


Private Sub AdjustDicomViewerLayout()
'------------------------------------------------
'���ܣ���ͼ����ӵ�����ͼdcmMiniature��
'������img���������DICOMͼ��
'���أ��ޣ�ֱ�ӽ�ͼ����ӵ�����ͼdcmMiniature��
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer
    
    '��������ͼ��ͼ�񲼾�
    If mintShowPhotoNumber < dcmMiniImageC.Images.Count + 1 Then
        ResizeRegion mintShowPhotoNumber, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImageC.Images.Count + 1, dcmMiniImageC.Width, dcmMiniImageC.Height, iRows, iCols
    End If
            
    dcmMiniImageC.MultiColumns = iCols
    dcmMiniImageC.MultiRows = iRows

    
'    '��������ͼ�ļ��UID������UID���޸�img��ֵ
'    subUniteUID img
'    dcmMiniature.Images.Add img
    
'    '��������ͼ�б�ѡ�е�״̬
'    If mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniature.Images.Count Then
'        dcmMiniature.Images(mintCurImgIndex).BorderColour = vbWhite
'    End If
    
    If dcmMiniImageC.Images.Count > 0 Then
         With dcmMiniImageC.Images(1)
            .BorderWidth = 1
            .BorderStyle = 6
            .BorderColour = vbRed
        End With
    End If
    
    mintCurImgIndex = 1
    
'    mintCurImgIndex = dcmMiniature.Images.Count
    '��ʾ������
    Call subDispScroll
End Sub


Private Sub LoadReportImages()
    
    On Error GoTo errH
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
    '��ʼ����������
    Call ClearReportImages
        
    '������ڱ������ݣ���ӱ��������ж�ȡ����ͼ�ͱ��ͼ������ӱ��浥��ʽ�ж�ȡ���ͼ
    If mlngReportID <> 0 Then
        strSql = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        If mblnMoved = True Then
            strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngReportID)
    Else
        strSql = "Select Id As ���Id From �����ļ��ṹ" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    End If
    
    iRImageCount = 0
    Do While Not rsTemp.EOF
    
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_�����ļ�����, mlngFileID, Val("" & rsTemp!���ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_���������, mlngReportID, Val("" & rsTemp!���ID))
        End If
        If blnGetTable Then
            
            '��¼ͼ�����ڱ��ID
            If pTableID = "" Then
                pTableID = cTable.ID
            Else
                pTableID = pTableID & ";" & cTable.ID
            End If
            
            '����viewer
            iRImageCount = iRImageCount + 1
            Load dcmReportImage(iRImageCount)
            dcmReportImage(iRImageCount).BorderStyle = 1
            dcmReportImage(iRImageCount).Labels.AddNew
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
            dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
            dcmReportImage(iRImageCount).Visible = True
            
            mSelViewerIndex = iRImageCount

            '��¼ͼ���Ŀ�Ⱥ͸߶ȣ��ÿ�߱������ں�����ͼ�����в���
            If cTable.ExtendTag <> "" Then
                If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                    dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                Else
                    dcmReportImage(iRImageCount).Tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                End If
            Else
                dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
            End If
            
            
            For i = 1 To cTable.Pictures.Count
                strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                Set oPicture = cTable.Pictures(i).OrigPic
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '��ʾ���ͼ�ͱ���ͼ
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture And dcmMark.Images.Count = 0 Then

                        'ֻ�����һ�����ͼ
                        dcmMark.Images.AddNew
                        
                        dcmMark.Images(1).FileImport strPicFile, "BMP"
                        dcmMark.Images(1).Tag = cTable.Pictures(i).ID
                        '������ͼ��������
                        Set pobjMarks = cTable.Pictures(i).PicMarks
                        pMarkImageID = cTable.Pictures(i).ID

                        mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                        '��ʾ��ע
                        If cTable.Pictures(i).PicMarks.Count > 0 Then
                            drawPicMarks dcmMark.Images(1), cTable.Pictures(i).PicMarks
                        End If
                    Else

                        dcmReportImage(iRImageCount).Images.AddNew
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).FileImport strPicFile, "BMP"
                        If cTable.Pictures(i).PicName = "" Then
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).Tag = mdcmGlobal.NewUID & ".jpg"
                        Else
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).Tag = cTable.Pictures(i).PicName
                        End If
                        
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderWidth = 3
                        dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).BorderColour = vbWhite
                        dcmReportImage(iRImageCount).CurrentIndex = 1
                        mselReportImgIndex = 1
                    End If
                    'ɾ����ʱͼ��
                    Kill strPicFile
                End If
            Next
        End If
        
        rsTemp.MoveNext
    Loop
    If dcmReportImage.Count > 1 Then dcmReportImage(1).Labels(1).ForeColour = vbRed
    Call picReportImage_Resize
    

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function DecideMarkImagesVisible() As Integer
'------------------------------------------------
'���ܣ��жϵ�ǰѡ�м����ͼ�Ƿ�ɼ�
'��������
'���أ�int���ͣ�1-��ʾ���ͼ  2-���ر��ͼ
'-----------------------------------------------
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim blnGetTable As Boolean
    Dim i As Integer
    
        
    '������ڱ������ݣ���ӱ��������ж�ȡ���ݣ�����ӱ��浥��ʽ�ж�ȡ����
    If mlngReportID <> 0 Then
        strSql = "Select Id As ���Id From ���Ӳ�������" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        If mblnMoved = True Then
            strSql = Replace(strSql, "���Ӳ�������", "H���Ӳ�������")
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "���������ж�ȡ", mlngReportID)
    Else
        strSql = "Select Id As ���Id From �����ļ��ṹ" & vbNewLine & _
            " Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
            " Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�ӱ��浥��ʽ�ж�ȡ", mlngFileID)
    End If
    
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_�����ļ�����, mlngFileID, Val("" & rsTemp!���ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_���������, mlngReportID, Val("" & rsTemp!���ID))
        End If
        
        If blnGetTable Then
            If cTable.Pictures.Count > 0 Then
                For i = 1 To cTable.Pictures.Count
                    If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                        DecideMarkImagesVisible = 1
                        Exit Do
                    Else
                        DecideMarkImagesVisible = 0
                    End If
                Next
            Else
                DecideMarkImagesVisible = 0
            End If
        End If
        rsTemp.MoveNext
    Loop

End Function


Private Sub drawPicMarks(img As DicomImage, thisMarks As cPicMarks)
'��ʾ��ע��ֻ֧�����ֱ�ű�ע
    Dim i As Integer
    Dim iLabelCount As Integer
    
    img.Labels.Clear
    For i = 1 To thisMarks.Count
        If thisMarks(i).���� = 6 Then   'Բ�α��
            With thisMarks(i)
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).BackColour = IIf(.���ɫ = 0, vbYellow, .���ɫ)
                img.Labels(iLabelCount).Transparent = False
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True
                
                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelEllipse
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Width = 20
                img.Labels(iLabelCount).Height = 20
                img.Labels(iLabelCount).ImageTied = True

                img.Labels.AddNew
                iLabelCount = img.Labels.Count
                img.Labels(iLabelCount).LabelType = doLabelText
                img.Labels(iLabelCount).Transparent = True
                img.Labels(iLabelCount).ForeColour = vbBlack
                img.Labels(iLabelCount).FontSize = 11
                img.Labels(iLabelCount).FontName = "Arial Bold"
                img.Labels(iLabelCount).Left = .X1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).Top = .Y1 * mdblMarkZoom - 10
                img.Labels(iLabelCount).AutoSize = True
                img.Labels(iLabelCount).Text = .����
                img.Labels(iLabelCount).ImageTied = True
            End With
        End If
    Next i
End Sub
 
Private Sub picMiniViewer_Resize()
On Error Resume Next
    ucMiniImageViewer.Left = 0
    ucMiniImageViewer.Top = 0
    ucMiniImageViewer.Width = picMiniViewer.ScaleWidth
    ucMiniImageViewer.Height = picMiniViewer.ScaleHeight
End Sub

Private Sub picReportImage_Resize()
    Dim i As Integer
    Dim rectH As Long, rectW As Long    'ͼ������ʹ�õ�������
    Dim picH As Long, picW As Long      'ͼ��ʵ�ʿ�ߣ���Ϊ����ʹ��
    Dim iCols As Integer, iRows As Integer
    Dim dImg As DicomImage
    
    If dcmReportImage.Count = 1 Then Exit Sub
    
    On Error Resume Next
    
    '���ȼ���ÿ��ͼ����ռ�õ������
    
    rectH = picReportImage.Height / (dcmReportImage.Count - 1)
    rectW = picReportImage.Width
    If rectH < 100 Or rectW < 100 Then Exit Sub
    
    For i = 1 To dcmReportImage.Count - 1
        '����ͼ�����������ͼ������ʵ��Ⱥ͸߶�
        picW = Val(Split(dcmReportImage(i).Tag, "|")(0))
        picH = Val(Split(dcmReportImage(i).Tag, "|")(1))
        
        dcmReportImage(i).Height = rectH - 100
        dcmReportImage(i).Width = rectW - 100
        
        dcmReportImage(i).Left = 0
        dcmReportImage(i).Top = rectH * (i - 1)
        
        dcmReportImage(i).Labels(1).Width = Abs(dcmReportImage(i).Width / Screen.TwipsPerPixelX - 2)
        dcmReportImage(i).Labels(1).Height = Abs(dcmReportImage(i).Height / Screen.TwipsPerPixelY - 1)

        
        '����ͼ����ʾ����
        ResizeRegion dcmReportImage(i).Images.Count, picW, picH, iRows, iCols
        dcmReportImage(i).MultiColumns = iCols
        dcmReportImage(i).MultiRows = iRows
    Next i
End Sub

Public Sub zlChangeFormat(FormatID As Long)

    On Error GoTo errH
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim blnHasMarkImage As Boolean
    Dim blnGetTable As Boolean
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    
    '��ʾ��ʽ�任��ͼ��򡢱��ͼ�ȵ������仯
    If FormatID = 0 Then     '��׼��ʽ���� �����ļ��ṹ
        strSql = "Select Id As ���Id From �����ļ��ṹ" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngFileID)
    Else        '���ĸ�ʽ���� ������������
        strSql = "Select Id As ���Id From ������������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, FormatID)
    End If
    If rsTemp.RecordCount > 0 Then
        If rsTemp.RecordCount < dcmReportImage.Count - 1 Then
            If MsgBoxD(Me, "�¸�ʽ��ͼ����������ڵ�ǰ��ʽ����ǰ�Ĳ���ͼ���ᱻɾ�����Ƿ������ʽ��", vbOKCancel) = vbCancel Then
                Exit Sub
            Else
                '��ɾ�������ͼ���
                For i = dcmReportImage.Count - 1 - rsTemp.RecordCount To 1 Step -1
                    Unload dcmReportImage(dcmReportImage.Count - 1)
                Next i
            End If
        End If
        
        '��ȡͼ����еı��ͼ�ͱ���ͼ
        iRImageCount = 0
        pTableID = ""
        Do While Not rsTemp.EOF
            Set cTable = New cEPRTable
            If FormatID = 0 Then
                blnGetTable = cTable.GetTableFromDB(cprET_�����ļ�����, mlngFileID, Val("" & rsTemp!���ID))
            Else
                blnGetTable = cTable.GetTableFromDB(cprET_ȫ��ʾ���༭, FormatID, Val("" & rsTemp!���ID))
            End If
            If blnGetTable Then
                iRImageCount = iRImageCount + 1
                If iRImageCount > dcmReportImage.Count - 1 Then
                    '����Viewer
                    Load dcmReportImage(iRImageCount)
                    dcmReportImage(iRImageCount).BorderStyle = 1
                    dcmReportImage(iRImageCount).Labels.AddNew
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LabelType = doLabelRectangle
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Left = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).Top = 1
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).LineWidth = 2
                    dcmReportImage(iRImageCount).Labels(dcmReportImage(iRImageCount).Labels.Count).ForeColour = vbWhite
                    dcmReportImage(iRImageCount).Visible = True
                End If
                mSelViewerIndex = iRImageCount
                
                '��¼ͼ���Ŀ�Ⱥ͸߶ȣ��ÿ�߱������ں�����ͼ�����в���
                If cTable.ExtendTag <> "" Then
                    If Val(Split(cTable.ExtendTag, "|")(0)) = 0 Then
                        dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                    Else
                        dcmReportImage(iRImageCount).Tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                    End If
                Else
                    dcmReportImage(iRImageCount).Tag = cTable.Width & "|" & cTable.Height
                End If
                
                '���±��ͼ
                For i = 1 To cTable.Pictures.Count
                    strPicFile = App.Path & "\PACSPic" & i & ".JPG"
                    If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                    Set oPicture = cTable.Pictures(i).OrigPic
                    SavePicture oPicture, strPicFile
                    If objFile.FileExists(strPicFile) Then
                        '��ʾ���ͼ
                        If cTable.Pictures(i).PictureType = EPRMarkedPicture Then
                            blnHasMarkImage = True
                            '�������ǰ���ͼ���ٸ���
                            dcmMark.Images.Clear
                            dcmMark.Images.AddNew
                            dcmMark.Images(1).FileImport strPicFile, "BMP"
                            dcmMark.Images(1).Tag = cTable.Pictures(i).ID
                            '�����ǰû�б�ǣ����ȡ�¸�ʽ�б��ͼ�ı��
                            If pobjMarks Is Nothing Then
                                Set pobjMarks = cTable.Pictures(i).PicMarks
                            End If
                            pMarkImageID = cTable.Pictures(i).ID
    
                            mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(i).Width * Screen.TwipsPerPixelX
                            '��ʾ��ע
                            If pobjMarks.Count > 0 Then
                                drawPicMarks dcmMark.Images(1), pobjMarks
                            End If
                        End If
                        'ɾ����ʱͼ��
                        Kill strPicFile
                    End If
                Next
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    If blnHasMarkImage = False Then
        '��ǰ��ʽû�б��ͼ��ɾ����ǰ��ʾ�ı��ͼ
        pMarkImageID = 0
        
        dcmMark.Images.Clear
        If Not (pobjMarks Is Nothing) Then
            For i = 1 To pobjMarks.Count
                pobjMarks.Remove 1
            Next i
        End If
    End If
    Call picReportImage_Resize

Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ucMiniImageViewer_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
    If ucMiniImageViewer.CurImageCount > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '�жϵ�ǰ˫���Ĳ�������
        If mintImageDblClick = 0 Then   'ֱ��д�뱨��
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '���ý���ǰͼ����ӵ�ͼ������
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            
        Else                            '�ȴ�ͼƬ�༭����
            '�ȹرմ�ͼ����
            ReleaseCapture      '�������
            frmShowImg.HideMe
            
            Call OpenImageProcessWind
        End If

    End If
    
    blnContinue = False
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
If Button = 2 Then Call ShowPopupImage(True)
End Sub

Private Sub ucMiniImageViewer_OnSelChange(ByVal lngSelectedIndex As Long)
    Set mSelMiniImg = ucMiniImageViewer.SelectImage
End Sub

Private Sub vscrollMini_Change()
    Dim iImgIndex As Integer
    
    If dcmMiniImageC.Images.Count > 0 And (mintCurImgIndex > 0 And mintCurImgIndex <= dcmMiniImageC.Images.Count) Then
        iImgIndex = vscrollMini.value + mintCurImgIndex - dcmMiniImageC.CurrentIndex
        If iImgIndex <= 0 Then
            iImgIndex = 1
        ElseIf iImgIndex > dcmMiniImageC.Images.Count Then
            iImgIndex = dcmMiniImageC.Images.Count
        End If
        dcmMiniImageC.CurrentIndex = vscrollMini.value
        
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbWhite
        mintCurImgIndex = iImgIndex
        dcmMiniImageC.Images(mintCurImgIndex).BorderColour = vbRed
    End If

End Sub

Private Sub LoadImages()
'------------------------------------------------
'���ܣ����ر���ͼ������ͼ
'������
'���أ��ޣ�ֱ�Ӽ���ͼ�񣬲��޸� blnLoadImages״̬
'-----------------------------------------------
    '�������ˢ��û�м���ͼ�������ͼ��
    If blnLoadImages = False Then
        '��ȡ����ʾ��ǰ��ѡ����ͼ��
        Call LoadMiniImages
        '���ݱ��浥��ʽ�����߱������ݸ�ʽ����ȡ���ͼ�ͱ���ͼ
        Call LoadReportImages
        '��Ǳ���ˢ���Ѿ�����ͼ��
        blnLoadImages = True
    End If
End Sub
