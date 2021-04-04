VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.0#0"; "zl9PacsControl.ocx"
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
   Begin VB.PictureBox picMiniCache 
      Height          =   3855
      Left            =   4080
      ScaleHeight     =   3795
      ScaleWidth      =   4155
      TabIndex        =   18
      Top             =   2760
      Width           =   4215
      Begin VB.ComboBox cboCache 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   2415
      End
      Begin zl9PacsControl.ucImagePreview ucMiniCache 
         Height          =   1215
         Left            =   240
         TabIndex        =   19
         Top             =   2400
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   2143
         BackColor       =   -2147483629
         ShowCheckbox    =   -1  'True
      End
   End
   Begin VB.PictureBox picMiniViewer 
      Height          =   1365
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   3615
      TabIndex        =   16
      Top             =   5280
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   975
         Left            =   45
         TabIndex        =   17
         Top             =   120
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1720
         BackColor       =   -2147483629
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
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   3675
      TabIndex        =   3
      Top             =   3600
      Width           =   3735
      Begin VB.VScrollBar vscrollMini 
         Height          =   975
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin DicomObjects.DicomViewer dcmMiniImageC 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3255
         _Version        =   262147
         _ExtentX        =   5741
         _ExtentY        =   1931
         _StockProps     =   35
         BackColor       =   -2147483629
      End
   End
   Begin VB.PictureBox picReportImage 
      Height          =   2055
      Left            =   3600
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   120
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
      Top             =   840
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

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
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
Private mrsImageCache As ADODB.Recordset
Private mdcmUID As New DicomGlobal
Private mlngReleationType As Integer    '1--������2--����
Private mlngCurDeptId As Long
Private mlngStudyState As Long
Private mstrTmpQueryValue As String
Private mblnUseAfterCapture As Boolean
Private mblnTmpUseAfterCapture As Boolean

Public Event AfterReleationImage(ByVal lngReleationType As Long)

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


Public Sub zlRefresh(ByVal lngAdviceId As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        dblBigImgZoom As Double, intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, _
        ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal lngStudyState As Long)
    Dim i As Integer
    Dim intShowMarkImage As Integer
    
    mlngCurDeptId = lngCurDeptId
    mlngStudyState = lngStudyState
    mlngAdviceID = lngAdviceId
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
    
    mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "���ú�̨�ɼ�", 1, True)) = 1
    
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.MouseMoveZoom = dblBigImgZoom
    ucMiniImageViewer.ShowPopup = False
    
    
    '�ж������ �������� ���� û�м��ع����� ���� ���ͼ״̬�Ѿ��ı䣬�����¼��س�ʼ������
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Then
        mintShowMarkImage = intShowMarkImage
        Call InitLoaclParas     '��ȡ��������
        Call InitFaceScheme     '��ʼ���������
    Else
        If mblnTmpUseAfterCapture <> mblnUseAfterCapture Then
            Call InitFaceScheme
        End If
    End If
    
    mblnTmpUseAfterCapture = mblnUseAfterCapture
    
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

Public Sub RefPacsPic(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg)
    '��ȡ����ʾ��ǰ��ѡ����ͼ��
    If mblnUseAfterCapture And mlngModule <> 1920 Then
        Call LoadMiniCache(lngEventType)
    End If
    
    If lngEventType <> vetAfterUpdateImg Then Call LoadMiniImages
End Sub

Private Sub cboCache_Click()
    Dim strQueryValue As String
    
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Do While Not mrsImageCache.EOF
        If "������" & Nvl(mrsImageCache!����) & "  ���ţ�" & Nvl(mrsImageCache!����) & "  ����" & Nvl(mrsImageCache!���к�) = cboCache.Text Then
            strQueryValue = Nvl(mrsImageCache!����UID)
            Exit Do
        End If
        
        mrsImageCache.MoveNext
    Loop
    
    Call ucMiniCache.RefreshImage(slSeries, strQueryValue, mblnMoved, True, True)
    mstrTmpQueryValue = strQueryValue
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case control.ID
        Case comMenu_Cap_Process 'ͼ����
            Call OpenImageProcessWind
        Case conMenu_Cap_DevSet
            If mblnUseAfterCapture And mlngModule <> 1290 Then Call ucMiniCache.ShowPageConfig
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
            
        Case conMenu_Edit_Import        '���뱨��ͼ������ͼ
            mlngReleationType = 2
            Call ReleationImage
            Call RefPacsPic
        
        Case conMenu_File_ExportAll     '��������ͼ������ͼ
            mlngReleationType = 1
            Call ReleationImage
            Call RefPacsPic
        
        Case conMenu_Manage_DeleteImage 'ɾ����ʱͼ��
            Call DelTempImage
            Call LoadMiniCache
        
        Case conMenu_Manage_RefreshImg  'ˢ�»���
            Call LoadMiniCache
    End Select
End Sub

Private Sub DelTempImage()
    Dim rsImageDatas As ADODB.Recordset
    Dim i As Long
    
    '�����ݿ��в�ѯͼ������
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "��ѡ����Ҫɾ���ļ��ͼ��", vbInformation, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "�Ƿ�ȷ��ɾ����ѡͼ��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    If DelTempImages(rsImageDatas) Then
        For i = ucMiniCache.CurImageCount To 1
            If ucMiniCache.ImgChecked(i) Then ucMiniCache.DeleteImage (i)
        Next
    End If
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'ɾ��ftp�������е��ļ�
    Dim objSrcFtp As New clsFtp
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strImageUID As String
    Dim strVirtualPath As String
    
    DelTempImages = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If

        'ɾ��ͼ���ļ�����ɾ��ʧ�ܺ����˳�ִ��
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        'ɾ�����ܴ��ڵı���ͼ��
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID & ".jpg")
    
        'ͼ��ɾ���ɹ���ͬ��ɾ�����ݿ��е�����
        Call zlDatabase.ExecuteProcedure("ZL_Ӱ����_ɾ����ʱͼ��(3,'" & strImageUID & "')", Me.Caption)
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend

    objSrcFtp.FuncFtpDisConnect
    
    DelTempImages = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    
    Set rsImageDatas = GetReleationImageIds()
    
    If rsImageDatas Is Nothing Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, Me.Caption)
        Exit Function
    End If
        
    '��ǰ���UID�����ݿ��в����ڣ����˳�������
    If rsImageDatas.RecordCount <= 0 Then
        Call MsgBoxD(Me, "��ѡ����Ҫ���й����ļ��ͼ��", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If mlngReleationType = 2 Then
        '����ͼ����ʾ
        strHint = GetReleationHintInfo(mlngAdviceID, rsImageDatas)
        
        If strHint = "" Then
            Call MsgBoxD(Me, "���ܲ�ѯ����Ҫ������������Ϣ������������", vbOKOnly, Me.Caption)
            Exit Function
        End If
        
        If MsgBoxD(Me, strHint, vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
        
    Else
        'ȡ��������ʾ
        If MsgBoxD(Me, "�Ƿ�ȷ�϶���ѡͼ�����ȡ������������", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '����2��ʾ����ͼ��
        ReleationImage = StartReleation(mlngAdviceID, rsImageDatas)
    Else
        ReleationImage = CancelReleation(mlngAdviceID, rsImageDatas)
    End If
    
    RaiseEvent AfterReleationImage(mlngReleationType)
End Function

'ȡ�ù�����ʾ��Ϣ
Private Function GetReleationHintInfo(lngAdviceId As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strResult As String
    Dim strStudyInf As String
    
    GetReleationHintInfo = ""
    
    If rsReleationImage.RecordCount <= 0 Then Exit Function
    
    Call rsReleationImage.MoveFirst
    While Not rsReleationImage.EOF
        strStudyInf = "[" & Nvl(rsReleationImage!����) & "(" & Nvl(rsReleationImage!����) & ") " & Nvl(rsReleationImage!�Ա�) & " " & Nvl(rsReleationImage!����) & "]"
        
        If InStr(strResult, strStudyInf) <= 0 Then
            If strResult <> "" Then strResult = strResult & "+"
        
            strResult = strResult & strStudyInf
        End If
        Call rsReleationImage.MoveNext
    Wend
    
    strSql = "select ����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    
    GetReleationHintInfo = "�Ƿ�ȷ�Ͻ�  " & strResult & "  ��ͼ����  [" & Nvl(rsTemp!����) & "(" & Nvl(rsTemp!����) & ") " & Nvl(rsTemp!�Ա�) & " " & Nvl(rsTemp!����) & "]  �ļ����й���������"
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'��ѯ��������Ҫȡ��������ͼ��ID
    Dim i As Long, j As Long
    Dim strSql As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String

    j = 0
    strUninTable = ""
    strFilter = ""
    strValue = ""
    
    '�����ѯ���
    If mlngReleationType = 1 Then
        For i = 1 To ucMiniImageViewer.CurImageCount
            If ucMiniImageViewer.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or ͼ��UID ='" & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucMiniImageViewer.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
    Else
        For i = 1 To ucMiniCache.CurImageCount
            If ucMiniCache.ImgChecked(i) Then
                If j > 79 Then
                    strFilter = strFilter & " Or ͼ��UID ='" & ucMiniCache.ImgViewer.Images(i).InstanceUID & "'"
                Else
                    If zlCommFun.ActualLen(strValue) > 3600 Then
                         strValues(j) = Mid(strValue, 2)
                         strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
                         
                         strValue = ""
                         j = j + 1
                    End If
                    
                    strValue = strValue & "," & ucMiniCache.ImgViewer.Images(i).InstanceUID
                End If
            End If
        Next
    End If
    
    If strValue <> "" Then
        strValues(j) = Mid(strValue, 2)
        strUninTable = strUninTable & " Union ALL  Select  Column_Value as ͼ��UID From Table(Cast(f_Str2list([" & j + 1 & "]) As zlTools.t_Strlist))  " & vbCrLf
    End If
    
    '���û����Ҫ���ҵ�ͼ��UID���򷵻ؿ����ݼ�
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
        Set GetReleationImageIds = Nothing
        Exit Function
    End If
    
'    If strFilter <> "" Then strFilter = " and ( " & Mid(strFilter, 4) & ")"
    If strFilter <> "" Then strFilter = strUninTable & " Union All Select ͼ��UID from [Ӱ��ͼ��] where  ( " & Mid(strFilter, 4) & ")"
    
    '�����ƶ��ķ���ͬ��Դͼ�п����ڡ�Ӱ����ʱ��¼�����ߡ�Ӱ�����¼����
    '����ʱ����ʱ��¼���Ƶ�������¼��ȡ������ʱ��������¼���Ƶ���ʱ��¼
    strSql = "Select /*+ RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ͼ�� A, Ӱ�������� B, Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID = E.ͼ��UID " & _
        "Union All " & _
        "Select /*+ RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ʱͼ�� A,Ӱ����ʱ���� B, Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ʱͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID= E.ͼ��UID"
        
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    
    Set GetReleationImageIds = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strValues(0), strValues(1), strValues(2), strValues(3), _
        strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10), _
        strValues(11), strValues(12), strValues(13), strValues(14), strValues(15), strValues(16), strValues(17), _
        strValues(18), strValues(19), strValues(20), strValues(21), strValues(22), strValues(23), strValues(24), strValues(25), strValues(26), _
        strValues(27), strValues(28), strValues(29), strValues(30), strValues(31), strValues(32), strValues(33), strValues(34), strValues(35), strValues(36), _
        strValues(37), strValues(38), strValues(39), strValues(40), strValues(41), strValues(42), strValues(43), strValues(44), strValues(45), strValues(46), _
        strValues(47), strValues(48), strValues(49), strValues(50), strValues(51), strValues(52), strValues(53), strValues(54), strValues(55), strValues(56), _
        strValues(57), strValues(58), strValues(59), strValues(60), strValues(61), strValues(62), strValues(63), strValues(64), strValues(65), strValues(66), _
        strValues(67), strValues(68), strValues(69), strValues(70), strValues(71), strValues(72), strValues(73), strValues(74), strValues(75), strValues(76), _
        strValues(77), strValues(78), strValues(79), strValues(80))
End Function

Private Function StartReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��ʼ����
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As String
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSql = "select ���UID,�������� from Ӱ�����¼ where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "�Ҳ����������ļ����Ϣ��", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!���uid)) = "" Or Trim(Nvl(rsTmp!��������)) = "" Then
        
        '��δ�ɼ�ͼ����Ҫ�����µļ��UID
        strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
            Exit Function
        End If
        
        '���´洢�豸��Ϣ
        strSql = "Zl_Ӱ����_�����豸(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else
        strNewStudyUID = Nvl(rsTmp!���uid)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
        
        Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '��ȡ����ͼ����Ϣ
    strSql = "Select ���UID,����ͼ�� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", lngAdviceId)
    
    strOldReportImages = ""
    lngReportImageLen = 0
    
    If rsReportImage.RecordCount > 0 Then
        strOldReportImages = Nvl(rsReportImage!����ͼ��)
        lngReportImageLen = Len(strOldReportImages)
    End If
        
    '�����µ�����UID
    strNewSeriesUid = CreateSeriesUid(mdcmUID.NewUID)
    
    strReportImageIds = ""
    rsImageDatas.MoveFirst
                
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        '����ͼ���������
        strSql = "Zl_Ӱ����_ͼ�����(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & Nvl(rsImageDatas!ͼ��UID) & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '���汨������
        If InStr(1, strOldReportImages & ";" & strReportImageIds, Nvl(rsImageDatas!ͼ��UID)) <= 0 And Len(strReportImageIds) < 4000 - lngReportImageLen - 60 Then
            If strReportImageIds <> "" Then strReportImageIds = strReportImageIds & ";"
            strReportImageIds = strReportImageIds & Nvl(rsImageDatas!ͼ��UID) & ".jpg"
        End If
    
        rsImageDatas.MoveNext
    Wend
    
    '�����Ҫ���ֱ���ͼ������Ҫ�Ȳ�ѯĿǰ�Ѿ����ֵı���ͼ��UID
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = IIf(strOldReportImages <> "", strOldReportImages & ";", "") & strReportImageIds
        strReportImageIds = Replace(strReportImageIds, ";;", ";")
    End If
    
    strSql = "Zl_Ӱ����_���±���ͼ(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    '�ύ����
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    StartReleation = True
    
    Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '�����׳�����
End Function

Private Function CancelReleation(ByVal lngAdviceId As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��������
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim strReportImageIds As String
    Dim strOldReportImages As Long
    Dim lngReportImageLen As Long
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim rsReportImage As ADODB.Recordset
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    CancelReleation = False
    
    curDate = zlDatabase.Currentdate
    
    '����ͼ�����
    strNewStudyUID = CreateStudyUid(mdcmUID.NewUID)
        
    Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
    If Trim(strNewFtpIp) = "" Then
        Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(rsImageDatas, strNewStudyUID, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd, objMoveList) Then
        Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
        Exit Function
    End If
    
    strSql = "Select ���UID,����ͼ�� From Ӱ�����¼ Where ҽ��ID=[1]"
    Set rsReportImage = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", mlngAdviceID)
    
    If rsReportImage.RecordCount > 0 Then
        strReportImageIds = Nvl(rsReportImage!����ͼ��)
        strReportImageIds = Replace(strReportImageIds, " ", "") '�ɼ�ͼ��ʱ�����ܻ��ڱ���ͼ���ݺ���ӿո�
    End If
    
    '��������
    rsImageDatas.MoveFirst
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    While Not rsImageDatas.EOF
        strSql = "Select D.���UID From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ����ʱ���� D " & _
                 "Where C.ҽ��ID=[1] And A.ͼ��UID=[2] And A.����UID=B.����UID And B.���UID=C.���UID And A.����UID = D.����UID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", mlngAdviceID, Nvl(rsImageDatas!ͼ��UID))

        If rsTmp.RecordCount > 0 Then strNewStudyUID = Nvl(rsTmp!���uid)
            
        strSql = "Zl_Ӱ����_��������(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!ͼ��UID) & "','" & strNewStudyUID & "','" & strNewDeviceNo & "'," & _
                                        "to_Date('" & Format(curDate, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'))"
                                        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        '�޸ı���ͼ����
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!ͼ��UID) & ".jpg;", "")
        strReportImageIds = Replace(strReportImageIds, Nvl(rsImageDatas!ͼ��UID) & ".jpg", "")
        
        rsImageDatas.MoveNext
    Wend
    
    '���±���ͼ��
    strSql = "Zl_Ӱ����_���±���ͼ(" & mlngAdviceID & ",'" & strReportImageIds & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
ErrHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo ErrHandle
'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strNewDirectory
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strNewDirectory = App.Path & "\TmpImage\" & Format(zlDatabase.Currentdate, "YYYYMMDD")
    
    If Not DirExists(strNewDirectory) Then MkDir strNewDirectory
    If Not DirExists(strNewDirectory & "\" & strNewStudyUID) Then MkDir strNewDirectory & "\" & strNewStudyUID
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strTmpFile = App.Path & "\TmpImage\" & Nvl(rsImageDatas!ͼ��UID)
        
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        
        strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
                
        If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            strCurFtpIp = Nvl(rsImageDatas!host)
            strCurFtpUser = Nvl(rsImageDatas!FtpUser)
            strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
            
            Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
        End If
        
'       Ϊ������������ͼ��������ش���ͼ���ļ������ý���ɾ��
        
        If FileExists(strTmpFile) Then Call Kill(strTmpFile)
        If FileExists(strTmpFile & ".jpg") Then Call Kill(strTmpFile & ".jpg")

        '�ƶ��ļ����µ�λ��
        Call MoveFile(App.Path & "\TmpImage\" & Nvl(rsImageDatas!Url) & "\" & Nvl(rsImageDatas!ͼ��UID), _
            strNewDirectory & "\" & strNewStudyUID & "\" & Nvl(rsImageDatas!ͼ��UID))
        
        Call objSrcFtp.FuncDelFile(strVirtualPath, strImageUID)
        
        'ɾ���յ�ftpĿ¼
        Call objSrcFtp.FuncFtpDelDir(Replace(strVirtualPath, strImageUID, ""), strImageUID)
                
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
Exit Sub
ErrHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

'����ͼ����ƶ�
Private Sub CancelImageMove(ByVal strFTPIP As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo ErrHandle

    Call objFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
ErrHandle:
    objFtp.FuncFtpDisConnect
End Sub

Public Function MoveImageToStudy(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String, _
    ByVal strFTPIP As String, ByVal strFtpUrl As String, ByVal strFtpVirtualPath As String, _
    ByVal strFTPUser As String, ByVal strFTPPwd As String, ByRef objMoveList As Collection) As Boolean
'------------------------------------------------
'���ܣ���ѡ���ļ��ͼ���ƶ���ftp��ָ���ļ����
'���أ�True--�ɹ���False��ʧ��
'------------------------------------------------
    Dim objSrcFtp As New clsFtp
    Dim objDestFtp As New clsFtp
    Dim strVirtualPath As String
    Dim strDestPath As String
    Dim strTmpFile As String
    Dim aFiles() As String
    Dim i As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim lngResult As Long       '��¼FTP�����Ľ��
    Dim strImageUID As String
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String
    Dim strFileList As String
    Dim blnIsMove As Boolean
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrHandle
    
    blnIsMove = False
    MoveImageToStudy = False
    If rsImageDatas.RecordCount <= 0 Then Exit Function

    '����Ŀ��Ftp
    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
    
    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""
    strVirtualPath = ""
    strFileList = ""
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        strImageUID = Nvl(rsImageDatas!ͼ��UID)
        
        'ȡ������
        If mlngReleationType = 1 Then
            strSql = "Select D.���UID From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ����ʱ���� D " & _
                     "Where C.ҽ��ID=[1] And A.ͼ��UID=[2] And A.����UID=B.����UID And B.���UID=C.���UID And A.����UID = D.����UID"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "����Ӱ��", mlngAdviceID, strImageUID)
    
            If rsTemp.RecordCount > 0 Then strNewStudyUID = Nvl(rsTemp!���uid)
            
            Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        End If
        
        If strVirtualPath <> Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url) Then
            strVirtualPath = Nvl(rsImageDatas!Root) & Nvl(rsImageDatas!Url)
            strFileList = ""
        End If
        
        '���ƶ����ļ�������ͬ��ftp��ַʱ����ʹ�����غ����ϴ��ķ�ʽת���ļ�
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
        
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
            
                strCurFtpIp = Nvl(rsImageDatas!host)
                strCurFtpUser = Nvl(rsImageDatas!FtpUser)
                strCurFtpPwd = Nvl(rsImageDatas!FtpPwd)
                
                Call objSrcFtp.FuncFtpConnect(strCurFtpIp, strCurFtpUser, strCurFtpPwd)
            End If
        
            strTmpFile = App.Path & "\TmpImage\" & strImageUID
            
            If strFileList = "" Then
                strFileList = objSrcFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objSrcFtp.FuncDownloadFile(strVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "���ع���ͼ��ʧ�ܡ� [ͼ��UID:" & strImageUID & " �ļ�����Ŀ¼:" & strVirtualPath & " ����·��:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
        
                lngResult = objDestFtp.FuncUploadFile(strFtpVirtualPath, strTmpFile, strImageUID)
                If lngResult <> 0 Then
                    objSrcFtp.FuncFtpDisConnect
                    objDestFtp.FuncFtpDisConnect
        
                    Call err.Raise(-1, "MoveImageToStudy", "�ϴ�����ͼ��ʧ�ܡ� [ͼ��UID:" & strImageUID & " �ϴ�����Ŀ¼:" & strFtpVirtualPath & " ����·��:" & strTmpFile & "]", err.HelpFile, err.HelpContext)
                    Exit Function
                End If
                
                blnIsMove = True
            End If
        Else
            If strFileList = "" Then
                strFileList = objDestFtp.FuncDirFiles(strVirtualPath)
            End If
            
            '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
            If InStr(strFileList, strImageUID) > 0 Then
                lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                If lngResult <> 0 Then
                    '����ļ��ƶ�ʧ�ܣ���˿���������һ��
                    Call objDestFtp.FuncFtpDisConnect
                    Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                    
                    lngResult = objDestFtp.FuncReNameFile(strVirtualPath & "/" & strImageUID, strFtpVirtualPath & "/" & strImageUID)
                    
                    If lngResult <> 0 Then
                        If mlngReleationType = 1 Then Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
                        
                        objSrcFtp.FuncFtpDisConnect
                        objDestFtp.FuncFtpDisConnect
                        
                        Call err.Raise(-1, "MoveImageToStudy", "��Ftp���ƶ��ļ�ʱʧ�ܡ� [ͼ��UID:" & strImageUID & " ԭ����Ŀ¼:" & strVirtualPath & " ������Ŀ¼:" & strFtpVirtualPath & "]", err.HelpFile, err.HelpContext)
                        Exit Function
                    End If
                End If
                
                blnIsMove = True
                
                '��¼�Ѿ����ƶ������ļ����Ա��ڴ�������ʧ�ܵ�ʱ�򣬻��ɶ��ƶ���ͼ����лָ�
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strVirtualPath & "/" & strImageUID & ">" & strFtpVirtualPath & "/" & strImageUID)
                End If
            End If
        End If
        
        If Nvl(rsImageDatas!host) <> strFTPIP Or Nvl(rsImageDatas!Root) <> strFtpUrl Then
            Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 0)
        Else
            Call MoveReportImage(strTmpFile, strImageUID, objSrcFtp, objDestFtp, strVirtualPath, strFtpVirtualPath, objMoveList, 1)
            End If
        
        rsImageDatas.MoveNext
        
        If Not rsImageDatas.EOF Then
            If strCurFtpIp <> Nvl(rsImageDatas!host) Or strCurFtpUser <> Nvl(rsImageDatas!FtpUser) Or strCurFtpPwd <> Nvl(rsImageDatas!FtpPwd) Then
                Call objSrcFtp.FuncFtpDisConnect
            End If
        End If
    Wend
    
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect
    
    '���һ���ļ���û�б��ƶ�����ֱ���˳�
    If Not blnIsMove Then Exit Function
    
    MoveImageToStudy = True
    
    Exit Function
ErrHandle:
    objSrcFtp.FuncFtpDisConnect
    objDestFtp.FuncFtpDisConnect

    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo ErrHandle
'�ƶ�����ͼ
    Dim dcmImages As New DicomImages
    Dim dcmImg As New DicomImage
    Dim lngResult As Long
    
    If lngWay = 0 Then
        Call objSrcFtp.FuncDelFile(strSourceVirtualPath, strImgUid & ".jpg")
        
        '��������д��ڴ�Դftp�����ص�dicomͼ����ͼ��ת����jpg�������浽Ŀ��ftp�豸��
        If FileExists(strDicomFile) Then
            Call dcmImages.Clear
            Set dcmImg = dcmImages.ReadFile(strDicomFile)
    
            Call dcmImg.FileExport(strDicomFile & ".jpg", "JPG")
            Call objDestFtp.FuncUploadFile(strDestVirtualPath, strDicomFile & ".jpg", strImgUid & ".jpg")
            
            If FileExists(strDicomFile & ".jpg") Then Call Kill(strDicomFile & ".jpg")
        End If
    Else
        '���Դftp�豸�в����ڸ�ͼ���򲻽����ƶ�
        If objDestFtp.FuncFtpFileExists(strSourceVirtualPath, strImgUid & ".jpg") Then
            lngResult = objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
            
            If lngResult <> 0 Then
                '����ļ��ƶ�ʧ�ܣ���˿���������һ��
                Call objDestFtp.FuncFtpDisConnect
'                Call objDestFtp.FuncFtpConnect(strFTPIP, strFTPUser, strFTPPwd)
                Call objDestFtp.ResotreFtpConnect
                
                Call objDestFtp.FuncReNameFile(strSourceVirtualPath & "/" & strImgUid & ".jpg", strDestVirtualPath & "/" & strImgUid & ".jpg")
                
                '��¼�Ѿ����ƶ������ļ����Ա��ڴ�������ʧ�ܵ�ʱ�򣬻��ɶ��ƶ���ͼ����лָ�
                If Not objMoveList Is Nothing Then
                    Call objMoveList.Add(strSourceVirtualPath & "/" & strImgUid & ".jpg" & ">" & strDestVirtualPath & "/" & strImgUid & ".jpg")
                End If
            End If
        End If
    End If
Exit Sub
ErrHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub

Private Sub GetStorageDevice(ByVal lngAdviceId As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFTPIP As String, _
    ByRef strFtpUrl As String, ByRef strVirtualPath As String, _
    ByRef strFTPUser As String, ByRef strFTPPwd As String)
'��ȡ�µĴ洢�豸��Ϣ������豸�洢��Ϣ�����ڣ�����Ҫ��������
'�����ȡ����������ʹ��strNewStudyUID�����ܴ����ݿ��в��ҵ���Ӧ������
'strDeviceNum:�豸��
'strFtpIp: ftp��ַ
'strFtpUrl: ftpĿ¼
'strVirtualPath: ftp����洢·��
'strFtpUser: ftp�û���
'strFtpPwd: ftp����



    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim objDestFtp As New clsFtp
    Dim curDate As Date
    
    strFTPIP = ""
    strFtpUrl = ""
    strFTPUser = ""
    strFTPPwd = ""
    strDeviceNO = ""
    
    strSql = "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1] Union All " & _
        "Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������," & _
        "D.IP��ַ As Host," & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL " & _
        "From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And C.���UID= [1]"
        
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
        
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNewStudyUID)
    
    blnIsGetNewDevice = False
    
    If rsData.RecordCount <= 0 Then
        blnIsGetNewDevice = True
    Else
        '���ִ�е����˵����ִ��ͼ�����,��Ҫ�жϵ�ǰ���Ĵ洢�豸�Ƿ���Ч�������Ч�������µĴ洢�豸
        If Trim(rsData!��������) = "" Then
            blnIsGetNewDevice = True
        Else
            strDeviceNO = Nvl(rsData!λ��һ)
            strFTPIP = Nvl(rsData!host)
            strFtpUrl = Nvl(rsData!Root)
            strFTPUser = Nvl(rsData!FtpUser)
            strFTPPwd = Nvl(rsData!FtpPwd)
            strVirtualPath = strFtpUrl & Nvl(rsData!Url)
        End If
    End If
    
    
    If blnIsGetNewDevice Then
        '�����µļ��UID�ʹ洢�豸,���ִ�е����˵����ȡ������
        
        If mlngModule = 1290 Then
            '��ѯҽ������վ�У��������Ӧ�Ĵ洢�豸
            strSql = "select d.����ֵ " & _
                        " from ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��DICOM����� c, Ӱ��DICOM������� d " & _
                        " Where a.����ID = b.ִ�в���id And a.ִ�м� = b.ִ�м� And a.����豸 = c.�豸�� " & _
                        " and c.������='ͼ�����' and c.����ID=d.����ID and d.��������='�洢�豸' and b.ҽ��id=[1]"
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceId)
            
            If rsTemp.RecordCount <= 0 Then
                MsgBoxD Me, "δ�ҵ�ͼ��洢�豸,��ȷ�ϵ�ǰ��������豸�Ƿ���Ӱ���豸Ŀ¼�ķ���������������ͼ��洢��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            strDeviceNO = Nvl(rsTemp!����ֵ)
        Else
            '��ѯ��ҽ������վ�е�ͼ��洢�豸
            strDeviceNO = GetDeptPara(mlngCurDeptId, "�洢�豸��")
            
            If Val(strDeviceNO) <= 0 Then
                MsgBoxD Me, "δ�ҵ�ͼ��洢�豸,��ȷ����Ӱ�����̹������Ƿ�Ըÿ���������ͼ��ɼ��洢�豸��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        strSql = "Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ " & _
                    " From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=[1] and NVL(״̬,0)=1"
                    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.tag, strDeviceNO)
        
        '����洢�豸ͣ�ã���ֱ���˳�
        If rsTemp.RecordCount <= 0 Then
            MsgBoxD Me, "δ�ҵ��洢�豸,��ȷ���豸��Ϊ [" & strDeviceNO & "] ���豸�Ƿ����á�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strFtpUrl = Nvl(rsTemp("URL"))
        strFTPIP = Nvl(rsTemp("IP��ַ"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strFTPPwd = Nvl(rsTemp("FTP����"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFTPIP, strFTPUser, strFTPPwd
        On Error GoTo ErrHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '����FTPĿ¼
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
ErrHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrHandle
    
    Select Case control.ID
        Case conMenu_Edit_Import
            If mlngAdviceID <= 0 Or mlngStudyState < 2 Or mlngStudyState > 4 Then control.Enabled = False
    End Select
    
    Exit Sub
ErrHandle:
    
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
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
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
    
    If Button = 2 Then Call ShowPopupImage(1)
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
    
    If Button = 2 Then Call ShowPopupImage(0)
End Sub

Private Sub ShowPopupCache()

End Sub

Private Sub ShowPopupImage(ByVal intType As Integer)
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'intType:0--����ͼ��1--����ͼ��2--����ͼ
'------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    If intType <> 2 Then
        If mblnUseActiveVideo Then
            If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
        Else
            '�������ͼû��ͼ�����ֹ�Ҽ�����
            If Me.dcmMiniImageC.Images.Count < 1 Then Exit Sub
        End If
    End If
    
    '����Ҽ������˵�
    Set cbrToolBar = cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        If intType = 0 Then
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelImage, "ɾ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveUp, "ǰ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_MoveDown, "����")
            Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_SelMiniImage, "��ȡ����ͼ")
        ElseIf intType = 1 Then
            Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "ͼ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "��ҳ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportAll, "����...")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "��ҳ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "����...")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_DeleteImage, "ɾ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_RefreshImg, "ˢ��")
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

Private Sub picMiniCache_Resize()
On Error Resume Next
    cboCache.Left = 0
    cboCache.Top = 0
    cboCache.Width = picMiniCache.Width
    
    ucMiniCache.Left = 0
    ucMiniCache.Top = cboCache.Top + cboCache.Height
    ucMiniCache.Width = picMiniCache.ScaleWidth
    ucMiniCache.Height = picMiniCache.ScaleHeight - ucMiniCache.Top
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
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, pane4 As Pane
    
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
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Or mlngModule = G_LNG_PATHSTATION_MODULE Then
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
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Set pane4 = dkpMain.CreatePane(4, 0, mlngCY3, DockBottomOf, Nothing)
        pane4.Title = "����ͼ"
        pane4.Handle = picMiniCache.hWnd
        pane4.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        pane4.AttachTo Pane3
    Else
        picMiniCache.Visible = False
    End If
    
    mblnIsInitFace = True
End Sub

Private Function LoadMiniCache(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg) As Boolean
    Dim i As Integer
    Dim strQueryValue As String
    Dim strSql As String
    
    strSql = "select A.����,A.����,A.�Ա�,A.����,A.�������� As �������,A.���UID,B.����UID,B.���к� " & _
            "from Ӱ����ʱ��¼ A, Ӱ����ʱ���� B where A.���uid = B.���uid And A.�������� Between Sysdate-7 And Sysdate " & _
            "order by �������� desc"
            
    Set mrsImageCache = zlDatabase.OpenSQLRecord(strSql, "")

    cboCache.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryValue = Nvl(mrsImageCache!����UID)
        
        cboCache.AddItem "������" & Nvl(mrsImageCache!����) & "  ���ţ�" & Nvl(mrsImageCache!����) & "  ����" & Nvl(mrsImageCache!���к�)
        If mstrTmpQueryValue = Nvl(mrsImageCache!����UID) And lngEventType <> vetAfterUpdateImg Then cboCache.ListIndex = i
        
        mrsImageCache.MoveNext
    Next
    
    ucMiniCache.ImgViewer.Images.Clear
    ucMiniCache.ShowCheckBox = 1
    
    If mrsImageCache.RecordCount > 0 Or lngEventType = vetAfterUpdateImg Then
        If cboCache.ListIndex < 0 And mrsImageCache.RecordCount > 0 Then
            cboCache.ListIndex = 0
            mstrTmpQueryValue = strQueryValue
        Else
            cboCache_Click
        End If
    End If
End Function

Private Function LoadMiniImages() As Boolean
    Dim lngMsgHwnd As Long
    
    
    If mblnUseActiveVideo Then
'        lngMsgHwnd = mobjStudyImage.hWnd
'
'        Call mobjStudyImage.RefreshImages(mlngAdviceID, mlngAdviceID, mblnMoved, True)
        ucMiniImageViewer.ShowCheckBox = 1
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
                    dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                Else
                    dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                End If
            Else
                dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
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
                        dcmMark.Images(1).tag = cTable.Pictures(i).ID
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
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
                        Else
                            dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count).tag = cTable.Pictures(i).PicName
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
        picW = Val(Split(dcmReportImage(i).tag, "|")(0))
        picH = Val(Split(dcmReportImage(i).tag, "|")(1))
        
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
                        dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
                    Else
                        dcmReportImage(iRImageCount).tag = (cTable.Width - Val(Split(cTable.ExtendTag, "|")(1)) - 30) & "|" & cTable.Height
                    End If
                Else
                    dcmReportImage(iRImageCount).tag = cTable.Width & "|" & cTable.Height
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
                            dcmMark.Images(1).tag = cTable.Pictures(i).ID
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

Private Sub ucMiniCache_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 1 Then ucMiniCache.ImgChecked(ucMiniCache.SelectIndex) = Not ucMiniCache.ImgChecked(ucMiniCache.SelectIndex)
End Sub

Private Sub ucMiniCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
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

Private Sub ucMiniImageViewer_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 1 Then ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex) = Not ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex)
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
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
        '��ȡ��̨�ɼ���ͼ��
        If mblnUseAfterCapture And mlngModule <> 1290 Then
            Call LoadMiniCache
        End If
        
        '��ȡ����ʾ��ǰ��ѡ����ͼ��
        Call LoadMiniImages
        '���ݱ��浥��ʽ�����߱������ݸ�ʽ����ȡ���ͼ�ͱ���ͼ
        Call LoadReportImages
        '��Ǳ���ˢ���Ѿ�����ͼ��
        blnLoadImages = True
    End If
End Sub
