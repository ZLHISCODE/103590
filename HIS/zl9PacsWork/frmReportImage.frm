VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#8.0#0"; "zl9PacsControl.ocx"
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
      TabIndex        =   15
      Top             =   2760
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPimg 
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2381
         _ExtentY        =   476
         _Version        =   393216
         Format          =   130023425
         CurrentDate     =   42674
      End
      Begin VB.ComboBox cboCache 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin zl9PacsControl.ucImagePreview ucMiniCache 
         Height          =   1215
         Left            =   240
         TabIndex        =   16
         Top             =   2400
         Width           =   3780
         _ExtentX        =   6694
         _ExtentY        =   2143
         BackColor       =   8421504
         ShowCheckbox    =   -1  'True
      End
   End
   Begin VB.PictureBox picMiniViewer 
      Height          =   1365
      Left            =   240
      ScaleHeight     =   1305
      ScaleWidth      =   3615
      TabIndex        =   13
      Top             =   5280
      Width           =   3675
      Begin zl9PacsControl.ucImagePreview ucMiniImageViewer 
         Height          =   975
         Left            =   45
         TabIndex        =   14
         Top             =   120
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1746
         BackColor       =   8421504
      End
   End
   Begin VB.PictureBox picMenu 
      Height          =   540
      Left            =   2100
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   12
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
         TabIndex        =   4
         Top             =   1300
         Width           =   2040
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":032E
            Height          =   510
            Index           =   1
            Left            =   490
            Picture         =   "frmReportImage.frx":0F70
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   510
            Width           =   510
         End
         Begin VB.CheckBox chkMark 
            DownPicture     =   "frmReportImage.frx":653E
            Height          =   510
            Index           =   3
            Left            =   1510
            Picture         =   "frmReportImage.frx":7180
            Style           =   1  'Graphical
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            ToolTipText     =   "�Զ����"
            Top             =   0
            Value           =   1  'Checked
            Width           =   510
         End
      End
      Begin DicomObjects.DicomViewer dcmMark 
         Height          =   975
         Left            =   120
         TabIndex        =   3
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

Private mdate As Date '���ڰ�ʱ����˺�̨ͼ
Private mintTagMaxTag As Integer '��ʶX�е����X�������ж��Ƿ���²˵�����Ϣ��
Private mintTagNow As Integer '��ǰ��ʶ
Private mintTagMax As Integer '����ʶ
Private mblDel As Boolean '�Ƿ�����ɾ��������ͬ��ˢ�²ɼ�ģ�����

Public mblnSingleWindow As Boolean     '�Ƿ�ʹ�ö���������ʾ����༭����True-����������ʾ��False-Ƕ��ʽ��ʾ
Public mlngAdviceID As Long    'ҽ��ID
Private mintEditType As Integer '����״̬ 0 ������1��д��2 �޶�
Private mlngReportID As Long    '��������ID
Private mlngFileID As Long      '���浥��ʽID
Private mlngShowBigImg As Long          '�Ƿ���ʾ��ͼ,0-����ʾ��1-����ƶ�ʱ��ʾ��2-��굥����ʾ��������
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
Private mlngCY1 As Long                 '���ͼ�ĸ߶�
Private mlngMarkW As Long               '���ͼ�Ŀ��
Private mlngCY2 As Long                 '����ͼ�ĸ߶�
Private mlngRptImgW As Long             '����ͼ�Ŀ��
Private mlngCY3 As Long                 '����ͼͼ�ĸ߶�

Public pMarkModified As Boolean        '���ͼ�ı���иĶ�
Public pImageModified As Boolean       '��¼����ͼ���Ƿ��޸ģ����û���޸ģ��򱣴汨���ʱ���ٱ���ͼ��
Public pobjMarks As cPicMarks          '��ǰ���ͼ�ı�ע����
Public pMarkImageID As Double            '��ǰ���ͼ�����ݿ�����Ӳ������ݡ����е�ID
Public pTableID As String              '��ǰͼ�����ڱ���ID�����á�;���ָ���Ӱ���ܷ񱣴汨��ͼ������108069�������������0���˵�����


Private mintShowMarkImage As Integer   '�Ƿ���ʾ���ͼ   0-���ر��ͼ  1-��ʾ���ͼ
Private mblnIsInitFace As Boolean        '�Ƿ��Ѿ����ش���
Private mobjImgCTables() As cEPRTable

Private blnLoadImages As Boolean        '��¼����ˢ���Ƿ������ͼ��


Private mdcmGlobal As New DicomGlobal    '����UIDRoot=1

Private mrsImageCache As New ADODB.Recordset
Private mdcmUID As New DicomGlobal
Private mlngReleationType As Integer    '1--������2--����
Private mlngCurDeptId As Long
Private mlngStudyState As Long
Private mstrTmpQueryPath As String
Private mblnUseAfterCapture As Boolean
Private mblnTmpUseAfterCapture As Boolean
Private mstrAfterImgPath As String
Private mobjFile As New FileSystemObject
Private mstrInstance As String
Private mblnImageShield As Boolean     '�Ƿ����δ�ͼ

Private WithEvents mobjImageProcess As zl9PacsControl.clsImageProcess
Attribute mobjImageProcess.VB_VarHelpID = -1
Private mobjPacsCapture As Object

Public Event OnImageCountChanged(ByVal intType As Integer, ByVal isNeedRefreshTitle As Boolean) 'ͼ�������ı� intType: 0:���͵����   1:ɾ��
Public Event AfterReleationImage(ByVal lngReleationType As Long)
Public Event AfterShowBigImage()


Private Enum MarkType
    �Զ���� = 0: ���1: ���2: ���3: ���4: ���5: ���6
End Enum

Property Get ImageCount() As Long
    ImageCount = ucMiniImageViewer.CurImageCount
End Property

Property Get ReportImageCount() As Long
    Dim i As Long
    Dim lngResult As Long
    
    lngResult = 0
    For i = 0 To dcmReportImage.Count - 1
        lngResult = lngResult + dcmReportImage(i).Images.Count
    Next i
    
    ReportImageCount = lngResult
End Property

Property Get dcmImages() As Object
    Set dcmImages = ucMiniImageViewer.ImgViewer.Images
End Property

Public Sub MovePage(ByVal lngPageType As TMoveType)
'�ƶ�����ͼҳ��
    ucMiniImageViewer.MovePage (lngPageType)
End Sub

Public Sub RefreshAfterImage()
    LoadMiniCache
End Sub

Public Sub zlRefresh(ByVal lngAdviceID As Long, FileID As Long, _
        ReportID As Long, blnSingleWindow As Boolean, lngShowBigImg As Long, _
        intImageDblClick As Integer, blnEditable As Boolean, _
        ByVal blnMoved As Boolean, ByVal intMinImageCount As Integer, blnFormIsSelected As Boolean, _
        ByVal lngModule As Long, ByVal lngCurDeptId As Long, ByVal lngStudyState As Long, _
        ByVal blnIsSaveRefresh As Boolean)
    Dim i As Integer
    Dim intShowMarkImage As Integer
    
        Call GetNowTag(True)
        
    mlngCurDeptId = lngCurDeptId
    mlngStudyState = lngStudyState
    mlngAdviceID = lngAdviceID
    mlngFileID = FileID
    mlngReportID = ReportID
    mlngShowBigImg = lngShowBigImg
    mintImageDblClick = intImageDblClick
    mblnEditable = blnEditable
    mblnMoved = blnMoved
    mintShowPhotoNumber = intMinImageCount
    mlngModule = lngModule
    mblnSingleWindow = blnSingleWindow
    mstrAfterImgPath = IIf(Len(App.Path) > 3, App.Path & "\TmpAfterImage\", App.Path & "\TmpAfterImage\")
    
    Call InitCTables
    
    intShowMarkImage = DecideMarkImagesVisible    '�жϱ��ͼ�Ƿ�ɼ�
    
    If mlngModule = 1291 Then
        mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "���ú�̨�ɼ�", 1, True)) = 1
    Else
        mblnUseAfterCapture = False
    End If
    ucMiniImageViewer.BigImageWay = lngShowBigImg
    ucMiniImageViewer.PreViewTime = Val(GetDeptPara(mlngCurDeptId, "�ƶ�Ԥ����ʱ", 0))
    ucMiniImageViewer.ShowPopup = False
    ucMiniImageViewer.ImgLoadType = IIf(GetServiceStatus = SERVICE_RUNNING, FileLoadType.Service, FileLoadType.Normal)
    
    Call GetLocalPar
    ucMiniImageViewer.ImageShield = mblnImageShield
    
    'ֻ���� ����ͼ�� �ֶ��еı���ͼ�����ֶ�����Ϊ�գ��ټ������б���ͼ
    ucMiniImageViewer.OnlyLoadReportImage = True
    
    
    '�ж������ �������� ���� û�м��ع����� ���� ���ͼ״̬�Ѿ��ı䣬�����¼��س�ʼ������
    If mblnSingleWindow Or (Not mblnIsInitFace Or (mintShowMarkImage <> intShowMarkImage)) Or (mblnTmpUseAfterCapture <> mblnUseAfterCapture) Then
        mintShowMarkImage = intShowMarkImage
'        Call InitLoaclParas     '��ȡ��������
        Call InitFaceScheme     '��ʼ���������
    End If
    
     mblnTmpUseAfterCapture = mblnUseAfterCapture
    
    '���³�ʼ���ڲ�����
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
    If blnFormIsSelected = True Then ' And Me.Visible
        '������Ҫ����ͼ��
        If Not blnIsSaveRefresh Then    '�����ǩ�����߱��汨�棬������ͼ��
            Call LoadImages
        Else
            blnLoadImages = True
        End If
    Else
        Call ClearReportImages
    End If
    
    
    '���ý���ؼ��Ƿ���Ա༭
    picMark.Enabled = mblnEditable
    picReportImage.Enabled = mblnEditable
    picMiniViewer.Enabled = mblnEditable
End Sub

Private Sub ClearReportImages()
    Dim i As Integer
    
    pTableID = ""
    '��ʼ����������
    For i = 1 To dcmReportImage.Count - 1
        Unload dcmReportImage(i)
    Next i
    dcmMark.Images.Clear
End Sub

Public Sub RefPacsPic(Optional ByVal lngEventType As TVideoEventType = vetUpdateImg)
    '��ȡ����ʾ��ǰ��ѡ����ͼ��
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Call LoadMiniCache
    End If
    
    If lngEventType <> vetAfterUpdateImg Then Call LoadMiniImages
End Sub

Private Sub cboCache_Click()
    Dim strQueryPath As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    
    ucMiniCache.ClearCurrentPageImage
    
    Set rsTmp = mrsImageCache
    rsTmp.Filter = ""
        
    If mrsImageCache.RecordCount <= 0 Then Exit Sub
    
    mrsImageCache.MoveFirst
    Set rsTmp = mrsImageCache

    rsTmp.Filter = "����='" & Trim(Mid(cboCache.Text, 1, 5)) & "'"

    If rsTmp.RecordCount < 1 Then Exit Sub
    strQueryPath = Nvl(rsTmp!·��)

    If strQueryPath = "" Then Exit Sub
    
    Call ucMiniCache.RefreshImage(slLocal, strQueryPath, mblnMoved)
    mstrTmpQueryPath = strQueryPath
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
    err.Clear
End Sub

Private Sub cboCache_DropDown()
On Error GoTo errHandle
    Call SendMessage(cboCache.hWnd, &H160, 500, 0)
errHandle:
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    
    Select Case Control.ID
        Case comMenu_Cap_Process 'ͼ����
            If Control.Caption = "ͼ���ע" Then
                Call OpenLabelProcessWnd
            Else
                Call OpenImageProcessWind
            End If
        Case conMenu_Cap_DevSet
            If mblnUseAfterCapture And mlngModule <> 1290 Then Call ucMiniCache.ShowPageConfig
            Call ucMiniImageViewer.ShowPageConfig
        Case conMenu_PacsReport_DelImage    'ɾ��ͼ��
            If dcmReportImage(mSelViewerIndex).Images.Count >= mselReportImgIndex And mselReportImgIndex <> 0 Then
                '�Ƴ�ͼ��֮ǰ����ɾ�����
                Call DelImgRptTag(dcmReportImage(mSelViewerIndex).Images(mselReportImgIndex))
                
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
            
            Set resImages = frmSelectRepImage.ShowMe(Me, mlngAdviceID)
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
            
        Case conMenu_Cap_SendToAdvice        '���͵����
            mlngReleationType = 2
            Call ReleationImage
        
        Case conMenu_Cap_SendToAfter     '���͵���̨
            mlngReleationType = 1
            Call ReleationImage
        
        Case conMenu_Manage_DeleteImage 'ɾ����ʱͼ��
            mlngReleationType = 2
                        mblDel = True
            Call DelTempImage
        
        Case conMenu_Manage_RefreshImg  'ˢ�»���
            Call LoadMiniCache
        
        Case conMenu_Cap_ImageShield    '���δ�ͼ
            Control.Checked = Not Control.Checked
            
            mblnImageShield = Control.Checked
            ucMiniImageViewer.ImageShield = mblnImageShield
            Call SaveLocalPar
    End Select
End Sub

Private Sub CheckSendOnImageCountChangedChanged(ByVal intType As Integer)
'intType 0:���͵����  1:ɾ��ͼ��
'isNeedRefreshTitle:�Ƿ���Ҫ��������

    If (InStr(cboCache.Text, "��ʶ" & zlStr.Lpad((mintTagMaxTag), 3, "0")) > 0) Then
        RaiseEvent OnImageCountChanged(intType, True)
    Else
        RaiseEvent OnImageCountChanged(intType, False)
    End If
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
    
    Call DelTempImages(rsImageDatas)
End Sub

Private Function DelTempImages(rsImageDatas As ADODB.Recordset) As Boolean
'ɾ�����ػ����е��ļ����ڽ�����ɾ��ucpre�ؼ���ѡ��ͼ��
    Dim blfinished As Boolean
    Dim i As Long
    Dim curTime As Date
    Dim intTMP As Integer
        
On Error GoTo errHandle
    If rsImageDatas.RecordCount <= 0 Then Exit Function
    
    Call rsImageDatas.MoveFirst
    
    While Not rsImageDatas.EOF
        If mobjFile.FileExists(rsImageDatas!·��) Then Call mobjFile.DeleteFile(rsImageDatas!·��)
        
        rsImageDatas.MoveNext
    Wend
        
    'ɾ������ͼ��
    blfinished = False
    For i = ucMiniCache.CurImageCount To 1 Step -1
        If ucMiniCache.ImgChecked(i) Then
            Call ucMiniCache.DeleteImage(i)
            blfinished = True
        End If
    Next

    If blfinished = False Then
        Call ucMiniCache.DeleteImage(ucMiniCache.SelectIndex)
    End If

    'ͬʱ��Ҫɾ��cbo��Ŀ
    Call ClearEmptyFolder(False)
    If ucMiniCache.CurImageCount = 0 Then
        curTime = zlDatabase.Currentdate
        '�ǵ��첢��ѡ�е��ǵ�ǰ��ʶ���Ͳ�������ղ���
        If Not ((Format(DTPimg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCache.Text, "��ʶ" & zlStr.Lpad((mintTagMaxTag), 3, "0")) > 0)) Then
            intTMP = cboCache.ListIndex
            Call cboCache.RemoveItem(cboCache.ListIndex)
            If cboCache.ListCount > intTMP - 1 Then
                cboCache.ListIndex = intTMP - 1
            Else
                cboCache.ListIndex = 0
            End If
        End If
    End If
    
    DelTempImages = True
    
    Exit Function
errHandle:
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Private Function ReleationImage() As Boolean
    Dim strHint As String
    Dim rsImageDatas As ADODB.Recordset
    Dim strTmpFile As String
    Dim i As Integer
    
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
        If MsgBoxD(Me, "�Ƿ�ȷ�Ͻ���ѡͼ���͵���̨��", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
    End If

    If mlngReleationType = 2 Then '����2��ʾ����ͼ��
        ReleationImage = StartReleation(mlngAdviceID, rsImageDatas)
        Call ClearEmptyFolder(False)
    Else
        ReleationImage = CancelReleation(mlngAdviceID, rsImageDatas)
    End If
        
    '�����������򣬷�ֹ����2������BUG
    For i = 1 To ucMiniImageViewer.CurImageCount
        ucMiniImageViewer.ImgViewer.Images(i).BorderColour = vbWhite
    Next
    
    If ReleationImage Then RefPacsPic
    RaiseEvent AfterReleationImage(mlngReleationType)
End Function

'ȡ�ù�����ʾ��Ϣ
Private Function GetReleationHintInfo(lngAdviceID As Long, rsReleationImage As ADODB.Recordset) As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strResult As String
    Dim strStudyInf As String
    
    GetReleationHintInfo = ""
    
    strSql = "select ����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    GetReleationHintInfo = "�Ƿ�ȷ�Ͻ�ѡ���ͼ���͵�[" & Nvl(rsTemp!����) & "(" & Nvl(rsTemp!����) & ") " & Nvl(rsTemp!�Ա�) & " " & Nvl(rsTemp!����) & "]�ļ���У�"
End Function

Private Function GetReleationImageIds() As ADODB.Recordset
'��ѯ��������Ҫȡ��������ͼ��ID
    Dim i As Long, j As Long
    Dim strSql As String
    Dim strValues(0 To 80) As String
    Dim strValue As String
    Dim strUninTable As String
    Dim strFilter As String
    Dim strTmpFile As String
    Dim rsImageDatas As ADODB.Recordset

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
        '������ͼ��û�б�ѡ�еĺ�㣬���к���ͼ����Ϊѡ��
                
        If Not ucMiniImageViewer.SelectImage Is Nothing And strValue = "" Then strValue = strValue & "," & ucMiniImageViewer.SelectImage.InstanceUID
                
    Else
        Set rsImageDatas = New ADODB.Recordset
        rsImageDatas.Fields.Append "����UID", adVarChar, 4000
        rsImageDatas.Fields.Append "���UID", adVarChar, 4000
        rsImageDatas.Fields.Append "·��", adVarChar, 4000
        rsImageDatas.Open
            
        For i = 1 To ucMiniCache.CurImageCount
            If ucMiniCache.ImgChecked(i) Then
                strTmpFile = ucMiniCache.ImgViewer.Images(i).tag.FilePath
                rsImageDatas.AddNew
                rsImageDatas!����UID = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                rsImageDatas!���UID = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                rsImageDatas!·�� = strTmpFile
                rsImageDatas.Update
            End If
        Next
        
        'û��ͼ���ں��״̬����ѡ���к���
        If rsImageDatas.RecordCount = 0 Then
            If ucMiniCache.CurImageCount > 0 Then
                If Not ucMiniCache.SelectImage Is Nothing Then
                    strTmpFile = ucMiniCache.SelectImage.tag.FilePath
                    rsImageDatas.AddNew
                    rsImageDatas!����UID = mobjFile.GetFolder(mobjFile.GetParentFolderName(strTmpFile)).Name
                    rsImageDatas!���UID = GetStudyUIDFromFolderName(mobjFile.GetFolder(mobjFile.GetParentFolderName(mobjFile.GetParentFolderName(strTmpFile))).Name)
                    rsImageDatas!·�� = strTmpFile
                    rsImageDatas.Update
                End If
            End If
        End If
                
        If rsImageDatas.RecordCount > 0 Then rsImageDatas.MoveFirst
                
        Set GetReleationImageIds = rsImageDatas
        Exit Function
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
    
    If strFilter <> "" Then strUninTable = strUninTable & " Union All Select ͼ��UID from [Ӱ��ͼ��] where  ( " & Mid(strFilter, 4) & ")"
    
    strSql = "Select /*+ RULE*/ D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd, Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ) as �豸��," & _
        "D.IP��ַ As Host,B.����UID,B.���UID,C.Ӱ�����, " & _
        "'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID As URL,A.ͼ��UID, c.����,c.�Ա�,c.����,c.���� " & _
        "From Ӱ����ͼ�� A, Ӱ�������� B, Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,(" & Replace(strUninTable, "[Ӱ��ͼ��]", "Ӱ����ͼ��") & ") E " & _
        "Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+)" & _
        "And A.����UID=B.����UID and B.���UID=C.���UID and A.ͼ��UID = E.ͼ��UID "
        
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

Private Function StartReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��ʼ����
On Error GoTo errHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim blnBeginTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim objFtp As New clsFtp
    Dim objMoveList As New Collection
    Dim strNewDeviceNo As String, strNewFtpIp As String, strNewFtpUrl As String, strNewFtpVirtualPath As String, strNewFtpUser As String, strNewFtpPwd As String
    
    blnBeginTrans = False
    StartReleation = False
    
    curDate = zlDatabase.Currentdate
    
    strSql = "select ���UID,�������� from Ӱ�����¼ where ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    
    If rsTmp.RecordCount <= 0 Then
        Call MsgBoxD(Me, "�Ҳ����������ļ����Ϣ��", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Trim(Nvl(rsTmp!���UID)) = "" Or Trim(Nvl(rsTmp!��������)) = "" Then
        '��δ�ɼ�ͼ����Ҫ�����µļ��UID
        strNewStudyUID = CreateStudyUid(rsImageDatas!���UID)
        
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
        strNewStudyUID = Nvl(rsTmp!���UID)
        
        Call GetStorageDevice(mlngAdviceID, strNewStudyUID, strNewDeviceNo, strNewFtpIp, strNewFtpUrl, strNewFtpVirtualPath, strNewFtpUser, strNewFtpPwd)
        If Trim(strNewFtpIp) = "" Then
            Call MsgBoxD(Me, "����ȡ����Ч�Ĵ洢�豸������洢�豸���á�", vbInformation, Me.Caption)
            Exit Function
        End If
    End If
    
    '����FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgBoxD(Me, "FTP����ʧ�ܣ������������á�", vbInformation, Me.Caption)
        Exit Function
    End If
    
    '�ƶ�ͼ���ļ�
    If Not MoveImageToStudy(objFtp, rsImageDatas, strNewFtpVirtualPath, objMoveList) Then Exit Function
          
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        '�����µ�����UID
        strNewSeriesUid = CreateSeriesUid(rsImageDatas!����UID, strNewStudyUID)
        
        '����ͼ���������
        strSql = "Zl_Ӱ����_ͼ����(" & mlngAdviceID & ",'" & strNewStudyUID & "','" & strNewSeriesUid & "','" & mobjFile.GetFileName(Nvl(rsImageDatas!·��)) & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    '�ύ����
    Call gcnOracle.CommitTrans
    
    '˵��ȫ���ϴ��ɹ�,ɾ��������ʱͼ��
    Call DelTempImages(rsImageDatas)
    
    StartReleation = True
    
    Exit Function
errHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    
    Call CancelImageMove(strNewFtpIp, strNewFtpUser, strNewFtpPwd, objMoveList)
    Call OutputDebug("StartReleation", err)
    
    Call RaiseErr(err)  '�����׳�����
End Function

Private Function CreateSeriesUid(ByVal strSeriesUID As String, ByVal strStudyUID As String) As String
'��������UID
    Dim rsData As New ADODB.Recordset
    Dim strSql As String
    Dim strNewSeriesUid As String
    
    strNewSeriesUid = strSeriesUID 'M_STR_SERIES_UID & "." & Format(Now, "yymmddhhmmss") & "." & Fix(Rnd(1000) * 1000)
    
    strSql = "select ����UID from Ӱ�������� where ����UID = [1] And ���UID <> [2]"
              
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�", strNewSeriesUid, strStudyUID)
    
    If rsData.RecordCount > 0 Then
        '����һ���µļ��UID
        strSql = "Select Ӱ����UID���_ID.Nextval From Dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "PACSͼ�񱣴�")
        
        If Len(strNewSeriesUid) <= 55 Then
            strNewSeriesUid = strNewSeriesUid & ".A" & rsData(0)
        Else
            strNewSeriesUid = Left(strNewSeriesUid, 55) & ".A" & rsData(0)
        End If
    End If
    
    CreateSeriesUid = strNewSeriesUid
End Function

Private Function CancelReleation(ByVal lngAdviceID As Long, rsImageDatas As ADODB.Recordset) As Boolean
'��������
On Error GoTo errHandle
    Dim strSql As String
    Dim strNewStudyUID As String, strNewSeriesUid As String
    Dim curDate As Date
    Dim objFtp As New clsFtp
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
    
    '����FTP
    If objFtp.FuncFtpConnect(strNewFtpIp, strNewFtpUser, strNewFtpPwd) = 0 Then
        Call MsgBoxD(Me, "FTP����ʧ�ܣ������������á�", vbInformation, Me.Caption)
        Exit Function
    End If
    
    If Not MoveImageToAfter(objFtp, rsImageDatas, objMoveList) Then Exit Function
    
    gcnOracle.BeginTrans
    
    blnBeginTrans = True
    
    '��������
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strSql = "Zl_Ӱ����_ͼ�񵼳�(" & mlngAdviceID & ",'" & Nvl(rsImageDatas!ͼ��UID) & "')"
                                        
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        rsImageDatas.MoveNext
    Wend
    
    Call gcnOracle.CommitTrans
    
    Call ClearFtpImage(rsImageDatas, strNewStudyUID)
    
    CancelReleation = True
Exit Function
errHandle:
    If blnBeginTrans Then Call gcnOracle.RollbackTrans
    Call OutputDebug("CancelReleation", err)
    Call RaiseErr(err)
End Function

Private Sub ClearFtpImage(rsImageDatas As ADODB.Recordset, ByVal strNewStudyUID As String)
On Error GoTo errHandle
'ת��ͼ��ɹ�����ɾ����ʱͼ���ԭ��FTP��ͼ���Ŀ¼���峡�������ִ�����Բ�����
    Dim objSrcFtp As New clsFtp
    Dim strTmpFile As String
    Dim strVirtualPath As String
    Dim strImageUID As String
    Dim strCurFtpIp As String, strCurFtpUser As String, strCurFtpPwd As String

    strCurFtpIp = ""
    strCurFtpUser = ""
    strCurFtpPwd = ""

    
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
errHandle:
    Call OutputDebug("ClearFtpImage", err)
End Sub

'����ͼ����ƶ�
Private Sub CancelImageMove(ByVal strFtpIp As String, ByVal strFTPUser As String, ByVal strFTPPwd As String, objMoveList As Collection)
    Dim i As Long
    Dim objFtp As New clsFtp
    Dim strDestFile As String
    Dim strMoveFile As String
    
    If objMoveList Is Nothing Then Exit Sub
    If objMoveList.Count <= 0 Then Exit Sub
    
On Error GoTo errHandle

    Call objFtp.FuncFtpConnect(strFtpIp, strFTPUser, strFTPPwd)
    
    For i = 1 To objMoveList.Count
        strDestFile = objMoveList.Item(i)
        
        strMoveFile = Mid(strDestFile, InStr(strDestFile, ">") + 1, 255)
        strDestFile = Mid(strDestFile, 1, InStr(strDestFile, ">") - 1)
        
        Call objFtp.FuncReNameFile(strMoveFile, strDestFile)
    Next i
        
errHandle:
    objFtp.FuncFtpDisConnect
End Sub

Private Function MoveImageToStudy(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, strNewFtpVirtualPath As String, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        If objFtp.FuncUploadFile(strNewFtpVirtualPath, rsImageDatas!·��, mobjFile.GetFileName(rsImageDatas!·��)) <> 0 Then
            'ʧ�ܺ�ɾ��֮ǰ�ϴ����ļ�
            For i = 0 To objMoveList.Count - 1
                Call objFtp.FuncDelFile(strNewFtpVirtualPath, objMoveList(i))
            Next
            
            Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
            
            Exit Function
        Else
            Call objMoveList.Add(rsImageDatas!·��)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    MoveImageToStudy = True
End Function

Private Function MoveImageToAfter(objFtp As clsFtp, rsImageDatas As ADODB.Recordset, ByRef objMoveList As Collection) As Boolean
    Dim i As Long
    Dim strDestPath As String
    
    rsImageDatas.MoveFirst
    While Not rsImageDatas.EOF
        strDestPath = GetAfterImagePath(rsImageDatas!ͼ��UID, rsImageDatas!����UID, rsImageDatas!���UID, rsImageDatas!Ӱ�����)
        If mobjFile.FolderExists(strDestPath) = False Then Call MkLocalDir(strDestPath)
        
        If objFtp.FuncDownloadFile(rsImageDatas!Root & rsImageDatas!Url, strDestPath & rsImageDatas!ͼ��UID, rsImageDatas!ͼ��UID) <> 0 Then
            'ʧ�ܺ�ɾ��֮ǰ���ص��ļ�
            For i = 0 To objMoveList.Count - 1
                Call mobjFile.DeleteFile(objMoveList(i))
            Next
            
            Call MsgBoxD(Me, "ͼ���ƶ�ʧ�ܣ�����FTP�����Ƿ�������", vbInformation, Me.Caption)
            
            Exit Function
        Else
            Call objMoveList.Add(strDestPath & rsImageDatas!ͼ��UID)
        End If
        
        rsImageDatas.MoveNext
    Wend
    
    Call MsgBoxD(Me, "�ѽ�ѡ��ͼ���͵�[���" & mintTagNow & "]��", vbInformation, "��ʾ")
        
    MoveImageToAfter = True
End Function

Public Function GetAfterImagePath(ByVal strImageName As String, ByVal strSeriesUID As String, ByVal strStudyUID As String, ByVal strModality As String) As String
    Dim strTmpPath As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim curDate As Date
    Dim strDate As String
    Dim intTMP As Integer
    
    curDate = zlDatabase.Currentdate
    strDate = Format(curDate, "yyyymmdd")
    
    strTmpPath = ""
    
    If mobjFile.FolderExists(mstrAfterImgPath & "\") Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterImgPath & "\").SubFolders   'ʱ���
            If objFolder1.Name = strDate Then 'ʱ��ֻ�жϵ���

                For Each objFolder2 In mobjFile.GetFolder(objFolder1.Path).SubFolders   '����
                
                    If InStr(objFolder2.Name, "���" & mintTagNow) > 0 Then '�ж��Ƿ���������+��ǰ��ʶ��Ŀ¼�����У�ֱ��ʹ�ã�
                        
                        For Each objFolder3 In mobjFile.GetFolder(objFolder2.Path).SubFolders   '���в�
                                strTmpPath = objFolder3.Path & "\"
                                GoTo step2
                        Next
                   
                    End If
                Next
                
                Exit For '��ֹʱ����ļ��е�����
            End If
        Next
    End If
    
    If strTmpPath = "" Then
        strTmpPath = mstrAfterImgPath & "\" & Format(curDate, "yyyymmdd") & "\" & "���" & mintTagNow & "-" & strStudyUID & "\" & strSeriesUID & "\"
    End If
    
    '�ҵ�Ŀ¼��ֹͣǰ��ı�����ֱ�ӽ���step2
step2:
    GetAfterImagePath = strTmpPath
End Function

Private Sub MoveReportImage(ByVal strDicomFile As String, ByVal strImgUid As String, _
    objSrcFtp As clsFtp, objDestFtp As clsFtp, ByVal strSourceVirtualPath As String, ByVal strDestVirtualPath As String, _
    objMoveList As Collection, Optional ByVal lngWay As Long = 0)
On Error GoTo errHandle
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
errHandle:
    Call OutputDebug("MoveReportImage", err)
End Sub

Private Sub GetStorageDevice(ByVal lngAdviceID As Long, ByVal strNewStudyUID As String, _
    ByRef strDeviceNO As String, ByRef strFtpIp As String, _
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
    
    strFtpIp = ""
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
            strFtpIp = Nvl(rsData!host)
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
                        
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
            
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
        strFtpIp = Nvl(rsTemp("IP��ַ"))
        strFTPUser = Nvl(rsTemp("FTP�û���"))
        strFTPPwd = Nvl(rsTemp("FTP����"))
        
        strFtpUrl = IIf(strFtpUrl = "/", "//", strFtpUrl)
        
        objDestFtp.FuncFtpConnect strFtpIp, strFTPUser, strFTPPwd
        On Error GoTo errHandle
            curDate = zlDatabase.Currentdate
            
            strVirtualPath = strFtpUrl & Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            '����FTPĿ¼
            objDestFtp.FuncFtpMkDir strFtpUrl, Format(curDate, "YYYYMMDD") & "/" & strNewStudyUID
            
        Call objDestFtp.FuncFtpDisConnect
errHandle:
        Call objDestFtp.FuncFtpDisConnect
    End If
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case Control.ID
        Case conMenu_Cap_SendToAdvice
            If mlngAdviceID <= 0 Or mlngStudyState < 2 Then Control.Enabled = False
    End Select
    
    Exit Sub
errHandle:
    
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
        pobjMarks(pobjMarks.Count).���ɫ = glngColor(pobjMarks.Count Mod 9 + 1)
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
'--------------------------------------------------------
'��  �ܣ���ͼ�����ڣ�����������ͼ
'��������
'���أ���
'-------------------------------------------------------
    Dim i As Long
    
    If Not mSelMiniImg Is Nothing Then
        For i = 1 To mSelMiniImg.Labels.Count
            If mSelMiniImg.Labels(i).tag = "SELECT" Or mSelMiniImg.Labels(i).tag = "BORDER" Or mSelMiniImg.Labels(i).tag = "����ͼ" Then
                mSelMiniImg.Labels(i).Visible = False
            End If
        Next
    End If
        
    mintCurImgIndex = ucMiniImageViewer.SelectIndex
    
    If mobjImageProcess Is Nothing Then
        Set mobjImageProcess = New zl9PacsControl.clsImageProcess
    
    End If
    
    mobjImageProcess.ShowImageProcess mlngAdviceID, mSelMiniImg, (ucMiniImageViewer.PageNumber - 1) * ucMiniImageViewer.PageImgCount + mintCurImgIndex, Me, mblnMoved, 0
'    Call frmReportImageEdit.zlShowMe(mSelMiniImg, Me, mintCurImgIndex, mSelViewerIndex, mlngModule)

    
    If Not mSelMiniImg Is Nothing Then
        For i = 1 To mSelMiniImg.Labels.Count
            mSelMiniImg.Labels(i).Visible = True
        Next
    End If
End Sub

Private Sub OpenLabelProcessWnd()
'--------------------------------------------------------
'��  �ܣ���ͼ���ע����
'��������
'���أ���
'-------------------------------------------------------
    If dcmMark.Images.Count <> 1 Then Exit Sub
    

    If mobjImageProcess Is Nothing Then
        Set mobjImageProcess = New zl9PacsControl.clsImageProcess
        mobjImageProcess.ShowImageProcess "", dcmMark.Images(1), 1, Me, mblnMoved, , 2
    End If
'    Call frmReportImageEdit.zlShowMe(dcmMark.Images(1), Me, 0, 1, mlngModule)

End Sub

Public Sub DcmAddMarkImage(dcmImage As DicomImage)
'------------------------------------------------
'���ܣ��滻���ͼ�������ͼ�滻��ͼ��dcmImage
'������ dcmImage --- �µı��ͼ
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim lngIndex As Long
    Dim oneLabel As DicomLabel
    Dim lblsTemp As DicomLabels
    Dim twoLabel As DicomLabel
    Dim threeLabel As DicomLabel
    Dim lngNumber As Long
    Dim blnSaveLabel As Boolean     '��ע�����ֱ���е���Բ�ͱ���ɫ������Ҫ�����Ѿ����ı��д����ˡ�
    Dim byt���� As Byte
    Dim col���ɫ As OLE_COLOR
    Dim strText As String
    Dim int��䷽ʽ As Integer
    
    On Error GoTo err
    
    If dcmImage Is Nothing Then Exit Sub
    
    '��ͼ����ӵ����ͼ����
    dcmMark.Images.Clear
    dcmMark.Images.Add dcmImage
    pMarkModified = True
    
    '�ؽ���ע֮��Ĺ���
    Call subLabelCopyRebuild(dcmImage, dcmMark.Images(1))
    
    '��ͼ��ı�ע����ӵ�pobjMarks������
    'pobjMarks(lngIndex).���Ͷ��� '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ), 6-˳���ţ�7-��ͷ��PACS�����ӣ�
    '����ձ�ע����
    For i = 1 To pobjMarks.Count
        pobjMarks.Remove 1
    Next i
    pMarkModified = True
    Set lblsTemp = dcmMark.Images(1).Labels
    col���ɫ = vbYellow
    
    '�������ӱ�ע����
    For i = 1 To lblsTemp.Count
        Set oneLabel = lblsTemp(i)
        blnSaveLabel = True
        
        int��䷽ʽ = -1        '�����
        '�㼯û�У�����
        If oneLabel.LabelType = doLabelText Then
            If oneLabel.tag = m_LabelTag_Number Then    '���ֱ�ţ���¼���֣�������
                blnSaveLabel = False
            Else         '��ͨ����
                byt���� = 0     '����
            End If
            strText = oneLabel.Text
        ElseIf oneLabel.LabelType = doLabelArrow Then
            byt���� = 7    '��ͷ
        ElseIf oneLabel.LabelType = doLabelEllipse Then
            If oneLabel.tag = m_LabelTag_Circle Then    '���ֱ�ţ���Ȧ��������
                blnSaveLabel = False
            ElseIf oneLabel.tag = m_LabelTag_Back Then  '���ֱ�ţ�����ɫ����Ҫͬʱ����������ע
                byt���� = 6 '���ֱ��
                col���ɫ = oneLabel.BackColour
                int��䷽ʽ = -2    'ʵ��
                strText = oneLabel.TagObject.Text
            Else        '��ͨ��Բ
                byt���� = 5     'Բ(��Բ)
            End If
        End If
        
        If blnSaveLabel = True Then
            lngIndex = pobjMarks.Count + 1
    
            pobjMarks.Add lngIndex
            pobjMarks(lngIndex).Selected = False
            
            pobjMarks(lngIndex).���� = byt����
            pobjMarks(lngIndex).���ɫ = col���ɫ
            pobjMarks(lngIndex).��䷽ʽ = int��䷽ʽ
            pobjMarks(lngIndex).���� = strText
            '����ɫ���գ�����ɫ����
            If oneLabel.LabelType = doLabelEllipse And oneLabel.tag = m_LabelTag_Back Then
                '���Ӳ��������У�������ֱ�ţ��ػ���ע��ʱ�����ԣ�X1-7��Y1-7����Ϊ���Ͻǵ�ģ���˱����ʱ����Ҫ���ƶ�7
                pobjMarks(lngIndex).X1 = oneLabel.Left / mdblMarkZoom + 7
                pobjMarks(lngIndex).Y1 = oneLabel.Top / mdblMarkZoom + 7
            Else
                pobjMarks(lngIndex).X1 = oneLabel.Left / mdblMarkZoom
                pobjMarks(lngIndex).Y1 = oneLabel.Top / mdblMarkZoom
            End If
            If oneLabel.LabelType = doLabelText And oneLabel.Width = 0 And oneLabel.Height = 0 Then
                pobjMarks(lngIndex).X2 = (oneLabel.Left + Len(oneLabel.Text) * 10) / mdblMarkZoom
                pobjMarks(lngIndex).Y2 = (oneLabel.Top + 14) / mdblMarkZoom
            Else
                pobjMarks(lngIndex).X2 = (oneLabel.Left + oneLabel.Width) / mdblMarkZoom
                pobjMarks(lngIndex).Y2 = (oneLabel.Top + oneLabel.Height) / mdblMarkZoom
            End If
            pobjMarks(lngIndex).���� = 1
            pobjMarks(lngIndex).�߿� = 2
            Set pobjMarks(lngIndex).���� = New StdFont  '����
        End If
        
        '������ı�ע��ɾ����
        'Call lblsTemp.Remove(1)
    Next i
    dcmMark.Refresh
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Public Sub DcmAddImage(dcmImage As DicomImage, SelViewerIndex As Integer)
'�ѵ�ǰͼ����ӵ�ͼ�����
    Dim i As Integer
    
    '�����û�д���ͼ������˳�
    If dcmReportImage.Count = 1 Then Exit Sub
    
    If Not dcmImage Is Nothing Then
        For i = 1 To dcmImage.Labels.Count
            If dcmImage.Labels(i).tag = "SELECT" Or dcmImage.Labels(i).tag = "BORDER" Or dcmImage.Labels(i).tag = "����ͼ" Then
                dcmImage.Labels(i).Visible = False
            End If
        Next
        
        dcmReportImage(SelViewerIndex).Images.Add dcmImage
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).tag = dcmReportImage(SelViewerIndex).Images(dcmReportImage(SelViewerIndex).Images.Count).InstanceUID & ".jpg"
        dcmReportImage(SelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
        
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = True
        Next
    End If
End Sub

Private Sub ChangeImgRptTag(imgIndex As Long, lngActionType As Long)
    '�ı�ͼ��ı���ͼ���
    Dim dcmImage As DicomImage
    
    Set dcmImage = ucMiniImageViewer.ImgViewer.Images(imgIndex)
    '���뱨��
    If lngActionType = 1 Then
        '���Ϊ""��0����ĳ�1�����>=1�����һ
        dcmImage.tag.ReportImage = Val(dcmImage.tag.ReportImage) + 1
    Else
        '�ӱ�����ɾ��
        '��һ�����<=0����Ϊ""
        dcmImage.tag.ReportImage = Val(dcmImage.tag.ReportImage) - 1
        
        If Val(dcmImage.tag.ReportImage) = 0 Then dcmImage.tag.ReportImage = ""
    End If
    
    Call ucMiniImageViewer.DrawReportImgTag(ucMiniImageViewer.ImgViewer.Images(imgIndex))
End Sub

Private Sub DelImgRptTag(dcmImage As DicomImage)
    'ɾ������ͼ���ж��Ƿ���Ҫɾ������ͼ�еġ�����ͼ�����
    Dim strInstanceUID As String
    Dim i As Long
    
    '����ͼ��uid���ж����ͼ���Ƿ��������ͼ��
    strInstanceUID = dcmImage.InstanceUID
    For i = 1 To ucMiniImageViewer.ImgViewer.Images.Count
        If strInstanceUID = ucMiniImageViewer.ImgViewer.Images(i).InstanceUID Then
            Call ChangeImgRptTag(i, 2)
        End If
    Next i
    
End Sub

Public Sub DcmAddXWImage(dcmImage As DicomImage)
'�ѵ�ǰͼ����ӵ�ͼ�����
    Dim i As Integer
    
    If Not dcmImage Is Nothing Then
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = False
        Next
        
        dcmReportImage(mSelViewerIndex).Images.Add dcmImage
        dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).BorderColour = vbWhite
        dcmReportImage(mSelViewerIndex).Images(dcmReportImage(mSelViewerIndex).Images.Count).tag = mdcmGlobal.NewUID & ".jpg"
        dcmReportImage(mSelViewerIndex).CurrentIndex = 1
        Call picReportImage_Resize
        pImageModified = True
        
        For i = 1 To dcmImage.Labels.Count
            dcmImage.Labels(i).Visible = True
        Next
    End If
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
        If ucMiniImageViewer.CurImageCount < 1 Then Exit Sub
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
            cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_ImageShield, "���δ�ͼ")
            Call GetLocalPar
            cbrControl.Checked = mblnImageShield
            
            If mlngModule = 1291 And mblnUseAfterCapture Then Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SendToAfter, "���͵���̨")
            cbrControl.BeginGroup = True
            
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "��ҳ����")
            Set cbrControl = .Add(xtpControlButton, conMenu_Cap_SendToAdvice, "���͵����")
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
        Set cbrControl = .Add(xtpControlButton, comMenu_Cap_Process, "ͼ���ע")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub DTPImg_Change()
    On Error GoTo errH
        
    mdate = DTPimg.value
    Call LoadMiniCache
    ucMiniCache.RedrawSelf
    Call dkpMain.RedrawPanes
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub

Private Sub Form_Activate()
    '������Ҫ����ͼ��
    
    'ע����Form��Activate��Paintʱ���б������LoadImages����
    '��Ϊ���ֻ��Activate�����е���LoadImages������������ɱ���ͼ�����ڵ�һʱ����ʾ�������������һ�±���ͼ�Ż���ʾ
    '���ֻ��Paint�����е���LoadImages���������ڸ÷�����ʹ����UnLoadж�ؿؼ����飬������ɡ����ܴӸ���������ж�ء��Ĵ���
    
    Call LoadImages
    Call GetNowTag(True)
End Sub

Private Sub Form_Load()

    DTPimg.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mdate = DTPimg.value
    Call LoadMiniCache
        
    '��Ǳ���ˢ���Ѿ�����ͼ��
    blnLoadImages = True
    
    '��Ǵ����Ѿ��״μ���
    mblnIsInitFace = False
        
    mintMoustType = MarkType.�Զ����
    
    '����UIDRoot=1
    mdcmGlobal.RegString("UIDRoot") = "1"
    
    Call InitLoaclParas     '��ȡ��������
'    Call InitFaceScheme     '��ʼ���������
    
    Call RegXWAddReportImgWindow(Me.hWnd, Me)
End Sub


Private Sub InitLoaclParas()
    Dim strRegPath As String
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReportImage"
    End If
    
    ucMiniImageViewer.PageImgCount = Val(GetSetting("ZLSOFT", strRegPath, "��������ͼ����", 5))
    If mlngModule = 1291 Then
        mblnUseAfterCapture = Val(GetDeptPara(mlngCurDeptId, "���ú�̨�ɼ�", 1, True)) = 1
    Else
        mblnUseAfterCapture = False
    End If
    
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
    
    Call ClearEmptyFolder(True)
    
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
    SaveSetting "ZLSOFT", strRegPath, "CY3", picMiniCache.Height
    
    If mblnSingleWindow = True Then
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport\SingleWindow"
    Else
        strRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    End If
    SaveSetting "ZLSOFT", strRegPath, "CX3", Me.Width
    SaveSetting "ZLSOFT", strRegPath, "CY3", Me.Height
    
    Call DisRegXWAddReportImgWindow(Me.hWnd)
End Sub

Private Sub mobjImageProcess_AfterSaveStady()
    Call LoadMiniImages
End Sub

Private Sub mobjImageProcess_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    If mstrInstance = dcmImage.InstanceUID Then Exit Sub
    Select Case lngImageType
        Case 0  '���ͼ
            Call DcmAddMarkImage(dcmImage)
        Case 1  '����ͼ
            Call DcmAddImage(dcmImage, mSelViewerIndex)
        Case 2  '���ͼ
            If mobjPacsCapture Is Nothing Then
                Set mobjPacsCapture = CreateObject("zl9PacsImageCap.clsPacsCapture")
                
                Call mobjPacsCapture.zlInitModule(gcnOracle, glngSys, mlngModule, gstrPrivs, mlngCurDeptId, Me.hWnd, Me, True, gblnUseDebugLog)
            End If
            
            Call mobjPacsCapture.SaveImageToStady(dcmImage, mlngAdviceID)
            
            Set mobjPacsCapture = Nothing
    End Select
    
    mstrInstance = dcmImage.InstanceUID
End Sub

Private Sub mobjImageProcess_OnUnload()
    Set mobjImageProcess = Nothing
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
    DTPimg.Left = 0
    DTPimg.Top = 0
    DTPimg.Width = 1400
    DTPimg.Height = 300
    
    cboCache.Left = DTPimg.Width
    cboCache.Top = 0
    cboCache.Width = picMiniCache.ScaleWidth - DTPimg.Width
    cboCache.Height = 300
    
    ucMiniCache.Left = 0
    ucMiniCache.Top = cboCache.Top + cboCache.Height
    ucMiniCache.Width = picMiniCache.ScaleWidth
    ucMiniCache.Height = picMiniCache.ScaleHeight - ucMiniCache.Top
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
    

    Pane2.Title = "����ͼ"
    Pane2.Handle = picReportImage.hWnd
    Pane2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set Pane3 = dkpMain.CreatePane(3, 0, mlngCY3, DockBottomOf, Nothing)
    Pane3.Title = "����ͼ"
    Pane3.Handle = picMiniViewer.hWnd
    Pane3.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        Set pane4 = dkpMain.CreatePane(4, 0, mlngCY3, DockBottomOf, Nothing)
        pane4.Title = "��̨ͼ"
        pane4.Handle = picMiniCache.hWnd
        pane4.Options = PaneNoCloseable Or PaneNoFloatable
        pane4.AttachTo Pane3
        picMiniCache.Visible = True
    Else
        picMiniCache.Visible = False
    End If
    
    Pane3.Selected = True
    
    mblnIsInitFace = True
End Sub

Private Function GetTag(ByVal FolderName As String, ByRef strType As String) As Integer
'�����ļ��������еı�ʶ�ţ�FolderName��Ŀ��Ŀ¼����strType�� ���ء���ʶ�� �� ����顱
On Error GoTo errH
    Dim i As Integer
    Dim strTmp As String
    
    strType = Mid(FolderName, 1, 2)
    strTmp = Mid(FolderName, 3, Len(FolderName) - 2)
    i = InStr(strTmp, "-")
    GetTag = Val(Mid(strTmp, 1, i - 1))
    
    Exit Function
errH:
    GetTag = 0
End Function

Private Function GetStudyUIDFromFolderName(ByVal FolderName As String) As String
'�����ļ��������еļ��UID�����أ����������ļ�����
On Error GoTo errH
    Dim i As Integer
    Dim j As Integer
    
    i = InStr(FolderName, "-")
    j = Len(FolderName)
    
    GetStudyUIDFromFolderName = Mid(FolderName, i + 1, j - i)
    Exit Function
errH:
    GetStudyUIDFromFolderName = FolderName
End Function

Function LoadMiniCache() As Boolean
    Dim i As Integer
    Dim strQueryPath As String
    Dim objFolder2 As Folder, objFolder3 As Folder, objFolder4 As Folder
    Dim strStudyUID As String, strSeriesUID As String
    Dim lngStudyNo As Long, lngSeriesNo As Long
    Dim strAfterTime As String
    Dim dtChose As Date
    Dim intTMP As Integer
    Dim strTag As String  '��λ���ı�ʶ
    Dim strType As String
    Dim curDate As Date
    
    If mblnUseAfterCapture = False Then Exit Function
    
    curDate = zlDatabase.Currentdate
    DTPimg = mdate

    Set mrsImageCache = New ADODB.Recordset
    mrsImageCache.Fields.Append "����", adVarChar, 100
    mrsImageCache.Fields.Append "����", adVarChar, 64
    mrsImageCache.Fields.Append "���UID", adVarChar, 64
    mrsImageCache.Fields.Append "���к�", adVarChar, 18
    mrsImageCache.Fields.Append "����UID", adVarChar, 64
    mrsImageCache.Fields.Append "�������", adVarChar, 20
    mrsImageCache.Fields.Append "·��", adVarChar, 4000
    mrsImageCache.Open
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders
            If InStr(objFolder2.Name, Format(mdate, "yyyymmdd")) > 0 Then ''�������ѡ���ʱ��������

                If objFolder2.SubFolders.Count > 0 Then
                        
                    For Each objFolder3 In objFolder2.SubFolders                            '���UID��
                            
                        If objFolder3.SubFolders.Count >= 0 Then

                            strAfterTime = Format(objFolder3.DateCreated, "YYYY-MM-DD HH:MM:SS")
 
                            strStudyUID = GetStudyUIDFromFolderName(objFolder3.Name)
                                                                  
                            lngStudyNo = lngStudyNo + 1
                            lngSeriesNo = 0
                                    
                            For Each objFolder4 In objFolder3.SubFolders                    '����UID��
                                    
                                strSeriesUID = objFolder4.Name
                                lngSeriesNo = lngSeriesNo + 1
                                        
                                mrsImageCache.AddNew
                                strTag = zlStr.Lpad(GetTag(objFolder3.Name, strType), 3, "0")
                                mrsImageCache!���� = strType & strTag

                                mrsImageCache!���� = str(lngStudyNo)
                                mrsImageCache!���UID = strStudyUID
                                mrsImageCache!���к� = lngSeriesNo
                                mrsImageCache!����UID = strSeriesUID
                                mrsImageCache!������� = strAfterTime

                                mrsImageCache!·�� = objFolder4.Path
                                mrsImageCache.Update
                                        
                            Next
                                  
                        End If
                    Next
                                        
                                        Exit For '��ʱ�Ѿ���������ѡʱ�䣬����ʱ��ѡ��
                End If
            End If 'ʱ��ѡ��
        Next
    End If
    
    If mrsImageCache.RecordCount > 0 Then
        mrsImageCache.Sort = "������� desc"
        mrsImageCache.MoveFirst
    End If

    cboCache.Clear
    ucMiniCache.ImgViewer.Images.Clear
    
    For i = 0 To mrsImageCache.RecordCount - 1
        If i = 0 Then strQueryPath = Nvl(mrsImageCache!·��)
        
        cboCache.AddItem Nvl(mrsImageCache!����) & "     ʱ�䣺" & Format(Nvl(mrsImageCache!�������), "HH:MM:SS")
        mrsImageCache.MoveNext
    Next
    
    If mrsImageCache.RecordCount > 0 Then
        If cboCache.ListIndex < 0 Then
            cboCache.ListIndex = 0
        End If
    End If
    
End Function

Public Function LoadMiniImages() As Boolean
    ucMiniImageViewer.ShowCheckBox = mlngModule <> 1290
    Call ucMiniImageViewer.RefreshImage(slAdvice, mlngAdviceID, mblnMoved, True)
End Function

Private Sub LoadReportImages()
    
    On Error GoTo errH
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable, oPicture As StdPicture
    Dim strPicFile As String
    Dim iRImageCount As Integer
    Dim objFile As New Scripting.FileSystemObject
    Dim i As Integer
    Dim j As Long
    Dim dcmImg As DicomImage
    
    '��ʼ����������
    Call ClearReportImages
     
    iRImageCount = 0
    For i = 1 To UBound(mobjImgCTables)
    
        Set cTable = mobjImgCTables(i)
            
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
        
        For j = 1 To cTable.Pictures.Count
            strPicFile = App.Path & "\PACSPic" & j & ".JPG"
            If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

            Set oPicture = cTable.Pictures(j).OrigPic
            SavePicture oPicture, strPicFile
            If objFile.FileExists(strPicFile) Then
                '��ʾ���ͼ�ͱ���ͼ
                If cTable.Pictures(j).PictureType = EPRMarkedPicture And dcmMark.Images.Count = 0 Then

                    'ֻ�����һ�����ͼ
                    dcmMark.Images.AddNew
                    
                    dcmMark.Images(1).FileImport strPicFile, "BMP"
                    dcmMark.Images(1).tag = cTable.Pictures(j).ID
                    '������ͼ��������
                    Set pobjMarks = cTable.Pictures(j).PicMarks
                    pMarkImageID = cTable.Pictures(j).ID

                    mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(j).Width * Screen.TwipsPerPixelX
                    '��ʾ��ע
                    If cTable.Pictures(j).PicMarks.Count > 0 Then
                        drawPicMarks dcmMark.Images(1), cTable.Pictures(j).PicMarks
                    End If
                Else

                    dcmReportImage(iRImageCount).Images.AddNew
                    Set dcmImg = dcmReportImage(iRImageCount).Images(dcmReportImage(iRImageCount).Images.Count)
                    dcmImg.FileImport strPicFile, "BMP"
                          
                    If cTable.Pictures(j).PicName = "" Then
                        dcmImg.tag = mdcmGlobal.NewUID & ".jpg"
                    Else
                        dcmImg.tag = cTable.Pictures(j).PicName
                    End If
                    
                    '����InstanceUID
                    dcmImg.BorderWidth = 3
                    dcmImg.BorderColour = vbWhite
                    dcmReportImage(iRImageCount).CurrentIndex = 1
                    mselReportImgIndex = 1
                End If
                'ɾ����ʱͼ��
                Kill strPicFile
            End If
        Next j
    Next i
    
    If dcmReportImage.Count > 1 Then dcmReportImage(1).Labels(1).ForeColour = vbRed
    
    Call picReportImage_Resize
Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function InitCTables(Optional ByVal lngFormatId As Long = 0) As Boolean
'��ʼ��������ʽ�е�ͼ�������
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim cTable As cEPRTable
    Dim blnGetTable As Boolean
    Dim lngUbound As Long
    Dim i As Long
    
    ReDim mobjImgCTables(0)
    
    InitCTables = False
    
    If lngFormatId <> 0 Then
      '���ĸ�ʽ���� ������������
        strSql = "Select Id As ���Id From ������������" & vbNewLine & _
        "       Where �ļ�id = [1] And �������� = 3 And Substr(��������, Instr(��������, ';', 1, 18) + 1, 1) = '2'" & vbNewLine & _
        "       Order By �������"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngFormatId)

        If rsTemp.RecordCount > 0 Then
            If rsTemp.RecordCount < dcmReportImage.Count - 1 Then
                If MsgBoxD(Me, "�¸�ʽ��ͼ����������ڵ�ǰ��ʽ����ǰ�Ĳ���ͼ���ᱻɾ�����Ƿ������ʽ��", vbOKCancel) = vbCancel Then
                    Exit Function
                Else
                    '��ɾ�������ͼ���
                    For i = dcmReportImage.Count - 1 - rsTemp.RecordCount To 1 Step -1
                        Unload dcmReportImage(dcmReportImage.Count - 1)
                    Next i
                End If
            End If
        End If
    Else
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
    End If
    
    
    Do While Not rsTemp.EOF
        Set cTable = New cEPRTable
        If mlngReportID = 0 Then
            blnGetTable = cTable.GetTableFromDB(cprET_�����ļ�����, mlngFileID, Val("" & rsTemp!���ID))
        Else
            blnGetTable = cTable.GetTableFromDB(cprET_���������, mlngReportID, Val("" & rsTemp!���ID))
        End If
        
        If blnGetTable Then
            lngUbound = UBound(mobjImgCTables) + 1
            ReDim Preserve mobjImgCTables(lngUbound)
            
            Set mobjImgCTables(lngUbound) = cTable
        End If
        
        Call rsTemp.MoveNext
    Loop
    
    InitCTables = True

End Function

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
    Dim j As Long
    
    For i = 1 To UBound(mobjImgCTables)
        Set cTable = mobjImgCTables(i)
        
        If cTable.Pictures.Count > 0 Then
            For j = 1 To cTable.Pictures.Count
                If cTable.Pictures(j).PictureType = EPRMarkedPicture Then
                    DecideMarkImagesVisible = 1
                    Exit Function
                Else
                    DecideMarkImagesVisible = 0
                End If
            Next
        Else
            DecideMarkImagesVisible = 0
        End If
    Next i
End Function


Private Sub drawPicMarks(img As DicomImage, thisMarks As cPicMarks)
'------------------------------------------------
'���ܣ���ʾ��ע��֧�����ֱ�ţ���ͷ��Բ�Σ����ֱ�ע
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim oneLabel As DicomLabel
    
    On Error GoTo err
    
    img.Labels.Clear
    'thisMarks(i).���Ͷ��� '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ), 6-˳���ţ�7-��ͷ��PACS�����ӣ�
    For i = 1 To thisMarks.Count
        With thisMarks(i)
            If thisMarks(i).���� = 0 Then       '�ı�
                img.Labels.Add GetNewLabel(doLabelText, .X1 * mdblMarkZoom, .Y1 * mdblMarkZoom, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.Text = .����
            ElseIf thisMarks(i).���� = 5 Then   '��Բ
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * mdblMarkZoom, .Y1 * mdblMarkZoom, (.X2 - .X1) * mdblMarkZoom, (.Y2 - .Y1) * mdblMarkZoom)
            ElseIf thisMarks(i).���� = 6 Then   '˳����
                'Բ�α���ɫ
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * mdblMarkZoom - 7, .Y1 * mdblMarkZoom - 7, 14, 14)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.XOR = False
                oneLabel.BackColour = IIf(.���ɫ = 0, vbYellow, .���ɫ)
                oneLabel.Transparent = False
                oneLabel.tag = m_LabelTag_Back
                
                'Բ�ο�
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * mdblMarkZoom - 7, .Y1 * mdblMarkZoom - 7, 14, 14)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.XOR = False
                oneLabel.ForeColour = vbBlack
                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Circle
                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)
                
                'Բ�α������
                img.Labels.Add GetNewLabel(doLabelText, .X1 * mdblMarkZoom - 7 + 1, .Y1 * mdblMarkZoom - 7, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.ForeColour = vbBlack
                oneLabel.XOR = False
                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Number
                oneLabel.FontSize = 8
                oneLabel.FontName = "Arial Bold"
                oneLabel.AutoSize = True
                oneLabel.Text = .����
                If Val(.����) < 10 Then  '10���µ����֣���Ҫ΢��һ��λ�ã����ֲ��ܳ�����ԲȦ�����м�
                    oneLabel.Left = oneLabel.Left + 3
                End If
                
                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)
                img.Labels(img.Labels.Count - 2).TagObject = oneLabel  'TagObject�γɱջ�
                
            ElseIf thisMarks(i).���� = 7 Then   '��ͷ
                img.Labels.Add GetNewLabel(doLabelArrow, .X1 * mdblMarkZoom, .Y1 * mdblMarkZoom, (.X2 - .X1) * mdblMarkZoom, (.Y2 - .Y1) * mdblMarkZoom)
            End If
        End With
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
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
    Dim i As Long
    Dim j As Integer
    
    '��ʼ����������
    Call ClearReportImages
    
    If InitCTables(FormatID) = True Then
        '��ȡͼ����еı��ͼ�ͱ���ͼ
        iRImageCount = 0
        pTableID = ""
        
        For i = 1 To UBound(mobjImgCTables)
            Set cTable = mobjImgCTables(i)

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
            For j = 1 To cTable.Pictures.Count
                strPicFile = App.Path & "\PACSPic" & j & ".JPG"
                If objFile.FileExists(strPicFile) Then objFile.DeleteFile strPicFile, True

                Set oPicture = cTable.Pictures(j).OrigPic
                SavePicture oPicture, strPicFile
                If objFile.FileExists(strPicFile) Then
                    '��ʾ���ͼ
                    If cTable.Pictures(j).PictureType = EPRMarkedPicture Then
                        blnHasMarkImage = True
                        '�������ǰ���ͼ���ٸ���
                        dcmMark.Images.Clear
                        dcmMark.Images.AddNew
                        dcmMark.Images(1).FileImport strPicFile, "BMP"
                        dcmMark.Images(1).tag = cTable.Pictures(j).ID
                        '�����ǰû�б�ǣ����ȡ�¸�ʽ�б��ͼ�ı��
                        If pobjMarks Is Nothing Then
                            Set pobjMarks = cTable.Pictures(j).PicMarks
                        End If
                        pMarkImageID = cTable.Pictures(j).ID

                        mdblMarkZoom = dcmMark.Images(1).SizeX / cTable.Pictures(j).Width * Screen.TwipsPerPixelX
                        '��ʾ��ע
                        If pobjMarks.Count > 0 Then
                            drawPicMarks dcmMark.Images(1), pobjMarks
                        End If
                    End If
                    'ɾ����ʱͼ��
                    Kill strPicFile
                End If
            Next j
        Next i
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
'    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
'    If Button = 1 Then ucMiniCache.ImgChecked(ucMiniCache.SelectIndex) = Not ucMiniCache.ImgChecked(ucMiniCache.SelectIndex)
End Sub

Private Sub ucMiniCache_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniCache.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(2)
End Sub

Private Sub ucMiniImageViewer_AfterSaveStudy()
    Call LoadMiniImages
End Sub

Private Sub ucMiniImageViewer_OnDbClick(ByVal lngSelectedIndex As Long, blnContinue As Boolean)
    If ucMiniImageViewer.CurImageCount > 0 And mSelViewerIndex <> 0 And dcmReportImage.Count > 1 Then
        '�жϵ�ǰ˫���Ĳ�������
        If mintImageDblClick = 0 Then   'ֱ��д�뱨��
            Dim dcmImage As DicomImage
            Set dcmImage = mSelMiniImg
            
            '���ý���ǰͼ����ӵ�ͼ������
            Call DcmAddImage(dcmImage, mSelViewerIndex)
            Call ChangeImgRptTag(lngSelectedIndex, 1)
        Else                            '�ȴ�ͼƬ�༭����
            Call OpenImageProcessWind
        End If

    End If
    
    blnContinue = False
End Sub

Private Sub AutoAddImg()
'------------------------------------------------
'���ܣ��Զ����Ѿ���ǵ�ͼ����ӵ�������
'������
'���أ���
'-----------------------------------------------
    
    Dim i As Long
    
    On Error GoTo err
    
    '�ӱ�������ͼ�У������Ѿ����Ϊ0��ͼ��
    For i = 1 To ucMiniImageViewer.ImgViewer.Images.Count
        If ucMiniImageViewer.ImgViewer.Images(i).tag.ReportImage = 0 Then
            ucMiniImageViewer.ImgViewer.Images(i).tag.ReportImage = 1
            Call DcmAddImage(ucMiniImageViewer.ImgViewer.Images(i), mSelViewerIndex)
        End If
    Next i
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ucMiniImageViewer_OnMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub

    If Button = 1 And mlngShowBigImg = 3 Then Call OpenImageProcessWind
'
'    If Button = 1 Then ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex) = Not ucMiniImageViewer.ImgChecked(ucMiniImageViewer.SelectIndex)
End Sub

Private Sub ucMiniImageViewer_OnMouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngShowBigImg = 1 Then RaiseEvent AfterShowBigImage
End Sub

Private Sub ucMiniImageViewer_OnMouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If ucMiniImageViewer.ImgViewer.Images.Count <= 0 Then Exit Sub
    If Button = 2 Then Call ShowPopupImage(1)
End Sub

Private Sub ucMiniImageViewer_OnSaveImage(ByVal dcmImage As DicomObjects.DicomImage, ByVal lngImageType As Long)
    If mstrInstance = dcmImage.InstanceUID Then Exit Sub
    Select Case lngImageType
        Case 0  '���ͼ
            Call DcmAddMarkImage(dcmImage)
        Case 1  '����ͼ
            Call DcmAddImage(dcmImage, mSelViewerIndex)
        Case 2  '���ͼ
            If mobjPacsCapture Is Nothing Then
                Set mobjPacsCapture = CreateObject("zl9PacsImageCap.clsPacsCapture")
                
                Call mobjPacsCapture.zlInitModule(gcnOracle, glngSys, mlngModule, gstrPrivs, mlngCurDeptId, Me.hWnd, Me, True, gblnUseDebugLog)
            End If
            
            Call mobjPacsCapture.SaveImageToStady(dcmImage, mlngAdviceID)
            
            Set mobjPacsCapture = Nothing
    End Select
    
    mstrInstance = dcmImage.InstanceUID
End Sub

Private Sub ucMiniImageViewer_OnSelChange(ByVal lngSelectedIndex As Long)
    Set mSelMiniImg = ucMiniImageViewer.SelectImage
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
        '�Զ������Ѿ�ѡ�еı���ͼ��
        Call AutoAddImg
        pImageModified = False  '�Զ����ر���ͼ�󣬽��������ó�δ�޸ĵ�״̬�������Զ������������ʾ�û�����
        '��Ǳ���ˢ���Ѿ�����ͼ��
        blnLoadImages = True
    End If
End Sub

Private Sub ClearEmptyFolder(ByVal blNoReason As Boolean)
'intType 0:���͵����  1:ɾ��ͼ��
'blNoReason �Ƿ���������ִ�б����̣����ڹرճ����ʱ��ִ��
'��տ�Ŀ¼����Ӧ������������ǰѡ�е��ǵ������±�ʶ����ִ�д˲���
'�����ж�������ǰѡ���Ƿ��ǵ������±�ʶ
    Dim curTime As Date
    Dim strTime As String
    Dim objFolder1 As Folder, objFolder2 As Folder, objFolder3 As Folder
    Dim strType As String
    Dim strTag As String
    Dim i As Long
    Dim blDT As Boolean
    Dim blTag As Boolean
    
    On Error GoTo errH
    blDT = False
    blTag = False
    
    If mblnUseAfterCapture And mlngModule <> 1290 Then
        If Not mblDel Then
            Call CheckSendOnImageCountChangedChanged(0)
        Else
            Call CheckSendOnImageCountChangedChanged(1)
        End If
    End If
    mblDel = False
   
    If blNoReason = False Then
        curTime = zlDatabase.Currentdate
        '�ǵ��첢��ѡ�е��ǵ�ǰ��ʶ���Ͳ�������ղ���

        If (Format(DTPimg.value, "yyyymmdd") = Format(curTime, "yyyymmdd")) And (InStr(cboCache.Text, "��ʶ" & zlStr.Lpad((mintTagMaxTag), 3, "0")) > 0) Then
            Debug.Print "��ֹ���"
            Exit Sub
        Else
            Debug.Print "�������"
        End If
        
    End If
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Sub

    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder1 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders '''ʱ��
            If objFolder1.Name = Format(curTime, "yyyymmdd") Then blDT = True
            If objFolder1.SubFolders.Count > 0 Then
            
                For Each objFolder2 In objFolder1.SubFolders '''���uid
                    If InStr(objFolder2.Name, "��ʶ" & mintTagMaxTag) > 0 Then blTag = True
                    If objFolder2.SubFolders.Count > 0 Then
                    
                        For Each objFolder3 In objFolder2.SubFolders '''����UID
                            If objFolder3.Files.Count = 0 Then
                                '���ǵ������±�ʶ�����Ŀ¼
                                If Not (blDT And blTag) Then Call mobjFile.DeleteFolder(objFolder3.Path)
                                
                            End If
                        Next
                        
                        If objFolder2.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder2.Path)
                    Else
                        Call mobjFile.DeleteFolder(objFolder2.Path)
                    End If
                Next
                
                If objFolder1.SubFolders.Count = 0 Then Call mobjFile.DeleteFolder(objFolder1.Path)
            Else
                Call mobjFile.DeleteFolder(objFolder1.Path)
            End If
        Next
    End If
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, Me.Caption)
End Sub

Private Function GetTodayTagMax(ByVal curDate As Date) As Integer
    '���㵱������ʶ
    Dim strDate As String
    Dim intTMP As Integer
    Dim strType As String
    Dim strStudyUID As String
    Dim objFolder2 As Folder, objFolder3 As Folder
    
    On Error GoTo errH
    
    mintTagMax = 1
    mintTagMaxTag = 1
    
    strDate = Format(curDate, "yyyymmdd")
    
    If mobjFile.FolderExists(mstrAfterImgPath) = False Then Exit Function
    
    If mobjFile.GetFolder(mstrAfterImgPath).SubFolders.Count > 0 Then
        For Each objFolder2 In mobjFile.GetFolder(mstrAfterImgPath).SubFolders
            If InStr(objFolder2.Name, strDate) > 0 Then                                 'ʱ��ѡ��
            
                If objFolder2.SubFolders.Count > 0 Then                                  'ʱ����Ƿ�����Ŀ¼
                
                    For Each objFolder3 In objFolder2.SubFolders                            '���UID��
                    
                        If objFolder3.SubFolders.Count > 0 Then

                            strStudyUID = GetStudyUIDFromFolderName(objFolder3.Name)
                            
                            intTMP = GetTag(objFolder3.Name, strType)
                            If intTMP > mintTagMax Then mintTagMax = intTMP
                            
                            If strType = "��ʶ" Then
                                If intTMP > mintTagMaxTag Then mintTagMaxTag = intTMP
                            End If
                            
                        End If

                    Next
                    
                End If 'ʱ����Ƿ�����Ŀ¼
                
            End If 'ʱ��ѡ��
        Next
    End If
    
    GetTodayTagMax = mintTagMax
    
    Exit Function
errH:
    BUGEX "GetTodayTagMax output= -1"
    GetTodayTagMax = -1
End Function

Private Function GetNowTag(ByVal blIsNeedAddOne As Boolean) As Integer
'��õ�ǰ��ʶ,blIsNeedAddOne:�Ƿ����+1 ������ػ��߳�ʼ����Ӧ��+1�����͵���̨�����������
    Dim curDate As Date
    
    curDate = zlDatabase.Currentdate
    '���µ�������ʶ
    mintTagNow = GetTodayTagMax(curDate)
    
    If blIsNeedAddOne = True Then mintTagNow = mintTagNow + 1

End Function

Public Sub UseAfterImgChanged(ByVal blUse As Boolean)
    Dim objImage As Pane, objCache As Pane, objTmp As Pane
    Dim blHavePane As Boolean
    Dim i As Integer
    
    mblnUseAfterCapture = blUse
    
    If blUse = True Then
    '�Ƿ���Ҫ�������ж�
        blHavePane = False
        
        For i = 1 To 5
            Set objTmp = dkpMain.FindPane(i)
            
            If Not objTmp Is Nothing Then
            
                If objTmp.Title = "����ͼ" Then Set objImage = dkpMain.FindPane(i)
                If objTmp.Title = "��̨ͼ" Then blHavePane = True
                
            End If
        Next
        
        If blHavePane = False And Not objImage Is Nothing Then
            
            picMiniCache.Visible = True
            
            Set objCache = dkpMain.CreatePane(4, 0, 400, DockLeftOf)
            objCache.Title = "��̨ͼ"
            objCache.Handle = picMiniCache.hWnd
            objCache.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            
            If objImage.Title = "����ͼ" Then
                objCache.AttachTo objImage
                objImage.Selected = True
                LoadMiniCache
            End If
            
        End If
    Else
    '�Ƿ���Ҫ���ٵ��ж�
        blHavePane = False

        For i = 1 To 5
            Set objTmp = dkpMain.FindPane(i)
            
            If Not objTmp Is Nothing Then
                If objTmp.Title = "��̨ͼ" Then
                    blHavePane = True
                    Exit For
                End If
            End If
        Next
    
        If blHavePane = True Then Call dkpMain.DestroyPane(objTmp)
        picMiniCache.Visible = False
        
    End If
    
    Exit Sub
errH:
    Call err.Raise(0, , "��̨ͼ��ǩ�������" & err.Description)
End Sub

Public Sub InitParaForAfterImage(ByVal lngCurDeptId As Long, ByVal lngModule As Long)
    mlngCurDeptId = lngCurDeptId
    mlngModule = lngModule
End Sub

Public Sub ReportImageAdd(strInstanceUID As String)
    '���ͼ��Ϊ����ͼ�����Զ����뱨����
    Dim dcmImage As DicomImage
    Dim i As Integer
    
    For i = 1 To ucMiniImageViewer.ImgViewer.Images.Count
        If ucMiniImageViewer.ImgViewer.Images(i).InstanceUID = strInstanceUID Then
            Set dcmImage = ucMiniImageViewer.ImgViewer.Images(i)
            If dcmImage.tag.ReportImage = "" Then
                '���á�����ͼ��������
                dcmImage.tag.ReportImage = 1
                '��������ͼ�����
                Call ucMiniImageViewer.DrawReportImgTag(ucMiniImageViewer.ImgViewer.Images(i))
                '��ӱ���ͼ
                Call DcmAddImage(ucMiniImageViewer.ImgViewer.Images(i), mSelViewerIndex)
            End If
        End If
    Next i
    
End Sub

Private Sub SaveLocalPar()
'���汾�ز���
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\frmReportImage", "���δ�ͼ", IIf(mblnImageShield, 1, 0)
End Sub

Private Sub GetLocalPar()
'��ȡ���ز���

    mblnImageShield = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\frmReportImage", "���δ�ͼ", 0)) = 1
End Sub

Public Sub CloseImageProcess()
    'ucMiniImageViewer.BigImageWay = 1 ����ƶ�ʱ��ʾ��ͼ
    If ucMiniImageViewer.BigImageWay = 1 Then Call ucMiniImageViewer.CloseImageProcess
End Sub

