VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmFilm 
   Caption         =   "��Ƭ��ӡԤ��"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   Icon            =   "frmFilm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimePass 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox picBak 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5568
      Left            =   720
      ScaleHeight     =   5565
      ScaleWidth      =   9135
      TabIndex        =   0
      Top             =   720
      Width           =   9135
      Begin VB.PictureBox picFilm 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H00000000&
         Height          =   4056
         Left            =   1800
         ScaleHeight     =   4020
         ScaleWidth      =   5340
         TabIndex        =   2
         Top             =   570
         Width           =   5376
         Begin DicomObjects.DicomViewer FilmViewer 
            DragIcon        =   "frmFilm.frx":000C
            Height          =   1140
            Index           =   0
            Left            =   3720
            TabIndex        =   3
            Top             =   2520
            Visible         =   0   'False
            Width           =   1230
            _Version        =   262147
            _ExtentX        =   2159
            _ExtentY        =   2011
            _StockProps     =   35
            BackColor       =   0
            UseScrollBars   =   0   'False
         End
      End
      Begin VB.VScrollBar VScro 
         Height          =   3888
         Left            =   4
         Max             =   1
         TabIndex        =   1
         Top             =   672
         Width           =   250
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   7875
      Width           =   11400
      _ExtentX        =   20108
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
            Object.Width           =   12647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ImageListMouse 
      Left            =   240
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":0CD6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":19B0
            Key             =   "��ѡ����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":1CCA
            Key             =   "ǰ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":1FE4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":2CBE
            Key             =   "��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":2FD8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":38B2
            Key             =   "�ü�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":3BCC
            Key             =   "��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":3EE6
            Key             =   "��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":4200
            Key             =   "��"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFilm.frx":451A
            Key             =   "��"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImgIcons 
      Left            =   240
      Top             =   2640
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmFilm.frx":4834
   End
   Begin XtremeCommandBars.CommandBars CommBar_Film 
      Left            =   300
      Top             =   1080
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFilm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event AfterPrinted(strImageUIDS As String)      '��ӡ���

'''''''���¼����Ӵ���'''''''''''''''
Public WithEvents mfrmFilmView As frmFilmView
Attribute mfrmFilmView.VB_VarHelpID = -1

'����Ĺ�������-------------------------------------

Public clsTruePrinter As clsDicomPrint      ''DICOM��ӡ��������
Public SelectedImage As DicomImage          ''��¼��ǰ��ѡ�е�ͼ���ṩ��ģ��ͳһ������λ���ܰ�ť
Public intMouseState As Integer             ''��¼����״̬��0���ޣ�1��������2�����Σ�3������;4-��;5-��ѡ����;6-�ü�:7-���ֱ�ע����frmFilmView�����н�������
Public pstrSideMarker As String             ''��¼��ǰ��Ҫ��ע����λ���֣���frmFilmView�����н�������
Public blnDefaultWW2 As Boolean             ''��¼˫����״̬���ṩ��ģ��ͳһ������λ���ܰ�ť
Public f As frmViewer

'�����˽�б���'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private imgsPrint As New DicomImages        ''�����ӡ��ͼ����������
Private mblnPrinted As Boolean              ''��¼�Ƿ���ͼ���Ѿ�����ӡ��
Private mintPrintFilmCount As Integer       ''��¼��ӡ�ɹ��Ĵ���
Private mintCellSpacing As Integer          ''��ʾ��ƬԤ����ʱ��ͼƬ֮��ļ�࣬��ʾ�ã���ʵ�ʴ�ӡ�޹�
Private mintFilmHeight As Integer           ''��Ƭ�ĸ߶ȣ���λӢ��
Private mintFilmWidth As Integer            ''��Ƭ�Ŀ�ȣ���λӢ��
Private mblnIsPortrait As Boolean           ''�Ƿ������ӡ
Private mblnIsRow As Boolean                ''��ǰҳ���У��Ƿ�������
Private mblnIsCustom As Boolean             ''��ǰҳ���У��Ƿ��������Զ���
Private mdubFilmRate As Double              ''��Ƭ�ĳ������
Private mdubScreenRate As Double            ''��Ļ�ĳ������
Private mblnBegin As Boolean

Private marrRCCount() As Integer            ''��ǰҳ���У�ÿ��/ÿ�е�ͼ����Ŀ����STANDARD����£�aRCCount(1)��ʾ����
Private marrPages() As FilmType             ''ÿһҳ�е�Viewer������DICOM��׼���ַ�ʽ������ά���ǽ�Ƭҳ��

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mintSelectedViewer As Integer       ''��¼��ǰ��ѡ�е�Viewer���
Private mintSelectedImage As Integer        ''��¼��ǰ��ѡ�е�ͼ�����
Private mintBaseX As Long                   ''��¼���ԭ����Xλ��
Private mintBaseY As Long                   ''��¼���ԭ����Yλ��
Private mdcmSelectLabel As DicomLabel       ''��ǰ��ѡ�еı�ע
Private mblnDcmViewDown As Boolean          ''�����ж�dcmView������Ƿ񱻰���
Private mblnLabelMoving As Boolean          ''�����ƶ��ü���
Private mblnCheckPrinter As Boolean         ''�Ƿ����ӡ����״̬
Private mblnInTest As Boolean               ''��¼�Ƿ��ڲ���״̬
Private mintPageRange As Integer            ''��ӡҳ����Χ��0-ȫ����1-��ǰҳ
Private mintTBMainPosition As Integer       ''��¼��������λ��
Private mintTBImageProcessPosition As Integer   ''��¼ͼ�����������λ��
Private mblnPrinting As Boolean             ''�Ƿ����ڴ�ӡ�Ĺ�����
Private mblnClearAfterPrint As Boolean      ''�Ƿ��ӡ�����ͼ��

''''''''''''''''�ü�''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutViewer                    '�ü������ڵ�viewer���
Private mintCutOutImage                     '�ü������ڵ�ͼ�����
Private mintCutOutLabel                     '�ü������ڵı�ע���
Private mdblCutOutRatio As Double           '�ü��ı���������ǹ̶�������ֱ�Ӽ�¼�������޹̶�������Ϊ0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��¼ÿ��Viwer��ͼ�������������
Private Type TLayout
    intRows As Integer
    intColumns As Integer
End Type

Private Type FilmType
    intViewerCount As Integer               '���ݲ��ּ��������Viewer������
    strPageFormat As String                 '��һҳ�Ĳ��֣�DICOM��׼����
    ViewerLayout() As TLayout               '��һҳ�У�ÿ��Viewer�����в���
    intImageCount As Integer                '��һҳ�е�ͼ������
End Type

Private Type ImageSize
    intWidth As Integer                     'ͼƬ�������
    intHeight As Integer                    'ͼƬ�����߶�
End Type

'-----��Ƭ��ӡ�õ�TAG����--------------------------------------------------
Private Const zlSpliter = "-ZL-"            'TAG�м�¼���ݵķָ���
Private Const TAG_ѡ�� = "1"
Private Const TAG_��λ = "2"
Private Const TAG_V��� = "3"
Private Const TAG_V�߶� = "4"


'����ĺ���'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub subConfig()
'------------------------------------------------
'���ܣ���frmFilmConf���������õ��Ű��ʽӦ�õ���Ƭ��ӡ�С�
'��������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strFilmFormat As String
    
    On Error GoTo err
    
    With frmFilmConf
        
        '��¼��Ƭ�Ŀ�Ⱥ͸߶�
        i = InStr(.cobSize.Text, "X")
        mintFilmWidth = Val(Mid(.cobSize.Text, 1, i - 1))
        mintFilmHeight = Val(Mid(.cobSize.Text, i + 1))
        
        '��¼��Ƭ����
        mblnIsPortrait = IIf(.cobAspect.Text = "����", True, False)       ''�Ƿ������ӡ
        
        'Option(0)--��׼���У�Option(1)---���Զ��壻Option(2) ---���Զ���
        
        '��ϲ���¼��Ƭ��ʽ
        If Not .Option(2) Then        '���Զ�����߱�׼��ʽ
            If Not .Option(0) Then
                strFilmFormat = "ROW\" & .txtC(1)
            Else
                strFilmFormat = "STANDARD\" & .txtCol & "," & .txtRow
            End If
        Else            '���Զ���
            strFilmFormat = "COL\" & .txtC(1)
        End If
        If Not .Option(0) Then
            If .Option(1) Then
                For i = 2 To Val(.txtRow)
                    strFilmFormat = strFilmFormat & "," & .txtC(i)
                Next i
            Else
                For i = 2 To Val(.txtCol)
                    strFilmFormat = strFilmFormat & "," & .txtC(i)
                Next i
            End If
        End If
    End With
    
    '���ò˵��϶�Ӧ�Ľ�Ƭ���ͽ�Ƭ��ʽ
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True).Text = frmFilmConf.cobSize.Text
    Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = strFilmFormat
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub CommBar_Film_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim intNowIndex As Integer
    Dim intTemp As Integer
    Dim thisControl As CommandBarControl
    Dim strLayout As String
    
    On Error GoTo err
    
    '''''''''''''''''''''''''''''[���ܼ����ô���λ����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 349 And control.Id <= 360 Then
        For i = 349 To 360
            If Not CommBar_Film.Item(3).FindControl(, i, , True) Is Nothing Then
                CommBar_Film.Item(3).FindControl(, i, , True).Checked = False
                If i = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
                    CommBar_Film.Item(3).FindControl(, i, , True).Checked = False
                End If
            End If
        Next
        control.Checked = True
        subFunctionWL CommBar_Film.Item(3).FindControl(, control.Id, , True), Me
        If control.Id = ID_Active_AdjustWindow_HandAdjustWindow_Custom Then
            CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
            CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow_Custom, , True).Checked = True
        End If
        
        Call subFilmViewButtonClick(CommBar_Film.Item(3).FindControl(, control.Id, , True))
        Exit Sub
    End If
    
    Select Case control.Id
    Case ID_frmFilm_FilmCol             '����
        If Not CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked Then
            CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = True
            CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = False
            mblnIsPortrait = True
            Call picBak_Resize
        End If
    Case ID_frmFilm_FilmRow             '����
        If Not CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked Then
            CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = True
            CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = False
            mblnIsPortrait = False
            Call picBak_Resize
        End If
    Case ID_frmFilm_RectPhotCase        '������ͼ���
        CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked = Not control.Checked
        Call picBak_Resize
    Case ID_frmFilm_FormatCustom        '��ʽ����
        Call subSetFilmFormat
    Case ID_frmFilm_TakePictures        '����
        Call CommBar_Execute_PrintFilm
        
    Case ID_frmFilm_FilmSize            '��Ƭ��С
        i = InStr(control.Text, "X")
        If i <> 0 Then
            mintFilmWidth = Val(Mid(control.Text, 1, i - 1))
            mintFilmHeight = Val(Mid(control.Text, i + 1))
            Call picBak_Resize
        End If
    Case ID_frmFilm_Format              '��Ƭ��ʽ
        '�޸Ľ�Ƭ��ʽ
        Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
        '������ʾ��һҳ
        Call subShowOnePage(Me.VScro.Value)
    Case ID_frmFilm_Camera              '��ӡ��
        Dim strThisFilmSize As String
        If Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text <> "" Then
            strThisFilmSize = cDICOMPrinter(Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text).strFilmSize
            Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text = strThisFilmSize
            CommBar_Film_Execute Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True)
        End If
    Case ID_frmFilm_Quit                '�˳�
        Unload Me
    Case ID_frmFilm_OpenImages          ''��ͼ��
        Dim strImageIDs As String
        strImageIDs = frmPACSImg.zlOpenImages(Me, f)
        '��ͼ��
        Call OpenImages(strImageIDs)
    Case ID_frmFilm_DeleteImg            ''ɾ��ͼ��
        Call subDelImage
    Case ID_Active_AdjustWindow_HandAdjustWindow             ''����
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 1, 0)
    Case ID_frmFilm_Pan                  ''����
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 2, 0)
    Case ID_frmFilm_Zoom                 ''����
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 3, 0)
    Case ID_frmFilm_FilterLengthUp       ''ƽ������
        If Not SelectedImage Is Nothing Then
            Call SubImageFiltering("miFilterLengthUp", SelectedImage)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FilterLengthDown     ''ƽ������
        If Not SelectedImage Is Nothing Then
            Call SubImageFiltering("miFilterLengthDown", SelectedImage)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RectZoom             ''��ѡ����
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 5, 0)
    Case ID_frmFilm_CutOut               ''�ü�
        subCheckToolBar control
        intMouseState = IIf(control.Checked = True, 6, 0)
        If intMouseState = 6 Then
            Call subCutOutClick
        End If
    Case ID_frmFilm_CutOut_Custom           ''���ɱ����ü�
        Set thisControl = CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True)
        Call subCheckToolBar(thisControl)
        thisControl.Checked = True
        Call subCutOutButtonState(control.Id)
        intMouseState = 6
        mdblCutOutRatio = 0
    Case ID_frmFilm_CutOut_14X17, ID_frmFilm_CutOut_11X14, ID_frmFilm_CutOut_10X14, _
        ID_frmFilm_CutOut_8X10, ID_frmFilm_CutOut_14X14, ID_frmFilm_CutOut_17X14, ID_frmFilm_CutOut_14X11, _
        ID_frmFilm_CutOut_14X10, ID_frmFilm_CutOut_10X8
        
        '�̶������ü�,ֻҪ�������ͽ���ü�״̬
        Set thisControl = CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True)
        Call subCheckToolBar(thisControl)
        thisControl.Checked = True
        intMouseState = 6
        Call subCutOutButtonState(control.Id)
        Call subCutOutRatio(control.Id)
    Case ID_frmFilm_Invert               ''����
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "Invert")
            Call subSynchronalImg(False, IMG_SYN_WINDOW)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RotateLeft           ''������ת90��
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "RotateAnticlockwise")
            Call subSynchronalImg(False, IMG_SYN_ROTATE)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_RotateRight          ''������ת90��
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "RotateClockwise")
            Call subSynchronalImg(False, IMG_SYN_ROTATE)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FlipHorizontal       ''���Ҿ���
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "FlipHorizontal")
            Call subSynchronalImg(False, IMG_SYN_FLIP)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_FlipVertical         ''���¾���
        If Not SelectedImage Is Nothing Then
            Call subFlipRotate(SelectedImage, "FlipVertical")
            Call subSynchronalImg(False, IMG_SYN_FLIP)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_Label_L             ''��ע����
        intMouseState = 7
        pstrSideMarker = "��"
        subCheckToolBar control
    Case ID_frmFilm_Label_R             ''��ע����
        intMouseState = 7
        pstrSideMarker = "��"
        subCheckToolBar control
    Case ID_frmFilm_Label_A             ''��ע����
        intMouseState = 7
        pstrSideMarker = "ǰ"
        subCheckToolBar control
    Case ID_frmFilm_Label_P             ''��ע����
        intMouseState = 7
        pstrSideMarker = "��"
        subCheckToolBar control
    Case ID_frmFilm_Label_S             ''��ע����
        intMouseState = 7
        pstrSideMarker = "��"
        subCheckToolBar control
    Case ID_frmFilm_Label_I             ''��ע����
        intMouseState = 7
        pstrSideMarker = "��"
        subCheckToolBar control
    Case ID_frmFilm_Label_Delete        ''�����ע����
        If Not SelectedImage Is Nothing Then
            For i = SelectedImage.Labels.Count To G_INT_SYS_LABEL_COUNT + 1 Step -1
                If SelectedImage.Labels(i).Text = "��" Or SelectedImage.Labels(i).Text = "��" Or _
                    SelectedImage.Labels(i).Text = "��" Or SelectedImage.Labels(i).Text = "��" Or _
                    SelectedImage.Labels(i).Text = "ǰ" Or SelectedImage.Labels(i).Text = "��" Then
                    
                    SelectedImage.Labels.Remove (i)
                End If
            Next i
            SelectedImage.Refresh False
        End If
        '�ѱ�ע����ͼ���ϴ���ԭʼͼ������
        Call subReloadImgsPrint
    Case ID_frmFilm_Resume               ''�ָ�
        If Not SelectedImage Is Nothing Then
            Call subSynchronalImg(True, IMG_SYN_All)
            Call subFilmViewButtonClick(control)
        End If
    Case ID_frmFilm_SelAll              ''ͼ��ȫѡ,�Զ�ͼ��ͬ��
        Call SelAllImage(True)
    Case ID_frmFilm_SelSeries           ''ѡ��ǰ����
        Call SelOneSeries
    Case ID_frmFilm_SelInverse          ''��ѡ
        Call SelectInverse
    Case ID_frmFilm_SelNone             ''ȫ��
        Call SelAllImage(False)
    Case ID_frmFilm_Divide              ''ͼ��ָ�
            '��ʾ�ָ�ѡ�񴰿�
            strLayout = frmFilmLayout.ShowMe(Me)
            If Len(strLayout) = 3 And Val(left(strLayout, 1)) <> 0 And Val(Right(strLayout, 1)) <> 0 _
                And Val(left(strLayout, 1)) <= 5 And Val(Right(strLayout, 1)) <= 5 Then
                
                '����ͼ��ָ�
                Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text, _
                    mintSelectedViewer, Val(Right(strLayout, 1)), Val(left(strLayout, 1)))
                    
                '������ʾ��һҳͼ��
                Call subShowOnePage(Me.VScro.Value)
            End If
            '���¶�ȡ��ӡ����ղ���
            mblnClearAfterPrint = IIf(GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "��ӡ�����", "1") = 0, False, True)
    Case ID_frmFilm_UnDivide            ''���Էָ���
        Dim imgTemp As DicomImage
        Dim arrImageSize() As ImageSize
        
        Set clsTruePrinter = funFillPrinterParams(False)
        
        If clsTruePrinter Is Nothing Then Exit Sub
        
        '����ÿһ��ͼ�����ֱ���
        Call subCalImageMaxSize(clsTruePrinter.strFilmSize, clsTruePrinter.strFormat, clsTruePrinter.intImageResolution, arrImageSize)
        
        If UBound(arrImageSize) >= mintSelectedViewer Then
            Set imgTemp = funAssembleImage(FilmViewer(mintSelectedViewer), arrImageSize(mintSelectedViewer).intWidth, arrImageSize(mintSelectedViewer).intHeight)
        Else
            Set imgTemp = funAssembleImage(FilmViewer(mintSelectedViewer))
        End If
        If imgTemp Is Nothing Then Exit Sub
        
        '���ͼ��
        imgsPrint.Add imgTemp
        'ͼ�������ˣ�����ҳ��
        Call subRecalPages
        
        '�������ӵ�ͼ���ڵ�ǰҳ����������ʾ��һҳ��ͼ��
        If imgsPrint.Count < funGetStartImgNo(Me.VScro.Value, 1, 1) + marrPages(Me.VScro.Value).intImageCount Then
            Call subShowPrintImages(Me.VScro.Value)
        End If
        
    Case ID_frmFilm_ImgIncrease     ''ͼ����������
        Call subImageSort(True)
    Case ID_frmFilm_ImgDecrease     ''ͼ����������
        Call subImageSort(False)
    End Select
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCheckToolBar(control As CommandBarControl)
    '�л��ؼ�ѡ��״̬
    Dim blnChecked As Boolean
    
    On Error Resume Next
    
    '����ǴӲü�״̬�˳�����Ҫ����ü���
    If CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).Checked = True Then
        If mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
            If mintCutOutViewer < FilmViewer.Count Then
                If mintCutOutImage <= FilmViewer(mintCutOutViewer).Images.Count Then
                    If mintCutOutLabel = FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then
                        'ɾ����ѡ�õ���ʱ��ע
                        FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Remove mintCutOutLabel
                        Set mdcmSelectLabel = Nothing
                        FilmViewer(mintCutOutViewer).Refresh
                        '���޸Ĺ���ͼ���ϴ���ԭʼͼ������
                        Call subReloadImgsPrint
                    End If
                End If
            End If
        End If
    End If
    
    If Not control Is Nothing Then blnChecked = control.Checked
    
    CommBar_Film.Item(3).FindControl(, ID_Active_AdjustWindow_HandAdjustWindow, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Pan, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Zoom, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_RectZoom, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).Checked = False
    
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_R, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_L, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_A, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_P, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_I, , True).Checked = False
    CommBar_Film.Item(3).FindControl(, ID_frmFilm_Label_S, , True).Checked = False
    
    If Not control Is Nothing Then
        CommBar_Film.Item(3).FindControl(, control.Id, , True).Checked = Not blnChecked
    End If
    
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
End Sub

Private Sub CommBar_Film_GetClientBordersWidth(left As Long, top As Long, Right As Long, Bottom As Long)
    If sbStatusBar.Visible Then
        Bottom = sbStatusBar.height
    End If
End Sub

Private Sub CommBar_Film_Resize()
    On Error Resume Next
    
    Dim left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.CommBar_Film.GetClientRect left, top, Right, Bottom
    If Right >= left And Bottom >= top Then
        picBak.Move left, top, Right - left, Bottom - top
    Else
        picBak.Move 0, 0, 0, 0
    End If
End Sub

Private Sub CommBar_Film_Update(ByVal control As XtremeCommandBars.ICommandBarControl)

    '�������״̬���²�����ʾ,���������״̬
    '����״̬��0���ޣ�1��������2�����Σ�3������;4-��;5-��ѡ����;6-�ü�:7-���ֱ�ע
    Select Case intMouseState
        Case 0
            Me.MousePointer = 0
            sbStatusBar.Panels(2).Text = "��קͼ�����ǰ���ƶ�ͼ�񣬰�ͼ��ŵ��հ�����Ϊ����ͼ��"
        Case 1
            Me.MouseIcon = ImageListMouse.ListImages("����").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "��ק��꣬�����ֶ����Ƶ�ͼ�󴰿�λ����ģʽ"
        Case 2
            Me.MouseIcon = ImageListMouse.ListImages("����").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "��ק��꣬�ڹ۲������ƶ�ͼ���λ�ã��Ա��ڸ��õع۲�"
        Case 3
            Me.MouseIcon = ImageListMouse.ListImages("����").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "��ק��꣬�ڹ۲�������С��Ŵ�ͼ��"
        Case 5
            Me.MouseIcon = ImageListMouse.ListImages("��ѡ����").Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "��ק��꣬��ѡ����Ҫ�Ŵ�������ɿ�����������"
        Case 6
            Me.MouseIcon = ImageListMouse.ListImages("�ü�").Picture
            '��Ϊ�ü��Ĺ����У��ƶ��ü���ʱ��ʹ�����ĸ���������ָ��
            If Me.MousePointer = 0 Then Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "��ק��꣬��ѡ����Ҫ�ü�������˫��ͼ����вü�"
        Case 7
            Me.MouseIcon = ImageListMouse.ListImages(pstrSideMarker).Picture
            Me.MousePointer = 99
            sbStatusBar.Panels(2).Text = "������꣬��עѡ����λ����"
    End Select
    
    '���²˵���ʾ״̬
    Select Case control.Id
        Case ID_frmFilm_TakePictures
            'û��ͼ���ʱ�����ఴťҪ�ҵ�
            control.Enabled = IIf(imgsPrint.Count >= 1, True, False)
            If control.Enabled = True Then control.Enabled = Not mblnPrinting
        Case ID_frmFilm_Label
            If pstrSideMarker = "" And control.Caption <> "��ע" Then
                control.Caption = "��ע"
                control.SetFocus
            ElseIf control.Caption <> pstrSideMarker And pstrSideMarker <> "" Then
                control.Caption = pstrSideMarker
                control.SetFocus
            End If
        Case ID_frmFilm_UnDivide
            '�ϲ����԰�ť����ȷ��������,�������״̬������
            control.Visible = mblnInTest
    End Select
    
    '���Ĺ�������ʾ
    If VScro.Max > 1 Then
        VScro.Visible = True
    Else
        VScro.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '�������루zl9PacsWork test��,����ʾ���ϲ����ԡ���ť
    Static strPass As String
    
    TimePass.Enabled = False
    If KeyCode = vbKeyF12 And Shift = 7 Then
        strPass = ""
        Exit Sub
    End If
    
    If KeyCode = vbKeyEscape Then
        Call subCheckToolBar(Nothing)
        intMouseState = 0
    End If
    
    If KeyCode <> vbKeyReturn Then
        '��¼��ǰ������ַ�
        If InStr(1, "1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyCode))) <> 0 Then
            strPass = strPass & UCase(Chr(KeyCode))
        End If
        
        '������ַ�=���룬����ʾ�ϲ����԰�ť
        If strPass = "ZL9PACSWORK TEST" Then
            mblnInTest = True
        Else
            mblnInTest = False
        End If
    End If
    
    TimePass.Enabled = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSQL As String
    
    mintFilmWidth = 11
    mintFilmHeight = 14
    mblnIsPortrait = True
    mblnIsRow = True
    mintPrintFilmCount = 0
    
    '��ʼ�����״̬�����������������������������ͬ������ʼ״̬
    '��¼����״̬��0���ޣ�1��������2�����Σ�3������;4-��;5-��ѡ����;6-�ü�:7-���ֱ�ע����frmFilmView�����н�������
    '����102������103������104��
    If cMouseUsage("102").lngMouseKey = 1 And Button_miWidthLevel Then  '����
        intMouseState = 1
    ElseIf cMouseUsage("103").lngMouseKey = 1 And Button_miCruise Then  '����
        intMouseState = 2
    ElseIf cMouseUsage("104").lngMouseKey = 1 And Button_miZoom Then  '����
        intMouseState = 3
    Else
        intMouseState = 0
    End If
    
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    mdblCutOutRatio = 0
    mblnPrinting = False
    mblnInTest = False
    mblnPrinted = False      'Ĭ��û�б���ӡ��
    
    '��ȡ����λ��
    Call RestoreWinState(Me, App.ProductName)
    
    '��ȡ������λ��
    mintTBMainPosition = Val(GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "��������λ��", 0))
    If mintTBMainPosition < 0 Or mintTBMainPosition > 3 Then mintTBMainPosition = 0
    mintTBImageProcessPosition = Val(GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "ͼ�����������λ��", 2))
    If mintTBImageProcessPosition < 0 Or mintTBImageProcessPosition > 3 Then mintTBImageProcessPosition = 0
    
    '�����˵�
    Call CreateBar
    '����״̬��ͼ��
    'Set sbStatusBar.Panels(1).Picture = f.ImgList24.ListImages("����ͼ��").Picture
    sbStatusBar.Panels(2).Text = "������ʾ"
    sbStatusBar.Panels(3).Text = "ҳ����"
    
    '����1  �ƽ�
    '���ò˵��������б��������
    With Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True)
        If blLocalRun = True Then
            strSQL = "SELECT ����ʶ as ���� FROM Ӱ��Ƭ���"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT ���� FROM Ӱ��Ƭ���"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
    
        While Not rsTemp.EOF
            .AddItem rsTemp!����
            rsTemp.MoveNext
        Wend
        If .ListCount > 1 Then
            .ListIndex = 1
            CommBar_Film_Execute Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True)
        End If
    End With
    
    With Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True)
        If blLocalRun = True Then
            strSQL = "SELECT ��ʽ��ʶ as ���� FROM Ӱ���ӡ��ʽ"
            Set rsTemp = cnAccess.Execute(strSQL, , adCmdText)
        Else
            strSQL = "SELECT ���� FROM Ӱ���ӡ��ʽ"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        End If
        While Not rsTemp.EOF
            .AddItem rsTemp!����
            rsTemp.MoveNext
        Wend
    End With
    
    With Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True)
        If cDICOMPrinter.Count > 0 Then
            For i = 1 To cDICOMPrinter.Count
                .AddItem cDICOMPrinter(i).strname
            Next
            .ListIndex = 0
        End If
    End With
    
    mintCellSpacing = 35

    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmSize", "14INX17IN")
    Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmFormat", "STANDARD\1,1")
    Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text = GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "PrinterName", "")
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = IIf(GetSetting("ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmPortrait", True) = "True", True, False)
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmRow, , True).Checked = Not Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    mblnIsPortrait = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    
    '��ʼ����ǰ��ҳ�沼��
    Call InitPageFormat(Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
    
    '��ʾ��һҳ����ҳ
    Call subShowOnePage(1)
    
    mblnCheckPrinter = IIf(GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "����ӡ��״̬", "0") = 1, True, False)
    
    mblnClearAfterPrint = IIf(GetSetting("ZLSOFT", "����ģ��\zlPacsCore", "��ӡ�����", "1") = 0, False, True)
    
    Me.VScro.Min = 1
    Me.VScro.Value = 1
     
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mblnBegin = True
    
End Sub

Private Sub subLoadViewer(intPage As Integer)
'------------------------------------------------
'���ܣ�����һҳViewer��ж�ض����Viewer��������intViewerCount����������װ��Viewer��
'������ intPage --- ��Ҫ��ʾ��ҳ��
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim intViewerCount As Integer
    
    If intPage > UBound(marrPages) Then Exit Sub
    
    On Error GoTo err
    
    intViewerCount = marrPages(intPage).intViewerCount
    
    'ж�ض����viewer
    For i = intViewerCount + 1 To FilmViewer.Count - 1
        Unload FilmViewer(i)
    Next
    
    '����װ��ȱ�ٵ�viewer
    For i = FilmViewer.Count To intViewerCount
        load FilmViewer(i)
        FilmViewer(i).Visible = True
        FilmViewer(i).CellSpacing = 2
    Next
    
    '���ԭ����ͼ������ÿ��Viewer��ͼ��������
    For i = 1 To FilmViewer.Count - 1
        FilmViewer(i).Images.Clear
        FilmViewer(i).MultiColumns = marrPages(intPage).ViewerLayout(i).intColumns
        FilmViewer(i).MultiRows = marrPages(intPage).ViewerLayout(i).intRows
    Next i
    
    '�����������в���
    Call subFillPageRCCount(marrPages(intPage).strPageFormat)
    
    '���°ڷ�Viewer
    Call picBak_Resize
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'ֻ�йر��������ʱ�򣬲�ж�ؽ�Ƭ��ӡ���壬���������ֻ�����ؽ�Ƭ��ӡ����
    If UnloadMode <> vbFormOwner Then
        Cancel = 1
        Me.Hide
        f.blnPrintFilm = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ''''''''''''''''''''''���������'''''''''''''''''''''''''''''''''
     '    ж��hook
    Call FilmUnhook(Me.hwnd, plngFilmPreWndProc)
    
    '�����������
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmSize", Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmFormat", Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "PrinterName", Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "FilmPortrait", Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked
    
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
    
    '���湤����λ��
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "��������λ��", Me.CommBar_Film(2).Position
    SaveSetting "ZLSOFT", "����ģ��\" & App.EXEName & "\frmFilm", "ͼ�����������λ��", Me.CommBar_Film(3).Position
    
    f.blnPrintFilm = False
    mblnBegin = False
    imgsPrint.Clear
    ReDim marrPages(0)
    
    
End Sub


Private Sub mfrmFilmView_AfterClose(dcmImage As DicomObjects.DicomImage, intViewerIndex As Integer, intImageIndex As Integer)
    '�ر�ͼ�����ڣ��Ѵ���õ�ͼ��ָ�����ǰ��ƬԤ����
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    On Error GoTo err
    If intViewerIndex = 0 Or intImageIndex = 0 Then Exit Sub
    If FilmViewer.Count < intViewerIndex Then Exit Sub
    If FilmViewer(intViewerIndex).Images.Count < intImageIndex Then Exit Sub
        
    FilmViewer(intViewerIndex).Images.Remove (intImageIndex)
    FilmViewer(intViewerIndex).Images.Add dcmImage
    Call FilmViewer(intViewerIndex).Images.Move(FilmViewer(intViewerIndex).Images.Count, intImageIndex)
    
    '����ͼ���λ�ú����ű���
    If dcmImage.StretchToFit = False Then
        lngWidth = mfrmFilmView.dcmViewer.width / mfrmFilmView.dcmViewer.MultiColumns
        lngHeight = mfrmFilmView.dcmViewer.height / mfrmFilmView.dcmViewer.MultiRows
        Call subScaleImage(FilmViewer(intViewerIndex).Images(intImageIndex), FilmViewer(intViewerIndex), lngWidth, lngHeight)
    End If
    
    'ж��mfrmFilmView����
    Set mfrmFilmView = Nothing
    
    '��ղü����
    mintCutOutViewer = 0
    mintCutOutImage = 0
    mintCutOutLabel = 0
    
    '�Ѹ����ύ��ԭʼͼ��
    '�ѱ�ע����ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
    
    'ͬ��ͼ��
    mintSelectedViewer = intViewerIndex
    mintSelectedImage = intImageIndex
    Set SelectedImage = FilmViewer(intViewerIndex).Images(intImageIndex)
    Call subSynchronalImg(False, IMG_SYN_All)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picBak_Resize()

'����һ����������Viewerλ�õĹ��̣����漰��Viewer�ļ��أ�ͼ��ļ��ص�

    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, t As Integer, w As Integer, h As Integer
    
    If Not mblnBegin Then Exit Sub
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�ȵ�����������λ��
    VScro.Move Me.picBak.ScaleWidth - VScro.width, 0, VScro.width, Me.picBak.ScaleHeight
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '���㽺Ƭ����Ļ�ĳ������߿�����
    mdubFilmRate = IIf(mblnIsPortrait, mintFilmWidth / mintFilmHeight, mintFilmHeight / mintFilmWidth)
    mdubScreenRate = (picBak.ScaleWidth - Me.VScro.width) / picBak.ScaleHeight
  
    '���ر�����picFilm
    picFilm.Visible = False
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '����picFilm��λ��
    If mdubFilmRate < mdubScreenRate Then
        picFilm.top = 0
        picFilm.height = picBak.ScaleHeight
        picFilm.width = picFilm.height * mdubFilmRate '- Me.VScro.width
        picFilm.left = Abs(picBak.ScaleWidth - picFilm.width - 250) / 2
    Else
        picFilm.left = 0
        picFilm.width = Abs(picBak.ScaleWidth - Me.VScro.width)
        picFilm.height = picFilm.width / mdubFilmRate
        picFilm.top = Abs(picBak.ScaleHeight - picFilm.height) / 2
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '�ڷ�Viewer��λ��
    k = 1
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 1 To UBound(marrRCCount)
        If mblnIsRow Then
            For j = 1 To marrRCCount(i)
                h = Abs(picFilm.ScaleHeight / UBound(marrRCCount) - mintCellSpacing * 2)
                w = Abs(picFilm.ScaleWidth / marrRCCount(i) - mintCellSpacing * 2)
                l = Abs(picFilm.ScaleWidth / marrRCCount(i) * (j - 1) + mintCellSpacing)
                t = Abs(picFilm.ScaleHeight / UBound(marrRCCount) * (i - 1) + mintCellSpacing)
                If Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked Then
                    If h > w Then
                         t = t + (h - w) / 2
                         h = w
                    Else
                         l = l + (w - h) / 2
                         w = h
                    End If
                End If
                FilmViewer(k).Move l, t, w, h
                k = k + 1
            Next
        Else
            For j = 1 To marrRCCount(i)
                h = Abs(picFilm.ScaleHeight / marrRCCount(i) - mintCellSpacing * 2)
                w = Abs(picFilm.ScaleWidth / UBound(marrRCCount) - mintCellSpacing * 2)
                l = Abs(picFilm.ScaleWidth / UBound(marrRCCount) * (i - 1) + mintCellSpacing)
                t = Abs(picFilm.ScaleHeight / marrRCCount(i) * (j - 1) + mintCellSpacing)
                If Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked Then
                    If h > w Then
                         t = t + (h - w) / 2
                         h = w
                    Else
                         l = l + (w - h) / 2
                         w = h
                    End If
                End If
                FilmViewer(k).Move l, t, w, h
                k = k + 1
            Next
        End If
    Next
    
    For i = 1 To FilmViewer.Count - 1
        '��ʾͼ��ѡ���
        Call subReScaleViewerFrame(FilmViewer(i))
        FilmViewer(i).Visible = True
    Next i
    
    picFilm.Visible = True
End Sub

Private Sub subLoadPrintImage(intPage As Integer)
'------------------------------------------------
'���ܣ���ʾһҳͼ����Viewer��˳��װ��imgsPrint�д洢��ͼ��
'������ intPage -- ��ʾ��ҳ��
'���أ���
'------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intStart As Integer
    Dim blnUpLoad As Boolean
    
    On Error GoTo err
    
    For i = 1 To marrPages(intPage).intViewerCount
        FilmViewer(i).Images.Clear
        
        '������ʾͼ��ѡ���
        Call subFilmDispframe(FilmViewer(i))
    Next i
    
    'ѭ��ÿһ��Viewer�����ͼ��
    intStart = funGetStartImgNo(intPage, 1, 1)
    For i = 1 To marrPages(intPage).intViewerCount
        
        If intStart > imgsPrint.Count Then Exit For
        
        '�����������ͼ��Viewer
        For j = 1 To FilmViewer(i).MultiColumns * FilmViewer(i).MultiRows
            If intStart > imgsPrint.Count Then Exit For
            '���ͼ��
            FilmViewer(i).Images.Add imgsPrint(intStart)
            
            '����ͼ���λ�ú����ű���
            If FilmViewer(i).Images(j).StretchToFit = False Then
                If Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V���)) <> 0 And Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V�߶�)) <> 0 Then
                    Call subScaleImage(FilmViewer(i).Images(j), FilmViewer(i), _
                        Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V���)), _
                        Val(funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_V�߶�)))
                Else
                    '˵����ͼ���������ӵģ�TAG��û�м�¼Viewer�Ŀ�Ⱥ͸߶ȣ���Ҫ��ӣ����Ҫreload����̨ͼ����
                    Call funSetTagVal(FilmViewer(i).Images(j), TAG_V���, CStr(FilmViewer(i).width / FilmViewer(i).MultiColumns))
                    Call funSetTagVal(FilmViewer(i).Images(j), TAG_V�߶�, CStr(FilmViewer(i).height / FilmViewer(i).MultiRows))
                    blnUpLoad = True
                End If
            End If
            
            '����ͼ��ѡ����
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                Call subImageSelect(i, j, True)
            End If
                
            intStart = intStart + 1
        Next j
        
    Next i
    
    '���ñ���ǰ�ͱ�ѡ�е�ͼ��
    If FilmViewer.Count > 1 And FilmViewer(1).Images.Count > 0 Then
        Call subImageCurrent(1, 1, True)
    Else
        mintSelectedViewer = 0
        mintSelectedImage = 0
        Set SelectedImage = Nothing
        mblnPrinted = False      'û��ͼ�񣬴�ӡ������ó�Ĭ��ֵFalse
    End If
    
    '��ͼ����Ϣ�ش���������
    If blnUpLoad Then
        Call subReloadImgsPrint
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub TimePass_Timer()
    Call Form_KeyDown(vbKeyF12, 7)   '�����̬����
End Sub

Private Sub FilmViewer_DblClick(Index As Integer)
    
    '��ͼ����˫��ʱ���������÷���
    '1����ͼ������
    '2������ͼ��ü�
    'ʹ�òü�������������÷�
    If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then
        '��ͼ������
        Call subOpenFilmView
        
    Else    'ͼ��ü�
        If mintCutOutViewer >= FilmViewer.Count Then Exit Sub
        If mintCutOutImage > FilmViewer(mintCutOutViewer).Images.Count Then Exit Sub
        If mintCutOutLabel <> FilmViewer(mintCutOutViewer).Images(mintCutOutImage).Labels.Count Then Exit Sub
        
        Dim Image As DicomImage
        Dim i As Integer
        Dim lblTemp As DicomLabel
        Dim sourceImage As DicomImage
        
        Set sourceImage = FilmViewer(mintCutOutViewer).Images(mintCutOutImage)
        Set Image = CutOutAImage(sourceImage)
        
        Image.Name = "ZLPIC"
        'ɾ����ѡ�õ���ʱ��ע
        sourceImage.Labels.Remove mintCutOutLabel
        Set mdcmSelectLabel = Nothing
        
        Call subWriteDicomPara(sourceImage, Image)
        
        '��ԭ��ͼ��ı�ע����ӵ����ڵ�ͼ����
        Image.Labels.Clear
        For i = 1 To sourceImage.Labels.Count
            Image.Labels.Add sourceImage.Labels(i)
        Next i
        
        '�������ɵ�ͼ����ӵ�Viewer��
        If mintCutOutImage = 1 And FilmViewer(mintCutOutViewer).Images.Count = 1 Then
            FilmViewer(mintCutOutViewer).Images.Clear
            FilmViewer(mintCutOutViewer).Images.Add Image
        Else
            FilmViewer(mintCutOutViewer).Images.Remove mintCutOutImage
            FilmViewer(mintCutOutViewer).Images.Add Image
            FilmViewer(mintCutOutViewer).Images.Move FilmViewer(mintCutOutViewer).Images.Count, mintCutOutImage
        End If
        
        'ͼ�����Viewer�к�������ʾ��ߣ����ʱ���ߺ͵�λ����׼ȷ��
        Call UpdateRuler(Image, True)
        
        mintCutOutViewer = 0
        mintCutOutImage = 0
        mintCutOutLabel = 0
        Me.MousePointer = vbArrow
        
        '�ύ����ͼ��
        Call subReloadImgsPrint
    End If
End Sub

Private Sub FilmViewer_DragDrop(Index As Integer, Source As control, x As Single, y As Single)
    '������ק��ͼ��
    Dim intOldImgIndex As Integer
    Dim intOldViewerIndex As Integer
    Dim intOldImgsPrintIndex As Integer
    Dim intNewImgIndex As Integer
    Dim intImgIndex As Integer
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim intOldDelta As Integer
    Dim blnNewMoveUp As Boolean
    
    On Error GoTo err
    
    If Source.Name = "FilmViewer" And Source.Images.Count > 0 Then
        '��ȡͼ��ľ�λ��
        intOldImgIndex = Val(Source.Tag)
        If intOldImgIndex <= 0 Then Exit Sub
        
        intOldViewerIndex = Val(Source.Index)
        If intOldViewerIndex <= 0 Or intOldViewerIndex > FilmViewer.Count Then Exit Sub
        
        '��ȡ��ͼ���������е�λ��
        intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, intOldViewerIndex, intOldImgIndex)
        If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
            Exit Sub
        End If
                            
        '��ȡͼ�����λ��
        intImgIndex = FilmViewer(Index).ImageIndex(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
        intNewImgIndex = funGetStartImgNo(VScro.Value, Index, intImgIndex)
        
        If intOldImgsPrintIndex = intNewImgIndex And intImgIndex <> 0 Then Exit Sub
        
        '��ʼ�ڷ�ͼ��֮ǰ���Ȱ�ԭ��ͼ���TAG�ύ��ͼ��������
        Call subReloadImgsPrint
        
        '���°ڷ�ͼ��
        If intImgIndex = 0 Then     '˵�����ϵ��˿յĵط�����Ҫ����������һ��ͼ
            '���ԭͼ�Ƿ�ѡ�У�����Ǳ�ѡ�еģ������ѡ�ƶ�
            If funGetTagVal(FilmViewer(intOldViewerIndex).Images(intOldImgIndex).Tag, TAG_ѡ��) = "Select" Then
                '��ǰͼ��ѡ�У������ѡ�ƶ�
                For i = 1 To FilmViewer.Count - 1
                    For j = 1 To FilmViewer(i).Images.Count
                        If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                            Set img = New DicomImage
                            '��ȡ��ͼ���������е�λ��
                            intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, i, j)
                            If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
                                Exit For
                            End If
                            
                            Set img = imgsPrint(intOldImgsPrintIndex)
                            imgsPrint.Add img
                        End If
                    Next j
                Next i
            Else
                'ֻ����һ��ͼ
                Set img = New DicomImage
                Set img = imgsPrint(intOldImgsPrintIndex)
                imgsPrint.Add img
            End If
        Else    '����ͼ��λ��
            '��������ƣ���λ��Ҫ-1
            If intNewImgIndex > intOldImgsPrintIndex Then
'                intNewImgIndex = intNewImgIndex - 1
                blnNewMoveUp = True
            Else
                blnNewMoveUp = False
            End If
            If intNewImgIndex <= 0 Or intNewImgIndex > imgsPrint.Count Then Exit Sub
            
            
            '���ԭͼ�Ƿ�ѡ�У�����Ǳ�ѡ�еģ������ѡ�ƶ�
            If funGetTagVal(FilmViewer(intOldViewerIndex).Images(intOldImgIndex).Tag, TAG_ѡ��) = "Select" Then
                '��ǰͼ��ѡ�У������ѡ�ƶ�
                intOldDelta = 0
                For i = 1 To FilmViewer.Count - 1
                    For j = 1 To FilmViewer(i).Images.Count
                        If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                            '����ͼ��ԭ����˳���ƶ���ѡ�е�ͼ��
                            
                            '��ȡ��ͼ���������е�λ��
                            intOldImgsPrintIndex = funGetStartImgNo(VScro.Value, i, j)
                            
                            If intOldImgsPrintIndex < intNewImgIndex Then
                                intOldImgsPrintIndex = intOldImgsPrintIndex + intOldDelta
                            End If
                            
                            If intOldImgsPrintIndex <= 0 Or intOldImgsPrintIndex > imgsPrint.Count Then
                                Exit For
                            End If
                            
                            '�����ǰ�ƣ�����֮ǰ�Ѿ����������Ƶ�Ч��������Ҫ+1
                            If (intNewImgIndex < intOldImgsPrintIndex) And blnNewMoveUp Then
                                intNewImgIndex = intNewImgIndex + 1
                                blnNewMoveUp = False
                            ElseIf (intNewImgIndex > intOldImgsPrintIndex) And blnNewMoveUp = False Then
                                intNewImgIndex = intNewImgIndex - 1
                                blnNewMoveUp = True
                            End If
                            
                            If intNewImgIndex <> intOldImgsPrintIndex Then
                                '�ƶ�ͼ��
                                Call imgsPrint.Move(intOldImgsPrintIndex, intNewImgIndex)
                                If intOldImgsPrintIndex > intNewImgIndex Then
                                    intNewImgIndex = intNewImgIndex + 1
                                Else
                                    intOldDelta = intOldDelta - 1
                                End If
                            End If
                        End If
                    Next j
                Next i
            Else
                '�ƶ�ͼ��
                Call imgsPrint.Move(intOldImgsPrintIndex, intNewImgIndex)
            End If
        End If
        '������ʾͼ��
        Call subShowPrintImages(Me.VScro.Value)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FilmViewer_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '����Del
    If KeyCode = 46 Then        'Delete
        Call subDelImage
    End If
End Sub

Private Sub FilmViewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim intImgIndex As Integer
    Dim ls As DicomLabels
    
    On Error GoTo err
    
    mintBaseX = x
    mintBaseY = y
    
    '�л���ͼ��
    If Index >= FilmViewer.Count Then Exit Sub
    intImgIndex = FilmViewer(Index).ImageIndex(x, y)
    If FilmViewer(Index).Images.Count > 0 And intImgIndex <> 0 Then
        '�л�ͼ��,�Ȼָ�ԭ��ͼ���ѡ���
        If mintSelectedViewer > 0 And mintSelectedViewer < FilmViewer.Count Then
            If mintSelectedImage > 0 And mintSelectedImage <= FilmViewer(mintSelectedViewer).Images.Count Then
                If funGetTagVal(FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Tag, TAG_ѡ��) = "Select" Then
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, True)
                Else
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, False)
                End If
                FilmViewer(mintSelectedViewer).Refresh
            End If
        End If
        
        '�����õ�ǰͼ���ѡ���
        Call subImageCurrent(Index, intImgIndex, True)
        
        '���ô���λ�����˵�
        Call subSetWidthLevelF(SelectedImage, Me)
        
        If Button = 1 Then
            'intMouseState ����״̬��0���ޣ�1��������2�����Σ�3������;4-��;5-��ѡ����;6-�ü�:7-���ֱ�ע
            If intMouseState = 6 Then '�ü�
                '�ü�״̬�µ����down�������ֲ�����1�����ü��򣨼�¼��ǣ���2���ƶ��ü���(�н���) ��3��˫�����вü�
                If mintCutOutViewer = 0 Or mintCutOutImage = 0 Or mintCutOutLabel = 0 Then  '���ü���
                    '�ж��ǹ̶��ü����������ɲü�
                    If (CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId <> ID_frmFilm_CutOut_Custom) _
                        And (CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId <> ID_frmFilm_CutOut) Then
                        '�̶��ü�
                        Call subCutOutRatio(CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId)
                    Else
                        '���ӿ�ѡ��ע
                        FilmViewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                        Set mdcmSelectLabel = FilmViewer(Index).Images(intImgIndex).Labels(FilmViewer(Index).Images(intImgIndex).Labels.Count)
                        mdcmSelectLabel.Tag = CUT_LABEL
                        mblnDcmViewDown = True
                        mintCutOutViewer = Index
                        mintCutOutImage = intImgIndex
                        mintCutOutLabel = FilmViewer(Index).Images(intImgIndex).Labels.Count
                    End If
                Else            '��ʼ�ƶ��ü���
                    Set ls = FilmViewer(Index).LabelHits(x, y, False, False, True)
                    If ls.Count <> 0 And Me.MousePointer <> vbArrow Then
                        '��ʼ�ƶ��ü���
                        If ls(1).Tag = CUT_LABEL And SelectedImage.Labels(SelectedImage.Labels.Count).Tag = CUT_LABEL Then
                            mblnLabelMoving = True
                        End If
                    End If
                End If
            End If
            If intMouseState = 5 Then      '��ѡ����
                '���ӿ�ѡ��ע
                FilmViewer(Index).Images(intImgIndex).Labels.Add GetNewLabel(doLabelRectangle, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                Set mdcmSelectLabel = FilmViewer(Index).Images(intImgIndex).Labels(FilmViewer(Index).Images(intImgIndex).Labels.Count)
                mblnDcmViewDown = True
            End If
            If intMouseState = 7 Then       '���ֱ�ע
                Dim dcmLabel As DicomLabel
                Set dcmLabel = GetNewLabel(doLabelText, FilmViewer(Index).ImageXPosition(x, y), FilmViewer(Index).ImageYPosition(x, y), 0, 0)
                FilmViewer(Index).Images(intImgIndex).Labels.Add dcmLabel
                dcmLabel.AutoSize = True
                dcmLabel.Margin = 0
                dcmLabel.Text = pstrSideMarker
                dcmLabel.Shadow = doShadowAll
                dcmLabel.ShowTextBox = True
                dcmLabel.Font.Bold = True
                dcmLabel.Tag = POSTURE_LABEL
                intMouseState = 0
                pstrSideMarker = ""
                '�ѱ�ע����ͼ���ϴ���ԭʼͼ������
                Call subReloadImgsPrint
            End If
            'Ctrl����ѡ��
            If Shift = 2 Then
                If funGetTagVal(FilmViewer(Index).Images(intImgIndex).Tag, TAG_ѡ��) = "Select" Then
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, False)
                Else
                    Call subImageSelect(mintSelectedViewer, mintSelectedImage, True)
                End If
                
                '��ѡ���ͼ���ϴ���ԭʼͼ������
                Call subReloadImgsPrint
            End If
            
            'Shift����ѡ��
            If Shift = 1 Then
                Call subShiftSelect(Index, intImgIndex)
            End If
            
            If intMouseState = 0 Then   '������κ�״̬����ʼ��ק
                If FilmViewer(Index).Images.Count > 0 Then
                    'tag ����һ���ֶΣ�Viewer����ͼ�����ڵ�����
                    FilmViewer(Index).Tag = FilmViewer(Index).ImageIndex(x, y)
                    FilmViewer(Index).Drag
                End If
            End If
        End If
        FilmViewer(Index).Refresh
    End If
    Exit Sub
err:
End Sub

Private Sub SelAllImage(blnSelect As Boolean)
'------------------------------------------------
'���ܣ�ȫѡ����ȫ��ѡ����ͼ����������ͼ���ѡ������ɫ
'       ��ѡ�е�ͼ��߿�Ϊ��ɫ��û�б�ѡ�е�ͼ��߿�Ϊ��ɫ
'������ blnSelect -- True ѡ��ͼ��False ȫ��ѡ��
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                Call subImageSelect(i, j, blnSelect)
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '��ѡ���ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FilmViewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim dblZoom As Double
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    
    On Error GoTo err
    
    If SelectedImage Is Nothing Then Exit Sub
    
    If (Button = 1 And intMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
        Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then  '����
        If SelectedImage.VOILUT = 1 Then SelectedImage.VOILUT = 0
        SelectedImage.width = SelectedImage.width + (x - mintBaseX) * lngWidthLevelStep / 5
        SelectedImage.Level = SelectedImage.Level + (y - mintBaseY) * lngWidthLevelStep / 5
        SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
        mintBaseX = x
        mintBaseY = y
        FilmViewer(Index).Refresh
    ElseIf (Button = 1 And intMouseState = 2) Or (Button = 4 And intMouseWheelDrag = 0) _
        Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) Then '����
        subCenterZoom SelectedImage, FilmViewer(Index), SelectedImage.ActualZoom
        SelectedImage.ScrollX = SelectedImage.ScrollX - (x - mintBaseX) * lngCruiseStep / 5
        SelectedImage.ScrollY = SelectedImage.ScrollY - (y - mintBaseY) * lngCruiseStep / 5
        mintBaseX = x
        mintBaseY = y
    ElseIf (Button = 1 And intMouseState = 3) Or (Button = 4 And intMouseWheelDrag = 1) _
        Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then '����
        '���ŵ�λ��0.01��
        dblZoom = SelectedImage.ActualZoom * (1 + (mintBaseY - y) * lngZoomStep / 5 * 0.001)
        If dblZoom < 0.01 Then dblZoom = 0.01
        If dblZoom > 64 Then dblZoom = 64
        subCenterZoom SelectedImage, FilmViewer(Index), dblZoom
        
        If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
                UpdateRuler SelectedImage, True
            End If
        End If
        
        mintBaseX = x
        mintBaseY = y
    ElseIf Button = 1 And (intMouseState = 5 Or intMouseState = 6) Then  '��ѡ����
        If mblnDcmViewDown = True Then
            mdcmSelectLabel.width = FilmViewer(Index).ImageXPosition(x, y) - mdcmSelectLabel.left
            mdcmSelectLabel.height = FilmViewer(Index).ImageYPosition(x, y) - mdcmSelectLabel.top
            FilmViewer(Index).Refresh
        End If
    End If
    
    If intMouseState = 6 And mintCutOutViewer <> 0 And mintCutOutImage <> 0 And mintCutOutLabel <> 0 Then
        Set ls = FilmViewer(Index).LabelHits(x, y, False, False, True)
        If Button = 1 Then          '��걻����
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(FilmViewer(Index), SelectedImage, x, y)
                Set lblCUT = SelectedImage.Labels(SelectedImage.Labels.Count)
                
                If (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then       '�����ƶ�
                    
                    lngXOffset = (FilmViewer(Index).ImageXPosition(x, y) - FilmViewer(Index).ImageXPosition(mintBaseX, mintBaseY))
                    If Abs(lblCUT.left - FilmViewer(Index).ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - FilmViewer(Index).ImageXPosition(x, y)) Then '�ұߵ��ƶ�
                            lblCUT.width = lblCUT.width + lngXOffset
                    Else    '������ƶ�
                            lblCUT.left = lblCUT.left + lngXOffset
                            lblCUT.width = lblCUT.width - lngXOffset
                    End If
                    If mdblCutOutRatio <> 0 Then    '���̶ֹ�����
                        lblCUT.height = lblCUT.width / mdblCutOutRatio
                    End If
                ElseIf (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then   '�����ƶ�
                    
                    lngYOffset = (FilmViewer(Index).ImageYPosition(x, y) - FilmViewer(Index).ImageYPosition(mintBaseX, mintBaseY))
                    If Abs(lblCUT.top - FilmViewer(Index).ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - FilmViewer(Index).ImageYPosition(x, y)) Then    '�����ߵ��ƶ�
                        lblCUT.height = lblCUT.height + lngYOffset
                        
                    Else    '�������ƶ�
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                    If mdblCutOutRatio <> 0 Then    '���̶ֹ�����
                        lblCUT.width = lblCUT.height * mdblCutOutRatio
                    End If
                ElseIf Me.MousePointer = vbSizePointer Then     '�����ƶ�
                
                    lngXOffset = (FilmViewer(Index).ImageXPosition(x, y) - FilmViewer(Index).ImageXPosition(mintBaseX, mintBaseY))
                    lngYOffset = (FilmViewer(Index).ImageYPosition(x, y) - FilmViewer(Index).ImageYPosition(mintBaseX, mintBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                mintBaseX = x
                mintBaseY = y
                FilmViewer(Index).Refresh
            End If
        ElseIf Button = 0 Then
            If ls.Count <> 0 Then
                If Abs(ls(1).left - FilmViewer(Index).ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - FilmViewer(Index).ImageXPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeNS
                    Else
                        Me.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - FilmViewer(Index).ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - FilmViewer(Index).ImageYPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeWE
                    Else
                        Me.MousePointer = vbSizeNS
                    End If
                Else
                    Me.MousePointer = vbSizePointer
                End If
            Else
                Me.MousePointer = vbArrow
            End If
        End If
    End If
    Exit Sub
err:
End Sub

Private Sub FilmViewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    
    On Error GoTo err
    
    If Button = 1 Then
        If intMouseState <> 0 Then
            If intMouseState = 5 And mblnDcmViewDown Then    '��ѡ����
                lngLeft = SelectedImage.Labels(SelectedImage.Labels.Count).left * SelectedImage.ActualZoom
                lngTop = SelectedImage.Labels(SelectedImage.Labels.Count).top * SelectedImage.ActualZoom
                lngWidth = SelectedImage.Labels(SelectedImage.Labels.Count).width * SelectedImage.ActualZoom
                lngHeight = SelectedImage.Labels(SelectedImage.Labels.Count).height * SelectedImage.ActualZoom
                
                '�������
                If lngWidth < 0 Then
                    lngLeft = lngLeft + lngWidth
                    lngWidth = -lngWidth
                End If
                
                If lngHeight < 0 Then
                    lngTop = lngTop + lngHeight
                    lngHeight = -lngHeight
                End If
                
                RectangleZoom FilmViewer(Index), SelectedImage, lngLeft, lngTop, lngWidth, lngHeight
                
                'ɾ����ѡ�õ���ʱ��ע
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                Set mdcmSelectLabel = Nothing
                FilmViewer(Index).Refresh
            ElseIf intMouseState = 6 Then
                If mblnDcmViewDown Then       '�ü�
                    '�����κβ���
                    '����ü���Ϊ0 ����ȡɾ���ü�������ü��ı��
                    If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                        'ɾ����ѡ�õ���ʱ��ע
                        SelectedImage.Labels.Remove SelectedImage.Labels.Count
                        Set mdcmSelectLabel = Nothing
                        FilmViewer(Index).Refresh
                        
                        mintCutOutViewer = 0
                        mintCutOutImage = 0
                        mintCutOutLabel = 0
                    End If
                End If
            End If
            'ͼ��Ϊ��
        End If
    End If
    
    'ͬ��,''intMouseState��0���ޣ�1��������2�����Σ�3������;4-��;5-��ѡ����;6-�ü�:7-���ֱ�ע
    If FilmViewer(Index).Images.Count > 0 Then
        If (Button = 1 And intMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
            Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then   '����
            Call subSynchronalImg(False, IMG_SYN_WINDOW)
        ElseIf (Button = 1 And (intMouseState = 2 Or intMouseState = 3 Or intMouseState = 5)) _
            Or (Button = 4 And (intMouseWheelDrag = 0 Or intMouseWheelDrag = 1)) _
            Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) _
            Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then
            Call subSynchronalImg(False, IMG_SYN_ZOOMPAN)
        End If
    End If
            
    mblnDcmViewDown = False
    mblnLabelMoving = False
    Exit Sub
err:
End Sub

Private Sub VScro_Change()
    
    If Me.VScro.Value = 0 Then
        sbStatusBar.Panels(3).Text = "ҳ����"
        Exit Sub
    End If
    
    If UBound(marrPages) = 0 Then Exit Sub
    
    '����ǰҳ�ĸ�ʽ���óɹ������˵�
    If marrPages(Me.VScro.Value).strPageFormat <> "" Then
        Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text = marrPages(Me.VScro.Value).strPageFormat
        Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Caption = marrPages(Me.VScro.Value).strPageFormat
    End If
    
    '������ʾ��ǰҳ��ͼ��
    Call subShowOnePage(Me.VScro.Value)
    
    sbStatusBar.Panels(3).Text = "ҳ����" & VScro.Value & "/" & VScro.Max
End Sub

Private Sub subFilmDispframe(v As DicomViewer)
'------------------------------------------------
'���ܣ���Viewer����ʾ���ο�
'������v������Ҫ��ʾ���ο��Viewer
'���أ���
'------------------------------------------------
    Dim w As Integer, h As Integer
    Dim l As DicomLabel
    Dim i As Integer
    
    v.Labels.Clear
    For i = 1 To v.MultiColumns * v.MultiRows
        w = v.width / Screen.TwipsPerPixelX / v.MultiColumns - 2
        h = v.height / Screen.TwipsPerPixelY / v.MultiRows - 2
        Set l = New DicomLabel
        l.LabelType = 2     '���α�ע
        l.width = w
        l.height = h
        l.left = ((i - 1) Mod v.MultiColumns) * (w + 2) + 1
        l.top = ((i - 1) \ v.MultiColumns) * (h + 2) + 1
        v.Labels.Add l
    Next i
    v.Refresh
End Sub

Private Sub subReScaleViewerFrame(v As DicomViewer)
'------------------------------------------------
'���ܣ�����Viewer������ο��λ��
'������v������Ҫ��ʾ���ο��Viewer
'���أ���
'------------------------------------------------
    Dim w As Integer, h As Integer
    Dim l As DicomLabel
    Dim i As Integer
    
    If v.Labels.Count <> v.MultiColumns * v.MultiRows Then
        '������ο���������ԣ������´���
        Call subFilmDispframe(v)
    Else
        For i = 1 To v.MultiColumns * v.MultiRows
            w = v.width / Screen.TwipsPerPixelX / v.MultiColumns - 2
            h = v.height / Screen.TwipsPerPixelY / v.MultiRows - 2
            Set l = v.Labels(i)
            l.LabelType = 2     '���α�ע
            l.width = w
            l.height = h
            l.left = ((i - 1) Mod v.MultiColumns) * (w + 2) + 1
            l.top = ((i - 1) \ v.MultiColumns) * (h + 2) + 1
        Next i
        v.Refresh
    End If
    
End Sub


Private Function subPrintFilm(clsOnePrinter As clsDicomPrint) As Boolean
'------------------------------------------------
'���ܣ�����ӡ�����ʹ�ӡ�źţ���ͼ���͸���ӡ����
'������clsOnePrinter������¼��ӡ����������
'���أ���
'�ϼ���������̣�CommBar_Film_Execute
'�¼���������̣���
'���õ��ⲿ������intViewerCount
'�����ˣ��ƽ�
'------------------------------------------------
    
    'ͼ�����ڵش� k = 1 To intViewerCount�� viewer(k)��
    
    '�ж�ͼ�������Ƿ���ڵ���1��û��ͼ����ֱ���˳�
    If FilmViewer.Count <= 1 Then
        Exit Function
    End If
    
    If clsOnePrinter Is Nothing Then
        Exit Function
    End If
    
    Dim printer As New DicomPrint, Thisim As DicomImage
    Dim k As Integer, i As Integer, j As Integer
    Dim strSQL As String
    Dim StrPrintLog As String, StrPrintPage As String
    Dim strImageUIDS As String  '��¼ͼ���ʵ��UID����,�ָ�
    Dim intCurPage As Integer   '��¼��ǰ��ʾ��ҳ��
    Dim arrImageSize() As ImageSize
    Dim arrTempUIDs() As String
    Dim strTempUIDs As String
    
    printer.Node = clsOnePrinter.strIPAddress
    printer.Port = clsOnePrinter.lngPort
    printer.CallingAE = clsOnePrinter.strSCUAETitle
    printer.CalledAE = clsOnePrinter.strAETitle
    intCurPage = Me.VScro.Value
    
    On Error GoTo err1
    
    
    'ѭ����ӡÿһҳ
    strImageUIDS = ","
    For j = 1 To Me.VScro.Max
        If mintPageRange = 0 Or j = intCurPage Then
            
            StrPrintPage = "," '��¼ÿ�Ž�Ƭ�ϴ�ӡ������UID,����¼�ظ���
            '��������FilmSession�еĲ�����Ȼ����Open��ӡ��
            
            ''''''''''FilmSession �Ĳ���''''''''''''''''''''
            '��ӡ����������
            If clsOnePrinter.lngCopies <> 0 Then
                printer.Copies = clsOnePrinter.lngCopies
            Else
                printer.Copies = 1
            End If
            
            'Print Priority ���ȼ�����ѡ
            If clsOnePrinter.strPriority <> "" Then
                printer.Session.Attributes.Add &H2000, &H20, clsOnePrinter.strPriority
            End If
            'Medium Type �������ͣ���ѡ
            If clsOnePrinter.strMedium <> "" Then
                printer.Session.Attributes.Add &H2000, &H30, clsOnePrinter.strMedium
            End If
            'Film Destination ����Ŀ�꣬��ѡ
            If clsOnePrinter.strFilmBox <> "" Then
                printer.Session.Attributes.Add &H2000, &H40, clsOnePrinter.strFilmBox
            End If
            
            '''''''''''''''''''''''''''''''''�򿪴�ӡ��''''''''''''''''''''''
            'Open��ʱ�򣬻��FilmSession�Ĳ�����N-CREATE�ķ�ʽ������ӡ��
            printer.Open
            
            '���ô�����
            On Error GoTo err2
            
            '����ӡ�����ص�״̬
            If mblnCheckPrinter = True Then
                If Not printer.printer Is Nothing Then
                    '���Printer�ģ�2110��0010��Printer Status�ͣ�2110��0020��Printer Status Info
                    If printer.printer.Attributes(&H2110, &H10).Exists And Not IsNull(printer.printer.Attributes(&H2110, &H10).Value) Then
                        If printer.printer.Attributes(&H2110, &H10).Value = "WARNING" Or _
                            printer.printer.Attributes(&H2110, &H10).Value = "FAILURE" Then
                            
                            '���־�����ߴ���
                            If printer.printer.Attributes(&H2110, &H20).Exists And Not IsNull(printer.printer.Attributes(&H2110, &H20).Value) Then
                                'ͬʱ����Printer Status�� Printer Status Info�ľ�����ߴ�����Ϣ
                                err.Raise vbObjectError + 101, 100, "Printer Status = " & CStr(printer.printer.Attributes(&H2110, &H10).Value) _
                                    & " Printer Status Info = " & CStr(printer.printer.Attributes(&H2110, &H20).Value)
                            Else
                                'Printer Status���־�����ߴ��󣬵���û����ϸ��Ϣ��ֱ�ӷ��ش�����߾���
                                err.Raise vbObjectError + 101, 100, "Printer Status = " & CStr(printer.printer.Attributes(&H2110, &H10).Value)
                            End If
                              
                        End If
                    End If
                End If
            End If
            
            '��ʾָ��ҳ��ͼ��
            Call subShowOnePage(j)
        
            '''''''''''''''''''FilmBox�Ĳ���''''''''''''''''''''''''''''''''''
            '��Ƭ���򣬱���
            If clsOnePrinter.strOrientation <> "" Then
                printer.Orientation = clsOnePrinter.strOrientation
            Else
                printer.Orientation = "PORTRAIT"
            End If
            '��Ƭ��С������
            If clsOnePrinter.strFilmSize <> "" Then
                printer.FilmSize = clsOnePrinter.strFilmSize
            Else
                printer.FilmSize = "14INX17IN"
            End If
            
            '��ӡͼ���λ��������
            If clsOnePrinter.lngBitDepth <> 0 Then
                printer.BitDepth = clsOnePrinter.lngBitDepth
            Else
                printer.BitDepth = 8
            End If
            
            '�Ŵ�ʽ,����
            If clsOnePrinter.strMagnification <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H60, clsOnePrinter.strMagnification
            Else
                printer.FilmBox.Attributes.Add &H2010, &H60, "CUBIC"
            End If
            'Smoothing Type 'ƽ��,��ѡ
            If clsOnePrinter.strSmooth <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H80, clsOnePrinter.strSmooth
            End If
            'border density ��Ե�ܶȣ�����
            If clsOnePrinter.strBorderDensity <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H100, clsOnePrinter.strBorderDensity
            Else
                printer.FilmBox.Attributes.Add &H2010, &H100, "BLACK"    'border density
            End If
            'empty image density �հ��ܶȣ�����
            If clsOnePrinter.strEmptyDensity <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H110, clsOnePrinter.strEmptyDensity
            Else
                printer.FilmBox.Attributes.Add &H2010, &H110, "BLACK"   'empty image density
            End If
            'min density ��С�ܶȣ�����
        '    If clsOnePrinter.strMinDensity <> "" Then
        '        printer.FilmBox.Attributes.Add &H2010, &H120, clsOnePrinter.strMinDensity
        '    Else
        '        printer.FilmBox.Attributes.Add &H2010, &H120, 16
        '    End If
            'max density ����ܶ�,����
        '    If clsOnePrinter.strMaxDensity <> "" Then
        '        printer.FilmBox.Attributes.Add &H2010, &H130, clsOnePrinter.strMaxDensity
        '    Else
        '        printer.FilmBox.Attributes.Add &H2010, &H130, 320
        '    End If
            'trim whether the film will be cut in to 2 or more films ������Ƭ,����
            If clsOnePrinter.strTrim <> "" Then
                printer.FilmBox.Attributes.Add &H2010, &H140, clsOnePrinter.strTrim
            Else
                printer.FilmBox.Attributes.Add &H2010, &H140, "NO"           'trim whether the film will be cut in to 2 or more films
            End If
            'Polarity ����,��ѡ
            If clsOnePrinter.strPolarity <> "" Then
                printer.FilmBox.Attributes.Add &H2020, &H20, clsOnePrinter.strPolarity
            End If
            'Requested Resolution ID �ֱ��ʣ���ѡ
            If clsOnePrinter.strResolution <> "" Then
                printer.FilmBox.Attributes.Add &H2020, &H50, clsOnePrinter.strResolution
            End If
        
            '��ӡ��ʽ������
            If marrPages(j).strPageFormat <> "" Then
                printer.Format = marrPages(j).strPageFormat
            Else
                printer.Format = "STANDARD\1,2"
            End If
            
            '����ÿһ��ͼ�����ֱ���
            Call subCalImageMaxSize(printer.FilmSize, printer.Format, clsOnePrinter.intImageResolution, arrImageSize)
            
            For k = 1 To (FilmViewer.Count - 1)
                If FilmViewer(k).Images.Count > 0 Then
                    Set Thisim = funAssembleImage(FilmViewer(k), arrImageSize(k).intWidth, arrImageSize(k).intHeight)
                    If Not Thisim Is Nothing Then
                        printer.PrintImage Thisim, False, True
                        For i = 1 To FilmViewer(k).Images.Count
                            If InStr(1, StrPrintPage, "," & FilmViewer(k).Images(i).SeriesUID & ",") <= 0 Then
                                StrPrintPage = StrPrintPage & FilmViewer(k).Images(i).SeriesUID & ","
                            End If
                            If InStr(1, strImageUIDS, "," & FilmViewer(k).Images(i).InstanceUID & ",") <= 0 Then
                                strImageUIDS = strImageUIDS & FilmViewer(k).Images(i).InstanceUID & ","
                            End If
                        Next
                    End If
                End If
            Next
            printer.PrintFilm
            printer.Close
            StrPrintLog = StrPrintLog & "|" & Mid(StrPrintPage, 2, Len(StrPrintPage) - 1)
        End If
    Next j
    
    StrPrintLog = Mid(StrPrintLog, 2) '��¼��Ƭʹ�����������Ǽ��UID�Ѵ�ӡ
    StrPrintPage = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
    For i = 0 To UBound(Split(StrPrintLog, "|"))
        strSQL = "Zl_��Ƭ��ӡ��¼_Insert('" & Split(StrPrintLog, "|")(i) & "','" & StrPrintPage & "')"
        zlDatabase.ExecuteProcedure strSQL, "��¼��Ƭʹ��"
    Next
    
    '��¼ͼ���ӡ��ǣ�2000���ַ��ͱ���һ�Σ�����һ�δ�ӡͼ��̫�࣬����oracle4000������
    If Len(strImageUIDS) > 2000 Then
        arrTempUIDs = Split(strImageUIDS, ",")
        strTempUIDs = ","
        For i = 1 To UBound(arrTempUIDs)
            strTempUIDs = strTempUIDs & arrTempUIDs(i) & ","
            If Len(strTempUIDs) > 2000 Then
                strSQL = "Zl_Ӱ��ͼ��Ƭ��ӡ_Update('" & strTempUIDs & "',1)"
                zlDatabase.ExecuteProcedure strSQL, "��¼ͼ��Ƭ��ӡ���"
                strTempUIDs = ","
            End If
        Next i
    Else
        strTempUIDs = strImageUIDS
    End If
    If Len(strTempUIDs) > 1 Then
        strSQL = "Zl_Ӱ��ͼ��Ƭ��ӡ_Update('" & strTempUIDs & "',1)"
        zlDatabase.ExecuteProcedure strSQL, "��¼ͼ��Ƭ��ӡ���"
    End If
    
    '������ӡ����¼�
    RaiseEvent AfterPrinted(strImageUIDS)
    
    subPrintFilm = True
    Exit Function
err1:
    MsgBox "��ӡ�����Ӵ���,�����ӡ������������." & vbCrLf & "��ӡ����Ϊ��" & clsOnePrinter.strname & " IPΪ:" _
            & clsOnePrinter.strIPAddress & " �˿�Ϊ��" & clsOnePrinter.lngPort & " ������룺 " & err.Number _
            & " ���������� " & err.Description, vbExclamation, gstrSysName, Me
    Exit Function
err2:
    If err.Number = vbObjectError + 101 Then
        MsgBox "��ӡ��û�д�������״̬�����ش��� " & err.Description & ", �����ӡ�������´�ӡ��", vbExclamation, gstrSysName, Me
    Else
        MsgBox "��ӡͼ����̳��ִ��������ӡ��ʽ����Ƭ��С�������Ƿ���ȷ��������� �� " & err.Number _
        & " ���������� " & err.Description, vbExclamation, gstrSysName, Me
    End If
    On Error Resume Next
    printer.Close
End Function

Public Function funFillPrinterParams(bShowFilmConfig As Boolean) As clsDicomPrint
'------------------------------------------------
'���ܣ����һ�����ڴ�ӡ��clsDicomPrint��ӡ����
'������bShowFilmConfig����True��ʹ�ý�Ƭ�����е���Ϣ����ӡ���ࣻFalse�Ӵ�ӡ���б��в�����Ϣ����ӡ���ࡣ
'���أ�DICOM��ӡ����
'�ϼ���������̣�frmFilmConf.cndOK_Click��frmFilm.CommBar_Film_Execute
'�¼���������̣���
'���õ��ⲿ������cDICOMPrinter���˵��ؼ�
'�����ˣ��ƽ�
'------------------------------------------------
    Dim clsOnePrinter As New clsDicomPrint
    Dim strPrinterName As String
    Dim i As Integer
    
    strPrinterName = Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    '�жϴ�ӡ���Ƿ����
    For i = 1 To cDICOMPrinter.Count
        If cDICOMPrinter(i).strname = strPrinterName Then
            Exit For
        End If
    Next i
    
    If i > cDICOMPrinter.Count Then
        MsgBox "��ӡ����" & strPrinterName & " û���ҵ���", vbInformation, gstrSysName, Me
        Exit Function
    End If
    '����ӡ�������ñ��浽clsOnePrinter��
    
    Set clsOnePrinter = cDICOMPrinter(strPrinterName)
    If bShowFilmConfig Then              '�ӽ�Ƭ�����л�ȡ��Ϣ
        With clsOnePrinter
            .strFilmBox = frmFilmConf.cboFilmBox.Text
            .strFilmSize = frmFilmConf.cboFilmSize.Text
            .strFormat = frmFilmConf.cboFormat.Text
            .strMagnification = frmFilmConf.cboMagnification.Text
            .strMedium = frmFilmConf.cboMedium.Text
            .strOrientation = frmFilmConf.cboOrientation.Text
            .strPriority = frmFilmConf.cboPriority.Text
            .strResolution = frmFilmConf.cboResolution.Text
            .strSmooth = frmFilmConf.cboSmooth.Text
            .strTrim = frmFilmConf.cboTrim.Text
            .lngCopies = frmFilmConf.lstCopies.list(frmFilmConf.lstCopies.TopIndex)
        End With
        mintPageRange = frmFilmConf.cboPageRange.ListIndex
    Else
        With clsOnePrinter
            .strOrientation = IIf(mblnIsPortrait, "PORTRAIT", "LANDSCAPE")
            .strFilmSize = CommBar_Film.FindControl(, ID_frmFilm_FilmSize, True).Text
            .strFormat = CommBar_Film.FindControl(, ID_frmFilm_Format, True).Text
        End With
        mintPageRange = 0
    End If
    Set funFillPrinterParams = clsOnePrinter
    
End Function

Private Sub CreateBar()
    '------------------------------------------------
    '���ܣ�                                  �����˵�
    '������
    '���أ�                                  ��
    '------------------------------------------------
    Dim ToolBar As CommandBar
    Dim control As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cboControl As CommandBarComboBox
    
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.CommBar_Film.VisualTheme = xtpThemeOffice2003
    Me.CommBar_Film.Icons = ImgIcons.Icons
    
    With Me.CommBar_Film.Options
        .ShowExpandButtonAlways = False     'ȥ����չ��ť
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    
    'Me.CommBar_Film.VisualTheme = IntComBarTheme                            'ͳһ���������
    Me.CommBar_Film.Item(1).Visible = False                                 '���ز˵���
    
    '������������
    Set ToolBar = Me.CommBar_Film.Add("��������", mintTBMainPosition)
    Call ToolBar.EnableDocking(xtpFlagAlignTop)
    
    With ToolBar.Controls
        Set control = .Add(xtpControlButton, ID_frmFilm_TakePictures, "����")
        control.Style = xtpButtonIconAndCaption ' xtpButtonIcon 'cbrControl.style = xtpButtonIconAndCaption
        control.IconId = 1001
        
        Set control = .Add(xtpControlButton, ID_frmFilm_FilmCol, "����")
        control.BeginGroup = True
        .Add xtpControlButton, ID_frmFilm_FilmRow, "����"
        .Add xtpControlButton, ID_frmFilm_RectPhotCase, "������ͼ��"
        .Add xtpControlButton, ID_frmFilm_FormatCustom, "��ʽ����"
        
        Set control = .Add(xtpControlButton, ID_frmFilm_OpenImages, "��")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1005
        control.ToolTipText = "��ͼ��"
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_DeleteImg, "ɾ��")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1002
        control.ToolTipText = "ɾ��ͼ��"
        
        Set control = .Add(xtpControlComboBox, ID_frmFilm_FilmSize, "��Ƭ��С")
        control.BeginGroup = True
        Set cboControl = .Add(xtpControlComboBox, ID_frmFilm_Format, "��ʽ")
        cboControl.width = 120
        Set control = .Add(xtpControlButton, ID_frmFilm_Divide, "ͼ��ָ�")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1004
        
        .Add xtpControlComboBox, ID_frmFilm_Camera, "���"
        Set control = .Add(xtpControlButton, ID_frmFilm_Quit, "�˳�")
        control.Style = xtpButtonIconAndCaption
        control.IconId = 1003
    End With
        
    '����ͼ�����������
    Set ToolBar = Me.CommBar_Film.Add("ͼ�������", mintTBImageProcessPosition)
    Call ToolBar.EnableDocking(xtpFlagAlignAny)
    ToolBar.ShowTextBelowIcons = True
    
    With ToolBar.Controls
        
        Set control = .Add(xtpControlSplitButtonPopup, ID_Active_AdjustWindow_HandAdjustWindow, " ����")
        
        control.IconId = 1006
        control.BeginGroup = True
        
        Set control = .Add(xtpControlButton, ID_frmFilm_Pan, " ����")
        control.IconId = 1008
        Set control = .Add(xtpControlButton, ID_frmFilm_Zoom, " ����")
        control.IconId = 1007
        Set control = .Add(xtpControlButton, ID_frmFilm_Invert, " ����")
        control.IconId = 1009
        Set control = .Add(xtpControlButton, ID_frmFilm_RotateLeft, " ����")
        control.IconId = 1010
        Set control = .Add(xtpControlButton, ID_frmFilm_RotateRight, " ����")
        control.IconId = 1011
        Set control = .Add(xtpControlButton, ID_frmFilm_FilterLengthDown, " ƽ����")
        control.IconId = 1012
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_FilterLengthUp, " ƽ����")
        control.IconId = 1013
        Set control = .Add(xtpControlButton, ID_frmFilm_RectZoom, " ��ѡ")
        control.IconId = 1014
        control.BeginGroup = True
        
        Set control = .Add(xtpControlSplitButtonPopup, ID_frmFilm_CutOut, " �ü�")
        control.BeginGroup = True
        control.ToolTipText = "���ɱ����ü����϶���������ѡ��ü�����˫��ͼ����вü�"
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_Custom, "���ɲü�")
        cbrPopControl.Checked = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X17, "14*17")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_11X14, "11*14")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_10X14, "10*14")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_8X10, "8*10")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X14, "14*14")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_17X14, "17*14")
        cbrPopControl.BeginGroup = True
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X11, "14*11")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_14X10, "14*10")
        Set cbrPopControl = control.CommandBar.Controls.Add(xtpControlButton, ID_frmFilm_CutOut_10X8, "10*8")
        
        Set control = .Add(xtpControlButton, ID_frmFilm_FlipHorizontal, "����")
        control.IconId = 1016
        Set control = .Add(xtpControlButton, ID_frmFilm_FlipVertical, "����")
        control.IconId = 1017
        Set control = .Add(xtpControlButtonPopup, ID_frmFilm_Label, "��ע")
        control.IconId = 1018
        control.BeginGroup = True
        control.Id = ID_frmFilm_Label
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_L, "L(��)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_R, "R(��)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_A, "A(ǰ)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_P, "P(��)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_S, "S(��)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_I, "I(��)"
        control.CommandBar.Controls.Add xtpControlButton, ID_frmFilm_Label_Delete, "�����ע"
        
        Set control = .Add(xtpControlButton, ID_frmFilm_SelAll, "ȫѡ")
        control.IconId = 1020
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_SelSeries, "����ѡ")
        control.IconId = 1021
        Set control = .Add(xtpControlButton, ID_frmFilm_SelInverse, "��ѡ")
        control.IconId = 1022
        Set control = .Add(xtpControlButton, ID_frmFilm_SelNone, "ȫ��")
        control.IconId = 1023
        
        Set control = .Add(xtpControlButton, ID_frmFilm_ImgIncrease, "����")
        control.BeginGroup = True
        Set control = .Add(xtpControlButton, ID_frmFilm_ImgDecrease, "����")
        
        Set control = .Add(xtpControlButton, ID_frmFilm_Resume, " �ָ� ")
        control.IconId = 1019
        
        control.BeginGroup = True
        
        .Add xtpControlButton, ID_frmFilm_UnDivide, "�ϲ�����"
    End With
    
    Me.CommBar_Film.FindControl(, ID_frmFilm_FilmCol, , True).Checked = True
    Me.CommBar_Film.FindControl(, ID_frmFilm_RectPhotCase, , True).Checked = False
End Sub

Public Sub subSynchronalImg(blnRestore As Boolean, intType As Integer)
'------------------------------------------------
'���ܣ��Խ�Ƭ��ӡ�е�ͼ������ͬ��
'������ blnRestore -True �ָ�ԭͼ�������False - ����ѡ����ͼ��ͬ��
'       intType   --- ͼ��ͬ�����ͣ��궨��
'���أ���
'------------------------------------------------
    
    If (Not SelectedImage Is Nothing) And funGetTagVal(SelectedImage.Tag, TAG_ѡ��) = "Select" Then
        Dim v As DicomViewer
        Dim img As DicomImage
        Dim i As Integer
        
        For Each v In FilmViewer
            For i = 1 To v.Images.Count
                If funGetTagVal(v.Images(i).Tag, TAG_ѡ��) = "Select" Then
                        Set img = v.Images(i)
                        If blnRestore = True Then
                            img.SetDefaultWindows
                            img.StretchToFit = True
                            img.FlipState = doFlipNormal
                            img.RotateState = doRotateNormal
                            img.UnsharpEnhancement = 0
                            img.UnsharpLength = 0
                            img.FilterLength = 0
                            If img.Labels.Count >= G_INT_SYS_LABEL_WWWL Then
                                img.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & img.width & "-L:" & img.Level
                            End If
                            If img.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
                                If img.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
                                    UpdateRuler img, True
                                End If
                            End If
                        Else
                            Call subImageInPhase(img, SelectedImage, intType)
                        End If
                End If
            Next i
        Next v
    End If
    '���޸Ĺ���ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
End Sub

Private Sub subReloadImgsPrint()
'------------------------------------------------
'���ܣ����޸Ĺ���ʾ������ͼ�����¼��ػ�imgsPrintͼ����
'��������
'���أ���
'�ϼ���������̣�
'�¼���������̣�
'���õ��ⲿ������
'�����ˣ��ƽ� 2006-2-17
'------------------------------------------------
    Dim v As DicomViewer
    Dim img As DicomImage
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim intStart As Integer
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    If FilmViewer.Count <= 1 Then Exit Sub
    '���㿪ʼװ�صĵ�һ��ͼ����
    intStart = funGetStartImgNo(VScro.Value, 1, 1)
    For i = 1 To FilmViewer.Count - 1
        lngWidth = FilmViewer(i).width / FilmViewer(i).MultiColumns
        lngHeight = FilmViewer(i).height / FilmViewer(i).MultiRows
        
        For j = 1 To FilmViewer(i).Images.Count
            'ɾ��imgsPrint�ж�Ӧ��ͼ��
            imgsPrint.Remove intStart
            '��imgsPrint������ͼ��
            '��¼ͼ������viewer��ռ�õ�ԭʼ��Ⱥ͸߶�
            Call funSetTagVal(FilmViewer(i).Images(j), TAG_V���, CStr(lngWidth))
            Call funSetTagVal(FilmViewer(i).Images(j), TAG_V�߶�, CStr(lngHeight))
            imgsPrint.Add FilmViewer(i).Images(j)
            '��������imgsPrint�е�ͼ���ƶ���ԭ�е�λ��
            imgsPrint.Move imgsPrint.Count, intStart
            intStart = intStart + 1
        Next j
    Next i
End Sub

Private Function funAssembleImage(AssembleViewer As DicomViewer, Optional intImgMaxWidth As Integer = 0, _
    Optional intImgMaxHeight As Integer = 0) As DicomImage
'------------------------------------------------
'���ܣ����viewer�е���ʾ������ͼ���һ��ͼ��
'������ AssembleViewer--��Ҫ��ϵ�Viewer
'       intImgMaxWidth -- ���ͼ��������
'       intImgMaxHeight -- ���ͼ������߶�
'���أ�������Ϻõ�ͼ��
'------------------------------------------------
    Dim Image As New DicomImage '��ͼ��
    Dim imgs As New DicomImages '��ʱ�洢��Ļ�ɼ���ͼ��
    Dim intWidth As Integer     '��ͼ��Ŀ��
    Dim intHeight As Integer    '��ͼ��ĸ߶�
    Dim Simg As New DicomImage
    Dim intLeft As Integer
    Dim intRight As Integer
    Dim intTop As Integer
    Dim intBottom As Integer
    Dim sZoom As Single
    Dim sOldZoom As Single          '��¼ԭ�������ű���
    Dim intImgRectWidth As Integer  '����ͼ���ռ�õ�������
    Dim intImgRectHeight As Integer '����ͼ���ռ�õ�����߶�
    Dim i As Integer
    Dim intMaxWidth As Integer      'ƴ�Ӻ�ͼ��������
    Dim intMaxHeight As Integer     'ƴ�Ӻ�ͼ������߶�
    Dim intBorder As Integer        'ͼ��֮��ı߾�
    Dim intImgX As Integer          'X�����ͼ������
    Dim intImgY As Integer          'Y�����ͼ������
    Dim intActualSizex As Integer   'ͼ����ת�任��X��������ص���
    Dim intActualSizey As Integer   'ͼ����ת�任��Y��������ص���
    Dim intOffsetX As Integer       'ƴ��ʱX�����λ��
    Dim intOffsetY As Integer       'ƴ��ʱY�����λ��
    Dim dlImgLabel As DicomLabel    'ͼ��ı�ע
    Dim j As Integer
    Dim dblX As Double, dblY As Double, intTemp As Integer
    Dim iMaxWidth As Integer, iMaxHeight As Integer
    Dim dblScaleZoom As Double
    Dim lngTempHeight  As Long
    Dim lngTempWidth As Long
    Dim lngImgLeft As Long
    Dim lngImgTop As Long
    Dim strPatiInfo(4) As String
        
    On Error GoTo err
    
    If AssembleViewer.Images.Count <= 0 Then
        '����һ����ͼ**************
        Exit Function
    End If
        
    
    '������ͼ��Ŀ�Ⱥ͸߶�
    '��ͼ��Ŀ�Ⱥ͸߶Ȳ��ܹ�����intMaxWidth��intMaxHeight����ȡ��߶ȣ�
    If intImgMaxWidth = 0 Then
        intMaxWidth = 3073
    Else
        intMaxWidth = intImgMaxWidth
    End If
    
    If intImgMaxHeight = 0 Then
        intMaxHeight = 3073
    Else
        intMaxHeight = intImgMaxHeight
    End If
    
    intBorder = 10
    intImgRectWidth = 0
    intImgRectHeight = 0
    
    '������ͼ��Ŀ�Ⱥ͸߶�
    'ʹ��ԭͼ��Ŀ�Ⱥ͸߶Ⱥͣ�����Viewer�ı�����������
    '����ͼ����¿��
    For i = 1 To AssembleViewer.Images.Count
        '������ת�任��ͼ���x�������
        intActualSizex = AssembleViewer.Images(i).sizex
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizex = AssembleViewer.Images(i).sizey
        End If
        
        '������ת�任��ͼ���y�������
        intActualSizey = AssembleViewer.Images(i).sizey
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizey = AssembleViewer.Images(i).sizex
        End If
        
        If intImgRectWidth < intActualSizex Then intImgRectWidth = intActualSizex
        If intImgRectHeight < intActualSizey Then intImgRectHeight = intActualSizey
    Next i
    
    '������������ͼ������
    intImgX = AssembleViewer.Images.Count
    If intImgX > AssembleViewer.MultiColumns Then intImgX = AssembleViewer.MultiColumns
    intImgY = (AssembleViewer.Images.Count - 1) \ AssembleViewer.MultiColumns + 1
    
    '��������ͼ����������
    If intImgRectWidth > intMaxWidth / intImgX Or intImgRectHeight > intMaxHeight / intImgY Then
        intImgRectWidth = intMaxWidth / intImgX
        intImgRectHeight = intMaxHeight / intImgY
    End If
    
    intWidth = intImgRectWidth * intImgX
    intHeight = intImgRectHeight * intImgY
    
    '����ͼ��Ŀ�ߣ����ܴ������ֵ
    '�������intMaxWidth��intMaxHeight�򣬰���ͼ���ܳ���ȣ�ʹ��С�ڵ���intMaxWidth��intMaxHeight��Ϊ�¿��,
    If intWidth > intMaxWidth Or intHeight > intMaxHeight Then
        If intHeight / intWidth > intMaxHeight / intMaxWidth Then
            intWidth = intWidth / intHeight * intMaxHeight
            intHeight = intMaxHeight
        Else
            intHeight = intHeight / intWidth * intMaxWidth
            intWidth = intMaxWidth
        End If
    End If
    
    '�ɼ�ͼ��
    '��ͼ��ɼ�����ʱͼ��
    lngImgTop = 1
    lngImgLeft = 1
    For i = 1 To AssembleViewer.Images.Count
        '����ɼ�ͼ��Ĵ�С
        
        intLeft = AssembleViewer.ImageXPosition(lngImgLeft, lngImgTop)
        intTop = AssembleViewer.ImageYPosition(lngImgLeft, lngImgTop)
        
         '������һ��ͼ��Left��Top,����λ��ʱ����С���ͽ�λ��+0.5������ֹͼ����ˣ��ۼ�λ�Ƴ���ƫ�
        lngImgLeft = lngImgLeft + AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX + 0.5
        If lngImgLeft >= AssembleViewer.width / Screen.TwipsPerPixelX Then
            lngImgLeft = 1
            lngImgTop = lngImgTop + AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY + 0.5
        End If
        
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            lngTempWidth = AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY / AssembleViewer.Images(i).ActualZoom
            lngTempHeight = AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX / AssembleViewer.Images(i).ActualZoom
        Else
            lngTempHeight = AssembleViewer.height / AssembleViewer.MultiRows / Screen.TwipsPerPixelY / AssembleViewer.Images(i).ActualZoom
            lngTempWidth = AssembleViewer.width / AssembleViewer.MultiColumns / Screen.TwipsPerPixelX / AssembleViewer.Images(i).ActualZoom
        End If
        
        If (AssembleViewer.Images(i).RotateState = doRotateLeft And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipVertical)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateRight And (AssembleViewer.Images(i).FlipState = doFlipBoth Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotate180 And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipNormal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateNormal And (AssembleViewer.Images(i).FlipState = doFlipHorizontal Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Then
            intLeft = intLeft - lngTempWidth
        End If
        
        If (AssembleViewer.Images(i).RotateState = doRotateLeft And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateRight And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotate180 And (AssembleViewer.Images(i).FlipState = doFlipNormal Or AssembleViewer.Images(i).FlipState = doFlipHorizontal)) Or _
            (AssembleViewer.Images(i).RotateState = doRotateNormal And (AssembleViewer.Images(i).FlipState = doFlipVertical Or AssembleViewer.Images(i).FlipState = doFlipBoth)) Then
            intTop = intTop - lngTempHeight
        End If
        
        intRight = lngTempWidth + intLeft
        intBottom = lngTempHeight + intTop

        '������ת�任��ͼ���x�������
        intActualSizex = AssembleViewer.Images(i).sizex
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizex = AssembleViewer.Images(i).sizey
        End If
        
        '������ת�任��ͼ���y�������
        intActualSizey = AssembleViewer.Images(i).sizey
        If AssembleViewer.Images(i).RotateState = doRotateLeft Or AssembleViewer.Images(i).RotateState = doRotateRight Then
            intActualSizey = AssembleViewer.Images(i).sizex
        End If
        
        '�������ű��� hj�޸�,�����ͼ�ϲ�ʱ���Ŵ��ͼ���޷������Ŵ������
        sZoom = intImgRectHeight / IIf((intBottom - intTop) > intActualSizey Or intBottom = 0, intActualSizey, (intBottom - intTop))
        If sZoom > intImgRectWidth / IIf((intRight - intLeft) > intActualSizex Or intRight = 0, intActualSizex, (intRight - intLeft)) Then
            sZoom = intImgRectWidth / IIf((intRight - intLeft) > intActualSizex Or intRight = 0, intActualSizex, (intRight - intLeft))
        End If
      
        '��ͼ�����±������ţ�Ȼ�����»����
        sOldZoom = AssembleViewer.Images(i).ActualZoom
        AssembleViewer.Images(i).StretchToFit = False
        AssembleViewer.Images(i).Zoom = sZoom
        
        If UpdateRuler(AssembleViewer.Images(i), True) = 1 Then
            '��߱�ע�����ڣ������ǡ��ϲ����ԡ�֮���ͼ���ٴν��С��ϲ����ԡ���
            MsgBox "ͼ��ı����Ϣ�����ڡ�" & vbCrLf & vbCrLf & "ԭ�������" & vbCrLf & "    1���ϲ����ԵĽ��ͼ�񣬲����ٴν��кϲ����ԡ�" & vbCrLf & "    2������δ֪����", vbOKOnly, "������ʾ"
            Exit Function
        End If
        
        '�����ô�ӡ�����С�������û��Լ����ı�ע
        Call subChangeLabelForPrint(AssembleViewer.Images(i), 1)

        '����ͼ����Ľ���Ϣ
        Call subDispImageInfo(AssembleViewer.Images(i), False, False, True)
                
        Set Simg = AssembleViewer.Images(i).PrinterImage(8, 1, True, sZoom, intLeft, intRight, intTop, intBottom)
        
        
        '��ʾͼ����Ľ���Ϣ
        '��ԭ��ͼ��ı�ע����ӵ����ڵ�ͼ���У���Ϊ�û��Լ����ı�ע����һ���Ѿ�������ͼ���У��������ֻ�ָ�ϵͳ��ע
        Simg.Labels.Clear
        For j = 1 To IIf(G_INT_SYS_LABEL_COUNT <= AssembleViewer.Images(i).Labels.Count, G_INT_SYS_LABEL_COUNT, AssembleViewer.Images(i).Labels.Count)
            Simg.Labels.Add AssembleViewer.Images(i).Labels(j)
            Simg.Labels(Simg.Labels.Count).Visible = False
        Next j
        '���ԭ����ͼ������
        Simg.Attributes.Add &H8, &H60, AssembleViewer.Images(i).Attributes(&H8, &H60)
        Call subDispImageInfo(Simg, True, False, False)     ''��ʾ�����Ľ���Ϣ�ʹ���λ��Ϣ
        '��Ϊ����Ѿ���ǰ�滭���ˣ�����������ر�ߵ���ʾ
        Call UpdateRuler(Simg, False)
        
        '���ô�ӡ�����С
        Call subChangeLabelForPrint(Simg, 1)

        '����ͼ����ʾ����ı�������ͼ������м䣬�Ľ����ַŵ�����������ĸ���
        dblX = Simg.sizex / (AssembleViewer.width / AssembleViewer.MultiColumns)
        dblY = Simg.sizey / (AssembleViewer.height / AssembleViewer.MultiRows)
        If dblX < dblY Then
            intTemp = dblY * AssembleViewer.width / AssembleViewer.MultiColumns
            intLeft = -(intTemp - Simg.sizex) / 2
            intRight = intTemp + intLeft
            intTop = 0
            intBottom = 0
        Else
            intTemp = dblX * AssembleViewer.height / AssembleViewer.MultiRows
            intLeft = 0
            intRight = 0
            intTop = -(intTemp - Simg.sizey) / 2
            intBottom = intTemp + intTop
        End If
        
        Set Simg = Simg.PrinterImage(8, 1, True, 1, intLeft, intRight, intTop, intBottom)
        
        '�ָ�ͼ��ԭ�������ű���
        AssembleViewer.Images(i).Zoom = sOldZoom
        
        '�ָ�ͼ��ԭ���ı�ע
        '��ʾͼ����Ľ���Ϣ
        Call subDispImageInfo(AssembleViewer.Images(i), True, False, True)
        
        '���ô�ӡ�����С�������û��Լ����ı�ע
        Call subChangeLabelForPrint(AssembleViewer.Images(i), 0)

        imgs.Add Simg
    Next i
     
    '��ȷ������ͼ��Ŀ�Ⱥ͸߶�
    intImgRectWidth = 0
    intImgRectHeight = 0
     
    For i = 1 To imgs.Count
        If intImgRectWidth < imgs(i).sizex Then intImgRectWidth = imgs(i).sizex
        If intImgRectHeight < imgs(i).sizey Then intImgRectHeight = imgs(i).sizey
    Next i
    
    If Not clsTruePrinter Is Nothing Then
        intBorder = clsTruePrinter.lngImageBorderWidth
    Else
        intBorder = 1
    End If
     
    intImgRectWidth = intImgRectWidth + intBorder
    intImgRectHeight = intImgRectHeight + intBorder
    intWidth = intImgRectWidth * intImgX
    intHeight = intImgRectHeight * intImgY
    
    '������ͼ��
    Image.Name = "print"
    Image.PatientID = "print001"
    Image.Attributes.Add &H8, &H16, doSOP_SecondaryCapture
    Image.Attributes.Add &H28, &H2, 1 ' samples/pixel
    Image.Attributes.Add &H28, &H4, "MONOCHROME2" ' photometric interpreation  'CT����MONOCHROME2,CR����MONOCHROME1��
    Image.Attributes.Add &H28, &H10, intHeight   'x,Rows
    Image.Attributes.Add &H28, &H11, intWidth 'Y,Columns
    Image.Attributes.Add &H28, &H100, 8 'bits allocated
    Image.Attributes.Add &H28, &H101, 8 ' bits stored
    Image.Attributes.Add &H28, &H102, 7 ' high bit
    ReDim pix(intWidth, intHeight) As Byte
    
    Image.Attributes.Add &H7FE0, &H10, pix
    
    '��ȡ����ͼ���Ⱥ͸߶�
    iMaxWidth = 0
    iMaxHeight = 0
    For i = 1 To imgs.Count
        If iMaxWidth < imgs(i).sizex Then
            iMaxWidth = imgs(i).sizex
            iMaxHeight = imgs(i).sizey
        End If
    Next i
    
    'ƴ����ͼ��
    For i = 1 To imgs.Count
        '����ͼ����λ��
        dblScaleZoom = iMaxWidth / imgs(i).sizex
        intOffsetX = (intImgRectWidth - imgs(i).sizex * dblScaleZoom) / 2
        intOffsetY = (intImgRectHeight - imgs(i).sizey * dblScaleZoom) / 2
        Image.Blt imgs(i), 0, 0, ((i - 1) Mod AssembleViewer.MultiColumns) * intImgRectWidth + intOffsetX, ((i - 1) \ AssembleViewer.MultiColumns) * intImgRectHeight + intOffsetY, imgs(i).sizex * dblScaleZoom, imgs(i).sizey * dblScaleZoom, 1, 1, dblScaleZoom, False
    Next i
    
    Set funAssembleImage = Image
    Exit Function
    
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetStartViewerNo(intPage As Integer, intViewer As Integer) As Integer
'------------------------------------------------
'���ܣ�ͨ����ǰҳ������ȡ�ӵ�һҳ����ǰ�У�Viewer��������
'������ intPage     --��ǰҳ������
'       intViewer   --��ǰViewer����
'���أ���intViewer ֮ǰ������viewer������
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetStartViewerNo = 0
    
    If intPage > UBound(marrPages) Then Exit Function
    If intViewer > marrPages(intPage).intViewerCount Then Exit Function
    
    For i = 1 To intPage - 1
        funGetStartViewerNo = funGetStartViewerNo + marrPages(i).intViewerCount
    Next i
    
    funGetStartViewerNo = funGetStartViewerNo + intViewer
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    funGetStartViewerNo = 0
End Function

Private Function funGetStartImgNo(intPage As Integer, intViewer As Integer, intImage As Integer) As Integer
'------------------------------------------------
'���ܣ�ͨ����ǰҳ������ȡ�ӵ�һҳ����ǰҳ֮ǰ��Images��������
'������ intPage     --��ǰҳ������
'       intViewer   --��ǰViewer������
'       intImage    --��ǰImage������
'���أ���ǰҳintImage֮ǰ��Images������
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    funGetStartImgNo = 0
    
    If intPage > UBound(marrPages) Then Exit Function
    If intViewer > marrPages(intPage).intViewerCount Then Exit Function
    If intImage > marrPages(intPage).ViewerLayout(intViewer).intColumns * marrPages(intPage).ViewerLayout(intViewer).intRows Then Exit Function
    
    For i = 1 To intPage - 1
        funGetStartImgNo = funGetStartImgNo + marrPages(i).intImageCount
    Next i
    
    For i = 1 To intViewer - 1
        funGetStartImgNo = funGetStartImgNo + marrPages(intPage).ViewerLayout(i).intColumns * marrPages(intPage).ViewerLayout(i).intRows
    Next i
    
    funGetStartImgNo = funGetStartImgNo + intImage
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    funGetStartImgNo = 0
End Function

Public Sub subDispReferLineFilm()
'------------------------------------------------
'���ܣ��ڽ�Ƭ��ӡ��������ʾ��λ��
'������
'����ֵ����
'2009 ��
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim v As DicomViewer
    Dim im As DicomImage
    Dim imm As DicomImage
    
    On Error Resume Next
    
    '��ɾ��ԭ���Ķ�λ��
    For Each v In Me.FilmViewer
        For Each im In v.Images
            subDeleteAppointLabel im, "RL"          'ɾ��ָ�����͵ı�ע
        Next
        v.Refresh
    Next
    
    '��ʾ���ж�λ��
    If Button_miAllReferLine = False Then Exit Sub
    
    For i = 1 To (FilmViewer.Count - 1)
        For Each im In Me.FilmViewer(i).Images
            If subGetReferImg(im) = True Then
                For j = 1 To (FilmViewer.Count - 1)
                    For Each imm In Me.FilmViewer(j).Images
                        If subGetReferImg(imm) = False Then
                            Call subDrawRefLine(imm, im, True, "RLL", False)
                        End If
                    Next
                Next
            End If
        Next
        Me.FilmViewer(i).Refresh
    Next
End Sub
Private Function subGetReferImg(img As DicomImage) As Boolean
    '����  ��ǰͼ���Ƿ��Ƕ�λ��
    '������ img --- ��Ҫ�жϵ�ͼ��
    '���أ� True -- �Ƕ�λ��False -- ���Ƕ�λ��
    
    Dim v As Variant
    Dim i As Integer
    Dim strAttr As String
    
    On Error GoTo err
    v = img.Attributes(&H8, &H8).Value
    
    If (VarType(v) > 8192) Then
        For i = LBound(v, 1) To UBound(v, 1)
            If i = 3 And v(i) = "LOCALIZER" Then
                subGetReferImg = True
                Exit Function
            End If
        Next
    End If
    
    '�����8,8����û�С�LOCALIZER����ǣ��ټ��ͼ�����������(8,103E)���Ƿ������LOC��
    'GE ��MR�����������а�����LOC���������ֵ�MR�����������а���"SURVEY"
    If img.Attributes(&H8, &H103E).Exists And Not IsNull(img.Attributes(&H8, &H103E).Value) Then
        strAttr = img.Attributes(&H8, &H103E).Value
        If InStr(UCase(strAttr), "LOC") <> 0 Or InStr(UCase(strAttr), "SURVEY") <> 0 Then
            subGetReferImg = True
            Exit Function
        End If
    End If
    
err:
    '��������
End Function

Private Sub subDeleteAllImages()
'------------------------------------------------
'���ܣ� ɾ��ȫ��ͼ��
'��������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    '�����ʵͼ��
    imgsPrint.Clear
        
    'ɾ��ͼ��֮��ͼ������ˣ����¼���ҳ��
    Call subRecalPages
    '������ʾͼ��
    Call subShowPrintImages(1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub CommBar_Execute_PrintFilm()
    Dim strPrinterName As String
    Dim i As Integer
    
    If mblnPrinted = True Then
        If MsgBox("�ý�Ƭ�Ѿ���ӡ���ˣ��Ƿ���Ҫ�ٴδ�ӡ��", vbYesNo, gstrSysName, Me) = vbNo Then
            Exit Sub
        End If
    End If
        
    strPrinterName = Me.CommBar_Film.FindControl(, ID_frmFilm_Camera, , True).Text
    If Len(Trim(strPrinterName)) <= 0 Then
        MsgBox "��ѡ��Ƭ��ӡ��!", vbInformation, gstrSysName, Me
        Exit Sub
    End If
    '����ӡ���Ƿ����
    For i = 1 To cDICOMPrinter.Count
        If cDICOMPrinter(i).strname = strPrinterName Then
            Exit For
        End If
    Next i
    
    If i > cDICOMPrinter.Count Then
        MsgBox "��ӡ����" & strPrinterName & " û���ҵ���" & vbCrLf & vbCrLf & "��ѡ��Ƭ��ӡ��!", vbInformation, gstrSysName, Me
        Exit Sub
    End If
        
    '�ж��Ƿ���Ҫ��ʾ��ӡ���ô���
    If bShowFilmConfig Then
        Set frmFilmConf.f = Me
        With frmFilmConf
            .sstabFilmConfig.TabVisible(0) = False
            .sstabFilmConfig.TabVisible(1) = True
            If strPrinterName = "" Then Exit Sub
            .cboFilmBox.Text = cDICOMPrinter(strPrinterName).strFilmBox
            .cboFilmSize.Text = Me.CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
            .cboFormat.Text = Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text
            .cboMagnification.Text = cDICOMPrinter(strPrinterName).strMagnification
            .cboMedium.Text = cDICOMPrinter(strPrinterName).strMedium
            .cboOrientation.Text = IIf(mblnIsPortrait, "PORTRAIT", "LANDSCAPE")
            .cboPriority.Text = cDICOMPrinter(strPrinterName).strPriority
            .cboResolution.Text = cDICOMPrinter(strPrinterName).strResolution
            .cboSmooth.Text = cDICOMPrinter(strPrinterName).strSmooth
            .cboTrim.Text = cDICOMPrinter(strPrinterName).strTrim
            .lstCopies = cDICOMPrinter(strPrinterName).lngCopies
            If .zlShowMe = False Then
                Exit Sub
            End If
        End With
    Else
        If MsgBox("�Ƿ��ӡ��Ƭ��", vbOKCancel, "PACS��ʾ", Me) = vbCancel Then
            Exit Sub
        End If
        Set clsTruePrinter = funFillPrinterParams(bShowFilmConfig)
    End If
        
    '��ӡ֮ǰ���������򴰿�
    mblnPrinting = True
    On Error GoTo err
    
    CommBar_Film.FindControl(, ID_frmFilm_TakePictures, , True).Enabled = False
    Me.Caption = "���ڴ�ӡ��Ƭ�����Ժ�......"
    If subPrintFilm(clsTruePrinter) = True Then
        mintPrintFilmCount = mintPrintFilmCount + 1
        Me.Caption = "��Ƭ��ӡԤ������ӡ�� " & mintPrintFilmCount & " ��"
        If blnPrintOkEcho = True Then
            MsgBox "��Ƭ��ӡԤ������ӡ�� " & mintPrintFilmCount & " �γɹ���", vbOKOnly, gstrSysName, Me
        End If
        
        '��ӡ�ɹ��������ͼ��
        If mblnClearAfterPrint = True Then
            Call subDeleteAllImages
        End If
        
        '��ʾ��ӡ��ɵ�����
        Call PrintFilmBeep(2)
    Else
        Me.Caption = "��Ƭ��ӡ���ɹ������������ú��ٴ�ӡ"
    End If
    mblnPrinting = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    mblnPrinting = False
End Sub

Private Sub OpenImages(strImageIDs As String)
'------------------------------------------------
'���ܣ���ͼ��ѡ�񴰿ڣ������û��Լ�ѡ��ͼ����ӵ���Ƭ��
'������ strImageIDs -- Ҫ�򿪵�����ID����'�����ǡ����к�1|1-3;5-27;33-100+���к�2|ȫ����,ȫ����ʾ��ȫ��ͼ��
'���أ���
'------------------------------------------------
    Dim blnAllImages As Boolean         '�Ƿ��ȫ��ͼ��
    Dim imgs As New DicomImages              '��Ҫ�򿪵�ͼ��
    Dim iSeriesID As Integer            '���к�
    Dim strSeries() As String
    Dim strImages() As String
    Dim i As Integer
    Dim j As Integer
    Dim k  As Integer
    Dim h As Integer
    Dim tmpImg As DicomImage
    
    On Error GoTo err
    
    '��ѡ�е�ͼ����ӵ�imgs ��
    strSeries = Split(strImageIDs, "+")
    For i = 0 To UBound(strSeries)
        iSeriesID = Split(strSeries(i), "|")(0)
        If Split(strSeries(i), "|")(1) = "ȫ��" Then
            For k = 1 To ZLSeriesInfos(iSeriesID).ImageInfos.Count
                Set tmpImg = funLoadAImage(iSeriesID, k, 0)
                
                If Not tmpImg Is Nothing Then
                    subInitAImage tmpImg, 0, Nothing
                    imgs.Add tmpImg
                End If
            Next k
        Else
            strImages = Split(Split(strSeries(i), "|")(1), ";")
            For j = 0 To UBound(strImages)
                For h = Split(strImages(j), "-")(0) To Split(strImages(j), "-")(1)
                    Set tmpImg = funLoadAImage(iSeriesID, h, 0)
                    If Not tmpImg Is Nothing Then
                        subInitAImage tmpImg, 0, Nothing
                        imgs.Add tmpImg
                    End If
                Next h
            Next j
        End If
    Next i
    
    '��ȡimgs �е�ͼ�񣬰�ͼ����ӵ���Ƭ��ӡ����
    For i = 1 To imgs.Count
        imgsPrint.Add imgs(i)
        subChangeLabelForPrint imgsPrint(imgsPrint.Count), 0
    Next i
    
    'ͼ�������ˣ�����ҳ��
    Call subRecalPages
    
    '������ʾ��һҳ��ͼ��
    Call subShowPrintImages(Me.VScro.Value)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subDelImage()
'------------------------------------------------
'���ܣ� ɾ��ͼ��
'       ���ж��Ƿ���ͼ��ѡ���ˣ�����У���ɾ����ѡ���ͼ��
'       ���û��ͼ��ѡ����ɾ����ǰ��굥������ͼ��
'��������
'���أ���
'------------------------------------------------
    Dim intStart As Integer
    Dim blnSelected As Boolean      '�Ƿ���ͼ��ѡ���ˣ��������ֻɾ����ѡ���ͼ�񣬷���ɾ����ǰͼ��
    Dim i As Integer
    Dim j As Integer
    Dim blnDeleted As Boolean
    Dim intDelViewer As Integer
    Dim intDelImage As Integer
    
    On Error GoTo err
    
    '���ж��Ƿ���ͼ��ѡ���ˣ�����У���ɾ����ѡ���ͼ��
    blnSelected = False
    For i = Me.FilmViewer.Count - 1 To 1 Step -1
        For j = Me.FilmViewer(i).Images.Count To 1 Step -1
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                intStart = funGetStartImgNo(Me.VScro.Value, i, j)
                imgsPrint.Remove intStart
                blnSelected = True
            End If
        Next
    Next
    
    '���û��ͼ��ѡ����ɾ����ǰ��굥������ͼ��
    If blnSelected = False And Not SelectedImage Is Nothing And mintSelectedViewer <> 0 And mintSelectedImage <> 0 Then
        
        intStart = funGetStartImgNo(Me.VScro.Value, mintSelectedViewer, mintSelectedImage)
        imgsPrint.Remove intStart
        
        intDelViewer = mintSelectedViewer
        intDelImage = mintSelectedImage
        
        blnDeleted = True
    End If
        
    '�����ͼ��ɾ���ˣ���������ʾͼ�񣬲���������һ����ѡ�е�ͼ�񣬷�������ɾ��
    If blnDeleted = True Or blnSelected = True Then
        'ɾ��ͼ��֮��ͼ������ˣ����¼���ҳ��
        Call subRecalPages
        '������ʾͼ��
        Call subShowPrintImages(Me.VScro.Value)
    
        '�����ɾ����ѡ�еļ���ͼ����ѡ��ͼ�����óɵ�һ�ţ�����Ҫ����
        '�����ɾ����ǰͼ������Ҫѡ������ǰһ��ͼ��
        If blnSelected = False Then
            If imgsPrint.Count > 1 Then
                
                intDelImage = intDelImage - 1
                If intDelImage = 0 Then
                    intDelViewer = intDelViewer - 1
                    If intDelViewer <> 0 Then
                        intDelImage = marrPages(Me.VScro.Value).ViewerLayout(intDelViewer).intColumns * marrPages(Me.VScro.Value).ViewerLayout(intDelViewer).intRows
                    Else
                        intDelImage = 0
                    End If
                End If
                
                If intDelViewer > 0 And intDelImage > 0 Then
                    '��ѡ��ͼ�����óɵ�ǰͼ��ǰһ��
                    Call subImageCurrent(1, 1, False)
                    Call subImageCurrent(intDelViewer, intDelImage, True)
                End If
            End If
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subSetFilmFormat()
'------------------------------------------------
'���ܣ����Ű��ʽ���ڣ������Զ������ý�Ƭ�ĸ�ʽ
'��������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim intTmp As Integer
    Dim intPage As Integer                       '''''��ҳ��
    
    On Error GoTo err
    
    Set frmFilmConf.f = Me
    
    '�����Ű��ʽ���ڵĿؼ�����
    With frmFilmConf
        .sstabFilmConfig.TabVisible(0) = True
        .sstabFilmConfig.TabVisible(1) = False
        .cobSize.Text = CommBar_Film.FindControl(, ID_frmFilm_FilmSize, , True).Text
        .cobAspect.Text = IIf(mblnIsPortrait, "����", "����")     ''�Ƿ������ӡ
        
        If Not mblnIsCustom Then     '��׼����
            .txtRow = UBound(marrRCCount)
            .txtCol = marrRCCount(1)
            .Option(0).Value = True
        Else
            If mblnIsRow Then        '���Զ���
                .txtRow = UBound(marrRCCount)
                .txtCol = marrRCCount(1)
                .Option(1).Value = True
            Else                    '���Զ���
                .txtCol = UBound(marrRCCount)
                .txtRow = marrRCCount(1)
                .Option(2).Value = True
            End If
            
            For i = 1 To UBound(marrRCCount)
                .txtC(i) = marrRCCount(i)
            Next
        End If
    End With
    
    '��ʾ�Ű��ʽ����
    If frmFilmConf.zlShowMe = True Then
        Call funChangeFormat(Me.VScro.Value, Me.CommBar_Film.FindControl(, ID_frmFilm_Format, , True).Text)
'       '������ʾ��һҳ
        Call subShowOnePage(Me.VScro.Value)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subShiftSelect(intViewerIndex As Integer, intImgIndex As Integer)
'------------------------------------------------
'���ܣ�ʹ��Shift+������ѡ��ͼ��
'������ intViewerIndex --- ��ǰ������Viewer����
'       intImgIndex  --  ��ǰ������ͼ������
'���أ���
'------------------------------------------------
    Dim intStartViewer As Integer
    Dim intStartImage As Integer
    Dim intEndViewer As Integer
    Dim intEndImage As Integer
    Dim i As Integer
    Dim j As Integer
    Dim blnSelected As Boolean
    
    On Error GoTo err
    blnSelected = False
    '�Ȳ��ҵ�һ����ѡ�е�ͼ��
    For i = 1 To FilmViewer.Count - 1
        For j = 1 To FilmViewer(i).Images.Count
            If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                blnSelected = True
                intStartViewer = i
                intStartImage = j
                Exit For
            End If
        Next j
        If blnSelected = True Then
            Exit For
        End If
    Next i
    
    '���û�б�ѡ�е�ͼ����ֻ����ǰͼ��
    If blnSelected = False Then
        Call subImageSelect(intViewerIndex, intImgIndex, True)
    Else
        '���ǰ���б�ѡ�е�ͼ�����Դ�ͼ��Ϊ��һ��ͼ��ѭ��ѡ�񵽵�ǰͼ��
        '�жϱ�ѡ�е�ͼ�����ڵ�ǰͼ���ǰ�滹�Ǻ���
        If intStartViewer < intViewerIndex Then
            intEndViewer = intViewerIndex
            intEndImage = intImgIndex
        ElseIf intStartViewer = intViewerIndex Then
            intEndViewer = intViewerIndex
            If intStartImage <= intImgIndex Then
                intEndImage = intImgIndex
            Else
                intEndImage = intStartImage
                intStartImage = intImgIndex
            End If
        Else
            intEndViewer = intStartViewer
            intEndImage = intStartImage
            intStartViewer = intViewerIndex
            intStartImage = intImgIndex
        End If
        
        'ѭ��ѡ��Χ�ڵ�ͼ��,������ͼ�񶼲�ѡ��
        For i = 1 To FilmViewer.Count - 1
            For j = 1 To FilmViewer(i).Images.Count
                If i = intStartViewer Then
                    If j >= intStartImage And j <= FilmViewer(intStartViewer).Images.Count Then
                        Call subImageSelect(i, j, True)
                    Else
                        Call subImageSelect(i, j, False)
                    End If
                ElseIf i = intEndViewer Then
                    If j >= 1 And j <= intEndImage Then
                        Call subImageSelect(i, j, True)
                    Else
                        Call subImageSelect(i, j, False)
                    End If
                ElseIf i > intStartViewer And i < intEndViewer Then
                    Call subImageSelect(i, j, True)
                Else
                    '��ѡ��
                    Call subImageSelect(i, j, False)
                End If
            Next j
            FilmViewer(i).Refresh
        Next i
    End If
    
    '��ѡ���ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subImageSelect(intViewerIndex As Integer, intImgIndex As Integer, blnSelected As Boolean)
'------------------------------------------------
'���ܣ�ѡ����߲�ѡ��ͼ��
'������ intViewerIndex --- ��Ҫѡ�����ȡ��ѡ���ͼ�����ڵ�����
'       intImgIndex --- ��Ҫѡ�����ȡ��ѡ���ͼ�����ڵ�����
'       blnSelected --- True ѡ��False ȡ��ѡ��
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    If blnSelected = True Then
        Call funSetTagVal(FilmViewer(intViewerIndex).Images(intImgIndex), TAG_ѡ��, "Select")
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = lngSelectedImageBorderColor
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = lngSelectedImageBorderLineStyle
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = lngSelectedImageBorderLineWidth
    Else
        Call funSetTagVal(FilmViewer(intViewerIndex).Images(intImgIndex), TAG_ѡ��, "")
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = vbWhite
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = 0
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = 1
    End If
    FilmViewer(intViewerIndex).Refresh
    Exit Sub
err:
    '��ʱ������

End Sub

Private Sub subImageCurrent(intViewerIndex As Integer, intImgIndex As Integer, blnCurrent As Boolean)
'------------------------------------------------
'���ܣ�����ָ��ͼ���Ƿ�ǰѡ�е�ͼ��
'������ intViewerIndex --- ��Ҫѡ�����ȡ��ѡ���ͼ�����ڵ�����
'       intImgIndex --- ��Ҫѡ�����ȡ��ѡ���ͼ�����ڵ�����
'       blnCurrent --- True �ǵ�ǰͼ��False ���ǵ�ǰͼ��
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    If blnCurrent = True Then
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = lngCurrentImageBorderColor
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = lngCurrentImageBorderLineStyle
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = lngCurrentImageBorderLineWidth
        Set SelectedImage = FilmViewer(intViewerIndex).Images(intImgIndex)
        mintSelectedViewer = intViewerIndex
        mintSelectedImage = intImgIndex
    Else
        FilmViewer(intViewerIndex).Labels(intImgIndex).ForeColour = vbWhite
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineStyle = 0
        FilmViewer(intViewerIndex).Labels(intImgIndex).LineWidth = 1
    End If
    FilmViewer(intViewerIndex).Refresh
    Exit Sub
err:
    '��ʱ������

End Sub

Private Sub SelOneSeries()
'------------------------------------------------
'���ܣ�ѡ��ǰ���е�ͼ��
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    Dim strSeriesUID As String
    
    On Error GoTo err
    '����ȡ��ǰͼ�������UID
    If SelectedImage Is Nothing Then Exit Sub
    strSeriesUID = SelectedImage.SeriesUID
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                Call subImageSelect(i, j, IIf(FilmViewer(i).Images(j).SeriesUID = strSeriesUID, True, False))
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '��ѡ���ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SelectInverse()
'------------------------------------------------
'���ܣ���ѡͼ��
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    If SelectedImage Is Nothing Then Exit Sub
    
    For i = 1 To FilmViewer.Count - 1
        If FilmViewer(i).Visible = True Then
            For j = 1 To FilmViewer(i).Images.Count
                If funGetTagVal(FilmViewer(i).Images(j).Tag, TAG_ѡ��) = "Select" Then
                    Call subImageSelect(i, j, False)
                Else
                    Call subImageSelect(i, j, True)
                End If
            Next j
            FilmViewer(i).Refresh
        End If
    Next i
    
    '��ѡ���ͼ���ϴ���ԭʼͼ������
    Call subReloadImgsPrint
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub subOpenFilmView()
'------------------------------------------------
'���ܣ���ͼ������
'������ thisViewer - �´򿪴��ڵ�Viewer,��Ҫ����ȡ��Ⱥ͸߶�
'���أ���
'------------------------------------------------
    Dim dcmNewImage As DicomImage
    
    On Error GoTo err
    
    '����Ѿ�����һ��ͼ�����壬���ٴ���һ��
    If Not mfrmFilmView Is Nothing Then Exit Sub
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer = 0 Then Exit Sub
    If mintSelectedImage = 0 Then Exit Sub
    If FilmViewer(mintSelectedViewer).Images.Count < mintSelectedImage Then Exit Sub
    
    Set mfrmFilmView = New frmFilmView
    
    Call mfrmFilmView.zlShowMe(SelectedImage, Me, mintSelectedViewer, mintSelectedImage)
    
    '    ���Ͻػ���Ϣ��hook�����ܷ���mfrmFilm��load�¼�
    plngFilmViewPreWndProc = FilmViewHook(mfrmFilmView.hwnd)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subFilmViewButtonClick(control As CommandBarControl)
'------------------------------------------------
'���ܣ���ͼ������
'������ thisViewer - �´򿪴��ڵ�Viewer,��Ҫ����ȡ��Ⱥ͸߶�
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    If mfrmFilmView Is Nothing Then Exit Sub
    
    Call mfrmFilmView.ZLToolButtonClick(control)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subCutOutRatio(lngControlID As Long)
'------------------------------------------------
'���ܣ����òü��Ĺ̶��������������̶������Ĳü���
'������ lngControlID --- �ü��˵���ID
'���أ���
'------------------------------------------------
    On Error GoTo err
    Dim intLeft As Integer
    Dim intTop As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    Dim dblScaleRatio As Double
    
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer = 0 Then Exit Sub
    If mintSelectedImage = 0 Then Exit Sub
    
    '���òü��ı���
    Select Case lngControlID
        Case ID_frmFilm_CutOut_14X17
            mdblCutOutRatio = 14 / 17
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_11X14
            mdblCutOutRatio = 11 / 14
            dblScaleRatio = 11 / 14
        Case ID_frmFilm_CutOut_10X14
            mdblCutOutRatio = 10 / 14
            dblScaleRatio = 10 / 14
        Case ID_frmFilm_CutOut_8X10
            mdblCutOutRatio = 8 / 10
            dblScaleRatio = 8 / 14
        Case ID_frmFilm_CutOut_14X14
            mdblCutOutRatio = 14 / 14
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_17X14
            mdblCutOutRatio = 17 / 14
            dblScaleRatio = 1
        Case ID_frmFilm_CutOut_14X11
            mdblCutOutRatio = 14 / 11
            dblScaleRatio = 14 / 17
        Case ID_frmFilm_CutOut_14X10
            mdblCutOutRatio = 14 / 10
            dblScaleRatio = 14 / 17
        Case ID_frmFilm_CutOut_10X8
            mdblCutOutRatio = 10 / 8
            dblScaleRatio = 10 / 17
    End Select
    
    '��ʾ�̶������Ĳü���
    If mintSelectedViewer < FilmViewer.Count Then
        If mintSelectedImage <= FilmViewer(mintSelectedViewer).Images.Count Then
            '����SelectImg����ü����λ��
            intHeight = SelectedImage.sizey
            intWidth = intHeight * mdblCutOutRatio
            If intWidth > SelectedImage.sizex Then
                '������Ⱥ͸߶�
                intWidth = SelectedImage.sizex
                intHeight = intWidth / mdblCutOutRatio
            End If
            
            '�������ű������¼����Ⱥ͸߶�
            intWidth = intWidth * dblScaleRatio
            intHeight = intHeight * dblScaleRatio
            
            '����Left��Top
            intTop = (SelectedImage.sizey - intHeight) / 2
            intLeft = (SelectedImage.sizex - intWidth) / 2
            
            '��ʾ�ü���
            FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Add GetNewLabel(doLabelRectangle, intLeft, intTop, intWidth, intHeight)
            Set mdcmSelectLabel = FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels(FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Count)
            mdcmSelectLabel.Tag = CUT_LABEL
            mintCutOutViewer = mintSelectedViewer
            mintCutOutImage = mintSelectedImage
            mintCutOutLabel = FilmViewer(mintSelectedViewer).Images(mintSelectedImage).Labels.Count
            FilmViewer(mintSelectedViewer).Refresh
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCutOutClick()
'------------------------------------------------
'���ܣ������ü���ť���Զ����������б��ж�Ӧ��ѡ�е���Ŀ
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    '�жϵ�ǰ�����ֲü���ʽ
    If CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_Custom, , True).Checked Then
        mdblCutOutRatio = 0
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X17, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X17)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_11X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_11X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_10X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_8X10, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_8X10)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_17X14, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_17X14)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X11, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X11)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X10, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_14X10)
    ElseIf CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X8, , True).Checked Then
        Call subCutOutRatio(ID_frmFilm_CutOut_10X8)
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCutOutButtonState(lngControlID As Long)
'------------------------------------------------
'���ܣ������ü��Ӳ˵���״̬
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_Custom, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X17, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_11X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_8X10, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_17X14, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X11, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_14X10, , True).Checked = False
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut_10X8, , True).Checked = False
        
        CommBar_Film.Item(3).FindControl(, lngControlID, , True).Checked = True
        CommBar_Film.Item(3).FindControl(, ID_frmFilm_CutOut, , True).IconId = lngControlID
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subCalImageMaxSize(strFilmSize As String, strFormat As String, intImageResolution As Integer, _
    arrImageSize() As ImageSize)
'------------------------------------------------
'���ܣ����ݽ�Ƭ�ߴ磬ͼ�񲼾֣�����ÿһ��ͼ�������ߴ�
'������ strFilmSize --- ��Ƭ�ߴ�
'       strFormat --- ��Ƭ��ʽ
'       intImageResolution --- �����ͼ��ֱ���
'       arrImageSize --- [OUT]����ÿ��ͼ�������ֱ���
'���أ���
'------------------------------------------------
    Dim intImageCount As Integer
    Dim lngFilmWidth As Long
    Dim lngFilmHeight As Long
    Dim strCurFormat As String
    Dim i As Integer
    Dim j As Integer
    Dim intX As Integer
    Dim intY As Integer
    Dim intCount As Integer
    
    ReDim arrImageSize(0)
    
    On Error GoTo err
    
    '������Ƭ�ĳߴ磬�������������8_5INX11IN������14INX17IN��
    If UBound(Split(UCase(strFilmSize), "X")) = 1 Then
        lngFilmWidth = Val(Replace(Split(UCase(strFilmSize), "X")(0), "_", ".")) * intImageResolution
        lngFilmHeight = Val(Replace(Split(UCase(strFilmSize), "X")(1), "_", ".")) * intImageResolution
    Else
        '��Ƭ�ߴ粻��ȷ���˳�����
        Exit Sub
    End If
    
    '���ݸ�ʽ����ͼ��������������ָ�ʽ��ʾ��������STANDARD\1,2�����У��У�����ROW\1,2������COL\2,3�����ֱ����
    If InStr(UCase(strFormat), "STANDARD\") > 0 Then
        strCurFormat = Mid(strFormat, 10)
        If UBound(Split(strCurFormat, ",")) = 1 Then
            intX = Val(Split(strCurFormat, ",")(0))
            intY = Val(Split(strCurFormat, ",")(1))
            ReDim arrImageSize(intX * intY)
            For i = 1 To intY
                For j = 1 To intX
                    arrImageSize((i - 1) * intX + j).intWidth = lngFilmWidth / intX
                    arrImageSize((i - 1) * intX + j).intHeight = lngFilmHeight / intY
                Next j
            Next i
        Else
            Exit Sub
        End If
    ElseIf InStr(UCase(strFormat), "ROW\") > 0 Then
        strCurFormat = Mid(strFormat, 5)
        intY = UBound(Split(strCurFormat, ",")) + 1
        intCount = 0
        For i = 1 To intY
            intX = Val(Split(strCurFormat, ",")(i - 1))
            ReDim Preserve arrImageSize(UBound(arrImageSize) + intX)
            For j = 1 To intX
                intCount = intCount + 1
                arrImageSize(intCount).intWidth = lngFilmWidth / intX
                arrImageSize(intCount).intHeight = lngFilmHeight / intY
            Next j
        Next i
    ElseIf InStr(UCase(strFormat), "COL\") > 0 Then
        strCurFormat = Mid(strFormat, 5)
        intX = UBound(Split(strCurFormat, ",")) + 1
        intCount = 0
        For i = 1 To intX
            intY = Val(Split(strCurFormat, ",")(i - 1))
            ReDim Preserve arrImageSize(UBound(arrImageSize) + intY)
            For j = 1 To intY
                intCount = intCount + 1
                arrImageSize(intCount).intWidth = lngFilmWidth / intX
                arrImageSize(intCount).intHeight = lngFilmHeight / intY
            Next j
        Next i
    Else
        '��Ƭ���ֲ���ȷ���˳�����
        Exit Sub
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subFillPageRCCount(strFilmFormat As String)
'------------------------------------------------
'���ܣ��������沼�ִ�����д���沼������Ͷ�Ӧ����
'������ strFilmFormat���沼�ִ�
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strFormatType As String
    Dim strFormatNumber As String
    Dim strRCDetail() As String
    Dim intRCCount As Integer
    
    On Error GoTo err
    
    '������Ƭ����
    If strFilmFormat <> "" Then
        i = InStr(strFilmFormat, "\")
        strFormatType = Mid(strFilmFormat, 1, i - 1)
        strFormatNumber = Mid(strFilmFormat, i + 1, Len(strFilmFormat) - i)
    End If
    
    '�ж��������ȣ������ȣ���׼��ʽ��
    If strFormatType = "ROW" Then
        mblnIsRow = True
        mblnIsCustom = True
    ElseIf strFormatType = "COL" Then
        mblnIsRow = False
        mblnIsCustom = True
    ElseIf strFormatType = "STANDARD" Then
        mblnIsRow = True
        mblnIsCustom = False
    End If

    '��ȡ�����еľ���������ֵ�����浽aRCCount������
    strRCDetail = Split(strFormatNumber, ",")
    If UBound(strRCDetail) >= 0 Then
        If Not mblnIsCustom Then                     '��׼����������
            intRCCount = Val(strRCDetail(1))
            ReDim marrRCCount(intRCCount)              '''ÿ��/ÿ�е�ͼ����Ŀ
            For i = 1 To UBound(marrRCCount)
                marrRCCount(i) = Val(strRCDetail(0))
            Next
        Else
            intRCCount = UBound(strRCDetail) + 1            ''''����������Ŀ
            ReDim marrRCCount(intRCCount)              '''ÿ��/ÿ�е�ͼ����Ŀ
            For i = 1 To UBound(marrRCCount)
                marrRCCount(i) = Val(strRCDetail(i - 1))
            Next
        End If
    End If
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subImageSort(blnIncrease As Boolean)
'------------------------------------------------
'���ܣ��Ե�ǰ��Ƭ�У�ѡ�е�ͼ��������򣬰���ͼ�������
'������ blnIncrease -����
'���أ���
'------------------------------------------------
    Dim v As DicomViewer
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedCount As Integer
    Dim blnSortSelected As Boolean  'ֻ�Ա�ѡ�е�ͼ���������
    Dim intCurrImgIndex As Integer
    Dim intImgsPrintStartIndex As Integer   '��ʼ���������ͼ��Index
    Dim intImgsPrintEndIndex As Integer     '�������������ͼ��Index
    
    On Error GoTo err
    
    '���Ƚ���ǰ��Select״̬���µ�������
    Call subReloadImgsPrint
    
    '�����ǰ��Ƭ���б�ѡ�е�2������ͼƬ����Ա�ѡ�е�ͼ��������򣬷���Խ�Ƭ�е�����ͼ���������
    intSelectedCount = 0
    For Each v In FilmViewer
        For i = 1 To v.Images.Count
            If funGetTagVal(v.Images(i).Tag, TAG_ѡ��) = "Select" Then
                intSelectedCount = intSelectedCount + 1
                If intSelectedCount >= 2 Then
                    blnSortSelected = True
                    Exit For
                End If
            End If
        Next i
        If blnSortSelected = True Then Exit For
    Next v
    
    '��ͼ���������
    intImgsPrintStartIndex = funGetStartImgNo(VScro.Value, 1, 1)
    intImgsPrintEndIndex = funGetStartImgNo(VScro.Value, FilmViewer.Count - 1, 1) + marrPages(VScro.Value).ViewerLayout(FilmViewer.Count - 1).intColumns * marrPages(VScro.Value).ViewerLayout(FilmViewer.Count - 1).intRows - 1
    If intImgsPrintEndIndex > imgsPrint.Count Then intImgsPrintEndIndex = imgsPrint.Count
    
    '������ͼ���н�������
    For i = intImgsPrintStartIndex To intImgsPrintEndIndex
        If (blnSortSelected = True And funGetTagVal(imgsPrint(i).Tag, TAG_ѡ��) = "Select") Or blnSortSelected = False Then
            '��ʼ����ȶԣ��������ͼ���λ��
            Call subImageSortAndMove(i, intImgsPrintEndIndex, blnSortSelected, blnIncrease)
        End If
    Next i
    
    'ȫ���������֮��������ʾ
    '���ù��̵���ͼ�����ʾ
    Call subShowPrintImages(Me.VScro.Value)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subImageSortAndMove(intImgsPrintIndex As Integer, intImgsPrintEndIndex As Integer, _
    blnSortSelected As Boolean, blnIncrease As Boolean)
'------------------------------------------------
'���ܣ���intImageIndex��ʼ�����������򣬵������ͼ���λ��
'������ intImgsPrintIndex -��ǰ��ʼ���ҵ�ͼ��index��ԭͼ�е�index
'       intImgsPrintEndIndex -- ������ͼ���в��ҵĽ���Index
'       blnSortSelected -- ֻ�Ա�ѡ�е�ͼ���������
'       blnIncrease -- True ��������False ��������
'���أ���
'------------------------------------------------
    Dim intNextImageIndex As Integer
    Dim lngCurrImgNum As Long       '��¼��ǰͼ���ͼ���
    Dim lngTestingImgNum As Long    '��¼�����Ե�ͼ���ͼ���
    Dim v As DicomViewer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '��¼��ǰͼ���ͼ���
    lngTestingImgNum = 0
    intNextImageIndex = 0
    If imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Exists And Not IsNull(imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Value) Then
        lngCurrImgNum = Val(imgsPrint(intImgsPrintIndex).Attributes(&H20, &H13).Value)
    End If
    
    For i = intImgsPrintIndex + 1 To intImgsPrintEndIndex
        If (blnSortSelected = True And funGetTagVal(imgsPrint(i).Tag, TAG_ѡ��) = "Select") Or blnSortSelected = False Then
            '��ȡͼ���
            If imgsPrint(i).Attributes(&H20, &H13).Exists And Not IsNull(imgsPrint(i).Attributes(&H20, &H13).Value) Then
                lngTestingImgNum = Val(imgsPrint(i).Attributes(&H20, &H13).Value)
            End If

            If (blnIncrease = True And lngTestingImgNum < lngCurrImgNum And lngTestingImgNum <> 0) _
                Or (blnIncrease = False And lngTestingImgNum > lngCurrImgNum And lngTestingImgNum <> 0) Then
                intNextImageIndex = i
                lngCurrImgNum = lngTestingImgNum
            End If
        End If
    Next i
    
    '�ƶ�ͼ��
    If intNextImageIndex <> 0 Then
        Call imgsPrint.Move(intNextImageIndex, intImgsPrintIndex)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subShowPrintImages(intPage As Integer)
'------------------------------------------------
'���ܣ����ݽ�Ƭ��ʽ��������������������ʾָ��ҳ��ͼ�񣬲���ʾͼ����صĶ�λ�ߵ�
'������
'���أ���
'------------------------------------------------

    On Error GoTo err
    
     'ǰ��������Viewer��������λ�ã��������������Ѿ��������
     
    If Not mblnBegin Then Exit Sub
    
    '������ʾһҳ��ͼ��
    Call subLoadPrintImage(intPage)
    
    '��ʾ��λ��
    Call subDispReferLineFilm
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function ZLAddImage(img As DicomImage, blnPrinted As Boolean, dblWidth As Double, dblHeight As Double) As Long
'------------------------------------------------
'���ܣ�����һ��ͼ��
'������ img -- ��Ҫ���ӵ�ͼ��
'       blnPrinted -- ��¼ͼ���Ƿ��Ѿ�����ӡ��
'       dblWidth --  ͼ��ԭ����ʾʱռ�õĿ�ȣ���������ͼ������ű������ƶ�
'       dblHeight -- ͼ��ԭ����ʾʱռ�õĸ߶ȣ���������ͼ������ű������ƶ�
'���أ�0 -- ��ȷ��1--����
'------------------------------------------------
    Dim AddedImage As DicomImage
    Dim dblScale As Double
    Dim thisViewer As DicomViewer
    
    On Error GoTo err
    
    imgsPrint.Add img
    Set AddedImage = imgsPrint(imgsPrint.Count)
    
    '�����ӡ���
    If blnPrinted = True Then mblnPrinted = True
    
    '����ͼ��ı�ע
    If AddedImage.Labels(G_INT_SYS_LABEL_PAT_INFO).Visible = False Then
        Call subDispImageInfo(AddedImage, True, False, True)        ''��ʾ�����Ľ���Ϣ�ʹ���λ��Ϣ
    End If
    
    '����Ŵ���ͼ�����ʾ
    Set thisViewer = FilmViewer(FilmViewer.Count - 1)
    
    'Ӧ��ͳһ���� subScaleImage ���̣������������ͼ����ʾ���ᣬ�Ȳ�ʹ��
'    Call subScaleImage(AddedImage, thisViewer, CLng(dblWidth), CLng(dblHeight))
    
    dblScale = ((thisViewer.width / thisViewer.MultiColumns) / dblWidth + (thisViewer.height / thisViewer.MultiRows) / dblHeight) / 2
    AddedImage.Zoom = AddedImage.Zoom * dblScale
    AddedImage.ScrollX = AddedImage.ScrollX * dblScale
    AddedImage.ScrollY = AddedImage.ScrollY * dblScale
    
    '������ͼ�񣬵�������
    Call subChangeLabelForPrint(AddedImage, 0)
    'ͼ�������ˣ�����ҳ��
    Call subRecalPages
    
    '�������ӵ�ͼ���ڵ�ǰҳ����������ʾ��һҳ��ͼ��
    If imgsPrint.Count < funGetStartImgNo(Me.VScro.Value, 1, 1) + marrPages(Me.VScro.Value).intImageCount Then
        Call subShowPrintImages(Me.VScro.Value)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    ZLAddImage = 1
End Function


Private Function funChangeFormat(intPage As Integer, strPageFormat As String, Optional intViewerIndex As Integer = 0, _
    Optional intRows As Integer = 0, Optional intCols As Integer = 0) As Long
'------------------------------------------------
'���ܣ�����ҳ�����ʾ���֣�������Ƭ���ֺ����ͼ�񲼾�
'������ intPage --      ��Ҫ������ҳ��
'       strPageFormat --ҳ���е�Viewer���֣�DICOM��ʽ��
'       intViewerIndex--����ѡ��ͼ�����ʱ�ã�0-����ͼ����ϣ�����-����ͼ����ϵ�Viewer��
'       intRows --      ����ѡ��ͼ�����ʱ�ã�0-����ͼ����ϣ�����-����ͼ����ϵ�����
'       intCols --      ����ѡ��ͼ�����ʱ�ã�0-����ͼ����ϣ�����-����ͼ����ϵ�����
'���أ�0 -- ��ȷ��1--����
'------------------------------------------------
    Dim intCurrViewerCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    funChangeFormat = 1
    If intPage > UBound(marrPages) Then Exit Function
    If intViewerIndex <> 0 Then
        If intRows = 0 Then intRows = 1
        If intCols = 0 Then intCols = 1
    End If
    
    '�ȼ�¼Viewer������ҳ�沼��
    intCurrViewerCount = funGetPageViewerCount(strPageFormat)
    
    '������ǰҳ���е�ͼ����ϲ���
    If intViewerIndex <> 0 And intViewerIndex <= intCurrViewerCount Then
        marrPages(intPage).ViewerLayout(intViewerIndex).intColumns = intCols
        marrPages(intPage).ViewerLayout(intViewerIndex).intRows = intRows
    End If
    
    '������ǰҳ��ͺ���ҳ���ҳ�沼��
    For i = intPage To UBound(marrPages)
        '��д��һҳ��Viewer������ͼ������
        marrPages(i).intViewerCount = intCurrViewerCount
        marrPages(i).strPageFormat = strPageFormat
        
        'Viewer�������ӻ��߼����ˣ���Ҫ�������ҳ���е�ViewerLayout����,ȷ������������������=1
        ReDim Preserve marrPages(i).ViewerLayout(intCurrViewerCount)
        
        For j = 1 To intCurrViewerCount
            '����ͼ����ϵĲ���
            If marrPages(i).ViewerLayout(j).intColumns = 0 Then marrPages(i).ViewerLayout(j).intColumns = 1
            If marrPages(i).ViewerLayout(j).intRows = 0 Then marrPages(i).ViewerLayout(j).intRows = 1
        Next j
        '������һҳ��ͼ������������ʵ��ͼ������������һҳ�ܹ��ܹ��ڷŵ�ͼ������
        marrPages(i).intImageCount = funGetPageImageCount(i)
    Next i
    
    '����ͼ������������������������¼��㽺Ƭҳ��
    Call subRecalPages
    
    funChangeFormat = 0
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub subRecalPages()
'------------------------------------------------
'���ܣ���������ͼ������������¼���ʵ��ҳ��
'������
'���أ�
'------------------------------------------------
    Dim intPageCount As Integer
    Dim intImageCount As Integer
    Dim intDefaultViewerCount As Integer
    Dim strDefaultFormat As String
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo err
    
    '����ͼ������������ȫ��ҳ�沼��
    intPageCount = 0
    intImageCount = 0
    
    '���һҳ�Ĳ�������ΪĬ�ϲ���
    intDefaultViewerCount = marrPages(UBound(marrPages)).intViewerCount
    strDefaultFormat = marrPages(UBound(marrPages)).strPageFormat
    
    '�����ǰһ��ͼ��û�У�ҳ�����ó�һҳ
    ReDim Preserve marrPages(IIf(imgsPrint.Count > 0, imgsPrint.Count, 1))
    
    For i = 1 To imgsPrint.Count
        '����������ҳ�棬����ǰĬ�ϸ�ʽӦ�õ�����ҳ�棬ֻӦ��ҳ�沼�֣���Ӧ��ͼ����ϲ���
        If marrPages(i).intViewerCount = 0 Then
            '��д��һҳ��Viewer������ͼ������
            marrPages(i).intViewerCount = intDefaultViewerCount
            marrPages(i).strPageFormat = strDefaultFormat
            
            'Viewer�������ӻ��߼����ˣ���Ҫ�������ҳ���е�ViewerLayout����,ȷ������������������=1
            ReDim Preserve marrPages(i).ViewerLayout(intDefaultViewerCount)
            
            For j = 1 To intDefaultViewerCount
                marrPages(i).ViewerLayout(j).intColumns = 1
                marrPages(i).ViewerLayout(j).intRows = 1
            Next j
            
            '������һҳ��ͼ������������ʵ��ͼ������������һҳ�ܹ��ܹ��ڷŵ�ͼ������
            marrPages(i).intImageCount = funGetPageImageCount(i)
        End If
        
        intPageCount = intPageCount + 1
        intImageCount = intImageCount + marrPages(i).intImageCount
        
        '����ۼӵ�ͼ��������������ͼ�����������˳�ѭ�����õ���ǰ�Ľ�Ƭҳ����
        If intImageCount >= imgsPrint.Count Then Exit For
    Next i
    
    ReDim Preserve marrPages(IIf(intPageCount > 0, intPageCount, 1))
    
    '���ù����������ֵ
    Me.VScro.Max = UBound(marrPages)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function funGetPageViewerCount(strPageFormat As String) As Integer
'------------------------------------------------
'���ܣ�����ҳ�沼�ַ�ʽ�����㵱ǰҳ���Viewer����
'������ strPageFormat -- DICOM��׼��ҳ�沼��
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim strFormatType As String
    Dim strFilmFormat As String
    Dim strRCDetail() As String
    
    funGetPageViewerCount = 0
    
    On Error GoTo err
    
    If strPageFormat = "" Then Exit Function
    
    '������Ƭ��ҳ�沼��
    i = InStr(strPageFormat, "\")
    strFormatType = Mid(strPageFormat, 1, i - 1)
    strFilmFormat = Mid(strPageFormat, i + 1, Len(strPageFormat) - i)
    
    strRCDetail = Split(strFilmFormat, ",")
    If UBound(strRCDetail) >= 0 Then
        If strFormatType = "STANDARD" Then
            '��׼���֣�ֱ��=����*����
            funGetPageViewerCount = Val(strRCDetail(0)) * Val(strRCDetail(1))
        Else
            '���Բ��֣������ȣ����������ȣ�������
            For i = 0 To UBound(strRCDetail)
                funGetPageViewerCount = funGetPageViewerCount + Val(strRCDetail(i))
            Next i
        End If
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    funGetPageViewerCount = 0
End Function

Private Function funGetPageImageCount(intPage As Integer) As Integer
'------------------------------------------------
'���ܣ�����ҳ�沼�ַ�ʽ�����㵱ǰҳ�������ɵ�ͼ������
'������ intPage -- ��Ҫ�����ҳ��
'���أ���ǰҳ�������ɵ�ͼ������
'------------------------------------------------
    Dim i As Integer
    Dim intImageCount As Integer
    
    On Error GoTo err
    
    funGetPageImageCount = 0
    If intPage > UBound(marrPages) Then Exit Function
    
    intImageCount = 0
    For i = 1 To marrPages(intPage).intViewerCount
        intImageCount = intImageCount + (marrPages(intPage).ViewerLayout(i).intColumns * marrPages(intPage).ViewerLayout(i).intRows)
    Next i
    
    funGetPageImageCount = intImageCount
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    
End Function

Private Sub InitPageFormat(strPageFormat As String)
'------------------------------------------------
'���ܣ���ʼ��ҳ��Ĳ�������
'������strPageFormat -- DICOM��ʽ��ҳ�沼��
'���أ���
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    'һ��ͼ��û�е�ʱ��marrPages���ó�1ҳ
    
    ReDim marrPages(1)
    marrPages(1).strPageFormat = strPageFormat

    '���¼�����һҳ�е�Viewer����
    marrPages(1).intViewerCount = funGetPageViewerCount(strPageFormat)
    ReDim marrPages(1).ViewerLayout(marrPages(1).intViewerCount)
    
    'ÿһ��Viewer�����ó�һ��Xһ��
    For i = 1 To marrPages(1).intViewerCount
        marrPages(1).ViewerLayout(i).intColumns = 1
        marrPages(1).ViewerLayout(i).intRows = 1
    Next i
    
    '��д��һҳ�����ͼ������
    marrPages(1).intImageCount = funGetPageImageCount(1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function funGetTagVal(strTag As String, TagID As String) As String
'------------------------------------------------
'���ܣ���TAG����ȡ��Ӧ��ֵ��ĿǰTAG�ܹ�����4��ֵ
'������ strTag --- ��Ҫ��ȡ��TAG
'       TagID --- TAG��������ʹ�ö���õĳ���
'���أ�TagID��Ӧ��ֵ
'------------------------------------------------
    Dim arrTags() As String
    Dim intTagID As Integer
    
    On Error GoTo err
    intTagID = Val(TagID)
    If intTagID <= 0 Or intTagID > 4 Then Exit Function
    If strTag = "" Then Exit Function
    
    arrTags = Split(strTag, zlSpliter)
    If UBound(arrTags) = 4 Then
        funGetTagVal = arrTags(intTagID)
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funSetTagVal(dcmImage As DicomImage, TagID As String, strTagVal As String) As Boolean
'------------------------------------------------
'���ܣ���TAG����ȡ��Ӧ��ֵ��ĿǰTag�ܹ�����4��ֵ
'������ dcmImage --- ��Ҫ��ȡ��TAG��DICOMͼ��
'       TagID --- TAG��������ʹ�ö���õĳ���
'       strTagVal --- Ҫ���õ�ֵ
'���أ�True�ɹ���False ʧ��
'------------------------------------------------
    
    Dim strTag As String
    Dim arrTags() As String
    Dim intTagID As Integer
    Dim i As Integer
    
    On Error GoTo err
    
    If dcmImage Is Nothing Then Exit Function
    
    strTag = dcmImage.Tag
    intTagID = Val(TagID)
    If intTagID <= 0 Or intTagID > 4 Then Exit Function
    arrTags = Split(strTag, zlSpliter)
    
    If UBound(arrTags) < intTagID Then
        ReDim Preserve arrTags(intTagID) As String
    End If
    arrTags(intTagID) = strTagVal
    
    strTag = ""
    For i = 1 To UBound(arrTags)
        strTag = strTag & zlSpliter & arrTags(i)
    Next i
    If i <= 4 Then
        For i = i To 4
            strTag = strTag & zlSpliter & ""
        Next i
    End If
    
    dcmImage.Tag = strTag
    funSetTagVal = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subShowOnePage(intPage As Integer)
'------------------------------------------------
'���ܣ���ʾ��ǰҳ���ͼ��������ͼ������û�䣬����ҳ����߲��ָı�����
'������ intPage -- ��Ҫ��ʾ��ҳ��
'���أ� ��
'------------------------------------------------

    On Error GoTo err
    
    '����ҳ�沼�֣�����ҳ���е�Viewer
    Call subLoadViewer(intPage)

    '���ù��̵���ͼ�����ʾ
    Call subShowPrintImages(intPage)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub MouseWheel(intDirection As Integer)
'------------------------------------------------
'���ܣ����������ֵ���Ϣ
'������intDirection--���ֹ������� 1-����Ϲ���2-����¹�
'���أ���
'------------------------------------------------
    Dim dblScale As Double
    
    '�������󣬲����κ���ʾ
    On Error Resume Next
    
    If SelectedImage Is Nothing Then Exit Sub
    If mintSelectedViewer > FilmViewer.Count Then Exit Sub
    If FilmViewer.Count = 1 Then Exit Sub
    
    If intDirection = 1 Then
        dblScale = 1 + (lngZoomStep * 0.01)
    Else
        dblScale = 1 - (lngZoomStep * 0.01)
    End If
    
    Call subCenterZoom(SelectedImage, FilmViewer(mintSelectedViewer), SelectedImage.ActualZoom * dblScale)
        
    If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
        If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
            Call UpdateRuler(SelectedImage, True)
        End If
    End If
    
    '����ͬ��
    Call subSynchronalImg(False, IMG_SYN_ZOOMPAN)
End Sub

