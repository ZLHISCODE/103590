VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.Form frmImageProcessV2 
   Caption         =   "ͼ����"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   Icon            =   "frmImageProcessV2.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zl9PacsControl.ucSplitter ucSplitter 
      Height          =   6375
      Left            =   4215
      TabIndex        =   8
      Top             =   480
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   11245
      Control1Name    =   "ucBgImages"
      Control2Name    =   "DViewer"
   End
   Begin zl9PACSWork.ucBgImgViewer ucBgImages 
      Height          =   6375
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
   End
   Begin VB.ListBox lstMemoText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8280
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.PictureBox picMemo 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11055
      TabIndex        =   2
      Top             =   7080
      Width           =   11055
      Begin VB.PictureBox picCboDropDown 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   8040
         Picture         =   "frmImageProcessV2.frx":6852
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
      Begin VB.ComboBox cbxMemoText 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   5655
      End
      Begin VB.CommandButton cmdFont 
         Height          =   375
         Left            =   9000
         Picture         =   "frmImageProcessV2.frx":6BAE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "���õ�ǰ��ע���塣"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdInsert 
         Height          =   375
         Left            =   8640
         Picture         =   "frmImageProcessV2.frx":6EF0
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "����ǰ��ע����Ϊ���ñ�ע"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   375
         Left            =   8280
         Picture         =   "frmImageProcessV2.frx":765A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "��ӱ�ע"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblMemoText 
         AutoSize        =   -1  'True
         Caption         =   "��ӱ�ע���֣�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   6
         Top             =   195
         Width           =   1470
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   2520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtInputText 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   6375
      Left            =   4350
      TabIndex        =   7
      Top             =   480
      Width           =   6255
      _Version        =   262147
      _ExtentX        =   11033
      _ExtentY        =   11245
      _StockProps     =   35
      BackColor       =   0
      UseScrollBars   =   0   'False
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmImageProcessV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TPoint
  X As Integer
  Y As Integer
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type


'Private gobjImageProcess As frmImageProcess

Private glngColor(10) As Long             '���ͼ��Բ�α��ʹ�õ�9����ɫ

Private Const G_STR_TAG = "Po=Ϣ��[+]E=������[+]M=��Ƕ[+]L=ճĤ�װ�[+]C=ʪ��[+]I=�����԰�[+]W=�����ɫ��Ƥ[+]AT=�쳣ת����[+]V=�ǵ���Ѫ��[+]P=��״Ѫ��[+]Xn=ֱ�ӻ�첿λ"

'ͼ����
Private Const conMenu_Process_Window = 501           '���ȶԱȶ�
Private Const conMenu_Process_Zoom = 502             '����
Private Const conMenu_Process_Corp = 512             '�϶�
Private Const conMenu_Process_RRotate = 503          '˳ʱ����ת
Private Const conMenu_Process_LRotate = 504          '��ʱ����ת
Private Const conMenu_Process_Sharpness = 505        '��
Private Const conMenu_Process_Filter = 506           'ƽ��
Private Const conMenu_Process_Arrow = 507            '��ͷ��ע
Private Const conMenu_Process_Ellipse = 508          'Բ�α�ע
Private Const conMenu_Process_Text = 509             '���ֱ�ע
Private Const conMenu_Process_RectZoom = 510         '�ü��ɼ�
Private Const conMenu_Process_RectCapture = 511      '�ü���ɼ�
Private Const conMenu_Process_Line = 520             'ֱ�߱�ע
Private Const conMenu_Process_Exit = 2613            '�˳�
Private Const conMenu_Process_Save = 3091            '����
Private Const conMenu_Process_SaveToReport = 3941    '���浽���
Private Const conMenu_Process_SaveToStudy = 3943     '���浽����
Private Const conMenu_Process_DelAllLabels = 8113    'ɾ��ȫ����ע��ʹ������ϵͳ��ͼ����
Private Const conMenu_Process_MoveLabel = 6891       '�ƶ���ɾ��ѡ�б�ע��ʹ������ϵͳ��ͼ����
Private Const conMenu_Process_LabelSetUp = 10003     '��ע��ť���ã�ʹ������ϵͳ��ͼ����
Private Const conMenu_Process_Restore = 8124         '�ָ�
Private Const conMenu_Process_TextTag = 5010         '�ı����
Private Const conMenu_Process_NumTag = 7405          '���ֱ��
Private Const conMenu_Process_Page = 1001
Private Const conMenu_Process_Num = 96
Private Const conMenu_Process_Word = 97



'Private mlngModule As Long
Private mlngAdviceId As Long

Private mImage As DicomImage
Private mintMouseState As TMouseState
Private mblnDcmViewDown As Boolean
Private mMouseDownPoint As TPoint
Private mInitScrollPoint As TPoint
Private mCorpSize As TPoint             '�϶�������ƫ��λ��

'����������ʹ�õ�����׼λ��
Private mlngBaseXX As Long
Private mlngBaseYY As Long
'�ƶ���עʹ�õ�����׼λ��
Private mlngBaseX As Long
Private mlngBaseY As Long

Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mMovingLabel As DicomLabel      '��ǰѡ��Ҫ�ƶ�����ɾ���ı�ע

Private mblnOk As Boolean
Private mOldImage As DicomImage
Private mlngImgIndex As Long            '������ѡ������ͼ������
Private mintTextIndex As Integer        '���ֱ�ע��ť������
Private mstrText As String              '���ֱ�ע����
Private mstrCustom  As String           '�Զ����ע����
Private mintNumberIndex As Integer      '���ֱ�Ű�ť������
Private mintAutoNumber As Integer       '�Զ�������ŵ�������
Private mStrTemp As String
Private mstrUser As String
 
Private mlngWinType As TImgProcessType  '�򿪴���ʱ��������

Private mlngPreViewTime As Long         '�ƶ�Ԥ����ʱ�ر�ʱ��
Private mlngState As Long               'Ԥ��ͼ�񴰿�״̬��1-Ԥ����2-����3-������
Private mblnMoved As Boolean
Private mstrQueryValue As String
Private mblnIsUnloud As Boolean         '��ǰ���λ���Ƿ��Զ��ر�
Private mblnDrag As Boolean
Private mintDisState As Integer
Private mblnIsChanged As Boolean
Private mblnCase As Boolean
Private mblnIsReportShow As Boolean
Private maryImgInfos() As Object

Private mrsTmp As ADODB.Recordset       'ͼ��ע��¼��

Private Enum TMouseState
    msNone = 0          '��״̬
    msWinLevel = 1      '����λ
    msZoom = 2          '����
    msRectangle = 3     '��ѡ����
    msline = 10         'ֱ��
    msArrow = 11        '��ͷ
    msEllipse = 12      '��Բ
    msText = 13         '����
    msDrag = 14         '�����϶�
    msNumber = 15       '���ֱ��
    msFixText = 16      '���ְ�ť
    msMove = 17         '�ƶ���ɾ����ע
End Enum


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Event OnUnload()
Public Event OnSaveImage(ByVal emImageType As TImageType, ByRef dcmImage As DicomImage)

Private mblnAllowSaveStudyImg As Boolean    '��������ͼ
Private mblnAllowSaveReportImg As Boolean   '�����汨��ͼ


Public Sub SetButtonState(ByVal blnStudyImgSaveState As Boolean, ByVal blnRepImgSaveState As Boolean)
    mblnAllowSaveStudyImg = blnStudyImgSaveState
    mblnAllowSaveReportImg = blnRepImgSaveState
End Sub



Property Get WinType() As Long
    WinType = mlngWinType
End Property

Private Sub LoadImgs(objImgInfos() As Object)
    Dim i As Long
    Dim objImgInf As clsBgImgInfo
    
    For i = 0 To UBound(objImgInfos)
        If Not objImgInfos(i) Is Nothing Then
            Set objImgInf = objImgInfos(i).CopyNew
            objImgInf.ImgCommand = icDownload
            objImgInf.LoadState = lsNone
            
            Call Me.ucBgImages.ConstructionImgData(objImgInf)
        End If
    Next
    
    Call Me.ucBgImages.Refresh
End Sub

Public Function ZlShowMe(objParent As Object, ByVal lngAdviceId As Long, _
    objSelImg As DicomImage, objImgInfos() As Object, _
    Optional lngWindowType As TImgProcessType = ptPreview, Optional lngPreviewTime As Long = 0, _
    Optional blnIsReportShow As Boolean) As Boolean
'lngType:�������ͣ�0-ͼ�����ڣ�1-ͼ��Ԥ�����ڣ�2-���ͼ������
    
    On Error GoTo err
    
    Dim i As Integer
    Dim oldWinType As TImgProcessType
    Dim arrImages() As String
 
    oldWinType = mlngWinType
     
    mlngWinType = lngWindowType
    mlngPreViewTime = lngPreviewTime
    mblnIsReportShow = blnIsReportShow
    mstrUser = GetUserInfo
    
    If mblnIsChanged Then
        If MsgBoxD(Me, "����δ�����ͼ�����ò����������Щ�����Ƿ������", vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    
    mblnDrag = False
    mblnIsChanged = False
    mblnCase = False
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    If mlngWinType = ptPreview Then mlngState = 1
    If mlngWinType = ptProcess Then mlngState = 2
    
    If mlngWinType <> ptMark Then
        Me.ucBgImages.IsShowCheck = False
        
        '������ͼ��
        If mlngAdviceId <> lngAdviceId Then
            Call Me.ucBgImages.ClearAll
            
            '���ֱ����ͼ������ͬʱ��������ͼ
            If mlngWinType = ptProcess Then
                Call LoadImgs(objImgInfos)
            Else
                '�Ƚ������鸳ֵ����timer1���Ӻ����
                maryImgInfos = objImgInfos
            End If
        Else
            '������ȵ���Ԥ�����ڣ�Ȼ���ڽ���ͼ��������Ҫ�ж�ͼ�������Ƿ�Ϊ0����Ϊ��Ԥ��ʱ����û�м�������ͼ��
            If mlngWinType = ptProcess And ucBgImages.ImgCount <= 0 Then
                Call LoadImgs(objImgInfos)
            End If
        End If
    Else
        Call Me.ucBgImages.ClearAll
    End If
    
    Set mOldImage = objSelImg
        
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add objSelImg
     
    If mlngAdviceId <> lngAdviceId Or Me.Visible = False Then
        Me.Show 0, objParent
        SetWindowPos Me.hwnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3  '�������ö�
    End If
    
    Call RefrshObjVisible
    
    
    If mlngWinType = ptPreview Then
        If DViewer.Images.Count > 0 Then
            Call DrawHintTag(DViewer.Images(1))
        End If
            
        Timer1.Enabled = True

        If lngPreviewTime > 0 Then
            Timer2.Interval = lngPreviewTime * 1000
            Timer2.Enabled = True
        End If
    Else
        refreshFace
    End If
    
    If oldWinType <> mlngWinType Then Call RestorceWinLayout
    
    mlngAdviceId = lngAdviceId
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function



Private Sub RefrshObjVisible()
    Dim blnVisible As Boolean
    
    If mlngWinType = ptMark Then
        Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = True
        Me.ucSplitter.Visible = False
        Me.ucBgImages.Visible = False
        Me.lblMemoText.Visible = False
        Me.cbxMemoText.Visible = False
        Me.picCboDropDown.Visible = False
        Me.cmdFont.Visible = False
        Me.cmdAdd.Visible = False
        Me.picMemo.Visible = False
        Me.lstMemoText.Visible = False
        Me.txtInputText.Visible = False
        
        Me.Caption = "���ͼ"
    Else
        If mlngState = 1 Then
            blnVisible = False
        Else
            blnVisible = True
        End If
        
        Me.ucBgImages.Visible = blnVisible
        Me.ucSplitter.Visible = blnVisible
        Me.lblMemoText.Visible = blnVisible
        Me.cbxMemoText.Visible = blnVisible
        Me.picCboDropDown.Visible = blnVisible
        Me.cmdInsert.Visible = blnVisible
        Me.cmdFont.Visible = blnVisible
        Me.cmdAdd.Visible = blnVisible
        Me.picMemo.Visible = blnVisible
        
        Me.cbrMain.FindControl(, conMenu_Process_Window).Parent.Visible = blnVisible

        If Me.lstMemoText.Visible Then Me.lstMemoText.Visible = blnVisible
        If Me.txtInputText.Visible Then Me.txtInputText.Visible = blnVisible
        
        Me.Caption = IIf(mlngWinType = ptPreview, "ͼ��Ԥ��", "ͼ����")
    End If
     
End Sub

Private Sub ClearLable(dcmImage As DicomImage)
    Dim i As Long
     'ȥ���߿�
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).tag = "SELECT" Or dcmImage.Labels(i).tag = "BORDER" Or dcmImage.Labels(i).tag = "HINT" Then
            dcmImage.Labels(i).Visible = False
        End If
    Next
    dcmImage.BorderColour = vbWhite
End Sub

Private Function getListIndex() As Integer
'���ݼ���������ȡ����
    Dim i As Integer

    getListIndex = -1
    
    If mrsTmp.RecordCount <= 0 Then Exit Function

    mrsTmp.MoveFirst
    
    If cbxMemoText.Text = "" Then Exit Function

    For i = 0 To mrsTmp.RecordCount - 1
        If InStr(Trim(nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Or InStr(Trim(nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Then
            getListIndex = i
            
            Exit For
        End If

        mrsTmp.MoveNext
    Next
End Function

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Select Case Control.ID
        Case conMenu_Process_Save           '���ͼ����
            If Control.Visible Then Control.Enabled = mblnAllowSaveReportImg
            
        Case conMenu_Process_SaveToStudy         '���浽���
            If Control.Visible Then Control.Enabled = mblnAllowSaveStudyImg
            
        Case conMenu_Process_SaveToReport           '���浽����
            If Control.Visible Then Control.Enabled = mblnAllowSaveReportImg
    End Select
Exit Sub
errHandle:
End Sub

Private Sub cbxMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
    End If
End Sub

Private Sub cmdAdd_Click()
'------------------------------------------------
'���ܣ���Ӳ���
'������
'���أ���
'------------------------------------------------
    On Error GoTo err

    'ƴ�ӷ���
    Call subAddMemoText


    Me.DViewer.Refresh

    '���ComboBox�ı�
    cbxMemoText.Text = ""

    '�ر�������
    lstMemoText.Visible = False
     
    Call EnterProcessState
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Function GetNewImage(emImageType As TImageType) As DicomImage
    Dim dcmImage As DicomImage
    Dim img As New DicomImage
    Dim iPlane As Integer
    Dim aryDcm() As Byte
      
    If Me.DViewer.Images.Count = 1 Then
        Set dcmImage = Me.DViewer.Images(1)
          
        If emImageType <> mtTagImage Then
On Error GoTo errRead
            'ת��һ��ͼƬ��ʽ�������ע
            aryDcm = dcmImage.ArrayExport("BMP")
            img.ArrayImport aryDcm, "BMP"
errRead:
            If err.Number <> 0 Then
                '���ͼƬ�ü�С�˺󱨴�����
                Set GetNewImage = dcmImage.PrinterImage(8, iPlane, True, 1, 0, dcmImage.SizeX, 0, dcmImage.SizeY)
            Else
                Set GetNewImage = img
            End If
            
            err.Clear
            
            GetNewImage.InstanceUID = CreateUID
            GetNewImage.SeriesUID = dcmImage.SeriesUID
            GetNewImage.StudyUID = dcmImage.StudyUID
            
            If emImageType = mtReportImage Then
                GetNewImage.BorderWidth = 1
                GetNewImage.BorderColour = vbWhite
            End If
        Else
            dcmImage.InstanceUID = CreateUID
            
            Set GetNewImage = dcmImage
        End If
    Else
        Set GetNewImage = Nothing
    End If
End Function

Private Sub SaveImage(emImageType As TImageType)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim objDcmImg As DicomImage
    Dim objSourceImgInfo As clsBgImgInfo
    
    Set objDcmImg = GetNewImage(emImageType)
    If objDcmImg Is Nothing Then Exit Sub
    
    If emImageType <> mtTagImage Then
        '���Ϊ���ͼ����ʱ���ǲ����ڶ�Ӧ����ͼ��ʾ��
        Set objSourceImgInfo = ucBgImages.ImageInfo(0).CopyNew()
        
        objSourceImgInfo.Key = objDcmImg.InstanceUID
        objSourceImgInfo.Filename = objDcmImg.InstanceUID
        objSourceImgInfo.ImgCommand = icReadly
        objSourceImgInfo.LoadState = lsNone
        objSourceImgInfo.Format = ifDcm
        objSourceImgInfo.JpgConvert = True
        objSourceImgInfo.IsReDrawed = False
        objSourceImgInfo.ErrorInfo = ""
        objSourceImgInfo.Redo = 0
        
        
        If FileExists(objSourceImgInfo.FilePath & objSourceImgInfo.Filename) = False Then
            objDcmImg.WriteFile objSourceImgInfo.FilePath & objSourceImgInfo.Filename, True, "1.2.840.10008.1.2.1"
        End If
        
        RaiseEvent OnSaveImage(emImageType, objDcmImg)
        
        strSQL = "select a.ͼ���,b.���к� from Ӱ����ͼ�� a , Ӱ�������� b where a.ͼ��UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯͼ���", objDcmImg.InstanceUID)
        
        If rsData.RecordCount > 0 Then
            objSourceImgInfo.ImageOrder = nvl(rsData!ͼ���)
            objSourceImgInfo.SeriesNoTag = nvl(rsData!���к�)
        End If
        
        ucBgImages.AddImg objSourceImgInfo
    Else
        RaiseEvent OnSaveImage(emImageType, objDcmImg)
    End If
    
    
    mblnIsChanged = False
 
    If emImageType = mtTagImage Then
        Unload Me
    End If
End Sub


Private Sub DViewer_Click()
On Error GoTo err
     
    Call EnterProcessState
        
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub DViewer_DblClick()
    Dim ls As DicomLabels
    Dim l As DicomLabel
    
    On Error GoTo err
    
    If mintMouseState = msMove Then
        Set ls = DViewer.LabelHits(mlngBaseXX, mlngBaseYY, False, False, True)
        If ls.Count > 0 Then
            If MsgBoxD(Me, "�Ƿ�ɾ�������ע��", vbOKCancel, "��ʾ") = vbOK Then
                Set l = ls(1)
                If l.tag <> "" Then
                    '�Ǳ�ű�ע����Ҫͬʱɾ��������ע����ɾ������
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject.TagObject))
                    End If
                    If DViewer.Images(1).Labels.IndexOf(l.TagObject) <> 0 Then
                        Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l.TagObject))
                    End If
                End If
                '����ͨ��ע�����߱�ŵ����һ����ע��ֱ��ɾ������
                Call DViewer.Images(1).Labels.Remove(DViewer.Images(1).Labels.IndexOf(l))
                DViewer.Refresh
            End If
        End If
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub DViewer_MouseWheel(ByVal Shift As Long, ByVal Delta As Integer, ByVal X As Long, ByVal Y As Long)
'    Call ucMiniature.MouseWheel(Delta)
'End Sub

Private Function ThumbnailImgCount()
On Error GoTo errHandle
    ThumbnailImgCount = UBound(maryImgInfos()) + 1
Exit Function
errHandle:
    ThumbnailImgCount = 0
End Function
  
Private Sub Form_Terminate()
    Dim i As Long
    
    Set mImage = Nothing
    Set mdcmSelectLabel = Nothing
    Set mMovingLabel = Nothing
    Set mOldImage = Nothing
    Set mrsTmp = Nothing
    
    For i = 0 To ThumbnailImgCount - 1
        Set maryImgInfos(i) = Nothing
    Next
    
    Erase maryImgInfos
End Sub

Private Sub lstMemoText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
End Sub

Private Sub picCboDropDown_Click()
    lstMemoText.ZOrder
    lstMemoText.Visible = Not lstMemoText.Visible
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex

    If lstMemoText.Visible Then lstMemoText.SetFocus
     
    Call EnterProcessState
End Sub

Private Sub EnterProcessState()
    mlngState = 3
    
    mlngPreViewTime = 0
    
    Timer1.Enabled = False
    Timer2.Enabled = False
     
    If mlngWinType = ptPreview Then mlngWinType = ptProcess
End Sub

Private Sub cbxMemoText_Change()
    lstMemoText.ZOrder
    
    If Not lstMemoText.Visible Then lstMemoText.Visible = True
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
End Sub

Private Sub cbxMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cbxMemoText.ListIndex = lstMemoText.ListIndex
        lstMemoText.Visible = False
        
        cbxMemoText.SelStart = 0
        cbxMemoText.SelLength = Len(cbxMemoText.Text)
        cbxMemoText.SetFocus
    End If
    
    If KeyAscii = vbKeyEscape Then lstMemoText.Visible = False
End Sub

Private Sub cmdFont_Click()
On Error GoTo errHandle
    diaFont.flags = 1
    diaFont.FontBold = Me.Font.Bold
    diaFont.FontItalic = Me.Font.Italic
    diaFont.FontName = Me.Font.Name
    diaFont.FontSize = Me.Font.Size
    diaFont.FontStrikethru = Me.Font.Strikethrough
    diaFont.FontUnderline = Me.Font.Underline

    
    diaFont.ShowFont
    
    Me.Font.Bold = diaFont.FontBold
    Me.Font.Italic = diaFont.FontItalic
    Me.Font.Name = diaFont.FontName
    Me.Font.Size = diaFont.FontSize
    Me.Font.Strikethrough = diaFont.FontStrikethru
    Me.Font.Underline = diaFont.FontUnderline
    
    Call EnterProcessState
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long
    
    On Error GoTo errHandle
     
    Call EnterProcessState
    
    mblnDrag = False
    Select Case Control.ID
        Case conMenu_Process_Window         '���ȶԱȶ�
            subSetMouseState 1
            'Control.Checked = True
            
        Case conMenu_Process_Zoom           '����
            subSetMouseState 2
            'Control.Checked = True
            
        Case conMenu_Process_RectZoom       '�ü�����
            subSetMouseState msRectangle
            'Control.Checked = True
        
        Case conMenu_Process_RectCapture         '�ü���ɼ�
            Call CaptureFrameSelectImage
            
        Case conMenu_Process_RRotate        '˳ʱ����ת
            subSetRotate True
            
        Case conMenu_Process_LRotate        '��ʱ����ת
            subSetRotate False
            
        Case conMenu_Process_Sharpness      '��
            subSetSharp True
            
        Case conMenu_Process_Filter         'ƽ��
            subSetSharp False
        
        Case conMenu_Process_Line           'ֱ�߱�ע
            subSetMouseState msline
        
        Case conMenu_Process_Arrow          '��ͷ��ע
            subSetMouseState msArrow
            
        Case conMenu_Process_Ellipse        'Բ�α�ע
            subSetMouseState msEllipse
            
        Case conMenu_Process_TextTag           '���ֱ�ע
            mstrText = ""
            subSetMouseState msText
            
        Case conMenu_Process_DelAllLabels   '�����ע
            lngCount = DViewer.Images(1).Labels.Count
            
            DViewer.Images(1).Labels.Clear
            DViewer.Refresh
            
            mintAutoNumber = 0
            
            If lngCount <> DViewer.Images(1).Labels.Count Then
                mblnIsChanged = True
            End If
            
        Case conMenu_Process_MoveLabel      '�ƶ�
            mblnDrag = True
            subSetMouseState msMove
            
        Case conMenu_Process_LabelSetUp     '��ע����
            Call subSetTextLabel
            
        Case conMenu_Process_Restore        '�ָ�
            DViewer.Images.Clear
            DViewer.Images.Add mOldImage
            
            If DViewer.Images.Count > 0 Then
                ClearLable DViewer.Images(1)
            End If
            
            '�ؽ���ע֮��Ĺ���
            Call subLabelCopyRebuild(mOldImage, Me.DViewer.Images(1))
            mintAutoNumber = 0  '�ָ���ͼ��ʱ��������
            
        Case conMenu_Process_Num * 100 To conMenu_Process_Num * 100 + 9
            mintNumberIndex = Val(Control.Category)
            subSetMouseState msNumber
        
        Case conMenu_Process_NumTag
            mintNumberIndex = 0
            subSetMouseState msNumber
            
        Case conMenu_Process_Word * 100 To conMenu_Process_Word * 100 + 99
            mstrText = Control.Caption
            subSetMouseState msFixText
        
        Case conMenu_Process_Save           '���ͼ����
            If mlngWinType = ptMark Then
                Call SaveImage(mtTagImage)
            End If
            
        Case conMenu_Process_SaveToStudy         '���浽���
            Call SaveImage(mtStudyImage)
            
            
        Case conMenu_Process_SaveToReport           '���浽����
            Call SaveImage(mtReportImage)
            
        Case conMenu_Process_Exit             '�˳�
            Unload Me
    End Select
    
'    If Control.ID <> conMenu_Process_LabelSetUp Then
'        Call setCmdLabelColor
'    End If
    If mlngWinType = ptPreview Then
        Me.Caption = "ͼ����"
        mlngWinType = ptProcess
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub



Private Sub subSetSharp(blnSharp As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ���ƽ������
'������blnSharp��ʾͼ����ķ���True=�񻯣�False=ƽ��
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        If blnSharp = True Then
            '�񻯴���
            If DViewer.Images(1).FilterLength <= 0 Then
                DViewer.Images(1).FilterLength = 0
                '��ǰû��ƽ������ֱ�ӽ����񻯴���
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement + 0.1
            Else
                '�����ǰ�Ѿ���ƽ���������ȵ���ƽ��Ч��
                DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength - 1
            End If
        Else
            'ƽ������
            '�ж�Zoom�Ƿ�1������ǣ����޸�Ϊ0.9999
            If DViewer.Images(1).ActualZoom = 1 Then
                DViewer.Images(1).Zoom = 0.9999
            End If
            
            If DViewer.Images(1).UnsharpEnhancement <= 0 Then
                DViewer.Images(1).UnsharpEnhancement = 0
                '��ǰû���񻯴���ֱ�ӿ�ʼƽ��
                '�ж�FilterLength�Ƿ�0����ǣ�����2/ActualZoom��2��FilterLength֮����е���
                If DViewer.Images(1).FilterLength = 0 Then
                    DViewer.Images(1).FilterLength = 2 / DViewer.Images(1).ActualZoom + 1
                Else    '���������FilterLength��1
                    DViewer.Images(1).FilterLength = DViewer.Images(1).FilterLength + 1
                End If
            Else
                '��ǰ�Ѿ������񻯴����ȵ����񻯵�Ч��
                DViewer.Images(1).UnsharpEnhancement = DViewer.Images(1).UnsharpEnhancement - 0.1
            End If
        End If
    End If
    
    mblnIsChanged = True
End Sub

Private Sub subSetRotate(blnClockwise As Boolean)
'------------------------------------------------
'���ܣ�dcmView��ͼ�����ת
'������blnClockwise��ת�ķ���,True=˳ʱ����ת��False=��ʱ����ת
'���أ��ޣ�ֱ�Ӵ���dcmView�е�ͼ��
'------------------------------------------------
    If DViewer.Images.Count > 0 Then
        Dim iRotateState As Integer
        
        iRotateState = DViewer.Images(1).RotateState
        If blnClockwise = True Then
            iRotateState = iRotateState - 1
        Else
            iRotateState = iRotateState + 1
        End If
        If iRotateState = -1 Then iRotateState = 3
        iRotateState = iRotateState Mod 4
        DViewer.Images(1).RotateState = iRotateState
    End If
End Sub


'DicomViewer�ü���ɼ�ͼ��
Private Sub CaptureFrameSelectImage()
    Dim imgResult As DicomImage
    Dim imgs As New DicomImages
    Dim iPlane As Integer
    Dim dblZoom As Double
    Dim iLeft As Integer
    Dim iRight As Integer
    Dim iTop As Integer
    Dim iBottom As Integer
    Dim iMax As Integer
    Dim img As DicomImage
    Dim lblFrame As DicomLabel
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    If Me.DViewer.Images(1).Labels.Count < 1 Then Exit Sub
    
    Set img = Me.DViewer.Images(1)
    Set lblFrame = Me.DViewer.Images(1).Labels(Me.DViewer.Images(1).Labels.Count)
    
    If Abs(lblFrame.Width) = 0 Or Abs(lblFrame.Height) = 0 Then
        MsgBoxD Me, "��ѡ��ͼ��������ٱ���", vbExclamation, "��ʾ"
        Exit Sub
    End If
    
    'ͼ�������=300
    iMax = 300
    
    '����label����ȡ����ѡ�е�ͼ��
    'ͼ��λ��,�ڰ�ͼ��Ϊ1����ɫͼ��Ϊ3
    iPlane = 1
    If Not IsNull(img.Attributes(&H28, &H4).value) And img.Attributes(&H28, &H4).Exists Then
        If img.Attributes(&H28, &H4).value = "RGB" Or img.Attributes(&H28, &H4).value = "YBR_FULL_422" Then
            iPlane = 3
        End If
    End If
    
    'ͼ����λ��
    If lblFrame.Width >= 0 Then
        iLeft = lblFrame.Left
        iRight = iLeft + lblFrame.Width
    Else
        iLeft = lblFrame.Left + lblFrame.Width
        iRight = lblFrame.Left
    End If
    
    If lblFrame.Height >= 0 Then
        iTop = lblFrame.Top
        iBottom = iTop + lblFrame.Height
    Else
        iTop = lblFrame.Top + lblFrame.Height
        iBottom = lblFrame.Top
    End If
    
    '���ƽ��ͼ��Ĵ�С��300*300֮��
    If (iRight - iLeft) > iMax Or (iBottom - iTop) > iMax Then
        dblZoom = iMax / (iRight - iLeft)
        If dblZoom > iMax / (iBottom - iTop) Then dblZoom = iMax / (iBottom - iTop)
    Else
        dblZoom = 1
    End If
    
    img.Labels(img.Labels.Count).Visible = False
    If (img.RotateState = doRotateLeft And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipHorizontal) Then
        'X����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, iTop, iBottom)
    ElseIf (img.RotateState = doRotateLeft And img.FlipState = doFlipBoth) _
        Or (img.RotateState = doRotateRight And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipVertical) Then
        'Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, img.SizeY - iBottom, img.SizeY - iTop)
    ElseIf (img.RotateState = doRotateRight And img.FlipState = doFlipHorizontal) _
        Or (img.RotateState = doRotateLeft And img.FlipState = doFlipVertical) _
        Or (img.RotateState = doRotate180 And img.FlipState = doFlipNormal) _
        Or (img.RotateState = doRotateNormal And img.FlipState = doFlipBoth) Then
        'X��Y����Ե�
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, img.SizeX - iRight, img.SizeX - iLeft, img.SizeY - iBottom, img.SizeY - iTop)
    Else
        Set imgResult = img.PrinterImage(8, iPlane, True, dblZoom, iLeft, iRight, iTop, iBottom)
    End If
    
    DViewer.Images.Clear
    DViewer.Images.Add imgResult
    
    mblnCase = True
    mblnIsChanged = True
End Sub

Private Sub subSetMouseState(intMoustState As TMouseState)
'------------------------------------------------
'���ܣ��������״̬��ͬʱ���¹�������ť��ѡ��״̬
'������intMoustState -- ���״̬
'���أ���
'------------------------------------------------
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_TextTag).Checked = False
    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_NumTag).Checked = False
'    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Line).Checked = False
     
    '�ı䵱ǰ���״̬
    If mintMouseState = intMoustState Then
        If intMoustState = msNumber Or intMoustState = msFixText Then
            If mintNumberIndex > 0 Or Len(mstrText) > 0 Then
                mintMouseState = intMoustState
                Exit Sub
            End If
        End If
        
        mintMouseState = msNone
    Else
        mintMouseState = intMoustState
        
        Select Case mintMouseState
            Case msWinLevel: cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = True
            Case msZoom: cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = True
            Case msRectangle: cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = True
            Case msArrow: cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = True
            Case msEllipse: cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = True
            Case msText
                If mstrText = "" Then
                    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_TextTag).Checked = True
                End If
            Case msNumber
                If mintNumberIndex = 0 Then
                    cbrMain.FindControl(xtpControlSplitButtonPopup, conMenu_Process_NumTag).Checked = True
                End If
            Case msMove: cbrMain.FindControl(xtpControlButton, conMenu_Process_MoveLabel).Checked = True
            Case msline: cbrMain.FindControl(xtpControlButton, conMenu_Process_Line).Checked = True
        End Select
    End If
    
End Sub

Private Sub cbrMain_Resize()
    Call refreshFace
     
End Sub

Public Sub refreshFace()
    '������ʾ�Ŀͻ�����
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
     
    If mlngWinType = ptMark Then
        Me.DViewer.Left = lngLeft
        Me.DViewer.Top = lngTop
        Me.DViewer.Width = lngRight
        Me.DViewer.Height = lngBottom - lngTop
    Else
        If mlngState <= 1 Then
            Me.DViewer.Left = lngLeft
            Me.DViewer.Top = lngTop
            Me.DViewer.Width = lngRight
            Me.DViewer.Height = lngBottom - lngTop
        Else
            Me.ucSplitter.Top = lngTop
            Me.ucSplitter.Height = lngBottom - lngTop - 600
            
            Me.ucBgImages.Left = lngLeft
            Me.ucBgImages.Top = lngTop
            Me.ucBgImages.Height = lngBottom - lngTop - 600
    
            '�ڷ�DViewer
            Me.DViewer.Left = lngLeft + Me.ucBgImages.Width + Me.ucSplitter.Width
            Me.DViewer.Top = lngTop
            Me.DViewer.Width = lngRight - lngLeft - Me.ucBgImages.Width - ucSplitter.Width
            Me.DViewer.Height = lngBottom - lngTop - 600
            
            ucSplitter.RePaint
            
            If lstMemoText.Visible Then lstMemoText.ZOrder
        End If
    End If
        
    Me.picMemo.Left = lngLeft
    Me.picMemo.Top = Me.DViewer.Top + Me.DViewer.Height
    Me.picMemo.Height = 600
    Me.picMemo.Width = lngRight
    
    Me.lstMemoText.Left = Me.cbxMemoText.Left
    Me.lstMemoText.Top = Me.picMemo.Top + Me.cbxMemoText.Top - Me.lstMemoText.Height
    Me.lstMemoText.Width = Me.cbxMemoText.Width - 10
    
    err.Clear
End Sub

Private Sub cmdInsert_Click()
    Dim strSQL As String, i As Integer
    Dim strUser As String
    
    Call EnterProcessState
    
    If Trim(cbxMemoText.Text) = "" Then
        MsgBoxD Me, "�����뱸ע���ݡ�", vbInformation, "��ʾ"
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    End If
    
    If cbxMemoText.ListIndex <> -1 Then
        MsgBoxD Me, "�ñ�ע�����Ѿ��ڳ��ñ�ע�С�", vbInformation, "��ʾ"
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    Else
        For i = 0 To cbxMemoText.ListCount - 1
            If UCase(Trim(cbxMemoText.list(i))) = UCase(Trim(cbxMemoText.Text)) Then
                MsgBoxD Me, "�ñ�ע���Ѿ��ڳ��ñ�ע�С�", vbInformation, "��ʾ"
                If cbxMemoText.Enabled Then cbxMemoText.SetFocus
                Exit Sub
            End If
        Next
    End If
        
    On Error GoTo errH
    
    strSQL = zlCommFun.zlGetSymbol(cbxMemoText.Text)
    strSQL = "zl_Ӱ��ͼ��ע_Insert('" & Replace(cbxMemoText.Text, "'", "''") & "','" & strSQL & "','" & mstrUser & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    AddComboItem cbxMemoText.hwnd, CB_ADDSTRING, 0, cbxMemoText.Text
    lstMemoText.AddItem cbxMemoText.Text
    
    MsgBoxD Me, "������Ϊ���ñ�ע��", vbInformation, "��ʾ"
    
    If cbxMemoText.Enabled Then cbxMemoText.SetFocus
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subAddMemoText()
'------------------------------------------------
'���ܣ���ͼ����ӱ�ע����
'������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    Dim img As DicomImage
    Dim iLeft As Integer
    Dim iWidth As Integer
    Dim iTop As Integer
    Dim iHeight As Integer
    Dim imgResult As New DicomImage
    Dim iPlane As Integer
    Dim lngWhiteX As Long
    Dim lngWhiteY As Long
    Dim lngFontHeight As Long
    
    If Me.DViewer.Images.Count <> 1 Then Exit Sub
    
    If Trim(cbxMemoText.Text) = "" Then Exit Sub
    
    lngFontHeight = ScaleY(TextHeight(cbxMemoText.Text), vbTwips, vbPixels) + 6
    
    '�ѱ�ע������ӵ�ͼ����
    Set img = Me.DViewer.Images(1)
    
    iLeft = 0
    iTop = 0
    iWidth = img.SizeX
    iHeight = img.SizeY + lngFontHeight

    'ʹ��PrinterImage���������Խ�ͼ���ϵı�ǩ����עͬʱ���л���
    Set imgResult = img.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight - lngFontHeight)
'

    '��ӱ�ע
    Dim dlMemoText As New DicomLabel
    
    dlMemoText.LabelType = doLabelText
    dlMemoText.ImageTied = True
    dlMemoText.Transparent = False
    dlMemoText.AutoSize = False
    dlMemoText.Left = 0
    dlMemoText.Top = img.SizeY
    dlMemoText.Width = iWidth
    dlMemoText.Height = lngFontHeight
    
    dlMemoText.BackColour = vbWhite
    dlMemoText.ForeColour = vbBlack
            
    dlMemoText.Font.Name = Me.Font.Name
    dlMemoText.Font.Italic = Me.Font.Italic
    dlMemoText.Font.Strikethrough = Me.Font.Strikethrough
    dlMemoText.Font.Underline = Me.Font.Underline
    dlMemoText.Font.Size = Me.Font.Size
    dlMemoText.Font.Bold = Me.Font.Bold
    dlMemoText.FontName = Me.Font.Name
    dlMemoText.FontSize = Me.Font.Size
    dlMemoText.ShowTextBox = True
    
    dlMemoText.Text = Me.cbxMemoText.Text & "                                                                                                                                 "
    
    imgResult.Labels.Add dlMemoText
    
    Set imgResult = imgResult.PrinterImage(8, iPlane, True, 1, 0, iWidth, 0, iHeight)

    '����ͼ��
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add imgResult
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub

Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim ls As DicomLabels
    Dim lngLeftD As Long
    
    If mlngWinType = ptPreview Then Exit Sub
    
    If Button = 1 And DViewer.Images.Count > 0 Then
        Dim intLabelType As Integer
        
        If mblnDrag Then
            Set ls = DViewer.LabelHits(X, Y, False, False, True)
            mlngBaseX = DViewer.ImageXPosition(X, Y)
            mlngBaseY = DViewer.ImageYPosition(X, Y)
            If ls.Count > 0 Then    '���ѡ�����κ�һ����ע
                '���Tag=""˵���Ǽ򵥱�ע���ǿ�˵�������ֱ�ű�ע����Ҫ�ҵ����ֱ�ע
                mintMouseState = msMove
                Set mMovingLabel = ls(1)
                If mMovingLabel.tag <> "" Then
                    If mMovingLabel.tag = m_LabelTag_Back Then
                        Set mMovingLabel = mMovingLabel.TagObject
                    ElseIf mMovingLabel.tag = m_LabelTag_Circle Then
                        Set mMovingLabel = mMovingLabel.TagObject.TagObject
                    End If
                End If
            Else
                mintMouseState = msDrag
            End If
        End If
                    
        mMouseDownPoint.X = DViewer.Images(1).ActualScrollX
        mMouseDownPoint.Y = DViewer.Images(1).ActualScrollY
          
        mInitScrollPoint.X = DViewer.Images(1).ScrollX + X
        mInitScrollPoint.Y = DViewer.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> msNone Then
            '��¼��ǰ���λ��
            mlngBaseXX = X
            mlngBaseYY = Y
            Select Case mintMouseState
                Case msline, msArrow, msEllipse, msText, msRectangle, msFixText, msNumber     'ֱ�ߣ���ͷ����Բ�����֣���ѡ���̶����֣�˳����
                    If mintMouseState = msArrow Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = msEllipse Or mintMouseState = msNumber Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = msText Or mintMouseState = msFixText Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = msRectangle Then
                        intLabelType = doLabelRectangle
                    ElseIf mintMouseState = msline Then
                        intLabelType = doLabelLine
                    End If
                    
                    If mintMouseState = msFixText Then
                        '����ǵ������֣�λ�Ƶ���Ҫ����
                        If mstrText = "�Զ���" Then
                            lngLeftD = IIf(Len(mstrCustom) = 1, 3, 7)
                        Else
                            lngLeftD = IIf(Len(Left(mstrText, InStr(mstrText & "=", "=") - 1)) = 1, 3, 7)
                        End If
                    Else
                        lngLeftD = 7
                    End If
                    DViewer.Images(1).Labels.Add GetNewLabel(intLabelType, DViewer.ImageXPosition(X, Y) - lngLeftD, DViewer.ImageYPosition(X, Y) - 7, 0, 0)
                    Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
                    If intLabelType = doLabelArrow Then
                        '��ͷ��Ҫʹ���߿�=2
                        mdcmSelectLabel.LineWidth = 4
                    ElseIf intLabelType = doLabelLine Then
                        mdcmSelectLabel.LineWidth = 2
                    ElseIf intLabelType = doLabelText Then
                        mdcmSelectLabel.XOR = False
                        mdcmSelectLabel.ForeColour = vbBlack
                        If mlngWinType <> ptMark Then
                            '���Ǳ��ͼ������������ӱ��������ͼ�������ӣ���Ϊ���Ӳ�����֧�֣���ӡ��ʱ��Ͳ�֧��
                            mdcmSelectLabel.Transparent = False
                            mdcmSelectLabel.ForeColour = vbWhite
                            mdcmSelectLabel.BackColour = vbBlack
                        End If
                        '���������С
                        If DViewer.Images(1).SizeX <= 256 Then
                            mdcmSelectLabel.FontSize = 10
                        ElseIf DViewer.Images(1).SizeX <= 512 Then
                            mdcmSelectLabel.FontSize = 15
                        Else
                            mdcmSelectLabel.FontSize = 18
                        End If
                        
                    End If
                    
                    mblnIsChanged = True
            End Select
        End If
    End If
End Sub

Private Sub DViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If mlngWinType = ptPreview Then Exit Sub
    
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        Select Case mintMouseState
            Case msWinLevel   '���ȶԱȶ�
                DViewer.Images(1).Width = DViewer.Images(1).Width + (X - mlngBaseXX)
                DViewer.Images(1).Level = DViewer.Images(1).Level + (Y - mlngBaseYY)
                mlngBaseXX = X
                mlngBaseYY = Y
                mblnIsChanged = True
            Case msZoom   '����
                Dim dblZoom As Double
                dblZoom = DViewer.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseYY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom DViewer.Images(1), DViewer, dblZoom, mCorpSize
                    mblnIsChanged = True
                End If
                mlngBaseYY = Y
'            Case msRectangle  '�ü�����
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseXX
'                dcmLabel.Top = mlngBaseYY
'                dcmLabel.Width = x - mlngBaseXX
'                dcmLabel.Height = y - mlngBaseYY
            Case msline, msArrow, msEllipse, msRectangle    'ֱ��,��ͷ��ע'Բ�α�ע,��ѡ
                mdcmSelectLabel.Width = DViewer.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = DViewer.ImageYPosition(X, Y) - mdcmSelectLabel.Top
                
                mblnIsChanged = True
            Case msDrag
                '�϶�ͼ��......
                DViewer.Images(1).ScrollX = mInitScrollPoint.X - X
                DViewer.Images(1).ScrollY = mInitScrollPoint.Y - Y
                
                mblnIsChanged = True
            Case msMove
                '�ƶ���ע
                If Not mMovingLabel Is Nothing Then
                    subaCorrectCursor DViewer, DViewer.Images(1), X, Y  '����ƶ��������ͼ��Χ�����������λ��
                    subMoveLable mMovingLabel, DViewer.ImageXPosition(X, Y) - mlngBaseX, DViewer.ImageYPosition(X, Y) - mlngBaseY
                    mlngBaseX = DViewer.ImageXPosition(X, Y)
                    mlngBaseY = DViewer.ImageYPosition(X, Y)
                    
                    mblnIsChanged = True
                End If
        End Select
        
        DViewer.Refresh
    End If
End Sub

Private Sub DViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mlngWinType = ptPreview Then Exit Sub

    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = msText Then      '���ֱ�ע
            
            txtInputText.Left = Me.ScaleX(X, vbPixels, vbTwips) + DViewer.Left
            txtInputText.Top = Me.ScaleY(Y, vbPixels, vbTwips) + DViewer.Top
            
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
            mblnIsChanged = True
        ElseIf mintMouseState = msRectangle Then   '�ü�����
            '��ʾͼ�񱣴�˵�
            Call ShowFrameSelectImagePopup
            
            'ɾ����ѡ�õ���ʱ��ע
            If DViewer.Images(1).Labels.Count > 0 Then
                DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            End If
            
            Set mdcmSelectLabel = Nothing
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseXX, mlngBaseYY, x - mlngBaseXX, y - mlngBaseYY
        ElseIf mintMouseState = msDrag Then
            '����ͼ�����ε�ƫ��λ��
            mCorpSize.X = mCorpSize.X + (DViewer.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (DViewer.Images(1).ActualScrollY - mMouseDownPoint.Y)
            
            mblnIsChanged = True
        ElseIf mintMouseState = msFixText Then
            '��ӹ̶�����
            If mstrText = "�Զ���" Then   '�Զ������ֱ�ע
                mdcmSelectLabel.Text = mstrCustom
            Else
                mdcmSelectLabel.Text = Left(mstrText, InStr(mstrText & "=", "=") - 1)
            End If
            mblnIsChanged = True
        ElseIf mintMouseState = msNumber Then
            Dim intText As Integer
            
            If mintNumberIndex = 0 Then '�Զ�˳����
                mintAutoNumber = mintAutoNumber + 1
                intText = mintAutoNumber
            Else
                intText = mintNumberIndex
            End If
            '���˳����
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.BackColour = glngColor(intText Mod 9 + 1)
            mdcmSelectLabel.Transparent = False
            mdcmSelectLabel.Width = 14
            mdcmSelectLabel.Height = 14
            mdcmSelectLabel.tag = m_LabelTag_Back
            
            '���˳����Բ�ε��������ӱ�ע��Բ�ο������
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelEllipse, mdcmSelectLabel.Left, mdcmSelectLabel.Top, 14, 14)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.tag = m_LabelTag_Circle
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            
            DViewer.Images(1).Labels.Add GetNewLabel(doLabelText, mdcmSelectLabel.Left + 1, mdcmSelectLabel.Top, 0, 0)
            Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
            mdcmSelectLabel.ForeColour = vbBlack
            mdcmSelectLabel.Transparent = True
            mdcmSelectLabel.XOR = False
            mdcmSelectLabel.tag = m_LabelTag_Number
            mdcmSelectLabel.FontSize = 8
            mdcmSelectLabel.FontName = "Arial Bold"
            mdcmSelectLabel.AutoSize = True
            mdcmSelectLabel.Text = intText
            If mdcmSelectLabel.Text < 10 Then
                mdcmSelectLabel.Left = mdcmSelectLabel.Left + 3
            End If
            mdcmSelectLabel.TagObject = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 1)
            DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count - 2).TagObject = mdcmSelectLabel    'TagObject�γɱջ�
            
            mblnIsChanged = True
        End If
        
        DViewer.Refresh
    End If
End Sub

Public Sub ShowFrameSelectImagePopup()
'------------------------------------------------
'���ܣ�������ѡͼ���ʱ�� ������Ҽ��ĵ����˵�
'������
'���أ���
'------------------------------------------------

Dim cbrControl As CommandBarControl
Dim cbrToolBar As CommandBar
Dim cbrToolPopup As CommandBarPopup
    
    
    '����Ҽ������˵�
    Set cbrToolBar = Me.cbrMain.Add("����Ҽ�", xtpBarPopup)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectCapture, "ȷ�ϲü�")
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub


Private Sub subCenterZoom(img As DicomImage, Viewer As DicomViewer, dblZoom As Double, corpSize As TPoint)
'------------------------------------------------
'���ܣ���ͼ��������š��Ե�ǰviewer���ĵ�Ϊ�������ĵ㡣
'������
'       img -- �������ŵ�ͼ��
'       viewer ���� ͼ�����ڵ�viewer
'       dblZoom ����ͼ���µ����ű���
'���أ��ޣ�ֱ�ӵ���ͼ������ű���
'�ϼ���������̣�frmViewer.Viewer_MouseMove
'�¼���������̣���
'���õ��ⲿ��������
'�����ˣ� �ƽ� 2006-2-10
'------------------------------------------------
    img.Zoom = dblZoom
    img.StretchToFit = False

            
    img.ScrollX = (img.SizeX * img.ActualZoom - ScaleX(Viewer.Width, vbTwips, vbPixels) / Viewer.MultiColumns) / 2 + corpSize.X
    img.ScrollY = (img.SizeY * img.ActualZoom - ScaleY(Viewer.Height, vbTwips, vbPixels) / Viewer.MultiRows) / 2 + corpSize.Y
End Sub


Private Sub Form_Load()
    On Error GoTo err

    '�ָ�����λ��
    Call RestorceWinLayout
    
    '����Ĭ����ɫ
    glngColor(1) = RGB(186, 186, 186)
    glngColor(2) = RGB(255, 215, 0)
    glngColor(3) = RGB(255, 0, 255)
    glngColor(4) = RGB(255, 0, 130)
    glngColor(5) = RGB(0, 255, 0)
    glngColor(6) = RGB(130, 255, 255)
    glngColor(7) = RGB(255, 255, 0)
    glngColor(8) = RGB(0, 0, 255)
    glngColor(9) = RGB(0, 160, 0)
    
    Call subLoadTextLabel
    
    '����������
    Call InitCommandBars
    
    Call LoadMemoFontStyle
    
    mCorpSize.X = 0
    mCorpSize.Y = 0
    mblnOk = False
    mintAutoNumber = 0
    
    'ͼ����������Ĭ���ǵ�����ͼ���ע��������Ĭ�����ƶ���ע
'    If mblnIsMark = True Then
'        Call subSetMouseState(msMove)
'    Else
'        Call subSetMouseState(msWinLevel)
'    End If
'
    ucBgImages.IsDrawOrder = False
    ucBgImages.IsDrawHint = False

    Call ReadEnjoin
    
    Call refreshFace
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

'���뱸ע������ʽ
Private Sub LoadMemoFontStyle()
    Dim strFontStyle As String
    Dim aryFontStyle() As String
    
    '������,12,B,U,S,I��
    
    strFontStyle = zlDatabase.GetPara("ͼ��ע����", glngSys, glngModul, "")
    
    strFontStyle = strFontStyle & ",,,,,,"
    
    aryFontStyle = Split(strFontStyle, ",")
    
    If aryFontStyle(0) <> "" Then Me.Font.Name = aryFontStyle(0)
    If Val(aryFontStyle(1)) <> 0 Then Me.Font.Size = Val(aryFontStyle(1))
    If UCase(aryFontStyle(2)) = "B" Then Me.Font.Bold = True
    If UCase(aryFontStyle(3)) = "U" Then Me.Font.Underline = True
    If UCase(aryFontStyle(4)) = "S" Then Me.Font.Strikethrough = True
    If UCase(aryFontStyle(5)) = "I" Then Me.Font.Italic = True
    
End Sub


Private Sub SaveMemoFontStyle()
    Dim strFontStyle As String
    
    strFontStyle = Me.Font.Name & "," & _
        Me.Font.Size & "," & _
        IIf(Me.Font.Bold, "B", "") & "," & _
        IIf(Me.Font.Underline, "U", "") & "," & _
        IIf(Me.Font.Strikethrough, "S", "") & "," & _
        IIf(Me.Font.Italic, "I", "")

    Call zlDatabase.SetPara("ͼ��ע����", strFontStyle, glngSys, glngModul)
End Sub


Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣�ñ�ע
    Dim strSQL As String, strPre As String
    Dim strUser As String
    
    On Error GoTo errH
    
    '��������
    strPre = cbxMemoText.Text '����󱣳�ԭ��ֵ
    cbxMemoText.Clear
    
    strSQL = _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա=[1]" & _
        " Union" & _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա is Null" & _
        " Order by ����"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrUser)
    Do While Not mrsTmp.EOF
        AddComboItem cbxMemoText.hwnd, CB_ADDSTRING, 0, mrsTmp!����
        
        lstMemoText.AddItem mrsTmp!����
        mrsTmp.MoveNext
    Loop
    cbxMemoText.Text = strPre
    
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Form_Unload(Cancel As Integer)
    If mblnIsChanged Then
        If MsgBoxD(Me, "ͼ������δ���棬�Ƿ񱣴棿", vbYesNo, "��ʾ") = vbYes Then
            If mlngWinType = ptMark Then
                Call SaveImage(mtTagImage)
            Else
                Call SaveImage(mtStudyImage)
            End If
        End If
    End If
    
    mlngAdviceId = 0
    
    Call SaveWinLayout
    
    Call SaveMemoFontStyle
    
    RaiseEvent OnUnload
End Sub

Private Sub SaveWinLayout()
'���洰��λ�ü����沼��
'����Ĭ�ϴ��ڴ�Сԭ��δʹ��ZL9COMLIB�еķ���
    Dim strCaption As String
    Dim strPrivateReg As String
    
    strCaption = GetWindowCaption
    
    strPrivateReg = GetPrivateRegPath(strCaption)
    
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinLeft", IIf(Me.Left < 0, 0, Me.Left))
    Call SaveSetting("ZLSOFT", strPrivateReg, "Wintop", IIf(Me.Top < 0, 0, Me.Top))
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinWidth", Me.Width)
    Call SaveSetting("ZLSOFT", strPrivateReg, "WinHeight", Me.Height)
    Call SaveSetting("ZLSOFT", strPrivateReg, "MiniatureW", ucBgImages.Width)
    Call SaveSetting("ZLSOFT", strPrivateReg, "����ͼ����", ucBgImages.PageRecordCount)
End Sub

Private Function GetWindowCaption() As String
    Dim lngCurWinType As TImgProcessType
    
    lngCurWinType = mlngWinType
    
    If lngCurWinType = ptMark Then
        GetWindowCaption = "���ͼ"
    Else
        GetWindowCaption = IIf(lngCurWinType = ptPreview, "ͼ��Ԥ��", "ͼ����")
    End If
End Function

Private Sub RestorceWinLayout()
    Dim strCaption As String
    Dim strPrivateReg As String
     
    strCaption = GetWindowCaption()
    
    strPrivateReg = GetPrivateRegPath(strCaption)
    
    Me.Left = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinLeft", Screen.Width / 4))
    Me.Top = nvl(GetSetting("ZLSOFT", strPrivateReg, "Wintop", Screen.Height / 4))
    Me.Width = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinWidth", Screen.Width / 2))
    Me.Height = nvl(GetSetting("ZLSOFT", strPrivateReg, "WinHeight", Screen.Height / 2))

    ucBgImages.Width = nvl(GetSetting("ZLSOFT", strPrivateReg, "MiniatureW", 3000))
    
    ucBgImages.PageRecordCount = Val(GetSetting("ZLSOFT", strPrivateReg, "����ͼ����", 8))
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 16, 16
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    'ͼ���������������
    Set cbrToolBar = Me.cbrMain.Add("ͼ�������", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '�ı���ʾ��ͼ���·�
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        If mlngWinType = ptMark Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_Save, "����"): cbrControl.ToolTipText = "����"
        Else
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToStudy, "��Ϊ���ͼ"): cbrControl.ToolTipText = "���浽���ͼ��"
            If mblnIsReportShow Then
                Set cbrControl = .Add(xtpControlButton, conMenu_Process_SaveToReport, "��Ϊ����ͼ"): cbrControl.ToolTipText = "���浽����ͼ��"
            End If
        End If
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "����"): cbrControl.ToolTipText = "��������/�Աȶ�": cbrControl.Visible = mlngWinType <> ptMark
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "����"): cbrControl.ToolTipText = "����ͼ��": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "�ü�"): cbrControl.ToolTipText = "�ü��ɼ�ͼ��": cbrControl.iconid = 3201: cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "��"): cbrControl.ToolTipText = "��": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "ƽ��"): cbrControl.ToolTipText = "ƽ��": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Line, "ֱ��"): cbrControl.ToolTipText = "ֱ�߱�ע": cbrControl.Visible = mlngWinType <> ptMark
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "��ͷ"): cbrControl.ToolTipText = "��ͷ��ע": cbrControl.Visible = mlngWinType <> ptMark
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "Բ��"): cbrControl.ToolTipText = "Բ�α�ע"
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_NumTag, "����"): cbrControl.ToolTipText = "���ֱ�ע"
        Call LoadComNumber(cbrControl)
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_TextTag, "�ı�"): cbrControl.ToolTipText = "�����ı���ע"
        Call LoadComText(cbrControl)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_MoveLabel, "�ƶ�"): cbrControl.ToolTipText = "ѡ�б�עʱ����������ק�ƶ���ע�������϶�ͼƬ��˫��ɾ����ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LabelSetUp, "���ñ�ע"): cbrControl.ToolTipText = "�������ֱ�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_DelAllLabels, "���"): cbrControl.ToolTipText = "���ȫ����ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Restore, "�ָ�"): cbrControl.ToolTipText = "�ָ�ͼ�񵽳�ʼ״̬"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Exit, "�˳�"): cbrControl.ToolTipText = "�˳�"
    End With
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
         cbrControl.Category = "Main" '���ó�������˵�
    Next
    cbrToolBar.Position = xtpBarTop
End Sub

Private Sub LoadComNumber(mnuParent As Object)
    Dim objControl As CommandBarControl
    Dim i As Long
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100, "��"): objControl.ToolTipText = "�Զ��������ֱ��": objControl.Category = 0: objControl.iconid = 0
    
    For i = 1 To 9
        Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Num * 100 + i, i): objControl.ToolTipText = "���ֱ��" & i: objControl.Category = i: objControl.iconid = 0
    Next
    
End Sub

Private Sub LoadComText(mnuParent As Object)
    Dim objControl As CommandBarControl
    Dim arrTemp() As String
    Dim i As Long
    
    arrTemp = Split(mStrTemp, "|")
    
    For i = 0 To UBound(arrTemp)
        If Len(arrTemp(i)) > 0 Then
            Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_Word * 100 + i + 1, arrTemp(i)): objControl.ToolTipText = "�ı���ע": objControl.Category = i + 1: objControl.iconid = 0
        End If
    Next
End Sub


Private Sub lstMemoText_DblClick()
    cbxMemoText.Text = lstMemoText.list(lstMemoText.ListIndex)
    lstMemoText.Visible = False
    
    cbxMemoText.SelStart = 0
    cbxMemoText.SelLength = Len(cbxMemoText.Text)
    cbxMemoText.SetFocus
End Sub

Private Sub lstMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
End Sub

Private Sub lstMemoText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
    End If
    
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then lstMemoText.Visible = False
End Sub

Private Sub picCboDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 1
End Sub

Private Sub picCboDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCboDropDown.BorderStyle = 0
End Sub
 

Private Sub picMemo_Resize()
    On Error Resume Next
    
    '�ڷű�ע����
    Me.lblMemoText.Left = 100
    Me.lblMemoText.Top = 200

    Me.cbxMemoText.Left = Me.lblMemoText.Left + Me.lblMemoText.Width
    Me.cbxMemoText.Top = Me.lblMemoText.Top - 100
    Me.cbxMemoText.Width = Me.ScaleWidth - Me.cbxMemoText.Left - 250 - cmdInsert.Width - cmdFont.Width - cmdAdd.Width
    
    Me.picCboDropDown.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width - 270
    Me.picCboDropDown.Top = Me.cbxMemoText.Top + 30
    
    Me.cmdAdd.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width
    Me.cmdAdd.Top = Me.cbxMemoText.Top
    
    Me.cmdInsert.Left = Me.cmdAdd.Left + Me.cmdAdd.Width
    Me.cmdInsert.Top = Me.cbxMemoText.Top

    Me.cmdFont.Left = Me.cmdInsert.Left + Me.cmdInsert.Width
    Me.cmdFont.Top = Me.cmdInsert.Top
    
    err.Clear
End Sub

Private Sub Timer1_Timer()
    Dim ptWin As POINTAPI
On Error GoTo errHandle
    If mlngState = 3 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        
        Exit Sub
    End If
    
    If mlngAdviceId = 0 Then
        Timer1.Enabled = False
        Exit Sub
    End If
    
    GetCursorPos ptWin

    If mlngState = 1 Then
        '����ƽ�����
        If ptWin.X >= Me.Left / 15 And ptWin.X <= (Me.Left + Me.Width) / 15 And ptWin.Y >= Me.Top / 15 And ptWin.Y <= (Me.Top + Me.Height) / 15 Then
 
            mlngState = 2
            
            If DViewer.Images.Count > 0 Then
                Call ClearHint(DViewer.Images(1))
            End If
            
            Call RefrshObjVisible
            Call refreshFace
            

            
            If ucBgImages.ImgCount <= 0 And ThumbnailImgCount > 0 Then
            '�����Ԥ������������һ���ƶ�������ʱ����ͼ��
                Call LoadImgs(maryImgInfos)
            End If
              
            Timer2.Enabled = False
        End If
    ElseIf mlngState = 2 Then
        '����Ƴ�����

        If ptWin.X < Me.Left / 15 Or ptWin.X > (Me.Left + Me.Width) / 15 Or ptWin.Y < Me.Top / 15 Or ptWin.Y > (Me.Top + Me.Height) / 15 Then
            mlngState = 1
            
            If DViewer.Images.Count > 0 Then
                Call DrawHintTag(DViewer.Images(1))
            End If
            
            Call RefrshObjVisible
            Call refreshFace
   
            If mlngPreViewTime > 0 And mlngWinType = ptPreview Then
                Timer2.Enabled = True
            End If
            
            
        End If
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
    End If
Exit Sub
errHandle:
    Debug.Print "Timer1 Bug:" & err.Description
End Sub



Private Sub Timer2_Timer()
    If mlngWinType = ptPreview And mlngPreViewTime > 0 Then
        Call UnloadMe
    End If
End Sub

Private Sub txtInputText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then  '''ESC�ͻس����˳�����
        txtInputText.Visible = False
        If Trim(txtInputText.Text) = "" Then
            'ɾ�����ֱ�ע
            DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            txtInputText = "1 "
        Else
            mdcmSelectLabel.Text = txtInputText.Text
            DViewer.Refresh
        End If
    End If
End Sub

Private Sub ShowPopupImage()
'------------------------------------------------
'���ܣ���������Ҽ������˵�
'intType:0--����ͼ��1--����ͼ��2--����ͼ
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
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Page, "��ҳ����")
            
    End With
    cbrToolBar.Visible = True
    cbrToolBar.ShowPopup
End Sub

Private Sub subaCorrectCursor(v As DicomViewer, im As DicomImage, xx As Long, Yy As Long)
'------------------------------------------------
'���ܣ�����ƶ��������ͼ��Χ�����������λ��
'������v--ͼ�����ڵ�viewer��im--������ڵ�ͼ��xx--������ڵ�x����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'      yy--������ڵ�y����λ�ã������곬��ͼ���򽫴�ֵ�޸ĵ�ͼ��֮�ڣ�
'���أ���
'------------------------------------------------
    Dim X As Integer, Y As Integer, w As Long, h As Long
    Dim i As DicomImage
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    w = v.Width / v.MultiColumns / Screen.TwipsPerPixelX - v.CellSpacing * 2
    h = v.Height / v.MultiRows / Screen.TwipsPerPixelY - v.CellSpacing * 2
    X = im.OriginX + v.CellSpacing
    Y = im.OriginY + v.CellSpacing
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If xx < X Then xx = X
    If xx > X + w Then xx = X + w
    If Yy < Y Then Yy = Y
    If Yy > Y + h Then Yy = Y + h
End Sub

Public Sub subMoveLable(la As DicomLabel, X As Long, Y As Long)
'------------------------------------------------
'���ܣ��ƶ�һ����ע
'������la--���ƶ��ı�ע��x--x�����ƶ���ͼ�����ؾ��룻y--y�����ƶ���ͼ�����ؾ���
'���أ���
'------------------------------------------------
    
    la.Left = la.Left + X
    la.Top = la.Top + Y
    
    '��������ֱ�ţ���Ҫͬʱ�ƶ�������ע
    If la.tag <> "" And Not la.TagObject Is Nothing Then
        la.TagObject.Left = la.TagObject.Left + X
        la.TagObject.Top = la.TagObject.Top + Y
        la.TagObject.TagObject.Left = la.TagObject.TagObject.Left + X
        la.TagObject.TagObject.Top = la.TagObject.TagObject.Top + Y
    End If
       
End Sub

Private Sub subSetTextLabel()
'------------------------------------------------
'���ܣ��������ֱ�ע��������
'������
'���أ���
'------------------------------------------------
    Dim strTemp As String
    Dim i As Integer

    On Error GoTo err
    
'    strTemp = InputBox("�������µ����ֱ�ע���ã���ʽΪ������1=˵��1|����2=˵��2|...����", "���ֱ�ע����", Replace(mstrTemp, "[+]", "|"))
    
    strTemp = frmInputBoxV2.ZlShowMe(Me, mStrTemp)
    
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        MsgBoxD Me, "����ĸ�ʽ����ȷ��Ӧ�ð��ա�����=˵������ʽ���룬������������á�", vbOKOnly, "��ʾ"
        Exit Sub
    End If
     
    '����ɹ���ʹ������µ����ֱ�ע��ͬʱ���浽ע�����
    mStrTemp = strTemp

    cbrMain.FindControl(, conMenu_Process_TextTag).CommandBar.Controls.DeleteAll
    
    LoadComText cbrMain.FindControl(, conMenu_Process_TextTag)
    Call SaveSetting("ZLSOFT", "����ģ��\zl9PACSWork\frmReportImageEdit", "�������ֱ�ע", Replace(mStrTemp, "|", "[+]"))
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subLoadTextLabel()
'------------------------------------------------
'���ܣ���ȡ���ֱ�ע
'������
'���أ���
'------------------------------------------------
    Dim strTemp As String
    Dim strtext() As String
    Dim i  As Integer
    
    On Error GoTo err
    
    mStrTemp = GetSetting("ZLSOFT", "����ģ��\zl9PACSWork\frmReportImageEdit", "�������ֱ�ע", G_STR_TAG)
    
    mStrTemp = Replace(mStrTemp, "[+]", "|")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub



Public Sub UnloadMe()
    Unload Me
End Sub

 

Private Sub ucBgImages_OnClick(ByVal lngSelIndex As Long)
On Error GoTo err
     
    Call EnterProcessState
    
    If mlngWinType = ptMark Then Exit Sub
    
 
    If DViewer.Images.Count > 0 Then
        If mblnIsChanged Then
            If MsgBoxD(Me, "ͼ�������δ���棬�Ƿ������", vbYesNo, "��ʾ") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    DViewer.Images.Clear
    
    Set mOldImage = ucBgImages.GetImage(lngSelIndex)
    
    If mOldImage Is Nothing Then Exit Sub
    DViewer.Images.Add ucBgImages.GetImage(lngSelIndex)
     
    mblnIsChanged = False
    mblnCase = False
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub AutoUnload()
    If mblnIsUnloud Then
        Timer2.Enabled = True
    End If
End Sub

Private Sub DrawHintTag(dcmImg As DicomImage)
    Dim lRpt As DicomLabel
    Dim i As Integer
    
    If mlngAdviceId = 0 Then Exit Sub
     
    Set lRpt = New DicomLabel
            
    With lRpt
        .LabelType = doLabelText
        .Width = 800
        .Height = 60
        .ImageTied = False
        .Transparent = True
        .ScaleWithCell = True
        .ScaleFontSize = 40
        .Font.Name = "����"
        .Font.Size = 22
        .Font.Bold = True
        .ForeColour = &HCBBECB
        .Left = 120
        .Top = 20
        .Text = "...�����������..."
        .Shadow = doShadowBottomRight
        .Alignment = doAlignCentre
        .Visible = True
        .tag = "HINT"
    End With
    
    dcmImg.Labels.Add lRpt
    
    dcmImg.Refresh False
End Sub

Private Sub ClearHint(dcmImage As DicomImage)
    Dim i As Long
    
    For i = 1 To dcmImage.Labels.Count
        If dcmImage.Labels(i).tag = "HINT" Then
            dcmImage.Labels.Remove i
            Exit For
        End If
    Next
    
    dcmImage.Refresh False
End Sub
 

Private Sub ucSplitter_OnMoveEnd()
    If lstMemoText.Visible Then lstMemoText.ZOrder
End Sub

