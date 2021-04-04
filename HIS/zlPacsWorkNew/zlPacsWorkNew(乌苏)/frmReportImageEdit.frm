VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmReportImageEdit 
   Caption         =   "����ͼƬ�༭"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12135
   Icon            =   "frmReportImageEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picCboDropDown 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6320
      Picture         =   "frmReportImageEdit.frx":0E42
      ScaleHeight     =   375
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   4350
      Width           =   255
   End
   Begin VB.ListBox lstMemoText 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   8400
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog diaFont 
      Left            =   4800
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFont 
      Height          =   375
      Left            =   6960
      Picture         =   "frmReportImageEdit.frx":119E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "���õ�ǰ��ע���塣"
      Top             =   4320
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��"
      Height          =   400
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdCur 
      Caption         =   "��һ��"
      Height          =   400
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
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
      Left            =   2040
      TabIndex        =   6
      Top             =   4320
      Width           =   4575
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   375
      Left            =   6600
      Picture         =   "frmReportImageEdit.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "����ǰ��ע����Ϊ���ñ�ע"
      Top             =   4320
      Width           =   375
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
      Left            =   9480
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   400
      Left            =   5040
      TabIndex        =   2
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���"
      Height          =   400
      Left            =   3840
      TabIndex        =   1
      Top             =   5040
      Width           =   1100
   End
   Begin DicomObjects.DicomViewer DViewer 
      Height          =   3495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _Version        =   262147
      _ExtentX        =   12515
      _ExtentY        =   6165
      _StockProps     =   35
      BackColor       =   -2147483638
      UseScrollBars   =   0   'False
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
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1470
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportImageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TPoint
  X As Integer
  Y As Integer
End Type

Private mlngModule As Long
Private mImage As DicomImage
Private mintMouseState As Integer
Private mblnDcmViewDown As Boolean
Private mMouseDownPoint As TPoint
Private mInitScrollPoint As TPoint
Private mCorpSize As TPoint             '�϶�������ƫ��λ��
Private mlngBaseX As Long
Private mlngBaseY As Long
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mblnOK As Boolean
Private mOldImage As DicomImage
Private mintCurImgIndex As Integer      '������ѡ������ͼ������
Private mfrmParent As Object            '������ģ�����
Private mSelViewerIndex As Integer      '�����屻ѡ�еı���ͼ���ID����1��ʼ����

Private mrsTmp As ADODB.Recordset       'ͼ��ע��¼��

Public Sub zlShowMe(ByVal img As DicomImage, frmParent As frmReportImage, _
    intCurImgIndex As Integer, SelViewerIndex As Integer, ByVal lngModule As Long)
    
    Set mOldImage = img

    mlngModule = lngModule
    mintCurImgIndex = intCurImgIndex
    mSelViewerIndex = SelViewerIndex
    Set mfrmParent = frmParent
    
    Me.DViewer.Images.Clear
    Me.DViewer.Images.Add img
    Me.Show 1, frmParent

End Sub

Private Sub ChangeImage(intType As Integer)
'intType �л����� 1 --��һ��ͼ��2--��һ��ͼ
    Dim i As Integer

    If mfrmParent.ImageCount <= 1 Then
        Exit Sub
    End If
    
    
    Me.DViewer.Images.Clear
    If intType = 1 Then  '��һ��ͼ
        If mintCurImgIndex <= 1 Then
            Call mfrmParent.MovePage(mtLast)
            mintCurImgIndex = mfrmParent.ImageCount
        Else
            mintCurImgIndex = mintCurImgIndex - 1
        End If
        
        Me.DViewer.Images.Add mfrmParent.dcmImages(mintCurImgIndex)
    ElseIf intType = 2 Then   '��һ��ͼ
        If mintCurImgIndex >= mfrmParent.ImageCount Then
            Call mfrmParent.MovePage(mtNext)
            mintCurImgIndex = 1
        Else
            mintCurImgIndex = mintCurImgIndex + 1
        End If
        
        
        Me.DViewer.Images.Add mfrmParent.dcmImages(mintCurImgIndex)
    End If
    
    '���ѡ��ͼ�εı߿���ɫ
    Me.DViewer.Images(1).BorderColour = vbRed
    
'    '�ж��Ƿ��ǵ�һ��ͼ ���� ���һ��ͼ,������ذ�ť
'    Me.cmdCur.Enabled = IIf(mintCurImgIndex = 1, False, True)
'    Me.cmdNext.Enabled = IIf(mintCurImgIndex = mfrmParent.ImageCount, False, True)
    
    '�Ը���������ͼ�ı߿���ɫ���д���
    For i = 1 To mfrmParent.ImageCount
        mfrmParent.dcmImages(i).BorderColour = vbWhite
    Next i
    
    Set mfrmParent.mSelMiniImg = mfrmParent.dcmImages(mintCurImgIndex)
    mfrmParent.mSelMiniImg.BorderColour = vbRed
    
    '���ComboBox�ı�
    zlControl.CboSetIndex cbxMemoText.hWnd, -1
    
    '�ر�������
    If lstMemoText.Visible Then lstMemoText.Visible = False
End Sub

Private Function getListIndex() As Integer
'���ݼ���������ȡ����
    Dim i As Integer

    getListIndex = -1
    
    If mrsTmp.RecordCount <= 0 Then Exit Function

    mrsTmp.MoveFirst
    
    If cbxMemoText.Text = "" Then Exit Function

    For i = 0 To mrsTmp.RecordCount - 1
        If InStr(Trim(Nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Or InStr(Trim(Nvl(mrsTmp!����)), UCase(cbxMemoText.Text)) > 0 Then
            getListIndex = i
            
            Exit For
        End If

        mrsTmp.MoveNext
    Next
End Function

Private Sub cbxMemoText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex
    End If
End Sub

Private Sub lstMemoText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlControl.CboSetText(cbxMemoText, lstMemoText.list(lstMemoText.ListIndex))
End Sub

Private Sub picCboDropDown_Click()
    lstMemoText.Visible = Not lstMemoText.Visible
    If lstMemoText.ListCount > 0 Then lstMemoText.ListIndex = getListIndex

    If lstMemoText.Visible Then lstMemoText.SetFocus
End Sub

Private Sub cbxMemoText_Change()
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

Private Sub cmdCur_Click()
'��һ��ͼ��
On Error GoTo errH
 
    Call ChangeImage(1)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdNext_Click()
'��һ��ͼ��
On Error GoTo errH
 
    Call ChangeImage(2)
 
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
'��Ӳ��������رմ���
    mblnOK = True
    'ƴ�ӷ���
    Call subAddMemoText
    
    If mblnOK Then
        If Me.DViewer.Images.Count = 1 Then
            Set mImage = Me.DViewer.Images(1)
        Else
            Set mImage = Nothing
        End If
    Else
        Set mImage = Nothing
    End If
    
    '��ƴ�Ӻ��ͼ��ı߿���д���
     If Me.DViewer.Images.Count > 0 Then
         With Me.DViewer.Images(1)
            .BorderWidth = 3
            .BorderStyle = 2
            .BorderColour = vbRed
        End With
    End If
    
    Call mfrmParent.DcmAddImage(mImage, mSelViewerIndex)
    Me.DViewer.Refresh
    
    '���ComboBox�ı�
    cbxMemoText.Text = ""
    
    '�ر�������
    lstMemoText.Visible = False
End Sub

Private Sub cmdExit_Click()
'���Viewer�ؼ�����ж�ش���
   ' Me.DViewer.Images.Clear
    Unload Me
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
  On Error GoTo errHandle
    Select Case control.ID
        Case conMenu_Process_Window         '���ȶԱȶ�
            subSetMouseState 1
            'Control.Checked = True
            
        Case conMenu_Process_Zoom           '����
            subSetMouseState 2
            'Control.Checked = True
            
        Case conMenu_Process_RectZoom       '�ü�����
            subSetMouseState 3
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
            
        Case conMenu_Process_Corp          '�϶�
           subSetMouseState 14
           'Control.Checked = True
            
        Case conMenu_Process_Arrow          '��ͷ��ע
            subSetMouseState 11
            'Control.Checked = True
            
        Case conMenu_Process_Ellipse        'Բ�α�ע
            subSetMouseState 12
            'Control.Checked = True
            
        Case conMenu_Process_Text           '���ֱ�ע
            subSetMouseState 13
            'Control.Checked = True
        Case conMenu_Process_Restore        '�ָ�
            DViewer.Images.Clear
            DViewer.Images.Add mOldImage
        
    End Select
    
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
        MsgBoxD Me, "��ѡ��ͼ��������ٱ���", vbExclamation, gstrSysName
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
    
End Sub


Private Sub subSetMouseState(intMouseState As Integer)
    
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = False
    cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = False
        
    '�ı䵱ǰ���״̬
    If mintMouseState = intMouseState Then
        mintMouseState = 0
        
    Else
        mintMouseState = intMouseState
        
        Select Case mintMouseState
            Case 1: cbrMain.FindControl(xtpControlButton, conMenu_Process_Window).Checked = True
            Case 2: cbrMain.FindControl(xtpControlButton, conMenu_Process_Zoom).Checked = True
            Case 3: cbrMain.FindControl(xtpControlButton, conMenu_Process_RectZoom).Checked = True
            Case 11: cbrMain.FindControl(xtpControlButton, conMenu_Process_Arrow).Checked = True
            Case 12: cbrMain.FindControl(xtpControlButton, conMenu_Process_Ellipse).Checked = True
            Case 13: cbrMain.FindControl(xtpControlButton, conMenu_Process_Text).Checked = True
            Case 14: cbrMain.FindControl(xtpControlButton, conMenu_Process_Corp).Checked = True
        End Select
    End If
    
End Sub


Private Sub cbrMain_Resize()
    '������ʾ�Ŀͻ�����
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    On Error Resume Next
    
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    '�ڷ�DViewer
    Me.DViewer.Left = lngLeft
    Me.DViewer.Top = lngTop
    Me.DViewer.Width = Abs(lngRight - lngLeft)
    Me.DViewer.Height = Abs(lngBottom - lngTop - 1300)
    
    '�ڷű�ע����
    Me.lblMemoText.Left = 100
    Me.lblMemoText.Top = Me.ScaleHeight - 1100
    
    Me.cbxMemoText.Left = Me.lblMemoText.Left + Me.lblMemoText.Width
    Me.cbxMemoText.Top = Me.lblMemoText.Top - 100
    Me.cbxMemoText.Width = Abs(Me.ScaleWidth - Me.cbxMemoText.Left - 250 - cmdInsert.Width - cmdFont.Width)
    
    Me.lstMemoText.Left = Me.cbxMemoText.Left
    Me.lstMemoText.Top = Me.cbxMemoText.Top - Me.lstMemoText.Height
    Me.lstMemoText.Width = Me.cbxMemoText.Width - 10
    
    Me.picCboDropDown.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width - 260
    Me.picCboDropDown.Top = Me.cbxMemoText.Top + 30
    
    Me.cmdInsert.Left = Me.cbxMemoText.Left + Me.cbxMemoText.Width
    Me.cmdInsert.Top = Me.cbxMemoText.Top
    
    Me.cmdFont.Left = Me.cmdInsert.Left + Me.cmdInsert.Width
    Me.cmdFont.Top = Me.cmdInsert.Top
    
    '�ڷš���ӡ������˳�����ť
    Me.cmdAdd.Left = Me.ScaleWidth - Me.cmdAdd.Width * 3
    Me.cmdAdd.Top = Me.ScaleHeight - 600
    
    Me.cmdExit.Left = Me.ScaleWidth - Me.cmdExit.Width * 1.8
    Me.cmdExit.Top = Me.cmdAdd.Top
    
    '�ڷš���һ����������һ������ť
    Me.cmdCur.Left = Me.ScaleWidth / 15
    Me.cmdCur.Top = Me.ScaleHeight - 600

    Me.cmdNext.Left = Me.cmdCur.Width + Me.cmdCur.Left + 200
    Me.cmdNext.Top = Me.cmdAdd.Top
End Sub



Private Sub cmdInsert_Click()
    Dim strSql As String, i As Integer
    
    If Trim(cbxMemoText.Text) = "" Then
        MsgBoxD Me, "�����뱸ע���ݡ�", vbInformation, gstrSysName
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    End If
    If cbxMemoText.ListIndex <> -1 Then
        MsgBoxD Me, "�ñ�ע�����Ѿ��ڳ��ñ�ע�С�", vbInformation, gstrSysName
        If cbxMemoText.Enabled Then cbxMemoText.SetFocus
        Exit Sub
    Else
        For i = 0 To cbxMemoText.ListCount - 1
            If UCase(Trim(cbxMemoText.list(i))) = UCase(Trim(cbxMemoText.Text)) Then
                MsgBoxD Me, "�ñ�ע���Ѿ��ڳ��ñ�ע�С�", vbInformation, gstrSysName
                If cbxMemoText.Enabled Then cbxMemoText.SetFocus
                Exit Sub
            End If
        Next
    End If
        
    On Error GoTo errH
    
    strSql = zlCommFun.zlGetSymbol(cbxMemoText.Text)
    strSql = "zl_Ӱ��ͼ��ע_Insert('" & Replace(cbxMemoText.Text, "'", "''") & "','" & strSql & "','" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, cbxMemoText.Text
    MsgBoxD Me, "������Ϊ���ñ�ע��", vbInformation, gstrSysName
    If cbxMemoText.Enabled Then cbxMemoText.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub subAddMemoText()
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
    
    If Trim(cbxMemoText.Text) <> "" Then
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
    End If
End Sub





Private Sub DViewer_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
   If Button = 1 And DViewer.Images.Count > 0 Then
        Dim intLabelType As Integer
        
        mMouseDownPoint.X = DViewer.Images(1).ActualScrollX
        mMouseDownPoint.Y = DViewer.Images(1).ActualScrollY
          
        mInitScrollPoint.X = DViewer.Images(1).ScrollX + X
        mInitScrollPoint.Y = DViewer.Images(1).ScrollY + Y
        
        mblnDcmViewDown = True
        If mintMouseState <> 0 Then
            '��¼��ǰ���λ��
            mlngBaseX = X
            mlngBaseY = Y
            Select Case mintMouseState
                'Case 14  'ͼ���϶�
                
                Case 11, 12, 13, 3    '��ͷ����Բ������,��ѡ
                    If mintMouseState = 11 Then
                        intLabelType = doLabelArrow
                    ElseIf mintMouseState = 12 Then
                        intLabelType = doLabelEllipse
                    ElseIf mintMouseState = 13 Then
                        intLabelType = doLabelText
                    ElseIf mintMouseState = 3 Then
                        intLabelType = doLabelRectangle
                    End If
                    
                    DViewer.Images(1).Labels.Add GetNewLabel(intLabelType, DViewer.ImageXPosition(X, Y), DViewer.ImageYPosition(X, Y), 0, 0)
                    
                    Set mdcmSelectLabel = DViewer.Images(1).Labels(DViewer.Images(1).Labels.Count)
                    
                    mdcmSelectLabel.LineWidth = 2
            End Select
        End If
    End If
End Sub


Public Function GetNewLabel(lType As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer) As DicomLabel
'------------------------------------------------
'���ܣ�����һ��LABEL���󣬲���������ʼ����
'������lType--��ע�����ͣ�lLeft--��ע��Leftֵ��lTop--��ע��Topֵ��lWidth--��ע��Widthֵ��lHeight--��ע��Heightֵ��
'���أ������ɵı�ע��
'�����ˣ��ƽ�
'------------------------------------------------
    Dim l As New DicomLabel
    l.LabelType = lType
    l.XOR = True
    l.ImageTied = True
    l.Left = lLeft
    l.Top = lTop
    l.Width = lWidth
    l.Height = lHeight
    l.Margin = 0
    l.AutoSize = True
    l.FontSize = 15
    l.LineWidth = 1
    
    If l.LabelType = 0 Then     '����
        l.Transparent = True
        l.Shadow = doShadowBottomRight

        l.Width = 200
        l.Height = 15
    End If
    
    Set GetNewLabel = l
End Function



Private Sub DViewer_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        Select Case mintMouseState
            Case 1  '���ȶԱȶ�
                DViewer.Images(1).Width = DViewer.Images(1).Width + (X - mlngBaseX)
                DViewer.Images(1).Level = DViewer.Images(1).Level + (Y - mlngBaseY)
                mlngBaseX = X
                mlngBaseY = Y
            Case 2  '����
                Dim dblZoom As Double
                dblZoom = DViewer.Images(1).ActualZoom
                dblZoom = dblZoom * (1 + (Y - mlngBaseY) * 0.001)
                If dblZoom < 64 And dblZoom > 0.01 Then
                    subCenterZoom DViewer.Images(1), DViewer, dblZoom, mCorpSize
                End If
                mlngBaseY = Y
'            Case 3  '�ü�����
'                Dim dcmLabel As DicomLabel
'                dcmView.Labels.Clear
'                Set dcmLabel = dcmView.Labels.AddNew
'                dcmLabel.LabelType = doLabelRectangle
'                dcmLabel.Left = mlngBaseX
'                dcmLabel.Top = mlngBaseY
'                dcmLabel.Width = x - mlngBaseX
'                dcmLabel.Height = y - mlngBaseY
            Case 11, 12, 3 '��ͷ��ע'Բ�α�ע,��ѡ
                mdcmSelectLabel.Width = DViewer.ImageXPosition(X, Y) - mdcmSelectLabel.Left
                mdcmSelectLabel.Height = DViewer.ImageYPosition(X, Y) - mdcmSelectLabel.Top
            Case 14
                '�϶�ͼ��......
                DViewer.Images(1).ScrollX = mInitScrollPoint.X - X
                DViewer.Images(1).ScrollY = mInitScrollPoint.Y - Y
        End Select
        
        DViewer.Refresh
    End If
End Sub


Private Sub DViewer_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If mblnDcmViewDown = True And Button = 1 And DViewer.Images.Count > 0 Then
        mblnDcmViewDown = False
        If mintMouseState = 13 Then     '���ֱ�ע
            
            txtInputText.Left = Me.ScaleX(X, vbPixels, vbTwips) + DViewer.Left
            txtInputText.Top = Me.ScaleY(Y, vbPixels, vbTwips) + DViewer.Top
            
            txtInputText.Text = ""
            txtInputText.Visible = True
            txtInputText.SetFocus
        ElseIf mintMouseState = 3 Then  '�ü�����
            
            '��ʾͼ�񱣴�˵�
            Call ShowFrameSelectImagePopup
            'ɾ����ѡ�õ���ʱ��ע
            If DViewer.Images(1).Labels.Count > 0 Then
                DViewer.Images(1).Labels.Remove DViewer.Images(1).Labels.Count
            End If
            
            Set mdcmSelectLabel = Nothing
            
            
'            dcmView.Labels.Clear
            
'            dcmView.Labels.Clear
'            RectangleZoom dcmView, dcmView.Images(1), mlngBaseX, mlngBaseY, x - mlngBaseX, y - mlngBaseY
        ElseIf mintMouseState = 14 Then
            '����ͼ�����ε�ƫ��λ��
            mCorpSize.X = mCorpSize.X + (DViewer.Images(1).ActualScrollX - mMouseDownPoint.X)
            mCorpSize.Y = mCorpSize.Y + (DViewer.Images(1).ActualScrollY - mMouseDownPoint.Y)
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
    
'    Call InitCommandBars    '����������
    
    '�ָ�����λ��
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitCommandBars
    
    Call LoadMemoFontStyle
    
'     '�ж��Ƿ��ǵ�һ��ͼ ���� ���һ��ͼ,������ذ�ť
'    Me.cmdCur.Enabled = IIf(mintCurImgIndex = 1, False, True)
'    Me.cmdNext.Enabled = IIf(mintCurImgIndex = mfrmParent.ImageCount, False, True)
    
    mCorpSize.X = 0
    mCorpSize.Y = 0
    mblnOK = False
    
    Call subSetMouseState(1)
    
    Call ReadEnjoin
End Sub

'���뱸ע������ʽ
Private Sub LoadMemoFontStyle()
    Dim strFontStyle As String
    Dim aryFontStyle() As String
    
    '������,12,B,U,S,I��
    
    strFontStyle = zlDatabase.GetPara("ͼ��ע����", glngSys, mlngModule, "")
    
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

    Call zlDatabase.SetPara("ͼ��ע����", strFontStyle, glngSys, mlngModule)
End Sub


Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣�ñ�ע
    Dim strSql As String, strPre As String
        
    On Error GoTo errH
    
    '��������
    strPre = cbxMemoText.Text '����󱣳�ԭ��ֵ
    cbxMemoText.Clear
    
    strSql = _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա=[1]" & _
        " Union" & _
        " Select ����,���� From Ӱ��ͼ��ע Where ���� is Not Null And ��Ա is Null" & _
        " Order by ����"
    Set mrsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����)
    Do While Not mrsTmp.EOF
        AddComboItem cbxMemoText.hWnd, CB_ADDSTRING, 0, mrsTmp!����
        
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
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
    
    Call SaveMemoFontStyle
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 32, 32
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    'ͼ���������������
    Set cbrToolBar = Me.cbrMain.Add("ͼ�������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '�ı���ʾ��ͼ���·�
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Window, "����"): cbrControl.ToolTipText = "��������/�Աȶ�"
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Zoom, "����"): cbrControl.ToolTipText = "����ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Corp, "�϶�"): cbrControl.ToolTipText = "�϶�ͼ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RectZoom, "�ü�"): cbrControl.ToolTipText = "�ü��ɼ�ͼ��": cbrControl.IconId = 3201
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_RRotate, "˳ʱ"): cbrControl.ToolTipText = "˳ʱ����ת"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_LRotate, "��ʱ"): cbrControl.ToolTipText = "��ʱ����ת"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Sharpness, "��"): cbrControl.ToolTipText = "��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Filter, "ƽ��"): cbrControl.ToolTipText = "ƽ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Arrow, "��ͷ"): cbrControl.ToolTipText = "��ͷ��ע"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ellipse, "Բ��"): cbrControl.ToolTipText = "Բ�α�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Text, "����"): cbrControl.ToolTipText = "���ֱ�ע"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Restore, "�ָ�"): cbrControl.ToolTipText = "�ָ�ͼ�񵽳�ʼ״̬"
        cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
         cbrControl.Category = "Main" '���ó�������˵�
    Next
    cbrToolBar.Position = xtpBarTop
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

