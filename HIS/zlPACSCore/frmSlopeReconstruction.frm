VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmSlopeReconstruction 
   Caption         =   "MPRб���ؽ�"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmSlopeReconstruction.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9915
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   2
      Left            =   5160
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   4
      Top             =   240
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   1
      Left            =   960
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin VB.PictureBox pic 
      Height          =   2775
      Index           =   3
      Left            =   5280
      ScaleHeight     =   2715
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   1455
         Index           =   3
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _Version        =   262147
         _ExtentX        =   2143
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   -2147483636
      End
   End
   Begin MSComctlLib.ImageList ImageListMouse 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":038A
            Key             =   "Stack"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":04EC
            Key             =   "Rotate2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0806
            Key             =   "Rotate"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0B20
            Key             =   "WindowWL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSlopeReconstruction.frx":0E3A
            Key             =   "Zoom"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   960
      Top             =   1800
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSlopeReconstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mParForm As frmViewer
Dim mImageViewer As DicomViewer      '��Ƭ�����У�ͼ�����ڵ�Viewer
Dim mAxialViewer As DicomViewer      'б���ؽ������У���λͼ�����ڵ�Viewer
Dim mCoronalViewer As DicomViewer     'б���ؽ������У���״λͼ�����ڵ�Viewer
Dim mSagittalViewer As DicomViewer   'Ь���ؽ������У�ʸ״λͼ�����ڵ�Viewer

Dim SelectedImage As DicomImage     '��ǰѡ�е�ͼ��
Dim SelectedLabel As DicomLabel     '��ǰѡ��ı�ע
Dim blnMoveLabel  As Boolean        '��ʼ�ƶ���ע�ı�ʶ
Dim blnMoveImage As Boolean         '��ʼ�϶����ı�ʶ���϶���꣬��״λ��ʸ״λͼ��ı�λ��
Dim lngBaseX As Long                '����ƶ���׼λ��
Dim lngBaseY As Long                '����ƶ���׼λ��
Dim lngBaseCenterX As Long          '�����ߵ����ĵ��׼λ��X
Dim lngBaseCenterY As Long          '�����ߵ����ĵ��׼λ��Y

Dim blnRebuild As Boolean           '�Ƿ�ɹ��ؽ���ͼ��

Private Enum RebuildType
    rt��ҳ = 0
    rtб���ؽ� = 1
    rtʸ��״λƽ�� = 2
End Enum

Dim mMouseAction As MouseAction
Private Enum MouseAction
    maĬ�� = 0
    ma��ת��ע = 1
    ma��ͼ = 2
    maƽ�Ʊ�ע = 3
End Enum

Public Function zlShowMe(parForm As frmViewer) As Boolean
    Dim iIndex As Integer
    
    On Error GoTo err
    
    blnRebuild = False
    
    Set mParForm = parForm
    
    If mParForm.intSelectedSerial = 0 Then
        MsgBox "����ѡ��һ��ͼ�����к��ٿ�ʼб���ؽ���"
        zlShowMe = True
        Exit Function
    End If
    
    '����Viewer������
    Set mImageViewer = mParForm.Viewer(mParForm.intSelectedSerial)
    Set mAxialViewer = Viewer(1)
    Set mCoronalViewer = Viewer(2)
    Set mSagittalViewer = Viewer(3)
    
    '�Ȱ������е�����ͼ�񶼼��ص�Viewer��
    Call funAddAllImages(mImageViewer)
    
    '�ж��Ƿ�����ʸ��״λ�ؽ�������
    If LeagelToACRebuild(mImageViewer.Images) = 1 Then
        zlShowMe = False   '�˳��ؽ�
        Exit Function
    End If
    
    '��ʼ������
    Call InitForm
    
    '��ʾ�м��ͼ��
    iIndex = mParForm.Viewer(mParForm.intSelectedSerial).Images.Count / 2
    If iIndex <= 0 Then
        iIndex = 1
    End If
    Call ShowAxialImage(mParForm.Viewer(mParForm.intSelectedSerial).Images(iIndex), iIndex)

    '�����ؽ���ͼ��
    If ShowImage = False Then
        zlShowMe = False
        Exit Function
    Else
        blnRebuild = True
    End If
    
    '����ʾ���壬�ټ����ؽ���ͼ��
    Me.Show 1, mParForm
    
    zlShowMe = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitForm()
'------------------------------------------------
'���ܣ���ʼ������
'��������
'���أ���
'------------------------------------------------
    Dim Pane1 As Pane
    Dim Pane2 As Pane
    Dim Pane3 As Pane
    Dim dGlabal As New DicomGlobal
    
    On Error GoTo err
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .TabPaintManager.BoldSelected = True
        .Options.DefaultPaneOptions = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set Pane1 = .CreatePane(1, 200, 200, DockTopOf)
        Set Pane2 = .CreatePane(2, 200, 200, DockBottomOf)
        Set Pane3 = .CreatePane(3, 600, 400, DockLeftOf, Pane1 And Pane2)
    End With
    
    '�ȴ�������MPRб���ؽ�������UID
    ZLMPRSlopeSeriesUID = dGlabal.NewUID
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Id = 1 Then
        Item.Handle = pic(2).hwnd
    ElseIf Item.Id = 2 Then
        Item.Handle = pic(3).hwnd
    ElseIf Item.Id = 3 Then
        Item.Handle = pic(1).hwnd
    End If
End Sub

Private Function ShowAxialImage(img As DicomImage, imageIndex As Integer) As Boolean
'------------------------------------------------
'���ܣ���ʾ��λͼ��MPR�ؽ����ͼ
'������ img -- ��ӵ���λͼλ�õ�ͼ��
'       imageIndex -- ͼ����������Viewer�е�ImageIndex
'���أ�True--�ɹ���False--ʧ��
'------------------------------------------------
    Dim oldImage As New DicomImage
    Dim blnCopyLabels As Boolean
    
    On Error GoTo err
    
    '���ԭ���Ѿ�����λͼ���ȱ������ͼ�񣬺�����Ҫ��������
    If mAxialViewer.Images.Count = 1 Then
        Set oldImage = mAxialViewer.Images(1)
        blnCopyLabels = True
    Else
        Set oldImage = Nothing
        blnCopyLabels = False
    End If
    
    mAxialViewer.Images.Clear
    mAxialViewer.Images.Add img
    '����Addͼ��֮�����½���һ��ͼ����tag�ı䣬��Ӱ��ԭͼ��
    mAxialViewer.Images(1).Tag = imageIndex
    
    '��ʼ����λͼ���MPR�ؽ�������
    Call funInitMPRControlLines(mAxialViewer.Images(1), False)
    
    '���Ʊ�ע
    If blnCopyLabels = True Then
        Call funCopyMPRControlLines(mAxialViewer.Images(1), oldImage)
    End If
    
    ShowAxialImage = True
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 
Private Function ShowImage() As Boolean
'------------------------------------------------
'���ܣ���ʾ��λͼ��MPR�ؽ����ͼ
'��������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    On Error GoTo err
    
    '��״λ�ؽ������Ͻ�Viewer(2)
    ShowImage = funMPRslope(mImageViewer, mAxialViewer, mCoronalViewer, mSagittalViewer, mParForm)
    
    Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim lngResult As Long
    Dim resImage As DicomImage
    
    If blnRebuild = True Then
        lngResult = MsgBox("�Ƿ񱣴��ؽ������", vbYesNoCancel, "��ʾ��Ϣ", Me)
        If lngResult = vbCancel Then
            Cancel = -1
            Exit Sub
        ElseIf lngResult = vbYes Then
            '�����ؽ����ͼ
            Set resImage = mAxialViewer.Images(1)
            resImage.SeriesUID = ZLMPRSlopeSeriesUID
            '������ͼ
            Call subSaveImage(resImage, mImageViewer.Images(1).SeriesUID)
            '��ͼ��׷�ӵ���Ƭվ��
            Call subOpenCurrentImage(mParForm, resImage)
        End If
        
        Set mImageViewer = Nothing
        Set mAxialViewer = Nothing
        Set mCoronalViewer = Nothing
        Set mSagittalViewer = Nothing
    End If
End Sub

Private Sub pic_Resize(Index As Integer)
'------------------------------------------------
'���ܣ�picture�����ı��С��ͬʱ�ı�����Viewer�Ĵ�С
'��������
'���أ���
'------------------------------------------------
    If Viewer.Count = 3 Then
        Viewer(Index).left = 0
        Viewer(Index).top = 0
        Viewer(Index).width = pic(Index).width
        Viewer(Index).height = pic(Index).height
    End If
End Sub

Private Sub Viewer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
'------------------------------------------------
'���ܣ�MouseDown��ѡ��ǰͼ�񣬵�ǰ��ע
'��������
'���أ���
'------------------------------------------------
    Dim ls As DicomLabels
    Dim j As Integer
    Dim m As Integer
    
    On Error GoTo err
    
    ''��¼���Ļ�׼λ��
    lngBaseX = Viewer(Index).ImageXPosition(x, y)
    lngBaseY = Viewer(Index).ImageYPosition(x, y)
    
    mMouseAction = maĬ��
    MousePointer = vbDefault
    
    '��굥��ʱ����ѡ��ͼ��
    If Viewer(Index).Images.Count <= 0 Then
        Set SelectedImage = Nothing
        Exit Sub
    Else
        Set SelectedImage = Viewer(Index).Images(1)
    End If
    
    If Button = 1 Then  '������
        Set ls = Viewer(Index).LabelHits(x, y, False, False, True)
        If ls.Count > 0 Then
            
            If Index = 1 And SelectedImage.Labels.IndexOf(ls(1)) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(ls(1)) <= G_INT_SYS_LABEL_MPR_POINT_O Then
                '����λͼ�ϵġ�MPR�ؽ�����صı�ע
                For j = 1 To ls.Count
                    If SelectedImage.Labels.IndexOf(ls(j)) > m Then m = SelectedImage.Labels.IndexOf(ls(j))
                Next
                Set SelectedLabel = SelectedImage.Labels(m)     'mΪ������ı�ע��
            ElseIf (Index = 2 Or Index = 3) And ((SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H) _
                Or (SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_V)) Then
                '�ǡ�ʸ��״�ؽ������ͼ�еĺ��ߺ�����
                Set SelectedLabel = ls(1)
                
                lngBaseCenterX = SelectedLabel.left + SelectedLabel.width / 2
                lngBaseCenterY = SelectedLabel.top + SelectedLabel.height / 2
        
                'ֻ�к��߿�����ת���ı������״
                If SelectedImage.Labels.IndexOf(ls(1)) = G_INT_SYS_LABEL_MPR_RESULT_H Then
                    If (Abs(SelectedLabel.width) > Abs(SelectedLabel.height) And _
                        (lngBaseX < SelectedLabel.left + SelectedLabel.width / 8 Or lngBaseX > SelectedLabel.left + SelectedLabel.width * 7 / 8)) _
                        Or (Abs(SelectedLabel.height) > Abs(SelectedLabel.width) And _
                        (lngBaseY < SelectedLabel.top + SelectedLabel.height / 8 Or lngBaseY > SelectedLabel.top + SelectedLabel.height * 7 / 8)) Then
                        MousePointer = vbCustom ' vbNoDrop 'vbCustom
                        MouseIcon = ImageListMouse.ListImages("Rotate").Picture
                        mMouseAction = ma��ת��ע
                    Else
                        MousePointer = vbSizeAll
                        mMouseAction = maƽ�Ʊ�ע
                    End If
                Else
                    MousePointer = vbSizeAll
                    mMouseAction = maƽ�Ʊ�ע
                End If
            End If
            
            blnMoveLabel = True
        Else
            'û��ѡ�б�ע�������ƶ�ͼ��
            blnMoveImage = True
            MousePointer = vbCustom
            MouseIcon = ImageListMouse.ListImages("Stack").Picture
            mMouseAction = ma��ͼ
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Viewer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
'------------------------------------------------
'���ܣ��ƶ���ע���ؽ�
'��������
'���أ���
'------------------------------------------------
    If SelectedImage Is Nothing Then Exit Sub
    
    On Error GoTo err

    '�ƶ���ע
    If blnMoveLabel And Not SelectedLabel Is Nothing Then
        subaCorrectCursor Viewer(Index), SelectedImage, x, y     ''''����ƶ��������ͼ��Χ�����������λ��
        
        '�ƶ���ע�����в����������ƶ�MPR�߲���ʾ�ؽ����ͼ
        subMoveSlopeLabel SelectedLabel, Viewer(Index).ImageXPosition(x, y), _
            Viewer(Index).ImageYPosition(x, y), lngBaseX, lngBaseY, Index
        
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        
        Viewer(Index).Refresh
    ElseIf blnMoveImage = True Then
        subaCorrectCursor Viewer(Index), SelectedImage, x, y     ''''����ƶ��������ͼ��Χ�����������λ��
        '��ק����ʱ���л���λ����״λ��ʸ״λͼ���λ��
        subMoveImage Viewer(Index).ImageXPosition(x, y), Viewer(Index).ImageYPosition(x, y), lngBaseX, lngBaseY, Index
        
        lngBaseX = Viewer(Index).ImageXPosition(x, y)
        lngBaseY = Viewer(Index).ImageYPosition(x, y)
        
        Viewer(Index).Refresh
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveSlopeLabel(la As DicomLabel, newX As Long, newY As Long, _
    baseX As Long, baseY As Long, ViewerIndex As Integer)
'------------------------------------------------
'���ܣ��ƶ�һ����ע��������λͼ��ʸ״λ�͹�״λ�Ŀ�����
'������ la -- ���ƶ��ı�ע
'       newX -- ��λ�õ�ͼ������X����
'       newY -- ��λ�õ�ͼ������Y����
'       basex -- ��λ�õ�ͼ������X����
'       baseY -- ��λ�õ�ͼ������Y����
'       ViewerIndex -- ͼ������Viewer��Index��1-��λͼ��2-��״λͼ��3-ʸ״λͼ
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    '����λͼ���ϣ��ƶ�������
    If SelectedImage.Labels.IndexOf(la) >= G_INT_SYS_LABEL_MPRV And SelectedImage.Labels.IndexOf(la) <= G_INT_SYS_LABEL_MPR_POINT_O Then ''[ʸ��״�ߵ��ƶ�]
        '�ƶ�ʸ��״�ؽ����Ƶ㡢�ߣ��������µ��ؽ�ͼ��
        Call subMoveAxialMPRLabel(la, newX, newY, baseX, baseY)
    ElseIf (SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_H) _
        Or (SelectedImage.Labels.IndexOf(la) = G_INT_SYS_LABEL_MPR_RESULT_V) Then
        '�ڹ�״λ��ʸ״λͼ���ϣ��ƶ�������
        Call subMoveCAndSLabel(la, SelectedImage, newX, newY, baseX, baseY, IIf(ViewerIndex = 2, True, False))
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveAxialMPRLabel(la As DicomLabel, xx As Long, Yy As Long, _
    baseX As Long, baseY As Long)
'------------------------------------------------
'���ܣ��ƶ���λͼʸ��״�ؽ����Ƶ㡢�ߣ��������µ��ؽ�ͼ��
'������
'       la -- ���ƶ���ʸ��״�ؽ����Ƶ������ߣ�
'       xx -- ��ע��λ����ͼ���ϵ�X���ꣻ
'       yy -- ��ע��λ����ͼ���ϵ�Y���ꣻ
'       basex -- ��ע��λ�õ�ͼ���ϵ�x���ꣻ
'       baseY -- ��ע��λ�õ�ͼ���ϵ�y���ꡣ
'���أ��ޣ�ֱ���ƶ�ʸ��״�ؽ��Ŀ��Ƶ���ߣ��������ؽ����ͼ��
'------------------------------------------------
    Dim intIndex As Integer
    Dim axialImage As DicomImage
    
    On Error GoTo err
    
    '����ֱ��ʹ��SelectedImage���������������ʱ��SelectedImage�����ǹ�״λ��ʸ״λͼ
    Set axialImage = mAxialViewer.Images(1)
    
    ''''''''''''''''''''''[���Ľǵ���ƶ�]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intIndex = axialImage.Labels.IndexOf(la)
    
    'ʸ��״���Ƶ����ĸ��ߵ���ƶ�����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRV), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 Then
        Call subPeriodMove(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), xx, Yy, _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPRH), _
                            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), axialImage)
    '''''''''''''''''''''''''''���ĵ���ƶ�'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_POINT_O Then
        axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top + Yy - baseY
        axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left + xx - baseX
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top < -G_INT_MPR_RADIUS / 2 + 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left < -G_INT_MPR_RADIUS / 2 + 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = -G_INT_MPR_RADIUS / 2 + 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top > axialImage.sizeY - G_INT_MPR_RADIUS / 2 - 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).top = axialImage.sizeY - G_INT_MPR_RADIUS - 1
        End If
        If axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left > axialImage.sizeX - G_INT_MPR_RADIUS / 2 - 1 Then
            axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O).left = axialImage.sizeX - G_INT_MPR_RADIUS - 1
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'ʸ��״���ĵ���ƶ�
        If xx <> baseX Then
            Call subPeriodMovee5X(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V1), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_V2), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPRV), xx, Yy, _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), baseX, baseY, axialImage)
        End If
        
        If Yy <> baseY Then
            Call subPeriodMovee5X(axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H1), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_H2), _
                                axialImage.Labels(G_INT_SYS_LABEL_MPRH), xx, Yy, _
                                axialImage.Labels(G_INT_SYS_LABEL_MPR_POINT_O), baseX, baseY, axialImage)
        End If
    End If
    
    '��ת��ע��ˢ��ͼ����ʾ
    Call axialImage.Refresh(False)
    
    ''''''''''�����ؽ�''''''''''''''''''''''''''''''''''''''''''''''
    '��ע��MPR���������ߵ������˵㣬������MPR���������ߵ����ĵ㣬��ʱ��Ҫ�ƶ�����MPR����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_V1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_V2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And xx <> baseX) Then
        
        '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
        If funGetCandSImageAndShow(axialImage.Labels(G_INT_SYS_LABEL_MPRV), mImageViewer, _
                                        mAxialViewer, mSagittalViewer, ToltalHeight, 1, False, True) = False Then
            '�ؽ������˳�MPR�ؽ�
            Exit Sub
        End If
    End If
    
    '��ע��MPR�����ߺ��ߵ������˵㣬������MPR�����ߵĺ������ĵ㣬��ʱ��Ҫ�ƶ�����MPR����
    If intIndex = G_INT_SYS_LABEL_MPR_POINT_H1 Or intIndex = G_INT_SYS_LABEL_MPR_POINT_H2 _
        Or (intIndex = G_INT_SYS_LABEL_MPR_POINT_O And Yy <> baseY) Then
        
        '���ݴ���Ŀ����ߣ���MPRViewer�ڵ�ͼ������ؽ������������ؽ��Ľ��ͼ��ʾ��ShowViewerIndexָ����Viewer��
        If funGetCandSImageAndShow(axialImage.Labels(G_INT_SYS_LABEL_MPRH), mImageViewer, _
                                        mAxialViewer, mCoronalViewer, ToltalHeight, 2, False, True) = False Then
            '�ؽ������˳�MPR�ؽ�
            Exit Sub
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub subMoveCAndSLabel(la As DicomLabel, img As DicomImage, newX As Long, newY As Long, _
    baseX As Long, baseY As Long, blnIsCoronal As Boolean)
'------------------------------------------------
'���ܣ��ƶ�ʸ��״�ؽ���״λ��ʸ״λͼ�Ŀ����ߣ����߿�����λͼ����Զ���ҳ�����߿��ƽ��ͼ
'������
'       la--���ƶ���ʸ��״�ؽ����Ƶ������ߣ�
'       img -- ��ע���ڵ�ͼ��
'       newX--��λ�õ�ͼ������X���ꣻ
'       newY--��λ�õ�ͼ������Y���ꣻ
'       basex--��λ�õ�ͼ������x���ꣻ
'       baseY--��λ�õ�ͼ������y����
'       blnIsCoronal -- �Ƿ��״λ��True-��״λ��False-ʸ״λ
'���أ��ޣ�ֱ���ƶ�ʸ��״�ؽ��Ľ����
'------------------------------------------------
    Dim iImageIndex As Integer
    Dim intIndex As Integer
    Dim rtType As RebuildType
    Dim k As Double
    Dim laAxial As DicomLabel
    
    On Error GoTo err
    
    '�ڹ�״λ��ʸ״λ���ƶ������ߣ�SelectedImage�ǹ�״λ��ʸ״λͼ
    '����ע�ߵ��ƶ���ͼ���ؽ������ֿ�����ͬ�ı�ע�ƶ�������������ʽ��ͬ
    
    intIndex = img.Labels.IndexOf(la)
    
    If intIndex = G_INT_SYS_LABEL_MPR_RESULT_H And (mMouseAction = maƽ�Ʊ�ע Or mMouseAction = ma��ͼ) Then    '����ƽ�ƣ���λͼ����ͼ�����б���ؽ�
        '���ж��Ǹ�����λͼ�񣬻���б���ؽ�
        If la.height = 0 Then
            If (blnIsCoronal = True And mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0) _
                Or (blnIsCoronal = False And mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H).height = 0) Then
                rtType = rt��ҳ
            Else
                rtType = rtб���ؽ�
            End If
        Else
            rtType = rtб���ؽ�
        End If

        '�ƶ���ע��
        If la.width <> 0 Then
            '����б��
            k = la.height / la.width
            If k = 0 Then
                'б��Ϊ0 ��ֱ���ƶ�Y����
                If la.top + (newY - baseY) > 0 And la.top + (newY - baseY) < img.sizeY Then
                    la.top = la.top + (newY - baseY)
                End If
            Else
                Call funGetLine(la, img, k, la.left + (newX - baseX), la.top + (newY - baseY))
            End If
        End If
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_H And mMouseAction = ma��ת��ע Then  '������ת��б���ؽ�
        '��ת���������
        Call subRotateLabel(la, newX, newY, img)
        
        rtType = rtб���ؽ�
    ElseIf intIndex = G_INT_SYS_LABEL_MPR_RESULT_V Then '����ƽ�ƣ����ؽ�ͼ��
        '���ƶ�ʸ��״�ؽ������
        If la.left + (newX - baseX) > 0 And la.left + (newX - baseX) < img.sizeX Then
            la.left = la.left + (newX - baseX)
        End If
        
        rtType = rtʸ��״λƽ��
    End If
    
    If rtType = rt��ҳ Then
        '���ݽ���ߵ�λ�ã���ʾ�µ���λͼ��
        iImageIndex = la.top / img.sizeY * mImageViewer.Images.Count
        If iImageIndex > 0 And iImageIndex <= mImageViewer.Images.Count Then
            '����Ӧ��ͼ����ʾ����λͼ��
            
            Call ShowAxialImage(mParForm.Viewer(mParForm.intSelectedSerial).Images(iImageIndex), iImageIndex)
            
            '������һ���ؽ�ͼ�ı�עλ��
            If blnIsCoronal = True Then
                Call subMPRSlopeDrawResultControlLabels(la, mSagittalViewer.Images(1), mImageViewer, mAxialViewer)
            Else
                Call subMPRSlopeDrawResultControlLabels(la, mCoronalViewer.Images(1), mImageViewer, mAxialViewer)
            End If
        End If
    ElseIf rtType = rtб���ؽ� Then
        '�����ת��ƽ���˹�״λ��ʸ״λ���ĺ�������ߣ�����Ҫ����ʸ״λ����״λ��ͼ�����ߵ�λ��
        '����ʸ״λͼ�͹�״λͼ�����ߵ����ĵ��غϣ���������ͬһ��ƽ����
        Call funTranslateLabel(blnIsCoronal)
        'б���ؽ�
        Call funGetSlopeImageAndShow
    ElseIf rtType = rtʸ��״λƽ�� Then
        '���ƽ���˹�״λ��ʸ״λ������������ߣ�����Ҫ����ʸ״λ����״λ��ͼ�����ߵ�λ��
        '����ʸ״λͼ�͹�״λͼ�����ߵ����ĵ��غϣ���������ͬһ��ƽ����
        Call funTranslateLabel(blnIsCoronal)
        '��������״λ��ʸ״λ���ؽ�
        If blnIsCoronal = True Then
            Set laAxial = mImageViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRV)
            laAxial.left = la.left
            laAxial.top = 0
            laAxial.width = 0
            laAxial.height = mImageViewer.Images(1).sizeY
            Call funGetCandSImageAndShow(laAxial, mImageViewer, mAxialViewer, mSagittalViewer, ToltalHeight, 2, False, True)
        Else
            Set laAxial = mImageViewer.Images(1).Labels(G_INT_SYS_LABEL_MPRH)
            laAxial.left = 0
            laAxial.top = la.left
            laAxial.width = mImageViewer.Images(1).sizeX
            laAxial.height = 0
            Call funGetCandSImageAndShow(laAxial, mImageViewer, mAxialViewer, mCoronalViewer, ToltalHeight, 1, False, True)
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Viewer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
    blnMoveLabel = False
    blnMoveImage = False
    MousePointer = vbDefault
    mMouseAction = maĬ��
End Sub

Private Sub subRotateLabel(la As DicomLabel, xNew As Long, yNew As Long, im As DicomImage)
'------------------------------------------------
'���ܣ����ߵ����ĵ�Ϊ���ģ���ת��״λ��ʸ״λͼ�ϵĺ���
'������
'       la--����ת��ʸ��״λͼ�еĿ����ߣ�
'       xNew--��λ����ͼ���е�X���ꣻ
'       yNew--��λ����ͼ���е�Y���ꣻ
'       im--���������ڵ�ͼ��
'���أ��ޣ�ֱ����ת��ע
'------------------------------------------------
    Dim x0 As Double, y0 As Double  'ֱ�����ĵ�����
    Dim k As Double             'ֱ�ߵ�б��
    
    On Error GoTo err
    
    '���ĵ�λ��
    x0 = lngBaseCenterX
    y0 = lngBaseCenterY
    
    '����б��
    k = (yNew - y0) / (xNew - x0)
    
    '����б�ʺ�һ���㣬������ע��
    Call funGetLine(la, im, k, x0, y0)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function funSlopeRebuild() As DicomImage
'------------------------------------------------
'���ܣ��Թ�״λ��ʸ״λͼ�ϵĺ���Ϊ��׼������λͼ�н���б���ؽ�
'������
'���أ��ޣ������ؽ����ͼ��
'------------------------------------------------
    Dim resImage As DicomImage
    Dim X1 As Double, Y1 As Double, Z1 As Double
    Dim X2 As Double, Y2 As Double, Z2 As Double
    Dim X3 As Double, Y3 As Double, Z3 As Double
    Dim X4 As Double, Y4 As Double, Z4 As Double
    Dim A As Double, B As Double, C  As Double, D As Double
    Dim zIndex As Long                  'б��ͼ����z���������
    Dim ZZ As Double
    Dim sizeX As Long, sizeY As Long    '��ͼ��Ŀ�Ⱥ͸߶�
    Dim i As Integer, j As Integer
    Dim v As Variant
    Dim lines() As Integer              '����ͼ��Ҷ�ֵ�Ķ�ά����
    Dim zIndexOld As Long               '�����ϴζ�ȡͼ���λ��
    
    '�㷨��
    '����֪�����������ֱ���P1(x1,y1,z1)��P2(x2,y2,z2)��P3(x3,y3,z3)��ʾ����P1��P2��P3����ͬһ��ֱ���ϡ���
    '��ͨ��P1��P2��P3�����ƽ�淽��ΪA(x - x1) + B(y - y1) + C(z - z1) = 0 ��
    '����Ϊһ��ʽ��Ax + By + Cz + D = 0��
    '��P1(x1,y1,z1)����ֵ���뷽��Ax + By + Cz + D = 0��
    '���ɵõ���Ax1 + By 1+ Cz1 + D = 0��
    '�����D = -(A * x1 + B * y1 + C * z1)��
    '����Ը���P1(x1,y1,z1)��P2(x2,y2,z2)��P3(x3,y3,z3)��������ֱ����A��B��C��ֵ�����£�
    'A = (y3 - y1)*(z3 - z1) - (z2 -z1)*(y3 - y1);
    'B = (x3 - x1)*(z2 - z1) - (x2 - x1)*(z3 - z1);
    'C = (x2 - x1)*(y3 - y1) - (x3 - x1)*(y2 - y1);
    '��D = -(A * x1 + B * y1 + C * z1)�����Կ������D��ֵ��
    '����õ�A��B��C��Dֵ����һ��ʽ���̾Ϳɵù�P1��P2��P3��ƽ�淽��:
    'Ax + By + Cz + D = 0 (һ��ʽ)
    
    '��״λͼ�У��������������ά�������е���������ΪP1��P2������P1(X1,Y1,Z1),P2(X2,Y2,Z2)
    'ʸ״λͼ�У��������������ά�������е���������ΪP3��P4������P3(X3,Y3,Z3),P4(X4,Y4,Z4)
    
    On Error GoTo err
    
    Set funSlopeRebuild = Nothing
    zIndexOld = -1
    
    'ֱ�Ӵӹ�״λ��ʸ״λͼ�ͺ��ߣ���ȡ��AB,CD������
    If funGetTwoPointsFromImg(mCoronalViewer.Images(1), X1, Y1, Z1, X2, Y2, Z2) = False Then
        Exit Function
    End If
    If funGetTwoPointsFromImg(mSagittalViewer.Images(1), X3, Y3, Z3, X4, Y4, Z4) = False Then
        Exit Function
    End If

    'ʹ������ȷ��һ��ƽ�棬����P1P2��P3P4֮��Ľ��㲻������Ԥ�Ƶ����ĵ㣬��ϵҲ����ͨ��P1,P2,P3Ҳ���Եõ����λ�ø�����б��ͼ
    '��ƽ�淽�̵�A,B,C,D��ʹ����֪�ĸ����е�������P1(X1,Y1,Z1),P2(X2,Y2,Z2),P3(X3,Y3,Z3)
    
'    '����㷨�д���
'    ''A = (y3 - y1)*(z3 - z1) - (z2 -z1)*(y3 - y1);
'    A = (Y3 - Y1) * (Z3 - Z1) - (Z2 - Z1) * (Y3 - Y1)
'    ''B = (x3 - x1)*(z2 - z1) - (x2 - x1)*(z3 - z1);
'    B = (X3 - X1) * (Z2 - Z1) - (X2 - X1) * (Z3 - Z1)
'    ''C = (x2 - x1)*(y3 - y1) - (x3 - x1)*(y2 - y1);
'    C = (X2 - X1) * (Y3 - Y1) - (X3 - X1) * (Y2 - Y1)
'    ''D = -(A * x1 + B * y1 + C * z1)
'    D = -(A * X1 + B * Y1 + C * Z1)
    
    '��һ�ּ���ABCD�ķ�����a=y1z2-y1z3-y2z1+y2z3+y3z1-y3z2,b=-x1z2+x1z3+x2z1-x2z3-x3z1+x3z2,
    'c=x1y2-x1y3-x2y1+x2y3+x3y1-x3y2,d=-x1y2z3+x1y3z2+x2y1z3-x2y3z1-x3y1z2+x3y2z1
    A = Y1 * Z2 - Y1 * Z3 - Y2 * Z1 + Y2 * Z3 + Y3 * Z1 - Y3 * Z2
    B = -X1 * Z2 + X1 * Z3 + X2 * Z1 - X2 * Z3 - X3 * Z1 + X3 * Z2
    C = X1 * Y2 - X1 * Y3 - X2 * Y1 + X2 * Y3 + X3 * Y1 - X3 * Y2
    D = -X1 * Y2 * Z3 + X1 * Y3 * Z2 + X2 * Y1 * Z3 - X2 * Y3 * Z1 - X3 * Y1 * Z2 + X3 * Y2 * Z1
    
    If C = 0 Then
        Exit Function
    End If
    
    '��ȡ�ؽ����ͼ�������ݣ����ؽ�б���У���ԭ�㿪ʼ���������ȡ����ֵ
    sizeX = mImageViewer.Images(1).sizeX
    sizeY = mImageViewer.Images(1).sizeY
    
    '���¶���ԭͼͼ��Ҷ�ֵ��ά����
    ReDim lines(sizeX, sizeY) As Integer
    
    '����ƽ���һ��ʽ���̣���ȡZ���� Ax + By + Cz + D = 0
    'z = (-D-Ax-By)/C
    If SafeArrayGetDim(aPixels) = 0 Then
        'MPR�Ļ�����ά������ά��=0��˵�������ڴ���ɣ���ֱ��ʹ��ͼ���������ؽ���ͼ��Խ�࣬�ؽ�Խ��
        For i = 1 To sizeX
            For j = 1 To sizeY
                ZZ = (-D - A * i - B * j) / C   '�ܸ߶��ϵ�z����
                
                '��zIndex�����ͼ���
                zIndex = mImageViewer.Images.Count / ToltalHeight * ZZ
                
                If zIndex < 1 Or zIndex > mImageViewer.Images.Count Then
                    lines(i, j) = 0
                Else
                    If zIndexOld <> zIndex Then
                        v = mImageViewer.Images(zIndex).Pixels
                        zIndexOld = zIndex
                    End If
                    lines(i, j) = v(i, j, 1)
                End If
            Next j
        Next i
    Else
        'ʹ����ά���鱣����ά�����ݣ���MPR�ؽ���ÿ���ؽ��ٶ���1������
        For i = 1 To sizeX
            For j = 1 To sizeY
                ZZ = (-D - A * i - B * j) / C   '�ܸ߶��ϵ�z����
                
                '��zIndex�����ͼ���
                zIndex = mImageViewer.Images.Count / ToltalHeight * ZZ
                
                If zIndex < 1 Or zIndex > mImageViewer.Images.Count Then
                    lines(i, j) = 0
                Else
                    lines(i, j) = aPixels(i, j, zIndex)
                End If
            Next j
        Next i
    End If
    
    'ƽ����������512��ͼ���ؽ��ٶ��������ȸߣ�����ƽ��
    If sizeX <= 512 Then
        Call funImageSmoothing(lines(), 1)
    End If
    
    '������ͼ��
    Set resImage = mImageViewer.Images(1).SubImage(0, 0, sizeX, sizeY, 1, 1)
    
    'ɾ��һЩ���õ�λ������
    resImage.Attributes.Remove &H18, &H50
    resImage.Attributes.Remove &H18, &H1110
    resImage.Attributes.Remove &H18, &H1111
    resImage.Attributes.Remove &H18, &H1120     'Tilt
    resImage.Attributes.Remove &H18, &H1140     'Rotation Direction
    resImage.Attributes.Remove &H18, &H5100     'Patient Position
    resImage.Attributes.Remove &H20, &H32       'Image Position(Patient)
    resImage.Attributes.Remove &H20, &H37       'Image Orientation (Patient)
    resImage.Attributes.Remove &H20, &H1041     'Slice Location
    
    '���ý��ͼ�����ԣ���Щ����Ҫ�޸�
'    resImage.Attributes.Add &H28, &H10, intToltalHeight
'    resImage.Attributes.Add &H28, &H11, iPointsCount
'    If intType = 1 Then
'        resImage.Attributes.Add &H20, &H11, LineLong(1).y
'    Else
'        resImage.Attributes.Add &H20, &H11, LineLong(1).x
'    End If
'    resImage.Attributes.Add &H20, &H13, intType
    resImage.Pixels = lines
    resImage.width = mImageViewer.Images(1).width
    resImage.Level = mImageViewer.Images(1).Level
    
    '���ؽ��ͼ
    Set funSlopeRebuild = resImage
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetTwoPointsFromImg(im As DicomImage, ByRef X1 As Double, _
    ByRef Y1 As Double, ByRef Z1 As Double, ByRef X2 As Double, ByRef Y2 As Double, _
    ByRef Z2 As Double) As Boolean
'------------------------------------------------
'���ܣ���ȡ��״λͼ��ʸ״λͼ���棬��������ά�������е�����
'������ im -- ��״λ��ʸ״λͼ
'       x1,y1,z1 -- ��һ���������
'       x2,y2,z2 -- �ڶ����������
'���أ�true -- �ɹ��� false -- ʧ��
'------------------------------------------------
    Dim isCoronal As Boolean
    Dim la As DicomLabel
    Dim AX As Double, AY  As Double, BX  As Double, BY  As Double 'ͼ���ϱ�ע��AB�����˵������
    Dim blnDoublePos As Boolean
    Dim Poss() As String
    
    On Error GoTo err
        
    '���ж��ǹ�״λ����ʸ״λ
    If im.Attributes(&H20, &H13).Value = 1 Then
        isCoronal = True
    Else
        isCoronal = False
    End If
    
    '��ȡ��״λ��ʸ״λͼ�еĺ��������
    Set la = im.Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
    
    '����la����ά�������е�����
    '��ά�����ݣ���ͷ��ſ������Ͻ�Ϊԭ�㣨0,0,0��������λͼ�����Ͻǵ�Ϊԭ�㣨0,0��
    '����la��ͼ���������˵�����꣬�������������λ��˳��
    blnDoublePos = False
    If la.Tag <> "" Then
        Poss = Split(la.Tag, ":")
        If UBound(Poss) = 3 Then
            AX = CDbl(Poss(0))
            AY = CDbl(Poss(1))
            BX = CDbl(Poss(0)) + CDbl(Poss(2))
            BY = CDbl(Poss(1)) + CDbl(Poss(3))
            blnDoublePos = True
        End If
    End If
    If blnDoublePos = False Then
        AX = la.left
        AY = la.top
        BX = la.left + la.width
        BY = la.top + la.height
    End If
    
    '��ά���꣬ת������ά����
    If isCoronal = True Then
        X1 = AX
        Y1 = im.Attributes(&H20, &H11).Value
        Z1 = AY
        X2 = BX
        Y2 = Y1
        Z2 = BY
    Else
        X1 = im.Attributes(&H20, &H11).Value
        Y1 = AX
        Z1 = AY
        X2 = X1
        Y2 = BX
        Z2 = BY
    End If
    
    funGetTwoPointsFromImg = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetSlopeImageAndShow() As Boolean
'------------------------------------------------
'���ܣ� ��б���ؽ�ͼ�񣬶�imageViewer�е�ͼ������ؽ����������ͼ��ʾ��mAxialViewer��
'������
'����:��
'------------------------------------------------
    Dim resImage As DicomImage
    
    On Error GoTo err
    
    '��ȡ�ؽ����ͼ
    Set resImage = funSlopeRebuild()
    
    '��ʾ���ͼ
    If resImage Is Nothing Then
        funGetSlopeImageAndShow = False
        Exit Function
    Else
        '�����ͼ��ӵ���λͼ��λ��
        mAxialViewer.Images.Clear
        mAxialViewer.Images.Add resImage
        If mAxialViewer.Images(1).Labels.Count = 0 Then
            Call subInitAImage(mAxialViewer.Images(1), 0, mAxialViewer)
        End If
    End If
    
    funGetSlopeImageAndShow = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funTranslateLabel(isCoronal As Boolean) As Boolean
'------------------------------------------------
'���ܣ� ƽ�ƺ�������ߣ�ȷ����״λͼ��ʸ״λͼ�ĺ�����������ĵ��غ�
'       �����������������߲ſ���ȷ��һ���ؽ���б��
'������
'����:��
'------------------------------------------------
    Dim lblRotate As DicomLabel         '��ת�ı�ע
    Dim lblTranslate As DicomLabel      'ƽ�Ƶı�ע
    Dim imgTranslate As DicomImage      'ƽ�Ʊ�ע��ͼ��
    Dim x0 As Double, y0 As Double      '��ת���ע�����ĵ�
    Dim xT0 As Double, yT0 As Double    'ƽ��ǰ��ע�����ĵ�
    Dim k As Double                     'б��
    
    On Error GoTo err
    
    If isCoronal = True Then    '��ת�˹�״λ�ĺ�������ߣ���ƽ��ʸ״λ�Ŀ�����
        Set lblRotate = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set lblTranslate = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set imgTranslate = mSagittalViewer.Images(1)
        x0 = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
    Else    '��ת��ʸ״λ�ĺ�������ߣ���ƽ�ƹ�״λ�Ŀ�����
        Set lblRotate = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set lblTranslate = mCoronalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Set imgTranslate = mCoronalViewer.Images(1)
        x0 = mSagittalViewer.Images(1).Labels(G_INT_SYS_LABEL_MPR_RESULT_V).left
    End If
    
    '����ƽ�Ʊ�ע����ƽ��Ľ���λ��
    
    xT0 = lblTranslate.left + lblTranslate.width / 2
    yT0 = lblTranslate.top + lblTranslate.height / 2
    
    If lblRotate.width = 0 Then
        x0 = xT0
        y0 = yT0
    Else
        y0 = (x0 - lblRotate.left) * lblRotate.height / lblRotate.width + lblRotate.top
    End If
        
    If xT0 = x0 And yT0 = y0 Then
        '���ô���
    Else
        If lblTranslate.width = 0 Then  '��ֱ��
            lblTranslate.left = x0
            lblTranslate.top = 0
            lblTranslate.width = 0
            lblTranslate.height = imgTranslate.sizeY
            lblTranslate.Tag = x0 & ":0:0:" & imgTranslate.sizeY
        ElseIf lblTranslate.height = 0 Then '�Ǻ���
            lblTranslate.left = 0
            lblTranslate.top = y0
            lblTranslate.width = imgTranslate.sizeX
            lblTranslate.height = 0
            lblTranslate.Tag = "0:" & y0 & ":" & imgTranslate.sizeY & ":0"
        Else
            '�ȼ���б��
            k = lblTranslate.height / lblTranslate.width
            Call funGetLine(lblTranslate, imgTranslate, k, x0, y0)
        End If
        imgTranslate.Refresh (False)
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function funGetLine(la As DicomLabel, im As DicomImage, k As Double _
    , x0 As Double, y0 As Double) As Boolean
'------------------------------------------------
'���ܣ� ����б�ʺ�һ���㣬������ע��
'������  la -- ��ע��
'       im -- ��ע�����ڵ�ͼ��
'       k -- б��
'       x0��y0 -- ֱ����һ���������
'����: ֱ�ӻ���ע��
'------------------------------------------------
    Dim xA As Double, yA As Double  'ֱ��AB��A������
    Dim xB As Double, yB As Double  'ֱ��AB��B������
    Dim xC As Double, yC As Double
    Dim xD As Double, yD As Double
    Dim xAA As Double, yAA As Double
    Dim xBB As Double, yBB As Double
    Dim strTag As String
    Dim intPoints As Integer
    
    On Error GoTo err
    
    If k = 0 Then Exit Function
    
    '�㷨��
    '������������ֱ�Ϊ��x1��y1����x2��y2��,б��ʽ
    '��б��: k = (y2 - y1) / (x2 - x1)
    'ֱ�߷��� y - y1 = k(x - x1)
    '�ٰ�k����y-y1=k(x-x1)���ɵõ�ֱ�߷��̡�
    'ʵ�����Ǽ���һ��ֱ�ߺ;��εĽ���
    '�����ΪEFGH,�ֱ������������ߵĽ���
    'E-------F
    '|       |
    'G-------H
    
    intPoints = 0
    
    '��EG�Ľ���
    xA = 0
    yA = k * (xA - x0) + y0
    If yA < 0 Then
        yA = 0
        xA = (yA - y0) / k + x0
    ElseIf yA > im.sizeY Then
        yA = im.sizeY
        xA = (yA - y0) / k + x0
    End If
    If xA >= 0 And xA <= im.sizeX Then
        intPoints = 1
        xAA = xA
        yAA = yA
    End If
    
    '��FH�Ľ���
    xB = im.sizeX
    yB = k * (xB - x0) + y0
    If yB < 0 Then
        yB = 0
        xB = (yB - y0) / k + x0
    ElseIf yB > im.sizeY Then
        yB = im.sizeY
        xB = (yB - y0) / k + x0
    End If
    If xB >= 0 And xB <= im.sizeX Then
        If intPoints = 1 Then
            xBB = xB
            yBB = yB
            intPoints = 2
        Else
            xAA = xB
            yAA = yB
            intPoints = 1
        End If
    End If
    
    '��EF�Ľ���
    If intPoints < 2 Then
        yC = 0
        xC = k * (xC - x0) + y0
        If xC < 0 Then
            xC = 0
            yC = k * (xC - x0) + y0
        ElseIf xC > im.sizeX Then
            xC = im.sizeX
            yC = k * (xC - x0) + y0
        End If
        If yC >= 0 And yC <= im.sizeY Then
            If intPoints = 1 Then
                xBB = xC
                yBB = yC
                intPoints = 2
            Else
                xAA = xC
                yAA = yC
                intPoints = 1
            End If
        End If
    End If
    
    '��GH�Ľ���
    If intPoints < 2 Then
        yD = im.sizeY
        xD = k * (xD - x0) + y0
        If xD < 0 Then
            xD = 0
            yD = k * (xD - x0) + y0
        ElseIf xD > im.sizeX Then
            xD = im.sizeX
            yD = k * (xD - x0) + y0
        End If
        If yD >= 0 And yD <= im.sizeY Then
            If intPoints = 1 Then
                xBB = xD
                yBB = yD
                intPoints = 2
            End If
        End If
    End If
    
    '���A����B����ұߣ�����AB���λ��
    If xBB < xAA Then
        xA = xAA
        yA = yAA
        xAA = xBB
        yAA = yBB
        xBB = xA
        yBB = yA
    End If
    
    '����ҵ��������㣬�ŵ�����ע��λ��
    If intPoints = 2 And Not (xAA = xBB And yAA = yBB) Then
        la.top = yAA
        la.left = xAA
        
        strTag = xAA & ":" & yAA
        
        If la.top = im.sizeY Then
            la.width = xBB - xAA
            la.height = yBB - yAA
            strTag = strTag & ":" & (xBB - xAA) & ":" & (yBB - yAA)
        ElseIf la.left = 0 Then
            la.width = xBB
            la.height = yBB - yAA
            strTag = strTag & ":" & xBB & ":" & (yBB - yAA)
        Else
            la.width = xBB - xAA
            la.height = yBB
            strTag = strTag & ":" & (xBB - xAA) & ":" & yBB
        End If
        la.Tag = strTag
    End If
    funGetLine = True
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub subMoveImage(newX As Long, newY As Long, baseX As Long, baseY As Long, ViewerIndex As Integer)
'------------------------------------------------
'���ܣ�����λ����״λ��ʸ״λ����ק��꣬�����ƶ����л�ͼ��
'������
'       newX -- ��λ�õ�ͼ������X����
'       newY -- ��λ�õ�ͼ������Y����
'       basex -- ��λ�õ�ͼ������X����
'       baseY -- ��λ�õ�ͼ������Y����
'       ViewerIndex -- ͼ������Viewer��Index��1-��λͼ��2-��״λͼ��3-ʸ״λͼ
'���أ���
'------------------------------------------------
    Dim la As DicomLabel
    Dim newImgX As Long, newImgY As Long, baseImgX As Long, baseImgY As Long
    Dim Size As Long
    Dim img As DicomImage
    
    '��ȡͼ��
    If ViewerIndex = 1 Then
        '��λ����������ƶ���ͼ���൱�ڹ�״λ�к��������ƶ�
        Set img = mCoronalViewer.Images(1)
    ElseIf ViewerIndex = 2 Then
        '��״λ����������ƶ���ͼ���൱��ʸ״λ���������ƶ�
        Set img = mSagittalViewer.Images(1)
    ElseIf ViewerIndex = 3 Then
        'ʸ״λ��������·�ͼ���൱�ڹ�״λ���������ƶ�
        Set img = mCoronalViewer.Images(1)
    End If
    
    If ViewerIndex = 1 Then
        Set la = img.Labels(G_INT_SYS_LABEL_MPR_RESULT_H)
        Size = img.sizeY
        
        baseImgX = la.left
        baseImgY = la.top
        newImgX = baseImgX
        newImgY = (newY - baseY) * Size / Viewer(ViewerIndex).Images(1).sizeY + baseImgY
        
    ElseIf ViewerIndex = 2 Or ViewerIndex = 3 Then
        Set la = img.Labels(G_INT_SYS_LABEL_MPR_RESULT_V)
        Size = img.sizeX
            
        baseImgX = la.left
        baseImgY = la.height / 2
        newImgY = baseImgY
        newImgX = Size / Viewer(ViewerIndex).Images(1).sizeY * (newY - baseY) + baseImgX
    End If
    
    Call subMoveCAndSLabel(la, img, newImgX, newImgY, baseImgX, baseImgY, IIf(ViewerIndex = 2, False, True))
    Call img.Refresh(False)
    
    On Error GoTo err
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
