VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form frmFilmView 
   Caption         =   "��Ƭ��ӡ--�鿴ͼ��"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7995
   Icon            =   "frmFilmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7995
   Begin DicomObjects.DicomViewer dcmViewer 
      Height          =   6000
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _Version        =   262147
      _ExtentX        =   12938
      _ExtentY        =   10583
      _StockProps     =   35
      BackColor       =   0
      UseScrollBars   =   0   'False
   End
End
Attribute VB_Name = "frmFilmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event AfterClose(dcmImage As DicomImage, intViewerIndex As Integer, intImageIndex As Integer)

'�����ڲ���������
Public SelectedImage As DicomImage      '��ǰ��ѡ�е�ͼ��
Public blnDefaultWW2 As Boolean         '��¼˫����λ��״̬

'�����ڲ�˽�б���
Private mfrmParent As frmFilm
Private mintMouseState As Integer
Private mdcmSelectLabel As DicomLabel   '��ǰ��ѡ�еı�ע
Private mblnDcmViewDown As Boolean      '�����ж�dcmView������Ƿ񱻰���
Private mblnLabelMoving As Boolean      '�����ƶ��ü���
Private intBaseX As Long                '��¼���ԭ����Xλ��
Private intBaseY As Long                '��¼���ԭ����Yλ��
Private mintSourceViewerIndex As Integer    '��¼��ǰ�����ͼ�����ڵ�Viewer����
Private mintSourceImageIndex As Integer     '��¼��ǰ�����ͼ�����ڵ�ͼ������
Private mdblViewerRatio As Double       '��¼Viewer�߶�/��ȵı���
Private mblnAfterShow As Boolean        '��ʾ���
Private mintDriverType As Integer       '��¼��ǰ�豸���ͣ�������ȡ��Ӧ�ĵ�����ݼ�

''''''''''''''''�ü�''''''''''''''''''''''''''''''''''''''''''''''
Private mintCutOutLabel                 '�ü������ڵı�ע���
''''''''''''''''�ü�''''''''''''''''''''''''''''''''''''''''''''''

Private Sub dcmViewer_DblClick()
    '��ͼ����˫��ʱ������ͼ��ü�
    If mintCutOutLabel = 0 Then Exit Sub
    If dcmViewer.Images.Count = 0 Then Exit Sub
    If mintCutOutLabel <> dcmViewer.Images(1).Labels.Count Then Exit Sub
    
    Dim Image As DicomImage
    Dim i As Integer
    Dim lblTemp As DicomLabel
    Dim sourceImage As DicomImage
    
    Set sourceImage = dcmViewer.Images(1)
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
    dcmViewer.Images.Clear
    dcmViewer.Images.Add Image
    
    'ͼ�����Viewer�к�������ʾ��ߣ����ʱ���ߺ͵�λ����׼ȷ��
    Call UpdateRuler(Image, True)
    
    mintCutOutLabel = 0
    Me.MousePointer = vbArrow
End Sub

Private Sub dcmViewer_KeyDown(KeyCode As Integer, Shift As Integer)
    '������λ��ݰ�ť��F2,F3-F12
    If KeyCode >= VK_F2 And KeyCode <= VK_F12 Then
        Call subWWWLShortCut(KeyCode)
    End If
End Sub

Private Sub dcmViewer_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    '����ͼ�����
    Dim ls As DicomLabels

    If dcmViewer.Images.Count = 0 Then Exit Sub
    intBaseX = x
    intBaseY = y

    If Button = 1 Then
        mintMouseState = mfrmParent.intMouseState

        'mintMouseState ����״̬��0���ޣ�1��������2�����Σ�3������;4-ѡ��ͼ��;5-��ѡ����;6-�ü�:7-���ֱ�ע
        If mintMouseState = 6 Then  '�ü�
            '�ü�״̬�µ����down�������ֲ�����1�����ü��򣨼�¼��ǣ���2���ƶ��ü���(�н���) ��3��˫�����вü�
            If mintCutOutLabel = 0 Then  '���ü���
                '���ӿ�ѡ��ע
                dcmViewer.Images(1).Labels.Add GetNewLabel(doLabelRectangle, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
                Set mdcmSelectLabel = dcmViewer.Images(1).Labels(dcmViewer.Images(1).Labels.Count)
                mdcmSelectLabel.Tag = CUT_LABEL
                mblnDcmViewDown = True
                mintCutOutLabel = dcmViewer.Images(1).Labels.Count
            Else    '��ʼ�ƶ��ü���
                Set ls = dcmViewer.LabelHits(x, y, False, False, True)
                If ls.Count <> 0 And Me.MousePointer <> vbArrow Then
                    '��ʼ�ƶ��ü���
                    If ls(1).Tag = CUT_LABEL Then
                        mblnLabelMoving = True
                    End If
                End If
            End If
        End If

        If mintMouseState = 5 Then  '��ѡ����
            '���ӿ�ѡ��ע
            dcmViewer.Images(1).Labels.Add GetNewLabel(doLabelRectangle, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
            Set mdcmSelectLabel = dcmViewer.Images(1).Labels(dcmViewer.Images(1).Labels.Count)
            mblnDcmViewDown = True
        End If

        If mintMouseState = 7 Then  '���ֱ�ע
            Dim dcmLabel As DicomLabel
            Set dcmLabel = GetNewLabel(doLabelText, dcmViewer.ImageXPosition(x, y), dcmViewer.ImageYPosition(x, y), 0, 0)
            dcmViewer.Images(1).Labels.Add dcmLabel
            dcmLabel.AutoSize = True
            dcmLabel.Margin = 0
            dcmLabel.Text = mfrmParent.pstrSideMarker
            dcmLabel.Shadow = doShadowAll
            dcmLabel.ShowTextBox = True
            dcmLabel.Font.Bold = True
            dcmLabel.Tag = POSTURE_LABEL
            mintMouseState = 0
            '���ø��������������
            mfrmParent.pstrSideMarker = ""
            mfrmParent.intMouseState = 0
        End If

        dcmViewer.Refresh
    End If
End Sub

Private Sub dcmViewer_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '��������ƶ��¼�
    Dim dblZoom As Double
    Dim ls As DicomLabels
    Dim lngXOffset As Long
    Dim lngYOffset As Long
    Dim lblCUT As DicomLabel
    
    On Error GoTo err
    
    If dcmViewer.Images.Count = 0 Then Exit Sub
    
    If (Button = 1 And mintMouseState = 1) Or (Button = 4 And intMouseWheelDrag = 2) _
        Or (Button = 2 And cMouseUsage("102").lngMouseKey = 2 And Button_miWidthLevel) Then  '����
        If SelectedImage.VOILUT = 1 Then SelectedImage.VOILUT = 0
        SelectedImage.width = SelectedImage.width + (x - intBaseX) * lngWidthLevelStep / 5
        SelectedImage.Level = SelectedImage.Level + (y - intBaseY) * lngWidthLevelStep / 5
        SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
        intBaseX = x
        intBaseY = y
        dcmViewer.Refresh
    ElseIf (Button = 1 And mintMouseState = 2) Or (Button = 4 And intMouseWheelDrag = 0) _
        Or (Button = 2 And cMouseUsage("103").lngMouseKey = 2 And Button_miCruise) Then '����
        subCenterZoom SelectedImage, dcmViewer, SelectedImage.ActualZoom
        SelectedImage.ScrollX = SelectedImage.ScrollX - (x - intBaseX) * lngCruiseStep / 5
        SelectedImage.ScrollY = SelectedImage.ScrollY - (y - intBaseY) * lngCruiseStep / 5
        intBaseX = x
        intBaseY = y
    ElseIf (Button = 1 And mintMouseState = 3) Or (Button = 4 And intMouseWheelDrag = 1) _
        Or (Button = 2 And cMouseUsage("104").lngMouseKey = 2 And Button_miZoom) Then '����
        '���ŵ�λ��0.01��
        dblZoom = SelectedImage.ActualZoom * (1 + (intBaseY - y) * lngZoomStep / 5 * 0.001)
        If dblZoom < 0.01 Then dblZoom = 0.01
        If dblZoom > 64 Then dblZoom = 64
        Call subCenterZoom(SelectedImage, dcmViewer, dblZoom)
        
        If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
            If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '���±�ߵ�λ
                Call UpdateRuler(SelectedImage, True)
            End If
        End If
        
        intBaseX = x
        intBaseY = y
    ElseIf Button = 1 And (mintMouseState = 5 Or mintMouseState = 6) Then  '��ѡ���źͲü�
        If mblnDcmViewDown = True Then
            mdcmSelectLabel.width = dcmViewer.ImageXPosition(x, y) - mdcmSelectLabel.left
            mdcmSelectLabel.height = dcmViewer.ImageYPosition(x, y) - mdcmSelectLabel.top
            dcmViewer.Refresh
        End If
    End If
    
    '��������ü�
    If mintMouseState = 6 And mintCutOutLabel <> 0 Then
        Set ls = dcmViewer.LabelHits(x, y, False, False, True)
        If Button = 1 Then  '��갴��
            If mblnLabelMoving = True Then
                Call subaCorrectCursor(dcmViewer, SelectedImage, x, y)
                Set lblCUT = SelectedImage.Labels(SelectedImage.Labels.Count)
                
                If (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then    '�����ƶ�
                    
                    lngXOffset = (dcmViewer.ImageXPosition(x, y) - dcmViewer.ImageXPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.left - dcmViewer.ImageXPosition(x, y)) > Abs(lblCUT.left + lblCUT.width - dcmViewer.ImageXPosition(x, y)) Then '�ұߵ��ƶ�
                        lblCUT.width = lblCUT.width + lngXOffset
                    Else    '��ߵ��ƶ�
                        lblCUT.left = lblCUT.left + lngXOffset
                        lblCUT.width = lblCUT.width - lngXOffset
                    End If
                ElseIf (Me.MousePointer = vbSizeNS And (SelectedImage.RotateState = doRotateNormal Or SelectedImage.RotateState = doRotate180)) _
                    Or (Me.MousePointer = vbSizeWE And (SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight)) Then    '�����ƶ�
                    
                    lngYOffset = (dcmViewer.ImageYPosition(x, y) - dcmViewer.ImageYPosition(intBaseX, intBaseY))
                    If Abs(lblCUT.top - dcmViewer.ImageYPosition(x, y)) > Abs(lblCUT.top + lblCUT.height - dcmViewer.ImageYPosition(x, y)) Then '�����ߵ��ƶ�
                        lblCUT.height = lblCUT.height + lngYOffset
                    Else    '�����ߵ��ƶ�
                        lblCUT.top = lblCUT.top + lngYOffset
                        lblCUT.height = lblCUT.height - lngYOffset
                    End If
                ElseIf Me.MousePointer = vbSizePointer Then     '�����ƶ�
                    lngXOffset = (dcmViewer.ImageXPosition(x, y) - dcmViewer.ImageXPosition(intBaseX, intBaseY))
                    lngYOffset = (dcmViewer.ImageYPosition(x, y) - dcmViewer.ImageYPosition(intBaseX, intBaseY))
                    lblCUT.top = lblCUT.top + lngYOffset
                    lblCUT.left = lblCUT.left + lngXOffset
                End If
                intBaseX = x
                intBaseY = y
                dcmViewer.Refresh
            End If
        ElseIf Button = 0 Then    '���û�б����£�ֻ�ı����ָ��
            If ls.Count <> 0 Then
                If Abs(ls(1).left - dcmViewer.ImageXPosition(x, y)) < 4 Or Abs(ls(1).left + ls(1).width - dcmViewer.ImageXPosition(x, y)) < 4 Then
                    If SelectedImage.RotateState = doRotateLeft Or SelectedImage.RotateState = doRotateRight Then
                        Me.MousePointer = vbSizeNS
                    Else
                        Me.MousePointer = vbSizeWE
                    End If
                ElseIf Abs(ls(1).top - dcmViewer.ImageYPosition(x, y)) < 4 Or Abs(ls(1).top + ls(1).height - dcmViewer.ImageYPosition(x, y)) < 4 Then
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

Private Sub dcmViewer_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    '������굯���¼�
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    
    On Error GoTo err
    
    If Button = 1 Then
        If mintMouseState <> 0 Then
            If mintMouseState = 5 And mblnDcmViewDown Then    '��ѡ����
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
                
                RectangleZoom dcmViewer, SelectedImage, lngLeft, lngTop, lngWidth, lngHeight
                
                'ɾ����ѡ�õ���ʱ��ע
                SelectedImage.Labels.Remove SelectedImage.Labels.Count
                Set mdcmSelectLabel = Nothing
                dcmViewer.Refresh
            ElseIf mintMouseState = 6 Then
                If mblnDcmViewDown Then       '�ü�
                    '�����κβ���
                    '����ü���Ϊ0 ����ȡɾ���ü�������ü��ı��
                    If mdcmSelectLabel.width = 0 Or mdcmSelectLabel.height = 0 Then
                        'ɾ����ѡ�õ���ʱ��ע
                        SelectedImage.Labels.Remove SelectedImage.Labels.Count
                        Set mdcmSelectLabel = Nothing
                        dcmViewer.Refresh
                        
                        mintCutOutLabel = 0
                    End If
                End If
            End If
        End If
    End If
    mblnDcmViewDown = False
    mblnLabelMoving = False
    Exit Sub
err:
End Sub

Private Sub Form_Load()
    '��ȡ����λ��
    Call RestoreWinState(Me, App.ProductName)
    
    'ÿ�δ򿪴��ڣ�������Ĭ��ֵ
    mintDriverType = 0
    blnDefaultWW2 = False
End Sub

Private Sub Form_Resize()
    Dim lngOldWidth As Long
    Dim lngOldHeight As Long
    
    lngOldWidth = dcmViewer.width
    lngOldHeight = dcmViewer.height
    
    dcmViewer.left = 0
    dcmViewer.top = 0
    dcmViewer.width = Me.ScaleWidth
    dcmViewer.height = Me.ScaleHeight
    
    '����ͼ���λ�ú����ű���
    If mblnAfterShow And Not SelectedImage Is Nothing Then
        If SelectedImage.StretchToFit = False Then
            Call subScaleImage(SelectedImage, dcmViewer, lngOldWidth, lngOldHeight)
        End If
    End If
End Sub


Public Sub zlShowMe(img As DicomImage, frmParent As frmFilm, intViewerIndex As Integer, intImageIndex As Integer)
'------------------------------------------------
'���ܣ���ͼ������
'������ img - ��Ҫ�������ʾ��ͼ��
'       frmParent - ��Ƭ��ӡԤ������
'       intViewerIndex - ��ǰ��ͼ�����ڵ�Viewer����
'       intImageIndex -- ��ǰ��ͼ�����ڵ�ͼ������
'���أ���
'------------------------------------------------
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    On Error GoTo err
    
    If img Is Nothing Then Exit Sub
    
    mblnAfterShow = False
    
    Set mfrmParent = frmParent
    mintSourceViewerIndex = intViewerIndex
    mintSourceImageIndex = intImageIndex
    
    lngWidth = frmParent.FilmViewer(mintSourceViewerIndex).width / frmParent.FilmViewer(mintSourceViewerIndex).MultiColumns
    lngHeight = frmParent.FilmViewer(mintSourceViewerIndex).height / frmParent.FilmViewer(mintSourceViewerIndex).MultiRows
    
    mdblViewerRatio = lngHeight / lngWidth
    
    dcmViewer.Images.Clear
    dcmViewer.Images.Add img
    Set SelectedImage = dcmViewer.Images(1)
    
    '����ͼ���ע����ʾ
    If SelectedImage.Labels.Count > 0 Then
        Call subChangeLabelForPrint(SelectedImage, 1)
    End If
    
    Me.height = mdblViewerRatio * Abs(Me.width - 115) + 510 '���ϱ���߶� 510,��Ե���115
    If Me.height < mdblViewerRatio * Abs(Me.width - 115) + 510 Then
        '�߶ȳ����ˣ�ʹ�ø߶ȼ��㿴���
        Me.width = Abs(Me.height - 510) / mdblViewerRatio + 115
    End If
    
    '����ͼ���λ�ú����ű���
    If img.StretchToFit = False Then
        Call subScaleImage(SelectedImage, dcmViewer, lngWidth, lngHeight)
    End If
    
    Me.Show , mfrmParent
    mblnAfterShow = True
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '    ж��hook
    Call FilmViewUnhook(Me.hwnd, plngFilmViewPreWndProc)
    
    '���ص�ǰ�����ͼ��
    If dcmViewer.Images.Count = 1 Then
        '�ָ�ͼ���ע����ʾ
        If dcmViewer.Images(1).Labels.Count > 0 Then
            Call subChangeLabelForPrint(dcmViewer.Images(1), 0)
        End If
        RaiseEvent AfterClose(dcmViewer.Images(1), mintSourceViewerIndex, mintSourceImageIndex)
    End If
    
    '���洰��λ��
    Call SaveWinState(Me, App.ProductName)
End Sub

Public Sub ZLToolButtonClick(control As CommandBarControl)
'------------------------------------------------
'���ܣ�����ƬԤ������Ĺ�������ť�¼�
'������ lngControlID -- ������ID
'���أ�ֱ���޸�ͼ��
'------------------------------------------------

    On Error GoTo err
    If SelectedImage Is Nothing Then Exit Sub
    
    '''''''''''''''''''''''''''''[���ܼ����ô���λ����]''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If control.Id >= 349 And control.Id <= 360 Then
        subFunctionWL control, Me
        Exit Sub
    End If
    
    Select Case control.Id
        Case ID_frmFilm_FilterLengthUp      'ƽ������
            Call SubImageFiltering("miFilterLengthUp", SelectedImage)
        Case ID_frmFilm_FilterLengthDown     ''ƽ������
            Call SubImageFiltering("miFilterLengthDown", SelectedImage)
        Case ID_frmFilm_Invert               ''����
            Call subFlipRotate(SelectedImage, "Invert")
        Case ID_frmFilm_RotateLeft           ''������ת90��
            Call subFlipRotate(SelectedImage, "RotateAnticlockwise")
        Case ID_frmFilm_RotateRight          ''������ת90��
            Call subFlipRotate(SelectedImage, "RotateClockwise")
        Case ID_frmFilm_FlipHorizontal       ''���Ҿ���
            Call subFlipRotate(SelectedImage, "FlipHorizontal")
        Case ID_frmFilm_FlipVertical         ''���¾���
            Call subFlipRotate(SelectedImage, "FlipVertical")
        Case ID_frmFilm_Resume               ''�ָ�
            SelectedImage.SetDefaultWindows
            SelectedImage.FlipState = doFlipNormal
            SelectedImage.RotateState = doRotateNormal
            SelectedImage.StretchToFit = True
            SelectedImage.UnsharpEnhancement = 0
            SelectedImage.UnsharpLength = 0
            SelectedImage.FilterLength = 0
            
            If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
                If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then  '���±�ߵ�λ
                    UpdateRuler SelectedImage, True
                End If
            End If
    End Select
    
    dcmViewer.Refresh
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub subWWWLShortCut(KeyCode As Integer)
'------------------------------------------------
'���ܣ�������λ��ݼ�
'������ KeyCode -- ��ʱ���µĿ�ݼ�
'���أ�ֱ���޸�ͼ��
'------------------------------------------------
    Dim intWidth As Integer
    Dim intLevel As Integer
    Dim strDriverType As String
    Dim i As Integer
    
    On Error GoTo err
    
    If KeyCode < VK_F2 Or KeyCode > VK_F12 Then Exit Sub
    If SelectedImage Is Nothing Then Exit Sub
    
    If KeyCode = VK_F2 Then 'Ĭ�ϴ���
        SelectedImage.VOILUT = 1
        '�ж��Ƿ�������Ĭ�ϴ���
        If blnDefaultWW2 = False Then
            '��ʾ�ڶ�������
            If SelectedImage.Attributes(&H28, &H1050).VM = 2 And SelectedImage.Attributes(&H28, &H1051).VM = 2 Then
                intWidth = SelectedImage.Attributes(&H28, &H1051).ValueByIndex(2)
                intLevel = SelectedImage.Attributes(&H28, &H1050).ValueByIndex(2)
                SelectedImage.width = intWidth
                SelectedImage.Level = intLevel
                blnDefaultWW2 = True
            Else
                SelectedImage.SetDefaultWindows
            End If
        Else
            SelectedImage.SetDefaultWindows
            blnDefaultWW2 = False
        End If
        
        If SelectedImage.Attributes(&H6000, &H15).Value = 1 Then
            If SelectedImage.Level = 0 Then SelectedImage.Level = 1
        End If
    Else    'Ԥ�贰��
        '���ж��Ƿ���Ҫ��ȡͼ�����
        If mintDriverType = 0 Then
            If IsNull(SelectedImage.Attributes(&H8, &H60).Value) Then Exit Sub         '��ȡModality
            strDriverType = SelectedImage.Attributes(&H8, &H60).Value
            
            For i = 1 To UBound(aPresetWinWL, 2)        '[�ҵ�ͼ���Ӧ�豸]
                If UCase(aPresetWinWL(3, i).strModality) = UCase(strDriverType) Then
                    mintDriverType = i
                    Exit For
                End If
            Next i
        End If
        
        If mintDriverType > 0 Then
            If aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).bInUse Then
                SelectedImage.width = aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).lngWinWidth
                SelectedImage.Level = aPresetWinWL(KeyCode - VK_F2 + 2, mintDriverType).lngWinLevel
            End If
        End If
    End If
    
    SelectedImage.Labels(G_INT_SYS_LABEL_WWWL).Text = "W:" & SelectedImage.width & "-L:" & SelectedImage.Level
    
    dcmViewer.Refresh
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    If dcmViewer.Images.Count <= 0 Then Exit Sub
    
    If intDirection = 1 Then
        dblScale = 1 + (lngZoomStep * 0.01)
    Else
        dblScale = 1 - (lngZoomStep * 0.01)
    End If
    
    Call subCenterZoom(SelectedImage, dcmViewer, SelectedImage.ActualZoom * dblScale)
        
    If SelectedImage.Labels.Count >= G_INT_SYS_LABEL_RULLER Then
        If SelectedImage.Labels(G_INT_SYS_LABEL_RULLER).Visible = True Then '���±�ߵ�λ
            Call UpdateRuler(SelectedImage, True)
        End If
    End If
    
End Sub
