VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form FrmMagnify 
   Caption         =   "�Ŵ�"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   Icon            =   "FrmMagnify.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkInvert 
      Caption         =   "����"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   3360
      Width           =   660
   End
   Begin VB.CheckBox chkOrganLens 
      Caption         =   "͸��"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3360
      Width           =   660
   End
   Begin VB.CommandButton CmdHid 
      Caption         =   "����"
      Height          =   345
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Width           =   585
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      Min             =   1
      Max             =   80
      SelStart        =   20
      TickStyle       =   3
      Value           =   20
   End
   Begin DicomObjects.DicomViewer Viewer1 
      Height          =   3165
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3855
      _Version        =   262147
      _ExtentX        =   6800
      _ExtentY        =   5583
      _StockProps     =   35
      BackColor       =   -2147483641
      AsyncReceive    =   -1  'True
      UseScrollBars   =   0   'False
   End
   Begin VB.Label lblZoomState 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3135
   End
End
Attribute VB_Name = "FrmMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public f As Form

Dim MP As POINTAPI
'���ڱ��游�������Ҫ�Ŵ��DICOM�ؼ�
Dim FrmClick As Boolean
Dim NowImg As Integer               '��ǰͼ��
Dim IntBC As Integer                '���ò���
Dim BeginWidth, BeginHeight As Integer  '��ǰ����λ��
Dim blnOrganLens As Boolean     '�Ƿ��������֯͸��״̬
Dim blnInvert As Boolean        '�Ƿ���뷴��״̬
Dim intMaxTop As Long           '��������TOP
Dim intMaxLeft As Long          '��������Left
Dim lngBaseXX As Long
Dim lngBaseYY As Long
Dim lngMagnifyWidth As Long
Dim lngMagnifyLevel As Long


Private Sub chkInvert_Click()
    blnInvert = IIf(chkInvert.Value = 1, True, False)
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Call subFlipRotate(Viewer1.Images(1), "Invert")
End Sub

Private Sub chkOrganLens_Click()
    blnOrganLens = IIf(chkOrganLens.Value = 1, True, False)
    If blnOrganLens Then
        subOrganLens
    Else
        If Me.Viewer1.Images.Count > 0 Then
            Me.Viewer1.Images(1).SetDefaultWindows
        End If
    End If
End Sub

'�˳�
Private Sub CmdHid_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '���ر�����
    Call zlcontrol.FormSetCaption(Me, False)
        
    Me.left = f.left + f.width / 2
    Me.top = f.top + f.height / 2
    intMaxTop = GetToolBarBottomOrRight(2)
    intMaxLeft = GetToolBarBottomOrRight(1)
End Sub
'����Ӧ����
Private Sub Form_Resize()
    On Error Resume Next
    'viewer1
    Me.Viewer1.top = 1
    Me.Viewer1.left = 1
    Me.Viewer1.width = Me.ScaleWidth - 2
    Me.Viewer1.height = Me.ScaleHeight - Me.Slider1.height - Me.lblZoomState.height - 5
    'lblZoomState
    Me.lblZoomState.top = Me.Viewer1.height + 2
    Me.lblZoomState.left = 1
    Me.lblZoomState.width = Me.ScaleWidth - Me.CmdHid.width - 1
    If Me.Viewer1.Images.Count > 0 Then
        Me.lblZoomState.Caption = "   �Ŵ�����" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
    Else
        Me.lblZoomState.Caption = "   �Ŵ�����"
    End If
    'slider1
    Me.Slider1.top = Me.lblZoomState.top + Me.lblZoomState.height + 2
    Me.Slider1.left = 1
    Me.Slider1.width = Abs(Me.ScaleWidth - Me.CmdHid.width - Me.chkOrganLens.width - Me.chkInvert.width - 1)
    'chkOrganLens
    Me.chkOrganLens.top = Me.Slider1.top
    Me.chkOrganLens.left = Me.Slider1.left + Me.Slider1.width
    'chkInvert
    Me.chkInvert.top = Me.Slider1.top
    Me.chkInvert.left = Me.chkOrganLens.left + Me.chkOrganLens.width
    'cmd
    Me.CmdHid.top = Me.Slider1.top - 2
    Me.CmdHid.left = Me.chkInvert.left + Me.chkInvert.width
    'ˢ��
    ImgMagnify
End Sub

'�ı����ű���
Private Sub Slider1_Change()
    On Error Resume Next
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Me.CmdHid.SetFocus
    ImgMagnify
    Me.lblZoomState.Caption = "   �Ŵ�����" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
End Sub

Private Sub Viewer1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    If Button = cMouseUsage("102").lngMouseKey Then
        FrmClick = True
        lngBaseXX = x
        lngBaseYY = y
    Else
        '���Կ�ʼ����
        FrmClick = True
        BeginWidth = x
        BeginHeight = y
    End If
End Sub

Private Sub Viewer1_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    '�����ƶ�����
    Dim TmpMp As POINTAPI
    Dim Mx, My As Integer
    On Error Resume Next
    If FrmClick = True Then
        If Button = cMouseUsage("102").lngMouseKey Then
            If Abs(y - lngBaseYY) >= lngWidthLevelStep / 5 Or Abs(x - lngBaseXX) >= lngWidthLevelStep / 5 Then  ''''������������
                Me.Viewer1.Images(1).width = Me.Viewer1.Images(1).width + (x - lngBaseXX) * lngWidthLevelStep / 5
                Me.Viewer1.Images(1).Level = Me.Viewer1.Images(1).Level + (y - lngBaseYY) * lngWidthLevelStep / 5
                lngMagnifyWidth = Me.Viewer1.Images(1).width
                lngMagnifyLevel = Me.Viewer1.Images(1).Level
                Me.Viewer1.Refresh
                lngBaseXX = x
                lngBaseYY = y
            End If
        Else
            GetCursorPos TmpMp
            Mx = (TmpMp.x * Screen.TwipsPerPixelX) - (BeginWidth * Screen.TwipsPerPixelX)
            My = (TmpMp.y * Screen.TwipsPerPixelY) - (BeginHeight * Screen.TwipsPerPixelY)
            Me.Move Mx, My
            ImgMagnify
            Me.lblZoomState.Caption = "   �Ŵ�����" & Format(Me.Viewer1.Images(1).ActualZoom, "###0.0000")
            'ˢ�·Ŵ���
            f.Refresh
        End If
    End If
End Sub
Private Sub Viewer1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    FrmClick = False
End Sub
'�Ŵ�
Public Sub ImgMagnify()
    '****************************************************************************
    '������MainFrm ȡ�ø��������
    '      DicView ȡ��DICOM����
    '���ã���DICOMͼ��Ŵ�
    '****************************************************************************
    Dim Ox1, Ox2, Oy1, Oy2 As Integer   '���ڵõ��������ڿؼ�����Ļ������
    Dim Ix1, Ix2, Iy1, Iy2 As Integer   '���ڵõ��Ŵ�������Ļ������
    Dim Dx, Dy As Integer               '�Ŵ����DICOM�ؼ������ĵ�����
    Dim HowImg As Integer               '�õ���ǰ�ǵڼ���ͼ
    Dim Sx1, Sx2, Sy1, Sy2 As Integer   'Сͼ�������
    Dim intRow, intCol As Integer       '������к��и���
    Dim MWidth, MHeight As Integer      '��С���ͼ���͸�
    Dim i As Integer                    '��ʱ����
    Dim ViewIndex As Integer
    Dim a As Double
'    On Error Resume Next
    'Dim A As New DicomImage
    ViewIndex = 0
   '�Ŵ���λ��
    With Me
        Ix1 = (.left / Screen.TwipsPerPixelX) + .Viewer1.left - intMaxLeft
        Ix2 = (.left / Screen.TwipsPerPixelX) + .Viewer1.left + .Viewer1.width - intMaxLeft
        Iy1 = (.top / Screen.TwipsPerPixelY) + .Viewer1.top - GetMenuHeight - intMaxTop - GetSystemMetrics(11) '+ 6
        Iy2 = (.top / Screen.TwipsPerPixelY) + .Viewer1.top + .Viewer1.height - intMaxTop - GetMenuHeight - GetSystemMetrics(11) '+ 6
    End With
    '�Ŵ������ĵ�
    Dx = (Ix2 - Ix1) / 2 + Ix1
    Dy = (Iy2 - Iy1) / 2 + Iy1
    
    ViewIndex = FunIsViewer(Dx * Screen.TwipsPerPixelX, Dy * Screen.TwipsPerPixelY)
    
    
    '�õ���������DIDOM�ؼ�λ��
    With f
        Ox1 = (.left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).left / Screen.TwipsPerPixelX)
        Ox2 = (.left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).left / Screen.TwipsPerPixelX) + (.Viewer(ViewIndex).width / Screen.TwipsPerPixelX)
        Oy1 = (.top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).top / Screen.TwipsPerPixelY)
        Oy2 = (.top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).top / Screen.TwipsPerPixelY) + (.Viewer(ViewIndex).height / Screen.TwipsPerPixelY)
    End With
    '����ͼ��
    With Me.Viewer1
        HowImg = f.Viewer(ViewIndex).ImageIndex(Dx - Ox1, Dy - Oy1)
        '����ͼ�������λ��(�������Ͽ������Ĳ���)
        Ox1 = Ox1 - f.Viewer(ViewIndex).Images(HowImg).ScrollX
        Oy1 = Oy1 - f.Viewer(ViewIndex).Images(HowImg).ScrollY
        If HowImg < 1 Then
            .Images.Clear
            Exit Sub
        End If
        '�ǵ�ǰͼ��ʱ��ˢ��ͼ��
            .Images.Clear
            .Images.Add f.Viewer(ViewIndex).Images(HowImg)
            .Images(1).MagnificationMode = doFilterBSpline
            .Images(1).Zoom = f.Viewer(ViewIndex).Images(HowImg).ActualZoom * (Me.Slider1.Value / 10)
            .Images(1).StretchToFit = False
            .Images(1).Labels.Clear
            If .Images(1).VOILUT = 1 Then
                .Images(1).width = f.Viewer(ViewIndex).Images(HowImg).width
                .Images(1).Level = f.Viewer(ViewIndex).Images(HowImg).Level
                .Images(1).VOILUT = 0
            End If
            If lngMagnifyWidth <> 0 And lngMagnifyLevel <> 0 Then
                .Images(1).width = lngMagnifyWidth
                .Images(1).Level = lngMagnifyLevel
            End If
    End With
    With f.Viewer(ViewIndex)
        '��ͼ�񳬹���ǰ��ʾͼ������ʱ����
        If (.MultiColumns * .MultiRows) < .Images.Count Then
            i = HowImg - f.Viewer(ViewIndex).CurrentIndex + 1
        Else
            i = HowImg
        End If
        '�õ���ǰͼ��λ��
        If (i Mod .MultiColumns) = 0 Then
            intRow = i / .MultiColumns
            intCol = .MultiColumns
        Else
            intRow = Int(i / .MultiColumns) + 1
            intCol = HowImg Mod .MultiColumns
            If intCol = 0 Then
                intCol = 1
            End If
        End If
        MWidth = (.width / .MultiColumns) / Screen.TwipsPerPixelX
        MHeight = (.height / .MultiRows) / Screen.TwipsPerPixelY
    End With
    '�Ŵ�Сͼ��λ����
    If intCol = 1 Then
        Sx1 = 0
        Sx2 = MWidth
    Else
        Sx1 = MWidth * (intCol - 1)
        Sx2 = MWidth * intCol
    End If
    If intRow = 1 Then
        Sy1 = 0
        Sy2 = MHeight
    Else
        Sy1 = MHeight * (intRow - 1)
        Sy2 = MHeight * intRow
    End If

    '����Ŵ���λ��
    With Viewer1
        If Dx > Ox1 And Dx < Ox2 And Dy > Oy1 And Dy < Oy2 Then
'            .Images(1).ScrollX = ((Dx - Ox1 - Sx1) * (Slider1.Value / 10)) - Abs(f.viewer(ViewIndex).Images(HowImg).ActualScrollX * (Slider1.Value / 10)) - (.width / 2)
'            .Images(1).ScrollY = ((Dy - Oy1 - Sy1) * (Slider1.Value / 10)) - Abs(f.viewer(ViewIndex).Images(HowImg).ActualScrollY * (Slider1.Value / 10)) - (.height / 2)
            .Images(1).ScrollX = (((Dx - Ox1 + Abs(f.Viewer(ViewIndex).Images(HowImg).ScrollX) - Sx1) - Abs(f.Viewer(ViewIndex).Images(HowImg).ActualScrollX)) / f.Viewer(ViewIndex).Images(HowImg).ActualZoom * .Images(1).ActualZoom) - (.width / 2)
            .Images(1).ScrollY = (((Dy - Oy1 + Abs(f.Viewer(ViewIndex).Images(HowImg).ScrollY) - Sy1) - Abs(f.Viewer(ViewIndex).Images(HowImg).ActualScrollY)) / f.Viewer(ViewIndex).Images(HowImg).ActualZoom * .Images(1).ActualZoom) - (.height / 2)
        Else
            Me.Viewer1.Images.Clear
        End If
    End With
    
    '�˾�
    If blnOrganLens Then subOrganLens
    '����
    If blnInvert Then Call subFlipRotate(Viewer1.Images(1), "Invert")
    
    '����һ��ˢ�£���ͼ��ӳ�ٶȼӿ졣
    Me.Viewer1.Refresh
End Sub

'�õ���ǰViewer����λ��
Function FunIsViewer(x As Long, y As Long) As Integer
    Dim v As DicomViewer

    With f
        FunIsViewer = 0
        For Each v In .Viewer
            If v.Visible And .left + v.left <= x And _
            .left + v.left + v.width >= x And _
            .top + v.top <= y And _
            .top + v.top + v.height >= y Then
                FunIsViewer = v.Index
                Exit Function
            End If
        Next
    End With
End Function

Private Sub subOrganLens()
    If Me.Viewer1.Images.Count = 0 Then Exit Sub
    Dim lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long
    Dim ww As Long, wl As Long
    
    '������֯͸���ĸ�������
    lngLeft = Me.Viewer1.Images(1).ActualScrollX / Me.Viewer1.Images(1).ActualZoom
    lngTop = Me.Viewer1.Images(1).ActualScrollY / Me.Viewer1.Images(1).ActualZoom
    lngWidth = Me.Viewer1.width / Me.Viewer1.Images(1).ActualZoom
    lngHeight = Me.Viewer1.height / Me.Viewer1.Images(1).ActualZoom
    
    '������֯͸�����´���λ����������Ӧ�����㷨
    If funAutoWinWL(Me.Viewer1.Images(1), lngLeft, lngTop, lngWidth, lngHeight, ww, wl) Then
        Me.Viewer1.Images(1).width = ww
        Me.Viewer1.Images(1).Level = wl
    End If
End Sub

Private Function GetToolBarBottomOrRight(LeftORTop As Integer) As Long
    '------------------------------------------------
    '���ܣ�                                  �õ���������Left��Right�ĸ�
    '������                                  LeftORRight 1=left  2= Right
    '���أ�                                  �������ĸ�
    '�ϼ���������̣�                        ImgMagnify
    '�¼���������̣�                        ��
    '���õ��ⲿ������                        f�����洰��
    '�����ˣ�                                ���� 2005-8-9
    '------------------------------------------------
    Dim intMaxTop  As Long                  '���ĸ�
    Dim intMaxLeft As Long                  '���ı�
    Dim intToolBarLeft  As Long             '������Left
    Dim intToolBarTop   As Long             '������Top
    Dim intToolBarRight As Long             '������Right
    Dim intToolBarBottom As Long            '������Bottom
    Dim a As CommandBar
    Dim i As Integer
    
    With f.ComToolBar
        If LeftORTop = 1 Then
            For i = 2 To 8
                If .Item(i).Position = xtpBarLeft Then
                    .Item(i).GetWindowRect intToolBarLeft, intToolBarTop, intToolBarRight, intToolBarBottom
                    If intMaxLeft < intToolBarLeft Or intMaxLeft = 0 Then
                        intMaxLeft = intToolBarLeft
                        GetToolBarBottomOrRight = GetToolBarBottomOrRight + (intToolBarRight - intToolBarLeft)
                    End If
                End If
            Next
            GetToolBarBottomOrRight = GetToolBarBottomOrRight / Screen.TwipsPerPixelX
        Else
            For i = 2 To 8
                If .Item(i).Position = xtpBarTop Then
                    .Item(i).GetWindowRect intToolBarLeft, intToolBarTop, intToolBarRight, intToolBarBottom
                     If intMaxTop < intToolBarTop Or intMaxTop = 0 Then
                        intMaxTop = intToolBarTop
                        GetToolBarBottomOrRight = GetToolBarBottomOrRight + (intToolBarBottom - intToolBarTop)
                    End If
                End If
            Next
            GetToolBarBottomOrRight = GetToolBarBottomOrRight / Screen.TwipsPerPixelY
        End If
    End With
End Function

Private Function GetMenuHeight() As Long
    '------------------------------------------------
    '���ܣ�                                  �õ��˵�������ͼ�ĸ߶�
    '������
    '���أ�                                  �˵�������ͼ�ĸ߶�
    '------------------------------------------------
    Dim lngToolBarLeft  As Long             '������Left
    Dim lngToolBarTop   As Long             '������Top
    Dim lngToolBarRight As Long             '������Right
    Dim lngToolBarBottom As Long            '������Bottom
    
    f.ComToolBar.Item(ToolBar_Menu).GetWindowRect lngToolBarLeft, lngToolBarTop, lngToolBarRight, lngToolBarBottom
    GetMenuHeight = (lngToolBarBottom - lngToolBarTop) / Screen.TwipsPerPixelY
    
    '�������ͼ����Ϊͣ����������ʾ������ͼ�����������ͼ�ĸ߶�
    If blnDockMiniImage = True And frmMiniSeries.Visible = True Then
        GetMenuHeight = GetMenuHeight + frmMiniSeries.height / Screen.TwipsPerPixelY
    End If
End Function








