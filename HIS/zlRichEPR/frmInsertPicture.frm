VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmInsertPicture 
   Caption         =   "����ͼƬ..."
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frmInsertPicture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   10830
   StartUpPosition =   3  '����ȱʡ
   Begin zlRichEPR.ucCanvas Canvas 
      Height          =   2175
      Left            =   2655
      TabIndex        =   5
      Top             =   3420
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3836
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1980
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   $"frmInsertPicture.frx":058A
   End
   Begin VB.PictureBox picFinal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   9405
      ScaleHeight     =   915
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   7065
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   45
      ScaleHeight     =   5910
      ScaleWidth      =   2265
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1395
      Width           =   2265
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   225
         TabIndex        =   2
         Top             =   135
         Width           =   1725
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4755
         Left            =   45
         TabIndex        =   3
         Top             =   900
         Width           =   2055
         _cx             =   3625
         _cy             =   8387
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInsertPicture.frx":0647
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Shape shpThis 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   0
         Top             =   810
         Width           =   330
      End
      Begin VB.Shape shpSearch 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   90
         Top             =   45
         Width           =   330
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8115
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14420
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   3015
      Top             =   270
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmInsertPicture.frx":06B4
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   90
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmInsertPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ID_Marked = 301                       '���ͼ
Private Const ID_Out = 302                          '�ⲿͼ
Private Const ID_InsAndExit = 303                   '���벢�˳�
Private Const ID_CancelAndExit = 304                'ȡ�����˳�
Private Const ID_Prev = 305                         'ǰһ��
Private Const ID_Next = 306                         '��һ��
Private Const ID_FitSize = 307                      '�ʺϳߴ�
Private Const ID_ActualSize = 308                   'ʵ�ʳߴ�
Private Const ID_Slider = 309                       '�õ�Ƭ
Private Const ID_ZoomIn = 310                       '�Ŵ�
Private Const ID_ZoomOut = 311                      '��С
Private Const ID_Deasil = 312                       '˳ʱ��
Private Const ID_AntiClockWise = 313                '��ʱ��
Private Const ID_Edit = 314                         '�༭
Private Const ID_Help = 315                         '����
Private Const ID_Scan = 316                         'ɨ����

Private mintType As Integer               'ͼƬ����
'-- λͼ����
Private WithEvents DIBFilter As cDIBFilter      ' DIB �˾�����(24 bpp)
Attribute DIBFilter.VB_VarHelpID = -1
Private WithEvents DIBDither As cDIBDither      ' DIB ��������(1, 4, 8 bpp)
Attribute DIBDither.VB_VarHelpID = -1
Private DIBPal               As New cDIBPal     ' DIB ��ɫ����� (1, 4, 8 bpp)
Private DIBSave              As New cDIBSave    ' Save ���� (BMP)  (1, 4, 8, 24 bpp)
Private DIBbpp               As Byte            ' ��ǰ��ɫ���
Private WithEvents cPicEditor As cPictureEditor     ' ͼƬ�༭����
Attribute cPicEditor.VB_VarHelpID = -1
Private m_LastFilename As String                    ' ���򿪵�ͼƬλ��
Private m_Temp As String                            ' ��ʱ�ļ�·��
Private m_AppID As Long
'-- GDI+
Private m_GDIpToken         As Long         ' ���ڹر� GDI+
Private m_MaxWidth As Long                  '����ȣ��������Ʋ���ͼƬ������ȣ�
Private m_MaxHeight As Long                 '���߶ȣ��������Ʋ���ͼƬ�����߶ȣ�

Private mfrmParent As Object                          '������


'################################################################################################################
'## ���ܣ�  ��ʾ���༭����
'##
'## ������  Parent                  :������
'##         blnIsMarkedPicture      :�Ƿ��Ǳ��ͼ
'##         lRow1,lCol1,lRow2,lCol2 :�����ͼƬ����Ե�Ԫλ��
'##         lngMaxWidth             :����ȣ��������Ʋ���ͼƬ������ȣ�
'##         lngMaxHeight            :���߶ȣ��������Ʋ���ͼƬ�����߶ȣ�
'################################################################################################################
Public Sub ShowMe(Parent As Object, Optional lngMaxWidth As Long = 0, Optional lngMaxHeight As Long = 0)
    Set mfrmParent = Parent
    m_MaxWidth = lngMaxWidth
    m_MaxHeight = lngMaxHeight
 
    Me.Show vbModal, Parent
End Sub

'################################################################################################################
'## ���ܣ�  ��ʾID=lID�ı��ͼ��ͼ���ݡ�
'################################################################################################################
Private Sub ShowPicture(strKey As String)
    mintType = EPRMarkedPicture
    stbThis.Panels(1).Text = "�����ͼ��"
    
    Dim strTemp As String, bSuccess As Boolean
    
    Screen.MousePointer = vbHourglass
    stbThis.Panels(2).Text = ""
    Set Canvas.DIB = New cDIB
    strTemp = zlBlobRead(0, strKey)
    If Len(strTemp) > 0 Then
        Call pvSetDIBPicture(pvGetStdPicture(strTemp, bSuccess))
        If bSuccess Then stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
        Kill strTemp
    End If
    Canvas.Resize
    Screen.MousePointer = vbNormal
End Sub

Private Sub Canvas_Crop()
    '-- Change to True color mode
    DIBbpp = 24
    Call pvSetPalMode(DIBbpp)
    
    '-- Update Info and Progress
    With Canvas.DIB
        stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
    End With
End Sub

Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = vbRightButton Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = Me.CommandBars.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set Control = .Add(xtpControlButton, ID_FitSize, "�ʺϳߴ�(&F)"): Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, ID_ActualSize, "ʵ�ʳߴ�(&S)")
            Set Control = .Add(xtpControlButton, ID_ZoomIn, "�Ŵ�(&Z)")
            Set Control = .Add(xtpControlButton, ID_ZoomOut, "��С(&O)")
            Popup.ShowPopup
        End With
    End If
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��������ť����¼�
    Dim strFileName As String, bSuccess As Boolean
    Select Case Control.ID
    Case ID_CancelAndExit
        Unload Me
    Case ID_Marked
        If picLeft.Visible = False Then
            picLeft.Visible = True
        Else
            picLeft.Visible = False
        End If
        mintType = EPRMarkedPicture
        stbThis.Panels(1).Text = "�����ͼ��"
        CommandBars_Resize
    Case ID_Out
        dlgThis.Filename = m_LastFilename
        dlgThis.CancelError = True
        On Error GoTo LL
        dlgThis.ShowOpen
        
        strFileName = dlgThis.Filename
        If Trim(strFileName) <> "" Then
            '-- Create DIB
'            DoEvents
            Screen.MousePointer = vbHourglass
            Call pvSetDIBPicture(pvGetStdPicture(strFileName, bSuccess))
            Screen.MousePointer = vbNormal
            
            If (bSuccess) Then
                m_LastFilename = strFileName
                mintType = EPROutPicture
                stbThis.Panels(1).Text = "���ⲿͼ��"
                stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
                stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
            End If
        End If
    Case ID_Scan
        'ɨ����
        Dim intReturn As Integer                        '���巵��ֵ
        intReturn = TWAIN_SelectImageSource(Me.hWnd)    '��ɨ����ѡ���Ի���
        
        If intReturn = 0 Then Exit Sub                  '�������ֵ�� 0 ˵���û����� ȡ��������ֵ�� 1 ����ʾ�����ɹ�
        
        intReturn = TWAIN_AcquireToClipboard(Me.hWnd, 0) ' 0 ��ɨ����ָ��ɫ��ģʽ
        'ɨ��ͼƬ������ɨ����ɨ���ͼƬ����������� ������Ĳ�����ֻɨ�����ɫģʽ
        
        If intReturn = 0 Then Exit Sub                  '����ֵ 0 ��ʾɨ��ʧ�ܣ� 1 ��ʾ�ɹ�
        
        Call pvSetDIBPicture(Clipboard.GetData(vbCFDIB))      '���������е�ͼƬ �浽 image ��
        mintType = EPROutPicture
        stbThis.Panels(1).Text = "���ⲿͼ��"
        stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
        stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
        
    Case ID_InsAndExit
        If Canvas.DIB.hDIB = 0 Then
            Exit Sub
        End If
        
        LockWindowUpdate Me.hWnd
        Call FitMaxSize     '����ͼƬ�ߴ磡����

        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '100%��ͼƬ����
        
        If gobjFSO.FileExists(strFileName) Then
            Set picFinal.Picture = LoadPicture(strFileName)
            gobjFSO.DeleteFile strFileName, True
            
            mfrmParent.InsertPicture mintType, picFinal.Picture, picFinal.Width, picFinal.Height
            Unload Me
        Else
            MsgBox "����ʧ�ܣ�", vbOKOnly + vbInformation, "����ͼƬ"
        End If
    Case ID_Prev
    Case ID_Next
    Case ID_FitSize
        DoZoomMenu 3
    Case ID_ActualSize
        DoZoomMenu 2
    Case ID_Slider
    Case ID_ZoomIn
        DoZoomMenu 0
    Case ID_ZoomOut
        DoZoomMenu 1
    Case ID_Deasil
        '-- Rotate +90��
        Screen.MousePointer = vbArrowHourglass
        Call Canvas.DIB.Orientation(True, True, True)
        Call Canvas.Resize
        Screen.MousePointer = vbDefault
    Case ID_AntiClockWise
        '-- Rotate +90��
        Screen.MousePointer = vbArrowHourglass
        Call Canvas.DIB.Orientation(True, False, False)
        Call Canvas.Resize
        Screen.MousePointer = vbDefault
'    Case ID_Edit
'        strFileName = m_Temp & "\R" & m_AppID & ".jpg"
'        Call mGdIpEx.SaveDIB(Me.Canvas.DIB, strFileName, [ImageJPEG], 90)         '100%��ͼƬ����
'        If gobjFSO.FileExists(strFileName) Then
'            Set picFinal.Picture = LoadPicture(strFileName)
'            gobjFSO.DeleteFile strFileName, True
'            cPicEditor.ShowPicEditor glngSys, gcnOracle, picFinal.Picture, , , Me
'        End If
    Case ID_Help
        ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
    End Select
LL:
End Sub

Private Sub DoZoomMenu(Index As Integer)
    Select Case Index
        Case 0 '-- Zoom +
            Canvas.Zoom = Canvas.Zoom + IIf(Canvas.Zoom < 25, 1, 0)
            Canvas.FitMode = False
        Case 1 '-- Zoom -
            Canvas.Zoom = Canvas.Zoom - IIf(Canvas.Zoom > 1, 1, 0)
            Canvas.FitMode = False
        Case 2 '-- 1 : 1
            Canvas.Zoom = 1
            Canvas.FitMode = False
        Case 3 '-- Best fit
            Canvas.FitMode = True
    End Select
    Call Canvas.Resize
    stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
End Sub

Private Sub CommandBars_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub CommandBars_Resize()
    'λ�õ���
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    CommandBars.GetClientRect Left, Top, Right, Bottom
    
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    If Right >= Left And Bottom >= Top Then
        If picLeft.Visible Then
            picLeft.Move Left + lX * 2, Top + lY * 2, picLeft.Width, (Bottom - Top) - lY * 4
            Canvas.Move picLeft.Left + picLeft.Width + lX * 2, picLeft.Top, (Right - Left) - picLeft.Width - lX * 4, picLeft.Height
        Else
            Canvas.Move Left, Top, (Right - Left), (Bottom - Top)
        End If
    End If
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    '��������ť�����¼�
    Select Case Control.ID
    Case ID_Marked
        If picLeft.Visible Then
            Control.Checked = True
        Else
            Control.Checked = False
        End If
    Case ID_Out
    Case ID_Prev
        Control.Enabled = False
    Case ID_Next
        Control.Enabled = False
    Case ID_FitSize
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
        Control.Checked = (Canvas.FitMode = True)
    Case ID_ActualSize
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_Slider
        Control.Enabled = False
    Case ID_ZoomIn
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_ZoomOut
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_Deasil
        Control.Enabled = False
    Case ID_AntiClockWise
        Control.Enabled = False
    Case ID_Edit
        Control.Enabled = (Canvas.DIB.hDIB <> 0)
    Case ID_Help
        
    End Select
End Sub

Private Sub picLeft_Resize()
    Dim lX As Long, lY As Long
    lX = Screen.TwipsPerPixelX
    lY = Screen.TwipsPerPixelY
    With picLeft
        txtSearch.Move lX, lY, .ScaleWidth - lX * 2
        shpSearch.Move txtSearch.Left - lX, txtSearch.Top - lY, txtSearch.Width + lX * 2, txtSearch.Height + lY * 2
        vfgThis.Move txtSearch.Left, shpSearch.Top + shpSearch.Height + lY, txtSearch.Width, .ScaleHeight - shpSearch.Height - lY * 2
        shpThis.Move vfgThis.Left - lX, vfgThis.Top - lY, vfgThis.Width + 2 * lX, vfgThis.Height + 2 * lY
    End With
End Sub

Private Sub txtSearch_GotFocus()
    zlControl.TxtSelAll txtSearch
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim i As Long
        For i = 1 To vfgThis.Rows - 1
            If UCase(vfgThis.Cell(flexcpText, i, 1)) Like UCase(Trim(txtSearch)) & "*" Or UCase(vfgThis.Cell(flexcpText, i, 2)) Like UCase(Trim(txtSearch)) & "*" Then
                vfgThis.Row = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub vfgThis_RowColChange()
    If vfgThis.Row = 0 Then Exit Sub
    ShowPicture vfgThis.Cell(flexcpText, vfgThis.Row, 0)
End Sub

Private Sub Form_Load()
Dim Bar���� As CommandBar                       '�������ؼ�
    '##########################################################################################
    '����λ�ûָ�
    Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Left", (Screen.Width - 10000) / 2)
    Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Top", (Screen.Width - 8000) / 2)
    Me.Width = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Width", 10000)
    Me.Height = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Height", 8000)
    m_LastFilename = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", App.Path)
    
    Dim GpInput As GdiplusStartupInput
    '-- ���� GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        Call MsgBox("���� GDI+ �����޷�����ͼƬ���룡���� GDI+ DLL �Ƿ���ڻ����𻵣�", vbInformation + vbOKOnly)
        Call Unload(Me)
        Exit Sub
    End If
    
    stbThis.Panels(2).Text = "׼������"
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    m_AppID = Me.hWnd
    Set DIBFilter = New cDIBFilter
    Set DIBDither = New cDIBDither
    Set cPicEditor = New cPictureEditor
    
    '##########################################################################################
    '## ��������ʼ��
    '##########################################################################################
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    'ͼ���
    CommandBars.Icons = ImageManager.Icons
    CommandBars.ActiveMenuBar.Visible = False
    
    Dim objControl As CommandBarControl                 '�������ؼ�
    Set Bar���� = CommandBars.Add("����", xtpBarTop)
    With Bar����.Controls
        Set objControl = .Add(xtpControlButton, ID_Marked, "���ͼ(&M)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_Out, "�ⲿͼ(&W)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_Scan, "����ɨ���ǻ������ Ctl+O")
        objControl.BeginGroup = True
        
'        Set objControl = .Add(xtpControlButton, ID_Prev, "��һ��ͼ��(���ͷ)")
'        objControl.BeginGroup = True
'        .Add xtpControlButton, ID_Next, "��һ��ͼ��(�Ҽ�ͷ)"
        Set objControl = .Add(xtpControlButton, ID_FitSize, "���ʺ� Ctl+B")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ActualSize, "ʵ�ʴ�С Ctl+A"
        .Add xtpControlButton, ID_Slider, "��ʼ�õ�Ƭ F11"
        Set objControl = .Add(xtpControlButton, ID_ZoomIn, "�Ŵ�(+)")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_ZoomOut, "��С(-)"
'        Set objControl = .Add(xtpControlButton, ID_Deasil, "˳ʱ����ת Ctl+K")
'        objControl.BeginGroup = True
'        .Add xtpControlButton, ID_AntiClockWise, "��ʱ����ת Ctl+L"
'        Set objControl = .Add(xtpControlButton, ID_Edit, "��λͼ�༭���༭��ͼƬ Ctl+E")
'        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_Help, "���� F1")
        objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, ID_InsAndExit, "����(&S)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, ID_CancelAndExit, "�ر�(&Q)")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
    End With

    '��ʾ��չ��ť
    CommandBars.Options.ShowExpandButtonAlways = False
    CommandBars.EnableCustomization (False)
    CommandBars.Options.UseDisabledIcons = True
    '�Ƿ���ʾ���в˵�
    CommandBars.Options.AlwaysShowFullMenus = True
   
    '##########################################################################################
    '�ȼ���
    CommandBars.KeyBindings.Add FCONTROL, vbKeyReturn, ID_InsAndExit
    CommandBars.KeyBindings.Add FCONTROL, vbKeyM, ID_Marked
    CommandBars.KeyBindings.Add FCONTROL, vbKeyW, ID_Out
    CommandBars.KeyBindings.Add FCONTROL, vbKeyS, ID_InsAndExit
    CommandBars.KeyBindings.Add FCONTROL, vbKeyQ, ID_CancelAndExit
    CommandBars.KeyBindings.Add FCONTROL, vbKeyReturn, ID_InsAndExit
    CommandBars.KeyBindings.Add 0, VK_ESCAPE, ID_CancelAndExit
    
    CommandBars.KeyBindings.Add 0, vbKeyLeft, ID_Prev
    CommandBars.KeyBindings.Add 0, vbKeyRight, ID_Next
    CommandBars.KeyBindings.Add FCONTROL, vbKeyB, ID_FitSize
    CommandBars.KeyBindings.Add FCONTROL, vbKeyA, ID_ActualSize
    CommandBars.KeyBindings.Add 0, vbKeyF11, ID_Slider
    CommandBars.KeyBindings.Add 0, vbKeyAdd, ID_ZoomIn
    CommandBars.KeyBindings.Add 0, vbKeySubtract, ID_ZoomOut
    CommandBars.KeyBindings.Add FCONTROL, vbKeyK, ID_Deasil
    CommandBars.KeyBindings.Add FCONTROL, vbKeyL, ID_AntiClockWise
    CommandBars.KeyBindings.Add FCONTROL, vbKeyE, ID_Edit
    CommandBars.KeyBindings.Add 0, vbKeyF1, ID_Help
    CommandBars.KeyBindings.Add FCONTROL, vbKeyO, ID_Scan
    '##########################################################################################
    
    Call FillGrid
    
    Me.KeyPreview = True
    vfgThis.Editable = flexEDNone
End Sub

'################################################################################################################
'## ���ܣ�  �����ͼͼƬ�б�
'################################################################################################################
Private Sub FillGrid()
Dim rsTemp As ADODB.Recordset                '��¼��
    gstrSQL = "select ����,����,���� from �������ͼ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "")
    vfgThis.Clear
    If Not rsTemp.EOF Then vfgThis.Rows = rsTemp.RecordCount + 1
    Dim i As Long
    i = 0
    vfgThis.Cell(flexcpText, 0, 0) = "����"
    vfgThis.Cell(flexcpText, 0, 1) = "����"
    vfgThis.Cell(flexcpText, 0, 2) = "����"
    vfgThis.ColAlignment(1) = flexAlignLeftCenter
    Do While Not rsTemp.EOF
        i = i + 1
        vfgThis.Cell(flexcpText, i, 0) = NVL(rsTemp("����"))
        vfgThis.Cell(flexcpText, i, 1) = NVL(rsTemp("����"))
        vfgThis.Cell(flexcpText, i, 2) = NVL(rsTemp("����"))
        rsTemp.MoveNext
    Loop
    rsTemp.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Canvas.DIB.Destroy
    LockWindowUpdate 0
    UpdateWindow Me.hWnd
    ' Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)

    '-- Free objects
    Set DIBFilter = Nothing
    Set DIBDither = Nothing
    Set DIBPal = Nothing
    Set DIBSave = Nothing
    Set cPicEditor = Nothing
    
    '���洰��λ��
    If Me.WindowState <> vbMinimized Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Left", Me.Left
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Top", Me.Top
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Width", Me.Width
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "Height", Me.Height
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "LastFilename", m_LastFilename
    End If
End Sub

'################################################################################################################
'## ���ܣ�  �������
'################################################################################################################
Private Sub txtSearch_Change()
End Sub

'################################################################################################################
'## ���ܣ�  ������غ���
'################################################################################################################
Private Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean) As StdPicture
    On Error Resume Next
    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
        '-- Use GDI+ loading
        Set pvGetStdPicture = mGdIpEx.LoadPictureEx(sFileName)
      Else
        '-- Use VB LoadPicture
        Set pvGetStdPicture = LoadPicture(sFileName)
    End If
    
    '-- Is there an image ?
    bSuccess = Not (pvGetStdPicture Is Nothing)
    
    If (bSuccess = False) Then
        '-- Nothing loaded
        Call MsgBox("����ͼƬʱ�����������", vbExclamation)
    End If

    On Error GoTo 0
End Function
    
Private Sub FitMaxSize()
    '����ͼƬ�ߴ������Χ��ҳ��ߴ磩����
    Dim dblH As Double, dblW As Double
    Dim lngWW As Long, lngHH As Long        '������С
    
    If m_MaxHeight = 0 And m_MaxWidth = 0 Then Exit Sub
    
    If Canvas.DIB.Height > m_MaxHeight Or Canvas.DIB.Width > m_MaxWidth Then
        If m_MaxHeight = 0 Then
            dblH = 1
        Else
            dblH = CDbl(Canvas.DIB.Height) / CDbl(m_MaxHeight)
        End If
        
        If m_MaxWidth = 0 Then
            dblW = 1
        Else
            dblW = CDbl(Canvas.DIB.Width) / CDbl(m_MaxWidth)
        End If
        
        If dblH > dblW Then
            lngWW = m_MaxHeight * (Canvas.DIB.Width / Canvas.DIB.Height)
            lngHH = m_MaxHeight
        Else
            lngWW = m_MaxWidth
            lngHH = m_MaxWidth * (Canvas.DIB.Height / Canvas.DIB.Width)
        End If
        If (lngWW <> Canvas.DIB.Width) Or (lngHH <> Canvas.DIB.Height) Then
'            DoEvents
            Screen.MousePointer = vbHourglass
            '-- Resize/Resample
            Call mGdIpEx.ScaleDIB(Canvas.DIB, lngWW, lngHH, True)
            '-- Remove Crop rectangle and resize canvas
            Call Canvas.RemoveCropRectangle
            Call Canvas.Resize
            '-- Update size info
            With Canvas.DIB
                stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
            End With
            Screen.MousePointer = vbNormal
        End If
    End If
End Sub

Private Sub pvSetDIBPicture(Image As StdPicture)
  Static lstW As Long
  Static lstH As Long

    If (Not Picture Is Nothing) Then

        '-- Save last DIB dimensions
        lstW = Canvas.DIB.Width
        lstH = Canvas.DIB.Height
        
        '-- Clear palette
        Call DIBPal.Clear
        
        '-- Create 32bpp DIB section from std. picture.
        '   Case source <=8bpp, palette saved in DIBPal, palette indexes in DIBDither.
        '   Return value: source color depth / 0 = Err.
        DIBbpp = Canvas.DIB.CreateFromStdPicture(Image, DIBPal, DIBDither)
        
        '-- Select current depth mode
        Call pvSetPalMode(DIBbpp)
        
        '-- Remove Crop rectangle and resize canvas
        Call Canvas.RemoveCropRectangle
        With Canvas.DIB
            If (lstW <> .Width Or lstH <> .Height) Then
                Call Canvas.Resize
              Else
                Call Canvas.Repaint
            End If
        End With
        
        '-- Show image info: Size + bpp
'        stbThis.Panels(2).Text = "ͼƬ��С��" & Canvas.DIB.Width & "��" & Canvas.DIB.Height & "��" & DIBbpp & "λ(Bpp)"
        stbThis.Panels(3).Text = Format(Canvas.Zoom, "0%")
    End If
End Sub

Private Sub pvSetPalMode(ByVal bpp As Long)
  Dim lIdxNew As Long
  Dim lIdxOld As Long
    
    Select Case bpp
        Case 1  '-- 2 colors / Black and White
            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
        Case 4  '-- 16 colors / 16 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
        Case 8  '-- 256 colors / 256 greys
            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
        Case 24 '-- True color
            lIdxNew = 8
        Case Else
            Exit Sub
    End Select
End Sub

Private Function pvGetExt(ByVal sFileName As String) As String
    pvGetExt = Mid(sFileName, Len(sFileName) - 2)
End Function
