VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucPictureEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ScaleHeight     =   4725
   ScaleWidth      =   8490
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   105
      ScaleHeight     =   4605
      ScaleWidth      =   5340
      TabIndex        =   0
      Top             =   30
      Width           =   5340
      Begin VB.PictureBox picBuff 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   3645
         ScaleHeight     =   660
         ScaleWidth      =   930
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   90
         ScaleHeight     =   187
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   231
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   3465
         Begin VB.PictureBox picTxt 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   315
            MousePointer    =   1  'Arrow
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   9
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "�ƶ���˫����������"
            Top             =   135
            Visible         =   0   'False
            Width           =   165
         End
         Begin VB.TextBox txtTmp 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   1530
            MultiLine       =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "������ǰ���������"
            Top             =   135
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   225
            MaxLength       =   250
            MouseIcon       =   "ucPictureEditor.ctx":0000
            MousePointer    =   99  'Custom
            MultiLine       =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   195
            Visible         =   0   'False
            Width           =   180
         End
      End
   End
   Begin zlRichEPR.ColorPicker ColorForeColor 
      Height          =   2190
      Left            =   6195
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   390
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      Color           =   0
   End
   Begin zlRichEPR.ColorPicker ColorLineColor 
      Height          =   2190
      Left            =   5925
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      Color           =   13209
      AutoColor       =   255
   End
   Begin zlRichEPR.ColorPicker ColorFillColor 
      Height          =   2190
      Left            =   5655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -15
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3863
      Color           =   16711680
      AutoColor       =   16711680
   End
   Begin MSComctlLib.ImageList imgCur 
      Left            =   6360
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0152
            Key             =   "Pen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":02B4
            Key             =   "Move"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":05CE
            Key             =   "Earse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":08E8
            Key             =   "Text"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0A4A
            Key             =   "Sel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFill 
      Left            =   7785
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0D64
            Key             =   "FILLSTYLE"
            Object.Tag             =   "565"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0DD0
            Key             =   "FILLNONE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0E3C
            Key             =   "FILLALL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0EA6
            Key             =   "FILLH"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0F18
            Key             =   "FILLV"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0F89
            Key             =   "FILLHV"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":0FFC
            Key             =   "FILLL"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":1077
            Key             =   "FILLR"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":10F0
            Key             =   "FILLLR"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7020
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":116F
            Key             =   "FILLCOLOR"
            Object.Tag             =   "562"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":12D9
            Key             =   "LINECOLOR"
            Object.Tag             =   "563"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucPictureEditor.ctx":1432
            Key             =   "FORECOLOR"
            Object.Tag             =   "564"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ucPictureEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'����������
Private mBar��ͼ As CommandBar
Private mBar���� As CommandBarPopup, mBar�߿� As CommandBarPopup, mBar�����ʽ As CommandBarPopup   '��ͼ�������Ӳ˵�

Private mfrmParent As frmMain
Private PicMarks As New cPicMarks           '�ֲ��ı�Ƕ�����ʱ���󣬿��Բ����棻���򣬱��浽�������mcPicture.PicMarks��
Private mobjTable As cEPRTable

Private mblnInDrawing As Boolean            '�Ƿ��ڻ�ͼģʽ
Private mlngSelFillColor As OLE_COLOR       '����ѡ������ɫֵ
Private mlngSelLineColor As OLE_COLOR       '����ѡ������ɫֵ
Private mlngSelForeColor As OLE_COLOR       '����ѡ������ɫֵ
Private mlngDrawModeID As Long              '��ǰ��ͼģʽ
Private mlngFillColor As Long, mlngLineColor As Long, mlngFillStyleID As Long, mlngLineWidthID As Long, mlngLineStyleID As Long
Private mvarOldPoint As POINTAPI, mvarFirstPoint As POINTAPI
Private mlngSelectedCount As Long

Private mvarPolyPoints() As POINTAPI
Private mblnModified As Boolean
Private mblnDblClick As Boolean             '�Ƿ�˫�� ����Σ�����߽�����������
Private mlngOrgX As Long, mlngOrgY As Long  '��ʼ��������
Private mblnEditInTable As Boolean          '�Ƿ����ڱ���б༭
Private mblnOK As Boolean

Private WithEvents cbsThis As CommandBars
Attribute cbsThis.VB_VarHelpID = -1
Private mblnInited As Boolean
Private zoomFactor As Double                '�������ӣ���ʾ�ߴ�/ʵ�ʳߴ磩

Public mcPicture As New cEPRPicture        'ԴͼƬ����
Public Event DblClick()

Public Property Let Modified(vData As Boolean)
    mblnModified = vData
End Property

Public Property Get Modified() As Boolean
    Modified = mblnModified
End Property

Public Property Get Inited() As Boolean
    Inited = mblnInited
End Property


'################################################################################################################
'   ��;��  ϵͳ��ڡ�
'################################################################################################################
Public Sub ShowMe(ByVal frmParent As frmMain, ByVal hWndParent As Long, ByVal cbsMain As CommandBars, ByVal cPicture As cEPRPicture, _
    ByVal lngLeft As Long, ByVal lngTop As Long, ByVal lngWidth As Long, ByVal lngHeight As Long, _
    Optional bEditInTable As Boolean = False, Optional objTable As cEPRTable = Nothing)
    
    Set cbsThis = cbsMain
    Set mobjTable = objTable
    Set mfrmParent = frmParent
    Set mcPicture = cPicture
    If cbsThis.FindControl(, ID_DRAW_SELECT) Is Nothing Then 'û�в˵����ؽ�
        mblnInited = True
        Call InitCommandBars
    End If
    If mcPicture.PictureType <> EPRFormulaPicture Then
        mBar��ͼ.Visible = True
        mBar����.Visible = True
        mBar�߿�.Visible = True
        mBar�����ʽ.Visible = True
        If bEditInTable Then
            UserControl.BorderStyle = 1
            UserControl.Appearance = 0
        Else
            UserControl.BorderStyle = 0
            UserControl.Appearance = 0
        End If
        
        '��������
        zoomFactor = CDbl(lngHeight) / CDbl(cPicture.OrigHeight)
            
        '##########################################################################################
        '�����Ļָ�
        ColorFillColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillColor", vbBlue)
        ColorLineColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineColor", vbRed)
        ColorForeColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "ForeColor", vbBlack)
        mlngFillStyleID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillStyle", ID_DRAW_FILLNONE)
        mlngLineStyleID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineStyle", ID_DRAW_LINECONTINUE)
        mlngLineWidthID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineWidth", ID_DRAW_LINEWIDTH1)
        mlngDrawModeID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "DrawMode", ID_DRAW_SELECT)
    
        mBar��ͼ.FindControl(, mlngDrawModeID).Checked = True
        mBar�����ʽ.CommandBar.FindControl(, mlngFillStyleID).Checked = True
        mBar�߿�.CommandBar.FindControl(, mlngLineWidthID).Checked = True
        mBar����.CommandBar.FindControl(, mlngLineStyleID).Checked = True
        '##########################################################################################
    End If
    SetParent UserControl.hwnd, hWndParent
    If objTable Is Nothing Then
        Set PicMarks = mcPicture.PicMarks
    Else
        If objTable.TableType = tte_����ͼƬ�� Then
            Set PicMarks = mcPicture.PicMarks
        Else
            Set PicMarks = ScalePicMarks(mcPicture.PicMarks, zoomFactor)
        End If
    End If
    mblnEditInTable = bEditInTable
    UserControl.Extender.Left = lngLeft
    UserControl.Extender.Top = lngTop
    UserControl.Extender.Width = lngWidth
    If mblnEditInTable Then
        lngHeight = (lngWidth * cPicture.Height) / cPicture.Width
        '������Զ�����ģʽ��ͼƬ����
        UserControl.Extender.Height = lngHeight + (mobjTable.GridLineWidth - 1) * 15
    Else
        UserControl.Extender.Height = lngHeight
    End If
    UserControl_Resize
    
    '����ͼƬ������ʾ
    picDraw.Width = UserControl.Extender.Width
    picDraw.Height = UserControl.Extender.Height
    picDraw.PaintPicture mcPicture.OrigPic, 0, 0, picDraw.Width / 15, picDraw.Height / 15
    
    '��ȡ��ǰ��ͼģʽ��Ϣ
    If mcPicture.PictureType <> EPRFormulaPicture Then
        If Not mblnInDrawing Then Call GetCurDrawMode
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
    End If
    UserControl.Extender.Visible = True
    picDraw.Move 0, 0
End Sub

Public Sub CloseMe(Optional pic As cEPRPicture = Nothing)
    If mcPicture.PictureType <> EPRFormulaPicture Then
        Call SaveState
        If pic Is Nothing Then
            If mblnModified Then Call SavePicture
        Else
            If mblnModified Then
                If mobjTable Is Nothing Then
                    Set pic.PicMarks = PicMarks
                Else
                    If mobjTable.TableType = tte_����ͼƬ�� Then
                        Set pic.PicMarks = PicMarks
                    Else
                        Set pic.PicMarks = ScalePicMarks(PicMarks, 1# / zoomFactor)
                    End If
                End If
            End If
        End If
         
        mBar����.Visible = False
        mBar�߿�.Visible = False
        mBar�����ʽ.Visible = False
        mBar��ͼ.Visible = False
    End If
    UserControl.Extender.Visible = False
End Sub

'################################################################################################################
'   ��;��  ��̬���¹���������ɫ��ͼ�ꡣ
'################################################################################################################
Private Sub SetColorIcon(Key As String, ID As Long, COLOR As OLE_COLOR)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Set ListImage = imgColor.ListImages(Key)
    
    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgColor.MaskColor
    
    ctlPictureBox.Picture = ListImage.ExtractIcon
    
    If COLOR = vbWhite Then COLOR = RGB(254, 254, 254)
    ctlPictureBox.Line (1, ctlPictureBox.Height * 0.6)-(ctlPictureBox.Width, ctlPictureBox.Height), COLOR, BF
    ctlPictureBox.Refresh

    'Replace icon
    imgColor.ListImages.Remove imgColor.ListImages(Key).Index
    imgColor.ListImages.Add 1, Key, ctlPictureBox.Image
'    Set imgColor.ListImages(Key).Picture = ctlPictureBox.Image

    'OK Now replace Tag property
    imgColor.ListImages(1).Tag = ID
    
    cbsThis.AddImageList imgColor

    cbsThis.RecalcLayout
    
    UserControl.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

'################################################################################################################
'   ��;��  ���������ʽͼ�ꡣ
'################################################################################################################
Private Sub SetFillIcon(ID As Long)
    Dim ctlPictureBox As VB.PictureBox
    Set ctlPictureBox = Controls.Add("VB.PictureBox", "ctlPictureBox1")
    Dim ListImage As ListImage
    Dim Key As String
    Select Case ID
    Case ID_DRAW_FILLNONE
        Key = "FILLNONE"
    Case ID_DRAW_FILLALL
        Key = "FILLALL"
    Case ID_DRAW_FILLH
        Key = "FILLH"
    Case ID_DRAW_FILLV
        Key = "FILLV"
    Case ID_DRAW_FILLHV
        Key = "FILLHV"
    Case ID_DRAW_FILLR
        Key = "FILLR"
    Case ID_DRAW_FILLL
        Key = "FILLL"
    Case ID_DRAW_FILLLR
        Key = "FILLLR"
    End Select
    Set ListImage = imgFill.ListImages(Key)
    
    ctlPictureBox.AutoRedraw = True
    ctlPictureBox.AutoSize = True
    ctlPictureBox.BackColor = imgFill.MaskColor
    ctlPictureBox.Picture = ListImage.ExtractIcon
    
    'Replace icon
    imgFill.ListImages.Remove imgFill.ListImages("FILLSTYLE").Index
    imgFill.ListImages.Add 1, "FILLSTYLE", ctlPictureBox.Image
    
    'OK Now replace Tag property
    imgFill.ListImages(1).Tag = ID_DRAW_FILLSTYLE
    
    cbsThis.AddImageList imgFill
    
    cbsThis.RecalcLayout
    
    UserControl.Controls.Remove ctlPictureBox
    Set ctlPictureBox = Nothing
End Sub

Private Sub picDraw_DblClick()
    If mcPicture.PictureType = EPRFormulaPicture Then
        RaiseEvent DblClick
        Exit Sub
    Else
        mblnDblClick = True
    End If
End Sub

Private Sub picTxt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mlngOrgX = x: mlngOrgY = y
End Sub

Private Sub picTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If txt.Left + x - mlngOrgX >= 0 And txt.Left + x - mlngOrgX + txt.Width <= picDraw.ScaleWidth Then
            picTxt.Left = picTxt.Left + x - mlngOrgX
            txt.Left = txt.Left + x - mlngOrgX
        End If
        If txt.Top + y - mlngOrgY >= 0 And txt.Top + y - mlngOrgY + txt.Height <= picDraw.ScaleHeight Then
            picTxt.Top = picTxt.Top + y - mlngOrgY
            txt.Top = txt.Top + y - mlngOrgY
        End If
        picDraw.Refresh
    End If
End Sub

Private Sub picTxt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    txt.SetFocus
End Sub

'################################################################################################################
'   ��;��  �ѱ༭��ı��ͼ�洢������С�
'################################################################################################################
Public Function SaveModifiedPictureToTable()
    Dim i As Long, lKey As Long

    ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks

    mblnModified = False
    RefreshSelectedMarks picDraw, PicMarks, 0, 0
    
    Set mfrmParent.tblThis.Cells("K" & mfrmParent.tblThis.SelectedCellKey).Picture = mcPicture.DrawFinalPic(mobjTable)
    mfrmParent.tblThis.Refresh False, False
End Function

'################################################################################################################
'   ��;��  �ѱ༭��ı��ͼ�洢���������RTF�С�
'################################################################################################################
Public Function SaveModifiedPictureToRTF()
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    
    lKey = mcPicture.Key     ' mlngCurPictureID
    bInKeys = FindKey(mfrmParent.Editor1, "P", lKey, lSS, lSE, lES, lEE, bNeeded)
    If bInKeys = False Then Exit Function

    Dim i As Long, p As Long
    ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks

    With mfrmParent.Editor1
        .TOM.TextDocument.Freeze
        .ForceEdit = True
        i = lSE
        Dim lS As Long, lE As Long
        lS = .Selection.StartPos
        lE = .Selection.EndPos
        .TOM.TextDocument.Range(i - 18, i + 17).Font.Protected = False
        .TOM.TextDocument.Range(i - 18, i + 17).Text = ""
        .TOM.TextDocument.Range(i - 18, i - 17).Select

        p = .InsertPicture(picDraw.Picture)
        .ForceEdit = True
        .TOM.TextDocument.Range(p, p) = vbCrLf & "PS(" & Format(lKey, "00000000") & ",1,0)"
        .TOM.TextDocument.Range(p + 19, p + 19) = "PE(" & Format(lKey, "00000000") & ",1,0) "   '�����ո��������������֣�
        .TOM.TextDocument.Range(p + 2, p + 18).Font.Hidden = True
        .TOM.TextDocument.Range(p + 19, p + 35).Font.Hidden = True
        .TOM.TextDocument.Range(p, p + 35).Font.Protected = True
        .Range(lS, lE).Selected
        
        .TOM.TextDocument.UnFreeze
        .LockAllOLEObjectSize
        .ForceEdit = False
    End With
    mblnModified = False
    RefreshSelectedMarks picDraw, PicMarks, 0, 0
End Function

Private Sub ColorFillColor_pOK()
    SendKeys "{ESCAPE}"
    mlngSelFillColor = IIf(ColorFillColor.COLOR = tomAutoColor, ColorFillColor.AutoColor, ColorFillColor.COLOR)
    Dim i As Long
    If mlngSelectedCount > 0 Then
        For i = 1 To PicMarks.Count
            If PicMarks(i).Selected Then
                PicMarks(i).���ɫ = mlngSelFillColor
            End If
        Next
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'        RefreshSelectedMarks picDraw, PicMarks, 0, 0
    End If
    SetColorIcon "FILLCOLOR", ID_DRAW_FILLCOLOR, mlngSelFillColor
End Sub

Private Sub ColorForeColor_GotFocus()
    ColorForeColor.Tag = "Focused"
End Sub

Private Sub ColorForeColor_pOK()
    SendKeys "{ESCAPE}"
    mlngSelForeColor = IIf(ColorForeColor.COLOR = tomAutoColor, ColorForeColor.AutoColor, ColorForeColor.COLOR)
    Dim i As Long
    If mlngSelectedCount > 0 Then
        For i = 1 To PicMarks.Count
            If PicMarks(i).Selected Then
                PicMarks(i).����ɫ = mlngSelForeColor
            End If
        Next
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'        RefreshSelectedMarks picDraw, PicMarks, 0, 0
    End If
    SetColorIcon "FORECOLOR", ID_DRAW_FONTCOLOR, mlngSelForeColor
End Sub

Private Sub ColorLineColor_GotFocus()
    ColorLineColor.Tag = "Focused"
End Sub

Private Sub ColorLineColor_pOK()
    SendKeys "{ESCAPE}"
    mlngSelLineColor = IIf(ColorLineColor.COLOR = tomAutoColor, ColorLineColor.AutoColor, ColorLineColor.COLOR)
    Dim i As Long
    If mlngSelectedCount > 0 Then
        For i = 1 To PicMarks.Count
            If PicMarks(i).Selected Then
                PicMarks(i).����ɫ = mlngSelLineColor
            End If
        Next
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'        RefreshSelectedMarks picDraw, PicMarks, 0, 0
    End If
    SetColorIcon "LINECOLOR", ID_DRAW_LINECOLOR, mlngSelLineColor
End Sub

Private Sub ColorFillColor_GotFocus()
    ColorFillColor.Tag = "Focused"
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long, j As Long
    Select Case Control.ID
    Case ID_DRAW_CLEARNUMBERS
        '��������Զ����
        For i = PicMarks.Count To 1 Step -1
            If PicMarks(i).���� = 6 Then
                PicMarks.Remove i
            End If
        Next
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
        mblnModified = True
        
    Case ID_DRAW_SELECT, ID_DRAW_MOVE, ID_DRAW_LINE, ID_DRAW_MLINE, ID_DRAW_RECT, ID_DRAW_MRECT, ID_DRAW_CIRCLE, ID_DRAW_TEXT, ID_DRAW_SEQUENCENUMBER
        mBar��ͼ.FindControl(, ID_DRAW_SELECT).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_MOVE).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_LINE).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_MLINE).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_RECT).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_MRECT).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_CIRCLE).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_TEXT).Checked = False
        mBar��ͼ.FindControl(, ID_DRAW_SEQUENCENUMBER).Checked = False
        
        Select Case Control.ID
        Case ID_DRAW_SELECT
            mBar��ͼ.FindControl(, ID_DRAW_SELECT).Checked = True
        Case ID_DRAW_MOVE
            mBar��ͼ.FindControl(, ID_DRAW_MOVE).Checked = True
        Case ID_DRAW_LINE
            mBar��ͼ.FindControl(, ID_DRAW_LINE).Checked = True
        Case ID_DRAW_MLINE
            mBar��ͼ.FindControl(, ID_DRAW_MLINE).Checked = True
        Case ID_DRAW_RECT
            mBar��ͼ.FindControl(, ID_DRAW_RECT).Checked = True
        Case ID_DRAW_MRECT
            mBar��ͼ.FindControl(, ID_DRAW_MRECT).Checked = True
        Case ID_DRAW_CIRCLE
            mBar��ͼ.FindControl(, ID_DRAW_CIRCLE).Checked = True
        Case ID_DRAW_TEXT
            mBar��ͼ.FindControl(, ID_DRAW_TEXT).Checked = True
        Case ID_DRAW_SEQUENCENUMBER
            mBar��ͼ.FindControl(, ID_DRAW_SEQUENCENUMBER).Checked = True
        End Select
        mlngDrawModeID = Control.ID
        Call SaveState

        Control.Checked = True
        
        If mblnInDrawing = False Then GetCurDrawMode
'            If txt.Visible Then FinishInputText
        'Public Const ID_DRAW_FILLCOLOR = 570
        'Public Const ID_DRAW_FILLNONE = 571
        'Public Const ID_DRAW_FILLALL = 572
        'Public Const ID_DRAW_FILLH = 573
        'Public Const ID_DRAW_FILLV = 574
        'Public Const ID_DRAW_FILLHV = 575
        'Public Const ID_DRAW_FILLR = 576
        'Public Const ID_DRAW_FILLL = 577
        'Public Const ID_DRAW_FILLLR = 578
        'Public Const ID_DRAW_LINECOLOR = 580
        'Public Const ID_DRAW_LINECONTINUE = 581
        'Public Const ID_DRAW_LINEDOT = 582
        'Public Const ID_DRAW_LINEDASH = 583
        'Public Const ID_DRAW_LINEDASHDOT = 584
        'Public Const ID_DRAW_LINEDASHDOT2 = 585
        'Public Const ID_DRAW_LINEWIDTH1 = 590
        'Public Const ID_DRAW_LINEWIDTH2 = 591
        'Public Const ID_DRAW_LINEWIDTH3 = 592
        'Public Const ID_DRAW_LINEWIDTH4 = 593
        'Public Const ID_DRAW_LINEWIDTH5 = 594
    Case ID_DRAW_DELETE
        If mblnInDrawing = False Then DeleteSelectedMarks: mblnModified = True
    Case ID_DRAW_RESET
        If mblnInDrawing = False Then ClearAllMarks: mblnModified = True
    Case ID_DRAW_FILLNONE, ID_DRAW_FILLALL, ID_DRAW_FILLH, ID_DRAW_FILLV, ID_DRAW_FILLHV, ID_DRAW_FILLR, ID_DRAW_FILLL, ID_DRAW_FILLLR
        SetFillIcon Control.ID
    
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLNONE).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLALL).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLH).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLV).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLHV).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLR).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLL).Checked = False
        mBar�����ʽ.CommandBar.FindControl(, ID_DRAW_FILLLR).Checked = False
        Control.Checked = True
        
        GetCurDrawMode
        If mlngSelectedCount > 0 Then
            mblnModified = True
            For i = 1 To PicMarks.Count
                If PicMarks(i).Selected Then
                    PicMarks(i).��䷽ʽ = gcurFillStyle
                End If
            Next
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'            RefreshSelectedMarks picDraw, PicMarks, 0, 0
        End If
    Case ID_DRAW_LINECONTINUE, ID_DRAW_LINEDOT, ID_DRAW_LINEDASH, ID_DRAW_LINEDASHDOT, ID_DRAW_LINEDASHDOT2
        mBar����.CommandBar.FindControl(, ID_DRAW_LINECONTINUE).Checked = False
        mBar����.CommandBar.FindControl(, ID_DRAW_LINEDOT).Checked = False
        mBar����.CommandBar.FindControl(, ID_DRAW_LINEDASH).Checked = False
        mBar����.CommandBar.FindControl(, ID_DRAW_LINEDASHDOT).Checked = False
        mBar����.CommandBar.FindControl(, ID_DRAW_LINEDASHDOT2).Checked = False
        Control.Checked = True
        GetCurDrawMode
        If mlngSelectedCount > 0 Then
            mblnModified = True
            For i = 1 To PicMarks.Count
                If PicMarks(i).Selected Then
                    PicMarks(i).���� = gcurPenStyle
                End If
            Next
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'            RefreshSelectedMarks picDraw, PicMarks, 0, 0
        End If
    Case ID_DRAW_LINEWIDTH1, ID_DRAW_LINEWIDTH2, ID_DRAW_LINEWIDTH3, ID_DRAW_LINEWIDTH4, ID_DRAW_LINEWIDTH5
        mBar�߿�.CommandBar.FindControl(, ID_DRAW_LINEWIDTH1).Checked = False
        mBar�߿�.CommandBar.FindControl(, ID_DRAW_LINEWIDTH2).Checked = False
        mBar�߿�.CommandBar.FindControl(, ID_DRAW_LINEWIDTH3).Checked = False
        mBar�߿�.CommandBar.FindControl(, ID_DRAW_LINEWIDTH4).Checked = False
        mBar�߿�.CommandBar.FindControl(, ID_DRAW_LINEWIDTH5).Checked = False
        Control.Checked = True
        GetCurDrawMode
        If mlngSelectedCount > 0 Then
            mblnModified = True
            For i = 1 To PicMarks.Count
                If PicMarks(i).Selected Then
                    PicMarks(i).�߿� = gcurPenWidth
                End If
            Next
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'            RefreshSelectedMarks picDraw, PicMarks, 0, 0
        End If
    Case ID_FILE_SAVE
        
    Case ID_COMMON_CANCEL
'        If mblnModified Then
'            Dim lngR As Long
'            lngR = MsgBox("�Ƿ񱣴��޸ģ�", vbYesNoCancel + vbQuestion, "����")
'            If lngR = vbYes Then
'                Set mcPicture.PicMarks = PicMarks.Clone
'                If mfrmParent.Name = "frmMain" Then
'                    'ͨ����������õ�ͼƬ�༭��
'                    Call SaveModifiedPictureToRTF
'                Else
'                    'ͨ�����༭�����õ�ͼƬ�༭��
'                    Call SaveModifiedPictureToTable
'                End If
'                mblnOK = True
'                Unload Me
'            ElseIf lngR = vbNo Then
'                Unload Me
'            End If
'        Else
'            Unload Me
'        End If
    Case ID_DRAW_FILLCOLOR
        Call ColorFillColor_pOK
    Case ID_DRAW_LINECOLOR
        Call ColorLineColor_pOK
    Case ID_DRAW_FONTCOLOR
        Call ColorForeColor_pOK
    End Select
End Sub

'################################################################################################################
'   �����ǣ�����ͼƬ
'################################################################################################################
Private Sub SavePicture()
    If mobjTable Is Nothing Then
        Set mcPicture.PicMarks = PicMarks
    Else
        If mobjTable.TableType = tte_����ͼƬ�� Then
            Set mcPicture.PicMarks = PicMarks
        Else
            Set mcPicture.PicMarks = ScalePicMarks(PicMarks, 1# / zoomFactor)
        End If
    End If
    If mblnEditInTable Then
        'ͨ�����༭�����õ�ͼƬ�༭��
        Call SaveModifiedPictureToTable
    Else
        'ͨ����������õ�ͼƬ�༭��
        Call SaveModifiedPictureToRTF
    End If
    mblnOK = True
End Sub

Private Sub UserControl_Resize()
    picBG.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If UserControl.Extender.Visible = False Then Exit Sub
    Select Case Control.ID
    Case ID_DRAW_DELETE
        Control.Enabled = (mlngSelectedCount > 0)
    Case ID_FILE_SAVE
        Control.Enabled = mblnModified
    End Select
End Sub

'Private Sub picDraw_GotFocus()
''    picHolder(0).Visible = True
''    picHolder(1).Visible = True
''    picHolder(2).Visible = True
''    picHolder(3).Visible = True
''    picHolder(4).Visible = True
''    picHolder(5).Visible = True
''    picHolder(6).Visible = True
''    picHolder(7).Visible = True
'
'    If mcPicture.IsmarkedPic Then
'        '��ȡ��ǰ��ͼģʽ��Ϣ
'        If Not mblnInDrawing Then Call GetCurDrawMode
'
'        '��ͼ��������ʾ
'        'frmParent.mBar��ͼ.Visible = True
'        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'    '    RefreshSelectedMarks picDraw, PicMarks, 0, 0
'        mlngPicPosition = mfrmParent.Editor1.SelStart
'        mfrmParent.Editor1.SelLength = 0
'    End If
'End Sub

'Private Sub picDraw_LostFocus()
'    '��ͼ����������
''    frmParent.mBar��ͼ.Visible = False
'
'    'ˢ��ͼƬ
'    If mcPicture.IsmarkedPic Then
'        picDraw.Picture = picDraw.Image
'        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
'    '    ShowModifiedPicture
'    End If
'End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '��ȡ��ǰ��ͼģʽ��Ϣ
    If mcPicture.PictureType = EPRFormulaPicture Then Exit Sub
    Dim lTxtID As Long, i As Long, X1 As Long, Y1 As Long, lp As POINTAPI
    mblnDblClick = False
    
    If Button = vbRightButton Then ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks: Exit Sub

'    If mcPicture.IsMarkedPic = False Then Exit Sub
    If Not mblnInDrawing Then Call GetCurDrawMode
    
    picDraw.Tag = "����ˢ��"
    
    If txt.Visible Then FinishInputText         '������ı�
    
    '��ʼ�����
    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
        mblnInDrawing = True
        RefreshSelectedMarks picDraw, PicMarks, 0, 0
    Case ID_DRAW_MOVE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_LINE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, mvarOldPoint.x, mvarOldPoint.y
        
        mblnInDrawing = True
    Case ID_DRAW_RECT
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_MLINE
        If mblnInDrawing = False Then
            '������ʼ��λ��
            mvarFirstPoint.x = x
            mvarFirstPoint.y = y
            mvarOldPoint.x = x
            mvarOldPoint.y = y
            ReDim mvarPolyPoints(1 To 1) As POINTAPI
            mvarPolyPoints(1).x = x: mvarPolyPoints(1).y = y
        End If
        mblnInDrawing = True
    Case ID_DRAW_MRECT
        If mblnInDrawing = False Then
            '������ʼ��λ��
            mvarFirstPoint.x = x
            mvarFirstPoint.y = y
            mvarOldPoint.x = x
            mvarOldPoint.y = y
            ReDim mvarPolyPoints(1 To 1) As POINTAPI
            mvarPolyPoints(1).x = x: mvarPolyPoints(1).y = y
        End If
        mblnInDrawing = True
    Case ID_DRAW_CIRCLE
        '������ʼ��λ��
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnInDrawing = True
    Case ID_DRAW_DELETE
        mblnInDrawing = True
    Case ID_DRAW_TEXT
           
        '���Ƿ�ѡ����ĳһ���ı�
        For i = 1 To PicMarks.Count
            If PicMarks(i).���� = 0 Then
                If x > PicMarks(i).X1 And x < PicMarks(i).X2 And y > PicMarks(i).Y1 - 2 And y < PicMarks(i).Y2 - 2 Then
                    lTxtID = i
                    Exit For
                End If
            End If
        Next i

        If lTxtID > 0 Then
            'ѡ��һ�������ı�
            With PicMarks(lTxtID)
                Set txt.Font = .����
                txt.Text = .����
                txt.Move .X1, .Y1, (.X2 - .X1), (.Y2 - .Y1)
            End With
            PicMarks.Remove lTxtID
            '���������
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
        Else
            '�½�һ���ı�
            txt.Text = ""
            
            txt.Top = y: txt.Left = x
            Call GetFitTxtSize(txt, txt.Text, X1, Y1)
            txt.Width = X1 + 10
            txt.Height = Y1 + 6
        End If
        picTxt.Top = txt.Top - picTxt.Height / 2
        picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
        txt.Visible = True
        picTxt.Visible = True
        txt.SetFocus
    Case ID_DRAW_SEQUENCENUMBER
        Call SetDrawStyleFromValue(picDraw.hDC, RGB(255, 255, 0), 0, 1, RGB(255, 255, 0), -2)
        Ellipse picDraw.hDC, x - 7, y - 7, x + 7, y + 7
        Call SetDrawStyleFromValue(picDraw.hDC, vbBlack, 0, 1, vbBlack, -1)
        Ellipse picDraw.hDC, x - 7, y - 7, x + 7, y + 7
        Dim Font As New StdFont
        Font.Bold = True
        Dim Num As Long
        Num = GetMaxNum
        Call TextOut(picDraw, Num, IIf(Len(CStr(Num)) > 1, x - 6, x - 2), y - 6, x + 14, y + 14, Font)

        picDraw.Refresh
        '��������
        PicMarks.Add
        With PicMarks.LastPicMark
            .���� = Num
            .X1 = x: .Y1 = y
            .X2 = x: .Y2 = y
            .���� = 6
            .��䷽ʽ = -2
            .���ɫ = vbBlack
            .�߿� = 1
            .����ɫ = vbBlack
            .���� = 1
        End With
        mblnInDrawing = False
        mblnModified = True
    End Select
End Sub

Private Function GetMaxNum() As Long
    '��ȡ�Զ���ŵ����ֵ
    Dim i As Long, j As Long
    For i = 1 To PicMarks.Count
        If PicMarks(i).���� = 6 Then
            If j < CLng(PicMarks(i).����) Then j = CLng(PicMarks(i).����)
        End If
    Next
    GetMaxNum = j + 1
End Function

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If mblnInDrawing = False Then Exit Sub
'    If mcPicture.IsMarkedPic = False Then Exit Sub
    Dim tmpX As Long, tmpY As Long, lp As POINTAPI
    
    '���߻��Ʊ߿�
    Call SetDrawStyleFromValue(picDraw.hDC, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), gcurPenWidth, mlngFillColor, -1)
    
    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '���߻��Ʊ߿�
        Call SetDrawStyleFromValue(picDraw.hDC, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), 1, mlngFillColor, -1)
        '����
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_MOVE
        '�ƶ�ѡ�б��
        '����
        tmpX = mvarOldPoint.x - mvarFirstPoint.x: tmpY = mvarOldPoint.y - mvarFirstPoint.y  '��ƫ����
        RefreshSelectedMarks picDraw, PicMarks, tmpX, tmpY    'ˢ��ѡ�еı�ǵ��µ�ַ
        
        '����
        tmpX = x - mvarFirstPoint.x: tmpY = y - mvarFirstPoint.y
        RefreshSelectedMarks picDraw, PicMarks, tmpX, tmpY    'ˢ��ѡ�еı�ǵ��µ�ַ
        picDraw.Refresh
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_LINE
        '������ǰ����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, mvarOldPoint.x, mvarOldPoint.y
        
        '�����µ�����
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_RECT
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = tmpX
        mvarOldPoint.y = tmpY
    Case ID_DRAW_MLINE
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, mvarOldPoint.x, mvarOldPoint.y
        
        '�����µ�����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_MRECT
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, mvarOldPoint.x, mvarOldPoint.y
        
        '�����µ�����
        picDraw.DrawMode = vbInvert
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = x
        mvarOldPoint.y = y
    Case ID_DRAW_CIRCLE
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbInvert
        Ellipse picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        '����
        Ellipse picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '�����µ�ĩβ��λ��
        mvarOldPoint.x = tmpX
        mvarOldPoint.y = tmpY
    Case ID_DRAW_DELETE
    
    End Select
    
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mcPicture.PictureType = EPRFormulaPicture Then Exit Sub '��ʽͼ����˵�
    If Button = vbRightButton And mblnInDrawing = False Then
        '�Ҽ��˵�����
        Dim Popup As CommandBar
        Dim objControl As CommandBarControl
        
        Set Popup = cbsThis.Add("Popup", xtpBarPopup)
        With Popup.Controls
            Set objControl = .Add(xtpControlButton, ID_DRAW_SELECT, "ѡ��(&E)")
            objControl.BeginGroup = True
            .Add xtpControlButton, ID_DRAW_MOVE, "�ƶ�(&M)"
            .Add xtpControlButton, ID_DRAW_LINE, "ֱ��(&L)"
            .Add xtpControlButton, ID_DRAW_MLINE, "����(&Z)"
            .Add xtpControlButton, ID_DRAW_RECT, "����(&R)"
            .Add xtpControlButton, ID_DRAW_MRECT, "�����(&W)"
            .Add xtpControlButton, ID_DRAW_CIRCLE, "��Բ(&C)"
            .Add xtpControlButton, ID_DRAW_TEXT, "����(&T)"
            .Add xtpControlButton, ID_DRAW_SEQUENCENUMBER, "˳����(&N)"
            
            Set objControl = .Add(xtpControlButton, ID_DRAW_CLEARNUMBERS, "���˳����(&K)")
            objControl.BeginGroup = True
            
            Set objControl = .Add(xtpControlButton, ID_DRAW_DELETE, "ɾ��(&D)")
            objControl.BeginGroup = True
        End With
        Popup.ShowPopup
        Exit Sub
    End If
    If mblnInDrawing = False Then Exit Sub
    
    '�ָ���䷽ʽ
    Call SetDrawStyleFromValue(picDraw.hDC, mlngLineColor, gcurPenStyle, gcurPenWidth, mlngFillColor, gcurFillStyle)
    Dim tmpX As Long, tmpY As Long
    Dim strTmp As String, i As Long, lp As POINTAPI

    Select Case mlngDrawModeID
    Case ID_DRAW_SELECT
        '����
        '���߻��Ʊ߿�
        Call SetDrawStyleFromValue(picDraw.hDC, mlngLineColor, IIf(gcurPenStyle = 0, 2, gcurPenStyle), 1, mlngFillColor, -1)
        picDraw.DrawMode = vbInvert
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y
        picDraw.Refresh
        mblnInDrawing = False
        
        'ѡ�з�ΧΪ��mvarFirstPoint,mvarOldPoint����
        '�����ж����б������Щ��ѡ�У���������ʾ
        Call HilightSelectMarks(mvarFirstPoint.x, mvarFirstPoint.y, mvarOldPoint.x, mvarOldPoint.y)
    Case ID_DRAW_MOVE
        '�����±�ǣ�ˢ��ͼ��
        tmpX = x - mvarFirstPoint.x: tmpY = y - mvarFirstPoint.y
        SaveSelectedMarks PicMarks, tmpX, tmpY
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
        mblnInDrawing = False
        RefreshSelectedMarks picDraw, PicMarks, 0, 0
        mblnModified = True
    Case ID_DRAW_LINE
        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        '��������
        PicMarks.Add
        With PicMarks.LastPicMark
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = x: .Y2 = y
            .���� = 1            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mblnInDrawing = False
        mblnModified = True
    Case ID_DRAW_RECT
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbCopyPen
        Rectangle picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '��������
        PicMarks.Add
        With PicMarks.LastPicMark
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = tmpX: .Y2 = tmpY
            .���� = 3            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mblnInDrawing = False
        mblnModified = True
    Case ID_DRAW_MLINE
'        If mvarFirstPoint.X = X And mvarFirstPoint.Y = Y And Button <> vbRightButton Then Exit Sub
        If Button = vbRightButton Then
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
            mblnInDrawing = False
            ReDim mvarPolyPoints(0)
            Exit Sub
        End If

        '���浱ǰ��
        ReDim Preserve mvarPolyPoints(1 To UBound(mvarPolyPoints) + 1) As POINTAPI
        mvarPolyPoints(UBound(mvarPolyPoints)).x = x
        mvarPolyPoints(UBound(mvarPolyPoints)).y = y
        
        If mblnDblClick And UBound(mvarPolyPoints) >= 2 Then
            '�������ݣ��˳���ͼ
            PicMarks.Add
            With PicMarks.LastPicMark
                .���� = 2            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
                .��䷽ʽ = gcurFillStyle
                .���ɫ = gcurFillColor
                .�߿� = gcurPenWidth
                .����ɫ = gcurPenColor
                .���� = gcurPenStyle
                For i = 1 To UBound(mvarPolyPoints)
                    If i = 1 Then
                        strTmp = strTmp & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    Else
                        strTmp = strTmp & ";" & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    End If
                Next i
                .�㼯 = strTmp              '����㼯����
            End With
            mblnInDrawing = False
        End If
        
        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnModified = True
    Case ID_DRAW_MRECT
'        If mvarFirstPoint.X = X And mvarFirstPoint.Y = Y And Button <> vbRightButton Then Exit Sub
        If Button = vbRightButton Then
            ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
            mblnInDrawing = False
            ReDim mvarPolyPoints(0)
            Exit Sub
        End If
                
        '���浱ǰ��
        ReDim Preserve mvarPolyPoints(1 To UBound(mvarPolyPoints) + 1) As POINTAPI
        mvarPolyPoints(UBound(mvarPolyPoints)).x = x
        mvarPolyPoints(UBound(mvarPolyPoints)).y = y
        
        If mblnDblClick And UBound(mvarPolyPoints) >= 2 Then
            '�������ն����
            picDraw.DrawMode = vbCopyPen
            Polygon picDraw.hDC, mvarPolyPoints(1), UBound(mvarPolyPoints)
            
            '�������ݣ��˳���ͼ
            PicMarks.Add
            With PicMarks.LastPicMark
                .���� = 4            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
                .��䷽ʽ = gcurFillStyle
                .���ɫ = gcurFillColor
                .�߿� = gcurPenWidth
                .����ɫ = gcurPenColor
                .���� = gcurPenStyle
                For i = 1 To UBound(mvarPolyPoints)
                    If i = 1 Then
                        strTmp = strTmp & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    Else
                        strTmp = strTmp & ";" & CStr(mvarPolyPoints(i).x) & "," & CStr(mvarPolyPoints(i).y)
                    End If
                Next i
                .�㼯 = strTmp              '����㼯����
            End With
            mblnInDrawing = False
        End If
        
        '������������
        picDraw.DrawMode = vbCopyPen
        MoveToEx picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, lp
        LineTo picDraw.hDC, x, y
        
        mvarFirstPoint.x = x
        mvarFirstPoint.y = y
        mvarOldPoint.x = x
        mvarOldPoint.y = y
        mblnModified = True
    Case ID_DRAW_CIRCLE
        tmpX = x: tmpY = y
        If Shift = 2 Then '������
            Call ForceSquare(mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY)
        End If
        '����
        picDraw.DrawMode = vbCopyPen
        Ellipse picDraw.hDC, mvarFirstPoint.x, mvarFirstPoint.y, tmpX, tmpY
        picDraw.Refresh
        '��������
        PicMarks.Add
        With PicMarks.LastPicMark
            .X1 = mvarFirstPoint.x: .Y1 = mvarFirstPoint.y
            .X2 = tmpX: .Y2 = tmpY
            .���� = 5            '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            .��䷽ʽ = gcurFillStyle
            .���ɫ = gcurFillColor
            .�߿� = gcurPenWidth
            .����ɫ = gcurPenColor
            .���� = gcurPenStyle
        End With
        mblnInDrawing = False
        mblnModified = True
    Case ID_DRAW_DELETE
        '����
        
        mblnModified = True
    End Select
    
    picDraw.DrawMode = vbCopyPen
    picDraw.Refresh

    '���浽������
'    gfrm��Ϣ.txtInfo.Text = "��ǰ��ʱ�㼯��Ŀ��" & UBound(mvarPolyPoints)
End Sub

'################################################################################################################
'   ��;��  �����ı���ǰ���ݵĺ��ʳߴ硣
'################################################################################################################
Private Sub GetFitTxtSize(objMain As Object, strText As String, Optional ByRef Width As Long, Optional ByRef Height As Long, Optional ByRef LineHeight As Long)
    '���أ�w,h�����ߴ�,h2���и߶�
    With objMain
        picTxt.FontName = .FontName
        picTxt.FontSize = .FontSize
        picTxt.FontBold = .FontBold
        picTxt.FontItalic = .FontItalic
        picTxt.FontUnderline = .FontUnderline
        picTxt.FontStrikethru = .FontStrikethru
        If strText = "" Then
            Width = picTxt.TextWidth("AA")
            Height = picTxt.TextHeight("A")
        Else
            Width = picTxt.TextWidth(strText & "A")
            Height = picTxt.TextHeight(strText)
        End If
        LineHeight = picTxt.TextHeight("A")
    End With
End Sub

'################################################################################################################
'   ��;��  ��ɵ�ǰ�������롣
'################################################################################################################
Public Sub FinishInputText()
    If txt.Visible Then
        '������״̬תΪȷ�����벢�˳�
        If Trim(Replace(txt.Text, vbCrLf, "")) <> "" Then
            '����������
            PicMarks.Add
            With PicMarks.LastPicMark
                .���� = 0
                .���� = txt.Text
                Set .���� = txt.Font
                .X1 = txt.Left
                .Y1 = txt.Top
                .X2 = txt.Left + txt.Width
                .Y2 = txt.Top + txt.Height
                
                TextOut picDraw, .����, .X1, .Y1, .X2, .Y2, .����
            End With
        End If
        txt.Text = ""
        txt.Visible = False
        picTxt.Visible = False
    End If
End Sub

'################################################################################################################
'   ��;��  �����ǰ���ͼ���б�ǡ�
'################################################################################################################
Public Sub ClearAllMarks()
    If PicMarks.Count = 0 Or picDraw.Visible = False Then Exit Sub
    If MsgBox("ȷ�����ͼ�����б��������", vbOKCancel + vbInformation, "ȷ�����") = vbCancel Then Exit Sub
    Set PicMarks = New cPicMarks
    mlngSelectedCount = 0
    'ˢ�½����
    ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
    picDraw.Visible = True
    picDraw.SetFocus
End Sub

'################################################################################################################
'   ��;��  ˢ�µ�ǰѡ�е����б�ǡ�
'################################################################################################################
Public Sub RefreshSelectedMarks(objPic As PictureBox, objMarks As cPicMarks, x As Long, y As Long)
    Dim arrTmp() As String, arrXY() As POINTAPI, lp As POINTAPI
    Dim i As Integer, j As Integer

'    Screen.MousePointer = 11
    LockWindowUpdate objPic.hwnd

    objPic.DrawMode = vbInvert

    '������Ԫ��
    For i = 1 To objMarks.Count
        With objMarks(i)
            If .Selected Then
                If .���� <> 0 Then
                    Call SetDrawStyleFromValue(objPic.hDC, .����ɫ, .����, .�߿�, .���ɫ, .��䷽ʽ)
                End If
                Select Case .����
'                    Case 0 '�ı�
'                        Call TextOut(objPic, .����, .X1 + x, .Y1 + y, .X2 + x, .Y2 + y, .����)
                    Case 1 '����
                        MoveToEx objPic.hDC, .X1 + x, .Y1 + y, lp
                        LineTo objPic.hDC, .X2 + x, .Y2 + y
                    Case 2 '����
                        arrTmp = Split(.�㼯, ";")
                        For j = 0 To UBound(arrTmp)
                            ReDim Preserve arrXY(j)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        Polyline objPic.hDC, arrXY(0), UBound(arrXY) + 1
                    Case 3 '����
                        Rectangle objPic.hDC, .X1 + x, .Y1 + y, .X2 + x, .Y2 + y
                    Case 4 '�����
                        arrTmp = Split(.�㼯, ";")
                        For j = 0 To UBound(arrTmp)
                            ReDim Preserve arrXY(j)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        Polygon objPic.hDC, arrXY(0), UBound(arrXY) + 1
                    Case 5 'Բ
                        Ellipse objPic.hDC, .X1 + x, .Y1 + y, .X2 + x, .Y2 + y
                End Select
            End If
        End With
    Next
    objPic.Refresh

    GetCurDrawMode

    LockWindowUpdate 0
    Screen.MousePointer = 0
End Sub

'################################################################################################################
'   ��;��  ɾ����ǰѡ�еı�ǡ�
'################################################################################################################
Public Sub DeleteSelectedMarks()
    If mlngSelectedCount = 0 Or picDraw.Visible = False Then Exit Sub
    Dim arrTmp() As String
    Dim i As Integer, j As Integer, strTmp As String
    j = 0
    For i = 1 To PicMarks.Count
        With PicMarks(i)
            If .Selected Then
                ReDim Preserve arrTmp(j) As String
                arrTmp(j) = .Key
                j = j + 1
            End If
        End With
    Next
    
    mlngSelectedCount = j
    If MsgBox("ȷ��ɾ��ѡ�е� " & mlngSelectedCount & " �������", vbOKCancel + vbInformation, "ȷ��ɾ��") = vbCancel Then Exit Sub
    
    For i = 0 To mlngSelectedCount - 1
        PicMarks.Remove arrTmp(i)
    Next i
    
    'ˢ�½����
    mlngSelectedCount = 0
    
    ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
    picDraw.Visible = True
    picDraw.SetFocus
End Sub

'################################################################################################################
'   ��;��  ��ȡ��ǰ��ͼģʽ��
'################################################################################################################
Public Sub GetCurDrawMode()
    With mBar��ͼ
        Select Case True
        Case .FindControl(, ID_DRAW_SELECT).Checked
            mlngDrawModeID = ID_DRAW_SELECT
        Case .FindControl(, ID_DRAW_MOVE).Checked
            mlngDrawModeID = ID_DRAW_MOVE
        Case .FindControl(, ID_DRAW_LINE).Checked
            mlngDrawModeID = ID_DRAW_LINE
        Case .FindControl(, ID_DRAW_MLINE).Checked
            mlngDrawModeID = ID_DRAW_MLINE
        Case .FindControl(, ID_DRAW_RECT).Checked
            mlngDrawModeID = ID_DRAW_RECT
        Case .FindControl(, ID_DRAW_MRECT).Checked
            mlngDrawModeID = ID_DRAW_MRECT
        Case .FindControl(, ID_DRAW_CIRCLE).Checked
            mlngDrawModeID = ID_DRAW_CIRCLE
        Case .FindControl(, ID_DRAW_TEXT).Checked
            mlngDrawModeID = ID_DRAW_TEXT
        Case .FindControl(, ID_DRAW_SEQUENCENUMBER).Checked
            mlngDrawModeID = ID_DRAW_SEQUENCENUMBER
        Case .FindControl(, ID_DRAW_DELETE).Checked
            mlngDrawModeID = ID_DRAW_DELETE
        End Select
    End With

    mlngFillColor = ColorFillColor.COLOR
    mlngLineColor = ColorLineColor.COLOR

    With mBar�����ʽ.CommandBar
        Select Case True
        Case .FindControl(, ID_DRAW_FILLNONE).Checked
            mlngFillStyleID = ID_DRAW_FILLNONE
        Case .FindControl(, ID_DRAW_FILLALL).Checked
            mlngFillStyleID = ID_DRAW_FILLALL
        Case .FindControl(, ID_DRAW_FILLH).Checked
            mlngFillStyleID = ID_DRAW_FILLH
        Case .FindControl(, ID_DRAW_FILLV).Checked
            mlngFillStyleID = ID_DRAW_FILLV
        Case .FindControl(, ID_DRAW_FILLHV).Checked
            mlngFillStyleID = ID_DRAW_FILLHV
        Case .FindControl(, ID_DRAW_FILLR).Checked
            mlngFillStyleID = ID_DRAW_FILLR
        Case .FindControl(, ID_DRAW_FILLL).Checked
            mlngFillStyleID = ID_DRAW_FILLL
        Case .FindControl(, ID_DRAW_FILLLR).Checked
            mlngFillStyleID = ID_DRAW_FILLLR
        End Select
    End With

    With mBar����.CommandBar
        Select Case True
        Case .FindControl(, ID_DRAW_LINECONTINUE).Checked
            mlngLineStyleID = ID_DRAW_LINECONTINUE
        Case .FindControl(, ID_DRAW_LINEDOT).Checked
            mlngLineStyleID = ID_DRAW_LINEDOT
        Case .FindControl(, ID_DRAW_LINEDASH).Checked
            mlngLineStyleID = ID_DRAW_LINEDASH
        Case .FindControl(, ID_DRAW_LINEDASHDOT).Checked
            mlngLineStyleID = ID_DRAW_LINEDASHDOT
        Case .FindControl(, ID_DRAW_LINEDASHDOT2).Checked
            mlngLineStyleID = ID_DRAW_LINEDASHDOT2
        End Select
    End With

    With mBar�߿�.CommandBar
        Select Case True
        Case .FindControl(, ID_DRAW_LINEWIDTH1).Checked
            mlngLineWidthID = ID_DRAW_LINEWIDTH1
        Case .FindControl(, ID_DRAW_LINEWIDTH2).Checked
            mlngLineWidthID = ID_DRAW_LINEWIDTH2
        Case .FindControl(, ID_DRAW_LINEWIDTH3).Checked
            mlngLineWidthID = ID_DRAW_LINEWIDTH3
        Case .FindControl(, ID_DRAW_LINEWIDTH4).Checked
            mlngLineWidthID = ID_DRAW_LINEWIDTH4
        Case .FindControl(, ID_DRAW_LINEWIDTH5).Checked
            mlngLineWidthID = ID_DRAW_LINEWIDTH5
        End Select

    End With

    '���������
    SetCursor mlngDrawModeID

    '���õ�ǰ��ͼģʽ�����ʡ���ˢ��
    SetDrawStyle picDraw.hDC
End Sub

'################################################################################################################
'   ��;��  �ж����б������Щ��ѡ�У���������ʾ��
'################################################################################################################
Private Sub HilightSelectMarks(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    '�ȴ洢ͼ��
    mlngSelectedCount = 0
    
    '����X1��Y1��X2��Y2��ʹ��(X1,Y1)�������Ͻǣ���(X2,Y2)�������½�
    Dim lTmp As Long
    If X1 > X2 Then
        lTmp = X2
        X2 = X1
        X1 = lTmp
    End If
    If Y1 > Y2 Then
        lTmp = Y2
        Y2 = Y1
        Y1 = lTmp
    End If
        
    Dim i As Long, j As Long, p As Long, q As Long, lSplit As Long, k As Long
    Dim T As Variant
    Dim lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, l As Long
    Dim arrXY() As POINTAPI, lp As POINTAPI
    Dim a As Double, b As Double, XX As Double, YY As Double
    
    i = giGetShiftState()
    If i <> vbShiftMask And i <> vbCtrlMask Then
        '����ס Shift ���� Control ��ѡ��ǡ�
        For i = 1 To PicMarks.Count
            PicMarks(i).Selected = False
        Next i
        ShowPicMarks picDraw, mcPicture.OrigPic, PicMarks
    End If
    For i = 1 To PicMarks.Count
        With PicMarks(i)
            If .���� <> 0 Then
                Call SetDrawStyleFromValue(picDraw.hDC, .����ɫ, .����, .�߿�, .���ɫ, .��䷽ʽ)
            End If
            picDraw.DrawMode = vbInvert
            Select Case .����   '0-�ı�,1-����,2,����,3-����,4-�����,5-Բ(��Բ)
            Case 0
                '
            Case 1
                '�ȿ�����߶ζ˵���һ���ھ����У���ѡ��֮��
                If (.X1 > X1 And .X1 < X2 And .Y1 > Y1 And .Y1 < Y2) Or (.X2 > X1 And .X2 < X2 And .Y2 > Y1 And .Y2 < Y2) Then
                    .Selected = Not .Selected
                    MoveToEx picDraw.hDC, .X1, .Y1, lp
                    LineTo picDraw.hDC, .X2, .Y2
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
                '�������ֶ�N(100)�ݣ�ȡ���ھ����У����ʾѡ��
                lSplit = IIf(Abs(.X2 - .X1) > Abs(.Y2 - .Y1), Abs(.X2 - .X1), Abs(.Y2 - .Y1))
                For j = 1 To lSplit
                    p = (j / lSplit) * (.X2 - .X1) + .X1
                    q = (j / lSplit) * (.Y2 - .Y1) + .Y1
                    '(p,q) �ھ�����
                    If p > X1 And p < X2 And q > Y1 And q < Y2 Then
                        .Selected = Not .Selected
                        MoveToEx picDraw.hDC, .X1, .Y1, lp
                        LineTo picDraw.hDC, .X2, .Y2
                        mlngSelectedCount = mlngSelectedCount + 1
                        GoTo LL
                    End If
                Next j
            Case 2
                '����
                'ͬ�������߷ֶΣ�ȡ���ھ����У����ʾѡ��
                T = Split(.�㼯, ";")
                For k = 1 To UBound(T)
                    lX1 = CLng(Split(T(k - 1), ",")(0))
                    lY1 = CLng(Split(T(k - 1), ",")(1))
                    lX2 = CLng(Split(T(k), ",")(0))
                    lY2 = CLng(Split(T(k), ",")(1))

                    lSplit = IIf(Abs(lX2 - lX1) > Abs(lY2 - lY1), Abs(lX2 - lX1), Abs(lY2 - lY1))
                    For j = 1 To lSplit
                        p = (j / lSplit) * (lX2 - lX1) + lX1
                        q = (j / lSplit) * (lY2 - lY1) + lY1
                        '(p,q) �ھ�����
                        If p > X1 And p < X2 And q > Y1 And q < Y2 Then
                            .Selected = Not .Selected
                            ReDim Preserve arrXY(UBound(T))
                            For l = 0 To UBound(T)
                                arrXY(l).x = CLng(Split(T(l), ",")(0))
                                arrXY(l).y = CLng(Split(T(l), ",")(1))
                            Next
                            Polyline picDraw.hDC, arrXY(0), UBound(T) + 1
                            mlngSelectedCount = mlngSelectedCount + 1
                            GoTo LL
                        End If
                    Next j
                Next k
            Case 3
                '����
                If ����������ཻ(X1, Y1, X2, Y2, .X1, .Y1, .X2, .Y2) Then
                    .Selected = Not .Selected
                    Rectangle picDraw.hDC, .X1, .Y1, .X2, .Y2
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
            Case 4
                '�����
                T = Split(.�㼯, ";")
                ReDim Preserve arrXY(UBound(T))
                For l = 0 To UBound(T)
                    arrXY(l).x = CLng(Split(T(l), ",")(0))
                    arrXY(l).y = CLng(Split(T(l), ",")(1))
                Next
                If �����������ཻ(X1, Y1, X2, Y2, arrXY) Then
                    .Selected = Not .Selected
                    ReDim Preserve arrXY(UBound(T))
                    For l = 0 To UBound(T)
                        arrXY(l).x = CLng(Split(T(l), ",")(0))
                        arrXY(l).y = CLng(Split(T(l), ",")(1))
                    Next
                    Polygon picDraw.hDC, arrXY(0), UBound(T) + 1
                    mlngSelectedCount = mlngSelectedCount + 1
                    GoTo LL
                End If
            Case 5
                '����4������Բ�н��㣡
                If ��������Բ�ཻ(X1, Y1, X2, Y2, .X1, .Y1, .X2, .Y2) Then
                    .Selected = Not .Selected
                    mlngSelectedCount = mlngSelectedCount + 1
                    Ellipse picDraw.hDC, .X1, .Y1, .X2, .Y2
                    GoTo LL
                End If
            End Select
        End With
LL:
    Next i
    GetCurDrawMode
End Sub

'################################################################################################################
'   ��;��  ����ѡ�б�ǵ��������ꡣ
'################################################################################################################
Public Sub SaveSelectedMarks(objMarks As cPicMarks, x As Long, y As Long)
    'objMarks=�����е�ǰ��Ŀ�ı��ͼ����
    'X,Y Ϊ����ƫ��
    Dim arrTmp() As String, arrXY() As POINTAPI
    Dim i As Integer, j As Integer, strTmp As String
     
    For i = 1 To objMarks.Count
        With objMarks(i)
            If .Selected Then
                Select Case .����
                    Case 1, 3, 5    '0 �ı�  1 ����  3  ���� 5  Բ
                        .X1 = .X1 + x
                        .Y1 = .Y1 + y
                        .X2 = .X2 + x
                        .Y2 = .Y2 + y
                    Case 2, 4 '����
                        arrTmp = Split(.�㼯, ";")
                        ReDim Preserve arrXY(UBound(arrTmp)) As POINTAPI
                        For j = 0 To UBound(arrTmp)
                            arrXY(j).x = CLng(Split(arrTmp(j), ",")(0)) + x
                            arrXY(j).y = CLng(Split(arrTmp(j), ",")(1)) + y
                        Next
                        strTmp = ""
                        For j = 0 To UBound(arrXY)
                            If j = 0 Then
                                strTmp = strTmp & CStr(arrXY(j).x) & "," & CStr(arrXY(j).y)
                            Else
                                strTmp = strTmp & ";" & CStr(arrXY(j).x) & "," & CStr(arrXY(j).y)
                            End If
                        Next j
                        .�㼯 = strTmp              '����㼯����
                End Select
            End If
        End With
    Next
End Sub

'################################################################################################################
'   ��;��  ���õ�ǰ����ꡣ
'################################################################################################################
Private Sub SetCursor(ByVal ID As Long)
    picDraw.MousePointer = 99
    Select Case ID
    Case ID_DRAW_SELECT
        Set picDraw.MouseIcon = imgCur.ListImages("Sel").Picture
'        picDraw.MousePointer = 1
    Case ID_DRAW_MOVE
        Set picDraw.MouseIcon = imgCur.ListImages("Move").Picture
    Case ID_DRAW_LINE, ID_DRAW_MLINE, ID_DRAW_RECT, ID_DRAW_MRECT, ID_DRAW_CIRCLE
        Set picDraw.MouseIcon = imgCur.ListImages("Pen").Picture
    Case ID_DRAW_TEXT
        Set picDraw.MouseIcon = imgCur.ListImages("Text").Picture
    Case ID_DRAW_SEQUENCENUMBER
        Set picDraw.MouseIcon = imgCur.ListImages("Pen").Picture
    End Select
    
End Sub

'################################################################################################################
'   ��;��  ���ݽ���״̬���õ�ǰ�Ļ��ʵĻ�ˢ��
'################################################################################################################
Private Sub SetDrawStyle(hDC As Long)
    Dim bytPenW As Byte
    Dim vBrush As LOGBRUSH
    Dim lngPen As Long, lngBrush As Long
    
    '�����ԭ�л��ʡ���ˢ
    If glngBrush <> 0 Then DeleteObject glngBrush
    If glngPen <> 0 Then DeleteObject glngPen
    
    '��������
    If mlngLineWidthID = ID_DRAW_LINEWIDTH1 Then
        bytPenW = 1
    ElseIf mlngLineWidthID = ID_DRAW_LINEWIDTH2 Then
        bytPenW = 2
    ElseIf mlngLineWidthID = ID_DRAW_LINEWIDTH3 Then
        bytPenW = 3
    ElseIf mlngLineWidthID = ID_DRAW_LINEWIDTH4 Then
        bytPenW = 4
    ElseIf mlngLineWidthID = ID_DRAW_LINEWIDTH5 Then
        bytPenW = 5
    End If
    
    gcurPenWidth = bytPenW '��¼ԭʼ����
    bytPenW = bytPenW * 1
    If bytPenW < 1 Then bytPenW = 1
    
    gcurPenColor = mlngLineColor
    
    If mlngLineStyleID = ID_DRAW_LINECONTINUE Then
        gcurPenStyle = PS_SOLID
        lngPen = CreatePen(PS_SOLID, bytPenW, mlngLineColor)
    ElseIf mlngLineStyleID = ID_DRAW_LINEDOT Then
        gcurPenStyle = PS_DOT
        lngPen = CreatePen(PS_DOT, bytPenW, mlngLineColor)
    ElseIf mlngLineStyleID = ID_DRAW_LINEDASH Then
        gcurPenStyle = PS_DASH
        lngPen = CreatePen(PS_DASH, bytPenW, mlngLineColor)
    ElseIf mlngLineStyleID = ID_DRAW_LINEDASHDOT Then
        gcurPenStyle = PS_DASHDOT
        lngPen = CreatePen(PS_DASHDOT, bytPenW, mlngLineColor)
    ElseIf mlngLineStyleID = ID_DRAW_LINEDASHDOT2 Then
        gcurPenStyle = PS_DASHDOTDOT
        lngPen = CreatePen(PS_DASHDOTDOT, bytPenW, mlngLineColor)
    End If
    glngPen = SelectObject(picDraw.hDC, lngPen)
    
    '��ˢ
    vBrush.lbColor = mlngFillColor
    gcurFillColor = vBrush.lbColor
    If mlngFillStyleID = ID_DRAW_FILLNONE Then
        vBrush.lbStyle = BS_NULL
        gcurFillStyle = -1
    ElseIf mlngFillStyleID = ID_DRAW_FILLALL Then
        vBrush.lbStyle = BS_SOLID
        gcurFillStyle = -2
    Else
        vBrush.lbStyle = BS_HATCHED
        If mlngFillStyleID = ID_DRAW_FILLH Then
            vBrush.lbHatch = HS_HORIZONTAL '====
        ElseIf mlngFillStyleID = ID_DRAW_FILLV Then
            vBrush.lbHatch = HS_VERTICAL '||||
        ElseIf mlngFillStyleID = ID_DRAW_FILLHV Then
            vBrush.lbHatch = HS_CROSS '++++
        ElseIf mlngFillStyleID = ID_DRAW_FILLL Then
            vBrush.lbHatch = HS_FDIAGONAL '\\\\
        ElseIf mlngFillStyleID = ID_DRAW_FILLR Then
            vBrush.lbHatch = HS_BDIAGONAL '////
        ElseIf mlngFillStyleID = ID_DRAW_FILLLR Then
            vBrush.lbHatch = HS_DIAGCROSS 'XXXX
        End If
        gcurFillStyle = vBrush.lbHatch
    End If
    lngBrush = CreateBrushIndirect(vBrush)
    glngBrush = SelectObject(picDraw.hDC, lngBrush)
End Sub


Private Sub txt_Change()
    Dim W As Long, h2 As Long
    Dim lngLines As Long
    
    Call GetFitTxtSize(txt, txt.Text, W, , h2)
    
    If txt.Left + W + 10 <= picDraw.ScaleWidth Then
        txt.Width = W + 10
        picTxt.Left = txt.Left + txt.Width - picTxt.Width / 2
    End If
    
    lngLines = SendMessage(txt.hwnd, EM_GETLINECOUNT, 0, 0)
    txt.Height = lngLines * h2 + 6
    picTxt.Top = txt.Top - picTxt.Height / 2
    mblnModified = True
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim h2 As Long, lngLines As Long
    
    If InStr("'%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub '�Ƿ�
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0: Beep: Exit Sub  '����
    
    If KeyAscii >= 32 Or KeyAscii = 13 Or KeyAscii < 0 Then
        txtTmp.FontSize = txt.FontSize
        txtTmp.FontName = txt.FontName
        txtTmp.FontBold = txt.FontBold
        txtTmp.FontItalic = txt.FontItalic
        txtTmp.FontUnderline = txt.FontUnderline
        txtTmp.FontStrikethru = txt.FontStrikethru
        txtTmp.Width = txt.Width
        txtTmp.Text = Left(txt.Text, txt.SelStart) & IIf(KeyAscii = 13, vbCrLf, Chr(KeyAscii)) & Mid(txt.Text, txt.SelStart + txt.SelLength + 1)
        lngLines = SendMessage(txtTmp.hwnd, EM_GETLINECOUNT, 0, 0)
        Call GetFitTxtSize(txt, "A", , , h2)
        If txt.Top + lngLines * h2 + 6 > picDraw.ScaleHeight Then KeyAscii = 0: Beep
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    If txt.Left + txt.Width > picDraw.ScaleWidth Or txt.Top + txt.Height > picDraw.Height Then
        Cancel = True
        MsgBox "�ı������޷��ڿɼ���Χ����ȫ��ʾ,������ı�λ�û����ݣ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub InitCommandBars()
    If mcPicture.PictureType = EPRFormulaPicture Then Exit Sub '��ʽͼ����˵�
    '##########################################################################################
    Dim cbpPopup As CommandBarPopup     '��ʱ����
    Dim cbpPopupSub As CommandBarPopup  '��ʱ����
    Dim objControl As CommandBarControl                 '�������ؼ�
    Dim objCustControl As CommandBarControlCustom       '�Զ���ؼ�
    Dim Combo As CommandBarComboBox     '������������ؼ�
    
 
    Set mBar��ͼ = cbsThis.Add("��ͼ", xtpBarBottom)
    mBar��ͼ.EnableDocking xtpFlagHideWrap
    mBar��ͼ.ModifyStyle XTP_CBRS_GRIPPER, 0
    With mBar��ͼ.Controls
        Set objControl = .Add(xtpControlButton, ID_DRAW_SELECT, "ѡ�� Ctrl+E")
        objControl.BeginGroup = True
        .Add xtpControlButton, ID_DRAW_MOVE, "�ƶ� Ctrl+M"
        .Add xtpControlButton, ID_DRAW_LINE, "ֱ�� Ctrl+L"
        .Add xtpControlButton, ID_DRAW_MLINE, "���� Ctrl+Z"
        .Add xtpControlButton, ID_DRAW_RECT, "���� Ctrl+R"
        .Add xtpControlButton, ID_DRAW_MRECT, "����� Ctrl+W"
        .Add xtpControlButton, ID_DRAW_CIRCLE, "��Բ Ctrl+C"
        .Add xtpControlButton, ID_DRAW_TEXT, "���� Ctrl+T"
        .Add xtpControlButton, ID_DRAW_SEQUENCENUMBER, "˳���� Ctrl+N"
        
        Set objControl = .Add(xtpControlButton, ID_DRAW_CLEARNUMBERS, "���˳����(&K)")
        objControl.BeginGroup = True
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FILLCOLOR, "�����ɫ")
        cbpPopup.BeginGroup = True
        cbpPopup.CloseSubMenuOnClick = True
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hwnd
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_LINECOLOR, "������ɫ")
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorLineColor.hwnd
        
        Set cbpPopup = .Add(xtpControlSplitButtonPopup, ID_DRAW_FONTCOLOR, "������ɫ")
        Set objCustControl = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hwnd
        
        Set mBar�����ʽ = .Add(xtpControlButtonPopup, ID_DRAW_FILLSTYLE, "���")
        mBar�����ʽ.BeginGroup = True
        mBar�����ʽ.Style = xtpButtonIconAndCaption
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLNONE, "�����"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLALL, "ʵ�����"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLH, "�������"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLV, "�������"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLHV, "�������"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLR, "��б�����"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLL, "��б�����"
        mBar�����ʽ.CommandBar.Controls.Add xtpControlButton, ID_DRAW_FILLLR, "���������"
        
        Set mBar���� = .Add(xtpControlButtonPopup, ID_DRAW_LINESTYLE, "����")
        mBar����.Style = xtpButtonIconAndCaption
        mBar����.CommandBar.SetPopupToolBar True
        mBar����.CommandBar.SetIconSize 80, 8
        mBar����.CommandBar.Width = 80
        mBar����.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINECONTINUE, "ʵ��"
        mBar����.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDOT, "����"
        mBar����.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASH, "����"
        mBar����.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT, "�㻮��"
        mBar����.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEDASHDOT2, "��㻮��"
        
        Set mBar�߿� = .Add(xtpControlButtonPopup, ID_DRAW_LINEWIDTH, "�߿�")
        mBar�߿�.Style = xtpButtonIconAndCaption
        mBar�߿�.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH1, "1�����"
        mBar�߿�.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH2, "2�����"
        mBar�߿�.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH3, "3�����"
        mBar�߿�.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH4, "4�����"
        mBar�߿�.CommandBar.Controls.Add xtpControlButton, ID_DRAW_LINEWIDTH5, "5�����"
        
        Set objControl = .Add(xtpControlButton, ID_DRAW_DELETE, "ɾ��")
        objControl.BeginGroup = True
'        objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, ID_DRAW_RESET, "����")
        objControl.BeginGroup = True
    End With
    DockingRightOf mBar��ͼ, mfrmParent.CommBar(ID_BAR_FORMAT)
    
    '##########################################################################################
    '�����Ļָ�
    ColorFillColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillColor", vbBlue)
    ColorLineColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineColor", vbRed)
    ColorForeColor.COLOR = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "ForeColor", vbBlack)
    mlngFillStyleID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillStyle", ID_DRAW_FILLNONE)
    mlngLineStyleID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineStyle", ID_DRAW_LINECONTINUE)
    mlngLineWidthID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineWidth", ID_DRAW_LINEWIDTH1)
    mlngDrawModeID = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "DrawMode", ID_DRAW_SELECT)
    '##########################################################################################
    'ͼ���ʼ��
    ColorFillColor_pOK
    ColorLineColor_pOK
    ColorForeColor_pOK
End Sub

'################################################################################################################
'## ���ܣ�  ��������A���õ�������B��ͬһ��
'##
'## ������  BarToDock   ������Ĺ�����
'##         BarOnLeft   ��λ����ߵĹ�����
'################################################################################################################
Private Sub DockingRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Top + Bottom) / 2, BarOnLeft.Position
End Sub

Private Sub SaveState()
    '���洰��λ��
    If UserControl.Extender.Visible Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillColor", ColorFillColor.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineColor", ColorLineColor.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "ForeColor", ColorForeColor.COLOR
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "FillStyle", mlngFillStyleID
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineStyle", mlngLineStyleID
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "LineWidth", mlngLineWidthID
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\ucPictureEditor", "DrawMode", mlngDrawModeID
    End If
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Set mBar��ͼ = Nothing
    Set mBar���� = Nothing
    Set mBar�߿� = Nothing
    Set mBar�����ʽ = Nothing
    Set mcPicture = Nothing
    Set PicMarks = Nothing
    Set mobjTable = Nothing
    Set cbsThis = Nothing
    Set mfrmParent = Nothing
End Sub
