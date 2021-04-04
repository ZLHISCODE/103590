VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ctlQueryItem 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   7860
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   360
      ScaleHeight     =   2580
      ScaleWidth      =   1800
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Label lblList 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         BackStyle       =   0  'Transparent
         Caption         =   "超级链接"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   315
         MouseIcon       =   "ctlQueryItem.ctx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   435
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Line ln 
         X1              =   1605
         X2              =   1605
         Y1              =   1515
         Y2              =   2880
      End
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   5475
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlQueryItem.ctx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlQueryItem.ctx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlQueryItem.ctx":0A3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlQueryItem.ctx":0DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ctlQueryItem.ctx":1172
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   90
      ScaleHeight     =   4425
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   645
      Width           =   6780
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1170
         Index           =   0
         Left            =   1830
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   1845
         _cx             =   3254
         _cy             =   2064
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   16761024
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   101
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1110
         Left            =   -20000
         ScaleHeight     =   1080
         ScaleWidth      =   1725
         TabIndex        =   8
         Top             =   645
         Visible         =   0   'False
         Width           =   1755
      End
      Begin zl9NewQuery.ctlPicture picDraw 
         Height          =   1140
         Index           =   0
         Left            =   4875
         TabIndex        =   6
         Top             =   -20000
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   2011
         Border          =   0
      End
      Begin VB.Shape shp 
         BorderColor     =   &H00FFC0C0&
         Height          =   2280
         Index           =   0
         Left            =   300
         Top             =   1950
         Visible         =   0   'False
         Width           =   3960
      End
      Begin VB.Label lblConnect 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "超级链接"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   4665
         MouseIcon       =   "ctlQueryItem.ctx":150C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   990
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTxt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "文本内容"
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   315
         TabIndex        =   3
         Top             =   330
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblReturn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "返回页首△"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   4920
         MouseIcon       =   "ctlQueryItem.ctx":1816
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2175
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   135
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image imgTitle 
         Height          =   240
         Index           =   0
         Left            =   1245
         Stretch         =   -1  'True
         Top             =   300
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "ctlQueryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarMaxVsb As Long
Private mvarMaxHsb As Long
Private mvarValueVsb As Long
Private mvarValueHsb As Long

Private mvarFactWidth As Single
Private mvarFactHeight As Single

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ConnectClick(ByVal PageNo As Long, ByVal OrderNo As Long)
Public Event ChangeNavigator()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event RefreshNavigator(ByVal W As Single, ByVal H As Single)

Private Sub imgTitle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblConnect_Click(Index As Integer)
    RaiseEvent ConnectClick(Val(Split(lblConnect(Index).Tag, ";")(0)), Val(Split(lblConnect(Index).Tag, ";")(1)))
End Sub

Private Sub lblConnect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If lblConnect(Index).ForeColor = &HFF0000 Then
        For i = 1 To lblConnect.UBound
            If lblConnect(i).ForeColor = &HFF& Then
                lblConnect(i).ForeColor = &HFF0000
            End If
        Next
        lblConnect(Index).ForeColor = &HFF&
    End If
    
End Sub

Private Sub lblHeader_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblList_Click(Index As Integer)
    Call GoPageItem(Index)
End Sub

Private Sub lblList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If lblList(Index).ForeColor = &HFF0000 Then
        For i = 1 To lblList.UBound
            If lblList(i).ForeColor = &HFF& Then
                lblList(i).ForeColor = &HFF0000
            End If
        Next
        lblList(Index).ForeColor = &HFF&
    End If
End Sub

Private Sub lblReturn_Click(Index As Integer)
    '
    Call GoPageItem(1)
End Sub

Private Sub lblReturn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If lblReturn(Index).ForeColor = &HFF0000 Then
        For i = 1 To lblReturn.UBound
            If lblReturn(i).ForeColor = &HFF& Then
                lblReturn(i).ForeColor = &HFF0000
            End If
        Next
        lblReturn(Index).ForeColor = &HFF&
    End If
End Sub

Private Sub lblTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 1 To lblConnect.UBound
        If lblConnect(i).ForeColor = &HFF& Then
            lblConnect(i).ForeColor = &HFF0000
        End If
    Next
        
    For i = 1 To lblReturn.UBound
        If lblReturn(i).ForeColor = &HFF& Then
            lblReturn(i).ForeColor = &HFF0000
        End If
    Next
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picDraw_PlayPaint(Index As Integer)
    Call picDraw(Index).ShowPictureByFile(picDraw(Index).Tag)
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long

    For i = 1 To lblList.UBound
        If lblList(i).ForeColor = &HFF& Then
            lblList(i).ForeColor = &HFF0000
        End If
    Next
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    ln.X1 = picList.Width - 15
    ln.X2 = picList.Width - 15
    ln.Y1 = 0
    ln.Y2 = picList.Height
    ln.BorderColor = &HC000C0
End Sub

Private Sub picNavigator_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Call ResizeControl(picList, 0, 0, picList.Width, UserControl.Height)
End Sub

Private Sub MovePageControl(ByVal Mode As Long, ByVal Off As Single)
'功能:移动页面的所有控件
'参数:Abs(Mode)=1为纵向滚动
'     Abs(Mode)<>1为横向滚动

    If Abs(Mode) = 1 Then
        picBack.Top = picBack.Top + Off
    Else
        picBack.Left = picBack.Left + Off
        picBack.Left = IIf(picBack.Left < 0, 0, picBack.Left)
    End If
End Sub

Private Sub GoPageItem(ByVal Item As Long)
    '直接翻到当页的第Item查询项目上
    Dim vPage As Long
    
    '当定位项在显示区域的上方
    If lblHeader(Item).Top < (0 - picBack.Top) Then
    
        vPage = 0 - Int(0 - (0 - picBack.Top - lblHeader(Item).Top + lblHeader(Item).Height) / 600)
        'vPage = Int((0 - picBack.Top - lblHeader(Item).Top + lblHeader(Item).Height) / 600)
        vPage = IIf(vPage < 0, 0, vPage)
        vPage = IIf(mvarValueVsb - vPage < 0, mvarValueVsb, vPage)
        mvarValueVsb = IIf(mvarValueVsb - vPage < 0, 0, mvarValueVsb - vPage)
        'mvarValueVsb = IIf(mvarValueVsb <= 0, 1, mvarValueVsb)
'        Call ChangeNavigator
        RaiseEvent ChangeNavigator
        Call MovePageControl(1, 600 * vPage)
        
        Exit Sub
    End If

    '当定位项是显示区域的下方
    If lblHeader(Item).Top > (UserControl.Height - picBack.Top) Then
        
        '确定要移动的步长
        vPage = Int((lblHeader(Item).Top + picBack.Top) / 600)
        vPage = IIf(vPage + mvarValueVsb > mvarMaxVsb, mvarMaxVsb - mvarValueVsb, vPage)
        
        mvarValueVsb = IIf(mvarValueVsb + vPage > mvarMaxVsb, mvarMaxVsb, mvarValueVsb + vPage)
        RaiseEvent ChangeNavigator
        Call MovePageControl(-1, -600 * vPage)

    End If

    
End Sub

Private Sub GridWallPaper()
'功能:设置表格控件的背景图片为表格背后的区域图片
    Dim i As Long
    
    On Error Resume Next
    For i = 1 To vsf.UBound
        pic.Cls
        pic.Width = vsf(i).Width
        pic.Height = vsf(i).Height
        pic.PaintPicture picBack.Image, 0, 0, pic.Width, pic.Height, vsf(i).Left, vsf(i).Top, vsf(i).Width, vsf(i).Height
        Set vsf(i).WallPaper = pic.Image
    Next
End Sub

Private Function AdjustTxtHeight(TxtObj As Label) As Single
    Dim strTxt As String
    Dim strTmp As String
    Dim intPos As Long
    
    strTxt = TxtObj.Caption
    picBack.FontName = TxtObj.FontName
    picBack.FontItalic = TxtObj.FontItalic
    picBack.FontBold = TxtObj.FontBold
    picBack.FontSize = TxtObj.FontSize
    picBack.FontStrikethru = TxtObj.FontStrikethru
    picBack.FontUnderline = TxtObj.FontUnderline
    
    intPos = InStr(strTxt, Chr(13) & Chr(10))
    If intPos = 0 Then AdjustTxtHeight = AdjustTxtHeight + picBack.TextHeight("测试行高") * (0 - Int(0 - picBack.TextWidth(strTxt) / TxtObj.Width))
    While intPos > 0
        strTmp = Mid(strTxt, 1, intPos - 1)
        strTxt = Mid(strTxt, intPos + 2)
        
        If strTmp <> "" Then
            AdjustTxtHeight = AdjustTxtHeight + picBack.TextHeight("测试行高") * (0 - Int(0 - picBack.TextWidth(strTmp) / TxtObj.Width))
        Else
            AdjustTxtHeight = AdjustTxtHeight + picBack.TextHeight("测试行高")
        End If
        intPos = InStr(strTxt, Chr(13) & Chr(10))
    Wend
    
End Function

Private Sub UsrCmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub vsf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'--------------------------------------------------------------------------------------------
'以下是属性及方法:
'--------------------------------------------------------------------------------------------

Public Property Get NextTxtIndex() As Long
    NextTxtIndex = lblTxt.UBound + 1
End Property

Public Property Get NextPicIndex() As Long
    NextPicIndex = picDraw.UBound + 1
End Property

Public Property Get NextVsfIndex() As Long
    NextVsfIndex = vsf.UBound + 1
End Property

Public Property Get NextConnectIndex() As Long
    NextConnectIndex = lblConnect.UBound + 1
End Property

Public Property Let CellFontName(ByVal Index As Long, ByVal vData As String)
    vsf(Index).CellFontName = vData
End Property

Public Property Let CellFontBold(ByVal Index As Long, ByVal vData As Boolean)
    vsf(Index).CellFontBold = vData
End Property

Public Property Let CellFontSize(ByVal Index As Long, ByVal vData As Single)
    vsf(Index).CellFontSize = vData
End Property

Public Property Let CellFontItalic(ByVal Index As Long, ByVal vData As Boolean)
    vsf(Index).CellFontItalic = vData
End Property

Public Property Let CellFontStrikethru(ByVal Index As Long, ByVal vData As Boolean)
    vsf(Index).CellFontStrikethru = vData
End Property

Public Property Let CellFontUnderline(ByVal Index As Long, ByVal vData As Boolean)
    vsf(Index).CellFontUnderline = vData
End Property

Public Property Let CellAlignment(ByVal Index As Long, ByVal vData As Byte)
    vsf(Index).CellAlignment = vData
End Property

Public Property Let CellForeColor(ByVal Index As Long, ByVal vData As Long)
    vsf(Index).CellForeColor = vData
End Property

Public Property Let TextMatrix(ByVal Index As Long, ByVal Row As Long, ByVal Col As Long, ByVal vData As String)
    vsf(Index).TextMatrix(Row, Col) = vData
End Property

Public Property Get CellFontName(ByVal Index As Long) As String
    CellFontName = vsf(Index).CellFontName
End Property

Public Property Get CellFontBold(ByVal Index As Long) As Boolean
    CellFontBold = vsf(Index).CellFontBold
End Property

Public Property Get CellFontSize(ByVal Index As Long) As Single
    CellFontSize = vsf(Index).CellFontSize
End Property

Public Property Get CellFontItalic(ByVal Index As Long) As Boolean
    CellFontItalic = vsf(Index).CellFontItalic
End Property

Public Property Get CellFontStrikethru(ByVal Index As Long) As Boolean
    CellFontStrikethru = vsf(Index).CellFontStrikethru
End Property

Public Property Get CellFontUnderline(ByVal Index As Long) As Boolean
    CellFontUnderline = vsf(Index).CellFontUnderline
End Property

Public Property Get CellAlignment(ByVal Index As Long) As Byte
    CellAlignment = vsf(Index).CellAlignment
End Property

Public Property Get CellForeColor(ByVal Index As Long) As Long
    CellForeColor = vsf(Index).CellForeColor
End Property

Public Property Get TextMatrix(ByVal Index As Long, ByVal Row As Long, ByVal Col As Long) As String
    TextMatrix = vsf(Index).TextMatrix(Row, Col)
End Property

Public Property Let Row(ByVal Index As Long, ByVal vData As Long)
    vsf(Index).Row = vData
End Property

Public Property Let Col(ByVal Index As Long, ByVal vData As Long)
    vsf(Index).Col = vData
End Property

Public Property Get Row(ByVal Index As Long) As Long
    Row = vsf(Index).Row
End Property

Public Property Get Col(ByVal Index As Long) As Long
    Col = vsf(Index).Col
End Property

Public Property Let ClientVisible(vData As Boolean)
    picBack.Visible = vData
'    picNavigator.Visible = vData
End Property

Public Sub InitLoad()
'功能:查询项初始化准备工作
        
    picList.Visible = False
    
    picBack.Left = 0
    picBack.Top = 0
    
    Call GridWallPaper
    
End Sub

Public Sub ResizePage(ByVal W As Single, ByVal H As Single)
    
    On Error Resume Next
    
    picBack.Width = IIf(W < UserControl.Width, UserControl.Width, W)
    picBack.Height = IIf(H < UserControl.Height, UserControl.Height, H)
    
        
    mvarFactWidth = picBack.Width
    mvarFactHeight = picBack.Height
    
    picBack.Width = picBack.Width + 1200
    picBack.Height = picBack.Height + 1200
End Sub

Public Sub BackPicture(ByVal vData As String, ByVal W As Single, ByVal H As Single)
    Dim vCountX As Long
    Dim vCountY As Long
    Dim i As Long
    Dim j As Long
    Dim X1 As Single
    Dim Y1 As Single
    Dim picObj As StdPicture
    
    On Error Resume Next
    vCountX = Int(picBack.Width / W) + 1
    vCountY = Int(picBack.Height / H) + 1
    
    Set picObj = VB.LoadPicture(vData)
    For j = 1 To vCountY
        For i = 1 To vCountX
            X1 = (i - 1) * W
            Y1 = (j - 1) * H
            picBack.PaintPicture picObj, X1, Y1, W, H
        Next
    Next
    
End Sub

Public Sub GoPageItemByOrder(ByVal Order As Long)
    Dim i As Long
'    If UsrCmd(3).Visible Then
        For i = 1 To lblList.UBound
            If Val(Split(lblList(i).Tag, ";")(1)) = Order Then
                Call GoPageItem(i)
                Exit Sub
            End If
        Next
'    End If
End Sub

Public Sub AddPageItemTitle(ByVal Index As Long, ByVal Y As Single, ByVal Title As String, ByVal Color As Long, ByVal TitleFont As StdFont, PicName As String, ByVal PageNo As Long, ByVal OrderNo As Long, ObjWidth As Single, ObjHeight As Single, Optional ByVal blnVisible As Boolean = True, Optional ByVal Alignment As Byte = 0)
        
    '添加标题文字内容
    Load lblHeader(Index)
    lblHeader(Index).ZOrder
    lblHeader(Index).Tag = PageNo & ";" & OrderNo
    lblHeader(Index).Caption = Title
    lblHeader(Index).ForeColor = Color
    lblHeader(Index).FontName = TitleFont.Name
    lblHeader(Index).FontSize = TitleFont.Size
    lblHeader(Index).FontBold = TitleFont.Bold
    lblHeader(Index).FontItalic = TitleFont.Italic
    
    lblHeader(Index).Left = IIf(Alignment = 0, IIf(Dir(PicName) <> "" And PicName <> "", 330, 90), IIf(Alignment = 1, UserControl.Width - lblHeader(Index).Width - 60, (UserControl.Width - lblHeader(Index).Width) / 2))
    lblHeader(Index).Top = Y
    
    lblHeader(Index).Visible = blnVisible
    
    '添加标题图标
    Load imgTitle(Index)
    imgTitle(Index).ZOrder
    If Dir(PicName) <> "" Then Set imgTitle(Index).Picture = LoadPicture(PicName)
    imgTitle(Index).Tag = Alignment
    imgTitle(Index).Top = Y
    imgTitle(Index).Left = lblHeader(Index).Left - imgTitle(Index).Width - 15
    imgTitle(Index).Visible = blnVisible
    
    '添加标题目录
    Load lblList(Index)
    lblList(Index).ZOrder
    
    lblList(Index).Caption = lblHeader(Index).Caption
    lblList(Index).ToolTipText = lblHeader(Index).Caption
    
    lblList(Index).Tag = lblHeader(Index).Tag
    lblList(Index).Left = 120
    lblList(Index).Top = 150 * Index + lblList(Index).Height * (Index - 1)
    lblList(Index).Visible = True
    
    
    '返回的参数
    ObjWidth = lblHeader(Index).Width
    ObjHeight = lblHeader(Index).Height
    
End Sub

Public Sub AddReturnFlag(ByVal X As Single, ByVal Y As Single, ObjHeight As Single)
    Dim Index As Long
    
    Index = lblReturn.UBound + 1
    Load lblReturn(Index)
    lblReturn(Index).ZOrder
    lblReturn(Index).Top = Y
    lblReturn(Index).Left = UserControl.Width - lblReturn(Index).Width - 60
    
    ObjHeight = lblReturn(Index).Height
    
    lblReturn(Index).Visible = True
End Sub

Public Sub AddPageItemIcon(ByVal Index As Long, ByVal Y As Single, PicName As String, Optional ByVal blnVisible As Boolean = True)
    '添加标题图标
    Load imgTitle(Index)
    
    imgTitle(Index).ZOrder
    imgTitle(Index).Top = Y
    imgTitle(Index).Left = 60
    imgTitle(Index).Width = 180
    
    On Error Resume Next
    Set imgTitle(Index).Picture = LoadPicture(App.Path & "\图形\" & PicName & ".ico")
    
    imgTitle(Index).Visible = blnVisible
End Sub

Public Sub AddPageItemTxt(ByVal Index As Long, ByVal X As Single, ByVal Y As Single, Text As String, strFont As String, ObjWidth As Single, ObjHeight As Single)
    Load lblTxt(Index)
    
    lblTxt(Index).ZOrder
    lblTxt(Index).Left = X
    lblTxt(Index).Top = Y
    lblTxt(Index).Caption = Text
            
    On Error Resume Next
    '有可能字体在某些计算机上不支持，此时跳过
    
    lblTxt(Index).FontName = Split(strFont, ";")(0)
    lblTxt(Index).FontSize = Val(Split(strFont, ";")(1))
    lblTxt(Index).FontBold = Val(Split(strFont, ";")(2))
    lblTxt(Index).FontItalic = Val(Split(strFont, ";")(3))
    
    On Error GoTo errHand
    
    lblTxt(Index).ForeColor = Val(Split(strFont, ";")(4))
    lblTxt(Index).Visible = True
    
    lblTxt(Index).Width = ObjWidth
    lblTxt(Index).Height = AdjustTxtHeight(lblTxt(Index))
    ObjHeight = lblTxt(Index).Height
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub AddPageItemPic(ByVal Index As Long, ByVal Alignment As Byte, ByVal Y As Single, PicName As String, ObjWidth As Single, ObjHeight As Single, ByVal W As Single, ByVal H As Single)
    Load picDraw(Index)
            
    picDraw(Index).ZOrder
    picDraw(Index).Left = IIf(Alignment = 0, 120, IIf(Alignment = 1, UserControl.Width - W - 60, (UserControl.Width - W) / 2))
    picDraw(Index).Top = Y
    picDraw(Index).Border = 0
    picDraw(Index).AutoSize = True
    
    On Error GoTo errHand
            
    picDraw(Index).Tag = PicName
    picDraw(Index).Visible = True
    Call picDraw(Index).ShowPictureByFile(PicName, True, W, H)
        
    ObjWidth = picDraw(Index).Width
    ObjHeight = picDraw(Index).Height
    
    Exit Sub
errHand:
End Sub

Public Sub AddPageItemGrid(ByVal Index As Long, ByVal Alignment As Byte, ByVal Y As Single, ByVal Rows As Long, ByVal Cols As Long, ByVal RowHeights As String, ByVal ColWidths As String, ByVal MergeRows As String, ByVal MergeCols As String, ObjWidth As Single, ObjHeight As Single)
    '功能:增加一个用户表格,并完成表格的行列数、行高度、列宽度以及合并行列
    Dim i As Long
    Dim strTmp As String
    Dim intPos As Long
    Dim svrRow As Long
    Dim svrCol As Long
    
    Load vsf(Index)
        
    vsf(Index).ZOrder
                    
    vsf(Index).Rows = Rows
    vsf(Index).Cols = Cols
    
    On Error Resume Next
    
    '1.设置每行的行高度
    For i = 0 To vsf(Index).Rows - 1
        vsf(Index).RowHeight(i) = 300
    Next
    For i = 0 To vsf(Index).Rows - 1
        vsf(Index).RowHeight(i) = Split(RowHeights, ";")(i)
    Next
    
    '2.设置每列的列宽度
    For i = 0 To vsf(Index).Cols - 1
        vsf(Index).ColWidth(i) = 1200
    Next
    For i = 0 To vsf(Index).Cols - 1
        vsf(Index).ColWidth(i) = Split(ColWidths, ";")(i)
    Next
    
    '3.设置要合并的行项
    strTmp = IIf(MergeRows <> "", MergeRows & ";", "")
    intPos = InStr(strTmp, ";")
    While intPos > 0
        vsf(Index).MergeRow(Val(Mid(strTmp, 1, intPos - 1)) - 1) = True
        strTmp = Mid(strTmp, intPos + 1)
        intPos = InStr(strTmp, ";")
    Wend
    
    '4.设置要合并的列项
    strTmp = IIf(MergeCols <> "", MergeCols & ";", "")
    intPos = InStr(strTmp, ";")
    While intPos > 0
        vsf(Index).MergeCol(Val(Mid(strTmp, 1, intPos - 1)) - 1) = True
        strTmp = Mid(strTmp, intPos + 1)
        intPos = InStr(strTmp, ";")
    Wend
    
    '5.调整表格的宽度和高度,以使不用滚动就能查看到
    svrRow = vsf(Index).Row
    svrCol = vsf(Index).Col
    vsf(Index).Width = 32760
    vsf(Index).Height = 32760
    vsf(Index).Row = vsf(Index).Rows - 1
    vsf(Index).Col = vsf(Index).Cols - 1
    vsf(Index).Width = vsf(Index).CellLeft + vsf(Index).CellWidth
    vsf(Index).Height = vsf(Index).CellTop + vsf(Index).CellHeight
    vsf(Index).Row = svrRow
    vsf(Index).Col = svrCol
    
    vsf(Index).Left = IIf(Alignment = 0, 120, IIf(Alignment = 1, UserControl.Width - vsf(Index).Width - 60, (UserControl.Width - vsf(Index).Width) / 2))
    vsf(Index).Top = Y
    
    Load shp(Index)
    shp(Index).ZOrder
    shp(Index).Left = vsf(Index).Left - 15
    shp(Index).Top = vsf(Index).Top - 15
    shp(Index).Width = vsf(Index).Width + 15
    shp(Index).Height = vsf(Index).Height + 15
    shp(Index).Visible = True
    
    On Error GoTo 0
    
    ObjWidth = vsf(Index).Width
    ObjHeight = vsf(Index).Height
    
    vsf(Index).Visible = True
    Exit Sub
errHand:
End Sub

Public Sub AddPageItemConnect(ByVal Index As Long, ByVal X As Single, ByVal Y As Single, Name As String, ByVal ToPage As Long, ByVal ToOrder As Long, ObjWidth As Single, ObjHeight As Single)
    Load lblConnect(Index)
        
    lblConnect(Index).ZOrder
    lblConnect(Index).Left = X
    lblConnect(Index).Top = Y
    lblConnect(Index).AutoSize = True
    lblConnect(Index).Caption = Name
    lblConnect(Index).Tag = ToPage & ";" & ToOrder
    lblConnect(Index).Visible = True
    
    ObjWidth = lblConnect(Index).Width
    ObjHeight = lblConnect(Index).Height
    
End Sub

Public Sub ClearAllPageItem()
    '清除页内项目的所有动态产生的控件
    Dim i As Long
    
    For i = lblHeader.UBound To 1 Step -1
        Unload lblHeader(i)
    Next
    
    For i = imgTitle.UBound To 1 Step -1
        Unload imgTitle(i)
    Next
    
    For i = lblTxt.UBound To 1 Step -1
        Unload lblTxt(i)
    Next
    
    For i = picDraw.UBound To 1 Step -1
        Unload picDraw(i)
    Next
    
    For i = vsf.UBound To 1 Step -1
        Unload vsf(i)
        Unload shp(i)
    Next
    
    For i = lblConnect.UBound To 1 Step -1
        Unload lblConnect(i)
    Next
    
    For i = lblList.UBound To 1 Step -1
        Unload lblList(i)
    Next
        
    For i = lblReturn.UBound To 1 Step -1
        Unload lblReturn(i)
    Next
    
    picList.Visible = False
    picBack.Cls
    
End Sub

Public Property Let MaxVsb(ByVal vData As Long)
    mvarMaxVsb = vData
End Property

Public Property Get MaxVsb() As Long
    MaxVsb = mvarMaxVsb
End Property

Public Property Let ValueVsb(ByVal vData As Long)
    mvarValueVsb = vData
End Property

Public Property Get ValueVsb() As Long
    ValueVsb = mvarValueVsb
End Property

Public Property Let MaxHsb(ByVal vData As Long)
    mvarMaxHsb = vData
End Property

Public Property Get MaxHsb() As Long
    MaxHsb = mvarMaxHsb
End Property

Public Property Let ValueHsb(ByVal vData As Long)
    mvarValueHsb = vData
End Property

Public Property Get ValueHsb() As Long
    ValueHsb = mvarValueHsb
End Property

Public Property Let FactWidth(ByVal vData As Single)
    mvarFactWidth = vData
End Property

Public Property Let FactHeight(ByVal vData As Long)
    mvarFactHeight = vData
End Property

Public Property Get FactHeight() As Long
    FactHeight = mvarFactHeight
End Property

Public Sub TurnToNextPage()
    If mvarValueVsb + 1 > mvarMaxVsb Then Exit Sub
    mvarValueVsb = mvarValueVsb + 1
    RaiseEvent ChangeNavigator
    Call MovePageControl(-1, -600)
End Sub

Public Sub TurnToLastPage()
    If mvarValueVsb - 1 < 0 Then Exit Sub
    mvarValueVsb = mvarValueVsb - 1
    RaiseEvent ChangeNavigator
    Call MovePageControl(1, 600)
End Sub

Public Sub TurnToLeftPage()
    If mvarValueHsb - 1 < 0 Then Exit Sub
    mvarValueHsb = mvarValueHsb - 1
    RaiseEvent ChangeNavigator
    Call MovePageControl(2, 600)
End Sub

Public Sub TurnToRightPage()
    If mvarValueHsb + 1 > mvarMaxHsb Then Exit Sub
    mvarValueHsb = mvarValueHsb + 1
    RaiseEvent ChangeNavigator
    Call MovePageControl(-2, -600)
End Sub

Public Sub ShowTreeList()
    Dim svrVsb As Long
    Dim svrHsb As Long
    Dim i As Integer
    
    picList.Visible = IIf(picList.Visible, False, True)
    svrVsb = mvarValueVsb
    svrHsb = mvarValueHsb
    If picList.Visible Then
        mvarValueHsb = 0
        svrHsb = 0
        picList.PaintPicture picBack.Image, 0, 0, picList.Width, picList.Height, 0, 0, picList.Width, picList.Height
        Call MovePageControl(2, 1800)
        RaiseEvent RefreshNavigator(mvarFactWidth + 1800, mvarFactHeight)
    Else
        Call MovePageControl(-2, -1800)
        RaiseEvent RefreshNavigator(mvarFactWidth - 1800, mvarFactHeight)
    End If
    
    mvarValueVsb = svrVsb
    mvarValueHsb = svrHsb
    RaiseEvent ChangeNavigator
            
    If svrHsb > 0 And picList.Visible = False Then
        For i = 1 To svrHsb
            Call TurnToRightPage
        Next
    End If
End Sub


Public Property Let Enabled(ByVal vData As Boolean)
    UserControl.Enabled = vData
End Property
