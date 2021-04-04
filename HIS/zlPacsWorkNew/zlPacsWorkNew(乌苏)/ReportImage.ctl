VERSION 5.00
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.UserControl ReportImage 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   2610
   ScaleWidth      =   7605
   ToolboxBitmap   =   "ReportImage.ctx":0000
   Begin VB.PictureBox picMiniImage 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.VScrollBar vscrollMini 
         Height          =   1455
         Left            =   6960
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
      Begin DicomObjects.DicomViewer dcmMiniImage 
         Height          =   1455
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6855
         _Version        =   262147
         _ExtentX        =   12091
         _ExtentY        =   2566
         _StockProps     =   35
         BackColor       =   4210752
      End
   End
End
Attribute VB_Name = "ReportImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const M_STR_SELECT_TAG As String = "SELECT"
Private Const M_STR_BORDER_TAG As String = "BORDER"


Private mintShowPhotoCount As Integer


Private mlngAdviceID As Long
Private mblnMoved As Boolean

Private mobjSelectedImg As DicomImage

Private mblnEnable As Boolean

Public Event SelectedChange(ByVal lngSelectedIndex As Long, ByVal blnSelected As Boolean)




Property Get SelectTag() As String
    SelectTag = M_STR_SELECT_TAG
End Property

Property Get BorderTag() As String
    BorderTag = M_STR_BORDER_TAG
End Property





Property Get Enable() As Boolean
    Enable = mblnEnable
End Property

Property Let Enable(value As Boolean)
    mblnEnable = value
End Property




Property Get dcmViewer() As DicomViewer
    Set dcmViewer = dcmMiniImage
End Property




Property Get ItemSelected(Index As Long) As Boolean
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    ItemSelected = False
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            ItemSelected = Not objLabs(i).Transparent
            Exit Property
        End If
    Next i
End Property

Property Let ItemSelected(Index As Long, value As Boolean)
    Dim i As Long
    Dim objLabs As DicomLabels
    
    Set objLabs = dcmMiniImage.Images(Index).Labels
    
    For i = 1 To objLabs.Count
        If objLabs(i).Tag = M_STR_SELECT_TAG Then
            objLabs(i).Transparent = Not value
            Call dcmMiniImage.Images(Index).Refresh(False)
            
            Exit Property
        End If
    Next i
End Property






Property Get CellSpacing() As Long
    CellSpacing = dcmMiniImage.CellSpacing
End Property

Property Let CellSpacing(value As Long)
    dcmMiniImage.CellSpacing = value
End Property




Property Get ShowPhotoCount() As Integer
    ShowPhotoCount = mintShowPhotoCount
End Property

Property Let ShowPhotoCount(value As Integer)
    mintShowPhotoCount = value
End Property





Property Get BackColor() As OLE_COLOR
    BackColor = dcmMiniImage.BackColour
End Property


Property Let BackColor(value As OLE_COLOR)
    dcmMiniImage.BackColour = value
End Property







Public Function SelectedCount() As Long
'获取选择的图像数量
    Dim i As Long
    Dim j As Long
    Dim lngCount As Long
    Dim objLabs As DicomLabels
    
    
    lngCount = 0
    For i = 1 To dcmMiniImage.Images.Count
        Set objLabs = dcmMiniImage.Images(i).Labels
        
        For j = 1 To objLabs.Count
            If objLabs(j).Tag = M_STR_SELECT_TAG Then
                If Not objLabs(j).Transparent Then lngCount = lngCount + 1
                Exit For
            End If
        Next j
    Next i
    
    SelectedCount = lngCount
End Function



Private Sub dcmMiniImage_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errHandle
    Dim i As Long
    Dim lngImgIndex As Long
    Dim objLabs As DicomLabels
    
        
    lngImgIndex = dcmMiniImage.ImageIndex(X, Y)
    
    If lngImgIndex > 0 And lngImgIndex <= dcmMiniImage.Images.Count Then
        
        If Not (mobjSelectedImg Is Nothing) Then mobjSelectedImg.BorderWidth = 0
        
        Set mobjSelectedImg = dcmMiniImage.Images(lngImgIndex)
        
        mobjSelectedImg.BorderWidth = 2
        mobjSelectedImg.BorderColour = vbRed
        
        If mblnEnable Then
        
            '设置选择框状态
            Set objLabs = dcmMiniImage.LabelHits(X, Y, False, True, True)
            For i = 1 To objLabs.Count
                If objLabs(i).Tag = M_STR_SELECT_TAG Then
                    objLabs(i).Transparent = Not objLabs(i).Transparent
                    
                    '触发选择改变事件
                    RaiseEvent SelectedChange(i, Not objLabs(i).Transparent)
                    
                    Exit For
                End If
            Next i
        
        End If
        
        Call mobjSelectedImg.Refresh(False)
    End If
errHandle:
End Sub

Private Sub UserControl_Initialize()
    Set mobjSelectedImg = Nothing
End Sub

Private Sub UserControl_InitProperties()
    mintShowPhotoCount = 6
    dcmMiniImage.CellSpacing = 3
    mblnEnable = True
End Sub







Public Sub ClearSelected()
'清除选择
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ItemSelected(i) = False
    Next i
End Sub



Public Sub SelectedAll()
'全选
    Dim i As Long
    
    For i = 1 To dcmMiniImage.Images.Count
        ItemSelected(i) = True
    Next i
End Sub


Public Sub ReInit()
'重新初始化
    Call dcmMiniImage.Images.Clear
    
    mlngAdviceID = -1
    mblnMoved = False
    
    Set mobjSelectedImg = Nothing
End Sub


Private Sub UserControl_Resize()
    picMiniImage.Left = 10
    picMiniImage.Top = 10
    picMiniImage.Width = Width - 20
    picMiniImage.Height = Height - 20
    
    dcmMiniImage.Left = 0
    dcmMiniImage.Top = 0
    dcmMiniImage.Width = picMiniImage.Width
    dcmMiniImage.Height = picMiniImage.Height - 50
    
    Call AdjustDicomViewerLayout
End Sub


Public Sub LoadReportImages(ByVal lngAdviceID As Long, ByVal blnMoved As Boolean, owner As Form)
'    If mlngAdviceID = lngAdviceID Then Exit Sub
    
    mlngAdviceID = lngAdviceID
    mblnMoved = blnMoved
    
    Call dcmMiniImage.Images.Clear
    
    '读取图像
    Call GetAllImages(owner, dcmMiniImage, blnMoved, 1, lngAdviceID, "", 100, mintShowPhotoCount)
    
    '绘制选择框
    Call DrawImageSelectBorder(dcmMiniImage)
    
    '配置滚动条
    Call subDispScroll
    
    If dcmMiniImage.Images.Count > 0 Then Set mobjSelectedImg = dcmMiniImage.Images(1)
End Sub


Private Sub AdjustDicomViewerLayout()
'------------------------------------------------
'功能：将图像添加到缩略图dcmMiniature中
'参数：img－－输入的DICOM图像
'返回：无，直接将图像添加到缩略图dcmMiniature中
'------------------------------------------------
    Dim iRows As Integer
    Dim iCols As Integer
    
    '自动对图像做布局
    '计算缩略图的图像布局
    If mintShowPhotoCount < dcmMiniImage.Images.Count Then
        ResizeRegion mintShowPhotoCount, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    Else
        ResizeRegion dcmMiniImage.Images.Count, dcmMiniImage.Width, dcmMiniImage.Height, iRows, iCols
    End If
    
    dcmMiniImage.MultiColumns = iCols
    dcmMiniImage.MultiRows = iRows
    
    '处理滚动条
    If vscrollMini.Visible = True Then
        dcmMiniImage.Width = picMiniImage.Width - vscrollMini.Width - 30
        vscrollMini.Height = dcmMiniImage.Height
        vscrollMini.Left = picMiniImage.Width - vscrollMini.Width - 30
    End If
End Sub




Public Sub subDispScroll()
'------------------------------------------------
'功能：自动判断是否需要显示或隐藏滚动条
'返回：无，直接显示或隐藏滚动条。
'------------------------------------------------
    Dim ii As Integer
    
    If dcmMiniImage.Images.Count > dcmMiniImage.MultiColumns * dcmMiniImage.MultiRows Then       '图像总数大于显示数，显示滚动条
        '摆放滚动条位置，并显示滚动条
        vscrollMini.Move dcmMiniImage.Width - vscrollMini.Width, dcmMiniImage.Top, vscrollMini.Width, dcmMiniImage.Height
        vscrollMini.Visible = True
        vscrollMini.ZOrder
        vscrollMini.Refresh
        
        ''''''''''''''''''[关于滚动条需要单独仔细分析]'''''''''''''''''''''''''
        vscrollMini.Min = 1
        vscrollMini.Max = IIf(dcmMiniImage.Images.Count Mod dcmMiniImage.MultiColumns = 0, _
                                Fix(dcmMiniImage.Images.Count / dcmMiniImage.MultiColumns), _
                                Fix(dcmMiniImage.Images.Count / dcmMiniImage.MultiColumns) + 1)
        
        vscrollMini.LargeChange = 1
        
        If GetImageRow(dcmMiniImage.CurrentIndex) >= vscrollMini.Max Then
            vscrollMini.value = vscrollMini.Max
            dcmMiniImage.CurrentIndex = (vscrollMini.Max - 1) * dcmMiniImage.MultiColumns + 1
        Else
            vscrollMini.value = GetImageRow(dcmMiniImage.CurrentIndex)
        End If
        
    Else    '图像数少于可显示数，隐藏滚动条
        vscrollMini.Visible = False
    End If
    
    If vscrollMini.Visible = True Then
        dcmMiniImage.Width = picMiniImage.Width - vscrollMini.Width - 30
        
        vscrollMini.Height = dcmMiniImage.Height - 40
        vscrollMini.Left = picMiniImage.Width - vscrollMini.Width - 30
    Else
        dcmMiniImage.Width = picMiniImage.Width
    End If
End Sub


Private Function GetImageRow(ByVal lngImageIndex As Long) As Integer
'取得当前所在行
    GetImageRow = CInt(lngImageIndex / dcmMiniImage.MultiColumns) + 1
End Function

Private Sub UserControl_Show()
    Call UserControl_Resize
End Sub



Private Sub vscrollMini_Change()
On Error GoTo errHandle
    dcmMiniImage.CurrentIndex = (vscrollMini.value - 1) * dcmMiniImage.MultiColumns + 1
errHandle:
End Sub



Private Sub DrawImageSelectBorder(dcmViewer As DicomViewer)
    Dim i As Long
    
    Dim lSelect As DicomLabel
    Dim lBorder As DicomLabel

    
    '循环每一个图像，画标注
    For i = 1 To dcmViewer.Images.Count
        Call dcmViewer.Images(i).Labels.Clear
        
        Set lBorder = New DicomLabel

        lBorder.LabelType = 2            '边框
        lBorder.Width = 1000
        lBorder.Height = 1000
        lBorder.Left = 0
        lBorder.Top = 0
        lBorder.LineWidth = 2


        lBorder.ForeColour = vbYellow
        lBorder.BackColour = vbYellow


        lBorder.Transparent = True
        lBorder.ScaleWithCell = True
        lBorder.Tag = M_STR_BORDER_TAG

        lBorder.Visible = True
        dcmViewer.Images(i).Labels.Add lBorder
        

    
    
        Set lSelect = New DicomLabel
        
        lSelect.LabelType = 2            '矩形
        lSelect.Width = 18
        lSelect.Height = 18
        lSelect.Left = 1
        lSelect.Top = 1
        lSelect.LineWidth = 2
        
        lSelect.ForeColour = vbYellow
        lSelect.BackColour = vbRed
        
                
        lSelect.Transparent = True
        lSelect.ScaleWithCell = False
        lSelect.ImageTied = False

        lSelect.Tag = M_STR_SELECT_TAG
        
        lSelect.Visible = True
        dcmViewer.Images(i).Labels.Add lSelect
        
        dcmViewer.Images(1).BorderStyle = vbRed
    Next i
End Sub





Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo errHandle
    
    dcmMiniImage.CellSpacing = PropBag.ReadProperty("CellSpacing", 3)
    mintShowPhotoCount = PropBag.ReadProperty("ShowPhotoCount", 6)
    dcmMiniImage.BackColour = PropBag.ReadProperty("BackColor", vbBlack)
    mblnEnable = PropBag.ReadProperty("Enable", True)
    
errHandle:
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo errHandle

    Call PropBag.WriteProperty("CellSpacing", dcmMiniImage.CellSpacing, 3)
    Call PropBag.WriteProperty("ShowPhotoCount", mintShowPhotoCount, 6)
    Call PropBag.WriteProperty("BackColor", dcmMiniImage.BackColour, vbBlack)
    Call PropBag.WriteProperty("Enable", mblnEnable, True)
    
errHandle:
End Sub
