VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################################
'##模 块 名：cEPRPicture.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：内部图片（文档中的图片和表格中的图片），表格中的图片与某单元格位置相关
'##版    本：
'###############################################################################################

Option Explicit
Public Enum EPRPictureTypeEnum
    EPROutPicture = 0      '外部图
    EPRMarkedPicture = 1   '标记图
    EPRInnerPicture = 2    '内部图
    EPRSignPicture = 3     '签名图
    EPRFormulaPicture = 4  '公式图
    EPRPatiSign = 5        '患者签名
End Enum

Private mvarKey As Long                     '集合中的关键字，等于Key，不是通过集合Add方法生成的图片则 Key＝0
Private mvarID As Double                      '内容图形的ID
Private mvar文件ID As Long                  '文件ID
Private mvar父ID As Double                    '图片所属父ID
Private mvar对象序号 As Double                '对象序号
Private mvar保留对象 As Boolean             '保留对象
Private mvar开始版 As Long                  '开始版，专用于“电子病历记录”
Private mvar终止版 As Long                  '终止版，专用于“电子病历记录”
Private mvar是否换行 As Boolean             '是否换行
Private mvar内容文本 As String              '公式图片的结构化数据
                                            '形式为：类型|数据。月经史 1|前辍|分子|分母|后辍|字号; 牙齿 2(恒牙)/3(乳牙)|左上|右上|左下|右下|字号; 胎心位置 4|上方|下方|左方|右方|字号
                                            
Private mvarRow As Long                     '绑定的行
Private mvarCol As Long                     '绑定的列
Private mvarPictureType As EPRPictureTypeEnum   '图片类型
Private mvarMarkedPicID As Long             '内容图形ID
Private mvarOrigPic As StdPicture           '原始图片
Private mvarPicMarks As cPicMarks           '标记集合
Private mvarLeft As Long                    '图片左上角 X 坐标（参照绑定单元左上角坐标）
Private mvarTop As Long                     '图片左上角 Y 坐标（参照绑定单元左上角坐标）
Private mvarOrigWidth As Long               '图片原始宽度
Private mvarOrigHeight As Long              '图片原始高度
Private mvarWidth As Long                   '图片最终宽度
Private mvarHeight As Long                  '图片最终高度
Private mvarZoomFactor As Double            '缩放比例
Private mvarZOrder As Long                  'Z轴顺序号。从上到下依次是：0、1、2、3......
Private mvarPicName  As String              '检查报告图形的名称,不含 ".JPG"
Private mvarAdviceID As Long                '检查报告的医嘱ID
Private mvarModified As Boolean             '底图是否被编辑过

Private mblnCompareSize As Boolean

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property

Public Property Let 对象序号(ByVal vData As Double)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Double
    对象序号 = mvar对象序号
End Property

Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = IIf(mvar开始版 = 0, 1, mvar开始版)
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
'    终止版 = IIf(mvar终止版 = 0, 0, mvar终止版)
    终止版 = 0
End Property

Public Property Let 是否换行(ByVal vData As Boolean)
    mvar是否换行 = vData
End Property

Public Property Get 是否换行() As Boolean
    是否换行 = mvar是否换行
End Property
Public Property Let 内容文本(ByVal vData As String)
    mvar内容文本 = vData
End Property

Public Property Get 内容文本() As String
    内容文本 = mvar内容文本
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let PictureType(ByVal vData As EPRPictureTypeEnum)
    mvarPictureType = vData
End Property

Public Property Get PictureType() As EPRPictureTypeEnum
    PictureType = mvarPictureType
End Property

Public Property Let MarkedPicID(ByVal vData As Long)
    mvarMarkedPicID = vData
End Property

Public Property Get MarkedPicID() As Long
    MarkedPicID = mvarMarkedPicID
End Property
Public Property Let PicName(ByVal vData As String)
    mvarPicName = vData
End Property

Public Property Get PicName() As String
    PicName = mvarPicName
End Property

Public Property Let AdviceID(ByVal vData As String)
    mvarAdviceID = vData
End Property

Public Property Get AdviceID() As String
    AdviceID = mvarAdviceID
End Property
Public Property Set OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Let OrigPic(ByRef vData As StdPicture)
    Set mvarOrigPic = vData
End Property

Public Property Get OrigPic() As StdPicture
    '检查有效性，如果无效则绘制一副无效图，保证图片有效，同时压缩图片为90%JPEG图片
    Set OrigPic = CheckValidPic(mvarOrigPic)
End Property

Public Property Let PicMarks(ByVal vData As cPicMarks)
    Set mvarPicMarks = vData
End Property

Public Property Set PicMarks(ByVal vData As cPicMarks)
    Set mvarPicMarks = vData
End Property

Public Property Get PicMarks() As cPicMarks
    Set PicMarks = mvarPicMarks
End Property

Public Property Let Left(ByVal vData As Long)
    mvarLeft = vData
End Property

Public Property Get Left() As Long
    Left = mvarLeft
End Property

Public Property Let Top(ByVal vData As Long)
    mvarTop = vData
End Property

Public Property Get Top() As Long
    Top = mvarTop
End Property

Public Property Let OrigWidth(ByVal vData As Long)
    mvarOrigWidth = vData
End Property

Public Property Get OrigWidth() As Long
    OrigWidth = mvarOrigWidth
End Property

Public Property Let OrigHeight(ByVal vData As Long)
    mvarOrigHeight = vData
End Property

Public Property Get OrigHeight() As Long
    OrigHeight = mvarOrigHeight
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let ZoomFactor(ByVal vData As Double)
    mvarZoomFactor = vData
End Property

Public Property Get ZoomFactor() As Double
    ZoomFactor = mvarZoomFactor
End Property

Public Property Let ZOrder(ByVal vData As Long)
    mvarZOrder = vData
End Property

Public Property Get ZOrder() As Long
    ZOrder = mvarZOrder
End Property

Public Property Let Modified(ByVal vData As Boolean)
    mvarModified = vData
End Property

Public Property Get Modified() As Boolean
    Modified = mvarModified
End Property

'################################################################################################################
'## 功能：  复制图片对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRPicture
    Dim cNode As New cEPRPicture
    
    If blnWithKeys Then
        cNode.Key = Me.Key
    End If
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.父ID = Me.父ID
    cNode.对象序号 = Me.对象序号
    cNode.保留对象 = Me.保留对象
    cNode.是否换行 = Me.是否换行
    cNode.内容文本 = Me.内容文本
    cNode.Row = Me.Row
    cNode.Col = Me.Col
    cNode.PictureType = Me.PictureType
    cNode.MarkedPicID = Me.MarkedPicID
    cNode.OrigPic = Me.OrigPic
    cNode.PicMarks = Me.PicMarks.Clone
    cNode.Left = Me.Left
    cNode.Top = Me.Top
    cNode.OrigWidth = Me.OrigWidth
    cNode.OrigHeight = Me.OrigHeight
    cNode.Width = Me.Width
    cNode.Height = Me.Height
    cNode.ZoomFactor = Me.ZoomFactor
    cNode.ZOrder = Me.ZOrder
    cNode.Modified = Me.Modified
    Set Clone = cNode
End Function

'################################################################################################################
'## 功能：  获取最终的按标记绘图后的图片（原始尺寸）
'##
'## 参数：  objPic :用于临时存储绘图结果的图片框控件，最终返回结果图片 StdPicture
'################################################################################################################
Public Function DrawFinalPic(Optional objTable As cEPRTable = Nothing) As StdPicture
    Dim ctlPic As VB.PictureBox
    Set ctlPic = gfrmPublic.Controls.Add("VB.PictureBox", "ctlPic" & CLng(Timer * 1000))
    ctlPic.AutoRedraw = True
    ctlPic.BorderStyle = 0
    If objTable Is Nothing Then
        ctlPic.Height = Me.Height
        ctlPic.Width = Me.Width
        ShowPicMarks ctlPic, Me.OrigPic, Me.PicMarks
    Else
        ctlPic.Height = objTable.Cell(mvarRow, mvarCol).Height
        ctlPic.Width = objTable.Cell(mvarRow, mvarCol).Width
        If objTable.AutoHeight Then
            '如果非自动换行模式，图片拉伸
            ctlPic.Height = ctlPic.Width * mvarOrigPic.Height / mvarOrigPic.Width ' objTable.Cell(mvarRow, mvarCol).Height + (objTable.GridLineWidth - 1) * 15
        End If
        ShowPicMarks ctlPic, Me.OrigPic, ScalePicMarks(Me.PicMarks, CDbl(ctlPic.Height) / CDbl(Me.Height))
    End If
    
    Set DrawFinalPic = ctlPic.Picture
    
    gfrmPublic.Controls.Remove ctlPic
    Set ctlPic = Nothing
End Function

'################################################################################################################
'## 功能：  获取当前左上角的行、列位置
'##
'## 参数：  objF1Book : 图片属于的F1Book控件
'##         lRow_Out:   返回的行值
'##         lCol_Out:   返回的列值
'################################################################################################################
Public Sub GetRowCol(objF1Book As Object, ByRef lRow_Out As Long, ByRef lCol_Out As Long)
    objF1Book.TwipsToRC Me.Left, Me.Top, lRow_Out, lCol_Out
End Sub

'################################################################################################################
'## 功能：  获取当前左上角的行位置
'##
'## 参数：  objF1Book : 图片属于的F1Book控件
'##         lngPicTop:  Top 缇值
'################################################################################################################
Public Function CurRow(objF1Book As Object, lngPicTop As Long) As Long
    CurRow = objF1Book.TwipsToRow(lngPicTop)
    If CurRow <= 0 Then CurRow = 1
End Function

'################################################################################################################
'## 功能：  获取当前左上角的列位置
'##
'## 参数：  objF1Book : 图片属于的F1Book控件
'##         lngPicLeft: Left 缇值
'################################################################################################################
Public Function CurCol(objF1Book As Object, lngPicLeft As Long) As Long
    CurCol = objF1Book.TwipsToCol(lngPicLeft)
    If CurCol <= 0 Then CurCol = 1
End Function

'################################################################################################################
'## 功能：  获取图片属性字符串
'################################################################################################################
Public Property Get 对象属性() As String
    Dim strR As String
    strR = Me.PictureType
    strR = strR & ";" & Me.Row
    strR = strR & ";" & Me.Col
    strR = strR & ";" & Me.Left
    strR = strR & ";" & Me.Top
    strR = strR & ";" & Me.Width
    strR = strR & ";" & Me.Height
    strR = strR & ";" & Me.OrigWidth
    strR = strR & ";" & Me.OrigHeight
    strR = strR & ";" & Me.ZoomFactor
    strR = strR & ";" & Me.ZOrder
    strR = strR & ";" & Me.PicName
    strR = strR & ";" & Me.AdviceID
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置图片属性
'################################################################################################################
Public Property Let 对象属性(vData As String)
    Dim T As Variant, i As Long
    T = Split(vData, ";")
    If UBound(T) > 0 Then
        Me.PictureType = CLng(T(0))
        Me.Row = IIf(T(1) = "", 0, T(1))
        Me.Col = IIf(T(2) = "", 0, T(2))
        Me.Left = IIf(T(3) = "", 0, T(3))
        Me.Top = IIf(T(4) = "", 0, T(4))
        Me.Width = IIf(T(5) = "", 0, T(5))
        Me.Height = IIf(T(6) = "", 0, T(6))
        Me.OrigWidth = IIf(T(7) = "", 0, T(7))
        Me.OrigHeight = IIf(T(8) = "", 0, T(8))
        Me.ZoomFactor = IIf(T(9) = "", 0, T(9))
        Me.ZOrder = IIf(T(10) = "", 0, T(10))
    End If
    If UBound(T) > 10 Then
        Me.PicName = T(11)
        Me.AdviceID = T(12)
    End If
    
End Property

'################################################################################################################
'## 功能：  检查对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_病历文件定义
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历文件结构")
    Case cprET_全文示范编辑
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历范文内容")
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

'################################################################################################################
'## 功能：  检查内容图片是否有效
'##
'## 参数：  objPic      :图片源
'##
'## 返回：  返回经过有效性处理的图片，如果底图丢失，则绘制一幅空白底图。
'################################################################################################################
Public Function CheckValidPic(ByRef objPic As StdPicture) As StdPicture
    If objPic = 0 Then
        '绘制无效图片（虽然底图丢失，但是要保证高度、宽度的正确性）
        Dim ctlPic As VB.PictureBox
        Set ctlPic = gfrmPublic.Controls.Add("VB.PictureBox", "ctlPic" & CLng(Timer * 1000))
        ctlPic.AutoRedraw = True
        ctlPic.Height = Me.Height
        ctlPic.Width = Me.Width
        ctlPic.BorderStyle = 0
        ctlPic.AutoSize = True
        ctlPic.BackColor = vbWhite
        
        ctlPic.Line (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(Me.Width - 2 * Screen.TwipsPerPixelX, Me.Height - 2 * Screen.TwipsPerPixelY), RGB(125, 125, 125), B
        ctlPic.PaintPicture gfrmPublic.imgErrPic.Picture, 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY
        
        ctlPic.Refresh
        ctlPic.Picture = ctlPic.Image
        Set CheckValidPic = ctlPic.Picture
        gfrmPublic.Controls.Remove ctlPic
        Set ctlPic = Nothing
    Else
        Set CheckValidPic = objPic
    End If
End Function

'################################################################################################################
'## 功能：  插入图片到Editor中
'## 参数：  edtThis             :当前的编辑器控件
'##         lngStartPos         :当前位置
'##         blnDrawFinalPic:    :是否绘制最终结果图
'################################################################################################################
Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1, Optional ByVal blnDrawFinalPic As Boolean = False) As Boolean
Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
Dim blnPrefix As Boolean, blnForce As Boolean

    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys And Me.PictureType <> EPRSignPicture Then InsertIntoEditor = False: Exit Function    '保证不能插入关键字内部,签名图除外，图片会在保存RTF前取掉，存完后再次插入

    If Me.PictureType = EPRSignPicture Or Me.PictureType = EPRFormulaPicture Or Me.PictureType = EPRPatiSign Then '前缀加空格还是加回车换行符
        blnPrefix = False
    Else
        blnPrefix = True
    End If
    
    With edtThis
         blnForce = .ForceEdit
        .Freeze
        .ForceEdit = True
        If lngStartPos = 0 Then
            .Range(lngStartPos, lngStartPos) = " "
            lngStartPos = 1
            .Range(0, 1).Font.Protected = False
            .Range(0, 1).Font.Hidden = False
        End If
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        .Range(lngStartPos, lngStartPos).Text = IIf(blnPrefix, vbCrLf, " ") & "PS(" & Format(Me.Key, "00000000") & ",1,0)"
        .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 18, 17)).Font.Protected = False
        .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 18, 17)).Font.Hidden = False
        
        If blnDrawFinalPic And blnPrefix Then '标记图需要绘制最终结果图，而签名图和公式图则插入原始图
            Call .InsertPicture(Me.DrawFinalPic, lStart:=lngStartPos + IIf(blnPrefix, 18, 17), lEnd:=lngStartPos + IIf(blnPrefix, 18, 17))
        Else
            Call .InsertPicture(Me.OrigPic, lStart:=lngStartPos + IIf(blnPrefix, 18, 17), lEnd:=lngStartPos + IIf(blnPrefix, 18, 17))
        End If
        .ReadOnly = False '-55508因频繁刷新出现"拒绝的权限",需要以下三行
        .ForceEdit = True
        .Range(lngStartPos + IIf(blnPrefix, 19, 18), lngStartPos + IIf(blnPrefix, 19, 18)).Font.Protected = False
        .Range(lngStartPos + IIf(blnPrefix, 19, 18), lngStartPos + IIf(blnPrefix, 19, 18)) = "PE(" & Format(Me.Key, "00000000") & ",1,0)" & " " & IIf(Me.是否换行, vbCrLf, "")
        If blnPrefix Then '回车或空格前辍可见
            .Range(lngStartPos, lngStartPos + 2).Font.Hidden = False
        Else
            .Range(lngStartPos, lngStartPos + 1).Font.Hidden = False
        End If
        .Range(lngStartPos + IIf(blnPrefix, 2, 1), lngStartPos + IIf(blnPrefix, 18, 17)).Font.Hidden = True
        .Range(lngStartPos + IIf(blnPrefix, 18, 17), lngStartPos + IIf(blnPrefix, 19, 18)).Font.Hidden = False
        If blnPrefix Then '独立成行的图片需要设置单倍行距，否则只显示一半高度
            .Range(lngStartPos + IIf(blnPrefix, 18, 17), lngStartPos + IIf(blnPrefix, 19, 18)).Para.SetLineSpacing cprLSSignle, 0 '设置为单倍行距
        End If
        .Range(lngStartPos + IIf(blnPrefix, 19, 18), lngStartPos + IIf(blnPrefix, 35, 34)).Font.Hidden = True
        .Range(lngStartPos + IIf(blnPrefix, 35, 34), lngStartPos + IIf(blnPrefix, 36, 35)).Font.Hidden = False '空格后辍可见
        If Me.是否换行 Then
            .Range(lngStartPos + IIf(blnPrefix, 36, 35), lngStartPos + IIf(blnPrefix, 38, 37)).Font.Hidden = False '回车后辍可见
            .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 38, 37)).Font.Protected = True
            .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 38, 37)).Font.ForeColor = tomAutoColor
            .Range(lngStartPos + IIf(blnPrefix, 38, 37), lngStartPos + IIf(blnPrefix, 38, 37)).Selected
        Else
            .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 36, 35)).Font.Protected = True
            .Range(lngStartPos, lngStartPos + IIf(blnPrefix, 36, 35)).Font.ForeColor = tomAutoColor
            .Range(lngStartPos + IIf(blnPrefix, 36, 35), lngStartPos + IIf(blnPrefix, 36, 35)).Selected
        End If
        .ForceEdit = blnForce
        .UnFreeze
    End With
    edtThis.LockAllOLEObjectSize
    InsertIntoEditor = True
End Function

''################################################################################################################
''## 功能：  刷新显示
''##
''## 参数：  edtThis         :当前的编辑器控件
''################################################################################################################
'Public Function Refresh(ByRef edtThis As Object) As Boolean
'    Dim strTmp As String, lngKey As Long, blnForce As Boolean
'    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
'
'    bFinded = FindKey(edtThis, "P", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
'    If bFinded Then
'        With edtThis
'            .Freeze
'            blnForce = .ForceEdit
'            .ForceEdit = True
''            Dim ParaFmt As New cParaFormat
''            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt
'
'            .Range(lKSS, lKEE) = ""
'            If Me.是否换行 Then Me.是否换行 = False
'            Me.InsertIntoEditor edtThis, lKSS
''            .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
'            .ForceEdit = blnForce
'            .UnFreeze
'        End With
'        Refresh = True
'    Else
'        Refresh = False
'    End If
'End Function

'################################################################################################################
'## 功能：  从编辑器中删除
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "P", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            If .Range(lKEE, lKEE + 2).Text = vbCrLf And .Range(lKEE, lKEE + 2).Font.Protected And Me.是否换行 Then   '图片后跟的换行符或空格且处于保护
                .Range(lKEE, lKEE + 2).Font.Protected = False
                .Range(lKEE, lKEE + 2).Text = ""
            End If
                
            If .Range(lKEE, lKEE + 1).Text = " " And .Range(lKEE, lKEE + 2).Font.Protected Then    '图片后跟的空格且处于保护
                .Range(lKEE, lKEE + 1).Font.Protected = False
                .Range(lKEE, lKEE + 1).Text = ""
            End If
            .Range(lKSS, lKEE).Text = ""
            .Range(lKSS, lKSS).Selected

            If .Range(lKSS - 2, lKSS).Text = vbCrLf And .Range(lKSS - 2, lKSS).Font.Protected Then   '回车前辍或空格前辍且处于保护
                .Range(lKSS - 2, lKSS).Font.Protected = False
                .Range(lKSS - 2, lKSS).Text = ""
                .Range(lKSS - 2, lKSS - 2).Selected
            ElseIf .Range(lKSS - 1, lKSS).Text = " " And .Range(lKSS - 1, lKSS).Font.Protected Then
                .Range(lKSS - 1, lKSS).Font.Protected = False
                .Range(lKSS - 1, lKSS).Text = ""
                .Range(lKSS - 1, lKSS - 1).Selected
            End If
            
            .ForceEdit = blnForce
            .UnFreeze
        End With
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function
Public Function FillPictureMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能:用数据记录填充类成员
Dim rsTemp As ADODB.Recordset, strTemp As String
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs("对象标记"), 0) <> 0 Then Me.Key = NVL(rs("对象标记"), 0)
    Me.ID = NVL(rs("ID"), 0)
    Me.文件ID = NVL(rs("文件ID"), 0)
    Me.父ID = NVL(rs("父ID"), 0)
    Me.对象序号 = NVL(rs("对象序号"), 0)
    Me.保留对象 = IIf(NVL(rs("保留对象"), 0) = 0, False, True)
    Me.对象属性 = NVL(rs("对象属性"))
    Me.是否换行 = IIf(NVL(rs("是否换行"), 0) = 0, False, True)
    Me.内容文本 = NVL(rs("内容文本"))
    
    If InStr(strTableName, "电子病历内容") > 0 Then '病历文件结构和病历范文内容没有这两个字段,但都共用这个函数
        Me.开始版 = NVL(rs!开始版, 1)
        Me.终止版 = NVL(rs!终止版, 0)
    End If
    
    If Me.AdviceID = 0 And Me.PicName = "" Then '检查报告图，不再存数据库
        Select Case strTableName
            Case "病历文件结构"
                strTemp = zlBlobRead(2, Me.ID)
            Case "病历范文内容"
                strTemp = zlBlobRead(4, Me.ID)
            Case "电子病历内容"
                strTemp = zlBlobRead(6, Me.ID)
            Case "H电子病历内容" '读出的是历史表时,图形数据也是转储
                strTemp = zlBlobRead(6, Me.ID, , True)
        End Select
    Else
        gstrSQL = "Select To_Char(L.接收日期, 'yyyymmdd') As 子目录, L.检查uid, L.报告图象, A1.Ftp目录 As Root1, A1.Ip地址 As Ip1," & vbNewLine & _
                "       A1.FTP用户名 As Usr1, A1.FTP密码 As Pwd1" & vbNewLine & _
                "From 影像检查记录 L, 影像设备目录 A1" & vbNewLine & _
                "Where L.位置一 = A1.设备号(+) And L.医嘱id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取目录", Me.AdviceID)
        
        If Not rsTemp.EOF Then
            '防止人为删除目录的情况
            If Not gobjFSO.FolderExists(App.Path & "\TmpImage") Then Call gobjFSO.CreateFolder(App.Path & "\TmpImage")
            If Not gobjFSO.FolderExists(App.Path & "\TmpImage\" & rsTemp!子目录) Then Call gobjFSO.CreateFolder(App.Path & "\TmpImage\" & rsTemp!子目录)
            If Not gobjFSO.FolderExists(App.Path & "\TmpImage\" & rsTemp!子目录 & "\" & rsTemp!检查uid) Then Call gobjFSO.CreateFolder(App.Path & "\TmpImage\" & rsTemp!子目录 & "\" & rsTemp!检查uid)
            
            If gobjFSO.FileExists(App.Path & "\TmpImage\" & rsTemp!子目录 & "\" & rsTemp!检查uid & "\" & Me.PicName) Then
                '本地已有图则不下载
                strTemp = App.Path & "\TmpImage\" & rsTemp!子目录 & "\" & rsTemp!检查uid & "\" & Me.PicName
            Else
                Dim cFtpNet As New cFtp, strPath As String, strFile As String
                If NVL(rsTemp!Ip1) <> "" Then
                    If cFtpNet.FuncFtpConnect(rsTemp!Ip1, NVL(rsTemp!Usr1), NVL(rsTemp!pwd1)) <> 0 Then
                        strPath = rsTemp!Root1 & "/" & rsTemp!子目录 & "/" & rsTemp!检查uid
                        strTemp = App.Path & "\TmpImage\" & rsTemp!子目录 & "\" & rsTemp!检查uid & "\" & Me.PicName
ReDownload:
                        Call cFtpNet.FuncDownloadFile(strPath, strTemp, Me.PicName)
                        
                        If mblnCompareSize Then
                            If ChechReportImgAndReDownload(cFtpNet, strTemp, strPath, Me.PicName) Then
                                GoTo ReDownload
                            End If
                        End If
                        cFtpNet.FuncFtpDisConnect
                    End If
                End If
            End If
        End If
    End If
    
    If Len(strTemp) > 0 Then
        If gobjFSO.FileExists(strTemp) Then
            Call SetPictureToOrigPic(strTemp)
            gobjFSO.DeleteFile strTemp, True
        End If
    Else
        Set Me.OrigPic = Me.OrigPic
    End If
    
    strTemp = ""
    gstrSQL = "Select 内容文本 " & _
        " From " & strTableName & _
        " Where 文件ID = [1] And 父id=[2] And 对象类型=6 " & _
        " Order By 内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRPicture", Me.文件ID, Me.ID)
    Do While Not rsTemp.EOF '---------------读取图片标记
        strTemp = strTemp & NVL(rsTemp!内容文本)
        rsTemp.MoveNext
    Loop
    
    '分解标记内容并存储到对象属性中
    Me.PicMarks.对象属性 = strTemp
    
    FillPictureMember = True
End Function

Private Function ChechReportImgAndReDownload(cFtpNet As cFtp, ByVal strSrcFile As String, strFtpFilePath As String, ByVal strFileName As String) As Boolean
'检查上传的文件和本地文件大小是否一致，不一致则返回true
    Dim blnResult As Boolean, blnReUpload As Boolean
    Dim lngFtpFileSzie As Long, lngDestFileSize As Long
    Dim StrMessage As String
    Dim objFileSystem As New FileSystemObject
    
On Error GoTo errHandle
    ChechReportImgAndReDownload = False
    
    '上传后对比一下大小，判断是否正常上传
    lngDestFileSize = objFileSystem.GetFile(strSrcFile).Size
    lngFtpFileSzie = cFtpNet.FuncFtpGetFileSize(strFtpFilePath, strFileName)

    If lngFtpFileSzie <> lngDestFileSize Then
        '强制重读
        Call cFtpNet.FuncDownloadFile(strFtpFilePath, strSrcFile, strFileName, True)
        
        lngDestFileSize = objFileSystem.GetFile(strSrcFile).Size
    End If

    If lngFtpFileSzie <> lngDestFileSize Then
        StrMessage = "下载后的本地文件大小[" & lngDestFileSize & "]与FTP原文件大小[" & lngFtpFileSzie & "]不一致" & vbCrLf & _
                     "本地文件：" & strSrcFile & vbCrLf & _
                     "FTP原文件：" & strFtpFilePath & "/" & strFileName & vbCrLf & _
                     "是否需要重新下载？"
        
        If MsgBox(StrMessage, vbQuestion + vbYesNo, "提示") = vbYes Then
            ChechReportImgAndReDownload = True
        End If
    End If
Exit Function
errHandle:
    If MsgBox("图像文件[本地:" & strSrcFile & "  FTP:" & strFtpFilePath & "/" & strFileName & "]一致性检查错误,错误原因:" & Err.Description & "。" & vbCrLf & "是否重试？", vbQuestion + vbYesNo, "提示") = vbYes Then
        ChechReportImgAndReDownload = True
    End If
End Function

Private Sub SetPictureToOrigPic(ByVal strPicFile As String)
On Error GoTo errHandle
'    Dim objFile As File
'
'    Set objFile = gobjFSO.GetFile(strPicFile)
'
'    '1*1像素jpg图片为631字节
'    If objFile.Size <= 500 Then Exit Sub
    
    Set Me.OrigPic = LoadPicture(strPicFile)
    
Exit Sub
errHandle:
End Sub


'################################################################################################################
'## 功能：  从数据库直接获取图片内容
'##
'## 参数：  eEdtType    :当前编辑方式
'##         lngFileID   :文件ID
'##         lngID       :数据库中的对象ID号
'##         blnWithKey  :是否包含Key值一起恢复
'##         strTableName:表名，如果省略，则根据当前的eEdtType来判断表名
'################################################################################################################
Public Function GetPictureFromDB(ByVal eEdtType As EditTypeEnum, _
    ByVal lngFileID As Long, ByVal lngId As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean
    
    Dim rs As New ADODB.Recordset, RsSub As New ADODB.Recordset, strTmp As String
    
    If strTableName = "" Then
        Select Case eEdtType
        Case cprET_病历文件定义
            strTableName = "病历文件结构"
        Case cprET_全文示范编辑
            strTableName = "病历范文内容"
        Case cprET_单病历编辑, cprET_单病历审核
            strTableName = "电子病历内容"
        End Select
    End If
    
    gstrSQL = "Select ID, 文件id," & IIf(InStr(strTableName, "电子病历内容") > 0, "开始版, 终止版,", "") & _
                " 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行 " & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRCompend", lngId)
    If Not rs.EOF Then
        If blnWithKey And NVL(rs("对象标记"), 0) <> 0 Then Me.Key = NVL(rs("对象标记"), 0)
        Me.ID = NVL(rs("ID"), 0)
        Me.文件ID = NVL(rs("文件ID"), 0)
        Me.父ID = NVL(rs("父ID"), 0)
        Me.对象序号 = NVL(rs("对象序号"), 0)
        Me.保留对象 = IIf(NVL(rs("保留对象"), 0) = 0, False, True)
        Me.对象属性 = NVL(rs("对象属性"))
        Me.是否换行 = IIf(NVL(rs("是否换行"), 0) = 0, False, True)
        Me.内容文本 = NVL(rs("内容文本"))
        If eEdtType = cprET_单病历审核 Or eEdtType = cprET_单病历编辑 Then
            Me.开始版 = NVL(rs("开始版"), 1)
            Me.终止版 = NVL(rs("终止版"), 0)
        End If

        '图片数据调入内存
        Dim strTemp As String
        
        Select Case eEdtType
        Case cprET_病历文件定义
            strTemp = zlBlobRead(2, Me.ID)
        Case cprET_全文示范编辑
            strTemp = zlBlobRead(4, Me.ID)
        Case cprET_单病历编辑, cprET_单病历审核
            strTemp = zlBlobRead(6, Me.ID, , strTableName = "H电子病历内容") '读出的是历史表时,图形数据也是转储
        End Select
        If Len(strTemp) > 0 Then
            Set Me.OrigPic = LoadPicture(strTemp)
            If gobjFSO.FileExists(strTemp) Then gobjFSO.DeleteFile strTemp, True
        Else
            Set Me.OrigPic = Me.OrigPic
        End If
        
        gstrSQL = "Select t.内容文本 " & _
            " From " & strTableName & " t " & _
            " Where 文件ID = [1] And 父id=[2] And 对象类型=6 " & _
            " Order By 内容行次"
        Set RsSub = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRPicture", lngFileID, Me.ID)
        strTmp = ""
        Do While Not RsSub.EOF '-------------读取图片标记
            strTmp = strTmp & NVL(RsSub("内容文本"))
            RsSub.MoveNext
        Loop
        
        '分解标记内容并存储到对象属性中
        Me.PicMarks.对象属性 = strTmp
        
        GetPictureFromDB = True
    Else
        GetPictureFromDB = False
    End If
End Function

'################################################################################################################
'## 功能：  获取图片保存到数据库的SQL语句
'##
'## 参数：  eEdtType        :IN，       当前编辑模式
'##         ArraySQL()      :IN/OUT，   SQL数组
'##         blnInTable      :In，       是否是表格中的图片
'##         lng行次         ;In，       如果是表格中图片，则必须指明需要保存的行次。
'##
'## 说明：  普通图片按照对象序号保存；表格中图片对象序号为 NULL，按照 内容行次 来保存。
'################################################################################################################
Public Function GetSaveSQL(ByVal eEdtType As EditTypeEnum, _
    ByRef ArraySQL() As String, _
    Optional blnInTable As Boolean = False, _
    Optional lng行次 As Long) As Boolean
    
    Dim lngCount As Long
    
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    Call CheckValidID(eEdtType)         '检查ID有效性！
    Select Case eEdtType
    Case cprET_病历文件定义
        gstrSQL = "Zl_病历文件结构_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            IIf(blnInTable, "NULL", Me.对象序号) & ",5," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "'," & _
            IIf(blnInTable, lng行次, "NULL") & ",'" & Me.内容文本 & "'," & IIf(blnInTable, "NULL", IIf(Me.是否换行, 1, 0)) & ")"
    Case cprET_全文示范编辑
        gstrSQL = "Zl_病历范文内容_Update(" & Me.ID & "," & Me.文件ID & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            IIf(blnInTable, "NULL", Me.对象序号) & ",5," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "'," & _
            IIf(blnInTable, lng行次, "NULL") & ",'" & Me.内容文本 & "'," & IIf(blnInTable, "NULL", IIf(Me.是否换行, 1, 0)) & ")"
    Case cprET_单病历编辑, cprET_单病历审核
        '加了 开始版，终止版 字段，其他相同
        gstrSQL = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
            IIf(blnInTable, "NULL", Me.对象序号) & ",5," & Me.Key & "," & IIf(Me.保留对象, 1, 0) & ",'" & Me.对象属性 & "'," & _
            IIf(blnInTable, lng行次, "NULL") & ",'" & Me.内容文本 & "'," & IIf(blnInTable, "NULL", IIf(Me.是否换行, 1, 0)) & ")"
    End Select
    ArraySQL(lngCount) = gstrSQL
    
    '同时添加标记的保存SQL
    Dim i As Long, strMarks As String, strSub As String, lngLen As Long, lngSub行次 As Long, lngId As Double
    strMarks = Me.PicMarks.对象属性
    lngLen = Len(strMarks)

    '按照4000为界分段存储。
    i = 0
    lngSub行次 = 1
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strMarks, i * 2000 + 1, 2000)

        Select Case eEdtType
        Case cprET_病历文件定义
            lngId = zlDatabase.GetNextId("病历文件结构")
            gstrSQL = "Zl_病历文件结构_Update(" & lngId & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                "6,NULL,NULL,NULL," & lngSub行次 & ",'" & strSub & "')"
        Case cprET_全文示范编辑
            lngId = zlDatabase.GetNextId("病历范文内容")
            gstrSQL = "Zl_病历范文内容_Update(" & lngId & "," & Me.文件ID & "," & Me.ID & ",NULL," & _
                "6,NULL,NULL,NULL," & lngSub行次 & ",'" & strSub & "')"
        Case cprET_单病历编辑, cprET_单病历审核
            lngId = GetEPRContentNextId
            '纯文本的 版本号 均为NULL，所有历史数据均保存在同一（最终）版本中。
            gstrSQL = "ZL_电子病历内容_Update(" & lngId & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & Me.ID & ",NULL," & _
                "6,NULL,NULL,NULL," & lngSub行次 & ",'" & strSub & "')"
        End Select
        ArraySQL(lngCount) = gstrSQL
       
        lngSub行次 = lngSub行次 + 1
        i = i + 1
    Loop
    GetSaveSQL = True
End Function

'################################################################################################################
'## 功能：  保存原始图片到数据库
'##
'## 参数：  eEdtType    :当前编辑模式
'################################################################################################################
Public Function SaveOrigPic(ByVal eEdtType As EditTypeEnum, ByRef arySql() As String) As Boolean
    On Error GoTo LL
    Dim m_Temp As String
    Dim rs As New ADODB.Recordset, RsSub As New ADODB.Recordset
    Dim strFileName As String
    
    If Me.PictureType = EPRInnerPicture And Me.PicName <> "" Then
        SaveOrigPic = True
        Exit Function '检查报告类图片不再存储数据库
    End If
    
    m_Temp = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strFileName = m_Temp & "\OrigPic" & Timer & ".jpg"
    If gobjFSO.FileExists(strFileName) Then gobjFSO.DeleteFile strFileName, True
    
    Dim GpInput As GdiplusStartupInput
    Dim m_GDIpToken         As Long         ' 用于关闭 GDI+
    Dim oDIB As New cDIB
    Dim DIBDither As New cDIBDither
    Dim DIBPal As New cDIBPal
    '-- 调入 GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (mGdIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        '按照BMP格式保存！会增大图片体积
        SavePicture Me.OrigPic, strFileName       '保存格式为BMP格式
    Else
        '采用JPEG压缩格式保存
        Call oDIB.CreateFromStdPicture(Me.OrigPic, DIBPal, DIBDither)
        '压缩存储
        Call mGdIpEx.SaveDIB(oDIB, strFileName, [ImageJPEG], 100)          '90%的JPEG图片压缩质量
    End If
    'Unload the GDI+ Dll
    Call mGdIpEx.GdiplusShutdown(m_GDIpToken)
    

    Select Case eEdtType
    Case cprET_病历文件定义
        gstrSQL = "select 对象ID from 病历文件图形 where 对象ID=[1]"
    Case cprET_全文示范编辑
        gstrSQL = "select 对象ID from 病历范文图形 where 对象ID=[1]"
    Case cprET_单病历编辑, cprET_单病历审核
        gstrSQL = "select 对象ID from 电子病历图形 where 对象ID=[1]"
    End Select
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRPicture", Me.ID)
        
    If rs.EOF Or Me.Modified Then
        '已可写方式打开记录
        Dim blnSave As Boolean
        Select Case eEdtType
        Case cprET_病历文件定义
            blnSave = zlBlobSql(2, Me.ID, strFileName, arySql)
        Case cprET_全文示范编辑
            blnSave = zlBlobSql(4, Me.ID, strFileName, arySql)
        Case cprET_单病历编辑, cprET_单病历审核
            blnSave = zlBlobSql(6, Me.ID, strFileName, arySql)
        End Select
        If blnSave Then
            Me.Modified = False
        Else
            GoTo LL
        End If
    Else
        '如果记录已经存在，则无需保存
    End If
    If gobjFSO.FileExists(strFileName) Then gobjFSO.DeleteFile strFileName, True        '删除临时文件
    
    SaveOrigPic = True
    Exit Function
LL:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    SaveOrigPic = False
End Function

Private Sub ReadFtpCfg()
On Error GoTo errHandle
    mblnCompareSize = IIf(Val(GetSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", 1)) <> 0, True, False)
    Call SaveSetting("ZLSOFT", "公共模块\Ftp", "启用FTP文件大小对比", IIf(mblnCompareSize, 1, 0))
Exit Sub
errHandle:
    mblnCompareSize = False
End Sub

Private Sub Class_initialize()
    Set mvarOrigPic = New StdPicture
    Set mvarPicMarks = New cPicMarks
    ZoomFactor = 1#
    
    Call ReadFtpCfg
End Sub

Private Sub Class_Terminate()
    Set mvarOrigPic = Nothing
    Set mvarPicMarks = Nothing
End Sub


