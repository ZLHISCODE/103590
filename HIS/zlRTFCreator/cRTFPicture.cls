VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFPicture"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'保持属性值的局部变量
Private mvarFileName As String          '实际图片16进制数据
Private mvarScaleX As Long              '水平缩放比例，默认：100
Private mvarScaleY As Long              '垂直缩放比例，默认：100
Private mvarCropLeft As Long            '左端剪切值，默认：0
Private mvarCropRight As Long           '右端剪切值，默认：0
Private mvarCropTop As Long             '上端剪切值，默认：0
Private mvarCropBottom As Long          '下端剪切值，默认：0
Private mvarWidth As Long               '图片象素宽度
Private mvarHeight As Variant           '图片象素高度
Private mvarWidthOld As Long            '图片原始宽度
Private mvarHeightOld As Long           '图片原始高度
Private mvarPicType As PicTypeEnum      '目前只能为jpegblip，默认：JPG
Private mvarPicID As Long               '图片ID值，必须为负值：如\bliptag-10292
Private mvarSingleLine As Boolean       '图片是否单独一行
Private mvarAlignment As AlignTypeEnum
Private mvarIsInTable As Boolean

Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "是否包含在表格中。默认:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.IsInTable = 5
    mvarIsInTable = vData
End Property


Public Property Get IsInTable() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property


Public Property Let Alignment(ByVal vData As AlignTypeEnum)
Attribute Alignment.VB_Description = "图片对齐方式。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Alignment = 5
    mvarAlignment = vData
End Property


Public Property Get Alignment() As AlignTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property


Public Property Let SingleLine(ByVal vData As Boolean)
Attribute SingleLine.VB_Description = "图片是否是单独一行。默认：True。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SingleLine = 5
    mvarSingleLine = vData
End Property


Public Property Get SingleLine() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SingleLine
    SingleLine = mvarSingleLine
End Property




Public Property Let PicID(ByVal vData As Long)
Attribute PicID.VB_Description = "图片ID值。必须为负数。目前没有使用。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PicID = 5
    mvarPicID = -Abs(vData)
End Property


Public Property Get PicID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PicID
    PicID = -Abs(mvarPicID)
End Property



Public Property Let PicType(ByVal vData As PicTypeEnum)
Attribute PicType.VB_Description = "图片类型。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PicType = 5
    mvarPicType = vData
End Property


Public Property Get PicType() As PicTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PicType
    PicType = mvarPicType
End Property



Public Property Let HeightOld(ByVal vData As Long)
Attribute HeightOld.VB_Description = "图片原始高度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.HeightOld = 5
    mvarHeightOld = vData
End Property


Public Property Get HeightOld() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.HeightOld
    HeightOld = mvarHeightOld
End Property



Public Property Let WidthOld(ByVal vData As Long)
Attribute WidthOld.VB_Description = "图片原始宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.WidthOld = 5
    mvarWidthOld = vData
End Property


Public Property Get WidthOld() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.WidthOld
    WidthOld = mvarWidthOld
End Property



Public Property Let Height(ByVal vData As Variant)
Attribute Height.VB_Description = "图片象素高度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.mHeight = 5
    mvarHeight = vData
End Property


Public Property Set Height(ByVal vData As Variant)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.mHeight = Form1
    Set mvarHeight = vData
End Property


Public Property Get Height() As Variant
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.mHeight
    If IsObject(mvarHeight) Then
        Set Height = mvarHeight
    Else
        Height = mvarHeight
    End If
End Property



Public Property Let Width(ByVal vData As Long)
Attribute Width.VB_Description = "图片象素宽度。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.mWidth = 5
    mvarWidth = vData
End Property


Public Property Get Width() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.mWidth
    Width = mvarWidth
End Property



Public Property Let CropBottom(ByVal vData As Long)
Attribute CropBottom.VB_Description = "下端剪切值。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CropBottom = 5
    mvarCropBottom = vData
End Property


Public Property Get CropBottom() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CropBottom
    CropBottom = mvarCropBottom
End Property



Public Property Let CropTop(ByVal vData As Long)
Attribute CropTop.VB_Description = "上端剪切值。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CropTop = 5
    mvarCropTop = vData
End Property


Public Property Get CropTop() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CropTop
    CropTop = mvarCropTop
End Property



Public Property Let CropRight(ByVal vData As Long)
Attribute CropRight.VB_Description = "右端剪切值。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CropRight = 5
    mvarCropRight = vData
End Property


Public Property Get CropRight() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CropRight
    CropRight = mvarCropRight
End Property



Public Property Let CropLeft(ByVal vData As Long)
Attribute CropLeft.VB_Description = "左端剪切值。默认：0"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.CropLeft = 5
    mvarCropLeft = vData
End Property


Public Property Get CropLeft() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CropLeft
    CropLeft = mvarCropLeft
End Property



Public Property Let ScaleY(ByVal vData As Long)
Attribute ScaleY.VB_Description = "垂直缩放比例。默认：100"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ScaleY = 5
    mvarScaleY = vData
End Property


Public Property Get ScaleY() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ScaleY
    ScaleY = mvarScaleY
End Property



Public Property Let ScaleX(ByVal vData As Long)
Attribute ScaleX.VB_Description = "水平缩放比例。默认：100"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ScaleX = 5
    mvarScaleX = vData
End Property


Public Property Get ScaleX() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ScaleX
    ScaleX = mvarScaleX
End Property



Public Property Let FileName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Data = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
Attribute FileName.VB_Description = "图片源的文件名。"
Attribute FileName.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Data
    FileName = mvarFileName
End Property


Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "获取图片的RTF字符串。"
'获取最终RTF文本
    Dim strR As String
    
    Dim strAlignment As String
    Select Case mvarAlignment
    Case alm2_居中对齐
        strAlignment = "\qc"
    Case alm3_居右对齐
        strAlignment = "\qr"
    Case alm4_两端对齐
        strAlignment = "\qj"
    Case alm5_分散对齐
        strAlignment = "\qd"
    Case Else
        strAlignment = "\ql"
    End Select
    Dim strIsInTable As String
    If mvarIsInTable Then strIsInTable = "\intbl"
    
    strR = "\pard\plain " & strAlignment & " \li0\ri0\nowidctlpar" & strIsInTable & "\aspalpha\aspnum\faauto\adjustright\rin0\lin0\yts15 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 "
    strR = strR & "{\*\shppict{\pict"
    If mvarScaleX > 0 Then strR = strR & "\picscalex" & mvarScaleX
    If mvarScaleY > 0 Then strR = strR & "\picscaley" & mvarScaleY
    If mvarCropLeft > 0 Then strR = strR & "\piccrop" & mvarCropLeft
    If mvarCropRight > 0 Then strR = strR & "\piccropr" & mvarCropRight
    If mvarCropTop > 0 Then strR = strR & "\piccropt" & mvarCropTop
    If mvarCropBottom > 0 Then strR = strR & "\piccropb" & mvarCropBottom
    If mvarHeight > 0 Then strR = strR & "\pichgoal" & mvarHeight
    If mvarWidth > 0 Then strR = strR & "\picwgoal" & mvarWidth
    If mvarHeightOld > 0 Then strR = strR & "\pich" & mvarHeightOld
    If mvarWidthOld > 0 Then strR = strR & "\picw" & mvarWidthOld
    If mvarPicType = pic_JPG Then   '目前只能有一种JPEG类型
        strR = strR & "\jpegblip "
    Else
        strR = strR & "\jpegblip "   '也用JPEG方法来处理
    End If
    'If mvarPicID > 0 Then  mvarpicID暂时没有使用。
    
    strR = strR & PicToASC(mvarFileName) & "}}"
    If SingleLine And Not mvarIsInTable Then strR = strR & "\par "

    GetTextRTF = strR
End Function

