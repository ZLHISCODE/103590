VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################################
'##模 块 名：cEPRSign.cls
'##创 建 人：吴庆伟
'##日    期：2005年11月24日
'##修 改 人：
'##日    期：
'##描    述：单个签名对象（对象类型＝8）
'##版    本：
'###############################################################################################

Option Explicit

Private mvarKey As Long                     '关键字，存储在“对象标记”中
Private mvarID As Double                      '用于保存的ID值
Private mvar文件ID As Long                  '文件ID
Private mvar父ID As Double                   '父ID
Private mvar开始版 As Long                  '开始版
Private mvar终止版 As Long                  '终止版
Private mvar对象序号 As Double                '每个签名按照文本顺序有个唯一序号（全文唯一）
Private mvar姓名 As String                  '签名用户的姓名，存储在“内容文本”中
Private mvar签名信息 As String              '数字签名字符串，存储在“要素值域”中
Private mvar前置文字 As String              '称呼或者顿号等前置文字，存储在“要素名称”中
Private mvar签名级别 As EPRSignLevelEnum    '0,1,2,3，存储在“要素表示”中
Private mvar签名人ID As Long                '记录签名人ID
Private mbln签名图片 As Boolean             '是否使用签名图片,存储在"输入形态"中

'下面5个属性组成“对象属性”，存储在“对象属性”字段中
Private mvar签名方式 As Long                '1－密码签名；2－数字签名
Private mvar签名规则 As Long                '从1开始编号，由程序自定义其涵义（表示不同版本的签名方法）
                                            '签名规则=1 使用RTF.Text做为数字签名原文，但生成原文的方式有问题，签名对象位置混乱时会出错
                                            '签名规则=2 使用RTF.Text做为数字签名原文 10.28.170版本修改见 frmman.Get
                                            '未定BY ZXH =3 使用保存数据库后的内容文本（不含签名要素）为数字签名原文
Private mvar证书ID As Long                  '数字证书ID号
Private mvar显示手签 As Boolean             '是否显示手签标志
Private mvar签名时间 As String              '签名时间，用于回退时恢复完成时间
Private mvar显示时间 As String              '是否显示签名时间及显示的格式串,空串表示不显示
Private mvar签名要素 As String             '是否为签名要素Key值
Private mvar时间戳 As String                '时间戳，存储在“要素单位”中
Private mvar时间戳信息 As String            '电子签名返回的时间戳的认证编码信息(Base64)

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

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = mvar开始版
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
    终止版 = mvar终止版
End Property

Public Property Let 对象序号(ByVal vData As Double)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Double
    对象序号 = mvar对象序号
End Property

Public Property Let 姓名(ByVal vData As String)
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
Attribute 姓名.VB_UserMemId = 0
    姓名 = mvar姓名
End Property
Public Property Let 签名人ID(ByVal vData As Long)
    mvar签名人ID = vData
End Property

Public Property Get 签名人ID() As Long
    签名人ID = mvar签名人ID
End Property

Public Property Let 签名信息(ByVal vData As String)
    mvar签名信息 = vData
End Property

Public Property Get 签名信息() As String
    签名信息 = mvar签名信息
End Property

Public Property Let 签名要素(ByVal vData As String)
    mvar签名要素 = vData
End Property

Public Property Get 签名要素() As String
    签名要素 = mvar签名要素
End Property

Public Property Let 前置文字(ByVal vData As String)
    mvar前置文字 = vData
End Property

Public Property Get 前置文字() As String
    前置文字 = mvar前置文字
End Property

Public Property Let 签名级别(ByVal vData As EPRSignLevelEnum)
    mvar签名级别 = vData
End Property

Public Property Get 签名级别() As EPRSignLevelEnum
    签名级别 = mvar签名级别
End Property

Public Property Let 签名方式(ByVal vData As Long)
    mvar签名方式 = vData
End Property

Public Property Get 签名方式() As Long
    签名方式 = mvar签名方式
End Property

Public Property Let 签名规则(ByVal vData As Long)
    mvar签名规则 = vData
End Property

Public Property Get 签名规则() As Long
    签名规则 = mvar签名规则
End Property

Public Property Let 证书ID(ByVal vData As Long)
    mvar证书ID = vData
End Property

Public Property Get 证书ID() As Long
    证书ID = mvar证书ID
End Property

Public Property Let 签名时间(ByVal vData As String)
    mvar签名时间 = vData
End Property

Public Property Get 签名时间() As String
    签名时间 = mvar签名时间
End Property

Public Property Let 显示时间(ByVal vData As String)
    mvar显示时间 = vData
End Property

Public Property Get 显示时间() As String
    显示时间 = mvar显示时间
End Property

Public Property Let 时间戳(ByVal vData As String)
    mvar时间戳 = vData
End Property

Public Property Get 时间戳() As String
    时间戳 = mvar时间戳
End Property

Public Property Let 时间戳信息(ByVal vData As String)
    mvar时间戳信息 = vData
End Property

Public Property Get 时间戳信息() As String
    时间戳信息 = mvar时间戳信息
End Property

Public Property Let 显示手签(ByVal vData As Boolean)
    mvar显示手签 = vData
End Property

Public Property Get 显示手签() As Boolean
    显示手签 = mvar显示手签
End Property

Public Property Let 对象属性(ByVal vData As String)
    Dim T As Variant
    T = Split(vData, ";")
    On Error Resume Next
    mvar签名方式 = Val(T(0))
    mvar签名规则 = Val(T(1))
    mvar证书ID = Val(T(2))
    mvar显示手签 = (Val(T(3)) = 1)
    mvar签名时间 = Format(T(4), "yyyy-mm-dd hh:mm:ss")
    mvar显示时间 = T(5)
    If UBound(T) >= 6 Then
        mvar签名要素 = CStr(T(6))
    Else
        mvar签名要素 = ""
    End If
    
    If UBound(T) >= 7 Then
        mvar时间戳信息 = CStr(T(7))
    Else
        mvar时间戳信息 = ""
    End If
End Property

Public Property Get 对象属性() As String
    对象属性 = CStr(mvar签名方式) & ";" & CStr(mvar签名规则) & ";" & _
        CStr(mvar证书ID) & ";" & IIf(mvar显示手签, 1, 0) & ";" & _
        Format(mvar签名时间, "yyyy-mm-dd hh:mm:ss") & ";" & mvar显示时间 & ";" & mvar签名要素 & ";" & mvar时间戳信息
End Property

'################################################################################################################
'## 功能：  复制签名对象
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRSign
    '用于拷贝数据，产生一个新类对象
    Dim cNode As New cEPRSign
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.父ID = Me.父ID
    cNode.开始版 = Me.开始版
    cNode.终止版 = Me.终止版
    cNode.对象序号 = Me.对象序号
    cNode.姓名 = Me.姓名
    cNode.签名信息 = Me.签名信息
    cNode.前置文字 = Me.前置文字
    cNode.签名级别 = Me.签名级别
    cNode.对象属性 = Me.对象属性
    cNode.时间戳 = Me.时间戳
    cNode.签名要素 = Me.签名要素
    cNode.签名人ID = Me.签名人ID
    cNode.签名图片 = Me.签名图片
    Set Clone = cNode
End Function
Public Function FillElementMember(ByVal rs As ADODB.Recordset, ByVal strTableName As String) As Boolean
'功能:用数据记录填充类成员
    If rs Is Nothing Then Exit Function
    If rs.EOF Then Exit Function
    
    If NVL(rs!对象标记, 0) <> 0 Then Me.Key = NVL(rs!对象标记, 0)
    Me.ID = rs!ID
    Me.文件ID = NVL(rs!文件ID, 0)
    Me.父ID = NVL(rs!父ID, 0)
    Me.对象序号 = NVL(rs!对象序号, 0)
    Me.姓名 = Split(NVL(rs!内容文本, ";"), ";")(0)
    Me.签名信息 = NVL(rs!要素值域)
    Me.前置文字 = NVL(rs!要素名称)
    Me.签名级别 = NVL(rs!要素表示)
    Me.对象属性 = NVL(rs!对象属性)
    Me.开始版 = NVL(rs!开始版, 1)
    Me.终止版 = NVL(rs!终止版, 0)
    Me.时间戳 = NVL(rs!要素单位)
    Me.签名图片 = NVL(rs!输入形态, 0) = 1
    If UBound(Split(NVL(rs!内容文本), ";")) > 0 Then
        Me.签名人ID = Val(Split(NVL(rs!内容文本), ";")(1))
    End If
    
    FillElementMember = True
End Function
'################################################################################################################
'## 功能：  从数据库直接获取签名内容
'##
'## 参数：  eEdtType    :当前编辑方式
'##         lngID       :数据库中的对象ID号
'##         blnWithKey  :是否包含Key值一起恢复
'################################################################################################################
Public Function GetSignFromDB(ByVal lngID As Double, _
    Optional blnWithKey As Boolean = False, _
    Optional ByVal strTableName As String = "") As Boolean

    Dim rs As New ADODB.Recordset
    If strTableName = "" Then strTableName = "电子病历内容"
    
    gstrSQL = "Select ID, 父id, 文件id, 对象标记, 对象序号, 对象属性, 内容文本, 要素值域, 要素名称, 要素表示, 要素单位, 输入形态, 开始版, 终止版" & vbNewLine & _
                "From " & strTableName & vbNewLine & _
                "Where 对象类型 = 8 And ID = [1]"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "cEPRSign", lngID)
    If Not rs.EOF Then
        If blnWithKey And NVL(rs("对象标记"), 0) <> 0 Then Me.Key = NVL(rs("对象标记"), 0)
        Me.ID = rs("ID")
        Me.文件ID = NVL(rs("文件ID"), 0)
        Me.父ID = NVL(rs("父ID"), 0)
        Me.对象序号 = NVL(rs("对象序号"), 0)
        Me.姓名 = Split(NVL(rs("内容文本"), ";"), ";")(0)
        Me.签名信息 = NVL(rs("要素值域"))
        Me.前置文字 = NVL(rs("要素名称"))
        Me.签名级别 = NVL(rs("要素表示"))
        Me.对象属性 = NVL(rs("对象属性"))
        Me.开始版 = NVL(rs("开始版"), 1)
        Me.终止版 = NVL(rs("终止版"), 0)
        Me.时间戳 = NVL(rs("要素单位"))
        Me.签名图片 = NVL(rs("输入形态"), 0) = 1
        If UBound(Split(NVL(rs!内容文本), ";")) > 0 Then
            Me.签名人ID = Val(Split(NVL(rs("内容文本")), ";")(1))
        End If
        GetSignFromDB = True
    Else
        GetSignFromDB = False
    End If
End Function
Public Function InsertIntoEditor(ByRef edtThis As Object, Optional ByVal lngStartPos As Long = -1, Optional ByVal blnForceInsert As Boolean = False, Optional ByRef objDocument As Object) As Boolean
    '******************************************************************************************************************
    '功能： 插入指定级别签名到Editor中
    '参数： edtThis         :当前的编辑器控件
    '       lngStartPos     :当前位置
    '       blnForceInsert  :是否强制插入，在修改要素时如果前后都保护，这时需要强制插入
    '返回：
    '******************************************************************************************************************
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim blnForce As Boolean
    Dim lngLen As Long
    Dim strTmp As String
    Dim StrText As String
    Dim intLoop As Integer
    Dim lESS As Long, lESE As Long, lEES As Long, lEEE As Long
    Dim blnNeeded As Boolean
    Dim blnFinded As Boolean
    
    '如果签名内容为空，也退出
    If Trim(Me.姓名) = "" Then Exit Function
    
    '形成签名显示内容
    strTmp = Format(Me.Key, "00000000") & ",0,0)"
    StrText = mvar前置文字 & Me.姓名 & IIf(mvar显示手签, "，手签：_____________", "")
    StrText = StrText & IIf(Trim(mvar显示时间) = "", "", "，" & Format(mvar签名时间, mvar显示时间))
    lngLen = Len(StrText)
        
    '先检查是否有对应的签名要素，如果有，则签到签名要素中
    '------------------------------------------------------------------------------------------------------------------
    If Not (objDocument Is Nothing) Then
        For intLoop = 1 To objDocument.Elements.Count
            If objDocument.Elements(intLoop).替换域 = 1 And Trim(objDocument.Elements(intLoop).内容文本) = "" Then
                Select Case objDocument.Elements(intLoop).要素名称
                Case "经治医师签名"
                    If Me.签名级别 = cprSL_经治 Then Exit For
                Case "主治医师签名"
                    If Me.签名级别 = cprSL_主治 Then Exit For
                Case "主任医师签名"
                    If Me.签名级别 = cprSL_主任 Or Me.签名级别 = cprSL_正高 Then Exit For
                End Select
            End If
        Next

        If intLoop <= objDocument.Elements.Count And objDocument.Elements.Count > 0 Then
            blnFinded = FindKey(edtThis, "E", objDocument.Elements(intLoop).Key, lESS, lESE, lEES, lEEE, blnNeeded)
            If blnFinded Then
                lngStartPos = lESS
                objDocument.Elements(intLoop).内容文本 = StrText
                objDocument.Elements(intLoop).Refresh edtThis
                Me.签名要素 = objDocument.Elements(intLoop).Key
            End If
        End If
    End If
    
    
    '如果没有对应的签名要素，则签到当前光标位置处
    '------------------------------------------------------------------------------------------------------------------
    If lngStartPos = -1 Then lngStartPos = edtThis.Selection.StartPos
    
    bInKeys = IsBetweenAnyKeys(edtThis, lngStartPos + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    
    '保证不能插入关键字内部
    If bInKeys Then
        InsertIntoEditor = False
        Exit Function
    End If
        
    If blnForceInsert = False And edtThis.Range(lngStartPos, lngStartPos).Font.ForeColor = PROTECT_FORECOLOR And edtThis.Range(lngStartPos, lngStartPos).Font.Protected Then
        Exit Function
    End If
    
    With edtThis
        .Freeze
        blnForce = .ForceEdit
        .ForceEdit = True
        
        .Range(lngStartPos, lngStartPos).Font.Protected = False
        .Range(lngStartPos, lngStartPos).Font.Hidden = False
        
        .Range(lngStartPos, lngStartPos).Text = "SS(" & strTmp & StrText & "SE(" & strTmp
            
        .Range(lngStartPos, lngStartPos + 32 + lngLen).Font.Protected = True
        .Range(lngStartPos, lngStartPos + 16).Font.Hidden = True
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Hidden = False
        '删除线
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Strikethrough = (Me.终止版 > 0)
        '下划线
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.Underline = IIf(Me.姓名 = "", ELE_UNDERLINE, cprNone)
        '前景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.ForeColor = GetCharColor(Me.开始版, Me.终止版 + 1)
        '背景色
        .Range(lngStartPos + 16, lngStartPos + 16 + lngLen).Font.BackColor = ELE_BACKCOLOR
        .Range(lngStartPos + 16 + lngLen, lngStartPos + 32 + lngLen).Font.Hidden = True
        lngStartPos = lngStartPos + 32 + lngLen

                If Me.签名要素 <> "" Then '有签名要素时光标定位到要素这后，因为要素被隐藏
            lngStartPos = lngStartPos + 32 + lngLen
        End If

        .Range(lngStartPos, lngStartPos).Selected
        
        .ForceEdit = blnForce
        .UnFreeze
    End With

    InsertIntoEditor = True
End Function

'################################################################################################################
'## 功能：  刷新签名组的显示
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function Refresh(ByRef edtThis As Object) As Boolean
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "S", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True
            Dim ParaFmt As New cParaFormat
            Set ParaFmt = edtThis.Range(lKSE, lKES).Para.GetParaFmt

            .Range(lKSS, lKEE) = ""
            Me.InsertIntoEditor edtThis, lKSS, True
            .Range(lKSE, lKSE + Len(Me.前置文字 & Me.姓名)).Selected
            .Range(lKSE, lKES).Para.SetParaFmt ParaFmt
            .ForceEdit = blnForce
            .UnFreeze
        End With
        Refresh = True
    Else
        Refresh = False
    End If
End Function

'################################################################################################################
'## 功能：  从编辑器中删除签名组
'##
'## 参数：  edtThis         :当前的编辑器控件
'################################################################################################################
Public Function DeleteFromEditor(ByRef edtThis As Object, Optional ByRef objDocument As Object) As Boolean
    '保存诊治要素编辑结果
    Dim strTmp As String, lngKey As Long, blnForce As Boolean
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean

    bFinded = FindKey(edtThis, "S", Me.Key, lKSS, lKSE, lKES, lKEE, bNeeded)
    If bFinded Then
        
        With edtThis
            .Freeze
            blnForce = .ForceEdit
            .ForceEdit = True

            If Not (objDocument Is Nothing) Then
                If Me.签名要素 <> "" Then
                    '在签名后显示出已隐藏的对应级别的签名要素，以便再次签名时自动定位
                    
                    Dim lESS As Long
                    Dim lESE As Long
                    Dim lEES As Long
                    Dim lEEE As Long
                    
                    Dim blnNeeded As Boolean
                    Dim blnFinded As Boolean
        
                    blnFinded = FindKey(edtThis, "E", Me.签名要素, lESS, lESE, lEES, lEEE, blnNeeded)
                
                    If blnFinded Then
                        .Range(lESE, lEES).Font.Hidden = False
                    End If
                End If
            End If
            
            .Range(lKSS, lKEE) = ""
            
            .ForceEdit = blnForce
            .UnFreeze
        End With
        
        If blnFinded Then
            objDocument.Elements("K" & Me.签名要素).内容文本 = ""
            objDocument.Elements("K" & Me.签名要素).Refresh edtThis
        End If
                    
        DeleteFromEditor = True
    Else
        DeleteFromEditor = False
    End If
End Function

'################################################################################################################
'## 功能：  获取指定的签名集合保存到数据库的SQL语句
'##
'## 参数：  只用于“电子病历记录”，其他情况不允许加入签名
'##         ArraySQL()          :IN/OUT，   SQL数组
'##         lng序号             :IN，       对象序号，同级别签名采用同一序号，然后用内容行次区别
'################################################################################################################
Public Function GetSaveSQL(ByRef ArraySQL() As String, _
    ByVal lng序号 As Long) As Boolean
    
    Dim lngCount As Long
    Call CheckValidID     '检查ID有效性！
    
    Dim i As Long, lng行次 As Long, lngEleKey As Long
    
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    
    gstrSQL = "ZL_电子病历内容_Update(" & Me.ID & "," & Me.文件ID & "," & Me.开始版 & "," & Me.终止版 & "," & IIf(Me.父ID = 0, "NULL", Me.父ID) & "," & _
        lng序号 & ",8," & Me.Key & ",1,'" & Me.对象属性 & "',NULL,'" & Me.姓名 & ";" & Me.签名人ID & "',0,NULL,NULL,NULL,NULL,NULL,Null," & _
        "'" & Me.前置文字 & "',1,50,0,'" & Me.时间戳 & "'," & Me.签名级别 & "," & IIf(Me.签名图片, 1, 0) & ",'" & Me.签名信息 & "')"
    
    ArraySQL(lngCount) = gstrSQL
    GetSaveSQL = True
End Function
'################################################################################################################
'## 功能：  检查对象的ID是否有效
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidID()
    If Me.ID = 0 Then Me.ID = GetEPRContentNextId
End Sub

Public Function IsFinished() As Boolean
    IsFinished = (Trim(Me.姓名) <> "")
End Function

Public Property Let 签名图片(ByRef vData As Boolean)
    mbln签名图片 = vData
End Property

Public Property Get 签名图片() As Boolean
     签名图片 = mbln签名图片
End Property
