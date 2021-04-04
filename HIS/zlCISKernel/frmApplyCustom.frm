VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Begin VB.Form frmApplyCustom 
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   Icon            =   "frmApplyCustom.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   11550
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTmp 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3120
      ScaleHeight     =   375
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser webSub 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11535
      ExtentX         =   20346
      ExtentY         =   14631
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmApplyCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '=0新增，=1修改，=2查看
Private mint场合 As Integer '0 住院医生工作站，1 门诊医生工作站；
Private mint调用场合 As Integer '申请单调用场合，0－医生站调用，1－医嘱编辑界面调用。为1时允许用缓存数据加载界面和保存为缓存数据。
Private mint调用类型 As Integer  '1-门诊,2-住院
Private mint服务对象 As Integer '1-门诊,2-住院

Private mrsAppend As ADODB.Recordset
Private mobjFile As New FileSystemObject     '文件操作对象
Private mlng病人ID As Long
Private mstr挂号单 As String
Private mlng主页ID As Long
Private mlng病区ID As Long
Private mlng病人科室id As Long '病人科室id/挂号执行科室id
Private mlng病人性质 As Long   '0-住院，1-门诊
Private mlng开单科室ID As Long
Private mstr开单科室 As String '申请科室名称
Private mstr病人姓名 As String, mstr门诊号 As String, mstr住院号 As String, mstr床号 As String
Private mrsPati As Recordset

Private mintPState As Integer
Private mdatTurn As Date
Private mstr入院时间 As String
Private mlng申请序号 As Long  '申请序号，修改，查看时传入，否则为空
Private mlng项目ID As Long
Private mlng医嘱ID As Long '修改查看时存在
Private mlngXML文件ID As Long 'XML对应的文件ID
Private mlngFileID As Long   '申请单文件ID
Private mint医嘱状态 As Integer

Private mclsMipModule As zl9ComLib.clsMipModule '消息平台对象

Private mlng前提ID As Long
Private mbytBaby As Byte  '婴儿序号
Private mstrXSL As String, mstrXSLPath As String, mstrXSLFileName As String
Private mstrXML As String, mstrXMLPath As String, mstrXMLFileName As String
Private mstrHTML As String, mstrHTMLPath As String, mstrHTMLFileName As String
Private mstrFold As String
Private mblnOK As Boolean
Private mrsDefine As Recordset, mobjVBA As Object, mobjScript As clsScript
Private mint险类 As Integer
Private mobjEmrInterface As Object
Private mobjXML As Object
Private mlngOut医嘱ID As Long
Private astr128Code()
Private astr128A() As Variant
Private astr128B() As Variant
Private astr128C() As Variant
Private astr128ID() As Variant
Public Enum ImageFileFormat
    Bmp = 1
    Jpg = 2
    Png = 3
    Gif = 4
End Enum
Private Const EncoderQuality             As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID(0 To 3)        As Long
    NumberOfValues      As Long
    Type                As EncoderParameterValueType
    value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal Bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As Long

'---------------------------------------------------------------------------------------------------
'医嘱相关属性(跟xslt中约定的固定属性)

Private mrsAdvice As New Recordset
Private mintWebCompLete As Integer '加载完成后的操作，0-无操作；7-预览，6-打印

Public Function ShowMe(frmParent As Object, ByVal int场合 As Integer, ByVal intType As Integer, ByVal lng病人ID As Long, ByVal str就诊ID As String, ByVal lng病人性质 As Long, _
    Optional ByVal lngFileID As Long, Optional ByRef lng申请序号 As Long, Optional ByVal lng科室id As Long, Optional ByVal lng开单科室ID As Long, _
    Optional ByVal lng病区ID As Long, Optional ByVal rsDefine As Recordset, Optional ByVal intPState As Integer, Optional ByVal datTurn As Date, Optional ByVal int调用场合 As Integer, _
    Optional ByRef objMip As Object, Optional ByVal lng前提ID As Long, Optional ByVal bytBaby As Byte, Optional ByVal int险类 As Integer, Optional ByRef lngOut医嘱ID As Long, Optional ByRef lng项目id As Long) As Boolean
'功能：公共接口
'参数：frmParent 父对象窗体；int场合 0 住院医生工作站，1 门诊医生工作站； lng病人性质 0-住院，1-门诊；lng病人ID；
'      str就诊ID 跟据 int场合 判断，主页id/挂号单；
'      intType 操作类型   0-新增，1-修改，2-查看,3-医嘱编辑调用；lng医嘱ID 传入的医嘱ID；
'      lng科室ID  病人科室id/挂号执行科室id；int调用场合 0－工作站界面，1－ 医嘱编辑界面； lng开单科室ID 开嘱科室id；
'      strDefine 医嘱内容格式串；
'      lng病区ID、intPState、datTurn 住院才有；objMip 消息对象用于消息发送 住院才有；
'      lng项目ID-医嘱编辑界面输入项目ID
'      lngOut医嘱ID=出参，本次新增或修改的医嘱ID，用于定位
    
    mint场合 = int场合
    Set mrsAdvice = Nothing
    If mint场合 = 0 Then
        mlng主页ID = Val(str就诊ID)
        mint调用类型 = 2
        mint服务对象 = 2
    Else
        mstr挂号单 = str就诊ID
        mint调用类型 = 1
        mint服务对象 = 1
    End If
    mint调用场合 = int调用场合
    mlng病人ID = lng病人ID
    mlng病人性质 = lng病人性质
    mlng病人科室id = lng科室id
    mlng病区ID = lng病区ID
    mlng开单科室ID = lng开单科室ID
    mlngFileID = lngFileID
    mlng申请序号 = lng申请序号
    mintPState = intPState
    mintType = intType
    mdatTurn = datTurn
    mint险类 = int险类
    mlng项目ID = lng项目id
    Set mrsDefine = rsDefine

    mlng前提ID = lng前提ID
    mbytBaby = bytBaby
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    ShowMe = mblnOK
    lngOut医嘱ID = mlngOut医嘱ID
End Function

Private Sub SetXMLForLoad(ByRef strText As String)
'功能：XML预处理
    Dim strSQL As String, rsTmp As Recordset
    Dim objNodes As Object
    Dim objNode As Object
    Dim objNewNode As Object
    Dim objAttribute As Object
    Dim objAttributeNode As Object
    Dim lng执行科室ID As Long

    On Error GoTo errH
    strSQL = "Select I.ID, I.中文名" & vbNewLine & _
        "From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
        "Where I.分类id = K.ID And (K.性质 = 1 And K.编码 = '06' And" & vbNewLine & _
        "      I.中文名 Not In ('门诊诊断', '一次住院诊断', '二次住院诊断', '上次住院诊断')  Or k.性质 = 6)" & vbNewLine & _
        "Order By I.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call mobjXML.loadXML(strText)
    On Error Resume Next
    mobjXML.selectSingleNode(".//" & "姓名").Text = mstr病人姓名
    mobjXML.selectSingleNode(".//" & "性别").Text = mrsPati!性别 & ""
    mobjXML.selectSingleNode(".//" & "年龄").Text = mrsPati!年龄 & ""
    mobjXML.selectSingleNode(".//" & "婚姻状况").Text = mrsPati!婚姻状况 & ""
    mobjXML.selectSingleNode(".//" & "床号").Text = mrsPati!当前床号 & ""
    mobjXML.selectSingleNode(".//" & "住院号").Text = mrsPati!住院号 & ""
    mobjXML.selectSingleNode(".//" & "门诊号").Text = mrsPati!门诊号 & ""
    mobjXML.selectSingleNode(".//" & "科室").Text = mstr开单科室
    mobjXML.selectSingleNode(".//" & "科别").Text = mstr开单科室
    mobjXML.selectSingleNode(".//" & "籍贯").Text = mrsPati!籍贯 & ""
    mobjXML.selectSingleNode(".//" & "地址").Text = mrsPati!家庭地址 & ""
    mobjXML.selectSingleNode(".//" & "病人来源").Text = IIF(mint调用类型 = 1, "门诊", "住院")
    mobjXML.selectSingleNode(".//" & "患者来源").Text = IIF(mint调用类型 = 1, "门诊", "住院")
    mobjXML.selectSingleNode(".//" & "民族").Text = mrsPati!民族 & ""
    mobjXML.selectSingleNode(".//" & "职业").Text = mrsPati!职业 & ""
    mobjXML.selectSingleNode(".//" & "身份证号").Text = mrsPati!身份证号 & ""
    mobjXML.selectSingleNode(".//" & "身份证").Text = mrsPati!身份证号 & ""
    mobjXML.selectSingleNode(".//" & "联系人电话").Text = mrsPati!联系人电话 & ""
    mobjXML.selectSingleNode(".//" & "联系电话").Text = mrsPati!联系人电话 & ""
    mobjXML.selectSingleNode(".//" & "家庭电话").Text = mrsPati!家庭电话 & ""
    If Val(mrsPati!当前病区ID & "") <> 0 Then
        mobjXML.selectSingleNode(".//" & "当前病区").Text = Sys.RowValue("部门表", Val(mrsPati!当前病区ID & ""), "名称")
        mobjXML.selectSingleNode(".//" & "病区").Text = Sys.RowValue("部门表", Val(mrsPati!当前病区ID & ""), "名称")
    End If
    mobjXML.selectSingleNode(".//" & "送检日期").Text = Format(zlDatabase.Currentdate, "YYYY年MM月DD日")
    mobjXML.selectSingleNode(".//" & "送检医师").Text = UserInfo.姓名
    mobjXML.selectSingleNode(".//" & "开单医师").Text = UserInfo.姓名
    mobjXML.selectSingleNode(".//" & "生效时间").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    mobjXML.selectSingleNode(".//" & "执行时间").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    mobjXML.selectSingleNode(".//" & "采集时间").Text = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    
    
    '病历数据（支持申请附项的内容，要求XML里面必须有这个节点）
    Do While Not rsTmp.EOF
        err.Clear
        mobjXML.selectSingleNode(".//" & rsTmp!中文名).Text = ""
        '如果报错，说明没有这个要素，就不读取
        If err.Number = 0 Then
            mobjXML.selectSingleNode(".//" & rsTmp!中文名).Text = GetAppendItemValue(rsTmp!中文名, rsTmp!ID, rsTmp!中文名)
        End If
        rsTmp.MoveNext
    Loop
    '动态的诊疗项目ID
    strSQL = "select C.ID,C.类别,C.名称,C.执行科室,D.简码 from  病历单据应用 B,诊疗项目目录 C,诊疗项目别名 D where  B.诊疗项目ID=C.ID And C.ID=D.诊疗项目ID AND D.码类=1 and B.病历文件ID=[1] and b.应用场合=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngFileID, mint调用类型)
    If rsTmp.RecordCount > 0 Then
        Set objNodes = mobjXML.selectSingleNode(".//" & "诊疗项目ID")
        objNodes.Text = ""
        Do While Not rsTmp.EOF
            Set objNode = mobjXML.createNode("1", "项目", "")
            Set objNewNode = mobjXML.createNode("1", "ID", "")
            objNewNode.Text = rsTmp!ID & ""
            objNode.appendChild objNewNode
            Set objNewNode = mobjXML.createNode("1", "名称", "")
            objNewNode.Text = rsTmp!名称 & ""
            objNode.appendChild objNewNode
            Set objNewNode = mobjXML.createNode("1", "简码", "")
            objNewNode.Text = rsTmp!简码 & ""
            objNode.appendChild objNewNode
            lng执行科室ID = Get诊疗执行科室ID(mlng病人ID, mlng主页ID, rsTmp!类别 & "", Val(rsTmp!ID & ""), 0, Val(rsTmp!执行科室 & ""), mlng病人科室id, mlng开单科室ID, 1, mint调用类型, , , mint调用类型)
            Set objNewNode = mobjXML.createNode("1", "缺省科室ID", "")
            objNewNode.Text = lng执行科室ID
            objNode.appendChild objNewNode
            
            If mlng项目ID <> 0 And mlng项目ID = Val(rsTmp!ID & "") Then
                Set objAttributeNode = mobjXML.selectSingleNode("root/indt")
                Set objNewNode = mobjXML.createNode("1", "诊疗项目ID", "")
                objNewNode.Text = rsTmp!ID & ""
                objAttributeNode.appendChild objNewNode
                
                Set objAttributeNode = mobjXML.selectSingleNode("root/indt")
                Set objNewNode = mobjXML.createNode("1", "诊疗项目名称", "")
                objNewNode.Text = rsTmp!名称 & ""
                objAttributeNode.appendChild objNewNode
                
            End If
            
            objNodes.appendChild objNode
            rsTmp.MoveNext
        Loop
    End If
    
    strText = mobjXML.xml
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetAppendItemValue(ByVal str项目 As String, ByVal lng要素ID As Long, ByVal str中文名 As String) As String
'功能：获取指定的申请附项值
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strText As String
    Dim arrItem As Variant, i As Long
    Dim lng就诊ID As Long
    Dim intType As Integer '1-门诊，2－住院
    
    On Error GoTo errH
    
    '4.未取到或未对应要素的，从病人之前已保存的医嘱中提取,以最后填写的为准
    strSQL = " Select 内容 From (" & _
        " Select B.内容 From 病人医嘱记录 A,病人医嘱附件 B" & _
        " Where A.ID=B.医嘱ID And A.病人ID=[1] And Nvl(A.婴儿,0)=[4]" & _
        IIF(mint调用类型 = 1, " And A.挂号单=[2]", " And A.主页ID=[3]") & _
        " And B.项目=[5] And B.内容 is Not Null" & _
        " Order by A.开嘱时间 Desc) Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, mlng主页ID, mbytBaby, str项目)
    If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
    
    '1.如果有对应要素，从要素提取函数读取
    If lng要素ID <> 0 And strText = "" Then
        '先老版，再新版
        If mint调用类型 = 1 Then '门诊
            strSQL = "Select Zl_Replace_Element_Value(B.中文名,[1],A.ID,1) as 内容" & _
                " From 病人挂号记录 A,诊治所见项目 B Where A.NO=[2] And B.ID=[3] And a.记录性质=1 And a.记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, lng要素ID)
        Else
            strSQL = "Select Zl_Replace_Element_Value(中文名,[1],[2],2) as 内容 From 诊治所见项目 Where ID=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng要素ID)
        End If
        If Not rsTmp.EOF Then strText = NVL(rsTmp!内容)
        If strText = "" Then
            
            If mint调用类型 = 1 Then
                strSQL = "select a.id From 病人挂号记录 A Where A.NO=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号单)
                lng就诊ID = Val(rsTmp!ID & "")
                intType = 1
            Else
                lng就诊ID = mlng主页ID
                intType = 2
            End If
            strText = GetOrderInspectInfo(mlng病人ID, str中文名, intType, lng就诊ID)
        End If
    End If
    
    GetAppendItemValue = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetOrderInspectInfo(ByVal lng病人ID As Long, ByVal strCondition As String, ByVal intType As Integer, ByVal lng就诊ID As Long) As String
'功能：读取指定病人的指定提纲在病历填写的信息，例如：主诉，诊断等
    Dim strText As String
    On Error Resume Next
    If mobjEmrInterface Is Nothing Then
        Set mobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    If Not mobjEmrInterface Is Nothing Then
        strText = mobjEmrInterface.GetOrderInspectInfoEx(intType, lng病人ID, lng就诊ID, strCondition)
        If err.Number <> 0 Then
            strText = mobjEmrInterface.GetOrderInspectInfo(lng病人ID, strCondition)
        End If
    End If
    GetOrderInspectInfo = strText
End Function

Private Function LoadFile(ByVal lng文件ID As Long, ByVal strFile As String, ByVal int类型 As Integer, ByVal lng医嘱ID As Long) As Boolean
'功能：读取和创建本地文件
    Dim strText As String
    Dim stmTmp As Stream
    Dim strFilename As String
    Dim objNodes As Object
    Dim objNewNode As Object
    
    strFilename = strFile
    strFile = mstrFold & "\" & strFile
    If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    If mintType = 0 Or lng医嘱ID = 0 And mintType = 2 Then
        strText = Sys.ReadLob(glngSys, 24, mlngFileID & "," & int类型, strFile, 1)
    Else
        strText = Sys.ReadLob(glngSys, 25, lng医嘱ID & "," & int类型, strFile, 1)
    End If
    'XML预处理
    If int类型 = 2 Then
        '新开时进行预处理读取信息
        If mintType = 0 Then Call SetXMLForLoad(strText)
        On Error Resume Next
        If Not gobjPlugIn Is Nothing Then
            If gobjPlugIn.AdviceLoadApplyCustom(glngSys, IIF(mint场合 = 0, p住院医生站, p门诊医生站), mlng病人ID, IIF(mstr挂号单 = "", mlng主页ID, mstr挂号单), lng文件ID, strText, lng医嘱ID) = False Then
                If err.Number = 0 Then Exit Function
            End If
            Call zlPlugInErrH(err, "AdviceLoadApplyCustom")
        End If
        '医技站修改已发送的申请
        If mint医嘱状态 <> 1 And mintType = 1 Then
            Set objNodes = Nothing
            Call mobjXML.loadXML(strText)
            Set objNodes = mobjXML.selectSingleNode("root/已发送")
            If objNodes Is Nothing Then
                Set objNodes = mobjXML.selectSingleNode("root")
                Set objNewNode = mobjXML.createNode("1", "已发送", "")
                objNewNode.Text = 1
                objNodes.appendChild objNewNode
            Else
                objNodes.Text = 1
            End If
            strText = mobjXML.xml
        End If
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    If Not mobjFile.FolderExists(mstrFold) Then Call mobjFile.CreateFolder(mstrFold)
    If Not mobjFile.FileExists(strFile) Then Call mobjFile.CreateTextFile(strFile, True)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText strText
    stmTmp.SaveToFile strFile, adSaveCreateOverWrite
    stmTmp.Close
    If int类型 = 1 Then
        mstrXSL = strText
        mstrXSLPath = strFile
        mstrXSLFileName = strFilename
    ElseIf int类型 = 2 Then
        mstrXML = strText
        mstrXMLPath = strFile
        mstrXMLFileName = strFilename
    ElseIf int类型 = 3 Then
        mstrHTML = strText
        mstrHTMLPath = strFile
        mstrHTMLFileName = strFilename
    End If

    If Not mobjFile.FileExists(strFile) Then
        MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Function
    End If
    LoadFile = True
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_SaveExit  '保存
            Control.Enabled = mintType <> 2
    End Select
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As Recordset
    Dim strFileXSL As String, strFileXML As String, strFileHTML As String
    
    InitCommandBar
    
    mblnOK = False
    mintWebCompLete = 0
    On Error GoTo errH
    Set mobjXML = CreateObject("MSXML2.DOMDocument")
    If mobjXML Is Nothing Then
        MsgBox "创建MSXML2.DOMDocument对象失败", vbExclamation, Me.Caption
        Unload Me
        Exit Sub
    End If
    
    mint医嘱状态 = 1
    
    If mlng申请序号 = 0 Then
        strSQL = "Select A.文件ID,A.文件名, A.类别,0 as 医嘱ID,B.名称,1 AS 医嘱状态 From 自定义申请单文件 A,病历文件列表 B Where A.文件ID=B.ID And B.ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngFileID)
    Else
        strSQL = "Select a.文件ID,a.文件名, a.类别,A.医嘱ID,B.名称,C.医嘱状态" & vbNewLine & _
                "From 医嘱申请单文件 A,病历文件列表 B,病人医嘱记录 C " & vbNewLine & _
                "Where A.文件ID=B.ID And A.医嘱ID=C.ID And C.申请序号 = [1] And C.相关id Is Null"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng申请序号)
    End If
    
    Screen.MousePointer = 11
    If mlng病人ID <> 0 Then
        strSQL = "select 0 诊疗项目ID,0 收费细目ID,'' 医生嘱托,0 执行科室ID ,0 采集科室ID,sysdate 生效时间,sysdate 采集时间,0 采集方式ID,'' 标本部位 from dual"
        Set mrsAdvice = zlDatabase.CopyNewRec(zlDatabase.OpenSQLRecord(strSQL, Me.Caption))
        strSQL = "Select 姓名,性别,年龄,住院号,门诊号,当前床号,婚姻状况,籍贯,家庭地址,民族,职业,身份证号,联系人电话,家庭电话,当前病区ID from 病人信息 Where 病人ID=[1]"
        Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        mstr病人姓名 = mrsPati!姓名
        mstr住院号 = mrsPati!住院号 & ""
        mstr门诊号 = mrsPati!门诊号 & ""
        mstr床号 = mrsPati!当前床号 & ""
        mstr开单科室 = Sys.RowValue("部门表", mlng开单科室ID, "名称")
    End If
    
    mstrFold = mobjFile.GetSpecialFolder(TemporaryFolder) & "\" & Decode(mintType, 0, "新增", 1, "修改", 2, "查看") & "_" & mlngFileID & "_" & mlng申请序号
    If rsTmp.RecordCount > 0 Then
        mlng医嘱ID = Val(rsTmp!医嘱ID)
        mlngXML文件ID = Val(rsTmp!文件ID & "")
        Me.Caption = rsTmp!名称 & ""
        mint医嘱状态 = Val(rsTmp!医嘱状态 & "")
    End If
    Do While Not rsTmp.EOF
        If LoadFile(Val(rsTmp!文件ID & ""), rsTmp!文件名 & "", Val(rsTmp!类别), Val(rsTmp!医嘱ID)) = False Then Exit Sub
        rsTmp.MoveNext
    Loop
    
    '加载网页
    webSub.Navigate mstrHTMLPath
    
    Me.Width = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "W", Me.Width)
    Me.Height = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "H", Me.Height)

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    webSub.Top = 530
    webSub.Move 0, webSub.Top, Me.Width - 240, Me.Height - webSub.Top - 580
End Sub

Private Sub InitCommandBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim objCbo As CommandBarComboBox
    
    '工具栏----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '生成工具栏
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
        objControl.IconId = 815
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印退出")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, " 保存退出(&S)")
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, " 退出(&X)"): objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsMain.KeyBindings
        .Add FALT, vbKeyX, conMenu_File_Exit
        .Add FALT, vbKeyS, conMenu_Edit_Save
    End With

End Sub

Private Sub PrintPreView(ByVal intType As Integer)
'功能：预览打印
'参数：intType=6打印；=7预览；=8打印设置
    Dim stmTmp As Stream
    Dim b() As Byte
    Dim strTmp As String
    Dim picSet As StdPicture
    Dim objNode As Object
    
    If (intType = 7 Or intType = 6) Then
        If mintType <> 2 Then
            If SaveData = False Then Exit Sub
        End If
        
        mintWebCompLete = intType
        webSub.Navigate mstrHTMLPath
        If mintType = 0 Then mintType = 1
    Else
        Call webSub.ExecWB(intType, 1)
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet: PrintPreView 8
        Case conMenu_File_Preview: PrintPreView 7
        Case conMenu_File_Print: PrintPreView 6
        Case conMenu_Edit_SaveExit  '保存
            If SaveData = True Then Unload Me
        Case conMenu_File_Exit
            Unload Me
    End Select
End Sub

Private Function SetXML(ByVal vTag As Object) As Boolean
    Dim objNodes As Object
    Dim objNode As Object
    Dim objNewNode As Object
    Dim objAttribute As Object
    Dim i As Long
    Dim j As Long
    Dim blnSelect As Boolean
    Dim blnIsSel As Boolean '是否存在""选中""节点
    Dim strNodeName As String
    Dim strGroupName As String
    Dim objNewNodeGroup As Object
    
    On Error Resume Next
    SetXML = True
    'ID不为空，且设置了ZLREAD属性的才解析
    If vTag.ID = "" Or vTag.zlread <> "1" Then Exit Function
    If vTag.zlnotnull = "1" Then
        If err.Number <> 0 Then
            err.Clear
        Else
            If vTag.value = "" Then
                MsgBox "项目：[" & vTag.ID & "]为必填项目，请检查并填写。", vbExclamation, Me.Caption
                SetXML = False
                Exit Function
            End If
        End If
    End If
    
    If UCase(vTag.tagName) = "INPUT" And (vTag.Type = "text" Or vTag.Type = "password" Or vTag.Type = "hidden") Or UCase(vTag.tagName) = "TEXTAREA" And UCase(vTag.Type) = "TEXTAREA" Then
        '单一文本框录入的处理
        err.Clear
        mobjXML.selectSingleNode("root/indt/" & vTag.ID).Text = vTag.value
        If err.Number <> 0 Then
            '创建节点
            strNodeName = "root/indt"
            Set objNodes = mobjXML.selectSingleNode(strNodeName)
            For i = 0 To UBound(Split(vTag.ID, "/"))
                Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & Split(vTag.ID, "/")(i))
                If objNodes Is Nothing Then
                    Set objNewNode = mobjXML.createNode("1", Split(vTag.ID, "/")(i), "")
                    If i = UBound(Split(vTag.ID, "/")) Then objNewNode.Text = vTag.value
                    Set objNodes = mobjXML.selectSingleNode(strNodeName)
                    objNodes.appendChild objNewNode
                End If
                strNodeName = strNodeName & "/" & Split(vTag.ID, "/")(i)
            Next
            err.Clear
        End If
    ElseIf UCase(vTag.tagName) = "SELECT" And vTag.Type = "select-one" Then
        '单选项选择处理
        Set objNodes = mobjXML.selectSingleNode("root/indt/" & vTag.ID)
        If Not objNodes Is Nothing And err.Number = 0 Then
            objNodes.Text = vTag.Options(vTag.selectedIndex).Text
            objNodes.Attributes.getNamedItem("valueid").value = vTag.value
        ElseIf vTag.ID <> "" Then
            Set objNodes = mobjXML.selectSingleNode("root/indt")
            Set objNewNode = mobjXML.createNode("1", vTag.ID, "")
            objNewNode.Text = vTag.Options(vTag.selectedIndex).Text
            Set objAttribute = mobjXML.createAttribute("valueid")
            objAttribute.value = vTag.value
            objNewNode.Attributes.setNamedItem objAttribute
            objNodes.appendChild objNewNode
        End If
    ElseIf UCase(vTag.tagName) = "INPUT" And (vTag.Type = "radio" Or vTag.Type = "checkbox") Then
        '单选项和多选项选择处理
        strGroupName = ""
        strGroupName = vTag.groupname
        err.Clear
        mobjXML.selectSingleNode("root/indt/" & IIF(strGroupName = "", "", strGroupName & "/") & vTag.ID).Text = IIF(vTag.Checked, 1, 0)
        If err.Number <> 0 Then
            '创建节点
            strNodeName = "root/indt"
            Set objNodes = mobjXML.selectSingleNode(strNodeName)
            For i = 0 To UBound(Split(vTag.ID, "/"))
                Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & IIF(strGroupName = "", "", strGroupName & "/") & Split(vTag.ID, "/")(i))
                If objNodes Is Nothing Then
                    Set objNewNode = mobjXML.createNode("1", Split(vTag.ID, "/")(i), "")
                    If i = UBound(Split(vTag.ID, "/")) Then objNewNode.Text = IIF(vTag.Checked, 1, 0)
                    If strGroupName <> "" Then
                        Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & strGroupName)
                        If objNodes Is Nothing Then
                            Set objNewNodeGroup = mobjXML.createNode("1", strGroupName, "")
                            Set objNodes = mobjXML.selectSingleNode(strNodeName)
                            objNodes.appendChild objNewNodeGroup
                            Set objNodes = mobjXML.selectSingleNode(strNodeName & "/" & strGroupName)
                        End If
                    Else
                        Set objNodes = mobjXML.selectSingleNode(strNodeName)
                    End If
                    
                    objNodes.appendChild objNewNode
                End If
                strNodeName = strNodeName & "/" & Split(vTag.ID, "/")(i)
            Next
            err.Clear
        End If
    End If
End Function

Private Function SaveData() As Boolean
    Dim vDoc, vTag
    Dim i As Long, j As Long
    Dim strSQL As String
    Dim rs项目 As Recordset, rs采集 As Recordset
    Dim datCur As Date, strDate As String
    Dim arrSQL As Variant
    Dim str检验计价性质 As String, str采集计价性质 As String, str检验执行性质 As String, str采集执行性质 As String
    Dim str医嘱内容 As String, lng序号 As Long, str摘要 As String, lng医嘱ID As Long, lng相关ID As Long
    Dim blnCancel As Boolean
    Dim str开始时间 As String, str采集时间 As String
    Dim rsTmp As Recordset, blnTrans As Boolean
    Dim arrTmp() As String
    Dim stmTmp As Stream
    Dim b() As Byte
    Dim strTmp As String
    Dim picSet As StdPicture
    Dim objNode As Object
    
    On Error GoTo errH
    Set vDoc = webSub.Document
    '组织XML
    Call mobjXML.loadXML(mstrXML)
    For i = 0 To vDoc.All.Length - 1
        If UCase(vDoc.All(i).tagName) = "INPUT" Or UCase(vDoc.All(i).tagName) = "SELECT" Or UCase(vDoc.All(i).tagName) = "TEXTAREA" Then
            Set vTag = vDoc.All(i)
            If vTag.Type = "text" Or vTag.Type = "password" Or vTag.Type = "hidden" Or vTag.Type = "select-one" Or UCase(vTag.Type) = "TEXTAREA" Then
            '----------------------------------------------------------------------------
            '获取xslt约定的项目值(以后在这里扩充)
                If vTag.ID = "诊疗项目ID" Then
                    mrsAdvice!诊疗项目ID = Val(vTag.value)
                ElseIf vTag.ID = "收费细目ID" Then
                    mrsAdvice!收费细目ID = Val(vTag.value)
                ElseIf vTag.ID = "医生嘱托" Then
                    mrsAdvice!医生嘱托 = vTag.value
                ElseIf vTag.ID = "执行科室ID" Then
                    mrsAdvice!执行科室ID = Val(vTag.value)
                ElseIf vTag.ID = "采集科室ID" Then
                    mrsAdvice!采集科室ID = Val(vTag.value)
                ElseIf vTag.ID = "生效时间" Then
                    If IsDate(vTag.value) Then mrsAdvice!生效时间 = CDate(vTag.value)
                ElseIf vTag.ID = "采集时间" Then
                    If IsDate(vTag.value) Then mrsAdvice!采集时间 = CDate(vTag.value)
                ElseIf vTag.ID = "采集方式ID" Then
                    mrsAdvice!采集方式ID = Val(vTag.value)
                ElseIf vTag.ID = "标本部位" Then
                    mrsAdvice!标本部位 = vTag.value
                End If
                If SetXML(vTag) = False Then Exit Function
            ElseIf vTag.Type = "submit" Then
                vTag.Click
            ElseIf vTag.Type = "radio" Or vTag.Type = "checkbox" Then
                If SetXML(vTag) = False Then Exit Function
            End If
        End If
    Next i
    mstrXML = mobjXML.xml
    webSub.Stop
    
    
    '转换和读取数据
    If mrsAdvice!诊疗项目ID = 0 Then
        MsgBox "未选择诊疗项目，请选择一个项目！", vbExclamation, Me.Caption
        Exit Function
    End If
    Set rs项目 = Get诊疗项目记录(mrsAdvice!诊疗项目ID)
    
    datCur = zlDatabase.Currentdate
    If mrsAdvice!生效时间 = CDate(0) Then mrsAdvice!生效时间 = datCur
    If mrsAdvice!采集时间 = CDate(0) Then mrsAdvice!采集时间 = datCur
    If mrsAdvice!标本部位 & "" = "" Then mrsAdvice!标本部位 = rs项目!标本部位
    mrsAdvice.Update
    
    
    '数据检查
    If CheckData(rs项目) = False Then Exit Function
    
    '数据保存
    arrSQL = Array()
    If mlng申请序号 <> 0 Then
        strSQL = "select a.id from 病人医嘱记录 a where a.申请序号=[1] and A.相关ID is null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng申请序号)
        For i = 1 To rsTmp.RecordCount
            If Not (mint医嘱状态 <> 1 And mintType = 1) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & rsTmp!ID & ",1)"
            End If
            lng相关ID = Val(rsTmp!ID)
            rsTmp.MoveNext
        Next
    End If
    If Not (mint医嘱状态 <> 1 And mintType = 1) Then
        str开始时间 = "To_Date('" & Format(mrsAdvice!生效时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        strDate = "To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        If mlng申请序号 = 0 Then
            strSQL = "Select 病人医嘱记录_申请序号.Nextval as 申请序号 From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            mlng申请序号 = Val(rsTmp!申请序号)
        End If
        lng序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, mbytBaby, mstr挂号单)
        If rs项目!类别 = "C" Then
            str采集时间 = "To_Date('" & Format(mrsAdvice!采集时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            str检验计价性质 = Val("" & rs项目!计价性质)
            str检验执行性质 = IIF("" & rs项目!执行科室 = "", "NULL", "" & rs项目!执行科室)
            str医嘱内容 = rs项目!名称
            lng序号 = lng序号 + 1
            str摘要 = ""
            str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(mrsAdvice!诊疗项目ID) & "||" & IIF(mlng病人性质 = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint调用类型, mlng病人ID, mlng主页ID, mint险类, "C", mrsAdvice!诊疗项目ID, mlng开单科室ID, UserInfo.姓名, mrsAdvice!执行科室ID, Val(rs项目!执行科室 & ""), str摘要 & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            lng医嘱ID = zlDatabase.GetNextID("病人医嘱记录")
            If lng相关ID = 0 Then lng相关ID = zlDatabase.GetNextID("病人医嘱记录")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng医嘱ID & "," & lng相关ID & "," & lng序号 & "," & mint调用类型 & "," & mlng病人ID & "," & _
                ZVal(mlng主页ID) & "," & mbytBaby & ",1,1,'C'," & mrsAdvice!诊疗项目ID & ",Null,Null,Null,1," & _
                "'" & str医嘱内容 & "',Null," & "'" & mrsAdvice!标本部位 & "','一次性',Null," & _
                "Null,Null,Null," & str检验计价性质 & "," & mrsAdvice!执行科室ID & _
                "," & str检验执行性质 & "," & 0 & "," & str开始时间 & ",Null," & mlng病人科室id & "," & _
                mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & strDate & "," & IIF(mstr挂号单 = "", "NULL", "'" & mstr挂号单 & "'") & "," & ZVal(mlng前提ID) & "," & _
                "Null,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                ",Null,Null,Null,Null," & mlng申请序号 & ")"
                
            '采集方式
            Set rsTmp = Get诊疗项目记录(mrsAdvice!采集方式ID)
            str采集计价性质 = Val("" & rsTmp!计价性质)
            str采集执行性质 = "" & rsTmp!执行科室
            str医嘱内容 = AdviceTextMake(rs项目)
            lng序号 = lng序号 + 1
            str摘要 = ""
            str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(mrsAdvice!采集方式ID) & "||" & IIF(mlng病人性质 = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint调用类型, mlng病人ID, mlng主页ID, mint险类, "E", mrsAdvice!采集方式ID, mlng开单科室ID, UserInfo.姓名, mrsAdvice!采集科室ID, Val(rsTmp!执行科室 & ""), str摘要 & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng相关ID & ",Null," & lng序号 & "," & mint调用类型 & "," & mlng病人ID & "," & _
                ZVal(mlng主页ID) & "," & mbytBaby & ",1,1,'E'," & mrsAdvice!采集方式ID & ",Null,Null,Null,1," & _
                "'" & str医嘱内容 & "','" & mrsAdvice!医生嘱托 & "'," & "'" & mrsAdvice!标本部位 & "','一次性',Null," & _
                "Null,Null,Null," & str采集计价性质 & "," & mrsAdvice!采集科室ID & _
                "," & str采集执行性质 & "," & 0 & "," & str采集时间 & ",Null," & mlng病人科室id & "," & _
                mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & strDate & "," & IIF(mstr挂号单 = "", "NULL", "'" & mstr挂号单 & "'") & "," & ZVal(mlng前提ID) & "," & _
                "Null,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                ",Null,Null,Null,Null," & mlng申请序号 & ")"
        Else
            str检验计价性质 = Val("" & rs项目!计价性质)
            str检验执行性质 = IIF("" & rs项目!执行科室 = "", "NULL", "" & rs项目!执行科室)
            str医嘱内容 = AdviceTextMake(rs项目)
            lng序号 = lng序号 + 1
            str摘要 = ""
            str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(mrsAdvice!诊疗项目ID) & "||" & IIF(mlng病人性质 = 1, 1, 2))
            blnCancel = CheckLISAppAdvice(mint调用类型, mlng病人ID, mlng主页ID, mint险类, "C", mrsAdvice!诊疗项目ID, mlng开单科室ID, UserInfo.姓名, mrsAdvice!执行科室ID, Val(rs项目!执行科室 & ""), str摘要 & "||0||0|| ||0")
            If Not blnCancel Then Exit Function
            
            If lng相关ID = 0 Then lng相关ID = zlDatabase.GetNextID("病人医嘱记录")
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & lng相关ID & ",NULL," & lng序号 & "," & mint调用类型 & "," & mlng病人ID & "," & _
                ZVal(mlng主页ID) & "," & mbytBaby & ",1,1,'" & rs项目!类别 & "'," & mrsAdvice!诊疗项目ID & ",Null,Null,Null,1," & _
                "'" & str医嘱内容 & "',Null," & "'" & mrsAdvice!标本部位 & "','一次性',Null," & _
                "Null,Null,Null," & str检验计价性质 & "," & mrsAdvice!执行科室ID & _
                "," & str检验执行性质 & "," & 0 & "," & str开始时间 & ",Null," & mlng病人科室id & "," & _
                mlng开单科室ID & ",'" & UserInfo.姓名 & "'," & strDate & "," & IIF(mstr挂号单 = "", "NULL", "'" & mstr挂号单 & "'") & "," & ZVal(mlng前提ID) & "," & _
                "Null,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                ",Null,Null,Null,Null," & mlng申请序号 & ")"
        End If
        
        If mstr挂号单 <> "" Then
            '门诊默认绑定所有门诊诊断
            strSQL = "Select A.ID from 病人诊断记录 A,病人挂号记录 B Where A.病人ID=B.病人ID AND A.主页ID=B.ID AND A.病人ID=[1] and B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
            Do While Not rsTmp.EOF
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(NULL,NULL," & rsTmp!ID & ",'" & lng相关ID & "')"
                rsTmp.MoveNext
            Loop
        End If
        
        If mintType = 0 Then
            Call mobjXML.loadXML(mstrXML)
            On Error Resume Next
            err.Clear
            Set objNode = mobjXML.selectSingleNode(".//" & "条码")
            If Not objNode Is Nothing Then
                Set picSet = DrawBarCode128Auto(pictmp, "1" & lng相关ID, 0, 2, True)
                'IE8对JPG有兼容问题，所以改为用PNG
                SaveStdPicToFile picSet, mstrFold & "条码.png", Png
                Open mstrFold & "条码.png" For Binary As #1
                ReDim b(LOF(1) - 1)
                Get #1, , b
                Close #1
                strTmp = Base64Encode(b)
                objNode.Text = "data:image/jpeg;base64," & strTmp
                
                mstrXML = mobjXML.xml
                
            End If
        End If
    End If
    '重新存储到本地，以便于打印预览
    If Not mobjFile.FolderExists(mstrFold) Then Exit Function
    If mobjFile.FileExists(mstrXMLPath) Then Call mobjFile.DeleteFile(mstrXMLPath): Call mobjFile.CreateTextFile(mstrXMLPath, True)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText mstrXML
    stmTmp.SaveToFile mstrXMLPath, adSaveCreateOverWrite
    stmTmp.Close
                
                
    '处理文件SQL
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        If gobjPlugIn.AdviceSaveApplyCustom(glngSys, IIF(mint场合 = 0, p住院医生站, p门诊医生站), mlng病人ID, IIF(mstr挂号单 = "", mlng主页ID, mstr挂号单), mlngXML文件ID, mstrXML, webSub, mlng医嘱ID) = False Then
            If err.Number = 0 Then Exit Function
        End If
        Call zlPlugInErrH(err, "AdviceSaveApplyCustom")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    For j = 1 To 3
         ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_医嘱申请单文件_Edit(" & mlngFileID & ",'" & Decode(j, 1, mstrXSLFileName, 2, mstrXMLFileName, 3, mstrHTMLFileName) & "'," & j & "," & lng相关ID & ")"
        ReDim arrTmp(0)
        If Not Sys.GetLobSql(glngSys, 25, lng相关ID & "," & j, Replace(Decode(j, 1, mstrXSL, 2, mstrXML, 3, mstrHTML), "'", "''"), arrTmp(), 1) Then
            MsgBox "文件添加到数据库失败，无法保存医嘱！", vbExclamation, Me.Caption
            Exit Function
        End If
        For i = LBound(arrTmp) To UBound(arrTmp)
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = arrTmp(i)
        Next
    Next

    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    If Not (mint医嘱状态 <> 1 And mintType = 1) Then
        Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, mstr病人姓名, mstr住院号, , IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mlng病区ID, , mlng病人科室id, "", , mstr床号, _
           lng相关ID, 0, 1, rs项目!类别, "", UserInfo.姓名, str开始时间, mlng开单科室ID, "", , , "")
    End If
    mlngOut医嘱ID = lng相关ID
    mblnOK = True
    SaveData = True
    Exit Function
errH:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData(rsSQL As Recordset) As Boolean
'功能：数据检查
    If rsSQL!类别 & "" = "C" Then
        If mrsAdvice!采集方式ID = 0 Then
            MsgBox "检验项目必须选择一个采集方式，请选择一个采集方式！", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    '检查表单数据
    On Error Resume Next
    If Not gobjPlugIn Is Nothing Then
        If gobjPlugIn.AdviceCheckApplyCustom(glngSys, IIF(mint场合 = 0, p住院医生站, p门诊医生站), mlng病人ID, IIF(mstr挂号单 = "", mlng主页ID, mstr挂号单), mlngXML文件ID, mstrXML, webSub, mlng医嘱ID) = False Then
            If err.Number = 0 Then Exit Function
        End If
        Call zlPlugInErrH(err, "AdviceCheckApplyCustom")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    CheckData = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mobjFile.FolderExists(mstrFold) Then mobjFile.DeleteFolder mstrFold, True
    Set mobjFile = Nothing
    Set mobjVBA = Nothing
    Set mobjEmrInterface = Nothing
    '根据文件ID来存储窗体大小
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "W", Me.Width
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mlngFileID & "", "H", Me.Height
    
End Sub

Private Function AdviceTextMake(rsSQL As Recordset) As String
'功能：获取医嘱内容文本
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean
    Dim strText As String, strSQL As String
    Dim strField As String, int频率范围 As Integer
    Dim i As Long, k As Long
    Dim blnDo As Boolean, str类别 As String
    Dim str中药 As String, str煎法 As String, str形态 As String
    Dim str麻醉 As String, str附术 As String
    Dim str检验 As String, str标本 As String
    Dim str部位 As String, str部位Last As String, str方法 As String
    Dim dbl数量 As Double, str中药名称 As String
    Dim str中药诊疗项目IDS As String, strSame As String
    
    On Error GoTo errH
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
    
    '确定是否定义
    str类别 = rsSQL!类别
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "诊疗类别='" & str类别 & "'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!医嘱内容)) = "" Then
            blnDefine = False
        End If
    End If
    
ReDoDefault: '用于按定义公式计算失败，重新按缺省规则进行组织
    strText = ""
    If blnDefine Then strText = mrsDefine!医嘱内容
    
    '产生医嘱内容
    Select Case str类别
    Case "C" '检验-------------------------------------------------------------
        str检验 = "": str标本 = ""

        str检验 = rsSQL!名称
        str标本 = mrsAdvice!标本部位 & ""
        If str标本 = "" Then str标本 = rsSQL!标本部位 & ""
        
        If Not blnDefine Then
            strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
        Else
            If InStr(strText, "[检验项目]") > 0 Then
                strField = str检验
                strText = Replace(strText, "[检验项目]", """" & strField & """")
            End If
            If InStr(strText, "[检验标本]") > 0 Then
                strField = str标本
                strText = Replace(strText, "[检验标本]", """" & strField & """")
            End If
            If InStr(strText, "[采集方法]") > 0 Then
                If mrsAdvice!采集方式ID <> 0 Then
                    strSQL = "select 类别,名称,标本部位 from 诊疗项目目录 where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mrsAdvice!采集方式ID)
                    strField = rsTmp!名称 & ""
                Else
                    strField = ""
                End If
                
                strText = Replace(strText, "[采集方法]", """" & strField & """")
            End If
        End If
    Case "D" '检查-------------------------------------------------------------
        str部位 = "": str方法 = ""
        
        strText = rsSQL!名称
    Case "F" '手术-------------------------------------------------------------
        str麻醉 = "": str附术 = ""
        strText = rsSQL!名称

    Case "8" '中药配方---------------------------------------------------------
        str中药 = "": str煎法 = "": str中药诊疗项目IDS = "": strSame = ""
        strText = rsSQL!名称
    Case "4" '卫材------------------------------------------------------------
        If Val(mrsAdvice!收费细目ID & "") <> 0 Then
            strSQL = "Select 名称,规格,产地 From 收费项目目录 Where ID=[1]"
            Set rsTmp = New ADODB.Recordset '清除Filter
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsAdvice!收费细目ID & ""))
            
            If Not blnDefine Then
                strText = rsSQL!名称
                If Not IsNull(rsTmp!规格) Then
                    strText = strText & " " & rsTmp!规格
                End If
            Else
                If InStr(strText, "[卫生材料]") > 0 Then
                    strField = rsTmp!名称
                    strText = Replace(strText, "[卫生材料]", """" & strField & """")
                End If
                If InStr(strText, "[规格]") > 0 Then
                    strField = NVL(rsTmp!规格)
                    strText = Replace(strText, "[规格]", """" & strField & """")
                End If
                If InStr(strText, "[产地]") > 0 Then
                    strField = NVL(rsTmp!产地)
                    strText = Replace(strText, "[产地]", """" & strField & """")
                End If
            End If
        End If
    Case "5", "6" '西成药，中成药---------------------------------------------
        strText = rsSQL!名称
    Case "K" '输血医嘱
        strText = rsSQL!名称
    Case Else '其它所有类别-----------------------------------------------------
        If Not blnDefine Then
            strText = rsSQL!名称
        Else
            If InStr(strText, "[诊疗项目]") > 0 Then
                strField = rsSQL!名称
                strText = Replace(strText, "[诊疗项目]", """" & strField & """")
            End If
        End If
        '术后医嘱特殊显示
        If str类别 = "Z" And (Val(rsSQL!操作类型 & "") = 4 Or Val(rsSQL!操作类型 & "") = 14) Then
            strText = "━━━" & strText & "━━━"
        End If
        '转科医嘱特殊显示
        If str类别 = "Z" And Val(rsSQL!操作类型 & "") = 3 Then
            strText = "━━━" & strText & "━━━"
        End If
    End Select
    
    '公共字段或可以公共处理的字段-------------------------------------------
    If blnDefine Then
        If InStr(strText, "[开始时间]") > 0 Then
            strField = mrsAdvice!生效时间
            strText = Replace(strText, "[开始时间]", """" & strField & """")
        End If
        If InStr(strText, "[医生嘱托]") > 0 Then
            strField = mrsAdvice!医生嘱托 & ""
            If mrsAdvice!医生嘱托 & "" <> "" Then
                If strField <> "" Then
                    strField = strField & "," & mrsAdvice!医生嘱托
                Else
                    strField = mrsAdvice!医生嘱托
                End If
            End If
            strText = Replace(strText, "[医生嘱托]", """" & strField & """")
        End If
    End If
            
    '计算医嘱内容
    If blnDefine Then
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            err.Clear: On Error GoTo errH
            blnDefine = False: GoTo ReDoDefault
        End If
    End If
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub webSub_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If mintWebCompLete <> 0 Then
        Call webSub.ExecWB(mintWebCompLete, 2)
        If mintWebCompLete = 6 Then Unload Me
        mintWebCompLete = 0
    End If
End Sub

Private Sub WebSub_DownloadBegin()
    webSub.Silent = True
End Sub

Private Sub WebSub_DownloadComplete()
    webSub.Silent = True
End Sub


Public Function DrawBarCode128Auto(ByVal PicObj As Object, ByVal strBarCode As String, sngPrintWidth As Single, _
                            Optional ByVal intLineWidth As Integer = 2, Optional ByVal blnShowBarCodeTxt As Boolean = True) As StdPicture
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                       传入图片对象画图像
    '参数                       PicObj              Picture对象(用于画图）
    '                           strBarCode          生成条码的内容
    '                           sngPrintWidth       打印时使用的固定长度单位（mm)(返回）
    '                   可选
    '                           intLineWidth        线宽，用于控制条码的宽度 默认为2
    '                           blnShowBarCodeTxt   是否显示条码内容，默认True为显示
    '返回                       Image对象
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strGetBarCode As String
    Dim intval As Integer
    Dim lngTxtHeight As Long
    
    If intLineWidth = 0 Then intLineWidth = 2
    
    With PicObj
        .Cls
        .BackColor = vbWhite
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .DrawWidth = intLineWidth
        .Height = 1335
    End With
    
    strGetBarCode = GetBarCode(strBarCode)
    PicObj.Width = (2 + Len(strGetBarCode) * intLineWidth) * Screen.TwipsPerPixelX
    
    If blnShowBarCodeTxt Then
        PicObj.FontSize = 12
        PicObj.FontName = "Arial"
        lngTxtHeight = PicObj.TextHeight(strBarCode)
        
        PicObj.CurrentX = (PicObj.ScaleWidth / 2) - PicObj.TextWidth(strBarCode) / 2   ' 水平坐标
        PicObj.CurrentY = PicObj.ScaleHeight - lngTxtHeight - 1 ' 垂直坐标
        PicObj.Print strBarCode
    End If
    
    For intLoop = 1 To Len(strGetBarCode)
        intval = Mid$(strGetBarCode, intLoop, 1)
        If blnShowBarCodeTxt Then
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - lngTxtHeight - 2), IIF(intval = 1, vbBlack, vbWhite), BF
        Else
            PicObj.Line (1 + (intLoop * intLineWidth), 1)-(1 + (intLoop * intLineWidth), PicObj.ScaleHeight - 1), IIF(intval = 1, vbBlack, vbWhite), BF
        End If
    Next
    
    PicObj.Width = (2 + intLoop * intLineWidth) * Screen.TwipsPerPixelX
    
    sngPrintWidth = PicObj.ScaleWidth / 8
    
    Set DrawBarCode128Auto = PicObj.Image
    
    '恢复缺省以避免其它作图受影响
    PicObj.ScaleMode = vbTwips
    PicObj.DrawWidth = 1
    PicObj.FontName = "宋体"
    PicObj.FontSize = 9
End Function

Private Sub initBarCode()
    '//539
    '初始化相当内容
    astr128Code = Array( _
             "11011001100", "11001101100", "11001100110", "10010011000", "10010001100", "10001001100", _
             "10011001000", "10011000100", "10001100100", "11001001000", "11001000100", "11000100100", _
             "10110011100", "10011011100", "10011001110", "10111001100", "10011101100", "10011100110", _
             "11001110010", "11001011100", "11001001110", "11011100100", "11001110100", "11101101110", _
             "11101001100", "11100101100", "11100100110", "11101100100", "11100110100", "11100110010", _
             "11011011000", "11011000110", "11000110110", "10100011000", "10001011000", "10001000110", _
             "10110001000", "10001101000", "10001100010", "11010001000", "11000101000", "11000100010", _
             "10110111000", "10110001110", "10001101110", "10111011000", "10111000110", "10001110110", _
             "11101110110", "11010001110", "11000101110", "11011101000", "11011100010", "11011101110", _
             "11101011000", "11101000110", "11100010110", "11101101000", "11101100010", "11100011010", _
             "11101111010", "11001000010", "11110001010", "10100110000", "10100001100", "10010110000", _
             "10010000110", "10000101100", "10000100110", "10110010000", "10110000100", "10011010000", _
             "10011000010", "10000110100", "10000110010", "11000010010", "11001010000", "11110111010", _
             "11000010100", "10001111010", "10100111100", "10010111100", "10010011110", "10111100100", _
             "10011110100", "10011110010", "11110100100", "11110010100", "11110010010", "11011011110", _
             "11011110110", "11110110110", "10101111000", "10100011110", "10001011110", "10111101000", _
             "10111100010", "11110101000", "11110100010", "10111011110", "10111101110", "11101011110", _
             "11110101110", "11010000100", "11010010000", "11010011100", "1100011101011" _
             )
             
    astr128A = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "NUL", "SOH", _
             "STX", "ETX", "EOT", "ENQ", "ACK", "BEL", _
             "BS", "HT", "LF", "VT", "FF", "CR", _
             "SO", "SI", "DLE", "DC1", "DC2", "DC3", _
             "DC4", "NAK", "SYN", "ETB", "CAN", "EM", _
             "SUB", "ESC", "FS", "GS", "RS", "US", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "CODEB", "FNC4", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    
    astr128B = Array( _
             "SP", "!", """", "#", "$", "%", _
             "&", "'", "(", ")", "*", "+", _
             ",", "-", ".", "/", "0", "1", _
             "2", "3", "4", "5", "6", "7", _
             "8", "9", ":", ";", "<", "=", _
             ">", "?", "@", "A", "B", "C", _
             "D", "E", "F", "G", "H", "I", _
             "J", "K", "L", "M", "N", "O", _
             "P", "Q", "R", "S", "T", "U", _
             "V", "W", "X", "Y", "Z", "[", _
             "\", "]", "^", "_", "`", "a", _
             "b", "c", "d", "e", "f", "g", _
             "h", "i", "j", "k", "I", "m", _
             "n", "o", "p", "q", "r", "s", _
             "t", "u", "v", "w", "x", "y", _
             "z", "{", "|", "}", "~", "DEL", _
             "FNC3", "FNC2", "SHIFT", "CODEC", "FNC4", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
             
    astr128C = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "CODEB", "CODEA", _
             "FNC1", "StartA", "StartB", "StartC", "Stop" _
             )
    astr128ID = Array( _
             "0", "1", "2", "3", "4", "5", _
             "6", "7", "8", "9", "10", "11", _
             "12", "13", "14", "15", "16", "17", _
             "18", "19", "20", "21", "22", "23", _
             "24", "25", "26", "27", "28", "29", _
             "30", "31", "32", "33", "34", "35", _
             "36", "37", "38", "39", "40", "41", _
             "42", "43", "44", "45", "46", "47", _
             "48", "49", "50", "51", "52", "53", _
             "54", "55", "56", "57", "58", "59", _
             "60", "61", "62", "63", "64", "65", _
             "66", "67", "68", "69", "70", "71", _
             "72", "73", "74", "75", "76", "77", _
             "78", "79", "80", "81", "82", "83", _
             "84", "85", "86", "87", "88", "89", _
             "90", "91", "92", "93", "94", "95", _
             "96", "97", "98", "99", "100", "101", _
             "102", "103", "104", "105", "106" _
             )
End Sub

Private Function FindArray(strChar As String, strArray() As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           查找字符在数据中的位置
    '参数           strChar 传入字符
    '               strArray() 要查找的数组
    '返回           位置index
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    FindArray = 0
    For intLoop = 0 To UBound(strArray)
        If strChar = strArray(intLoop) Then
            FindArray = intLoop
            Exit For
        End If
    Next
    
End Function


Private Function FindCode(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:              按规则查找固定的字符
    '参数:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar 传入字符
    '返回:              对应的编码规则
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindCode = astr128Code(intIndex)
    End If
    
    If intType = 4 Then
        intIndex = FindArray(strChar, astr128ID)
        FindCode = astr128Code(intIndex)
    End If
End Function

Private Function GetBarCode(strChar As String) As String
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           按传入字符成生条码线规则
    '参数           strChar = 传入字符串
    '返回           条码成生的规则
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    Dim strTmp As String
    Dim intType As Integer
    Dim lngCheckCount As Long
    Dim intRow As Integer
    intType = 0
    intRow = 0
    
    For intLoop = 1 To Len(strChar) Step 2
        strTmp = Mid$(strChar, intLoop, 2)
        If Len(strTmp) = 2 And IsNumeric(strTmp) = True Then
            '128C码
            If intType = 0 Then
                GetBarCode = FindCode(3, "StartC")
                lngCheckCount = FindID(3, "StartC")
                intRow = intRow + 1
                intType = 3
            End If
            If intType <> 3 Then
                '转为128C码
                GetBarCode = GetBarCode & FindCode(1, "CODEC")
                lngCheckCount = lngCheckCount + intRow * FindID(1, "CODEC")
                intRow = intRow + 1
            End If
            
            
            GetBarCode = GetBarCode & FindCode(3, Val(strTmp))
            lngCheckCount = lngCheckCount + intRow * FindID(3, Val(strTmp))
            intRow = intRow + 1
        
            intType = 3
        Else
             '128A码
            If intType = 0 Then
                GetBarCode = FindCode(1, "StartA")
                lngCheckCount = FindID(1, "StartA")
                intRow = intRow + 1
                intType = 1
            End If
            If intType <> 1 Then
                '转为128A码
                GetBarCode = GetBarCode & FindCode(3, "CODEA")
                lngCheckCount = lngCheckCount + intRow * FindID(3, "CODEA")
                intRow = intRow + 1
            End If
            If Len(strTmp) = 1 Then
                GetBarCode = GetBarCode & FindCode(1, strTmp)
                lngCheckCount = lngCheckCount + intRow * FindID(1, strTmp)
                intRow = intRow + 1
            Else
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 1, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 1, 1))
                intRow = intRow + 1
              
                GetBarCode = GetBarCode & FindCode(1, Mid(strTmp, 2, 1))
                lngCheckCount = lngCheckCount + intRow * FindID(1, Mid(strTmp, 2, 1))
                intRow = intRow + 1
            End If
            intType = 1
        End If
    Next
    lngCheckCount = lngCheckCount Mod 103
    GetBarCode = GetBarCode & FindCode(4, CStr(lngCheckCount)) & FindCode(3, "Stop")
End Function

Private Function FindID(intType As Integer, strChar As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能:              按规则查找固定的字符的ID
    '参数:              intType(1=128A,2=128B,3=128C,4=ID)
    '                   strChar 传入字符
    '返回:              对应的编码的ID
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intIndex As Integer
    
    Call initBarCode
    
    If intType = 1 Then
        intIndex = FindArray(strChar, astr128A)
        FindID = Val(astr128ID(intIndex))
    End If
    
    If intType = 3 Then
        intIndex = FindArray(strChar, astr128C)
        FindID = Val(astr128ID(intIndex))
    End If
End Function


Function Base64Encode(Str() As Byte) As String                                  'Base64 编码
    On Error GoTo over                                                          '排错
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(Str) + 1) Mod 3   '除以3的余数
    Length = UBound(Str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIF(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10 + (Str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (Str(Length + 1) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function

Public Function SaveStdPicToFile(Stdpic As StdPicture, ByVal FileName As String, _
                              Optional ByVal FileFormat As ImageFileFormat = Jpg, _
                              Optional ByVal JpgQuality As Long = 80, _
                              Optional Resolution As Single) As Boolean
                              
    Dim CLSID(3)        As Long
    Dim Bitmap          As Long
    Dim Token           As Long
    Dim Gsp             As GdiplusStartupInput

    Gsp.GdiplusVersion = 1                      'GDI+ 1.0版本
    GdiplusStartup Token, Gsp                   '初始化GDI+
    GdipCreateBitmapFromHBITMAP Stdpic.Handle, Stdpic.hPal, Bitmap
    If Bitmap <> 0 Then                          '说明我们成功的将StdPic对象转换为GDI+的Bitmap对象了
        GdipBitmapSetResolution Bitmap, Resolution, Resolution
        Select Case FileFormat
        Case ImageFileFormat.Bmp
            If Not GetEncoderClsID("Image/bmp", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Jpg                    'JPG格式可以设置保存的质量
            Dim aEncParams()        As Byte
            Dim uEncParams          As EncoderParameters
            If GetEncoderClsID("Image/jpeg", CLSID) <> -1 Then
                uEncParams.Count = 1                                        ' 设置自定义的编码参数，这里为1个参数
                If JpgQuality < 0 Then
                    JpgQuality = 0
                ElseIf JpgQuality > 100 Then
                    JpgQuality = 100
                End If
                ReDim aEncParams(1 To Len(uEncParams))
                With uEncParams.Parameter
                    .NumberOfValues = 1
                    .Type = EncoderParameterValueTypeLong                   ' 设置参数值的数据类型为长整型
                    Call CLSIDFromString(StrPtr(EncoderQuality), .GUID(0))  ' 设置参数唯一标志的GUID，这里为编码品质
                    .value = VarPtr(JpgQuality)                                ' 设置参数的值：品质等级，最高为100，图像文件大小与品质成正比
                End With
                CopyMemory aEncParams(1), uEncParams, Len(uEncParams)
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), aEncParams(1)) = 0)
            End If
        Case ImageFileFormat.Png
            If Not GetEncoderClsID("Image/png", CLSID) = -1 Then
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        Case ImageFileFormat.Gif
            If Not GetEncoderClsID("Image/gif", CLSID) = -1 Then                '如果原始的图像是24位，则这个函数会调用系统的调色板来将图像转换为8位，转换的效果会不尽人意,但也有可能系统不自动转换，保存失败
                SaveStdPicToFile = (GdipSaveImageToFile(Bitmap, StrPtr(FileName), CLSID(0), ByVal 0) = 0)
            End If
        End Select
    End If
    GdipDisposeImage Bitmap      '注意释放资源
    GdiplusShutdown Token       '关闭GDI+。
End Function

Private Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim Size        As Long
    Dim i           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, Size               '得到解码器数组的大小
    If Size <> 0 Then
       ReDim Info(1 To Num) As ImageCodecInfo       '给数组动态分配内存
       ReDim Buffer(1 To Size) As Byte
       GdipGetImageEncoders Num, Size, Buffer(1)            '得到数组和字符数据
       CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)     '复制类头
       For i = 1 To Num             '循环检测所有解码
           If (StrComp(PtrToStrW(Info(i).MimeType), strMimeType, vbTextCompare) = 0) Then         '必须把指针转换成可用的字符
               CopyMemory ClassID(0), Info(i).ClassID(0), 16  '保存类的ID
               GetEncoderClsID = i      '返回成功的索引值
               Exit For
           End If
       Next
    End If
End Function

Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function

