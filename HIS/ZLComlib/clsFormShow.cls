VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：调用功能展示另一个界面的处理
'命名规则：函数过程名称为通俗易懂的英文即可
'上级类：clsDatabase
'##############################################################################

Public Function VerifyPassWord(frmParent As Object, ByVal strPass As String, Optional ByVal strName As String, Optional ByVal strSex As String, Optional ByVal strOld As String, Optional blnPassEncode As Boolean = True) As Boolean
'功能：对密码进行验证
'参数：frmParent=显示的父窗体
'      strPass=正确的密码
'      strName,strSex,strOld=可选参数，病人姓名、性别、年龄，当不传入时不显示这个区域。
'      blnPassEncode-strPass是否传入的加密串
'返回：True=密码验证通过,False=取消输入，或连续3次输入错误的密码
    VerifyPassWord = gobjComLib.zlCommFun.VerifyPassWord(frmParent, strPass, strName, strSex, strOld, blnPassEncode)
End Function

Public Sub StopFlash()
    '功能： 停止并关闭等待的动态窗体
    '参数：
    '返回：
    Call gobjComLib.zlCommFun.StopFlash
End Sub

Public Sub ShowFlash(Optional strNote As String, Optional frmParent As Object)
    '功能： 显示等待的动态窗体
    '参数：
    '   strNote:提示信息
    '   frmParent：关于窗体的父窗体
    '返回：
    Call gobjComLib.zlCommFun.ShowFlash(strNote, frmParent)
End Sub

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String, Optional ByVal strSelectCaption As String, Optional ByRef strSelectInput As String, Optional ByVal strTextCaption As String, Optional ByVal lngTextLength As Long, Optional ByRef strTextInput As String, Optional ByVal strSort As String = "1,2,3", Optional ByVal blnSelectMust As Boolean) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=传入的日期的标题，如果<>""则显示日期控件，供用户输入日期，将DateInput返回。
'      strDateFormat=时间格式 格式""yyyy-MM-dd hh:mm:ss"其中HH为大写是24小时制"
'      strSelectCaption=选项的标题:选项1|1(1设为缺省),选项2|0|1(选择本选项时才允许填写，1代表日期，2代表文本框，0代表所有)···
'      strSelectInput=选项选择的返回值(返回选择的名称)
'      strTextCaption=文本框标题
'      lngTextLength=文本框最大录入长度
'      strTextInput=文本框的返回值
'      strSort=日期、选项、文本框的排序，日期=1，选项=2，文本=3，默认排序"1,2,3"
'      blnSelectMust=如果有单选项，则必须选择一个，否则提示。
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    ShowMsgBox = gobjComLib.zlCommFun.ShowMsgBox(strCaption, strInfo, strCmds, frmParent, vStyle, strDateCaption, DateInput, strDateFormat, strSelectCaption, strSelectInput, strTextCaption, lngTextLength, strTextInput, strSort, blnSelectMust)
End Function

Public Sub ShowTipInfo(ByVal lnghwnd As Long, ByVal strInfo As String, Optional blnMultiRow As Boolean, Optional blnOutline As Boolean, Optional lngMaxWidth As Long)
'功能：显示或者隐藏提示
'参数：lngHwnd=提示所针对的控件句柄,当传入为0时隐藏提示
'      strInfo=提示信息,当传入为空时隐藏提示
'      blnMultiRow=以一定的间距分行显示多行信息，每行按vbcrlf分隔
'      blnOutline=是否将每行文本中字符|前的文字做为提纲单独一行显示
'      lngMaxWidth=窗口的最大窗度，缺省为0表示按设计状态的窗体最大宽度为准
    Call gobjComLib.zlCommFun.ShowTipInfo(lnghwnd, strInfo, blnMultiRow, blnOutline, lngMaxWidth)
End Sub

Public Sub DeviceSetup(frmParent As Object, ByVal lngSys As Long, ByVal lngModule As Long)
'功能:根据不同的系统和模块,调用不同的设备配置
    Call gobjComLib.zlCommFun.DeviceSetup(frmParent, lngSys, lngModule)
End Sub

Public Function PatiIdentify(frmParent As Object, ByVal lngSys As Long, _
    ByVal lng病人ID As Long, ByVal cur金额 As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True, _
    Optional blnFamilyMoney As Boolean, _
    Optional strFamilyPatiIDs As String = "", _
    Optional bln刷卡验证 As Boolean = True, _
    Optional bln无密码不验卡 As Boolean) As Boolean
    '功能:对指定病人进行刷卡身份验证
    '入参:lngSys=调用程序的系统编号
    '       lng病人ID=要验证身份的病人ID
    '       cur金额=本次消费金额
    '       lngModul-模块号
    '       bytOperationType-业务类型(0-余额汇总显示;1-只显示门诊余额;2-只显示住院余额)
    '       lngDefaultCardTypeID-缺省的刷卡类别ID
    '       blnCheckPassWord-是否验证密码(true表示需要验证密码,False表示不验证密码,只验证卡)
    '        47216
    '       blnFamilyMoney-是否读取家属预交余额
    '       strFamilyPatiIDs-病人家属的病人ID，79868
    '       bln刷卡验证-是否进行刷卡验证，主要用于不刷卡验证时读取家属IDs
    '       bln无密码不验卡-病人的所有医疗卡都没有设置密码时是否验卡，当为True时，只要有一张卡设置了密码都要进行验卡,112418
    '返回:验证成功能返回True，失败或取消返回False
    '注意：1.只适用于标准版
    '      2.调用该函数的地方要具有"病人信息,病人余额,病人医疗卡信息"的查询权限
    PatiIdentify = gobjComLib.zlDatabase.PatiIdentify(frmParent, lngSys, lng病人ID, cur金额, lngModul, bytOperationType, _
                                                    lngDefaultCardTypeID, blnCheckPassWord, blnFamilyMoney, strFamilyPatiIDs, _
                                                    bln刷卡验证, bln无密码不验卡)
End Function

Public Function UserIdentifyByUser(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strFunc As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDefaultPreUser As Boolean, Optional ByRef strDBUser As String) As String
'参数：strNote=提示信息(简短)
'      lngProgID=程序序号
'      strFunc=授权功能
'      cnNew=要返回的连接,需要返回时,必须传入非Nothing的对象,并且需要由调用程序关闭连接；如果是当前登录用户,返回Nothing
'      blnDefaultPreUser-缺省显示上次登录人
'返回：成功返回人员姓名
'      strDBUser=返回输入的输入的数据库用户
    UserIdentifyByUser = gobjComLib.zlDatabase.UserIdentifyByUser(frmParent, strNote, lngSys, lngProgId, strFunc, cnNew, blnDefaultPreUser, strDBUser)
End Function

Public Function UserIdentify(frmParent As Object, ByVal strNote As String, ByVal lngSys As Long, ByVal lngProgId As Long, ByVal strUser As String, Optional cnNew As ADODB.Connection, Optional ByVal blnDBUser As Boolean, Optional ByRef strDBUser As String, Optional cnNewOLEDB As ADODB.Connection) As String
'参数：strNote=提示信息(简短)
'      lngProgID=程序序号
'      strUser=指定的用户
'      cnNew=要返回的连接,需要返回时,必须传入非Nothing的对象,并且需要由调用程序关闭连接；如果是当前登录用户,返回Nothing
'      blnDBUser=用数据库用户直接验证登录，并返回该用户创建的连接，此时参数lngProgId，strFunc，返回值无效
'      cnNewOLEDB=需要获取的OLEDB连接，和CNNEW都是同一用户，但是连接不同。当该参数不是Nothing时，才返回 OLEDB连接，否则不返回
'返回：成功返回人员姓名
'      strDBUser=返回输入的输入的数据库用户
    UserIdentify = gobjComLib.zlDatabase.UserIdentify(frmParent, strNote, lngSys, lngProgId, strUser, cnNew, blnDBUser, strDBUser, cnNewOLEDB)
End Function

Public Function ShowPatiColorTip(frmParent As Object) As Boolean
'功能:在frmParent窗口右下角显示一窗体，内容为各种病人类型的颜色说明
'说明：必须具有“病人类型”查询权限
    ShowPatiColorTip = gobjComLib.zlDatabase.ShowPatiColorTip(frmParent)
End Function

Public Function ShowRecSelect(ByVal frmMain As Object, ByVal objControl As Object, ByVal rsBindings As ADODB.Recordset, Optional ByRef rsOutSel As ADODB.Recordset, Optional ByVal blnShowHead As Boolean = False, _
                                                    Optional ByVal strHideCols As String = "", Optional ByVal lngSys As Long, _
                                                   Optional ByVal lngModule As Long, Optional ByVal str参数名 As String = "", Optional ByVal strColAlignment As String = "") As Boolean
    '功能:选择器入口
    '入参:frmMain-调用的主窗口
    '     lngSys-系统号
    '     lngModule-模块号
    '     objControl-控件对象(目前只支:textBox,Combox,VSFlexGrid,BILLEDIT)
    '     rsBindings-绑定的记录集(不能为空,主要字段,ID,......)(ID不再必须）
    '     blnShowHead-是否显示现列头
    '     str参数-个性化保存的参数名
    '     strColAlignment-需要对齐设置的各个列名称以及其对齐模式
    '         格式为：列名1|0,列名2|1,列名3|2,...其中0、1、2...8表示的对齐方式分别为：左上、左中、左下、中上、中中、中下、右上、右中、右下
    '出参:rsOutSel-选择后的记录集
    '返回:选中返回True, 否则返回False(可以按Esc进行返回)
    Dim frmNew As New frmListSel
    ShowRecSelect = frmNew.ShowSelect(frmMain, objControl, rsBindings, rsOutSel, blnShowHead, strHideCols, lngSys, lngModule, str参数名, strColAlignment)
End Function

Public Function ShowSQLSelectEx(frmParent As Object, ByVal objControl As Object, ByVal strSQL As String, bytStyle As Byte, _
                                                ByVal strTitle As String, ByVal bln末级 As Boolean, ByVal strSeek As String, ByVal strNote As String, _
                                                ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, ByVal blnNoneWin As Boolean, ByRef Cancel As Boolean, _
                                                ByVal blnMultiOne As Boolean, ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
                                                ByVal strOtherInfo As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：多功能选择器，使用ADO.Command打开,允许使用[x]参数
'           (替代以前的ShowSelect、ShowSQLMultiSelect、ShowSQLSelect）
'           传入输入框对象，就可以自动定位
'参数：
'     frmParent=显示的父窗体
'     objControl=调用界面输入框
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'             双表风格如果列名存在含Check结尾的字段，则该字段作为是否勾选的值存储字段。=1为勾选，0=不勾选。
'             双表风格如果列名存在*名称，*简码，*编码的，则显示右上角的查询功能，以供查询项目，
'                    编码列必须整个匹配，匹配成功后定位到该分类的该项目上，按F3支持查找下一个。
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin=处理成非窗体风格
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     blnMulti=是否允许多选
'     strOtherInfo=格式为：项目名称1=内容1#项目2=内容2#......
'                当前项目有：bytSize=0,1;字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体
'                            ColSet=列宽设置|列名1,宽度1,0;列名2,宽度2,1;.....|悬浮提示|列名。 其中宽度后面的一个参数表示该列的对齐方式,0、1和2分别表示左对齐、右对齐和中间对齐
'                            NotShowNon=0,1;0-默认处理，显示没有子项的分类，1-不显示没有子项的分类;bytStyle=2有作用
'                            HeadCap=SQL列名1,列表展示列名1;SQL列名2,列表展示列名2；该项目用来手工指定SQL列在列表中展示名称，一般用于编码名称列，但是不改变列的Key
'                            MultiCheckReturn=0,1：多选时只返回勾选行，由于多选点确定默认返回当前行所以增加该参数控制，该控制启用后，不支持默认行的返回，但是仍旧支持双击行自动返回。
'                            HideNullCols=0,1;是否隐藏SQl中的null as 写法的列
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    Dim frmNew As New frmPubSel
    Dim arrPar() As Variant
    arrPar = arrInput
    Set ShowSQLSelectEx = frmNew.ShowSelectV2(frmParent, objControl, strSQL, bytStyle, strTitle, bln末级, strSeek, strNote, _
                                                 blnShowSub, blnShowRoot, blnNoneWin, Cancel, blnMultiOne, blnSearch, blnMulti, _
                                                 strOtherInfo, arrPar)
End Function

Public Function ShowILLSelect(frmParent As Object, ByVal str类别 As String, ByVal lng病人科室ID As Long, Optional ByVal str性别 As String, Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True, Optional ByVal strSel As String, Optional ByVal lngSys As Long = 100) As ADODB.Recordset
'功能：ICD10疾病编码/疾病诊断选择器
'参数：str类别=多个疾病编码类别,形如"D,M",为空时表示所有
'         D:ICD-10疾病编码
'         Y:损伤中毒的外部原因
'         M:肿瘤形态学编码
'         S:ICD-9-CM3手术编码
'         B:中医疾病编码
'         Z:中医病征编码
'         1:疾病诊断-西医
'         2:疾病诊断-中医
'      str性别=病人性别,用于处理疾病的性别限制
'      blnMultiSel=是否允许进行复选
'      blnICD10=是否显示疾病编码选择器(str类别=D,Y,M,S,B,Z)，否则显示疾病诊断选择器(str类别=1,2)
'      strSel:已经选择的编码，（以逗号分割）
'      lngSys=调用的系统的系统号

'返回：包含选择疾病项目的记录集，
'      字段：项目ID,编码,序号,附码,附码ID,附码名称,名称,说明,编者,分类id,简码,疗效限制,分娩,是否病人,疾病ID,诊断ID
'      标准版调用：
'               ICD-10疾病编码选择器：项目ID,编码,序号,附码,名称,说明,诊断ID(其余字段为空或不读取)
'               疾病诊断选择器：项目ID,编码,名称,说明,编者,疾病ID(其余字段为空或不读取）
'      病案系统调用：项目ID,编码,序号,附码,附码ID,附码名称,名称,说明,简码,疗效限制,分娩,是否病人(其余字段为空或不读取)
'      如果取消选择，返回为Nothing
'注意：标准版或其他调用该函数时需具有以下权限
'               zl_疾病编码科室_Insert,Zl_疾病编码科室_Delete
'               疾病编码目录,疾病编码科室,疾病编码分类,疾病编码类别
'               Zl_疾病诊断科室_Insert,Zl_疾病诊断科室_Delete
'               疾病诊断目录,疾病诊断分类,疾病诊断属类
'               部门表,部门人员,部门性质说明,上机人员表
'               疾病诊断科室、疾病诊断对照
'       病案系统调用：zl_疾病编码科室_Insert,Zl_疾病编码科室_Delete
'               疾病编码目录,疾病编码科室,疾病编码分类,疾病编码类别
    Set ShowILLSelect = gobjComLib.zlDatabase.ShowILLSelect(frmParent, str类别, lng病人科室ID, str性别, blnMultiSel, blnICD10, strSel, lngSys)
End Function

