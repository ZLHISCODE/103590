VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================================================================
'                0、模块变量与公共类型
'===========================================================================
'文件类型
Public Enum FileType
    FT_Public = 0                   '产品公共部件
    FT_Apply = 1                    '产品应用部件
    FT_Help = 2                     '产品帮助文件
    FT_AdditionFile = 3             '产品附加文件
    FT_Other = 4                    '三方产品文件
    FT_System = 5                   '系统文件
End Enum

'===========================================================================
'                1、保存属性的变量
'===========================================================================
'FileName文件名称(R/W)
Private mstrFileName                As String
'FileSTName文件标准名称，全大写(R)
Private mstrFileSTName              As String
'FileType文件类型(R/W)
Private mftFileType                 As FileType
'FilePath文件全路径(R/W)
Private mstrFilePath                As String
'SetupPath安装路径(R/W)
Private mstrSetupPath               As String
'FileVerion文件版本(R/W)
Private mstrFileVerion              As String
'ModifyDate最后修改时间(R/W)
Private mdatModifyDate              As Date
'MD5文件MD5(R/W)
Private mstrMD5                     As String
'IsForceCover强制覆盖(R/W)
Private mblnForceCover              As Boolean
'RFTReg注册(R/W)
Private mrftReg                     As RegFileType
'===========================================================================
'                2、属性定义
'===========================================================================
'FileName文件名称(R)
Public Property Get FileName() As String
    FileName = mstrFileName
End Property
'FileSTName文件标准名称，全大写(R)
Public Property Get FileSTName() As String
    FileSTName = mstrFileSTName
End Property
'FileType文件类型(R)
Public Property Get FileType() As FileType
    FileType = mftFileType
End Property
'FilePath文件全路径(R)
Public Property Get FilePath() As String
    FilePath = mstrFilePath
End Property
'SetupPath安装路径(R)
Public Property Get SetupPath() As String
    SetupPath = mstrSetupPath
End Property
'FileVerion文件版本(R)
Public Property Get FileVerion() As String
    FileVerion = mstrFileVerion
End Property
'ModifyDate最后修改时间(R)
Public Property Get ModifyDate() As Date
    ModifyDate = mdatModifyDate
End Property
'MD5文件MD5(R)
Public Property Get MD5() As String
    MD5 = mstrMD5
End Property
'IsForceCover强制覆盖(R)
Public Property Get IsForceCover() As Boolean
    IsForceCover = mblnForceCover
End Property
'注册Type(R)
Public Property Get RegType() As RegFileType
    RegType = mrftReg
End Property
'===========================================================================
'                3、公共方法
'===========================================================================
Public Function InitFile(ByVal blnUpBase As Boolean, ByVal strFileName As String, ByVal ftFileType As FileType, ByVal strSetupPath As String, ByVal strFileVerion As String, _
                        ByVal datModifyDate As Date, ByVal strMD5 As String, Optional ByVal blnForceCover As Boolean, Optional ByVal rftRegType As RegFileType, _
                        Optional ByVal strRelatedComponent As String) As Boolean
'功能：初始化文件信息
'参数：
'       blnUpBase:是否升级所需要的基础部件。基础部件在预升级时不会放在提前升级目录，会正式升级的部件。
'               升级基础部件包括：ZLRUNAS.EXE,ZLHISCRUST.EXE,REGCOM.DLL,7Z.EXE,7Z.DLL
'               ZLRUNAS.EXE-使自动升级以管理员运行的工具。ZLHISCRUST.EXE-自动升级工具。REGCOM.DLL-注册部件的部件（已经弃用了）
'               7Z.EXE,7Z.DLL:压缩解压部件，进行服务器部件的解压缩
'       strFileName:文件名称
'       ftFileType：文件类型
'       strSetupPath:安装路径
'       strFileVerion:文件版本
'       datModifyDate:最后修改时间
'       strMD5:文件MD5
'       blnForceCover:强制覆盖
'       blnReg:注册
'返回：是否成功
    On Error GoTo ErrH
    mstrFileName = strFileName
    mstrFileSTName = UCase(strFileName)
    mftFileType = ftFileType
    mstrSetupPath = strSetupPath
    mstrFileVerion = strFileVerion
    mdatModifyDate = datModifyDate
    mstrMD5 = strMD5
    mblnForceCover = blnForceCover
    mrftReg = rftRegType
    mstrFilePath = GetActualPath(strSetupPath, ftFileType, strFileName)
    InitFile = mstrFilePath = ""
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'===========================================================================
'                4、私有方法
'===========================================================================

