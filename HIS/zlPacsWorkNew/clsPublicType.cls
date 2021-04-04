VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum TImageType
    mtTagImage = 0      '标记图
    mtReportImage = 1   '报告图
    mtStudyImage = 2    '检查图
End Enum


Public Enum TReportImgType
    ritReport = 0   '报告图
    ritMark = 1     '标记图
End Enum


Public Enum TOutlineType
    otNone = 0
    otDesc = 1      '描述
    otOpin = 2      '意见   '诊断意见
    otAdvi = 3      '建议
End Enum


Public Enum TImgProcessType '图像处理窗口类型
    ptPreview = 0
    ptProcess = 1
    ptMark = 2
End Enum

Public Type TReportImgTag
    strKey As String        '对象关键字
    lngFromAdvice As Long    '图像来源的医嘱
    lngFileId As Double       '文件ID
    lngTableId As Double      '表格ID
    strImgFile As String    '图像文件
    strObjectTag As String  '对象标记
    strPros As String       '对象属性
    strImgMarks As String   '图像标记
    lngStartVer As String   '开始版本
    lngImgType As TReportImgType    '图像类型
End Type


Public Enum TImgMarkType
    imtNormal = 0       '不进行标记
    imtAuto = 1         '自动标记
    imtSpecify = 2      '指定标记
End Enum

Public Type TReportSignInfo
    ID As Double
    文件ID As Double
    父ID As Double
    Key As String
    对象序号 As Double
    姓名 As String
    签名信息 As String
    前置文字 As String
    签名级别 As String
    对象属性 As String
    签名版本 As Long
    开始版 As Long
    终止版 As Long
    签名方式 As Long
    签名规则 As Long
    时间戳 As String
    时间戳信息 As String
    证书ID As Long
    签名图片 As Boolean
    签名人ID As String
    签名要素 As String
    显示手签 As Boolean
    签名时间 As String
    显示时间 As String
End Type



Public Type TDicomBaseInfo
    lngAdviceId As Long         '医嘱ID
    lngSendNo As Long           '发送号
    '-----------------------------------------
    strModality As String       '设备
    lngID As Long               '患者ID
    strName As String           '患者姓名
    strSex As String            '性别
    strAge As String            '年龄
    strBirthDate As String      '出生日期
    
    strStudyUID As String       '检查uid
    strSeriesUID As String      '序列UID
    strInstanceUID As String    '实例UID
    lngSeriesNo As Long         '序列号
    lngImgNo As Long            '图像号
    strReceiveFullTime As String    '接受日期
    
    strInstitution As String    '机构名称
    
    strSeriesDes As String      '序列描述
    strImageDes As String       '图像描述
    
    strMediaEncode As String    '多媒体编码
    lngMediaTag As Long         '多媒体标记
    lngMediaLen As Long         '多媒体长度
    strDeviceNo As String       '设备号
End Type
