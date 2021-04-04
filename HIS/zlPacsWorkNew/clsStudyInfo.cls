VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStudyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public lngPatId As Long                 '1 病人ID
Public lngPageID As Long                  '2 主页ID
Public lngAdviceId As Long                '3 医嘱ID
Public lngSendNo As Long                  '4 发送号
Public strPatientName As String           '5 病人姓名
Public strPatientSex As String
Public strPatientAge As String
Public strNO As String
Public lngRecordKind As Long
Public blnIsPrinted As Boolean           '已经打印过

Public lngPatDept As Long                 '6 病人所属科室
Public strRegNo As String                 '7 挂号单
'Public lngRegId As Long                   '8 挂号id
Public blnMoved As Boolean                 '9 是否转出
Public intState As Integer                '10 检查状态
Public intStep As Integer                 '11 检查过程
Public lngUnit As Long                    '12 病区ID
Public strStudyUID As String              '13 检查UID
Public blnCanPrint As Boolean             '14 是否能够打印
Public blnIsInsidePatient As Boolean      '15 是否门诊或住院病人
Public lngExeDepartmentId As Long         '16 执行部门ID
Public strDoDoctor As String              '17 检查技师
Public strExeRoom As String               '18 执行间
Public lngPatientFrom As Long             '19 病人来源
Public strStudyNum As String              '20 检查号
'Public strBedNum As String                '21 床号
Public strMarkNum As String               '22 标志号
Public lngBaby As Long                    '23 婴儿
Public blnInfancy As Boolean                     '独立身份小于一岁也算婴儿(pacs业务专用)
Public strReportDoctor As String          '25 报告人
Public strReportOperation As String       '26 报告操作
Public lngLinkId As Long                  '27 关联ID
Public strImgType As String               '28 影像类别
Public intImageLocation As Integer        '29 PACS影像所在的位置，0在中联PACS；1在新网PACS
'Public strPatientType As String           '30 病人类型
Public intFilmGiveOut As Integer          '胶片发放
Public intReportGiveOut As Integer        '报告发放
Public intDangerState As Integer          '危急状态
Public intEmergentTag As Integer          '紧急状态
Public intGreenChannel As Integer         '绿色通道
Public strAdviceContext As String             '医嘱内容
Public strAdviceDepartAndMethod As String     '部位方法
Public strStuStateDesc As String          '检查状态描述
Public blnIsTechincalSure As Boolean      '是否技师确认
Public lngMoneyState As Long            '费用状态描述
Public blnIsReported As Boolean           '已经有报告
Public blnIsInfectious As Boolean             '是否传染病  LSQ待修正

Public strReportQuality As String   '报告质量
Public strImageQuality As String   '影像质量
Public strFollowUpDescribe As String '随访描述
Public intPathoType As Integer '病理检查类别
Public intPositive As Integer '结果阳性
Public strAccord As String '符合情况
Public strIIDNumber As String '身份证号
Public strCollectionInfo As String '收藏信息
Public lngReportType As Long    '报告编辑器类型（根据实际执行科室进行关联）1-pacs报告编辑器，2-病历编辑器，3-智能编辑器
Public lngReportEditState As Long   '报告可编辑状态
Public dtFirstTime As Date          '首次时间

Public dtRefreshDate As Date '上次刷新检查信息时间
Public lngClinicId As Long   '诊疗项目ID
Public lng门诊执行状态 As Long   '门诊执行状态



Public Function IsEquals(objStudyInfo As clsStudyInfo) As Boolean
    IsEquals = False
    
    With objStudyInfo
        If .blnCanPrint <> blnCanPrint Then Exit Function
        If .blnInfancy <> blnInfancy Then Exit Function
        
        If .blnIsInfectious <> blnIsInfectious Then Exit Function
        If .blnIsInsidePatient <> blnIsInsidePatient Then Exit Function
'        If .blnIsReported <> blnIsReported Then Exit Function
        If .blnIsTechincalSure <> blnIsTechincalSure Then Exit Function
        If .blnMoved <> blnMoved Then Exit Function
'        If .dtRefreshDate <> dtRefreshDate Then Exit Function
        If .intDangerState <> intDangerState Then Exit Function
        If .intEmergentTag <> intEmergentTag Then Exit Function
'        If .intFilmGiveOut <> intFilmGiveOut Then Exit Function
        If .intGreenChannel <> intGreenChannel Then Exit Function
'        If .intImageLocation <> intImageLocation Then Exit Function
        If .intPathoType <> intPathoType Then Exit Function
        If .intPositive <> intPositive Then Exit Function
'        If .intReportGiveOut <> intReportGiveOut Then Exit Function
        If .intState <> intState Then Exit Function
        If .intStep <> intStep Then Exit Function
        If .lngAdviceId <> lngAdviceId Then Exit Function
        If .lngBaby <> lngBaby Then Exit Function
        If .lngExeDepartmentId <> lngExeDepartmentId Then Exit Function
        If .lngLinkId <> lngLinkId Then Exit Function
        If .lngMoneyState <> lngMoneyState Then Exit Function
        If .lngPageID <> lngPageID Then Exit Function
        If .lngPatDept <> lngPatDept Then Exit Function
        If .lngPatId <> lngPatId Then Exit Function
        If .lngPatientFrom <> lngPatientFrom Then Exit Function
'        If .lngRecordKind <> lngRecordKind Then Exit Function
'        If .lngReportType <> lngReportType Then Exit Function
        If .lngSendNo <> lngSendNo Then Exit Function
        If .lngUnit <> lngUnit Then Exit Function
'        If .strAccord <> strAccord Then Exit Function
        If .strAdviceContext <> strAdviceContext Then Exit Function
'        If .strAdviceDepartAndMethod <> strAdviceDepartAndMethod Then Exit Function
'        If .strCollectionInfo <> strCollectionInfo Then Exit Function
        If .strDoDoctor <> strDoDoctor Then Exit Function
        If .strExeRoom <> strExeRoom Then Exit Function
'        If .strFollowUpDescribe <> strFollowUpDescribe Then Exit Function
        If .strIIDNumber <> strIIDNumber Then Exit Function
'        If .strImageQuality <> strImageQuality Then Exit Function
        If .strImgType <> strImgType Then Exit Function
        If .strMarkNum <> strMarkNum Then Exit Function
        If .strNO <> strNO Then Exit Function
        If .strPatientAge <> strPatientAge Then Exit Function
        If .strPatientName <> strPatientName Then Exit Function
        If .strPatientSex <> strPatientSex Then Exit Function
        If .strRegNo <> strRegNo Then Exit Function
'        If .strReportDoctor <> strReportDoctor Then Exit Function
'        If .strReportOperation <> strReportOperation Then Exit Function
'        If .strReportQuality <> strReportQuality Then Exit Function
        If .strStudyNum <> strStudyNum Then Exit Function
        If .strStudyUID <> strStudyUID Then Exit Function
        If .dtFirstTime <> dtFirstTime Then Exit Function
'        If .strStuStateDesc <> strStuStateDesc Then Exit Function
    End With
    
    IsEquals = True
End Function
