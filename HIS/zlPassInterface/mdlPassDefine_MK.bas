Attribute VB_Name = "mdlPassDefine_MK"
Option Explicit

'PASS接口函数，具体说明参见PASS接口文档描述
'--------------------------------------------------------------------------------------------------------------------------------------
'美康接口定义    version = 3.0
'--------------------------------------------------------------------------------------------------------------------------------------
'说明：ShellRunAs.dll需要安装在程序或系统目录
'      DIFPassDll.dll为Pass系统自动获取并注册路径
'注册服务器
Public Declare Function RegisterServer Lib "ShellRunAs.dll" () As Integer
'PASS初始化
Public Declare Function PassInit Lib "DIFPassDll.dll" ( _
                                 ByVal UserName As String, _
                                 ByVal DepartMentName As String, _
                                 ByVal WorkstationType As Integer) As Integer
'PASS运行模式设置
Public Declare Function PassSetControlParam Lib "DIFPassDll.dll" ( _
                                            ByVal SaveCheckResult As Integer, _
                                            ByVal AllowAllegen As Integer, _
                                            ByVal CheckMode As Integer, _
                                            ByVal DisqMode As Integer, _
                                            ByVal UseDiposeIdea As Integer) As Integer
'AllowAllegen 是否管理病人过敏史状态０－不管理；１－由用户传入；２－ＰＡＳＳ管理；３－ＰＡＳＳ强制管理

'传病人基本信息
Public Declare Function PassSetPatientInfo Lib "DIFPassDll.dll" ( _
                                           ByVal PatientID As String, _
                                           ByVal VisitID As String, _
                                           ByVal Name As String, _
                                           ByVal Sex As String, _
                                           ByVal Birthday As String, _
                                           ByVal Weight As String, _
                                           ByVal cHeight As String, _
                                           ByVal DepartMentName As String, _
                                           ByVal Doctor As String, _
                                           ByVal LeaveHospitalDate As String) As Integer
'传病人药品信息
Public Declare Function PassSetRecipeInfo Lib "DIFPassDll.dll" ( _
                                          ByVal OrderUniqueCode As String, _
                                          ByVal DrugCode As String, _
                                          ByVal DrugName As String, _
                                          ByVal SingleDose As String, _
                                          ByVal DoseUnit As String, _
                                          ByVal Frequency As String, _
                                          ByVal StartOrderDate As String, _
                                          ByVal StopOrderDate As String, _
                                          ByVal RouteName As String, _
                                          ByVal GroupTag As String, _
                                          ByVal OrderType As String, _
                                          ByVal OrderDoctor As String) As Integer

'传入病人过敏史
Public Declare Function PassSetAllergenInfo Lib "DIFPassDll.dll" _
                                            (ByVal AllergenIndex As String, _
                                             ByVal AllergenCode As String, _
                                             ByVal AllergenDesc As String, _
                                             ByVal AllergenType As String, _
                                             ByVal Reaction As String) As Integer
'参数:
'     AllergenIndex-过敏原在医嘱中的顺序编号，要求唯一
'     AllergenCode-过敏原编码，药品Id
'     AllergenDesc-过敏原名称
'     AllergenType-固定传人DrugName
'     Reaction-过敏症状，传人空串

'传入病生状态
Public Declare Function PassSetMedCond Lib "DIFPassDll.dll" _
                                       (ByVal MedCondIndex As String, _
                                        ByVal MedCondCode As String, _
                                        ByVal MedCondDesc As String, _
                                        ByVal MedCondType As String, _
                                        ByVal StartDate As String, _
                                        ByVal EndDate As String) As Integer
'参数:
'     MedCondIndex-诊断序号，唯一即可
'     MedCondCode-诊断编码
'     MedCondDesc-诊断名称
'     MedCondType-诊断类型(User)
'     StartDate-开始日期 当前时间， 精确到天，yyyy-mm-dd
'     EndDate-结束日期 当前时间，精确到天，yyyy-mm-dd


'设置需要进行单药警告的药品
Public Declare Function PassSetWarnDrug Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'信息查询药品传入
Public Declare Function PassSetQueryDrug Lib "DIFPassDll.dll" ( _
                                         ByVal DrugCode As String, _
                                         ByVal DrugName As String, _
                                         ByVal DoseUnit As String, _
    ByVal RouteName As String) As Integer
'获取右键菜单是否可用值
Public Declare Function PassGetState Lib "DIFPassDll.dll" (ByVal QueryItemNo As String) As Integer
'PASS功能调用
Public Declare Function PassDoCommand Lib "DIFPassDll.dll" (ByVal CommandNo As Integer) As Integer
'获取药品警示级别
Public Declare Function PassGetWarn Lib "DIFPassDll.dll" (ByVal DrugUniqueCode As String) As Integer
'设置药品浮动窗口位置
Public Declare Function PassSetFloatWinPos Lib "DIFPassDll.dll" ( _
    ByVal Left As Integer, ByVal Top As Integer, _
    ByVal Right As Integer, ByVal Bottom As Integer) As Integer
'PASS退出函数
Public Declare Function PassQuit Lib "DIFPassDll.dll" () As Integer

'----------------------------------------------------------------------------------------------------------------------------------
'--------------         美康接口声明   version 4.0 ------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------
'*******PASS4.0**1-美康嵌入代码开始（DLL函数声明）*****************************

'1、PASS初始化
Public Declare Function MDC_Init Lib "PASS4Invoke.dll" (ByVal pcCheckMode As String, ByVal pcHisCode As String, ByVal pcDoctorCode As String) As Integer
'传入参数:
'pcCheckMode: 字符串，审查模式，传入使用系统设置定义的模式，根据传入值的不同审查结果不同。维护工具调用DLL时传空，不显示工具条。
'pcHisCode: 字符串，医院编码，单医院模式传空字符串或者his提供的医院编码，区域模式传his提供的医院编码。
'pcDoctorCode: 字符串，医生编码，传入登录医生编码，必须是医生字典表里有的。用来登录互动平台。

'返回值：整型，1-成功
'0-失败
'-1-执行命令超时
'-2-连接PASS服务器失败
'-3-获取审查、查询列表出错
'-4-初始化工具条出错
'-5-更新资源文件出错
'调用: 系统必须首先调用MDC_Init成功后才能调用其他功能函数


'2、获取PASS系统最后一次错误信息函数
Public Declare Function MDC_GetLastError Lib "PASS4Invoke.dll" () As String
'传入参数:
'无
'返回值: 字符串 -错误信息

'3、审查类函数

'3-1 传入审查对象信息类函数
'3-1-1 传病人基本记录信息
Public Declare Function MDC_SetPatient Lib "PASS4Invoke.dll" ( _
                    ByVal pcPatCode As String, _
                    ByVal pcInHospNo As String, _
                    ByVal pcVisitCode As String, _
                    ByVal pcName As String, _
                    ByVal pcSex As String, _
                    ByVal pcBirthday As String, _
                    ByVal pcHeightCM As String, _
                    ByVal pcWeighKG As String, _
                    ByVal pcDeptCode As String, _
                    ByVal pcDeptName As String, _
                    ByVal pcDoctorCode As String, _
                    ByVal pcDoctorName As String, _
                    ByVal piPatStatus As Integer, _
                    ByVal piIsLactation As Long, _
                    ByVal piIsPregnancy As Long, _
                    ByVal pcPregStartDate As String, _
                    ByVal piHepDamageDegree As Long, _
                    ByVal piRenDamageDegree As Long) As Integer
'传入参数:
'pcPatCode：字符串类型，表示病人ID，与参数pcVisitCode唯一确定一个病人，此参数不能为空。
'pcInHospNo:符串类型，表示病人处方号或住院号，此参数不能为空。
'pcVisitCode：字符串类型，表示病人就诊次数或住院次数，与参数pcPatCode唯一确定一个病人，如果HIS系统没有此信息，则可传入"1"。
'pcName：字符串类型，表示病人姓名。
'pcSex：字符串类型，表示病人性别，格式为"男"、"女"、"不详"，如果没有赋值将会影响病人"妊娠"、"哺乳"、"性别"模块的审查。
'pcBirthday：字符串类型，表示病人出生日期，格式为"yyyy-mm-dd"，如果没有赋值将会影响病人"剂量"、"儿童警告"、"老年人警告"、"成人"、"肝、肾剂量"模块的审查。例如："1976-08-12"
'pcHeightCM：字符串类型，表示病人以厘米为单位的身高值，例如某病人身高为175厘米，则应传入"175"。如果HIS系统没有管理病人身高信息，则应传入空字符串。
'pcWeighKG：字符串类型，表示病人以公斤为单位的体重值，例如某病人体重为23.5公斤，则应传入"23.5"，由于传入身高时不能传入单位，所以如果HIS系统病人身高不是以公斤为单位，则要求必须换算成公斤后，再传入数值。如果HIS系统没有管理病人体重信息，则应传入空字符串。与剂量，肝、肾剂量损害相关.
'pcDeptCode: 字符串类型，表示科室编码。
'pcDeptName：字符串类型，表示科室名称。
'pcDoctorCode：字符串类型，表示主治/挂号医生编码。
'pcDoctorName：字符串类型，表示主治/挂号医生名称。
'piPatStatus：整型，表示病人状态：1表示住院病人（默认），2表示门诊病人，3表示急诊病人。
'piIsLactation：整型，表示病人哺乳状态，优先于通过PassSetMedCond（）函数传入"哺乳期"方式的审查，取值： -1-无法获取哺乳状态（默认）;0-不是;1-是
'piIsPregnancy：整型，表示病人妊娠状态，优先于通过PassSetMedCond（）函数传入"妊娠期"方式的审查，取值： -1-无法获取妊娠状态（默认）;0-不是;1-是
'pcPregStartDate：字符串类型，表示妊娠开始日期，格式为yyyy-mm-dd。
'piHepDamageDegree：整型，表示病人肝损害程度，优先于通过PassSetMedCond（）函数传入肝损害类诊断的审查，取值： -1-不确定（默认）；0-无肝损害；1-肝功能不全；2-轻度肝损害；3-中度肝损害；4-重度肝损害
'piRenDamageDegree整型，表示病人肾损害程度，优先于通过PassSetMedCond（）函数传入肾损害类诊断的审查，取值： -1-不确定（默认）；0-无肾损害；1-肾功能不全；2-轻度肾损害；3-中度肾损害；4-重度肾损害
'
'返回值：整型，1-成功
'0-失败
'调用：病人的基本信息发生变化之后，调用该接口。


'3-1-2 传病人药品记录信息
Public Declare Function MDC_AddScreenDrug Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, ByVal piOrderNo As Integer, _
                    ByVal pcDrugUniqueCode As String, ByVal pcDrugName As String, _
                    ByVal pcDosePerTime As String, ByVal pcDoseUnit As String, _
                    ByVal pcFrequency As String, ByVal pcRouteCode As String, _
                    ByVal pcRouteName As String, ByVal pcStartTime As String, _
                    ByVal pcEndTime As String, ByVal pcExecuteTime As String, _
                    ByVal pcGroupTag As String, ByVal pcIsTempDrug As String, _
                    ByVal pcOrderType As String, ByVal pcDeptCode As String, _
                    ByVal pcDeptName As String, ByVal pcDoctorCode As String, _
                    ByVal pcDoctorName As String, _
                    ByVal pcRecipNo As String, ByVal pcNum As String, _
                    ByVal pcNumUnit As String, ByVal pcPurpose As String, _
                    ByVal pcOprCode As String, ByVal pcMediTime As String, ByVal pcRemark As String) As Integer

'传入参数:
'pcIndex：字符串类型，表示医嘱唯一码，PASS系统将根据此参数来识别和区分传入的各条医嘱记录，审查后HIS系统只能通过此参数来获取PASS审查的结果值。在同一循环传入时，要求各记录的pcIndex值必须唯一，例如，可传入记录的行号值。
'piOrderNo：整型，表示医嘱编号,表示同一次审查传入药品的顺序号，用于确定审查问题属于哪个药。如果传入-1，则由系统根据调用接口顺序自动排序。
'pcDrugUniqueCode：字符串类型，表示药品唯一码，要求与PASS系统配对时采用的药品唯一码完全一致，否则PASS系统无法识别药品信息。此参数不能为空。
'pcDrugName：字符串类型，表示药品名称。
'pcDosePerTime：字符串类型，表示每次使用剂量的数字部分，传入此参数主要用于PASS对病人每次服用剂量的审查。注意：此处要求是转化为与药品配对剂量单位完全一致单位后的数值。例如药品配对剂量单位为"mg"，而病人的每次服用剂量为"0.5g"，此时就不能传入"0.5"，而应换算为"500mg"后，传入"500"。此参数如果为空，则不能审查剂量。
'pcDoseUnit：字符串类型，表示每次服用剂量单位，要求与药品配对剂量单位完全一致，否则可能造成剂量审查不正确。
'pcFrequency：字符串类型，表示药品服用频次信息。注意，要求与PASS系统配对时采用的频次编码完全一致。
'pcRouteCode：字符串类型，表示给药途径编码。注意，要求与PASS系统配对时采用的给药途径编码完全一致，由于PASS系统审查与给药途径关系密切，此参数传入错误，将直接导致审查错误；如果传空，则导致PASS系统无法审查与给药途径相关的审查项目。
'pcRouteName：字符串类型，表示给药途径名称。
'pcStartTime：字符串类型，表示开立医嘱日期。格式为"yyyy-mm-dd hh:mm:ss "，例如开嘱日期为1999年3月12日，则应传入"1999-03-12 00:00:00"。
'pcEndTime：字符串类型，传入参数，表示停嘱日期，格式为"yyyy-mm-dd hh:mm:ss "，例如停嘱日期为1999年3月12日，则应传入"1999-03-12 00:00:00"。临嘱停嘱日期等于开嘱日期，未停长期医嘱停嘱日期传空字符串。
'pcExecuteTime：字符串类型，表示执行医嘱时间。格式为"yyyy-mm-dd hh:mm:ss"。
'pcGroupTag：字符串类型，表示成组医嘱标记。主要用于PASS系统进行注射剂体外配伍审查识别注射剂是否配在一起使用，在循环传入的医嘱中，如果此参数值相同，则表示是配制在一起用，此种情况下才有可能存在体外配伍问题。
'pcIsTempDrug：字符串类型，表示医嘱是长期医嘱还是临时医嘱，'0'-表示长期医嘱； '1'-表示临时医嘱；
'pcOrderType：字符串类型，表示医嘱类别，取值'0'-在用（默认）；'1'-已作废；'2'-已停嘱；'3'-出院带药（根据系统设置参与审查），已作废医院不参与审查，并且会删除与此医嘱pcindex有关的所有审查结果，已停嘱不参与审查，但不影响停嘱前的审查结果。
'pcDeptCode：字符串类型，表示开嘱科室编码。
'pcDeptName：字符串类型，表示开嘱科室名称。
'pcDoctorCode：字符串类型，表示开嘱医生编码。
'pcDoctorName：字符串类型，表示开嘱医生名称。
'pcRecipNo：字符串类型，处方号，门诊处方专用，住院传空。此参数主要具有如下功能：
'(1)、用于"统计分析"显示处方号，便于查询和核对。
'（2）、用于门诊同一病人的多处方审查，处方号相同的才审药物与药物间的审查项目，与疾病相关审查规则为，适应症审查与处方号相关，禁忌症和不良反应与处方号无关。
'pcNum：字符串类型，药品开出数量，门诊处方审查专用，住院传空。为审7日用量预留。
'pcNumUnit：字符串类型，药品开出数量单位，门诊处方审查专用，住院传空。为审7日用量预留。
'pcPurpose：字符串类型，用药目的(0默认, 1可能预防，2可能治疗，3预防，4治疗，5预防+治疗)
'pcOprCode：字符串类型，手术编号，如果对应多手术，用'，'隔开，表示该药为该编号对应的手术用药
'pcMediTime：字符串类型，用药时机  0：非手术用药
'                                  1：术前0.5h以内用药
'                                  2：术前0.5-2h内
'3:                                   术前大于2h用药
'4:                                   术中用药
'5:                                   术后用药
'pcRemark：字符串类型，表示医嘱备注信息。
'返回值：整型，1-成功
'0-失败
'调用：：如果当前病人有多条用药信息记录时时，要求循环调用传入。

'传入病人过敏史记录信息
Public Declare Function MDC_AddAller Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, _
                    ByVal pcAllerCode As String, _
                    ByVal pcAllerName As String, _
                    ByVal pcAllerSymptom As String) As Integer

'传入参数:
'    pcIndex：字符串类型，表示过敏源序号，在同一循环传入时，要求各记录的pcIndex值必须唯一。
'pcAllerCode: 字符串类型，表示过敏源唯一码，要求与PASS系统配对时采用的过敏源唯一码完全一致，否则PASS系统无法识别此过敏信息。此参数不能为空。
'pcAllerName：字符串类型，表示过敏源名称。
'pcAllerSymptom：字符串类型，表示过敏源症状。
'返回值：整型，1-成功
'0-失败
'调用：如果当前病人有多条过敏信息记录时，要求循环调用传入。

'3-1-4 传入病人诊断记录信息
Public Declare Function MDC_AddMedCond Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, ByVal pcDiseaseCode As String, _
                    ByVal pcDiseaseName As String, ByVal pcRecipNo As String) As Integer
'传入参数:
'pcIndex：字符串类型，表示诊断序号，在同一循环传入时，要求各记录的pcIndex值必须唯一。
'pcDiseaseCode：字符串类型，表示诊断唯一码，要求与PASS系统配对时采用的诊断唯一码完全一致，否则PASS系统无法识别此诊断信息。此参数不能为空。
'pcDiseaseName：字符串类型，表示诊断名称。
'pcRecipNo：处方号。
'返回值：整型，2-成功，但是是重复传入的pcDiseaseCode
'1-成功
'0-失败
'调用：如果当前病人有多条诊断信息记录时时，要求循环调用传入。
                    
'3-1-5 传入病人手术记录信息
Public Declare Function MDC_AddOperation Lib "PASS4Invoke.dll" ( _
                    ByVal pcIndex As String, _
                    ByVal pcOprCode As String, _
                    ByVal pcOprName As String, _
                    ByVal pcIncisionType As String, _
                    ByVal pcOprStartDateTime As String, _
                    ByVal pcOprEndDateTime As String) As Integer
'传入参数:
'pcIndex：字符串类型，表示手术序号，在同一循环传入时，要求各记录的pcIndex值必须唯一。
'pcOprCode：字符串类型，表示手术唯一码，要求与PASS系统配对时采用的手术唯一码完全一致，否则PASS系统无法识别此手术信息。此参数不能为空。
'pcOprName：字符串类型，表示手术名称。
'pcIncisionType：字符串类型，表示手术切口类型。
'pcOprStartDateTime：字符串类型，表示手术开始时间，格式为"yyyy-mm-dd hh:mm:ss"。
'pcOprEndDateTime：字符串类型，表示手术结束时间，格式为"yyyy-mm-dd hh:mm:ss"。
'返回值：整型，1-成功
'0-失败
'调用：如果当前病人有多条手术信息记录时时，要求循环调用传入。
'pcIndex:HIS中医嘱ID;pcOprCode:疾病编码目录.编码（类别=’S’）pcOprName:疾病编码目录.名称 。pcOprStartDateTime:病人医嘱记录.手术时间，pcOprEndDateTime:HIS中取不到终止时间允许传空串。
'不传，审查不了手术用药 ，是否用抗菌药   ，抗菌药品种是否超出。但这个功能是在合理用药的增值包中。标准版是没有这个功能的。

'3-2审查函数

'3-2-1合理用药审查函数
Public Declare Function MDC_DoCheck Lib "PASS4Invoke.dll" (ByVal piShowMode As Integer, ByVal piIsSave As Integer) As Integer
'传入参数:
'piShowMode：整型，表示审查结果显示模式， 0-不显示界面 1-显示界面。
'piIsSave：整型，表示审查结果采集模式，0-不采集 1-采集。
'返回值：整型，1-成功
'0-失败

'3-3 获取审查结果函数
'3-3-1 获取药品医嘱警示级别
Public Declare Function MDC_GetWarningCode Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'传入参数:
'    pcIndex：字符串类型，表示医嘱唯一码，要求与调用MDC_AddScreenDrug（）函数传入的pcIndex 值完全一致。
'特别注意：pcIndex传空时，函数将返回去本次审查所有医嘱中最高的警示级别值。
'返回值：整型，具体含义如下：
'返回值小于0：表示可能出现异常，或者医嘱的一些错误信息：取值可能如下：
'-1-该药品在PASS中不存在或未配对。
'-2-该药品由于参数设置监测结果被过滤掉了。
'-3-医嘱已停嘱，不进行监测。
'-4-医嘱已作废，不进行监测。
'                -5-系统设置出院带药不进行监测。
'                -9-无开始和结束时间。
'        返回值大于或等于0：严重程度，取值可能为以下：
'                0-正常监测，无监测结果，蓝灯。
'1-正常监测，结果为禁忌或严重，黑灯。
'                2-正常监测，结果为不推荐，红灯。
'                3-正常监测，结果为慎用，橙灯。
'                4-正常监测，结果为关注，黄灯。



'获取一条药品医嘱的审查结果提示窗口函数
Public Declare Function MDC_ShowWarningHint Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'传入参数:
'pcIndex：字符串类型，表示医嘱唯一码，要求与调用MDC_AddScreenDrug（）函数传入的pcIndex 值完全一致。
'返回值：整型，1-成功

'关闭一条药品医嘱的审查结果提示窗口函数
Public Declare Function MDC_CloseWarningHint Lib "PASS4Invoke.dll" () As Integer
'传入参数:
'无
'返回值：整型，1-成功
'0-失败


'3-3-2获取审查结果条数函数
Public Declare Function MDC_GetResultItemCount Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As Integer
'传入参数:
'pcIndex：字符串类型，表示医嘱唯一码，要求与调用MDC_AddScreenDrug（）函数传入的pcIndex 值完全一致。
'返回值：整型，表示该医嘱所审查出的问题条数。

'3-3-3 获取审查结果详细信息函数
Public Declare Function MDC_GetResultDetail Lib "PASS4Invoke.dll" (ByVal pcIndex As String) As String
'传入参数:
'pcIndex：字符串类型，表示医嘱唯一码，要求与调用MDC_AddScreenDrug（）函数传入的pcIndex 值完全一致。（传空返回所有监测结果，传入哪条返回哪条的结果）
'返回值：字符串，以XML格式返回该医嘱所审查出的问题详细信息。
'

'4、信息查询类函数
'4-1 获取查询项目有效性函数
Public Declare Function MDC_GetDrugRefEnabled Lib "PASS4Invoke.dll" (ByVal pcDrugUniqueCode As String, ByVal piQueryType As Integer) As String
'传入参数:
'pcDrugUniqueCode：字符串类型，表示药品唯一码，要求与PASS系统配对时采用的药品唯一码完全一致，否则PASS系统无法识别药品信息。此参数不能为空。
'piQueryType:整型,表示查询模块。具体如下（特别注意：如果不传返回所有模块）：
'11-药品说明书
'21-药物专论
'31-病人用药教育
'41-中国药典
'51-简要信息(浮动窗口)
'61-相互作用
'62-药食作用
'63-体外配伍
'64-配伍浓度
'65-药物禁忌症
'66-药物适应症
'67-不良反应
'68-肝损害剂量
'69-肾损害剂量
'70-儿童用药
'71-妊娠用药
'72-哺乳用药
'73-老人用药
'74-成人用药
'75-性别用药
'76-细菌耐药率
'
'返回值:
'1、 如果传入piQueryType参数，则返回该项目查询信息是否可用的整型值，0-不可用 >0-可用。
'2、 如果没有传入piQueryType参数，返回按模块顺序组织好的字符串。

'4-2 查询药品信息函数
Public Declare Function MDC_GetDrugQueryInfo Lib "PASS4Invoke.dll" ( _
                    ByVal pcDrugUniqueCode As String, _
                    ByVal pcDrugName As String, _
                    ByVal piQueryType As Integer, _
                    ByVal X As Integer, _
                    ByVal Y As Integer) As Integer
'传入参数:
'pcDrugUniqueCode：字符串类型，表示药品唯一码，要求与PASS系统配对时采用的药品唯一码完全一致，否则PASS系统无法识别药品信息。此参数不能为空。
'pcDrugName：字符串类型，表示药品名称。
'piQueryType：整型,表示查询模块。具体如下（特别注意：如果不传返回所有模块）：
'                                        11-药品说明书
'                                        21-药物专论
'                                        31-病人用药教育
'                                        41-中国药典
'                                        51-简要信息(浮动窗口)
'                                        61-相互作用
'                                        62-药食作用
'                                        63-体外配伍
'                                        64-配伍浓度
'                                        65-药物禁忌症
'                                        66-药物适应症
'                                        67-不良反应
'                                        68-肝损害剂量
'                                        69-肾损害剂量
'                                        70-儿童用药
'                                        71-妊娠用药
'                                        72-哺乳用药
'                                        73-老人用药
'                                        74-成人用药
'                                        75-性别用药
'                                        76-细菌耐药率
'X：整型，表示X坐标。
'Y：整型，表示Y坐标。
'返回值：整型，1-成功
'0-失败

'传入一个药品信息函数
Public Declare Function MDC_DoSetDrug Lib "PASS4Invoke.dll" (ByVal pcDrugUniqueCode As String, _
                            ByVal pcDrugName As String) As Integer
'传入参数:
'pcDrugUniqueCode：字符串类型，表示药品唯一码，要求与PASS系统配对时采用的药品唯一码完全一致，否则PASS系统无法识别药品信息。此参数不能为空。
'pcDrugName：字符串类型，表示药品名称。
'返回值：整型，1-成功
'0-失败


'查询已传入药品说明书有效性函数
Public Declare Function MDC_DoRefDrugEnable Lib "PASS4Invoke.dll" (ByVal piQueryType As Integer) As String
'传入参数:
'piQueryType: 整型 , 11 - 药品说明书
'返回值：整型，>0表示有效。


'查询某一个药品信息函数

Public Declare Function MDC_DoRefDrug Lib "PASS4Invoke.dll" (ByVal piQueryType As Integer) As Integer
'传入参数:
'piQueryType:     整型 , 表示查询模块?具体如下:
'11-药品说明书
'                                        51-简要信息(浮动窗口)
'返回值：整型，1-成功
'0-失败

'4-3关闭浮动窗口函数
Public Declare Function MDC_CloseDrugHint Lib "PASS4Invoke.dll" () As Integer
'传入参数:无
'返回值：整型，1-成功
'0-失败

'6、调用药研究窗口函数
Public Declare Function MDC_DoMediStudy Lib "PASS4Invoke.dll" (ByVal pcUseTime As String) As Integer
'传入参数:
'pcUseTime：字符串，表示审查日期，嵌入医生工作站时要求传空，测试程序调用药研究传审查时间。
'返回值：整型，1-成功
'0-失败

'20 PASS退出函数
Public Declare Function MDC_Quit Lib "PASS4Invoke.dll" () As Integer

'附加信息传入
Public Declare Function MDC_AddJsonInfo Lib "PASS4Invoke.dll" (ByVal pcJson As String) As Integer
'传入参数:pcJson：字符串类型，JSON格式。druginfo为药品滴速补充信息，diseaseinfo为诊断补充信息
'/*传入格式
'{
'            "type":"jsontype",
'            "screentype":"1"
'        },
'        {
'            "type":"druginfo",
'            "index":"drug001",
'            "driprate":"60",
'            "driptime":"120"
'        },
'        {
'            "type":"diseaseinfo",
'            "index":"dis001",
'            "starttime":"2015-12-31 09:11:11",
'            "endtime":"2016-08-02 09:11:11"
'        },
'        {
'            "type":"otherrecipinfo",
'            "hiscode":"his001",
'            "index":" drug001",
'            "recipno":"2016-08-02 09:11:11",
'            "drugsource":"USER",
'            "druguniquecode":"123456",
'            "drugname":"阿莫西林胶囊",
'            "doseunit":"g",
'            "routesource":"USER",
'            "routecode":"1"",
'            "routename":"口服""
'        }
'*/
'返回值：整型，1-成功;0-失败
'调用：可以组织一小段JSON后调用，也可以组织完整JSON调用。



'*******PASS4.0**1-美康嵌入代码结束（DLL函数声明）*****************************


'*******PASS4.0**1-美康药师干预系统（DLL函数声明）*****************************
Public Declare Function MDC_GetTaskStatus Lib "PASS4Invoke.dll" ( _
    ByVal pcPatCode As String, ByVal pcInHospNo As String, _
    ByVal pcVisitCode As String, ByVal pcRecipNo As String, _
    ByVal piTaskType As Integer) As Integer
'传入参数:
'pcPatCode：字符串类型，表示病人ID，与参数pcVisitCode唯一确定一个病人，此参数不能为空。要求与MDC_SetPatient函数传入的pcVisitCode参数值完全一致。
'pcInHospNo:符串类型，表示病人门诊号或住院号，此参数不能为空。要求与MDC_SetPatient函数传入的pcInHospNo参数值完全一致。
'pcVisitCode：字符串类型，表示病人就诊次数或住院次数，与参数pcPatCode唯一确定一个病人，如果HIS系统没有此信息，则可传入"1"。要求与MDC_SetPatient函数传入的pcInHospNo参数值完全一致。
'pcRecipNo：字符串类型，门诊传处方号，住院传医嘱唯一码。门诊要求与MDC_AddScreenDrug函数传入的pcRecipNo参数值完全一致，
'           住院要求与MDC_AddScreenDrug函数传入的pcindex参数值完全一致。（注：该参数可以传空，表示取该任务的状态，不到具体处方或医嘱上）
'piTaskType：整型，表示病人类型：1表示住院病人（默认），2表示门诊病人。
'
'返回值：整型，表示药师干预结果：1-通过，0-不能通过
'调用：在门诊或住院医生工作站调用了PASS4的用药审查接口MDC_DoCheck后，有PASS审查结果时会弹出界面1，没有PASS审查结果时会弹出界面2-1（有超时设置时弹出界面2-2）
                                   
'*******PASS4.0**1-美康药师干预系统结束（DLL函数声明）*****************************
Public Function MK_GetPara() As Boolean
        Dim arrList As Variant
        Dim strPara As String
        
        On Error GoTo errH
100     strPara = zlDatabase.GetPara(90001, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
        '格式服务器IP&&服务器端口号
102     If strPara = "" Then strPara = "0" & G_STR_SPLIT & "" & G_STR_SPLIT & "0"
        '启用药师干预（1-开启;0-关闭）;医院编码:（默认为空按站点传入;不为空传入指定值）;是否显示备孕(1-是;0-否);是否启用静默式审查((1-是;0-否))
104     arrList = Split(strPara, G_STR_SPLIT)
106     If UBound(arrList) >= 2 Then
            gblnPharmReview = Val(arrList(0)) = 1
            gstrHOSCODE = arrList(1)     '医院编码
            gblnPrePregnancy = Val(arrList(2)) = 1
            If UBound(arrList) >= 3 Then gblnTEST = Val(arrList(3)) = 1
        Else
            gblnPharmReview = False
            gstrHOSCODE = ""     '医院编码
            gblnPrePregnancy = False
            gblnTEST = False
            Exit Function
        End If
        MK_GetPara = True
        Exit Function
errH:
146     MsgBox "读取参数失败！" & vbNewLine & "HZYY_GetPara:第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function MK_SetPara() As String
    MK_SetPara = IIf(gblnPharmReview, 1, 0) & G_STR_SPLIT & gstrHOSCODE & G_STR_SPLIT & IIf(gblnPrePregnancy, 1, 0) & G_STR_SPLIT & IIf(gblnTEST, 1, 0)
End Function
