Attribute VB_Name = "mdlPluginDefine"
Option Explicit


'报告内容分隔符
Public Const SPLITER_REPORT = "[[@]]"
Public Const SPLITER_ELEMENT = "[[;]]"


'报告窗体
Public Const Report_Form_frmReportES  As String = "内镜基本信息"
Public Const Report_Form_frmReportPL As String = "病理妇科液基薄层信息"
Public Const Report_Form_frmReportUS As String = "B超心脏测量信息"


Public Const ReportViewType_检查所见 = "检查所见"
Public Const ReportViewType_诊断意见 = "诊断意见"
Public Const ReportViewType_建议 = "建议"
Public Const ReportViewType_病理诊断 = "病理诊断"
Public Const ReportViewType_活检部位 = "活检部位"



Public glngAdviceId As Long
Public glngReportId As Long
Public gblnMoved As Boolean
Public gobjParent As Object
Public gblnEditable As Boolean
Public gcnOracle As ADODB.Connection


Public gModified As Boolean     '记录是否有修改





Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
