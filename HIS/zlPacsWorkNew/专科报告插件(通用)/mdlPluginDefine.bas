Attribute VB_Name = "mdlPluginDefine"
Option Explicit


'�������ݷָ���
Public Const SPLITER_REPORT = "[[@]]"
Public Const SPLITER_ELEMENT = "[[;]]"


'���洰��
Public Const Report_Form_frmReportES  As String = "�ھ�������Ϣ"
Public Const Report_Form_frmReportPL As String = "������Һ��������Ϣ"
Public Const Report_Form_frmReportUS As String = "B�����������Ϣ"


Public Const ReportViewType_������� = "�������"
Public Const ReportViewType_������ = "������"
Public Const ReportViewType_���� = "����"
Public Const ReportViewType_������� = "�������"
Public Const ReportViewType_��첿λ = "��첿λ"



Public glngAdviceId As Long
Public glngReportId As Long
Public gblnMoved As Boolean
Public gobjParent As Object
Public gblnEditable As Boolean
Public gcnOracle As ADODB.Connection


Public gModified As Boolean     '��¼�Ƿ����޸�





Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
