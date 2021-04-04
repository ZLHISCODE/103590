Attribute VB_Name = "mdlTaxBill"
Option Explicit
Public gobjTax As New Beijing_tax.Tax
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrSql As String
Public gstrSysName As String
Public gstrUnitName As String

'Public Declare Function BJ_Normal_Invoice Lib "BeiJing_Tax.DLL" ( _
'    ByVal Invoice_Kind As Long, _
'    ByVal Invoice_NO As String, _
'    ByVal S_Consumer_Name As String, _
'    ByVal s_Oper_Name As String, _
'    ByVal InvoiceData As String, _
'    ByVal errMessage As String) As Long
'--------------------------------------------------------------------
'                   类型    名称        最大长度    备注
'Invoice_Kind       Integer 发票种类                1-医疗服务收费专用发票；2-医疗服务门诊收费专用发票
'Invoice_NO         PChar   发票号      18          发票号只能是数字，
'S_Consumer_Name    PChar   付款单位                Invoice_Kind= 1时最大长度为60；Invoice_Kind= 2时最大长度为76
's_Oper_Name        PChar   收费员      16
'InvoiceData        PChar   发票金额数据
'ErrMessage         PChar   操作返回提示错误信息
'--------------------------------------------------------------------

'Public Declare Function BJ_Other_Invoice Lib "BeiJing_Tax.DLL" ( _
'    ByVal Inv_Type As Long, _
'    ByVal Invoice_Kind As Long, _
'    ByVal Invoice_NO As String, _
'    ByVal s_Oper_Name As String, _
'    ByVal AdditionData As String, _
'    ByVal errMessage As String) As Long

'--------------------------------------------------------------------
'               名称        类型    最大长度    备注
'Inv_Type       开票类型    Integer             退票为1；废票为2；错票为3；定额票4
'Invoice_Kind   发票种类    Integer             1-医疗服务收费专用发票；2-医疗服务门诊收费专用发票
'Invoice_NO     发票号      PChar   18          定额票时可以为空,对应于开票软件中的"机打票号"项
's_Oper_Name    操作员名称  PChar   16
'AdditionData               PChar               退票时为空，废票时为空，错票时对应于开票软件中的"原始票号"项，定额票时对应于定额票的金额
'ErrMessage     操作返回提示错误信息    PChar
'--------------------------------------------------------------------

Public Function zStr(ByVal strText As String) As String
    If InStr(strText, Chr(0)) > 0 And strText <> "" Then
        zStr = Trim(Mid(strText, 1, InStr(strText, Chr(0)) - 1))
    Else
        zStr = Trim(strText)
    End If
End Function
