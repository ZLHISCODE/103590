Attribute VB_Name = "mdlTaxBill"
Option Explicit
Public gobjTax As New Beijing_tax.Tax
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
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
'                   ����    ����        ��󳤶�    ��ע
'Invoice_Kind       Integer ��Ʊ����                1-ҽ�Ʒ����շ�ר�÷�Ʊ��2-ҽ�Ʒ��������շ�ר�÷�Ʊ
'Invoice_NO         PChar   ��Ʊ��      18          ��Ʊ��ֻ�������֣�
'S_Consumer_Name    PChar   ���λ                Invoice_Kind= 1ʱ��󳤶�Ϊ60��Invoice_Kind= 2ʱ��󳤶�Ϊ76
's_Oper_Name        PChar   �շ�Ա      16
'InvoiceData        PChar   ��Ʊ�������
'ErrMessage         PChar   ����������ʾ������Ϣ
'--------------------------------------------------------------------

'Public Declare Function BJ_Other_Invoice Lib "BeiJing_Tax.DLL" ( _
'    ByVal Inv_Type As Long, _
'    ByVal Invoice_Kind As Long, _
'    ByVal Invoice_NO As String, _
'    ByVal s_Oper_Name As String, _
'    ByVal AdditionData As String, _
'    ByVal errMessage As String) As Long

'--------------------------------------------------------------------
'               ����        ����    ��󳤶�    ��ע
'Inv_Type       ��Ʊ����    Integer             ��ƱΪ1����ƱΪ2����ƱΪ3������Ʊ4
'Invoice_Kind   ��Ʊ����    Integer             1-ҽ�Ʒ����շ�ר�÷�Ʊ��2-ҽ�Ʒ��������շ�ר�÷�Ʊ
'Invoice_NO     ��Ʊ��      PChar   18          ����Ʊʱ����Ϊ��,��Ӧ�ڿ�Ʊ����е�"����Ʊ��"��
's_Oper_Name    ����Ա����  PChar   16
'AdditionData               PChar               ��ƱʱΪ�գ���ƱʱΪ�գ���Ʊʱ��Ӧ�ڿ�Ʊ����е�"ԭʼƱ��"�����Ʊʱ��Ӧ�ڶ���Ʊ�Ľ��
'ErrMessage     ����������ʾ������Ϣ    PChar
'--------------------------------------------------------------------

Public Function zStr(ByVal strText As String) As String
    If InStr(strText, Chr(0)) > 0 And strText <> "" Then
        zStr = Trim(Mid(strText, 1, InStr(strText, Chr(0)) - 1))
    Else
        zStr = Trim(strText)
    End If
End Function
