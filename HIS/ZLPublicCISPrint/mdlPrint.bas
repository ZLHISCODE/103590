Attribute VB_Name = "mdlPrint"
Option Explicit

Public gobjCommFun As Object
Public gobjDatabase As Object
Public gobjComlib As Object
Public gobjReport As Object
Public glngSys As Long
Public gcnOracle As New ADODB.Connection

'��ҳ��Ϣ;ҽ����¼;סԺ����;�����¼;������;���Ʊ���;����֤��;֪���ļ�;�ٴ�·��
'�ڲ�Ӧ��ģ��Ŷ���
Public Enum Enum_Inside_Program
    p���Ӳ������� = 2250
    p�°�סԺ���� = 2252
    p�°����ﲡ�� = 2251
    p����������д = 1249
    p���ﲡ������ = 1250
    pסԺ�������� = 1251
    p����ҽ���´� = 1252
    pסԺҽ���´� = 1253
    pסԺҽ������ = 1254
    p�����¼���� = 1255
    p�ٴ�·��Ӧ�� = 1256
    pҽ�����ѹ��� = 1257
    p���Ʊ������ = 1258
    p���Ӳ������� = 1259
    p����ҽ��վ = 1260
    pסԺҽ��վ = 1261
    pסԺ��ʿվ = 1262
    pҽ������վ = 1263
    P�°滤ʿվ = 1265
    p������ϲο� = 1270
    pҩƷ���Ʋο� = 1271
    p���˲������� = 1273
    p��Ƭ���߹��� = 1289
    p������� = 1132
    pסԺ���� = 1133
    p���ò�ѯ = 1139
    p���������� = 1113
    p�Ŷӽк�����ģ�� = 1160
    p������ҩ��� = 1266
    p������˹��� = 1267
    p���Ӳ������ = 1560
    p��Ѫ��˹��� = 1268
    p����ӿ� = 2425
    p������Ȩ���� = 1080
    p��Һ�������� = 1345
    P����·��Ӧ�� = 1248
    P�������Ĵ�ӡ = 1566
    P����ڲ��ӿ� = 2150
End Enum

Public Function GetPatiInfo(ByVal lngPatiID As Long, ByVal lngVisitID As Long) As ADODB.Recordset
'����:��ȡ������Ϣ
    Dim strSQL As String
        
    strSQL = "Select ��Ժ����id,����,סԺ��,����ת�� From ������ҳ Where ����id = [1] And ��ҳid = [2]"
    On Error GoTo errH
    Set GetPatiInfo = gobjDatabase.OpenSQLRecord(strSQL, "GetPatiInfo", lngPatiID, lngVisitID)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun���ڸú���
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

