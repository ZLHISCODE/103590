VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisDev_TECAN_F50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mstrCode As String

'���̣���ʿ�ۿ�
'�ͺţ�ø����TECAN���ͺţ�F50
'�û�����������ҽԺ
'�ӿڷ�ʽ�����.asc�ļ�
Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function
Public Sub Analyse(ByVal strSource As String, strResult As String, strReserved As String, _
                   strCmd As String)
    On Error GoTo errHand '�򿪴�������
    '------------------------------------------------
    '�豸��׼�ӿ�
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����

    '����������H02��ʼH03����()

    Dim str_S As Long, str_E As Long, str_SS As Long
    Dim str_Code As String, strLine As String, varLine As Variant
    Dim VAR_DAT As Variant, aItem As Variant
    Dim i As Integer, intLoop As Integer
    Dim str_time As String
    Dim str���� As String, str�걾�� As String, str������ As String, strData As String, str��Ŀ As String, strͨ���� As String
    Dim str_Line As String

    str���� = ""
    str�걾�� = ""
    str��Ŀ = ""
    str������ = ""
    strData = ""

    '---���ݻ���
    mstrCode = mstrCode & strSource
    str_S = 0
    str_E = 0

    '������������������ݽ���Ӧ��

    strCmd = ""

    '�ж��ַ�����������
    str_S = InStr(mstrCode, Chr(&H2))

    If str_S > 0 Then
        str_E = InStr(str_S, mstrCode, Chr(&H3))
    Else
        str_E = 0
    End If

    If str_S > 0 And str_E > 0 And Len(mstrCode) > 10 Then

        Do While str_S > 0 And str_E > 0 '�ж�һ���������ж����ʼ�ͽ����Ŀ�ʼ

            str_Code = Mid(mstrCode, str_S, str_E - str_S)
            mstrCode = Mid(mstrCode, str_E + 1)
            'HOA�ֶ�
            VAR_DAT = Split(str_Code, Chr(&HA))
            '�ļ�û�����ݣ�����
            If InStr(str_Code, "-") = 0 Then
                Exit Do
            End If
            '�жϵ�λ����������ֹ�����걾�ŵ�ʱ���������
            If InStr(str_Code, vbTab) > 0 And Len(str_Code) - Len(Replace(str_Code, vbTab, "")) > 2 Then
                'ȡͨ����
                If InStr(str_Code, ".mth") > 0 Then
                    str��Ŀ = Replace(Right(Split(str_Code, ".m")(0), 5), vbCrLf, "")
                End If

                For i = LBound(VAR_DAT) To UBound(VAR_DAT) '���һ��Ϊ�գ������±�Խ��
                    strLine = VAR_DAT(i)
                    If Len(strLine) - Len(Replace(strLine, vbTab, "")) > 3 Then
                        str_Line = trimString(strLine)

                        str_time = Format(Now, "yyyy-mm-dd hh:mm:ss")
                        str���� = str_time

                        If Len(str_Line) - Len(Replace(str_Line, vbTab, "")) > 2 Then
                            str�걾�� = Val(Split(str_Line, vbTab)(2))
                        End If

                        If str�걾�� <> 0 Then
                            str������ = Replace(Replace(Replace(Replace(Trim(Split(str_Line, vbTab)(0)), "?", ""), "", "*"), "-", "����"), "+", "����") & "^0^0^" & Trim(Split(str_Line, vbTab)(3))
                            strData = IIf(Trim(strData) = "", "", strData & "||") & str���� & "|" & str�걾�� & "| |ѪҺ|0" & "|" & Trim(str��Ŀ) & "|" & Replace(Trim(str������), "", "")
                        End If
                    End If
                Next
                'ѭ�������ж�
                str_S = InStr(mstrCode, Chr(&H2))
                If str_S > 0 Then
                    str_E = InStr(str_E, mstrCode, Chr(&H3))
                Else
                    str_E = 0
                End If

            End If
        Loop
    End If
    '���������ؽ��
    strResult = strData
    strReserved = ""
    '------------------------------------------------
    Exit Sub
    '----------------
errHand:
    Dim lngErrFileNo As Long
    lngErrFileNo = FreeFile
    Open App.Path & "\LisErrmsg.txt" For Append As lngErrFileNo
    Print #lngErrFileNo, "ErrMsg:" & Format$(Now, "YYYY-MM-DD HH:MM:SS") & "," & Err.Number & "," & Err.Description & ",Analyse,clsLisDev_DSI905"
    Close #lngErrFileNo
    
End Sub

Private Function trimString(str As String) As String
        Dim strTemp       As String
        
        strTemp = Trim(str)
        Do While InStr(1, strTemp, Space(2)) > 0
                strTemp = Replace(strTemp, Space(2), Space(1))
        Loop
        trimString = strTemp
End Function