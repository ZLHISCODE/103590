VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_Getein1100_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Ѫ������
'�����ͺţ�Getein1100_1
'�û�������������ҽԺ
'��ս�У�������
'���ڣ�2014-09-22
'
Private mstrCode As String    '��������

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


'----------------------�ı���ʽ�ӿں���---------------------------------------------------
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
                               ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '��׼���������ļ���ȡ������
    '   strFile������·�����ļ���
    '   lngSampleNO���걾�š�Ϊ-1��ʾѡȡ����ʱ�䷶Χ�ڵı걾
    '   dtStart����ʼʱ�䡣���ָ���˱걾�ţ���ֻ��ȡ�ò������������걾��dtEnd��Ч��
    '   dtEnd������ʱ�䡣ֻ�е�ѡȡ����걾��lngSampleNO=-1��ʱ���ò�����Ч�������ָ���ò���ֵ����ʱ�䷶ΧΪ>=dtStart��
    '   ���صļ���������(������������밴���±�׼��֯���)
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������

End Function


'----------------------------------------------------------------------------------------

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
    '   ��1��Ԫ�أ�������� ����� ^1 ��ʾ���� ��Ӽ� ^����
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
    Dim lng_S As Long, lng_E As Long
    Dim str_Sample As String, str_Day As String
    Dim str_No As String, str_Item As String, str_Value As String
    Dim INR_Value As String
    Dim str��� As String
    Dim strCode As String
    Dim J As Integer
    Dim lngID As String
    Dim strLine As String
    Dim Var_date As Variant
    Dim var_date1 As Variant

    '    mstrCode = ""
    mstrCode = mstrCode & strSource
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S + 1, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If

    Do While lng_S > 0 And lng_E > 0

        strCode = Mid(mstrCode, lng_S, lng_E - lng_S)
        mstrCode = Mid(mstrCode, lng_E)

        If UBound(Split(strCode, "|")) > 2 Then

            Var_date = Split(strCode, "|")
            str_Day = Var_date(13)
            str_Day = Format(str_Day, "yyyy/mm/dd hh:mm:ss")
            str_No = Trim(Var_date(3)) '��ȡ�걾��

            str_Item = Trim(Var_date(5))
            str_Value = Trim(Var_date(7))

            If InStr(str_Item, "^") > 0 Then

                var_date1 = Split(str_Item, "^")

                For J = LBound(var_date1) To UBound(var_date1)
                    '                    strLine = ""
                    str_Value = Trim(Var_date(7))
                    str_Value = Split(str_Value, "^")(J)
                    If str_Item = "CK-MB" Then
                        If Val(str_Value) < 2.5 Then
                            str_Value = "<2.5"
                        ElseIf Val(str_Value) > 80 Then
                            str_Value = ">80"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "cTnI" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 50 Then
                            str_Value = ">50"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "Myo" Then
                        If Val(str_Value) < 30 Then
                            str_Value = "<30"
                        ElseIf Val(str_Value) > 600 Then
                            str_Value = ">600"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "D-Dimer" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 30 Then
                            str_Value = ">30"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "NT-proBNP" Then
                        If Val(str_Value) < 100 Then
                            str_Value = "<100"
                        ElseIf Val(str_Value) > 35000 Then
                            str_Value = ">35000"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "hs-CRP" Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 200 Then
                            str_Value = ">200"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "PCT" Then
                        If Val(str_Value) < 0.1 Then
                            str_Value = "<0.1"
                        ElseIf Val(str_Value) > 50 Then
                            str_Value = ">50"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "mAlb" Then
                        If Val(str_Value) < 10 Then
                            str_Value = "<10"
                        ElseIf Val(str_Value) > 200 Then
                            str_Value = ">200"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "CysC" Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 10 Then
                            str_Value = ">10"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf InStr(str_Item, "MG") > 0 Then
                        If Val(str_Value) < 0.5 Then
                            str_Value = "<0.5"
                        ElseIf Val(str_Value) > 20 Then
                            str_Value = ">20"
                        Else
                            str_Value = str_Value
                        End If
                    ElseIf str_Item = "NGAL" Then
                        If Val(str_Value) < 50 Then
                            str_Value = "<50"
                        ElseIf Val(str_Value) > 1500 Then
                            str_Value = ">1500"
                        Else
                            str_Value = str_Value
                        End If
                    End If
                    strLine = strLine & "|" & Split(str_Item, "^")(J) & "|" & str_Value

                Next
            Else
                If str_Item = "CK-MB" Then
                    If Val(str_Value) < 2.5 Then
                        str_Value = "<2.5"
                    ElseIf Val(str_Value) > 80 Then
                        str_Value = ">80"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "cTnI" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 50 Then
                        str_Value = ">50"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "Myo" Then
                    If Val(str_Value) < 30 Then
                        str_Value = "<30"
                    ElseIf Val(str_Value) > 600 Then
                        str_Value = ">600"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "D-Dimer" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 30 Then
                        str_Value = ">30"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "NT-proBNP" Then
                    If Val(str_Value) < 100 Then
                        str_Value = "<100"
                    ElseIf Val(str_Value) > 35000 Then
                        str_Value = ">35000"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "hs-CRP" Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 200 Then
                        str_Value = ">200"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "PCT" Then
                    If Val(str_Value) < 0.1 Then
                        str_Value = "<0.1"
                    ElseIf Val(str_Value) > 50 Then
                        str_Value = ">50"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "mAlb" Then
                    If Val(str_Value) < 10 Then
                        str_Value = "<10"
                    ElseIf Val(str_Value) > 200 Then
                        str_Value = ">200"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "CysC" Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 10 Then
                        str_Value = ">10"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf InStr(str_Item, "MG") > 0 Then
                    If Val(str_Value) < 0.5 Then
                        str_Value = "<0.5"
                    ElseIf Val(str_Value) > 20 Then
                        str_Value = ">20"
                    Else
                        str_Value = str_Value
                    End If
                ElseIf str_Item = "NGAL" Then
                    If Val(str_Value) < 50 Then
                        str_Value = "<50"
                    ElseIf Val(str_Value) > 1500 Then
                        str_Value = ">1500"
                    Else
                        str_Value = str_Value
                    End If
                End If
                strLine = strLine & "|" & str_Item & "|" & str_Value

            End If
            str��� = str��� & "||" & str_Day & "|" & str_No & "|������|ѪҺ|0"
            str��� = str��� & strLine
            '                  str��� = ""
            strLine = ""
        End If

        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S + 1, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop

    strResult = Mid(str���, 3)
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

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    'strSampleInfo:������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ���
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
    
End Function


'---------------------------ø���ǽӿ��Զ���ָ��--------------------
Public Function CmdAnalyse(strCmdType As String) As String

End Function

Public Function StrValue(strN As String) As String

If InStr(strN, "(") > 0 Then
    StrValue = Split(strN, "(")(0)
Else
    StrValue = strN
End If


End Function