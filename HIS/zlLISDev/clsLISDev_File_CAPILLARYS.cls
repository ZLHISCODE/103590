VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_CAPILLARYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:CAPILLARYS ������
'����:Sebia
'����:�¶�
'�û�:��̨ɽҽԺ
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
    
        Dim aResults() As String
        Dim cnAccess As New ADODB.Connection
        Dim strFilePath As String, objFile As New Scripting.FileSystemObject
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim lngResultCount As Long, strBh As String
        Dim blnUnion As Boolean, strimg As String, iField As Integer, i As Integer
    
        Dim cnErs As New ADODB.Connection
        Dim strFileErs As String, rsErs As ADODB.Recordset
        Dim strImgErs As String, str���� As String
        Dim strChar As String
        Dim intMaxValue As Integer
        On Error GoTo hErr
100     strFilePath = objFile.GetParentFolderName(strFile)
102     If Len(strFilePath) = 0 Then strFilePath = App.Path
    
104     strChar = Replace(strFile, strFilePath & "\", "")
106     strChar = Mid(strChar, 4, 1)
    
108     strFileErs = strFile
    
110     If Trim(Dir(strFileErs)) <> "" Then
112         With cnErs
114         .ConnectionString = "DBQ=" & strFileErs & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
116         .Open
118         If .State = adStateClosed Then strFileErs = ""
            End With
        Else
120         strFileErs = ""
        End If
        
122     lngResultCount = -1
   
124     strImgErs = ""
126     WriteLog "��ʼ!", "", strFileErs
128     If strFileErs <> "" Then
130         strSQL = "SELECT val(seq) As ˳���,programma as ����, data_analisi as ����, " & _
                     "A.fraz_1,fraz_2,fraz_3,fraz_4,fraz_5,fraz_6,fraz_7,fraz_8,fraz_9,fraz_10,rapp1,Curva " & _
                     " FROM Anagrafica AS A Where Format(A.data_analisi,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-mm-dd") & "' And Format(A.data_analisi,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-mm-dd") & _
                     "' " & IIf(lngSampleNO = -1, "", " And val(seq)=" & lngSampleNO)
132         Set rsErs = cnErs.Execute(strSQL)
134         'WriteLog "ִ��Sql", rsErs.EOF, strSQL
136         Do Until rsErs.EOF

140             If Val("" & rsErs("˳���")) <> 0 Then
142                 lngResultCount = lngResultCount + 1
144                 ReDim Preserve aResults(lngResultCount)
                
146                 aResults(lngResultCount) = Format(rsErs("����"), "yyyy-MM-dd")  '����ʱ��
                
148                 If "" & rsErs!���� = "S" Then
150                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rsErs("˳���"))   '������� ѪҺ�걾+100
                    Else
152                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                            Val("" & rsErs("˳���")) + 200 '�������
                    End If
154                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          " "  '������
156                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & _
                          IIf("" & rsErs!���� = "S", "ѪҺ", "��Һ")  '�걾
                      
158                 aResults(lngResultCount) = aResults(lngResultCount) & "|0"
160                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "01|" & rsErs!fraz_1
162                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "02|" & rsErs!fraz_2
164                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "03|" & rsErs!fraz_3
166                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "04|" & rsErs!fraz_4
168                 aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "05|" & rsErs!fraz_5
                
170                 If "" & rsErs!���� = "S" Then
172                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "06|" & rsErs!rapp1
                    Else
174                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "06|" & rsErs!fraz_6
176                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "07|" & rsErs!fraz_7
178                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "08|" & rsErs!fraz_8
180                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "09|" & rsErs!fraz_9
182                     aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsErs!���� & "10|" & rsErs!fraz_10
                    End If
184                 strImgErs = "" & rsErs!Curva
            
186                 If strImgErs <> "" Then
                        Dim maxVal As Long
                    
188                     strImgErs = SebiaImg(Mid(strImgErs, 5, Len(strImgErs) - 8))
190                     maxVal = Val(Split(strImgErs, "|")(0))
192                     If maxVal <= 0 Then maxVal = 3800
194                     strImgErs = "��;0;" & maxVal + 200 & ",300,0-;" & Split(strImgErs, "|")(1)
196                     aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & strImgErs
198                     'WriteLog "��Ӿ", Val("" & rsErs("˳���")), strImgErs
                    End If
                End If
200             rsErs.MoveNext
            Loop
        End If
202     'WriteLog "����!", "", strFileErs
204     ResultFromFile = aResults
        Exit Function
hErr:
206     WriteLog "��Ӿ", CStr(Erl()), Err.Description

End Function

Private Function SebiaImg(ByVal strimg As String) As String
'   ����ͼ�����ݴ�
    Dim maxVal As Long
    Do While strimg <> ""
        If maxVal < Val("&H" & Left(strimg, 4)) And Val("&H" & Left(strimg, 4)) < 4000 Then
            maxVal = Val("&H" & Left(strimg, 4))
        End If
        If Val("&H" & Left(strimg, 4)) < 0 Then
            SebiaImg = "1;" & SebiaImg
        ElseIf Val("&H" & Left(strimg, 4)) < 4000 Then
            SebiaImg = CStr(Val("&H" & Left(strimg, 4))) & ";" & SebiaImg
        End If
        strimg = Mid(strimg, 5)
    Loop
    If Right(SebiaImg, 1) = ";" Then SebiaImg = maxVal & "|" & Mid(SebiaImg, 1, Len(SebiaImg) - 1)

End Function

