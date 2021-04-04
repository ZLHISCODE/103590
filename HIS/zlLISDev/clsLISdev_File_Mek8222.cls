VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISdev_File_Mek8222"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����:���Mek8222
'ҽԺ:��Ϫ�к�ʮ��ҽԺ
'����:����ǰ
'ʱ��:2010/08/16
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
    Dim textline
    Dim strFilePath As String
    Dim arrRow
    Dim str�걾�� As String, str���� As String, str�ʿ� As String
    Dim str��� As String
    Dim strNo As Long
    Dim lng_count As Long
    Dim strPicPath As String, str_imgName As String
    Dim strimg As String
    Dim arrCoo As String, j As Integer
    Dim objFileSys As New FileSystemObject
    Dim frmPic As New frmGraph
    On Error GoTo errH
     '��������±����
100 ReDim aResults(0)
102 aResults(0) = ""
104 lng_count = -1
106 ResultFromFile = aResults
    
108 strFilePath = strFile
110 strPicPath = strFile
112 Open strFile For Input As #1
114 Do While Not EOF(1)
116     Line Input #1, textline
118     arrRow = Split(textline, ",")
120     str��� = ""
122     strNo = lngSampleNO
124     If strNo > 0 Then
126         If strNo = Val(arrRow(9)) Then
128             j = 10
130             For j = 10 To UBound(arrRow)
132                 arrRow(j) = Replace(Replace(arrRow(j), "H", ""), "L", "")
134             Next j
136             str���� = "20" & Mid(arrRow(7), 9, 2) & "-" & TrunMonth(Mid(arrRow(7), 4, 3)) & "-" & Mid(arrRow(7), 1, 2) & " " & arrRow(8)
138             aResults(0) = str���� & "|" & strNo & "| |ȫѪ|0" & _
                     "|WBC|" & nvl_sp(arrRow(10)) & "|NE%|" & nvl_sp(arrRow(11)) & "|LY%|" & nvl_sp(arrRow(12)) & "|MO%|" & nvl_sp(arrRow(13)) & "|EO%|" & nvl_sp(arrRow(14)) & "|BA%|" & nvl_sp(arrRow(15)) & _
                     "|NE#|" & nvl_sp(arrRow(16)) & "|LY#|" & nvl_sp(arrRow(17)) & "|MO#|" & nvl_sp(arrRow(18)) & "|EO#|" & nvl_sp(arrRow(19)) & "|BA#|" & nvl_sp(arrRow(20)) & "|RBC|" & nvl_sp(arrRow(21)) & _
                     "|HGB|" & nvl_sp(arrRow(22)) & "|HCT|" & nvl_sp(arrRow(23)) & "|MCV|" & nvl_sp(arrRow(24)) & "|MCH|" & nvl_sp(arrRow(25)) & "|MCHC|" & nvl_sp(arrRow(26)) & "|RDW|" & nvl_sp(arrRow(27)) & _
                     "|PLT|" & nvl_sp(arrRow(28)) & "|PCT|" & nvl_sp(arrRow(29)) & "|MPV|" & nvl_sp(arrRow(30)) & "|PDW|" & nvl_sp(arrRow(31))
                
                Exit Do
            End If
        Else
140         str���� = "20" & Mid(arrRow(7), 9, 2) & "-" & TrunMonth(Mid(arrRow(7), 4, 3)) & "-" & Mid(arrRow(7), 1, 2)
142         If Format(str����, "yyyy-MM-dd") <= Format(dtEnd, "yyyy-MM-dd") And Format(str����, "yyyy-MM-dd") >= Format(DateAdd("d", -1, dtStart), "yyyy-MM-dd") Then
           
144             str���� = str���� & " " & arrRow(8)
            
146             str�걾�� = Val(arrRow(9))
148             j = 10
150             For j = 10 To UBound(arrRow)
152                 arrRow(j) = Replace(Replace(arrRow(j), "H", ""), "L", "")
154             Next j
156             str��� = "|WBC|" & nvl_sp(arrRow(10)) & "|NE%|" & nvl_sp(arrRow(11)) & "|LY%|" & nvl_sp(arrRow(12)) & "|MO%|" & nvl_sp(arrRow(13)) & "|EO%|" & nvl_sp(arrRow(14)) & "|BA%|" & nvl_sp(arrRow(15)) & _
                         "|NE#|" & nvl_sp(arrRow(16)) & "|LY#|" & nvl_sp(arrRow(17)) & "|MO#|" & nvl_sp(arrRow(18)) & "|EO#|" & nvl_sp(arrRow(19)) & "|BA#|" & nvl_sp(arrRow(20)) & "|RBC|" & nvl_sp(arrRow(21)) & _
                         "|HGB|" & nvl_sp(arrRow(22)) & "|HCT|" & nvl_sp(arrRow(23)) & "|MCV|" & nvl_sp(arrRow(24)) & "|MCH|" & nvl_sp(arrRow(25)) & "|MCHC|" & nvl_sp(arrRow(26)) & "|RDW|" & nvl_sp(arrRow(27)) & _
                         "|PLT|" & nvl_sp(arrRow(28)) & "|PCT|" & nvl_sp(arrRow(29)) & "|MPV|" & nvl_sp(arrRow(30)) & "|PDW|" & nvl_sp(arrRow(31))
                
158             str��� = Replace(Replace(str���, "H", ""), "L", "")
            
                'ͼ������;ͼ������;[ͼƬ������Ƿ�ɾ��];ȫ·��
                '����RBC;103;1;C:\TEMP\ABC.BMP
160             strimg = ""
162             strPicPath = ""
164             strPicPath = Mid(strFile, 1, InStr(strFile, "\mek_history_list.csv") - 1) & "\" & Mid(arrRow(0), 1, InStr(arrRow(0), ".mek") - 1)
166             If Dir(strPicPath & "_h1.bmp", vbDirectory) <> "" Then
                    'strimg = "RBC;113;1;" & strPicPath & "_h1.bmp"
168                 frmPic.Picture1.AutoSize = True
170                 frmPic.Picture1.Picture = LoadPicture(strPicPath & "_h1.bmp")
172                 Call PicInvertAndSave(frmPic.Picture1, strPicPath & "_h1.jpg", "JPG")
174                 strimg = "RBC;113;1;" & strPicPath & "_h1.jpg"
                End If
176             If Dir(strPicPath & "_h2.bmp", vbDirectory) <> "" Then
178                 frmPic.Picture1.AutoSize = True
180                 frmPic.Picture1.Picture = LoadPicture(strPicPath & "_h2.bmp")
182                 Call PicInvertAndSave(frmPic.Picture1, strPicPath & "_h2.jpg", "JPG")
                    
184                 If strimg <> "" Then strimg = strimg & "^"
186                 strimg = strimg & "PLT;113;1;" & strPicPath & "_h2.jpg"
                End If
188             If Dir(strPicPath & "_s1.bmp", vbDirectory) <> "" Then
190                 frmPic.Picture1.AutoSize = True
192                 frmPic.Picture1.Picture = LoadPicture(strPicPath & "_s1.bmp")
                      
194                 Call PicInvertAndSave(frmPic.Picture1, strPicPath & "_s1.jpg", "JPG")
                    
196                  If strimg <> "" Then strimg = strimg & "^"
198                  strimg = strimg & "DIFF1;113;1;" & strPicPath & "_s1.jpg"
                End If
200             If Dir(strPicPath & "_s2.bmp", vbDirectory) <> "" Then
202                 frmPic.Picture1.AutoSize = True
204                 frmPic.Picture1.Picture = LoadPicture(strPicPath & "_s2.bmp")
206                 Call PicInvertAndSave(frmPic.Picture1, strPicPath & "_s2.jpg", "JPG")
                    
208                 If strimg <> "" Then strimg = strimg & "^"
210                 strimg = strimg & "DIFF2;113;1;" & strPicPath & "_s2.jpg"
                End If
212             If Dir(strPicPath & "_s3.bmp", vbDirectory) <> "" Then
214                 frmPic.Picture1.AutoSize = True
216                 frmPic.Picture1.Picture = LoadPicture(strPicPath & "_s3.bmp")
218                 Call PicInvertAndSave(frmPic.Picture1, strPicPath & "_s3.jpg", "JPG")
                    
220                  If strimg <> "" Then strimg = strimg & "^"
222                 strimg = strimg & "DIFF3;113;1;" & strPicPath & "_s3.jpg"
                End If
            End If
            
224         If str��� <> "" Then
226             If strimg <> "" Then str��� = str��� & vbCrLf & strimg
228              ReDim Preserve aResults(lng_count + 1)
230              aResults(UBound(aResults)) = str���� & "|" & str�걾�� & "| |ȫѪ|0" & str���
            End If
232         lng_count = lng_count + 1
    
        End If
    Loop
234 Close #1
        
236 ResultFromFile = aResults
    Exit Function
errH:
238     WriteLog "Mek8222", "��" & CStr(Erl()) & "��", Err.Description
End Function




Private Function TrunMonth(ByVal strMonth As String) As String
    'ת��Ӣ�ĵ��·�Ϊ����
    Select Case Mid(UCase(strMonth), 1, 3)
        Case "JAN": TrunMonth = "01"
        Case "FEB": TrunMonth = "02"
        Case "MAR": TrunMonth = "03"
        Case "APR": TrunMonth = "04"
        Case "MAY": TrunMonth = "05"
        Case "JUN": TrunMonth = "06"
        Case "JUL": TrunMonth = "07"
        Case "AUG": TrunMonth = "08"
        Case "SEP": TrunMonth = "09"
        Case "OCT": TrunMonth = "10"
        Case "NOV": TrunMonth = "11"
        Case "DEC": TrunMonth = "12"
    End Select

End Function


Private Function nvl_sp(ByVal strIn As String) As String
    'Ϊ�յķ��ؿո�
    Dim strR As String
    strR = strIn
    If strR <> "" Then
    nvl_sp = strIn
    Else
    nvl_sp = " "
    End If
End Function


