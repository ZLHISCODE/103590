VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_KY90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:�����
'ҽԺ:������Ժ�����
'����:����ǿ
'ʱ��:2015/1/17
'����:


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
    Dim cnsql As New ADODB.Connection
    Dim cnstring As String
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim lngResultCount As Long
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim blnNew As Boolean           '�Ƿ��±걾��
    Dim lngSampleNumber As Long     '�걾ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim lng_count As Long
    Dim st���� As String, en���� As String, str���� As String
    Dim strSer As String, strDB As String, strPWD As String
    Dim strValues As String, varValues As Variant
    Dim strItemValue As String, i As Integer
    Dim str�걾�� As String
    Dim str��� As String


    '��������±����
    ReDim aResults(0)
    aResults(0) = ""
    lng_count = -1
    ResultFromFile = aResults

    With cnsql
        Set cnsql = New ADODB.Connection
        If cnsql.State = adStateOpen Then cnsql.Close
        On Error Resume Next
        cnstring = "DSN=zllis;UID=root;PSW=1234"    'dsnϵͳodbc,uid �û���,psw����
        cnsql.Open cnstring
        st���� = Format(dtStart, "yyyy-mm-dd")
        en���� = Format(dtEnd, "yyyy-mm-dd")
        If lngSampleNO > 0 Then
            strSQL = "SELECT a.date,a.SampleCode,a.RedCell ,a.WhiteCell ,a.EpithelialCell,a.MucousSilk,a.Crystallize,a.UricAcidCrystal,a.Trichomonad,a.Fungus,a.Sperm,a.HyalineCast," & _
                   " a.RedCellCast,a.WhiteCellCast,a.WaxyCast,a.ThinGrainCast,a.ThickGrainCast,a.CircleEpi,a.CellBlock,a.unnormal_red_cell,a.unknow_cell," & _
                   " b.colour,b.sharpness,b.glucose,b.bilirubin,b.ketone,b.propotion,b.PHvalue,b.protein,b.URO,b.nitrite,b.occultblood,b.Whitecell as bWhitecell," & _
                   " b.vitaminC , b.calcium, b.creatinine, b.albumin " & _
                   " from nzjyjg a,drychemical b " & _
                   " where a.Date=b.Date and a.SampleCode=b.SampleCode and a.Date between '" & st���� & "' and '" & en���� & "' and a.SampleCode=" & lngSampleNO
        Else
            strSQL = "SELECT a.date,a.SampleCode,a.RedCell ,a.WhiteCell ,a.EpithelialCell,a.MucousSilk,a.Crystallize,a.UricAcidCrystal,a.Trichomonad,a.Fungus,a.Sperm,a.HyalineCast," & _
                   " a.RedCellCast,a.WhiteCellCast,a.WaxyCast,a.ThinGrainCast,a.ThickGrainCast,a.CircleEpi,a.CellBlock,a.unnormal_red_cell,a.unknow_cell," & _
                   " b.colour,b.sharpness,b.glucose,b.bilirubin,b.ketone,b.propotion,b.PHvalue,b.protein,b.URO,b.nitrite,b.occultblood,b.Whitecell as bWhitecell," & _
                   " b.vitaminC , b.calcium, b.creatinine, b.albumin " & _
                   " from nzjyjg a,drychemical b " & _
                   " where a.Date=b.Date and a.SampleCode=b.SampleCode and a.Date between '" & st���� & "' and '" & en���� & "'"
        End If

        Set rsTmp = cnsql.Execute(strSQL)
        Do Until rsTmp.EOF
            strItemValue = ""
            str�걾�� = Val(rsTmp!SampleCode)
            str���� = Format(rsTmp!date, "YYYY-MM-DD")
            intLoop = rsTmp.Fields.count - 1
            For i = 2 To intLoop
                strItemValue = strItemValue & "|" & i & "|" & Replace(Replace(Replace(Nvl(rsTmp.Fields(i), " "), vbCrLf, ""), vbCr, ""), "|", "")
            Next
            If strItemValue <> "" Then
                If aResults(UBound(aResults)) <> "" Then
                    ReDim Preserve aResults(UBound(aResults) + 1)
                End If
                str��� = str���� & "|" & str�걾�� & "| |��Һ|0" & strItemValue
                aResults(UBound(aResults)) = str���
            End If
            strItemValue = ""
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
    '    ResultFromFile = aResults
End Function

Public Sub AnalyConf(strItem As String, strCO As String)
    Dim arrLine
    Dim strLine As String
    Dim strFile As String
    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As textStream
    On Error GoTo ErrHand
    Dim intLoop As Long


    '�������ļ��ж�ȡ�������ļ��е�����cotf
    strFile = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "U2_Config.ini"
    If objFilesys.FileExists(strFile) Then
        Set objStream = objFilesys.OpenTextFile(strFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            If strLine = "" Then Exit Do
            arrLine = Split(strLine, "=")
            If arrLine(0) = strItem Then
                strCO = arrLine(1)
                Exit Sub
            End If
        Loop
        objStream.Close
        blnOpen = False
    End If

    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub

Public Function Fun_���(ByVal StrInput As String, ByVal strKey As String) As String
    Dim str��� As String
    Dim strVar As Variant
    Dim i As Long
    Fun_��� = ""
    If StrInput = "" Then Exit Function
    StrInput = Replace(Replace(Replace(StrInput, Chr(10), ""), Chr(13), ""), "<p", "p")
    strVar = Split(StrInput, ";")
    For i = LBound(strVar) To UBound(strVar)
        If Split(Replace(strVar(i), vbCrLf, ""), "=")(0) = strKey Then
            Fun_��� = Trim(Split(strVar(i), "=")(1))
            Exit Function
        End If
    Next i
End Function









