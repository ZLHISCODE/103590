VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_UrisedE77"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'����:Urised E77 (�����������)
'����:δ֪
'����:�¶�
'�û�:

Private mstrCode As String

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
        
    '���ܣ���������
    '������
    'strSource��Ҫ������ԭʼ��
    'strResult�����صļ�����(����������������밴���±�׼��֯���)
    '   ÿ���������||�ָ�,Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������^�Ƿ���^�����  ����ֻ���������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   �ӵ�5��Ԫ�ؿ�ʼΪ��������ÿ2��Ԫ�ر�ʾһ��������Ŀ��
    '       �磺��5i��Ԫ��Ϊ������Ŀ����5i+1��Ԫ��Ϊ������
    'strReserved���������ɽ�����ԭʼ������Ҫ��������������
    'strCmd�������Ҫ���ɷ������豸���͵�����
   
    Dim lng_S As Long, lng_E As Long
    Dim str_Code As String 'һ��������¼
    Dim str_Line As String 'һ������
    Dim str_�걾�� As String, str_��Ŀ As String, str_��� As String, str_���� As String, str_Items As String
    Dim var_Code As Variant, i As Integer
    
    Dim objFile As New FileSystemObject
    
    Dim strImgPath As String
    Dim strFilename As String, strImg As String
    
    If strSource <> "" Then mstrCode = mstrCode & strSource
    
    lng_S = InStr(mstrCode, Chr(&H2))
    If lng_S > 0 Then
        lng_E = InStr(lng_S, mstrCode, Chr(&H3))
    Else
        lng_E = InStr(mstrCode, Chr(&H3))
    End If
    Do While lng_S > 0 And lng_E > 0
        str_Code = Mid(mstrCode, lng_S, lng_E - lng_S + 1)
        mstrCode = Mid(mstrCode, lng_E + 1)
        '-------------------------------------------------------------
        var_Code = Split(str_Code, Chr(&HD) & Chr(&HA))
        str_�걾�� = "": str_���� = "": str_��Ŀ = "": str_��� = "": str_Items = ""
        For i = LBound(var_Code) To UBound(var_Code)
            str_Line = var_Code(i)
            If str_Line Like "Sample ID :*" Then
                str_�걾�� = Mid(str_Line, 41)
            ElseIf str_Line Like "Date&Time :*" Then
                str_���� = Format(Now, "yyyy-MM-dd HH:mm")
            ElseIf InStr(",BIL,UBG,KET,ASC,GLU,PRO,BLD,PH ,NIT,LEU,SG.,", "," & Left(str_Line, 3) & ",") > 0 Then
                str_��Ŀ = Trim(Left(str_Line, 3))
                str_��� = Trim(Mid(str_Line, 61, 20))
                If str_��� <> "" And str_��Ŀ <> "" Then
                    str_Items = str_Items & "|" & str_��Ŀ & "|" & str_���
                End If
            ElseIf InStr(",RBC,.RBC,.RBCi,.RBCd,.RBC-G1,.RBC-oth,WBC,.WBC,.WBCc,CRY,.CRY,.CaOx,.CaOxm,.CaOxd,.TRI,.URI,.CaPh,.AMO,.CYS,.LEU,.TYR,.ATY,HYA,PAT,.PAT,.C-HGR,.C-GRA,.C-NEC,.C-RBC,.C-WBC,.C-CRY,.C-MIC,.C-FAT,.C-WAX,.C-MIX,NEC,.NEC,.s-TRA,.d-TRA,.REN,EPI,YEA,BAC,.BAC,.BACr,.BACc,MUC,SPRM,LIP,.LIP,.REN-L,.LDR,.CHOL,TRV,SCH,ART,", "," & Trim(Left(str_Line, 18)) & ",") > 0 Then
                str_��Ŀ = Trim(Left(str_Line, 18))
                str_��� = Replace(Trim(Mid(str_Line, 21, 20)), ",", ".")
                
                If str_��� <> "" And str_��Ŀ <> "" Then
                    str_Items = str_Items & "|" & str_��Ŀ & "|" & str_���
                End If
            End If
        Next
        If Val(str_�걾��) > 0 And str_���� <> "" And str_Items <> "" Then
            strResult = strResult & "||" & str_���� & "|" & str_�걾�� & "| |��Һ|0" & str_Items
            Call ReadConf(App.Path & "\UrisedE77.ini", "IMG_PATH", strImgPath)
            strImgPath = Trim(strImgPath)
            If strImgPath <> "" Then
                
                If objFile.FolderExists(Trim(strImgPath)) = True Then
                    strFilename = Dir(strImgPath & "\" & str_�걾�� & "_*_*_*_*_*_*.JPG")      ' ��Ѱ��һ�
                     
                    strImg = ""
                    
                    Do While strFilename <> ""   ' ��ʼѭ����
                        '9074_090421_02_59_11_-_01.JPG
                        strImg = strImg & "^" & Replace(UCase(Split(strFilename, "_")(6)), ".JPG", "") & ";110;1;" & strImgPath & "\" & strFilename
                        strFilename = Dir   ' ������һ���ļ���
                    Loop
                    If strImg <> "" Then
                        strImg = Mid(strImg, 2)
                        If strResult <> "" Then strResult = strResult & vbCrLf & strImg
                    End If
                End If
            End If
        End If
        '-------------------------------------------------------------
        lng_S = InStr(mstrCode, Chr(&H2))
        If lng_S > 0 Then
            lng_E = InStr(lng_S, mstrCode, Chr(&H3))
        Else
            lng_E = InStr(mstrCode, Chr(&H3))
        End If
    Loop
    
    If strResult <> "" Then strResult = Mid(strResult, 3)

End Sub

Public Function SendSample(ByVal strSampleInfo As String, iSendStep As Integer, blnSuccess As Boolean, Optional ByVal strResponse As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As String
    '������Ҫ���������͵ı걾��Ϣ
    '   Ԫ��֮����|�ָ�
    '   ��0��Ԫ�أ�����ʱ��
    '   ��1��Ԫ�أ��������
    '   ��2��Ԫ�أ�������
    '   ��3��Ԫ�أ��걾
    '   ��4��Ԫ�أ��Ƿ��ʿ�Ʒ
    '   ��5��Ԫ�أ�����
    '   ��6��9Ԫ�أ�ϵͳ����
    '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        SendSample = ""
End Function


Private Sub ReadConf(strIniFile As String, ByVal strItems As String, ByRef strReturn As String)
    '�������ļ�
    'strIniFile :ini�����ļ���
    'strItems   :Ҫ��ȡ���ֶΣ��磺DATABASE NAME=th_xlb�е� DATABASE NAME,����ֶ���,�ָ�
    'strReturn  :��ȡ�ķ���ֵ����:DATABASE NAME=th_xlb�е� th_xlb ���������ֵ��,�ָ�
    
    Dim arrLine
    Dim strLine As String

    Dim blnOpen As Boolean
    Dim objFilesys As New FileSystemObject
    Dim objStream As TextStream

    Dim intLoop As Long, intCount As Integer
    Dim arrItems As Variant
    
    On Error GoTo ErrHand
    
    '�������ļ��ж�ȡ�������ļ��е�����cotf
    intCount = UBound(Split(strItems, ","))
    ReDim arrItems(intCount, 2)
    For intLoop = 0 To intCount
        arrItems(intLoop, 0) = Split(strItems, ",")(intLoop)
    Next
    
    If objFilesys.FileExists(strIniFile) Then
        Set objStream = objFilesys.OpenTextFile(strIniFile)
        blnOpen = True
        Do While Not objStream.AtEndOfStream
            strLine = UCase(objStream.ReadLine)
            
            arrLine = Split(strLine, "=")
            If UBound(arrLine) = 1 Then
            
                For intLoop = 0 To intCount
                    If Trim(arrLine(0)) = arrItems(intLoop, 0) Then
                        arrItems(intLoop, 1) = arrLine(1)
                        Exit For
                    End If
                Next
           End If
        Loop
        objStream.Close
        strReturn = ""
        For intLoop = 0 To intCount
            strReturn = strReturn & "," & arrItems(intLoop, 1)
        Next
        If strReturn <> "" Then strReturn = Mid(strReturn, 2)
        blnOpen = False
    End If
    
    
    Exit Sub
ErrHand:
    If Error = 1 Then
        Resume
    End If
    If blnOpen Then objStream.Close
End Sub