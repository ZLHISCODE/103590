VERSION 5.00
Begin VB.UserControl ctrlComm 
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   Picture         =   "ctrlComm.ctx":0000
   ScaleHeight     =   1560
   ScaleWidth      =   1710
   ToolboxBitmap   =   "ctrlComm.ctx":0842
   Begin VB.Timer timInData 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   15
      Top             =   525
   End
End
Attribute VB_Name = "ctrlComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strBuffer As String '���ݻ�����
'strSampleInfo�����͵ı걾��Ϣ
'iSendStep�����Ͳ��衣��1��ʼ������0����ִ�з���
Private strSampleInfo As String, iSendStep As Integer, dtSendTime As Date, mblnUndo As Boolean, miType As Integer

'Public Event DataReceived()
'Public Event DevOnComm(ByVal comPort As String, ByVal lngEvent As Long, ByVal strR As String)  ' ��ʾ��־�¼�
'Public Event DevSenComm(ByVal comPort As String, ByVal strR As String, ByVal intErr As Integer)
Public Event DevDecode(ByVal commport As String, ByVal str��� As String)
Public Event DevRefresh(ByVal lngID As Long)

Public Event ItemUnknown(ByVal commport As String, ByVal strItems As String) '����δ֪��
Public Event ReturnCompute(ByVal strReturn As String)  '�����Զ�������

Private mstrReceiveDir As String  'ͨѶ����Ŀ¼
Private mfsoTmp As New FileSystemObject  '�ļ�����

Private mintIndex As Integer     '����ؼ��������ţ����Դ�g����������ȡ��������Ϣ
Private mintMicrobe As Integer    '�Ƿ���΢���� 1= ΢����
'Private mItem() As Variant       '����ͨ����
Private mlng�������Ѻ��ձ걾 As Long   '˫��ͨѶ��ʹ�õ�һ������

Private mlngManID As Long        '������ID,���Ϊ״̬ʱ����

Private mlngDeviceID As Long     '����ID,����������Ϊ�������Ϊ������ID
Private mlngExeDeptID As Long    '����С��ID
Private mstrAutoCheckMan As String '�Զ������
Private mintAutoQCCalc  As Integer  '�Զ������ʿ� 0-������ 1-Ҫ����

Private int��촦��ʽ As Integer  '1-��ʾ��2-������3-������
Private int���ﴦ��ʽ As Integer  '1-��ʾ��2-������3-������
Private intסԺ����ʽ As Integer  '1-��ʾ��2-������3-������
Private intԺ�⴦��ʽ As Integer  '1-��ʾ��2-������3-������

Private mItem() As Variant

Public Property Get CommSetting() As String
    '�������ò���
    CommSetting = ""
End Property

Public Property Get DevProgName() As String
    '�������ò���
    DevProgName = ""
End Property

Public Function SendSample(ByVal lngDeviceID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, Optional strAdviceIDs As String = "", Optional ByVal blnUndo As Boolean = False, Optional ByVal iType As Integer = 0) As Boolean
    '�������������ͱ걾��Ϣ
    
    Dim strSendData As String
    On Error GoTo errH
   
    strSampleInfo = GetSampleInfo(lngDeviceID, mlngManID, strSampleDate, strSampleNO, "", strAdviceIDs, iType)
    If strSampleInfo <> "" Then
        strSampleInfo = strSampleInfo & ";" & IIf(blnUndo, 1, 0) & ";" & iType
        Call WriteToSendDir(strSampleInfo)
    End If
    SendSample = True
    
    If strSampleInfo <> "" Then
        gstrSQL = "ZL_����걾��¼_����(" & lngDeviceID & ",To_Date('" & strSampleDate & "','yyyy-MM-dd HH24:mi:ss'),'" & strSampleNO & "',1," & iType & ")"
        gobjDatabase.ExecuteProcedure gstrSQL, "���ͱ�־"
    End If
    
    Exit Function
errH:
    WriteLog "SendSample", LOG_������־, Err.Number, Err.Description
End Function

Public Property Get PortOpened() As Boolean
    PortOpened = False
End Property

Public Property Get DeviceID() As Long
    DeviceID = mlngDeviceID
End Property

Public Sub InitContrl(ByVal intIndex As Integer, Optional strCmd As String = "")
    '��ʼ���ؼ�
    '   intindex: ����
    '   strCmd  : ʵʼ����Ҫ����ͨѶ�����ָ�������ResetExe-����ͨѶ����CloseExe-�ر�ͨѶ����
        Dim tsmTmp As TextStream
        Dim lngSaveAsID As Long, rsTmp As adodb.Recordset, strSQL As String
        Dim strVer As String, strDevVer As String
        On Error GoTo errH
    
100     ReDim mItem(1, 2) As Variant
102     mItem(1, 0) = -1
104     mItem(1, 1) = 0
106     mItem(1, 2) = 2
    
108     timInData.Enabled = False
    
110     If g����(intIndex).ID > 0 Then
            '�������Ŀ¼�Ƿ���ڣ��������򴴽�������ͨѶ�����ȥ
112         mintIndex = intIndex
114         mstrReceiveDir = g����(intIndex).ͨѶĿ¼
116         mstrAutoCheckMan = Trim(g����(intIndex).�Զ������)
118         If mstrAutoCheckMan <> "" Then
120             int��촦��ʽ = Val(gobjDatabase.GetPara("��첡����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
122             intԺ�⴦��ʽ = Val(gobjDatabase.GetPara("Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
124             intסԺ����ʽ = Val(gobjDatabase.GetPara("סԺ������Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
126             int���ﴦ��ʽ = Val(gobjDatabase.GetPara("���ﲡ����Ϣ��һ�µĴ���ʽ", glngSys, 1208, True, 1))
            End If
128         mintAutoQCCalc = Val(g����(intIndex).�Զ������ʿ�)
130         If Not mstrReceiveDir Like "?:\*" Then mstrReceiveDir = App.Path & "\Dev_" & g����(intIndex).ID
132         g����(intIndex).ͨѶĿ¼ = mstrReceiveDir
134         If Not mfsoTmp.FolderExists(mstrReceiveDir) Then
136             Call mfsoTmp.CreateFolder(mstrReceiveDir)
            End If
        
138         If Dir(mstrReceiveDir & "\zlLisReceiveSend.exe") = "" Then mfsoTmp.CopyFile App.Path & "\zlLisReceiveSend.exe", mstrReceiveDir & "\"
        
140         If Dir(mstrReceiveDir & "\ReceiveSend.ini") = "" Then
142             Set tsmTmp = mfsoTmp.CreateTextFile(mstrReceiveDir & "\ReceiveSend.ini")
144             tsmTmp.WriteLine "[RECEIVE_SET]"
146             tsmTmp.WriteLine "���� = " & g����(intIndex).����
            
148             tsmTmp.WriteLine "COM�˿� = " & g����(intIndex).COM��
150             tsmTmp.WriteLine "������ = " & g����(intIndex).������
152             tsmTmp.WriteLine "����λ = " & g����(intIndex).����λ
154             tsmTmp.WriteLine "ֹͣλ = " & g����(intIndex).ֹͣλ
156             tsmTmp.WriteLine "У��λ = " & g����(intIndex).У��λ
158             tsmTmp.WriteLine "���� = " & g����(intIndex).����
160             tsmTmp.WriteLine "�����С = 2048"
            
162             tsmTmp.WriteLine "IP = " & g����(intIndex).IP
164             tsmTmp.WriteLine "IP�˿� = " & g����(intIndex).IP�˿�
166             tsmTmp.WriteLine "���� = " & g����(intIndex).����
            
168             tsmTmp.WriteLine "�Զ�Ӧ�� = " & g����(intIndex).�Զ�Ӧ��
170             tsmTmp.WriteLine "�ַ�ģʽ = " & g����(intIndex).�ַ�ģʽ
172             tsmTmp.WriteLine "ͨѶ���� = " & g����(intIndex).ͨѶ����
174             tsmTmp.WriteLine "ͨѶ���� = 0.5"
176             tsmTmp.Close
178             Set tsmTmp = Nothing
            End If
            '����Ƿ�������û����������
180         If Dir(mstrReceiveDir & "\Lock.txt") = "" Then
                '�Զ�����������������
182             strVer = mfsoTmp.GetFileVersion(App.Path & "\zlLisReceiveSend.exe")
184             strDevVer = mfsoTmp.GetFileVersion(mstrReceiveDir & "\zlLisReceiveSend.exe")
186             If strVer > strDevVer And strVer <> "" And strDevVer <> "" Then
188                 mfsoTmp.CopyFile App.Path & "\zlLisReceiveSend.exe", mstrReceiveDir & "\"
                End If
            
190             If strCmd = "" Then Call Shell(mstrReceiveDir & "\zlLisReceiveSend.exe", vbNormalNoFocus)

            Else
                '����������д��Ҫreceiveִ�е�����,�������ӿڣ��رյ�.
192             If strCmd <> "" Then
194                 Set tsmTmp = mfsoTmp.CreateTextFile(mstrReceiveDir & "\Send\" & strCmd & ".txt")
196                 tsmTmp.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss")
198                 tsmTmp.Close
200                 Set tsmTmp = Nothing
                End If
            End If
        
202         timInData.Enabled = False
204         If mfsoTmp.FolderExists(mstrReceiveDir & "\Result") Then
            
206             mlng�������Ѻ��ձ걾 = g����(intIndex).�ɷ��Ѻ˱걾
208             mlngManID = g����(intIndex).ID
210             mlngDeviceID = mlngManID

            
                '��ʼ��ͨѶ����,ʼ�մ�������ȡ
212             strSQL = "Select ͨѶ������,nvl(΢����,0) as ΢����,ʹ��С��ID From �������� Where ID=[1]"
214             Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "ȡ�Ƿ�΢����", mlngManID)
            
216             If Not rsTmp.EOF Then
218                 mintMicrobe = Nvl(rsTmp(1), 0)
220                 mlngExeDeptID = Nvl(rsTmp(2), 0)
222                 glngExeDeptID = mlngExeDeptID
                End If


                '-----
                '������Ϊģʽ���ã���Ϊ �������� ��ID
224             If g����(intIndex).SaveAsID > 0 Then mlngDeviceID = g����(intIndex).SaveAsID
            
                '���ݲ���������������������ͨ�����Ǵ��ĸ��ط�ȡ
226             If g����(intIndex).���Ϊͨ���� = 0 Then mlngManID = mlngDeviceID
                '------
            
228             If mintMicrobe = 1 Then
230                 strSQL = "Select ͨ������,������ID As ��ĿID, 2 as С��λ��,b.����||nvl(b.����,b.������) as ���� From ����ϸ������ A, �����ÿ����� B Where a.������id = b.Id And a.����id = [1] "
                Else
232                 strSQL = "Select a.ͨ������, a.��Ŀid, Nvl(a.С��λ��, 2) As С��λ��, b.���� || '-' || Nvl(b.Ӣ����, b.������) As ����," & vbNewLine & _
                            "       LPad(Decode(c.�������, Null, b.����, c.�������), 10, '0') As ����" & vbNewLine & _
                            "From ������Ŀ C, ����������Ŀ B, ����������Ŀ A" & vbNewLine & _
                            "Where a.��Ŀid = b.Id And a.��Ŀid = c.������Ŀid And a.����id = [1] " & vbNewLine & _
                            "Order By LPad(Decode(c.�������, Null, b.����, c.�������), 10, '0')"

                
                    '2011-12-07 ���������޸ģ�3/5 - ָ������
                End If
            
234             Set rsTmp = gobjDatabase.OpenSqlRecord(strSQL, "ȡͨ����", mlngManID)
236             If Not rsTmp.EOF Then
238                 mItem = rsTmp.GetRows
                End If

            
                ' ������ʱ������ʼ��ⷵ������
                timInData.Interval = 500
240             timInData.Enabled = True
            End If
        End If
        Exit Sub
errH:
242     WriteLog "initcontrl", LOG_������־, Err.Number, CStr(Erl()) & "��," & Err.Description
End Sub

Private Sub timInData_Timer()
    'ÿ���봥���������ݵĹ���
    Dim strData As String
    Dim ii As Long
    On Error GoTo errH
    
    '���붨ʱ����ʱ���ȹرն�ʱ��
    timInData.Enabled = False
    
    '���ȴ��� ���������ļ�
    Call ReadResultDirFileIQ(True)

    '�������ļ�
    Call ReadResultDirFileRE(True)
        
    '��������������֮���ٿ�����ʱ��
    timInData.Enabled = True
        
    Exit Sub
errH:
    'Resume
    If Err.Number <> 9 Then
        WriteLog "timInData", LOG_������־, Err.Number, Err.Description
    End If
End Sub

Private Function ReadResultDirFileIQ(ByVal blnDelete As Boolean) As String
    '�������ļ�
    Dim objFolder As Folder
    Dim objStream As TextStream
    Dim objFiles As Files
    Dim objOneFile As File
    Dim strFolder As String
    Dim strFileName As String
    Dim i As Long
    
 
    Dim strLine As String, lngCount As Long
    
    On Error GoTo errH
    
    strFolder = mstrReceiveDir & "\Result"
    Set objFolder = mfsoTmp.GetFolder(strFolder)
    Set objFiles = objFolder.Files
    For Each objOneFile In objFiles
        'Ϊ�˷�ֹ�ļ����ı�������δд�����ݾͱ���ȡ,�ļ��Ĵ���ʱ��С�ڵ�ǰʱ��10����ļ��Żᱻ��ȡ
        If objOneFile.Name Like "IQ" & Format(Now, "yyyyMMdd") & "_*.txt" And objOneFile.DateCreated < Format(Now - 0.0001, "yyyy-mm-dd hh:mm:ss") Then
            '����ҵ���ƥ����ļ��Ͷ�ȡ���ļ�
            strFileName = objOneFile.Path
            If mfsoTmp.FileExists(strFileName) Then
                Set objStream = mfsoTmp.OpenTextFile(strFileName, ForReading)
                strLine = ""
                Do
                    If objStream.AtEndOfStream Then Exit Do
                    strLine = strLine & objStream.ReadLine
                Loop
                objStream.Close
                Set objStream = Nothing
                If mfsoTmp.FileExists(strFileName) And blnDelete = True Then mfsoTmp.DeleteFile (strFileName)
'                    ReadResultDirFile = strLine
                '����ȡ���ļ����ݱ��浽���ݿ�
                If strLine <> "" Then
                    WriteLog "timInData-��������", LOG_ͨѶ��־, 0, strLine
                    Call WriteSampleInfo(strLine)   '= True Then RaiseEvent DataReceived
                End If
            End If
        End If
        DoEvents
    Next
    Set objFolder = Nothing
    Set objFiles = Nothing
    
    Exit Function
errH:
    WriteLog "ReadResultDirFile", LOG_������־, Err.Number, Err.Description
End Function

Private Function ReadResultDirFileRE(ByVal blnDelete As Boolean) As String
    '�������ļ�
    Dim objFolder As Folder
    Dim objStream As TextStream
    Dim objFiles As Files
    Dim objOneFile As File
    Dim strFolder As String
    Dim strFileName As String
    Dim ii As Long
    Dim i As Long
    
 
    Dim strLine As String, lngCount As Long
    
    On Error GoTo errH
    
    strFolder = mstrReceiveDir & "\Result"
    Set objFolder = mfsoTmp.GetFolder(strFolder)
    Set objFiles = objFolder.Files
    For Each objOneFile In objFiles
        'Ϊ�˷�ֹ�ļ����ı�������δд�����ݾͱ���ȡ,�ļ��Ĵ���ʱ��С�ڵ�ǰʱ��10����ļ��Żᱻ��ȡ
        If objOneFile.Name Like "RE" & Format(Now, "yyyyMMdd") & "_*.txt" And objOneFile.DateCreated < Format(Now - 0.0001, "yyyy-mm-dd hh:mm:ss") Then
            strFileName = objOneFile.Path
            If mfsoTmp.FileExists(strFileName) Then
                Set objStream = mfsoTmp.OpenTextFile(strFileName, ForReading)
                strLine = ""
                Do
                    If objStream.AtEndOfStream Then Exit Do
                    strLine = strLine & objStream.ReadLine
                Loop
                objStream.Close
                Set objStream = Nothing
                If mfsoTmp.FileExists(strFileName) And blnDelete = True Then mfsoTmp.DeleteFile (strFileName)
                If strLine <> "" Then
                    ii = UBound(mItem, 2)
                    strLine = Replace(strLine, "CHR(10) CHR(13)", vbCrLf)
                    WriteLog "TimInData-��������", LOG_ͨѶ��־, 0, strLine
                    Call InDataBase(strLine)  '= True ' Then RaiseEvent DataReceived
                End If
            End If
        End If
        DoEvents
    Next
    Set objFolder = Nothing
    Set objFiles = Nothing
        
    Exit Function
errH:
    WriteLog "ReadResultDirFile", LOG_������־, Err.Number, Err.Description
End Function

Private Sub UserControl_Initialize()
    strBuffer = ""
    iSendStep = 0 '��ʼ��ִ�з���
End Sub

Private Sub Return_Decode(ByVal strDecode As String)
    '���ؽ�����
    If strDecode = "" Then Exit Sub
    If g����(mintIndex).���� = 0 Then
        RaiseEvent DevDecode(g����(mintIndex).COM��, strDecode)
    Else
        RaiseEvent DevDecode(g����(mintIndex).IP, strDecode)
    End If
End Sub

Private Function WriteSampleInfo(ByVal strResult As String) As Boolean
    '˫��ͨѶ�У�ȡ�ñ걾��Ϣ,Ȼ��д��ͨѶĿ¼
    
    Dim aSamples() As String, aSampleInfo() As String, i As Integer, strSampleInfo As String, iType As Integer
    Dim strSampleNO As String, aTmp() As String, strBarcode As String
    
    On Error GoTo errH
    
    If Len(strResult) > 0 Then 'Ҫ���������ͱ걾��Ϣ
        aSamples = Split(strResult, "||")
        strSampleInfo = "": miType = 0
        For i = 0 To UBound(aSamples)
            aSampleInfo = Split(aSamples(i), "|")
            If UBound(aSampleInfo) > 0 Then
                aTmp = Split(aSampleInfo(1), "^")
                If UBound(aTmp) = 0 Then
                    strSampleNO = Val(aTmp(0)): miType = 0: strBarcode = ""
                Else
                    strSampleNO = Val(aTmp(0)): miType = Val(aTmp(1)): strBarcode = ""
                    If UBound(aTmp) > 1 Then
                        strBarcode = Trim(aTmp(2))
                    End If
                End If
                
                'д���ݵ� ͨѶ����Ŀ¼
                strSampleInfo = GetSampleInfo(mlngDeviceID, mlngManID, Format(aSampleInfo(0), "yyyy-MM-dd"), strSampleNO, strBarcode, , miType)
                If strSampleInfo <> "" Then
                    strSampleInfo = strSampleInfo & ";0;" & miType
                    Call WriteToSendDir(strSampleInfo)
                    If UBound(Split(strSampleInfo, "|")) > 2 Then strSampleNO = Split(strSampleInfo, "|")(1)
                    gstrSQL = "ZL_����걾��¼_����(" & mlngDeviceID & ",To_Date('" & Format(aSampleInfo(0), "yyyy-MM-dd") & "','yyyy-MM-dd'),'" & strSampleNO & "',1," & miType & ")"
                    WriteLog "д���ͱ�־", LOG_ͨѶ��־, 0, gstrSQL
                    gobjDatabase.ExecuteProcedure gstrSQL, "���ͱ�־"
                End If
                
            End If
        Next

    End If
    Exit Function
errH:
    WriteLog "WriteSampleInfo", LOG_������־, Err.Number, Err.Description
End Function

Private Sub WriteToSendDir(ByVal strInput As String, Optional ByVal strFileType As String)
    'strFileType: д��Ҫ����Ŀ¼���ļ����ͣ�SendSampleΪҪ���͸�������ָ�
    
    
    Dim strFileName As String
    Dim lngCount As Long, lngFileNo As Long
    On Error GoTo errH
    
    If mfsoTmp.FolderExists(mstrReceiveDir & "\Send") = False Then mfsoTmp.CreateFolder (mstrReceiveDir & "\Send")
    lngCount = 0
    If strFileType = "" Then strFileType = "SendSample"
    strFileName = Dir(mstrReceiveDir & "\Send\" & strFileType & "_*.txt")
    If strFileName <> "" Then lngCount = Val(Split(strFileName, "_")(1))
    Do While lngCount < 1000
        lngCount = lngCount + 1
        strFileName = mstrReceiveDir & "\Send\" & strFileType & "_" & Format(lngCount, "000") & ".txt"
        If mfsoTmp.FileExists(strFileName) = False Then
            lngFileNo = FreeFile
            Open strFileName For Binary Access Read Write Lock Read Write As lngFileNo
            Put lngFileNo, , strInput
            Close lngFileNo
            Exit Do
        End If
    Loop
    Exit Sub
errH:
    WriteLog "WriteToSendDir", LOG_������־, Err.Number, Err.Description
End Sub

Private Function GetSampleInfo(ByVal lngDeviceID As Long, ByVal lngMainID As Long, ByVal strSampleDate As String, ByVal strSampleNO As String, ByVal strBarcode As String, Optional strAdviceIDs As String = "", Optional ByVal iType As Integer = 0) As String
        '��ȡ��Ҫ���������͵ı걾��Ϣ
        '���أ��걾��Ϣ��
        '   �걾֮����||�ָ�
        '   Ԫ��֮����|�ָ�
        '   ��0��Ԫ�أ�����ʱ��
        '   ��1��Ԫ�أ��������
        '   ��2��Ԫ�أ���������
        '   ��3��Ԫ�أ��걾����
        '   ��4��Ԫ�أ������־
        '   ��5��Ԫ�أ���������
        '   ��6��Ԫ�أ��̺ţ�����
        '   ��7��Ԫ�أ�����ID^�Ա�^��������^����^����ȫƴ^ϡ�ͱ���(Ԥ�����ݴ���)     2013��11��07�� modify by �¶�,
        '   ��8��9Ԫ�أ�ϵͳ����
        '   �ӵ�10��Ԫ�ؿ�ʼΪ��Ҫ�ļ�����Ŀ��
        '  lngDeviceID = ����ID
        '  strSampleDate = ���� ��ʽΪ YYYY-MM-DD
        '  strSampleNO = �걾��
        '  strBarcode = ����
        '  strAdviceIDs =???
        '  iType = �걾���
        Dim objDevice As Object
        Dim rsTmp As New adodb.Recordset
        Dim lngAdviceID As Long, aAdviceIDs() As String, i As Integer
        Dim bln����ʱָ������ As Boolean
        Dim strAddInfo As String
        Dim str�걾�� As String, int_���� As Integer
    
        On Error GoTo DBErr
        '�������������ݣ�����������ǲ����գ�����ҽ��������ִ��״̬Ϊ0 ����ָ���˱��ŵ�������Ҫ�Ⱥ���,��д���ź��ٷ��ͣ�����ָ�����ŵ������Ͳ���ִ��״̬��
    
100     If mlng�������Ѻ��ձ걾 = 0 Then
102         bln����ʱָ������ = False
104         gstrSQL = "Select ����ʱָ������ From �������� Where Id = [1]"
106         Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "ȡ��������", lngDeviceID)
108         Do Until rsTmp.EOF
110             bln����ʱָ������ = Val("" & rsTmp!����ʱָ������) = 1
112             rsTmp.MoveNext
            Loop
        Else
114         bln����ʱָ������ = True
        End If
    
116     If Len(strAdviceIDs) = 0 Or Val(strAdviceIDs) = 0 Then
118         If Len(Trim(strBarcode)) = 0 Then
                '���걾��Ų�ѯ, 2013-11-12 ���ϼ�����Ϊ�յĲ���
120             gstrSQL = "Select TO_CHAR(A.����ʱ��, 'MM-DD HH24:MI') AS �걾ʱ��,A.�걾��� AS �걾��,F.����,D.������,D.Ӣ����,C.ͨ������,A.�걾����,A.��������, A.����, E.������־, A.�걾��� ,A.��������,A.�Ա�,A.����,A.����id" & _
                    " From ����걾��¼ A,������ͨ��� B,����������Ŀ C,����������Ŀ D,����ҽ����¼ E,������Ϣ F,������Ŀ G,����ҽ������ H " & _
                    " Where A.ID+0=B.����걾ID And A.������=B.��¼���� And B.������ĿID+0=C.��ĿID And C.����ID=[6] And B.������ĿID+0=D.ID" & _
                    " And A.ҽ��ID+0=E.ID And E.����ID+0=F.����ID And D.ID=G.������ĿID And A.����ID=[1]" & _
                    " And A.����ʱ�� BETWEEN [2] AND [3] And E.id=H.ҽ��ID " & IIf(bln����ʱָ������ = True, "", " And H.ִ��״̬ = 0 ") & _
                    " And B.������ Is Null And A.�걾���=[4] And G.��Ŀ���<>3 And C.ͨ������<>'0' " & IIf(gblnEmerge, " and nvl(a.�걾���,0)  = [5] ", "")
122             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "������������", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                    CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), iType, lngMainID)
124             Call WriteLog("GetSampleInfo", LOG_ͨѶ��־, 0, "���걾��:" & lngDeviceID & "," & strSampleNO & "," & strSampleDate & "," & iType & "," & mlng�������Ѻ��ձ걾)
            
            Else
                '���������  2013-11-12 ���ϼ�����Ϊ�յĲ���
126             gstrSQL = "Select TO_CHAR(A.����ʱ��, 'MM-DD HH24:MI') AS �걾ʱ��,A.�걾��� AS �걾��,F.����,D.������,D.Ӣ����,C.ͨ������,A.�걾����,A.��������, A.����, E.������־, A.�걾��� ,A.��������,A.�Ա�,A.����,A.����id " & _
                    " From ����걾��¼ A,������ͨ��� B,����������Ŀ C,����������Ŀ D,����ҽ����¼ E,������Ϣ F,������Ŀ G,����ҽ������ H" & _
                    " Where A.ID+0=B.����걾ID And A.������=B.��¼���� And B.������ĿID+0=C.��ĿID And C.����ID=[7] And B.������ĿID+0=D.ID" & _
                    " And A.ҽ��ID+0=E.ID And E.����ID+0=F.����ID And D.ID=G.������ĿID And A.����ID=[1]" & _
                    " And A.����ʱ�� BETWEEN [2] AND [3] And E.id=H.ҽ��ID " & IIf(bln����ʱָ������ = True, "", " And H.ִ��״̬ = 0 ") & _
                    " And B.������ Is Null And A.��������=[5] And G.��Ŀ���<>3 And C.ͨ������<>'0' "
128             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "������������", lngDeviceID, CDate(Format(strSampleDate, "yyyy-mm-dd") & " 00:00:00"), _
                    CDate(Format(strSampleDate, "yyyy-mm-dd") & " 23:59:59"), Val(strSampleNO), strBarcode, iType, lngMainID)
130             If rsTmp.EOF Then
                    '����ҽ��  ҽ��״̬=8 ����ҽ���������������ͺ������ֹͣ
132                 gstrSQL = "Select TO_CHAR(F.����ʱ��, 'MM-DD HH24:MI') AS �걾ʱ��,0 AS �걾��,A.������־," & _
                        "C.����||Decode(A.Ӥ��,0,'',Null,'','(Ӥ��)') As ����,Y.ͨ������,A.�걾��λ As �걾����,F.��������,'' as ����, 0 as �걾���,C.��������,C.�Ա�,C.����,C.����ID " & _
                        "FROM ����ҽ����¼ A," & _
                        "������Ϣ C,����ҽ������ F,���鱨����Ŀ G,������Ŀ I,����������Ŀ Y " & _
                        "WHERE A.������� = 'C' " & _
                        "AND A.����ID=C.����ID " & IIf(bln����ʱָ������ = True, "", " And F.ִ��״̬ = 0 ") & _
                        "AND A.���id IS NOT NULL " & _
                        "AND A.ҽ��״̬=8 AND A.ID=F.ҽ��id " & _
                        "AND A.������Ŀid=G.������Ŀid AND G.ϸ��ID Is Null AND G.������Ŀid=Y.��Ŀid " & _
                        "AND G.������ĿID=I.������ĿID " & _
                        "AND Y.����ID+0=[1] " & _
                        "And F.��������=[2] " & _
                        " And Y.ͨ������<>'0' "

134                 Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "������������", lngMainID, strBarcode, iType)
                End If
            End If
136         GetSampleInfo = ""
138         If Not rsTmp.EOF Then
140             int_���� = IIf(gblnEmerge, Val("" & rsTmp!������־), 0)

142             GetSampleInfo = Format(rsTmp("�걾ʱ��"), "yyyy-MM-dd HH:mm:ss")
144             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("�걾��"), " ")
146             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("����"), " ")
148             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("�걾����"), " ") & "|" & int_����
                '2013-11-07 add by cd ҽ���Ժ��������ˮ����Ҫ����Ϣ���
                 
150             strAddInfo = Val("" & rsTmp!����id) & "^" & Trim("" & rsTmp!�Ա�) & "^" & Format(rsTmp!��������, "YYYY-MM-DD") & "^" & Trim("" & rsTmp!����) & "^" & _
                             gobjCommFun.mGetFullPY("" & rsTmp("����")) & "^"  '�˴�ϡ�ͱ���Ϊ�գ������°�LIS
152             GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("��������"), " ") & "|" & IIf(Trim("" & rsTmp("����")) = "", " ", Trim("" & rsTmp("����"))) & "|" & strAddInfo & "| | "
154             Do While Not rsTmp.EOF
156                 GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("ͨ������"), " ")
158                 rsTmp.MoveNext
                Loop
            End If
        Else '��ҽ��ID��ѯ
160         aAdviceIDs = Split(strAdviceIDs, ",")
162         GetSampleInfo = ""
164         For i = 0 To UBound(aAdviceIDs)
166             lngAdviceID = Val(aAdviceIDs(i))
        
168             gstrSQL = "Select TO_CHAR(A.����ʱ��, 'MM-DD HH24:MI') AS �걾ʱ��,A.�걾��� AS �걾��,F.����,D.������,D.Ӣ����,C.ͨ������,A.�걾����,'' As ��������, A.����, E.������־, A.�걾���,A.��������,A.�Ա�,A.����,A.����ID " & _
                    " From ����걾��¼ A,������Ŀ�ֲ� B,����������Ŀ C,����������Ŀ D,����ҽ����¼ E,������Ϣ F,������Ŀ G,����ҽ������ H " & _
                    " Where A.ID=B.�걾ID+0 And B.��ĿID+0=C.��ĿID And C.����ID=[4] And B.������ĿID+0=D.ID" & _
                    " And B.ҽ��ID=E.ID And E.����ID+0=F.����ID And D.ID=G.������ĿID And A.����ID=[1] And E.id=H.ҽ��ID " & IIf(bln����ʱָ������ = True, "", " And H.ִ��״̬ = 0 ") & _
                    " And B.ҽ��ID=[2] And G.��Ŀ���<>3 And C.ͨ������<>'0' " & IIf(gblnEmerge, " and nvl(a.�걾���,0)  = [3] ", "")
170             Set rsTmp = gobjDatabase.OpenSqlRecord(gstrSQL, "������������", lngDeviceID, lngAdviceID, iType, lngMainID)
172             If Not rsTmp.EOF Then
174                 If Len(GetSampleInfo) = 0 Then
176                     int_���� = Val("" & rsTmp!������־)

178                     GetSampleInfo = Format(rsTmp("�걾ʱ��"), "yyyy-MM-dd HH:mm:ss")
180                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("�걾��"), " ")
182                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("����"), " ")
184                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("�걾����"), " ") & "|" & int_����
                        '2013-11-07 add by cd ҽ���Ժ��������ˮ����Ҫ����Ϣ���
186                     strAddInfo = Val("" & rsTmp!����id) & "^" & Trim("" & rsTmp!�Ա�) & "^" & Format(rsTmp!��������, "YYYY-MM-DD") & "^" & Trim("" & rsTmp!����) & _
                                     "^" & gobjCommFun.mGetFullPY("" & rsTmp("����")) & "^"   '�˴�ϡ�ͱ���Ϊ�գ������°�LIS'�˴�ϡ�ͱ���Ϊ�գ������°�LIS
188                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("��������"), " ") & "|" & IIf(Trim("" & rsTmp("����")) = "", " ", Trim("" & rsTmp("����"))) & "|" & strAddInfo & "| | "
                    End If
190                 Do While Not rsTmp.EOF
192                     GetSampleInfo = GetSampleInfo & "|" & Nvl(rsTmp("ͨ������"), " ")
                
194                     rsTmp.MoveNext
                    Loop
                End If
            Next
        End If
196     Call WriteLog("getSampleInfo", LOG_ͨѶ��־, 0, GetSampleInfo)
        Exit Function
DBErr:
198     GetSampleInfo = ""
200     Call WriteLog("GetSampleInfo", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function

Private Function InDataBase(ByVal strResult As String) As Boolean
        '�������ݵ����ݿ�
        Dim strErr As String, lngErr As Long, strQCComputeInfo As String, strUnkonw As String
        Dim strIDs As String, varIds As Variant, i As Integer
        Dim strlogs As String
        On Error GoTo hErr
    
100     If SaveToDataBase(mlngDeviceID, mlngManID, mlngExeDeptID, mintMicrobe, mintAutoQCCalc, mstrAutoCheckMan, strResult, mItem, strUnkonw, strQCComputeInfo, lngErr, strErr, strIDs, strlogs) = True Then
102         InDataBase = True
104         If strIDs <> "" Then
106             varIds = Split(strIDs, ",")
108             For i = LBound(varIds) To UBound(varIds)
110                 If Val("" & varIds(i)) <> 0 Then
112                     RaiseEvent DevRefresh(Val("" & varIds(i)))
                    End If
                Next
            End If
114         If strQCComputeInfo <> "" Then
116             RaiseEvent ReturnCompute(strQCComputeInfo)
            End If
118         If strUnkonw <> "" Then
120             If g����(mintIndex).���� = 0 Then
122                 RaiseEvent ItemUnknown(g����(mintIndex).COM��, strUnkonw)
                Else
124                 RaiseEvent ItemUnknown(g����(mintIndex).IP, strUnkonw)
                End If
            End If
126         If strlogs <> "" Then
128             Call WriteToSendDir(strlogs, "SaveDataLog")
            End If
        End If
        Exit Function
hErr:
130     Call WriteLog("GetSampleInfo", LOG_������־, Err.Number, CStr(Erl()) & "�г��ִ���  " & Err.Description)
End Function








