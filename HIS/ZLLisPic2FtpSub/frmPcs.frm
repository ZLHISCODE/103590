VERSION 5.00
Begin VB.Form frmPcs 
   Caption         =   "�ӽ���-���ɼ�"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4800
   Icon            =   "frmPcs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   4800
   StartUpPosition =   2  '��Ļ����
   Begin VB.Label lblPro 
      Caption         =   "��ǰ����:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSta 
      AutoSize        =   -1  'True
      Caption         =   $"frmPcs.frx":6852
      Height          =   720
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "frmPcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSave As Boolean
Private mclsFtp As New clsFtp   'FTP��

Private Sub Form_Load()
    Dim strCmd As String
    Dim strServerName As String, strUserName As String
    Dim strUserPwd As String, strPro As String
    Dim intUserPosition As Integer, intPwdPosition As Integer
    Dim intServerPosition As Integer
    Dim strFtp As String

    On Error GoTo errH
    '��������潫��¼��Ϣ����,��ʽ:·�� �û��� ����
    '"Path\zlLisPic2FtpSub.exe zlUserName=" & gstrUserName & "zlPassword=HIS" & "zlServer=" & gstrServer
    strCmd = Command
    If strCmd <> "" Then
        intUserPosition = InStr(1, strCmd, "zlUserName=") + Len("zlUserName=")
        intPwdPosition = InStr(1, strCmd, "zlPassword=") + Len("zlPassword=")
        intServerPosition = InStr(1, strCmd, "zlServer=") + Len("zlServer=")
        
        strUserName = Mid(Left(strCmd, InStr(1, strCmd, "zlPassword=") - 1), intUserPosition)
        strUserPwd = Mid(Left(strCmd, InStr(1, strCmd, "zlServer=") - 1), intPwdPosition)
        strServerName = Mid(strCmd, intServerPosition)
        
        '���ݿ����ӳɹ�,��ִ�в���,����д�����
        If OraDataOpen(strServerName, strUserName, strUserPwd) Then
            '��ز�����ע����ж�ȡ,��ʽ: ת������(1-FTP 2-���汾��);������Դ(1-�ɰ�LIS 2-�°�LIS);���̺�;��ʼʱ��;����ʱ��;��ʱ·��;FTP·��
            strPro = GetSetting("LISͼƬת��", "ת������", "��������")
        
            'ֱ���ϴ���FTPģʽ��,��ȡFTP���û�\����\IP\·��
            If Split(strPro, ";")(0) = 1 Then
                strFtp = GetSetting("LISͼƬת��", "ת������", "FTP·��")
                mclsFtp.FuncFtpConnect Split(strFtp, ";")(2), Split(strFtp, ";")(0), Split(strFtp, ";")(1)
                mclsFtp.FuncChangeDir Split(strFtp, ";")(3)
            End If
            
            ImgUpload Split(strPro, ";")(0), Split(strPro, ";")(1), Split(strPro, ";")(2), Split(strPro, ";")(3), Split(strPro, ";")(4), Split(strPro, ";")(5), Split(strPro, ";")(6), Split(strPro, ";")(7)
                

        Else
            SaveSetting "LISͼƬת��", "ת������", "ת������", "���ݿ�����ʱ��������"
            WriteErrLog "���ݿ�����ʱ��������"
        End If
    End If
    
    Unload Me   'ִ�������,�˳�
    Exit Sub
    
errH:
    SaveSetting "LISͼƬת��", "ת������", "ת������", "��ʼ������ʱ��������"
    WriteErrLog Err.Description
    Unload frmPcs
End Sub


Private Sub ImgUpload(ByVal intType As Integer, ByVal intSource, ByVal intProc As Integer, ByVal intProcNum As Integer, ByVal strStart As String, ByVal strEnd As String, ByVal strTmpPath As String, ByVal strFtpPath As String)
    '����:�Խ��̺���Ϊѭ��ÿ���ͼƬ��Step,��ͼƬ�ϴ���FTP����������ת�浽����
    '����˵��: intType ת������ 1-ͬ�� 2-�첽  ; intSource 1-�ɰ�LIS 2-�°�LIS ;intProc ���̺� ;intProcNum ������; strStart-ת����ʼʱ�� ;strEnd -ת������ʱ��; strTmpPath-��ʱĿ¼; strFtpPath -FTPĿ¼
    Dim DateMin As Date, DateMax As Date, iDays As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsExp As ADODB.Recordset
    Dim blnDo As Boolean, strDes As String, strFile As String
    Dim strSrc As String, iDay As Integer, strͼ��λ�� As String
    Dim DateS As Date, DateE As Date
    Dim j As Integer, i As Long
    
    On Error GoTo errH
    '��ȡ��¼�Ŀ�ʼʱ��ͽ���ʱ��
    DateMin = CDate(strStart)
    DateMax = CDate(strEnd)

    iDays = DateDiff("d", DateMin, DateMax)
    strSrc = strTmpPath & "\"  '����Ŀ¼
    
    If intSource = 1 Then
        strSQL = "Select id From ����ͼ����_EXP_TEMP"
        Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")
    Else
        strSQL = "Select id From ���鱨��ͼ��_EXP_TEMP"
        Set rsExp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetExp")
    End If
    
    'ѭ��ÿ��ļ�¼,��ȡͼƬ
    For iDay = intProc - 1 To iDays Step intProcNum
        
        DateS = Format(DateAdd("d", -iDay, DateMax), "yyyy-MM-dd 00:00:00")
        DateE = Format(DateAdd("d", -iDay, DateMax), "yyyy-MM-dd 23:59:59")
        
        If GetSetting("LISͼƬת��", "ת������", "ת������") <> "" Then '������������̷�������,����ֹ
            WriteErrLog "�н����������,ת����ֹ"
            Exit Sub
        End If
        
        If CheckProcExist("zllispic2ftp.exe") = 0 Then   '�����̱���ֹ
            If CheckProcExist("zlSvrStudio.exe") = 0 Then
                WriteErrLog "�������������,ת����ֹ"
                Exit Sub
            End If
        End If
        
        If intSource = 1 Then
            strSQL = "Select /*+ rule */ b.ID,b.�걾id,a.����ID,b.ͼ������,b.ͼ��� " & vbNewLine & _
                            " From ����걾��¼ A, ����ͼ���� B " & vbNewLine & _
                            " Where a.����ʱ�� Between [1]  And  [2] And a.����� Is Not Null And a.Id = b.�걾id And b.ͼ��λ�� Is Null and b.ͼ��� Is Not Null"
        Else
            strSQL = "Select /*+ rule */ b.ID,b.�걾id,a.����ID,b.ͼ������,b.ͼ��� " & vbNewLine & _
                            " From ���鱨���¼ A, ���鱨��ͼ�� B " & vbNewLine & _
                            " Where a.����ʱ�� Between [1]  And  [2] And a.����� Is Not Null And a.Id = b.�걾id And b.ͼ��λ�� Is Null and b.ͼ��� Is Not Null"
        End If
        
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "GetImg", CDate(DateS), CDate(DateE))
        
        With rsTmp
            If .RecordCount > 0 Then
                Do While Not .EOF
                    '�жϵ�ǰ��¼�Ƿ񱣴��ڵ�������
                    blnDo = True
                    If rsExp.RecordCount > 0 Then
                        rsExp.Filter = "id=" & !id
                        blnDo = rsExp.RecordCount = 0
                        rsExp.Filter = 0
                    End If
                    
                    If blnDo Then
                        strFile = !�걾ID & "_" & !ͼ������     'ͼƬ����
                        strͼ��λ�� = "110;" & "/" & strFtpPath & "/Dev_" & !����ID & "/" & Format(DateS, "yyyyMM") & "/" & strFile & ".JPG"   '���ݿ��б��������
                        
                        If intType = 1 Then  'ѡ���ģʽ��ֱ���ϴ���FTP
                            strDes = "Dev_" & !����ID & "/" & Format(DateS, "yyyyMM")  'FTP����·��
                            If ImgSaveAsJpg(!ͼ������, !ͼ���, strTmpPath, strFile) Then  'ͼƬת�汾��
                                'ͼƬ�ϴ�
                                '·������  : Dev_����ID/ʱ��/
                                'ÿ�δ�������Ҫ��֤��ǰ·���Ǹ�·��
                                If i > 0 Then
                                    mclsFtp.FuncFtpCommand strFtpPath, "cdup"
                                    mclsFtp.FuncFtpCommand strFtpPath, "cdup"
                                End If
                                mclsFtp.FuncFtpCommand strFtpPath, "mkd " & "Dev_" & !����ID   '����Ҫ�ȴ����׼�Ŀ¼
                                mclsFtp.FuncFtpCommand strFtpPath, "mkd " & strDes
                                
                                If mclsFtp.FuncUploadFile(strDes, strSrc & strFile & ".JPG", strFile & ".JPG") <> 0 Then
                                    '����ϴ�ʧ��,��ɾ����ͼƬ,�´������ϴ�
                                    If gobjFile.FileExists(strSrc & strFile & ".JPG") Then
                                        Kill strSrc & strFile & ".JPG"
                                    End If
                                    WriteErrLog "�ϴ�ͼƬʱ��������,ת����ֹ,����ͼ��IDΪ:" & !id
                                    Exit Sub
                                Else
                                    If intSource = 1 Then
                                        Call ExecuteProcedure("Zl_����ͼ����_Temp_Insert(" & !id & "," & !�걾ID & ",'" & !ͼ������ & "','" & strͼ��λ�� & "')", Me.Caption)
                                    Else
                                        Call ExecuteProcedure("Zl_���鱨��ͼ��_Temp_Insert(" & !id & "," & !�걾ID & ",'" & !ͼ������ & "','" & strͼ��λ�� & "')", Me.Caption)
                                    End If
                                End If
                                If gobjFile.FileExists(strSrc & strFile & ".JPG") Then
                                    Kill strSrc & strFile & ".JPG"
                                End If
                            Else
                                WriteErrLog "����ͼƬʱ��������,ת����ֹ,����ͼ��IDΪ:" & !id
                                Exit Sub
                            End If
                            i = i + 1
                        Else
                            'ֱ�ӽ����ļ��б����ڱ���
                            If Not gobjFile.FolderExists(strTmpPath & "\" & "Dev_" & !����ID) Then
                                gobjFile.CreateFolder strTmpPath & "\" & "Dev_" & !����ID
                            End If
                            If Not gobjFile.FolderExists(strTmpPath & "\" & "Dev_" & !����ID & "\" & Format(DateS, "yyyyMM")) Then
                                gobjFile.CreateFolder strTmpPath & "\" & "Dev_" & !����ID & "\" & Format(DateS, "yyyyMM")
                            End If
                            strDes = strTmpPath & "\Dev_" & !����ID & "\" & Format(DateS, "yyyyMM")
                            If Not ImgSaveAsJpg(!ͼ������, !ͼ���, strDes, strFile) Then
                                WriteErrLog "����ͼƬʱ��������,ת����ֹ,����ͼ��IDΪ:" & !id
                                Exit Sub
                            Else
                                If intSource = 1 Then
                                    Call ExecuteProcedure("Zl_����ͼ����_Temp_Insert(" & !id & "," & !�걾ID & ",'" & !ͼ������ & "','" & strͼ��λ�� & "')", Me.Caption)
                                Else
                                    Call ExecuteProcedure("Zl_���鱨��ͼ��_Temp_Insert(" & !id & "," & !�걾ID & ",'" & !ͼ������ & "','" & strͼ��λ�� & "')", Me.Caption)
                                End If
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
        j = j + 1
        SaveSetting "LISͼƬת��", "ת������", "����" & intProc, j '���洦�����
    Next
    SaveSetting "LISͼƬת��", "ת������", "����" & intProc, j & ";" '�������,�ڽ��Ⱥ������һ�� ";" ��ʶ��ǰ�߳�ת�����
    Exit Sub
errH:
    SaveSetting "LISͼƬת��", "ת������", "ת������", "ת��ͼƬʱ��������"
    WriteErrLog Err.Description & " ����ͼƬ:" & strFile
    Unload frmPcs
End Sub


Private Function ImgSaveAsJpg(ByVal strType As String, ByVal strImgStrem As String, ByVal strPath As String, strFileName As String) As Boolean
    '����:���ݴ����ͼ���ת����ͼƬ���汾��,�ɹ�����True
    '����:strImgStrem-ͼ�����  strFileName-ͼƬ����    strPath-λ��
    Dim strFile As String
    Dim aryChunk() As Byte, intLayOut As Integer
    Dim lngFileNum As Long, lngBound As Long
    Dim frmObj As frmLisChart
    
    On Error GoTo errH
            
    '����Ŀ¼
    If Not gobjFile.FolderExists(strPath) Then
        gobjFile.CreateFolder strPath
    End If
    '�ж�ͼƬ����
    If Val(Mid(strImgStrem, 1, 3)) >= 100 And Val(Mid(strImgStrem, 1, 3)) <= 227 And Mid(strImgStrem, 4, 1) = ";" Then
        '���ݿ��б������ͼƬ����
        If Mid(strImgStrem, 1, 3) >= 100 And Mid(strImgStrem, 1, 3) <= 107 Then
            strFile = strPath & "\" & strFileName & "_Tmp.bmp"
        ElseIf Mid(strImgStrem, 1, 3) >= 110 And Mid(strImgStrem, 1, 3) <= 117 Then
            strFile = strPath & "\" & strFileName & "_Tmp.jpg"
        ElseIf Mid(strImgStrem, 1, 3) >= 120 And Mid(strImgStrem, 1, 3) <= 127 Then
            strFile = strPath & "\" & strFileName & "_Tmp.gif"
        ElseIf Mid(strImgStrem, 1, 3) >= 200 And Mid(strImgStrem, 1, 3) <= 227 Then
            If gobjFile.FolderExists(strPath & "\ZLLIS_ZIP") = False Then
                gobjFile.CreateFolder strPath & "\ZLLIS_ZIP"
            End If
            If gobjFile.FolderExists(strPath & "\ZLLIS_ZIP\" & strFileName) = False Then
                gobjFile.CreateFolder strPath & "\ZLLIS_ZIP\" & strFileName
            End If
            strFile = strPath & "\ZLLIS_ZIP\" & strFileName & "\ZLISPIC.ZIP"
        End If
    
        '������ݿ��б������ͼƬ����,��ֱ�ӱ��汾��,�ٽ��л�ͼ
        intLayOut = Val(Mid(strImgStrem, 1, 3))
        strImgStrem = Mid(Replace(Replace(Trim(strImgStrem), vbCr, ""), vbLf, ""), 5)
        If gobjFile.FileExists(strFile) Then
            Kill strFile '����ļ�����,��ɾ��
        End If
    
        '�����ļ�,ת���󱣴汾��
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        ReDim aryChunk(Len(strImgStrem) / 2 - 1) As Byte
        For lngBound = LBound(aryChunk) To UBound(aryChunk)
            aryChunk(lngBound) = CByte("&H" & Mid(strImgStrem, lngBound * 2 + 1, 2))
        Next
        Put lngFileNum, , aryChunk()
        Close lngFileNum
        
        strImgStrem = intLayOut & ";" & strFile
    Else
        strImgStrem = Replace(Replace(Trim(strImgStrem), vbCr, ""), vbLf, "")
    End If
    
    '����chart2D��ͼ ת����JPG
    Set frmObj = New frmLisChart
    If frmObj.DrawImg(strType, strImgStrem, strPath & "\" & strFileName & ".JPG", 1) Then
        ImgSaveAsJpg = True
    End If
    Unload frmObj
    
    Exit Function
errH:
    SaveSetting "LISͼƬת��", "ת������", "ת������", "ͼƬ���汾�ط�������"
    WriteErrLog Err.Description & " ����ͼƬ:" & strFileName
    Unload Me
End Function
