VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIS�����ϴ� v1.2"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6090
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdPara 
      Caption         =   "��������(&P)"
      Height          =   350
      Left            =   3000
      TabIndex        =   5
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4800
      TabIndex        =   4
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Timer TimerTrans 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "���ݿ�����(&D)"
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   5800
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "��ʼ�ϴ�(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox lstLog 
      Height          =   5280
      ItemData        =   "frmMain.frx":030A
      Left            =   120
      List            =   "frmMain.frx":030C
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOutConnect As Boolean   '�ⲿ���ݿ��Ƿ�����

Private mlngҩ��id As Long
Private mstrҩ������ As String
Private mlng��ѯ��� As Long
Private mint��ѯ���� As Integer
Private mstr���� As String
Private mstr��ʼʱ�� As String
Private mstr����ʱ�� As String
Private mstrUpdate As String        'Ҫ���µ����ݣ�����,NO|����,NO

Private Function GetHisData() As Variant
    '��ȡHIS����
    '��δ��ҩƷ��¼��ȡ��Ӧ��NO
    '����NO��ȡҩƷ��ϸ��Ϣ������Ƶ�ηֽ��ÿһ������Ϣ
    '��NO��ҩƷ�ж������ʱ��ֻ����һ��
    
    Dim rsData As ADODB.Recordset
    Dim rsDataDrug As ADODB.Recordset
    Dim rsGetNext As ADODB.Recordset
    Dim n As Integer
    Dim strReturn As String
    Dim strLastTime As String
    Dim intCount As Integer
    Dim varReturn As Variant
    Dim str��ҩ���ű��� As String
    Dim strPatiid As String
    Dim str�ְ��豸��� As String
    Dim str����     As String
    Dim strNO As String
    Dim strҩ������ As String
    Dim lngҩƷid As Long
    Dim int���� As Integer
    Dim strDeptId As String     '��������
    Dim intMBType As Integer    '�������ҵ��ݹ���0-���ϴ�����,1-ֻ�ϴ����۵�,2-ֻ�ϴ��շѵ�,3-���е��ݶ��ϴ�
    Dim intFMBType As Integer   '���������ҵ��ݹ���0-���ϴ�����,1-ֻ�ϴ����۵�,2-ֻ�ϴ��շѵ�,3-���е��ݶ��ϴ�
    
    str�ְ��豸��� = "1"
    
    varReturn = Array()
    GetHisData = Array()
    
    On Error GoTo errHandle
    
    strDeptId = GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "��������", "")
    intMBType = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "�������ҵ�������", 0))
    intFMBType = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "���������ҵ�������", 0))
    
    '�������ֵΪ���ϴ����ݣ���ִ�в�ѯ
    If intMBType = 0 And intFMBType = 0 Then Exit Function
    
    '��ȡָ��ʱ��ε�δ�ϴ�������Ϣ
    gstrSql = "Select f.���� as ��ҩ���ű���, g.���� ҩ������, a.����id,a.����,a.����, a.No From δ��ҩƷ��¼ A, ���ű� F, ���ű� G "
    
    'δ����ҩƷ�շ���¼��ע�͵�
'    If mstr���� <> "" Then
'        gstrSql = gstrSql & " ,ҩƷ��� C, ҩƷ���� D, Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)) E "
'    End If

    gstrSql = gstrSql & " Where a.�Է�����id = f.Id and a.�ⷿid=g.id And a.���� In (8, 9) And Nvl(a.�Ƿ��ϴ�, 0) = 0 And a.�ⷿid = [1] And a.�������� Between [2] And [3] "

    'δ����ҩƷ�շ���¼��ע�͵�
'    If mstr���� <> "" Then
'        gstrSql = gstrSql & " And b.ҩƷid = c.ҩƷid And c.ҩ��id = d.ҩ��id And d.ҩƷ���� = e.Column_Value"
'    End If
    
    '�������ҵ����ϴ�����
    If strDeptId = "" Then
        '���û��ѡ���������ң������п��Ҷ��Ƿ��������ҹ�����
        If intFMBType = 1 Then
            'ֻ���۵�
            gstrSql = gstrSql & " And A.���շ� = 0 "
        ElseIf intFMBType = 2 Then
            'ֻ�շѵ�
            gstrSql = gstrSql & " And A.���շ� = 1 "
        ElseIf intFMBType = 3 Then
            '���۵����շѵ���������������
        End If
    Else
        'ѡ�����������ң� ���ֱ�Ĺ�����
        If intMBType = 3 And intFMBType = 3 Then
            '������е��ݶ����򲻼�����
        Else
            gstrSql = gstrSql & " And ("
            
            '�������Ҵ������
            If intMBType = 1 Then
                'ֻ���۵�
                gstrSql = gstrSql & " (Instr([5], ',' || a.�Է�����id || ',', 1) > 0 And ���շ� = 0) "
            ElseIf intMBType = 2 Then
                'ֻ�շѵ�
                gstrSql = gstrSql & " (Instr([5], ',' || a.�Է�����id || ',', 1) > 0 And ���շ� = 1) "
            ElseIf intMBType = 3 Then
                '���۵����շѵ�����
                gstrSql = gstrSql & " Instr([5], ',' || a.�Է�����id || ',', 1) > 0  "
            End If
            
            gstrSql = gstrSql & " Or "
            
            '���������Ҵ������
            If intFMBType = 1 Then
                'ֻ���۵�
                gstrSql = gstrSql & " (Instr([5], ',' || a.�Է�����id || ',', 1) = 0 And ���շ� = 0) "
            ElseIf intFMBType = 2 Then
                'ֻ�շѵ�
                gstrSql = gstrSql & " (Instr([5], ',' || a.�Է�����id || ',', 1) = 0 And ���շ� = 1) "
            ElseIf intFMBType = 3 Then
                '���۵����շѵ�����
                gstrSql = gstrSql & " Instr([5], ',' || a.�Է�����id || ',', 1) = 0 "
            End If
             
            gstrSql = gstrSql & ")"
        End If
    End If
    
    gstrSql = gstrSql & " Order By f.����, a.����id, a.No "

    Set rsData = OpenSQLRecord(gstrSql, "HisTransData", mlngҩ��id, CDate(mstr��ʼʱ��), CDate(mstr����ʱ��), mstr����, strDeptId)
    
    If rsData.RecordCount = 0 Then Exit Function
    
    str��ҩ���ű��� = rsData!��ҩ���ű���
    strPatiid = rsData!����id
    strNO = rsData!NO
    strҩ������ = rsData!ҩ������
    
    Do While Not rsData.EOF
'        '��NO��֯����
'        If str��ҩ���ű��� & "," & strPatiid & "," & strNO <> rsData!��ҩ���ű��� & "," & rsData!����id & "," & rsData!NO And strReturn <> "" Then
'            strReturn = str��ҩ���ű��� & ";" & str���� & ";" & str�ְ��豸��� & ";" & strNO & "|" & strReturn
'
'            ReDim Preserve varReturn(UBound(varReturn) + 1)
'            varReturn(UBound(varReturn)) = strReturn
'
'            strReturn = ""
'        End If
                
        str��ҩ���ű��� = rsData!��ҩ���ű���
        strPatiid = rsData!����id
        strNO = rsData!NO
        strҩ������ = rsData!ҩ������
        int���� = rsData!����
        
        '��ʼ����
        lngҩƷid = 0
        strReturn = ""
        
        '��ȡҩƷ��Ϣ����ָ��NO������±���Ҫ��ҩƷID����
        gstrSql = " Select a.�շ�id, a.סԺ��, a.����id,A.����, a.���ұ���, a.��������, a.������, a.����, a.�÷�, a.ҽ������ ,a.ҩƷ����, a.ҩƷ����, a.���, a.����ϵ��, a.������λ,a.���㵥λ, a.��������,A.�ܸ�����," & vbNewLine & _
        "                   a.�״�ʱ��, a.ĩ��ʱ��, a.��ʼִ��ʱ��, a.Ƶ�ʼ��, a.�����λ, a.ִ��ʱ�䷽��, Nvl(b.��������, 0) As ���� ,a.����,a.Ч��, a.��ҩ����, ����װ,a.ִ��Ƶ��,a.����,a.ҩƷid " & vbNewLine & _
        "            From (Select Distinct a.Id As �շ�id, b.��ʶ�� As סԺ��, b.����id, b.����, c.���� As ���ұ���, c.���� As ��������, b.������, '' As ����, a.�÷�,h.ҽ������," & vbNewLine & _
        "                                  d.���� As ҩƷ����, d.���� As ҩƷ����, d.���, e.����ϵ��, d.���㵥λ,f.���㵥λ As ������λ, h.��������/e.����ϵ��  As ��������,round(H.�ܸ�����,0) �ܸ�����,g.�״�ʱ��, g.ĩ��ʱ��," & vbNewLine & _
        "                                  h.��ʼִ��ʱ��, h.Ƶ�ʼ��, h.�����λ, h.ִ��ʱ�䷽��, h.���id, g.���ͺ�, a.ʵ������ * Nvl(a.����, 1) / e.�����װ As ��ҩ����," & vbNewLine & _
        "                                  Decode(Mod(a.ʵ������ * Nvl(a.����, 1), e.ҩ���װ), 0, 1, 0) ����װ,h.ִ��Ƶ��,h.���� ,a.����,a.Ч��,a.ҩƷid " & vbNewLine & _
        "                  From ҩƷ�շ���¼ A, ������ü�¼ B, ���ű� C, �շ���ĿĿ¼ D, ҩƷ��� E, ������ĿĿ¼ F, ����ҽ������ G, ����ҽ����¼ H" & vbNewLine & _
        "                  Where a.����id = b.Id And b.��������id= c.Id And a.ҩƷid = d.Id And b.��¼״̬ in (1,3) and a.ҩƷid = e.ҩƷid And e.ҩ��id = f.Id And" & vbNewLine & _
        "                        b.ҽ����� = g.ҽ��id And b.No = g.No And b.ҽ����� = h.Id And a.�ⷿid = [1] And  a.���� = [2] And a.No = [3]  ) A, ����ҽ������ B" & vbNewLine & _
        "             Where a.���id = b.ҽ��id(+) And a.���ͺ� = b.���ͺ�(+) And a.�÷�='�ڷ�'" & vbNewLine & _
        "            Order By a.ҩƷid "
        Set rsDataDrug = OpenSQLRecord(gstrSql, "HisTransData", mlngҩ��id, rsData!����, rsData!NO)
        
        str���� = ""
        With rsDataDrug
            'ѭ���������ݵ�ҩƷ
            Do While Not .EOF
                
                If str���� = "" Then str���� = NVL(!����)
                
                If lngҩƷid <> !ҩƷid Then
                    '����ҩ������ҩƷ��ͬһ��NO�����ж�����Σ�ֻ��һ�����εĽ��зֽ⣬���ҩƷID��ͬ�򲻴���
                    lngҩƷid = !ҩƷid
                    
                    If Val(NVL(!Ƶ�ʼ��, 0)) = 0 Or NVL(!�����λ, "") = "" Or NVL(!ִ��ʱ�䷽��, "") = "" Then
                        intCount = 1
                    Else
                        intCount = Val(!����)
                        If intCount = 0 Then
                            gstrSql = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) From Dual "
                            Set rsGetNext = OpenSQLRecord(gstrSql, "ȡ�´�ִ��ʱ��", CDate(!��ʼִ��ʱ��), CDate(!�״�ʱ��), CDate(!ĩ��ʱ��), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                            If Not rsGetNext.EOF Then
                                intCount = Val(rsGetNext.Fields(0).Value)
                            End If
                        End If
                        If intCount = 0 Then
                            intCount = 1
                        End If
                    End If
                    
                    For n = 1 To intCount
                        strReturn = IIf(strReturn = "", "", strReturn & "|")
                        strReturn = strReturn & !�շ�id
                        strReturn = strReturn & ";" & NVL(!סԺ��)
                        strReturn = strReturn & ";" & NVL(!����id)
                        strReturn = strReturn & ";" & Replace(Replace(!����, ";", ""), "|", "")
                        strReturn = strReturn & ";" & !���ұ���
                        strReturn = strReturn & ";" & Replace(Replace(!��������, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(!������, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(NVL(!����, ""), ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(NVL(!�÷�, ""), ";", ""), "|", "")
                        strReturn = strReturn & ";" & ""    '����ʱ��˵��
                        strReturn = strReturn & ";" & !ҩƷ����
                        strReturn = strReturn & ";" & Replace(Replace(!ҩƷ����, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(!���, ";", ""), "|", "")
                        strReturn = strReturn & ";" & NVL(!����ϵ��, 1) * NVL(!��������, 1)
                        strReturn = strReturn & ";" & !������λ
                        strReturn = strReturn & ";" & !�ܸ�����
                        
                        If n = 1 Then
                            strLastTime = Format(!�״�ʱ��, "YYYY-MM-DD HH:MM:SS")
                        Else
                            gstrSql = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) From Dual "
                            Set rsGetNext = OpenSQLRecord(gstrSql, "ȡ�´�ִ��ʱ��", CDate(!��ʼִ��ʱ��), CDate(strLastTime), Val(!Ƶ�ʼ��), !�����λ, !ִ��ʱ�䷽��)
                            If Not rsGetNext.EOF Then
                                strLastTime = Format(rsGetNext.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                            End If
                        End If
                        
                        strReturn = strReturn & ";" & strLastTime
                        strReturn = strReturn & ";" & "1"           '�ְ��豸���
                        strReturn = strReturn & ";" & "0"           '���ȱ��
                        strReturn = strReturn & ";" & "1"           '����
                        strReturn = strReturn & ";" & Replace(Replace(!��������, ";", ""), "|", "")
                        strReturn = strReturn & ";" & !ִ��Ƶ��
                        strReturn = strReturn & ";" & Format(!�״�ʱ��, "YYYY-MM-DD HH:MM:SS")
                        strReturn = strReturn & ";" & Format(!����, "0.0")
                        strReturn = strReturn & ";" & !ִ��ʱ�䷽��
                        strReturn = strReturn & ";" & NVL(!����)
                        strReturn = strReturn & ";" & NVL(!Ч��)
                        strReturn = strReturn & ";" & NVL(!��������, 1)
                    Next
                End If
                
                .MoveNext
            Loop
        End With
        
        If strReturn <> "" Then
            '��NO��֯�ϴ�����
            strReturn = str��ҩ���ű��� & ";" & str���� & ";" & str�ְ��豸��� & ";" & strNO & ";" & int���� & "|" & strReturn
            
            ReDim Preserve varReturn(UBound(varReturn) + 1)
            varReturn(UBound(varReturn)) = strReturn
            
'            '��¼Ҫ���µ�����
'            If InStr(1, mstrUpdate, rsData!���� & "," & rsData!NO) = 0 Then
'                mstrUpdate = IIf(mstrUpdate = "", "", mstrUpdate & "|") & rsData!���� & "," & rsData!NO
'            End If
            
            Call OutputLog("" & Now & vbCrLf & strReturn)
        End If
                        
        rsData.MoveNext
        
'        If rsData.EOF And strReturn <> "" Then
'            '����û�м�¼ʱ���������ݣ�������û�д��ݳɹ����շ�ID
'            strReturn = str��ҩ���ű��� & ";" & str���� & ";" & str�ְ��豸��� & ";" & strNO & "|" & strReturn
'
'            ReDim Preserve varReturn(UBound(varReturn) + 1)
'            varReturn(UBound(varReturn)) = strReturn
'
'        End If
    Loop
    
    GetHisData = varReturn
    
    Exit Function
    
errHandle:
'    If gobjComLib.ErrCenter() = 1 Then
'        Resume
'    End If
'    Call gobjComLib.SaveErrLog
    Call LogListItem(Err.Description)
    Call OutputLog("�쳣��" & Err.Description)
    Set varReturn = Nothing
End Function

Private Sub AutoTrans()
    Dim arrTrans As Variant
    Dim strReturn As String
    Dim i As Integer
    Dim int���� As Integer
    Dim strNO As String
    Dim strTmp As String
    
    On Error GoTo errHandle
    
    '�������ڷ�Χ
    Call UpdateDateValue
    
    '��ȡHIS����
    arrTrans = GetHisData()
       
    If UBound(arrTrans) = -1 Then
        LogListItem "���������ݣ�" & Now
        Exit Sub
    End If
    
    mstrUpdate = ""
    
    Me.cmdStart.Enabled = False
    
    '�����ϴ�����
    For i = 0 To UBound(arrTrans)
        strReturn = TranToPacker(CStr(arrTrans(i)))
        If strReturn <> "" Then
            LogListItem "�ϴ�ʧ�ܵ��շ�ID��" & strReturn
        Else
            '��¼�ύ�ɹ��ĵ��ݺ͵��ݺţ�������Ҫ�����ϴ���־
            strTmp = Left(arrTrans(i), InStr(arrTrans(i), "|") - 1)
            strNO = Split(strTmp, ";")(3)      'NO
            int���� = Split(strTmp, ";")(4)     '����
            If InStr(1, mstrUpdate, int���� & "," & strNO) = 0 Then
                mstrUpdate = IIf(mstrUpdate = "", "", mstrUpdate & "|") & int���� & "," & strNO
            End If
        End If
    Next
    
    '�����ϴ���־
    If mstrUpdate <> "" Then
         gstrSql = "Zl_δ��ҩƷ��¼_�����ϴ���־("
        '��ҩID,���
        gstrSql = gstrSql & mlngҩ��id
        '����,NO
        gstrSql = gstrSql & ",'" & mstrUpdate & "'"
        gstrSql = gstrSql & ")"
        Call ExecuteProcedure(gstrSql, "�����ϴ���־")
    End If
    
    Me.cmdStart.Enabled = True
    
    LogListItem "�����ϴ�������ɣ�" & Now
    
    Exit Sub
errHandle:
    Me.cmdStart.Enabled = True
    LogListItem Err.Description
End Sub

Private Function TranToPacker(ByVal strData As String) As String
'���ܣ� ����ҩƷ�Զ��ְ�����
'������ �ְ������ַ���
'��ʽ�� ��������;�ⷿ���;�ְ��豸���;NO;����|�շ�ID1;������;...|�շ�ID2;������;...|�շ�ID3;������;...
'���� �շ�ID,������,����ID,����,��������,��������,ҩʦ����,����,���÷���,��ҩʱ��˵��,
'       ҩƷ����,ҩƷ����,���,����,������λ,��������,����ʱ��,�ְ��豸���,ҽ������
'����ֵ��δ�ɹ����͵��շ�ID�ַ���
    Dim arrPrimary As Variant, arrSecondly As Variant, arrSecondlyVals As Variant
    Dim strInsert As String, strTmp As String, strID As String, strPageNO As String
    Dim i As Integer, j As Integer, intPageNO As Integer
    Dim rsInsert As New ADODB.Recordset
    Dim blnRollback As Boolean, blnInsert As Boolean, blnInserted As Boolean
    
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        MsgBox "��δ�������ݿ⣬����ִ��DBConnect()������", vbCritical, GSTR_MESSAGE
        TranToPacker = "NOT"
        Exit Function
    End If
    
    strTmp = Trim(strData)
    If strTmp = "" Then Exit Function
'     Exit Function
    arrPrimary = Split(Mid(strTmp, 1, InStr(1, strTmp, "|") - 1), ";")
    
    strTmp = Mid(strTmp, InStr(1, strTmp, "|") + 1)
    arrSecondly = Split(strTmp, "|")
    
''    ȡPageNO��
'    strTmp = "select convert(char(6),getdate(),12) + right('000000'+cast(isnull(max(substring(page_no,7,len(page_no))),0)+1 as varchar(4)),4) max_no " _
'           & "from dbo.atf_ypxx where convert(char(6),getdate(),12)=left(page_no,6)"
'    rsInsert.Open strTmp, gcnOutside
'    strPageNO = rsInsert!max_no
'    rsInsert.Close

    'ȡNO��ΪPageNO��ɽ����ú����Ҫ��
    strPageNO = arrPrimary(3)
    
    '�ȴ��ͱ�����(��)
'    intPageNO = 1   '����
'    intAbate = 0    '�ع���
    strInsert = "insert into dbo.atf_ypxx " _
              & "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" _
              & ",specification,dosage,dos_unit,total,occ_time,atf_no,pri_flag,Mz_flag,dept_name,freq,start_times,days,script,lot,expiredate,amount,page_no) " & Chr(13)
    strTmp = ""
    For i = LBound(arrSecondly) To UBound(arrSecondly)
        '�õ�Ԫ��
        arrSecondlyVals = Split(arrSecondly(i), ";")
        '��֯�ַ���
        strTmp = strTmp & "select "
        For j = LBound(arrSecondlyVals) To UBound(arrSecondlyVals)
            Select Case j
            Case 0
                strTmp = strTmp & "'" & arrSecondlyVals(j) & "'"
            Case 1 To 12, 14, 20, 16 To 18, 21 To 26
                strTmp = strTmp & ",'" & arrSecondlyVals(j) & "'"
            Case 13, 15, 19, 27
                strTmp = strTmp & "," & arrSecondlyVals(j)
            End Select
        Next
        strTmp = strTmp & ",'" & strPageNO & "'"
        strTmp = strTmp & " union all " & Chr(13)
        '�ж�������¼�Ƿ�Ϊͬһ�շ�ID
        strID = arrSecondlyVals(0)
        If i = UBound(arrSecondly) Then
            blnInsert = True
        Else
            If Mid(arrSecondly(i + 1), 1, InStr(1, arrSecondly(i + 1), ";") - 1) = strID Then
                blnInsert = False
            Else
                blnInsert = True
            End If
        End If
        '�Ƿ�ִ��Insert���
        If blnInsert = True Then
            blnRollback = False
            strTmp = Left(strTmp, Len(strTmp) - 11)
            
            gcnOutside.BeginTrans
            On Error GoTo errRollback
            rsInsert.Open strInsert & strTmp, gcnOutside
            On Error GoTo 0
            If blnRollback = False Then
                gcnOutside.CommitTrans
                blnInserted = True
            Else
'                intPageNO = intPageNO - intAbate - 1
                '��¼δ�ύ���շ�ID
                TranToPacker = TranToPacker & strID & ";"
            End If
            If rsInsert.State = adStateOpen Then rsInsert.Close
            strTmp = ""
'            intAbate = 0
        Else
            strTmp = strTmp & Chr(13)
            '��¼��������ͬ��
'            intAbate = intAbate + 1
        End If
'        intPageNO = intPageNO + 1
    Next
    If rsInsert.State = adStateOpen Then rsInsert.Close
    
    '�ȴ��ͱ�����(��)
    If blnInserted Then
        blnRollback = False
        strTmp = "insert into dbo.atf_yp_page_no (ward_sn,group_no,atf_no,submit_time,page_no,flag) " & Chr(13)
        strTmp = strTmp & "select "
        For i = LBound(arrPrimary) To UBound(arrPrimary)
            Select Case i
            Case 0 To 2
                strTmp = strTmp & "'" & arrPrimary(i) & "',"
'            Case 3
'                strTmp = strTmp & "getdate(),"
'            Case 4
'                strTmp = strTmp & "'" & strPageNO & "'"
            End Select
        Next
        strTmp = strTmp & "getdate(),'" & strPageNO & "',0"
        'strTmp = Left(strTmp, Len(strTmp) - 1)
        '�ύ����
        gcnOutside.BeginTrans
        On Error GoTo errRollback
        rsInsert.Open strTmp, gcnOutside
        On Error GoTo 0
        If blnRollback = False Then
            gcnOutside.CommitTrans
        Else
            '�����������ʧ�ܣ�ͬ��ɾ���ӱ��Ӧ����
            strTmp = "delete dbo.atf_ypxx where page_no='" & strPageNO & "'"
            On Error Resume Next
            If rsInsert.State = adStateOpen Then rsInsert.Close
            rsInsert.Open strTmp, gcnOutside
            If rsInsert.State = adStateOpen Then rsInsert.Close
            '���������շ�ID�ַ���
            strID = "": TranToPacker = ""
            For i = LBound(arrSecondly) To UBound(arrSecondly)
                If Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1) <> strID Then
                    strID = Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1)
                    TranToPacker = TranToPacker & strID & ";"
                End If
            Next
        End If
    End If
    'If gcnOutside.State = adStateOpen Then gcnOutside.Close
    If Trim(TranToPacker) <> "" Then
        '�����շ�ID�ַ���
        TranToPacker = Left(TranToPacker, Len(TranToPacker) - 1)
    End If
    
    Exit Function

errRollback:
    Call OutputLog("TranToPacker: " & Err.Description)
    gcnOutside.RollbackTrans
    blnRollback = True
    Resume Next
End Function


Private Sub cmdConnect_Click()
    frmOutsideLinkSet.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPara_Click()
    frmPara.Show 1, Me
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Tag = "0" Then
        cmdStart.Tag = "1"
        cmdStart.Caption = "ֹͣ�ϴ�(&S)"
        
        '��ʼ�ϴ�
        TimerTrans.Enabled = True
        
        LogListItem "��ʼ�ϴ���" & Now
        
        cmdConnect.Enabled = False
        cmdPara.Enabled = False
    Else
        cmdStart.Tag = "0"
        cmdStart.Caption = "��ʼ�ϴ�(&S)"
        
        'ֹͣ�ϴ�
        TimerTrans.Enabled = False
        
        LogListItem "ֹͣ�ϴ�" & Now
        
        cmdConnect.Enabled = True
        cmdPara.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    '��ʼ����������

'    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
'
'    If gobjComLib Is Nothing Then
'        MsgBox "��ʼ����������ʧ�ܣ�", vbInformation, ""
'        Unload Me
'    End If
    
    '�����ⲿ���ݿ�
    mblnOutConnect = DBConnect
    
    '��ȡע������
    mlngҩ��id = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "ҩ��ID"))
    mstrҩ������ = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "ҩ������"))
    mlng��ѯ��� = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "��ѯ���", 60))
    mint��ѯ���� = Val(GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "��ѯ����", 0))
    mstr���� = GetSetting("ZLSOFT", "����ģ��\����ҩ����ҩ��", "����", "")
    
    If mlng��ѯ��� > 60 Then
        mlng��ѯ��� = 60
    End If
    TimerTrans.Interval = mlng��ѯ��� * 1000
    
    '�������ڷ�Χ
    Call UpdateDateValue
    
End Sub

Private Sub UpdateDateValue()
    If mint��ѯ���� = 0 Then
        'Ĭ���ǵ���
        mstr��ʼʱ�� = Format(Currentdate, "YYYY-MM-DD")
        mstr����ʱ�� = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    Else
        'ָ��������
        If mint��ѯ���� > 3 Then
            mint��ѯ���� = 3
        End If
        
        mstr��ʼʱ�� = Format(DateAdd("d", -mint��ѯ����, Currentdate), "YYYY-MM-DD")
        mstr����ʱ�� = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStart.Enabled = False Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set gobjComLib = Nothing
End Sub

Private Sub TimerTrans_Timer()
    TimerTrans.Enabled = False
    
    On Error GoTo errHandle
    
    '�������
    If gcnOracle.State <> adStateOpen Then
        gcnOracle.Open
    End If
    If gcnOutside.State <> adStateOpen Then
        gcnOutside.Open
    End If
    
    DoEvents
    '�����Զ��ϴ�����
    Call AutoTrans
    DoEvents
    TimerTrans.Enabled = True
    
    Exit Sub
    
errHandle:
    Call LogListItem("�쳣��" & Err.Description)
    TimerTrans.Enabled = True
End Sub

Private Sub LogListItem(ByVal strLog As String)
    Const INT_MAX_LINES As Integer = 200

    Me.lstLog.AddItem strLog
    Me.lstLog.Selected(Me.lstLog.ListCount - 1) = True
    Me.lstLog.TopIndex = Me.lstLog.ListCount - 1
    If lstLog.ListCount >= INT_MAX_LINES Then lstLog.RemoveItem 0

End Sub
