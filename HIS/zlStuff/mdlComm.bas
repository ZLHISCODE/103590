Attribute VB_Name = "mdlComm"
Option Explicit
Public Enum gEditType
     g���� = 0
     g�޸� = 1
     g��� = 2
     gȡ�� = 3
     g�鿴 = 4
End Enum
Public Enum RecBillStatus  '��¼״̬��Ϣ
    ������¼ = 1
    ������¼ = 2
    ��������¼ = 3
End Enum
Public Enum ErrBillStatusInfor  '����״̬��Ϣ
    ������� = 1
    �Ѿ�ɾ��
    �Ѿ����
    �Ѿ�����
End Enum
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum

Public gstrProductName As String

Public gblnCode As Boolean                   '�Ƿ���������������Ȩ��

Public Enum gС������
    g_���� = 0
    g_�ɱ���
    g_�ۼ�
    g_���
End Enum

Private Type m_С��λ
    ����С�� As Integer
    �ɱ���С�� As Integer
    ���ۼ�С�� As Integer
    ���С�� As Integer
End Type

Public Type m_��λС��
     obj_ɢװС�� As m_С��λ
     obj_��װС�� As m_С��λ
     obj_���С�� As m_С��λ
End Type

Public g_С��λ�� As m_��λС��

'С����ʽ����
Public Type g_FmtString
    FM_���� As String
    FM_�ɱ��� As String
    FM_���ۼ� As String
    FM_��� As String
    FM_ɢװ���ۼ� As String
End Type
Private Type mSystem_para
    int���뷽ʽ As Integer
    Para_���뷽ʽ   As String
    para_������¿��ÿ�� As Boolean
    bln����վ�� As Boolean      '�Ƿ����վ�����
    bln���￨������ʾ As Boolean  'true,������ʾ,false ��ʾˢ���Ŀ���
    str���￨ǰ׺�� As String    ' ��ž��￨���е���ĸǰ׺,��ͬǰ׺��|�ָ�,��:AA|BB|CC...
    P156_�����㷨 As Integer    '0-��������;1-Ч������
End Type
Public gSystem_Para As mSystem_para


'С��λ������
Public Const GFM_VBXS As String = "###0.000;-###0.000;0.000; "    '����ϵ��
Public Const GFM_VBCJL  As String = "#####0.00000;-#####0.00000;0.00000;"    'ָ�������
Public Const GFM_VBKL  As String = "#####0.0000;-#####0.0000;0.0000;"    '����
Public Const GFM_VBJCL  As String = "#####0.00;-#####0.00;0.00;"    '�ӳ���


Public Const GFM_XS As String = "'999999999990.999'"    '����ϵ��
Public Const GFM_CJL  As String = "'999999999990.99999'"    'ָ�������
Public Const GFM_KL  As String = "'999999999990.99999'"    '����
Public Const GFM_JCL  As String = "'999999999990.99'"    '�ӳ���

'���Ŀ���ѯ�У������α�����������ɫ����
Public Const glng���� As Long = &HC00000
Public Const glng���� As Long = &H80000008
Public Const glngͣ�� As Long = &HC0

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long


Public Function ExistsColObject(Col, index) As Boolean
    '�жϼ������Ƿ����ָ������(�ؼ���)�ĳ�Ա
    On Error GoTo ErrorHandler
    
    Dim v As Variant
    
    If TypeName(Col(index)) = "Collection" Then
        '������Ӧ�ĳ�Ա�Ǽ���ʱ
        ExistsColObject = True
        Exit Function
    Else
        '������Ӧ�ĳ�Ա�ǷǼ���ʱ
        v = Col(index)
        ExistsColObject = True
        Exit Function
    End If
ErrorHandler:
    '�쳣ʱ��ʾ��������Ӧ�ĳ�Ա
    ExistsColObject = False
End Function
Public Function GetCodePrivs() As Boolean
    '�ж�����ϵͳ�Ƿ������������Ȩ��
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From Zltools.zlRegFunc Where ϵͳ = 1 And ��� = 1711 And ���� = '�����������'"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, GetCodePrivs)
    GetCodePrivs = Not rsTemp.EOF
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Public Function getDept(strDept As String, Optional ByRef strID As String = "", Optional ByRef strName As String = "") As Boolean
'    Dim rsTemp As New ADODB.Recordset, strSQL As String
'    getDept = True
'    If strDept <> "" Then
'        If IsNumeric(strDept) Then
'            strSQL = "Select * From ��Ӧ�� Where ĩ��=1 And ����=[1]"
'        Else
'            strSQL = "Select * From ��Ӧ�� Where ĩ��=1 And (����=[1] Or ����=[1])"
'        End If
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��Ӧ�̼��", UCase(strDept))
'
'        If rsTemp.EOF Then
'            getDept = False
'        Else
'            strID = rsTemp!Id
'            strName = rsTemp!����
'        End If
'    End If
'End Function
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Public Sub SetCtlBackColor(objCtl As Object)
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:���ÿؼ��ı���ɫ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    If objCtl.Enabled Then
        objCtl.BackColor = &H80000005
    Else
        objCtl.BackColor = &H8000000A
    End If
End Sub

'ȡָ����ͷ����λ��
Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    Dim i As Integer
    
    On Error GoTo errH
    GetCol = -1
    With mshFlex
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = ColName Then
                GetCol = i
                Exit Function
            End If
        Next
        
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ��鵥��(ByVal lng���� As Long, ByVal strNo As String, Optional blnMsg As Boolean = True) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lng����_Last As Long, lng����_Cur As Long
    
    '������ĵļ۸��Ƿ�Ϊ���µļ۸������������
    '�����ڱ���ǰ�жϺ��鷳���Ҹ��ֵ��ݵı���б�������ݲ�һ������ˣ����������֮�����ύǰ���ѱ�������ݽ��м��
    '������ͬ�ļ�¼�Թ�
    On Error GoTo ErrHandle

    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = [2] And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(b.�ּ�, " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = [2] And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = [2] And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ")<>round(b.ƽ���ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " Order By ����, ����id, ���"

    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�ļ۸�", strNo, lng����)
      
    If rsTemp.EOF Then
        ��鵥�� = True
        Exit Function
    End If
    
    lng����_Last = 0
    With rsTemp
        Do While Not .EOF
            lng����_Cur = !����ID
            If lng����_Cur <> lng����_Last Then
                If blnMsg = True Then
                    If MsgBox("��" & !��� & "�����ĵ�" & !���� & "�������¼۸��Ƿ�������浥�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                Else
                    Exit Function
                End If
            End If
            
            lng����_Last = lng����_Cur
            .MoveNext
        Loop
        ��鵥�� = True
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ���ϵ������(ByVal str������ As String) As Boolean
    
    '���ϵ������ʱ���Ƿ��ж�������������ˣ��䷵����˽��
    Dim blnBillVerify As Boolean
    
    ���ϵ������ = True
    
    '���޴˹���,ԭ������������û��ҩƷ������ô��.��Ԥ���˲���
    blnBillVerify = Val(zldatabase.GetPara(64, glngSys, 0)) = 1
    If Not blnBillVerify Then Exit Function
    
    ���ϵ������ = (Trim(str������) <> Trim(UserInfo.�û���))
    If Not ���ϵ������ Then MsgBox "������������˲�����ͬһ�ˣ����飡", vbInformation, gstrSysName
End Function

Public Function ���������(ByVal lng�ⷿID As Long, ByVal lng����ID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln����Ƿ���� As Boolean, bln���� As Boolean, bln�ⷿ As Boolean
    
    'ͨ������ѡ������������ʱ��������Ŀ���е�������Ӳ������ʡ�����Ŀ¼�еķ��������жϳ��Ĳ�һ�£��򱨴�
    
    ��������� = False
    On Error GoTo ErrHandle
    
    '���û�п���¼����ֱ���˳�
    gstrSQL = "" & _
        "   Select Count(*) ��¼�� From ҩƷ��� " & _
        "   Where �ⷿID=[1] And ����=1 And ҩƷID=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����������Ƿ����", lng�ⷿID, lng����ID)
    If rsTemp!��¼�� = 0 Then
        ��������� = True
        Exit Function
    End If
    
    
    '���ڷ�����¼���������
    gstrSQL = " Select Count(*) ���� From ҩƷ��� " & _
              " Where �ⷿID=[1] And ����=1 And Nvl(����,0)<>0 And ҩƷID=[2]"
              
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "����������Ƿ����", lng�ⷿID, lng����ID)
    
    bln����Ƿ���� = (rsTemp!���� <> 0)
    
    '���ж��Ƿ��ǿⷿ
    gstrSQL = "select ����ID from ��������˵�� where (�������� like '���ϲ���' Or �������� like '%�Ƽ���') And ����id=[1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng�ⷿID)
    bln�ⷿ = (rsTemp.EOF)
        
    '�ж϶�Ӧ������Ŀ¼�еķ�������
    gstrSQL = "" & _
        "   Select Nvl(�ⷿ����,0) as �ⷿ����,nvl(���÷���,0) ���÷��� " & _
        "   From �������� Where ����ID=[1]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ����Ŀ¼�еķ�������", lng����ID)
    
    If bln�ⷿ Then
        bln���� = (rsTemp!�ⷿ���� = 1)
    Else
        bln���� = (rsTemp!���÷��� = 1)
    End If
    ��������� = (bln����Ƿ���� = bln����)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng����� As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    '��ָ���п�ʼ�������
    
    With mshBill
        lngRows = .Rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng�����) = lngRow
        Next
    End With
End Sub

Public Sub CheckLapse(ByVal strЧ�� As String)
    'ʧЧҩƷ���
    If Not IsDate(strЧ��) Then Exit Sub
    If Format(strЧ��, "yyyy-MM-dd") < Format(sys.Currentdate, "yyyy-MM-dd") Then
        MsgBox "�����������Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
    End If
End Sub

'ת����ֵΪ����
Public Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Public Function ��ͬ����(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_���� As Boolean, blnSecond_���� As Boolean
    ��ͬ���� = False
    
    blnFirst_���� = (sinFirst <= 0)
    blnSecond_���� = (sinSecond <= 0)
    
    ��ͬ���� = (blnFirst_���� = blnSecond_����)
End Function

'��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Public Function Get������(ByVal lng�ⷿID As Long) As Integer
    Dim rsSystemPara As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select Nvl(��鷽ʽ,0) ������ From ���ϳ����� Where �ⷿID=[1]"
    
    Set rsSystemPara = zldatabase.OpenSQLRecord(gstrSQL, "������", lng�ⷿID)
    
    If rsSystemPara.EOF Then
        Get������ = 0
        Exit Function
    End If
    Get������ = rsSystemPara!������
    rsSystemPara.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteSql(ByRef arrSQL As Variant, strTitle As String, _
Optional ByVal blnCommit As Boolean = True, Optional ByVal blnBeginTrans As Boolean = True) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer
    Dim intouter As Integer
    Dim intInner As Integer
    

    ExecuteSql = False
    If UBound(arrSQL) >= 0 Then
        '��SQL���а�ҩƷID��������
        intouter = UBound(arrSQL) - 1
        If Split(arrSQL(UBound(arrSQL)), ";" & vbCrLf)(0) = "����" Then
            intouter = UBound(arrSQL) - 2
        Else
            intouter = UBound(arrSQL) - 1
        End If
        
        
        For i = 0 To intouter
            For j = i + 1 To intouter + 1
                If CLng(Split(arrSQL(j), ";" & vbCrLf)(0)) < CLng(Split(arrSQL(i), ";" & vbCrLf)(0)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        If blnBeginTrans Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            zldatabase.ExecuteProcedure CStr(Split(arrSQL(i), ";" & vbCrLf)(1)), strTitle
        
'            Call SQLTest(App.ProductName, strTitle, CStr(Split(arrSql(i), ";" & vbCrLf)(1)))
'            gcnOracle.Execute CStr(Split(arrSql(i), ";" & vbCrLf)(1)), , adCmdStoredProc
'            Call SQLTest
        Next
        If blnCommit Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
       
errH:
    If blnBeginTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ReturnSQL(ByVal lng�ⷿID As Long, ByVal strCaption As String, _
    Optional ByVal bln���� As Boolean = True, _
    Optional ByRef strOutSQL As String = "", _
    Optional ByVal lngModuleNO As Long = 0) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:������ص�SQL����
    '���:
    '����:strOutSQL-������ص�SQL���
    '����:
    '����:���˺�
    '����:2008-08-22 17:26:47
    '-----------------------------------------------------------------------------------------------------------
        
    Dim str�ⷿ���� As String, str�������� As String, strվ������ As String
    '��������������Ʊ�����ݣ���ȡ�Է��ⷿ
    '-----------------����-----------------
    '���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (1"������Է��ⷿ",3"��˫����ͨ")
    '�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (2"���������ڿⷿ",3"��˫����ͨ")
    '-----------------����-----------------
    '���ڿⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� In (2"���������ڿⷿ",3"��˫����ͨ")
    '�Է��ⷿ�ǵ�ǰ�ⷿ�ģ���ȡ���� IN (1"������Է��ⷿ",3"��˫����ͨ")
    Dim bln���ϲ��� As Boolean  '��ʾ�ⷿΪ���ϲ���
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    strվ������ = GetDeptStationNode(lng�ⷿID)
    If lngModuleNO = 1716 Or lngModuleNO = 1717 Or lngModuleNO = 1722 Then
        str�ⷿ���� = "('V','W','K')"
    Else
        str�ⷿ���� = "('V','W','K','12')"
    End If
    
    str�������� = "" & _
    ",( Select �Է��ⷿID ID From �����������" & _
    "   Where ���ڿⷿID=[1] And ���� In (" & IIf(bln����, 1, 2) & ",3)" & _
    "   Union" & _
    "   Select ���ڿⷿID ID From �����������" & _
    "   Where �Է��ⷿID=[1] And ���� In (" & IIf(bln����, 2, 1) & ",3)) D "
    
    If bln���� Then
        'ȷ��ֻ�Ƿ��ϲ���
        gstrSQL = " Select a.ID From ���ű� a,��������˵�� B where A.id=B.����ID and A.ID=[1] and b.��������  in ('���Ŀ�','�Ƽ���')"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���Ƿ��ϲ��ŵĲ���", lng�ⷿID)
        
        bln���ϲ��� = rsTemp.RecordCount = 0
        rsTemp.Close
        If bln���ϲ��� Then
            '���ֻ�Ƿ��ϲ������ԵĲ��ţ���ֻ���Ƶ��ⷿ
            If lngModuleNO = 1716 Or lngModuleNO = 1722 Then    '�ƿ�����������
                gstrSQL = "" & _
                    "   Select DISTINCT a.id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
                    "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��" & _
                    "   From ���ű� a,��������˵�� B " & str�������� & _
                    "   where A.id=B.����ID and a.id=d.id and b.��������  in ('���Ŀ�','�Ƽ���') " & _
                    "           AND (a.����ʱ�� is null or a.����ʱ��>= to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                gstrSQL = "" & _
                    "   Select DISTINCT a.id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
                    "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��" & _
                    "   From ���ű� a,��������˵�� B " & str�������� & _
                    "   where A.id=B.����ID and a.id=d.id  and b.��������  in ('���Ŀ�','�Ƽ���','����ⷿ') " & _
                    IIf(strվ������ <> "", " and (a.վ�� = [2] or a.վ�� is null) ", "") & "" & _
                    "           AND (a.����ʱ�� is null or a.����ʱ��>= to_date('3000-01-01','yyyy-mm-dd'))"
            End If

            strOutSQL = gstrSQL
            gstrSQL = gstrSQL & " Order by a.����"
            Set ReturnSQL = zldatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ⷿID, strվ������)
            Exit Function
        End If
    End If

    '������ζԷ��ϲ��ŷ�����.
    gstrSQL = "" & _
    " SELECT DISTINCT a.id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
    "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��" & _
    " FROM ��������˵�� c, �������ʷ��� b, ���ű� a" & str�������� & _
    " Where c.�������� = b.���� AND b.���� in " & str�ⷿ���� & _
    "       AND a.id = c.����id And A.ID=D.ID" & _
    "       AND (a.����ʱ�� is null or a.����ʱ��>= to_date('3000-01-01','yyyy-mm-dd')) " & _
    IIf(strվ������ <> "" And lngModuleNO <> 1716 And lngModuleNO <> 1722, " and (a.վ�� = [2] or a.վ�� is null) ", "")
    
    strOutSQL = Replace(gstrSQL, "[1]", lng�ⷿID)
    
    gstrSQL = gstrSQL & " Order by a.����"
    Set ReturnSQL = zldatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ⷿID, strվ������)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetBillInfo(ByVal lng���� As Long, ByVal strNo As String, Optional ByVal bln�������� As Boolean = True, Optional ByVal bln��ҩ���� As Boolean = False) As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡ���ݵ�����޸�ʱ��
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select to_char(Max(" & IIf(bln��������, "��������", IIf(bln��ҩ����, "��ҩ����", "�������")) & "),'yyyyMMddhh24miss') ���� " & _
        "   From ҩƷ�շ���¼ " & _
        "   Where ����=[1] And NO=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ���ݵ�����޸�ʱ��", lng����, strNo)
    
    With rsTemp
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then Exit Function
        If IsNull(!����) Then Exit Function
        GetBillInfo = !����
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AutoAdd������(str���� As String, str���� As String, Optional strTittle As String = "����������", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�����������
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int���� As Integer, strCode As String, strSpecify As String
    
    AutoAdd������ = False
    If blnMsg = True Then
        If MsgBox("û���ҵ�������Ĳ��������̣���Ҫ���������������������", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    err = 0
    On Error GoTo ErrHand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ����������"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int���� = rsTemp!Length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM ����������"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlStr.GetCodeByVB(str����)
    
    
    gstrSQL = "ZL_����������_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    Call zldatabase.ExecuteProcedure(gstrSQL, strTittle)
    str���� = strCode
    AutoAdd������ = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function Where����ʱ��(Optional strAlias As String) As String
    If strAlias = "" Then
        Where����ʱ�� = " (����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or ����ʱ�� is null) "
    Else
        Where����ʱ�� = " (" & strAlias & ".����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or " & strAlias & ".����ʱ�� is null) "
    End If
End Function

'ȡʱ�������������ʱ���Ƿ��������Ӽ���
Public Function Get�Ӽ���() As Boolean
    Get�Ӽ��� = Val(zldatabase.GetPara(82, glngSys, 0)) = 1
End Function

Public Function Get���۵�λ() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:ȡָ�������۵Ķ��۵�λ
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
   '20050106���˺����˲���
    Get���۵�λ = Val(zldatabase.GetPara(88, glngSys, 0))
End Function

Public Function GetDigit() As Integer
    '��ȡ���С��λ��
    Dim intС�� As Integer
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    gstrSQL = "Select nvl(����,2) as ����  From ҩƷ���ľ��� Where ����=0 and ��� = 2 And ���� = 4 And ��λ = 5"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ����")
    If rsTemp.RecordCount = 0 Then
        GetDigit = 2
    Else
        GetDigit = rsTemp!����
    End If
    Exit Function
ErrHandle:
    GetDigit = 2
End Function

Public Function GetDigit����() As Integer
    '��ȡ���С��λ��
    Dim intС�� As Integer
    
    Dim rsTemp As New ADODB.Recordset

    On Error GoTo ErrHandle
    intС�� = Trim(zldatabase.GetPara(9, glngSys, 0))
    If intС�� = 0 Then intС�� = 2
    GetDigit���� = Val(intС��)
    Exit Function
ErrHandle:
    GetDigit���� = 2
End Function

Public Function IS��������() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Ƿ����ν�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    IS�������� = Val(zldatabase.GetPara(83, glngSys, 0)) = 1
End Function

Public Function IS�����ƿ�() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Ƿ����ν����ƿ�
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    IS�����ƿ� = Val(zldatabase.GetPara(280, glngSys, 0)) = 1
End Function

Public Function isʱ������ȡ�ϴ��ۼ�() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:�����⹺���ȡ�ϴ��ۼ�
    '���:
    '����:
    '����:����true-ȡ�ϴ��ۼ�,���򷵻�false-Ĭ�Ϸ�ʽ
    '------------------------------------------------------------------------------------------------------
    isʱ������ȡ�ϴ��ۼ� = Val(zldatabase.GetPara(229, glngSys, 0)) = 1
End Function

Public Function Get�ֶμӳ���(ByVal dbl���� As Double) As Double
    '����:��ȡ�ֶμӳ���
    '����:���ؼӳ���
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    gstrSQL = "Select �ӳ��� From ���ϼӳɷ���  " & _
             " Where  ([1] >��ͼ� And [1] <=��߼�)  " & _
             "        Or ([1] <=��߼� And nvl(��ͼ�,0)=0) " & _
             "        Or ([1] >��ͼ� And nvl(��߼�,0)=0)"
                 
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�ֶμӳ���", dbl����)
             
    If rsTemp.EOF Then
        ShowMsgBox "δ���ý���Ϊ:" & dbl���� & " �ļӳ��ʣ� " & vbCrLf & "��������Ŀ¼���������ã�����15%Ϊ�ӳ��ʼ���!"
        Get�ֶμӳ��� = 15
    Else
        Get�ֶμӳ��� = Val(zlStr.NVL(rsTemp!�ӳ���))
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ֶμӳ��ۼ�(ByVal dbl�ɹ��� As Double, ByVal dbl����ϵ�� As Double, ByVal strFormCaption As String, ByRef sng�ۼ� As Double) As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ȡ�ֶϼӳɺ���ۼ�
    '���:dbl�ɹ���-�ɹ���
    '     dbl����ϵ��-����ϵ��
    '     strFormCaption-��������
    '����:dbl�ۼ�-���صķֶϼӳɺ���ۼ�
    '����:�����������ȷ������true,���򷵻�false
    '�޸���:���˺�
    '�޸�ʱ��:2007/2/26
    '------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim byt���㷽�� As Byte
    Dim dbl�޼� As Double, dbl���޼� As Double
    Dim dblTemp As Double
    Dim dbl�ۼ�1 As Double
    Dim dbl�ɱ��� As Double
    
    err = 0: On Error GoTo ErrHand:
    dbl�ɱ��� = dbl�ɹ��� / IIf(dbl����ϵ�� = 0, 1, dbl����ϵ��)
    
    gstrSQL = "Select ���,��ͼ�,��߼�,1+�ӳ���/100 as �ӳ���,���㷽��,�޼� from ���ϼӳɷ��� order by ���"
    zldatabase.OpenRecordset rsTemp, gstrSQL, strFormCaption
    If rsTemp.EOF Then
        ShowMsgBox "δ���÷ֶϼӳ���,��������Ŀ¼����������!"
        Exit Function
    End If
    byt���㷽�� = Val(zlStr.NVL(rsTemp!���㷽��))
    dbl���޼� = Val(zlStr.NVL(rsTemp!�޼�))
    dbl�ۼ�1 = 0
    dblTemp = 0
    
    '2010-8�µ״����Ĵ�ʡҪ�����ӷֶ��޼ۿ��ơ������32282
    If rsTemp!��� = 0 Then rsTemp.MoveNext
    If byt���㷽�� = 0 Then
        '������㷨
        With rsTemp
            Do While Not .EOF
                If (dbl�ɱ��� > Val(zlStr.NVL(!��ͼ�)) And dbl�ɱ��� <= Val(zlStr.NVL(!��߼�))) Or _
                   (dbl�ɱ��� > Val(zlStr.NVL(!��ͼ�)) And Val(zlStr.NVL(!��߼�)) = 0) _
                Then
                    dbl�޼� = Val(zlStr.NVL(!�޼�))
                    dblTemp = Val(zlStr.NVL(!�ӳ���))
                    If dbl�޼� > 0 Then
                        If Round((dblTemp - 1) * dbl�ɱ���, 7) > Round(dbl�޼�, 7) Then
                            dbl�ۼ�1 = dbl�ɱ��� + dbl�޼�
                        Else
                            dbl�ۼ�1 = dblTemp * dbl�ɱ���
                        End If
                    Else
                        dbl�ۼ�1 = dblTemp * dbl�ɱ���
                    End If
                    GoTo Check�ۼ�:
                End If
                .MoveNext
            Loop
        End With
        ShowMsgBox "δ���ý���Ϊ:" & dbl�ɱ��� & " �ļӳ��ʣ� " & vbCrLf & "��������Ŀ¼����������!"
        Exit Function
    End If
    '�ֶϼ��㷨
    '�㶫���ļӳ��㷨:
    '    1000Ԫ���£�1000*10%��
    '    1000Ԫ���ϣ������Σ�1000*10%+(�ɹ���-1000)*8%
    With rsTemp
        Do While Not .EOF
            dbl�޼� = Val(zlStr.NVL(rsTemp!�޼�))
            If dbl�ɱ��� <= Val(zlStr.NVL(!��߼�)) Or Val(zlStr.NVL(!��߼�)) = 0 Then
                dblTemp = (dbl�ɱ��� - Val(zlStr.NVL(!��ͼ�))) * (Val(zlStr.NVL(!�ӳ���)) - 1)
                If dbl�޼� > 0 Then
                    If Round(dblTemp, 7) > Round(dbl�޼�, 7) Then
                        dbl�ۼ�1 = dbl�ɱ��� + dbl�ۼ�1 + dbl�޼�
                    Else
                        dbl�ۼ�1 = dbl�ɱ��� + dbl�ۼ�1 + dblTemp
                    End If
                Else
                    dbl�ۼ�1 = dbl�ɱ��� + dbl�ۼ�1 + dblTemp
                End If
                GoTo Check�ۼ�:
            ElseIf dbl�ɱ��� > Val(zlStr.NVL(!��߼�)) Then
                dblTemp = (Val(zlStr.NVL(!��߼�)) - Val(zlStr.NVL(!��ͼ�))) * (Val(zlStr.NVL(!�ӳ���)) - 1)
                If dbl�޼� > 0 Then
                    If Round(dblTemp, 7) > Round(dbl�޼�, 7) Then
                        dbl�ۼ�1 = dbl�ۼ�1 + dbl�޼�
                    Else
                        dbl�ۼ�1 = dbl�ۼ�1 + dblTemp
                    End If
                Else
                    dbl�ۼ�1 = dbl�ۼ�1 + dblTemp
                End If
            End If
            .MoveNext
        Loop
    End With
    ShowMsgBox "δ���ý���Ϊ:" & dbl�ɱ��� & " �ļӳ��ʣ� " & vbCrLf & "��������Ŀ¼����������!"
    Exit Function
Check�ۼ�:
    If Round(dbl�ۼ�1 - dbl�ɱ���, 7) > Round(dbl���޼�, 7) Then
        ShowMsgBox "�ӳɣ���" & Format(dbl�ۼ�1 - dbl�ɱ���, "###0.0000000;-###0.0000000;0;0") _
                 & ")��������޼�(��" & Format(dbl���޼�, "###0.0000000;-###0.0000000;0;0") _
                 & ")����Ĭ������޼ۼӳɣ�"
        'Exit Function
        dbl�ۼ�1 = dbl�ɱ��� + dbl���޼�
    End If
    sng�ۼ� = dbl�ۼ�1 * dbl����ϵ��
    
    Get�ֶμӳ��ۼ� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IS�ֶμӳ���() As Boolean
    '����:����Ƿ�ӳ����Էֶν������ȡ
    IS�ֶμӳ��� = Val(zldatabase.GetPara(121, glngSys, 0)) = 1
End Function

Public Function isʱ������ֱ��ȷ���ۼ�() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ȡ�Ƿ���isʱ������ֱ��ȷ���ۼ۵ķ�ʽ���
    '���:
    '����:
    '����:����ֱ��ȷ���ۼ۵ķ�ʽ,����true,���򷵻�false
    '�޸���:���˺�
    '�޸�ʱ��:2007/1/25
    '------------------------------------------------------------------------------------------------------
    isʱ������ֱ��ȷ���ۼ� = Val(zldatabase.GetPara(136, glngSys, 0)) = 1
End Function

Public Sub ��ʼС��λ��()
    '------------------------------------------------------------------------------------------------------
    '����:��ʼС��λ��
    '���:
    '����:
    '����:
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    
    '    ���    Number(1)   1-ҩƷ,2-����
    '    ����    Number(1)   1-�ɱ��ۣ�2-���ۼ�,3-����
    '    ��λ    Number(1)   1,2,3,4��ҩƷ�ֱ�Ϊ�ۼۡ����סԺ��ҩ�ⵥλ�����ķֱ�Ϊɢװ����װ��λ��
    '    ����    Number(1)   ȡֵΪ2-4��
    On Error GoTo ErrHandle
    strSql = "Select * from ҩƷ���ľ��� where ���=[1] and ����=0 order by ��λ"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "��ȡ�������ϵ�С��λ������", 2)
    With g_С��λ��
        With .obj_��װС��
            .�ɱ���С�� = 7
            .���ۼ�С�� = 7
            If glngModul = 1723 Then '���ķ���ȡ���ý��ȣ�����ȡҩƷ���������õľ���
                .���С�� = GetDigit����
            Else
                .���С�� = GetDigit
            End If
            .����С�� = 3
        End With
        With .obj_ɢװС��
            .�ɱ���С�� = 7
            .���ۼ�С�� = 7
            If glngModul = 1723 Then '���ķ���ȡ���ý��ȣ�����ȡҩƷ���������õľ���
                .���С�� = GetDigit����
            Else
                .���С�� = GetDigit
            End If
            .����С�� = 3
        End With
        With .obj_���С��
            .�ɱ���С�� = 7
            .���С�� = 5
            .����С�� = 5
            .���ۼ�С�� = 7
        End With
    End With
    
    With gOraFmt_Max
        .FM_��� = GetFmtString(-1, gС������.g_���, True)
        .FM_�ɱ��� = GetFmtString(-1, gС������.g_�ɱ���, True)
        .FM_���ۼ� = GetFmtString(-1, gС������.g_�ۼ�, True)
        .FM_���� = GetFmtString(-1, gС������.g_����, True)
        .FM_ɢװ���ۼ� = GetFmtString(-1, gС������.g_�ۼ�, True)
    End With
    
    If rsTemp.EOF Then Exit Sub
    Do While Not rsTemp.EOF
        If Val(zlStr.NVL(rsTemp!��λ)) = 2 Then
            '��װ��λ
            If Val(zlStr.NVL(rsTemp!����)) = 1 Then
                g_С��λ��.obj_��װС��.�ɱ���С�� = Val(zlStr.NVL(rsTemp!����))
            ElseIf Val(zlStr.NVL(rsTemp!����)) = 2 Then
                g_С��λ��.obj_��װС��.���ۼ�С�� = Val(zlStr.NVL(rsTemp!����))
            ElseIf Val(zlStr.NVL(rsTemp!����)) = 3 Then
                g_С��λ��.obj_��װС��.����С�� = Val(zlStr.NVL(rsTemp!����))
            End If
        ElseIf Val(zlStr.NVL(rsTemp!��λ)) = 1 Then
            'ɢװ��λ
            If Val(zlStr.NVL(rsTemp!����)) = 1 Then
                g_С��λ��.obj_ɢװС��.�ɱ���С�� = Val(zlStr.NVL(rsTemp!����))
            ElseIf Val(zlStr.NVL(rsTemp!����)) = 3 Then
                g_С��λ��.obj_ɢװС��.����С�� = Val(zlStr.NVL(rsTemp!����))
            Else
                g_С��λ��.obj_ɢװС��.���ۼ�С�� = Val(zlStr.NVL(rsTemp!����))
            End If
        ElseIf Val(zlStr.NVL(rsTemp!��λ)) = 5 Then
            '���
            If glngModul <> 1723 Then '���ķ��ϵĻ� �Ͳ����øþ��ȣ�����ֱ���÷���ҵ�����õľ���
                g_С��λ��.obj_��װС��.���С�� = Val(zlStr.NVL(rsTemp!����))
                g_С��λ��.obj_ɢװС��.���С�� = Val(zlStr.NVL(rsTemp!����))
            End If
        End If
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFmtString(ByVal int��λ As Integer, ByVal С������ As gС������, _
    Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '����:����ָ����С����ʽ��
    '���:int��λ-0-ɢװ��λ,1-��װ��λ,<>0 or 1:���ݿ����С��λ
    '     lngС��λ��-С��λ��
    '     blnOracle-������oracle�ĸ�ʽ������Vb�ĸ�ʽ��
    '����:
    '����:����ָ���ĸ�ʽ��
    '�޸���:���˺�
    '�޸�ʱ��:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim strFmt As String
    Dim intλ�� As Integer
    Select Case С������
    Case g_����
        Select Case int��λ
        Case 0  ' ɢװ��λ
            intλ�� = g_С��λ��.obj_ɢװС��.����С��
        Case 1  '��װ��λ
            intλ�� = g_С��λ��.obj_��װС��.����С��
        Case Else  '-1������ݿⵥλ
            intλ�� = g_С��λ��.obj_���С��.����С��
        End Select
    Case g_���
        Select Case int��λ
        Case 0  ' ɢװ��λ
            intλ�� = g_С��λ��.obj_ɢװС��.���С��
        Case 1  '��װ��λ
            intλ�� = g_С��λ��.obj_��װС��.���С��
        Case Else  '-1������ݿⵥλ
            intλ�� = g_С��λ��.obj_���С��.���С��
        End Select
    Case g_�ɱ���
        Select Case int��λ
        Case 0  ' ɢװ��λ
            intλ�� = g_С��λ��.obj_ɢװС��.�ɱ���С��
        Case 1  '��װ��λ
            intλ�� = g_С��λ��.obj_��װС��.�ɱ���С��
        Case Else  '-1������ݿⵥλ
            intλ�� = g_С��λ��.obj_���С��.�ɱ���С��
        End Select
    Case g_�ۼ�
        Select Case int��λ
        Case 0  ' ɢװ��λ
            intλ�� = g_С��λ��.obj_ɢװС��.���ۼ�С��
        Case 1  '��װ��λ
            intλ�� = g_С��λ��.obj_��װС��.���ۼ�С��
        Case Else  '-1������ݿⵥλ
            intλ�� = g_С��λ��.obj_���С��.���ۼ�С��
        End Select
    Case Else
        intλ�� = 0
    End Select
    If blnOracle Then
       GetFmtString = "'9999999999990." & String(intλ��, "9") & "'"
    Else
       GetFmtString = "#0." & String(intλ��, "0")
    End If
End Function

Public Function InitSystemPara() As Boolean
    '------------------------------------------------------------------------------------------------------
    '����:��ʼ����ص�ϵͳ����
    '���:
    '����:
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '�޸���:���˺�
    '�޸�ʱ��:2007/6/28
    '------------------------------------------------------------------------------------------------------
    Dim strValue As String
    With gSystem_Para
        '0-ƴ����,1-�����,2-����
        .int���뷽ʽ = Val(zldatabase.GetPara("���뷽ʽ"))
        '��1λ1-ȫ����ֻ�����,��2λ1-ȫ��ĸֻ�����,��HIS��������������
        .Para_���뷽ʽ = zldatabase.GetPara(44, glngSys, 0): .Para_���뷽ʽ = IIf(.Para_���뷽ʽ = "", "11", .Para_���뷽ʽ)
        .para_������¿��ÿ�� = Val(zldatabase.GetPara(95, glngSys, 0)) = 1
        '������ʾ��ʽ
        .bln���￨������ʾ = Val(zldatabase.GetPara(12, glngSys)) = 1
        .str���￨ǰ׺�� = zldatabase.GetPara(27, glngSys)
        .P156_�����㷨 = zldatabase.GetPara(156, glngSys, 0, 0)
     End With
    'վ��������
    Call Initվ����Ϣ
    InitSystemPara = True
End Function
 
Public Function Check��Ժ����(ByVal strPrivs As String, ByVal lng���� As Long, ByVal strNo As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer, Optional ByVal lng����id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ժ�����Ƿ�������,��Ҫ����Ȩ�޿���(���û��Ȩ�ޡ����˳�Ժ���˴����������������ϲ���)
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 15:05:49
    '-----------------------------------------------------------------------------------------------------------

    '����˵���������ǰ������סԺ���ˣ�
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng��ҳID As Long
    
    On Error GoTo ErrHandle
    If lng���� = 24 Then
        Check��Ժ���� = True
        Exit Function
    End If
    
    '���δ���벡��ID�����Զ���ȡ
    gstrSQL = "Select A.����ID,c.��ҳid From ������ü�¼ A, ҩƷ�շ���¼ B,����ҽ����¼ C Where A.ID = B.����ID  And A.ҽ�����=C.id And b.���� = [1] And b.No = [2] And Rownum = 1 "
    
    If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
    End If
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ����ID", lng����, strNo)
    
    '����������Ҳ�������ID�򲻽�����һ�����
    If rsTemp.EOF Then
        Check��Ժ���� = True
        Exit Function
    End If
    
    lng����id = rsTemp!����ID
    lng��ҳID = NVL(rsTemp!��ҳid, 0)

    'ȡ��������
    gstrSQL = "Select ���� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng����id)

    str���� = rsTemp!����
    
    '�����ǰ������סԺ���ˣ����û��Ȩ�ޡ����˳�Ժ���˴���������������ҩ����
    If zlStr.IsHavePrivs(strPrivs, "���˳�Ժ���˴���") = False Then
        '��鲡����Ԥ��Ժ���Ժ
        gstrSQL = " Select 1 From ������ҳ" & _
                  " Where ����ID=[1] and ��ҳid=[2] " & _
                  " And (��Ժ���� Is Not NULL)"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѳ�Ժ", lng����id, lng��ҳID)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "�ڴ���[" & strNo & "]�У����ˡ�" & str���� & "���ѳ�Ժ����û�ж��ѳ�Ժ���˵Ĵ������з��ϡ����ϵ�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check��Ժ���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Check���ʴ���(ByVal strPrivs As String, ByVal lng���� As Long, ByVal strNo As String, ByVal str��� As String, ByVal int��¼���� As Integer, ByVal int�����־ As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��鴦���Ƿ��Ѿ�������,���ʵĴ������ܷ����ϲ���
    '���:  lng����    ����ǰ��������
    '       strNO      ����ǰ���ݺ�
    '       lng����ID  �����Զಡ�˵���Ч
    '       str��ţ���ص������,��,����
    '����:
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 14:58:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If lng���� = 24 Then
        Check���ʴ��� = True
        Exit Function
    End If
    
    '���û��Ȩ�ޡ����˽��ʴ����������ô����Ƿ��ѽ��ʣ��ѽ��ʴ������������ϲ���
    If zlStr.IsHavePrivs(strPrivs, "���˽��ʴ���") = 0 Then
    
        gstrSQL = "Select Nvl(Sum(Nvl(���ʽ��,0)),0) AS ���ʽ��   " & _
                 "  From ������ü�¼   " & _
                 "  Where Instr([1], ',' || ��� || ',') > 0 " & _
                 "  And Mod(��¼����,10) = 2 and NO = [2]"
        If int��¼���� = 1 Or (int��¼���� = 2 And (int�����־ = 1 Or int�����־ = 4)) Then
        Else
            gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End If
        gstrSQL = gstrSQL & " Order By ���ʽ�� Desc"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ��ѽ���", "," & str��� & ",", strNo)
        If zlStr.NVL(rsTemp!���ʽ��, 0) <> 0 Then
            MsgBox "�ô���[" & strNo & "]�ѽ��ʣ���û�ж��ѽ��ʴ������з��ϡ����ϵ�Ȩ�ޣ�������ֹ��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    Check���ʴ��� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '����:�жϿؼ��Ƿ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo ErrHand:
    
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub Initվ����Ϣ()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��վ��������Ϣ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-09-01 11:32:00
    '-----------------------------------------------------------------------------------------------------------
    gSystem_Para.bln����վ�� = gstrNodeNo <> "-"
 End Sub
 
Public Function zl_��ȡվ������(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str���� As String = "") As String
    '����:��ȡվ����������:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str���� = "", "", str���� & ".") & "վ��"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
     zl_��ȡվ������ = strWhere
End Function

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
    '���ܣ��ж�ָ���ı����е�ǰ�����Ƿ���ˢ��,���ݴ���������ʾ
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long
    
    '��ǰ�������ʾ������(��δ��ʾ����)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
    '�ж��Ƿ���ˢ��
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf gSystem_Para.str���￨ǰ׺�� <> "" Then
        arrMask = Split(gSystem_Para.str���￨ǰ׺��, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    'ˢ��ʱ�����Ƿ�������ʾ
    If blnCard Then
       txtInput.PasswordChar = IIf(gSystem_Para.bln���￨������ʾ = False, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    InputIsCard = blnCard
End Function

Public Sub CheckKeyPress���￨��(ByVal txtPati As TextBox, KeyAscii As Integer)
    '�����￨�ŵ�����Ϸ���
    
    If InStr(1, ":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Public Function GetDeptStationNode(ByVal lngDeptId As Long) As String
'��ȡ��������վ����Ϣ
    Dim rsSQL As ADODB.Recordset
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    strTmp = "select վ�� from ���ű� where id=[1]"
    Set rsSQL = zldatabase.OpenSQLRecord(strTmp, "��ȡ��������վ����Ϣ", lngDeptId)
    If Not rsSQL.EOF Then
        GetDeptStationNode = zlStr.NVL(rsSQL!վ��)
    End If
    rsSQL.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '���ݴ�����ַ������зֽ⣬����ָ���ַ����Ⱦ���Ҫ���зֽ⣬������浽������
    '��Σ�strInput-������ַ�����strSplitChar-�ַ��������ݵķָ���
    '���أ����飬���������Ա���ַ����Ȳ�����ָ������
    Dim strArray As Variant
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '����ָ���ַ�ʱ����Ҫ�ֽ�
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '�޷ָ���ʱ
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '�зָ���ʱ
            arrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(arrTmp)
        
            For i = 0 To lngCount
                If arrTmp(i) <> "" Then
                    '�зָ�������Ҫ���ַָ���֮���ַ��������ԣ����ܰѷָ���֮����ַ���
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = arrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    GetArrayByStr = strArray
End Function
