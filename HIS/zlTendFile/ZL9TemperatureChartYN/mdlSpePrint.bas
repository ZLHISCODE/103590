Attribute VB_Name = "mdlSpePrint"
Option Explicit

Private mlng���²�����ʾ��ʽ As Long
Private mintBaby As Integer  '�Ƿ���Ӥ��
Private mlngBreatheHeight  '�������̶��߶�

Public Function PrintOrPreviewBodyStateNew(objOut As Object, _
                                        ByVal lng����ID As Long, _
                                        ByVal lng��ҳID As Long, _
                                        ByVal lng�ļ�ID As Long, _
                                        ByVal intBaby As Integer, _
                                        ByVal lngSectID As Long, _
                                        ByVal lngBeginY As Long, _
                                        ByVal lngBeginX As Long, _
                                        ByVal objParent As Object, _
                                        Optional ByVal blnKeepOn As Boolean = False, _
                                        Optional ByVal intBeginPage As Integer = -1, _
                                        Optional ByVal intEndPage As Integer = -1, _
                                        Optional ByVal intPageNo As Integer = -1, _
                                        Optional ByVal sngScale As Single = 1, _
                                        Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '����:��ӡ��Ԥ��ĳ�����ר���¶ȱ�
    '����:objOut=�������,����ΪPrinter��һ������(�����а����ؼ�����picPage)
    '      lngCaseRecordID=������¼id
    '      lngBeginY=��ʼ������
    '      blnKeepOn=�Ƿ񱣳�����
    '      objParent=�����ô���
    '      intBeginPage=Ҫ��ʼҳ�����,��Ϊ-1ʱ��ʾ�������.
    '      intEndPage=����ҳ������intEndPage����ʵ��ҳ����ֻ��ӡ��ʵ��ҳ��
    '      intPageNO=��ʼ��ҳ��,���Ϊ-1��ʾ����ʾҳ��
    '      sngScale=�������
    
    '����:���δ�ӡ�����Ƿ�ɹ�
    '******************************************************************************************************************
    Dim strSQL As String, strNewSql As String
    '�������ò�������
    Dim intOpDays As Integer  '�������ע����
    Dim blnStopFlag As Boolean '�ٴ�����ֹͣǰ�α�ע
    Dim intOpFormat As Integer '��������ȱʡ��ʽ
    Dim bytδ����ʾλ�� As Byte 'δ��˵����ʾλ��
    Dim blnӤ�����µ���ʾ��Ժ As Boolean 'Ӥ�����µ���ʾ��Ժ��Ϣ
    Dim bln���µ���ʾ��� As Boolean '���µ���ʾ���
    Dim intRepairRows As Integer  '�����ʾ����
    Dim bln��ʾƤ�� As Boolean '���µ������ʾƤ�Խ��
    Dim bln��ӡҽԺ���� As Boolean '���µ��Ƿ��ӡҽԺ����
    Dim bln�����ʾ��Ժ As Boolean
    Dim bln���� As Boolean
    Dim bln���ܵ��� As Boolean '���µ���������ʱ������ܽ��컹�ǽ���������������
    Dim bln¼��Сʱ As Boolean '���µ�ȫ���������¼�����ʾ����Сʱ��
    Dim bln��ӡ������� As Boolean  '���µ���ӡʱ�Ƿ��ӡ�������
    Dim bln����ӡ������ As Boolean '���µ���ӡʱ�Ƿ��ӡ������(�������ʵ���Ӧ������Ч��ֻ�ǲ���ӡ�̶��У����������)
    Dim lngCurveRow As Long '�������߹̶��������
    Dim bln��Ժ As Boolean
    Dim lngSignColor As Long '����:���µ������ʾ��ɫ
    
    '������ͼ����
    Dim i As Integer, j As Integer
    Dim lngPicPageIndex As Integer 'Ԥ��ʱPIC������
    Dim blnPrint As Boolean  '�Ƿ��ӡ
    Dim strInfo As String '˵����Ϣ
    Dim intAllOpt As Single  '��ӡ���ܹ�����
    Dim intCurOpt As Single  '��ӡ���е��ڼ���
    Dim objDraw As Object '��ͼ����
    Dim lngHwnd As Long '���
    Dim lngDC As Long  '��ͼ�����DC
    Dim lngFont As Long
    Dim lngOldFont As Long
    Dim stdSet As StdFont
    Dim lngLableStep As Long '�̶������п�
    Dim lngColStep As Long ' ���������п�
    Dim lngInitRowStep As Long '���������и�
    
    Dim lngCountPage As Long '����ҳ��
    Dim lngPage As Long
    Dim strBeginDate  As String, strBeginDate1 As String '��ʼʱ��
    Dim strEndDate As String '��ֹʱ��
    Dim strTmpDay As String, strEndDay As String
    Dim dtBegin As Date, dtEnd As Date
    Dim intDrawLineRows As Integer '���µ�����������
    Dim intDrawLineCOL As Integer '���µ��̶���������
    Dim strTmp As String, strTime As String, strTmp1 As String
    Dim lngValue As Long 'סԺ����
    Dim T_Rect As RECT
    Dim rsPart As New ADODB.Recordset  '�������²�λ��Ϣ
    Dim rsTemp As New ADODB.Recordset  '�˼�¼���벻Ҫ˳��ʹ��
    Dim rsTmp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset 'ʹ����˲��˵����л�����Ŀ��Ϣ
    Dim rsDrawItems As New ADODB.Recordset '���µ�������Ŀ��Ϣ
    Dim rsPoints As New ADODB.Recordset '�������µ��ļ���
    Dim rsNotes As New ADODB.Recordset   '����˵����Ϣ
    Dim rsDownTab As New ADODB.Recordset '���±��������Ϣ
    Dim H_16pt As Long, W_16pt As Long
    Dim int����Ӧ�� As Integer
    Dim str���ʷ���  As String
    Dim arrTmpValue() As Variant, arrTmpNote() As Variant
    Dim arrValues() As String
    Dim strPart As String '��λ
    Dim SinX As Single, sinY As Single
    Dim intCOl As Integer
    Dim blnAdd As Boolean, blnAllow As Boolean
    Dim dbl��ֵ As Double, dblMinValue As Double, dblMaxValue As Double
    Dim lng��Ŀ��� As Long
    Dim str����˵�� As String
    Dim bln���� As Boolean  '�����Ƿ�Ϊ���
    Dim sngHTab As Single  '���±��߶�
    Dim sngHPrint As Single '�ɴ�ӡ����
    
    Dim strBegin As String, strEnd As String
    Dim str��� As String
    Dim strItemName As String, strItems As String
    Dim intƵ�� As Integer
    Dim intCol1 As Integer
    Dim str��Ŀ���� As String
    Dim int��Ŀ���� As Integer, int��Ŀ���� As Integer, int��Ժ�ײ� As Integer
    Dim int����ѹ As Integer, int����ѹ As Integer, Int�к� As Integer
    Dim blnColor As Boolean

    '���˻�����Ϣ
    Dim strPatiInfo As String
    Dim VarPatiInfo As Variant
    Dim lng����ȼ� As Long
    
    '--������������ �ڼ�¼���²���ʱ����ʱ�������
    Dim strTmpString0 As String  '��¼��ǰʱ��
    Dim strTmpString2 As String '��¼סԺ����
    Dim strTmpString1 As String '��¼����������
    Dim strNewTmpString As String
    Dim ArrNewTmpString() As String '��¼�����Ŀ��������ÿһ��ֵ����Ϣ
    Dim ArrNewString() As String '��¼���б����Ŀ��Ϣ
    Dim intDays As String '��������
    Dim strOpdays() As String
    Dim strOpValue() As String
    Dim arrOperDay
    Dim strEditors() As Variant    '��¼������Ŀ��Ϣ(��Ŀ���||��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼��||��¼ɫ||���ֵ||��Сֵ||�ٽ�ֵ��
    Dim ArrComTable() As Variant '��¼���еı��±����Ŀ (��Ŀ���||��λ+��Ŀ����|��Ŀ��λ||��Ŀֵ��||��¼Ƶ��||��Ŀ����||��Ŀ��ʾ||��Ժ�ײ�)
    Dim lng���� As Long  '��¼��������
    Dim bln������ʾ As Boolean '����������14���Ժ�����ʾ
    Dim str����ʱ�� As String
    Dim bln��Ʋ�ת��Ժ As Boolean
    
    '������Ϣ
    Dim lngLeft As Long, lngTop As Long
    Dim lngRight As Long, lngButtom As Long
    Dim X As Long, Y As Long
    Dim lngCurX As Long, lngCurY As Long
    Dim dblSureW As Double, dblSureH As Double
    
    Dim M_DrawClient As DrawClient
    Dim lng�̶ȿ�� As Long
    
    On Error GoTo ErrPrint
    
    msngTwips = 1
    
    mintBaby = intBaby
    '����ԭʼֵ:
    
    M_DrawClient.ƫ����X = T_DrawClient.ƫ����X
    M_DrawClient.ƫ����Y = T_DrawClient.ƫ����Y
    M_DrawClient.�̶����� = T_DrawClient.�̶�����
    M_DrawClient.�̶ȵ�λ = T_DrawClient.�̶ȵ�λ
    M_DrawClient.�������� = T_DrawClient.��������
    M_DrawClient.�е�λ = T_DrawClient.�е�λ
    M_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ
    M_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ
    M_DrawClient.�е�λ = T_DrawClient.�е�λ
    M_DrawClient.˫�� = T_DrawClient.˫��
    M_DrawClient.������ = T_DrawClient.������
    M_DrawClient.���������� = T_DrawClient.����������
    M_DrawClient.�������������� = T_DrawClient.��������������
    lng�̶ȿ�� = T_BodyStyle.lng�̶ȿ��
    
    mintBmpW = gintBmpW
    mintBmpH = gintBmpH
    '��ȡ���²�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    intOpDays = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10"))
    blnStopFlag = (Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
    bytδ����ʾλ�� = Abs(Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0")))
    blnӤ�����µ���ʾ��Ժ = (zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    bln���µ���ʾ��� = (zlDatabase.GetPara("���µ���ʾ���", glngSys, 1255, 1) = 1)
    intRepairRows = T_BodyStyle.lng������ + GetRows(bln����, T_BodyItem.str�����Ŀ)
    bln��ʾƤ�� = (Val(zlDatabase.GetPara("���µ���ʾƤ�Խ��", glngSys, 1255, "0")) = 1)
    bln��ӡҽԺ���� = (Val(zlDatabase.GetPara("��ӡҽԺ����", glngSys, 1255, "1")) = 1)
    bln���ܵ��� = (Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    bln��ӡ������� = (Val(zlDatabase.GetPara("����ӡ�������ͼ��", glngSys, 1255, "0")) = 0)
    bln����ӡ������ = (Val(zlDatabase.GetPara("���µ�����ӡ������", glngSys, 1255, "0")) = 1)
    mlng���²�����ʾ��ʽ = Val(zlDatabase.GetPara("���²�����ʾ��ʽ", glngSys, 1255, "0"))
    bln������ʾ = (Val(zlDatabase.GetPara("����������14���Ժ�����ʾ", glngSys, 1255, "0")) = 1)
    bln��Ʋ�ת��Ժ = (Val(zlDatabase.GetPara("��Ʊ�ʶ���Զ�ת��Ϊ��Ժ", glngSys, 1255, "0")) = 0)
    '62989:������,2013-07-24,���µ������ʾ��ɫ
    lngSignColor = Val(zlDatabase.GetPara("���µ������ʾ��ɫ", glngSys, 1255, "255"))
    
    lngCurveRow = T_BodyStyle.lng���߿���
    
    '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
    bln¼��Сʱ = (Val(zlDatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    
    '51338,������,2012-07-06
    strTmp = zlDatabase.GetPara("��������ȱʡ��ʽ", glngSys, 1255, "2")
    If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
        intOpFormat = Val(strTmp)
    Else
        intOpFormat = 0
    End If
    '���˱䶯�����ʾ����
    '------------------------------------------------------------------------------------------------------------------
    Call InitPara(T_BodyStyle.blnר��)

    blnPrint = TypeName(objOut) = "Printer"
    
    '���ڴ�ӡ������Ļ�����ز�ͬ���˴���Ҫȡ���Ե�����
    If blnPrint = True Then
        T_TwipsPerPixel.X = Printer.TwipsPerPixelX
        T_TwipsPerPixel.Y = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        Printer.Font.Size = 9
        Printer.FontName = "����"
    Else
        T_TwipsPerPixel.X = Screen.TwipsPerPixelX
        T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
        msngTwips = 1
    End If
    
    mlngBreatheHeight = 300 \ T_TwipsPerPixel.Y
    Screen.MousePointer = 11
    intAllOpt = 5
    
    '������ȴ���
    '------------------------------------------------------------------------------------------------------------------
    strInfo = "����" & IIf(blnPrint, "׼����ӡ���±�", "����Ԥ��") & ",���Ժ�..."
    Call ShowFlash(strInfo, , objParent)
    
    '��ӡǰ�����
    If blnKeepOn = False Then
        If Not blnPrint Then
            For i = objOut.picPage.UBound To 0 Step -1
                If i = 0 Then
                    objOut.picPage(i).Cls
                Else
                    Unload objOut.picPage(i)
                End If
            Next
            Set objDraw = objOut.picPage(0)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    Else
        If Not blnPrint Then
            i = objOut.picPage.UBound + 1
            Load objOut.picPage(i)
            Set objDraw = objOut.picPage(objOut.picPage.UBound)
            objDraw.Width = Printer.Width * sngScale
            objDraw.Height = Printer.Height * sngScale
        Else
            Set objDraw = Printer
        End If
    End If
    
    bln��Ժ = False
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ)����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strSQL = getSQLString("��ȡ�ļ�ʱ�䷶Χ", blnMoved)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        lngCountPage = DateDiff("d", rsTemp!��ʼ, rsTemp!��ֹ) + 1
        lngCountPage = IIf(lngCountPage / T_BodyStyle.lng���� = Fix(lngCountPage / T_BodyStyle.lng����), lngCountPage / T_BodyStyle.lng����, Fix(lngCountPage / T_BodyStyle.lng����) + 1)
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strBeginDate1 = strBeginDate
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        bln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        CloseRs rsTemp
        GoTo ErrPrint '�������˱䶯��Ϣ�˳�
    End If
    
    gbln��Ժ = bln��Ժ
    If bln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
        strEndDate = Format(RetrunEndTimeNew(CDate(strBeginDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
    End If
    
    bln�����ʾ��Ժ = False
    
    If CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) > CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln�����ʾ��Ժ = True
    ElseIf T_BodyFlag.��Ժ = 0 And CDate(Format(strBeginDate, "YYYY-MM-DD HH:MM:SS")) = CDate(Format(strBeginDate1, "YYYY-MM-DD HH:MM:SS")) Then
        bln�����ʾ��Ժ = True
    End If
            
    intCurOpt = intCurOpt + 1
    
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '------------------------------------------------------------------------------------------------------------------
    '��1���ݣ����˵Ļ�����Ϣ
    '��ȡ���˻�����Ϣ
    
    '"����'����'�Ա�'�Ʊ�'����'��Ժ����'סԺ��:
    strPatiInfo = "''''''"
    VarPatiInfo = Split(strPatiInfo, "'")
    
    strSQL = " Select  NVL(A.����,b.����) ����,A.סԺ��,A.��Ժ���� ��Ժʱ��,NVL(A.�Ա�,b.�Ա�) �Ա�,NVL(A.����,B.����) ���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then
        VarPatiInfo(0) = zlCommFun.Nvl(rsTemp("����").Value)
        VarPatiInfo(6) = zlCommFun.Nvl(rsTemp("סԺ��").Value)
        VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("��Ժʱ��").Value), "yyyy-MM-dd")
        VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("�Ա�").Value)
        VarPatiInfo(1) = zlCommFun.Nvl(rsTemp("����").Value)
    End If
    
    '��Ժʱ��(������µ���ʼʱ�������Ժʱ��������ʱ��Ϊ׼)
    strSQL = "select ��ʼʱ�� from ���˱䶯��¼ where ����id=[1] And ��ҳid=[2] and ��ʼԭ��=2 order by ��ʼʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then
        If bln�����ʾ��Ժ = True Then
            VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("��ʼʱ��").Value), "yyyy-MM-dd")
        End If
    End If
    
    If intBaby <> 0 Then
        
        VarPatiInfo(1) = ""
        VarPatiInfo(2) = ""
        
        strSQL = "Select Decode(a.Ӥ������,Null,NVL(C.����,B.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,a.Ӥ���Ա�,a.����ʱ�� " & _
            " From ������Ϣ B,������ҳ C,������������¼ A " & _
            " Where B.����ID=C.����ID And C.����ID=A.����ID And C.��ҳID=A.��ҳID And C.����id=[1] And C.��ҳid=[2] And a.���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID, intBaby)
        If rsTemp.BOF = False Then
            VarPatiInfo(0) = rsTemp("Ӥ������").Value
            VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("Ӥ���Ա�").Value)
            VarPatiInfo(1) = "������"
            If IsNull(rsTemp("����ʱ��").Value) = False Then VarPatiInfo(5) = Format(zlCommFun.Nvl(rsTemp("����ʱ��").Value), "yyyy-MM-dd")
        End If
        
    End If
    
    If bln���µ���ʾ��� Then ReDim Preserve VarPatiInfo(UBound(VarPatiInfo) + 1)
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '��ȡ���˻���ȼ�
     lng����ȼ� = Get����ȼ�(lng����ID, lng��ҳID)
    
    '��ȡ���ü�¼��
    Call InitPublicData
    
    '�������Ӧ�÷�ʽ
    int����Ӧ�� = 2
    str���ʷ��� = ""
    strSQL = "Select a.Ӧ�÷�ʽ,b.��¼�� From �����¼��Ŀ a,���¼�¼��Ŀ b Where a.��Ŀ���=-1 And a.��Ŀ���=b.��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint")
    If rsTemp.BOF = False Then
        int����Ӧ�� = zlCommFun.Nvl(rsTemp("Ӧ�÷�ʽ").Value, 2)
        str���ʷ��� = zlCommFun.Nvl(rsTemp("��¼��").Value, "��")
    Else
        int����Ӧ�� = 0
    End If
    
    Dim int���� As Integer, int���� As Integer
    
    '-------------------------------------------------------------------------------------------------------------------
    '2��ȡ����������Ŀ(�����µ��̶�������������������-2)
    strSQL = getSQLString("��ȡ����������Ŀ", blnMoved)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����������Ŀ", T_BodyItem.str������Ŀ)
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        rsTemp.Filter = "��¼��=1"
        intDrawLineCOL = rsTemp.RecordCount
        rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
        If rsTemp.RecordCount > 0 And bln����ӡ������ Then
            rsTemp.Filter = 0
            intDrawLineCOL = intDrawLineCOL - 1
        Else
            rsTemp.Filter = 0
        End If
        If intDrawLineCOL <= 0 Then intDrawLineCOL = 1
    Else
        CloseRs rsTemp
        MsgBox "���κ�����������Ŀ��", vbExclamation, gstrSysName
        GoTo ErrExit
    End If
    strEditors = Array()
    int���� = -1: int���� = -1
    rsTemp.Filter = 0
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            strTmp = Nvl(!��Ŀ���, 0) & "|| " & Nvl(!��¼��) & "|| " & Nvl(!��λ) & "|| " & Nvl(!��Ŀֵ��) & "|| " & _
                 Nvl(!��¼��) & "|| " & Nvl(!��¼ɫ) & "||" & Nvl(!���ֵ) & "||" & Nvl(!��Сֵ) & "||" & Nvl(!�ٽ�ֵ)
                
            ReDim Preserve strEditors(UBound(strEditors) + 1)
            strEditors(UBound(strEditors)) = strTmp
            If zlCommFun.Nvl(!��Ŀ���, 0) = gint���� Then
                int���� = UBound(strEditors)
            End If
        .MoveNext
        Loop
        .MoveFirst
    End With
    If int����Ӧ�� = 2 And int���� <> -1 Then
        ReDim Preserve strEditors(UBound(strEditors) + 1)
        strTmp = "-1||����||" & Split(strEditors(int����), "||")(2) & "||" & Split(strEditors(int����), "||")(3) & "||" & str���ʷ��� & "||" & RGB_RED & "||" & _
            Split(strEditors(int����), "||")(6) & "||" & Split(strEditors(int����), "||")(7) & "||" & Split(strEditors(int����), "||")(8)
        strEditors(UBound(strEditors)) = strTmp
    End If
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    
    '3����ȡ����������Ŀ��Ϣ�������Ŀ�����Ŀ���ܴ���һ����Ŀ�����λҲҪ��ȡ��
    ArrComTable = Array()
    strTmp = ""
    strTime = ""
    
    '��ȡ���ǻ�����Ŀ
    strSQL = getSQLString("��ȡ���ǻ�����Ŀ", blnMoved)
    If blnMoved Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
        strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
        strSQL = Replace(strSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If

    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, "ȡ��ʼ��", lng�ļ�ID, lng����ID, lng��ҳID, intBaby, Int(CDate(Format(strBeginDate, "yyyy-mm-dd hh:mm:ss"))), CDate(Format(strEndDate, "yyyy-mm-dd hh:mm:ss")), lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID, T_BodyItem.str�����Ŀ)
    bln���� = False
    With rsItems
        Do While Not .EOF
            str��Ŀ���� = ""
            If Val(Nvl(!��Ŀ����, 1)) = 2 Then
                str��Ŀ���� = Trim(Nvl(!��λ)) & Nvl(!��Ŀ����)
            Else
                str��Ŀ���� = Nvl(!��Ŀ����)
            End If
            
            intƵ�� = Val(zlCommFun.Nvl(!��¼Ƶ��))
            
            If zlCommFun.Nvl(!��Ŀ��ʾ) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!��Ŀ���))) Then
                If intƵ�� > 2 Then intƵ�� = 2
            End If
            
            strTmp = zlCommFun.Nvl(!��Ŀ���) & "||" & Replace(str��Ŀ����, ";", ":") & "||" & zlCommFun.Nvl(!��Ŀ��λ) & "||" & _
                zlCommFun.Nvl(!��Ŀֵ��) & "||" & intƵ�� & "||" & zlCommFun.Nvl(!��Ŀ����, 1) & "||" & _
                zlCommFun.Nvl(!��Ŀ��ʾ) & "||" & zlCommFun.Nvl(!��Ŀ����) & "||" & zlCommFun.Nvl(!��Ժ�ײ�, 0)
            If Val(zlCommFun.Nvl(!��Ŀ���)) = gint���� Then
                bln���� = True
            End If
            
            ReDim Preserve ArrComTable(UBound(ArrComTable) + 1)
            ArrComTable(UBound(ArrComTable)) = strTmp
        .MoveNext
        Loop
    End With

    If rsItems.RecordCount > 0 Then rsItems.MoveFirst
    
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '------------------------------------------------------------------------------------------------------------------
    '4��ȷ��X��Y������λ��
    '�߽���Ϣ(Twip)
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
    
    dblSureH = 0
    dblSureW = 0
    If blnPrint = True Then
        '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
        dblSureW = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH), 4)
        dblSureH = Round(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT), 4)
        On Error Resume Next
        dblSureH = (objDraw.Height * dblSureH) / T_TwipsPerPixel.Y
        dblSureW = (objDraw.Width * dblSureW) / T_TwipsPerPixel.X
    End If

    lngRight = gPrinter.lngRight
    lngButtom = gPrinter.lngBottom
     
    lngRight = lngRight * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale
    If lngRight < dblSureW Then lngRight = dblSureW
    lngButtom = lngButtom * (conRatemmToTwip / T_TwipsPerPixel.Y) * sngScale
    If lngButtom < dblSureH Then lngButtom = dblSureH
    lngLeft = lngBeginX * (conRatemmToTwip / T_TwipsPerPixel.X) * sngScale
    If lngLeft < dblSureW Then lngLeft = dblSureW
    lngTop = (lngBeginY / T_TwipsPerPixel.X) * sngScale
    If lngTop < dblSureH Then lngTop = dblSureH
    
    H_16pt = objDraw.TextHeight("��") / T_TwipsPerPixel.Y
    W_16pt = objDraw.TextWidth("��") / T_TwipsPerPixel.X
    
    X = lngLeft: Y = lngTop
    lngCurX = X: lngCurY = Y
        
    T_DrawClient.�̶�����.Left = lngCurX
    T_DrawClient.�̶�����.Right = lngCurX + T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale
    
    lngColStep = (T_BodyStyle.lng�����п� / T_TwipsPerPixel.X) * sngScale
    lngInitRowStep = (T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y) * sngScale
    
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Right = T_DrawClient.�̶�����.Right + (T_BodyStyle.lng������ * T_BodyStyle.lng���� * lngColStep)
    
    Dim sigSign As Single
    sigSign = 1
    If T_DrawClient.��������.Right > objDraw.Width / T_TwipsPerPixel.X - lngRight Then
        sigSign = Round((T_DrawClient.��������.Right - (objDraw.Width / T_TwipsPerPixel.X - lngRight)) / (T_DrawClient.��������.Right - T_DrawClient.�̶�����.Right), 2)
        sigSign = Round((1 - sigSign), 2)
        If sigSign < 0.8 Then sigSign = 0.8
        T_BodyStyle.lng�̶ȿ�� = Fix(T_BodyStyle.lng�̶ȿ�� * sigSign)
        lngColStep = Fix(lngColStep * sigSign)
    End If
    If T_BodyStyle.lng�����п� / T_TwipsPerPixel.X > W_16pt Then
        If lngColStep < W_16pt Then lngColStep = W_16pt
    Else
        lngColStep = (T_BodyStyle.lng�����п� / T_TwipsPerPixel.X) * sngScale
    End If
    
    If lngColStep < gintBmpW Then
        mintBmpW = lngColStep
        mintBmpH = lngColStep
    End If
    
    lngLableStep = Fix((T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X / intDrawLineCOL) * sngScale)
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�̶�����.Right = lngCurX + (T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale)
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Right = T_DrawClient.�̶�����.Right + (T_BodyStyle.lng������ * T_BodyStyle.lng���� * lngColStep) * sngScale
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.�е�λ = lngInitRowStep
    T_DrawClient.ʱ���е�λ = T_BodyStyle.lng���߶� / T_TwipsPerPixel.Y * sngScale
    T_DrawClient.ƫ����X = lngLeft

    '------------------------------------------------------------------------------------------------------------------
    '������п�����߱���ܹ��ж�����
    '������±���Ŀ��������
    intDrawLineRows = Get������(dbl��ֵ, lngCurveRow)
    If intDrawLineRows = 0 Then GoTo ErrPrint

    T_DrawClient.������ = intDrawLineRows

    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    '5��ѭ��������ҳ��ѭ��
    intCurOpt = 0
    intAllOpt = 100
    intCurOpt = intCurOpt + 1
    Call ShowFlash(strInfo, intCurOpt / intAllOpt, objParent)
    If blnPrint = False Then
        lngPicPageIndex = objOut.picPage.UBound + 1
    End If
    
    '��ʽ��ʼ���Ĳ���ѭ��ÿһҳ
    '------------------------------------------------------------------------------------------------------------------
    For lngPage = 1 To lngCountPage
        strTmpDay = Format(CDate(strBeginDate) + T_BodyStyle.lng���� * (lngPage - 1), "YYYY-MM-DD")  '��õ�ǰҳ��ĵ�һ��������ʱ��
        If CDate(strTmpDay) < CDate(strBeginDate) Then strTmpDay = strBeginDate
        If CDate(strEndDate) < CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")) And Not bln��Ժ Then strEndDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        strEndDay = Format(CDate(strTmpDay) + T_BodyStyle.lng���� - 1, "YYYY-MM-DD") & " 23:59:59"
        If CDate(strEndDay) > CDate(strEndDate) Then strEndDay = Format(strEndDate, "YYYY-MM-DD HH:mm:ss")
        intCurOpt = lngPage / lngCountPage
        strInfo = "����" & IIf(blnPrint, "��ӡ���±�", "Ԥ��") & ",���Ժ�..."
        Call ShowFlash(strInfo, intCurOpt, objParent)
        
        '��ҳ�Ŵ�ӡ
        If intBeginPage > 0 Then  'ֻ��ӡָ��ҳ���
            If lngPage >= intBeginPage And lngPage <= intEndPage Then
                If lngPage > intBeginPage Then  '���ڶ�ҳʱ��ʼ��ʼ��ֽ�Ż�ҳ��
                    If Not blnPrint Then
                        Load objOut.picPage(lngPicPageIndex)
                        Set objDraw = objOut.picPage(lngPicPageIndex)
                        objDraw.Cls
                        objDraw.Width = Printer.Width * sngScale
                        objDraw.Height = Printer.Height * sngScale
                        lngPicPageIndex = lngPicPageIndex + 1
                    Else
                        Printer.NewPage
                    End If
                End If
            Else
                GoTo NOPageSub
            End If
        Else  '��ӡ����ʱ
            If lngPage > 1 Then
                If Not blnPrint Then
                    Load objOut.picPage(lngPicPageIndex)
                    Set objDraw = objOut.picPage(lngPicPageIndex) 'PictureBox
                    objDraw.Cls
                    objDraw.Width = Printer.Width * sngScale
                    objDraw.Height = Printer.Height * sngScale
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Printer.NewPage
                End If
            End If
        End If
        
         'ҳüͼ�����
        Call frmTendFileRead.PrintRTBData(objDraw, True, lngTop)
        
        '��ȡ�����DC
        Call ReleaseFontIndirect(objDraw)
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, objDraw.hDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(objDraw.hDC, lngFont)
        lngDC = objDraw.hDC
        '67934:������,2013-12-03,��͸��״̬���л�ͼ
        Call SetBkMode(lngDC, TRANSPARENT)
        '��������
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '��ӡ�ʿغ�
        strTmp = zlDatabase.GetPara("�ʿغ�", glngSys, 1255, "")
        Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
        lngCurX = T_DrawClient.��������.Right - T_Size.W
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, , , , sngScale)
        Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '�Ƿ��ӡҽԺ���ƣ��е�ҽԺ���µ�ҽԺ�����ܴ�����������Ҫ��ҳü��ʵ�֡���ʱ�Ͳ��ڴ�ӡע���ļ��е�ҽԺ��Ϣ��
        If bln��ӡҽԺ���� = True Then
            '��ȡҽԺ����
            Set stdSet = New StdFont
            stdSet.Name = Split(T_BodyStyle.str��������, ",")(0)
            stdSet.Size = Split(T_BodyStyle.str��������, ",")(1) * sngScale
            If InStr(1, T_BodyStyle.str��������, "��") > 0 Then stdSet.Bold = True
            If InStr(1, T_BodyStyle.str��������, "б") > 0 Then stdSet.Italic = True
            Call SetFontIndirect(stdSet, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            strTmp = IIf(GetUnitName = "-", "", GetUnitName) & IIf(intBaby <> 0, "Ӥ��", "") & T_BodyStyle.str�����ı�
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            lngCurY = T_Size.H \ 2 + lngCurY
            Call GetTextRect(objDraw, 0, lngCurY, strTmp, objDraw.Width / T_TwipsPerPixel.X, True, T_Size.H, sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
            objDraw.Font.Size = 9 * sngScale
            Y = lngCurY + T_Size.H \ 2 + 12 * msngTwips
        Else
            Y = lngCurY + 12 * msngTwips
        End If
        lngCurX = X
        lngCurY = Y
        '��ȡ���˿��ҡ����ŵ���Ϣ
    
        VarPatiInfo(3) = ""
        VarPatiInfo(4) = ""
        strTmp = "": strTime = ""
        
        strSQL = getSQLString("��ȡ���Ҵ���", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˿��ҡ����ŵ���Ϣ", lng����ID, lng��ҳID, CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")))
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
                
                If zlCommFun.Nvl(rsTmp("����").Value) <> strTmp And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                
                    strTmp = zlCommFun.Nvl(rsTmp("����").Value)
                    
                    If VarPatiInfo(3) = "" Then
                        VarPatiInfo(3) = strTmp
                    Else
                        VarPatiInfo(3) = VarPatiInfo(3) & "->" & strTmp
                    End If
                    
                End If
    
                If zlCommFun.Nvl(rsTmp("����").Value) <> strTime And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                
                    strTime = zlCommFun.Nvl(rsTmp("����").Value)
                    
                    If VarPatiInfo(4) = "" Then
                        VarPatiInfo(4) = strTime
                    Else
                        VarPatiInfo(4) = VarPatiInfo(4) & "->" & strTime
                    End If
                    
                End If
                            
                rsTmp.MoveNext
            Loop
            
            If Left(VarPatiInfo(3), 2) = "->" Then VarPatiInfo(3) = Mid(VarPatiInfo(3), 3)
            If Left(VarPatiInfo(4), 2) = "->" Then VarPatiInfo(4) = Mid(VarPatiInfo(4), 3)
        End If
        
        If bln���µ���ʾ��� Then
            '��ȡ��ϵ���Сʱ��
            strTmp = GetDiagnoseMinTime(lng����ID, lng��ҳID, CDate(strTmpDay), blnMoved)
            '��ȡ���������Ϣ
            strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As ������ From Dual"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������", "������", lng����ID, lng��ҳID, CDate(Format(strTmp, "yyyy-mm-dd hh:mm:ss")))
            If rsTmp.BOF = False Then
                If intBaby = 0 Then
                    VarPatiInfo(UBound(VarPatiInfo)) = zlCommFun.Nvl(rsTmp("������").Value)
                Else
                    VarPatiInfo(UBound(VarPatiInfo)) = ""
                End If
            Else
                VarPatiInfo(UBound(VarPatiInfo)) = ""
            End If
        End If
        strPatiInfo = Join(VarPatiInfo, "'")
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        stdSet.Italic = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '���������Ϣ
        Call DrawPatiInfo(lngDC, objDraw, strPatiInfo, lngCurX, lngCurY, T_DrawClient.��������.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '---��ʼ�����µ��ϱ��(סԺ����,סԺ����,����,ʱ��)
        Y = lngCurY: lngCurX = X: lngCurY = Y
        '1.��ȡסԺ��ʼ����
        lngValue = 0: strTmp = "": strTime = ""
        strSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡסԺ����", lng�ļ�ID, lng����ID, lng��ҳID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        For i = 0 To T_BodyStyle.lng���� - 1
            strTmp = Format(CDate(strTmpDay) + i, "YYYY-MM-DD")
            If Right(strTmp, 5) = "01-01" Then
                'һ��ĵ�һ��
                strTime = strTmp
            ElseIf strTmp = Format(strBeginDate, "yyyy-MM-dd") Then
                '��Ժ��һ�죬д�����
                strTime = strTmp
            ElseIf i = 0 Then 'ÿҳ�ĵ�һ��
                '70299:������,2014-4-4,ÿҳ����������ʾΪ������(1-��-��-��,0:Ĭ�ϸ�ʽ:��������ʾ)
                If Val(zlDatabase.GetPara("�������ڸ�ʽ", glngSys, 1255, "0")) = 1 Then
                    strTime = strTmp
                Else
                    strTime = Right(strTmp, 5)
                End If
            ElseIf Right(strTmp, 2) = "01" Then
                strTime = Right(strTmp, 5)
            Else
                strTime = Right(strTmp, 2)
            End If

            strTmpString0 = strTmpString0 & "'" & strTime
            strTmpString2 = strTmpString2 & "'" & lngValue + i
        Next i
        strTmpString0 = Mid(strTmpString0, 2)
        strTmpString2 = Mid(strTmpString2, 2)
        '2.��ȡ����ʱ��ʹ���
        strTime = ""
        '��ʾ��ǰ�ε��������
        strSQL = getSQLString("��ȡ��ǰ������Ϣ", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", lng�ļ�ID, intBaby, lng����ID, lng��ҳID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")) - 14), CDate(strEndDay))
        
        ReDim strOpdays(1 To T_BodyStyle.lng����) As String
        ReDim strOpValue(1 To T_BodyStyle.lng����) As String
        
        str����ʱ�� = strEndDay
        Do While Not rsTmp.EOF
            strTime = Format(rsTmp("ʱ��"), "YYYY-MM-DD")
            
            '�����:56005,����,2013-04-27
            If Not rsTmp.EOF Then
                If bln������ʾ And DateDiff("d", CDate(Format(strTime, "YYYY-MM-DD")), str����ʱ��) < 14 Then
                    strEndDay = Format(DateAdd("D", T_BodyStyle.lng���� - 1, CDate(strTmpDay)), "YYYY-MM-DD") & " 23:59:59"
                End If
            End If
            
            For i = 1 To T_BodyStyle.lng����
                If DateDiff("d", strTmpDay, str����ʱ��) + 1 >= i Then
                    intDays = DateDiff("d", strTime, strTmpDay) + (i - 1)

                    Select Case intDays
                        Case 0 '��ǰ�����ڵ�������ʼʱ��
                             'Modify 2012-03-05 �޸�һ������ж������
                            If Trim(strOpdays(i)) <> "" Then
                                strOpdays(i) = strTime & "/" & strOpdays(i)
                            Else
                                strOpdays(i) = strTime
                            End If
                        Case Else
                            If intDays >= 1 And intDays <= intOpDays Then '������ʼ����
                                If blnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                                    strOpValue(i) = intDays
                                Else
                                    If Trim(strOpValue(i)) <> "" Then
                                        If intOpFormat = 3 Then
                                            strOpValue(i) = strOpValue(i) & "/" & intDays
                                        Else
                                            strOpValue(i) = intDays & "/" & strOpValue(i)
                                        End If
                                    Else
                                        strOpValue(i) = intDays
                                    End If
                                End If
                            End If
                    End Select
                End If
            Next i
            rsTmp.MoveNext
        Loop
        
        '��ȡ��ǰ��ʼ����-14��ǰ��������¼��Ϣ
        strSQL = getSQLString("��ȡ14��֮ǰ��������Ϣ", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������", lng�ļ�ID, intBaby, lng����ID, lng��ҳID, Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))))
        
        lng���� = 0
        If rsTmp.BOF = False Then lng���� = Val(rsTmp("����"))
        
        For i = 1 To T_BodyStyle.lng����
            If DateDiff("d", Int(CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss"))), Int(CDate(Format(str����ʱ��, "yyyy-mm-dd hh:mm:ss")))) + 1 >= i Then
                If Trim(strOpdays(i)) <> "" Then
                    arrOperDay = Split(strOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng����
                If Trim(strOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng���� = lngValue + j
                        '�����:57771,����,2013-05-02
                        If intOpFormat = 3 Then
                            strTmp1 = Switch(lng���� = 1, "����", lng���� = 2, "��2", lng���� = 3, "��3", lng���� = 4, "��4", lng���� = 5, "��5", lng���� = 6, _
                            "��6", lng���� = 7, "��7", lng���� = 8, "��8", lng���� = 9, "��9", lng���� = 10, "��10", lng���� = 11, "��11", lng���� = 12, "��12")
                        Else
                            strTmp1 = Switch(lng���� = 1, "��", lng���� = 2, "��", lng���� = 3, "��", lng���� = 4, "��", lng���� = 5, "��", lng���� = 6, _
                            "��", lng���� = 7, "��", lng���� = 8, "��", lng���� = 9, "��", lng���� = 10, "��", lng���� = 11, "��", lng���� = 12, "��")
                        End If
    
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If blnStopFlag Then Exit For
                    Next j
                    lng���� = lngValue + UBound(arrOperDay) + 1
                    If blnStopFlag Then '������ע�������ڴ�����ʱֹͣǰһ�α�ע
                        Select Case intOpFormat
                            Case 1 '��ʾ0
                                strOpValue(i) = 0
                            Case 2 '��ʾ��������
                                If strTmp = "��" Then
                                    strOpValue(i) = 0
                                Else
                                    strOpValue(i) = strTmp & "-0"
                                End If
                            Case 3
                                If strTmp = "��1" Then
                                    strOpValue(i) = "����"
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '����ʾ
                                strOpValue(i) = ""
                        End Select
                    Else
                        Select Case intOpFormat
                            Case 1 '��ʾ0
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = 0 & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = 0
                                End If
                            Case 2 '��ʾ��������
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strTmp & "/" & strOpValue(i)
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case 3
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strOpValue(i) & "/" & strTmp
                                Else
                                    strOpValue(i) = strTmp
                                End If
                            Case Else '����ʾ
                                If Trim(strOpValue(i)) <> "" Then
                                    strOpValue(i) = strOpValue(i)
                                Else
                                    strOpValue(i) = ""
                                End If
                        End Select
                    End If
                End If
            End If
        Next i
        
        strTmpString1 = Join(strOpValue, "'")
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '3��ʼ���סԺ���ڣ�������������Ϣ
        Call DrawUpTableNew(lngDC, objDraw, strTmpString0 & "||" & strTmpString2 & "||" & strTmpString1, lngCurX, lngCurY, T_DrawClient.��������.Right, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '----------------------------------------------------------------------------------------------
         '�˴�����ɴ�ӡ���� �Ӷ��������µ���ӡ���и�
        T_DrawClient.ʱ���е�λ = T_BodyStyle.lng�±��߶� / T_TwipsPerPixel.Y * sngScale
        If intRepairRows = 0 Then
            sngHTab = intRepairRows
        Else
            '�����̶�Ϊ300
            sngHTab = intRepairRows * T_DrawClient.ʱ���е�λ + IIf(bln���� = True, mlngBreatheHeight - T_DrawClient.ʱ���е�λ, 0)
        End If
        
        sngHTab = sngHTab + msngTwips * 12
        sngHPrint = Format(objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHTab, "#0.00;-#0.00;0.00")
        T_DrawClient.�е�λ = (sngHPrint - 2 * T_DrawClient.�е�λ) / (T_DrawClient.������ + T_DrawClient.��������������)
        T_DrawClient.�е�λ = Round(T_DrawClient.�е�λ - 0.05, 1) * sngScale
        If T_DrawClient.�е�λ > T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.�е�λ = T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale
        If T_DrawClient.�е�λ < T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.�е�λ = T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale
        
        '�����иߺ��ڼ������µ��ɴ�ӡ�ı������
        If intRepairRows > 0 Then
            sngHPrint = (T_DrawClient.������ + T_DrawClient.��������������) * T_DrawClient.�е�λ + 2 * T_DrawClient.�е�λ
            sngHTab = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHPrint - (msngTwips * 12)
            sngHTab = sngHTab - IIf(bln���� = True, mlngBreatheHeight - T_DrawClient.ʱ���е�λ, 0)
            If Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3) < intRepairRows Then intRepairRows = Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3)
        End If
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '4��ʼ���̶������������������̶�ֵ��Ϣ
        T_DrawClient.ƫ����Y = lngCurY
        mbln�������� = False
        
        rsTemp.Filter = 0
        rsTemp.Sort = "�������"
        rsTemp.MoveFirst
        str����˵�� = DrawCanvasNew(lngDC, objDraw, rsTemp, rsDrawItems, bln����ӡ������, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        
        '5.��ȡ�����������ݺ����ת�ȱ����Ϣ
        '��ʼ�� ���µ��¼�������ת�ȱ����Ϣ
        
        '���е�ı��ּ���
        '   �ص��Ƿ��ص����.
        '   �ص���Ŀ��¼�ص���Ŀ
        '   �Ͽ�������:����һ��������,����δ��˵��
        '   ��ע:������ʱ��¼ԭֵ
        '   ����:������ע���²���������ֵС�ڵ�����Ŀ��Сֵ���ڵ�����Ŀ���ֵ�ǵ��������.����Ĭ��Ϊ��

        gstrFields = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
             "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
             "����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
             "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
             "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
        Call Record_Init(rsPoints, gstrFields)
    
        '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,13-����,99-δ��˵��)
        '���ñ�ʾ��Ϣ�Ƿ����
        gstrFields = "ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|����," & adDouble & ",2|" & _
            "����," & adLongVarChar & ",200|��ɫ," & adLongVarChar & ",20|X����," & adDouble & ",20|" & _
            "Y����," & adDouble & ",20|�߶�," & adDouble & ",20|��ӡX����," & adDouble & ",20|" & _
            "����," & adInteger & ",1|��ʾ," & adDouble & ",1"
        Call Record_Init(rsNotes, gstrFields)
        
        Dim rs���� As New ADODB.Recordset
        Dim strFileds As String, strValues As String
        
        '��¼������Ϣ
        strFileds = "��Ŀ���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|X����," & adDouble & ",5|ʱ��," & adLongVarChar & ",20"
        Call Record_Init(rs����, strFileds)
        
        Dim int��� As Integer
        
        '----��ȡ���в�λ��Ϣ
        strSQL = "select ��Ŀ���,��λ,ȱʡ�� from ���²�λ"
        Call zlDatabase.OpenRecordset(rsPart, strSQL, "���²�λ")
        '----��ȡ�����������ݺ�δ��˵��
        strSQL = getSQLString("��ȡ�����������ݺ�δ��˵��", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������Ŀ����", lng�ļ�ID, lng����ID, lng��ҳID, CDate(strTmpDay), CDate(strEndDay), T_BodyItem.str������Ŀ)
         
        strTmpString0 = ""
        strTmpString1 = ""
        strTmpString2 = ""
        With rsTmp
            Do While Not .EOF
                strTmp = ""
                blnAllow = False
                strPart = zlCommFun.Nvl(!���²�λ)
                lng��Ŀ��� = Val(zlCommFun.Nvl(!��Ŀ���))
                Select Case lng��Ŀ���
                    Case gint����
                        int��� = 1
                    Case Else
                        int��� = Val(zlCommFun.Nvl(!��¼���))
                End Select
                If strPart = "" Then
                    rsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
                    If rsPart.BOF = False Then
                        strPart = zlCommFun.Nvl(rsPart!��λ)
                    Else
                        Select Case lng��Ŀ���
                            Case gint����
                                strPart = "Ҹ��"
                            Case gint����
                                strPart = "��������"
                            Case Else
                                strPart = ""
                        End Select
                    End If
                End If
                
                SinX = GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinateNew(SinX, Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinateNew(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                
                '��¼����������Ϣ
                If lng��Ŀ��� = gint���� Then
                    strFileds = "��Ŀ���|��ֵ|X����|ʱ��"
                    strValues = lng��Ŀ��� & "|" & zlCommFun.Nvl(!��ֵ) & "|" & SinX & "|" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs����, strFileds, strValues)
                End If
                
                If (Not IsNull(!δ��˵��)) And zlCommFun.Nvl(!��ֵ) <> "����" Then
                    rsNotes.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���)) & " AND X����=" & SinX
                    blnAdd = (rsNotes.RecordCount = 0)
                    '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
                    gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����|��ʾ"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
                    gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & !��Ŀ��� & "|99|" & _
                        !δ��˵�� & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & zlCommFun.Nvl(!��ʾ)
                   
                    If blnAdd Then
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                         Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsNotes!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsNotes!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                             blnAllow = GetCanvasCenterNew(CDate(Format(rsNotes!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
    
                        If blnAllow = True Then
                            If Val(rsNotes!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsNotes, gstrFields, gstrValues, "ʱ��|" & Format(rsNotes!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(rsNotes, gstrFields, gstrValues, "ʱ��|" & Format(rsNotes!ʱ��, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                    End If
                Else
                    blnAdd = False
                    
                    rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " AND X����=" & SinX & " And ���=" & int���
                    
                    blnAdd = (rsPoints.RecordCount = 0)
                    
                    dbl��ֵ = Val(zlCommFun.Nvl(!��ֵ))
                    
                    dblMinValue = GetMaxMinValue(0, lng��Ŀ���, strEditors)
                    dblMaxValue = GetMaxMinValue(1, lng��Ŀ���, strEditors)

                    '��ָ�����ţ���Ŀ���ݲ������ֵ����Сֵ����Ŀ���������ʾ
                    If dbl��ֵ <= dblMinValue Then
                        dbl��ֵ = dblMinValue
                        'strTmp = "��"
                    End If
                    
                    
                    If dbl��ֵ >= dblMaxValue Then
                        dbl��ֵ = dblMaxValue
                        'strTmp = "��"
                    End If
                    
                     '���²���������ʾ��35�̶�
                    If Trim(Nvl(!��ֵ)) = "����" And lng��Ŀ��� = gint���� Then dbl��ֵ = 35
                    
                    sinY = Val(GetYCoordinate(objDraw, rsDrawItems, !��Ŀ���, dbl��ֵ, lngDC, True))
                    
                    gstrFields = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                    gstrValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & strPart & "|" & int��� & "|" & _
                                 Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng��Ŀ��� & "|" & Val(zlCommFun.Nvl(!���Ժϸ�, 0)) & "|" & IIf(zlCommFun.Nvl(!��ֵ, 0) = "����", 1, 0) & "|��|0|" & _
                                 SinX & "|" & sinY & "||" & strTmp & "|" & zlCommFun.Nvl(!��ʾ, 0)
                    If blnAdd Then '���
                        Call Record_Add(rsPoints, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                            blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                       '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(rsPoints!��ʾ) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                gstrFields = "��ʾ"
                                gstrValues = "2"
                                Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                            End If
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
                
        '�����Ѿ��õ���������Ŀ��������Ϣ���������������º���������������
        rsPoints.Filter = ""
        arrTmpValue = Array()
        If int����Ӧ�� = 2 Then
            rsPoints.Filter = "��Ŀ���=" & gint����
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !X���� & ";" & Format(!ʱ��, "yyyy-MM-DD HH:mm:ss")
                .MoveNext
                Loop
            End With
        End If
        
        '������Ϊ��������ʱ����������Ƿ�����Ϊ����
        If int���� <> -1 Then
            For i = 0 To UBound(arrTmpValue)
                '��������Ƿ����������Ӧ
                rs����.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                
                rsPoints.Filter = "��Ŀ���=" & gint���� & " and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                If rsPoints.RecordCount = 0 Then
                    If rs����.RecordCount = 0 Then
                        rsPoints.Filter = ""
                        gstrFields = "��Ŀ���": gstrValues = gint����
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                    Else
                        rsPoints.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                        rsPoints.Delete
                    End If
                End If
            Next i
        End If
        
        If int����Ӧ�� = 2 Then
            Set rs���� = New ADODB.Recordset
            strFileds = "���," & adDouble & ",18|��ֵ," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",200|" & _
                        "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|" & _
                        "����," & adDouble & ",1|�Ͽ�," & adDouble & ",1|�ص���Ŀ," & adLongVarChar & ",50|" & _
                        "�ص�," & adDouble & ",5|X����," & adDouble & ",5|Y����," & adDouble & ",5|��ע," & adLongVarChar & ",50|" & _
                        "����," & adLongVarChar & ",10|��ʾ," & adDouble & ",1"
            Call Record_Init(rs����, strFileds)
            
            rsPoints.Filter = "��Ŀ���=" & gint����
            With rsPoints
                Do While Not .EOF
                    rs����.AddNew
                    For i = 0 To .Fields.Count - 1
                        rs����.Fields(.Fields(i).Name).Value = .Fields(i).Value
                    Next i
                    rs����.Update
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "��Ŀ���=" & gint����
            Do While Not rsPoints.EOF
                rsPoints.Delete
                rsPoints.MoveNext
            Loop
            
            rs����.Filter = ""
            rs����.Sort = "ʱ��"
            With rs����
                Do While Not .EOF
                    blnAdd = False
                    blnAllow = False
                    
                    SinX = Val(zlCommFun.Nvl(!X����))
                    sinY = Val(zlCommFun.Nvl(!Y����))
                    rsPoints.Filter = "��Ŀ���=" & Val(zlCommFun.Nvl(!��Ŀ���, 0)) & " AND X����=" & SinX
                    blnAdd = IIf(rsPoints.RecordCount = 0, True, False)
                    
                    strFileds = "���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����|��ʾ"
                    strValues = Val(zlCommFun.Nvl(!���)) & "|" & !��ֵ & "|" & zlCommFun.Nvl(!��λ) & "|" & Val(zlCommFun.Nvl(!���, 0)) & "|" & _
                                 Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|0|" & Val(zlCommFun.Nvl(!�Ͽ�)) & "|��|0|" & _
                                 SinX & "|" & sinY & "||" & zlCommFun.Nvl(!����) & "|" & Val(zlCommFun.Nvl(!��ʾ, 0))
                    
                    If blnAdd Then '���
                        Call Record_Add(rsPoints, strFileds, strValues)
                    Else
                        If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                            blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '��ȡ�ӽ��м�ʱ����ֵ��Ϊ����ֵ
                        If blnAllow = True Then
                            If Val(rsPoints!��ʾ) = 2 Then
                                arrValues = Split(strValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                strValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(rsPoints, strFileds, strValues, "���|" & rsPoints!���)
                        Else
                            If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                strFileds = "��ʾ"
                                strValues = "2"
                                Call Record_Update(rsPoints, strFileds, strValues, "���|" & rsPoints!���)
                            End If
                        End If
                    End If
                .MoveNext
                Loop
            End With
        End If
        
        '��������������
        arrTmpValue = Array()
        rsPoints.Filter = "��Ŀ���=1 and ���=0"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "��Ŀ���=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount <> 0 Then
                gstrFields = "��ע": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & zlCommFun.Nvl(rsPoints!���))
            End If
        Next i
        
        arrTmpValue = Array()
        rsPoints.Filter = "��Ŀ���=1 and ���=1"
        With rsPoints
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        rsPoints.Filter = "��Ŀ���=1"
        If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=1 and ���=0 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If rsPoints.RecordCount = 0 Then
                rsPoints.Filter = "��Ŀ���=1 and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                rsPoints.Delete
            End If
        Next i
    
        'ɾ����ʾΪ2������
        rsPoints.Filter = ""
        rsPoints.Filter = "��ʾ=2"
        Do While Not rsPoints.EOF
            rsPoints.Delete
        rsPoints.MoveNext
        Loop
        
        rsNotes.Filter = ""
        rsNotes.Filter = "��ʾ=2"
        Do While Not rsNotes.EOF
            rsNotes.Delete
        rsNotes.MoveNext
        Loop
        
        '����δ��˵�����������ݸ���ʾ��һ��
        rsNotes.Filter = ""
        rsPoints.Filter = ""
        
        arrTmpValue = Array()
        arrTmpNote = Array()
        rsNotes.Sort = "��Ŀ���,X����"
        With rsNotes
            Do While Not .EOF
                SinX = Val(!X����)
                blnAllow = False
                rsPoints.Filter = "��Ŀ���=" & Val(!��Ŀ���) & " And X����=" & SinX
                If rsPoints.RecordCount > 0 Then
                    If (zlCommFun.Nvl(rsPoints!��ʾ, 0) = 1 And zlCommFun.Nvl(!��ʾ, 0) = 1) Or (zlCommFun.Nvl(rsPoints!��ʾ, 0) <> 1 And zlCommFun.Nvl(!��ʾ, 0) <> 1) Then
                        blnAllow = GetCanvasCenterNew(CDate(Format(rsPoints!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!��ʾ, 0) = 1 Then
                        blnAllow = True
                    End If
                    If blnAllow = True Then
                        ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                        arrTmpValue(UBound(arrTmpValue)) = !��Ŀ��� & ";" & SinX
                    Else
                        ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                        arrTmpNote(UBound(arrTmpNote)) = !��Ŀ��� & ";" & SinX
                    End If
                End If
            .MoveNext
            Loop
        End With
        
        For i = 0 To UBound(arrTmpValue)
            rsPoints.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
            Do While Not rsPoints.EOF
                rsPoints.Delete
            rsPoints.MoveNext
            Loop
        Next i
        
        For i = 0 To UBound(arrTmpNote)
            rsNotes.Filter = "��Ŀ���=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X����=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
            Do While Not rsNotes.EOF
                rsNotes.Delete
            rsNotes.MoveNext
            Loop
        Next i
    
'        '�������²��� ����Ϊ������Ҫ��35��������������²�������
        rsPoints.Filter = "��Ŀ���=" & gint���� & " and ��ֵ='����' and ���<>1"
        rsPoints.Sort = "ʱ��"
        With rsPoints
            Do While Not .EOF
                strTmpString0 = strTmpString0 & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!��Ŀ���)) & "|99|" & _
                      "����|" & RGB_BLUE & "|" & !X���� & "|0|0|0|0"
                strTmpString2 = strTmpString2 & ";" & !X����
            .MoveNext
            Loop
        End With
        
        '--------���¶Ͽ����
        '����֮����δ��˵���Ͽ���ʱ�����һ��Ͽ�,���²����Ͽ�
        rsPoints.Filter = ""
        
        gstrFields = "�Ͽ�"
        gstrValues = "1"
        rsNotes.Filter = ""
        
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst
        With rsNotes
            Do While Not .EOF
                If int����Ӧ�� = 2 And !��Ŀ��� = -1 Then
                    rsPoints.Filter = "��Ŀ���=" & gint���� & " And X����<=" & !X����
                Else
                    If !��Ŀ��� = 1 Then
                        rsPoints.Filter = "��Ŀ���=" & !��Ŀ��� & " And  ���<>1 And X����<" & !X����
                    Else
                        rsPoints.Filter = "��Ŀ���=" & !��Ŀ��� & " And X����<" & !X����
                    End If
                End If
                rsPoints.Sort = "ʱ��"
                If rsPoints.RecordCount <> 0 Then
                    rsPoints.MoveLast
                    Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & rsPoints!���)
                End If
      
            .MoveNext
            Loop
        End With
        'ʱ�䳬��һ��
        strTime = ""
        strTmp = ""
        rsPoints.Filter = ""
        
        rsPoints.Sort = "��Ŀ���,ʱ��,���"
        With rsPoints
            Do While Not .EOF
                If Not IsNull(!���) Then
                    If Not (Val(!��Ŀ���) = 1 And Val(!���) = 1) Then
                        If lng��Ŀ��� <> 0 Then
                            If lng��Ŀ��� <> !��Ŀ��� Then strTime = ""
                        End If
                        lng��Ŀ��� = !��Ŀ���
                        If strTime <> "" Then
                            If DateDiff("D", CDate(strTime), CDate(Format(!ʱ��, "YYYY-MM-DD"))) > 1 Then
                                strTmp = strTmp & "," & lngValue
                            End If
                        End If
                        strTime = Format(rsPoints!ʱ��, "YYYY-MM-DD")
                        lngValue = Val(rsPoints!���)
                    End If
                End If
                .MoveNext
            Loop
        End With
        
        strTmp = Mid(strTmp, 2)
        For i = 0 To UBound(Split(strTmp, ","))
            Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & Split(strTmp, ",")(i))
        Next i
        
        '�������²�����.��ǰһ����ĶϿ���־����Ϊ1
        rsPoints.Filter = ""
        rsPoints.Filter = "��Ŀ���=" & gint���� & " and ���<>1"
        rsPoints.Sort = "ʱ��,���"
        With rsPoints
            Do While Not .EOF
                If !��ֵ = "����" And .AbsolutePosition <> 1 Then
                    .MovePrevious '������һ�жϿ����
                    If Val(!�Ͽ�) <> 1 Then
                        lngValue = !���
                        Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & lngValue)
                    End If
                    .MoveNext
                End If
            .MoveNext
            Loop
        End With
    
        '��������δ��˵����ͬһX��������ͬ��˵��ֵ���һ��
        rsNotes.Filter = ""
        rsNotes.Sort = "X����"
        With rsNotes
            Do While Not .EOF
                If lngValue = !X���� Then
                    If InStr(1, "," & strTmp & ",", "," & zlCommFun.Nvl(!����) & ",") <> 0 Then
                       rsNotes.Delete
                    Else
                        strTmp = strTmp & "," & zlCommFun.Nvl(!����)
                    End If
                Else
                    lngValue = !X����
                    strTmp = zlCommFun.Nvl(!����)
                End If
            .MoveNext
            Loop
        End With
        
        '--��ȡ���Ժ,�����ȱ��˵��
        Dim bytShow As Byte
        Dim str���� As String
        Dim lng�к� As Long, lngColor As Long
        
        '��ȡ���������±���Ϣ
        '-----------------------------------------------------------------------
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
        strSQL = getSQLString("��ȡ���������±���Ϣ", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���������±����Ϣ", lng�ļ�ID, lng����ID, lng��ҳID, Int(CDate(strTmpDay)), CDate(strEndDay), intBaby, lng����ȼ�)
        With rsTmp
            Do While Not .EOF
                bytShow = 1
                str���� = Trim(zlCommFun.Nvl(!��¼����))
               
                lng�к� = IIf(!��¼���� = 2, 10, IIf(!��¼���� = 6, 11, 4))
                
                '����������ʾ��Ҫ���⴦��
                If !��¼���� = 4 Then
                    str���� = Trim(zlCommFun.Nvl(!��Ŀ����))
                    
                    If str���� = "����" Then
                        bytShow = T_BodyFlag.����
                    ElseIf str���� = "����" Then
                        bytShow = T_BodyFlag.����
                    Else
                        bytShow = T_BodyFlag.����
                    End If
                    
                    If bytShow = 2 Then
                        str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                    Else
                        str���� = !��Ŀ����
                    End If
                    lngColor = lngSignColor
                Else
                    lngColor = IIf(Not IsNumeric(Nvl(!δ��˵��)), RGB_BLUE, Val(Nvl(!δ��˵��)))
                End If
                
                If bytShow > 0 Then
                    SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                    
                    rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=" & !��¼���� & " And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
                    If rsNotes.BOF Then
                        gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|" & !��¼���� & "|" & _
                            str���� & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(rsNotes, gstrFields, gstrValues)
                    Else
                        rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                        rsNotes!���� = str����
                        rsNotes.Update
                    End If
                End If
                rsNotes.Filter = ""
                .MoveNext
            Loop
        End With
        
        '��ȡ���ת����Ϣ
        '-----------------------------------------------------------------------
        '������Ҫ������ı�����(����:2-�ϱ�;3-���ת;4-������;6-�±�,99-δ��˵��)
        '1-��Ժ��2-��ƣ�3-ת�ƣ�4-����
        gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
        Set rsTmp = GetDataFromHis(lng����ID, lng��ҳID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), 2)
        With rsTmp
            Do While Not .EOF
                If Trim(zlCommFun.Nvl(!����)) <> "" Then
                    bytShow = 0
                    lng�к� = Val(!�к�)
                    str���� = zlCommFun.Nvl(!����)
                    Select Case Val(!�к�)
                    Case 5
                        bytShow = T_BodyFlag.��Ժ
                    Case 6, 3 '6ת�룬3ת��
                        bytShow = T_BodyFlag.ת��
                    Case 7
                        bytShow = T_BodyFlag.����
                    Case 8
                        bytShow = T_BodyFlag.��Ժ
                        If intBaby > 0 Then
                            bytShow = IIf(blnӤ�����µ���ʾ��Ժ, bytShow, 0)
                        End If
                    Case 9
                        bytShow = T_BodyFlag.���
                    End Select
                    
                    If bytShow > 0 Then
                        If lng�к� = 9 And bln�����ʾ��Ժ = True And bln��Ʋ�ת��Ժ = True Then str���� = "��Ժ"
                        'Ŀǰ3��4 �����ת�� 3-��ʾ˵���Ϳ��� 4 ��ʾ˵�������ң�ʱ��
                        If bytShow = 2 Then
                            str���� = str���� & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                        ElseIf bytShow = 3 Then
                            str���� = str���� & gstrCaveSplit & zlCommFun.Nvl(!����)
                        ElseIf bytShow = 4 Then
                            str���� = str���� & gstrCaveSplit & zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                        ElseIf bytShow = 1 Then
                            str���� = str����
                        End If
                        
                        SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                        rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=3 And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
                        
                        If rsNotes.BOF Then
                            gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|3|" & _
                                str���� & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(rsNotes, gstrFields, gstrValues)
                        Else
                            rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                            rsNotes!���� = str����
                            rsNotes.Update
                        End If
                    End If
                    rsNotes.Filter = ""
                End If
                .MoveNext
            Loop
        End With
        
        '��ȡӤ��������Ϣ
        If intBaby > 0 Then
            gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"  '���תȱʡ�Ǻ�ɫ,���±꼰δ��˵��ȱʡ����ɫ
            Set rsTmp = GetDataFromHis(lng����ID, lng��ҳID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), 3)
            With rsTmp
                Do While Not .EOF
                    bytShow = 0
                    If Trim(zlCommFun.Nvl(!����)) <> "" Then
                        lng�к� = 12
                        bytShow = T_BodyFlag.����
                        If bytShow > 0 Then
                            Select Case bytShow
                                Case 1
                                    str���� = zlCommFun.Nvl(!����)
                                Case 2
                                    str���� = zlCommFun.Nvl(!����) & gstrCaveSplit & ConvertTimeToChinese(Format(!ʱ��, "HH:mm"))
                            End Select
                            
                            SinX = Val(GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), strTmpDay))
                            rsNotes.Filter = "X����=" & SinX & " and ��Ŀ���=" & lng�к� & " and ����=13 And ʱ��='" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "'"
                            
                            If rsNotes.BOF Then
                                gstrValues = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & lng�к� & "|13|" & _
                                    str���� & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
                                Call Record_Add(rsNotes, gstrFields, gstrValues)
                            Else
                                rsNotes!ʱ�� = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                                rsNotes!���� = str����
                                rsNotes.Update
                            End If
                        End If
                    End If
                    rsNotes.Filter = ""
                .MoveNext
                Loop
            End With
        End If
        '51512,������,2012-07-11,δ��˵����ʾλ�� 0-��ʾ������,1-��ʾ������,2-����ʾ(����)
        '��ҽ��ԺҪ��δ��˵������ʾ������ע��δ�ǵ����ߵ��������߲�����
        strTmp = ""
        Dim arrString() As String
        '�������²��� ���²���ʼ����ʾ�� 35 �����棬ֻ��δ��˵����ʾ�������������Ž���������δ��˵���У���������������±���
        If Left(strTmpString0, 1) = ";" Then
            gstrFields = "ʱ��|��Ŀ���|����|����|��ɫ|X����|Y����|�߶�|��ӡX����|����"
            If mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2 Then
                arrString = Split(strTmpString0, "|")
                arrString(3) = "�� "
                strTmpString0 = Join(arrString, "|")
            End If
            strTmpString0 = Mid(strTmpString0, 2)
            strTmpString2 = Mid(strTmpString2, 2)
            For i = 0 To UBound(Split(strTmpString0, ";"))
                strTmp = Split(strTmpString0, ";")(i)
                rsNotes.Filter = "����=" & IIf(bytδ����ʾλ�� = 1, 99, 6) & " and X����=" & Val(Split(strTmpString2, ";")(i))
                rsNotes.Sort = "��Ŀ���"
                If rsNotes.RecordCount > 0 Then
                    rsNotes!���� = IIf(mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2, "�� ", "����") & ";" & zlCommFun.Nvl(rsNotes!����)
                    rsNotes.Update
                Else
                    If mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 2 Then strTmp = Replace(strTmp, "����", "�� ")
                    Call Record_Add(rsNotes, gstrFields, strTmp)
                    rsNotes!���� = IIf(bytδ����ʾλ�� = 1, 99, 6)
                    rsNotes.Update
                End If
            Next i
        End If
        
        '���δ��˵������ʾ����ȡ����¼��rsNote������Ϊ99�ļ�¼
        If bytδ����ʾλ�� = 2 Then
            rsNotes.Filter = "����=99"
            Do While Not rsNotes.EOF
                rsNotes.Delete
                rsNotes.MoveNext
            Loop
            rsNotes.Filter = ""
        End If
        rsPoints.Filter = 0
        '6 ������֯�ظ��ĵ�
        Call GetConverPoint(rsPoints)
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '7 ��ʼ�����Ϣ������
        strTmp = ShowPointsNew(lngDC, objDraw, rsPoints, strEditors, int����Ӧ��, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '8.����������������
        rsPoints.Filter = ""
        If strTmp <> "" And bln��ӡ������� = True Then Call CreatePolyNew(rsPoints, objDraw, lngDC, strTmpDay, strTmp, int����Ӧ�� = 2)
        '9���˵����Ϣ
        '�ȴ���δ��˵�����±�˵��
        Dim strText As String
        Dim SinY35 As Single, SinY42 As Single
        Dim intAscCharNum As Integer
        
        strTime = ""
        strTmp = ""
        blnAllow = False
        SinX = 0: sinY = 0
        SinY35 = GetYCoordinate(objDraw, rsDrawItems, gint����, 35, lngDC)
        SinY42 = GetYCoordinate(objDraw, rsDrawItems, gint����, 42, lngDC)
        
        rsNotes.Filter = ""
        rsNotes.Sort = "X����,��Ŀ���"
        With rsNotes
            Do While Not .EOF
                strTmp = ""
                For i = 0 To UBound(Split(!����, ";"))
                    If Not (Split(!����, ";")(i) = "����" And bytδ����ʾλ�� = 0 And Nvl(!����) = 99) And Split(!����, ";")(i) <> "" Then
                        If InStr(1, strTmp, Split(!����, ";")(i)) = 0 Then
                            strTmp = strTmp & ";" & Split(!����, ";")(i)
                        End If
                    End If
                Next i
                If Left(strTmp, "1") = ";" Then strTmp = Mid(strTmp, 2)
                If strTmp <> "" Then
                    strTime = Replace(strTmp, ";", " ")
                    If zlCommFun.Nvl(!����) = 99 Then
                        If bytδ����ʾλ�� = 1 Then '��ʾ�����µ�����
                            If blnAllow = True Then
                                If Val(zlCommFun.Nvl(!X����)) <> SinX Then
                                    sinY = SinY35
                                Else
                                    strTime = " " & strTime
                                End If
                            Else
                                sinY = SinY35
                            End If
                            SinX = Val(zlCommFun.Nvl(!X����))
                            For i = 1 To Len(strTime)
                                If sinY < T_DrawClient.�̶�����.Bottom Then
                                    strText = Mid(strTime, i, 1)
                                    T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                    T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                    If T_DrawClient.�̶�����.Bottom - sinY > T_Size.H Then
                                        Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(!��ɫ))
                                    End If
                                    If Asc(strText) < 0 Then
                                        sinY = sinY + T_Size.H
                                    Else
                                        sinY = sinY + T_Size.H / 2
                                    End If
                                End If
                            Next i
                            rsNotes!���� = 1
                            blnAllow = True
                        Else
                            rsNotes!���� = strTime
                            rsNotes!Y���� = SinY42
                            blnAllow = False
                        End If
                    ElseIf zlCommFun.Nvl(!����) = 6 Then
                        If blnAllow = True Then
                            If Val(zlCommFun.Nvl(!X����)) <> SinX Then
                                sinY = SinY35
                            Else
                                strTime = " " & strTime
                            End If
                        Else
                            sinY = SinY35
                        End If
                        SinX = Val(zlCommFun.Nvl(!X����))
                        For i = 1 To Len(strTime)
                            If i < 3 Then intAscCharNum = 0
                            If sinY < T_DrawClient.�̶�����.Bottom Then
                                strText = Mid(strTime, i, 1)
                                T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                If Asc(strText) < 0 Then
                                    If intAscCharNum Mod 2 = 1 Then sinY = sinY + T_Size.H / 2
                                End If
                                '���������Ϣ
                                If T_DrawClient.�̶�����.Bottom - sinY > T_Size.H Then
                                    Call DrawRotateText(objDraw, lngDC, SinX, sinY, strText, Val(zlCommFun.Nvl(!��ɫ)))
                                End If
                                If Asc(strText) < 0 Then
                                    sinY = sinY + T_Size.H
                                    intAscCharNum = 0
                                Else
                                    sinY = sinY + T_Size.H / 2
                                    intAscCharNum = intAscCharNum + 1
                                End If
                            End If
                        Next i
                        rsNotes!���� = 1
                        blnAllow = False
                        sinY = 0
                    Else
                        '���ת�ȱ����Ϣ ��ʼY�����������Ϊ42
                        rsNotes!Y���� = SinY42
                    End If
                End If
            .MoveNext
            Loop
        End With
        If rsNotes.RecordCount > 0 Then rsNotes.MoveFirst: rsNotes.Update
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call OutPutTextNew(objDraw, rsDrawItems, lngDC, rsNotes, strTmpDay, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '��ʼ������±��������Ŀ����
        ReDim ArrNewString(0)
        Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
        
        '��֯��ȡ���±����Ϣ
        For i = 0 To UBound(ArrComTable)
            lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
            str��Ŀ���� = Split(ArrComTable(i), "||")(1)
            If lng��Ŀ��� <> 4 Then
                strItemName = str��Ŀ����
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next i
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        If Not mbln�������� Then strItems = strItems & ",'����'"
        strItems = strItems & ",'����ѹ','����ѹ'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        dtBegin = Int(CDate(strTmpDay) - 1)
        dtEnd = CDate(CDate(Format(strEndDay, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(strBeginDate, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss"))

        
        '��ȡ���б����Ŀ������Ϣ
        strSQL = getSQLString("��ȡ���б����Ŀ������Ϣ", blnMoved, strItems)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Print", _
                                            lng�ļ�ID, _
                                            lng����ID, _
                                            lng��ҳID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, intBaby, lng����ȼ�, IIf(intBaby = 0, 1, 2), lngSectID, T_BodyItem.str�������)
                                                    
        ReDim Preserve ArrNewString(UBound(ArrComTable))
        For i = 0 To UBound(ArrComTable)
            If Split(ArrComTable(i), "||")(0) = 3 Then '������Ŀ
                lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
                strNewTmpString = String(T_BodyStyle.lng������ * T_BodyStyle.lng����, ";")
                arrTmpString0 = Split(String(T_BodyStyle.lng������ * T_BodyStyle.lng����, ";"), ";")
                arrTmpString1 = Split(String(T_BodyStyle.lng������ * T_BodyStyle.lng����, ";"), ";")
                arrTmpString2 = Split(String(T_BodyStyle.lng������ * T_BodyStyle.lng����, ";"), ";")
                
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                rsTmp.Filter = "��Ŀ���=" & gint����
                With rsTmp
                    Do While Not .EOF
                        blnAdd = False
                        If CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")) Then
                            intCOl = GetCurveColumnNew(CDate(!ʱ��), CDate(strTmpDay), gintHourBegin)
                            If intCOl > LBound(ArrNewTmpString) And intCOl <= UBound(ArrNewTmpString) Then
                            
                                If arrTmpString1(intCOl) <> "" Then
                                    If (Val(arrTmpString2(intCOl)) = 0 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 0) Or _
                                        (Val(arrTmpString2(intCOl)) = 1 And Val(zlCommFun.Nvl(!��ʾ, 0)) = 1) Then
                                        
                                        '����Ǹ����ص�ʱ�����
                                        SinX = GetXCoordinateNew(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss"), Format(strTmpDay, "YYYY-MM-DD HH:mm:ss"))
                                        blnAdd = GetCanvasCenterNew(CDate(Format(arrTmpString1(intCOl), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strTmpDay, "YYYY-MM-DD HH:mm:ss")), SinX)
                                    ElseIf Val(arrTmpString2(intCOl)) = 1 Then
                                        blnAdd = False
                                    Else
                                        blnAdd = True
                                    End If
                                    If blnAdd = True Then
                                        If Val(arrTmpString2(intCOl)) = 2 Then
                                            arrTmpString0(intCOl) = zlCommFun.Nvl(!���) & "," & zlCommFun.Nvl(!���²�λ)
                                            arrTmpString1(intCOl) = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    Else
                                        If Val(zlCommFun.Nvl(!��ʾ, 0)) = 2 Then
                                            arrTmpString2(intCOl) = 2
                                            GoTo ErrNext
                                        End If
                                    End If
                                Else
                                    blnAdd = True
                                End If
                                
                                If blnAdd = True Then
                                    arrTmpString0(intCOl) = zlCommFun.Nvl(!���) & "," & zlCommFun.Nvl(!���²�λ)
                                    arrTmpString1(intCOl) = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                                    arrTmpString2(intCOl) = Val(zlCommFun.Nvl(!��ʾ, 0))
                                End If
                                
                            End If
                        End If
ErrNext:
                    .MoveNext
                    Loop

                    For intCOl = LBound(ArrNewTmpString) To UBound(ArrNewTmpString)
                        ArrNewTmpString(intCOl) = IIf(Val(arrTmpString2(intCOl)) = 2, "", arrTmpString0(intCOl))
                    Next intCOl
                    
                    strNewTmpString = Join(ArrNewTmpString, "||")
                End With
                ArrNewString(i) = strNewTmpString
            Else
                blnColor = False
                intƵ�� = Val(Split(ArrComTable(i), "||")(4))
                strTmp = Val(Split(ArrComTable(i), "||")(6)) '��Ŀ��ʾ 4��ʾ������Ŀ
                lng��Ŀ��� = Val(Split(ArrComTable(i), "||")(0))
                str��Ŀ���� = Split(ArrComTable(i), "||")(1)
                int��Ŀ���� = Val(Split(ArrComTable(i), "||")(5))
                int��Ŀ���� = Val(Split(ArrComTable(i), "||")(7))
                int��Ժ�ײ� = Val(Split(ArrComTable(i), "||")(8))
                
                If Val(strTmp) = 4 Or IsWaveItem(lng��Ŀ���) Then
                    If intƵ�� > 2 Then intƵ�� = 2 '����/������ĿƵ��ֻ���� 1 �� 2
                End If
                
                blnColor = (int��Ŀ���� = 2 And int��Ŀ���� = 1 And Val(strTmp) = 0)
                strNewTmpString = String(Val(intƵ��) * T_BodyStyle.lng����, ";")
              
                ArrNewTmpString = Split(strNewTmpString, ";")
                
                For j = 0 To T_BodyStyle.lng���� - 1
                    strBegin = DateAdd("D", j, CDate(strTmpDay))
                    If CDate(strBegin) > CDate(strEndDay) Then strBegin = strEndDay
                    int����ѹ = 0
                    int����ѹ = 0
                    Int�к� = 0
                    strTime = ""
                    intCOl = 0
                    
                    Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(strBegin)), CDate(strBeginDate), lng��Ŀ��� & ";" & str��Ŀ���� & ";" & _
                                    intƵ�� & ";" & Val(strTmp) & ";" & int��Ŀ���� & ";" & int��Ժ�ײ�, bln���ܵ���, bln¼��Сʱ)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "ʱ��,��Ŀ���,���"
                    With rsDownTab
                        Do While Not .EOF
                            lngColor = 0
                            str��� = zlCommFun.Nvl(!��¼����)
                            intCOl = Val(!���)
                            intCOl = intCOl + j * intƵ��
                            If blnColor Then lngColor = Val(zlCommFun.Nvl(!δ��˵��))
                            
                            Select Case zlCommFun.Nvl(!��Ŀ����)
                                Case "����ѹ"
                                    If int����ѹ <> intCOl Then
                                        If Trim(ArrNewTmpString(intCOl)) <> "" Or str��� <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = Trim(Split(ArrNewTmpString(intCOl), "/")(0)) & "/" & str���
                                            Else
                                                ArrNewTmpString(intCOl) = "/" & str���
                                            End If
                                            
                                            mrsCurInfo.Filter = "����='" & str��� & "'"
                                            If Not mrsCurInfo.EOF Then ArrNewTmpString(intCOl) = str���
                                        End If
                                         int����ѹ = intCOl
                                         If ArrNewTmpString(intCOl) = "/" Then ArrNewTmpString(intCOl) = ""
                                    End If
                                Case "����ѹ"
                                    If int����ѹ <> intCOl Then
                                        If ArrNewTmpString(intCOl) <> "" Or str��� <> "" Then
                                            If InStr(1, ArrNewTmpString(intCOl), "/") > 0 Then
                                                ArrNewTmpString(intCOl) = str��� & "/" & Trim(Split(ArrNewTmpString(intCOl), "/")(1))
                                            Else
                                                ArrNewTmpString(intCOl) = str��� & "/"
                                            End If
                                        End If
                                        int����ѹ = intCOl
                                    End If
                                Case Else
                                    If Int�к� <> intCOl Then
                                        ArrNewTmpString(intCOl) = Replace(str���, "-#", "") & "-#" & lngColor
                                        Int�к� = intCOl
                                    End If
                            End Select
                        .MoveNext
                        Loop
                    End With
                    
                    If Format(strBegin, "YYYY-MM-DD") = Format(strEndDay, "YYYY-MM-DD") Then Exit For
                Next j
                strNewTmpString = Join(ArrNewTmpString, "||")
                ArrNewString(i) = strNewTmpString
            End If
        Next i
        
        '��Ŀ���||��λ+��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼Ƶ��||��Ŀ����||��Ŀ��ʾ
        For i = 0 To UBound(ArrComTable)
            strTmpString0 = ""

            If Trim(CStr(Split(ArrComTable(i), "||")(2))) <> "" Then
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1))) & "(" & Trim(CStr(Split(ArrComTable(i), "||")(2))) & ")"
            Else
                strTmpString0 = Trim(CStr(Split(ArrComTable(i), "||")(1)))
            End If
           
            ArrNewString(i) = Trim(CStr(Split(ArrComTable(i), "||")(0))) & ";" & strTmpString0 & ";" & ArrNewString(i)
        Next i
        
        '��ʾƤ�Խ��
        If bln��ʾƤ�� = True Then
            strSQL = getSQLString("��ʾƤ�Խ��", blnMoved)
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˹�����¼��Ϣ", lng����ID, lng��ҳID, intBaby, CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")))

            strNewTmpString = String(T_BodyStyle.lng����, ";")
            ArrNewTmpString = Split(strNewTmpString, ";")
            intCOl = 0

            Do While Not rsTmp.EOF
                intCOl = DateDiff("D", CDate(Format(strTmpDay, "YYYY-MM-DD")), CDate(Format(rsTmp!ʱ��, "YYYY-MM-DD"))) + 1
                ArrNewTmpString(intCOl) = Nvl(rsTmp!ҩ����)
                rsTmp.MoveNext
            Loop
            strNewTmpString = Join(ArrNewTmpString, "||")
            ReDim Preserve ArrNewString(UBound(ArrNewString) + 1)
            ArrNewString(UBound(ArrNewString)) = "-999;Ƥ�Խ��" & ";" & strNewTmpString
        End If
        
        lngCurX = X

        '��ʼ�滭�����Ŀ��չʾ����
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call DrawBodyRecordItemNew(lngDC, objDraw, ArrNewString, rsItems, lngCurX, T_DrawClient.��������.Bottom, T_DrawClient.��������.Right, intRepairRows, lngCurY, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        
        lngCurX = X
        lngCurY = lngCurY
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        '��ʼ��ӡ ҳ�� סԺ���� �� ����˵����Ϣ
        Call DrawBodyPageFooterNew(lngDC, objDraw, lngCurX, lngCurY, T_DrawClient.��������.Right, intPageNo, intEndPage, str����˵��, sngScale)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        'ҳ��ͼ�����
        Call frmTendFileRead.PrintRTBData(objDraw, False, lngButtom)
        
        If Not blnPrint Then objDraw.Refresh
NOPageSub:  Next

    If blnPrint = False Then Call DrawDeviceCapsNew(lngDC, objDraw)
     
    Call ShowFlash
    PrintOrPreviewBodyStateNew = True
    Screen.MousePointer = 0
    Set stdSet = Nothing
    GoTo ErrClare
    Exit Function
ErrPrint:
    Call ShowFlash
    Screen.MousePointer = 0
    
    If ErrCenter = 1 Then
        Resume
    End If
    GoTo ErrClare
    Call SaveErrLog
ErrExit:
    Call ShowFlash
    Screen.MousePointer = 0
    msngTwips = 1
    Err.Clear
    PrintOrPreviewBodyStateNew = False
    Set stdSet = Nothing
    GoTo ErrClare
ErrClare:
    Call ClearData(M_DrawClient.ƫ����X, M_DrawClient.ƫ����Y, M_DrawClient.�̶ȵ�λ, M_DrawClient.�е�λ, M_DrawClient.ʱ���е�λ, M_DrawClient.ʱ���е�λ, _
                    M_DrawClient.�е�λ, M_DrawClient.˫��, M_DrawClient.������, M_DrawClient.��������������, lng�̶ȿ��)
    T_DrawClient.�̶����� = M_DrawClient.�̶�����
    T_DrawClient.�������� = M_DrawClient.��������
    T_DrawClient.���������� = M_DrawClient.����������
    Call ErrEmpty
    Set stdSet = Nothing
End Function

Public Sub DrawUpTableNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmpString As String, _
    ByVal lngX As Long, ByVal lngY As Long, ByVal lngLeft As Long, lngOutY As Long, Optional sngScale As Single)
'-----------------------------------------------------------------------------------------------------------------------
'���һ����Ŀ����Ϣ������ סԺ����,����,������������ʱ������
'����:lngDC ��ͼ�����DC��strTmpString ��סԺ���ڣ����� ������������ɵ��ַ���
'     lngX ��߾�,lngY�ϱ߾�,lngLeft �ұ߾�(���Ի�ͼ������ұ߾�)
'����:lngOutY ���ػ�ͼ����ϱ߾�
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim ArrCode() As String
    Dim strTmp As String
    Dim arrTmpTime() As String 'סԺʱ��
    Dim arrTmpDay() As String  'סԺ����
    Dim arrOptDay() As String '��������
    Dim lngCurX As Long, lngCurY As Long, lngStartY As Long, lngStartX As Long, lngTmpX As Long
    Dim lngColor As Long
    Dim intBold As Integer, intFine As Integer
    Dim str���� As String
    Dim strסԺ���� As String
    Dim str���������� As String
    Dim strʱ�� As String
    
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    str���� = Split(T_BodyStyle.str��ͷ����, "@")(0)
    strסԺ���� = Split(T_BodyStyle.str��ͷ����, "@")(1)
    str���������� = Split(T_BodyStyle.str��ͷ����, "@")(2)
    strʱ�� = Split(T_BodyStyle.str��ͷ����, "@")(3)
    
    ArrCode = Split(strTmpString, "||")
    strTmp = strTmpString & String(2 - UBound(ArrCode), "||")
    ArrCode = Split(strTmp, "||")
    arrOptDay = Split(ArrCode(2), "'")
    arrTmpTime = Split(ArrCode(0), "'")
    arrTmpDay = Split(ArrCode(1), "'")

    lngCurX = lngX: lngStartX = lngX
    lngCurY = lngY: lngStartY = lngY
    
    '��ʼ�������
    
    'X
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ + 6
    Call DrawLine(lngDC, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    'Y
    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = T_DrawClient.�̶�����.Right

    Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)

    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = lngCurX + T_DrawClient.�е�λ * T_BodyStyle.lng������
        Call DrawLine(lngDC, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    Next i
    
    lngCurX = T_DrawClient.�̶�����.Right
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    'ʱ��
    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = T_DrawClient.�̶�����.Right + i * T_DrawClient.�е�λ * T_BodyStyle.lng������
        For j = 1 To T_BodyStyle.lng������ - 1
            lngCurX = lngCurX + T_DrawClient.�е�λ
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ + 6, PS_SOLID, intFine, RGB_BLACK)
        Next j
    Next i
    
    '��ʼ�����Ϣ
    '������Ϣ
    lngCurY = lngStartY
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str����, Len(str����), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, str����, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str����, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    For i = 0 To UBound(arrTmpTime)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpTime(i)), Len(CStr(arrTmpTime(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpTime(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpTime(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 1
    'סԺ����
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strסԺ����, Len(strסԺ����), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, strסԺ����, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, strסԺ����, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To UBound(arrTmpDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, RGB_BLUE)
        Call GetTextExtentPoint32(lngDC, CStr(arrTmpDay(i)), Len(CStr(arrTmpDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpDay(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDC, CStr(arrTmpDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    '��/�������
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 2
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, str����������, Len(str����������), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, str����������, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, str����������, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    '51283,������,2012-07-11,����������ɫ
    lngColor = Val(zlDatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "255"))
    For i = 0 To UBound(arrOptDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDC, lngColor)
        Call GetTextExtentPoint32(lngDC, CStr(arrOptDay(i)), Len(CStr(arrOptDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrOptDay(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDC, CStr(arrOptDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    lngColor = 0
    'ʱ��
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strʱ��, Len(strʱ��), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, strʱ��, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDC, strʱ��, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        '�����������ʱ��
        For j = 0 To T_BodyStyle.lng������ - 1
            strTmp = ""
            
            strTmp = gintHourBegin + T_BodyStyle.lngʱ���� * j

            lngColor = GetTimeColor(Val(strTmp))
            lngTmpX = lngCurX + T_DrawClient.�е�λ * j
            Call SetTextColor(lngDC, lngColor)
            Call GetTextExtentPoint32(lngDC, strTmp, Len(strTmp), T_Size)
            Call GetTextRect(objDraw, lngTmpX - 1, lngCurY + (T_DrawClient.ʱ���е�λ + 6) / 2, strTmp, T_DrawClient.�е�λ, True, , sngScale)
            Call DrawText(lngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Next j
    Next i
    lngOutY = lngStartY + T_DrawClient.ʱ���е�λ * 4 + 6
End Sub


Public Function DrawCanvasNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, rsDrawItems As ADODB.Recordset, Optional ByVal bln����ӡ������ As Boolean = False, Optional sngScale As Single = 1) As String
'------------------------------------------------------------------------------------------------------
'����:���̶������������������̶�ֵ��Ϣ
'����:lngDC ��ͼ�����DC��objDraw �滭����.rsTemp:����������Ŀ��¼��(A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,C.��Ŀ��λ ��λ,A.�����-2 AS �����,B.��λ)
'����:���ظ������ߵľ�����Ϣ����( "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ")
'����˵����Ϣ(��Ŀ�ķ���)
'-------------------------------------------------------------------------------------------------------
    Dim str˵�� As String
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Long   '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Long   '�߽�
    Dim lngCurAlerY As Long '����������
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln˫�� As Boolean                  '�˲������û�ָ��,bln˫��=TRUE��ʾֻ��ʾ����;������ʾʮ��
    Dim bln���� As Boolean                  '�˲������û�ָ��,���зֽ��Ǵ��߻���ϸ��
    Dim blnAche As Boolean                  '�Ƿ�����ʹ��������
    '���¶��Ǳ�׼�߶�
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '������Ϊһ�д�ӡ���
    Dim sinAlertness  As Single              '������,��������
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single, lngInitRowStep As Long
    Dim arrTemp()     As String
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngFont As Long, lngOldFont As Long
    Dim sinY��λ As Single '���ߵ�λ�����Bottom
    Dim lngY As Long, lngCurveRows As Long, lng�̶ȿ�� As Long, lngX As Long
    
    '�������ͼ�������(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    Dim sin�̶� As Single, bln��ʾ�̶� As Boolean, blnFirst As Boolean
    Dim sin�̶ȼ�� As Single, sinBegin�̶� As Single, dbl��λֵ As Double
    Dim sinCurAlerY As Single
    
    Dim str���ֵ���� As String, str��Сֵ���� As String

    On Error GoTo Errhand
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '����������Ŀ����ͼ����(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
    gstrFields = "��Ŀ���," & adDouble & ",18|���ֵ," & adDouble & ",18|��Сֵ," & adDouble & ",18|" & "��λֵ," & adDouble & _
        ",18|���ֵ����," & adLongVarChar & ",20|��Сֵ����," & adLongVarChar & ",20|" & "��λ�̶�," & adLongVarChar & ",20|��ʾģʽ," & adDouble & ",5|��ɫ," & adDouble & ",18"
    Call Record_Init(rsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '����ֵ
    intLineMode = PS_SOLID
    lngColStep = T_DrawClient.�е�λ
    lngInitRowStep = T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale
    sinRowStep = T_DrawClient.�е�λ
    lngLableStep = T_DrawClient.�̶ȵ�λ
    lng�̶ȿ�� = T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale
    
    '���µ��Ե�����ʾ(������ѡ����˫����ʾ��û�����̶���ʾһ��) 1��������ʾ 0��˫����ʾ
    If zlDatabase.GetPara("���µ���ʾ��ʽ", glngSys, 1255, 0) = 1 Then
        bln˫�� = False
    Else
        bln˫�� = True
    End If
    'True��ʾ����ֻ���һ��,Ч����һ���̶�ֻ��ʾ������;����һ���̶���ʾʮ��,���û�������������,��blnDoubleRow�޹�
    bln���� = True
    
    If Not bln���� Then intLineMode = PS_DASHDOTDOT
    '�����
    rsTemp.Filter = "��¼��=1"
    intLables = rsTemp.RecordCount
    rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
        If rsTemp.RecordCount > 0 And bln����ӡ������ = True Then
        rsTemp.Filter = 0
        intLables = intLables - 1
    Else
        rsTemp.Filter = 0
    End If
    If intLables <= 0 Then intLables = 1
    
    lngCurX = T_DrawClient.ƫ����X
    lngCurY = T_DrawClient.ƫ����Y
    lngMaxX = (intLables * lngLableStep) + (T_BodyStyle.lng���� * T_BodyStyle.lng������ * lngColStep) + T_DrawClient.ƫ����X    '�̶�+7*���+ƫ����X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.������ * sinRowStep + T_DrawClient.ƫ����Y '��Ϊ����С�����������ʼY���꣩
       
    str˵�� = ""
        
    SlngMaxY = lngMaxY
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�е�λ = sinRowStep
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.˫�� = blnDoubleRow
    
    For lngRow = 1 To intLables
        lngX = lngCurX
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        If lngRow = intLables Then
            lngCurX = lngCurX + lng�̶ȿ�� - ((intLables - 1) * lngLableStep)
        Else
            lngCurX = lngCurX + lngLableStep
        End If
        
        Call DrawLine(lngDC, lngX, lngCurY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDC, lngX, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    
    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Top = lngCurY
    T_DrawClient.�̶�����.Right = lngCurX
    T_DrawClient.�̶�����.Bottom = lngMaxY
    
    'Ĭ�����һ��������ʾ��Ŀ����
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(lngDC, T_DrawClient.ƫ����X, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    '�����µ�������
    For lngRow = 0 To T_DrawClient.������ - 1
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '�����µ���������
        If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
            Call DrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln����, intBold, intFine), RGB_BLACK)
        End If
    Next
    
    lngCurY = T_DrawClient.�̶�����.Top
    
    '�����µ�������
    For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
        lngCurX = lngCurX + lngColStep
        Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng������ = 0, RGB_RED, RGB_BLACK))
    Next
        
    lngCurX = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.��������.Right = lngMaxX
    T_DrawClient.��������.Bottom = lngMaxY
    
    '�������������
    Call DrawLine(lngDC, T_DrawClient.��������.Left, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)

    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    intLables = 1
    rsTemp.Filter = "��¼��=1"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            If Not (bln����ӡ������ = True And !��Ŀ��� = gint����) Then
                '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
                lngCurX = T_DrawClient.�̶�����.Left + ((intLables - 1) * T_DrawClient.�̶ȵ�λ)
                If .AbsolutePosition = .RecordCount Then
                    lngLableStep = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) - ((intLables - 1) * T_DrawClient.�̶ȵ�λ)
                Else
                    lngLableStep = T_DrawClient.�̶ȵ�λ
                End If
                lngCurY = T_DrawClient.�̶�����.Top
                Set gstdSet = New StdFont
                gstdSet.Name = "����"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                '���������Ŀ������
                Call SetTextColor(lngDC, zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.TextHeight(zlCommFun.Nvl(!��¼��)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��¼��)), lngLableStep, , , sngScale)
'                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!��¼��)), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!��¼��)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                '���������С
                Set gstdSet = New StdFont
                gstdSet.Name = "����"
                gstdSet.Size = 8 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
    
                '�����Ŀ��λ
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngInitRowStep * 2 + objDraw.TextHeight(zlCommFun.Nvl(!��λ)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��λ)), lngLableStep, , , sngScale)
'                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!��λ, 0)), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                Call DrawText(lngDC, Trim(zlCommFun.Nvl(!��λ, 0)), -1, T_LableRect, DT_CENTER)
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                sinY��λ = T_LableRect.Bottom
                intLables = intLables + 1
            End If
            objDraw.Font.Size = 9 * sngScale
            'ǿ���趨����������Ŀ����ʾģʽ
            Select Case !��Ŀ���

                Case gint����  '��������ʱ����̶�
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 1)
                    dbl��λֵ = 0.1
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 37)
                    arrTemp = Split(zlCommFun.Nvl(!��¼��, "��,��,��,��"), ",")
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(����" & arrTemp(0) & ",Ҹ��" & arrTemp(1) & ",����" & arrTemp(2) & ",����" & arrTemp(3) & ")"

                Case gint����, gint����  '����/������10�ı�������̶�
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 10)
                    dbl��λֵ = 2
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)

                    If !��Ŀ��� = gint���� Then
                        str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(ȱʡ��¼��" & zlCommFun.Nvl(!��¼��, "+") & ",����H)"
                    Else
                        str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "��") & ")"
                    End If

                Case gint����  '������5�ı�������̶�
                    mbln�������� = True
                    dbl��λֵ = 1
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, 5)
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(��������" & zlCommFun.Nvl(!��¼��, "*") & ",������R)"
                Case Else
                    dbl��λֵ = Val(zlCommFun.Nvl(!��λֵ, 0))
                    sin�̶ȼ�� = zlCommFun.Nvl(!�̶ȼ��, Val(zlCommFun.Nvl(!��λֵ, 0)) * 10)
                    If sin�̶ȼ�� > Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ)) Then
                        sin�̶ȼ�� = Val(zlCommFun.Nvl(!���ֵ)) - Val(zlCommFun.Nvl(!��Сֵ))
                    End If
                    sinAlertness = zlCommFun.Nvl(!��ʾ��, 0)
                    str˵�� = str˵�� & "��" & zlCommFun.Nvl(!��¼��) & "(" & zlCommFun.Nvl(!��¼��, "*") & ")"
            End Select

            '����ֵ
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow) '�̶�ǰ?�еĸ߶Ȳ�����̶�

            '��������ж�λ����Чλ��
            lngCurY = lngCurY + (T_DrawClient.�е�λ * zlCommFun.Nvl(!�����, 0))
            blnFirst = False
            Do While True
                bln��ʾ�̶� = False
                If blnFirst = False Then     '�ս���ѭ������ʱȡ�����ֵ
                    sin�̶� = zlCommFun.Nvl(!���ֵ, 0)
                    sinBegin�̶� = sin�̶�
                    str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                    blnFirst = True
                Else                    '����õ�ÿ���̶ȵ�ֵ
                    sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                End If
                
                '�������õĿ̶ȼ����ʾ�̶�ֵ
                If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - IIf(T_DrawClient.˫��, sin�̶ȼ�� * 2, sin�̶ȼ��)
                If sinBegin�̶� < Val(Format(zlCommFun.Nvl(!��Сֵ), "#0.00")) Then sinBegin�̶� = Val(Format(zlCommFun.Nvl(!��Сֵ), "#0.00"))
                
                If bln��ʾ�̶� And Not (bln����ӡ������ = True And !��Ŀ��� = gint����) Then
                    '�������ֵ�������ߵ�λ�ظ�
                    If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY < sinY��λ Then
                        Call GetTextRect(objDraw, lngCurX, sinY��λ, Format(sin�̶�, "#0"), lngLableStep, , , sngScale)
                    ElseIf lngCurY = T_DrawClient.�̶�����.Bottom Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y), Format(sin�̶�, "#0"), lngLableStep, , , sngScale)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), lngLableStep, , , sngScale)
                    End If
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Format(sin�̶�, "#0"), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                    Call DrawText(lngDC, Format(sin�̶�, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                '���������Ч��Χ�ڣ����߳����������˳�
                If Val(Format(sin�̶�, "#0.00")) <= Val(Format(zlCommFun.Nvl(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.�̶�����.Bottom Then
                    str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                    '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                    gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ"
                    gstrValues = zlCommFun.Nvl(!��Ŀ���) & "|" & zlCommFun.Nvl(!���ֵ, 0) & "|" & zlCommFun.Nvl(!��Сֵ, 0) & _
                    "|" & dbl��λֵ & "|" & str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & sin�̶ȼ�� & "|" & !��¼ɫ
                    Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                    '�����߻�ʾ��
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!���ֵ)) And sinAlertness > Val(Nvl(!��Сֵ))) Then
                        lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!��Ŀ���)), sinAlertness))
                        Call DrawLine(lngDC, T_DrawClient.��������.Left, lngCurAlerY, lngMaxX, lngCurAlerY, intLineMode, intBold, RGB_RED)
                    End If
                    
                    Exit Do
                End If
                lngCurY = lngCurY + T_DrawClient.�е�λ
            Loop
            sinBegin�̶� = 0
            sin�̶� = 0                 '���ƴӵ�һ�п�ʼ���
            .MoveNext
        Loop
    End With
    
    '��ɶ������߲��ֵ����
    rsTemp.Filter = "��¼��=3"
    rsTemp.Sort = "�������"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = T_DrawClient.ƫ����X
            lngCurveRows = ((Val(Nvl(!���ֵ, 0)) - Val(Nvl(!��Сֵ, 0))) / Val(Nvl(!��λֵ)))
            If Val(Nvl(!�����, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(Nvl(!�����, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * sinRowStep + lngCurY
                '��ɿ̶�����Ļ���
                Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX + Fix(T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale / 2), lngCurY, lngCurX + Fix(T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale / 2), lngMaxY, PS_SOLID, intFine, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX + lng�̶ȿ��, lngCurY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngCurX, lngMaxY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                blnAche = Nvl(!��¼��) Like "��ʹǿ��*"
                '��������еĻ���
                lngCurX = lngCurX + lng�̶ȿ��
                For lngRow = 1 To lngCurveRows
                    '�����µ���������
                    If lngRow <> 0 Then
                        lngCurY = lngCurY + sinRowStep
                    End If
                    If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
                        If blnAche = True Then
                            Call DrawLine(lngDC, lngCurX - lng�̶ȿ�� + Fix(T_BodyStyle.lng�̶ȿ�� / T_TwipsPerPixel.X * sngScale / 2) + 1, lngCurY, lngCurX, lngCurY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_BLACK)
                        End If
                        Call DrawLine(lngDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 And blnAche = False, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln���� And blnAche = False, intBold, intFine), RGB_BLACK)
                    End If
                Next
                '������
                Call DrawLine(lngDC, lngCurX, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                lngCurY = lngY
                '�����µ�������
                For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
                    lngCurX = lngCurX + lngColStep
                    Call DrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng������ = 0, RGB_RED, RGB_BLACK))
                Next
                
                '�����Ŀ���ƺͿ̶ȵ����
                lngX = T_DrawClient.�̶�����.Left
                lngCurX = lngX
                lngCurY = lngY
                '���������Ŀ������
                '��������
                Set gstdSet = New StdFont
                gstdSet.Name = "����"
                gstdSet.Size = 9 * sngScale
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                Call SetTextColor(lngDC, Nvl(!��¼ɫ, RGB_BLACK))
                T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                If T_Size.H * Len(Nvl(!��¼��)) >= lngCurveRows * sinRowStep Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!��¼��)))) \ 2
                End If
                For lngRow = 1 To Len(Nvl(!��¼��))
                    Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Nvl(!��¼��), lngRow, 1), lng�̶ȿ�� \ 2, False)
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Nvl(!��¼��), lngRow, 1), Nvl(!��¼ɫ, RGB_BLACK))
                    Call DrawText(lngDC, Mid(Nvl(!��¼��), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                '�����Ŀ��λ
                lngCurY = lngY: If Nvl(!��¼��) <> "" Then lngCurX = T_LableRect.Right
                If Trim(Nvl(!��λ)) <> "" And Nvl(!��¼��) <> "" Then
                    '���������С
                    Set gstdSet = New StdFont
                    gstdSet.Name = "����"
                    gstdSet.Size = 8 * sngScale
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDC, lngFont)
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(Nvl(!��λ))) >= lngCurveRows * sinRowStep Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!��λ)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(Nvl(!��λ)))
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Trim(Nvl(!��λ)), lngRow, 1), 0, False)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Trim(Nvl(!��λ)), lngRow, 1), Nvl(!��¼ɫ, RGB_BLACK))
                        Call DrawText(lngDC, Mid(Trim(Nvl(!��λ)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call SelectObject(lngDC, lngOldFont)
                    Call DeleteObject(lngFont)
                    Call ReleaseFontIndirect(objDraw)
                End If
                objDraw.Font.Size = 9 * sngScale
                '���������С
                dbl��λֵ = Val(Nvl(!��λֵ, 0))
                sin�̶ȼ�� = Nvl(!�̶ȼ��, Val(Nvl(!��λֵ, 0)) * 10)
                If sin�̶ȼ�� > Val(Nvl(!���ֵ)) - Val(Nvl(!��Сֵ)) Then
                    sin�̶ȼ�� = Val(Nvl(!���ֵ)) - Val(Nvl(!��Сֵ))
                End If
                sinAlertness = Nvl(!��ʾ��, 0)
                str˵�� = str˵�� & "��" & Nvl(!��¼��) & "(" & Nvl(!��¼��, "*") & ")"
                lngCurY = lngY + (sinRowStep * Val(Nvl(!�����, 0)))
                blnFirst = False
                Do While True
                    bln��ʾ�̶� = False
                    If blnFirst = False Then     '�ս���ѭ������ʱȡ�����ֵ
                        sin�̶� = Nvl(!���ֵ, 0)
                        sinBegin�̶� = sin�̶�
                        str���ֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                        blnFirst = True
                    Else                    '����õ�ÿ���̶ȵ�ֵ
                        sin�̶� = sin�̶� - dbl��λֵ     '���Ŀǰ��ʾģʽΪ˫������˫���ۼ�
                    End If
    
                    '�������õĿ̶ȼ����ʾ�̶�ֵ
                    If Val(Format(sin�̶�, "#0.00")) = Val(Format(sinBegin�̶�, "#0.00")) Then bln��ʾ�̶� = True
                    If bln��ʾ�̶� = True Or sin�̶� < sinBegin�̶� Then sinBegin�̶� = sinBegin�̶� - sin�̶ȼ��
                    If sinBegin�̶� < Val(Format(Nvl(!��Сֵ), "#0.00")) Then sinBegin�̶� = Val(Format(Nvl(!��Сֵ), "#0.00"))
    
                    If bln��ʾ�̶� Then
                        '�������ֵ�������ߵ�λ�ظ�
'                        lngCurX = lngX + lng�̶ȿ�� - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin�̶�, "#0.0"))), vbTwips, vbPixels)
'                        lngCurX = lngCurX - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        lngCurX = lngX + lng�̶ȿ�� - Fix(lng�̶ȿ�� / 4) - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin�̶�, "#0.0"))) / 2, vbTwips, vbPixels)
                        If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY = lngY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY + (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        ElseIf lngCurY = lngMaxY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin�̶�, "#0.0")))
                        End If
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Val(Format(sin�̶�, "#0.0")), Nvl(!��¼ɫ, RGB_BLACK))
                        Call DrawText(lngDC, Val(Format(sin�̶�, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin�̶�, "#0.00")) <= Val(Format(Nvl(!��Сֵ), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        str��Сֵ���� = T_DrawClient.��������.Left & "," & lngCurY
                        '��Ӹ���Ŀ(��Ŀ���,���ֵ,��Сֵ,��λֵ,���ֵ����,��Сֵ����,��λ�̶�,��ʾģʽ)
                        gstrFields = "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ"
                        gstrValues = zlCommFun.Nvl(!��Ŀ���) & "|" & zlCommFun.Nvl(!���ֵ, 0) & "|" & zlCommFun.Nvl(!��Сֵ, 0) & _
                        "|" & dbl��λֵ & "|" & str���ֵ���� & "|" & str��Сֵ���� & "|" & T_DrawClient.�е�λ & "," & T_DrawClient.�е�λ & "|" & sin�̶ȼ�� & "|" & !��¼ɫ
                        Call Record_Add(rsDrawItems, gstrFields, gstrValues)
                    
                        '���������
                        If blnDoubleRow = False And sinAlertness > Val(Nvl(!��Сֵ)) And sinAlertness < Val(Nvl(!���ֵ)) Then
                            '�������ֵ�뵱ǰֵ֮��Ĳ��,�Լ���Сֵ,����õ������ٸ��̶�,�ٸ��ݵ�λ�̶ȵõ�ʵ������
                            lngCurAlerY = Val(GetYCoordinate(objDraw, rsDrawItems, Val(Nvl(!��Ŀ���)), sinAlertness))
                            Call DrawLine(lngDC, lngX + lng�̶ȿ��, lngCurAlerY, lngMaxX, lngCurAlerY, PS_SOLID, 1, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + T_DrawClient.�е�λ
                Loop
                sinBegin�̶� = 0
                sin�̶� = 0
            End If
        .MoveNext
        Loop
       T_DrawClient.��������.Bottom = 2 * mintNullRow * lngInitRowStep + (T_DrawClient.������ + T_DrawClient.��������������) * sinRowStep + T_DrawClient.ƫ����Y
    End With
    str˵�� = "˵��:" & Mid(str˵��, 2)
    
    DrawCanvasNew = str˵��
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowPointsNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsPoint As ADODB.Recordset, _
    strEditors() As Variant, Optional int�������� As Integer = 1, Optional ByVal sngScale As Single = 1) As String
'-------------------------------------------------------------------------------------
'����:���������Ŀ�����ߺ�ͼ�����
'����::lngDC ��ͼ�����DC��objDraw �滭����.rsPoint ������Ŀ��ļ���(���|��ֵ|��λ|���|ʱ��|��Ŀ���|����|�Ͽ�|�ص���Ŀ|�ص�|X����|Y����|��ע|����)
'strEditors ���£����ʣ���������������Ϣ(��Ŀ���||��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼��||��¼ɫ)
'����:���ʵ�ļ��� !X���� & ";" & !Y���� & "," & !X���� & ";" & !Y����
'-------------------------------------------------------------------------------------
    Dim sinԭX As Single, sinԭY As Single
    Dim lng��Ŀ��� As Long
    Dim SinX As Single, sinY As Single  '������ʹ��
    Dim dblvalue As Double
    Dim dblMaxValue As Double, dblMinValue As Double
    Dim lngRGB As Long
    Dim strChar As String, str��λ As String, strTmp As String, strPic As String
    Dim str���� As String
    Dim lngCount As Long '�ص���Ŀ����
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnLine As Boolean
    Dim i As Integer
    Dim X1 As Single
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim bln�������� As Boolean
    Dim lngWith As Long
    Dim bln���� As Boolean
    
    On Error GoTo Errhand
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    
    If blnPrinter = True Then
        intBold = 4
        intFine = 4
    Else
        intBold = 2
        intFine = 1
    End If
    rsPoint.Filter = ""
    rsPoint.Sort = "��Ŀ���,ʱ��"
    '���Ƚ�������
    With rsPoint
        Do While Not .EOF
            For i = 0 To UBound(strEditors)
                If Val(Split(strEditors(i), "||")(0)) = Val(zlCommFun.Nvl(!��Ŀ���)) Then
                     Exit For
                End If
            Next i
            If Not (zlCommFun.Nvl(!��Ŀ���) = gint���� And Val(zlCommFun.Nvl(!���)) = 1) Then
                If zlCommFun.Nvl(!��Ŀ���) <> lng��Ŀ��� Then
                    sinԭX = 0
                    sinԭY = 0
                    lngRGB = Split(CStr(strEditors(i)), "||")(5)
                    lng��Ŀ��� = zlCommFun.Nvl(!��Ŀ���)
                End If
                If int�������� = 2 Then
                    If !��Ŀ��� = -1 Then
                        blnLine = False
                    Else
                        blnLine = True
                    End If
                Else
                    blnLine = True
                End If
                
                '�����:56886,����,2013-05-06,ԲȦ���Ų�������
                bln���� = Get����(!�ص�, !�ص���Ŀ, !��Ŀ���, !����, !��λ, strEditors, !���)
                lngWith = 0
                If bln���� Then
                    lngWith = objDraw.TextWidth("��") / 4 / T_TwipsPerPixel.X
                End If
                
                If sinԭX <> 0 And blnLine Then
                    Call DrawLine(lngDC, sinԭX + T_DrawClient.�е�λ / 2, sinԭY, !X���� + T_DrawClient.�е�λ / 2 - lngWith, !Y����, PS_SOLID, intFine, lngRGB)
                End If
                If !�Ͽ� = 0 Then
                    sinԭX = zlCommFun.Nvl(!X����, 0) + lngWith
                    sinԭY = zlCommFun.Nvl(!Y����, 0)
                Else
                    sinԭX = 0
                End If
                
                If !��Ŀ��� = gint���� Then
                    If zlCommFun.Nvl(!����) = 1 Then '���Ժϸ�
                        Call SetTextColor(lngDC, lngRGB)
                        Call GetTextRect(objDraw, !X����, !Y���� - T_DrawClient.�е�λ, "v", T_DrawClient.�е�λ, True, , sngScale)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), "v", lngRGB)
                        Call DrawText(lngDC, "v", -1, T_LableRect, DT_CENTER)
                    End If
                End If
                
                dblMinValue = GetMaxMinValue(0, Val(zlCommFun.Nvl(!��Ŀ���)), strEditors)
                dblMaxValue = GetMaxMinValue(1, Val(zlCommFun.Nvl(!��Ŀ���)), strEditors)
                    
                If Not (Val(Nvl(!��Ŀ���)) = gint���� And Trim(Nvl(!��ֵ)) = "����") Then
                    dblvalue = Val(zlCommFun.Nvl(!��ֵ))
                    If dblvalue > dblMaxValue Then
                        Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
                    ElseIf dblvalue < dblMinValue Then
                        Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
                    End If
                End If
            Else
                '���µ�������
                dblvalue = Split(!��ע, ",")(0)
                SinX = Val(Split(Split(!��ע, ",")(1), ";")(0))
                sinY = Val(Split(Split(!��ע, ",")(1), ";")(1))
                T_Size.H = objDraw.TextHeight("��") / T_TwipsPerPixel.Y

                If Val(!��ֵ) > Val(dblvalue) Then
                    '������ʧ�ܣ�������ͷ�ĺ�ɫʵ�ߣ��ַ��̶��á�
                    'Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y����, SinX + T_DrawClient.�е�λ / 2, sinY, PS_SOLID, intFine, RGB_RED, True)
                    '����ʧ��ҲΪ����(ҽԺҪ��)
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� + (T_Size.H / 4), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                ElseIf Val(!��ֵ) < Val(dblvalue) Then
                    '�����³ɹ�������ɫ���ߣ��ַ��̶��á�
                    Call DrawLine(lngDC, !X���� + T_DrawClient.�е�λ / 2, !Y���� - (T_Size.H / 2), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                End If
            End If
            .MoveNext
        Loop
    End With
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    '������е��ͼ��
    With rsPoint
        Do While Not .EOF
            str��λ = ""
            strTmp = ""
            For i = 0 To UBound(strEditors)
                If Split(CStr(strEditors(i)), "||")(0) = Val(zlCommFun.Nvl(!��Ŀ���)) Then
                     Exit For
                End If
            Next i
            If zlCommFun.Nvl(!�ص�) = 0 And zlCommFun.Nvl(!�ص���Ŀ) = "��" Then 'δ�ص�����Ŀ
                lngRGB = Split(CStr(strEditors(i)), "||")(5)
                If zlCommFun.Nvl(!��Ŀ���) = -1 And int�������� = 2 Then lngRGB = RGB_RED
                str��λ = zlCommFun.Nvl(!��λ)
                If str��λ = "" Then
                    Select Case lng��Ŀ���
                        Case gint����
                            str��λ = "Ҹ��"
                        Case gint����
                            str��λ = "��������"
                        Case Else
                            str��λ = ""
                    End Select
                End If
                strTmp = Split(CStr(strEditors(i)), "||")(4)
                strPic = ""
                strChar = ""
                Select Case zlCommFun.Nvl(!��Ŀ���)
                    Case gint����
                        strTmp = strTmp & String(3 - UBound(Split(strTmp, ",")), ",")
                        If str��λ = "����" Then
                            strChar = Split(strTmp, ",")(0)
                        ElseIf str��λ = "Ҹ��" Then
                            strChar = Split(strTmp, ",")(1)
                        ElseIf str��λ = "����" Then
                            strChar = Split(strTmp, ",")(2)
                        Else
                            strChar = Split(strTmp, ",")(3)
                        End If
                        If zlCommFun.Nvl(!���) = 1 Then '�����·���
                            lngRGB = RGB_RED
                            strChar = "��"
                        Else
                            If strChar = "" Then strChar = "��"
                        End If
                    Case gint����
                        strChar = IIf(strTmp = "", "��", strTmp)
                    Case gint����
                        If str��λ = "����" Then
                            strPic = "PACEMAKER"
                        Else
                            strChar = IIf(strTmp = "", "+", strTmp)
                        End If
                    Case gint����
                        If str��λ = "��������" Then
                            strChar = IIf(strTmp = "", "*", strTmp)
                        Else
                            strPic = "BREATH"
                        End If
                    Case Else
                        strChar = strTmp
                End Select
                If Trim(zlCommFun.Nvl(!����)) <> "" Then
                    strChar = Trim(zlCommFun.Nvl(!����))
                    strPic = ""
                End If
                
                If !��Ŀ��� = gint���� And Trim(Nvl(!��ֵ)) = "����" And (mlng���²�����ʾ��ʽ = 0 Or mlng���²�����ʾ��ʽ = 1) Then
                    bln�������� = False
                Else
                    bln�������� = True
                End If
                                
                If strPic = "" And bln�������� Then
                    Call SetTextColor(lngDC, lngRGB)
                    Call GetTextRect(objDraw, !X����, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
                    T_LableRect.Left = T_LableRect.Left - 1
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                    Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                    'Debug.Print T_LableRect.Left & ";" & T_LableRect.Right
                Else
                    Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y���� - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                        objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y���� + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), True)
                End If
            
            Else  'չʾ�ص���λͼ��
                strPic = ""
                strChar = ""
                If zlCommFun.Nvl(!�ص���Ŀ) <> "��" Then '�ص�=1�Ĳ����κδ���
                    lngCount = UBound(Split(zlCommFun.Nvl(!�ص���Ŀ), ","))
                    strTmp = zlCommFun.Nvl(!�ص���Ŀ)
                    If Trim(strTmp) <> "" Then
                        str��λ = zlCommFun.Nvl(!��λ)
                        lngCount = lngCount + 2
                        strTmp = zlCommFun.Nvl(!��Ŀ���) & "," & strTmp
                        If InStr(1, "," & strTmp & ",", ",1,") <> 0 Then

                            strSQL = "SELECT A.���,A.��Ƿ���,A.�����ɫ" & vbNewLine & _
                                    " FROM �����ص���� A," & vbNewLine & _
                                    "     (SELECT �ϼ����, COUNT(*) ����" & vbNewLine & _
                                    "     FROM �����ص����" & vbNewLine & _
                                    "     WHERE ��Ŀ��� IN (" & strTmp & ")" & vbNewLine & _
                                    "     GROUP BY �ϼ����) B" & vbNewLine & _
                                    " WHERE A.�ص���Ŀ = B.����" & vbNewLine & _
                                    " AND A.��� = B.�ϼ���� AND A.���=[1]"
                        Else
                            strSQL = "Select A.���, A.��Ƿ���, A.�����ɫ" & vbNewLine & _
                                "  From �����ص���� A," & vbNewLine & _
                                "       (Select �ϼ����, Count(1) ����" & vbNewLine & _
                                "          from �����ص����" & vbNewLine & _
                                "         where ��Ŀ��� in (" & strTmp & ")" & vbNewLine & _
                                "         group by �ϼ����) B" & vbNewLine & _
                                " Where A.�ص���Ŀ = B.����" & vbNewLine & _
                                "   And A.��� = B.�ϼ����"
                        End If
                        
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ص�", Val(str��λ))
                        
                        If rsTmp.RecordCount > 0 Then
                            If IsNull(rsTmp!��Ƿ���) Then
                                strPic = zlBlobRead(9, zlCommFun.Nvl(rsTmp!���))
                            Else
                                strChar = Trim(zlCommFun.Nvl(rsTmp!��Ƿ���))
                                lngRGB = Val(zlCommFun.Nvl(rsTmp!�����ɫ, 0))
                            End If
                            If strPic = "" Then
                                Call SetTextColor(lngDC, lngRGB)
                                Call GetTextRect(objDraw, !X���� - 1, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
'                                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                                Call DrawText(lngDC, Trim(strChar), -1, T_LableRect, DT_CENTER)
                            Else
                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), objDraw.ScaleY(!Y���� - mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), _
                                    objDraw.ScaleX(!X���� + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), objDraw.ScaleY(!Y���� + mintBmpH * IIf(blnPrinter = True, msngTwips, 1) / 2, vbPixels, vbTwips), False)
                                
                                Call FileSystem.Kill(strPic)
                            End If
                        End If
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '��ȡ�������ʵ���Ϣ
    If rsPoint.RecordCount > 0 Then rsPoint.MoveFirst
    rsPoint.Filter = "��Ŀ���=" & gint����
    With rsPoint
        Do While Not .EOF
            str���� = str���� & "," & !X���� & ";" & !Y����
        .MoveNext
        Loop
    End With
    If str���� <> "" Then str���� = Mid(str����, 2)
    
    ShowPointsNew = str����
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub DrawBodyRecordItemNew(ByVal lngDC As Long, ByVal objDraw As Object, strValue() As String, ByVal rsItems As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long, _
    ByVal lngLeft As Long, ByVal intRepairRows As Integer, lngOutY As Long, Optional sngScale As Single = 1)
'-----------------------------------------------------------------------------------------------------------------------
'������˻�����Ϣ
'����:lngDC ��ͼ�����DC��strValue() ���б����Ŀ����Ϣ (��ʽ��������:��Ŀ���;����;����,��λ||����,��λ/(����) ��Ŀ���;����;����||����) ���ݺͲ�λ��ɵ������ʾ����Ŀ�ж�����
'    rsItems �������±������Ŀ, lngX ��߾�,lngY�ϱ߾�,lngLeft �ұ߾�(���Ի�ͼ������ұ߾�),intRepairRows Ҫ��ӡ�����Ŀ��������
'����:lngOutY ���ػ�ͼ����ϱ߾�
'-----------------------------------------------------------------------------------------------------------------------
    Dim lngX1 As Long, lngY1 As Long, lngCurY As Long, lngCurX As Long
    Dim lngRowHeiht As Long, lngTestisHeight As Long, arrTestis
    Dim arrTmpString0() As String, arrTmpString1() As String
    Dim arrTmp() As String, arrText() As String, arrData
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer, j As Integer
    Dim int������������ʽ As Integer
    Dim bln�೦����Է��ӷ�ĸ��ʾ As Boolean
    Dim strTmp As String, strPart As String
    Dim strPic As String
    Dim blnValue As Boolean
    Dim intValue As Integer, int����λ�� As Integer
    Dim intRowCount As Integer
    Dim intƵ�� As Integer '��¼Ƶ��
    Dim blnDataTrue As Boolean
    Dim lngColor As Long
    Dim intNum As Integer
    Dim blnOutText As Boolean '�Ƿ�����ı�
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim sgnSize As Single
    Dim sngLen As Single, lngLen As Long
    Dim LPoint As T_LPoint
    Dim lngFont As Long, lngOldFont As Long
    Dim bln��ʾƤ�� As Boolean
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    
    On Error GoTo Errhand
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
        intBold = 6
        intFine = 2
    Else
        msngTwips = 1
        intBold = 2
        intFine = 1
    End If
    
    lngCurY = lngY
    lngCurX = lngX
    blnValue = False
    intValue = 0
    int����λ�� = 0
    int������������ʽ = zlDatabase.GetPara("����������", glngSys, 1255, 0)
    bln�೦����Է��ӷ�ĸ��ʾ = (Val(zlDatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    
    strPic = ""
    If InStr(1, strValue(0), ";") > 0 Then
        bln��ʾƤ�� = IIf(Split(strValue(UBound(strValue)), ";")(0) = "-999", True, False)
        
        For intRow = LBound(strValue) To UBound(strValue)
            arrTmpString0 = Split(strValue(intRow), ";")
            arrTmpString1 = Split(arrTmpString0(2), "||")
            
            If intRepairRows > 0 And intRepairRows > intRowCount Then
            
                If arrTmpString0(0) = "3" Then '������Ŀ
                    '��ȡ�����ɫ
                    rsItems.Filter = 0
                    rsItems.Filter = "��Ŀ���=" & gint����
                    If rsItems.RecordCount > 0 Then
                        lngColor = Val(Nvl(rsItems!��¼ɫ, RGB_RED))
                    Else
                        lngColor = RGB_RED
                    End If
                    intRowCount = intRowCount + 1
                    arrTmpString1 = Split(arrTmpString0(2), "||")
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '��ͷ
                            Call SetTextColor(lngDC, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                            LPoint.X = lngX
                            LPoint.Y = lngY
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX
                            LPoint.H = mlngBreatheHeight
                            
                            Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            Call DrawLine(lngDC, lngX, lngY, lngX, lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            Call DrawLine(lngDC, lngX, lngY + mlngBreatheHeight, T_DrawClient.�̶�����.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY, T_DrawClient.�̶�����.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            lngX1 = T_DrawClient.�̶�����.Right
                            lngY1 = lngCurY
                        Else
                            arrTmpString1(intCOl) = arrTmpString1(intCOl) & String(1 - UBound(Split(arrTmpString1(intCOl), ",")), ",")
                            strTmp = Split(arrTmpString1(intCOl), ",")(0)
                            strPart = Split(arrTmpString1(intCOl), ",")(1)
                            If strPart = "" Then strPart = "��������"
                            strPic = ""
                            '��ӡ����ֵ���������ӡ�� ��һ��ʼ��������
                            If IsNumeric(strTmp) Then
                                If strPart = "��������" Then
                                    Call SetTextColor(lngDC, lngColor)
                                    T_Size.H = objDraw.TextHeight(strTmp) / T_TwipsPerPixel.Y
                                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                Else
                                    strPic = "BREATH"
                                End If
                                
                                If blnValue = False Then
                                    intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                                    blnValue = True
                                    int����λ�� = 2
                                End If
                                
                                If int������������ʽ = 0 Then '˳��������ʾ
                                    If intCOl Mod 2 = intValue Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                vbPixels, vbTwips), objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                Else        '������ʱ����֮��������ʾ
                                    If int����λ�� = 2 Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                        Else
                                            Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                        End If
                                    End If
                                    
                                   
                                    int����λ�� = int����λ�� + 1
                                    If int����λ�� > 2 Then int����λ�� = 1
                                End If
                                
                            End If
                            lngX1 = lngX1 + T_DrawClient.�е�λ
                        End If
                    Next intCOl
                    lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ
                    lngY1 = lngY + mlngBreatheHeight
                    
                    '�����������е���
                    For intCOl = 1 To T_BodyStyle.lng���� * T_BodyStyle.lng������
                        If intCOl Mod T_BodyStyle.lng������ = 0 Then
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Else
                            Call DrawLine(lngDC, lngX1, lngY, lngX1, lngY1, PS_SOLID, intFine, RGB_BLACK)
                        End If
                        lngX1 = lngX1 + T_DrawClient.�е�λ
                    Next intCOl
                    Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                    
                    '��ǰY������
                    lngCurY = lngY1
                ElseIf arrTmpString0(0) <> "-999" Then '����Ƥ�Խ��
                    
                    rsItems.Filter = ""
                    rsItems.Filter = "���=" & intRow
                    If rsItems.RecordCount > 0 Then
                        intƵ�� = CInt(zlCommFun.Nvl(rsItems!��¼Ƶ��, 2))
                        If Val(Nvl(rsItems!��Ŀ��ʾ)) = 4 Or IsWaveItem(Val(Nvl(rsItems!��Ŀ���))) Then
                            If intƵ�� > 2 Then intƵ�� = 2 '����/������ĿƵ��ֻ���� 1 �� 2
                        End If
                        '���Ŀ����Ƿ�������ݣ������ھͲ���ӡ����
                        If zlCommFun.Nvl(rsItems!��Ŀ����) = 2 Then
                            
                            If Trim(Replace(arrTmpString0(2), "||", "")) = "" Then
                                blnDataTrue = False
                            Else
                                blnDataTrue = True
                            End If
                        Else
                            blnDataTrue = True
                        End If
                    Else
                        blnDataTrue = False
                    End If
                    
                    If blnDataTrue = True Then
                        lngY1 = lngCurY
                        lngX1 = lngCurX
                        
                        '����Ƶ�μ���Ҫ��ӡ�ı�������Ƿ񳬳��û����õı������
                        
                        intNum = 0
                        Select Case intƵ��
                            Case 1, 2, 6
                                intRowCount = intRowCount + 1
                            Case 3
                                intRowCount = intRowCount + 3
                            Case 4
                                intRowCount = intRowCount + 2
                            Case Else
                                intRowCount = intRowCount + 1
                        End Select
                        
                        If intRowCount > intRepairRows Then
                            intNum = intRowCount - intRepairRows
                            intRowCount = intRepairRows
                        End If
                        blnOutText = False
                        
                        For intCOl = 0 To UBound(arrTmpString1)
                            If intCOl = 0 Then '��ʼ����ͷ��Ϣ������������
                                Select Case intƵ��
                                    Case 1, 2, 6
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ
                                        lngRowHeiht = T_DrawClient.ʱ���е�λ / 2
                                    Case 3
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ * (3 - intNum)
                                        lngRowHeiht = (T_DrawClient.ʱ���е�λ * (3 - intNum)) / 2
                                    Case 4
                                        lngY1 = lngY1 + T_DrawClient.ʱ���е�λ * (2 - intNum)
                                        lngRowHeiht = (T_DrawClient.ʱ���е�λ * (2 - intNum)) / 2
                                End Select

                                Call SetTextColor(lngDC, RGB_BLACK)
                                T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                                LPoint.X = lngX1
                                LPoint.Y = lngY1 - lngRowHeiht * 2
                                LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                                LPoint.H = lngRowHeiht * 2
                                Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                
                                lngY1 = lngCurY
                                lngX1 = T_DrawClient.�̶�����.Right
                            Else  '��ʼ���л������
                                strTmp = CStr(arrTmpString1(intCOl))
                               
                                If InStr(1, strTmp, "-#") <> 0 Then
                                    If Not IsNumeric(Split(strTmp, "-#")(1)) Then
                                        lngColor = 0
                                    Else
                                        lngColor = Val(Split(strTmp, "-#")(1))
                                        strTmp = Split(strTmp, "-#")(0)
                                    End If
                                Else
                                    lngColor = 0
                                End If
                                
                                If strTmp = "*" And Val(arrTmpString0(0)) = gint��� Then strTmp = "��"
                                
                                Call SetTextColor(lngDC, lngColor)
                                
                                T_Size.H = objDraw.TextHeight(strTmp) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                                blnOutText = True
                                
                                If InStr(1, ",3,4,", "," & intƵ�� & ",") = 0 Then
                                    LPoint.X = lngX1
                                    LPoint.Y = lngCurY
                                    LPoint.W = T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                                ElseIf intƵ�� = 3 Then
                                    LPoint.W = T_DrawClient.�е�λ * T_BodyStyle.lng������
                                    If intCOl Mod intƵ�� = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ * 2
                                        If intNum <> 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * T_BodyStyle.lng������
                                    ElseIf intCOl Mod intƵ�� = 2 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ
                                        If intNum > 1 Then blnOutText = False
                                    Else
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                    End If
                                    
                                ElseIf intƵ�� = 4 Then
                                    LPoint.W = T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                    If intCOl Mod 4 = 3 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                    ElseIf intCOl Mod 4 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY + T_DrawClient.ʱ���е�λ
                                        If intNum > 0 Then blnOutText = False
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                    ElseIf intCOl Mod 2 = 0 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                        lngX1 = lngX1 - T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                    ElseIf intCOl Mod 4 = 1 Then
                                        LPoint.X = lngX1
                                        LPoint.Y = lngCurY
                                        lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                    End If
                                End If
                                LPoint.H = T_DrawClient.ʱ���е�λ
                                
                                If blnOutText = True Then
                                    If AnsyGrade(Val(arrTmpString0(0)), strTmp, arrText) = True Then
                                        Call DrawAnsyGrade(lngDC, objDraw, arrText, LPoint, lngColor, bln�೦����Է��ӷ�ĸ��ʾ, sngScale)
                                    Else
                                        Call DrawTabTextNew(lngDC, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale)
                                    End If
                                End If
                   
                            End If
                        Next intCOl
                        
                        '����Ԫ������
                        If InStr(1, ",2,3,4,", "," & intƵ�� & ",") = 0 Then
                            lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                            lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                            For intCOl = 1 To intƵ�� * T_BodyStyle.lng����
                                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod intƵ�� = 0, intBold, intFine), RGB_BLACK)
                                lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                            Next intCOl
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                        ElseIf intƵ�� = 3 Then
                            intRowCount = intRowCount - (intƵ�� - intNum)
                            intValue = intRowCount
                            For i = 1 To 3 - intNum
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * T_BodyStyle.lng������
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                For intCOl = 1 To T_BodyStyle.lng����
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * T_BodyStyle.lng������
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                
                                lngCurY = lngY1
                            Next i
                        ElseIf InStr(1, ",2,4,", "," & intƵ�� & ",") <> 0 Then
                            intRowCount = intRowCount - (intƵ�� / 2 - intNum)
                            intValue = intRowCount
                            For i = 1 To (intƵ�� / 2 - intNum)
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                For intCOl = 1 To T_BodyStyle.lng���� * 2
                                    Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, intBold, intFine), RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                lngCurY = lngY1
                            Next i
                        End If
                        
                        lngCurY = lngY1
                    End If
                End If
                
                intNum = 0
                arrTestis = Array()
                'Ƥ�Խ��,ֻ�����������ݣ�����ڲ������д���
                If arrTmpString0(0) = "-999" Then
                    lngY1 = lngCurY
                    lngX1 = lngCurX
                    intƵ�� = 1
                    
                    arrTestis = Array(0) 'Ƥ�Խ��������ʾ�����ڴ��ÿһ��Ƥ�Խ�������߶�
                    arrTestis(0) = Val(Format(T_DrawClient.ʱ���е�λ * T_TwipsPerPixel.Y, "#0"))
                    
                    lngTestisHeight = Val(Format(T_DrawClient.ʱ���е�λ * T_TwipsPerPixel.Y, "#0")) 'Ƥ�Խ��ռ�õ����߶�
                    '�õ�Ƥ�Խ��ռ�õ�������
                    LPoint.W = T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                    For intCOl = 1 To UBound(arrTmpString1)
                        intNum = 1
                        strTmp = CStr(arrTmpString1(intCOl))
                        If strTmp = "" Then strTmp = "-#"
                        arrTmp = Split(strTmp, ",")
                        T_Size.H = 0
                        If UBound(arrTmp) > UBound(arrTestis) Then
                            ReDim Preserve arrTestis(UBound(arrTmp))
                        End If
                        For i = LBound(arrTmp) To UBound(arrTmp)
                            strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
                            If Trim(strTmp) <> "" Then
                                sgnSize = GetFontSize(objDraw, CStr(strTmp) & "L", LPoint.W, sngScale)
                                With frmTendFileRead.txtLength
                                    .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(blnPrinter, 12, 0)
                                    .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                                    .FontName = "����"
                                    .FontSize = sgnSize * sngScale
                                    .FontBold = False
                                    .FontItalic = False
                                End With
                                
                                arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
                                '����ĳһ��Ƥ�Խ���ĸ߶�
                                If Val(objDraw.TextHeight("��") * (UBound(arrData) + 1)) < Val(Format(T_DrawClient.ʱ���е�λ * T_TwipsPerPixel.Y, "#0")) Then
                                    lngRowHeiht = Val(Format(T_DrawClient.ʱ���е�λ * T_TwipsPerPixel.Y, "#0"))
                                Else
                                    lngRowHeiht = objDraw.TextHeight("��") * (UBound(arrData) + 1)
                                End If
                                T_Size.H = T_Size.H + lngRowHeiht
                                If Val(arrTestis(i)) < lngRowHeiht Then arrTestis(i) = lngRowHeiht
                                intNum = intNum + 1
                                If intRowCount + intNum > intRepairRows Then Exit For
                            End If
                        Next i
                        If lngTestisHeight < T_Size.H Then lngTestisHeight = T_Size.H
                    Next intCOl
                    Call ReleaseFontIndirect(objDraw)
                    lngTestisHeight = Val(Format(lngTestisHeight / T_TwipsPerPixel.Y, "#0"))
                    
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '��ʼ����ͷ��Ϣ������������
                            lngY1 = lngY1 + lngTestisHeight
                            lngRowHeiht = lngTestisHeight / 2
                               
                            Call SetTextColor(lngDC, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                
                            LPoint.X = lngX1
                            LPoint.Y = lngY1 - lngTestisHeight
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                            LPoint.H = lngTestisHeight
                            Call DrawTabTextNew(lngDC, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            
                            lngY1 = lngCurY
                            lngX1 = T_DrawClient.�̶�����.Right
                        Else  '��ʼ���л������
                            intNum = 1
                            strTmp = CStr(arrTmpString1(intCOl))
                            If strTmp = "" Then strTmp = "-#"
                            LPoint.X = lngX1
                            LPoint.Y = lngCurY
                            LPoint.W = T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                            '��ʼ�����Ƿ���Ҫ����
                            strPart = ""
                            
                            arrTmp = Split(strTmp, ",")
                            
                            For i = LBound(arrTmp) To UBound(arrTmp)
                                lngColor = Val(Split(arrTmp(i), "-#")(0))
                                '����������ɫ
                                Call SetTextColor(lngDC, lngColor)
                                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") 'Ƥ�Խ��
                                If Trim(strTmp) <> "" Then
                                    sgnSize = GetFontSize(objDraw, CStr(strTmp) & "L", LPoint.W, sngScale)
                                    '����Ƥ�Խ�������ʵ������
                                    With frmTendFileRead.txtLength
                                        .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(blnPrinter, 12, 0)
                                        .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                                        .FontName = "����"
                                        .FontSize = sgnSize
                                        .FontBold = False
                                        .FontItalic = False
                                    End With
                                    arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
                                    
                                    Set gstdSet = New StdFont
                                    gstdSet.Name = "����"
                                    gstdSet.Size = sgnSize
                                    gstdSet.Bold = False
                                    gstdSet.Italic = False
                                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                                    lngFont = CreateFontIndirect(T_Font)
                                    lngOldFont = SelectObject(lngDC, lngFont)
                                    lngY1 = LPoint.Y
                                    If Val((UBound(arrData) + 1) * objDraw.TextHeight("��")) < Val(arrTestis(i)) Then
                                        LPoint.Y = LPoint.Y + (Val(arrTestis(i)) - ((UBound(arrData) + 1) * objDraw.TextHeight("��"))) / T_TwipsPerPixel.Y / 2
                                    End If
                                    
                                    '��ʼ�������
                                    For j = 0 To UBound(arrData)
                                        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(arrData(j)), , False, , sngScale)
                                        Call DrawText(lngDC, CStr(arrData(j)), -1, T_LableRect, DT_CENTER)
                                        LPoint.X = lngX1
                                        LPoint.Y = LPoint.Y + Val(Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0"))
                                    Next j
                                    LPoint.Y = lngY1 + Val(Format(Val(arrTestis(i)) / T_TwipsPerPixel.Y, "#0"))
                                    Call SelectObject(lngDC, lngOldFont)
                                    Call DeleteObject(lngFont)
                                    Call ReleaseFontIndirect(objDraw)
                                    
                                    intNum = intNum + 1
                                    If intRowCount + intNum > intRepairRows Then GoTo ErrNext
                                End If
                            Next i
ErrNext:
                            lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                        End If
                    Next intCOl
                End If
            End If
        Next intRow
        '��Ƥ�Խ���������
        arrData = Array()
        lngTestisHeight = 0
        For i = 0 To UBound(arrTestis)
            '˵����Ƥ�Խ��
            If Val(arrTestis(i)) >= Val(Format(T_DrawClient.ʱ���е�λ * T_TwipsPerPixel.Y, "#0")) Then
                ReDim Preserve arrData(UBound(arrData) + 1)
                arrData(UBound(arrData)) = Val(Format(Val(arrTestis(i)) / T_TwipsPerPixel.Y, "#0"))
                lngTestisHeight = lngTestisHeight + Val(arrData(UBound(arrData)))
            End If
        Next i
        lngX1 = lngCurX
        lngY1 = lngCurY
        For i = 0 To UBound(arrData)
            If i = 0 Then
                Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, T_DrawClient.��������.Right, lngCurY, T_DrawClient.��������.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDC, lngX1, lngCurY + lngTestisHeight, T_DrawClient.�̶�����.Right, lngCurY + lngTestisHeight, PS_SOLID, IIf(intRowCount + UBound(arrData) + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
            End If
            For intCOl = 0 To T_BodyStyle.lng���� - 1
                If intCOl = 0 Then
                    Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1 + Val(arrData(i)), T_DrawClient.��������.Right, lngY1 + Val(arrData(i)), PS_SOLID, IIf(intRowCount + i + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
                Else
                    lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * T_BodyStyle.lng������) * intCOl
                    Call DrawLine(lngDC, lngX1, lngY1, lngX1, lngY1 + Val(arrData(i)), PS_SOLID, intBold, RGB_BLACK)
                End If
            Next intCOl
            lngY1 = lngY1 + Val(arrData(i))
        Next i
        
        lngCurY = lngCurY + lngTestisHeight
        intRowCount = intRowCount + UBound(arrData) + 1
        
        '������
        If intRepairRows > 0 And intRepairRows > intRowCount Then
            intRowCount = intRowCount + 1
            For intRow = intRowCount To intRepairRows
                lngX1 = lngCurX
                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                
                '�ո�ÿ��1��
                For intCOl = 0 To T_BodyStyle.lng����
                    If intCOl = 0 Then
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Call DrawLine(lngDC, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                    Else
                        
                        lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * T_BodyStyle.lng������) * intCOl
                        Call DrawLine(lngDC, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        If intCOl = T_BodyStyle.lng���� Then
                            Call DrawLine(lngDC, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        End If
                    End If
                Next intCOl
                lngCurY = lngY1
            Next intRow
        End If
        
        lngOutY = lngCurY + 2 * msngTwips
    Else
        lngOutY = lngCurY + 2 * msngTwips
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawBodyPageFooterNew(ByVal lngDC As Long, objDraw As Object, X As Long, Y As Long, ByVal LeftX As Long, ByVal intPageNo As Integer, _
    ByVal intBeginPage As Integer, Optional ByVal strInfo As String, Optional ByVal sngScale As Single = 1)
    '--------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�������ײ�˵��
    '����:intPageNO=ҳ��
    '--------------------------------------------------------------------------------------------------------------------------------
    Dim blnWeek As Boolean
    Dim blnPageNo As Boolean
    Dim blnOper As Boolean
    Dim blnPrintCurveInfo As Boolean
    Dim strNOPage As String
    Dim lngX As Long
    Dim blnPrinter As Boolean
    
    blnPrinter = False
    If TypeName(objDraw) = "Printer" Then
        blnPrinter = True
    Else
        msngTwips = 1
    End If
    blnPrintCurveInfo = (Val(zlDatabase.GetPara("���µ�����ӡ����˵��", glngSys, 1255, "0")) = 1)
    If blnPrintCurveInfo = False Then
        '��ӡ����˵����Ϣ
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strInfo, Len(strInfo), T_Size)
        Call GetTextRect(objDraw, X, Y, strInfo, 0, False, , sngScale)
        Call DrawText(lngDC, strInfo, -1, T_LableRect, DT_CENTER)
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 14
    Else
        Y = Y + IIf(blnPrinter = True, msngTwips, 1) * 6
    End If
    
    blnWeek = (Val(zlDatabase.GetPara("��ӡ����", glngSys, 1255, "0")) = 1)
    blnPageNo = (Val(zlDatabase.GetPara("��ӡҳ��", glngSys, 1255, "1")) = 1)
    '67405:������,2013-11-25,���"��ӡ��ӡ��"
    blnOper = (Val(zlDatabase.GetPara("��ӡ��ӡ��", glngSys, 1255, "0")) = 1)
    
    '��ӡҳ��
    '------------------------------------------------------------------------------------------------------------------
    If intPageNo > -1 And blnPageNo Then
        intPageNo = intPageNo + intBeginPage - 1
        strNOPage = "��   " & CStr(intPageNo) & "   ҳ"
    End If
    
    If blnWeek Then
        If strNOPage = "" Then
            strNOPage = "��   " & CStr(intBeginPage) & "   ��"
        Else
            strNOPage = strNOPage & "(�� " & CStr(intBeginPage) & " ��)"
        End If
    End If
    
    Call SetTextColor(lngDC, RGB_BLACK)
    Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
    Call GetTextRect(objDraw, 0, Y, strNOPage, objDraw.Width / T_TwipsPerPixel.X, False, , sngScale)
    Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    
    '�����ӡ��,����ǰ����Ա����
    '------------------------------------------------------------------------------------------------------------------
    If blnOper = True Then
        strNOPage = "��ӡ��:" & gstrUserName
    
        Call SetTextColor(lngDC, RGB_BLACK)
        Call GetTextExtentPoint32(lngDC, strNOPage, Len(strNOPage), T_Size)
        Call GetTextRect(objDraw, LeftX - objDraw.TextWidth(strNOPage) / T_TwipsPerPixel.X, Y, strNOPage, 0, False, , sngScale)
        Call DrawText(lngDC, strNOPage, -1, T_LableRect, DT_CENTER)
    End If

    Y = Y + T_Size.H / 2
    '--------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub DrawDeviceCapsNew(ByVal lngDC As Long, ByVal objDraw As Object)
    Dim dblSureW As Double, dblSureH As Double
    '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
    dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
    dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    On Error Resume Next
    Call DrawRect(lngDC, (objDraw.Width * dblSureW) / T_TwipsPerPixel.X, (objDraw.Height * (1 - dblSureH)) / T_TwipsPerPixel.Y, _
    (objDraw.Width * (1 - dblSureW)) / T_TwipsPerPixel.X, objDraw.Height * dblSureH / T_TwipsPerPixel.Y, PS_DOT, 1, RGB_FleetGRAY)
End Sub

Private Sub CloseRs(RS As ADODB.Recordset)
    '���ܣ��ر�Recordset����
    On Error Resume Next
    If RS.State = ADODB.adStateOpen Then RS.Close
    Set RS = Nothing
End Sub

Private Sub ErrEmpty()
    msngTwips = 1
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
End Sub

Public Function GetFontSize(ByVal objDraw As Object, ByVal strTmp As String, sinWidth As Single, Optional sngScale As Single = 1) As Single
'---------------------------------------------------
'���� ������±���������
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, sgnSize As Single
    Dim stdSet As StdFont
    Dim sngD As Single
    Dim blnChage As Boolean
    Dim arrText, blnGrade As Boolean
    
    On Error GoTo Errhand
    blnChage = False
    
    sgnSize = 9
    objDraw.Font.Size = sgnSize * sngScale
    objDraw.Font.Name = "����"
    objDraw.Font.Bold = False
    objDraw.Font.Italic = False
    
    If objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X > sinWidth Then
ErrGoTo:
        sngD = Round((objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X - sinWidth) / sinWidth, 4)
        If sngD > 0 Then
            sgnSize = CInt(Round((1 - sngD), 2) * sgnSize - 0.5)
            If sgnSize < 7 Then sgnSize = 7
            objDraw.Font.Size = sgnSize * sngScale
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > sinWidth And sgnSize > 7 Then GoTo ErrGoTo
        End If
    Else
        sgnSize = 9
    End If
    
    objDraw.Font.Size = sgnSize * sngScale
    
    GetFontSize = sgnSize
    Exit Function
Errhand:
    objDraw.Font.Size = 9 * sngScale
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawTabTextNew(ByVal lngDC As Long, ByVal objDraw As Object, ByVal strTmp As String, ByVal nCount As Long, ByVal wFormat As Long, LPoint As T_LPoint, Optional sngScale As Single = 1, Optional ByVal bytCenterType As Byte = 2, Optional sgnFontSize As Single = 9)
'---------------------------------------------------
'���� ������±���������
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, sgnSize As Single
    Dim stdSet As StdFont
    Dim sngD As Single
    Dim blnChage As Boolean
    Dim arrText, blnGrade As Boolean
    Dim arrData, i As Integer
    Dim lngFontHeight As Long
    Dim lngCurX As Long, lngCurY As Long
    
    On Error GoTo Errhand
    blnChage = False
    
    sgnSize = sgnFontSize
    objDraw.Font.Size = sgnSize * sngScale
    objDraw.Font.Name = "����"
    objDraw.Font.Bold = False
    objDraw.Font.Italic = False
    
    If objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X > LPoint.W Then
ErrGoTo:
        sngD = Round((objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X - LPoint.W) / LPoint.W, 4)
        If sngD > 0 Then
            sgnSize = Int(Round((1 - sngD), 2) * sgnSize - 0.5)
            If sgnSize < 7 Then sgnSize = 7
            objDraw.Font.Size = sgnSize * sngScale
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > LPoint.W And sgnSize > 7 Then GoTo ErrGoTo
            blnChage = True
        End If
    Else
        sgnSize = sgnFontSize
    End If
    
    arrData = Array()
    If blnChage = True Then
        With frmTendFileRead.txtLength
            .Width = Val(Format(LPoint.W * T_TwipsPerPixel.X, "#0")) + IIf(TypeName(objDraw) = "Printer", 12, 0)
            .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
            .FontName = "����"
            .FontSize = sgnSize * sngScale
            .FontBold = False
            .FontItalic = False
        End With
        arrData = GetData(frmTendFileRead.txtLength.Text, frmTendFileRead.txtLength)
        lngFontHeight = Val(Format((objDraw.TextHeight("��") / T_TwipsPerPixel.Y) * (UBound(arrData) + 1), "#0"))
    Else
        lngFontHeight = Val(Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0"))
    End If
    
    Set stdSet = New StdFont
    stdSet.Name = "����"
    stdSet.Size = sgnSize * sngScale
    stdSet.Bold = False
    stdSet.Italic = False
    Call SetFontIndirect(stdSet, lngDC, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDC, lngFont)
    
    Select Case bytCenterType
        Case 1 '����
            lngCurY = LPoint.Y
        Case 2 '����
            If lngFontHeight < LPoint.H Then
                lngCurY = LPoint.Y + (LPoint.H - lngFontHeight) / 2
            Else
                lngCurY = LPoint.Y
            End If
        Case 3 '����
            If lngFontHeight < LPoint.H Then
                lngCurY = LPoint.Y + (LPoint.H - lngFontHeight)
            Else
                lngCurY = LPoint.Y
            End If
    End Select
    lngCurX = LPoint.X
    
    '�������
    If UBound(arrData) > 0 Then
        For i = 0 To UBound(arrData)
            Call GetTextRect(objDraw, lngCurX, lngCurY, CStr(arrData(i)), , False, , sngScale)
            Call DrawText(lngDC, CStr(arrData(i)), nCount, T_LableRect, wFormat)
            lngCurY = lngCurY + Val(Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0"))
        Next i
    Else
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, LPoint.W, False, , sngScale)
        Call DrawText(lngDC, strTmp, nCount, T_LableRect, wFormat)
    End If
    
    Call SelectObject(lngDC, lngOldFont)
    Call DeleteObject(lngFont)
    Call ReleaseFontIndirect(objDraw)
    Set stdSet = Nothing
    Exit Sub
Errhand:
    objDraw.Font.Size = 9 * sngScale
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub DrawAnsyGrade(ByVal lngDC As Long, ByVal objDraw As Object, arrText() As String, LPoint As T_LPoint, ByVal lngColor As Long, Optional ByVal blnFormat As Boolean = False, Optional sngScale As Single = 1)
'---------------------------------------------------
'���� ���������
'˵�� AnsyGrade=True���ܵ��ô˺���
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer
    Dim stdSet As StdFont, stdOldset As StdFont
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    Dim lngMaxWidth As Long
    
    On Error GoTo Errhand
    
    If UBound(arrText) < 2 Then Exit Sub
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        '60529:������,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            strTmp = str1 & str2
        Else
            strTmp = str1 & str3
        End If
    Else
        strTmp = str1 & str2 & "/" & str3
    End If
    intSize = 9
    objDraw.Font.Size = 9 * sngScale
    Set stdSet = New StdFont
    stdSet.Name = "����"
    stdSet.Size = intSize * sngScale
    stdSet.Bold = False
    Set stdOldset = stdSet
    
    LPoint.Y = LPoint.Y + Val(Format(LPoint.H / 2, "#0"))
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , sngScale)
    '������
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + msngTwips
        Call ReleaseFontIndirect(objDraw)
    Else
        lngX = T_LableRect.Left
    End If
    
    If blnFormat = True Then '���ӷ�ĸ��ʾ
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        '60529:������,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            lngMaxWidth = objDraw.TextWidth(str2) / T_TwipsPerPixel.X
        Else
            lngMaxWidth = objDraw.TextWidth(str3) / T_TwipsPerPixel.X
        End If
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = intSize * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str2) / T_TwipsPerPixel.X) \ 2
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        If T_LableRect.Top < LPoint.Y - Val(Format(LPoint.H / 2, "#0")) Then T_LableRect.Top = LPoint.Y - Val(Format(LPoint.H / 2, "#0"))
        T_LableRect.Bottom = LPoint.Y + Val(Format(LPoint.H / 2, "#0"))
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '������
        objDraw.Font.Size = 9 * sngScale
        Call DrawLine(lngDC, lngX, lngY, lngX + lngMaxWidth, lngY)
        '�����ĸ
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        lngY = lngY
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str3) / T_TwipsPerPixel.X) \ 2
        T_LableRect.Top = lngY
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = intSize * sngScale
        Call SetFontIndirect(stdSet, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngColor)
        Call DrawText(lngDC, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
    Else
        If str1 <> "" Then
            '����ϱ�
            intSize = 7
            objDraw.Font.Size = intSize * sngScale
            Set stdSet = New StdFont
            stdSet.Name = "����"
            stdSet.Size = intSize * sngScale
            Call SetFontIndirect(stdSet, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < (LPoint.Y - Val(Format(LPoint.H / 2, "#0"))) Then T_LableRect.Top = (LPoint.Y - Val(Format(LPoint.H / 2, "#0")))
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            Call ReleaseFontIndirect(objDraw)
            '�����벿��
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDC, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        Else
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngColor)
            Call DrawText(lngDC, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        End If
    End If
    
    objDraw.Font.Size = 9 * sngScale
    Set stdSet = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Public Function GetXCoordinateNew(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln���� As Boolean = True) As String

    '����ʱ��õ�X��������X����ת��Ϊʱ�䷶Χ
    Dim SinX   As Single

    Dim intDO  As Integer, intMax As Integer

    Dim intDay As Integer, intTime As Integer

    Dim strDay As String, strTime As String
    
    Dim int������ As Integer

    On Error GoTo Errhand
    
    int������ = T_BodyStyle.lng������
    
    If bln���� Then
        '��һ����0,��������6
        strDay = Split(strInput, " ")(0)

        If InStr(1, strInput, " ") <> 0 Then
            strTime = Split(strInput, " ")(1)
        Else
            strTime = "00:00:00"
        End If

        intDay = DateDiff("d", CDate(strBeginDate), CDate(strInput))
        
        '�õ�����Ŀ̶�
        intMax = int������ - 1

        For intDO = 0 To intMax

            If strTime >= Split(gvarTime(intDO), ",")(0) And strTime <= Split(gvarTime(intDO), ",")(1) Then
                intTime = intDO
                Exit For
            End If
        Next
        
        '����õ�X����(ÿ��6��,������*�е�λ�õ�����)
        SinX = Format(T_DrawClient.��������.Left + (T_DrawClient.�е�λ * (intDay * int������ + intTime)), "#0.0")
        GetXCoordinateNew = SinX
    Else
        '����õ������ٸ��̶�
        SinX = Val(strInput)
        intTime = (SinX - T_DrawClient.��������.Left) \ T_DrawClient.�е�λ
        intDay = intTime \ int������
        intTime = intTime Mod int������
        
        strDay = Format(DateAdd("d", intDay, strBeginDate), "yyyy-MM-dd")
        strTime = gvarTime(intTime)
        GetXCoordinateNew = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    End If
    
    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetCurveDateNew(ByVal intCOl As Integer, _
                             ByVal dtBeginDateTime As Date, _
                             Optional ByVal intHourBegin As Integer = 4) As String

    '-------------------------------------------------------------------------------------
    '����:�����м����ʱ�䷶Χ
    '���� intCol ��ǰ��    dtBeginDateTime ��ʼʱ��
    '���ظ�ʽΪ:��ʼʱ��;��ֹʱ��
    '-------------------------------------------------------------------------------------
    Dim varTime  As Variant

    Dim intDays  As Integer

    Dim strBegin As String

    Dim strEnd   As String

    Dim lngLoop  As Long

    Dim lng�к�  As Long
    
    Dim int������ As Integer

    On Error GoTo Errhand
    
    GetCurveDateNew = -1
    
    int������ = T_BodyStyle.lng������
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin, int������, T_BodyStyle.lngʱ����)
    
    '���㵱ǰ�кͿ�ʼʱ�� ��������,�����¼����еĿ�ʼʱ��
    intDays = (intCOl - 1) \ int������
    strBegin = DateAdd("d", intDays, Int(dtBeginDateTime))
    strEnd = strBegin
    
    '���������ڵ�ʱ�䷶Χ
    lng�к� = (intCOl - 1) Mod int������
    
    strBegin = Format(strBegin & " " & Split(varTime(lng�к�), ",")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(strEnd & " " & Split(varTime(lng�к�), ",")(1), "YYYY-MM-DD HH:mm:ss")

    GetCurveDateNew = strBegin & ";" & strEnd

    Exit Function

Errhand:

    If ErrCenter = 1 Then

        Resume

    End If

End Function



Public Function GetCurveColumnNew(ByVal dtDateTime As Date, _
                               ByVal dtBeginDateTime As Date, _
                               Optional ByVal intHourBegin As Integer = 4) As Integer

    '******************************************************************************************************************
    '���ܣ� ��ʱ��������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim varTime As Variant

    Dim strTmp  As String

    Dim intDays As Integer

    Dim intLoop As Integer
    
    Dim int������ As Integer
    
    Dim int���� As Integer
    On Error GoTo Errhand
    
    GetCurveColumnNew = -1
    
    int������ = T_BodyStyle.lng������
    int���� = T_BodyStyle.lng����
    
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin, T_BodyStyle.lng������, T_BodyStyle.lngʱ����)

    '���㵱ǰ���ʱ������һ��ĵڼ���λ����
    strTmp = Format(dtDateTime, "HH:mm:ss")
    
    For intLoop = 0 To int������
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next
    
    If intLoop < int������ Then
        '���㵱���ڵ�ǰ���µ�ҳ���ǵڼ��죨0��ʾ��һ�죻1��ʾ�ڶ���.....��
        intDays = DateDiff("d", Int(dtBeginDateTime), Int(dtDateTime))
        GetCurveColumnNew = intDays * int������ + intLoop + 1
    End If
    
    Exit Function

Errhand:

    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CalcMinMaxColNew(ByVal strDate As String, _
                              MinCol As Integer, _
                              MaxCol As Integer) As Boolean

    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �����С���ʱ�䷶Χ
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String

    Dim dtTmp      As Date

    Dim strTmp     As String
    
    'If mvarEdit = False Then Exit Function
    
    aryValue = Split(strDate, ";")
    gintHourBegin = T_BodyStyle.lng��ʼʱ��
    MinCol = GetCurveColumnNew(CDate(aryValue(0)), CDate(aryValue(0)), gintHourBegin)
    MaxCol = GetCurveColumnNew(CDate(aryValue(1)), CDate(aryValue(0)), gintHourBegin)
    
End Function


Public Sub CreatePolyNew(rsPoint As ADODB.Recordset, ByVal objDraw As Object, ByVal lngDC As Long, ByVal strBeginDate As String, ByVal str�������� As String, ByVal bln�������� As Boolean)

'rsPoint ��¼�� �������  ��Ŀ���,X����,Y����
    Dim arrData, arrPt
    Dim bln���� As Boolean      '����������ǵ�Ե�,���ʱ����Ӧ���������γ����������
    Dim bln�� As Boolean, bln�� As Boolean, bln��ǰ As Boolean, bln�Ͽ� As Boolean, bln��Ч As Boolean
    Dim intDO   As Integer, intMax As Integer             'intLast��¼���һ����Ч������
    Dim recttmp As RECT, SinX As Single, sinY As Single, sin��X As Single, sin��X As Single
    Dim str��ǰ As String, str�� As String, str�� As String
    Dim str���� As String, str���� As String
    Dim PtIn����() As POINTAPI
    Dim PtIn����() As POINTAPI
    Dim lng��䷽ʽ As Long

    Dim PtInPoly() As POINTAPI, intCOl As Integer, intCols As Integer, intCount As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngWith As Long
    Dim i As Integer, j As Integer
    On Error GoTo Errhand

    '1�����ʶ�Ӧ1��3������,����������ÿһ�춼��ֵ,�����γ�����
    '�γɵ����򼯺ϱ�����������,����,��װ������,�ٵ���װ������,�γ�������һ������
    '�ɵ���ɵķ������,��DrawPoly����ɷ�����������
    
    lng��䷽ʽ = Val(zlDatabase.GetPara("���������䷽ʽ", glngSys, 1255, "0"))
    
    If TypeName(objDraw) = "Printer" Then
        intBold = 4
        intFine = 4
        blnPrinter = True
    Else
        intBold = 2
        intFine = 1
        blnPrinter = False
    End If
    
    rsPoint.Sort = "��Ŀ���,ʱ��"
    arrData = Split(str��������, ",")
    intMax = UBound(arrData)
    

'
    For intDO = 0 To intMax

        SinX = Val(Split(arrData(intDO), ";")(0))
        sinY = Val(Split(arrData(intDO), ";")(1))
        '����ǰ���ʼ������򼯺�
        intCount = intCount + 1
        ReDim Preserve PtInPoly(intCount)
        str���� = str���� & "," & SinX + T_DrawClient.�е�λ / 2 & ";" & sinY
        
        '��������,�������е���������
        If Not bln���� Then
            bln�� = False
            rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����<" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
               rsPoint.Sort = "X���� DESC"
                bln�Ͽ� = (rsPoint!�Ͽ� = 1)
                If Not bln�Ͽ� Then
                    rsPoint.Sort = "X���� DESC"
                    sin��X = rsPoint!X����
                
                    '���ݵ�ǰ�����ȡʱ��
                    str�� = GetXCoordinateNew(sin��X, strBeginDate, False)
                    str��ǰ = GetXCoordinateNew(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                    '��ǰ���ǰһʱ�����һ��û�����ݾͶϿ�
                    If DateDiff("d", CDate(Split(str��, ",")(0)), CDate(Split(str��ǰ, ",")(0))) < 2 Then
                        recttmp.Left = rsPoint!X����
                        recttmp.Top = rsPoint!Y����
                        '���������������򼯺�
                        intCount = intCount + 1
                        ReDim Preserve PtInPoly(intCount)
                        str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
                        bln�� = True
                    End If
                End If
            End If
        End If
        
        bln��ǰ = False
        'ȱʡ�Ǻ͵�ǰ�е���������
        rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����=" & Val(Split(arrData(intDO), ";")(0))
        bln��ǰ = (rsPoint.RecordCount <> 0)

        If bln��ǰ Then
            If Not bln�� Then
                recttmp.Left = rsPoint!X����
                recttmp.Top = rsPoint!Y����
            End If

            bln�Ͽ� = (rsPoint!�Ͽ� = 1)

            '����ǰ�����������򼯺�
            If Not bln���� Then
                intCount = intCount + 1
                ReDim Preserve PtInPoly(intCount)
                str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
            End If
        End If

        bln�� = False

        If Not bln�Ͽ� Then
            rsPoint.Filter = "��Ŀ���=" & gint���� & " And X����>" & Val(Split(arrData(intDO), ";")(0))
            
            If rsPoint.RecordCount <> 0 Then
                rsPoint.Sort = "X���� ASC"
                sin��X = rsPoint!X����
            
                '���ݵ�ǰ�����ȡʱ��
                str�� = GetXCoordinateNew(sin��X, strBeginDate, False)
                str��ǰ = GetXCoordinateNew(Val(Split(arrData(intDO), ";")(0)), strBeginDate, False)
                '��ǰ�����һʱ�����һ��û�����ݾͶϿ�
                If DateDiff("d", CDate(Split(str��ǰ, ",")(0)), CDate(Split(str��, ",")(0))) < 2 Then
                    bln�� = True
                    recttmp.Right = rsPoint!X����
                    recttmp.Bottom = rsPoint!Y����
                    '���������������򼯺�
                    intCount = intCount + 1
                    ReDim Preserve PtInPoly(intCount)
                    str���� = str���� & "," & rsPoint!X���� + T_DrawClient.�е�λ / 2 & ";" & rsPoint!Y����
                End If
            End If
        End If
        
        '�Ȱ���߷��
        If bln���� = False Then
            If bln��ǰ = True Then
                '�����л�ǰ�е���������
                Call DrawLine(lngDC, recttmp.Left + T_DrawClient.�е�λ / 2, recttmp.Top, SinX + T_DrawClient.�е�λ / 2, sinY, PS_SOLID, intFine, RGB_RED)
            End If

            bln���� = (bln�� Or bln��) And bln��ǰ
        End If
        
        '�ҵ��ұߵķ������������
        If bln���� Then
            bln���� = False

            If bln�� = True Then
                '�жϵ�ǰ���ʶ�Ӧ����һ����������һ������X�����Ƿ����,����Ⱦͷ������
                If intDO < intMax Then
                    If recttmp.Right = Val(Split(arrData(intDO + 1), ";")(0)) Then
                        bln���� = True
                    End If
                End If
            End If
            
            
            If Not bln���� Then
                '��֯����,��������ʼ,Ȼ��ת������(���ʴ����ʼ,�ٻص�֮ǰ������,�ٻص���һ������,�γɷ������)
                intCount = 1
                str���� = Mid(str����, 2)
                arrPt = Split(str����, ",")
                intCols = UBound(arrPt)
                i = 0
                ReDim Preserve PtIn����(intCols)
                For intCOl = 0 To intCols
                    PtIn����(i).X = Split(arrPt(intCOl), ";")(0)
                    PtIn����(i).Y = Split(arrPt(intCOl), ";")(1)
                    i = i + 1
                 Next
                
           
                For intCOl = 0 To intCols
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

                str���� = Mid(str����, 2)
                arrPt = Split(str����, ",")
                intCols = UBound(arrPt)
                
                i = 0
                ReDim Preserve PtIn����(intCols)
                For intCOl = 0 To intCols
                    PtIn����(i).X = Split(arrPt(intCOl), ";")(0)
                    PtIn����(i).Y = Split(arrPt(intCOl), ";")(1)
                    i = i + 1
                Next

                For intCOl = intCols To 0 Step -1
                    PtInPoly(intCount).X = Split(arrPt(intCOl), ";")(0)
                    PtInPoly(intCount).Y = Split(arrPt(intCOl), ";")(1)
                    intCount = intCount + 1
                Next

'                '��������γɷ������
                ReDim Preserve PtInPoly(intCount)
                PtInPoly(intCount).X = PtInPoly(1).X
                PtInPoly(intCount).Y = PtInPoly(1).Y
                
                '��������
                Call DrawPoly(lngDC, PtInPoly, lng��䷽ʽ, UBound(Split(str����, ",")) + 1)
                '76697,LPF,����66628��������Ĵ���
                '����ţ�66628,�޸��ˣ�����,�������Ṳ��ʱ�����ͼ�Σ�ֱ�����������ʵ�������ߡ�
                If lng��䷽ʽ = 2 And bln�������� Then
                    i = 0: j = 0
                    For i = 0 To UBound(PtIn����)
                        For j = 0 To UBound(PtIn����)
                            If PtIn����(j).X = PtIn����(i).X Then
                                Call DrawLine(lngDC, PtIn����(j).X, PtIn����(j).Y, PtIn����(i).X, PtIn����(i).Y, PS_SOLID, intFine, RGB_RED)
                            End If
                        Next
                    Next
                End If
            End If
          
        End If

        If Not bln���� Then
            intCount = 0
            str���� = ""
            str���� = ""
            ReDim Preserve PtInPoly(intCount)
            ReDim Preserve PtIn����(intCount)
            ReDim Preserve PtIn����(intCount)
        End If

    Next
    
    rsPoint.Filter = ""

    Exit Sub

Errhand:

    If ErrCenter() = 1 Then

        Resume

    End If

End Sub


Public Function GetCanvasCenterNew(ByVal dtBegin As Date, ByVal dtEnd As Date, ByVal dtBeginDate As Date, ByVal SinX As Single) As Boolean
'---------------------------------------------------------
'����:�жϸ�ʱ����Ƿ����м�ֵ
'����:dtbegin:���Ƚϵ�ʱ���.  dtend:Ҫ�Ƚϵ�ʱ��� . dtBeginDate ��ҳ���µ��Ŀ�ʼʱ�� .sinx��ǰ���X����
'---------------------------------------------------------
    Dim blnTrue As Boolean
    Dim strTime As String, strTmp As String
    Dim intDay As Integer, intTime As Integer, strDay As String
    Dim int������ As Integer
    Dim intʱ���� As Integer
    
    int������ = T_BodyStyle.lng������
    intʱ���� = T_BodyStyle.lngʱ����
    
    intTime = (SinX - T_DrawClient.��������.Left) \ T_DrawClient.�е�λ
    intDay = intTime \ int������
    intTime = intTime Mod int������
        
    strDay = Format(DateAdd("d", intDay, dtBeginDate), "yyyy-MM-dd")
    strTmp = strDay & " " & Split(gvarTime(intTime), ",")(0) & "," & strDay & " " & Split(gvarTime(intTime), ",")(1)
    
    If intTime <= UBound(gvarTime) Then
        If gintHourBegin + intTime * intʱ���� = 24 Then
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(strDay, "YYYY-MM-DD") & " " & gintHourBegin + intTime * 4 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    If CDate(strTime) > CDate(Split(strTmp, ",")(1)) Then strTime = Format(Split(strTmp, ",")(1), "YYYY-MM-DD HH:mm:ss")
    
    If Abs(DateDiff("s", Format(dtBegin, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) > _
        Abs(DateDiff("s", Format(dtEnd, "YYYY-MM-DD HH:mm:ss"), Format(strTime, "YYYY-MM-DD HH:mm:ss"))) Then
        blnTrue = True
    End If

    GetCanvasCenterNew = blnTrue
End Function

Public Function RetrunEndTimeNew(ByVal dtBegin As Date, ByVal dtEnd As Date, Optional ByVal intHourBegin As Integer = 4) As Date
'**********************************************************************************
'���ܣ�������µ���ֹʱ��Ϳ�ʼʱ���Ƿ���ͬһ��Ԫ�������ͬһ��Ԫ����Ҫ����ֹʱ���Ƶ���һ��Ԫ��
'������strBegin ���µ���ʼʱ��,strEnd ���µ���ֹʱ��(���˳�Ժʱ��)
'����ֵ�����µ���ֹʱ��
'**********************************************************************************
'���󣺶��ڲ��˳�Ժ����Ժʱ����ͬһ�����ӣ���Ҫ¼����Ժ���£�ҲҪ¼���Ժ���£�����Ժ����¼�뵽��һ�����ӡ�

    Dim varTime As Variant
    Dim intLoop As Integer, strTmp As String
    Dim intBegin As Integer, intEnd As Integer
    Dim strEnd As String
    Dim int������ As Integer
    Dim intʱ���� As Integer
    
    int������ = T_BodyStyle.lng������
    intʱ���� = T_BodyStyle.lngʱ����
    RetrunEndTimeNew = dtEnd
    If Format(dtBegin, "YYYY-MM-DD") <> Format(dtEnd, "YYYY-MM-DD") Then Exit Function
    '��ʼ��ʱ�䷶Χ����
    Call InitDateTimeRange(varTime, intHourBegin, int������, intʱ����)
    '1/���㿪ʼʱ�����ֹʱ���ڵڼ�������
    strTmp = Format(dtBegin, "HH:mm:ss")
    For intLoop = 0 To int������
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intBegin = intLoop
            Exit For
        End If
    Next
    strTmp = Format(dtEnd, "HH:mm:ss")
    For intLoop = 0 To int������
        If Format(strTmp, "HH:mm:ss") >= Format(Split(varTime(intLoop), ",")(0), "HH:mm:ss") And Format(strTmp, "HH:mm:ss") <= Format(Split(varTime(intLoop), ",")(1), "HH:mm:ss") Then
            intEnd = intLoop
            Exit For
        End If
    Next
    '2 ����ͬһ�о��˳�
    If intBegin <> intEnd Then Exit Function
    If intEnd > int������ - 1 Then Exit Function
    '3 �����ֹʱ������¸�ֵ
    If intEnd > int������ - 2 Then
        strEnd = Format(DateAdd("D", 1, dtEnd), "YYYY-MM-DD") & " " & Format(Split(varTime(0), ",")(1), "HH:mm:ss")
    Else
        strEnd = Format(dtEnd, "YYYY-MM-DD") & " " & Format(Split(varTime(intEnd + 1), ",")(1), "HH:mm:ss")
    End If
    
    RetrunEndTimeNew = CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss"))
End Function



Public Sub OutPutTextNew(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal lngDC As Long, ByVal rsNote As ADODB.Recordset, ByVal mstrBeginDate As String, Optional ByVal sngScale As Single = 1)

    'rsDrawItems  ��¼��Ŀ��������� ��λֵ�Ȼ�����Ϣ
    'rsNote ����˵����Ϣ
    'mstrBeginDate ���µ�ÿҳ��ʼʱ��
    '���������Ϣ:��Ժ,���,ת��,��Ժ,��������,δ��˵��,�ϱ�˵��������
    'δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ
    '��δ��˵�����ϱ�˵����,���ת����Ϣ��һ���̶ȷ������ʱ,����д������̶���,�������̶�Ҳ����Ϣ,˳��
    Dim lngMaxX As Long     '���µ����X����
    Dim lngX    As Long '��һ�е�X����
    Dim lngY    As Long 'Y����
    Dim lngY1   As Long '40 �ȹ̶�����
    Dim i       As Integer, j As Integer
    Dim X, Y As Long '�������ʱ������
    Dim strComment    As String, strText As String
    Dim intAscCharNum As Integer
    Dim rsTemp  As New ADODB.Recordset
    Dim strDate As String
    Dim bln�ϱ� As Boolean
    Dim bln�¼���ʾ���� As Boolean '����:���±�־��˳��������
    Dim blnLessenSize As Boolean  '����:���±�־����42�̶���С������ʾ
    Dim arrX, arrCurX
    Dim blnBigSize As Boolean '�Ƿ��Ծź�������ʾ
    Dim lngFont As Long, lngOldFont As Long
    Dim dblCurveHeight As Double  '���µ�42��40�ĸ߶�
    Dim dblHeight As Double
    Dim blnCenter As Boolean
    
    On Error GoTo Errhand
    
    arrX = Array(): arrCurX = Array()
    bln�¼���ʾ���� = (Val(zlDatabase.GetPara("���±�־��˳��������", glngSys, 1255, 0)) = 1)
    blnLessenSize = (Val(zlDatabase.GetPara("���±�־����40�̶���С������ʾ", glngSys, 1255, 0)) = 1)
    
    lngMaxX = T_DrawClient.��������.Right - T_DrawClient.�е�λ
    dblCurveHeight = Format(GetYCoordinate(objDraw, rsDrawItems, gint����, 40) - GetYCoordinate(objDraw, rsDrawItems, gint����, 42), "#0.00")
    
    rsNote.Filter = "����<>1"

    '���ȼ��������ת������������Ϣ
    If rsNote.RecordCount = 0 Then Exit Sub
    
    '70228:������,2014-02-18,�����Զ���ʶ��ʾ�޸ġ�
    '����
    '   1�����±�־��˳��������=True��ÿҳ��ѭ������������ʾ�����ǣ�һ��ʱ�������ʾ�������(��С���崦��).����ڵ�����ʾ���꣬ʣ���ǲ�������ʾ��
    '   2�����±�־��˳��������=False,ÿҳ��˳������������ʾ��һ��ʱ��ֻ��ʾһ��������ڱ�ҳ���һ�л���ʾ���꣬�������һ��������ʾʣ���ǡ�
    rsNote.Sort = "X����,ʱ��,��Ŀ���"
    lngX = rsNote!X����
    j = 1
    With rsNote
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!����)) <> "" Then
                If Not (!���� = 2 Or !���� = 99) Then
                    '���±�־��˳��������
                    If bln�¼���ʾ���� = True Then
                        If Val(!X����) > lngX Then j = 1
                        If lngX <= lngMaxX Then
                            strDate = Format(Split(GetXCoordinateNew(lngX, mstrBeginDate, False), ",")(0), "YYYY-MM-DD")
                            If CDate(strDate) > CDate(Format(!ʱ��, "YYYY-MM-DD")) Then
                                !���� = 1
                            End If
                        Else
                            lngX = lngMaxX
                            !���� = 1
                        End If
                    Else
                        '����x���꣬��������������x���꣬�����У��
                        If lngX > lngMaxX Then lngX = lngMaxX
                    End If
                    
                    !��ӡX���� = IIf(lngX <= Val(!X����), !X����, lngX)
                    !�߶� = GetFontHeight(objDraw, zlCommFun.Nvl(!����))
                    .Update
                    
                    If lngX <= !X���� Then lngX = !X����
                    
                    '70228:ĳ�д��ڶ����ǣ������ʾ����(����X����)
                    If Not (bln�¼���ʾ���� = True And j Mod 2 = 1) Then
                        ReDim Preserve arrX(UBound(arrX) + 1)
                        arrX(UBound(arrX)) = lngX
                        lngX = lngX + T_DrawClient.�е�λ
                        j = 0
                    End If
                    If bln�¼���ʾ���� = True Then j = j + 1
                Else
                    !�߶� = GetFontHeight(objDraw, zlCommFun.Nvl(!����))
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        '���������Զ���־�ĸ߶�
        '�������һ��Ҫ���������־������С���塣�������Ƿ�ѡ�˲���"���±�־����40�̶���С������ʾ"����ѡ����С����
        .Filter = "����<>1"
        .Sort = "X����,ʱ��,��Ŀ���"
        Do While Not .EOF
            If Not (!���� = 2 Or !���� = 99) Then
                blnBigSize = True
                If bln�¼���ʾ���� = True Then
                    For i = 0 To UBound(arrX)
                        If Val(arrX(i)) = Val(Nvl(!��ӡX����)) Then
                            blnBigSize = False
                            Exit For
                        End If
                    Next i
                End If
                
                If blnBigSize = True And blnLessenSize = True Then
                    If GetFontHeight(objDraw, zlCommFun.Nvl(!����)) > dblCurveHeight Then
                        blnBigSize = False
                    End If
                End If
                
                If blnBigSize = False Then
                    gstdSet.Name = "����"
                    gstdSet.Size = 7.5
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDC, lngFont)
                    dblHeight = GetFontHeight(objDraw, zlCommFun.Nvl(!����))
                    Call SelectObject(lngDC, lngOldFont)
                    Call DeleteObject(lngFont)
                    '��ԭ����
                    gstdSet.Name = "����"
                    gstdSet.Size = 9
                    Call SetFontIndirect(gstdSet, lngDC, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    Call SelectObject(lngDC, lngFont)
                    !�߶� = dblHeight
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        lngY = GetYCoordinate(objDraw, rsDrawItems, gint����, 42)
        '�������ת ���������䵽�����X�����ж���ʽ��Y����
        .Filter = "��ӡX����=" & lngMaxX & " And ����<>1"
        .Sort = "ʱ��,��Ŀ���"

        Do While Not .EOF
            !Y���� = lngY
            .Update
            lngY = lngY + Val(!�߶�) + T_DrawClient.�е�λ
            .MoveNext
        Loop
        
        '����δ��˵�����ϱ����ʾλ��(Y����).
        '˵��:��û�����ת��������Ϣ������� ��ӡ�� 42-40��֮�䣬�����ӡ��40�����´�ӡ
        .Filter = "����<>1"
        .MoveFirst
        .Sort = "X����,ʱ��,��Ŀ���"
        Set rsTemp = .Clone

        Do While Not .EOF
            If (!���� = 2 Or !���� = 99) Then
                
                rsTemp.Filter = "(��ӡX����=" & !X���� & " And ����<>1 and ����=99) or (��ӡX����=" & !X���� & " And ����<>1 and ����=2)"
                
                If rsTemp.BOF Then
                    rsTemp.Filter = "��ӡX����=" & !X���� & " And ����<>1"
                End If
                
                If rsTemp.RecordCount > 0 Then
                    bln�ϱ� = False
                    lngY = 0
                    Do While Not rsTemp.EOF
                        If bln�ϱ� = False Then
                            bln�ϱ� = IIf(rsTemp!���� = 2 Or rsTemp!���� = 99, True, False)
                            lngY1 = Val(rsTemp!Y����)
                        End If
                        
                        If lngY < lngY1 + rsTemp!�߶� + T_DrawClient.�е�λ Then lngY = lngY1 + rsTemp!�߶� + T_DrawClient.�е�λ
                        lngY1 = lngY
                        
                        rsTemp.MoveNext
                    Loop
                    
                    lngY1 = GetYCoordinate(objDraw, rsDrawItems, gint����, 40)

                    If lngY > lngY1 Or bln�ϱ� Then lngY1 = lngY
                    
                Else '�������κ���Ϣ ��42��ʼ��ӡ
                    lngY1 = Val(!Y����)
                End If
                
                !Y���� = lngY1
                !��ӡX���� = !X����
                .Update
            End If

            .MoveNext
        Loop
        
        '70228:����һ����ʾ������ǵĴ�ӡX���꣬��δ��������ڴ����ϱ���ʾλ�õĺ���
        '��������Ϊ7.5
        If bln�¼���ʾ���� = True Then
            gstdSet.Name = "����"
            gstdSet.Size = 7.5
            objDraw.Font.Name = gstdSet.Name
            objDraw.Font.Size = gstdSet.Size
            
            For i = 0 To UBound(arrX)
                .Filter = "��ӡX����=" & Val(arrX(i)) & " And ����<>1 And ����<>2 And ����<>99"
                .Sort = "X����,ʱ��,��Ŀ���"
                If .RecordCount > 1 Then
                    lngX = !��ӡX���� - Abs(T_DrawClient.�е�λ - (objDraw.TextWidth("��") / T_TwipsPerPixel.X) * 2) / 2
                    !��ӡX���� = lngX
                    .Update
                    ReDim Preserve arrCurX(UBound(arrCurX) + 1)
                    arrCurX(UBound(arrCurX)) = !���� & "," & !��Ŀ��� & "," & !��ӡX���� & "," & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    .MoveNext
                    !��ӡX���� = lngX + objDraw.TextWidth("��") / T_TwipsPerPixel.X
                    .Update
                    ReDim Preserve arrCurX(UBound(arrCurX) + 1)
                    arrCurX(UBound(arrCurX)) = !���� & "," & !��Ŀ��� & "," & !��ӡX���� & "," & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                End If
            Next i
            '��ԭ����Ϊ9������
            gstdSet.Name = "����"
            gstdSet.Size = 9
            objDraw.Font.Name = gstdSet.Name
            objDraw.Font.Size = gstdSet.Size
        End If
        '��ʼ�������
        .Filter = "����<>1"
        .MoveFirst
        .Sort = "X����,ʱ��,��Ŀ���"
        Dim sigNum As Single
        Do While Not .EOF
            '�������
            strComment = Trim(zlCommFun.Nvl(!����))

            If strComment <> "" Then
                X = Val(IIf(Trim(!��ӡX����) <> "", !��ӡX����, !X����))
                Y = Val(!Y����)
                intAscCharNum = 0
                
                '70228:һ����ʾ������ǽ���������С����
                blnBigSize = True
                blnCenter = True
                If bln�¼���ʾ���� = True Then
                    For i = 0 To UBound(arrCurX)
                        If !���� & "," & !��Ŀ��� & "," & !��ӡX���� & "," & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss") = CStr(arrCurX(i)) Then
                            blnBigSize = False
                            Exit For
                        End If
                    Next i
                End If
                blnCenter = blnBigSize
                '���һ��ֻ��һ����ǣ����ұ�����ݳ���40�̶ȣ�����С���塣
                If blnBigSize = True And blnLessenSize = True And Not (!���� = 2 Or !���� = 99) Then
                    If GetFontHeight(objDraw, strComment) > dblCurveHeight Then
                        blnBigSize = False
                    End If
                End If
                
                gstdSet.Name = "����"
                gstdSet.Size = IIf(blnBigSize = True, 9, 7.5)
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDC, lngFont)
                T_Size.H = objDraw.TextHeight("1") / T_TwipsPerPixel.Y
                    
                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.�̶�����.Bottom Then
                        strText = Mid(strComment, i, 1)
                        
                        If Asc(strText) < 0 Then
                            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If

                        '���������Ϣ
                        Call DrawRotateText(objDraw, lngDC, X, Y, strText, !��ɫ, sngScale, IIf(blnCenter = True, -999, objDraw.TextWidth("��") / T_TwipsPerPixel.X))

                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
                Call SelectObject(lngDC, lngOldFont)
                Call DeleteObject(lngFont)
                
                gstdSet.Name = "����"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, lngDC, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                Call SelectObject(lngDC, lngFont)
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub





Public Function GetAppendGridItemNew(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal int����ȼ� As Integer, ByVal intӤ�� As Long, dt��ʼʱ�� As Date, dt����ʱ�� As Date, ByVal byt���ò��� As Byte, ByVal lng����ID As Long, ByVal str�����Ŀ As String, Optional blnMove As Boolean = False) As ADODB.Recordset
    '**************************************************************************
    '����:��ȡ������ݵ����±����Ŀ�Լ��̶������Ŀ
    '**************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo Errhand
    
    Set rsTemp = GetGridItem(int����ȼ�, byt���ò���, lng����ID, 2)
    If rsTemp.RecordCount = 0 Then
        '�����ڻ��Ŀֱ����ȡ�̶������Ŀ
        Set rsTemp = GetGridItemNew(str�����Ŀ)
        Set GetAppendGridItemNew = rsTemp
        Exit Function
    End If
    With rsTemp
        Do While Not .EOF
            strSQL = IIf(strSQL = "", "select " & !��Ŀ��� & " ��Ŀ��� from dual", strSQL & " UNION ALL select " & !��Ŀ��� & "  ��Ŀ��� from dual ")
            .MoveNext
        Loop
    End With
    
    strSQL = "(" & strSQL & ") F"
    '��ȡ���Ŀ
    gstrSQL = "Select distinct D.�������,D.��Ŀ���,C.���²�λ,C.���²�λ || D.��¼��  ��¼��,D.��¼��,D.��¼��,D.��¼ɫ,D.���ֵ,D.��Сֵ,D.��λֵ,nvl(D.��¼Ƶ��,2) ��¼Ƶ��,D.��Ժ�ײ�," & _
        "   E.��Ŀ����,E.������,E.��Ŀֵ��,E.��Ŀ��ʾ,E.��Ŀ����,E.��Ŀ����,E.��ĿС��,E.��Ŀ��λ ��λ" & _
        "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E," & strSQL & _
        "   Where  B.ID=A.�ļ�ID And A.ID = c.��¼ID  AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[5]  AND B.����id=[2]    AND B.��ҳid=[3] AND d.��Ŀ���=C.��Ŀ��� " & _
        "   AND c.��¼����=1 And E.��Ŀ����=2  AND E.��Ŀ���=D.��Ŀ���  AND E.����ȼ�>=[4]   AND a.����ʱ�� BETWEEN [6] And [7] And c.��ֹ�汾 Is Null " & _
        "   AND d.��¼��=2 and D.��Ŀ���=F.��Ŀ���"
    
    '��ȡ�̶������Ŀ
    strSQL = "Select A.�������,A.��Ŀ���,'' ���²�λ,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,nvl(D.C2,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C,TABLE(CAST(F_NUM2LIST2([8]) AS ZLTOOLS.T_NUMLIST2)) D" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And B.��Ŀ���=D.C1 And A.��¼��=2 And B.��Ŀ����=1"



    
    gstrSQL = "Select /*+ Rule*/ �������,��Ŀ���,���²�λ,��¼��,��¼��,��¼��,��¼ɫ,���ֵ,��Сֵ,��λֵ,��¼Ƶ��,��Ժ�ײ�,��Ŀ����," & _
        "   ������,��Ŀֵ��,��Ŀ��ʾ,��Ŀ����,��Ŀ����,��ĿС��,��λ" & _
        "   From (" & gstrSQL & vbCrLf & " UNION ALL " & vbCrLf & strSQL & ") order by Decode(��Ŀ���,3 ,0,1 ),�������,��¼��"
    If blnMove Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "", lng�ļ�ID, lng����ID, lng��ҳID, int����ȼ�, intӤ��, CDate(Format(dt��ʼʱ��, "yyyy-mm-dd hh:mm:ss")), CDate(Format(dt����ʱ��, "yyyy-mm-dd hh:mm:ss")), str�����Ŀ)
    
    Set GetAppendGridItemNew = rsTemp

    Exit Function

Errhand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetGridItemNew(ByVal str�����Ŀ As String) As ADODB.Recordset

    '**********************************************************************************
    '����:��ȡר�����±����Ŀ
    '**********************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo Errhand
    
   '��ȡ�����Ŀ
   gstrSQL = "Select A.�������,A.��Ŀ���,'' ���²�λ,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,nvl(D.C2,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C,TABLE(CAST(F_NUM2LIST2([1]) AS ZLTOOLS.T_NUMLIST2)) D" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And B.��Ŀ���=D.C1 And A.��¼��=2 And B.��Ŀ����=1" & _
        "   order by Decode(��Ŀ���,3 ,0,1 ),�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�̶����±����Ŀ", str�����Ŀ)
    Set GetGridItemNew = rsTemp

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetRows(bln���� As Boolean, ByVal strValue As String) As Long
    Dim strOld() As String
    Dim intRow As Integer, i As Integer
    Dim intRows As Integer
    strOld = Split(strValue, ",")
    For i = 0 To UBound(strOld)
        If InStr(1, strOld(i), ":") > 0 Then
            If Split(strOld(i), ":")(0) = 3 Then
                bln���� = True
                intRows = intRows + 1
            Else
                If Split(strOld(i), ":")(0) <> 5 Then
                    Select Case Split(strOld(i), ":")(1)
                        Case 1
                            intRow = 1
                        Case 2
                            intRow = 1
                        Case 3
                            intRow = 3
                        Case 4
                            intRow = 2
                        Case 6
                            intRow = 1
                    End Select
                    intRows = intRows + intRow
                End If
            End If
        End If
    Next
    GetRows = intRows
End Function


Private Function getSQLString(ByVal strText As String, ByVal blnMoved As Boolean, Optional ByVal strItems As String) As String
    Dim strNewSql As String
    Dim strSQL As String
    Dim strSQLText As String
    Dim lngColor As Long
    Select Case strText
        Case "��ȡ�ļ�ʱ�䷶Χ"

             strNewSql = "   (SELECT ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0  AND C.��� = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.�������� = COLUMN_VALUE) And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
            '��ȡ���˳�Ժǰ��ʱ����Ϣ
            '------------------------------------------------------------------------------------------------------------------
            strSQL = " SELECT /*+ RULE */ DECODE(D.��ʼʱ��,NULL,DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��)," & vbNewLine & _
                "               DECODE(SIGN(D.��ʼʱ�� - DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))," & vbNewLine & _
                "                      1," & vbNewLine & _
                "                      D.��ʼʱ��," & vbNewLine & _
                "                      DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))) AS ��ʼ," & vbNewLine & _
                "       DECODE(D.����ʱ��," & vbNewLine & _
                "               NULL," & vbNewLine & _
                "               DECODE(E.��¼," & vbNewLine & _
                "                      0," & vbNewLine & _
                "                      DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, NVL(E.Ӥ��ʱ��, A.��ֹ), D.����ʱ��)," & vbNewLine & _
                "                      NVL(E.Ӥ��ʱ��, A.��ֹ))," & vbNewLine & _
                "               DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, D.����ʱ��, NVL(E.Ӥ��ʱ��, A.��ֹ))) ��ֹ," & vbNewLine & _
                "       DECODE(D.����ʱ��, NULL, E.��¼, 1) ��¼" & vbNewLine & _
                " FROM (SELECT ����ID, ��ҳID, MIN(��ʼʱ��) AS ��ʼ, MAX(NVL(��ֹʱ��, SYSDATE)) AS ��ֹ" & vbNewLine & _
                "       FROM ���˱䶯��¼" & vbNewLine & _
                "       WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
                "       GROUP BY ����ID, ��ҳID) A," & vbNewLine & _
                "     (SELECT ����ID, ��ҳID, ����ʱ�� FROM ������������¼ WHERE ����ID = [2] AND ��ҳID = [3] AND ��� = [4]) B," & vbNewLine & _
                "     (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
                "       FROM (SELECT MAX(B.����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
                "              FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
                "              WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND A.Ӥ�� = [4])) D," & vbNewLine & _
                "  " & strNewSql & vbNewLine & _
                " WHERE A.����ID = E.����ID AND A.��ҳID = E.��ҳID AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)"
            
            strSQLText = strSQL
        
        Case "��ȡ����������Ŀ"
            
            strSQL = " Select A.��Ŀ���,A.�������,A.��¼��,C.��Ŀֵ��,A.��¼��,A.��¼��,A.��¼ɫ,nvl(A.���ֵ,0) ���ֵ ,nvl(A.��Сֵ,0) ��Сֵ,A.�ٽ�ֵ," & _
                "nvl(A.��λֵ,0) ��λֵ,A.�̶ȼ��,A.��ʾ��,C.��Ŀ��λ ��λ,Decode(��¼��,3,A.�����,nvl(A.�����,2)-2) AS �����,B.��λ " & _
                " From ���¼�¼��Ŀ A,���²�λ B,�����¼��Ŀ C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D" & _
                " Where A.��Ŀ���=B.��Ŀ���(+) And B.ȱʡ��(+)=1" & _
                " And A.��Ŀ���=C.��Ŀ��� And A.��¼��<>2 And NOT (NVL(C.Ӧ�÷�ʽ,0)=2 And C.��Ŀ���=-1) And C.��Ŀ���=D.COLUMN_VALUE" & _
                " Order by �������"
                
            strSQLText = strSQL
        Case "��ȡ���ǻ�����Ŀ"
            
            gstrSQL = _
            " Select A.�������,A.��Ŀ���,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,B.��Ŀֵ��,nvl(D.C2,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����,'' ��λ," & _
            "   B.��Ŀ����,B.��Ŀ����,B.��Ŀ��ʾ,B.��ĿС��,B.��Ŀ��λ ��Ŀ��λ" & _
            "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C,TABLE(CAST(F_NUM2LIST2([10]) AS ZLTOOLS.T_NUMLIST2)) D" & _
            "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And B.��Ŀ���=D.C1 And A.��¼��=2 And B.��Ŀ����=1" & _
            "   UNION ALL " & _
            " Select Distinct  B.�������,B.��Ŀ���,B.��¼��,B.��¼��,B.��¼��,B.��¼ɫ,C.��Ŀֵ��,nvl(B.��¼Ƶ��,2) ��¼Ƶ��,B.��Ժ�ײ�,C.��Ŀ����, A.��λ," & _
                "   C.��Ŀ����,C.��Ŀ����,C.��Ŀ��ʾ,C.��ĿС��,C.��Ŀ��λ ��Ŀ��λ" & _
                "            From (Select ��Ŀ���, DECODE(��Ŀ���,3,'',���²�λ) ��λ" & vbNewLine & _
                "                           From ���˻����ļ� a, ���˻������� b, ���˻�����ϸ c" & vbNewLine & _
                "                           Where a.Id = b.�ļ�id And b.Id = c.��¼id And a.Id = [1] And Nvl(a.Ӥ��, 0) = [4] And a.����id = [2] And" & vbNewLine & _
                "                                       a.��ҳid = [3] And c.��¼���� = 1 And b.����ʱ�� Between [5] And [6] And ��ֹ�汾 Is Null) a, ���¼�¼��Ŀ b," & vbNewLine & _
                "                       �����¼��Ŀ c" & vbNewLine & _
                "            Where b.��Ŀ��� = a.��Ŀ��� And b.��Ŀ��� = c.��Ŀ��� And b.��¼�� = 2 And C.��Ŀ����=2" & _
                "   And nvl(C.Ӧ�÷�ʽ,0)=1 And nvl(C.����ȼ�,0)>=[7] And nvl(C.���ò���,0) In (0,[8])" & _
                "   And (C.���ÿ���=1 Or (C.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=C.��Ŀ��� And D.����id=[9])))"
        
            strSQL = "Select Rownum-1 ��� ,��Ŀ���, Decode(��Ŀ���, 4, 'Ѫѹ',��¼��) ��Ŀ����,��¼ɫ,��Ŀ��λ,��Ŀֵ��, ��λ,��¼Ƶ��,��Ժ�ײ�,��Ŀ����,��Ŀ��ʾ,��Ŀ���� " & _
                " From ( select �������, ��Ŀ���, ��¼��, ��¼��, ��¼��, ��¼ɫ, ��Ŀֵ��, Nvl(��¼Ƶ��, 2) ��¼Ƶ��, ��Ժ�ײ�, ��Ŀ����, ��λ, ��Ŀ����, ��Ŀ����,��Ŀ��ʾ, ��ĿС��, ��Ŀ��λ ��Ŀ��λ " & _
                      "From(" & gstrSQL & ") A where ��Ŀ���<>5 order by Decode(A.��Ŀ���,3 ,0,1 ),A.�������,a.��¼��,a.��λ) "
                      
            strSQLText = strSQL
        Case "��ȡ�����������ݺ�δ��˵��"
                
            strSQL = _
                    " SELECT /*+ Rule*/ C.ID ���, a.����ʱ�� As ʱ��,C.��ʾ,C.��¼���� As ��ֵ,C.���²�λ,c.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵�� " & _
                    " FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E,Table(Cast(f_num2list([6]) As zlTools.t_Numlist)) F " & _
                    " Where B.ID=A.�ļ�ID  " & _
                    "   AND A.ID = C.��¼ID " & _
                    "   AND B.ID=[1] " & _
                    "   AND B.����id=[2] " & _
                    "   AND B.��ҳid=[3] " & _
                    "   AND D.��Ŀ���=C.��Ŀ��� " & _
                    "   AND C.��¼����=1 " & _
                    "   AND E.��Ŀ���=D.��Ŀ��� " & _
                    "   AND E.��Ŀ���=F.COLUMN_VALUE " & _
                    "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null And D.��¼��<>2 " & _
                    " Order By a.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
            
            strSQLText = strSQL
        Case "��ȡ���б����Ŀ������Ϣ"
            
            strSQL = "SELECT /*+ Rule*/  C.Id,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & vbNewLine & _
                "  DECODE(E.��Ŀ����,2,C.���²�λ || D.��¼��,D.��¼��) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,E.��Ŀ���� " & _
                "  FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E" & _
                "  Where B.ID = A.�ļ�ID" & vbNewLine & _
                "  AND A.ID = C.��¼ID" & vbNewLine & _
                "  AND B.ID = [1]" & vbNewLine & _
                "  AND B.����id = [2]" & vbNewLine & _
                "  AND B.��ҳid = [3]" & vbNewLine & _
                "  AND Nvl(B.Ӥ��, 0) = [7]" & vbNewLine & _
                "  AND INSTR([6], DECODE(E.��Ŀ����, 2,C.���²�λ || D.��¼��, D.��¼��)) > 0" & vbNewLine & _
                "  AND D.��Ŀ��� = C.��Ŀ���" & vbNewLine & _
                "  AND Mod(c.��¼����,10) = 1" & vbNewLine & _
                "  AND E.��Ŀ��� = D.��Ŀ���" & vbNewLine & _
                "  AND E.����ȼ� >= [8]" & vbNewLine & _
                "  AND A.����ʱ�� BETWEEN [4] And [5]" & vbNewLine & _
                "  AND D.��¼�� = 2" & vbNewLine & _
                "  UNION ALL "
             '��ȡ�����±��Ļ�����Ŀ�����±�������Ŀ������ܴ��ڷ�������Ŀ��
            strSQL = strSQL & vbNewLine & _
                "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
                "   D.��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
                "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, 1 ��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
                "       WHERE A.��Ŀ���=B.��� AND NOT EXISTS (SELECT C.COLUMN_VALUE FROM Table(Cast(f_num2list([11]) As zlTools.t_Numlist)) C,���������Ŀ E WHERE C.COLUMN_VALUE=E.��� AND C.COLUMN_VALUE=A.��Ŀ���)" & vbNewLine & _
                "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
                "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
                "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
                "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
                "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"
                
            strSQL = _
                "   Select ID,ʱ��,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & strSQL & ")" & _
                "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
                
            strSQLText = strSQL
        Case "��ȡ���������±���Ϣ"
            strSQL = "" & _
                 " Select B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.��¼����,C.��Ŀ����,C.δ��˵��" & _
                 " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
                 " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]   AND Nvl(A.Ӥ��, 0)=[6] AND A.����id=[2] AND A.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
                 " AND mod(c.��¼����,10) <> 1  AND B.����ʱ�� BETWEEN [4]  And [5]"
            strSQLText = strSQL
        Case "��ʾƤ�Խ��"
            lngColor = RGB(0, 0, 255)
            strSQL = _
               "SELECT ʱ��,F_LIST2STR(CAST(COLLECT(ҩ����) AS T_STRLIST)) ҩ���� FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(A.��ʼִ��ʱ��,'YYYY-MM-DD') ʱ��,DECODE(Ƥ�Խ��,'(+)',255,'(����)',255," & lngColor & ") || '-#' || REPLACE(REPLACE(REPLACE(DECODE(B.�Թܱ���,NULL,A.ҽ������,B.�Թܱ���),',',''),'-#',''),'Ƥ��','') || A.Ƥ�Խ��  ҩ����" & vbNewLine & _
                "   FROM ����ҽ����¼ A,������ĿĿ¼ B " & vbNewLine & _
                "   WHERE  A.����ID=[1] AND A.��ҳID=[2] AND A.Ӥ��=[3] AND A.Ƥ�Խ�� IS NOT NULL" & vbNewLine & _
                "   AND A.��ʼִ��ʱ��  BETWEEN [4] AND [5] AND A.������Ŀid=b.id(+)" & vbNewLine & _
                "   ORDER BY A.��ʼִ��ʱ��,A.Ƥ�Խ��)" & vbNewLine & _
                "GROUP BY ʱ��"
            strSQLText = strSQL
        Case "��ȡ���Ҵ���"
            strSQL = " Select  c.���� As ����,b.���� As ����,a.����,a.��ʼԭ�� " & _
                " From ���˱䶯��¼ a,���ű� b,���ű� c " & _
                " Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id And NVL(A.���Ӵ�λ,0)=0 " & _
                " And a.��ʼʱ��-" & T_BodyStyle.lngʱ���� & "/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] Order By a.��ʼʱ��"
            strSQLText = strSQL
        Case "��ȡ��ǰ������Ϣ"
            strSQL = "Select B.����ʱ�� ʱ��" & vbNewLine & _
                " From ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & vbNewLine & _
                " Where A.Id=B.�ļ�ID And B.Id=C.��¼ID And A.Id=[1] And  nvl(A.Ӥ��,0)=[2]" & vbNewLine & _
                " And A.����ID=[3] and A.��ҳID=[4] and C.��¼����=4 And NVL(C.���Ժϸ�,0)<>1 and C.��ֹ�汾 is null" & vbNewLine & _
                " And B.����ʱ�� between [5] and [6] order by B.����ʱ��"
            strSQLText = strSQL
        Case "��ȡ14��֮ǰ��������Ϣ"
            strSQL = "select Nvl(Count(B.����ʱ��),0) ����" & _
                "   from ���˻����ļ� A, ���˻������� B,���˻�����ϸ C" & _
                "   where A.ID=B.�ļ�ID and B.ID=C.��¼ID and A.ID=[1] and nvl(A.Ӥ��,0)=[2]" & _
                "   and A.����ID=[3] and A.��ҳID=[4] and C.��¼����=4 And NVL(C.���Ժϸ�,0)<>1 and C.��ֹ�汾 is null" & _
                "   and B.����ʱ�� <[5] "
            strSQLText = strSQL
    End Select
    If blnMoved Then
        strSQLText = Replace(strSQLText, "���˻����ļ�", "H���˻����ļ�")
        strSQLText = Replace(strSQLText, "���˻�������", "H���˻�������")
        strSQLText = Replace(strSQLText, "���˻�����ϸ", "H���˻�����ϸ")
        strSQLText = Replace(strSQLText, "���˹�����¼", "H���˹�����¼")
    End If
    getSQLString = strSQLText
End Function

Public Function GetDiagnoseMinTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strTime As Date, Optional ByVal blnMoved As Boolean = False) As String
'����:��ȡ��С���ʱ��
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, strSQL As String
    On Error GoTo Errhand
    strTmp = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    strSQL = "SELECT /*+Rule */" & vbNewLine & _
        " MIN(��¼����) �������" & vbNewLine & _
        " FROM ������ϼ�¼ a, TABLE(CAST(f_Num2list('1,2') AS Zltools.t_Numlist)) b" & vbNewLine & _
        " WHERE MOD(a.�������, 10) = b.Column_Value AND a.����id = [1] AND a.��ҳid = [2] And a.��¼��Դ>1"
    If blnMoved = True Then
        strSQL = Replace(strSQL, "������ϼ�¼", "H������ϼ�¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��С���ʱ��", lng����ID, lng��ҳID)
    If rsTmp.BOF = False Then
        If IsDate(Nvl(rsTmp!�������)) Then
            If CDate(rsTmp!�������) >= CDate(strTmp) Then
                strTmp = Format(rsTmp!�������, "yyyy-MM-dd HH:mm:ss")
                strTmp = DateAdd("s", 1, CDate(strTmp))
            End If
        End If
    End If
    GetDiagnoseMinTime = strTmp
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub ClearData(ByVal lngƫ����X As Long, ByVal lngƫ����Y As Long, ByVal lng�̶ȵ�λ As Long, _
                        ByVal sin�е�λ As Single, ByVal sinʱ���е�λ As Single, ByVal sinʱ���е�λ As Single, ByVal lng�е�λ As Long, ByVal bln˫�� As Boolean, ByVal lng������ As Long, _
                        ByVal lng�������������� As Long, ByVal lng�̶ȿ�� As Long)
    
    T_DrawClient.ƫ����X = lngƫ����X
    T_DrawClient.ƫ����Y = lngƫ����Y
    T_DrawClient.�̶ȵ�λ = lng�̶ȵ�λ
    T_DrawClient.�е�λ = sin�е�λ
    T_DrawClient.ʱ���е�λ = sinʱ���е�λ
    T_DrawClient.ʱ���е�λ = sinʱ���е�λ
    T_DrawClient.�е�λ = lng�е�λ
    T_DrawClient.˫�� = bln˫��
    T_DrawClient.������ = lng������
    T_DrawClient.�������������� = lng��������������
    T_BodyStyle.lng�̶ȿ�� = lng�̶ȿ��
End Sub


Private Function Get����ȼ�(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lng����ȼ� As Long
    
    
    strSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ȼ�", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then lng����ȼ� = zlCommFun.Nvl(rsTemp("����ȼ�"), 0)
    
    Get����ȼ� = lng����ȼ�
End Function


Private Function Get������(ByVal dbl��ֵ As Double, ByVal lngCurveRow As Long) As Integer
    Dim intDrawLineRows As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim intRows As Integer
    
    strSQL = "Select Count(A.��Ŀ���) ��¼�� " & _
    "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C " & _
    "   Where A.��Ŀ���=B.��Ŀ��� And B.��Ŀ���=C.COLUMN_VALUE AND A.��¼��<>2 AND NOT (NVL(B.Ӧ�÷�ʽ,0)=2 And B.��Ŀ���=-1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", T_BodyItem.str������Ŀ)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        intDrawLineRows = zlCommFun.Nvl(rsTmp!��¼��, 0)
    Else
        CloseRs rsTmp
        Get������ = 0
        Exit Function
    End If

    If intDrawLineRows < 1 Then
        Get������ = 0
        Exit Function
    End If
    
     strSQL = "Select nvl(A.���ֵ,0) ���ֵ,nvl(A.��Сֵ,0) ��Сֵ ,nvl(A.��λֵ,0.1) ��λֵ ,Decode(��¼��,3,A.�����,nvl(A.�����,2)-2) AS �����,A.��¼��,A.��Ŀ���" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C " & _
        "   Where A.��Ŀ���=B.��Ŀ��� And b.��Ŀ���=c.Column_value AND A.��¼��<>2 AND NOT (NVL(B.Ӧ�÷�ʽ,0)=2 And B.��Ŀ���=-1)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", T_BodyItem.str������Ŀ)

    rsTmp.Filter = "��¼��=1 And ��Ŀ���=" & gint����
    If rsTmp.RecordCount > 0 Then
        '�޸����⣺51442
        dbl��ֵ = Val(zlCommFun.Nvl(rsTmp!��Сֵ, 0))
        intDrawLineRows = (Val(rsTmp!���ֵ) - IIf(dbl��ֵ > 34, 35, dbl��ֵ)) / 0.1 + IIf(Val(rsTmp!�����) < 0, 0, Val(rsTmp!�����))
        intDrawLineRows = intDrawLineRows + lngCurveRow
    Else
        intDrawLineRows = glngMaxRows + lngCurveRow
    End If
    
       
    T_DrawClient.�������������� = 0
    rsTmp.Filter = "��¼��=3"
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
             '�޸����⣺51442
            intRows = (Val(rsTmp!���ֵ) - Val(zlCommFun.Nvl(rsTmp!��Сֵ, 0))) / Val(zlCommFun.Nvl(rsTmp!��λֵ, 0)) + IIf(Val(rsTmp!�����) < 0, 0, Val(rsTmp!�����))
            If intRows Mod 2 = 1 Then intRows = intRows + 1
            T_DrawClient.�������������� = T_DrawClient.�������������� + intRows
            rsTmp.MoveNext
        Loop
    End If
    
    Get������ = intDrawLineRows
End Function





Private Function Get����(ByVal lng�ص� As Long, ByVal str�ص���Ŀ As String, ByVal lng��Ŀ��� As Long, ByVal str���� As String, ByVal strPosition As String, strEditors() As Variant, ByVal lng��� As Long) As Boolean
    '��ȡ���·���
    Dim str��λ As String
    Dim strTmp As String
    Dim strChar As String
    Dim strPic As String
    Dim i As Integer
    
    On Error GoTo Errhand
    
    
    If lng�ص� = 0 And str�ص���Ŀ = "��" Then 'δ�ص�����Ŀ
         For i = 0 To UBound(strEditors)
            If Split(CStr(strEditors(i)), "||")(0) = lng��Ŀ��� Then
                 Exit For
            End If
        Next i
        str��λ = strPosition
        If str��λ = "" Then
            Select Case lng��Ŀ���
                Case gint����
                    str��λ = "Ҹ��"
                Case gint����
                    str��λ = "��������"
                Case Else
                    str��λ = ""
            End Select
        End If
        strTmp = Split(CStr(strEditors(i)), "||")(4)
        strPic = ""
        strChar = ""
        Select Case lng��Ŀ���
            Case gint����
                strTmp = strTmp & String(3 - UBound(Split(strTmp, ",")), ",")
                If str��λ = "����" Then
                    strChar = Split(strTmp, ",")(0)
                ElseIf str��λ = "Ҹ��" Then
                    strChar = Split(strTmp, ",")(1)
                ElseIf str��λ = "����" Then
                    strChar = Split(strTmp, ",")(2)
                Else
                    strChar = Split(strTmp, ",")(3)
                End If
                If lng��� = 1 Then '�����·���
                    strChar = "��"
                Else
                    If strChar = "" Then strChar = "��"
                End If
            Case gint����
                strChar = IIf(strTmp = "", "��", strTmp)
            Case gint����
                If str��λ = "����" Then
                    strPic = "PACEMAKER"
                Else
                    strChar = IIf(strTmp = "", "+", strTmp)
                End If
            Case gint����
                If str��λ = "��������" Then
                    strChar = IIf(strTmp = "", "*", strTmp)
                Else
                    strPic = "BREATH"
                End If
            Case Else
                strChar = strTmp
        End Select
        If Trim(str����) <> "" Then
            strChar = Trim(str����)
            strPic = ""
        End If
    End If
    
    If strChar <> "��" Then
        Get���� = False
    Else
        Get���� = True
    End If
        
        Exit Function
Errhand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function


Public Sub DrawTextPrint(objDraw As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, Optional ByVal ForeColor As Long = 0)
    '��(X,Y)�����Text�ı�
    Dim lngSaveForeColor As Long
    
    With objDraw
        lngSaveForeColor = .ForeColor
        .ForeColor = ForeColor
        .CurrentX = X
        .CurrentY = Y
        objDraw.FontTransparent = True
        objDraw.Print Text
        .ForeColor = lngSaveForeColor
    End With
End Sub

Private Function GetMaxMinValue(ByVal bytType As Byte, ByVal lngNO As Long, arrEditors() As Variant) As Double
'����:��ȡ������Ŀ���ٽ�ֵ(�������ֵ����Сֵ)
'����:bytType=0 ��Сֵ,1-���ֵ
'     arrEditors:'��¼������Ŀ��Ϣ(��Ŀ���||��Ŀ����||��Ŀ��λ||��Ŀֵ��||��¼��||��¼ɫ||���ֵ||��Сֵ||�ٽ�ֵ��
    Dim dblvalue As Double
    Dim dblMax As Double, dblMin As Double
    Dim strValue As String
    Dim i As Integer
    
    For i = 0 To UBound(arrEditors)
        If Val(Split(arrEditors(i), "||")(0)) = lngNO Then
             Exit For
        End If
    Next i
    
    If i <= UBound(arrEditors) Then
        dblMax = Val(Split(arrEditors(i), "||")(6))
        dblMin = Val(Split(arrEditors(i), "||")(7))
    End If
    
    strValue = Split(arrEditors(i), "||")(8)
    If bytType = 0 Then
        dblvalue = dblMin
        If InStr(1, strValue, ";") <> 0 Then
            strValue = Split(strValue, ";")(0)
        Else
            strValue = ""
        End If
        If IsNumeric(strValue) = True And Val(strValue) <= dblMax And Val(strValue) >= dblMin Then
            dblvalue = Val(strValue)
        Else
            '���������Сֵ��Ч�����������СֵΪ35
            If lngNO = gint���� And dblvalue < 35 Then dblvalue = 35
        End If
    Else
        dblvalue = dblMax
        If InStr(1, strValue, ";") <> 0 Then strValue = Split(strValue, ";")(1)
        If IsNumeric(strValue) = True And Val(strValue) <= dblMax And Val(strValue) >= dblMin Then dblvalue = Val(strValue)
    End If
    
    GetMaxMinValue = dblvalue
End Function
