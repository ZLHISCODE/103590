Attribute VB_Name = "mdlSpePrint"
Option Explicit

Private mlng���²�����ʾ��ʽ As Long
Private mintBaby As Integer  '�Ƿ���Ӥ��
Private mlngBreatheHeight  '�������̶��߶�
Public Const glngBreashLableWidth As Long = 600 '�������߿̶��п�(�)

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
    Dim lngNoRecordColor As Long '����:δ��˵����ʾ��ɫ
    Dim blnCurBeginTop As Boolean '����:���±�־���λ��
    '������ͼ����
    Dim i As Integer, j As Integer
    Dim lngPicPageIndex As Integer 'Ԥ��ʱPIC������
    Dim blnPrint As Boolean  '�Ƿ��ӡ
    Dim strInfo As String '˵����Ϣ
    Dim intAllOpt As Single  '��ӡ���ܹ�����
    Dim intCurOpt As Single  '��ӡ���е��ڼ���
    Dim objDraw As Object '��ͼ����
    Dim lngHwnd As Long '���
    Dim lngDc As Long  '��ͼ�����DC
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
    Dim T_RECT As RECT
    Dim rsPart As New ADODB.Recordset  '�������²�λ��Ϣ
    Dim rsTemp As New ADODB.Recordset  '�˼�¼���벻Ҫ˳��ʹ��
    Dim rsTmp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset 'ʹ����˲��˵����л�����Ŀ��Ϣ
    Dim rsDrawItems As New ADODB.Recordset '���µ�������Ŀ��Ϣ
    Dim rsPoints As New ADODB.Recordset '�������µ��ļ���
    Dim rsNotes As New ADODB.Recordset   '����˵����Ϣ
    Dim rsDownTab As New ADODB.Recordset '���±��������Ϣ
    Dim H_16pt As Long, W_16pt As Long
    Dim int����Ӧ�� As Integer, str������Ϣ As String
    Dim str���ʷ���  As String, lng������ɫ As Long
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
    Dim int��Ŀ���� As Integer, int��Ŀ���� As Integer, strTabItemTemp As String
    Dim int����ѹ As Integer, int����ѹ As Integer, Int�к� As Integer
    Dim blnColor As Boolean

    '���˻�����Ϣ
    Dim strPatiInfo As String, strBabyDay As String
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
    Dim ArrComTable() As Variant '��¼���еı��±����Ŀ (��Ŀ���||��λ+��Ŀ����|��Ŀ��λ||��Ŀֵ��||��¼Ƶ��||��Ŀ����||��Ŀ��ʾ||��Ժ�ײ�||��λ||��¼ɫ)
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
    '71950:������,2014-06-11,���µ�δ��˵����ʾ��ɫ
    lngNoRecordColor = Val(zlDatabase.GetPara("δ��˵����ʾ��ɫ", glngSys, 1255, "16711680"))
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
    '70452:������,2014-08-02,���±�־���λ��:���߶���ʼ������42��ʼ
    blnCurBeginTop = (Val(zlDatabase.GetPara("���±�־���λ��", glngSys, 1255, "0")) = 1)
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
    intAllOpt = 4
    
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
    
    '����:'����:'�Ա�:'�Ʊ�:'����:'����:'סԺ��:
    strPatiInfo = "''''''"
    VarPatiInfo = Split(strPatiInfo, "'")
    
    strSQL = " Select  NVL(A.����,b.����) ����,A.סԺ��,A.��Ժ���� ��Ժʱ��,NVL(A.�Ա�,b.�Ա�) �Ա�,NVL(A.����,B.����) ���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID)
    If rsTemp.BOF = False Then
        VarPatiInfo(0) = zlCommFun.Nvl(rsTemp("����").Value)
        VarPatiInfo(6) = zlCommFun.Nvl(rsTemp("סԺ��").Value)
        VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("�Ա�").Value)
    End If
    
    If intBaby <> 0 Then
        VarPatiInfo(2) = ""
        strSQL = "Select Decode(a.Ӥ������,Null,NVL(C.����,B.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,a.Ӥ���Ա�,a.����ʱ�� " & _
            " From ������Ϣ B,������ҳ C,������������¼ A " & _
            " Where B.����ID=C.����ID And C.����ID=A.����ID And C.��ҳID=A.��ҳID And C.����id=[1] And C.��ҳid=[2] And a.���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID, intBaby)
        If rsTemp.BOF = False Then
            VarPatiInfo(0) = rsTemp("Ӥ������").Value
            VarPatiInfo(2) = zlCommFun.Nvl(rsTemp("Ӥ���Ա�").Value)
            strBabyDay = Format(Nvl(rsTemp("����ʱ��").Value), "YYYY-MM-DD")
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
    str���ʷ��� = "��"
    lng������ɫ = RGB_RED
    str������Ϣ = ""
    strSQL = "Select a.Ӧ�÷�ʽ,a.��Ŀ����,b.��¼��,b.��¼ɫ From �����¼��Ŀ a,���¼�¼��Ŀ b Where a.��Ŀ���=-1 And a.��Ŀ���=b.��Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint")
    If rsTemp.BOF = False Then
        int����Ӧ�� = zlCommFun.Nvl(rsTemp("Ӧ�÷�ʽ").Value, 2)
        str���ʷ��� = zlCommFun.Nvl(rsTemp("��¼��").Value, "��")
        str������Ϣ = int����Ӧ�� & "'" & rsTemp!��Ŀ���� & "'" & str���ʷ��� & "'" & lng������ɫ
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
        rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
        gbln�������� = rsTemp.RecordCount > 0
        rsTemp.Filter = "��¼��=1"
        intDrawLineCOL = rsTemp.RecordCount
        rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
        If rsTemp.RecordCount > 0 And bln����ӡ������ Then
            rsTemp.Filter = 0
            intDrawLineCOL = intDrawLineCOL - 1
        Else
            rsTemp.Filter = 0
        End If
        If gbln�������� = True Then intDrawLineCOL = intDrawLineCOL - 1
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
        strTmp = "-1||����||" & Split(strEditors(int����), "||")(2) & "||" & Split(strEditors(int����), "||")(3) & "||" & str���ʷ��� & "||" & lng������ɫ & "||" & _
            Split(strEditors(int����), "||")(6) & "||" & Split(strEditors(int����), "||")(7) & "||" & Split(strEditors(int����), "||")(8)
        strEditors(UBound(strEditors)) = strTmp
    End If
    
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
    If T_DrawClient.��������.Right + IIf(gbln�������� = True, T_BodyStyle.lngBreashLabelStep, 0) > objDraw.Width / T_TwipsPerPixel.X - lngRight Then
        sigSign = Round((T_DrawClient.��������.Right + IIf(gbln�������� = True, T_BodyStyle.lngBreashLabelStep, 0) - (objDraw.Width / T_TwipsPerPixel.X - lngRight)) / (T_DrawClient.��������.Right - T_DrawClient.�̶�����.Right), 2)
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
        lngDc = objDraw.hDC
        '67934:������,2013-12-03,��͸��״̬���л�ͼ
        Call SetBkMode(lngDc, TRANSPARENT)
        '��������
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '��ӡ�ʿغ�
        strTmp = zlDatabase.GetPara("�ʿغ�", glngSys, 1255, "")
        Call GetTextExtentPoint32(lngDc, strTmp, Len(strTmp), T_Size)
        T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
        lngCurX = T_DrawClient.��������.Right - T_Size.W
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, , , , sngScale)
        Call DrawText(lngDc, strTmp, -1, T_LableRect, DT_CENTER)
        Call SelectObject(lngDc, lngOldFont)
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
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            strTmp = IIf(GetUnitName = "-", "", GetUnitName) & IIf(intBaby <> 0, "Ӥ��", "") & T_BodyStyle.str�����ı�
            Call GetTextExtentPoint32(lngDc, strTmp, Len(strTmp), T_Size)
            lngCurY = T_Size.H \ 2 + lngCurY
            Call GetTextRect(objDraw, 0, lngCurY, strTmp, objDraw.Width / T_TwipsPerPixel.X, True, T_Size.H, sngScale)
            Call DrawText(lngDc, strTmp, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDc, lngOldFont)
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
        VarPatiInfo(1) = GetElementValue("����", lng����ID, lng��ҳID, intBaby, strTmpDay)
        VarPatiInfo(3) = ""
        VarPatiInfo(5) = ""
        VarPatiInfo(4) = ""
        strTmp = "": strTime = "": strTmp1 = ""
        
        strSQL = getSQLString("��ȡ���Ҵ���", blnMoved)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˿��ҡ����ŵ���Ϣ", lng����ID, lng��ҳID, CDate(Format(strEndDay, "yyyy-mm-dd hh:mm:ss")), CDate(Format(strTmpDay, "yyyy-mm-dd hh:mm:ss")))
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF
            If zlCommFun.Nvl(rsTmp("����").Value) <> strTmp1 And zlCommFun.Nvl(rsTmp("����").Value) <> "" Then
                    strTmp1 = zlCommFun.Nvl(rsTmp("����").Value)
                    If VarPatiInfo(4) = "" Then
                        VarPatiInfo(4) = strTmp1
                    Else
                        VarPatiInfo(4) = VarPatiInfo(4) & "->" & strTmp1
                    End If
                End If
                
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
                    If VarPatiInfo(5) = "" Then
                        VarPatiInfo(5) = strTime
                    Else
                        VarPatiInfo(5) = VarPatiInfo(5) & "->" & strTime
                    End If
                End If
                            
                rsTmp.MoveNext
            Loop
            
            If Left(VarPatiInfo(3), 2) = "->" Then VarPatiInfo(3) = Mid(VarPatiInfo(3), 3)
            If Left(VarPatiInfo(4), 2) = "->" Then VarPatiInfo(4) = Mid(VarPatiInfo(4), 3)
            If Left(VarPatiInfo(5), 2) = "->" Then VarPatiInfo(5) = Mid(VarPatiInfo(5), 3)
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
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '���������Ϣ
        Call DrawPatiInfo(lngDc, objDraw, strPatiInfo, lngCurX, lngCurY, T_DrawClient.��������.Right + IIf(gbln�������� = True, T_BodyStyle.lngBreashLabelStep, 0), lngCurY, intBaby & "'" & strBabyDay, sngScale)
        Call SelectObject(lngDc, lngOldFont)
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
                    str����ʱ�� = Format(DateAdd("D", T_BodyStyle.lng���� - 1, CDate(strTmpDay)), "YYYY-MM-DD") & " 23:59:59"
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
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '3��ʼ���סԺ���ڣ�������������Ϣ
        Call DrawUpTableNew(lngDc, objDraw, strTmpString0 & "||" & strTmpString2 & "||" & strTmpString1, lngCurX, lngCurY, T_DrawClient.��������.Right, lngCurY, sngScale)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '----------------------------------------------------------------------------------------------
        '��ȡ����������Ŀ��Ϣ�������Ŀ�����Ŀ���ܴ���һ����Ŀ�����λҲҪ��ȡ��
        ArrComTable = Array()
        Set rsItems = GetAppendGridItemNew(lng�ļ�ID, lng����ID, lng��ҳID, lng����ȼ�, intBaby, CDate(strTmpDay), CDate(strEndDate), IIf(intBaby = 0, 1, 2), lngSectID, T_BodyItem.str�����Ŀ, blnMoved)
        bln���� = False
        With rsItems
            Do While Not .EOF
                If Val(Nvl(!��Ŀ���)) <> 5 Then
                    If Val(Nvl(!��Ŀ���)) = 4 Then
                        str��Ŀ���� = "Ѫѹ"
                    Else
                        str��Ŀ���� = Nvl(!��¼��)
                    End If
                    
                    intƵ�� = Val(zlCommFun.Nvl(!��¼Ƶ��))
                    If zlCommFun.Nvl(!��Ŀ��ʾ) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!��Ŀ���))) Then
                        If intƵ�� > 2 Then intƵ�� = 2
                    End If
                    
                    strTmp = zlCommFun.Nvl(!��Ŀ���) & "||" & Replace(str��Ŀ����, ";", ":") & "||" & zlCommFun.Nvl(!��λ) & "||" & _
                        zlCommFun.Nvl(!��Ŀֵ��) & "||" & intƵ�� & "||" & zlCommFun.Nvl(!��Ŀ����, 1) & "||" & _
                        zlCommFun.Nvl(!��Ŀ��ʾ) & "||" & zlCommFun.Nvl(!��Ŀ����) & "||" & zlCommFun.Nvl(!��Ժ�ײ�, 0) & "||" & Nvl(!���²�λ) & "||" & Nvl(!��¼ɫ, RGB_BLUE)
                    If Val(zlCommFun.Nvl(!��Ŀ���)) = gint���� Then
                        bln���� = True
                    End If
                    
                    ReDim Preserve ArrComTable(UBound(ArrComTable) + 1)
                    ArrComTable(UBound(ArrComTable)) = strTmp
                End If
            .MoveNext
            Loop
        End With
        
        '�˴��ļ������DrawBodyPageFooterNew����һ��
         sinY = 0
         If (Val(zlDatabase.GetPara("���µ�����ӡ����˵��", glngSys, 1255, "0")) = 1) = False And gbln�������� = False Then
            sinY = 14 * msngTwips
         Else
            sinY = 6 * msngTwips
         End If
         If (Val(zlDatabase.GetPara("��ӡҳ��", glngSys, 1255, "1")) = 1) Or (Val(zlDatabase.GetPara("��ӡ����", glngSys, 1255, "0")) = 1) Or _
            (Val(zlDatabase.GetPara("��ӡ��ӡ��", glngSys, 1255, "0")) = 1) Then
            sinY = 12 * msngTwips
        End If
        
         '�˴�����ɴ�ӡ���� �Ӷ��������µ���ӡ���и�
        T_DrawClient.ʱ���е�λ = T_BodyStyle.lng�±��߶� / T_TwipsPerPixel.Y * sngScale
'        If intRepairRows = 0 Then
'            sngHTab = intRepairRows
'        Else
'            '�����̶�Ϊ300
'            sngHTab = intRepairRows * T_DrawClient.ʱ���е�λ + IIf(bln���� = True, mlngBreatheHeight - T_DrawClient.ʱ���е�λ, 0)
'        End If
'        sngHTab = sngHTab + sinY
'        sngHPrint = Format(objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHTab, "#0.00;-#0.00;0.00")
'        T_DrawClient.�е�λ = (sngHPrint - 2 * T_DrawClient.�е�λ) / (T_DrawClient.������ + T_DrawClient.��������������)
'        T_DrawClient.�е�λ = Round(T_DrawClient.�е�λ - 0.05, 1) * sngScale
'        If T_DrawClient.�е�λ > T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.�е�λ = T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale
'        If T_DrawClient.�е�λ < T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale Then T_DrawClient.�е�λ = T_BodyStyle.lng�����и� / T_TwipsPerPixel.Y * sngScale
'
        '�����иߺ��ڼ������µ��ɴ�ӡ�ı������
        If intRepairRows > 0 Then
            sngHPrint = (T_DrawClient.������ + T_DrawClient.��������������) * T_DrawClient.�е�λ + 2 * T_DrawClient.�е�λ
            sngHTab = objDraw.Height / T_TwipsPerPixel.Y - lngCurY - lngButtom - sngHPrint - sinY
            sngHTab = sngHTab - IIf(bln���� = True, mlngBreatheHeight - T_DrawClient.ʱ���е�λ, 0)
            If Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3) < intRepairRows Then intRepairRows = Fix(sngHTab / T_DrawClient.ʱ���е�λ + 0.3)
        End If
        
        'Ϊ�˱������̫������ݣ��˴��ɸ��ݱ�������Լ�ֽ����Ϣ���¼�����߶�
        '˵�������ڻ��ߵ�λֻ�������������߲��������Զ����ſ��ܻᵼ�³��ֶ���հײ���(��ʱ+1���ܵ��±���������,��+�����²��ֿհ�)
        sngHPrint = sngHTab - (intRepairRows * T_DrawClient.ʱ���е�λ + IIf(bln���� = True, mlngBreatheHeight - T_DrawClient.ʱ���е�λ, 0)) '�հײ���
        If sngHPrint > 0 Then
            If bln���� = True Then
                If intRepairRows - 1 > 0 And sngHPrint - mlngBreatheHeight > 0 Then
                    T_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ + (sngHPrint - mlngBreatheHeight) \ (intRepairRows - 1)
                End If
            Else
                T_DrawClient.ʱ���е�λ = T_DrawClient.ʱ���е�λ + sngHPrint \ intRepairRows
            End If
        End If
        If T_DrawClient.ʱ���е�λ > 20 * msngTwips Then T_DrawClient.ʱ���е�λ = 20 * msngTwips
        
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '4��ʼ���̶������������������̶�ֵ��Ϣ
        T_DrawClient.ƫ����Y = lngCurY
        gbln�������� = False
        
        rsTemp.Filter = 0
        rsTemp.Sort = "�������"
        rsTemp.MoveFirst
        str����˵�� = DrawCanvasNew(lngDc, objDraw, rsTemp, rsDrawItems, bln����ӡ������, str������Ϣ, sngScale)
        Call SelectObject(lngDc, lngOldFont)
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
                        !δ��˵�� & "|" & lngNoRecordColor & "|" & SinX & "|0|0|0|0|" & zlCommFun.Nvl(!��ʾ)
                   
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
                    
                    sinY = Val(GetYCoordinate(objDraw, rsDrawItems, !��Ŀ���, dbl��ֵ, lngDc, True))
                    
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
        
        '����������,��ʹ��ʹ����
        For j = 0 To 1
            lng��Ŀ��� = IIf(j = 0, gint����, gint��ʹǿ��)
            arrTmpValue = Array()
            rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " And ���=0"
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "��Ŀ���=" & lng��Ŀ���
            If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
            For i = 0 To UBound(arrTmpValue)
                rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                If rsPoints.RecordCount <> 0 Then
                    gstrFields = "��ע": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                    Call Record_Update(rsPoints, gstrFields, gstrValues, "���|" & zlCommFun.Nvl(rsPoints!���))
                End If
            Next i
            
            arrTmpValue = Array()
            rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1"
            With rsPoints
                Do While Not .EOF
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !��� & ";" & !��Ŀ��� & ";" & !��ֵ & ";" & !X���� & ";" & !Y���� & ";" & Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                .MoveNext
                Loop
            End With
            
            rsPoints.Filter = "��Ŀ���=" & lng��Ŀ���
            If rsPoints.RecordCount > 0 Then rsPoints.MoveFirst
            For i = 0 To UBound(arrTmpValue)
                rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=0 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                If rsPoints.RecordCount = 0 Then
                    rsPoints.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ���=1 and X����=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And ʱ��='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                    rsPoints.Delete
                End If
            Next i
        Next j
    
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
                    If Val(!��Ŀ���) = gint���� Or Val(!��Ŀ���) = gint��ʹǿ�� Then
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
                    If Not ((Val(!��Ŀ���) = gint���� Or Val(!��Ŀ���) = gint��ʹǿ��) And Val(!���) = 1) Then
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
                    Case 15
                        bytShow = T_BodyFlag.ת����
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
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '7 ��ʼ�����Ϣ������
        strTmp = ShowPointsNew(lngDc, objDraw, rsPoints, strEditors, int����Ӧ��, sngScale)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '8.����������������
        rsPoints.Filter = ""
        If strTmp <> "" And bln��ӡ������� = True Then Call CreatePolyNew(rsPoints, objDraw, lngDc, strTmpDay, strTmp, int����Ӧ�� = 2)
        '9���˵����Ϣ
        '�ȴ���δ��˵�����±�˵��
        Dim strText As String
        Dim SinY35 As Single, SinY42 As Single
        Dim intAscCharNum As Integer
        
        strTime = ""
        strTmp = ""
        blnAllow = False
        SinX = 0: sinY = 0
        SinY35 = GetYCoordinate(objDraw, rsDrawItems, gint����, 35, lngDc)
        If blnCurBeginTop = False Then
            SinY42 = GetYCoordinate(objDraw, rsDrawItems, gint����, 42, lngDc)
        Else
            SinY42 = T_DrawClient.��������.Top
        End If
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
                                If sinY < T_DrawClient.��������.Bottom Then
                                    strText = Mid(strTime, i, 1)
                                    T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                    T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                    If T_DrawClient.��������.Bottom - sinY >= T_Size.H - (1 * msngTwips) Then
                                        Call DrawRotateText(objDraw, lngDc, SinX, sinY, strText, Val(!��ɫ))
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
                            If sinY < T_DrawClient.��������.Bottom Then
                                strText = Mid(strTime, i, 1)
                                T_Size.H = objDraw.TextHeight(strText) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(strText) / T_TwipsPerPixel.X
                                If Asc(strText) < 0 Then
                                    If intAscCharNum Mod 2 = 1 Then sinY = sinY + T_Size.H / 2
                                End If
                                '���������Ϣ
                                If T_DrawClient.��������.Bottom - sinY >= T_Size.H - (1 * msngTwips) Then
                                    Call DrawRotateText(objDraw, lngDc, SinX, sinY, strText, Val(zlCommFun.Nvl(!��ɫ)))
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
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call OutPutTextNew(objDraw, rsDrawItems, lngDc, rsNotes, strTmpDay, blnCurBeginTop, sngScale)
        Call SelectObject(lngDc, lngOldFont)
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
        If Not gbln�������� Then strItems = strItems & ",'����'"
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
                strTabItemTemp = Val(Split(ArrComTable(i), "||")(8)) & ";" & Split(ArrComTable(i), "||")(9)
                
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
                                    intƵ�� & ";" & Val(strTmp) & ";" & int��Ŀ���� & ";" & strTabItemTemp, bln���ܵ���, bln¼��Сʱ)
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
           
            ArrNewString(i) = Trim(CStr(Split(ArrComTable(i), "||")(0))) & "'" & strTmpString0 & "'" & Val(CStr(Split(ArrComTable(i), "||")(4))) & "'" & _
                 Val(CStr(Split(ArrComTable(i), "||")(5))) & "'" & Val(CStr(Split(ArrComTable(i), "||")(10))) & "'" & ArrNewString(i)
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
            ArrNewString(UBound(ArrNewString)) = "-999'Ƥ�Խ��'1'1'0" & "'" & strNewTmpString
        End If
        
        lngCurX = X

        '��ʼ�滭�����Ŀ��չʾ����
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call DrawBodyRecordItemNew(lngDc, objDraw, ArrNewString, lngCurX, T_DrawClient.����������.Bottom, T_DrawClient.��������.Right, intRepairRows, lngCurY, sngScale)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        
        lngCurX = X
        lngCurY = lngCurY
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = 9 * sngScale
        stdSet.Bold = False
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        '��ʼ��ӡ ҳ�� סԺ���� �� ����˵����Ϣ
        Call DrawBodyPageFooterNew(lngDc, objDraw, lngCurX, lngCurY, T_DrawClient.��������.Right, intPageNo, intEndPage, str����˵��, sngScale)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        'ҳ��ͼ�����
        Call frmTendFileRead.PrintRTBData(objDraw, False, lngButtom)
        
        If Not blnPrint Then objDraw.Refresh
NOPageSub:  Next

    If blnPrint = False Then Call DrawDeviceCapsNew(lngDc, objDraw)
     
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

Public Sub DrawUpTableNew(ByVal lngDc As Long, ByVal objDraw As Object, ByVal strTmpString As String, _
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
    Call DrawLine(lngDc, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDc, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDc, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ
    Call DrawLine(lngDc, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intFine, RGB_BLACK): lngCurY = lngCurY + T_DrawClient.ʱ���е�λ + 6
    Call DrawLine(lngDc, lngCurX, lngCurY, lngLeft, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    'Y
    Call DrawLine(lngDc, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = T_DrawClient.�̶�����.Right

    Call DrawLine(lngDc, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    
    '����Ϊ�����ϱ����չ����
    If gbln�������� = True Then
        Call DrawLine(lngDc, lngLeft, lngStartY, lngLeft + T_BodyStyle.lngBreashLabelStep, lngStartY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDc, lngLeft + T_BodyStyle.lngBreashLabelStep, lngStartY, lngLeft + T_BodyStyle.lngBreashLabelStep, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    End If

    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = lngCurX + T_DrawClient.�е�λ * T_BodyStyle.lng������
        Call DrawLine(lngDc, lngCurX, lngStartY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
    Next i
    
    lngCurX = T_DrawClient.�̶�����.Right
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    'ʱ��
    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = T_DrawClient.�̶�����.Right + i * T_DrawClient.�е�λ * T_BodyStyle.lng������
        For j = 1 To T_BodyStyle.lng������ - 1
            lngCurX = lngCurX + T_DrawClient.�е�λ
            Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ + 6, PS_SOLID, intFine, RGB_BLACK)
        Next j
    Next i
    
    '��ʼ�����Ϣ
    '������Ϣ
    lngCurY = lngStartY
    Call SetTextColor(lngDc, RGB_BLACK)
    Call GetTextExtentPoint32(lngDc, str����, Len(str����), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, str����, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDc, str����, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    For i = 0 To UBound(arrTmpTime)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDc, RGB_BLUE)
        Call GetTextExtentPoint32(lngDc, CStr(arrTmpTime(i)), Len(CStr(arrTmpTime(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpTime(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDc, CStr(arrTmpTime(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 1
    'סԺ����
    Call SetTextColor(lngDc, RGB_BLACK)
    Call GetTextExtentPoint32(lngDc, strסԺ����, Len(strסԺ����), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, strסԺ����, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDc, strסԺ����, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To UBound(arrTmpDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDc, RGB_BLUE)
        Call GetTextExtentPoint32(lngDc, CStr(arrTmpDay(i)), Len(CStr(arrTmpDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrTmpDay(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDc, CStr(arrTmpDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    
    '��/�������
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 2
    Call SetTextColor(lngDc, RGB_BLACK)
    Call GetTextExtentPoint32(lngDc, str����������, Len(str����������), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, str����������, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDc, str����������, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    '51283,������,2012-07-11,����������ɫ
    lngColor = Val(zlDatabase.GetPara("����������ʾ��ɫ", glngSys, 1255, "255"))
    For i = 0 To UBound(arrOptDay)
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        Call SetTextColor(lngDc, lngColor)
        Call GetTextExtentPoint32(lngDc, CStr(arrOptDay(i)), Len(CStr(arrOptDay(i))), T_Size)
        Call GetTextRect(objDraw, lngCurX, lngCurY + T_DrawClient.ʱ���е�λ / 2, CStr(arrOptDay(i)), T_DrawClient.�е�λ * T_BodyStyle.lng������, True, , sngScale)
        Call DrawText(lngDc, CStr(arrOptDay(i)), -1, T_LableRect, DT_CENTER)
    Next i
    lngColor = 0
    'ʱ��
    lngCurY = lngStartY + T_DrawClient.ʱ���е�λ * 3
    Call SetTextColor(lngDc, RGB_BLACK)
    Call GetTextExtentPoint32(lngDc, strʱ��, Len(strʱ��), T_Size)
    Call GetTextRect(objDraw, lngStartX, lngCurY + T_DrawClient.ʱ���е�λ / 2, strʱ��, T_DrawClient.�̶�����.Right - lngStartX, True, , sngScale)
    Call DrawText(lngDc, strʱ��, -1, T_LableRect, DT_CENTER)
    lngCurX = T_DrawClient.�̶�����.Right
    
    For i = 0 To T_BodyStyle.lng���� - 1
        lngCurX = T_DrawClient.�̶�����.Right + i * T_BodyStyle.lng������ * T_DrawClient.�е�λ
        '�����������ʱ��
        For j = 0 To T_BodyStyle.lng������ - 1
            strTmp = ""
            
            strTmp = gintHourBegin + T_BodyStyle.lngʱ���� * j

            lngColor = GetTimeColor(Val(strTmp))
            lngTmpX = lngCurX + T_DrawClient.�е�λ * j
            Call SetTextColor(lngDc, lngColor)
            Call GetTextExtentPoint32(lngDc, strTmp, Len(strTmp), T_Size)
            Call GetTextRect(objDraw, lngTmpX - 1, lngCurY + (T_DrawClient.ʱ���е�λ + 6) / 2, strTmp, T_DrawClient.�е�λ, True, , sngScale)
            Call DrawText(lngDc, strTmp, -1, T_LableRect, DT_CENTER)
        Next j
    Next i
    lngOutY = lngStartY + T_DrawClient.ʱ���е�λ * 4 + 6
End Sub


Public Function DrawCanvasNew(ByVal lngDc As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, rsDrawItems As ADODB.Recordset, _
    Optional ByVal bln����ӡ������ As Boolean = False, Optional ByVal str������Ϣ As String = "", Optional sngScale As Single = 1) As String
'------------------------------------------------------------------------------------------------------
'����:���̶������������������̶�ֵ��Ϣ
'����:lngDC ��ͼ�����DC��objDraw �滭����.rsTemp:����������Ŀ��¼��(A.��Ŀ���,A.�������,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,C.��Ŀ��λ ��λ,A.�����-2 AS �����,B.��λ)
'       str������Ϣ=""���ʽ��ã���Ϊ�ո�ʽΪ:int����Ӧ�� & "'" & ��Ŀ���� & "'" & str���ʷ��� & "'" & lng������ɫ
'����:���ظ������ߵľ�����Ϣ����( "��Ŀ���|���ֵ|��Сֵ|��λֵ|���ֵ����|��Сֵ����|��λ�̶�|��ʾģʽ|��ɫ")
'����˵����Ϣ(��Ŀ�ķ���)
'-------------------------------------------------------------------------------------------------------
    Dim str˵�� As String
    Dim i As Integer, strTmp As String, ArrCode, int����Ӧ�� As Integer
    Static SlngMaxY As Long                 '��¼��һ�ε����߶ȣ��Ծ��������Ƿ���Ҫ�ػ�
    Dim lngCurX     As Long, lngCurY As Long   '��ǰλ��
    Dim lngMaxX     As Long, lngMaxY As Long   '�߽�
    Dim lngCurAlerY As Long '����������
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln˫�� As Boolean                  '�˲������û�ָ��,bln˫��=TRUE��ʾֻ��ʾ����;������ʾʮ��
    Dim bln���� As Boolean                  '�˲������û�ָ��,���зֽ��Ǵ��߻���ϸ��
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
    
    ArrCode = Split(str������Ϣ, "'")
    If UBound(ArrCode) = -1 Then
        int����Ӧ�� = 0
    Else
        int����Ӧ�� = Val(ArrCode(0))
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
    rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
    gbln�������� = rsTemp.RecordCount > 0
    rsTemp.Filter = "��¼��=1"
    intLables = rsTemp.RecordCount
    rsTemp.Filter = "��Ŀ���=" & gint���� & " And ��¼��=1"
        If rsTemp.RecordCount > 0 And bln����ӡ������ = True Then
        rsTemp.Filter = 0
        intLables = intLables - 1
    Else
        rsTemp.Filter = 0
    End If
    If gbln�������� = True Then intLables = intLables - 1
    If intLables <= 0 Then intLables = 1
    
    lngCurX = T_DrawClient.ƫ����X
    lngCurY = T_DrawClient.ƫ����Y
    lngMaxX = lng�̶ȿ�� + (T_BodyStyle.lng���� * T_BodyStyle.lng������ * lngColStep) + T_DrawClient.ƫ����X    '�̶�+7*���+ƫ����X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.������ * sinRowStep + T_DrawClient.ƫ����Y '��Ϊ����С�����������ʼY���꣩
       
    str˵�� = ""
        
    SlngMaxY = lngMaxY
    T_DrawClient.�̶ȵ�λ = lngLableStep
    T_DrawClient.�е�λ = sinRowStep
    T_DrawClient.�е�λ = lngColStep
    T_DrawClient.˫�� = blnDoubleRow
    
    For lngRow = 1 To intLables
        Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        If lngRow = intLables Then
            lngCurX = lngCurX + lng�̶ȿ�� - ((intLables - 1) * lngLableStep)
        Else
            lngCurX = lngCurX + lngLableStep
        End If
        
        Call DrawLine(lngDc, lngCurX - lngLableStep, lngCurY, lngCurX, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDc, lngCurX - lngLableStep, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    
    T_DrawClient.�̶�����.Left = T_DrawClient.ƫ����X
    T_DrawClient.�̶�����.Top = lngCurY
    T_DrawClient.�̶�����.Right = lngCurX
    T_DrawClient.�̶�����.Bottom = lngMaxY
    
     'Ĭ�������������ʾ��Ŀ����
    For i = 1 To mintNullRow
        If mintNullRow > 1 Then
            Call DrawLine(lngDc, T_DrawClient.ƫ����X + lng�̶ȿ��, lngCurY + (lngInitRowStep * 2 * i), lngMaxX, lngCurY + (lngInitRowStep * 2 * i), PS_SOLID, intFine, RGB_BLACK)
        Else
            Call DrawLine(lngDc, T_DrawClient.ƫ����X, lngCurY + (lngInitRowStep * 2 * i), lngMaxX, lngCurY + (lngInitRowStep * 2 * i), PS_SOLID, intFine, RGB_BLACK)
        End If
    Next i
    lngCurY = lngCurY + lngInitRowStep * 2 * mintNullRow
    '�����µ�������
    For lngRow = 0 To T_DrawClient.������ - 1
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '�����µ���������
        If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
            Call DrawLine(lngDc, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln����, intBold, intFine), RGB_BLACK)
        End If
    Next
    
    lngCurY = T_DrawClient.�̶�����.Top
    
    '�����µ�������
    For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
        lngCurX = lngCurX + lngColStep
        Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng������ = 0 And lngRow < T_BodyStyle.lng������ * T_BodyStyle.lng����, RGB_RED, RGB_BLACK))
    Next
        
    lngCurX = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Left = T_DrawClient.�̶�����.Right
    T_DrawClient.��������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.��������.Right = lngMaxX
    T_DrawClient.��������.Bottom = lngMaxY
    
    T_DrawClient.����������.Left = T_DrawClient.�̶�����.Left
    T_DrawClient.����������.Top = T_DrawClient.�̶�����.Top
    T_DrawClient.����������.Right = lngMaxX
    T_DrawClient.����������.Bottom = lngMaxY
    
    '�������������
    Call DrawLine(lngDc, T_DrawClient.��������.Left, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    
    Set mobjPart = New Collection
    '���̶ȿ�ı�ߣ��ӹ̶������10�п�ʼ��ʶ��
    intLables = 1
    rsTemp.Filter = "��¼��=1"
    rsTemp.Sort = "���,�������" '�������������
    With rsTemp
        Do While Not .EOF
            If Not (bln����ӡ������ = True And !��Ŀ��� = gint����) Then
                '��ʾ�̶ȿ���Ŀ�����Ƽ�����,�����¡�
                If !��Ŀ��� = gint���� Then
                    lngCurX = T_DrawClient.��������.Right
                    lngCurY = T_DrawClient.�̶�����.Top
                    sinY��λ = T_DrawClient.ƫ����Y
                Else
                    lngCurX = T_DrawClient.�̶�����.Left + ((intLables - 1) * T_DrawClient.�̶ȵ�λ)
                    If .AbsolutePosition = (.RecordCount - IIf(gbln�������� = True, 1, 0)) Then
                        lngLableStep = (T_DrawClient.�̶�����.Right - T_DrawClient.�̶�����.Left) - ((intLables - 1) * T_DrawClient.�̶ȵ�λ)
                    Else
                        lngLableStep = T_DrawClient.�̶ȵ�λ
                    End If
                    lngCurY = T_DrawClient.�̶�����.Top
                
                    Set gstdSet = New StdFont
                    gstdSet.Name = "����"
                    gstdSet.Size = 9 * sngScale
                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDc, lngFont)
                    '���������Ŀ������
                    Call SetTextColor(lngDc, zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                    Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.TextHeight(zlCommFun.Nvl(!��¼��)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��¼��)), lngLableStep, , , sngScale)
    '                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!��¼��)), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                    Call DrawText(lngDc, Trim(zlCommFun.Nvl(!��¼��)), -1, T_LableRect, DT_CENTER)
                    Call SelectObject(lngDc, lngOldFont)
                    Call DeleteObject(lngFont)
                    Call ReleaseFontIndirect(objDraw)
                    '���������С
                    Set gstdSet = New StdFont
                    gstdSet.Name = "����"
                    gstdSet.Size = 8 * sngScale
                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDc, lngFont)
        
                    '�����Ŀ��λ
                    Call GetTextRect(objDraw, lngCurX, lngCurY + lngInitRowStep * 2 + objDraw.TextHeight(zlCommFun.Nvl(!��λ)) / T_TwipsPerPixel.Y / 2, Trim(zlCommFun.Nvl(!��λ)), lngLableStep, , , sngScale)
    '                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(zlCommFun.Nvl(!��λ, 0)), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                    Call DrawText(lngDc, Trim(zlCommFun.Nvl(!��λ, 0)), -1, T_LableRect, DT_CENTER)
                    Call SelectObject(lngDc, lngOldFont)
                    Call DeleteObject(lngFont)
                    Call ReleaseFontIndirect(objDraw)
                    If T_LableRect.Bottom > lngInitRowStep * 2 * mintNullRow + T_DrawClient.�̶�����.Top Then
                        sinY��λ = T_LableRect.Bottom
                    Else
                        sinY��λ = lngInitRowStep * 2 * mintNullRow + T_DrawClient.�̶�����.Top
                    End If
                End If
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
            If !��Ŀ��� = gint���� Then
                mobjPart.Add Array("" & !��Ŀ���, "����", arrTemp(0), Nvl(!��¼ɫ, RGB_BLACK), "B"), "B" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, "Ҹ��", arrTemp(1), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, "����", arrTemp(2), Nvl(!��¼ɫ, RGB_BLACK), "C"), "C" & !��Ŀ���
                mobjPart.Add Array("" & !��Ŀ���, "����", arrTemp(3), Nvl(!��¼ɫ, RGB_BLACK), "D"), "D" & !��Ŀ���
            ElseIf !��Ŀ��� = gint���� And int����Ӧ�� = 2 Then
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
                mobjPart.Add Array("" & gint����, CStr(ArrCode(1)), CStr(ArrCode(2)), Val(ArrCode(3)), "A"), "A" & gint����
            Else
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
            End If
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
                    If !��Ŀ��� = gint���� And sin�̶� = Val(Nvl(!���ֵ, 0)) Then
                        strTmp = Trim(Nvl(!��¼��))
                        lngY = lngCurY - Len(strTmp) / 2 * (objDraw.TextHeight("��") / T_TwipsPerPixel.Y)
                        If lngY < sinY��λ Then lngY = sinY��λ
                        For i = 1 To Len(strTmp)
                            Call GetTextRect(objDraw, lngCurX, lngY, Mid(strTmp, i, 1), False, False)
                            Call DrawText(lngDc, Mid(strTmp, i, 1), -1, T_LableRect, DT_CENTER)
                            lngY = lngY + objDraw.TextHeight("��") / T_TwipsPerPixel.Y
                        Next i
                    Else
                        If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY < sinY��λ Then
                            Call GetTextRect(objDraw, lngCurX, sinY��λ, Format(sin�̶�, "#0"), IIf(!��Ŀ��� = gint����, 0, lngLableStep), , , sngScale)
                        ElseIf lngCurY = T_DrawClient.�̶�����.Bottom Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.TextHeight("1") / 2 / T_TwipsPerPixel.Y), Format(sin�̶�, "#0"), IIf(!��Ŀ��� = gint����, 0, lngLableStep), , , sngScale)
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Format(sin�̶�, "#0"), IIf(!��Ŀ��� = gint����, 0, lngLableStep), , , sngScale)
                        End If
    '                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Format(sin�̶�, "#0"), zlCommFun.Nvl(!��¼ɫ, RGB_BLACK))
                        Call DrawText(lngDc, Format(sin�̶�, "#0"), -1, T_LableRect, DT_CENTER)
                    End If
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
                        Call DrawLine(lngDc, T_DrawClient.��������.Left, lngCurAlerY, lngMaxX, lngCurAlerY, intLineMode, intBold, RGB_RED)
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
                Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDc, lngCurX + lng�̶ȿ��, lngCurY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDc, lngCurX, lngMaxY, lngCurX + lng�̶ȿ��, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                '��������еĻ���
                lngCurX = lngCurX + lng�̶ȿ��
                For lngRow = 1 To lngCurveRows
                    '�����µ���������
                    If lngRow <> 0 Then
                        lngCurY = lngCurY + sinRowStep
                    End If
                    If ((blnDoubleRow Or bln˫��) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln˫��) Then
                        Call DrawLine(lngDc, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sinRowStep >= 4 And bln����, intBold, intFine), RGB_BLACK)
                    End If
                Next
                '������
                Call DrawLine(lngDc, lngCurX, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                lngCurY = lngY
                '�����µ�������
                For lngRow = 1 To T_BodyStyle.lng������ * T_BodyStyle.lng����
                    lngCurX = lngCurX + lngColStep
                    Call DrawLine(lngDc, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng������ = 0, intBold, intFine), IIf(lngRow Mod T_BodyStyle.lng������ = 0 And lngRow < T_BodyStyle.lng������ * T_BodyStyle.lng����, RGB_RED, RGB_BLACK))
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
                Call SetFontIndirect(gstdSet, lngDc, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDc, lngFont)
                Call SetTextColor(lngDc, Nvl(!��¼ɫ, RGB_BLACK))
                T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                If T_Size.H * Len(Nvl(!��¼��)) >= lngCurveRows * sinRowStep Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!��¼��)))) \ 2
                End If
                For lngRow = 1 To Len(Nvl(!��¼��))
                    Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Nvl(!��¼��), lngRow, 1), lng�̶ȿ�� \ 2, False)
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Nvl(!��¼��), lngRow, 1), Nvl(!��¼ɫ, RGB_BLACK))
                    Call DrawText(lngDc, Mid(Nvl(!��¼��), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call SelectObject(lngDc, lngOldFont)
                Call DeleteObject(lngFont)
                Call ReleaseFontIndirect(objDraw)
                '�����Ŀ��λ
                lngCurY = lngY: If Nvl(!��¼��) <> "" Then lngCurX = T_LableRect.Right
                If Trim(Nvl(!��λ)) <> "" And Nvl(!��¼��) <> "" Then
                    '���������С
                    Set gstdSet = New StdFont
                    gstdSet.Name = "����"
                    gstdSet.Size = 8 * sngScale
                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDc, lngFont)
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight("��"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(Nvl(!��λ))) >= lngCurveRows * sinRowStep Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * sinRowStep) - (T_Size.H * Len(Nvl(!��λ)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(Nvl(!��λ)))
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Trim(Nvl(!��λ)), lngRow, 1), 0, False)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Mid(Trim(Nvl(!��λ)), lngRow, 1), Nvl(!��¼ɫ, RGB_BLACK))
                        Call DrawText(lngDc, Mid(Trim(Nvl(!��λ)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call SelectObject(lngDc, lngOldFont)
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
                mobjPart.Add Array("" & !��Ŀ���, Nvl(!��¼��), Nvl(!��¼��), Nvl(!��¼ɫ, RGB_BLACK), "A"), "A" & !��Ŀ���
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
                        lngCurX = lngX + lng�̶ȿ�� - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin�̶�, "#0.0"))), vbTwips, vbPixels)
                        lngCurX = lngCurX - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        If sin�̶� = Val(Nvl(!���ֵ, 0)) And lngCurY = lngY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY + (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        ElseIf lngCurY = lngMaxY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin�̶�, "#0.0")))
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin�̶�, "#0.0")))
                        End If
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Val(Format(sin�̶�, "#0.0")), Nvl(!��¼ɫ, RGB_BLACK))
                        Call DrawText(lngDc, Val(Format(sin�̶�, "#0.0")), -1, T_LableRect, DT_CENTER)
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
                            Call DrawLine(lngDc, lngX + lng�̶ȿ��, lngCurAlerY, lngMaxX, lngCurAlerY, PS_SOLID, 1, RGB_RED)
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
       T_DrawClient.����������.Bottom = 2 * mintNullRow * lngInitRowStep + (T_DrawClient.������ + T_DrawClient.��������������) * sinRowStep + T_DrawClient.ƫ����Y
    End With
    '����Ϊ�����ٻ�����
    If gbln�������� = True Then
        Call DrawLine(lngDc, T_DrawClient.��������.Right + T_BodyStyle.lngBreashLabelStep, T_DrawClient.��������.Top, T_DrawClient.��������.Right + T_BodyStyle.lngBreashLabelStep, T_DrawClient.����������.Bottom, PS_SOLID, intBold, RGB_BLACK)
    End If
    
    str˵�� = "˵��:" & Mid(str˵��, 2)
    
    DrawCanvasNew = str˵��
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowPointsNew(ByVal lngDc As Long, ByVal objDraw As Object, ByVal rsPoint As ADODB.Recordset, _
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
    Dim strWaveReview As String '����:���¸��Ժϸ�ķ���
    On Error GoTo Errhand
    
    strWaveReview = zlDatabase.GetPara("���¸��Ժϸ����", glngSys, 1255, "v")
    
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
            If Not ((zlCommFun.Nvl(!��Ŀ���) = gint���� Or zlCommFun.Nvl(!��Ŀ���) = gint��ʹǿ��) And Val(zlCommFun.Nvl(!���)) = 1) Then
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
                    Call DrawLine(lngDc, sinԭX + T_DrawClient.�е�λ / 2, sinԭY, !X���� + T_DrawClient.�е�λ / 2 - lngWith, !Y����, PS_SOLID, intFine, lngRGB)
                End If
                If !�Ͽ� = 0 Then
                    sinԭX = zlCommFun.Nvl(!X����, 0) + lngWith
                    sinԭY = zlCommFun.Nvl(!Y����, 0)
                Else
                    sinԭX = 0
                End If
                
                If !��Ŀ��� = gint���� Then
                    If zlCommFun.Nvl(!����) = 1 Then '���Ժϸ�
                        Call SetTextColor(lngDc, lngRGB)
                        Call GetTextRect(objDraw, !X����, !Y���� - T_DrawClient.�е�λ, strWaveReview, T_DrawClient.�е�λ, True, , sngScale)
'                        Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), strWaveReview, lngRGB)
                        Call DrawText(lngDc, strWaveReview, -1, T_LableRect, DT_CENTER)
                    End If
                End If
                
                dblMinValue = GetMaxMinValue(0, Val(zlCommFun.Nvl(!��Ŀ���)), strEditors)
                dblMaxValue = GetMaxMinValue(1, Val(zlCommFun.Nvl(!��Ŀ���)), strEditors)
                    
                If Not (Val(Nvl(!��Ŀ���)) = gint���� And Trim(Nvl(!��ֵ)) = "����") Then
                    dblvalue = Val(zlCommFun.Nvl(!��ֵ))
                    If dblvalue > dblMaxValue Then
                        Call DrawLine(lngDc, !X���� + T_DrawClient.�е�λ / 2, !Y���� - T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
                    ElseIf dblvalue < dblMinValue Then
                        Call DrawLine(lngDc, !X���� + T_DrawClient.�е�λ / 2, !Y���� + T_DrawClient.�е�λ * 2, !X���� + T_DrawClient.�е�λ / 2, !Y����, PS_SOLID, intFine, lngRGB, True)
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
                    Call DrawLine(lngDc, !X���� + T_DrawClient.�е�λ / 2, !Y���� + (T_Size.H / 4), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
                ElseIf Val(!��ֵ) < Val(dblvalue) Then
                    '�����³ɹ�������ɫ���ߣ��ַ��̶��á�
                    Call DrawLine(lngDc, !X���� + T_DrawClient.�е�λ / 2, !Y���� - (T_Size.H / 2), SinX + T_DrawClient.�е�λ / 2, sinY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, RGB_RED, False)
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
                    Select Case zlCommFun.Nvl(!��Ŀ���)
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
                        If Val(Nvl(!��Ŀ���)) = gint��ʹǿ�� And zlCommFun.Nvl(!���) = 1 Then '��ʹ��ʹ����
                            lngRGB = RGB_RED
                            strChar = "��"
                        Else
                            strChar = strTmp
                        End If
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
                    Call SetTextColor(lngDc, lngRGB)
                    Call GetTextRect(objDraw, !X����, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
                    T_LableRect.Left = T_LableRect.Left - 1
'                    Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                    Call DrawText(lngDc, Trim(strChar), -1, T_LableRect, DT_CENTER)
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
                                Call SetTextColor(lngDc, lngRGB)
                                Call GetTextRect(objDraw, !X���� - 1, !Y����, Trim(strChar), T_DrawClient.�е�λ, True, , sngScale)
'                                Call DrawTextPrint(objDraw, objDraw.ScaleX(T_LableRect.Left, vbPixels, vbTwips), objDraw.ScaleY(T_LableRect.Top, vbPixels, vbTwips), Trim(strChar), lngRGB)
                                Call DrawText(lngDc, Trim(strChar), -1, T_LableRect, DT_CENTER)
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


Public Sub DrawBodyRecordItemNew(ByVal lngDc As Long, ByVal objDraw As Object, strValue() As String, ByVal lngX As Long, ByVal lngY As Long, _
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
    Dim arrTmp() As String, arrText() As String, arrData, arrItem
    Dim blnRed As Boolean, lngMoveX As Long
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer, j As Integer, k As Integer
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
    Dim PartRect As RECT
    '73316�����޸���ر�������
    Dim arrBreathe, blnBreathe As Boolean, intBegin  As Integer, intEnd As Integer
    Dim lngBreatheX As Long, lngBreatheY As Long, lngBottomY As Long
    Dim blnBreatheShowType As Boolean
    
    On Error GoTo Errhand
    
    If UBound(strValue) < 0 Then Exit Sub
    If IsEmpty(strValue) = True Then Exit Sub
    
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
    If int������������ʽ < 0 Or int������������ʽ > 3 Then int������������ʽ = 0
    bln�೦����Է��ӷ�ĸ��ʾ = (Val(zlDatabase.GetPara("�೦������ʾ��ʽ", glngSys, 1255, 0)) = 1)
    '73316:������,2014-06-26,���첿��ҽԺҪ��:
    '��1����������ɫ���ں�������Ӧʱ������д���������κ������½�����д�����Ϻ���
    '��2������������ʶ������ʼ��Ӧʱ������ɫ�ֱ������µ������������Ϸ�����
    '��д�������������á�������ʶ��ʼ����ֹ�ԡ�������ʶ���������趨Ƶ�������ֱ�ʾ������
    'ɫ���ں�������Ӧʱ������д���������κ������½�����д�����Ϻ�
    '2----��ʼ����������� ������Ϊͼ�����
    blnBreatheShowType = (Val(zlDatabase.GetPara("�����������������ʽ", glngSys, 1255, 0)) = 1)
    
    strPic = ""
    If InStr(1, strValue(0), "'") > 0 Then
        arrTestis = Array()
        bln��ʾƤ�� = IIf(Split(strValue(UBound(strValue)), "'")(0) = "-999", True, False)
        
        For intRow = LBound(strValue) To UBound(strValue)
            arrTmpString0 = Split(strValue(intRow), "'")
            arrTmpString1 = Split(arrTmpString0(5), "||")
            
            If intRepairRows > 0 And intRepairRows > intRowCount Then
            
                If arrTmpString0(0) = "3" Then '������Ŀ
                    '��ȡ�����ɫ
                    lngColor = Val(arrTmpString0(4))
                    intRowCount = intRowCount + 1
                    arrTmpString1 = Split(arrTmpString0(5), "||")
                    arrBreathe = Array(): blnBreathe = False
                    For intCOl = 0 To UBound(arrTmpString1)
                        If intCOl = 0 Then '��ͷ
                            Call SetTextColor(lngDc, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                            LPoint.X = lngX
                            LPoint.Y = lngY
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX
                            LPoint.H = mlngBreatheHeight
                            
                            Call DrawTabTextNew(lngDc, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            Call DrawLine(lngDc, lngX, lngY, lngX, lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            Call DrawLine(lngDc, lngX, lngY + mlngBreatheHeight, T_DrawClient.�̶�����.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                            Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY, T_DrawClient.�̶�����.Right, _
                                lngY + mlngBreatheHeight, PS_SOLID, intBold, RGB_BLACK)
                            lngX1 = T_DrawClient.�̶�����.Right
                            lngY1 = lngCurY
                        Else
                            arrTmpString1(intCOl) = arrTmpString1(intCOl) & String(1 - UBound(Split(arrTmpString1(intCOl), ",")), ",")
                            strTmp = Split(arrTmpString1(intCOl), ",")(0)
                            strPart = Split(arrTmpString1(intCOl), ",")(1)
                            If strPart = "" Then strPart = "��������"
                            strPic = ""
                            
                            If strPart <> "��������" And IsNumeric(strTmp) Then
                                If blnBreathe = False Then
                                    ReDim Preserve arrBreathe(UBound(arrBreathe) + 1)
                                    arrBreathe(UBound(arrBreathe)) = intCOl & ";" & intCOl
                                    blnBreathe = True
                                Else
                                    arrBreathe(UBound(arrBreathe)) = Split(arrBreathe(UBound(arrBreathe)), ";")(0) & ";" & intCOl
                                End If
                            Else
                                blnBreathe = False
                            End If
                            
                            '��ӡ����ֵ���������ӡ�� ��һ��ʼ��������
                            If IsNumeric(strTmp) Then
                                If strPart = "��������" Or blnBreatheShowType = True Then
                                    Call SetTextColor(lngDc, lngColor)
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
                                
                                If int������������ʽ = 0 Or int������������ʽ = 2 Then '˳��������ʾ
                                    If intCOl Mod 2 = intValue Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            If int������������ʽ = 0 Then
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                            Else
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                            End If
                                        Else
                                            If int������������ʽ = 0 Then
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            Else
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                    vbPixels, vbTwips), objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                            End If
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            If int������������ʽ = 0 Then
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                            Else
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                            End If
                                        Else
                                            If int������������ʽ = 0 Then
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), _
                                                    vbPixels, vbTwips), objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                            Else
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            End If
                                        End If
                                    End If
                                    
                                Else        '������ʱ����֮��������ʾ
                                    If int����λ�� = 2 Then
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            If int������������ʽ = 1 Then
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                            Else
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                            End If
                                        Else
                                            If int������������ʽ = 1 Then
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            Else
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                            End If
                                        End If
                                    Else
                                        If strPic = "" Then
                                            LPoint.X = lngX1
                                            LPoint.Y = lngY
                                            LPoint.W = T_DrawClient.�е�λ
                                            LPoint.H = mlngBreatheHeight
                                            If int������������ʽ = 1 Then
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 3)
                                            Else
                                                Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale, 1)
                                            End If
                                        Else
                                            If int������������ʽ = 1 Then
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + (mlngBreatheHeight - mintBmpH * IIf(blnPrinter = True, msngTwips, 1)), vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + mlngBreatheHeight, vbPixels, vbTwips), True)
                                            Else
                                                Call DrawPicture(objDraw, strPic, objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1, vbPixels, vbTwips), _
                                                    objDraw.ScaleX(lngX1 + ((T_DrawClient.�е�λ - mintBmpW * IIf(blnPrinter = True, msngTwips, 1)) / 2) + mintBmpW * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), _
                                                    objDraw.ScaleY(lngY + 1 + mintBmpH * IIf(blnPrinter = True, msngTwips, 1), vbPixels, vbTwips), True)
                                            End If
                                        End If
                                    End If
                                   
                                    int����λ�� = int����λ�� + 1
                                    If int����λ�� > 2 Then int����λ�� = 1
                                End If
                                
                            End If
                            lngX1 = lngX1 + T_DrawClient.�е�λ
                        End If
                    Next intCOl
                    
                    If blnBreatheShowType = True Then
                        '��ʼ�ں������Ϸ��������
                        lngBottomY = T_DrawClient.����������.Bottom
                        For i = 0 To UBound(arrBreathe)
                            intBegin = Split(arrBreathe(i), ";")(0)
                            intEnd = Split(arrBreathe(i), ";")(1)
                            '�������������
                            strPart = "������"
                            Call SetTextColor(lngDc, lngColor)
                            T_Size.H = objDraw.TextHeight("��") / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth("��") / T_TwipsPerPixel.X
                            '����GetTextRect����Ĭ�ϸ�X+1���Դ˴�-1
                            If intBegin = intEnd Then
                                If T_DrawClient.�е�λ >= T_Size.W + (6 * msngTwips) Then
                                    lngBreatheX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ + ((T_DrawClient.�е�λ - T_Size.W - (6 * msngTwips)) \ 2) - 1
                                Else
                                    lngBreatheX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ - ((T_Size.W + (6 * msngTwips) - T_DrawClient.�е�λ)) - 1
                                End If
                            Else
                                If T_DrawClient.�е�λ >= T_Size.W + (3 * msngTwips) Then
                                    lngBreatheX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ + ((T_DrawClient.�е�λ - T_Size.W - (3 * msngTwips)) \ 2) - 1
                                Else
                                    lngBreatheX = T_DrawClient.��������.Left + (intBegin - 1) * T_DrawClient.�е�λ - ((T_Size.W + (3 * msngTwips) - T_DrawClient.�е�λ)) - 1
                                End If
                            End If
                            lngBreatheY = lngBottomY - T_Size.H * Len(strPart)
                            For j = 1 To Len(strPart)
                                Call GetTextRect(objDraw, lngBreatheX, lngBreatheY, Mid(strPart, j, 1), 0, False)
                                Call DrawText(lngDc, Mid(strPart, j, 1), -1, T_LableRect, DT_CENTER)
                                lngBreatheY = lngBreatheY + T_Size.H
                            Next j
                            '��ʼ�����ϼ�ͷ�����������¼�ͷ
                            If intBegin = intEnd Then
                                lngBreatheY = T_Size.H * Len(strPart) - T_Size.H
                                lngBreatheX = lngBreatheX + T_Size.W + (3 * msngTwips)
                                Call DrawLine(lngDc, lngBreatheX, lngBottomY - lngBreatheY - (T_Size.H \ 2), lngBreatheX, lngBottomY - (T_Size.H \ 2), PS_SOLID, intFine, lngColor, True)
                                lngBreatheX = lngBreatheX + (3 * msngTwips)
                                Call DrawLine(lngDc, lngBreatheX, lngBottomY - (T_Size.H \ 2), lngBreatheX, lngBottomY - lngBreatheY - (T_Size.H \ 2), PS_SOLID, intFine, lngColor, True)
                            Else
                                lngBreatheY = T_Size.H * Len(strPart) - T_Size.H
                                lngBreatheX = lngBreatheX + T_Size.W + (3 * msngTwips)
                                Call DrawLine(lngDc, lngBreatheX, lngBottomY - lngBreatheY - (T_Size.H \ 2), lngBreatheX, lngBottomY - (T_Size.H \ 2), PS_SOLID, intFine, lngColor, True)
                                lngBreatheX = T_DrawClient.��������.Left + (intEnd - 1) * T_DrawClient.�е�λ + T_DrawClient.�е�λ \ 2
                                Call DrawLine(lngDc, lngBreatheX, lngBottomY - (T_Size.H \ 2), lngBreatheX, lngBottomY - lngBreatheY - (T_Size.H \ 2), PS_SOLID, intFine, lngColor, True)
                            End If
                        Next i
                    End If
                    
                    lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ
                    lngY1 = lngY + mlngBreatheHeight
                    
                    '�����������е���
                    For intCOl = 1 To T_BodyStyle.lng���� * T_BodyStyle.lng������
                        If intCOl Mod T_BodyStyle.lng������ = 0 Then
                            Call DrawLine(lngDc, lngX1, lngY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Else
                            Call DrawLine(lngDc, lngX1, lngY, lngX1, lngY1, PS_SOLID, intFine, RGB_BLACK)
                        End If
                        lngX1 = lngX1 + T_DrawClient.�е�λ
                    Next intCOl
                    Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                    
                    '��ǰY������
                    lngCurY = lngY1
                ElseIf arrTmpString0(0) <> "-999" Then '����Ƥ�Խ��
                    intƵ�� = Val(arrTmpString0(2))
                    '���Ŀ����Ƿ�������ݣ������ھͲ���ӡ����
                    If Val(arrTmpString0(3)) = 2 Then
                        
                        If Trim(Replace(arrTmpString0(5), "||", "")) = "" Then
                            blnDataTrue = False
                        Else
                            blnDataTrue = True
                        End If
                    Else
                        blnDataTrue = True
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

                                Call SetTextColor(lngDc, RGB_BLACK)
                                T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                                T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                            
                                LPoint.X = lngX1
                                LPoint.Y = lngY1 - lngRowHeiht * 2
                                LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                                LPoint.H = lngRowHeiht * 2
                                Call DrawTabTextNew(lngDc, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                                Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                Call DrawLine(lngDc, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                
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
                                
                                Call SetTextColor(lngDc, lngColor)
                                
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
                                        Call DrawAnsyGrade(lngDc, objDraw, arrText, LPoint, lngColor, bln�೦����Է��ӷ�ĸ��ʾ, sngScale)
                                    Else
                                        Call DrawTabTextNew(lngDc, objDraw, strTmp, -1, DT_CENTER, LPoint, sngScale)
                                    End If
                                End If
                   
                            End If
                        Next intCOl
                        
                        '����Ԫ������
                        If InStr(1, ",2,3,4,", "," & intƵ�� & ",") = 0 Then
                            lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                            lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                            For intCOl = 1 To intƵ�� * T_BodyStyle.lng����
                                Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod intƵ�� = 0, intBold, intFine), RGB_BLACK)
                                lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / intƵ��)
                            Next intCOl
                            Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                        ElseIf intƵ�� = 3 Then
                            intRowCount = intRowCount - (intƵ�� - intNum)
                            intValue = intRowCount
                            For i = 1 To 3 - intNum
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * T_BodyStyle.lng������
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                For intCOl = 1 To T_BodyStyle.lng����
                                    Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * T_BodyStyle.lng������
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
                                
                                lngCurY = lngY1
                            Next i
                        ElseIf InStr(1, ",2,4,", "," & intƵ�� & ",") <> 0 Then
                            intRowCount = intRowCount - (intƵ�� / 2 - intNum)
                            intValue = intRowCount
                            For i = 1 To (intƵ�� / 2 - intNum)
                                lngY1 = lngCurY + T_DrawClient.ʱ���е�λ
                                lngX1 = T_DrawClient.�̶�����.Right + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                For intCOl = 1 To T_BodyStyle.lng���� * 2
                                    Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, IIf(intCOl Mod 2 = 0, intBold, intFine), RGB_BLACK)
                                    lngX1 = lngX1 + T_DrawClient.�е�λ * (T_BodyStyle.lng������ / 2)
                                Next intCOl
                                intRowCount = intValue + i
                                Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRowCount = intRepairRows, intBold, intFine), RGB_BLACK)
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
                               
                            Call SetTextColor(lngDc, RGB_BLACK)
                            T_Size.H = objDraw.TextHeight(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.Y
                            T_Size.W = objDraw.TextWidth(arrTmpString0(intCOl + 1)) / T_TwipsPerPixel.X
                
                            LPoint.X = lngX1
                            LPoint.Y = lngY1 - lngTestisHeight
                            LPoint.W = T_DrawClient.�̶�����.Right - lngX1
                            LPoint.H = lngTestisHeight
                            Call DrawTabTextNew(lngDc, objDraw, arrTmpString0(intCOl + 1), -1, DT_CENTER, LPoint, sngScale)
                            
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
                                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                                    lngFont = CreateFontIndirect(T_Font)
                                    lngOldFont = SelectObject(lngDc, lngFont)
                                    lngY1 = LPoint.Y
                                    If Val((UBound(arrData) + 1) * objDraw.TextHeight("��")) < Val(arrTestis(i)) Then
                                        LPoint.Y = LPoint.Y + (Val(arrTestis(i)) - ((UBound(arrData) + 1) * objDraw.TextHeight("��"))) / T_TwipsPerPixel.Y / 2
                                    End If
                                    
                                    '89797:LPF,Ƥ�����ݺͽ��֮���Կո�����,��ֻ�н������ɫ����(���ݺͽ����'չʾ����)
                                    '��ʼ�������
                                    Call SetTextColor(lngDc, 0)
                                    blnRed = False
                                    For j = 0 To UBound(arrData)
                                        '������ʾ
                                        lngMoveX = Format((frmTendFileRead.txtLength.Width - objDraw.TextWidth(CStr(arrData(j)))) \ T_TwipsPerPixel.X \ 2, "#0") - msngTwips
                                        If lngMoveX < 0 Then lngMoveX = 0
                                        arrItem = Split(CStr(arrData(j)), "$")
                                        For k = 0 To UBound(arrItem)
                                            If k > 0 And blnRed = False Then
                                                Call SetTextColor(lngDc, lngColor)
                                                blnRed = True
                                            End If
                                            strTmp = CStr(arrItem(k))
                                            '�ų����һ���ַ�(��һ��)���һ���ַ�(�ڶ���)Ϊ'�����
                                            If k > 0 Then
                                                If arrItem(k - 1) <> "" And arrItem(k) <> "" Then strTmp = " " & strTmp
                                            End If
                                            If strTmp <> "" Then
                                                Call GetTextRect(objDraw, LPoint.X + lngMoveX, LPoint.Y, strTmp, , False, , sngScale)
                                                Call DrawText(lngDc, strTmp, -1, T_LableRect, DT_CENTER)
                                                lngMoveX = lngMoveX + Format(objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X, "#0")
                                            End If
                                        Next k
                                        LPoint.X = lngX1
                                        LPoint.Y = LPoint.Y + Val(Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0"))
                                    Next j
                                    LPoint.Y = lngY1 + Val(Format(Val(arrTestis(i)) / T_TwipsPerPixel.Y, "#0"))
                                    Call SelectObject(lngDc, lngOldFont)
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
                Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDc, T_DrawClient.��������.Right, lngCurY, T_DrawClient.��������.Right, lngCurY + lngTestisHeight, PS_SOLID, intBold, RGB_BLACK)
                Call DrawLine(lngDc, lngX1, lngCurY + lngTestisHeight, T_DrawClient.�̶�����.Right, lngCurY + lngTestisHeight, PS_SOLID, IIf(intRowCount + UBound(arrData) + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
            End If
            For intCOl = 0 To T_BodyStyle.lng���� - 1
                If intCOl = 0 Then
                    Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1 + Val(arrData(i)), T_DrawClient.��������.Right, lngY1 + Val(arrData(i)), PS_SOLID, IIf(intRowCount + i + 1 = intRepairRows, intBold, intFine), RGB_BLACK)
                Else
                    lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * T_BodyStyle.lng������) * intCOl
                    Call DrawLine(lngDc, lngX1, lngY1, lngX1, lngY1 + Val(arrData(i)), PS_SOLID, intBold, RGB_BLACK)
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
                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        Call DrawLine(lngDc, lngX1, lngY1, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
                        Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngCurY, T_DrawClient.�̶�����.Right, lngY1, PS_SOLID, intBold, RGB_BLACK)
                    Else
                        
                        lngX1 = T_DrawClient.�̶�����.Right + (T_DrawClient.�е�λ * T_BodyStyle.lng������) * intCOl
                        Call DrawLine(lngDc, lngX1, lngCurY, lngX1, lngY1, PS_SOLID, intBold, RGB_BLACK)
                        If intCOl = T_BodyStyle.lng���� Then
                            Call DrawLine(lngDc, T_DrawClient.�̶�����.Right, lngY1, T_DrawClient.��������.Right, lngY1, PS_SOLID, IIf(intRow = intRepairRows, intBold, intFine), RGB_BLACK)
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
    
    '����Ϊ����
    If gbln�������� = True Then
        Call DrawLine(lngDc, lngLeft, lngY, lngLeft + T_BodyStyle.lngBreashLabelStep, lngY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDc, lngLeft + T_BodyStyle.lngBreashLabelStep, lngY, lngLeft + T_BodyStyle.lngBreashLabelStep, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        Call DrawLine(lngDc, lngLeft, lngCurY, lngLeft + T_BodyStyle.lngBreashLabelStep, lngCurY, PS_SOLID, intBold, RGB_BLACK)
        '�������˵����Ϣ
        PartRect.Left = lngLeft + msngTwips
        PartRect.Top = lngY
        PartRect.Right = lngLeft + T_BodyStyle.lngBreashLabelStep
        PartRect.Bottom = lngOutY
        Call PrintCurveInfo(objDraw, PartRect, blnPrinter)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DrawBodyPageFooterNew(ByVal lngDc As Long, objDraw As Object, X As Long, Y As Long, ByVal LeftX As Long, ByVal intPageNo As Integer, _
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
    If blnPrintCurveInfo = False And gbln�������� = False Then
        '��ӡ����˵����Ϣ
        Call SetTextColor(lngDc, RGB_BLACK)
        Call GetTextExtentPoint32(lngDc, strInfo, Len(strInfo), T_Size)
        Call GetTextRect(objDraw, X, Y, strInfo, 0, False, , sngScale)
        Call DrawText(lngDc, strInfo, -1, T_LableRect, DT_CENTER)
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
    
    Call SetTextColor(lngDc, RGB_BLACK)
    Call GetTextExtentPoint32(lngDc, strNOPage, Len(strNOPage), T_Size)
    Call GetTextRect(objDraw, 0, Y, strNOPage, objDraw.Width / T_TwipsPerPixel.X, False, , sngScale)
    Call DrawText(lngDc, strNOPage, -1, T_LableRect, DT_CENTER)
    
    '�����ӡ��,����ǰ����Ա����
    '------------------------------------------------------------------------------------------------------------------
    If blnOper = True Then
        strNOPage = "��ӡ��:" & gstrUserName
    
        Call SetTextColor(lngDc, RGB_BLACK)
        Call GetTextExtentPoint32(lngDc, strNOPage, Len(strNOPage), T_Size)
        Call GetTextRect(objDraw, LeftX - objDraw.TextWidth(strNOPage) / T_TwipsPerPixel.X, Y, strNOPage, 0, False, , sngScale)
        Call DrawText(lngDc, strNOPage, -1, T_LableRect, DT_CENTER)
    End If

    Y = Y + T_Size.H / 2
    '--------------------------------------------------------------------------------------------------------------------------------
End Sub


Private Sub DrawDeviceCapsNew(ByVal lngDc As Long, ByVal objDraw As Object)
    Dim dblSureW As Double, dblSureH As Double
    '����Ǵ�ӡԤ��,Ӧ����ӡ���Ŀɴ�ӡ�Ŀ�ʼ����ʼԤ��
    dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
    dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
    On Error Resume Next
    Call DrawRect(lngDc, (objDraw.Width * dblSureW) / T_TwipsPerPixel.X, (objDraw.Height * (1 - dblSureH)) / T_TwipsPerPixel.Y, _
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

Private Sub DrawTabTextNew(ByVal lngDc As Long, ByVal objDraw As Object, ByVal strTmp As String, ByVal nCount As Long, ByVal wFormat As Long, LPoint As T_LPoint, Optional sngScale As Single = 1, Optional ByVal bytCenterType As Byte = 2, Optional sgnFontSize As Single = 9)
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
    
    If Trim(strTmp) = "" Then Exit Sub
    
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
            If (objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X) > LPoint.W Then
                If sgnSize > 7 Then
                    GoTo ErrGoTo
                Else
                    blnChage = True
                End If
            End If
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
    Call SetFontIndirect(stdSet, lngDc, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDc, lngFont)
    
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
            Call DrawText(lngDc, CStr(arrData(i)), nCount, T_LableRect, wFormat)
            lngCurY = lngCurY + Val(Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0"))
        Next i
    Else
        Call GetTextRect(objDraw, lngCurX, lngCurY, strTmp, LPoint.W, False, , sngScale)
        Call DrawText(lngDc, strTmp, nCount, T_LableRect, wFormat)
    End If
    
    Call SelectObject(lngDc, lngOldFont)
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


Private Sub DrawAnsyGrade(ByVal lngDc As Long, ByVal objDraw As Object, arrText() As String, LPoint As T_LPoint, ByVal lngColor As Long, Optional ByVal blnFormat As Boolean = False, Optional sngScale As Single = 1)
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
        Call SetFontIndirect(stdOldset, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngColor)
        Call DrawText(lngDc, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
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
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngColor)
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str2) / T_TwipsPerPixel.X) \ 2
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        If T_LableRect.Top < LPoint.Y - Val(Format(LPoint.H / 2, "#0")) Then T_LableRect.Top = LPoint.Y - Val(Format(LPoint.H / 2, "#0"))
        T_LableRect.Bottom = LPoint.Y + Val(Format(LPoint.H / 2, "#0"))
        Call DrawText(lngDc, str2, -1, T_LableRect, 0)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(objDraw)
        '������
        objDraw.Font.Size = 9 * sngScale
        Call DrawLine(lngDc, lngX, lngY, lngX + lngMaxWidth, lngY)
        '�����ĸ
        intSize = 7
        objDraw.Font.Size = intSize * sngScale
        lngY = lngY
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str3) / T_TwipsPerPixel.X) \ 2
        T_LableRect.Top = lngY
        Set stdSet = New StdFont
        stdSet.Name = "����"
        stdSet.Size = intSize * sngScale
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngColor)
        Call DrawText(lngDc, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
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
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < (LPoint.Y - Val(Format(LPoint.H / 2, "#0"))) Then T_LableRect.Top = (LPoint.Y - Val(Format(LPoint.H / 2, "#0")))
            Call DrawText(lngDc, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            Call ReleaseFontIndirect(objDraw)
            '�����벿��
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDc, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        Else
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngColor)
            Call DrawText(lngDc, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDc, lngOldFont)
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


Public Sub CreatePolyNew(rsPoint As ADODB.Recordset, ByVal objDraw As Object, ByVal lngDc As Long, ByVal strBeginDate As String, ByVal str�������� As String, ByVal bln�������� As Boolean)

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
                Call DrawLine(lngDc, recttmp.Left + T_DrawClient.�е�λ / 2, recttmp.Top, SinX + T_DrawClient.�е�λ / 2, sinY, PS_SOLID, intFine, RGB_RED)
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
                Call DrawPoly(lngDc, PtInPoly, lng��䷽ʽ, UBound(Split(str����, ",")) + 1)
                '76697,LPF,����66628��������Ĵ���
                '����ţ�66628,�޸��ˣ�����,�������Ṳ��ʱ�����ͼ�Σ�ֱ�����������ʵ�������ߡ�
                If lng��䷽ʽ = 2 And bln�������� Then
                    i = 0: j = 0
                    For i = 0 To UBound(PtIn����)
                        For j = 0 To UBound(PtIn����)
                            If PtIn����(j).X = PtIn����(i).X Then
                                Call DrawLine(lngDc, PtIn����(j).X, PtIn����(j).Y, PtIn����(i).X, PtIn����(i).Y, PS_SOLID, intFine, RGB_RED)
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



Public Sub OutPutTextNew(ByVal objDraw As Object, ByVal rsDrawItems As ADODB.Recordset, ByVal lngDc As Long, ByVal rsNote As ADODB.Recordset, _
    ByVal mstrBeginDate As String, Optional ByVal blnBeginTop As Boolean = False, Optional ByVal sngScale As Single = 1)

    'rsDrawItems  ��¼��Ŀ��������� ��λֵ�Ȼ�����Ϣ
    'rsNote ����˵����Ϣ
    'mstrBeginDate ���µ�ÿҳ��ʼʱ��
    'blnBeginTop ���������Զ���־�Ǵ����߶����������42�ȿ�ʼ���
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
    If blnBeginTop = False Then
        dblCurveHeight = Format(GetYCoordinate(objDraw, rsDrawItems, gint����, 40) - GetYCoordinate(objDraw, rsDrawItems, gint����, 42), "#0.00")
    Else
        dblCurveHeight = Format(GetYCoordinate(objDraw, rsDrawItems, gint����, 40) - T_DrawClient.��������.Top, "#0.00")
    End If
    
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
                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    lngOldFont = SelectObject(lngDc, lngFont)
                    dblHeight = GetFontHeight(objDraw, zlCommFun.Nvl(!����))
                    Call SelectObject(lngDc, lngOldFont)
                    Call DeleteObject(lngFont)
                    '��ԭ����
                    gstdSet.Name = "����"
                    gstdSet.Size = 9
                    Call SetFontIndirect(gstdSet, lngDc, objDraw)
                    lngFont = CreateFontIndirect(T_Font)
                    Call SelectObject(lngDc, lngFont)
                    !�߶� = dblHeight
                    .Update
                End If
            End If
            .MoveNext
        Loop
        
        If blnBeginTop = False Then
            lngY = GetYCoordinate(objDraw, rsDrawItems, gint����, 42)
        Else
            lngY = T_DrawClient.��������.Top
        End If
        
        '�������ת ���������䵽�����X�����ж���ʽ��Y����
        If bln�¼���ʾ���� = False Then
            .Filter = "��ӡX����=" & lngMaxX & " And ����<>1"
            .Sort = "ʱ��,��Ŀ���"
    
            Do While Not .EOF
                !Y���� = lngY
                .Update
                lngY = lngY + Val(!�߶�) + T_DrawClient.�е�λ
                .MoveNext
            Loop
        End If
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
                Call SetFontIndirect(gstdSet, lngDc, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                lngOldFont = SelectObject(lngDc, lngFont)
                T_Size.H = objDraw.TextHeight("1") / T_TwipsPerPixel.Y
                    
                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.��������.Bottom Then
                        strText = Mid(strComment, i, 1)
                        
                        If Asc(strText) < 0 Then
                            If intAscCharNum Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If

                        '���������Ϣ
                        Call DrawRotateText(objDraw, lngDc, X, Y, strText, !��ɫ, sngScale, IIf(blnCenter = True, -999, objDraw.TextWidth("��") / T_TwipsPerPixel.X))

                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
                Call SelectObject(lngDc, lngOldFont)
                Call DeleteObject(lngFont)
                
                gstdSet.Name = "����"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, lngDc, objDraw)
                lngFont = CreateFontIndirect(T_Font)
                Call SelectObject(lngDc, lngFont)
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
    Dim strSQL As String, strItemNO As String
    Dim strItemInfo As String, ArrCode, arrItem, i As Integer
    Dim blnOK As Boolean, intƵ�� As Integer
    Dim intHour As Integer, strBegin As String, strEnd As String, dtDate As Date, strCurTime As String
    Dim bln���ܵ��� As Boolean
    
    On Error GoTo Errhand
    
    Set rsTemp = GetGridItem(int����ȼ�, byt���ò���, lng����ID, 2)
    If rsTemp.RecordCount = 0 Then
        '�����ڻ��Ŀֱ����ȡ�̶������Ŀ
        Set rsTemp = GetGridItemNew(str�����Ŀ)
        Set GetAppendGridItemNew = rsTemp
        Exit Function
    End If
    strItemInfo = ""
    With rsTemp
        Do While Not .EOF
            strItemNO = strItemNO & "," & !��Ŀ���
            strItemInfo = strItemInfo & ";" & Nvl(!��Ŀ���) & "," & Nvl(!��Ŀ����, 0) & "," & Nvl(!��Ŀ��ʾ, 0) & "," & Nvl(!��¼Ƶ��, 2)
            .MoveNext
        Loop
    End With
    strItemInfo = Mid(strItemInfo, 2)
    ArrCode = Split(strItemInfo, ";")
    strItemNO = "," & Mid(strItemNO, 2) & ","
    
    bln���ܵ��� = (Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    
    If bln���ܵ��� = True Then
        strBegin = Format(dt��ʼʱ��, "yyyy-mm-dd hh:mm:ss")
        strEnd = Format(DateAdd("d", 1, dt����ʱ��), "yyyy-mm-dd hh:mm:ss")
        dtDate = dt����ʱ��
    Else
        strBegin = Format(DateAdd("d", -1, dt��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
        strEnd = Format(dt����ʱ��, "yyyy-mm-dd hh:mm:ss")
        dtDate = DateAdd("d", -1, dt��ʼʱ��)
    End If
    '86368:��ȡ����ֵ�Ļ��Ŀ,����ܴ��ڻ�Ļ�����Ŀ,����Ҫ���⴦��
    gstrSQL = " Select Max(a.����ʱ��) ���ʱ��, Min(a.����ʱ��) ��Сʱ��, c.���²�λ, c.��Ŀ���" & vbNewLine & _
        " From ���˻����ļ� b, ���˻������� a, ���˻�����ϸ c" & vbNewLine & _
        " Where b.Id = a.�ļ�id And a.Id = c.��¼id And b.Id = [1] And b.����id = [2] And b.��ҳid = [3] And Nvl(b.Ӥ��, 0) = [4] And" & vbNewLine & _
        "      c.��¼���� = 1 And a.����ʱ�� Between [5] And [6] And c.��ֹ�汾 Is Null And" & vbNewLine & _
        "      Instr([7], ',' || c.��Ŀ��� || ',') > 0" & vbNewLine & _
        " Group By c.��Ŀ���, c.���²�λ"
    If blnMove Then
        gstrSQL = Replace(gstrSQL, "���˻����ļ�", "H���˻����ļ�")
        gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        gstrSQL = Replace(gstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���Ŀ", lng�ļ�ID, lng����ID, lng��ҳID, intӤ��, CDate(strBegin), CDate(strEnd), strItemNO)
     
    strItemInfo = ""
    For i = 0 To UBound(ArrCode)
        arrItem = Split(ArrCode(i), ",")
        intƵ�� = Val(arrItem(3))
        rsTemp.Filter = "��Ŀ���=" & Val(arrItem(0))
        Do While Not rsTemp.EOF
            blnOK = False
            If Format(rsTemp!���ʱ��, "YYYY-MM-DD HH:mm:ss") >= Format(dt��ʼʱ��, "yyyy-mm-dd hh:mm:ss") And Format(rsTemp!���ʱ��, "YYYY-MM-DD HH:mm:ss") <= Format(dt����ʱ��, "yyyy-mm-dd hh:mm:ss") Or _
                    Format(rsTemp!��Сʱ��, "YYYY-MM-DD HH:mm:ss") >= Format(dt��ʼʱ��, "yyyy-mm-dd hh:mm:ss") And Format(rsTemp!��Сʱ��, "YYYY-MM-DD HH:mm:ss") <= Format(dt����ʱ��, "yyyy-mm-dd hh:mm:ss") Then
                    blnOK = True
            End If
            If Val(arrItem(1)) = 0 And Val(arrItem(2)) = 4 And blnOK = False Then '���ܻ��Ŀ
                If intƵ�� = 0 Or intƵ�� > 2 Then intƵ�� = 2
                mrsTabTime.Filter = "����=2 and Ƶ��=" & intƵ��
                Do While Not mrsTabTime.EOF
                    intHour = CInt(24 / intƵ��)
                    strBegin = Format(IIf(IsDate(Trim(Nvl(mrsTabTime!��ʼ))) = False, (Val(Nvl(mrsTabTime!���)) - 1) * intHour & ":00:00", mrsTabTime!��ʼ), "HH:mm:ss")
                    strEnd = Format(IIf(IsDate(Trim(Nvl(mrsTabTime!����))) = False, Val(Nvl(mrsTabTime!���)) * intHour - 1 & ":59:59", mrsTabTime!����), "HH:mm:ss")
                    'ȷ��Ƶ��ʱ�䷶Χ
                    If Val(Nvl(mrsTabTime!���)) = intƵ�� Then
                        If strBegin >= strEnd Then
                            strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                            strEnd = Format(DateAdd("d", 1, CDate(dtDate)), "YYYY-MM-DD") & " " & strEnd
                        Else
                            strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                            strEnd = Format(dtDate, "YYYY-MM-DD") & " " & strEnd
                        End If
                    Else
                        If strBegin >= strEnd Then
                            strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                            strEnd = strBegin
                        Else
                            strBegin = Format(dtDate, "YYYY-MM-DD") & " " & strBegin
                            strEnd = Format(dtDate, "YYYY-MM-DD") & " " & strEnd
                        End If
                    End If
                    strBegin = Format(strBegin, "YYYY-MM-DD HH:mm:ss")
                    strEnd = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
                    
                    If bln���ܵ��� = True Then
                        strCurTime = Format(rsTemp!��Сʱ��, "YYYY-MM-DD HH:mm:ss")
                    Else
                        strCurTime = Format(rsTemp!���ʱ��, "YYYY-MM-DD HH:mm:ss")
                    End If
                    If strCurTime >= strBegin And strCurTime <= strEnd Then
                        blnOK = True
                        Exit Do
                    End If
                    mrsTabTime.MoveNext
                Loop
            End If
            If blnOK = True Then
                If InStr(1, "," & strItemInfo & ",", "," & rsTemp!��Ŀ��� & ":" & Nvl(rsTemp!���²�λ) & ",") = 0 Then
                    strItemInfo = strItemInfo & "," & rsTemp!��Ŀ��� & ":" & Nvl(rsTemp!���²�λ)
                End If
            End If
            rsTemp.MoveNext
        Loop
     Next i
     strItemInfo = Mid(strItemInfo, 2)
    '��ȡ�̶������Ŀ
    strSQL = strSQL & "Select A.�������,A.��Ŀ���,'' ���²�λ,A.��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,nvl(D.C2,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
        "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
        "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C,TABLE(CAST(F_NUM2LIST2([1]) AS ZLTOOLS.T_NUMLIST2)) D" & _
        "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And B.��Ŀ���=D.C1 And A.��¼��=2 And B.��Ŀ����=1"
    
    '��ȡ���Ŀ
    If strItemInfo <> "" Then
        strSQL = strSQL & vbCrLf & " UNION ALL " & vbCrLf & _
            "   Select A.�������,A.��Ŀ���,D.C2 ���²�λ,D.C2 || A.��¼�� ��¼��,A.��¼��,A.��¼��,A.��¼ɫ,A.���ֵ,A.��Сֵ,A.��λֵ,NVL(A.��¼Ƶ��,2) ��¼Ƶ��,A.��Ժ�ײ�,B.��Ŀ����," & _
            "   B.������,B.��Ŀֵ��,B.��Ŀ��ʾ,B.��Ŀ����,B.��Ŀ����,B.��ĿС��,B.��Ŀ��λ ��λ" & _
            "   From ���¼�¼��Ŀ A,�����¼��Ŀ B,����������Ŀ C,TABLE(CAST(F_str2LIST2([2]) AS ZLTOOLS.T_strLIST2)) D" & _
            "   Where A.��Ŀ���=B.��Ŀ��� And B.��ĿID=C.Id(+) And B.��Ŀ���=D.C1 And A.��¼��=2 And B.��Ŀ����=2"
    End If
    gstrSQL = "Select /*+ Rule*/ �������,��Ŀ���,���²�λ,��¼��,��¼��,��¼��,��¼ɫ,���ֵ,��Сֵ,��λֵ,��¼Ƶ��,��Ժ�ײ�,��Ŀ����," & _
        "   ������,��Ŀֵ��,��Ŀ��ʾ,��Ŀ����,��Ŀ����,��ĿС��,��λ" & _
        "   From (" & strSQL & ") order by Decode(��Ŀ���,3 ,0,1 ),�������,��¼��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����Ŀ", str�����Ŀ, strItemInfo)
    
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
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0  AND B.������� = 'Z'" & vbNewLine & _
                "                AND Instr(',3,5,11,', ',' || c.�������� || ',') > 0 And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
            '��ȡ���˳�Ժǰ��ʱ����Ϣ
            '------------------------------------------------------------------------------------------------------------------
            strSQL = " SELECT  DECODE(D.��ʼʱ��,NULL,DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��)," & vbNewLine & _
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
            
            strSQL = " Select /*+ RULE */ A.��Ŀ���,A.�������,A.��¼��,C.��Ŀֵ��,A.��¼��,A.��¼��,A.��¼ɫ,nvl(A.���ֵ,0) ���ֵ ,nvl(A.��Сֵ,0) ��Сֵ,A.�ٽ�ֵ," & _
                "nvl(A.��λֵ,0) ��λֵ,A.�̶ȼ��,A.��ʾ��,C.��Ŀ��λ ��λ,Decode(��¼��,3,A.�����,nvl(A.�����,2)-2) AS �����,B.��λ,DECODE(A.��Ŀ���,3,1,0) ���" & _
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
            '83477:LPF,Ƥ�Խ����ȡSQL����
            strSQL = _
                " Select ʱ��, f_List2str(Cast(Collect(ҩ����) As t_Strlist)) ҩ����" & vbNewLine & _
                " From (Select To_Char(a.��ʼִ��ʱ��, 'YYYY-MM-DD') ʱ��," & vbNewLine & _
                "              Decode(Ƥ�Խ��, '(+)', 255, '(����)', 255, " & lngColor & ") || '-#' ||" & vbNewLine & _
                "               Replace(Replace(Replace(Decode(b.�Թܱ���, Null, a.ҽ������, b.�Թܱ���), ',', ''), '-#', ''), 'Ƥ��', '') || '$' || Replace(Replace(a.Ƥ�Խ��,'(',''),')','') ҩ����" & vbNewLine & _
                "       From ����ҽ����¼ a, ������ĿĿ¼ b" & vbNewLine & _
                "       Where a.������Ŀid = b.Id And a.������� = 'E' And b.�������� = '1' And a.ҽ��״̬ = 8 And a.Ƥ�Խ�� Is Not Null And a.Ƥ�Խ�� <> '����' And" & vbNewLine & _
                "             a.����id = [1] And a.��ҳid = [2] And a.Ӥ�� = [3] And a.��ʼִ��ʱ�� Between [4] And [5]" & vbNewLine & _
                "       Order By a.��ʼִ��ʱ��, a.Ƥ�Խ��)" & vbNewLine & _
                " Group By ʱ��"
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
        strSQLText = Replace(strSQLText, "����ҽ����¼", "H����ҽ����¼")
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

Public Sub GetPainDegreeNO()
    Dim rsTemp As New ADODB.Recordset
'����:��ȡ��ʹǿ�ȵ���Ŀ���
    '74547:������,2014-07-07,��ʹǿ��Ҳ����¼���ʹ���������µ�������
    '������ʹǿ�Ȳ��ǹ̶���Ŀ��Ŀǰֻ�ܸ�����Ŀ��������
    gint��ʹǿ�� = -999
    On Error GoTo Errhand
    gstrSQL = "Select ��Ŀ��� from ���¼�¼��Ŀ Where ��¼��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���¼�¼��Ŀ", "��ʹǿ��")
    If rsTemp.RecordCount > 0 Then
        gint��ʹǿ�� = Val(rsTemp!��Ŀ���)
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetAdviceOutTime(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngӤ�� As Long) As String
'����:��ȡ���˻�Ӥ����ҽ����Ժʱ��
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String, strTime As String
    On Error GoTo Errhand
    If lngӤ�� = 0 Then
        strTmp = ",5,11,"
    Else
        strTmp = ",3,5,11,"
    End If
    gstrSQL = "Select ��ʼִ��ʱ��" & vbNewLine & _
        " From ����ҽ����¼ b, ������ĿĿ¼ c" & vbNewLine & _
        " Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And instr([4],',' || c.�������� || ',',1)>0 And" & vbNewLine & _
        "      b.����id = [1] And b.��ҳid = [2] And b.Ӥ�� = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҽ����Ժʱ��", lng����ID, lng��ҳID, lngӤ��, strTmp)
    If rsTemp.RecordCount > 0 Then strTime = Format(rsTemp!��ʼִ��ʱ��, "YYYY-MM-DD HH:mm:ss")
    GetAdviceOutTime = strTime
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetElementValue(ByVal strElementName As String, ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lngӤ�� As Long = 0, Optional ByVal strDate As String = "") As String
'����:��ȡ��ӦҪ����Ϣ
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strValue As String
    
    On Error GoTo Errhand
    If IsDate(strDate) Then
        strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,[4],[5]) As ��Ϣֵ From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Replace_Element_Value", strElementName, lng����ID, lng��ҳID, lngӤ��, CDate(strDate))
    Else
        strSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,[4]) As ��Ϣֵ From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Zl_Replace_Element_Value", strElementName, lng����ID, lng��ҳID, lngӤ��)
    End If
    If rsTmp.BOF = False Then
        strValue = Nvl(rsTmp!��Ϣֵ)
    Else
        strValue = ""
    End If
    GetElementValue = strValue
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub PrintCurveInfo(ByVal objDraw As Object, T_RECT As RECT, Optional ByVal blnPrint As Boolean = False)
'���±����������������߲�λ������Ϣ
    Dim lngDc As Long
    Dim i As Integer, j As Integer, lngColor As Long
    Dim arrPar, strPart As String, strSymbol As String
    Dim LPoint As T_LPoint, lngLeading As Long
    Dim lngHeight As Long, lngWidth As Long
    Dim objHeat As New Collection
    Dim lngFont As Long, lngOldFont As Long
    Dim stdSet As New StdFont
    
    If gbln�������� = False Or mobjPart Is Nothing Then Exit Sub
    '�������ƶ���ǰ��
    For j = 0 To 1
        For i = 1 To mobjPart.Count
            arrPar = mobjPart.Item(i)
            If (Val(arrPar(0)) = gint���� And j = 0) Or (Val(arrPar(0)) <> gint���� And j = 1) Then
                objHeat.Add arrPar, CStr(arrPar(4)) & Val(arrPar(0))
            End If
        Next i
    Next j
    
    lngWidth = T_RECT.Right - T_RECT.Left
    lngHeight = T_RECT.Bottom - T_RECT.Top
    '�����м��
    lngLeading = (lngHeight - objHeat.Count * (objDraw.TextHeight("��") / T_TwipsPerPixel.Y)) \ objHeat.Count
    If lngLeading > Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0") Then
        lngLeading = Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0")
    End If
    
    stdSet.Size = 9
    stdSet.Name = "����"
    Call SetFontIndirect(stdSet, lngDc, objDraw)
    lngFont = CreateFontIndirect(T_Font)
    lngOldFont = SelectObject(lngDc, lngFont)


    LPoint.Y = T_RECT.Top + lngLeading
    lngDc = objDraw.hDC
    For i = 1 To objHeat.Count
        arrPar = objHeat.Item(i)
        strPart = CStr(arrPar(1))
        strSymbol = CStr(arrPar(2))
        lngColor = Val(arrPar(3))
        LPoint.X = T_RECT.Left
        Call SetTextColor(lngDc, RGB_BLACK)
        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strPart, , False)
        Call DrawText(lngDc, strPart, -1, T_LableRect, DT_CENTER)
        LPoint.X = LPoint.X + Format(objDraw.TextWidth(strPart) / T_TwipsPerPixel.X, "#0")
        LPoint.W = T_BodyStyle.lngBreashLabelStep - Format(objDraw.TextWidth(strPart) / T_TwipsPerPixel.X, "#0")
        LPoint.X = LPoint.X + (LPoint.W - Format(objDraw.TextWidth(strSymbol) / T_TwipsPerPixel.X, "#0")) \ 2 - IIf(blnPrint = False, 1, 0)
        Call SetTextColor(lngDc, lngColor)
        Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strSymbol, , False)
        Call DrawText(lngDc, strSymbol, -1, T_LableRect, DT_CENTER)
        LPoint.Y = LPoint.Y + Format(objDraw.TextHeight("��") / T_TwipsPerPixel.Y, "#0") + lngLeading
    Next i
    
    Call SelectObject(lngDc, lngOldFont)
    Call DeleteObject(lngFont)
    Call ReleaseFontIndirect(objDraw)
End Sub

