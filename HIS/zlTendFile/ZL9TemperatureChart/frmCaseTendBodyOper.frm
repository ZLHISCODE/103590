VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCaseTendBodyOper 
   Caption         =   "��������/����"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7785
   Icon            =   "frmCaseTendBodyOper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7785
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   -15
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4995
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   135
      ScaleHeight     =   4965
      ScaleWidth      =   8100
      TabIndex        =   1
      Top             =   255
      Width           =   8130
      Begin zl9TemperatureChart.VsfGrid vsfOper 
         Height          =   3810
         Left            =   -15
         TabIndex        =   2
         Top             =   510
         Width           =   7665
         _ExtentX        =   7011
         _ExtentY        =   1005
      End
      Begin VB.PictureBox picDate 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   135
         ScaleHeight     =   360
         ScaleWidth      =   2505
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   2505
         Begin MSComCtl2.DTPicker dtpDate 
            Height          =   300
            Left            =   510
            TabIndex        =   6
            Top             =   30
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   122880003
            CurrentDate     =   42285
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "����:"
            Height          =   180
            Left            =   45
            TabIndex        =   7
            Top             =   75
            Width           =   450
         End
         Begin VB.Image imgDefault 
            Height          =   255
            Left            =   600
            Top             =   840
            Width           =   255
         End
         Begin VB.Image imgbtn 
            Height          =   240
            Index           =   0
            Left            =   2250
            Picture         =   "frmCaseTendBodyOper.frx":6852
            Top             =   45
            Width           =   240
         End
         Begin VB.Image imgbtn 
            Height          =   240
            Index           =   1
            Left            =   1920
            Picture         =   "frmCaseTendBodyOper.frx":7254
            Stretch         =   -1  'True
            Top             =   45
            Width           =   255
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5445
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodyOper.frx":7C56
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10821
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "��Ϣ��ʾ��Ϣ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLDataType"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsDate 
      Left            =   9015
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":84EA
            Key             =   "preGreen"
            Object.Tag             =   "preGreen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":8EFC
            Key             =   "preGray"
            Object.Tag             =   "preGray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":990E
            Key             =   "nextGreen"
            Object.Tag             =   "nextGreen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":A320
            Key             =   "nextGray"
            Object.Tag             =   "nextGray"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":AD32
            Key             =   "preLight"
            Object.Tag             =   "preLight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyOper.frx":B744
            Key             =   "nextLight"
            Object.Tag             =   "nextLight"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   15
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodyOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng�ļ�ID As Long
    lngӤ�� As Long
    lng����ID As Long
    lng����ȼ� As Long
    lng����ID As Long
    lng��ʽID As Long
End Type
Private mT_Patient As type_Patient

Private Enum TYPE_Oper
    Col_OperNull = 0
    Col_OperTime = 1
    Col_OperType = 2
End Enum

Private mcbrToolBar As CommandBar
Private mblnChage  As Boolean
Private Const mFontSize As Integer = 9 '���������ʼ��СΪ9������
Private mstrTime As String
Private mstrDate As String
Private mstrBTime As String     '���µ���ʼʱ��
Private mstrETime As String     '���µ�����ʱ��
Private mstrOverDate As String
Private mstrPreOutDate As String
Private mintPreDays As Integer
Private mlngHours As Long
Private mstrSQL As String
Private mintBigSize As Integer  '�����С
Private mblnMove As Boolean     '�Ƿ�ת��
Private mblnFileBack As Boolean '�Ƿ�鵵
Private mbln��Ժ As Boolean
Private mblnOK As Boolean 'ˢ�����µ���ͼ


Private mrsOper As New ADODB.Recordset '����


Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int����Ӧ�� As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������µ��༭����
    '����:frmParent ������,strParam ��ʽ:����ID;��ҳId;�ļ�ID;Ӥ��;����ID;������ȼ�  strTime ĳ��ʱ���ʱ�䷶Χ ����:2011-01-25 00:00:00;2011-01-25 05:59:59
    
    '     strDayTime һ�ܿ�ʼʱ��; int����Ӧ��=2 ��ʾ���������ʹ��� blnMove ��ʷ�����Ƿ�ת��
    '     bytSize 0-9������ 1-12������
    '----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
    mT_Patient.lng����ID = 0
    mT_Patient.lng����ȼ� = 3
    mblnMove = False
    mblnOK = False
    mT_Patient.lng����ID = Val(arrParam(0))
    mT_Patient.lng��ҳID = Val(arrParam(1))
    mT_Patient.lng�ļ�ID = Val(arrParam(2))
    mT_Patient.lngӤ�� = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng����ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng����ȼ� = arrParam(5)
    
    If mT_Patient.lng����ID = 0 And mT_Patient.lng��ҳID = 0 And mT_Patient.lng����ID = 0 Then
        MsgBox "�ļ�ID,����ID,��ҳID����Ϊ��,����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not OpenPatientInfo Then Exit Function
    mstrDate = strDayTime
    mstrTime = strTime
    If Not ChekPatientOut(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mblnMove = blnMove
    
    '����ļ��Ƿ�鵵
    mblnFileBack = CheckFileBack(mT_Patient.lng�ļ�ID, mblnMove)
    Call InitCommandBars
    '��ȡ����
    Call InitTabOper
    Call zlRefreshData
    Me.Show 1
    
    ShowEditor = mblnOK
End Function


Public Function ChekPatientOut(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBaby As Long) As Boolean
    '-----------------------------------------------------------------------------------------------
    '����:��ȡ���µ���ʼʱ��ͽ���ʱ�� ����鲡���Ƿ��Ժ
    '-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    mbln��Ժ = False
    On Error GoTo Errhand
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys)), 1, 6))
    If mintPreDays < 0 Then mintPreDays = 0
    
    '��ȡ����Ԥ��Ժʱ��
    strSQL = "Select ��ʼʱ�� From ���˱䶯��¼ where ����ID=[1] and ��ҳID=[2] And ��ʼԭ��=10"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then mstrPreOutDate = Format(rsTemp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ),����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "(SELECT " & vbNewLine & _
                "        ����ID, ��ҳID, Ӥ��ʱ��, DECODE(NVL(Ӥ��, 0), 0, DECODE(NVL(��Ժ����, ''), '', 0, 1), DECODE(NVL(Ӥ��ʱ��, ''), '', 0, 1)) ��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID, A.��ҳID, B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����, B.Ӥ��" & vbNewLine & _
                "              FROM ������ҳ A," & vbNewLine & _
                "                   (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                     FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                     WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND NVL(B.Ӥ��, 0) <> 0 AND B.������� = 'Z' " & vbNewLine & _
                "                      AND Instr(',3,5,6,11,', ',' || c.�������� || ',') > 0 AND B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "              WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "              ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '˵��:Ŀǰ����ר�����µ������˿���ͬʱ���ڶ�����µ������µ���ʼʱ�����ֹʱ��Ĺ�������:
    '����ļ��Ŀ�ʼʱ�䲻Ϊ�ղ��Ҵ��ڵ��ڲ�����Ժʱ���Ӥ������ʱ��,���µ��Ŀ�ʼʱ�����ļ���ʼʱ��Ϊ׼,�����Բ�����Ժʱ���Ӥ������ʱ��Ϊ׼
    '����ļ�����ֹʱ�䲻Ϊ�ղ���С�ڵ��ڲ��˻�Ӥ����Ժʱ�䣨δ��Ժ���ܴ��ڵ�ǰʱ�䣩,���µ�����ʱ�����ļ���ʼʱ��Ϊ׼���������µ�����ʱ���Բ��˻�Ӥ����Ժʱ��Ϊ׼(δ��ԺΪ��ǰʱ��)
    '����ļ�����ֹʱ��Ϊ��,����ԭ�з�ʽ,��������Ѿ���Ժ�����ѳ�Ժʱ��Ϊ׼,δ��Ժ���ѵ�ǰʱ������ݽ���ʱ��Ϊ׼.
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

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        mbln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        MsgBox "�޴˲��˱���סԺ��Ϣ,����!", vbInformation, gstrSysName '�������˱䶯��Ϣ�˳�
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln��Ժ Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '���˳�Ժ�Գ�Ժʱ��Ϊ��ֹʱ��
    If mbln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
        mstrETime = Format(RetrunEndTimeNew(CDate(mstrBTime), CDate(mstrETime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        strMaxDate = Format(mstrETime, "YYYY-MM-DD")
    Else
        intDay = mintPreDays - DateDiff("D", CDate(strCurrDate), CDate(mstrETime))
        If intDay < 0 Then intDay = 0
        strMaxDate = Format(DateAdd("d", intDay, CDate(mstrETime)), "yyyy-MM-dd")
        If CDate(mstrETime) < CDate(Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")) Then
            mstrETime = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    mstrETime = Format(strMaxDate & " " & Format(mstrETime, "HH:mm:ss"), "yyyy-MM-DD HH:mm:ss")
    
    dtpDate.Value = Format(mstrTime, "YYYY-MM-DD")
    dtpDate.MaxDate = Format(strMaxDate, "YYYY-MM-DD")
    dtpDate.MinDate = Format(mstrBTime, "YYYY-MM-DD")
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientInfo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo Errhand
    '��ȡ������Ϣ
    mstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then
        mT_Patient.lng����ID = Val(zlCommFun.Nvl(rsTmp("��Ժ����ID").Value))
    End If
    
    '��ȡ����ȼ�
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then mT_Patient.lng����ȼ� = zlCommFun.Nvl(rsTmp("����ȼ�"), 3)

    OpenPatientInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'����:����ļ��Ƿ�鵵
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSQL = "Select 1 From ���˻����ļ� Where Id=[1] And �鵵�� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ļ��Ƿ�鵵", lngID)
    If blnMove = True Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    If rsTemp.RecordCount > 0 Then
        CheckFileBack = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'����:��ʼ��������
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim cboChild As CommandBarPopup
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    
     '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '����������
    Set mcbrToolBar = cbsMain.Add("��׼", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    '��λ������
    '------------------------------------------------------------------------------------------------------------------
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    
    '�����
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '����
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse 'ȡ��
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitTabOper() As String
    '-------------------------------------------------------
    '����:��ʼ����������¼����
    '-------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    On Error Resume Next
    
    With vsfOper
        .Rows = 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "ʱ��", 1000 + 1000 * mintBigSize / 3, 4, , 4
        .NewColumn "����", 2000 + 2000 * mintBigSize / 3, 4, "����|����|��������|����", 1
        .NewColumn "", 255, 4
        .ExtendLastCol = True
        .Body.RowHeightMin = 300 + 300 * mintBigSize / 3
        .FixedCols = 1
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.WordWrap = False
        .Body.AllowUserResizing = flexResizeNone

        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Function


Private Function zlRefreshData() As Boolean
    Dim strTime As String
    Dim rsTmp  As New ADODB.Recordset
    
    On Error GoTo Errhand
    '����ˢ������
    gstrFields = "���," & adDouble & ",18|��Ŀ���," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��¼����," & adDouble & ",1|����," & _
            adLongVarChar & ",100|��Ŀ����," & adLongVarChar & ",20|δ��˵��," & adLongVarChar & ",20|��¼���," & adDouble & ",1|������Դ," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
             "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1"
    Call Record_Init(mrsOper, gstrFields)
    gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"

    '��ȡ������Ϣ
    mstrSQL = "" & _
         " Select C.ID ���, B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.δ��˵��,C.��¼����,C.��¼���,C.��Ŀ����,C.������Դ,C.��ʾ,C.��ԴID,C.����" & _
         " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
         " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]  AND Nvl(A.Ӥ��, 0)=[4] AND a.����id=[2] AND a.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
         " AND c.��¼����=4  AND B.����ʱ�� BETWEEN [5]  And [6]"

    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
        mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
    End If

    strTime = CDate(Format(mstrTime, "YYYY-MM-DD") & " 23:59:59")
    If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")

    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������±����Ϣ", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, _
        mT_Patient.lngӤ��, Int(CDate(Format(mstrTime, "YYYY-MM-DD"))), CDate(strTime))
    With rsTmp
        Do While Not .EOF
            gstrValues = zlCommFun.Nvl(!���) & "|" & zlCommFun.Nvl(!��Ŀ���, 0) & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!��¼����) & "|" & _
                zlCommFun.Nvl(!��¼����) & "|" & zlCommFun.Nvl(!��Ŀ����) & "|" & Nvl(!δ��˵��) & "|" & zlCommFun.Nvl(!��¼���, 0) & "|" & Val(zlCommFun.Nvl(!������Դ, 0)) & "|" & _
                Val(zlCommFun.Nvl(!��ʾ, 0)) & "|" & Val(zlCommFun.Nvl(!��ԴID, 0)) & "|" & Val(zlCommFun.Nvl(!����, 0)) & "|0"
            Call Record_Add(mrsOper, gstrFields, gstrValues)
        .MoveNext
        Loop
    End With
    
    '���������Ϣ
    mrsOper.Filter = 0
    mrsOper.Sort = "ʱ��"
    With mrsOper
        vsfOper.Rows = vsfOper.FixedRows
        Do While Not .EOF
            vsfOper.Rows = vsfOper.Rows + 1
            vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperTime) = Format(!ʱ��, "HH:mm")
            vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperType) = Nvl(!��Ŀ����, "����")
            If InStr(1, ",0,3,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0 Then
                vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = 255
            Else
                vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = &H80000012
            End If
            vsfOper.RowData(vsfOper.Rows - 1) = Val(!���)
        .MoveNext
        Loop
        vsfOper.Rows = vsfOper.Rows + 1
    End With
        vsfOper.Row = 1
        vsfOper.Col = 1
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function UpData(ByVal intRow As Integer, ByVal intCOl As Integer, _
    Optional blnComList As Boolean = False) As Boolean
    Dim strName As String
    Dim strTime As String
    Dim strValue As String
    Dim lngNo As String
    Dim lngID As Long

    
    On Error GoTo Errhand
    lngNo = 4
    If blnComList = True Then
        strName = vsfOper.EditText
        strTime = Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm")
    Else
        strName = vsfOper.TextMatrix(intRow, Col_OperType)
        strTime = Format(vsfOper.EditText, "HH:mm")
        If Not IsDate(strTime) Then strTime = ""
    End If
    mrsOper.Filter = "��¼����=" & lngNo & " And ���=" & Val(vsfOper.RowData(intRow))
    If mrsOper.RecordCount <> 0 Then
        If Val(mrsOper!״̬) <> 1 And Val(mrsOper!״̬) <> 3 Then 'his��ȡ������
            mrsOper!״̬ = 2
            If Trim(strTime) = "" Or strName = "" Then
                mrsOper!��Ŀ���� = ""
                mrsOper!���� = ""
            ElseIf Trim(strTime) <> "" And strName <> "" Then
                mrsOper!��Ŀ���� = strName
                mrsOper!���� = strName
            End If
            If Trim(strTime) <> "" Then mrsOper!ʱ�� = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
        Else
            If Trim(strTime) = "" Or strName = "" Then
                mrsOper!״̬ = 3
                mrsOper!��Ŀ���� = ""
                mrsOper!���� = ""
            Else
                mrsOper!״̬ = 1
                mrsOper!��Ŀ���� = strName
                mrsOper!���� = strName
            End If
            If Trim(strTime) <> "" Then mrsOper!ʱ�� = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
        End If
        mrsOper.Update
    Else
        If Trim(strTime) = "" Or strName = "" Then
            strValue = ""
        Else
            strValue = 1
            strTime = SetDate(Format(Format(dtpDate.Value, "YYYY-MM-DD") & " " & strTime & ":00", "YYYY-MM-DD HH:mm:ss"))
        End If
        
        If strValue <> "" Then
            strValue = strName
            lngID = GetMaxID(mrsOper)
            gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
            gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNo & "|" & strValue & "|" & strName & "||0|0|0|0|0|1"
            vsfOper.RowData(intRow) = lngID
            Call Record_Add(mrsOper, gstrFields, gstrValues)
        End If
    End If
    
    If strName <> vsfOper.TextMatrix(intRow, Col_OperType) Or Format(strTime, "HH:mm") <> Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm") Then
        mblnChage = True
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SetDate(ByVal strTime As String) As String
'---------------------------------------------------------
' �������
'---------------------------------------------------------
    Dim strVTime As String
    If Not IsDate(strTime) Then Exit Function
    strVTime = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    If CDate(strTime) < CDate(mstrBTime) Then
        strVTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strVTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    SetDate = strVTime
End Function


Private Function GetMaxID(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'����:��ȡ��¼���е�������
'----------------------------------------------------
    rsTmp.Filter = 0
    rsTmp.Sort = "��� Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(rsTmp!���) + 1
    End If
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save '����
            If Not SaveData Then Exit Sub
            Call zlRefreshData
        Case conMenu_Edit_Reuse 'ȡ��
            Call zlRefreshData
            mblnChage = False
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    picOper.Height = 5000 + 5000 * mintBigSize / 3
    Bottom = stbThis.Height
    
    With picDate
        .Left = 0
        .Top = 0
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = (lblDate.Width + dtpDate.Width + 520) + (lblDate.Width + dtpDate.Width + 520) * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    With lblDate
        .Left = 30
        .Top = 60
        .Height = picDate.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
        .Top = 0
        .Left = lblDate.Left + lblDate.Width
    End With

    With imgbtn(1)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = lblDate.Width + dtpDate.Width + 20
    End With
    
    With imgbtn(0)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = lblDate.Width + dtpDate.Width + imgbtn(1).Width + 30
    End With
    
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With
    
    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("��")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With


End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    With picOper
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
             Control.Enabled = IIf(mblnChage = True, True, False)
    End Select
    
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    End If
    
End Sub

Private Sub dtpDate_Change()
    Dim strDate As String
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then Exit Sub
    imgbtn(1).Enabled = True
    imgbtn(0).Enabled = True
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
End Sub


Private Function dtpDateChageDate(ByVal strValue As String) As Boolean
'------------------------------------------------------------------------------
'��¼ʱ��Ϸ�ʱ�������仯��ˢ������
'------------------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String, strTime As String
    Dim i As Integer
    Dim strCurrDate As String
    Dim intBound As Integer
    Dim strBegin As String, strEnd As String
    Dim intCOl As Integer
    Dim strCurDate As String
    Dim intDay As Integer
    Dim strBTime As String
    On Error GoTo Errhand
    
    lblStb.Tag = lblStb.Caption
    
    If Format(strValue, "YYYY-MM-DD") > Format(mstrETime, "YYYY-MM-DD") Then
        If mbln��Ժ = False Then
            strErrMsg = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
        Else
            strErrMsg = "¼������ڲ��ܴ���[���˳�Ժʱ����ļ�����ʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "¼������ڲ���С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If Format(strValue, "YYYY-MM-DD") = mstrETime Then
        strDate = Format(Format(mstrETime, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    ElseIf Format(strValue, "YYYY-MM-DD") = mstrBTime Then
        strDate = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        strTime = strDate
    Else
        strDate = Format(Format(strValue, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(Format(strValue, "YYYY-MM-DD") & " 23:59:00", "YYYY-MM-DD HH:mm:ss")
    End If
    
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, strTime, strCurrDate) Then
        strErrMsg = "¼���ʱ��[" & strValue & "]����[�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        GoTo ErrInfo
    End If
    
    mstrTime = Format(dtpDate.Value, "YYYY-MM-DD hh:mm:ss")
    If mblnChage Then
        mblnChage = False
        If MsgBox("�����Ѿ������ı�,�����Ƿ���б���?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            If Not SaveData Then Exit Function
        End If
    End If
    Call zlRefreshData
    dtpDateChageDate = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub dtpDate_CloseUp()
    vsfOper.SetFocus
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChage = True Then
        If MsgBox("�������������Ѿ������ı�,�����Ƿ���Ҫ���棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If

    mblnChage = False
    mblnMove = False
    mbln��Ժ = False
    
    If Not (mrsOper Is Nothing) Then Set mrsOper = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgbtn_Click(Index As Integer)
    Select Case Index
        Case 1
            dtpDate.Value = dtpDate.Value - 1
            Call dtpDate_Change
        Case 0
            dtpDate.Value = dtpDate.Value + 1
            Call dtpDate_Change
    End Select
    vsfOper.SetFocus
End Sub

Private Sub picOper_Paint()
     picOper.BackColor = &H8000000F
End Sub

Private Sub picOper_Resize()
    On Error Resume Next
    With vsfOper
        .Left = 5
        .Top = picDate.Top + picDate.Height + 20
        .Width = picOper.Width
        .Height = picOper.Height - .Top
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
End Sub

Private Function SaveData() As Boolean
    '--------------------------------------------------------
    '����:���������޸ı���
    '--------------------------------------------------------
    Dim lngItemCode As Long
    Dim strTime As String
    Dim strEnd As String
    Dim strMarkTime As String
    Dim strSQL As String
    Dim strValue As String
    Dim int������ As Integer
    Dim int��Ŀ�״� As Integer
    Dim i As Integer
    Dim blnTran As Boolean
    Dim arrSQL() As String
    
    On Error GoTo Errhand
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    With mrsOper
        .Filter = 0
        .Sort = "ʱ��"
        '��ɾ�����޸ĵ�������Ϣ,һ��������ö���������������ʱ�����������ʱ����ͬ����������ʱ��Ļ����ᵼ����������ʱ�䷢���仯
        Do While Not .EOF
            If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then
                lngItemCode = 4
                If Val(!״̬) = 2 Then
                    strTime = Format(!ԭʼʱ��, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime
                    int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strMarkTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & lngItemCode & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & 0 & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "NULL" & ","
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & "NULL,"
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    strSQL = strSQL & "NULL" & ","
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & 0 & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
                
                strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                strValue = Trim(zlCommFun.Nvl(!����))
                If strValue <> "" Then
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strMarkTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & lngItemCode & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & 0 & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & "NULL,"
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & IIf(strValue = "����", "1", "NULL") & ","
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!δ��˵��) & "'", "NULL") & ","
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & int��Ŀ�״� & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnTran = True
    '��ִ�����ݱ仯
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������������"):
'        Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    mblnChage = False
    mblnOK = True
    
    SaveData = True
    Screen.MousePointer = 0
    
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call SaveErrLog
End Function


Private Function ISCheckDept(ByVal str����ʱ�� As String) As Boolean
'���ܣ��Ƿ���Zl_���µ�����_Update�н��п��Ҽ��
    'mstrOverDate<=mstrETime ���Ҳ����Ѿ���Ժ���϶��ǲ��˳�Ժʱ�����Ժʱ����һ�У��������Ľ����
    If mbln��Ժ = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
            ISCheckDept = False
        Else
            ISCheckDept = True
        End If
    Else
        ISCheckDept = True
    End If
End Function

Private Sub vsfOper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    lblStb.Caption = ""
    vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
End Sub

Private Sub vsfOper_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '����Ƿ���ͬ������������
    Dim lngID As Long, intState As Integer
    lngID = Val(vsfOper.RowData(Row))
    If lngID > 0 Then
        mrsOper.Filter = "��¼����=4 And ���=" & lngID
        intState = mrsOper!״̬
        If InStr(1, ",0,3,9,", "," & Val(Nvl(mrsOper!������Դ, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "ͬ������������,�������������ɾ��."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
        
        '������ݵ�ɾ������
        If intState = 0 Or intState = 2 Then '��ʾ��ԭ������
            mrsOper!���� = ""
            mrsOper!��Ŀ���� = ""
            mrsOper!״̬ = 2
        Else '��ʾ��������
            mrsOper.Delete
        End If
        mrsOper.Update
        mblnChage = True
    End If
End Sub

Private Sub vsfOper_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intRow As Integer
    '�����һ��û��¼��ʱ���������Ϣ ���ܽ�����һ��
    If Row >= vsfOper.FixedRows And Col >= vsfOper.FixedCols Then
        If vsfOper.TextMatrix(Row, Col_OperTime) = "" Or (vsfOper.TextMatrix(Row, Col_OperType) = "" And vsfOper.EditText = "") Then Cancel = True
    End If
End Sub

Private Sub vsfOper_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfOper
        If .EditMode(NewCol) = 1 Then
            .Body.FocusRect = flexFocusSolid
        Else
            .Body.FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfOper_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    If Trim(vsfOper.TextMatrix(Row, Col_OperTime)) <> "" Then
        Call UpData(Row, Col, True)
    End If
End Sub

Private Sub vsfOper_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnFileBack = True Then
        Cancel = True
        vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    '����Ƿ���ͬ������������
    If Val(vsfOper.RowData(Row)) > 0 Then
        mrsOper.Filter = "��¼����=4 And ���=" & Val(vsfOper.RowData(Row))
        If InStr(1, ",0,3,9,", "," & Val(Nvl(mrsOper!������Դ, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "ͬ������������,��������������޸�."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    End If
End Sub

Private Sub vsfOper_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�������ݺϷ��Լ��
    Dim strText As String
    Dim strInfo As String, strDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If Row < vsfOper.FixedRows Then Exit Sub
    If vsfOper.EditText = vsfOper.TextMatrix(Row, Col) Then Exit Sub
    With vsfOper
        strText = .EditText
        If Col = Col_OperTime Then
            If Trim(strText) = "" Then
                .TextMatrix(Row, Col_OperType) = ""
                GoTo ErrEnd
            End If
            Select Case Len(strText)
            Case 3, 4
                strText = String(4 - Len(strText), "0") & strText
                strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
            Case Is < 3
                strText = String(2 - Len(strText), "0") & strText
                strText = Format(Now, "HH") & ":" & strText
            End Select
            
            '�Ϸ��Լ��
            If Mid(strText, 3, 1) <> ":" Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
                GoTo ErrInfo
            End If
            If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
                GoTo ErrInfo
            End If
            If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
                GoTo ErrInfo
            End If
            .EditText = Format(strText, "HH:mm")
            
            '���¼���ʱ���Ƿ��Ѿ�������������Ϣ
            strDate = Format(dtpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")
            gstrSQL = "select 1 from ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & _
                " Where A.ID=B.�ļ�ID And B.ID=C.��¼ID And A.ID=[1] And B.����ʱ��=[2] And C.��¼����=4"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��������", mT_Patient.lng�ļ�ID, CDate(strDate))
            If rsTemp.RecordCount > 0 Then
                strInfo = "��ʱ���Ѿ�����������Ϣ�����飡 ʱ��[" & strDate & "]"
                GoTo ErrInfo
            End If
            If Not CheckDateTime(Row, "ʱ��", Format(dtpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")) Then
                Cancel = True
            End If
ErrEnd:
            If Cancel = False Then Call UpData(Row, Col, IIf(Col = Col_OperType, True, False))
        End If
    End With
    
    Exit Sub
ErrInfo:
    lblStb.Caption = strInfo
    lblStb.ForeColor = 255
    Cancel = True
End Sub


Private Function CheckDateTime(ByVal lngRow As Long, ByVal strName As String, ByVal strTime As String) As Boolean
'------------------------------------------------------------------
'����:��¼����ʱ����������÷�Χ
'------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String
    Dim strCurrDate As String
    Dim strInfo As String
    
    On Error GoTo Errhand
    If lngRow <> 0 Then
        strInfo = "��" & lngRow & "��"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstrETime, "YYYY-MM-DD HH:mm") Then
        If mbln��Ժ = False Then
            strErrMsg = strInfo & "��¼����ʱ���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ! "
        Else
            strErrMsg = strInfo & "��¼����ʱ�䲻�ܴ���[���˳�Ժʱ����ļ�����ʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstrBTime, "YYYY-MM-DD HH:mm") Then
        strErrMsg = strInfo & "��¼����ʱ�䲻��С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, strTime, strCurrDate) Then
        strErrMsg = strInfo & "��¼����ʱ��[" & strTime & "]����![�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        GoTo ErrInfo
    End If
    
    CheckDateTime = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsAllowInput(ByVal lng����ID As Long, ByVal lng��ҳID As Long, lngӤ�� As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    'ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    On Error GoTo Errhand
    
    IsAllowInput = True
    If lngӤ�� <> 0 And mbln��Ժ = True Then
        strBabyOutTime = GetAdviceOutTime(lng����ID, lng��ҳID, lngӤ��)
        If strBabyOutTime <> "" Then
            strTime = Format(DateAdd("H", mlngHours, strBabyOutTime), "yyyy-MM-dd HH:mm")
            GoTo GONext
        End If
    End If
    gstrSQL = "" & _
              " SELECT DECODE(��ֹԭ��,1,'��Ժ',3,'ת��',10,'Ԥ��Ժ',15,'ת����',DECODE(��ʼԭ��,10,'��Ժ','δ����')) AS ����,��ֹʱ�� AS ʱ��" & _
              " From ���˱䶯��¼" & _
              " WHERE (��ֹԭ�� IN (1,3,10,15) OR ��ʼԭ��=10) And ����ID=[1] And ��ҳID=[2] And [3] <= ��ֹʱ��" & _
              " ORDER BY ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��", lng����ID, lng��ҳID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    'ֻȡ��һ�����ϵļ�¼
    strTime = Format(DateAdd("H", mlngHours, rsTemp!ʱ��), "yyyy-MM-dd HH:mm")
GONext:
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
