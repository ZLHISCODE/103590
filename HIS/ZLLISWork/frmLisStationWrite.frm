VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLisStationWrite 
   BorderStyle     =   0  'None
   Caption         =   "��ͨ������д"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLisStationWrite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog CommDialog 
      Left            =   6000
      Top             =   4500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraTitle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8595
      Begin VB.CheckBox chkYiQiTiShi 
         Appearance      =   0  'Flat
         Caption         =   "���������ʾ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5580
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox chkYiQiBiaoShi 
         Appearance      =   0  'Flat
         Caption         =   "������ʶ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6420
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.PictureBox PicFilter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5160
         MouseIcon       =   "frmLisStationWrite.frx":0E42
         Picture         =   "frmLisStationWrite.frx":0F94
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   45
         Width           =   240
      End
      Begin VB.CheckBox chkOriginal 
         Appearance      =   0  'Flat
         Caption         =   "ԭʼ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   780
         TabIndex        =   15
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkLast 
         Appearance      =   0  'Flat
         Caption         =   "�ϴ�"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1470
         TabIndex        =   14
         Top             =   30
         Width           =   690
      End
      Begin VB.CheckBox chkSign 
         Appearance      =   0  'Flat
         Caption         =   "��־"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2190
         TabIndex        =   13
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkUnit 
         Appearance      =   0  'Flat
         Caption         =   "��λ"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2910
         TabIndex        =   12
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkReferrence 
         Appearance      =   0  'Flat
         Caption         =   "�ο�"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3630
         TabIndex        =   11
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkMB 
         Appearance      =   0  'Flat
         Caption         =   "ø��"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4350
         TabIndex        =   10
         Top             =   30
         Width           =   660
      End
      Begin VB.CheckBox chkChina 
         Appearance      =   0  'Flat
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   60
         TabIndex        =   9
         Top             =   30
         Width           =   690
      End
      Begin VB.Label lblLow 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Height          =   210
         Left            =   5535
         TabIndex        =   21
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ƫ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   5895
         TabIndex        =   20
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblHigh 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Height          =   210
         Left            =   6315
         TabIndex        =   19
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ƫ��"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   6645
         TabIndex        =   18
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblExigency 
         BackColor       =   &H000040C0&
         Height          =   210
         Left            =   7095
         TabIndex        =   17
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "��ʾ"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   7425
         TabIndex        =   16
         Top             =   60
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2685
      Left            =   5490
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���ѡ��"
         Object.Width           =   2999
      EndProperty
   End
   Begin zl9LisWork.VsfGrid vsf 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   5027
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   4845
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
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
   Begin VB.Frame fraComment 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   3420
      Width           =   7050
      Begin VB.TextBox txtDiagnose 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   3960
         Locked          =   -1  'True
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   60
         Width           =   3000
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   450
         Locked          =   -1  'True
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   60
         Width           =   3000
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3540
         TabIndex        =   7
         Top             =   90
         Width           =   405
      End
      Begin VB.Label lblComment 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���鱸ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   30
         TabIndex        =   6
         Top             =   90
         Width           =   405
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   7020
      Top             =   4440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLisStationWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mlngKey As Long    '�걾ID
Private mDeviceID As Long
Private mstrType As String '��������
Private mblnEdit As Boolean '�Ƿ�����༭
Private mbytRedoNumber As Long '��������
Private mblnLoadHistory As Boolean '�Ƿ�װ����ʷ����
Private mSelectRedo As Boolean '�Ƿ�ѡ��������
Private mblnChangeEdit As Boolean, mblnEvent As Boolean
Private mLngPatientID As Long                               '����ID
Private mstrPatientName As String                           '��������
Private lngReferenceLow As Long                             '�ο�����ɫ
Private lngReferenceHigh As Long                            '�ο�����ɫ
Private lngReferenceExigency As Long                        '�ο���ʾ��ɫ
Public mblnPatientFind As Boolean                           '�Ƿ񰴲������鿴
Const mintColCount As Integer = 29                          '��ʾ�б��з�����ʾһ�����ж��ٸ�COL

Private Enum mCol
    ������Ŀ = 1
    ԭʼ���
    ������
    ��λ
    CV
    �����־
    �ϴν��
    �ϴ�ʱ��
    ����ο�
    �������
    ����id
    ���㹫ʽ
    �����Χ
    �̶���Ŀ
    С��
    ��������
    ��������
    ������Ŀid
    �������
    �걾ID
    od
    CUTOFF
    COV
    ø���ID
    ���챨��
    ���쾯ʾ
    ������ʾ
    ������˱�ʶ
End Enum

Public Event StartEdit(Cancel As Boolean)

Private Function CalcDefaultFlag(ByVal strValue As String, ByVal strReference As String, Optional ByVal bytMode As Byte = 1, _
    Optional ByVal strAlarmLow As String, Optional ByVal strAlarmHigh As String, Optional ByVal lngItemID As Long) As String
    
    '--------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    
    If Len(Trim(strValue)) = 0 Then CalcDefaultFlag = "": Exit Function
    
    CalcDefaultFlag = ""
    
    If InStr(strReference, vbCrLf) > 0 Then strReference = Mid(strReference, 1, InStr(strReference, vbCrLf) - 1)
    If Trim(strReference) = "" Then Exit Function
                
    If bytMode = 2 Or (bytMode = 3 And IsNumeric(strValue) = False) Then  '���ԡ��붨��
        If bytMode = 2 Or InStr(strReference, "��") = 0 Or Trim(strValue) Like "*��*" Or Trim(strValue) Like "*+*" Or _
            Trim(strValue) Like "*��*" Or Trim(strValue) Like "*��*" Or Trim(strValue) Like "*-*" Then
            '���Ի��޷�Χ�ο��İ붨��
            If (Len(Trim(strReference)) > 0 And (Trim(strReference) Like (Trim(strValue) & "*") Or Trim(strReference) Like ("*" & Trim(strValue)))) Or _
                (Not (Trim(strValue) Like "*��*" Or Trim(strValue) Like "*+*" Or Trim(strValue) Like "*��*")) Then
                CalcDefaultFlag = ""
            Else
                CalcDefaultFlag = "�쳣"
            End If
            Exit Function
        Else
            '��ȡ�붨��ֵ
            For i = 1 To Len(Trim(strValue))
                If InStr("01234567890.", Mid(strValue, i, 1)) > 0 Then Exit For
            Next
            If i > Len(Trim(strValue)) Then Exit Function
            strValue = Val(Mid(strValue, i))
        End If
    End If
    
'    If InStr(strValue, ">") Then CalcDefaultFlag = "��": Exit Function
'    If InStr(strValue, "<") Then CalcDefaultFlag = "��": Exit Function
    strValue = Replace(strValue, "<", "")
    strValue = Replace(strValue, ">", "")
    
    '����������־Ͳ�������ֱ���˳�
    If IsNumeric(strValue) = False Then Exit Function
    
    
    
    If InStr(strReference, "��") > 0 Then
        
        '���С�ڲο���ֵ
        If Val(strValue) < Val(Mid(strReference, 1, InStr(strReference, "��") - 1)) And _
            Len(Trim(Mid(strReference, 1, InStr(strReference, "��") - 1))) > 0 Then
            CalcDefaultFlag = "��"
        End If
        
        '������ڲο���ֵ
        If Val(strValue) > Val(Mid(strReference, InStr(strReference, "��") + 1)) And _
            Len(Trim(Mid(strReference, InStr(strReference, "��") + 1))) > 0 Then
            CalcDefaultFlag = "��"
        End If
        
        If CalcDefaultFlag <> "" Then
            '�ߵ��ж�
            If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
                If Val(strValue) < Val(strAlarmLow) Then
                    CalcDefaultFlag = "����"
                    Exit Function
                End If
            End If
            If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
                If Val(strValue) > Val(strAlarmHigh) Then
                    CalcDefaultFlag = "����"
                    Exit Function
                End If
            End If
        End If
    Else
        '�ߵ��ж�
        If Len(Trim(strAlarmLow)) > 0 And Val(strAlarmLow) <> 0 Then
            If Val(strValue) < Val(strAlarmLow) Then
                CalcDefaultFlag = "����"
                Exit Function
            End If
        End If
        If Len(Trim(strAlarmHigh)) > 0 And Val(strAlarmHigh) <> 0 Then
            If Val(strValue) > Val(strAlarmHigh) Then
                CalcDefaultFlag = "����"
                Exit Function
            End If
        End If
    End If
    
    gstrSql = "select nvl(��ο�,0) as ��ο� from ������Ŀ where ������Ŀid = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    If rsTmp.EOF = False Then
        If rsTmp("��ο�") = 1 Then
            CalcDefaultFlag = ""
        End If
    End If

End Function

Private Function CalcExpress(ByVal Vsf As Object, ByVal strExPress As String) As String
    
    '--------------------------------------------------------------------------------------------------------
    '����:�ڱ���м���ĳһ���ʽ�Ľ��
    '����:vsf           ������ݵı��
    '     strExpress    Ҫ����ı��ʽ
    '����:������ֵ
    '--------------------------------------------------------------------------------------------------------
    
    Dim strTmpPress As String
    Dim rs As New ADODB.Recordset
    
    Dim lngTmpID As Long
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim lngLoop As Long
    Dim sglValue As String
    Dim intCol As Integer, intCols As Integer
    
    On Error GoTo errH
    
    CalcExpress = 0
    
    strTmpPress = strExPress
    If strTmpPress <> "" Then
        
        intCols = GetColCount(Vsf.Cols)
        If intCols = 0 Then intCols = 1
        
        lngLeftPos = InStr(strTmpPress, "[")
        lngRightPos = InStr(strTmpPress, "]")
        
        Do While lngLeftPos > 0
        
            lngTmpID = Val(Mid(strTmpPress, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
            
            '�ж�lngTmpID�Ƿ�Ҳ�Ǽ�����Ŀ
            For intCol = 0 To intCols - 1
                For lngLoop = 1 To Vsf.Rows - 1
'                    If Val(Vsf.RowData(lngLoop)) = lngTmpID Then
                    If Val(Me.Vsf.Cell(flexcpData, lngLoop, intCol * mintColCount, lngLoop, intCol * mintColCount)) = lngTmpID Then
                        If Trim(Vsf.TextMatrix(lngLoop, mCol.���㹫ʽ + intCol * mintColCount)) <> "" Then
                            '�Ǽ�����Ŀ,�ȼ�����˽��
                            sglValue = CalcExpress(Vsf, Trim(Vsf.TextMatrix(lngLoop, mCol.���㹫ʽ + intCol * mintColCount)))
                        Else
                            '���Ǽ�����Ŀ,ֱ��ȡ�˽��
                            sglValue = Vsf.TextMatrix(lngLoop, mCol.������ + intCol * mintColCount)
                            If sglValue = "" Then
                                CalcExpress = ""
                                Exit Function
                            Else
                                sglValue = Val(sglValue)
                            End If
                        End If
                        
                        Exit For
                        
                    End If
                Next
                If Val(sglValue) <> 0 Then Exit For
            Next
            
            '�ڵ�ǰ�����û�д˼�����Ŀ,��Ϊ���Ϊ��
            If lngLoop = Vsf.Rows Then sglValue = 0
                                        
            '�Խ��������ʽ�еļ�������
            strTmpPress = Mid(strTmpPress, 1, lngLeftPos - 1) & sglValue & Mid(strTmpPress, lngRightPos + 1)
            
            '����һ���������ӵ�λ��
            lngLeftPos = InStr(strTmpPress, "[")
            lngRightPos = InStr(strTmpPress, "]")
            sglValue = ""
        Loop
                
        '������ʽ�Ľ��
        On Error Resume Next
        Set rs = zlDatabase.OpenSQLRecord("SELECT " & strTmpPress & " AS ��� FROM DUAL", Me.Caption)
        If rs.BOF = False Then CalcExpress = zlCommFun.Nvl(rs("���"), 0)
        On Error GoTo 0
        
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowValue(ByVal intType As Integer, Optional intAttr As Integer = 2, Optional lngItemID As Long = 0)
    'intType��1���г����ֵ��2���г���עֵ
    'intAttr��2�����֡�3���붨��
    Dim rs As New ADODB.Recordset
    Dim strsql As String, strValue As String, i As Long, aValues() As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    
    Select Case intType
        Case 1
            strsql = "SELECT ROWNUM AS ID,����,���� As ȡֵ FROM ���������� A " & _
                " WHERE ����=[1]"
            intColCount = GetColCount(Vsf.Col)
            Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, Vsf.TextMatrix(Vsf.Row, mCol.�����Χ + intColCount * mintColCount))
            With lvwSelect
                .ListItems.Clear
                .Tag = 1
                
                Do While Not rs.EOF
                    .ListItems.Add , "_" & rs("ID"), Nvl(rs("ȡֵ"))
                
                    rs.MoveNext
                Loop
            End With
        
            If intAttr <> 1 Then '�Ƕ�����ȡֵ����
                strsql = "SELECT ȡֵ���� FROM ������Ŀ WHERE ������ĿID=[1]"
                Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lngItemID)
                If rs.EOF Then
                    strValue = "-|��|+|++|+++|++++"
                Else
                    strValue = Nvl(rs("ȡֵ����"), "-|��|+|++|+++|++++")
                    strValue = Replace(strValue, ";", "|")
                End If
                aValues = Split(strValue, "|")
                With lvwSelect
                    For i = 0 To UBound(aValues)
                        .ListItems.Add , "V" & i, aValues(i)
                    Next
                End With
            End If
        Case 2
            strsql = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ȡֵ FROM ���鱸ע���� A " & _
                "WHERE A.���� Is Null Or A.����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mstrType)
            With lvwSelect
                .ListItems.Clear
                .Tag = 2
                
                Do While Not rs.EOF
                    .ListItems.Add , "_" & rs("ID"), Nvl(rs("ȡֵ"))
                
                    rs.MoveNext
                Loop
            End With
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadPatient() As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim lngPatientID As Long
    Dim mbytMode As Integer '0�������걾
    Dim mlngLoop As Long
    Dim strTmp As String
    Dim lngAdvice As Long     'ҽ��ID
    Dim intColCount As Integer, intCol As Integer
    Dim blnMoved As Boolean                                         '�Ƿ��Ƴ�
    Dim strSQLbak As String
    Dim strStart As String
    Dim strEnd As String
    
    On Error GoTo ErrHand
'    If mblnLoadHistory Then ReadPatient = True: Exit Function
    mblnLoadHistory = True
    
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpForeColor, 1, 0, 1, 1) = vbBack
    
    strStart = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";", ";")(0), 2)
    
    If strStart = "�Զ���" Then
        strStart = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Now & ";" & Now, ";")(1)
        strStart = Format(strStart, "yyyy-mm-dd 00:00:00")
        strEnd = Split(zlDatabase.GetPara("�걾��Χ", 100, 1208, "��  ��") & ";" & Now & ";" & Now, ";")(2)
        strEnd = Format(strEnd, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    End If

    mstrSQL = "Select /*+ rule */ Distinct A.�걾ID ,a.������Ŀid ,A.����, a.�������, a.�̶���Ŀ, a.Id, a.������Ŀ, a.ԭʼ���, a.�ϴν��, a.�ϴ�ʱ��, a.Cv," & vbNewLine & _
                "            Decode(a.���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', a.���ν��) As ���ν��, Rownum As ���, a.���㹫ʽ," & vbNewLine & _
                "            a.�������, a.��־, a.����id, a.�걾���, a.����ʱ��, a.�걾���, a.�걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��," & vbNewLine & _
                "            a.��ǰ����, a.��ҳid, a.�����Χ, Nvl(G.С��λ��,2) as С��, a.��������, a.��������, a.��λ,a.����ο� as �ο�, " & vbNewLine & _
                "                           Trim(Replace(Replace(' ' || Zlgetreference(a.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
                "                                                                                                                   a.����id, a.����,a.�������id), ' .', '0.'), '��.', '��0.')) As �ο�1," & vbNewLine & _
                "            a.OD,a.CUTOFF,a.COV,a.ø���ID,a.���챨��,a.���쾯ʾ,lpad(����,10,'0') as ����,a.������,a.�걾���� " & vbNewLine & _
                "From (Select A.id as �걾ID ,b.������Ŀid, decode(d.�������,Null,nvl(h.����,C.����),d.�������) as ����, Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ," & vbNewLine & _
                "                           b.������Ŀid As Id, " & vbNewLine & _
                "                           " & IIf(chkChina.Value = 1, " c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ ", "d.��д as ������Ŀ ") & vbNewLine & _
                "                           , b.ԭʼ���," & vbNewLine & _
                "                           '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv, b.������ As ���ν��, d.���㹫ʽ, d.�������," & vbNewLine & _
                "                           Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                "                           Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
                "                           Decode(a.����id, Null," & vbNewLine & _
                "                                           To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                "                                           a.�걾���) As �걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�걾����,a.��������,a.�����, a.סԺ��," & vbNewLine & _
                "                           a.���� As ��ǰ����, a.��ҳid, d.�����Χ, d.��������, d.��������, d.��λ,b.OD,B.CUTOFF,B.SCO as COV,b.ø���ID, " & vbNewLine & _
                "                           d.���챨���� as  ���챨��,d.���쾯ʾ�� as ���쾯ʾ,b.����ο�,a.������,a.�������ID " & vbNewLine & _
                "            From ����걾��¼ a, ������ͨ��� b, ����������Ŀ c, ������Ŀ d, ������ĿĿ¼ h" & vbNewLine & _
                "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And" & vbNewLine & _
                "                        b.������Ŀid = h.Id(+) And b.��¼���� = 0 And a.����ID = [1] and a.����ʱ�� between [2] and [3] " & vbNewLine & _
                "            ) A ,����������Ŀ G" & _
                "  Where A.����id = G.����id(+) And A.ID = G.��Ŀid(+) "


    
    If blnMoved Then
        strSQLbak = mstrSQL
        strSQLbak = Replace(strSQLbak, "����걾��¼", "H����걾��¼")
        strSQLbak = Replace(strSQLbak, "������ͨ���", "H������ͨ���")
        strSQLbak = Replace(strSQLbak, "����������Ŀ", "H����������Ŀ")
        mstrSQL = mstrSQL & " Union ALL " & strSQLbak
    End If
    
    mstrSQL = mstrSQL & " Order by ����,�������,����ʱ�� desc "
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mLngPatientID, CDate(strStart), CDate(strEnd))
    
    If rs.BOF = False Then
        '��ʼ�걾��Ϣ
        mDeviceID = rs("����ID")
        Me.txtComment.Text = Nvl(rs("���鱸ע"))
        
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Call ReadVsf_Patient(rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
        rs.MoveFirst
        
        Call FormatVsfCell(Vsf, mCol.������, "0.0######", IIf(Nvl(rs("�������"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.С��, -1))
                
        Call FormatVsfCell(Vsf, mCol.ԭʼ���, "0.0######", IIf(Nvl(rs("�������"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.С��, -1))
        
'        If chkLast.Value Then LoadLastValue
        '--ÿ�ζ�������ʷ���
        LoadLastValue
    Else
        mDeviceID = -1
        Me.txtComment.Text = ""
        ResetVsf Vsf
    End If
    
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Call ApplyResultColor(Vsf, mlngLoop, mCol.������ + intCol * mintColCount, _
                Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 6, "����", 5, 1))
        Next
    Next
    
    'д�������Ϣ
    Me.txtDiagnose.Text = ""
    gstrSql = "Select b.ҽ��id, b.��Ŀ, b.����, b.����" & vbNewLine & _
                "From ����걾��¼ a, ����ҽ������ b" & vbNewLine & _
                "Where a.ҽ��id = b.ҽ��id and a.ID = [1] " & vbNewLine & _
                "Order By ҽ��id, ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("��Ŀ")) & ":" & Replace(Nvl(rs("����")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txtDiagnose.Text = strTmp
    ReadPatient = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadVsf_Patient(ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    Dim intColCount  As Integer
    Dim intCol As Integer, intRow As Integer
    Dim lngHeight As Long
    Dim blnShowType As Boolean
    Dim lngItem As Long                                     '������ĿID
    Dim intItemCount As Integer                             '��ǰ�м�����Ŀ
    Dim rsTmp As New ADODB.Recordset
    blnShowType = zlDatabase.GetPara("����Ӧ��ʾ���", 100, 1208, False)
    If fraComment.Tag <> "" Then blnShowType = True
    
    If blnClear Then
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        For lngLoop = 0 To Vsf.Cols - 1
            Vsf.TextMatrix(1, lngLoop) = ""
            Vsf.Cell(flexcpData, 1, lngLoop, 1, lngLoop) = ""
        Next
        lngRow = 0
        Vsf.Cols = mintColCount
    Else
        'Ԥ����һ����
        With Vsf
            intColCount = GetColCount(.Cols)
            If intColCount = 0 Then intColCount = 1
            For intCol = 0 To intColCount - 1
                For intRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) = 0 Then
                        lngRow = intRow - 1
                        intColCount = intCol
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    
    With Vsf.Body
        If .ClientHeight < .CellHeight * 15 Then
            lngHeight = .CellHeight * 15
        Else
            lngHeight = .ClientHeight
        End If
    End With
    
    Do While Not rsData.EOF
        lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsData("ID"))))
'        lngCurrRow = -1
        If lngCurrRow = -1 Then
            '--------------------------------------������Ŀ��һ��ʱ������������Ŀ------------------------------------------
            If lngItem <> rsData("������ĿID") Then
                intItemCount = intItemCount + 1
                gstrSql = "select ���� from ������ĿĿ¼ where id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsData("������Ŀid")))
                
                With Vsf.Body

                    If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                        intColCount = intColCount + 1
                        lngRow = 1
                        With Vsf
                            .NewColumn "#", 300, 7
                            .NewColumn "������Ŀ", 2100, 1
                            .NewColumn "ԭʼ���", 0, 1
                            .NewColumn "���ν��", 1200, 1, , 1
                            .NewColumn "��λ", 1000, 1
                            .NewColumn "CV", 0, 1
                            .NewColumn "��־", 450, 1
                            .NewColumn "�ϴν��", 0, 1
                            .NewColumn "�ϴ�ʱ��", 0, 1
                            .NewColumn "�ο�", 1300, 1
                            .NewColumn "�������", 0, 1
                            .NewColumn "����id", 0, 1
                            .NewColumn "���㹫ʽ", 0, 1
                            .NewColumn "�����Χ", 0, 1
                            .NewColumn "�̶���Ŀ", 0, 1
                            .NewColumn "С��", 0, 1
                            .NewColumn "��������", 0, 1
                            .NewColumn "��������", 0, 1
                            .NewColumn "������ĿID", 0, 1
                            .NewColumn "�������", 0, 1
                            .NewColumn "�걾ID", 0, 1
                            .NewColumn "OD", 700, 1, , 1
                            .NewColumn "CUTOFF", 700, 1
                            .NewColumn "COV", 700, 1
                            .NewColumn "ø���ID", 0, 1
                            .NewColumn "���챨��", 0, 1
                            .NewColumn "���쾯ʾ", 0, 1
                            .NewColumn "������ʾ", 1000, 1
                            .NewColumn "������˱�ʶ", 1200, 1
                        End With
                    Else
                        lngRow = lngRow + 1
                    End If
                End With
                
                lngCurrRow = lngRow
                
            
            
                If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
                
                On Error Resume Next
                
                On Error GoTo ErrHand
                
                For lngLoop = 1 To mintColCount - 1
                    '���úϲ�
                    Me.Vsf.Body.MergeCol(lngLoop) = True
                    Me.Vsf.Body.MergeRow(lngCurrRow) = True
                    Me.Vsf.Body.MergeCells = flexMergeFree
                    intCol = intColCount * mintColCount + lngLoop
                    Vsf.TextMatrix(lngCurrRow, intCol) = rsTmp("����") & "(" & rsData("������") & " " & Format(rsData("����ʱ��"), "yyyy-mm-dd") & ")"
                Next
                '������ɫ
                Vsf.Cell(flexcpBackColor, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount + mintColCount - 1) = &HFDD6C6
                Me.Vsf.Cell(flexcpFontBold, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount + mintColCount - 1) = True
                            
            End If
            '-----------------------------------------------------------------------------------------------------
            With Vsf.Body

                If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                    intColCount = intColCount + 1
                    lngRow = 1
                    With Vsf
                        .NewColumn "#", 300, 7
                        .NewColumn "������Ŀ", 2100, 1
                        .NewColumn "ԭʼ���", 0, 1
                        .NewColumn "���ν��", 1200, 1, , 1
                        .NewColumn "��λ", 1000, 1
                        .NewColumn "CV", 0, 1
                        .NewColumn "��־", 450, 1
                        .NewColumn "�ϴν��", 0, 1
                        .NewColumn "�ϴ�ʱ��", 0, 1
                        .NewColumn "�ο�", 1300, 1
                        .NewColumn "�������", 0, 1
                        .NewColumn "����id", 0, 1
                        .NewColumn "���㹫ʽ", 0, 1
                        .NewColumn "�����Χ", 0, 1
                        .NewColumn "�̶���Ŀ", 0, 1
                        .NewColumn "С��", 0, 1
                        .NewColumn "��������", 0, 1
                        .NewColumn "��������", 0, 1
                        .NewColumn "������ĿID", 0, 1
                        .NewColumn "�������", 0, 1
                        .NewColumn "�걾ID", 0, 1
                        .NewColumn "OD", 700, 1, , 1
                        .NewColumn "CUTOFF", 700, 1
                        .NewColumn "COV", 700, 1
                        .NewColumn "ø���ID", 0, 1
                        .NewColumn "���챨��", 0, 1
                        .NewColumn "���쾯ʾ", 0, 1
                        .NewColumn "������ʾ", 1000, 1
                        .NewColumn "������˱�ʶ", 1200, 1
                    End With
                Else
                    lngRow = lngRow + 1
                End If
            End With
            
            lngCurrRow = lngRow
        
            If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
            
            On Error Resume Next
'            Vsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            Vsf.Cell(flexcpData, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount) = CStr(Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            For lngLoop = 0 To mintColCount - 1
                intCol = intColCount * mintColCount + lngLoop
                
                If Trim(Vsf.TextMatrix(0, intCol)) <> "" Then
                    If Vsf.TextMatrix(0, intCol) = "#" Then
                        Vsf.TextMatrix(lngCurrRow, intCol) = IIf(intColCount > 0, intColCount * (Vsf.Body.Rows - 1) + lngCurrRow, lngCurrRow) - intItemCount
                        Vsf.Cell(flexcpBackColor, lngCurrRow, intCol, lngCurrRow, intCol) = &HFDD6C6
                    Else
                        On Error Resume Next
                        strMask = ""
                        strMask = MaskArray(intCol)
                                                
                        On Error GoTo ErrHand

                         
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol)))
                        End If
                        Vsf.TextMatrix(lngCurrRow, intCol) = strNewValue
                    End If
                End If
                
            Next
        End If
        lngItem = Val(Nvl(rsData("������ĿID"), 0))
        rsData.MoveNext
    Loop
'    Call chkOriginal_Click: Call chkLast_Click: Call chkSign_Click
'    Call chkUnit_Click: Call chkReferrence_Click: Call chkMB_Click
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.������Ŀ + intCol * mintColCount) = IIf(chkChina.Value, 2100, 1000)
        Vsf.Body.ColWidth(mCol.ԭʼ��� + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.�ϴν�� + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.�ϴ�ʱ�� + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.�����־ + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
        Vsf.Body.ColWidth(mCol.��λ + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.����ο� + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.������ʾ + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.������˱�ʶ + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ReadData() As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSQL As String
    Dim lngPatientID As Long
    Dim mbytMode As Integer '0�������걾
    Dim mlngLoop As Long
    Dim strTmp As String
    Dim lngAdvice As Long     'ҽ��ID
    Dim intColCount As Integer, intCol As Integer
    Dim blnMoved As Boolean                                         '�Ƿ��Ƴ�
    Dim strSQLbak As String
    
    On Error GoTo ErrHand
    If mblnLoadHistory Then ReadData = True: Exit Function
    mblnLoadHistory = True
    
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpForeColor, 1, 0, 1, 1) = vbBack

     '1-������2-ƫ�͡�3-ƫ�ߡ�4-����(�쳣)��5-�������ޡ�6-��������
    mstrSQL = "Select a.ҽ��Id,a.������, a.����id, a.��ҳid, a.��������, a.������, a.����ʱ��, a.�����, a.���ʱ��,����ID,����,a.������,a.����ʱ��  " & vbNewLine & _
            "From ����걾��¼ a" & vbNewLine & _
            "Where a.Id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
    If Not rs.EOF Then
        lngPatientID = Nvl(rs("����ID"), 0)
        mstrType = Nvl(rs("��������"))
        '�Ƿ�������������ʾ������Ŀ
        lngAdvice = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0)
        If lngAdvice = 0 Then
            lngAdvice = 0
        Else
            lngAdvice = Nvl(rs("ҽ��ID"), 0)
        End If
        mbytRedoNumber = IIf(mSelectRedo, mbytRedoNumber - 1, Nvl(rs("������"), 0))
        
        mSelectRedo = False
        
        mbytMode = IIf(IsNull(rs("����ID")), 0, 1)
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�" & Nvl(rs("������"))
            .Panels(2).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
            If Nvl(rs("�����")) <> "" Then
                .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
            Else
                If Nvl(rs("������")) <> "" Then
                    .Panels(3).Text = "�����ˣ�" & Nvl(rs("������"))
                    .Panels(4).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                Else
                    .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                    .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                End If
            End If
        End With
        blnMoved = MovedByDate(CDate(Format(Nvl(rs("����ʱ��")), "yyyy-MM-dd hh:mm:ss")))
        mLngPatientID = Nvl(rs("����ID"), 0)
        mstrPatientName = Nvl(rs("����"))
    Else
        mstrSQL = "Select a.ҽ��Id,a.������, a.����id, a.��ҳid, a.��������, a.������, a.����ʱ��, a.�����, a.���ʱ��,����ID,����,a.������,a.����ʱ�� " & vbNewLine & _
            "From h����걾��¼ a" & vbNewLine & _
            "Where a.Id = [1] "
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mlngKey)
        If Not rs.EOF Then
            lngPatientID = Nvl(rs("����ID"), 0)
            mstrType = Nvl(rs("��������"))
            '�Ƿ�������������ʾ������Ŀ
            lngAdvice = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0)
            If lngAdvice = 0 Then
                lngAdvice = 0
            Else
                lngAdvice = Nvl(rs("ҽ��ID"), 0)
            End If
            mbytRedoNumber = IIf(mSelectRedo, mbytRedoNumber - 1, Nvl(rs("������"), 0))
            
            mSelectRedo = False
            
            mbytMode = IIf(IsNull(rs("����ID")), 0, 1)
            
            With sbrInfo
                .Panels(1).Text = "�����ˣ�" & Nvl(rs("������"))
                .Panels(2).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                If Nvl(rs("�����")) <> "" Then
                    .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                    .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                Else
                    If Nvl(rs("������")) <> "" Then
                        .Panels(3).Text = "�����ˣ�" & Nvl(rs("������"))
                        .Panels(4).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                    Else
                        .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                        .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                    End If
                End If
            End With
            blnMoved = MovedByDate(CDate(Format(Nvl(rs("����ʱ��")), "yyyy-MM-dd hh:mm:ss")))
            mLngPatientID = Nvl(rs("����ID"), 0)
            mstrPatientName = Nvl(rs("����"))
        Else
            lngPatientID = 0
            mstrType = ""
            mbytRedoNumber = 0
            
            mbytMode = 0
            
            With sbrInfo
                .Panels(1).Text = "�����ˣ�"
                .Panels(2).Text = "����ʱ�䣺"
                .Panels(3).Text = "����ˣ�"
                .Panels(4).Text = "���ʱ�䣺"
            End With
        End If
    End If
    '1-������2-ƫ�͡�3-ƫ�ߡ�4-����(�쳣)��5-�������ޡ�6-��������
    If mbytMode = 1 Then

'        mstrSQL = "Select /*+ rule */ �걾ID, ������Ŀid, �������, �̶���Ŀ, Id, ������Ŀ, ԭʼ���, �ϴν��, �ϴ�ʱ��, Cv," & vbNewLine & _
'                    "            Decode(���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', ���ν��) As ���ν��, Rownum As ���, ���㹫ʽ," & vbNewLine & _
'                    "            �������, ��־, ����id, �걾���, ����ʱ��, �걾���, �걾����ʾ, ���鱸ע, ����, �Ա�, ����, �����, סԺ��," & vbNewLine & _
'                    "            ��ǰ����, ��ҳid, �����Χ, С��, ��������, ��������, ��λ, �ο�" & vbNewLine & _
'                    "From (Select a.ID as �걾ID,b.������Ŀid, h.����, Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ," & vbNewLine & _
'                    "                           b.������Ŀid As Id, c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, b.ԭʼ���," & vbNewLine & _
'                    "                           '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv, b.������ As ���ν��, d.���㹫ʽ, d.�������," & vbNewLine & _
'                    "                           Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
'                    "                           Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
'                    "                           Decode(a.����id, Null," & vbNewLine & _
'                    "                                           To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
'                    "                                           a.�걾���) As �걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��," & vbNewLine & _
'                    "                           a.���� As ��ǰ����, a.��ҳid, d.�����Χ, Nvl(g.С��λ��, 2) As С��, d.��������, d.��������, d.��λ," & vbNewLine & _
'                    "                           Trim(Replace(Replace(' ' || Zlgetreference(c.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
'                    "                                                                                                                   a.����id, a.����), ' .', '0.'), '��.', '��0.')) As �ο�" & vbNewLine & _
'                    "            From ����걾��¼ a, ������ͨ��� b, ����������Ŀ c, ������Ŀ d, ����������Ŀ g, ������ĿĿ¼ h" & vbNewLine & _
'                    "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And" & vbNewLine & _
'                    "                        (g.����id = a.����id + 0 Or g.����id Is Null Or a.����id Is Null) And b.������Ŀid = g.��Ŀid(+) And" & vbNewLine & _
'                    "                        b.������Ŀid = h.Id(+) And b.��¼���� = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.ҽ��ID = [4] ")
'        mstrSQL = mstrSQL & " Union All" & vbNewLine & _
'                    "           Select a.ID as �걾ID,b.������Ŀid, h.����, Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ," & vbNewLine & _
'                    "                           b.������Ŀid As Id, c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, b.ԭʼ���," & vbNewLine & _
'                    "                           '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv, b.������ As ���ν��, d.���㹫ʽ, d.�������," & vbNewLine & _
'                    "                           Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
'                    "                           Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
'                    "                           Decode(a.����id, Null," & vbNewLine & _
'                    "                                           To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
'                    "                                           a.�걾���) As �걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��," & vbNewLine & _
'                    "                           a.���� As ��ǰ����, a.��ҳid, d.�����Χ, Nvl(g.С��λ��, 2) As С��, d.��������, d.��������, d.��λ," & vbNewLine & _
'                    "                           Trim(Replace(Replace(' ' || Zlgetreference(c.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
'                    "                                                                                                                   a.����id, a.����), ' .', '0.'), '��.', '��0.')) As �ο�" & vbNewLine & _
'                    "            From ����걾��¼ a, ������ͨ��� b, ����������Ŀ c, ������Ŀ d, ����������Ŀ g, ������ĿĿ¼ h" & vbNewLine & _
'                    "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And" & vbNewLine & _
'                    "                        (g.����id = a.����id + 0 Or g.����id Is Null Or a.����id Is Null) And b.������Ŀid = g.��Ŀid(+) And" & vbNewLine & _
'                    "                        b.������Ŀid = h.Id(+) And b.��¼���� = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.ҽ��ID = [4] ") & vbNewLine & _
'                    "            Order By ����, �������)"
        '2008-02-13 �޸����� �¶�
        mstrSQL = "Select /*+ rule */ Distinct A.�걾ID ,a.������Ŀid ,A.����, a.�������, a.�̶���Ŀ, a.Id, a.������Ŀ, a.ԭʼ���, a.�ϴν��, a.�ϴ�ʱ��, a.Cv," & vbNewLine & _
                    "            Decode(a.���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', a.���ν��) As ���ν��, Rownum As ���, a.���㹫ʽ," & vbNewLine & _
                    "            a.�������, a.��־, a.����id, a.�걾���, a.����ʱ��, a.�걾���, a.�걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��," & vbNewLine & _
                    "            a.��ǰ����, a.��ҳid, a.�����Χ, Nvl(G.С��λ��,2) as С��, " & vbNewLine & _
                    "            Trim(Replace(Replace(' ' || Zl_Get_Reference(4,a.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
                    "                           a.����id, a.����,a.�������ID), ' .', '0.'), '��.', '��0.')) As ��������," & vbNewLine & _
                    "            Trim(Replace(Replace(' ' || Zl_Get_Reference(3,a.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
                    "                           a.����id, a.����,a.�������ID), ' .', '0.'), '��.', '��0.')) As ��������," & vbNewLine & _
                    "            a.��λ,a.����ο� as �ο�, " & vbNewLine & _
                    "                           Trim(Replace(Replace(' ' || Zlgetreference(a.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������," & vbNewLine & _
                    "                                                                                                                   a.����id, a.����,a.�������ID), ' .', '0.'), '��.', '��0.')) As �ο�1," & vbNewLine & _
                    "            a.OD,a.CUTOFF,a.COV,a.ø���ID,a.���챨��,a.���쾯ʾ,lpad(����,4,'0') as ����,a.�걾����,a.������ʾ,decode(a.������˱�ʶ,1,'��',0,'��','') as ������˱�ʶ  " & vbNewLine & _
                    "From (Select A.id as �걾ID ,b.������Ŀid, decode(d.�������,Null,nvl(h.����,C.����),d.�������) as ����, Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ,"
                    
          mstrSQL = mstrSQL & " " & _
                    "                           b.������Ŀid As Id, " & vbNewLine & _
                    "                           " & IIf(chkChina.Value = 1, " c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ ", "d.��д as ������Ŀ ") & vbNewLine & _
                    "                           , b.ԭʼ���," & vbNewLine & _
                    "                           '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv, b.������ As ���ν��, d.���㹫ʽ, d.�������," & vbNewLine & _
                    "                           Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                    "                           Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
                    "                           Decode(a.����id, Null," & vbNewLine & _
                    "                                           To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                    "                                           a.�걾���) As �걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�걾����,a.��������,a.�����, a.סԺ��," & vbNewLine & _
                    "                           a.���� As ��ǰ����, a.��ҳid, d.�����Χ, d.��������, d.��������, d.��λ,b.OD,B.CUTOFF,B.SCO as COV,b.ø���ID, " & vbNewLine & _
                    "                           d.���챨���� as  ���챨��,d.���쾯ʾ�� as ���쾯ʾ,b.����ο�,a.�������ID ,e.�����Ƿ���� as  ������˱�ʶ,e.�������  as ������ʾ " & vbNewLine & _
                    ",Zl_To_Number(Zl_Get_Reference(1, b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������,a.����id, a.����,a.�������ID)) as �ο�ID " & vbNewLine & _
                    "            From ����걾��¼ a, ������ͨ��� b, ����������Ŀ c, ������Ŀ d, ������ĿĿ¼ h,������ˮ��ָ��  e" & vbNewLine & _
                    "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And  b.������Ŀid = h.Id(+) and  b.����걾id=e.�걾id(+) and  b.������Ŀid = e.��Ŀid(+) And b.��¼���� = [1] And " & IIf(lngAdvice = 0, " a.Id = [2] ", " a.ҽ��ID = [4] ")
        mstrSQL = mstrSQL & " Union All" & vbNewLine & _
                    "           Select a.id as �걾ID ,b.������Ŀid, decode(d.�������,Null,nvl(h.����,C.����),d.�������) as ����, Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ," & vbNewLine & _
                    "                           b.������Ŀid As Id, " & _
                    "                           " & IIf(chkChina.Value = 1, "c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ", "d.��д as ������Ŀ ") & vbNewLine & _
                    "                           , b.ԭʼ���," & vbNewLine & _
                    "                           '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv, b.������ As ���ν��, d.���㹫ʽ, d.�������," & vbNewLine & _
                    "                           Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                    "                           Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
                    "                           Decode(a.����id, Null," & vbNewLine & _
                    "                                           To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                    "                                           a.�걾���) As �걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�걾����,a.��������, a.�����, a.סԺ��," & vbNewLine & _
                    "                           a.���� As ��ǰ����, a.��ҳid, d.�����Χ, d.��������, d.��������, d.��λ,b.OD,B.CUTOFF,B.SCO as COV,b.ø���ID, " & vbNewLine & _
                    "                           d.���챨���� as ���챨��,d.���쾯ʾ�� as ���쾯ʾ,b.����ο�,a.�������ID ,e.�����Ƿ���� as  ������˱�ʶ ,e.������� as ������ʾ   " & vbNewLine & _
                    ",Zl_To_Number(Zl_Get_Reference(1, b.������Ŀid, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������,a.����id, a.����,a.�������ID)) as �ο�ID " & vbNewLine & _
                    "            From ����걾��¼ a, ������ͨ��� b, ����������Ŀ c, ������Ŀ d,  ������ĿĿ¼ h,������ˮ��ָ��  e" & vbNewLine & _
                    "            Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And" & vbNewLine & _
                    "                        b.������Ŀid = h.Id(+) And b.��¼���� = [1]  and  b.����걾id=e.�걾id(+) and  b.������Ŀid = e.��Ŀid(+)  And a.�ϲ�id = [2]" & vbNewLine & _
                    "            ) A ,����������Ŀ G,������Ŀ�ο� F" & _
                    "  Where A.����id = G.����id(+) And A.ID = G.��Ŀid(+) and a.�ο�id=f.id(+)"

    Else
'        mstrSQL = "Select /*+ rule */ A.*,Rownum As ��� From (SELECT a.����걾ID as �걾ID,A.������ĿID,A.�������,B.ID," & _
'                        "B.������||DECODE(C.��д,NULL,'','('||C.��д||')') AS ������Ŀ," & _
'                        "A.ԭʼ���," & _
'                        "'' As �ϴν��,'' as �ϴ�ʱ�� ,''As CV," & _
'                        "A.������ As ���ν��," & _
'                        "C.���㹫ʽ," & _
'                        "C.�������," & _
'                        "DECODE(A.�����־,3,'��',2,'��',1,'',4,'�쳣',5,'����',6,'����','') AS ��־," & _
'                        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,D.�걾����,0,NULL,D.����ID),' .','0.'),'��.','��0.')) AS �ο�," & _
'                        "Nvl(D.����ID,-1) As ����ID,Nvl(D.�걾���,0) As �걾���,D.����ʱ��,D.�걾���,D.���鱸ע,C.�����Χ,0 As �̶���Ŀ,Nvl(X.С��λ��,2) AS С��," & _
'                        "C.��������,C.��������,C.��λ " & _
'                    "FROM ������ͨ��� A,����������Ŀ B,������Ŀ C,����걾��¼ D,����������Ŀ X " & _
'                    "WHERE A.������Ŀid = B.ID " & _
'                        "AND B.ID = C.������ĿID " & _
'                        "AND A.��¼���� = [1] " & _
'                        "AND D.ID=A.����걾ID " & _
'                        "AND A.������Ŀid=X.��ĿID(+) AND (X.����ID=D.����ID+0 OR X.����ID IS NULL OR D.����ID IS NULL) " & _
'                        "AND D.ID= [2] Order By B.����) A"
        '2008-02-13 �޸����� �¶�
        mstrSQL = "Select /*+ rule */ Distinct A.�걾ID,A.������Ŀid, A.����, a.�������, a.Id, a.������Ŀ, a.ԭʼ���, a.�ϴν��, a.�ϴ�ʱ��, a.Cv," & _
                  "a.���ν��,a.���㹫ʽ,a.�������,a.��־,a.������ʾ,decode(a.������˱�ʶ,1,'��',0,'��','') as ������˱�ʶ," & vbNewLine & _
                  "Trim(REPLACE(REPLACE(' '||zlGetReference(A.ID,A.�걾����,0,NULL,A.����ID),' .','0.'),'��.','��0.')) AS �ο�1,a.�걾����,a.����ο� as �ο�, " & _
                  " a.����ID,a.�걾���,a.����ʱ��,a.�걾���,a.���鱸ע,a.�����Χ,a.�̶���Ŀ,Nvl(X.С��λ��,2) AS С��," & _
                  " Trim(REPLACE(REPLACE(' '||Zl_Get_Reference(4,A.ID,A.�걾����,0,NULL,A.����ID),' .','0.'),'��.','��0.')) AS ��������, " & _
                  " Trim(REPLACE(REPLACE(' '||Zl_Get_Reference(3,A.ID,A.�걾����,0,NULL,A.����ID),' .','0.'),'��.','��0.')) AS ��������, " & _
                  "a.��λ" & _
                  ",Rownum As ���,A.OD,A.CUTOFF,A.COV,a.ø���Id,a.���챨��,a.���쾯ʾ,lpad(����,4,'0') as ���� From (SELECT D.id as �걾ID,A.������ĿID,A.�������,B.ID," & _
                        IIf(chkChina.Value = 1, "B.������||DECODE(C.��д,NULL,'','('||C.��д||')') AS ������Ŀ", "C.��д as ������Ŀ ") & vbNewLine & _
                        ",decode(c.�������,Null,nvl(h.����,b.����),c.�������) as ����," & _
                        "A.ԭʼ���," & _
                        "'' As �ϴν��,'' as �ϴ�ʱ�� ,''As CV," & _
                        "A.������ As ���ν��," & _
                        "C.���㹫ʽ," & _
                        "C.�������," & _
                        "DECODE(A.�����־,3,'��',2,'��',1,'',4,'�쳣',5,'����',6,'����','') AS ��־," & _
                        "Nvl(D.����ID,-1) As ����ID,Nvl(D.�걾���,0) As �걾���,D.����ʱ��,D.�걾���,D.���鱸ע,C.�����Χ,0 As �̶���Ŀ," & _
                        "C.��������,C.��������,C.��λ,D.�걾���� ,A.OD,A.CUTOFF,A.SCO as COV,a.ø���ID,c.���챨���� as ���챨��,c.���쾯ʾ�� as ���쾯ʾ,a.����ο� ,e.�����Ƿ���� as  ������˱�ʶ,e.�������  as ������ʾ " & _
                        ",Zl_To_Number(Zl_Get_Reference(1, a.������Ŀid, d.�걾����, Decode(d.�Ա�, '��', 1, 'Ů', 2, 0), d.��������,d.����id, d.����,d.�������ID)) as �ο�ID " & vbNewLine & _
                    "FROM ������ͨ��� A,����������Ŀ B,������Ŀ C,����걾��¼ D,������ĿĿ¼ H ,������ˮ��ָ��  e " & _
                    "WHERE A.������ĿID=H.ID(+) And A.������Ŀid = B.ID " & _
                        "AND B.ID = C.������ĿID and  a.����걾id=e.�걾id(+) and  a.������Ŀid = e.��Ŀid(+)  " & _
                        "AND A.��¼���� = [1] " & _
                        "AND D.ID=A.����걾ID " & _
                        "AND D.ID= [2] ) A,����������Ŀ X,������Ŀ�ο� F where a.����ID=X.����ID(+) and A.ID=X.��ĿID(+) and A.�ο�id=f.id(+)"
    End If
    
    If blnMoved Then
        strSQLbak = mstrSQL
        strSQLbak = Replace(strSQLbak, "����걾��¼", "H����걾��¼")
        strSQLbak = Replace(strSQLbak, "������ͨ���", "H������ͨ���")
        strSQLbak = Replace(strSQLbak, "����������Ŀ", "H����������Ŀ")
        mstrSQL = mstrSQL & " Union ALL " & strSQLbak
    End If
    
    mstrSQL = mstrSQL & " Order by ����,������� "
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mbytRedoNumber, IIf(mlngKey = 0, -1, mlngKey), lngPatientID, lngAdvice)
    
    If rs.BOF = False Then
        '��ʼ�걾��Ϣ
        mDeviceID = rs("����ID")
        Me.txtComment.Text = Nvl(rs("���鱸ע"))
        
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Call ReadVsf(rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
        rs.MoveFirst
        
        Call FormatVsfCell(Vsf, mCol.������, "0.0######", IIf(Nvl(rs("�������"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.С��, -1))
                
        Call FormatVsfCell(Vsf, mCol.ԭʼ���, "0.0######", IIf(Nvl(rs("�������"), 0) = 1, 0, 1), _
                IIf(mDeviceID > 0, mCol.С��, -1))
        
'        If chkLast.Value Then LoadLastValue
        '--ÿ�ζ�������ʷ���
        LoadLastValue
    Else
        mDeviceID = -1
        Me.txtComment.Text = ""
        ResetVsf Vsf
    End If
    
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Call ApplyResultColor(Vsf, mlngLoop, mCol.������ + intCol * mintColCount, _
                Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 6, "����", 5, 1))
        Next
    Next
    
    'д�������Ϣ
    Me.txtDiagnose.Text = ""
    gstrSql = "Select b.ҽ��id, b.��Ŀ, b.����, b.����" & vbNewLine & _
                "From ����걾��¼ a, ����ҽ������ b" & vbNewLine & _
                "Where a.ҽ��id = b.ҽ��id and a.ID = [1] " & vbNewLine & _
                "Order By ҽ��id, ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("��Ŀ")) & ":" & Replace(Nvl(rs("����")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txtDiagnose.Text = strTmp
    
    If mbytRedoNumber > 0 Then
        gstrSql = "select ���鱸ע from ������ͨ��� where ����걾id = [1] and ��¼���� = [2] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, mbytRedoNumber)
        If rs.EOF = False Then
            txtComment.Text = rs("���鱸ע") & ""
        End If
    End If
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub FormatVsfCell(objVsf As Object, ByVal lngCol As Long, ByVal strFormat As String, Optional ByVal iType As Integer = -1, Optional ByVal iTypeCol As Integer = -1)
    'iType�����ʽ������������
    '  0�����֡�1���ַ���2�����ڡ�3���߼���-1�����ޣ�ȱʡ��
    'iTypeCol��С��λ���Ĵ洢�ֶ����
    Dim lngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    intColCount = GetColCount(objVsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For lngLoop = 1 To objVsf.Rows - 1
            If iType = 0 And IsNumeric("-" & objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount)) And iTypeCol <> -1 Then
                If InStr(UCase(objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount)), "E") = 0 Then
                    objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount) = Format(objVsf.TextMatrix(lngLoop, lngCol + intCol * mintColCount), _
                        IIf(Val(objVsf.TextMatrix(lngLoop, iTypeCol + intCol * mintColCount)) = 0, "#0", "0." & String(Val(objVsf.TextMatrix(lngLoop, iTypeCol + intCol * mintColCount)), "0")))
                End If
            Else
                '�����޸ģ��붨���Ͷ��ԵĲ���Ҫ��ʽ��(�п�ѧ������)
    '            If IsNumeric("-" & objVsf.TextMatrix(lngLoop, lngCol)) Then objVsf.TextMatrix(lngLoop, lngCol) = Format(objVsf.TextMatrix(lngLoop, lngCol), strFormat)
            End If
        Next
    Next
End Sub

Public Function zlRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mblnLoadHistory = False ' IIf(mlngKey <> lngKey, False, mblnLoadHistory)
    mlngKey = lngKey
    fraComment.Tag = ""
    Call Form_Resize
'    SetEditState False
    '��ʼ�����б�
    If ReadData = False Then Exit Function
    
    zlRefresh = True
End Function
Public Function zlRefreshPatient(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ���ݣ�������)
    '������lngkey ����ID
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mLngPatientID = lngKey
    fraComment.Tag = "����ʾ"
    zlRefreshPatient = True
    Call ReadPatient
End Function

Public Function ZlEditStart(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ��༭����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer
    
    SetEditState True
    
    If mlngKey <> lngKey Then
        mlngKey = lngKey
        If ReadData = False Then Exit Function
    End If
    
    mblnChangeEdit = False
    ZlEditStart = True
    intColCount = GetColCount(Vsf.Col)
    With Vsf
        If .Col = mCol.������ + intColCount * mintColCount Or .Col = mCol.�����־ + intColCount * mintColCount _
           Or .Col = mCol.od + intColCount * mintColCount Then
            Select Case .Col
                Case mCol.������ + intColCount * mintColCount
                    .EditMode(.Col) = 1
                Case mCol.�����־ + intColCount * mintColCount
                    If .TextMatrix(.Row, .Col) <> "����" And .TextMatrix(.Row, .Col) <> "����" Then .EditMode(.Col) = 1
                Case mCol.od + intColCount * mintColCount
                    .EditMode(.Col) = 1
            End Select
        Else
            .Col = mCol.������ + intColCount * mintColCount
        End If
        '����������걾��������޽����ָ�꿪ʼ��д
'        If mDeviceID > 0 And Not mblnEvent Then
'            For i = 1 To .Rows - 1
'                If Len(Trim(.TextMatrix(i, mCol.������))) = 0 Then Exit For
'            Next
'            If i <= .Rows - 1 And i >= 1 Then
'                .Row = i
'                .ShowCell .Row, mCol.������
'            End If
'        End If
        mblnEvent = False
        
        .SetFocus
    End With
End Function

Public Function ZlSave() As Boolean
    If SaveData() = False Then Exit Function

    ZlSave = True
End Function

Public Function ZlCancel() As Boolean
    '��ʾ�Ƿ񱣴�
    SetEditState False
    '������ʾ������
    mblnLoadHistory = False
    
    Vsf.EditMode(mCol.������) = 1: Vsf.EditMode(mCol.�����־) = 0
    Vsf.Rows = 2
    Vsf.Cols = mintColCount
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    Vsf.Cell(flexcpData, 1, 0, Vsf.Rows - 1, Vsf.Cols - 1) = 0
    'Call ReadData
    
    ZlCancel = True
End Function

Public Function ZlClearForm() As Boolean
    '��ս��
    mblnLoadHistory = False
    mlngKey = 0
    With sbrInfo
        .Panels(1).Text = "�����ˣ�"
        .Panels(2).Text = "����ʱ�䣺"
        .Panels(3).Text = "����ˣ�"
        .Panels(4).Text = "���ʱ�䣺"
    End With
    Me.txtComment = ""
    ResetVsf Vsf
End Function

Private Sub SetEditState(ByVal blnEdit As Boolean)
    Dim intColCount As Integer
    mblnEdit = blnEdit
'    vsf.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    txtComment.Locked = Not blnEdit
    Me.lvwSelect.Visible = blnEdit
    intColCount = GetColCount(Me.Vsf.Col)
    If Me.lvwSelect.Visible Then
        ShowValue 1, Val(Vsf.TextMatrix(Vsf.Row, mCol.������� + intColCount * mintColCount)), Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)
    End If
    Call Form_Resize
End Sub

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim bytResultFlag As Byte, mlngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    Dim intLoop As Integer

    Dim strsql() As String

    Dim strTmp As String, rsTmp As ADODB.Recordset
    If Vsf.Rows > 1 Then
        Vsf.Row = Vsf.Row - 1
        Vsf.Row = Vsf.Row + 1
    Else
        Vsf.Row = Vsf.Row + 1
        Vsf.Row = Vsf.Row - 1
    End If
    If Not mblnChangeEdit Then SaveData = True: Exit Function

    On Error GoTo ErrHand
    ReDim strsql(1 To 1)
    '��ȡ����ʱ��
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    '��ȡ����걾��¼��Ϣ
    strTmp = "Select ҽ��id ,����ʱ�� ,������ , ����ʱ�� ,nvl(����,0) as ����, " & _
        "�������� , �������� , ִ�п���id , ������ ,����ʱ�� ,�걾���,����� " & _
        "From ����걾��¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, Me.Caption, mlngKey)
    If rsTmp.EOF Then
        SaveData = False
        Exit Function
    Else
        If rsTmp!����� & "" <> "" Then
            MsgBox "�ñ걾�ѱ������û���ˣ�", vbInformation, gstrSysName
            SaveData = False
            mblnChangeEdit = False
            Call frmLabMain.zlRefreshData
            Exit Function
        End If
    End If
    
    Vsf.SetFocus
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For mlngLoop = 1 To Vsf.Rows - 1
            'If Val(vsf.RowData(mlngLoop)) > 0 Then
            If Val(Vsf.Cell(flexcpData, mlngLoop, intCol * mintColCount, mlngLoop, intCol * mintColCount)) > 0 Then
                bytResultFlag = 0
                If Trim(Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount)) <> "" Then
                    bytResultFlag = Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1)
                End If
                '����ֻ�����˽�����ϵ㱣������
                If mlngLoop = Vsf.Row And mCol.������ + intCol * mintColCount = Vsf.Col And Vsf.EditText <> "" Then
                    Vsf.TextMatrix(mlngLoop, mCol.������ + intCol * mintColCount) = Vsf.EditText
                End If
                strsql(ReDimArray(strsql)) = "ZL_����걾��¼_������д(" & CLng(Vsf.TextMatrix(mlngLoop, mCol.�걾ID + intCol * mintColCount)) & "," & _
                Val(Vsf.Cell(flexcpData, mlngLoop, intCol * mintColCount, mlngLoop, intCol * mintColCount)) & "," & _
                mbytRedoNumber & ",'" & Vsf.TextMatrix(mlngLoop, mCol.������ + intCol * mintColCount) & "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                IIf(bytResultFlag = 0, "NULL", bytResultFlag) & ",'" & Vsf.TextMatrix(mlngLoop, mCol.����ο� + intCol * mintColCount) & "',1,NULL,0," & IIf(intCol = 0 And mlngLoop = 1, 1, 0) & _
                ",'" & Vsf.TextMatrix(mlngLoop, mCol.ԭʼ��� + intCol * mintColCount) & "'," & Vsf.TextMatrix(mlngLoop, mCol.������Ŀid + intCol * mintColCount) & _
                "," & IIf(Vsf.TextMatrix(mlngLoop, mCol.������� + intCol * mintColCount) = "", Vsf.TextMatrix(mlngLoop, intCol * mintColCount), Vsf.TextMatrix(mlngLoop, mCol.������� + intCol * mintColCount)) & _
                ",'" & Vsf.TextMatrix(mlngLoop, mCol.od + intCol * mintColCount) & _
                "','" & Vsf.TextMatrix(mlngLoop, mCol.CUTOFF + intCol * mintColCount) & "','" & Vsf.TextMatrix(mlngLoop, mCol.COV + intCol * mintColCount) & _
                "'," & IIf(Vsf.TextMatrix(mlngLoop, mCol.ø���ID + intCol * mintColCount) = "", "Null", Vsf.TextMatrix(mlngLoop, mCol.ø���ID + intCol * mintColCount)) & _
                ",'" & txtComment & "','" & UserInfo.���� & "',1)"
                intLoop = intLoop + 1
            End If
        Next
    Next
    
    If intLoop = 0 Then
        strsql(ReDimArray(strsql)) = "Zl_������ͨ���_Delete(" & mlngKey & ")"
    End If
    
    
    '���¼��������Ŀ
    strsql(ReDimArray(strsql)) = "Zl_���¼�����_Cale(" & mlngKey & ")"
    
    blnTran = True

    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strsql)
        If strsql(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(Replace(strsql(mlngLoop), ",,", ",Null,"), Me.Caption)
    Next
    gcnOracle.CommitTrans
    gstrSql = "ZL_����걾��¼_����ѡ��(" & mlngKey & "," & mbytRedoNumber & ",'" & txtComment & "')"
    zlDatabase.ExecuteProcedure gstrSql, gstrSysName
    
    If Signature(mlngKey, gstrDBUser, "����") = False Then
        Exit Function
    End If

    

    SaveData = True
    mblnEdit = False
    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim p As POINTAPI
    Dim intColCount As Integer
    Dim lngItemID As Long
    
    If Control.ID = 100 Then
        
        intColCount = GetColCount(Me.Vsf.Col)
        lngItemID = Val(Vsf.Cell(flexcpData, Me.Vsf.Row, intColCount * mintColCount, Me.Vsf.Row, intColCount * mintColCount))
        Call GetCursorPos(p)
        With frmLisStationWriteInfo
            .Top = p.Y * Screen.TwipsPerPixelY
            .Left = p.X * Screen.TwipsPerPixelX
            .ShowME Me, lngItemID
        End With
    Else
        mbytRedoNumber = Control.ID
        mSelectRedo = True
        mblnLoadHistory = False
        ReadData
    End If
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID = mbytRedoNumber + 1 Then Control.Checked = True
End Sub

Private Sub chkChina_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)

    mblnLoadHistory = False
    If fraComment.Tag = "" Then
        Call ReadData
    Else
        ReadPatient
    End If
End Sub

Private Sub chkLast_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
        
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.�ϴν�� + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.�ϴ�ʱ�� + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
    Next
'    vsf.Body.ColWidth(mCol.CV) = IIf(chkLast.Value, 400, 0)
    
    If chkLast.Value Then LoadLastValue
End Sub

Private Sub chkLastDate_Click()

End Sub

Private Sub chkMB_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
    Next
End Sub

Private Sub chkOriginal_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.ԭʼ��� + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
    Next
End Sub

Private Sub chkReferrence_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.����ο� + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
    Next
End Sub

Private Sub chkSign_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.�����־ + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
    Next
End Sub

Private Sub chkUnit_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.��λ + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
    Next
End Sub

Private Sub chkYiQiBiaoShi_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.������˱�ʶ + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
End Sub

Private Sub chkYiQiTiShi_Click()
    Dim intColCount As Integer, intCol As Integer
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.������ʾ + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
    Next
End Sub

Private Sub Form_Load()
    With Vsf
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "������Ŀ", 2100, 1
        .NewColumn "ԭʼ���", 0, 1
        .NewColumn "���ν��", 1200, 1, , 1
        .NewColumn "��λ", 1000, 1
        .NewColumn "CV", 0, 1
        .NewColumn "��־", 450, 1
        .NewColumn "�ϴν��", 0, 1
        .NewColumn "�ϴ�ʱ��", 0, 1
        .NewColumn "�ο�", 1300, 1
        .NewColumn "�������", 0, 1
        .NewColumn "����id", 0, 1
        .NewColumn "���㹫ʽ", 0, 1
        .NewColumn "�����Χ", 0, 1
        .NewColumn "�̶���Ŀ", 0, 1
        .NewColumn "С��", 0, 1
        .NewColumn "��������", 0, 1
        .NewColumn "��������", 0, 1
        .NewColumn "������ĿID", 0, 1
        .NewColumn "�������", 0, 1
        .NewColumn "�걾ID", 0, 1
        .NewColumn "OD", 700, 1, , 1
        .NewColumn "CUTOFF", 700, 1
        .NewColumn "COV", 700, 1
        .NewColumn "ø���ID", 0, 1
        .NewColumn "���챨��", 0, 1
        .NewColumn "���쾯ʾ", 0, 1
        .NewColumn "������ʾ", 1000, 1
        .NewColumn "������˱�ʶ", 1200, 1
        .FixedCols = 0
    End With
    lvwSelect.Tag = 1 'Ĭ��ѡ��ָ����
    mblnLoadHistory = False
    
    If mblnPatientFind = False Then
        'ȡ����ѡ��
        chkOriginal.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴ԭʼ���", 100, 1208, 0))
        chkLast.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴�ϴν��", 100, 1208, 0))
        chkSign.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴��־", 100, 1208, 1))
        chkUnit.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴��λ", 100, 1208, 1))
        chkReferrence.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴�ο�", 100, 1208, 1))
        chkMB.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴ø��", 100, 1208, 1))
        chkChina.Value = Val(zlDatabase.GetPara("frmLisStationWrite_�鿴����", 100, 1208, 1))
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������ʾ", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������˱�ʶ", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
        
    End If
    Vsf.Body.ColWidth(mCol.�ϴν��) = IIf(chkLast.Value, 900, 0)
    Vsf.Body.ColWidth(mCol.ԭʼ���) = IIf(chkOriginal.Value, 900, 0)
    Vsf.Body.ColWidth(mCol.od) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.CUTOFF) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.COV) = IIf(chkMB.Value, 700, 0)
    Vsf.Body.ColWidth(mCol.������Ŀ) = IIf(chkChina.Value, 1000, 2100)
    Vsf.Body.ColWidth(mCol.������ʾ) = IIf(chkYiQiTiShi.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.������˱�ʶ) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    SetEditState False
    
    '��ȡ��ɫ
    lngReferenceLow = Val(zlDatabase.GetPara("�ο���ɫ_ƫ��", 100, 1208, 0))
    If lngReferenceLow = 0 Then lngReferenceLow = 8454143
    lblLow.BackColor = lngReferenceLow
    lngReferenceHigh = Val(zlDatabase.GetPara("�ο���ɫ_ƫ��", 100, 1208, 0))
    If lngReferenceHigh = 0 Then lngReferenceHigh = 8438015
    lblHigh.BackColor = lngReferenceHigh
    lngReferenceExigency = Val(zlDatabase.GetPara("�ο���ɫ_��ʾ", 100, 1208, 0))
    If lngReferenceExigency = 0 Then lngReferenceExigency = 16576
    lblExigency.BackColor = lngReferenceExigency
    
    
    Call RestoreFlexState(Vsf, Me.Name)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fraComment
        .Left = 0
        .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        .Width = Me.ScaleWidth - .Left
    End With
    With txtComment
'        .Width = Me.fraComment.Width - .Left - txtDiagnose.Width - Me.Label2.Width
        .Width = Me.fraComment.Width / 2
        .Height = fraComment.Height - 20
    End With
    
    With Me.Label2
        .Left = Me.txtComment.Left + Me.txtComment.Width + 20
    End With
    
    With Me.txtDiagnose
        .Left = Me.Label2.Left + Me.Label2.Width + 20
        .Width = fraComment.Width - txtComment.Left - txtComment.Width - 400
        .Height = txtComment.Height
    End With
        
    With fraTitle
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
    End With
    
    With lvwSelect
        .Left = Me.ScaleWidth - .Width - 30
        .Top = fraTitle.Top + fraTitle.Height + 30
        .Height = fraComment.Top - .Top + 30
    End With
    If fraComment.Tag = "" Then
        fraComment.Visible = zlDatabase.GetPara("��ʾ���鱸ע", 100, 1208, True)
    Else
        fraComment.Visible = False
    End If
    
    With Vsf
        .Left = -15
        .Top = fraTitle.Top + fraTitle.Height + 30
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) - 30 - .Left
        If fraComment.Visible Then
            .Height = fraComment.Top - .Top + 30
        Else
            .Height = fraComment.Top + fraComment.Height - .Top + 30
        End If
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Vsf, Me.Name)
    
    zlDatabase.SetPara "frmLisStationWrite_�鿴ԭʼ���", Me.chkOriginal.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴�ϴν��", Me.chkLast.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴��־", Me.chkSign.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴��λ", Me.chkUnit.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴�ο�", Me.chkReferrence.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴ø��", Me.chkMB.Value, 100, 1208
    zlDatabase.SetPara "frmLisStationWrite_�鿴����", Me.chkChina.Value, 100, 1208
    fraComment.Tag = ""
    '�˳�����ʱ��ԭ����
    mblnEdit = False
End Sub

Private Sub lblExigency_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblExigency.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("�ο���ɫ_��ʾ", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lblHigh_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblHigh.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("�ο���ɫ_ƫ��", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lblLow_DblClick()
    CommDialog.ShowColor
    If CommDialog.COLOR <> 0 Then
        lblLow.BackColor = CommDialog.COLOR
        Call zlDatabase.SetPara("�ο���ɫ_ƫ��", CommDialog.COLOR, 100, 1208)
    End If
End Sub

Private Sub lvwSelect_DblClick()
    Dim intColCount As Integer
    
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    If Not mblnEdit Then Exit Sub
    
    On Error GoTo errH
    
    intColCount = GetColCount(Vsf.Col)
    
    Select Case Val(lvwSelect.Tag)
        Case 1 'ѡ����
            If Val(Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.������ + intColCount * mintColCount) = lvwSelect.SelectedItem.Text
                '����ȱʡ�Ľ����־
                Vsf.TextMatrix(Vsf.Row, mCol.�����־ + intColCount * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(Vsf.Row, mCol.������ + intColCount * mintColCount)), _
                Trim(Vsf.TextMatrix(Vsf.Row, mCol.����ο� + intColCount * mintColCount)), Val(Vsf.TextMatrix(Vsf.Row, mCol.������� + intColCount * mintColCount)), _
                    Vsf.TextMatrix(Vsf.Row, mCol.�������� + intColCount * mintColCount), Vsf.TextMatrix(Vsf.Row, mCol.�������� + intColCount * mintColCount))
                
                '���ݽ��Ӧ����ɫ��־
                Call ApplyResultColor(Vsf, Vsf.Row, mCol.������ + intColCount * mintColCount, _
                    Decode(Vsf.TextMatrix(Vsf.Row, mCol.�����־ + intColCount * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
                
                Vsf.SetFocus
                
                mblnChangeEdit = True
            End If
        Case 2 'ѡ��ע
            Me.txtComment.SelText = lvwSelect.SelectedItem.Text
            
            mblnChangeEdit = True
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picFilter_Click()
    Dim p As POINTAPI
    Dim blnOriginal As Boolean
    Dim blnLast As Boolean
    Dim blnChina As Boolean
    Dim blnSign As Boolean
    If Me.picFilter.Tag = "" Then
        Me.picFilter.Tag = "True"
    Else
        Me.picFilter.Tag = ""
    End If
    Call GetCursorPos(p)
    With frmLabMainSizer
        .Top = p.Y * Screen.TwipsPerPixelY
        .Left = p.X * Screen.TwipsPerPixelX
        .ShowME Me, "frmLisStationWrite", IIf(Me.picFilter.Tag = "", True, False)
    End With

    If mblnPatientFind = False Then
'        'ȡ����ѡ��
        chkOriginal.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴ԭʼ���", 100, 1208, 0), 1, 0)
        chkLast.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴�ϴν��", 100, 1208, 0), 1, 0)
        chkSign.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴��־", 100, 1208, 1), 1, 0)
        chkUnit.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴��λ", 100, 1208, 1), 1, 0)
        chkReferrence.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴�ο�", 100, 1208, 1), 1, 0)
        chkMB.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴ø��", 100, 1208, 1), 1, 0)
        chkChina.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴����", 100, 1208, 1), 1, 0)
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������ʾ", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������˱�ʶ", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
    End If
End Sub

Private Sub picFilter_LostFocus()
    frmLabMainSizer.ShowME Me, "frmLisStationWrite", True
    Me.picFilter.Tag = ""
    If mblnPatientFind = False Then
'        'ȡ����ѡ��
        chkOriginal.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴ԭʼ���", 100, 1208, 0), 1, 0)
        chkLast.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴�ϴν��", 100, 1208, 0), 1, 0)
        chkSign.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴��־", 100, 1208, 1), 1, 0)
        chkUnit.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴��λ", 100, 1208, 1), 1, 0)
        chkReferrence.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴�ο�", 100, 1208, 1), 1, 0)
        chkMB.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴ø��", 100, 1208, 1), 1, 0)
        chkChina.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_�鿴����", 100, 1208, 1), 1, 0)
        chkYiQiTiShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������ʾ", 100, 1208, 1), 1, 0)
        chkYiQiBiaoShi.Value = IIf(zlDatabase.GetPara("frmLisStationWrite_������˱�ʶ", 100, 1208, 1), 1, 0)
    Else
        chkChina.Value = 1
        chkSign.Value = 1
    End If
End Sub

Private Sub picFilter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFilter.BackColor = &HFFFFFF
    picFilter.BorderStyle = 1
End Sub

Private Sub picFilter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFilter.BackColor = &H8000000F
    picFilter.BorderStyle = 0
End Sub

Private Sub txtComment_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtComment_GotFocus()
    With txtComment
'        .SelStart = 0
'        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 2
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        RaiseEvent StartEdit(False)
        mblnChangeEdit = True
        txtComment.SetFocus
'        txtComment.SelLength = 0
    End If
End Sub

Private Sub vsf_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mblnChangeEdit = True
'    Call RenumVsf(vsf, 0)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strReference As String, strItemIDs As String
    Dim lngCount As Long, mlngLoop As Long
    Dim intColCount As Integer
    Dim intCol As Integer, intCols As Integer
    Dim str���� As String
    On Error GoTo errH
    
    intColCount = GetColCount(Col)
    Select Case Col
        Case mCol.�����־ + intColCount * mintColCount
            Select Case Val(Left(Vsf.TextMatrix(Row, mCol.�����־ + intColCount * mintColCount), 1))
                Case 3
                    Vsf.TextMatrix(Row, Col) = "��"
                Case 2
                    Vsf.TextMatrix(Row, Col) = "��"
                Case 4
                    Vsf.TextMatrix(Row, Col) = "�쳣"
                Case 5
                    Vsf.TextMatrix(Row, Col) = "����"
                Case 6
                    Vsf.TextMatrix(Row, Col) = "����"
                Case Else
                    Vsf.TextMatrix(Row, Col) = ""
            End Select
            Call ApplyResultColor(Vsf, Row, mCol.������ + intColCount * mintColCount, _
                Decode(Vsf.TextMatrix(Row, mCol.�����־ + intColCount * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
        Case mCol.������ + intColCount * mintColCount
            '�ȴ����Ƿ����ģ��
            If Left(Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount), 1) = "/" Then
                If LoadModel(Mid(Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount), 2)) Then
                    mblnChangeEdit = True
                    Exit Sub
                End If
            End If
            '��ʽ�����
            If Val(Vsf.TextMatrix(Row, mCol.������� + intColCount * mintColCount)) <> 2 And Val(Vsf.TextMatrix(Row, mCol.������� + intColCount * mintColCount)) <> 3 And IsNumeric(Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount)) Then
                If InStr(Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount), "E") = 0 Then
                    Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount) = Format(Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount), _
                        "0" & IIf(Val(Vsf.TextMatrix(Row, mCol.С�� + intColCount * mintColCount)) = 0, "", "." & String(Val(Vsf.TextMatrix(Row, mCol.С�� + intColCount * mintColCount)), "0")))
                
                    str���� = Get������(mlngKey, Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount), Vsf.TextMatrix(Row, mCol.������ + intColCount * mintColCount))
                    If str���� <> "" Then
                        Vsf.TextMatrix(Row, mCol.������Ŀ + intColCount * mintColCount) = Trim(Replace(Vsf.TextMatrix(Row, mCol.������Ŀ + intColCount * mintColCount), "�踴��", "")) & " " & str����
                        Vsf.Cell(flexcpForeColor, Row, mCol.������Ŀ + intColCount * mintColCount, Row, mCol.������Ŀ + intColCount * mintColCount) = COLOR.��ɫ
                    Else
                        Vsf.TextMatrix(Row, mCol.������Ŀ + intColCount * mintColCount) = Trim(Replace(Vsf.TextMatrix(Row, mCol.������Ŀ + intColCount * mintColCount), "�踴��", ""))
                    End If
                End If
            End If
            
            '����ȱʡ�Ľ����־
            Vsf.TextMatrix(Row, mCol.�����־ + intColCount * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(Row, Col)), Trim(Vsf.TextMatrix(Row, mCol.����ο� + intColCount * mintColCount)), Val(Vsf.TextMatrix(Row, mCol.������� + intColCount * mintColCount)), _
                Vsf.TextMatrix(Row, mCol.�������� + intColCount * mintColCount), Vsf.TextMatrix(Row, mCol.�������� + intColCount * mintColCount), _
                Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount))
            
            '���ݽ��Ӧ����ɫ��־
            Call ApplyResultColor(Vsf, Row, mCol.������ + intColCount * mintColCount, _
                Decode(Vsf.TextMatrix(Row, mCol.�����־ + intColCount * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
            
            '�Զ����������Ŀ���
            intCols = GetColCount(Vsf.Cols)
            If intCols = 0 Then intCols = 1
            For intCol = 0 To intCols - 1
                For mlngLoop = 1 To Vsf.Rows - 1
                    If Trim(Vsf.TextMatrix(mlngLoop, mCol.���㹫ʽ + intCol * mintColCount)) <> "" Then
                        
                        
                        Vsf.TextMatrix(mlngLoop, mCol.������ + intCol * mintColCount) = Format(CalcExpress(Vsf, Trim(Vsf.TextMatrix(mlngLoop, mCol.���㹫ʽ + intCol * mintColCount))), _
                            "0" & IIf(Val(Vsf.TextMatrix(mlngLoop, mCol.С�� + intCol * mintColCount)) = 0, "", "." & String(Val(Vsf.TextMatrix(mlngLoop, mCol.С�� + intCol * mintColCount)), "0")))
                        If CalcExpress(Vsf, Trim(Vsf.TextMatrix(mlngLoop, mCol.���㹫ʽ + intCol * mintColCount))) = "" Then
                            Vsf.TextMatrix(mlngLoop, mCol.������ + intCol * mintColCount) = ""
                        End If
    
                        '����ȱʡ�Ľ����־
                        Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(mlngLoop, mCol.������ + intCol * mintColCount)), Trim(Vsf.TextMatrix(mlngLoop, mCol.����ο� + intCol * mintColCount)), Val(Vsf.TextMatrix(mlngLoop, mCol.������� + intCol * mintColCount)), _
                            Vsf.TextMatrix(mlngLoop, mCol.�������� + intCol * mintColCount), Vsf.TextMatrix(mlngLoop, mCol.�������� + intCol * mintColCount))
                
                        '���ݽ��Ӧ����ɫ��־
                        Call ApplyResultColor(Vsf, mlngLoop, mCol.������ + intCol * mintColCount, _
                            Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־ + intCol * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
                    End If
                Next
            Next
        Case mCol.������Ŀ + intColCount * mintColCount
            strItemIDs = GetLabItems(Vsf, mstrType, Vsf.TextMatrix(Row, Col)): gintSelectFocus = 3: ' lvwSelect.SetFocus
'            vsf.SetFocus
            If strItemIDs <> "" Then
                Call AddItems(strItemIDs): Vsf.EditMode(mCol.������ + intColCount * mintColCount) = 1
                Vsf.Col = mCol.������ + intColCount * mintColCount: Vsf.ShowCell Vsf.Row, Vsf.Col
                ShowValue 1, Val(Vsf.TextMatrix(Row, mCol.������� + intColCount * mintColCount)), Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)
            Else
                Vsf.TextMatrix(Row, Col) = ""
            End If
    End Select

    mblnChangeEdit = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intColCount As Integer
    Dim lngItemID As Long
    
    On Error GoTo errH
     
    
    intColCount = GetColCount(NewCol)
    lngItemID = Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount))
    frmLisStationWriteInfo.SelectItem lngItemID
    If lngItemID = 0 Then
        Vsf.Col = mCol.������Ŀ + intColCount * mintColCount
        Exit Sub
    End If
    If OldRow = NewRow Then Exit Sub
    
    If mblnEdit Then
        ShowValue 1, Val(Vsf.TextMatrix(NewRow, mCol.������� + intColCount * mintColCount)), Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub vsf_BeforeComboList(ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    '1-������2-ƫ�͡�3-ƫ�ߡ�4-����
    '1:����,2:���֣�3��������(+-)
    Dim intCol As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(NewCol)
    
    If NewCol = mCol.�����־ + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(Vsf.Row, mCol.������� + intColCount * mintColCount))
            Case 1  '����
                ComboList = "1-����|2-ƫ��|3-ƫ��"
            Case 2  '����
                ComboList = "1-����|4-�쳣"
            Case 3  '�붨��
                ComboList = "1-����|2-ƫ��|3-ƫ��|4-�쳣"
        End Select
    ElseIf NewCol = mCol.������ + intColCount * mintColCount Then
        ComboList = "" '"|-|+|--|++|+-"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim intColCount As Integer
    Dim lng�걾ID As Long
    Dim lng��ĿID As Long
    Dim strsql As String
    Dim rsTmp As New ADODB.Recordset
    If Not mblnEdit Then Cancel = True: Exit Sub
    intColCount = GetColCount(Col)
    lng��ĿID = Val(Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount))
    lng�걾ID = Val(Vsf.TextMatrix(Row, mCol.�걾ID + intColCount * mintColCount))
    
    strsql = "Select Distinct C.������Ŀid" & vbNewLine & _
            "From ����걾��¼ A, ������Ŀ�ֲ� B, ���鱨����Ŀ C, ����ҽ����¼ D" & vbNewLine & _
            "Where A.Id = B.�걾id And B.ҽ��id = D.���id And D.������Ŀid = C.������Ŀid And A.Id = [1] And C.������Ŀid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng�걾ID, lng��ĿID)
    
    If rsTmp.EOF = False Then Cancel = True
    
'    If Val(vsf.TextMatrix(Row, mCol.�̶���Ŀ + intColCount * mintColCount)) = 1 Then Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intColCount As Integer
    If Not mblnEdit Then Cancel = True: Exit Sub
    intColCount = GetColCount(Col)
    If Val(Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)) = 0 Then Cancel = True
End Sub

Private Sub Vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim intCol As Integer
    Dim intColCount As Integer
    If Not mblnEdit Then Exit Sub
    '����������͵�
    On Error Resume Next
    
    intColCount = GetColCount(NewCol)
    
    
    If Not mblnEdit Then Exit Sub
    
    If OldCol <> NewCol And mblnEdit Then
        Vsf.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.������ + intColCount * mintColCount
                Vsf.EditMode(NewCol) = 1
            Case mCol.od + intColCount * mintColCount
                Vsf.EditMode(NewCol) = 1
            Case mCol.�����־ + intColCount * mintColCount
                If Vsf.TextMatrix(NewRow, NewCol) <> "����" And Vsf.TextMatrix(NewRow, NewCol) <> "����" Then
                    Vsf.EditMode(NewCol) = 1
                Else
                    Vsf.EditMode(mCol.������ + intColCount * mintColCount) = 1
                End If
            Case mCol.od + intColCount * mintColCount, mCol.CUTOFF + intColCount * mintColCount, mCol.COV + intColCount * mintColCount
                '����ø��ļ���ֵ�޸�
                Vsf.EditMode(NewCol) = 1
            Case Else
                Vsf.EditMode(mCol.������ + intColCount * mintColCount) = 1
        End Select
    End If
    
    If NewCol = mCol.������ + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(NewRow, mCol.������� + intColCount * mintColCount))
            Case 3
                Vsf.ComboList(mCol.������ + intColCount * mintColCount) = " "
                Vsf.VsfComboList = "" '"|-|+|--|++|+-"
'                If Len(Trim(vsf.TextMatrix(NewRow, mCol.������ + intColCount * mintColCount))) = 0 Then vsf.TextMatrix(NewRow, mCol.������ + intColCount * mintColCount) = "-"
            Case Else
                Vsf.ComboList(mCol.������ + intColCount * mintColCount) = ""
                Vsf.VsfComboList = ""
        End Select
    ElseIf NewCol = mCol.������Ŀ + intColCount * mintColCount Then
        If Val(Vsf.Cell(flexcpData, NewRow, intColCount * mintColCount, NewRow, intColCount * mintColCount)) = 0 Then
            Vsf.EditMode(mCol.������Ŀ + intColCount * mintColCount) = 1
            Vsf.ComboList(mCol.������Ŀ + intColCount * mintColCount) = "..."
            Vsf.VsfComboList = "..."
        Else
            Vsf.EditMode(mCol.������Ŀ + intColCount * mintColCount) = 0
            Vsf.ComboList(mCol.������Ŀ + intColCount * mintColCount) = ""
            Vsf.VsfComboList = ""
        End If
    ElseIf NewCol = mCol.�����־ + intColCount * mintColCount Then
        Vsf.ComboList(NewCol) = " "
        
        Select Case Val(Vsf.TextMatrix(NewRow, mCol.������� + intColCount * mintColCount))
            Case 1  '����
                Vsf.VsfComboList = "1-����|2-ƫ��|3-ƫ��"
            Case 2  '����
                Vsf.VsfComboList = "1-����|4-�쳣"
            Case 3  '�붨��
                Vsf.VsfComboList = "1-����|2-ƫ��|3-ƫ��|4-�쳣"
        End Select
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strItemIDs As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)
    Select Case Col
        Case mCol.������Ŀ + intColCount * mintColCount
            strItemIDs = GetLabItems(Vsf, mstrType): gintSelectFocus = 3: lvwSelect.SetFocus
            Vsf.SetFocus
            If strItemIDs <> "" Then
                Call AddItems(strItemIDs): Vsf.EditMode(mCol.������ + intColCount * mintColCount) = 1: Vsf.Col = mCol.������ + intColCount * mintColCount
                ShowValue 1, Val(Vsf.TextMatrix(Row, mCol.������� + intColCount * mintColCount)), Vsf.Cell(flexcpData, Row, intColCount * mintColCount, Row, intColCount * mintColCount)
                mblnChangeEdit = True
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_GotFocus()
    Dim intColCount As Integer
    intColCount = GetColCount(Me.Vsf.Col)
    If mblnEdit Then
        ShowValue 1, Val(Vsf.TextMatrix(Vsf.Row, mCol.������� + intColCount * mintColCount)), Vsf.Cell(flexcpData, Vsf.Row, intColCount * mintColCount, Vsf.Row, intColCount * mintColCount)
    End If
End Sub

Private Sub Vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel  As Boolean
    Dim intColCount As Integer
    Dim intCol As Integer, intRow As Integer, intTemp As Integer
    Dim intItem As Integer
    Dim lngLoop As Long
    Dim intStart As Integer
    Dim lngColor As Long, lngForeColor As Long
    
    On Error GoTo ErrHand
    
    If KeyCode = vbKeyDelete Then
        If Shift = 0 And Vsf.Body.Editable <> flexEDNone Then
            'ɾ�����м�����
            KeyCode = 0
            blnCancel = False
            
            Call vsf_BeforeDeleteRow(Vsf.Row, Vsf.Col, blnCancel)
            
            If blnCancel Then Exit Sub
            
            intColCount = GetColCount(Vsf.Cols)
            
            If intColCount = 0 Then
                '���д���
                If Vsf.Rows > 1 Then
                    If Vsf.Rows = 2 And Vsf.Row = 1 Then
                        For lngLoop = 0 To Vsf.Cols - 1
                            Vsf.TextMatrix(1, lngLoop) = ""
                        Next
                        Vsf.RowData(1) = ""
                    Else
                        Vsf.RemoveItem Vsf.Row
                    End If
                    Call vsf_AfterDeleteRow(Vsf.Row, Vsf.Col)
                End If
            Else
                '���д���
                If intColCount = 0 Then intColCount = 1
                intTemp = GetColCount(Vsf.Col)
                intStart = Vsf.Row
                For intCol = intTemp To intColCount - 1
                    For intRow = intStart To Vsf.Rows - 1
                        With Vsf
                            lngColor = &H80000005
                            lngForeColor = COLOR.Ĭ��ǰ��ɫ
                            If intRow < Vsf.Rows - 1 Then
                                .Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = _
                                .Cell(flexcpData, intRow + 1, intCol * mintColCount, intRow + 1, intCol * mintColCount)
                                Vsf.Cell(flexcpBackColor, intRow, mCol.������ + intCol * mintColCount, intRow, mCol.������ + intCol * mintColCount) = lngColor
                                Vsf.Cell(flexcpForeColor, intRow, mCol.������ + intCol * mintColCount, intRow, mCol.������ + intCol * mintColCount) = lngForeColor
                                For intItem = 1 To 20
                                    .TextMatrix(intRow, intItem + intCol * mintColCount) = .TextMatrix(intRow + 1, intItem + intCol * mintColCount)
                                Next
                            Else
                                '������ʾ��
                                If intCol + 1 <= intColCount - 1 Then
                                    .Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = _
                                    .Cell(flexcpData, 1, (intCol + 1) * mintColCount, 1, (intCol + 1) * mintColCount)
                                    Vsf.Cell(flexcpBackColor, intRow, mCol.������ + intCol * mintColCount, intRow, mCol.������ + intCol * mintColCount) = lngColor
                                    Vsf.Cell(flexcpForeColor, intRow, mCol.������ + intCol * mintColCount, intRow, mCol.������ + intCol * mintColCount) = lngForeColor
                                    For intItem = 1 To 20
                                        .TextMatrix(intRow, intItem + intCol * mintColCount) = .TextMatrix(1, intItem + (intCol + 1) * mintColCount)
                                    Next
                                    intStart = 1
                                End If
                            End If
                        End With
                    Next
                Next
                lngLoop = 0
                For intCol = 0 To intColCount - 1
                    For intRow = 1 To Vsf.Rows - 1
                        If Val(Vsf.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) <> 0 Then
                            lngLoop = lngLoop + 1
                            Vsf.TextMatrix(intRow, intCol * mintColCount) = lngLoop
                            Vsf.Cell(flexcpBackColor, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = &HFDD6C6
                        Else
                            Vsf.TextMatrix(intRow, intCol * mintColCount) = ""
                            Vsf.Cell(flexcpBackColor, intRow, intCol * mintColCount, intRow, intCol * mintColCount) = &H80000005
                        End If
                    Next
                Next
            End If
        End If
    End If
    
    Exit Sub
    
ErrHand:
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strSvrText As String, strItemIDs As String
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If InStr(Vsf.EditText, "'") > 0 Then
            Cancel = True
            Exit Sub
        End If

        Select Case Col
            Case mCol.������Ŀ + intColCount * mintColCount
                Vsf.Cell(flexcpData, Row, Col, Row, Col) = Vsf.EditText
                Call vsf_AfterEdit(Row, Col)
            Case mCol.������ + intColCount * mintColCount
                If Row = Vsf.Rows - 1 And Col + mintColCount <= Vsf.Cols Then
                    Vsf.Row = 1
                    Vsf.Col = mCol.������ + (intColCount + 1) * mintColCount
                Else
                    If Row = Vsf.Rows - 1 Then
                        Vsf.Rows = Vsf.Rows + 1
                        Vsf.Row = Vsf.Row + 1
                        Vsf.Col = mCol.������Ŀ + intColCount * mintColCount
                        Vsf.ShowCell Vsf.Row, Vsf.Col
                    Else
                        Vsf.Row = Vsf.Row + 1
                        Vsf.Col = mCol.������ + intColCount * mintColCount
                    End If
                    
                End If
            Case mCol.od + intColCount * mintColCount
                Vsf.SetFocus
        End Select
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    Dim intColCount As Integer
    
    On Error GoTo errH
    intColCount = GetColCount(Col)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Row < Vsf.Rows - 1 Then
            Vsf.Row = Vsf.Row + 1
        ElseIf Col + mintColCount <= Vsf.Cols Then
            Vsf.Row = 1
            Vsf.Col = Vsf.Col + mintColCount
        Else
            Vsf.Rows = Vsf.Rows + 1
            Vsf.Row = Vsf.Row + 1
            Vsf.Col = mCol.������Ŀ + intColCount * mintColCount
            Vsf.ShowCell Vsf.Row, Vsf.Col
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim intColCount As Integer
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    If Chr(KeyAscii) = "'" Then KeyAscii = 0: Exit Sub
    intColCount = GetColCount(Col)
    If Col = mCol.������ + intColCount * mintColCount Then
        Select Case Val(Vsf.TextMatrix(Vsf.Row, mCol.������� + intColCount * mintColCount))
            Case 1
                KeyAscii = FilterKeyAscii(KeyAscii, 2)
        End Select
        mblnChangeEdit = True
    End If
End Sub
'ͨ�������ȡ���鱸ע
Private Function GetComment(ByVal strCode As String, ByVal strTYPE As String)
    Dim rsTmp As ADODB.Recordset
    Dim objPoint As POINTAPI, mstrSQL As String
    Dim sglX As Single, sglY As Single
    
    mstrSQL = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ���� FROM ���鱸ע���� A " & _
        "WHERE (Instr(A.����,[1])>0 Or Instr(A.����,[1])>0) And (A.���� Is Null Or A.����=[2])"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, UCase(strCode), mstrType)
    If rsTmp.EOF Then
        GetComment = strCode
    Else
        If rsTmp.RecordCount = 1 Then
            GetComment = Nvl(rsTmp("����"))
        Else
            Call ClientToScreen(txtComment.hWnd, objPoint)
    
            sglX = objPoint.X * 15 - 30
            sglY = objPoint.Y * 15 - 2000
            If frmSelectList.ShowSelect(Me, rsTmp, "����,800,0,0;����,1500,0,0;����,2500,0,0;����,5500,0,0", sglX, sglY, Me.txtComment.Width, 2000, Me.Name & "\���鱸עѡ��", "��ѡ����鱸ע") Then
                GetComment = Nvl(rsTmp("����"))
            Else
                GetComment = strCode
            End If
        End If
    End If
End Function

Private Sub AddItems(ByVal strItemIDs As String)
'��Ӽ�����Ŀ(����΢������Ŀ)
'strItemIDs��������ĿID�����ԣ��ָ�
    Dim strsql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strsql = "SELECT " & mlngKey & " as �걾ID, B.ID,B.������||DECODE(C.��д,NULL,'','('||C.��д||')') AS ������Ŀ,'' As �ϴν��,''As CV," & _
        "'' As ԭʼ���,Decode(C.�������,3,Nvl(C.Ĭ��ֵ,'-'),2,C.Ĭ��ֵ,'') As ���ν��,C.���㹫ʽ,C.�������," & _
        "'' AS ��־,'' as OD,'' as CUTOFF,'' as COV, '' as ø���ID,c.���챨���� as ���챨��,c.���쾯ʾ�� as ���쾯ʾ, " & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,A.�걾��λ,0,NULL,[1]),' .','0.'),'��.','��0.')) AS �ο�," & _
        "[1] As ����ID,C.�����Χ,0 As �̶���Ŀ,Nvl(E.С��λ��,2) As С��,C.��������,C.��������,C.��λ,'' as �ϴ�ʱ��,a.id as ������ĿID,'' as ������� " & _
        ",Zl_To_Number(Zl_Get_Reference(1, b.id, A.�걾��λ, 0, Null,[1])) as �ο�ID " & vbNewLine & _
        "FROM ������ĿĿ¼ A,���鱨����Ŀ D,����������Ŀ B,������Ŀ C,����������Ŀ E " & _
        "WHERE A.ID = D.������ĿID And D.������ĿID=B.ID " & _
                    "AND B.ID = C.������ĿID And D.������ĿID=E.��ĿID(+) And E.����ID(+)=[1] " & _
                    "AND D.ϸ��ID IS NULL AND C.��Ŀ���<>2 " & _
                    "AND A.ID In (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)))  "
                    
    strsql = "Select a.�걾id,a.id,a.������Ŀ,a.�ϴν��,a.cv,a.ԭʼ���,a.���ν��,a.���㹫ʽ,a.�������,a.��־,a.od,a.cutoff,a.cov,a.ø���id" & _
           ",a.���챨��,a.���쾯ʾ,a.�ο�,a.����id,a.�����Χ,a.�̶���Ŀ,a.С��,f.��ʾ���� as ��������,f.��ʾ���� as ��������,a.��λ,a.�ϴ�ʱ��,a.������Ŀid,a.�������,null as ������ʾ,null as ������˱�ʶ" & _
           " From (" & strsql & ") a,������Ŀ�ο� F Where a.�ο�id=f.id(+) Order By A.ID,a.�������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mDeviceID, strItemIDs)
    
    If Not rsTmp.EOF Then
        Vsf.TextMatrix(0, 0) = "#"
'        Call FillGrid_UQ(vsf, rsTmp, Array("", "", "", ""), False)
        Call ReadVsf(rsTmp, Array("", "", "", ""), False)
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LoadModel(ByVal strCode As String) As Boolean
'���뱨��ģ��(����΢������Ŀ)
'strCode��ģ���������
    Dim strsql As String, rsTmp As ADODB.Recordset
    Dim lngCurrRow As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    On Error GoTo errH
    
    LoadModel = False
    strsql = "SELECT B.ID,B.������||DECODE(C.��д,NULL,'','('||C.��д||')') AS ������Ŀ,'' As �ϴν��,''As CV," & _
        "'' As ԭʼ���,A.������ As ���ν��,C.���㹫ʽ,C.�������," & _
        "'' AS ��־," & _
        "Trim(REPLACE(REPLACE(' '||zlGetReference(B.ID,'',0,NULL,''),' .','0.'),'��.','��0.')) AS �ο�," & _
        "[2] As ����ID,C.�����Χ,0 As �̶���Ŀ,2 As С��,C.��������,C.��������,C.��λ " & _
        ",zl_Get_Reference(1,B.ID,'',0,NULL,'') as �ο�id " & _
        "FROM ����ģ������ A,����������Ŀ B,������Ŀ C,����ģ��Ŀ¼ D " & _
        "WHERE A.��ĿID=B.ID AND B.ID = C.������ĿID And D.ID=A.ģ��ID " & _
                    "AND A.ϸ��ID IS NULL AND (D.����=[1] Or D.����=[1])"
    strsql = "Select a.ID,a.������Ŀ,a.�ϴν��,a.ԭʼ���,a.���ν��,a.���㹫ʽ,a.�������,a.��־,a.�ο�,a.����id" & _
            ",a.�����Χ,a.�̶���Ŀ,a.С��,f.��ʾ���� as ��������,f.��ʾ���� as ��������,a.��λ" & _
            " From (" & strsql & ") a,������Ŀ�ο� F where a.�ο�id=F.ID(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, strCode, mDeviceID)
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsTmp("ID"))))
            If lngCurrRow > 0 Then
                intColCount = GetColCount(Vsf.Cols)
                If intColCount = 0 Then intColCount = 1
                For intCol = 0 To intColCount - 1
                    If Val(Vsf.Cell(flexcpData, lngCurrRow, intCol * mintColCount, lngCurrRow, intCol * mintColCount)) = Nvl(rsTmp("ID")) Then
                        Exit For
                    End If
                Next
                Vsf.TextMatrix(lngCurrRow, mCol.������ + intCol * mintColCount) = Nvl(rsTmp("���ν��"))
                Vsf.TextMatrix(lngCurrRow, mCol.����ο� + intCol * mintColCount) = Nvl(rsTmp("�ο�"))
                '����ȱʡ�Ľ����־
                Vsf.TextMatrix(lngCurrRow, mCol.�����־ + intCol * mintColCount) = CalcDefaultFlag(Trim(Vsf.TextMatrix(lngCurrRow, mCol.������ + intCol * mintColCount)), _
                    Trim(Vsf.TextMatrix(lngCurrRow, mCol.����ο� + intCol * mintColCount)), Val(Vsf.TextMatrix(lngCurrRow, mCol.������� + intCol * mintColCount)), _
                    Vsf.TextMatrix(lngCurrRow, mCol.�������� + intCol * mintColCount), Vsf.TextMatrix(lngCurrRow, mCol.�������� + intCol * mintColCount))
                
                '���ݽ��Ӧ����ɫ��־
                Call ApplyResultColor(Vsf, lngCurrRow, mCol.������ + intCol * mintColCount, _
                    Decode(Vsf.TextMatrix(lngCurrRow, mCol.�����־ + intCol * mintColCount), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
            End If
            
            rsTmp.MoveNext
        Loop
        
        LoadModel = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadLastValue()
    '���ܣ�װ���ϴν������
    Dim lngDates As Long, lngTimes As Long
    Dim strRows As String, aryRows() As String
    Dim strCols As String, aryCols() As String
    Dim dblCurCV As Double     '�����CV
    Dim lngDays As Long, mstrEndTime As String, rsTemp As ADODB.Recordset
    Dim lngRow As Long, lngCurrKey As Long
    Dim intFindMode As Integer          '���˲���ģʽ
    Dim intColCount As Integer, intCol As Integer
    Dim dblCalc As Double
    Dim strTag As String                '��ʾ��ʶ
    Dim dbl���챨�� As Double, dbl���쾯ʾ As Double
    Dim intSampleType As Integer
    
    If mlngKey = 0 Then Exit Sub
    
    Err = 0: On Error GoTo ErrHand
    
    intFindMode = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
    intSampleType = zlDatabase.GetPara("�ϴν�������ձ걾����", 100, 1208, 0)
    
    '��õ�ǰ�����ʱ�䡢��ĿҪ��ĸ���������ȡ��Ŀ�����ģ�
    gstrSql = "Select Nvl(L.����ʱ��, Sysdate) As ����ʱ��, Nvl(Max(��������), 0) As ����" & vbNewLine & _
            "From ������Ŀѡ�� O, ���鱨����Ŀ X, ������ͨ��� R, ����걾��¼ L" & vbNewLine & _
            "Where O.������Ŀid(+) = X.������Ŀid And X.������Ŀid = R.������Ŀid And R.����걾id = L.ID And L.ID = [1]" & vbNewLine & _
            "Group By Nvl(L.����ʱ��, Sysdate)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    If rsTemp.RecordCount > 0 Then
        lngDays = rsTemp!����
        mstrEndTime = Format(rsTemp!����ʱ��, "yyyy-MM-dd hh:mm:ss")
    Else
        lngDays = 30
        mstrEndTime = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    End If
    lngDays = IIf(lngDays = 0, 30, lngDays)
    lngDates = lngDays
    
    '��ѯ��������װ�룺
    gstrSql = "Select L.������Ŀid As ID, V.��д As Ӣ����, L.����, L.����ʱ��, L.������, V.���챨����" & vbNewLine & _
            "From (Select L.������Ŀid, L.����, L.����ʱ��, L.������" & vbNewLine & _
            "       From (Select L.����id,L.����,L.�Ա�,L.����, L.ID As ����, L.����ʱ��, R.������Ŀid, R.������,L.�걾���� " & vbNewLine & _
            "              From ����걾��¼ L, ������ͨ��� R" & vbNewLine & _
            "              Where L.ID = R.����걾id AND L.������=R.��¼���� AND " & vbNewLine & _
            "                    L.����ʱ��  Between To_Date([2], 'yyyy-mm-dd hh24:mi:ss') - [3] And" & vbNewLine & _
            "                    To_Date([2], 'yyyy-mm-dd hh24:mi:ss') And L.ID<>[1] And " & vbNewLine & _
            "                    " & IIf(intFindMode = 0, " L.����id = [4] ", " L.����ID in (select ����id from ������Ϣ where ���� = [5] )") & ") L," & vbNewLine & _
            "            (Select L.����id,L.����,L.�Ա�,L.����,R.������Ŀid,L.�걾���� " & vbNewLine & _
            "              FROM ����걾��¼ L,������ͨ��� R" & vbNewLine & _
            "              WHERE L.ID = [1] AND L.ID = R.����걾id AND L.������=R.��¼����) C" & vbNewLine & _
            "       Where " & IIf(intFindMode = 0, " L.����id = C.����id ", " L.���� = C.���� ") & vbNewLine & _
            "             AND L.������Ŀid+0 =C.������Ŀid" & IIf(intSampleType = 0, " And nvl(L.�걾����,'') = nvl(C.�걾����,'')", "") & ") L, ������Ŀ V " & vbNewLine & _
            "Where L.������Ŀid = V.������Ŀid" & vbNewLine & _
            "Order By L.���� Desc"
'    gstrSql = "Select /*+ rule */" & vbNewLine & _
                " L.������Ŀid As ID, V.��д As Ӣ����, L.����, L.����ʱ��, L.������, V.���챨����" & vbNewLine & _
                " From (Select L.����id, L.����, L.�Ա�, L.����, L.ID As ����, L.����ʱ��, R.������Ŀid, R.������" & vbNewLine & _
                "       From ����걾��¼ L, ������ͨ��� R" & vbNewLine & _
                "       Where L.ID = R.����걾id And L.������ = R.��¼���� And" & vbNewLine & _
                "             L.����ʱ�� Between To_Date([2], 'yyyy-mm-dd hh24:mi:ss') - [3] And " & vbNewLine & _
                "             To_Date([2], 'yyyy-mm-dd hh24:mi:ss') And L.ID<>[1] " & vbNewLine & _
                "             And " & IIf(intFindMode = 0, " L.����id = [4] ", " L.����=[5] ") & ") L, " & vbNewLine & _
                "       ������Ŀ V" & vbNewLine & _
                " Where L.������Ŀid = V.������Ŀid" & vbNewLine & _
                " Order By L.���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, mstrEndTime, lngDates, mLngPatientID, mstrPatientName)
    
    
    Err = 0: On Error GoTo 0
    With Me.Vsf
        lngCurrKey = 0: lngTimes = 0
        Do While Not rsTemp.EOF
            If lngCurrKey <> rsTemp("����") Then
                'ֻ����һ�ν��
                If lngTimes = 1 Then Exit Do
                lngTimes = lngTimes + 1
                lngCurrKey = rsTemp("����")
            End If
            lngRow = FindRepeatLine(Vsf, rsTemp("ID"))
            If lngRow > 0 Then
                intColCount = GetColCount(Vsf.Cols)
                If intColCount = 0 Then intColCount = 1
                For intCol = 0 To intColCount - 1
                    If Val(.Cell(flexcpData, lngRow, intCol * mintColCount, lngRow, intCol * mintColCount)) = Nvl(rsTemp("ID")) Then
                        .TextMatrix(lngRow, mCol.�ϴν�� + intCol * mintColCount) = Nvl(rsTemp("������"))
                        If Nvl(rsTemp("������")) <> "" Then
                            .TextMatrix(lngRow, mCol.�ϴ�ʱ�� + intCol * mintColCount) = Format(Nvl(rsTemp("����ʱ��")), "YYYY-MM-DD")
                            dblCalc = 0
                            If Val(.TextMatrix(lngRow, mCol.������ + intCol * mintColCount)) <> 0 Then
                                dblCalc = (Val(Nvl(rsTemp("������"))) - Val(.TextMatrix(lngRow, mCol.������ + intCol * mintColCount))) / _
                                Val(.TextMatrix(lngRow, mCol.������ + intCol * mintColCount)) * 100
                                dblCalc = Format(dblCalc, "00#.##")
                            End If
                            strTag = ""
                            
                            dbl���챨�� = Val(.TextMatrix(lngRow, mCol.���챨�� + intCol * mintColCount))
                            dbl���쾯ʾ = Val(.TextMatrix(lngRow, mCol.���쾯ʾ + intCol * mintColCount))
                            
                            If dblCalc > 0 Then
                                If dblCalc >= dbl���쾯ʾ And dbl���쾯ʾ <> 0 Then
                                    strTag = "����"
                                ElseIf dblCalc >= 10 And dbl���챨�� <> 0 Then
                                    strTag = "��"
                                End If
                            Else
                                If Abs(dblCalc) >= dbl���쾯ʾ And dbl���쾯ʾ <> 0 Then
                                    strTag = "����"
                                ElseIf Abs(dblCalc) >= dbl���챨�� And dbl���챨�� <> 0 Then
                                    strTag = "��"
                                End If
                            End If
                            .TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount) = Replace(.TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount), "��", "")
                            .TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount) = Replace(.TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount), "��", "")
                            .TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount) = .TextMatrix(lngRow, mCol.������Ŀ + intCol * mintColCount) & strTag
                            Select Case strTag
'                                Case "��"
'                                    .Cell(flexcpForeColor, lngRow, mCol.������Ŀ + intCol * mintColCount, lngRow, mCol.������Ŀ + intCol * mintColCount) = COLOR.���걳��ɫ
'                                Case "��"
'                                    .Cell(flexcpForeColor, lngRow, mCol.������Ŀ + intCol * mintColCount, lngRow, mCol.������Ŀ + intCol * mintColCount) = COLOR.�ͱ걳��ɫ + 300
                                Case "����", "����"
                                    .Cell(flexcpForeColor, lngRow, mCol.������Ŀ + intCol * mintColCount, lngRow, mCol.������Ŀ + intCol * mintColCount) = COLOR.��ɫ
                            End Select
                        End If
                        
                    End If
                Next
            End If
        
            rsTemp.MoveNext
        Loop
        
'        '�����ʼ�����д�ͱ���ɫ����
'        For lngRow = .FixedRows To .Rows - 1
'            .TextMatrix(lngRow, mCol.��λ + 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, mCol.��λ + 1)), " .", "0."), " ", "")
'            For lngCol = mCol.��λ + 4 To .Cols - 1 Step 2
'                .TextMatrix(lngRow, lngCol - 1) = Replace(Replace(" " & Trim(.TextMatrix(lngRow, lngCol - 1)), " .", "0."), " ", "")
'                If Val(.TextMatrix(lngRow, lngCol - 1)) = 0 Or Val(.TextMatrix(lngRow, mCol.��λ + 1)) = 0 Then
'                    dblCurCV = 0
'                Else
'                    dblCurCV = (Val(.TextMatrix(lngRow, lngCol - 1)) - Val(.TextMatrix(lngRow, mCol.��λ + 1))) / Val(.TextMatrix(lngRow, mCol.��λ + 1)) * 100
'                End If
'                .TextMatrix(lngRow, lngCol) = Format(dblCurCV, "0.00;-0.00; ; ")
'                If Val(.TextMatrix(lngRow, mCol.������)) <> 0 And Abs(dblCurCV) > Val(.TextMatrix(lngRow, mCol.������)) Then
'                    .Cell(flexcpBackColor, lngRow, lngCol) = RGB(248, 194, 169)
'                End If
'            Next
'        Next
'        .Redraw = flexRDDirect
    End With
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    
    Call FormatVsfCell(Vsf, mCol.�ϴν��, "0.0######", 0, IIf(mDeviceID > 0, mCol.С��, -1))
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim rsTmp As New ADODB.Recordset
    
    If Button <> vbRightButton Then Exit Sub
    
    On Error GoTo errH
    
    gstrSql = "select distinct nvl(��¼����,0) as ��¼���� from ������ͨ��� where ����걾id = [1] order by nvl(��¼����,0) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngKey)
    Set cbrPopupBar = Me.cbrthis.Add("�����˵�", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, 100, "�ٴ�����")
    If rsTmp.RecordCount > 1 Then
        Do Until rsTmp.EOF
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, rsTmp(0) + 1, "ѡ���" & rsTmp(0) + 1 & "��")
            rsTmp.MoveNext
        Loop
'        cbrPopupBar.ShowPopup
    End If
    cbrPopupBar.ShowPopup
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then
        RaiseEvent StartEdit(Cancel)
        If mblnPatientFind = True Then Cancel = True
        If Cancel = False Then mblnEvent = True
    End If
    
End Sub


Public Sub Resize()
    '�����������
    Call Form_Resize
End Sub

Private Function ReadVsf(ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long, lngCurrRow As Long
    Dim strOldValue As String, strNewValue As String
    Dim intColCount  As Integer
    Dim intCol As Integer, intRow As Integer
    Dim lngHeight As Long
    Dim blnShowType As Boolean
    Dim str���� As String
    blnShowType = zlDatabase.GetPara("����Ӧ��ʾ���", 100, 1208, False)
    If fraComment.Tag <> "" Then blnShowType = True
    
    If blnClear Then
        Vsf.Rows = 2
        Vsf.RowData(1) = 0
        For lngLoop = 0 To Vsf.Cols - 1
            Vsf.TextMatrix(1, lngLoop) = ""
            Vsf.Cell(flexcpData, 1, lngLoop, 1, lngLoop) = ""
        Next
        lngRow = 0
        Vsf.Cols = mintColCount
    Else
        'Ԥ����һ����
        With Vsf
            intColCount = GetColCount(.Cols)
            If intColCount = 0 Then intColCount = 1
            For intCol = 0 To intColCount - 1
                For intRow = 1 To .Rows - 1
                    If Val(.Cell(flexcpData, intRow, intCol * mintColCount, intRow, intCol * mintColCount)) = 0 Then
                        lngRow = intRow - 1
                        intColCount = intCol
                        Exit For
                    End If
                Next
            Next
        End With
    End If
    
    
    With Vsf.Body
        If .ClientHeight < .CellHeight * 15 Then
            lngHeight = .CellHeight * 15
        Else
            lngHeight = .ClientHeight
        End If
    End With
    Do While Not rsData.EOF
        lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsData("ID"))))
'        lngCurrRow = -1
        If lngCurrRow = -1 Then
            With Vsf.Body

                If (.CellHeight + 15) * (lngRow + 2) > lngHeight And blnShowType = True Then
                    intColCount = intColCount + 1
                    lngRow = 1
                    With Vsf
                        .NewColumn "#", 300, 7
                        .NewColumn "������Ŀ", 2100, 1
                        .NewColumn "ԭʼ���", 0, 1
                        .NewColumn "���ν��", 1200, 1, , 1
                        .NewColumn "��λ", 1000, 1
                        .NewColumn "CV", 0, 1
                        .NewColumn "��־", 450, 1
                        .NewColumn "�ϴν��", 0, 1
                        .NewColumn "�ϴ�ʱ��", 0, 1
                        .NewColumn "�ο�", 1300, 1
                        .NewColumn "�������", 0, 1
                        .NewColumn "����id", 0, 1
                        .NewColumn "���㹫ʽ", 0, 1
                        .NewColumn "�����Χ", 0, 1
                        .NewColumn "�̶���Ŀ", 0, 1
                        .NewColumn "С��", 0, 1
                        .NewColumn "��������", 0, 1
                        .NewColumn "��������", 0, 1
                        .NewColumn "������ĿID", 0, 1
                        .NewColumn "�������", 0, 1
                        .NewColumn "�걾ID", 0, 1
                        .NewColumn "OD", 700, 1, , 1
                        .NewColumn "CUTOFF", 700, 1
                        .NewColumn "COV", 700, 1
                        .NewColumn "ø���ID", 0, 1
                        .NewColumn "���챨��", 0, 1
                        .NewColumn "���쾯ʾ", 0, 1
                        .NewColumn "������ʾ", 1000, 1
                        .NewColumn "������˱�ʶ", 1200, 1
                    End With
                Else
                    lngRow = lngRow + 1
                End If
            End With
            
            lngCurrRow = lngRow
        
            If Vsf.Rows < lngRow + 1 Then Vsf.Rows = lngRow + 1
            
            On Error Resume Next
'            Vsf.RowData(lngCurrRow) = CStr(zlCommFun.Nvl(rsData("ID")))
            Vsf.Cell(flexcpData, lngCurrRow, intColCount * mintColCount, lngCurrRow, intColCount * mintColCount) = CStr(Nvl(rsData("ID")))
            
            On Error GoTo ErrHand
            
            str���� = Get������(Val("" & rsData("�걾ID")), Val("" & rsData("ID")), "" & rsData("���ν��"))
            
            For lngLoop = 0 To mintColCount - 1
                intCol = intColCount * mintColCount + lngLoop
                
                If Trim(Vsf.TextMatrix(0, intCol)) <> "" Then
                    If Vsf.TextMatrix(0, intCol) = "#" Then
                        Vsf.TextMatrix(lngCurrRow, intCol) = IIf(intColCount > 0, intColCount * (Vsf.Body.Rows - 1) + lngCurrRow, lngCurrRow)
                        Vsf.Cell(flexcpBackColor, lngCurrRow, intCol, lngCurrRow, intCol) = &HFDD6C6
                    Else
                        On Error Resume Next
                        strMask = ""
                        strMask = MaskArray(intCol)
                                                
                        On Error GoTo ErrHand
                        
                        If strMask <> "" Then
                            strNewValue = Format(zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol))), strMask)
                        Else
                            strNewValue = zlCommFun.Nvl(rsData(Vsf.TextMatrix(0, intCol)))
                        End If
                        If str���� <> "" Then
                            If rsData(Vsf.TextMatrix(0, intCol)).Name = "������Ŀ" Then
                                strNewValue = strNewValue & " " & str����
                                Vsf.Cell(flexcpForeColor, lngCurrRow, intCol, lngCurrRow, intCol) = COLOR.��ɫ
                            End If
                        End If
                        Vsf.TextMatrix(lngCurrRow, intCol) = strNewValue
                    End If
                End If
                
            Next
        End If
        
        rsData.MoveNext
    Loop
'    Call chkOriginal_Click: Call chkLast_Click: Call chkSign_Click
'    Call chkUnit_Click: Call chkReferrence_Click: Call chkMB_Click
    intColCount = GetColCount(Vsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        Vsf.Body.ColWidth(mCol.������Ŀ + intCol * mintColCount) = IIf(chkChina.Value, 2100, 1000)
        Vsf.Body.ColWidth(mCol.ԭʼ��� + intCol * mintColCount) = IIf(chkOriginal.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.�ϴν�� + intCol * mintColCount) = IIf(chkLast.Value, 900, 0)
        Vsf.Body.ColWidth(mCol.�ϴ�ʱ�� + intCol * mintColCount) = IIf(chkLast.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.�����־ + intCol * mintColCount) = IIf(chkSign.Value, 450, 0)
        Vsf.Body.ColWidth(mCol.��λ + intCol * mintColCount) = IIf(chkUnit.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.����ο� + intCol * mintColCount) = IIf(chkReferrence.Value, 1300, 0)
        Vsf.Body.ColWidth(mCol.od + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.CUTOFF + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.COV + intCol * mintColCount) = IIf(chkMB.Value, 700, 0)
        Vsf.Body.ColWidth(mCol.������ʾ + intCol * mintColCount) = IIf(chkYiQiTiShi.Value, 1000, 0)
        Vsf.Body.ColWidth(mCol.������˱�ʶ + intCol * mintColCount) = IIf(chkYiQiBiaoShi.Value, 1200, 0)
    Next
    
    Exit Function
    
ErrHand:

    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function GetColCount(Col As Long) As Integer
    '����               ���ص�ǰ�ǵڼ���������Ŀ
    '����               ��ǰ����
    Dim dblTmp As Double
    If Col <= mintColCount Then
        GetColCount = 0
    Else
        dblTmp = Col / mintColCount
        If InStr(dblTmp, ".") > 0 Then
            GetColCount = Mid(dblTmp, 1, InStr(dblTmp, ".") - 1)
        Else
            GetColCount = dblTmp
        End If
    End If
End Function
Private Function FindRepeatLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '����:����RowData����strSeekID����
    '����:
    '����:�кŻ�-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer, intCol As Integer
    FindRepeatLine = -1
    intColCount = GetColCount(objMsf.Cols)
    If intColCount = 0 Then intColCount = 1
    For intCol = 0 To intColCount - 1
        For i = 1 To objMsf.Rows - 1
            If Val(Me.Vsf.Cell(flexcpData, i, intCol * mintColCount, i, intCol * mintColCount)) = strSeekID Then
                FindRepeatLine = i
                Exit For
            End If
'            If objMsf.RowData(i) = strSeekID Then Exit For
        Next
    Next
    If i <= objMsf.Rows - 1 Then FindRepeatLine = i
End Function

Private Function Get������(ByVal lng�걾ID As Long, ByVal lng��ĿID As Long, ByVal str������ As String) As String
    Dim str���� As String, strWhere As String
    Dim rsTmp As ADODB.Recordset
    Dim bln�������� As Boolean
    Dim lng�ο�ID As Long
    str���� = ""
    
    If Not IsNumeric(str������) Then
        Exit Function
    End If
    
    gstrSql = "Select Zl_Get_Reference(1, " & lng��ĿID & ", a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����,a.�������iD) As �ο�id" & vbNewLine & _
                "From ����걾��¼ A" & vbNewLine & _
                "Where a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng�걾ID)
    If rsTmp.EOF = True Then Exit Function
    lng�ο�ID = Val("" & rsTmp!�ο�id)
    If lng�ο�ID <> 0 Then
        gstrSql = "Select ��������, �������� From ������Ŀ�ο� A Where a.Id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng�ο�ID)
        Do Until rsTmp.EOF
            If "" & rsTmp!�������� <> "" And "" & rsTmp!�������� <> "" Then
                If Val(str������) < Val("" & rsTmp!��������) Or Val(str������) > Val("" & rsTmp!��������) Then
                    Get������ = "�踴��"
                    Exit Function
                End If
            ElseIf "" & rsTmp!�������� = "" And "" & rsTmp!�������� <> "" Then
                If Val(str������) > Val("" & rsTmp!��������) Then
                    Get������ = "�踴��"
                    Exit Function
                End If
            ElseIf "" & rsTmp!�������� <> "" And "" & rsTmp!�������� = "" Then
                If Val(str������) < Val("" & rsTmp!��������) Then
                    Get������ = "�踴��"
                    Exit Function
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
End Function
