VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����������Ǩ"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "����ʱ���趨(ϵͳ����������)"
      ForeColor       =   &H8000000D&
      Height          =   1665
      Index           =   1
      Left            =   3690
      TabIndex        =   6
      Top             =   240
      Width           =   2955
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101646339
         CurrentDate     =   40540
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ��1 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   690
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101384195
         CurrentDate     =   40540.0833333333
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   1080
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101384195
         CurrentDate     =   40540.1666666667
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������Ǩ��ʼʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   7
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ӡ������ʼʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   9
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������ʱ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   675
         TabIndex        =   11
         Top             =   1140
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ģʽ"
      Height          =   1245
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   3375
      Begin VB.OptionButton opt 
         Caption         =   "��Ǩ��ʷ����(��̨)"
         Height          =   180
         Index           =   2
         Left            =   690
         TabIndex        =   5
         Top             =   900
         Width           =   2295
      End
      Begin VB.OptionButton opt 
         Caption         =   "ϵͳ����"
         Height          =   180
         Index           =   1
         Left            =   690
         TabIndex        =   4
         Top             =   600
         Width           =   1845
      End
      Begin VB.OptionButton opt 
         Caption         =   "����ǰ׼��"
         Height          =   180
         Index           =   0
         Left            =   690
         TabIndex        =   3
         Top             =   300
         Width           =   1845
      End
   End
   Begin VB.Timer tim 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   150
      Top             =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��̨����"
      Height          =   350
      Left            =   5370
      TabIndex        =   14
      Top             =   2070
      Width           =   1100
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2685
   End
   Begin VB.Label Label2 
      Caption         =   "��ѡ��һ�ֵ���ģʽ"
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   3285
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng����ID As Long
Dim strSQL As String
Dim rsTemp As New ADODB.Recordset
Dim rsPati As New ADODB.Recordset
Dim rsFile As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset

'��������˵��:
'����ǰ׼�������������ʷ������Ǩ,����ֻҪ������ͻ���ÿ��12�����賿6��֮�����������Ǩ����ӡ��������,��Ҫ��������,��һ�����������Ժ����(����Ժ����)��ʷ���ݵĵ���,�ڶ��������ʷ���˵����ݵ��빤��


'--��ȡ65txyy���ݿ���������ݽ��к˶�
'--64757,1|661573,1|386712,1|470068,1  ���ĸ����˵Ļ������ݷֱ���>300,>400,>900,>2000
'--select ����ID,סԺ���� from ������Ϣ where סԺ��=124631;
'select * from ������Ϣ where ����ID=661573;
'select * from ������ҳ where ����ID=661573 and ��ҳID=1;
'select count(*) from ���˻����¼ where ����ID=661573 and ��ҳID=1;
'select count(*) from ���˻����¼ A,���˻������� B Where A.����ID=661573 and A.��ҳID=1 And A.ID=B.��¼ID ;
'select count(*) from ���˻����ļ� where ����ID=661573 and ��ҳID=1;
'select A.�ļ�ID,count(*) from ���˻������� A,���˻����ļ� B where A.�ļ�ID=B.ID And B.����ID=661573 and B.��ҳID=1 group by A.�ļ�ID ;
'select A.ID AS �ļ�ID,count(*) from ���˻����ļ� A,���˻������� B,���˻�����ϸ C Where A.����ID=661573 and A.��ҳID=1 And A.ID=B.�ļ�ID And B.ID=C.��¼ID Group by A.ID ;
'select * from ���˻����ļ� where ����ID=661573 and ��ҳID=1;
'select count(*) from ���˻����¼ B,���˻������� C where B.ID=C.��¼ID And B.����ID=661573 And B.��ҳID=1 And B.Ӥ��=0 and ��Ŀ���� IN ('1)����������Ŀ','2)���±����Ŀ');



Private Sub DataUpgrade()
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim i As Integer, j As Integer, intStart As Integer, intEnd As Integer
    Dim datStart As Date, datEnd As Date
    Dim lng�ļ�ID As Long
    Dim lng����ID As Long, lng��ҳID As Long, intӤ�� As Integer, lng��ʽID As Long, lng���� As Long
    Dim str�鵵�� As String, str�鵵ʱ�� As String, strMsg As String
    Dim bln���� As Boolean      '�������µ����ݺ���Ϊ��
    Dim bln���� As Boolean      '�Ƿ�����������(���˱仯���ļ���ʽ�仯ʱ)
    Dim blnFirst As Boolean     '��һ������
    Dim blnError As Boolean     '��������
    Dim blnCommit As Boolean    '�л�����ʱ��Ҫ�ύ����
    Dim blnDataMoved As Boolean '����ת����־
    Dim rsDate As New ADODB.Recordset
    Dim strʱ�� As String
    On Error GoTo errHand
    '�ļ���ʽ��ͬ�Ĳ��ظ���������,�������ݵĿ�ʼ����Ϊ��Ժ����,��������Ϊ��Ժ����
    '�����������ļ��嵥�����Ժ�,�ٲ�����������,����bln������صĴ���ָ�����,���ڵ�ģʽΪ:����Ҳ�������ʱ���Ӧ�Ļ����ļ�,���Ըø�ʽ�ļ����һ������Ϊ׼����
    Set objStream = objFile.OpenTextFile("C:\" & IIf(gintAutoRUN = 1, "AUTO", "") & "Data_LOG" & Format(Now, "yyyyMMddHHmmss") & ".txt", ForWriting, True)
    strʱ�� = "sysdate - 30"
    
    Command1.Enabled = False
    blnFirst = True
    gcnOracle.BeginTrans
    
    If Me.cbo����.ListIndex = 0 Then
        intStart = 1
        intEnd = Me.cbo����.ListCount - 1   '�����"���в���",��˴�1��ʼѭ��,ѭ�����μ�ȥ���ӵ����в���
    Else
        intEnd = Me.cbo����.ListIndex
        intStart = intEnd
    End If

redo:
    For i = intStart To intEnd
        Call WriteLog(objStream, String(50, "-"))
        Select Case gintMode
        Case 0  '����ǰ׼������ǰ��Ժ���˼����30���Ժ���˵���ʷסԺ�Ļ������ݣ�
            strSQL = "" & _
                    "SELECT  /*+ RULE */ ����ID,��ҳID,Ӥ�� " & vbNewLine & _
                    "FROM (" & vbNewLine & _
                    "    SELECT  B.����ID,B.��ҳID,0 AS Ӥ��" & vbNewLine & _
                    "    FROM ������Ϣ C,������ҳ B," & vbNewLine & _
                    "        (SELECT A.����ID" & vbNewLine & _
                    "        FROM ������Ϣ A" & vbNewLine & _
                    "        WHERE A.��Ժ=1 And A.��ǰ����ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.����ID" & vbNewLine & _
                    "        FROM ������ҳ A" & vbNewLine & _
                    "        WHERE A.��ǰ����ID=[1] And A.��Ժ����>=" & strʱ�� & ") A" & vbNewLine & _
                    "    WHERE B.����ID=C.����ID AND B.��ҳID<>C.סԺ���� AND C.����ID=A.����ID" & vbNewLine
            strSQL = strSQL & _
                    "    UNION" & vbNewLine & _
                    "    SELECT B.����ID,B.��ҳID,B.��� AS Ӥ��" & vbNewLine & _
                    "    FROM ������Ϣ C,������������¼ B," & vbNewLine & _
                    "        (SELECT A.����ID" & vbNewLine & _
                    "        FROM ������Ϣ A" & vbNewLine & _
                    "        WHERE A.��Ժ=1 And A.��ǰ����ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.����ID" & vbNewLine & _
                    "        FROM ������ҳ A" & vbNewLine & _
                    "        WHERE A.��ǰ����ID=[1] And A.��Ժ����>=" & strʱ�� & ") A" & vbNewLine & _
                    "    WHERE C.����ID=B.����ID AND C.סԺ����<>B.��ҳID AND C.����ID=A.����ID" & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT ����ID,��ҳID,Ӥ�� From ������Ǩ��¼" & vbNewLine & _
                    "    ) " & vbNewLine & _
                    "ORDER BY ����ID,��ҳID DESC ,Ӥ��"
        Case 1  '������������������Ժ���˼����30���Ժ���˵����л������ݣ�
            strSQL = "" & _
                    "SELECT  /*+ RULE */ ����ID,��ҳID,Ӥ�� " & vbNewLine & _
                    "FROM ( " & vbNewLine & _
                    "    SELECT  B.����ID,B.��ҳID,0 AS Ӥ��" & vbNewLine & _
                    "    FROM ������ҳ B," & vbNewLine & _
                    "        (SELECT A.����ID" & vbNewLine & _
                    "        FROM ������Ϣ A" & vbNewLine & _
                    "        WHERE A.��Ժ=1 And A.��ǰ����ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.����ID" & vbNewLine & _
                    "        FROM ������ҳ A" & vbNewLine & _
                    "        WHERE A.��ǰ����ID=[1] ANd A.��Ժ����>=" & strʱ�� & ") A" & vbNewLine & _
                    "    WHERE B.����ID=A.����ID" & vbNewLine
            strSQL = strSQL & _
                    "    UNION" & vbNewLine & _
                    "    SELECT B.����ID,B.��ҳID,B.��� AS Ӥ��" & vbNewLine & _
                    "    FROM ������������¼ B," & vbNewLine & _
                    "        (SELECT A.����ID" & vbNewLine & _
                    "        FROM ������Ϣ A" & vbNewLine & _
                    "        WHERE A.��Ժ=1 And A.��ǰ����ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.����ID" & vbNewLine & _
                    "        FROM ������ҳ A" & vbNewLine & _
                    "        WHERE A.��ǰ����ID=[1] And A.��Ժ����>=" & strʱ�� & ") A" & vbNewLine & _
                    "    WHERE B.����ID=A.����ID" & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT ����ID,��ҳID,Ӥ�� From ������Ǩ��¼" & vbNewLine & _
                    "    ) " & vbNewLine & _
                    "ORDER BY ����ID,��ҳID DESC ,Ӥ��"
        Case 2  '��̨ת��ʷ����
            strSQL = "" & _
                    "SELECT /*+ RULE */  ����ID,��ҳID,Ӥ�� " & vbNewLine & _
                    "FROM (" & _
                    "    SELECT A.����ID,A.��ҳID,0 AS Ӥ��" & vbNewLine & _
                    "    FROM ������ҳ A" & vbNewLine & _
                    "    WHERE A.��ǰ����ID=[1]" & vbNewLine & _
                    "    UNION" & vbNewLine & _
                    "    SELECT A.����ID,A.��ҳID,A.��� AS Ӥ��" & vbNewLine & _
                    "    FROM ������������¼ A,������ҳ B" & vbNewLine & _
                    "    WHERE A.����ID=B.����ID AND A.��ҳID=B.��ҳID AND B.��ǰ����ID=[1]) " & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT ����ID,��ҳID,Ӥ�� From ������Ǩ��¼" & vbNewLine & _
                    "ORDER BY ����ID,��ҳID desc ,Ӥ��"
        End Select
        Set rsPati = OpenSQLRecord(strSQL, "��ȡָ���������в���", CLng(Me.cbo����.ItemData(i)))
        Call WriteLog(objStream, "��ȡָ���������в����嵥...���,����:" & CLng(Me.cbo����.ItemData(i)) & ",סԺ�˴�:" & rsPati.RecordCount)
        
        With rsPati
            Do While Not .EOF
                '�����˷����仯ʱ�Ͳ������һ���ļ��Ļ�������
                If lng����ID <> 0 Then   'ֻҪ�����ݾ������µ��뻤���¼��
                    If Not blnError Then
                        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼��������������")
                        If Not InsertCollect(lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID) Then
                            blnError = True
                            strMsg = "���ڴ���[" & Me.cbo����.List(i) & "]����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID:" & lng��ʽID & "  �Ļ�������ʱ��������"
                            objStream.WriteLine "��������,�����ò���:" & vbCrLf & strMsg
                        Else
                            'Call WriteLog(objStream, "������������...���")
                        End If
                    End If
                    '������Ǩ��¼
                    gcnOracle.Execute "zl_������Ǩ��¼_Insert(" & lng����ID & "," & lng��ҳID & "," & intӤ�� & "," & IIf(blnError, "1", "0") & ",'" & Replace(Replace(Replace(Replace(strMsg, "'", ""), vbLf, ""), "[", ""), "]", "") & "')", , adCmdStoredProc
                End If
EXITDO:
                If blnCommit Or blnError Then
                    If blnError = False Then
                        gcnOracle.CommitTrans
                        blnError = False
                        blnCommit = False
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        blnError = False
                        blnCommit = False
                        gcnOracle.BeginTrans
                    End If
                End If
                
                strMsg = ""
                '�²��˽������¸�ֵ
                bln���� = False
                lng����ID = !����ID
                lng��ҳID = !��ҳID
                intӤ�� = !Ӥ��
                lng��ʽID = 0
                Me.Caption = "[" & Me.cbo����.List(i) & "] ����:" & rsPati.AbsolutePosition & "/" & rsPati.RecordCount
                
                '����Ƿ�鵵
                strSQL = " Select   �鵵��,�鵵ʱ�� From ���˻����¼ Where ����ID=[1] And ��ҳID=[2]"
                Set rsTemp = OpenSQLRecord(strSQL, "����Ƿ�鵵", lng����ID, lng��ҳID)
                If rsTemp.RecordCount <> 0 Then
                    str�鵵�� = NVL(rsTemp!�鵵��)
                    str�鵵ʱ�� = Format(rsTemp!�鵵ʱ��, "yyyy-MM-dd HH:mm:ss")
                End If
                
                '��ȡ�������Ժʱ��
                gstrSQL = " Select   ��Ժ����,nvl(��Ժ����,sysDate) as ��Ժ����,NVL(����ת��,0) AS ת�� From ������ҳ " & _
                          " Where ����ID=[1] ANd ��ҳID=[2] And [3]=0 " & _
                          " UNION " & _
                          " Select A.����ʱ�� AS ��Ժʱ��,nvl(B.��Ժ����,sysDate) as ��Ժ����,NVL(B.����ת��,0) AS ת�� From ������������¼ A,������ҳ B" & _
                          " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.����ID=[1] And A.��ҳID=[2] And A.���=[3]"
                Set rsDate = OpenSQLRecord(gstrSQL, "��ȡ�������Ժʱ��", lng����ID, lng��ҳID, intӤ��)
                blnDataMoved = rsDate!ת��
                
                '��ԭ������ȡ���˵Ļ����ļ��б�
                'Call WriteLog(objStream, "��ԭ������ȡ���˵Ļ����ļ��б�")
                strSQL = "" & _
                        " SELECT   DISTINCT A.ID, A.���, A.���� AS �ļ�," & vbNewLine & _
                        "        A.��ʼ,A.��ֹ," & vbNewLine & _
                        "        A.����ID, B.���� AS ����, 0 AS ������,A.�ļ�����,����" & vbNewLine & _
                        " FROM (" & vbNewLine & _
                        "        SELECT F.ID, F.���, F.����, R.��ʼ, R.��ֹ, R.����ID, ����,�ļ�����" & vbNewLine & _
                        "        FROM (" & vbNewLine & _
                        "        SELECT ID, ���, ����, 3 AS �ļ�����, ͨ��, 0 AS ����ID,���� FROM �����ļ��б� WHERE ����=3 AND ����<0" & vbNewLine & _
                        "        UNION ALL" & vbNewLine & _
                        "        SELECT L.ID, L.���, L.����, F.���� AS �ļ�����, L.ͨ��, A.����ID,L.����" & vbNewLine & _
                        "        FROM �����ļ��б� L, ����ҳ���ʽ F, ����Ӧ�ÿ��� A" & vbNewLine & _
                        "        WHERE L.���� = 3 AND L.���� = 0 AND L.���� = F.���� AND L.��� = F.��� AND L.ID = A.�ļ�ID(+)) F," & vbNewLine & _
                        "      (SELECT R.����ID, NVL(MIN(R.������),3) AS ������, MIN(R.����ʱ��) AS ��ʼ, MAX(R.����ʱ��) AS ��ֹ" & vbNewLine & _
                        "      FROM ���˻����¼ R" & vbNewLine & _
                        "      WHERE R.������Դ = 2 AND R.����ID = [1] AND NVL(R.��ҳID, 0) = [2] AND NVL(R.Ӥ��,0)=[3]" & vbNewLine & _
                        "      GROUP BY R.����ID) R" & vbNewLine & _
                        "        WHERE (F.����<0 OR F.ͨ�� = 1 OR F.ͨ�� = 2 AND R.����ID IN (SELECT T.����ID FROM �������Ҷ�Ӧ T WHERE T.����ID=F.����ID)) AND F.�ļ����� >= R.������) A, ���ű� B" & vbNewLine & _
                        " WHERE A.����ID = B.ID" & vbNewLine & _
                        " ORDER BY A.����,A.�ļ�����,A.��� DESC, TO_CHAR(A.��ʼ, 'YYYY-MM-DD HH24:MI') || ' �� ' || TO_CHAR(A.��ֹ, 'YYYY-MM-DD HH24:MI')"
                If blnDataMoved Then
                    strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                    strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                End If
                Set rsFile = OpenSQLRecord(strSQL, "��ԭ������ȡ���˵Ļ����ļ��б�", lng����ID, lng��ҳID, intӤ��)
                'Call WriteLog(objStream, "��ԭ������ȡ���˵Ļ����ļ��б�...���,��¼��:" & rsFile.RecordCount)
                
'                bln���� = False
                '�Ȳ����������л����ļ�(�����µ�)������,Ȼ��������ѭ�������ò��˵Ļ�������
                Do While Not rsFile.EOF
                    
                    datStart = rsDate!��Ժ����
                    datEnd = rsDate!��Ժ����
                    
                    '����ʽ�����仯ʱ�Ͳ�����������
                    '���µ���������������
'                    If bln���� Then
'                        If rsFile!���� <> -1 And lng���� <> -1 And lng��ʽID <> rsFile!ID And lng��ʽID <> 0 Then
'                            'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼��������������")
'                            If Not InsertCollect(lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID) Then
'                                blnError = True
'                                strMsg = "���ڴ���[" & Me.cbo����.List(i) & "]����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID:" & lng��ʽID & " �Ļ�������ʱ��������"
'                                objStream.WriteLine "��������,�����ò���:" & vbCrLf & strMsg
'                                Exit Do
'                            Else
'                                'Call WriteLog(objStream, "������������...���")
'                            End If
'                        End If
'                    End If
                    lng��ʽID = rsFile!ID
                    lng���� = rsFile!����
                    
                    '�Ȳ��������ļ��б�
                    If (rsFile!���� = -1 And bln���� = False) Or rsFile!���� <> -1 Then
                        lng�ļ�ID = GetNextId("���˻����ļ�")
                        strSQL = "insert into ���˻����ļ�(ID,����ID,����ID,��ҳID,Ӥ��,��ʽID,�ļ�����," & _
                                                          "��ʼʱ��,����ʱ��,����ID,�鵵��,�鵵ʱ��,������,����ʱ��)" & _
                                " Values (" & lng�ļ�ID & "," & CLng(rsFile!����ID) & "," & rsPati!����ID & "," & rsPati!��ҳID & "," & rsPati!Ӥ�� & "," & lng��ʽID & ",'" & "[" & rsFile!���� & "]" & rsFile!�ļ� & "'," & _
                                         "to_date('" & Format(datStart, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),NULL,NULL,'" & str�鵵�� & "'," & _
                                         "" & IIf(str�鵵ʱ�� = "", "NULL", "to_date('" & str�鵵ʱ�� & "','yyyy-MM-dd hh24:mi:ss')") & ",'ZLHIS',sysdate)"
                        gcnOracle.Execute strSQL
                    End If
                    
                    If rsFile!���� = -1 Then   '���µ�
                        If bln���� = False Then
                            bln���� = True
                            
                            '--3.1��������Դȫ������Ϊ�ֹ�¼������ݣ��������Աʹ���°����ʱ��ֱ�Ӹ��ģ��󶨱����У�����ID����ҳID��Ӥ������ʼʱ�䣬����ʱ��
                            '--     ���µ�������Ŀ����������Ŀ��Ҳ���ǲ��˻����¼�еļ�¼���� IN (1,5)����ֹ�汾Ϊ�յ�����Ϊ���µ�����
                            '--     ���µ��е���/�±ꡢ������ǵ����ݣ�Ҳ���ǲ��˻����¼�еļ�¼���� NOT IN ��1��5��
                            'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼���������µ���������")
                            '�������µ���������(ɾ������: And A.����ID=[1],���µ���ȡ��������)
                            strSQL = "" & _
                                    " SELECT   A.ID,A.����ʱ��,A.���汾,A.������,A.����ʱ��" & vbNewLine & _
                                    " FROM ���˻����¼ A,���˻������� B,���¼�¼��Ŀ C" & vbNewLine & _
                                    " WHERE A.ID=B.��¼ID ANd A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4] AND A.����ʱ�� BETWEEN [5] AND [6] And B.��Ŀ���=C.��Ŀ���" & vbNewLine & _
                                    " UNION" & vbNewLine & _
                                    " SELECT A.ID,A.����ʱ��,A.���汾,A.������,A.����ʱ��" & vbNewLine & _
                                    " FROM ���˻����¼ A,���˻������� B" & vbNewLine & _
                                    " WHERE A.ID=B.��¼ID ANd A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4] AND A.����ʱ�� BETWEEN [5] AND [6] And B.��¼���� NOT IN (1,5,9)"
                            strSQL = "" & _
                                    " INSERT INTO ���˻�������(ID,�ļ�ID,����ʱ��,���汾,������,����ʱ��)" & vbNewLine & _
                                    " SELECT ���˻�������_ID.Nextval,[7],A.����ʱ��,A.���汾,A.������,A.����ʱ��" & vbNewLine & _
                                    " From (" & strSQL & ") A"
                            If blnDataMoved Then
                                strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                                strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                            End If
                            Set rsTemp = OpenSQLRecord(strSQL, "�������µ���������", CLng(rsFile!����ID), lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID)
                            'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼���������µ���������...���")
                            
                            '�������µ���ϸ������
                            'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼���������µ���ϸ������")
                            strSQL = "" & _
                                    " SELECT /*+ RULE */  A.ID, D.ID AS ��¼ID, A.��¼����, A.��Ŀ����, A.��ĿID, A.��Ŀ���, A.��Ŀ����, A.��Ŀ����, A.��¼����, A.��Ŀ��λ, A.��¼���," & vbNewLine & _
                                    "      A.���²�λ, A.��¼���, A.���Ժϸ�,9 AS ������Դ, A.δ��˵��, A.��ʼ�汾, A.��ֹ�汾, A.��¼��,B.����ʱ�� AS �޸�ʱ��" & vbNewLine & _
                                    " FROM ���˻������� A,���˻����¼ B,���¼�¼��Ŀ C,���˻������� D" & vbNewLine & _
                                    " WHERE A.��¼ID=B.ID AND B.����ID=[2] AND B.��ҳID=[3] AND B.Ӥ��=[4] AND B.����ʱ�� BETWEEN [5] AND [6] And A.��Ŀ���=C.��Ŀ���" & vbNewLine & _
                                    " And B.����ʱ��=D.����ʱ�� And D.�ļ�ID=[7]" & vbNewLine & _
                                    " UNION" & vbNewLine & _
                                    " SELECT A.ID, D.ID AS ��¼ID, A.��¼����, A.��Ŀ����, A.��ĿID, A.��Ŀ���, A.��Ŀ����, A.��Ŀ����, A.��¼����, A.��Ŀ��λ, A.��¼���," & vbNewLine & _
                                    "     A.���²�λ, A.��¼���, A.���Ժϸ�,9 AS ������Դ, A.δ��˵��, A.��ʼ�汾, A.��ֹ�汾, A.��¼��,B.����ʱ�� AS �޸�ʱ��" & vbNewLine & _
                                    " FROM ���˻������� A,���˻����¼ B,���˻������� D" & vbNewLine & _
                                    " WHERE A.��¼ID=B.ID AND B.����ID=[2] AND B.��ҳID=[3] AND B.Ӥ��=[4] AND B.����ʱ�� BETWEEN [5] AND [6] And A.��¼���� NOT IN (1,5,9)" & vbNewLine & _
                                    " And B.����ʱ��=D.����ʱ�� And D.�ļ�ID=[7]"
                            strSQL = "" & _
                                    " INSERT INTO ���˻�����ϸ(ID, ��¼ID, ��¼����, ��Ŀ����, ��ĿID, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���," & vbNewLine & _
                                    "                       ���²�λ, ��¼���, ���Ժϸ�,������Դ, δ��˵��, ��ʼ�汾, ��ֹ�汾, ��¼��,��¼ʱ��)" & vbNewLine & _
                                    " SELECT ���˻�����ϸ_ID.Nextval, ��¼ID, ��¼����, ��Ŀ����, ��ĿID, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���," & vbNewLine & _
                                    "      ���²�λ, ��¼���, ���Ժϸ�, ������Դ, δ��˵��, ��ʼ�汾, ��ֹ�汾, ��¼��,�޸�ʱ��" & vbNewLine & _
                                    " From (" & strSQL & ") "
                            If blnDataMoved Then
                                strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                                strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                            End If
                            Set rsTemp = OpenSQLRecord(strSQL, "�������µ���ϸ������", CLng(rsFile!����ID), lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID)
                            'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼���������µ���ϸ������...���")
                        End If
                    Else        '�����¼��
                        '--3.2��ѭ����ȡ�ò���ָ���Ļ����ļ����ݣ����ݲ����ļ���ʽ�������������°�Ĳ��˻������ݡ����˻�����ϸ�У��󶨱����У��ļ�ID������ID����ҳID��Ӥ������ʼʱ�䡢����ʱ��
                        '--     a)�����ļ��Ļ���������ǰ�����������
                        '--�ϰ�������¼��û�л��Ŀ�ĸ���,��������Щ��Ŀ��ֻ��ʾ��Щ��Ŀ,��������ʱ��������Ŀ�Լ����µ����е����ת,����,���±����Ϣ
                        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼�����������¼����������" & IIf(rsFile!���� = -1, "", "[" & rsFile!���� & "]") & rsFile!�ļ�)
                        '���뻤���ļ���������
                        strSQL = "" & _
                                " INSERT INTO ���˻�������(ID,�ļ�ID,����ʱ��,���汾,������,����ʱ��)" & vbNewLine & _
                                " Select ���˻�������_ID.Nextval,[8],����ʱ��,���汾,������,����ʱ��" & vbNewLine & _
                                " FROM (" & vbNewLine & _
                                "   SELECT /*+ RULE */  DISTINCT ID,����ʱ��,���汾,������,����ʱ��" & vbNewLine & _
                                "   FROM (" & vbNewLine & _
                                "       SELECT C.*" & vbNewLine & _
                                "       FROM" & vbNewLine & _
                                "         (SELECT * FROM �����ļ��ṹ WHERE ��ID=(SELECT DISTINCT ID FROM �����ļ��ṹ WHERE ��������=1 AND �������=4 AND �ļ�ID=[2] )) A," & vbNewLine & _
                                "         �����¼��Ŀ B,���˻����¼ C,���˻������� D" & vbNewLine & _
                                "       WHERE �ļ�ID=[2] AND ��������=4 AND Ҫ������=B.��Ŀ����" & vbNewLine & _
                                "       And C.����ID=[1] AND C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                "       AND C.����ʱ�� BETWEEN [6] AND [7]" & vbNewLine & _
                                "       AND D.��¼ID=C.ID AND D.��¼���� IN (1,5) AND D.��Ŀ��� =B.��Ŀ���" & vbNewLine & _
                                "       UNION"
                        strSQL = strSQL & _
                                "       SELECT  C.*" & vbNewLine & _
                                "       FROM ���˻����¼ C,���˻������� D," & vbNewLine & _
                                "           (SELECT DISTINCT C.*" & vbNewLine & _
                                "           FROM" & vbNewLine & _
                                "               (SELECT * FROM �����ļ��ṹ WHERE ��ID=(SELECT DISTINCT ID FROM �����ļ��ṹ WHERE ��������=1 AND �������=4 AND �ļ�ID=[2] )) A," & vbNewLine & _
                                "               �����¼��Ŀ B,���˻����¼ C,���˻������� D" & vbNewLine & _
                                "           WHERE �ļ�ID=[2] AND ��������=4 AND Ҫ������=B.��Ŀ����" & vbNewLine & _
                                "           And C.����ID=[1] AND C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                "           AND C.����ʱ�� BETWEEN [6] AND [7]" & vbNewLine & _
                                "           AND D.��¼ID=C.ID AND D.��¼����=1 AND D.��Ŀ��� =B.��Ŀ���) A" & vbNewLine & _
                                "       WHERE C.����ID=[1] And C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                "       AND C.����ʱ�� between [6] AND [7]" & vbNewLine & _
                                "       AND D.��¼ID=C.ID AND D.��¼����=5 And C.ID=A.ID))"
                        If blnDataMoved Then
                            strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                            strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                        End If
                        Set rsTemp = OpenSQLRecord(strSQL, "���뻤���ļ���������", CLng(rsFile!����ID), lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID)
                        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼�����������¼����������...���")
                        
                        '���뻤���ļ���ϸ������
                        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼�����������¼����ϸ������")
                        strSQL = "" & _
                                " SELECT /*+ RULE */ D.ID, Z.ID AS ��¼ID, D.��¼����, D.��Ŀ����, D.��ĿID, D.��Ŀ���, D.��Ŀ����, D.��Ŀ����, D.��¼����, D.��Ŀ��λ, D.��¼���," & vbNewLine & _
                                "      D.���²�λ, D.��¼���, D.���Ժϸ�,0 AS ������Դ, D.δ��˵��, D.��ʼ�汾, D.��ֹ�汾, D.��¼��,C.����ʱ�� AS �޸�ʱ��" & vbNewLine & _
                                " FROM" & vbNewLine & _
                                "    (SELECT * FROM �����ļ��ṹ WHERE ��ID=(SELECT DISTINCT ID FROM �����ļ��ṹ WHERE ��������=1 AND �������=4 AND �ļ�ID=[2] )) A," & vbNewLine & _
                                "    �����¼��Ŀ B,���˻����¼ C,���˻������� D,���˻������� Z" & vbNewLine & _
                                " WHERE A.�ļ�ID=[2] AND ��������=4 AND Ҫ������=B.��Ŀ����" & vbNewLine & _
                                " And C.����ID=[1] AND C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                " AND C.����ʱ�� BETWEEN [6] AND [7] And C.����ʱ��=Z.����ʱ�� And Z.�ļ�ID=[8]" & vbNewLine & _
                                " AND D.��¼ID=C.ID AND D.��¼����=1 AND D.��Ŀ��� =B.��Ŀ���" & vbNewLine & _
                                " UNION"
                        strSQL = strSQL & _
                                " SELECT  D.ID, Z.ID AS ��¼ID, D.��¼����, D.��Ŀ����, ��ĿID, D.��Ŀ���,D.��Ŀ����, D.��Ŀ����, D.��¼����, D.��Ŀ��λ, D.��¼���," & vbNewLine & _
                                "     D.���²�λ, D.��¼���, D.���Ժϸ�,0 AS ������Դ, D.δ��˵��, D.��ʼ�汾, D.��ֹ�汾, D.��¼��,C.����ʱ�� AS �޸�ʱ��" & vbNewLine & _
                                " FROM ���˻����¼ C,���˻������� D," & vbNewLine & _
                                "    (SELECT DISTINCT C.*" & vbNewLine & _
                                "    FROM" & vbNewLine & _
                                "       (SELECT * FROM �����ļ��ṹ WHERE ��ID=(SELECT DISTINCT ID FROM �����ļ��ṹ WHERE ��������=1 AND �������=4 AND �ļ�ID=[2] )) A," & vbNewLine & _
                                "    �����¼��Ŀ B,���˻����¼ C,���˻������� D" & vbNewLine & _
                                "    WHERE A.�ļ�ID=[2] AND ��������=4 AND Ҫ������=B.��Ŀ����" & vbNewLine & _
                                "    And C.����ID=[1] AND C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                "    AND C.����ʱ�� BETWEEN [6] AND [7]" & vbNewLine & _
                                "    AND D.��¼ID=C.ID AND D.��¼����=1 AND D.��Ŀ��� =B.��Ŀ���) A,���˻������� Z" & vbNewLine & _
                                " WHERE C.����ID=[1] And C.����ID=[3] AND C.��ҳID=[4] AND NVL(C.Ӥ��,0)=[5]" & vbNewLine & _
                                " AND C.����ʱ�� between [6] AND [7] And C.����ʱ��=Z.����ʱ�� And Z.�ļ�ID=[8]" & vbNewLine & _
                                " AND D.��¼ID=C.ID AND D.��¼����=5 And C.ID=A.ID"
                        strSQL = "" & _
                                " INSERT INTO ���˻�����ϸ(ID, ��¼ID, ��¼����, ��Ŀ����, ��ĿID, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���," & vbNewLine & _
                                "                       ���²�λ, ��¼���, ���Ժϸ�,������Դ, δ��˵��, ��ʼ�汾, ��ֹ�汾, ��¼��,��¼ʱ��)" & vbNewLine & _
                                " SELECT ���˻�����ϸ_ID.Nextval, ��¼ID, ��¼����, ��Ŀ����, DECODE(��¼����,5,NULL,��ĿID), ��Ŀ���, DECODE(��¼����,5,NULL,��Ŀ����), ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���," & vbNewLine & _
                                "      ���²�λ, ��¼���, ���Ժϸ�, ������Դ, δ��˵��, ��ʼ�汾, ��ֹ�汾, ��¼��,�޸�ʱ��" & vbNewLine & _
                                " FROM (" & strSQL & ")"
                        If blnDataMoved Then
                            strSQL = Replace(strSQL, "���˻����¼", "H���˻����¼")
                            strSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                        End If
                        Set rsTemp = OpenSQLRecord(strSQL, "���뻤���ļ���ϸ������", CLng(rsFile!����ID), lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID)
                        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼�����������¼����ϸ������...���")
                        
                    End If
                    '��һ�������ļ�
                    rsFile.MoveNext
                    DoEvents
                Loop
                
                '�Զ�ִ��ʱ���ʱ�� , ǰ4Сʱ��Ǩ����, ��2Сʱ�������ݴ�ӡ����
                If gintAutoRUN = 1 Then
                    If Format(Now, "HH:mm") >= gstrNextTime Then
                        blnFirst = False    '�����ٴν�����һ��ѭ��
                        GoTo todoPrint
                    End If
                End If
                
                '��һ������
                .MoveNext
            Loop
            
            blnCommit = True 'ѭ������Ӧ���ύ��
        End With
    Next
    
todoPrint:
    '���һ�����˵����һ���ļ���ʽ��Ҫ������������
    If lng����ID <> 0 Then
        'Call WriteLog(objStream, "����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "׼��������������")
        If Not InsertCollect(lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd, lng�ļ�ID) Then
            blnError = True
            strMsg = "���ڴ���[" & Me.cbo����.List(i) & "]����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "  �Ļ�������ʱ��������"
            objStream.WriteLine "��������,�����ò���:" & vbCrLf & strMsg
        Else
            'Call WriteLog(objStream, "������������...���")
        End If
        
        '������Ǩ��¼
        gcnOracle.Execute "zl_������Ǩ��¼_Insert(" & lng����ID & "," & lng��ҳID & "," & intӤ�� & "," & IIf(blnError, "1", "0") & ",'" & Replace(strMsg, "'", "") & "')", , adCmdStoredProc
    End If
    
    If Not blnError Then
        gcnOracle.CommitTrans
    Else
        gcnOracle.RollbackTrans
    End If
    
    '�������4��,�򲻽������ѭ��,�����Ѷ�blnFirst��ֵΪ��
    If gintAutoRUN = 1 Then
        If blnFirst Then
            blnFirst = False
            gcnOracle.BeginTrans
            GoTo redo
        End If
    End If
    
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & "������Ǩ�ɹ�!"
    objStream.Close
    
    Me.Caption = "���ڽ��д�ӡ���ݽ���,���Ժ�..."
    Call DoPrintData
    
    If gintAutoRUN = 1 Then Unload Me
    Command1.Enabled = True
    Exit Sub
errHand:
    blnError = True
    strMsg = "���ڴ���[" & Me.cbo����.List(i) & "]����ID=" & lng����ID & ";��ҳID=" & lng��ҳID & ";Ӥ��=" & intӤ�� & ";�ļ���ʽID=" & lng��ʽID & "  �Ļ�������ʱ��������:" & Err.Description
    objStream.WriteLine "��������,�����ò���:" & vbCrLf & strMsg
    GoTo EXITDO
End Sub

Private Function InsertCollect(ByVal lng��ʽID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, _
    ByVal datStart As Date, ByVal datEnd As Date, ByVal lng�ļ�ID As Long) As Boolean
    Dim datCur As Date
    Dim lngID As Long
    Dim strSQL As String
    Dim str���� As String, lng������� As Long, lng�ļ�ID_Cur As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
'    '--һ������һ���ļ�ֻͳ��һ��
'    '--3.3����ǩ�����ݸ���
'    strSQL = " SELECT distinct B.ID,Nvl(D.Ƹ�μ���ְ��,5)-1 AS ǩ������" & _
'             " From ���˻����ļ� A,���˻������� B,���˻�����ϸ C,��Ա�� D" & _
'             " Where A.ID=[1] And A.ID=B.�ļ�ID And B.ID=C.��¼ID And C.��¼����=5 And B.������=D.����"
'    Set rsTemp = OpenSQLRecord(strSQL, "����ǩ������", lng�ļ�ID)
'    Do While Not rsTemp.EOF
'        gcnOracle.Execute "Update ���˻������� Set ǩ����=������,ǩ��ʱ��=sysdate,ǩ������=" & rsTemp!ǩ������ & " Where ID=" & rsTemp!ID
'        rsTemp.MoveNext
'    Loop
    
'    '--3.4�����������¼���Ļ�������,��ǰ��չ������ʱ������ܵ�,�󶨱���:�ļ�ID,����ID,��ҳID,Ӥ��
'    '--���ݻ����ļ���ʽ������ȡ��������,����ѭ����������,��ϸ��
'    '--ע��:С��������Ϊ��Ŀ����,����ʱ��+���������Ϊ����ʱ��,��������Ϊ��¼����,����+������ŷ����仯ʱ,ȡ���˻����¼_Nextval
'    'ѭ�������������
'    strSQL = "" & _
'            " SELECT A.����,A.�������,A.С������,A.��ʼʱ��,A.����ʱ��,B.��Ŀ����,B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ,SUM(zl_to_number(B.��¼����,2)) AS ������" & vbNewLine & _
'            " FROM" & vbNewLine & _
'            "    (SELECT" & vbNewLine & _
'            "        B.����,A.�������,A.С������,B.����||' '||DECODE(SIGN(LENGTH(A.��ʼʱ��)-8),-1,'0','')||A.��ʼʱ�� AS ��ʼʱ��," & vbNewLine & _
'            "        DECODE(SIGN(TO_NUMBER(SUBSTR(A.��ʼʱ��,1,INSTR(A.��ʼʱ��,':',1)-1))-TO_NUMBER(SUBSTR(A.����ʱ��,1,INSTR(A.����ʱ��,':',1)-1)))," & vbNewLine & _
'            "            1,TO_CHAR(TO_DATE(B.����,'YYYY-MM-DD')+1,'YYYY-MM-DD'),B.����)||' '||DECODE(SIGN(LENGTH(A.����ʱ��)-8),-1,'0','')||A.����ʱ�� AS ����ʱ��" & vbNewLine & _
'            "    FROM"
'    strSQL = strSQL & _
'            "        (SELECT �������,С������," & vbNewLine & _
'            "        ��ʼʱ��||DECODE(INSTR(��ʼʱ��,':',1),0,':00:00',':00') AS ��ʼʱ��," & vbNewLine & _
'            "        ����ʱ��||DECODE(INSTR(����ʱ��,':',1),0,':59:59',':59') AS ����ʱ��" & vbNewLine & _
'            "        FROM (" & vbNewLine & _
'            "            SELECT �������,SUBSTR(�����ı�,1,INSTR(�����ı� ,',',1,1)-1) AS С������," & vbNewLine & _
'            "                   replace(SUBSTR(�����ı�,INSTR(�����ı� ,',',1,1)+1,INSTR(�����ı� ,',',1,2)-INSTR(�����ı� ,',',1,1)-1),'��',':') AS ��ʼʱ��," & vbNewLine & _
'            "                   replace(SUBSTR(�����ı�,INSTR(�����ı� ,',',1,2)+1,LENGTH(�����ı�)-INSTR(�����ı� ,',',1,2)),'��',':') AS ����ʱ��" & vbNewLine & _
'            "            FROM �����ļ��ṹ" & vbNewLine & _
'            "            WHERE ��ID=(SELECT DISTINCT ID FROM �����ļ��ṹ WHERE ��������=1 AND �������=5 AND �ļ�ID=[1]))) A," & vbNewLine & _
'            "        (SELECT DISTINCT TO_CHAR(����ʱ��,'YYYY-MM-DD') AS ���� FROM ���˻����¼ A WHERE A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4] ) B) A," & vbNewLine & _
'            "    (SELECT A.����ʱ��,C.��Ŀ���,C.��Ŀ����,B.��¼����,C.��Ŀ��λ" & vbNewLine & _
'            "    FROM ���˻����¼ A,���˻������� B,�����¼��Ŀ C,�����ļ��ṹ D" & vbNewLine & _
'            "    WHERE A.ID=B.��¼ID AND A.����ID=[2] AND A.��ҳID=[3] AND A.Ӥ��=[4] AND A.����ʱ�� between [5] AND [6]" & vbNewLine & _
'            "    AND B.��Ŀ���=C.��Ŀ��� AND D.Ҫ������=C.��Ŀ���� AND NVL(D.Ҫ�ر�ʾ,0)=1 AND D.�ļ�ID=[1]) B" & vbNewLine & _
'            "WHERE B.����ʱ�� BETWEEN TO_DATE(A.��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') AND TO_DATE(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
'            "GROUP BY A.����,A.�������,A.С������,A.��ʼʱ��,A.����ʱ��,B.��Ŀ���,B.��Ŀ����,B.��Ŀ��λ" & vbNewLine & _
'            "ORDER BY ����,�������"
'    Set rsTemp = OpenSQLRecord(strSQL, "���뻤���ļ���ϸ������", lng��ʽID, lng����ID, lng��ҳID, intӤ��, datStart, datEnd)
'    With rsTemp
'        Do While Not .EOF
'            If str���� <> !���� Or lng������� <> !������� Then
'                str���� = !����
'                lng������� = !�������
'
'                '���ݽ���ʱ��ȡ���ڿ���(�п�����һת�ڶ�)
'                datCur = !����ʱ��
'                strSQL = "" & _
'                        " SELECT A.����ID,B.ID AS �ļ�ID" & vbNewLine & _
'                        " FROM ���˱䶯��¼ A,���˻����ļ� B" & vbNewLine & _
'                        " WHERE A.����ID=B.����ID And A.����ID=B.����ID And A.��ҳID=B.��ҳID " & vbNewLine & _
'                        " And B.��ʽID=[5] And A.����ID=[1] AND A.��ҳID=[2] And B.Ӥ��=[3]" & vbNewLine & _
'                        " AND [4] BETWEEN A.��ʼʱ�� AND NVL(A.��ֹʱ��,SYSDATE)"
'                Set rsDept = OpenSQLRecord(strSQL, "��ȡ����ʱ�䲡����������", lng����ID, lng��ҳID, intӤ��, datCur, lng��ʽID)
'                lngID = 0
'                lng�ļ�ID_Cur = 0
'                'lng�ļ�ID_Cur = lng�ļ�ID      '������ȱʡֵ,���ĸ��ļ��ľͲ������ĸ��ļ���ȥ
'
'                If rsDept.RecordCount <> 0 Then
'                    lng�ļ�ID_Cur = rsDept!�ļ�ID
'                End If
'
'                If lng�ļ�ID_Cur <> 0 Then
'                    lngID = GetNextId("���˻�������")
'                    '��������¼
'                    strSQL = "" & _
'                            " INSERT INTO ���˻�������(ID,�ļ�ID,����ʱ��,���汾,������,����ʱ��,�������,�����ı�,���ܱ��)" & vbNewLine & _
'                            " Values (" & lngID & "," & lng�ļ�ID_Cur & ",to_date('" & Format(DateAdd("s", !�������, !����ʱ��), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
'                                     "NULL,'ZLHIS',sysdate," & -1 * !������� & ",'" & !С������ & "',0)"
'                    gcnOracle.Execute strSQL
'                End If
'            End If
'
'            If lngID <> 0 Then
'                '������ϸ����
'                strSQL = "" & _
'                        " INSERT INTO ���˻�����ϸ(ID, ��¼ID, ��¼����, ��Ŀ����, ��ĿID, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���," & vbNewLine & _
'                        "                       ���²�λ, ��¼���, ���Ժϸ�,������Դ, δ��˵��, ��ʼ�汾, ��ֹ�汾, ��¼��,��¼ʱ��)" & vbNewLine & _
'                        " Values (���˻�����ϸ_ID.Nextval," & lngID & ",1,'" & !��Ŀ���� & "',NULL," & !��Ŀ��� & ",'" & !��Ŀ���� & "',0,'" & !������ & "','" & NVL(!��Ŀ��λ) & "',0," & _
'                                 "NULL,NULL,0,0,NULL,1,NULL,'ZLHIS',sysdate)"
'                gcnOracle.Execute strSQL
'            End If
'
'            .MoveNext
'        Loop
'    End With
    
    InsertCollect = True
errHand:
    Exit Function
End Function

Private Sub WriteLog(ByVal objStream As TextStream, ByVal strLog As String)
    objStream.WriteLine "ʱ��:" & Format(Now, "yyyy-MM-dd HH:mm:ss") & ";" & strLog
End Sub

Private Sub DoPrintData()
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim arrData
    Dim lngRows As Long, lngParent As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsFormat As New ADODB.Recordset
    On Error GoTo errHand
    '���ļ�ѭ������������ز��˵����ݽ��д�ӡ���ݽ���
    Set objStream = objFile.OpenTextFile("C:\" & IIf(gintAutoRUN = 1, "AUTO", "") & "PRINT_LOG" & Format(Now, "yyyyMMddHHmmss") & ".txt", ForWriting, True)
    
    strSQL = " Select ID,���,���� From �����ļ��б� Where ����=3 And ����<>-1 and ͨ��<>0 Order by ���"
    Call OpenRecordset(rsFile, strSQL, "��ȡ�����ļ��б�")
    
    With rsFile
        Do While Not .EOF
            objStream.WriteLine String(50, "-")
            objStream.WriteLine "�ļ���" & rsFile!ID & "��" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "��ʼ���д�ӡ����"
            
            '��ȡҳ���ʽ������ദ��Ҫʹ�ã�
            '(ֽ��|ֽ��|��|��|�ϱ߾�|�±߾�|��߾�|�ұ߾�|�и�|�̶�����|������������|�����������С|�����ı�|������������|�����������С|�������ı�)
            strSQL = "" & _
                    " SELECT B.ID,A.����,A.���,B.����," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,1,INSTR(A.��ʽ,';',1,1)-1) AS PAGE ," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,1)+1,INSTR(A.��ʽ,';',1,2)-INSTR(A.��ʽ,';',1,1)-1) AS Orient," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,2)+1,INSTR(A.��ʽ,';',1,3)-INSTR(A.��ʽ,';',1,2)-1) AS HEIGHT ," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,3)+1,INSTR(A.��ʽ,';',1,4)-INSTR(A.��ʽ,';',1,3)-1) AS WIDTH ," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,4)+1,INSTR(A.��ʽ,';',1,5)-INSTR(A.��ʽ,';',1,4)-1) AS LEFT ," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,5)+1,INSTR(A.��ʽ,';',1,6)-INSTR(A.��ʽ,';',1,5)-1) AS RIGHT," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,6)+1,INSTR(A.��ʽ,';',1,7)-INSTR(A.��ʽ,';',1,6)-1) AS TOP," & vbNewLine & _
                    "       SUBSTR(A.��ʽ,INSTR(A.��ʽ,';',1,7)+1,DECODE(INSTR(A.��ʽ,';',1,8),0,LENGTH(��ʽ)+1,INSTR(A.��ʽ,';',1,8))-INSTR(A.��ʽ,';',1,7)-1) AS BOTTOM" & vbNewLine & _
                    " FROM ����ҳ���ʽ A,�����ļ��б� B " & _
                    " WHERE A.����=B.���� AND A.���=B.��� AND B.����=3 AND B.����<>-1 And B.ID=[1]" & vbNewLine & _
                    " ORDER BY ���"
            Set rsFormat = OpenSQLRecord(strSQL, "��ȡ�����ļ���ʽ", CLng(rsFile!ID))
            arrData = Split(rsFormat!Page & "," & rsFormat!orient & "," & rsFormat!Height & "," & rsFormat!Width & "," & rsFormat!Left & "," & rsFormat!Right & "," & rsFormat!Top & "," & rsFormat!Bottom, ",")
            
            '���û�н��������ļ�һҳ����ʾ��������Ч�������Ƚ��н���������
            strSQL = "select ��ID,�����ı�,�������� from �����ļ��ṹ where ��ID=(select ID from �����ļ��ṹ where �ļ�ID=[1] and �������=1 and ��ID is null)"
            Set rsTemp = OpenSQLRecord(strSQL, "��ȡ��ǰ�ļ���Ч������", CLng(rsFile!ID))
            lngParent = rsTemp!��ID
            rsTemp.Filter = "��������='��Ч������'"
            If rsTemp.RecordCount <> 0 Then
                lngRows = rsTemp!�����ı�
            Else
                lngRows = frmPreview.ShowMe(Me, rsFile!ID, arrData)
                
                '�������ݹ��Ժ�ʹ��
                gcnOracle.Execute "insert into �����ļ��ṹ(ID,�ļ�ID,��ID,�������,��������,��������,�����ı�,Ҫ������) select �����ļ��ṹ_ID.Nextval," & rsFile!ID & "," & lngParent & ",13,4,'��Ч������'," & lngRows & ",'��Ч������' from dual"
            End If
            
            '������ѭ�������������ݽ��н���
            gcnOracle.BeginTrans
            Call frmPreview.AnaliseData(Me, rsFile!ID, arrData, objStream)
            gcnOracle.CommitTrans
            objStream.WriteLine "�ļ���" & rsFile!ID & "��" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "��ɴ�ӡ������"
            
            If gintAutoRUN = 1 Then
                If Format(Now, "HH:mm") >= gstrEndTime Then
                    Exit Do
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & "���"
    objStream.Close
    Unload frmPreview
    
    If gintAutoRUN = 0 Then MsgBox "�������ݴ�ӡ������ɣ�"
    Exit Sub
errHand:
    MsgBox Err.Description
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & Err.Description
    objStream.Close
End Sub

Private Sub Command1_Click()
    If gintMode = -1 Then
        MsgBox "����ѡ��һ�����ݵ���ģʽ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrStartTime = Format(Me.dtp��ʼʱ��.Value, "HH:mm")
    gstrNextTime = Format(Me.dtp��ʼʱ��1.Value, "HH:mm")
    gstrEndTime = Format(Me.dtp����ʱ��.Value, "HH:mm")
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��", gstrStartTime
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��1", gstrNextTime
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "����ʱ��", gstrEndTime
    
    If gintMode = 1 Then
        Call DataUpgrade
    Else
        Me.Hide
        tim.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
    frmSet.Show 1, Me
End Sub

Private Sub Form_Activate()
    If gintAutoRUN = 1 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    If gintAutoRUN = 1 Then
        If gintMode < 0 Then
            gintAutoRUN = 0
            MsgBox "�Զ�������ָ������ģʽ�����ڽ����ֶ�ģʽ��", vbInformation, gstrSysName
        ElseIf gintMode > 2 Then
            gintAutoRUN = 0
            MsgBox "ָ���ĵ���ģʽ����(0,1,2)�����ڽ����ֶ�ģʽ��", vbInformation, gstrSysName
        End If
    End If
    
    strSQL = "" & _
            " SELECT DISTINCT A.ID,A.����,A.����" & vbNewLine & _
            " FROM ���ű� A,��������˵�� B" & vbNewLine & _
            " WHERE A.ID=B.����ID AND B.������� IN(1,2,3) AND B.��������='����'" & vbNewLine & _
            " AND (A.����ʱ�� IS NULL OR TRUNC(A.����ʱ��)=TO_DATE('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
            " ORDER BY A.����"
    Call OpenRecordset(rsTemp, strSQL, "��ȡ��Ժ���в���")
    With rsTemp
        Me.cbo����.Clear
        Me.cbo����.AddItem "���в���"
        Do While Not .EOF
            Me.cbo����.AddItem !����
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            .MoveNext
        Loop
        Me.cbo����.ListIndex = 0
    End With
    
    gstrStartTime = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��", "00:00")
    gstrNextTime = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "��ʼʱ��1", "02:00")
    gstrEndTime = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\����������Ǩ", "����ʱ��", "04:00")
    dtp��ʼʱ��.Value = gstrStartTime
    dtp��ʼʱ��1.Value = gstrNextTime
    dtp����ʱ��.Value = gstrEndTime
End Sub

Private Sub opt_Click(Index As Integer)
    Command1.Caption = "��̨����"
    Select Case Index
    Case 0
        Label2.Caption = "    ��ǰ��Ժ���˼����30���Ժ���˵���ʷ��������(��ʱ������Ӱ��)"
    Case 1
        Label2.Caption = "    ������������Ժ���˼����30���Ժ���˵����л�������"
        Command1.Caption = "������Ǩ"
    Case 2
        Label2.Caption = "    ������ʷ���˻�������(��ʱ������Ӱ��)"
    End Select
    
    If gintAutoRUN = 1 Then Exit Sub
    gintMode = Index
End Sub

Private Sub tim_Timer()
    '������Ǩ����Զ����д�ӡ����,���Դ˴�ֻ�ж�������Ǩ����Чʱ�伴��
    If gintMode = 1 Then Exit Sub   '����ģʽֱ������
    If Not (Format(Now, "HH:mm") >= gstrStartTime And Format(Now, "HH:mm") <= gstrNextTime) Then Exit Sub
    If Not Command1.Enabled Then Exit Sub
    
    Call DataUpgrade
End Sub
