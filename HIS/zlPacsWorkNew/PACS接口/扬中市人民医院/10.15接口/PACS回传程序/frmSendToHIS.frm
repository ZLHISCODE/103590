VERSION 5.00
Begin VB.Form frmSendToHIS 
   Caption         =   "PACS�ش�����"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "��������"
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "ֹͣ����"
      Height          =   350
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   350
      Left            =   3960
      TabIndex        =   4
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "��������"
      Height          =   350
      Left            =   3960
      TabIndex        =   3
      Top             =   720
      Width           =   1100
   End
   Begin VB.TextBox txtInterval 
      Height          =   285
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer tmlistener 
      Interval        =   5000
      Left            =   3600
      Top             =   720
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   375
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "ʱ������"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Height          =   2145
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmSendToHIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim lngInterval As Long
    
    If Val(Me.txtInterval.Text) < 1 Then
        lngInterval = 1
    ElseIf Val(Me.txtInterval.Text) > 65 Then
        lngInterval = 65
    Else
        lngInterval = Val(Me.txtInterval.Text)
    End If
    Me.txtInterval.Text = lngInterval
    
    Me.tmlistener.Enabled = False
    Me.tmlistener.Interval = lngInterval * 1000
    Me.tmlistener.Enabled = True
    glngInterval = lngInterval
    MsgBox "ʱ�������ñ���ɹ���"
End Sub

Private Sub cmdStart_Click()
    Me.tmlistener.Enabled = True
    Me.cmdStop.Enabled = True
    Me.cmdStart.Enabled = False
End Sub

Private Sub cmdStop_Click()
    Me.tmlistener.Enabled = False
    Me.cmdStart.Enabled = True
    Me.cmdStop.Enabled = False
    
End Sub


Private Sub Form_Load()
    If glngInterval < 1 Or glngInterval > 65 Then glngInterval = 1
    Me.tmlistener.Interval = glngInterval * 1000
    Me.txtInterval.Text = Me.tmlistener.Interval / 1000
    Me.cmdStart.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "PACS�ش�", "HISIP��ַ", gstrHISIP
    SaveSetting "ZLSOFT", "PACS�ش�", "HIS�û���", gstrUser
    SaveSetting "ZLSOFT", "PACS�ش�", "HIS����", gstrPassw
    SaveSetting "ZLSOFT", "PACS�ش�", "�������", glngInterval
    
    
    SaveSetting "ZLSOFT", "PACS�ش�", "PACS�û���", gstrPACSUser
    SaveSetting "ZLSOFT", "PACS�ش�", "PACS����", gstrPACSPassw
    SaveSetting "ZLSOFT", "PACS�ش�", "PACSsid", gstrPACSsid

End Sub

Private Sub tmlistener_Timer()
    subSendDataToHIS
End Sub

Private Sub subSendDataToHIS()
    Dim dsReport As New ADODB.Recordset
    Dim dsAdvice As New ADODB.Recordset
    Dim dsState As New ADODB.Recordset
    Dim strSQL As String
    Dim strCheckDocID As String     '���ҽ��
    Dim strWriteDocID As String     '���ҽ��
    Dim intRowCount As Integer      '�ɹ��޸ĵļ�¼����
    
    On Error GoTo errLog
    '��ѯ��ʱ�� PACS_TMP���˲�����¼����������ݣ���ʼ���ݻش�
    
    strSQL = "select id,����ID,����ID,����ID,��������,��д��ID,��д��,to_date(��д����) as ��д����,������ID,������,��������,��¼���� from PACS_TMP���˲�����¼"
    Set dsReport = gcnOracle.Execute(strSQL)
    If dsReport.EOF Then        'û�����ݣ��˳�����
        Exit Sub
    Else
        dsReport.MoveFirst
        While Not dsReport.EOF
            '���� PACS_TMP���˲�����¼.id=����ҽ������.����id����ѯ��Ӧ�ġ�ҽ��ID��
            strSQL = "select ҽ��ID from ����ҽ������ where ����ҽ������.����id = " & dsReport!����ID
            Set dsAdvice = gcnOracle.Execute(strSQL)
            
            If Not dsAdvice.EOF Then
                '�����ж�Ӧҽ��ID�ı��棬����PACS_TMP���˲�����¼.��¼�����жϻش�����д�ˡ����ǡ������ˡ����ش���ɾ����¼
                'PACS��ҽ��ID��ӦHIS�ļ���
                lblStatus.Caption = Date & " " & Time & vbCrLf & " ���ڻش�:ID--" & dsReport!����ID & " ��д��--" & dsReport!��д�� & " ������--" & dsReport!������
                strWriteDocID = Format(dsReport!��д��ID, "0000")
                strCheckDocID = Format(dsReport!������ID, "0000")
                
                If Len(strWriteDocID) > 4 Then strWriteDocID = Left(strWriteDocID, 4)
                If Len(strCheckDocID) > 4 Then strCheckDocID = Left(strCheckDocID, 4)
                
                If dsReport!��¼���� = 1 Then   '������д��ɣ��ش���д�ˣ���дʱ��ͱ�����ɱ�־=1
                
                    strSQL = "update pacs_bldak set zdys ='" & strWriteDocID & "', bgdate = '" _
                        & Format(dsReport!��д����, "YYYY-MM-DD HH:mm:ss") & "',bgzt='1' where jcdh = '" _
                        & dsAdvice!ҽ��ID & "'"
                    gcnSQL2K.Execute strSQL, intRowCount
                    
                    'HIS��û�м�¼�����ģ��������˺ͱ���ʱ��д����־
                    If intRowCount <= 0 Then
                        subLogErr 100, "������д�����,��д��=" & dsReport!��д�� & " ��д����=" _
                                       & dsReport!��д���� & " ���� = " & dsAdvice!ҽ��ID
                    End If
                    
                ElseIf dsReport!��¼���� = 2 Then   '���汻��ˣ��ش������
                    
                    '��鱨���Ƿ����״̬������Ѿ���ˣ���ش������
                    '�����Ǳ��汻�޸ģ�ֱ��ɾ��������¼
                    strSQL = "select ִ��״̬,ִ�й��� from ����ҽ������ where ҽ��ID =" & dsAdvice!ҽ��ID
                    Set dsState = gcnOracle.Execute(strSQL)
                    
                    If dsState!ִ��״̬ = 1 And dsState!ִ�й��� = 6 Then
                        
                        strSQL = "update pacs_bldak set zdys ='" & strWriteDocID & "', bgdate = '" _
                            & Format(dsReport!��д����, "YYYY-MM-DD HH:mm:ss") & "',shys ='" _
                            & strCheckDocID & "' where jcdh = '" & dsAdvice!ҽ��ID & "'"
                        gcnSQL2K.Execute strSQL, intRowCount
                        
                        If intRowCount <= 0 Then    'HIS��û�м�¼�����ģ��������˺ͱ���ʱ��д����־
                            subLogErr 101, "�����д�����,��д��=" & dsReport!��д�� & " ��д����=" _
                                       & dsReport!��д���� & " ����� = " & dsReport!������ _
                                       & " ���� = " & dsAdvice!ҽ��ID
                        End If
                        
                    End If
                ElseIf dsReport!��¼���� = 3 Then   '���汻���أ��ش�������ɱ�־=0,Ŀǰ����Ҫ���״̬
                
                    'strSQL = "update pacs_bldak set bgzt='1' where jcdh = '" & dsAdvice!ҽ��ID & "'"
                End If
            
            End If
            'û�ж�Ӧҽ���ı���Ϊ���뵥��ֱ��ɾ����¼
            '�����˴洢��ɺ�ֱ��ɾ����¼
            strSQL = "delete from PACS_TMP���˲�����¼ where id = " & dsReport!id
            gcnOracle.Execute (strSQL)
            
            lblStatus.Caption = Date & " " & Time & vbCrLf & " �ش����:ID--" & dsReport!����ID
            dsReport.MoveNext
        Wend
    End If
    
    Exit Sub
errLog:
    subLogErr Err.Number, Err.Description
End Sub


Private Sub subLogErr(lngErrNo As Long, strDesc As String)
    On Error Resume Next
    Dim lngID As Long
    Dim strSQL As String
    Dim dsRecord As ADODB.Recordset
    
    Me.lblStatus.Caption = Date & " " & Time & vbCrLf & " �������󣬴�����룺" & lngErrNo & " ����������" & strDesc
    strSQL = "SELECT MAX(ID) as mID FROM PACS_ERR"
    Set dsRecord = gcnOracle.Execute(strSQL)
    If Not dsRecord.EOF Then
        lngID = dsRecord!Mid + 1
    End If
    strSQL = "insert into PACS_ERR (ID,�����,��������,����ʱ��) values(" & lngID & "," _
            & lngErrNo & ",'" & strDesc & "',sysdate)"
    gcnOracle.Execute strSQL
End Sub


