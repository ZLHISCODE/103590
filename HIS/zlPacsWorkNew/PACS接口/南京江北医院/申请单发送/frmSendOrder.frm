VERSION 5.00
Begin VB.Form frmSendOrder 
   Caption         =   "HIS���������뵥�����ͷ���"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5100
   Icon            =   "frmSendOrder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5100
   StartUpPosition =   2  '��Ļ����
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
Attribute VB_Name = "frmSendOrder"
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
    
    SaveSetting "ZLSOFT", gstrRegPath, "HIS�û���", gstrHISUser
    SaveSetting "ZLSOFT", gstrRegPath, "HIS����", gstrHISPassw
    SaveSetting "ZLSOFT", gstrRegPath, "HISsid", gstrHISsid
     
    SaveSetting "ZLSOFT", gstrRegPath, "�������", glngInterval
    
    SaveSetting "ZLSOFT", gstrRegPath, "PACSIP��ַ", gstrPACSIP
    SaveSetting "ZLSOFT", gstrRegPath, "PACS�û���", gstrPACSUser
    SaveSetting "ZLSOFT", gstrRegPath, "PACS����", gstrPACSPassw
    SaveSetting "ZLSOFT", gstrRegPath, "PACSsid", gstrPACSsid
    SaveSetting "ZLSOFT", gstrRegPath, "gstrPACSport", gstrPACSport

End Sub

Private Sub tmlistener_Timer()
    subSendDataToPACS
End Sub

Private Sub subSendDataToPACS()
    Dim adoCmd As New ADODB.Command
    Dim dsOrder As ADODB.Recordset
    Dim adoParaReturn As Parameter  '�洢����ֵ
    Dim strSQL As String
    
    
    
    On Error GoTo errLog
    '��ѯ��ʱ�� PACS_TMP���˲�����¼����������ݣ���ʼ���ݻش�
    
    strSQL = "Select ID,��������,�������,ҽ��ID,��ʶ��,����ID,����,Ӣ����,�Ա�,��������,���֤��," & _
             "��ͥ�绰,��ͥ��ַ,����,����,Ӱ�����,�����Ŀ����,�����Ŀ����,����ҽ��,��������,��ʷ," & _
             "�ٴ����,ע������,��ע From ZLPACS�ӿ�KODAK where ��������=1 or ��������=2"
    Set dsOrder = gcnHIS.Execute(strSQL)
    If dsOrder.EOF Then        'û�����ݣ��˳�����
        Exit Sub
    Else
        dsOrder.MoveFirst
        While Not dsOrder.EOF
            '��ÿһ����¼�Ľ����֯�ɵ������뵥��ͨ��SP_EOrder_For_Kodak���͸��´�RIS
            
            '���ô�����ֵ�Ĵ洢����
    
            adoCmd.CommandText = "ZLHIS.SP_EOrder_For_Kodak"
            adoCmd.CommandType = adCmdStoredProc
            adoCmd.ActiveConnection = gcnPACS
            
            
            adoCmd.Parameters.Append adoCmd.CreateParameter("��������", adInteger, adParamInput, , Nvl(dsOrder!��������))
            adoCmd.Parameters.Append adoCmd.CreateParameter("�������", adInteger, adParamInput, , Nvl(dsOrder!�������))
            adoCmd.Parameters.Append adoCmd.CreateParameter("ҽ��ID", adVarChar, adParamInput, 32, Nvl(dsOrder!ҽ��ID))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��ʶ��", adVarChar, adParamInput, 64, Nvl(dsOrder!��ʶ��))
            adoCmd.Parameters.Append adoCmd.CreateParameter("����ID", adVarChar, adParamInput, 64, Nvl(dsOrder!����ID))
            adoCmd.Parameters.Append adoCmd.CreateParameter("����", adVarChar, adParamInput, 64, Nvl(dsOrder!����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("Ӣ����", adVarChar, adParamInput, 64, Nvl(dsOrder!Ӣ����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("�Ա�", adVarChar, adParamInput, 1, Nvl(dsOrder!�Ա�))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��������", adVarChar, adParamInput, 16, Nvl(dsOrder!��������))
            adoCmd.Parameters.Append adoCmd.CreateParameter("���֤��", adVarChar, adParamInput, 32, Nvl(dsOrder!���֤��))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��ͥ�绰", adVarChar, adParamInput, 128, Nvl(dsOrder!��ͥ�绰))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��ͥ��ַ", adVarChar, adParamInput, 256, Nvl(dsOrder!��ͥ��ַ))
            adoCmd.Parameters.Append adoCmd.CreateParameter("����", adVarChar, adParamInput, 32, Nvl(dsOrder!����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("����", adVarChar, adParamInput, 32, Nvl(dsOrder!����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("Ӱ�����", adVarChar, adParamInput, 128, Nvl(dsOrder!Ӱ�����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("�����Ŀ����", adVarChar, adParamInput, 1024, Nvl(dsOrder!�����Ŀ����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("�����Ŀ����", adVarChar, adParamInput, 1024, Nvl(dsOrder!�����Ŀ����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("����ҽ��", adVarChar, adParamInput, 128, Nvl(dsOrder!����ҽ��))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��������", adVarChar, adParamInput, 128, Nvl(dsOrder!��������))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��ʷ", adVarChar, adParamInput, 1024, Nvl(dsOrder!��ʷ))
            adoCmd.Parameters.Append adoCmd.CreateParameter("�ٴ����", adVarChar, adParamInput, 1024, Nvl(dsOrder!�ٴ����))
            adoCmd.Parameters.Append adoCmd.CreateParameter("ע������", adVarChar, adParamInput, 1024, Nvl(dsOrder!ע������))
            adoCmd.Parameters.Append adoCmd.CreateParameter("��ע", adVarChar, adParamInput, 156, Nvl(dsOrder!��ע))
            
            '����ֵ
           ' Set adoParaReturn = adoCmd.CreateParameter("����ֵ", adVarChar, adParamOutput, 4)
            '����Է��洢����ʹ�õ��Ƿ���ֵ���������������ʹ����һ���ȡ����ֵ
             Set adoParaReturn = adoCmd.CreateParameter("����ֵ", adVarChar, adParamReturnValue, 4)
            adoCmd.Parameters.Append adoParaReturn
            
            adoCmd.Execute
    
            '�жϷ���ֵ 1001��ִ�гɹ�,1002��ִ��ʧ�ܣ���ҽ�����Ѵ���,1003��ִ��ʧ�ܣ�δ֪����
            '�������ʧ�ܣ����¼������־
            If adoParaReturn.Value = "1002" Or adoParaReturn.Value = "1003" Then
                subLogErr adoParaReturn.Value, IIf(adoParaReturn.Value = "1002", "ִ��ʧ�ܣ���ҽ�����Ѵ���", "ִ��ʧ�ܣ�δ֪����") & _
                    " ,�������� = " & Nvl(dsOrder!��������) & " , " & _
                    "ҽ��ID = " & Nvl(dsOrder!ҽ��ID) & " , " & "���� = " & Nvl(dsOrder!����)
            End If
            
            '������ɺ�ɾ���������뵥
            
            strSQL = "delete from ZLPACS�ӿ�KODAK where ID = " & dsOrder!ID
            gcnHIS.Execute (strSQL)
            
            dsOrder.MoveNext
        Wend
    End If
    
    Exit Sub
errLog:
    subLogErr Err.Number, Err.Description
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub subLogErr(lngErrNo As Long, strDesc As String)
    On Error Resume Next
    Dim lngID As Long
    Dim strSQL As String
    Dim dsRecord As ADODB.Recordset
    
    Me.lblStatus.Caption = Date & " " & Time & vbCrLf & " �������󣬴�����룺" & lngErrNo & " ����������" & strDesc
    strSQL = "SELECT MAX(ID) as mID FROM ZLPACS�ӿ�KODAK_ERR"
    Set dsRecord = gcnHIS.Execute(strSQL)
    If Not dsRecord.EOF Then
        lngID = dsRecord!Mid + 1
    End If
 '   strSQL = "insert into ZLPACS�ӿ�KODAK_ERR (ID,�����,��������,����ʱ��) values(" & lngID & "," _
 '           & lngErrNo & ",'" & Replace(strDesc, "'", "''") & "',sysdate)"
 '   gcnHIS.Execute strSQL
     '����Ŀ¼������������򴴽�
     '����洢�����������ڷ�ʽ�洢
    If Dir("D:\ZLPACS�ӿ�KODAK_ERR", vbDirectory) = "" Then
        MkDir "D:\ZLPACS�ӿ�KODAK_ERR\"
    End If
    
    '�����ļ�
    Err = 0
    Open "D:\ZLPACS�ӿ�KODAK_ERR\" & Date & ".txt" For Append As #1
    Print #1, Date & " " & Time & vbCrLf & "ҽ��id:" & lngID & " �������󣬴�����룺" & lngErrNo & " ����������" & strDesc
    Close #1

End Sub


