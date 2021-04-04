VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPreOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ԥ��Ժ"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmPreOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   2
      Top             =   1875
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2475
      TabIndex        =   1
      Top             =   1875
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   1710
      Left            =   120
      TabIndex        =   4
      Top             =   15
      Width           =   4710
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   1845
         TabIndex        =   0
         Top             =   1125
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmPreOut.frx":058A
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblInfo 
         Caption         =   "������ ""XXX"" ��Ԥ��Ժʱ�䣬Ԥ��Ժ֮�󣬲��������Ȩ�޵���Ա�����ٶԲ��˼Ʒѣ�ָ��ʱ��֮����Զ�����Ҳ���ٷ�����"
         ForeColor       =   &H00C00000&
         Height          =   525
         Left            =   855
         TabIndex        =   6
         Top             =   330
         Width           =   3600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ��Ժʱ��"
         Height          =   180
         Left            =   840
         TabIndex        =   5
         Top             =   1185
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   3
      Top             =   1875
      Width           =   1100
   End
End
Attribute VB_Name = "frmPreOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr���� As String
Private mstrPrivs As String
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str���� As String, ByVal strPrivs As String) As Boolean
    Set mfrmParent = frmParent
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr���� = str����
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = mblnOk
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date, dMax As Date
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "��������ȷ��ʱ��ֵ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ��)
    curDate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > curDate Then
        If CDate(txtDate.Text) - curDate > 7 Then
            MsgBox "Ԥ��Ժʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("Ԥ��Ժʱ������˵�ǰϵͳʱ��,ȷʵҪԤ��Ժ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "����Ԥ��Ժʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    strSQL = "zl_���˱䶯��¼_PreOut(" & mlng����ID & "," & mlng��ҳID & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    On Error GoTo 0
    mblnOk = True
    
    If mclsMipModule.IsConnect = True Then
        strSQL = _
                " Select a.����, a.�Ա�, a.סԺ��, a.��ǰ����id, b.���ơ���ǰ����, a.��Ժ����id ��ǰ����id," & _
                        " c.���� ��ǰ����, d.����� ��ǰ����, a.��Ժ���� ��ǰ����, e.Id  �䶯id" & _
                " From ������ҳ a,���˱䶯��¼ e, ��λ״����¼ d, ���ű� b, ���ű� c" & _
                " Where a.����id = e.����id And a.��ҳid = e.��ҳid And a.����id = d.����id(+)  And a.��ǰ����id = d.����id(+) And a.��Ժ���� = d.����(+) " & _
                    " And a.��ǰ����id = b.Id(+) And a.��Ժ����id = c.Id(+) And a.����id = [1] And  a.��ҳid = [2] And e.��ʼԭ�� = [3] And Nvl(e.���Ӵ�λ, 0) = 0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˱䶯��¼", mlng����ID, mlng��ҳID, 10)
        
        mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", Nvl(rsTmp!����), xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", Nvl(rsTmp!�Ա�), xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", Nvl(rsTmp!סԺ��), xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        'out_prehospital     ����Ԥ��Ժ  1
        mclsXML.AppendNode "out_prehospital"
        'change_id       ���id  1   N
        mclsXML.appendData "change_id", Nvl(rsTmp!�䶯id), xsNumber
        'out_date        Ԥ��Ժʱ��  1   s
        mclsXML.appendData "out_date", Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss"), xsString
        'out_area_id     ��ǰ����id  0..1    N
        mclsXML.appendData "out_area_id", Nvl(rsTmp!��ǰ����ID, 0), xsNumber
        'out_area_title      ��ǰ����    0..1    S
        mclsXML.appendData "out_area_title", Nvl(rsTmp!��ǰ����), xsString
        'out_dept_id     ��ǰ����id    1   N
        mclsXML.appendData "out_dept_id", Nvl(rsTmp!��ǰ����id, 0), xsNumber
        'out_dept_title      ��ǰ����  1   S
        mclsXML.appendData "out_dept_title", Nvl(rsTmp!��ǰ����id), xsString
        'out_room        ��ǰ����    0..1    S
        mclsXML.appendData "out_room", Nvl(rsTmp!��ǰ����), xsString
        'out_bed     ��ǰ����    1   S
        mclsXML.appendData "out_bed", Nvl(rsTmp!��ǰ����), xsString
        'order_id        ҽ��id  0..1    N
        mclsXML.AppendNode "out_prehospital", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_009", mclsXML.XmlText
    End If
    
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    mblnOk = False
    '--55791:������,2012-11-13,���ϳ�Ժҽ�����ܳ�����Ժ
    If gblnҽ��������ܳ�Ժ Then
        If Not Checkҽ���´��Ժҽ��(mlng����ID, mlng��ҳID) Then
            MsgBox "ҽ����δ�´��Ժ(��תԺ������)ҽ��������ֱ�ӽ���Ԥ��Ժ������", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    
    lblInfo.Caption = Replace(lblInfo.Caption, "XXX", mstr����)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub txtDate_GotFocus()
    Call zlControl.TxtSelAll(txtDate)
End Sub
