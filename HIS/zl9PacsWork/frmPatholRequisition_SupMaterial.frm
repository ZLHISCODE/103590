VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_SupMaterial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ȡ������"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   Icon            =   "frmPatholRequisition_SupMaterial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   720
      Width           =   3945
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�� ��(&E)"
      Height          =   400
      Left            =   3960
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   160038915
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label labRequestDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽʦ��"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   780
      Width           =   900
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ�䣺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholRequisition_SupMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private mufgParentRequest As ucFlexGrid
Private mufgParentContext As ucFlexGrid

Private mlngPatholAdviceId As Long

Private mfrmOwner As Form


Public blnIsOk As Boolean


Public Function ShowSupMaterialWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, owner As Form) As Boolean
'��ʾ��ȡ�����봰��
    Set mufgParentRequest = ufgParentRequestGrid
    Set mufgParentContext = ufgParentContextGrid
    
    Set mfrmOwner = owner

    mlngPatholAdviceId = lngPatholAdviceId
    
    blnIsOk = False


    dtpRequestTime.value = zlDatabase.Currentdate
    txtRequestDoctor.Text = UserInfo.����

    Call Me.Show(1, owner)
End Function



Private Sub cmdExit_Click()
On Error Resume Next
    blnIsOk = False
    Call Me.Hide
End Sub


Private Sub SaveSupMaterialRequest()
'���油ȡ������
    Dim lngNewRow As Long
    
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    

    '��Ӽ��������Ϣ
    strSql = "select Zl_��������_����([1],[2],[3],[4],[5],[6],[7]) as ����ֵ from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mlngPatholAdviceId, _
                                            txtRequestDoctor.Text, _
                                            CDate(dtpRequestTime.value), _
                                            4, 0, _
                                            0, _
                                            txtDescription.Text)
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "SaveSpeExamRequest", "δ�ɹ���ȡ�����������ID,����ʧ�ܡ�")
        Exit Sub
    End If

    '���ý�����Ϣ
    lngNewRow = mufgParentRequest.NewRow
    
    mufgParentRequest.Text(lngNewRow, gstrRequisition_����ID) = rsData!����ֵ
    mufgParentRequest.Text(lngNewRow, gstrRequisition_������) = txtRequestDoctor.Text
    mufgParentRequest.Text(lngNewRow, gstrRequisition_��������) = "��ȡ��"
    mufgParentRequest.Text(lngNewRow, gstrRequisition_����ʱ��) = dtpRequestTime.value
    mufgParentRequest.Text(lngNewRow, gstrRequisition_��������) = txtDescription.Text
    mufgParentRequest.Text(lngNewRow, gstrRequisition_��ǰ״̬) = "������"
                                            
    
    '��λ��������
    Call mufgParentRequest.LocateRow(lngNewRow)

End Sub




Private Sub cmdSure_Click()
'ȷ������
On Error GoTo errHandle
    
    '�ж�������ϸ�б��Ƿ�Ϊ��Ƭ��Ŀ��ϸ�б�
    If mufgParentContext.GetColIndex(gstrRequest_Material_ȡ��ʱ��) < 0 Then
        mufgParentContext.ColNames = gstrRequest_Material_Cols
        mufgParentContext.ColConvertFormat = gstrRequest_MaterialConvertFormat
        
        '�л�������Ŀ��ƽ���
        Call mfrmOwner.ChangeControlFace(4)
    End If
    
    '����ȡ������
    Call SaveSupMaterialRequest
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
