VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCriticalValues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Σ����������"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   FillStyle       =   0  'Solid
   Icon            =   "frmCriticalValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtTechnical 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1140
      Width           =   2175
   End
   Begin VB.Frame frmControl 
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2520
      TabIndex        =   15
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "ȷ��(&S)"
      Height          =   350
      Left            =   1200
      TabIndex        =   14
      Top             =   4560
      Width           =   1100
   End
   Begin VB.TextBox txtRisTime 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3780
      Width           =   2175
   End
   Begin VB.TextBox txtSubscriber 
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3300
      Width           =   2175
   End
   Begin VB.TextBox txtResult 
      Height          =   1020
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2100
      Width           =   2175
   End
   Begin VB.ComboBox cboNotifyStaff 
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   1620
      Width           =   2175
   End
   Begin VB.ComboBox cboStyle 
      Height          =   300
      ItemData        =   "frmCriticalValues.frx":179A
      Left            =   1440
      List            =   "frmCriticalValues.frx":17AA
      TabIndex        =   8
      Top             =   660
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   180
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483643
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   130023427
      CurrentDate     =   38222
   End
   Begin VB.Label lblSubscriber 
      AutoSize        =   -1  'True
      Caption         =   "�� �� ��"
      Height          =   180
      Left            =   360
      TabIndex        =   5
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   2120
      Width           =   720
   End
   Begin VB.Label lblStyle 
      AutoSize        =   -1  'True
      Caption         =   "֪ͨ��ʽ"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblNotifyStaff 
      AutoSize        =   -1  'True
      Caption         =   "������Ա"
      Height          =   180
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lblTechnical 
      AutoSize        =   -1  'True
      Caption         =   "���ܿ���"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label lblRisTime 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ�ʱ��"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   3780
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "֪ͨʱ��"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmCriticalValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mlngAdviceID As Long
Public mblnSave As Boolean


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String
    
    On Error GoTo ErrorHand
    'Ϊ���ж�
    If cboStyle.Text = "" Then
        MsgBoxD Me, "֪ͨ��ʽ����Ϊ��!", vbExclamation, Me.Caption
        cboStyle.SetFocus
        Exit Sub
    End If
    
    If txtTechnical.Text = "" Then
        MsgBoxD Me, "���ܿ��Ҳ���Ϊ��!", vbExclamation, Me.Caption
        txtTechnical.SetFocus
        Exit Sub
    End If
    
    If cboNotifyStaff.Text = "" Then
        MsgBoxD Me, "������Ա����Ϊ��!", vbExclamation, Me.Caption
        cboNotifyStaff.SetFocus
        Exit Sub
    End If
    
    strSQL = "Zl_Ӱ��Σ��ֵ��¼_�Ǽ�(" & mlngAdviceID & "," & _
                                    To_Date(dtpTime) & ",'" & _
                                    cboStyle.Text & "','" & _
                                    txtTechnical.Text & "','" & _
                                    cboNotifyStaff.Text & "','" & _
                                    Nvl(txtResult.Text, "") & "','" & _
                                    txtSubscriber.Text & "'," & _
                                    To_Date(txtRisTime.Text) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    mblnSave = True
    Unload Me
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(lngAdviceID As Long, Optional owner As Form = Nothing)
'��ʾ����
    mlngAdviceID = lngAdviceID
    mblnSave = False
    Call Me.Show(1, owner)
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim dtServicesTime As Date
    
    On Error GoTo ErrorHand
    dtServicesTime = zlDatabase.Currentdate
    
    txtRisTime.Text = Format(dtServicesTime, "yyyy-mm-dd hh:mm:ss")
    dtpTime.value = dtServicesTime
    txtSubscriber.Text = UserInfo.����
    
    If cboStyle.ListCount > 0 And cboStyle.ListIndex = -1 Then cboStyle.ListIndex = 0

    strSQL = "select a.id,a.����,a.����,b.����ҽ�� " & _
             "from ���ű� a ,(select ��������id,����ҽ�� from ����ҽ����¼ where id=[1]) b " & _
             "where a.id=b.��������id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
    
    If Not rsTemp.EOF Then
        txtTechnical.Text = rsTemp!���� & "-" & rsTemp!����
        InitDoctors rsTemp!ID, rsTemp!����ҽ��
    End If
    
    strSQL = "select ֪ͨʱ��,֪ͨ��ʽ,���ܿ���,������Ա,������,�Ǽ���,�Ǽ�ʱ�� from Ӱ��Σ��ֵ��¼ where ҽ��id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)
    '����Σ����Ϣ���ȡΣ����Ϣ
    If rsTemp.RecordCount > 0 Then
        dtpTime.value = Nvl(rsTemp!֪ͨʱ��, dtServicesTime)
        cboStyle.Text = Nvl(rsTemp!֪ͨ��ʽ)
        txtTechnical.Text = Nvl(rsTemp!���ܿ���)
        cboNotifyStaff.Text = Nvl(rsTemp!������Ա)
        txtResult.Text = Nvl(rsTemp!������)
        txtSubscriber.Text = Nvl(rsTemp!�Ǽ���, UserInfo.����)
        txtRisTime.Text = Format(Nvl(rsTemp!�Ǽ�ʱ��, dtServicesTime), "yyyy-mm-dd hh:mm:ss")
    End If
    
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitDoctors(ByVal lng����ID As Long, ByVal strAdviceName As String)
'���ܣ���ȡ��ǰ�����а�����������Ա
'lng����ID:��������id
'strAdviceName:����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrorHand
    strSQL = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b, ��Ա����˵�� c" & vbNewLine & _
                " Where a.��Աid = b.Id And b.Id = c.��Աid And c.��Ա���� = 'ҽ��' And" & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and a.����id = [1] " & vbNewLine & _
                " Order By ���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    cboNotifyStaff.Clear
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cboNotifyStaff.AddItem rsTmp!����
            If rsTmp!���� = strAdviceName Then cboNotifyStaff.ListIndex = cboNotifyStaff.NewIndex
            rsTmp.MoveNext
        Loop
        If cboNotifyStaff.ListCount > 0 And cboNotifyStaff.ListIndex = -1 Then cboNotifyStaff.ListIndex = 0
        cboNotifyStaff.Enabled = True
    End If
    Exit Sub
ErrorHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtResult_GotFocus()
    Call GetFocus(txtResult)
End Sub

Private Sub txtTechnical_GotFocus()
    Call GetFocus(txtTechnical)
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub
