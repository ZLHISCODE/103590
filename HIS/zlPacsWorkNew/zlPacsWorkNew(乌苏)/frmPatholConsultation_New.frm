VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholConsultation_New 
   Caption         =   "��������"
   ClientHeight    =   3420
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   6945
   Icon            =   "frmPatholConsultation_New.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   6945
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ComboBox cbxConsultationDoctor 
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   720
      Width           =   1905
   End
   Begin VB.TextBox txtConsultationUnit 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1905
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3975
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "ȷ ��(&S)"
      Height          =   400
      Left            =   4200
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�� ��(&E)"
      Height          =   400
      Left            =   5520
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cbxConsultationType 
      Height          =   300
      ItemData        =   "frmPatholConsultation_New.frx":179A
      Left            =   1200
      List            =   "frmPatholConsultation_New.frx":179C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1905
   End
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   1905
   End
   Begin VB.TextBox txtDescription 
      Height          =   1095
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker dtpStartTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   154796035
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   300
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   154796035
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ﵥλ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ֹʱ�䣺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   17
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.Label labConsultationDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽʦ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3840
      TabIndex        =   13
      Top             =   780
      Width           =   900
   End
   Begin VB.Label labConsultation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������ͣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   12
      Top             =   300
      Width           =   900
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������ϣ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   1680
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
      Left            =   3840
      TabIndex        =   10
      Top             =   300
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
      TabIndex        =   9
      Top             =   1260
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholConsultation_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentGrid As ucFlexGrid

Private mlngPatholAdviceId As Long
Private mlngCurDepartmentId As Long


Public Function ShowConsultationWindow(ufgParentGrid As ucFlexGrid, ByVal lngPatholAdviceId As Long, _
    ByVal lngDepartmentId As Long, owner As Form) As Boolean
'��ʾ��Ƭ���봰��
    Dim curDate As Date
    
    Set mufgParentGrid = ufgParentGrid
    
    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurDepartmentId = lngDepartmentId

    curDate = zlDatabase.Currentdate

    dtpStartTime.value = curDate
    dtpEndTime.value = Format(curDate + 1, "yyyy-mm-dd 23:59:59")
    txtRequestDoctor.Text = UserInfo.����

    Call CloseProcessHint
    
    '���뵱ǰ���ҵ�ҽʦ
    Call LoadConsultationDoctor(lngDepartmentId)

    Call Me.Show(1, owner)
End Function





Private Sub LoadConsultationDoctor(ByVal lngDepartmentId As Long)
'��ȡ����ҽ������
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select a.���� from ��Ա�� a, ������Ա b where a.id=b.��ԱID and b.����ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDepartmentId)
    
    Call cbxConsultationDoctor.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    Do While Not rsData.EOF
        Call cbxConsultationDoctor.AddItem(rsData!����)
        Call rsData.MoveNext
    Loop
    
End Sub


Private Sub LoadConsultationType()
'�����������
    Call cbxConsultationType.AddItem("0-���ڻ���")
    Call cbxConsultationType.AddItem("1-Ժ�����")
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'��ʾ������Ϣ
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub CloseProcessHint()
'�رմ�����ʾ
    picShow.Visible = False
End Sub


Private Function CheckDataIsValid() As Boolean
'��������Ƿ���Ч
    CheckDataIsValid = True
    
    If cbxConsultationType.Text = "" Then
        CheckDataIsValid = False
        Call ShowProcessHint("��ѡ����ʵĻ������͡�")
        
        cbxConsultationType.SetFocus
        
        Exit Function
    End If
End Function



Private Sub SaveConsultationData()
'�����������
    Dim lngNewRow As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select Zl_�������_����([1],[2],[3],[4],[5],[6],[7],[8]) as ����ֵ from dual"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                            mlngPatholAdviceId, _
                                            txtRequestDoctor.Text, _
                                            txtConsultationUnit.Text, _
                                            cbxConsultationDoctor.Text, _
                                            CDate(dtpStartTime.value), _
                                            CDate(dtpEndTime.value), _
                                            Val(cbxConsultationType.Text), _
                                            txtDescription.Text)
                                            
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "SaveConsultationData", "δ�ɹ���ȡ������Ļ����¼ID,����ʧ�ܡ�")
        Exit Sub
    End If
    
'    mvfgConsultation.mCurFlexGrid.Rows = mvfgConsultation.mCurFlexGrid.Rows + 1
    
    lngNewRow = mufgParentGrid.NewRow
    
    mufgParentGrid.Text(lngNewRow, gstrConsultation_ID) = rsData!����ֵ
    mufgParentGrid.Text(lngNewRow, gstrConsultation_����ҽʦ) = txtRequestDoctor.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_���ﵥλ) = txtConsultationUnit.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_����ҽʦ) = cbxConsultationDoctor.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_����ʱ��) = dtpStartTime.value
    mufgParentGrid.Text(lngNewRow, gstrConsultation_��ֹʱ��) = dtpEndTime.value
    mufgParentGrid.Text(lngNewRow, gstrConsultation_��������) = GetConsultationTypeValue(Val(cbxConsultationType.Text))
    mufgParentGrid.Text(lngNewRow, gstrConsultation_�������) = txtDescription.Text
    mufgParentGrid.Text(lngNewRow, gstrConsultation_��ǰ״̬) = "������"
    
    '��λ��������
    Call mufgParentGrid.LocateRow(lngNewRow)
                                            
End Sub


Private Function GetConsultationTypeValue(ByVal lngConsultationType As Long) As String
'��ȡ��������ȡֵ
    Select Case lngConsultationType
        Case 0:
            GetConsultationTypeValue = "���ڻ���"
        Case 1:
            GetConsultationTypeValue = "Ժ�����"
    End Select

End Function



Private Sub cbxConsultationType_Click()
On Error GoTo errHandle
    txtConsultationUnit.Text = ""
    If cbxConsultationType.Text = "" Then Exit Sub
    
    If Val(cbxConsultationType.Text) = 0 Then txtConsultationUnit.Text = "������"
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
'��ӻ����¼
On Error GoTo errHandle
    If Not CheckDataIsValid Then Exit Sub
    
    '�����������
    Call SaveConsultationData
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    '����д���ﱨ���ʱ����Ҫ��ʾ����ǰ��
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    '�����������
    Call LoadConsultationType
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub
