VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmҽ���������_�ڽ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ������"
   ClientHeight    =   4395
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7470
   Icon            =   "frmҽ���������.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTCDQBM 
      Height          =   300
      ItemData        =   "frmҽ���������.frx":000C
      Left            =   5190
      List            =   "frmҽ���������.frx":000E
      TabIndex        =   23
      Top             =   225
      Width           =   1830
   End
   Begin VB.ComboBox cmbDZLB 
      Height          =   300
      ItemData        =   "frmҽ���������.frx":0010
      Left            =   1680
      List            =   "frmҽ���������.frx":001A
      TabIndex        =   22
      Text            =   "����"
      Top             =   1275
      Width           =   1830
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   375
      Left            =   2155
      TabIndex        =   21
      Top             =   3855
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   135
      TabIndex        =   14
      Top             =   2280
      Width           =   7095
      Begin VB.TextBox txtDZJE 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   19
         Top             =   795
         Width           =   1830
      End
      Begin VB.TextBox txtDZCOUNT 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   795
         Width           =   1830
      End
      Begin VB.TextBox txtDZQK 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   270
         Width           =   5310
      End
      Begin VB.Label Label10 
         Caption         =   "���ʽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3675
         TabIndex        =   20
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "������������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   18
         Top             =   825
         Width           =   1365
      End
      Begin VB.Label Label8 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   300
         Width           =   1365
      End
   End
   Begin VB.TextBox txtJE 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5190
      TabIndex        =   12
      Top             =   1800
      Width           =   1830
   End
   Begin VB.TextBox txtCOUNT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   10
      Top             =   1800
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpKSRQ 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   750
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   78643200
      CurrentDate     =   38646
   End
   Begin VB.TextBox txtHOSPID 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   225
      Width           =   1830
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5805
      TabIndex        =   1
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "����"
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   3855
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpZZRQ 
      Height          =   300
      Left            =   5190
      TabIndex        =   8
      Top             =   750
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   78643200
      CurrentDate     =   38646
   End
   Begin VB.Label Label7 
      Caption         =   "�ϴ��ܶ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   13
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "�ϴ���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   11
      Top             =   1845
      Width           =   1365
   End
   Begin VB.Label Label5 
      Caption         =   "��ֹ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   9
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "��ʼ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   7
      Top             =   780
      Width           =   1365
   End
   Begin VB.Label Label3 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   5
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "ͳ���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3795
      TabIndex        =   4
      Top             =   255
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "ҽԺ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   285
      TabIndex        =   3
      Top             =   255
      Width           =   1365
   End
End
Attribute VB_Name = "frmҽ���������_�ڽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub cmd����_Click()
    Dim StrInput As String, strOutput As String
    Dim strArr
    '������ 20051027
    If txtCOUNT.Text = "" Then Exit Sub
    If ҽ����ʼ��_�ɶ��ڽ� = False Then Exit Sub
    StrInput = txtHOSPID & vbTab & cmbTCDQBM & vbTab & IIf(cmbDZLB.Text = "����", 0, 1) & vbTab & _
               Format(dtpKSRQ, "yyyyMMdd") & vbTab & Format(dtpZZRQ, "yyyyMMdd") & vbTab & _
               Lpad(txtCOUNT * 100, 10, "0") & vbTab & Lpad(txtJE * 100, 10, "0")
    '���ö���
    Call DebugTool("׼�����ö���")
    If ҵ������_�ɶ��ڽ�(���϶���_�ڽ�, StrInput, strOutput) = False Then Exit Sub
    Call DebugTool("���ö��ʽ���")
    strArr = Split(strOutput, vbTab)
    Select Case strArr(0)
        Case 0
            txtDZQK = "0 �ɹ�"
        Case 1
            txtDZQK = "1 ������,��������"
        Case 2
            txtDZQK = "2 ����,�������"
        Case 3
            txtDZQK = "3 ����,��������"
        Case Else
            txtDZQK = strArr(0)
    End Select
    txtDZCOUNT = strArr(1) / 100
    txtDZJE = strArr(2) / 100
    If mblnInit = False Then ҽ����ʼ��_�ɶ��ڽ�
    
    gstrSQL = "ZL_������־_INSERT('" & StrInput & "','" & strOutput & "')"
    ExecuteProcedure_ZLNJ "���������־"
End Sub

Private Sub Form_Load()
    Dim rsTC As New ADODB.Recordset
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_�ɶ��ڽ�
    Call OpenRecordset(rsTC, "ҽԺ����")
    txtHOSPID = Rpad(rsTC!ҽԺ����, 5)
    
    gstrSQL = "Select Distinct substr(����֤��,1,instr(����֤��,'|')-1) As ͳ��������� From �����ʻ� Where ����=" & TYPE_�ɶ��ڽ�
    Call OpenRecordset(rsTC, "ȡ������")
    Do Until rsTC.EOF
        If cmbTCDQBM.Text = "" Then
            cmbTCDQBM.Text = Nvl(rsTC!ͳ���������)
        End If
        cmbTCDQBM.AddItem Nvl(rsTC!ͳ���������)
        rsTC.MoveNext
    Loop
End Sub

Private Sub OKButton_Click()
    Dim rsDz As New ADODB.Recordset
    Dim rsCount As New ADODB.Recordset
    Dim curJE As Currency
    Dim lngCount As Long
    
    If cmbDZLB.Text = "����" Then
        gstrSQL = "Select  B.֧��˳���,sum(nvl(A.��Ԥ��,0)) ��� From ����Ԥ����¼ A,���ս����¼ B,�����ʻ� C,���㷽ʽ D " & _
                  " Where A.���㷽ʽ=D.���� And D.���� between 3 and 4  " & _
                  " And A.����ID=B.��¼ID And B.����=1 And B.����=" & TYPE_�ɶ��ڽ� & _
                  " And B.����ID=C.����ID And substr(C.����֤��,1,instr(C.����֤��,'|')-1)='" & Trim(cmbTCDQBM.Text) & "'" & _
                  " And A.�տ�ʱ�� between to_date('" & Format(dtpKSRQ, "yyyy-MM-dd") & "','YYYY-MM-DD') And to_date('" & _
                  Format(dtpZZRQ + 1, "yyyy-MM-dd") & "','YYYY-MM_DD')" & _
                  " Group by B.֧��˳��� having sum(nvl(A.��Ԥ��,0))<>0"
        '������ 20051027
        Call OpenRecordset(rsDz, "���ս����¼")
        Do Until rsDz.EOF
            curJE = curJE + Nvl(rsDz!���, 0)
            gstrSQL = "Select count(distinct ����ID) as �ϴ���¼�� From ҽ��������Ϣ Where ҽ����ˮ��='" & rsDz!֧��˳��� & "'"
            Call OpenRecordset(rsCount, "�ϴ���¼")
            If rsCount.EOF = False Then
                lngCount = lngCount + rsCount!�ϴ���¼��
            End If
            rsDz.MoveNext
        Loop
    Else
        gstrSQL = "Select  b.��¼id,sum(nvl(A.��Ԥ��,0)) ��� From ����Ԥ����¼ A,���ս����¼ B,�����ʻ� C,���㷽ʽ D " & _
                  " Where A.���㷽ʽ=D.���� And D.����<>'����ӯ��' And D.���� between 3 and 4  " & _
                  " And A.����ID=B.��¼ID And B.����=2 And B.����=" & TYPE_�ɶ��ڽ� & _
                  " And B.����ID=C.����ID And substr(C.����֤��,1,instr(C.����֤��,'|')-1)='" & Trim(cmbTCDQBM.Text) & "'" & _
                  " And A.�տ�ʱ�� between to_date('" & Format(dtpKSRQ, "yyyy-MM-dd") & "','YYYY-MM-DD') And to_date('" & _
                  Format(dtpZZRQ + 1, "yyyy-MM-dd") & "','YYYY-MM_DD')" & _
                  "group by b.��¼id having sum(nvl(A.��Ԥ��,0))<>0"
        Call OpenRecordset(rsDz, "���ս����¼")
        Do Until rsDz.EOF
           curJE = curJE + Nvl(rsDz!���, 0)
           lngCount = lngCount + 1
           rsDz.MoveNext
        Loop
    End If
    
    txtJE = Format(curJE, "0.00")
    txtCOUNT = Format(lngCount, "0.00")
End Sub
