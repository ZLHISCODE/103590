VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmNurse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ļ���ȼ�"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmNurse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   375
      TabIndex        =   11
      Top             =   2415
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   90
      TabIndex        =   12
      Top             =   15
      Width           =   5445
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3390
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   1845
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   690
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1005
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3390
         TabIndex        =   6
         Top             =   1005
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd HH:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1785
         Width           =   4275
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1395
         Width           =   4260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   2610
         TabIndex        =   21
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   375
         TabIndex        =   20
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2760
         TabIndex        =   18
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4110
         TabIndex        =   17
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˴�λ"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Чʱ��"
         Height          =   180
         Left            =   2610
         TabIndex        =   15
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�»���"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1845
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         Height          =   180
         Left            =   375
         TabIndex        =   13
         Top             =   1455
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   10
      Top             =   2415
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2895
      TabIndex        =   9
      Top             =   2415
      Width           =   1100
   End
End
Attribute VB_Name = "frmNurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mblnBed As Boolean
Private mlng����ID As Long, mlng��ҳID As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, strLevel, lng����ID As Long
    
    On Error GoTo errH
    
    gblnOK = False
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    With frmManageCourse
        If mblnBed Then
            txt����.Text = .mrsBeds!����
            txt�Ա�.Text = IIf(IsNull(.mrsBeds!�Ա�), "", .mrsBeds!�Ա�)
            txt����.Text = IIf(IsNull(.mrsBeds!����), "", .mrsBeds!����)
            txtסԺ��.Text = IIf(IsNull(.mrsBeds!סԺ��), "", .mrsBeds!סԺ��)
            txt����.Text = .mrsBeds!��ǰ����
            txtPre.Text = Nvl(.mrsBeds!����ȼ�)
            
            '���ܰ���
            .mrsCBeds.Filter = "����ID=" & .mrsBeds!����ID
            Do While Not .mrsCBeds.EOF
                txt����.Text = txt����.Text & "," & .mrsCBeds!����
                .mrsCBeds.MoveNext
            Loop
            txt����.Text = Mid(txt����.Text, 2)
            
            mlng����ID = .mrsBeds!����ID
            mlng��ҳID = .mrsBeds!��ҳID
            lng����ID = Nvl(.mrsBeds!����ȼ�ID, 0)
        Else
            txt����.Text = .mrsFamily!����
            txt�Ա�.Text = IIf(IsNull(.mrsFamily!�Ա�), "", .mrsFamily!�Ա�)
            txt����.Text = IIf(IsNull(.mrsFamily!����), "", .mrsFamily!����)
            txtסԺ��.Text = IIf(IsNull(.mrsFamily!סԺ��), "", .mrsFamily!סԺ��)
            txt����.Text = .mrsFamily!��ǰ����
            txt����.Text = "��ͥ����"
            txtPre.Text = Nvl(.mrsFamily!����ȼ�)
            
            mlng����ID = .mrsFamily!����ID
            mlng��ҳID = .mrsFamily!��ҳID
            lng����ID = Nvl(.mrsFamily!����ȼ�ID, 0)
        End If
        gstrSQL = "Select ID as ���,����,���� From �շ���ĿĿ¼" & _
            " Where ���='H' And ��Ŀ����>=1 And (����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL) And ID<>[1]  Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    End With
    
'    Set rsTmp = New ADODB.Recordset
'    rsTmp.CursorLocation = adUseClient
'    Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'    rsTmp.Open gstrSQL, gcnOracle, adOpenKeyset
'    Call SQLTest
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboNew.AddItem rsTmp!���� & "-" & rsTmp!����
            cboNew.ItemData(i - 1) = rsTmp!���
            rsTmp.MoveNext
        Next
        cboNew.ListIndex = 0
    Else
        MsgBox "���ܶ�ȡ����ȼ�����,���ȵ�����ȼ����������ã�", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, strSql As String, Curdate As Date
    
    If cboNew.ListIndex = -1 Then
        MsgBox "��ѡ���µĻ���ȼ���", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
    If Not IsDate(txtDate.Text) Then
        MsgBox "������Ϸ�����Чʱ�䣡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "��Чʱ�������ڸò����ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
        
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 30 Then
            MsgBox "��Чʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("��Чʱ������˵�ǰϵͳʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
        
    strSql = "zl_���˱䶯��¼_Nurse(" & mlng����ID & "," & mlng��ҳID & "," & _
        cboNew.ItemData(cboNew.ListIndex) & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    gblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
