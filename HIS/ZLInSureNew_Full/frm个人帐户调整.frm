VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�����ʻ����� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ʻ�����"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frm�����ʻ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   6405
      Begin VB.TextBox txt���֤�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   18
         TabIndex        =   17
         Top             =   1125
         Width           =   2085
      End
      Begin VB.ComboBox Cbo��Ա��� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   2085
      End
      Begin VB.ComboBox cbo�Ա� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1515
         Width           =   1635
      End
      Begin VB.ComboBox cbo����1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1125
         Width           =   1635
      End
      Begin VB.TextBox txtסԺ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1515
         Width           =   855
      End
      Begin VB.TextBox txtҽ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Left            =   2580
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt�ʻ���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1905
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker dtp��������1 
         Height          =   300
         Left            =   1230
         TabIndex        =   11
         Top             =   1905
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   86245379
         CurrentDate     =   36526
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&I)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   16
         Top             =   1185
         Width           =   990
      End
      Begin VB.Label lbl��λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5760
         TabIndex        =   22
         Top             =   1965
         Width           =   180
      End
      Begin VB.Label lbl��Ա���1 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&K)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   14
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl��������1 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   150
         TabIndex        =   10
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   510
         TabIndex        =   8
         Top             =   1575
         Width           =   630
      End
      Begin VB.Label lblҽ������1 
         AutoSize        =   -1  'True
         Caption         =   "ҽ������(&R)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   510
         TabIndex        =   6
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lblסԺ���� 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����(&S)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   18
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lblҽ���� 
         AutoSize        =   -1  'True
         Caption         =   "ҽ����(&Y)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   330
         TabIndex        =   4
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   510
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl�ʻ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ʻ����(&L)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3090
         TabIndex        =   20
         Top             =   1965
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -270
      TabIndex        =   46
      Top             =   3780
      Width           =   7275
   End
   Begin VB.TextBox txt������ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1110
      TabIndex        =   39
      Top             =   3360
      Width           =   1275
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   3060
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�����ʻ�����.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   705
      Left            =   150
      TabIndex        =   31
      Top             =   2550
      Width           =   6405
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   4170
         MaxLength       =   9
         TabIndex        =   36
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   33
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5760
         TabIndex        =   37
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3450
         TabIndex        =   35
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ԫ"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2910
         TabIndex        =   34
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   32
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   45
      Top             =   3990
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   43
      Top             =   3990
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   44
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   2415
      Left            =   150
      TabIndex        =   23
      Top             =   90
      Width           =   6405
      Begin MSComctlLib.ListView lvw��Ա��� 
         Height          =   1125
         Left            =   1260
         TabIndex        =   30
         Top             =   1140
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1984
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImgLvw"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox cbo����2 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   330
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtp��������2 
         Height          =   300
         Left            =   1260
         TabIndex        =   27
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   86245379
         CurrentDate     =   36526
      End
      Begin VB.Label lbl��Ա���2 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&K)"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������ǰ�����Ĳ���"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   28
         Top             =   780
         Width           =   1620
      End
      Begin VB.Label lbl��������2 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lblҽ������2 
         AutoSize        =   -1  'True
         Caption         =   "ҽ������(&R)"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "�˶�(&T)"
      Height          =   350
      Left            =   2910
      TabIndex        =   42
      Top             =   3990
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt˵�� 
      Height          =   300
      Left            =   3540
      MaxLength       =   200
      TabIndex        =   41
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label lbl˵�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2820
      TabIndex        =   40
      Top             =   3420
      Width           =   630
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   225
      TabIndex        =   38
      Top             =   3420
      Width           =   810
   End
End
Attribute VB_Name = "frm�����ʻ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mint����ģʽ As Integer                     '1-��������;2-��������;3-�޸�;4-����
Private mlng��¼ID As Long                          'ָ�޸ļ�¼��ID
Private mblnOK As Boolean                           '�Ƿ�������ݿ�
Private mblnУ�� As Boolean
Private mblnStart As Boolean

Public Function ShowME(ByVal frmParent As Object, ByVal int����ģʽ As Integer, _
ByVal int���� As Integer, Optional ByVal lng��¼ID As Long = 0) As Boolean
    mblnOK = False
    
    mint����ģʽ = int����ģʽ
    mint���� = int����
    mlng��¼ID = lng��¼ID
    Me.Show 1, frmParent
    ShowME = mblnOK
End Function

Private Function InitFace() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '������������
    
    gstrSQL = "Select ����,���� ID From �Ա�"
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.cbo�Ա�, rsTemp, True)
    Me.cbo�Ա�.ListIndex = 0
    
    gstrSQL = "Select ����,��� ID From ��������Ŀ¼ Where ����=" & mint����
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.cbo����1, rsTemp, True)
    cbo����1.ListIndex = 0
    cbo����2.Clear
    cbo����2.AddItem "����ҽ������"
    cbo����2.ItemData(cbo����2.NewIndex) = 0
    Call zlControl.CboAddData(Me.cbo����2, rsTemp, False)
    Me.cbo����2.ListIndex = 0
    
    gstrSQL = "Select ����,��� ID From ������Ⱥ Where ����=" & mint����
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.Cbo��Ա���, rsTemp, True)
    Cbo��Ա���.ListIndex = 0
    With rsTemp
        .MoveFirst
        lvw��Ա���.ListItems.Clear
        lvw��Ա���.ListItems.Add , "K_0", "������Ա���", , 1
        Do While Not .EOF
            lvw��Ա���.ListItems.Add , "K_" & !ID, !����, , 1
            .MoveNext
        Loop
    End With
    txt������ = gstrUserName
    If mint����ģʽ < 3 Then
        InitFace = True
        Exit Function
    End If
    
    '������޸Ļ���ģ���Ҫ����ԭʼ���ݣ��޸�ʱ���ʻ�������޸�ǰ����
    gstrSQL = "Select B.ID,A.����,A.����,A.ҽ����,C.����ID,C.����,A.����ID, " & _
             " C.�Ա�,C.��������,A.��ְ ��Ա���,Nvl(A.�ʻ����,0) �ʻ����,C.���֤��,A.����֤��, " & _
             " ltrim(to_char(B.���,'900090000.00')) ���,B.������,Nvl(D.סԺ�����ۼ�,0) ��Ժ,Nvl(D.��ԺסԺ����,0) ��Ժ, " & _
             " To_char(B.ʱ��,'yyyy-MM-dd hh24:mi:ss') ʱ��,˵��  " & _
             " From �����ʻ� A,�ʻ��䶯��¼ B,������Ϣ C , " & _
             " (Select * From �ʻ������Ϣ Where ���=to_char(Sysdate,'yyyy')) D " & _
             " Where A.����=B.���� And A.����ID=B.����ID And A.����ID=C.����ID  " & _
             " And A.����=D.����(+) And A.����ID=D.����ID(+) And A.����=" & mint���� & " And B.ID=" & mlng��¼ID
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.EOF Then
        MsgBox "û�ҵ�ָ�����ʻ��䶯��¼�������Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteCons(rsTemp)
    If mint����ģʽ = 3 Then
        InitFace = True
        Exit Function
    End If
    
    Call DisableCons
    InitFace = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function Calc���(ByVal cur�ʻ���� As Currency) As Currency
    '����ʵ�ʵĵ�����
    If Val(txt������.Text) <> 0 Then
        Calc��� = Val(txt������.Text)
    Else
        If Val(txt����.Text) < 0 Then
            Calc��� = Val(cur�ʻ����) * Abs(txt����.Text) / 100 * -1
        Else
            Calc��� = Val(cur�ʻ����) * Abs(txt����.Text) / 100
        End If
    End If
End Function

Private Sub cmdOK_Click()
    Dim lngNextID As Long, cur��� As Currency
    Dim rsAccount As New ADODB.Recordset
    If Not ValidData Then Exit Sub
    If mint����ģʽ = 4 Then
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo errHand
    
    '���������Ϊ���Ҹ���ҲΪ�㣬���˳�
    If Val(txt������) = 0 And Val(txt����) = 0 Then
        MsgBox "�����������򸡶�������", vbInformation, gstrSysName
        txt������.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    Select Case mint����ģʽ
    Case 1
        '��������ҽ�����˵��ʻ����
        cur��� = Calc���(Val(txt�ʻ����.Text))
        lngNextID = zlDatabase.GetNextID("�ʻ��䶯��¼")
        gstrSQL = "ZL_�ʻ��䶯��¼_INSERT(" & _
                  lngNextID & "," & mint���� & ",1," & Val(txt����.Tag) & "," & _
                  cur��� & ",'" & txt������.Text & "','" & txt˵��.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call ����ʻ���Ϣ_����(Val(txt����.Tag), True, False)
    Case 2
        '��������
        gstrSQL = " Select A.����ID,Nvl(A.�ʻ����,0) �ʻ����" & _
                  " From �����ʻ� A,������Ϣ B" & _
                  " Where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & mint���� & GetSQL
        Call OpenRecordset(rsAccount, "ͳ�Ƽ�¼�����Ա����")
        
        Do While Not rsAccount.EOF
            cur��� = Calc���(rsAccount!�ʻ����)
            lngNextID = zlDatabase.GetNextID("�ʻ��䶯��¼")
            gstrSQL = "ZL_�ʻ��䶯��¼_INSERT(" & _
                      lngNextID & "," & mint���� & ",1," & rsAccount!����ID & "," & _
                      cur��� & ",'" & txt������.Text & "','" & txt˵��.Text & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Call ����ʻ���Ϣ_����(rsAccount!����ID, True, False)
            rsAccount.MoveNext
        Loop
    Case 3
        '�޸�
        cur��� = Calc���(Val(txt�ʻ����.Text))
        gstrSQL = "ZL_�ʻ��䶯��¼_UPDATE(" & _
            mlng��¼ID & "," & cur��� & ",'" & txt������.Text & "','" & txt˵��.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call ����ʻ���Ϣ_����(Val(txt����.Tag), True, False)
    End Select
    
    mblnOK = True
    'ֻ�Ե������ʻ����д�ӡ���޸Ļ�ɾ���Ļ��ʻ�����ʱ��ʼ���ʻ������û��Լ��ڹ������ѹ����嵥�����
    If mint����ģʽ = 1 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1607", Me, "��¼ID=" & lngNextID, 1)
    End If
    
    gcnOracle.CommitTrans
    
    '������޸����˳�
    If mint����ģʽ = 3 Then
        Unload Me
        Exit Sub
    End If
    
    'Ϊ����������׼������
    Call ClearAllCons
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdSelect_Click()
    gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
            " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��,A.�ʻ����" & _
            " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
            "  where A.����ID=B.����ID and A.����=" & mint���� & _
            "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)"
    
    Call Get�ʻ����
    zlControl.TxtSelAll txt����
    txt����.SetFocus
End Sub

Private Sub cmdTest_Click()
    Dim strMsg As String
    Dim rsAccount As New ADODB.Recordset
    '��������ִ��ǰ������ǰ�趨��������ͳ�ƹ��ж���ҽ�����˻�������Ա����Ա���жԱ�
    
    'ͳ�ƹ��ж���ҽ�����˵��ʻ��������
    gstrSQL = " Select Count(*) ��¼��" & _
              " From �����ʻ� A,������Ϣ B" & _
              " Where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & mint���� & GetSQL
    Call OpenRecordset(rsAccount, "ͳ�Ƽ�¼�����Ա�˶������Ƿ���ȷ")
    
    If rsAccount!��¼�� = 0 Then
        strMsg = "û�з��������ļ�¼��"
        mblnУ�� = False
        cmdOK.Enabled = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    Else
        strMsg = "����ǰ�趨��������ͳ�Ƴ��� " & rsAccount!��¼�� & " ��ҽ���ʻ����������"
    End If
    
    '�����ֱ������ĵ����ͳ���Ƿ����ڵ���Ϊ�������
    If Val(txt������.Text) < 0 Then
        gstrSQL = " Select Count(*) ��¼��" & _
                  " From �����ʻ� A,������Ϣ B" & _
                  " Where A.����ID=B.����ID And Nvl(A.�Ҷȼ�,0)<>9 And A.����=" & mint���� & GetSQL & _
                  " And Nvl(�ʻ����,0)<" & Val(Abs(txt������))
        Call OpenRecordset(rsAccount, "ͳ�Ƴ��ʻ��������Ϊ���������м�¼")
        If rsAccount!��¼�� <> 0 Then
            strMsg = strMsg & vbCrLf & "��������" & rsAccount!��¼�� & "��ҽ���ʻ����������Ϊ����"
            mblnУ�� = False
            cmdOK.Enabled = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    
    cmdOK.Enabled = True
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Function ValidData() As Boolean
    '����������ݺϷ���
    If mint����ģʽ <> 2 Then
        If Val(txt����.Tag) = 0 Then
            MsgBox "������ҽ�����˵Ŀ��ţ�", vbInformation, gstrSysName
            txt����.SetFocus
            Exit Function
        End If
    End If
    
    If Val(txt����.Text) <> 0 Then
        If Abs(txt����.Text) > 100 Then
            MsgBox "������ܴ���100%��", vbInformation, gstrSysName
            txt����.SetFocus
            Exit Function
        End If
    End If
    If Val(txt������.Text) <> 0 Then
        If Abs(txt������.Text) > 100000000000000# Then
            MsgBox "����������ֵ��", vbInformation, gstrSysName
            txt������.SetFocus
            Exit Function
        End If
        
        '������ܴ����ʻ����
        If mint����ģʽ <> 2 Then
            If Val(txt�ʻ����.Text) + Val(txt������.Text) < 0 Then
                MsgBox "������ܴ����ʻ���", vbInformation, gstrSysName
                txt������.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Trim(txt������.Text) = "" Then
        MsgBox "���������˵�ǰ�û���Ӧ����Ա��", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(txt˵��.Text) > 200 Then
        MsgBox "˵�������ݳ��������100�����ֻ�200���ַ�����", vbInformation, gstrSysName
        txt˵��.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function GetSQL() As String
    Dim strSQL As String, str��Ա��� As String
    Dim intItem As Integer
    Dim bln���� As Boolean
    Dim rs���� As New ADODB.Recordset
    '�����û��趨��SQL��
    
    bln���� = ��������(mint����)
    strSQL = "": str��Ա��� = ""
    If cbo����2.ListIndex <> 0 Then
        If bln���� Then
            strSQL = strSQL & " And A.����=" & cbo����2.ItemData(cbo����2.ListIndex)
        End If
    End If
    If Not IsNull(dtp��������2.Value) Then
        strSQL = strSQL & " And B.��������<to_date('" & Format(dtp��������2.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    End If
    With lvw��Ա���
        If Not .ListItems(1).Checked Then
            For intItem = 2 To .ListItems.Count
                If .ListItems(intItem).Checked Then str��Ա��� = str��Ա��� & IIf(str��Ա��� = "", "", ",") & Mid(.ListItems(intItem).Key, 3)
            Next
        End If
    End With
    If str��Ա��� <> "" Then
        strSQL = strSQL & " And A.��ְ in (" & str��Ա��� & ")"
    End If
    GetSQL = strSQL
End Function

Private Sub Form_Activate()
    If mblnStart = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnStart = False
    fra����.Visible = (mint����ģʽ = 2)
    fra����.Visible = Not (mint����ģʽ = 2)
    cmdTest.Visible = (mint����ģʽ = 2)
    If cmdTest.Visible Then cmdOK.Enabled = False
    
    mblnStart = InitFace
End Sub

Private Sub lvw��Ա���_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim intItems As Integer
    Dim blnState As Boolean
    
    If Item.Key = "K_0" Then
        For intItems = 1 To lvw��Ա���.ListItems.Count
            lvw��Ա���.ListItems(intItems).Checked = Item.Checked
        Next
    Else
        '������µ�ȫ��ѡ���ȫ��δѡ������µ�һ���״̬
        blnState = lvw��Ա���.ListItems(2).Checked             '������һ����Ա���
        For intItems = 2 To lvw��Ա���.ListItems.Count
            If blnState <> lvw��Ա���.ListItems(intItems).Checked Then
                lvw��Ա���.ListItems(1).Checked = False
                Exit Sub
            End If
        Next
        If blnState Then
            lvw��Ա���.ListItems(1).Checked = blnState
        Else
            '����ѡ��һ����Ա���
            lvw��Ա���.ListItems(2).Checked = True
        End If
    End If
End Sub

Private Sub txt������_Change()
    On Error Resume Next
    If Me.ActiveControl.Name = "txt������" Then txt����.Text = ""
End Sub

Private Sub txt������_GotFocus()
    Call zlControl.TxtSelAll(txt������)
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If Not (InStr(1, "0123456789.-", Chr(KeyAscii)) <> 0 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt������_Validate(Cancel As Boolean)
    txt������.Text = Format(txt������.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt����_Change()
    On Error Resume Next
    If Me.ActiveControl.Name = "txt����" Then txt������.Text = ""
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If Not (InStr(1, "0123456789.-", Chr(KeyAscii)) <> 0 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = Format(txt����.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub ClearAllCons()
    Select Case mint����ģʽ
    Case 1
        txt����.Text = ""
        txt����.Tag = ""
        txtҽ����.Text = ""
        txt����.Text = ""
        txt���֤��.Text = ""
        txtסԺ����.Text = "0/0"
        txt�ʻ����.Text = ""
    Case 2
        
    End Select
    
    txt������.Text = ""
    txt����.Text = ""
    txt˵��.Text = ""
End Sub

Private Sub txt����_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Len(txt����.Text) = txt����.MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = Replace(Trim(UCase(txt����.Text)), "'", "")
        If strCode = "" Then Exit Sub
        
        If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then 'ˢ��
            str���� = " and A.����='" & strCode & "' and A.����=" & cbo����1.ItemData(cbo����1.ListIndex)
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
            str���� = " and A.����ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��(��ס(��)Ժ�Ĳ���)
            str���� = " and B.סԺ��=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����(�������ﲡ��)
            str���� = " and B.�����=" & Mid(strCode, 2)
        Else '��������
            str���� = " and A.����='" & strCode & "'"
        End If
    
        gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��,A.�ʻ����" & _
                " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                "  where A.����ID=B.����ID and A.����=" & mint���� & _
                "  and A.����=C.���� and A.����=C.��� And Nvl(A.�Ҷȼ�,0)<>9 and A.����ID=D.ID(+)" & str����
        
        Call Get�ʻ����
    End If
End Sub

Private Sub txt����_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , txt����.Text, "", False, True)
    If Not rs�ʻ� Is Nothing Then
        txt����.Text = rs�ʻ�("����")
        txt����.Tag = rs�ʻ�!ID
        '�������õ�����
        txtҽ����.Text = IIf(IsNull(rs�ʻ�("ҽ����")), "", rs�ʻ�("ҽ����"))
        txt����.Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txt���֤��.Text = IIf(IsNull(rs�ʻ�("���֤��")), "", rs�ʻ�("���֤��"))
        
        Call SetComboByText(cbo�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
        If IsNull(rs�ʻ�("��������")) = False Then
            dtp��������1.Value = rs�ʻ�("��������")
        End If
        
        For lngIndex = 0 To cbo����1.ListCount - 1
            If cbo����1.ItemData(lngIndex) = rs�ʻ�("����ID") Then
                cbo����1.ListIndex = lngIndex
                Exit For
            End If
        Next
        txt�ʻ���� = Format(rs�ʻ�!�ʻ����, "#####0.00;-#####0.00; ;")
        txt�ʻ����.Enabled = False
        
        '�ٶ����ʻ������Ϣ
        gstrSQL = "select * from �ʻ������Ϣ where ����=" & mint���� & _
            " and ����ID=" & rs�ʻ�("ID") & " and ���=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = False Then
            '�����ʻ����
            txtסԺ����.Text = Nvl(rsTemp("סԺ�����ۼ�"), "0") & "/" & Nvl(rsTemp("��ԺסԺ����"), "0")
        Else
            txtסԺ����.Text = "0/0"
        End If
    End If
End Sub

Private Sub WriteCons(ByVal rsObj As ADODB.Recordset)
    Dim cur�ʻ���� As Currency, cur������ As Currency
    
    '������д�����
    txt����.Text = rsObj!����
    txt����.Tag = rsObj!����ID
    txtҽ���� = rsObj!ҽ����
    txt���� = rsObj!����
    Call zlControl.CboLocate(cbo�Ա�, rsObj!�Ա�)
    dtp��������1.Value = Format(rsObj!��������, "yyyy-MM-dd")
    Call zlControl.CboLocate(cbo����1, rsObj!����, True)
    Call zlControl.CboLocate(Cbo��Ա���, rsObj!��Ա���, True)
    txt���֤��.Text = Nvl(rsObj!���֤��)
    txtסԺ����.Text = Nvl(rsObj!��Ժ, "0") & "/" & Nvl(rsObj!��Ժ, 0)
    
    cur�ʻ���� = Nvl(rsObj!�ʻ����, 0)
    cur������ = Nvl(rsObj!���, 0)
    cur�ʻ���� = cur�ʻ���� - cur������
    txt�ʻ����.Text = Format(cur�ʻ����, "#####0.00;-#####0.00; ;")
    txt������.Text = Format(cur������, "#####0.00;-#####0.00; ;")
    txt˵��.Text = Nvl(rsObj!˵��)
End Sub

Private Sub DisableCons()
    txt����.Enabled = False
    cmdSelect.Enabled = False
    txt������.Enabled = False
    txt����.Enabled = False
    txt˵��.Enabled = False
    cmdTest.Visible = False
    cmdOK.Visible = False
    cmdCancel.Caption = "ȷ��(&O)"
End Sub

Private Sub txt˵��_GotFocus()
    Call zlControl.TxtSelAll(txt˵��)
End Sub
