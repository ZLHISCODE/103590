VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeSortEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ѱ�����"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmChargeSortEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra���� 
      Caption         =   "�������"
      Height          =   645
      Left            =   3360
      TabIndex        =   11
      Top             =   150
      Width           =   2595
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "סԺ"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "���ÿ���"
      Height          =   3345
      Left            =   3360
      TabIndex        =   15
      Top             =   960
      Width           =   2595
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   350
         Left            =   1320
         TabIndex        =   19
         Top             =   2820
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   350
         Left            =   180
         TabIndex        =   18
         Top             =   2820
         Width           =   1100
      End
      Begin VB.ListBox lst���� 
         Enabled         =   0   'False
         Height          =   2220
         Left            =   180
         TabIndex        =   17
         Top             =   540
         Width           =   2235
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���п���(&L)"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Value           =   1  'Checked
         Width           =   1305
      End
   End
   Begin VB.Frame frm���� 
      Caption         =   "�������"
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3150
      Begin VB.CheckBox chkȱʡ 
         Caption         =   "ȱʡ(&E)"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   180
         TabIndex        =   30
         Top             =   3240
         Width           =   945
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1560
         TabIndex        =   28
         Top             =   1440
         Width           =   1450
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100663299
         CurrentDate     =   40871
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "���޳���(&F)"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2520
         Width           =   1485
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "��̬����Ŀ(&Y)"
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   2880
         Width           =   1485
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���Ψһ��Ŀ(&I)"
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   2220
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   900
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "����"
         Top             =   1026
         Width           =   2055
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   900
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "����"
         Top             =   648
         Width           =   2055
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "����"
         Top             =   270
         Width           =   645
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1560
         TabIndex        =   29
         Top             =   1785
         Width           =   1450
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   100663299
         CurrentDate     =   40871
      End
      Begin VB.Label lblȱʡ 
         Caption         =   "���ѡ�����������ѱ�ı������Զ�ȡ����"
         Height          =   420
         Left            =   180
         TabIndex        =   31
         Top             =   3600
         Width           =   2880
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   2220
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��Ч��������(&P)"
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   1845
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��Ч��ʼ����(&B)"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   5
         Top             =   1470
         Width           =   1350
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   1
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.TextBox txtEdit 
      Height          =   750
      Index           =   3
      Left            =   1080
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   21
      Tag             =   "˵��"
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -60
      TabIndex        =   22
      Top             =   5400
      Width           =   6270
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   210
      TabIndex        =   25
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3360
      TabIndex        =   23
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   24
      Top             =   5640
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "˵��(&X)"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   630
   End
End
Attribute VB_Name = "frmChargeSortEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    text���� = 2
    Text˵�� = 3
    Text��ʼ = 4
    Text���� = 5
End Enum

Dim mstr���� As String         '��ǰ�༭�ѱ��ԭ����
Dim mblnChange As Boolean      '�Ƿ�ı���
Dim mBoundSelect As Integer    '��ǰ����Χ

Private Sub cmdAdd_Click()
    Dim blnRe  As Boolean, lngIndex As Long
    Dim strID As String, str���� As String, str���� As String, strԭ��ID As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select distinct id,�ϼ�id,����,���� from ���ű� " & _
              "where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD') " & _
              "start with ID In " & _
              "(Select ����ID From ��������˵�� Where ������� In (" & _
              Switch(opt����(0).Value, "1,", opt����(1).Value, "2,", opt����(2).Value, "1,2,") & _
              "3)) connect by prior �ϼ�id=ID Order By ����"
    blnRe = frmTreeSel.ShowTree(gstrSQL, strID, str����, str����, strԭ��ID, "�ѱ����ÿ���", "���п���", False)
    
    If blnRe = True Then
        For lngIndex = 0 To lst����.ListCount - 1
            If lst����.ItemData(lngIndex) = Val(strID) Then
                '�Ѿ��иÿ��ң������ټ���
                Exit Sub
            End If
        Next

        lst����.AddItem "��" & str���� & "��" & str����
        lst����.ItemData(lst����.NewIndex) = strID
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    If lst����.ListIndex < 0 Then Exit Sub
    lst����.RemoveItem lst����.ListIndex
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "��Ա�ȼ�����"
        lbl����.Caption = "��Ա�ȼ�����"
        lblEdit(5).Caption = "�ȼ�˵��(&X)"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save�ѱ�() = False Then Exit Sub
    
    Call frmChargeSortGrade.FillList
    If mstr���� <> "" Then
        '�޸ĳɹ�
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    '��������
    For i = 1 To txtEdit.UBound
        txtEdit(i).Text = ""
    Next
    chkȱʡ.Value = 0
    txtEdit(text����).Text = sys.MaxCode("�ѱ�", "����", 2)
    txtEdit(text����).SetFocus
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
'����:���������йطѱ�������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = txtEdit.LBound To txtEdit.UBound
        strTemp = Trim(txtEdit(i).Text)
        If zlCommFun.StrIsValid(strTemp, txtEdit(i).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(i)
            txtEdit(i).SetFocus
            Exit Function
        End If
    Next
    
    If Trim(txtEdit(text����).Text) = "" Then
        MsgBox "���벻��Ϊ�ա�", vbInformation, gstrSysName
        txtEdit(text����).Text = ""
        txtEdit(text����).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(Text����).Text) = "" Then
        MsgBox "���Ʋ���Ϊ�ա�", vbInformation, gstrSysName
        txtEdit(Text����).Text = ""
        txtEdit(Text����).SetFocus
        Exit Function
    End If
    
    If IsDate(dtpBegin.Value) And IsDate(dtpEnd.Value) Then
        If CDate(dtpBegin.Value) > CDate(dtpEnd.Value) Then
            MsgBox "��Ч�ڵĿ�ʼ���ڲ��ܴ��ڽ������ڡ�", vbInformation, gstrSysName
            dtpBegin.SetFocus
            Exit Function
        End If
    End If
    If chk����.Value = 0 And lst����.ListCount = 0 Then
        If MsgBox("���ѱ�����ÿ��Ҳ������п��ң�����ûѡ��ָ�����ҡ�" & vbCrLf & "�Ƿ������", _
                    vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chk����.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save�ѱ�() As Boolean
'����:����༭�����ݵ��ѱ����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    On Error GoTo ErrHandle
    Dim lngCount As Long
    Dim str��ʼ���� As String, str�������� As String
    Dim strָ������ As String
    
    str��ʼ���� = IIF(dtpBegin.Enabled = False Or IsDate(dtpBegin.Value) = False, _
                    "null", "to_date('" & dtpBegin.Value & "','YYYY-MM-dd')")
    str�������� = IIF(dtpEnd.Enabled = False Or IsDate(dtpEnd.Value) = False, _
                    "null", "to_date('" & dtpEnd.Value & "','YYYY-MM-dd')")
    If lst����.Enabled = True Then
        For lngCount = 0 To lst����.ListCount - 1
            strָ������ = strָ������ & lst����.ItemData(lngCount) & ","
        Next
    End If
    
    gstrSQL = Trim(txtEdit(text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "','" & _
            Trim(txtEdit(text����).Text) & "','" & Trim(txtEdit(Text˵��).Text) & "'," & _
            str��ʼ���� & "," & str�������� & "," & IIF(chk����.Value = 1, 1, 2) & "," & _
            IIF(opt����(0).Value = True, 1, 2) & "," & chk����.Value & "," & chkȱʡ.Value & ",'" & strָ������ & "'," & _
            Switch(opt����(0).Value, 1, opt����(1).Value, 2, opt����(2).Value, 3) & ")"
    If mstr���� = "" Then       '����һ����¼
        gstrSQL = "zl_�ѱ�_Insert('" & gstrSQL
    Else    '�޸�
        gstrSQL = "zl_�ѱ�_update('" & mstr���� & "','" & gstrSQL
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save�ѱ� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭�ѱ�(ByVal str���� As String) As Boolean
'����:��������õķѱ�����ڽ���ͨѶ�ĳ���
'����:str����          ��ǰ�༭�ķѱ�����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rs�ѱ� As New ADODB.Recordset
    Dim i As Integer
    
    mstr���� = str����
    
    On Error GoTo ErrHandle
    If str���� <> "" Then
        rs�ѱ�.CursorLocation = adUseClient
        
        gstrSQL = "Select ����, ����, ����, ��Ч��ʼ, ��Ч����, ���ÿ���, ����, ���޳���, ȱʡ��־, �������, ˵�� From �ѱ� Where ���� =[1] "
        Set rs�ѱ� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����)
                
        txtEdit(text����).Text = rs�ѱ�("����")
        txtEdit(Text����).Text = rs�ѱ�("����")
        txtEdit(text����).Text = IIF(IsNull(rs�ѱ�("����")), "", rs�ѱ�("����"))
        If rs�ѱ�("��Ч��ʼ") & "" <> "" Then
            dtpBegin.Value = CDate(Format(rs�ѱ�("��Ч��ʼ"), "yyyy-MM-dd"))
        Else
            dtpBegin.Value = Null
        End If
        If rs�ѱ�("��Ч����") & "" <> "" Then
            dtpEnd.Value = CDate(Format(rs�ѱ�("��Ч����"), "yyyy-MM-dd"))
        Else
            dtpEnd.Value = Null
        End If
        txtEdit(Text˵��).Text = IIF(IsNull(rs�ѱ�("˵��")), "", rs�ѱ�("˵��"))
        
        opt����(IIF(rs�ѱ�("����") = 2, 1, 0)).Value = True
        
        chk����.Value = IIF(rs�ѱ�("���޳���") = 1, 1, 0)
        opt����(IIF(IsNull(rs�ѱ�("�������")), 2, Val(rs�ѱ�("�������")) - 1)).Value = True
        chkȱʡ.Value = IIF(rs�ѱ�("ȱʡ��־") = 1, 1, 0)
        chk����.Value = IIF(rs�ѱ�("���ÿ���") = 2, 0, 1) '2��ʾָ������
        
        lst����.Clear
        If chk����.Value = 0 Then
            '����ָ������
            gstrSQL = "select B.ID,B.����,B.���� from �ѱ����ÿ��� A,���ű� B " & _
                      " where A.�ѱ�=[1] and A.����ID=B.ID order by B.����"
            Set rs�ѱ� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����)
                        
            Do Until rs�ѱ�.EOF
                lst����.AddItem "��" & rs�ѱ�("����") & "��" & rs�ѱ�("����")
                lst����.ItemData(lst����.NewIndex) = rs�ѱ�("ID")
                
                rs�ѱ�.MoveNext
            Loop
        End If
    Else
        txtEdit(text����).Text = sys.MaxCode("�ѱ�", "����", 2)
        dtpBegin.Value = Null
        dtpEnd.Value = Null
    End If
    Call SetEnable
    
    mblnChange = False
    
    For i = 0 To Me.opt����.Count - 1
        If Me.opt����(i).Value = True Then
            mBoundSelect = i
            Exit For
        End If
    Next
        
    
    frmChargeSortEdit.Show vbModal
    �༭�ѱ� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub opt����_Click(Index As Integer)
    '���ϴ�һ��ʱ�˳�
    If mBoundSelect = Index Then Exit Sub
    If Index = 2 Then
        mBoundSelect = 2
        Exit Sub
    End If
    If Me.lst����.ListCount > 0 Then
        If MsgBox("��ѡ������һ�������������ѡ�еĿ��ҽ���������Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Me.lst����.Clear
            mBoundSelect = Index
        Else
            Me.opt����(mBoundSelect).Value = True
            Me.opt����(mBoundSelect).SetFocus
        End If
    End If
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text���� Then
        txtEdit(text����).Text = zlStr.GetCodeByVB(txtEdit(Text����).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = Text���� Or Index = Text˵�� Then
        OS.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Dim strDate As String
    
    If Index = Text��ʼ Or Index = Text���� Then
        '��������
        strDate = zlCommFun.AddDate(txtEdit(Index).Text)
        If IsDate(strDate) Then
            txtEdit(Index).Text = Format(CDate(strDate), "yyyy-MM-dd")
        Else
            txtEdit(Index).Text = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
          KeyAscii = 0
          SendKeys "{TAB}"
    ElseIf KeyAscii = Asc(":") Or KeyAscii = Asc(",") Then  '����ʵ�ս��ĺ���Zl_Actualmoney���صĴ��õ���:�ŷָ���
        KeyAscii = 0
    End If
End Sub

Private Sub chkȱʡ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chkȱʡ_Click()
    Call SetEnable
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk����_Click()
    Call SetEnable
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub chk����_Click()
    Call SetEnable
End Sub

Private Sub opt����_Click(Index As Integer)
    Call SetEnable
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub SetEnable()
'���ݿؼ�ȡֵ�Ĳ�ͬ��������ؿؼ��Ŀ�����
    Dim bln��ͨ As Boolean
    
    mblnChange = True
    If chkȱʡ.Value = 1 Then
        bln��ͨ = False
        'ֻ��ʹ���ض�ֵ
        chk����.Value = 1
        opt����(0).Value = True
        chk����.Value = 0
    Else
        bln��ͨ = True
    End If
    lblEdit(Text��ʼ).Enabled = bln��ͨ
    lblEdit(Text����).Enabled = bln��ͨ
    dtpBegin.Enabled = bln��ͨ
    dtpEnd.Enabled = bln��ͨ
    opt����(0).Enabled = bln��ͨ
    opt����(1).Enabled = bln��ͨ
    chk����.Enabled = bln��ͨ
    
    If opt����(1).Value = True Then
        chk����.Value = 0
    End If
    chk����.Enabled = bln��ͨ And opt����(0).Value
    
    lst����.Enabled = (chk����.Value = 0)
    cmdAdd.Enabled = lst����.Enabled
    cmdDelete.Enabled = lst����.Enabled
End Sub
