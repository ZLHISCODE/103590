VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPrintList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ӡ��ϸ��"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "FrmPrintList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView Tvw 
      Height          =   2505
      Left            =   1620
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4419
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTree"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintList.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintList.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPrintList.frx":3B60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.Animation Avi 
      Height          =   1005
      Left            =   4890
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1773
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   85
      FullHeight      =   67
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4980
      TabIndex        =   16
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4980
      TabIndex        =   15
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4980
      TabIndex        =   14
      Top             =   390
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Height          =   2775
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4665
      Begin VB.ComboBox cbo����δ��˵��� 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1920
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker Dtp��ʼ���� 
         Height          =   300
         Left            =   1440
         TabIndex        =   6
         Top             =   1140
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   114491395
         CurrentDate     =   37648
      End
      Begin VB.ComboBox Cbo�ⷿ 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1755
      End
      Begin VB.ComboBox Cbo��λ 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   1755
      End
      Begin VB.CommandButton CmdSelect 
         Caption         =   "��"
         Height          =   300
         Left            =   3810
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2310
         Width           =   300
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1440
         TabIndex        =   12
         Top             =   2310
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker Dtp�������� 
         Height          =   300
         Left            =   1440
         TabIndex        =   8
         Top             =   1530
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   114491395
         CurrentDate     =   37648
      End
      Begin VB.Label lbl����δ��˵��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��δ�󵥾�(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1980
         Width           =   1170
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   7
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Lbl��ʼ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   5
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label lbl6�ⷿ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&K)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   1
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Lbl��λ 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   810
         Width           =   630
      End
      Begin VB.Label Lbl��;���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   780
         TabIndex        =   11
         Top             =   2370
         Width           =   630
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "��ӡ(&P)"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintSET 
         Caption         =   "Ԥ��(&V)"
         Index           =   1
      End
      Begin VB.Menu mnuPrintSET 
         Caption         =   "��ӡ(&P)"
         Index           =   2
      End
      Begin VB.Menu mnuPrintSET 
         Caption         =   "�����&Excel"
         Index           =   3
      End
   End
End
Attribute VB_Name = "FrmPrintList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPath As String
Private mintState As Integer '1=��ʼ��ӡ;2-��ͣ;3-����
Private mintPrint As Integer '��ӡģʽ
Private mblnStart As Boolean
Private mrs���� As New ADODB.Recordset
Private mrs���� As New ADODB.Recordset
Public mstrPrivs As String

Private Sub cmdCancel_Click()
    If CmdCancel.Caption = "�˳�(&X)" Or CmdCancel.Caption = "ȡ��(&C)" Then
        Unload Me
        Exit Sub
    End If
    mintState = 2
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrHand
    If mrs����.State = 0 Then
        If Val(Txt����.Tag) = 0 Then
            MsgBox "��ѡ��һ�����ķ��࣡", vbInformation, gstrSysName
            CmdSelect.SetFocus
            Exit Sub
        End If
        CmdSelect.Enabled = False
        mintPrint = 1
        Call PopupMenu(mnuPrint, 2)
        
        '��ҩƷ��¼��
        gstrSQL = "" & _
            " Select '['||d.����||']'||d.���� as ����,A.����ID " & _
            " From �������� A,������ĿĿ¼ c,�շ���ĿĿ¼ d" & _
            " Where a.����id=c.id and a.����id=d.id And (d.վ��=[2] or d.վ�� is null) " & _
            "       and c.����ID in (Select ID From ���Ʒ���Ŀ¼ where ����=7 start with ID = [1] connect by prior id=�ϼ�id)" & _
            " Order by d.����"
        Set mrs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt����.Tag), gstrNodeNo)
    End If
    
    mintState = 1
    Me.CmdCancel.Caption = "��ͣ(&A)"
    Me.cmdOk.Enabled = False
    On Error Resume Next
    Avi.AutoPlay = True
    Avi.Play
    err = 0
    On Error GoTo ErrHand
    Do While Not mrs����.EOF
        DoEvents
        If mintState = 2 Then
            Me.CmdCancel.Caption = "�˳�(&X)"
            Me.cmdOk.Caption = "����(&P)"
            Me.cmdOk.Enabled = True
            On Error Resume Next
            Avi.AutoPlay = False
            Avi.Open mstrPath & "\�����ļ�\��ӡ.avi"
            Exit Sub
        End If
        
        '��ӡ
        '��󸽼Ӳ���:[0]=����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3=�����Excel
        If cbo�ⷿ.Text = "���пⷿ" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "����=" & mrs����!���� & "|" & mrs����!����ID, "�ⷿ=���пⷿ|is not null", "��λ=" & IIf(cbo��λ.ListIndex = 0, "ɢװ��λ", "��װ��λ") & "|" & cbo��λ.ListIndex, "��ʼ����=" & Format(Me.dtp��ʼ����.Value, "yyyy-MM-DD"), "��������=" & Format(Me.dtp��������.Value, "yyyy-MM-DD"), "����δ��˵���=" & IIf(cbo����δ��˵���.ListIndex = 0, " And 1=1", " And A.����� Is Not NULL"), mintPrint)
        Else
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1721_2", Me, "����=" & mrs����!���� & "|" & mrs����!����ID, "�ⷿ=" & cbo�ⷿ.Text & "|=  " & cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), "��λ=" & IIf(cbo��λ.ListIndex = 0, "ɢװ��λ", "��װ��λ") & "|" & cbo��λ.ListIndex, "��ʼ����=" & Format(Me.dtp��ʼ����.Value, "yyyy-MM-DD"), "��������=" & Format(Me.dtp��������.Value, "yyyy-MM-DD"), "����δ��˵���=" & IIf(cbo����δ��˵���.ListIndex = 0, " And 1=1", " And A.����� Is Not NULL"), mintPrint)
        End If
        
        mrs����.MoveNext
        DoEvents
    Loop
    
    If mrs����.EOF Then
        '��ʾ��ӡ����
        On Error Resume Next
        Avi.AutoPlay = False
        Avi.Open mstrPath & "\�����ļ�\��ӡ.avi"
        err = 0
        Unload Me
        Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelect_Click()
    With tvw
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab): Exit Sub
    If KeyCode = vbKeyEscape Then mintState = 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim str���� As String
    mintState = 0
    mblnStart = False
    mstrPath = gstrAviPath
    Set mrs���� = New ADODB.Recordset
    
    On Error GoTo ErrHandle
    With cbo��λ
        .Clear
        .AddItem "1-ɢװ��λ"
        .AddItem "2-��װ��λ"
        .ListIndex = 0
    End With
    With cbo����δ��˵���
        .Clear
        .AddItem "��δ��˵���"
        .AddItem "����ѯ����˵���"
        .ListIndex = 1
    End With
    
    gstrSQL = "Select distinct a.ID,a.����,a.���� From ���ű� a,��������˵�� b,�������ʷ��� C " & _
             " Where a.id=b.����id And b.��������=c.���� And C.���� In('�Ƽ���','���Ŀ�','���ϲ���','����ⷿ') and (a.վ��=[2] or a.վ�� is null) " & _
             IIf(InStr(1, mstrPrivs, "���пⷿ") <> 0, "", " And A.id In (Select ����ID From ������Ա Where ��ԱID=[1])") & _
             "   and (to_char(a.����ʱ��,'yyyy-mm-dd')='3000-01-01' or a.����ʱ�� is null) "
    Set mrs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, gstrNodeNo)
    
    With mrs����
        If .RecordCount = 0 Then
            MsgBox "���ʼ�����Ŀⷿ�����ķ��ϲ��ţ�[���Ź���]", vbInformation, gstrSysName
            Exit Sub
        End If
        
        cbo�ⷿ.Clear
        If InStr(1, mstrPrivs, "���пⷿ") <> 0 Then cbo�ⷿ.AddItem "���пⷿ"
        Do While Not .EOF
            cbo�ⷿ.AddItem !����
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = !Id
            .MoveNext
        Loop
        cbo�ⷿ.ListIndex = 0
    End With
    
    Me.dtp��ʼ����.Value = Format(DateAdd("m", -1, Sys.Currentdate), "yyyy��MM��dd��")
    Me.dtp��������.Value = Format(Sys.Currentdate, "yyyy��MM��dd��")
    dtp��ʼ����.MaxDate = Format(Sys.Currentdate, "yyyy��MM��dd��")
    dtp��������.MaxDate = Format(Sys.Currentdate, "yyyy��MM��dd��")
    Me.Txt���� = ""
    Me.Txt����.Tag = ""
    
    '��ҩƷ��;����
        gstrSQL = ""
        gstrSQL = "" & _
            "   Select ID,����,����,�ϼ�ID,1 as ĩ��" & _
            "   From ���Ʒ���Ŀ¼ " & _
            "   Where ����=7" & _
            "   Start With �ϼ�ID is null  Connect By Prior ID=�ϼ�ID"
    
        zlDatabase.OpenRecordset mrs����, gstrSQL, Me.Caption
    With mrs����
        If .RecordCount = 0 Then
            MsgBox "���ʼ�����ķ��࣡[����Ŀ¼����]", vbInformation, gstrSysName
            Exit Sub
        End If
        Call LoadTvw
    End With
    
    On Error Resume Next
    With Avi
        .AutoPlay = False
        .Open mstrPath & "\��ӡ.avi"
    End With
    
    mblnStart = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTvw()
    tvw.Nodes.Clear
    If mrs����.RecordCount = 0 Then Exit Sub
    
    With mrs����
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                    tvw.Nodes.Add , , "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
            Else
                If !ĩ�� = 1 Then
                    tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 2, 2
                Else
                    tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !Id, "[" & !���� & "]" & !����, 3, 3
                End If
            End If
            tvw.Nodes("K_" & !Id).Tag = !ĩ��
            .MoveNext
        Loop
    End With
    tvw.Nodes(1).Selected = True
End Sub

Private Sub mnuPrintSET_Click(Index As Integer)
    mintPrint = Index
End Sub

Private Sub tvw_DblClick()
    If tvw.SelectedItem.Tag = -1 Then Exit Sub
    Txt����.Text = tvw.SelectedItem.Text
    Txt����.Tag = Mid(tvw.SelectedItem.Key, 3)
    cmdOk.SetFocus
End Sub

Private Sub Tvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call tvw_DblClick
End Sub

Private Sub Tvw_LostFocus()
    tvw.Visible = False
End Sub
