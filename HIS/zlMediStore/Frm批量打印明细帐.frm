VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm������ӡ��ϸ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ӡ��ϸ��"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Frm������ӡ��ϸ��.frx":0000
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
            Picture         =   "Frm������ӡ��ϸ��.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm������ӡ��ϸ��.frx":1E56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm������ӡ��ϸ��.frx":3B60
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
         Format          =   275709955
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
      Begin VB.TextBox Txt��;���� 
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
         Format          =   275709955
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
         Caption         =   "��;����(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   420
         TabIndex        =   11
         Top             =   2370
         Width           =   990
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
Attribute VB_Name = "Frm������ӡ��ϸ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strPath As String
Private intState As Integer '1=��ʼ��ӡ;2-��ͣ;3-����
Private intPrint As Integer '��ӡģʽ
Private blnStart As Boolean
Private rs��;���� As New ADODB.Recordset
Private rsҩƷ As New ADODB.Recordset

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "�˳�(&X)" Or cmdCancel.Caption = "ȡ��(&C)" Then
        Unload Me
        Exit Sub
    End If
    intState = 2
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHand
    If rsҩƷ.State = 0 Then
        If Val(Txt��;����.Tag) = 0 Then
            MsgBox "��ѡ��һ��ҩƷ��;���࣡", vbInformation, gstrSysName
            CmdSelect.SetFocus
            Exit Sub
        End If
        CmdSelect.Enabled = False
        intPrint = 1
        Call PopupMenu(mnuPrint, 2)
        
        '��ҩƷ��¼��
        gstrSQL = "Select '['||F.����||']'||F.���� ����,A.ҩƷID  " & _
                 " From ҩƷ��� A,�շ���ĿĿ¼ F, " & _
                 "  (Select ID ҩ��ID  " & _
                 "  From ������ĿĿ¼   " & _
                 "  Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And ����ID IN  " & _
                 "      (Select ID From ���Ʒ���Ŀ¼ " & _
                 "      where ���� in (1,2,3) " & _
                 "      Start With ID=[1] Connect By Prior ID=�ϼ�ID)) B " & _
                 " Where (F.վ�� = [2] Or f.վ�� is Null) And A.ҩ��ID=B.ҩ��ID And A.ҩƷID=F.ID " & _
                 " Order by F.����"
        Set rsҩƷ = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Txt��;����.Tag), gstrNodeNo)
    End If
    
    intState = 1
    Me.cmdCancel.Caption = "��ͣ(&A)"
    Me.cmdOK.Enabled = False
    On Error Resume Next
    Avi.AutoPlay = True
    Avi.Play
    Err = 0
    On Error GoTo ErrHand
    Do While Not rsҩƷ.EOF
        DoEvents
        If intState = 2 Then
            Me.cmdCancel.Caption = "�˳�(&X)"
            Me.cmdOK.Caption = "����(&P)"
            Me.cmdOK.Enabled = True
            On Error Resume Next
            Avi.AutoPlay = False
            Avi.Open strPath & "\�����ļ�\��ӡ.avi"
            Exit Sub
        End If
        
        '��ӡ
        '��󸽼Ӳ���:[0]=����(������Ԥ��),1=ֱ�ӵ�Ԥ��,2=ֱ�Ӵ�ӡ,3=�����Excel
        If cbo�ⷿ.Text = "���пⷿ" Then
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "ҩƷ=" & rsҩƷ!���� & "|" & rsҩƷ!ҩƷid, "�ⷿ=���пⷿ|is not null", "��λ=" & Choose(cbo��λ.ListIndex + 1, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(cbo��λ.ListIndex + 1, 1, 3, 2, 4), "��ʼ����=" & Format(Me.dtp��ʼ����.Value, "yyyy-MM-DD"), "��������=" & Format(Me.dtp��������.Value, "yyyy-MM-DD"), "����δ��˵���=" & IIf(cbo����δ��˵���.ListIndex = 0, " And 1=1", " And A.����� Is Not NULL"), intPrint)
        Else
            Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "ҩƷ=" & rsҩƷ!���� & "|" & rsҩƷ!ҩƷid, "�ⷿ=" & cbo�ⷿ.Text & "|=  " & cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex), "��λ=" & Choose(cbo��λ.ListIndex + 1, "�ۼ۵�λ", "���ﵥλ", "ҩ�ⵥλ", "סԺ��λ") & "|" & Choose(cbo��λ.ListIndex + 1, 1, 3, 2, 4), "��ʼ����=" & Format(Me.dtp��ʼ����.Value, "yyyy-MM-DD"), "��������=" & Format(Me.dtp��������.Value, "yyyy-MM-DD"), "����δ��˵���=" & IIf(cbo����δ��˵���.ListIndex = 0, " And 1=1", " And A.����� Is Not NULL"), intPrint)
        End If
        
        rsҩƷ.MoveNext
        DoEvents
    Loop
    
    If rsҩƷ.EOF Then
        '��ʾ��ӡ����
        On Error Resume Next
        Avi.AutoPlay = False
        Avi.Open strPath & "\�����ļ�\��ӡ.avi"
        Err = 0
        Unload Me
        Exit Sub
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CmdSelect_Click()
    With Tvw
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub Form_Activate()
    If Not blnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab): Exit Sub
    If KeyCode = vbKeyEscape Then intState = 2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Str���� As String
    intState = 0
    blnStart = False
    strPath = gstrAviPath
    Set rsҩƷ = New ADODB.Recordset
    
    On Error GoTo errHandle
    With cbo��λ
        .Clear
        .AddItem "1-�ۼ۵�λ"
        .AddItem "2-���ﵥλ"
        .AddItem "3-ҩ�ⵥλ"
        .AddItem "4-סԺ��λ"
        .ListIndex = 0
    End With
    With cbo����δ��˵���
        .Clear
        .AddItem "��δ��˵���"
        .AddItem "����ѯ����˵���"
        .ListIndex = 1
    End With
    
    gstrSQL = "Select distinct a.ID,a.����,a.���� From ���ű� a,��������˵�� b,�������ʷ��� C " & _
              "Where (a.վ�� = [2] Or a.վ�� is Null) And a.id=b.����id And b.��������=c.���� And Instr('HIJKLMN',c.����,1)>0 " & _
              IIf(zlStr.IsHavePrivs(gstrStockSearchPrivs, "���пⷿ"), "", " And A.id In (Select ����ID From ������Ա Where ��ԱID=[1]) ") & _
              "  and (to_char(a.����ʱ��,'yyyy-mm-dd')='3000-01-01' or a.����ʱ�� is null) "
    Set rs��;���� = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.�û�ID, gstrNodeNo)
       
    With rs��;����
        If .RecordCount = 0 Then
            MsgBox "���ʼ��ҩƷ�ⷿ��[���Ź���]", vbInformation, gstrSysName
            Exit Sub
        End If
        
        cbo�ⷿ.Clear
        If zlStr.IsHavePrivs(gstrStockSearchPrivs, "���пⷿ") Then cbo�ⷿ.AddItem "���пⷿ"
        Do While Not .EOF
            cbo�ⷿ.AddItem !����
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = !id
            .MoveNext
        Loop
        cbo�ⷿ.ListIndex = 0
    End With
    
    Me.dtp��ʼ����.Value = Format(DateAdd("m", -1, Sys.Currentdate), "yyyy��MM��dd��")
    Me.dtp��������.Value = Format(Sys.Currentdate, "yyyy��MM��dd��")
    dtp��ʼ����.MaxDate = Format(Sys.Currentdate, "yyyy��MM��dd��")
    dtp��������.MaxDate = Format(Sys.Currentdate, "yyyy��MM��dd��")
    Me.Txt��;���� = ""
    Me.Txt��;����.Tag = ""
    
    '��ҩƷ��;����
    gstrSQL = "": Str���� = ""
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "����ҩ") Then
        Str���� = Str���� & ",'����ҩ'"
        gstrSQL = gstrSQL & _
                  " Select -1 id,'1' ����,'����ҩ' ����,to_number(NULL,0) �ϼ�ID,-1 ĩ�� from dual" & _
                  " Union all"
    End If
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "�г�ҩ") Then
        Str���� = Str���� & ",'�г�ҩ'"
        gstrSQL = gstrSQL & _
                " Select -2 id,'3' ����,'�г�ҩ' ����,to_number(NULL,0) �ϼ�ID,-1 ĩ�� from dual" & _
                " Union all"
    End If
    If zlStr.IsHavePrivs(gstrStockSearchPrivs, "�в�ҩ") Then
        Str���� = Str���� & ",'�в�ҩ'"
        gstrSQL = gstrSQL & _
                " Select -3 id,'2' ����,'�в�ҩ' ����,to_number(NULL,0) �ϼ�ID,-1 ĩ�� from dual" & _
                " Union all"
    End If
    If Str���� = "" Then
        MsgBox "��û��Ȩ��ʹ��������ӡ��ϸ�ʣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Str���� = Mid(Str����, 2)
    gstrSQL = gstrSQL & _
        " Select ID,����,����,Decode(�ϼ�ID,null,-1*����,Nvl(�ϼ�ID,0)) �ϼ�ID,1" & _
        " From ���Ʒ���Ŀ¼ A" & _
        " Where A.���� In (1,2,3)" & _
        " Start With Nvl(�ϼ�ID,0)=0" & _
        " Connect By Prior ID=�ϼ�ID"
    Set rs��;���� = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-ҩƷ��;����")
    
    With rs��;����
        If .RecordCount = 0 Then
            MsgBox "���ʼ��ҩƷ��;���࣡[ҩƷ��;����]", vbInformation, gstrSysName
            Exit Sub
        End If
        Call LoadTvw
    End With
    
    On Error Resume Next
    With Avi
        .AutoPlay = False
        .Open strPath & "\��ӡ.avi"
    End With
    
    blnStart = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTvw()
    Tvw.Nodes.Clear
    If rs��;����.RecordCount = 0 Then Exit Sub
    
    With rs��;����
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                If !ĩ�� = 1 Then
                    Tvw.Nodes.Add , , "K_" & !id, "[" & !���� & "]" & !����, 2, 2
                Else
                    Tvw.Nodes.Add , , "K_" & !id, "[" & !���� & "]" & !����, 3, 3
                End If
            Else
                If !ĩ�� = 1 Then
                    Tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !id, "[" & !���� & "]" & !����, 2, 2
                Else
                    Tvw.Nodes.Add "K_" & !�ϼ�ID, 4, "K_" & !id, "[" & !���� & "]" & !����, 3, 3
                End If
            End If
            Tvw.Nodes("K_" & !id).Tag = !ĩ��
            .MoveNext
        Loop
    End With
    Tvw.Nodes(1).Selected = True
End Sub

Private Sub mnuPrintSET_Click(Index As Integer)
    intPrint = Index
End Sub

Private Sub Tvw_DblClick()
    If Tvw.SelectedItem.Tag = -1 Then Exit Sub
    Txt��;����.Text = Tvw.SelectedItem.Text
    Txt��;����.Tag = Mid(Tvw.SelectedItem.Key, 3)
    cmdOK.SetFocus
End Sub

Private Sub Tvw_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call Tvw_DblClick
End Sub

Private Sub Tvw_LostFocus()
    Tvw.Visible = False
End Sub


