VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffUnitMgr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�б굥λ"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmStuffUnitMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.OptionButton optApply 
      Caption         =   "Ӧ����������������"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   17
      Top             =   5040
      Width           =   2055
   End
   Begin VB.OptionButton optApply 
      Caption         =   "Ӧ���ڴ˷��������й��"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   5077
      Width           =   2535
   End
   Begin VB.OptionButton optApply 
      Caption         =   "Ӧ���ڱ�Ʒ�����й��"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
   End
   Begin VB.OptionButton optApply 
      Caption         =   "Ӧ���ڱ����"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   4717
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdStuff 
      Caption         =   "��"
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "����"
      ToolTipText     =   "��*��ѡ����"
      Top             =   818
      Width           =   285
   End
   Begin VB.CheckBox chk�б���� 
      Caption         =   "�б�����(&U)"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1425
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   5430
      TabIndex        =   7
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   210
      Picture         =   "frmStuffUnitMgr.frx":1CFA
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   4320
      TabIndex        =   6
      Top             =   5415
      Width           =   1100
   End
   Begin VB.TextBox txtStuff 
      Height          =   300
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   2
      Top             =   810
      Width           =   4980
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&C)"
      Height          =   350
      Left            =   1410
      Picture         =   "frmStuffUnitMgr.frx":1E44
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   2700
      Picture         =   "frmStuffUnitMgr.frx":1F8E
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5400
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffUnitMgr.frx":20D8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffUnitMgr.frx":2672
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfUnit 
      Height          =   2775
      Left            =   210
      TabIndex        =   5
      Top             =   1695
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4895
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4921
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�б굥λѡ�� 
      Height          =   2565
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmStuffUnitMgr.frx":2C0C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "���      ���ƣ�       ��λ��ƿ"
      Height          =   180
      Left            =   1260
      TabIndex        =   3
      Top             =   1200
      Width           =   2970
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ѡ�����ĺ�ָ�����б굥λ���б��������ʱ���乩Ӧ�̱��������б굥λ"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Width           =   5685
   End
   Begin VB.Label lblStuff 
      AutoSize        =   -1  'True
      Caption         =   "ָ������(&M)"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   870
      Width           =   990
   End
End
Attribute VB_Name = "frmStuffUnitMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lblTag As String
Public frmTag As String
Public strPrivs As String
Dim strTemp As String
Dim objItem As ListItem
Dim rsTemp As New ADODB.Recordset

Private Sub chk�б����_Click()
    msfUnit.Active = (chk�б����.Value = 1)
    If chk�б����.Value = 0 Then
        Call cmdClear_Click
    Else
        Call cmdRestore_Click
    End If
End Sub

Private Sub chk�б����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub cmdClear_Click()
    msfUnit.ClearBill
    msfUnit.TextMatrix(1, 0) = "1"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdStuff_Click()

    err = 0: On Error GoTo ErrHand
    
    gstrSQL = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I" & _
            " where I.���='4'" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "��δ�����������ϵ�����Ϣ��", vbExclamation, gstrSysName
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag: Me.txtStuff.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblStuff.Tag <> !Id Then
                Me.lblStuff.Tag = !Id
                Me.txtStuff.Tag = "[" & !���� & "]" & !����
                Me.txtStuff.Text = Me.txtStuff.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���أ�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtStuff.Name
        .Left = Me.txtStuff.Left
        .Top = Me.txtStuff.Top + Me.txtStuff.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    Dim str��λ As String
    Dim intApply As Integer
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If Val(Me.lblStuff.Tag) = 0 Then
        MsgBox "��δѡ���������ϣ����ܱ��棡", vbInformation, gstrSysName
        txtStuff.SetFocus
        Exit Sub
    End If
    
    For lngRow = 1 To msfUnit.Rows - 1
        If Val(msfUnit.TextMatrix(lngRow, 3)) > 1000000 Then
            MsgBox "��" & lngRow & "�гɱ��۳������ֵ1000000�����ܱ��棡", vbInformation, gstrSysName
            Exit Sub
        End If
    Next
    
    If optApply(0).Value = False Then
        For i = 0 To optApply.UBound
            If optApply(i).Value = True Then
                If MsgBox("�������б굥λӦ�÷�ΧΪ��" & optApply(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    'str��λ��ʽ����λid,�б����|��λid,�б����....
    With msfUnit
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) <> 0 Then
                str��λ = IIf(str��λ = "", "", str��λ & "|") & Val(.RowData(lngRow)) & "," & .TextMatrix(lngRow, 2) & "," & .TextMatrix(lngRow, 3)
            End If
        Next
    End With
    
    For i = 0 To optApply.UBound
        If optApply(i).Value = True Then
            intApply = i
            Exit For
        End If
    Next
    
    gstrSQL = "ZL_�����б굥λ_INSERT(" & Val(lblStuff.Tag) & ",'" & str��λ & "'," & intApply & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Me.cmdClose.Tag = "����" Then
        Me.msfUnit.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
        Me.cmdStuff.Enabled = False
        Me.txtStuff.Enabled = False
        Me.chk�б����.Enabled = False
        
        
    End If
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I" & _
            " where I.���='4' and I.ID=[1]" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblStuff.Tag))
    If rsTemp.State <> 1 Then
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag
    End If
    With rsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag
        Else
            Me.lblStuff.Tag = !Id
            Me.txtStuff.Tag = "[" & !���� & "]" & !����
            Me.txtStuff.Text = Me.txtStuff.Tag
            Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
            Call ShowData
        End If
    End With
    If Me.txtStuff.Enabled Then Me.txtStuff.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If Msf�б굥λѡ��.Visible Then
            Msf�б굥λѡ��.Visible = False
            msfUnit.TxtSetFocus
            Exit Sub
        Else
            cmdClose_Click
        End If
    Case Else
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Tag = frmTag
    Me.lblStuff.Tag = lblTag
    
    With msfUnit
        .Rows = 2
        .Cols = 4
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "��λ����"
        .TextMatrix(0, 2) = "�б����"
        .TextMatrix(0, 3) = "�ɱ���"
        .TextMatrix(1, 0) = "1"
        .ColData(0) = 5
        .ColData(1) = 1
        .ColData(2) = 4
        .ColData(3) = 4
        .ColWidth(0) = 300
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 1000
        .Add , "���", "���", 1200
        .Add , "����", "����", 1200
        .Add , "��λ", "��λ", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call ShowData
End Sub

Private Sub msfUnit_EnterCell(Row As Long, Col As Long)
    With msfUnit
        If Col = 1 Then
            .TextMask = ""
        ElseIf Col = 2 Then
            .TxtCheck = True
            .TextMask = "1234567890"
            .MaxLength = 50
        ElseIf Col = 3 Then
            .TxtCheck = True
            .TextMask = "1234567890."
            .MaxLength = 50
        End If
    End With
End Sub

Private Sub msfUnit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strFind As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    With msfUnit
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then Exit Sub
        If .Text = "" Then Exit Sub
        
        strKey = GetMatchingSting(UCase(.Text))
        strFind = " And (���� Like [1]" & _
                    " Or upper(����) Like [1]" & _
                    " Or ���� Like [1])"
    End With
    
    gstrSQL = " Select ID,����,����,���� From ��Ӧ�� " & _
             " Where ĩ��=1 And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) " & strFind & " Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey)
    With rsTemp
        If .EOF Then
            MsgBox "û���ҵ�ƥ������Ĺ�Ӧ�̣����������룡", vbInformation, gstrSysName
            Cancel = True
            msfUnit.TxtSetFocus
            Exit Sub
        End If
        
        With Msf�б굥λѡ��
            .Clear
            Set .DataSource = rsTemp
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Top = msfUnit.Top + msfUnit.CellTop + msfUnit.MsfObj.CellHeight
            If .Top + .Height > Me.Height Then .Top = msfUnit.Top + msfUnit.CellTop - .Height
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Cancel = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtStuff_GotFocus()
    Me.txtStuff.SelStart = 0: Me.txtStuff.SelLength = 100
End Sub

Private Sub txtStuff_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtStuff.Text))
    If strTemp = "" Then Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    strTemp = GetMatchingSting(strTemp)
    
    gstrSQL = "select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
            " where I.ID=N.�շ�ϸĿID and I.���='4'" & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [1] or N.���� like [1] or N.���� like [1])"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "δ�ҵ�ָ�������������ϣ�������ָ����", vbExclamation, gstrSysName
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag: Me.txtStuff.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblStuff.Tag <> !Id Then
                Me.lblStuff.Tag = !Id
                Me.txtStuff.Tag = "[" & !���� & "]" & !����
                Me.txtStuff.Text = Me.txtStuff.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���أ�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtStuff.Name
        .Left = Me.txtStuff.Left
        .Top = Me.txtStuff.Top + Me.txtStuff.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtStuff_LostFocus()
    Me.txtStuff.Text = Me.txtStuff.Tag
End Sub

Private Sub ShowData()
    '��ʾ�ѳ�ʼ�����б굥λ
    msfUnit.ClearBill
    msfUnit.TextMatrix(1, 0) = "1"
    
    On Error GoTo ErrHandle
    gstrSQL = "Select C.ID,'['||C.����||']'||C.���� ��λ,B.�ɱ���,�б���� From �������� A,�����б굥λ B,��Ӧ�� C" & _
            " Where A.����ID=B.����ID And substr(C.����,5,1)=1 And B.��λID=C.ID And A.����ID=[1]"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(lblStuff.Tag))
    With rsTemp
        If .RecordCount <> 0 Then chk�б����.Value = 1
        Do While Not .EOF
            msfUnit.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            msfUnit.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!��λ), "", !��λ)
            msfUnit.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!�б����), "", !�б����)
            msfUnit.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!�ɱ���), "", !�ɱ���)
            msfUnit.RowData(.AbsolutePosition) = !Id
            If msfUnit.Rows - 1 >= .AbsolutePosition Then msfUnit.Rows = msfUnit.Rows + 1
            .MoveNext
        Loop
        If msfUnit.RowData(msfUnit.Rows - 1) = 0 And msfUnit.Rows > 2 Then msfUnit.Rows = msfUnit.Rows - 1
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If Me.lblStuff.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblStuff.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtStuff.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.txtStuff.Text = Me.txtStuff.Tag
            Call ShowData
        End If
        Me.txtStuff.SetFocus
        Call OS.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfUnit_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    
    '�޸������
    With msfUnit
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfUnit_AfterDeleteRow()
    Dim lngCurRow As Long
    
    '�޸������
    With msfUnit
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfUnit_CommandClick()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID,����,����,���� From ��Ӧ�� Where ĩ��=1 And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) Order By ���� "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    With rsTemp
        If .EOF Then
            MsgBox "���ʼ�����Ĺ�Ӧ�̣���Ӧ�̣���", vbInformation, gstrSysName
            msfUnit.SetFocus
            Exit Sub
        End If
        
        With Msf�б굥λѡ��
            .Clear
            Set .DataSource = rsTemp
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Top = msfUnit.Top + msfUnit.CellTop + msfUnit.MsfObj.CellHeight
            If .Top + .Height > Me.Height Then .Top = msfUnit.Top + msfUnit.CellTop - .Height
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msf�б굥λѡ��_DblClick()
    Dim LngFindReturn As Long, lngRow As Long, lngID As Long
    
    '�ȼ���Ƿ������ͬ���б굥λ���������ֹѡ��
    lngID = Val(Msf�б굥λѡ��.TextMatrix(Msf�б굥λѡ��.Row, 0))
    With msfUnit
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) = lngID Then
                MsgBox "�Ѿ����ڸ��б굥λ��������ѡ��", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
    End With
    
    With msfUnit
        .TextMatrix(.Row, 0) = .Row
        .Text = "[" & Msf�б굥λѡ��.TextMatrix(Msf�б굥λѡ��.Row, 1) & "]" & Msf�б굥λѡ��.TextMatrix(Msf�б굥λѡ��.Row, 2)
        .TextMatrix(.Row, 1) = .Text
        .RowData(.Row) = lngID
    End With
    
    With msfUnit
'        If .Row = .Rows - 1 Then
'            .Rows = .Rows + 1
'            .Row = .Row + 1
'            .TextMatrix(.Row, 0) = .Row
'        End If
        .Col = 2
        .SetFocus
    End With
End Sub

Private Sub msf�б굥λѡ��_GotFocus()
    If Msf�б굥λѡ��.Rows - 1 = 1 Then Call msf�б굥λѡ��_DblClick
End Sub

Private Sub msf�б굥λѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call msf�б굥λѡ��_DblClick
End Sub

Private Sub msf�б굥λѡ��_LostFocus()
    With Msf�б굥λѡ��
        .ZOrder 1
        .Visible = False
    End With
End Sub
