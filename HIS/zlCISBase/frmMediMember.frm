VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ���"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmMediMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdMedi 
      Caption         =   "��"
      Height          =   285
      Left            =   7440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   788
      Width           =   285
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      Height          =   350
      Left            =   6625
      TabIndex        =   5
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   90
      Picture         =   "frmMediMember.frx":058A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   5535
      TabIndex        =   4
      Top             =   3780
      Width           =   1100
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   1125
      MaxLength       =   50
      TabIndex        =   1
      Top             =   780
      Width           =   6285
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ�����(&C)"
      Height          =   350
      Left            =   1275
      Picture         =   "frmMediMember.frx":06D4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   2565
      Picture         =   "frmMediMember.frx":081E
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   4590
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
            Picture         =   "frmMediMember.frx":0968
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediMember.frx":0F02
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfMember 
      Height          =   2055
      Left            =   90
      TabIndex        =   3
      Top             =   1620
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   3625
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
      Height          =   2505
      Left            =   960
      TabIndex        =   9
      Top             =   4425
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   4419
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
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "���      �����̣�       ��λ��ƿ"
      Height          =   180
      Left            =   1125
      TabIndex        =   11
      Top             =   1110
      Width           =   3150
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmMediMember.frx":149C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ѡ��������ҩƷ����������λָ�����������ҩƷ��δָ������ɣ����������ɣ�����������ΪЭ��ҩƷ��"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   150
      Width           =   7065
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "ָ��ҩƷ(&M)"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblMember 
      AutoSize        =   -1  'True
      Caption         =   "���ҩƷ(&E)��"
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   1395
      Width           =   1170
   End
End
Attribute VB_Name = "frmMediMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ���ʣ���me.tag����,�ֱ�Ϊ"5","6","7"
'   2����ǰ״̬����me.cmdClose.tag���棬�ֱ�Ϊ"�޸�"��"����"�����ϼ�������
'   3��ָ��ҩƷ����me.lblMedi.tag���棬���ϼ���������Դ��ݣ�Ҳ���Բ�����
'   4����ǰ�༭���ݣ���Me.msfMember.Tag���棬�ֱ�Ϊ"Э��"��"����"
'---------------------------------------------------
'Э�����ҩƷֻ��Ϊͬ���ʵ�ҩƷ��
'����ԭ��ҩƷ�Ĳ��ʹ�ϵ:
'   ����ҩ������ԭ��ҩƷֻ��Ϊ������ҩ����ԭ����ҩ��
'   �г�ҩ������ԭ��ҩƷ����Ϊ���г�ҩ���͡��в�ҩ����ԭ��ҩ��
'   �в�ҩ������ԭ��ҩƷֻ��Ϊ���в�ҩ����ԭ��ҩ��
'---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Private Const colƷ�� As Integer = 1
Private Const col��� As Integer = 2
Private Const col���� As Integer = 3
Private Const col������ As Integer = 4
Private Const col��λ As Integer = 5

Private Sub cmdClear_Click()
    Me.msfMember.ClearBill
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlMemberRef(Me.lblMedi.Tag)
End Sub

Private Sub cmdSave_Click()
    If Val(Me.lblMedi.Tag) = 0 Then MsgBox "δ��ȷָ��ҩƷ��", vbExclamation, gstrSysName: Me.txtMedi.SetFocus: Exit Sub
    gstrSql = "": strTemp = ""
    With Me.msfMember
        For intCount = 1 To .Rows - 1
            If .RowData(intCount) <> 0 Then
                If Val(.TextMatrix(intCount, col������)) = 0 Then
                    MsgBox intCount & "�����ҩƷ�Ĳ�����û�����룡", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "��ҩƷ��ǰ�淢���ظ���", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount) & "^" & Val(.TextMatrix(intCount, col������))
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    If Me.msfMember.Tag = "Э��" Then
        gstrSql = "zl_Э��ҩƷ����_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSql & "')"
    Else
        gstrSql = "zl_����ҩƷ����_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSql & "')"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtMedi.Text & Me.msfMember.Tag & "����ɹ���", vbExclamation, gstrSysName
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMedi_Click()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ F" & _
            " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and  I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "��δ��������������ҩƷ��", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
'            If Me.Tag <> "7" Then
                objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
'            End If
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    With Me.msfMember
        .MsfObj.FixedCols = 1: .Cols = 6

        .TextMatrix(0, 0) = "": .TextMatrix(0, colƷ��) = "ҩƷ����"
        .TextMatrix(0, col���) = "���"
        If Me.Tag <> "7" Then
            .TextMatrix(0, col����) = "������"
        Else
            .TextMatrix(0, col����) = "������"
        End If
        .TextMatrix(0, col������) = "������": .TextMatrix(0, col��λ) = "������λ"
        
        .ColAlignment(colƷ��) = 1: .ColAlignment(col���) = 1: .ColAlignment(col����) = 1: .ColAlignment(col��λ) = 7
        
        .ColWidth(0) = 300: .ColWidth(colƷ��) = 2800
        .ColWidth(col���) = 1200: .ColWidth(col����) = 1200: .ColWidth(col������) = 1000: .ColWidth(col��λ) = 800

        .ColData(0) = 5: .ColData(colƷ��) = 1
        .ColData(col���) = 5: .ColData(col����) = 5: .ColData(col������) = 4: .ColData(col��λ) = 5
        
        .PrimaryCol = colƷ��: .LocateCol = colƷ��
        .TextMatrix(1, 0) = "1": .Row = 1: .Col = colƷ��
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 1000
        .Add , "���", "���", 1200
        If Me.Tag <> "7" Then
            .Add , "����", "������", 1200
        Else
            .Add , "����", "������", 1200
        End If
        .Add , "��λ", "������λ", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    
    
    If Me.msfMember.Tag = "Э��" Then
        Me.Caption = "Э��ҩƷ���"
        Me.lblnote.Caption = "    ѡ��������ҩƷ����������λָ�����������ҩƷ��" & _
                "δָ������ɣ����������ɣ�����ҩƷ��������ΪЭ��ҩƷ��"
        Me.lblMember.Caption = "���ҩƷ(&E)��"
    Else
        Me.Caption = "����ҩƷ����"
        Me.lblnote.Caption = "    ѡ��������ҩƷ����������λָ��������ԭ��ҩƷ��" & _
                "δָ����ԭ��ҩƷ�������������ԭ�ϣ�����ҩƷ��������Ϊ����ҩƷ��"
        Me.lblMember.Caption = "ԭ��ҩƷ(&E)��"
    End If
    If Me.cmdClose.Tag = "����" Then
        Me.msfMember.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfMember.Active = True
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ F" & _
            " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and I.���=[1] and I.ID=[2] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !���� & "]" & !����
            Me.txtMedi.Text = Me.txtMedi.Tag
            If Me.Tag <> "7" Then
                Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                    "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                    "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
            Else
                Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                    "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                    "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
            End If
            Call zlMemberRef(Me.lblMedi.Tag)
        End If
    End With
    Me.txtMedi.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        If Me.lvwItems.Tag = Me.txtMedi.Name Then
            Me.txtMedi.SetFocus
        Else
            Me.msfMember.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'
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
        If .Tag = Me.txtMedi.Name Then
            If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1) & _
                        "   �����̣�" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & _
                        "   ������λ��" & .SelectedItem.SubItems(.ColumnHeaders("��λ").Index - 1)
                Else
                    Me.lblSpec.Caption = "���" & .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1) & _
                        "   �����̣�" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & _
                        "   ������λ��" & .SelectedItem.SubItems(.ColumnHeaders("��λ").Index - 1)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Me.txtMedi.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Me.msfMember.RowData(Me.msfMember.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfMember.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, 0) = msfMember.Rows - 1
            Me.msfMember.TextMatrix(Me.msfMember.Row, colƷ��) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col���) = .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col����) = .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col��λ) = .SelectedItem.SubItems(.ColumnHeaders("��λ").Index - 1)
            Me.msfMember.SetFocus
        End If
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

Private Sub msfMember_AfterAddRow(Row As Long)
    With Me.msfMember
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMember_AfterDeleteRow()
    With Me.msfMember
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMember_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfMember.Tag = "Э��" Then
        gstrSql = "select I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
                " from �շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ F" & _
                " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and  I.���=[1] " & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and I.ID<>[2] "
    Else
        gstrSql = "select I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
                " from �շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ F,ҩƷ���� T" & _
                " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and F.ID=T.ҩ��ID" & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and I.ID<>[2] "
        Select Case Me.Tag
        Case "5"
            gstrSql = gstrSql & "      and I.���='5' and T.�Ƿ�ԭ��=1"
        Case "6"
            gstrSql = gstrSql & "      and I.��� in ('6','7') and T.�Ƿ�ԭ��=1"
        Case "7"
            gstrSql = gstrSql & "      and I.���='7' and T.�Ƿ�ԭ��=1"
        End Select
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "��δ��������Ϊ���ҩ���ҩƷ���", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !ID
            Me.msfMember.Text = "[" & !���� & "]" & !����
            Me.msfMember.TextMatrix(Me.msfMember.Row, colƷ��) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col���) = IIf(IsNull(!���), "", !���)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col����) = IIf(IsNull(!����), "", !����)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col��λ) = IIf(IsNull(!��λ), "", !��λ)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
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
        .Tag = Me.msfMember.Name
        .Left = Me.msfMember.Left + 500
        .Top = Me.msfMember.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfMember_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Me.msfMember
        If .TxtVisible = False Then Exit Sub
        If .Col <> 1 Then
            If Trim(.Text) = "" Then
                MsgBox "�������������", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Not IsNumeric(.Text) Then
                MsgBox "�������к��зǷ��ַ���", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Val(.Text) > 10000000 Then
                MsgBox "�������������ֵ��", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Val(.Text) <= 0 Then
                MsgBox "�������������0��", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            .Text = Format(.Text, "0.00000"): .TextMatrix(.Row, col������) = .Text
            Exit Sub
        End If
    End With
    
    strTemp = UCase(Trim(Me.msfMember.Text))
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfMember.Tag = "Э��" Then
        gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
                " from �շ���ĿĿ¼ I,�շ���Ŀ���� N,ҩƷ��� S,������ĿĿ¼ F" & _
                " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and I.ID=N.�շ�ϸĿID and I.���=[1] " & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])" & _
                "       and I.ID<>[4] "
    Else
        gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,F.���㵥λ as ��λ" & _
                " from �շ���ĿĿ¼ I,�շ���Ŀ���� N,ҩƷ��� S,������ĿĿ¼ F,ҩƷ���� T" & _
                " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and I.ID=N.�շ�ϸĿID and F.ID=T.ҩ��ID" & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])" & _
                "       and I.ID<>[4] "
        Select Case Me.Tag
        Case "5"
            gstrSql = gstrSql & "      and I.���='5' and T.�Ƿ�ԭ��=1"
        Case "6"
            gstrSql = gstrSql & "      and I.��� in ('6','7') and T.�Ƿ�ԭ��=1"
        Case "7"
            gstrSql = gstrSql & "      and I.���='7' and T.�Ƿ�ԭ��=1"
        End Select
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%", Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .EOF Then
            MsgBox "δ�ҵ����ҩƷ�����������룡", vbInformation, gstrSysName: Cancel = True: Me.msfMember.TxtSetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !ID
            Me.msfMember.Text = "[" & !���� & "]" & !����
            Me.msfMember.TextMatrix(Me.msfMember.Row, colƷ��) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col���) = IIf(IsNull(!���), "", !���)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col����) = IIf(IsNull(!����), "", !����)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col��λ) = IIf(IsNull(!��λ), "", !��λ)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
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
        .Tag = Me.msfMember.Name
        .Left = Me.msfMember.Left + 500
        .Top = Me.msfMember.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True: Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,�շ���Ŀ���� N,ҩƷ��� S,������ĿĿ¼ F" & _
            " where I.ID=S.ҩƷID and S.ҩ��ID=F.ID and I.ID=N.�շ�ϸĿID and I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ������ҩƷ��������ָ����", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   �����̣�" & IIf(IsNull(!����), "", !����) & _
                        "   ������λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
'            If Me.Tag <> "7" Then
                objItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
'            End If
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub zlMemberRef(lngMediId As Long)
    '--------------------------------------------------------
    '���ܣ�ˢ��ָ��ҩƷ��Э�����ҩƷ��ԭ��ҩƷ
    '��Σ�lngMediId-ָ����ҩ��id
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand

    If Me.msfMember.Tag = "Э��" Then
        gstrSql = "select I.ID,I.����,I.����,I.���,I.����,M.���㵥λ as ��λ,P.���� as ������" & _
                " from Э��ҩƷ���� P,�շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ M" & _
                " where P.Э��ҩƷID=I.ID and I.ID=S.ҩƷID and S.ҩ��id=M.ID" & _
                "       and P.ҩƷID=[1]"
    Else
        gstrSql = "select I.ID,I.����,I.����,I.���,I.����,M.���㵥λ as ��λ,P.���� as ������" & _
                " from ����ҩƷ���� P,�շ���ĿĿ¼ I,ҩƷ��� S,������ĿĿ¼ M" & _
                " where P.ԭ��ҩƷID=I.ID and I.ID=S.ҩƷID and S.ҩ��id=M.ID" & _
                "       and P.����ҩƷID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        Me.msfMember.ClearBill
        Do While Not .EOF
            If Me.msfMember.Rows < .AbsolutePosition + 1 Then Me.msfMember.Rows = Me.msfMember.Rows + 1
            Me.msfMember.RowData(.AbsolutePosition) = !ID
            Me.msfMember.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfMember.TextMatrix(.AbsolutePosition, colƷ��) = "[" & !���� & "]" & !����
            Me.msfMember.TextMatrix(.AbsolutePosition, col���) = IIf(IsNull(!���), "", !���)
            Me.msfMember.TextMatrix(.AbsolutePosition, col����) = IIf(IsNull(!����), "", !����)
            Me.msfMember.TextMatrix(.AbsolutePosition, col������) = Format(!������, "0.00000")
            Me.msfMember.TextMatrix(.AbsolutePosition, col��λ) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

