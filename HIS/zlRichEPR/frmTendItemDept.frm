VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemDept 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ÿ���"
   ClientHeight    =   5625
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5025
   Icon            =   "frmTendItemDept.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Index           =   0
      Left            =   585
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ȫ��(&E)"
      Height          =   350
      Index           =   1
      Left            =   1665
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4665
      Width           =   1100
   End
   Begin VB.OptionButton optApply 
      Caption         =   "�ݲ�ʹ��(&0)"
      Height          =   195
      Index           =   0
      Left            =   585
      TabIndex        =   8
      Top             =   1305
      Value           =   -1  'True
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "ȫԺͨ����Ŀ(&1)"
      Height          =   195
      Index           =   1
      Left            =   585
      TabIndex        =   7
      Top             =   1590
      Width           =   1950
   End
   Begin VB.OptionButton optApply 
      Caption         =   "���������²���(&2)"
      Height          =   195
      Index           =   2
      Left            =   585
      TabIndex        =   6
      Top             =   1890
      Width           =   1950
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   -45
      TabIndex        =   5
      Top             =   540
      Width           =   5115
   End
   Begin VB.CheckBox chkSelect 
      Caption         =   "����ʾѡ����(&L)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   2925
      TabIndex        =   4
      Top             =   4740
      Width           =   1830
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   3
      Top             =   5085
      Width           =   5115
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2460
      TabIndex        =   2
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   1
      Top             =   5190
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwBakup 
      Height          =   2475
      Left            =   -840
      TabIndex        =   0
      Tag             =   "10"
      Top             =   2175
      Visible         =   0   'False
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwApply 
      Height          =   2475
      Left            =   585
      TabIndex        =   9
      Tag             =   "10"
      Top             =   2190
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "���Ը���ҽѧרҵ�Ĳ�ͬҪ��ָ������Ŀ�����ڲ��ֲ��Ż�ȫԺͨ�á�"
      Height          =   360
      Left            =   795
      TabIndex        =   14
      Top             =   75
      Width           =   3960
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmTendItemDept.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����:   ����"
      Height          =   180
      Left            =   270
      TabIndex        =   13
      Top             =   705
      Width           =   1440
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      Caption         =   "ʹ�÷�Χ(&S)"
      Height          =   180
      Left            =   270
      TabIndex        =   12
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmTendItemDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mintKind As Integer       '��������
Private mlngFileId As Long        '�����ļ�ID
Private mblnOK As Boolean

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim lngCount As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileId As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    mlngFileId = lngFileId
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ��Ŀ����, ���ÿ��� From �����¼��Ŀ Where ��Ŀ��� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "�����¼��Ŀ��ʧ(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Exit Function
        lblFile.Caption = "�����¼��Ŀ:   " & !��Ŀ����
        optApply(IIf(IsNull(!���ÿ���), 0, !���ÿ���)).Value = True
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '��ѡ��������ѡ�����б�
    With Me.lvwBakup.ColumnHeaders
        .Add , "_����", "����", 900
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 800
    End With
    With Me.lvwApply.ColumnHeaders
        .Add , "_����", "����", 900
        .Add , "_����", "����", 2000
        .Add , "_����", "����", 800
    End With
    With Me.lvwApply
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.Id, d.����, d.����, d.����, Decode(s.����id, Null, 0, 1) As ѡ��" & _
            " From ���ű� d, ��������˵�� m, (Select ����id From �������ÿ��� Where ��Ŀ��� = [1]) s" & _
            " Where d.Id = m.����id And d.Id = s.����id(+) And m.�������� = '�ٴ�' And m.������� In (2, 3)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileId)
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwBakup.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = "" & !����
            If !ѡ�� = 1 Then objItem.Checked = True
            
            Set objItem = Me.lvwApply.ListItems.Add(, "_" & !ID, !����)
            objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwBakup.ColumnHeaders("_����").Index - 1) = "" & !����
            If !ѡ�� = 1 Then objItem.Checked = True
            .MoveNext
        Loop
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOK: Unload Me
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkSelect_Click()
    Dim objAdd As ListItem

    Me.lvwApply.ListItems.Clear
    If Me.chkSelect.Value Then
        For Each objItem In Me.lvwBakup.ListItems
            If objItem.Checked Then
                Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
                objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
                objAdd.Checked = objItem.Checked
            End If
        Next
    Else
        For Each objItem In Me.lvwBakup.ListItems
            Set objAdd = Me.lvwApply.ListItems.Add(, objItem.Key, objItem.Text)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
            objAdd.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1) = objItem.SubItems(Me.lvwApply.ColumnHeaders("_����").Index - 1)
            objAdd.Checked = objItem.Checked
        Next
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strSelected As String
    strSelected = ""
    For Each objItem In Me.lvwApply.ListItems
        If objItem.Checked Then strSelected = strSelected & ";" & Mid(objItem.Key, 2)
    Next
    If strSelected <> "" Then strSelected = Mid(strSelected, 2)
    
    If Me.optApply(0).Value Then
        gstrSQL = "Zl_�������ÿ���_Apply(" & mlngFileId & ",0,Null)"
    ElseIf Me.optApply(1).Value Then
        gstrSQL = "Zl_�������ÿ���_Apply(" & mlngFileId & ",1,Null)"
    Else
        If strSelected = "" Then MsgBox "û��ѡ����ң�", vbInformation, gstrSysName: Me.lvwApply.SetFocus: Exit Sub
        gstrSQL = "Zl_�������ÿ���_Apply(" & mlngFileId & ",2,'" & strSelected & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    For Each objItem In Me.lvwBakup.ListItems
        objItem.Checked = IIf(Index = 0, True, False)
    Next
    Call chkSelect_Click
    Me.lvwApply.SetFocus
End Sub

Private Sub lvwApply_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwApply.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwApply.SortOrder = IIf(Me.lvwApply.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwApply.SortKey = ColumnHeader.Index - 1
        Me.lvwApply.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwApply_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Me.lvwBakup.ListItems(Item.Key).Checked = Item.Checked
End Sub

Private Sub optApply_Click(Index As Integer)
    Me.lvwApply.Enabled = Me.optApply(2).Value
    Me.chkSelect.Enabled = Me.optApply(2).Value
    Me.cmdSelect(0).Enabled = Me.optApply(2).Value
    Me.cmdSelect(1).Enabled = Me.optApply(2).Value
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


