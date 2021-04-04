VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWritImp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˲�������"
   ClientHeight    =   3150
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6075
   Icon            =   "frmWritImp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4740
      TabIndex        =   3
      Top             =   135
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   2565
      Left            =   825
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3090
      Visible         =   0   'False
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvwWrits 
      Height          =   2355
      Left            =   150
      TabIndex        =   2
      Top             =   720
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtPati 
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   840
      MaxLength       =   11
      TabIndex        =   1
      ToolTipText     =   "�밴""-����ID""��""+סԺ��""��""*�����""��ʽ�����ֱ��������������"
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   4740
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1035
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4740
      TabIndex        =   4
      Top             =   495
      Width           =   1200
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   5100
      Picture         =   "frmWritImp.frx":08CA
      Top             =   2310
      Width           =   480
   End
   Begin VB.Label lblWrit 
      AutoSize        =   -1  'True
      Caption         =   "��Ժ��¼��"
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   525
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "����:        �Ա�:    ����:  "
      Height          =   180
      Left            =   1860
      TabIndex        =   6
      Top             =   165
      Width           =   2610
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmWritImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lngFileId As Long
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Private Sub cmdCancel_Click()
    lngFileId = 0
    Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    lngFileId = Mid(Me.lvwWrits.SelectedItem.Key, 2)
    Me.Hide
End Sub

Private Sub Form_Activate()
    gstrSql = "select ���� from �����ļ�Ŀ¼ where ID=" & Me.lblWrit.Tag
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Me.lblWrit.Caption = !���� & ":"
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then cmdCancel_Click
End Sub

Private Sub Form_Load()
    
    Me.lvwWrits.ListItems.Clear
    With Me.lvwWrits.ColumnHeaders
        .Clear
        .Add , "���", "���", 600
        .Add , "��д��", "��д��", 900
        .Add , "��д����", "��д����", 1700
    End With
    With Me.lvwWrits
        .SortKey = .ColumnHeaders("���").Index - 1: .SortOrder = lvwAscending
    End With
    
    With Me.lvwPati.ColumnHeaders
        .Clear
        .Add , "����ID", "����ID", 800
        .Add , "�����", "�����", 800
        .Add , "סԺ��", "סԺ��", 800
        .Add , "����", "����", 900
        .Add , "�Ա�", "�Ա�", 600
        .Add , "����", "����", 600
    End With
    With Me.lvwPati
        .SortKey = .ColumnHeaders("����ID").Index - 1: .SortOrder = lvwAscending
    End With

End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwPati.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwPati.SortOrder = IIf(Me.lvwPati.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwPati.SortKey = ColumnHeader.Index - 1
        Me.lvwPati.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwPati_DblClick()
    If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwPati
        If Val(Me.txtPati.Tag) <> Val(.SelectedItem.Text) Then
            Me.txtPati.Tag = .SelectedItem.Text
            Me.txtPati.Text = Me.txtPati.Tag
            Me.lblInfo.Caption = "����:" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & _
                    Space(2) & "�Ա�:" & .SelectedItem.SubItems(.ColumnHeaders("�Ա�").Index - 1) & _
                    Space(2) & "����:" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
            Me.lblInfo.Tag = .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
        End If
        Me.txtPati.SetFocus
        Call RefereshWrits
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwPati.SelectedItem Is Nothing Then Exit Sub
        Call lvwPati_DblClick
    End Select
End Sub

Private Sub lvwPati_LostFocus()
    Me.lvwPati.Visible = False
End Sub

Private Sub lvwWrits_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwWrits.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwWrits.SortOrder = IIf(Me.lvwWrits.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwWrits.SortKey = ColumnHeader.Index - 1
        Me.lvwWrits.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwWrits_DblClick()
    If Me.lvwWrits.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub lvwWrits_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwWrits.SelectedItem Is Nothing Then Exit Sub
    Call cmdOK_Click
End Sub

Private Sub txtPati_GotFocus()
    Me.txtPati.SelStart = 0: Me.txtPati.SelLength = 100
End Sub

Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If InStr("~!@#$^&()|=`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Me.txtPati.Text = Trim(Me.txtPati.Text)
    If Me.txtPati.Text = "" Then Me.txtPati.Text = Me.txtPati.Tag: Exit Sub
    
    Select Case Left(Me.txtPati.Text, 1)
    Case "-", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0" '����ID
        gstrSql = "select ����ID,�����,סԺ��,����,�Ա�,����" & _
                " from ������Ϣ" & _
                " where ����id=" & Abs(Val(Me.txtPati.Text))
    Case "+"        'סԺ��
        gstrSql = "select ����ID,�����,סԺ��,����,�Ա�,����" & _
                " from ������Ϣ" & _
                " where סԺ��=" & Val(Me.txtPati.Text)
    Case "*"        '�����
        gstrSql = "select ����ID,�����,סԺ��,����,�Ա�,����" & _
                " from ������Ϣ" & _
                " where �����=" & Val(Mid(Me.txtPati.Text, 2))
    Case Else       '��������
        gstrSql = "select ����ID,�����,סԺ��,����,�Ա�,����" & _
                " from ������Ϣ" & _
                " where ���� like '" & Me.txtPati.Text & "%'"
    End Select
    
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .BOF Or .EOF = 1 Then
            MsgBox "δ�ҵ�ָ������", vbExclamation, gstrSysName
            Me.txtPati.Text = "": Me.txtPati.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Val(Me.txtPati.Tag) <> !����ID Then
                Me.txtPati.Tag = !����ID: Me.txtPati.Text = Me.txtPati.Tag
                Me.lblInfo.Caption = "����:" & Trim(!����) & _
                        Space(2) & "�Ա�:" & IIf(IsNull(!�Ա�), "", !�Ա�) & _
                        Space(2) & "����:" & IIf(IsNull(!����), "", !����)
                Me.lblInfo.Tag = !����
            End If
            Call RefereshWrits
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwPati.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwPati.ListItems.Add(, "_" & !����ID, !����ID)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("�����").Index - 1) = IIf(IsNull(!�����), "", !�����)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("סԺ��").Index - 1) = IIf(IsNull(!סԺ��), "", !סԺ��)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("�Ա�").Index - 1) = IIf(IsNull(!�Ա�), "", !�Ա�)
            objItem.SubItems(Me.lvwPati.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            .MoveNext
        Loop
        Me.lvwPati.ListItems(1).Selected = True
    End With
    With Me.lvwPati
        .Left = Me.txtPati.Left
        .Top = Me.txtPati.Top + Me.txtPati.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPati_LostFocus()
    Me.txtPati.Text = Me.txtPati.Tag
End Sub

Private Sub RefereshWrits()
    gstrSql = "select ID,Rownum As ���,��д��,��д���� From ���˲�����¼ where ����ID=" & Me.txtPati.Tag & " and �ļ�id=" & Me.lblWrit.Tag
    Err = 0: On Error GoTo ErrHand
    Me.cmdOK.Enabled = False
    Me.lvwWrits.ListItems.Clear
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        Do While Not .EOF
            Set objItem = Me.lvwWrits.ListItems.Add(, "_" & !ID, !���)
            objItem.SubItems(Me.lvwWrits.ColumnHeaders("��д��").Index - 1) = IIf(IsNull(!��д��), "", !��д��)
            objItem.SubItems(Me.lvwWrits.ColumnHeaders("��д����").Index - 1) = IIf(IsNull(!��д����), "", Format(!��д����, "YYYY-MM-DD HH:MM"))
            .MoveNext
        Loop
    End With
    If Me.lvwWrits.ListItems.Count > 0 Then Me.cmdOK.Enabled = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
