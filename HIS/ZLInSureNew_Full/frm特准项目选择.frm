VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm��׼��Ŀѡ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��׼��Ŀѡ��"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frm��׼��Ŀѡ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdDel 
      Caption         =   "ɾ��(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5190
      TabIndex        =   11
      Top             =   2430
      Width           =   1100
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   5190
      TabIndex        =   10
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5190
      TabIndex        =   14
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5190
      TabIndex        =   13
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5190
      TabIndex        =   12
      Top             =   3330
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "����(&R)"
      Height          =   1695
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4875
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1140
         Width           =   2685
      End
      Begin VB.CommandButton Cmd�������� 
         Caption         =   "��"
         Height          =   300
         Left            =   4050
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   750
         Width           =   285
      End
      Begin VB.TextBox Txt�������� 
         Height          =   300
         Left            =   1650
         TabIndex        =   5
         Top             =   750
         Width           =   2415
      End
      Begin VB.CommandButton Cmd��ʼ���� 
         Caption         =   "��"
         Height          =   300
         Left            =   4050
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox Txt��ʼ���� 
         Height          =   300
         Left            =   1650
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   960
         TabIndex        =   7
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl��ʼ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼ����(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   5370
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":1CFA
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":2014
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":232E
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":2648
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":2962
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm��׼��Ŀѡ��.frx":2EFC
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm��׼��Ŀѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public frmParent As Object
Public lng���� As Long
Public bln��ϸ As Boolean
Private strSelect As String

Private Sub CmdAdd_Click()
    Dim str��ʼ���� As String, str�������� As String, strSql As String
    Dim lvsItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    '���û�����ķ�Χ�ڵ���Ŀ�����뵽�б����
    str��ʼ���� = Trim(Txt��ʼ����.Tag)
    str�������� = Trim(Txt��������.Tag)
    
    '����SQL
    If str��ʼ���� <> "" And str�������� <> "" Then
        strSql = " And A.���� Between '" & str��ʼ���� & "' And '" & str�������� & "'"
    Else
        If str��ʼ���� <> "" Then
            strSql = " And A.����>='" & str��ʼ���� & "'"
        ElseIf str�������� <> "" Then
            strSql = " And A.����<='" & str�������� & "'"
        Else
            MsgBox "�����뿪ʼ���룺", vbInformation, gstrSysName
            Txt��ʼ����.SetFocus
        End If
    End If
    
    If bln��ϸ Then
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,B.����,B.����,A.���,A.��� " & _
                 "   FROM �շ�ϸĿ A,�շѱ��� B WHERE A.ID=B.�շ�ϸĿID " & strSql
    Else
        gstrSQL = "Select A.ID,A.����,A.����,A.���� " & _
                 "   FROM ����֧������ A WHERE ����=" & lng���� & strSql
    End If
    Call OpenRecordset(rsTemp, "���û��趨������������¼")
    
    Do While Not rsTemp.EOF
        If InStr(1, strSelect & "|", "|" & rsTemp!ID & "|") = 0 Then
            strSelect = strSelect & "|" & rsTemp!ID
            Call addLvw(rsTemp)
        End If
        rsTemp.MoveNext
    Loop
    
    cmdOK.Enabled = (lvwDetail.ListItems.Count <> 0)
    CmdDel.Enabled = cmdOK.Enabled
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDel_Click()
    Dim lngItem As Long
    
    With lvwDetail
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        strSelect = strSelect & "|"
        For lngItem = 1 To .ListItems.Count
            If lngItem > .ListItems.Count Then Exit For
            If .ListItems(lngItem).Selected Then
                strSelect = Replace(strSelect, "|" & Mid(.ListItems(lngItem).Key, 2) & "|", "|")
                .ListItems.Remove .ListItems(lngItem).Key
                lngItem = lngItem - 1
            End If
        Next
        If Len(strSelect) > 1 Then
            strSelect = Mid(strSelect, 1, Len(strSelect) - 1)
        Else
            strSelect = ""
        End If
        
        If .ListItems.Count <> 0 Then .ListItems(1).Selected = True
        cmdOK.Enabled = (.ListItems.Count <> 0)
        CmdDel.Enabled = cmdOK.Enabled
    End With
End Sub

Private Sub cmdOK_Click()
    Dim strExist As String
    Dim objLvw As ListView, objItem As ListItem
    '�����������е�����
    If bln��ϸ Then
        Set objLvw = frmParent.Lvw��ϸ
    Else
        Set objLvw = frmParent.lvw����
    End If
    
    '��ȡ�Ѵ�����Ŀ��ID��
    strExist = ""
    With objLvw
        For Each objItem In .ListItems
            strExist = strExist & "|" & Mid(objItem.Key, 2)
        Next
    End With
    
    '���뵱ǰѡ�����Ŀ
    With lvwDetail
        For Each objItem In .ListItems
            If InStr(1, strExist & "|", "|" & Mid(objItem.Key, 2) & "|") = 0 Then
                With objLvw
                    .ListItems.Add , objItem.Key, "[" & objItem.Text & "]" & objItem.SubItems(1), IIf(bln��ϸ, "Fix", "Limit"), IIf(bln��ϸ, "Fix", "Limit")
                    If bln��ϸ Then
                        .ListItems(objItem.Key).SubItems(1) = objItem.SubItems(2)
                        .ListItems(objItem.Key).SubItems(2) = objItem.SubItems(3)
                    Else
                        .ListItems(objItem.Key).SubItems(1) = objItem.SubItems(2)
                    End If
                End With
                strExist = strExist & "|" & Mid(objItem.Key, 2)
            End If
        Next
    End With
    
    Unload Me
End Sub

Private Sub Cmd��������_Click()
    Dim strID As String, str���� As String, str���� As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If bln��ϸ Then
        If frm�շ�ϸĿѡ��.ShowTree(strID, str����, str����) = True Then
            If Get��Ŀ(strID, rsTemp) = False Then Exit Sub
            Txt��������.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            Txt��������.Tag = rsTemp!����
        End If
    Else
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,A.����,A.���� " & _
                 "   FROM ����֧������ A Where  ����=" & lng����
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            '����ѡ����
            If rsTemp.RecordCount > 1 Then
                '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                blnReturn = frmListSel.ShowSelect(lng����, rsTemp, "���", "ѡ����", "��ѡ����Ŀ��")
            Else
                blnReturn = True
            End If
        End If
        
        If blnReturn = False Then
            '��¼����û�п�ѡ�������
            Exit Sub
        Else
            Txt��������.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            Txt��������.Tag = rsTemp!����
        End If
    End If
End Sub

Private Sub Cmd��ʼ����_Click()
    Dim strID As String, str���� As String, str���� As String
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    If bln��ϸ Then
        If frm�շ�ϸĿѡ��.ShowTree(strID, str����, str����) = True Then
            If Get��Ŀ(strID, rsTemp) = False Then Exit Sub
            Txt��ʼ����.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            Txt��ʼ����.Tag = rsTemp!����
        End If
    Else
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,A.����,A.���� " & _
                 "   FROM ����֧������ A  Where ����=" & lng����
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.RecordCount > 0 Then
            '����ѡ����
            If rsTemp.RecordCount > 1 Then
                '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
                blnReturn = frmListSel.ShowSelect(lng����, rsTemp, "���", "ѡ����", "��ѡ����Ŀ��")
            Else
                blnReturn = True
            End If
        End If
        
        If blnReturn = False Then
            '��¼����û�п�ѡ�������
            Exit Sub
        Else
            Txt��ʼ����.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            Txt��ʼ����.Tag = rsTemp!����
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo����
        .Clear
        .AddItem "0-����"
        .AddItem "1-����"
        .AddItem "2-�ų�"
        .ListIndex = 0
    End With
    
    strSelect = ""
    Call initLvw
End Sub

Private Sub initLvw()
    lvwDetail.ColumnHeaders.Clear
    If bln��ϸ = False Then
        lvwDetail.ColumnHeaders.Add , "K1", "����", 800
        lvwDetail.ColumnHeaders.Add , "K2", "����", 2000
        lvwDetail.ColumnHeaders.Add , "K3", "����", 1000
    Else
        lvwDetail.ColumnHeaders.Add , "K1", "����", 800
        lvwDetail.ColumnHeaders.Add , "K2", "����", 2000
        lvwDetail.ColumnHeaders.Add , "K3", "���", 1000
        lvwDetail.ColumnHeaders.Add , "K4", "����", 1000
    End If
End Sub

Private Sub addLvw(ByVal rsTemp As ADODB.Recordset)
    With lvwDetail
        .ListItems.Add , "K" & rsTemp!ID, rsTemp!����, IIf(bln��ϸ, "Fix", "Limit"), IIf(bln��ϸ, "Fix", "Limit")
        .ListItems("K" & rsTemp!ID).SubItems(1) = rsTemp!����
        If bln��ϸ Then
            .ListItems("K" & rsTemp!ID).SubItems(2) = Nvl(rsTemp!���)
            .ListItems("K" & rsTemp!ID).SubItems(3) = cbo����.Text
        Else
            .ListItems("K" & rsTemp!ID).SubItems(2) = cbo����.Text
        End If
    End With
End Sub

Private Function Get��Ŀ(ByVal strID As String, rsTemp As ADODB.Recordset) As Boolean
'���ܣ�������ĿID���õ���Ŀ����
    On Error GoTo errHandle
    
    If Trim(strID) = "" Then Exit Function
    If bln��ϸ Then
        gstrSQL = "Select ID,����,����,���,��� From �շ�ϸĿ Where ID=" & strID
    Else
        gstrSQL = "Select ID,����,���� From ����֧������ Where ID=" & strID
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Get��Ŀ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Txt��������_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Txt��������.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = Txt��������.Text
    If bln��ϸ Then
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,B.����,B.����,A.���,A.��� " & _
                 "   FROM �շ�ϸĿ A,�շѱ��� B WHERE A.ID=B.�շ�ϸĿID And (" & _
                    zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & ")"
    Else
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,A.����,A.���� " & _
                 "   FROM ����֧������ A WHERE ����=" & lng���� & " And (" & _
                    zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & ")"
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(lng����, rsTemp, "���", "ѡ����", "��ѡ����Ŀ��")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll Txt��������
        Exit Sub
    Else
        Txt��������.Text = "[" & rsTemp!���� & "]" & rsTemp!����
        Txt��������.Tag = rsTemp!����
    End If
    zlControl.TxtSelAll Txt��������
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Txt��ʼ����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Txt��ʼ����.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = Txt��ʼ����.Text
    If bln��ϸ Then
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,B.����,B.����,A.���,A.��� " & _
                 "   FROM �շ�ϸĿ A,�շѱ��� B WHERE A.ID=B.�շ�ϸĿID And (" & _
                    zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & ")"
    Else
        gstrSQL = "Select Distinct Rownum as ���, A.ID,A.����,A.����,A.���� " & _
                 "   FROM ����֧������ A WHERE ����=" & lng���� & " And (" & _
                    zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & " or " & zlCommFun.GetLike("B", "����", strText) & ")"
    End If
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(lng����, rsTemp, "���", "ѡ����", "��ѡ����Ŀ��")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll Txt��ʼ����
        Exit Sub
    Else
        Txt��ʼ����.Text = "[" & rsTemp!���� & "]" & rsTemp!����
        Txt��ʼ����.Tag = rsTemp!����
    End If
    zlControl.TxtSelAll Txt��ʼ����
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
