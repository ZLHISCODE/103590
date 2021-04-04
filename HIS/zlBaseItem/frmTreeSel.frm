VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeSel 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5700
   Icon            =   "frmTreeSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picOpt 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5685
      Left            =   4365
      ScaleHeight     =   5685
      ScaleWidth      =   1335
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.CommandButton cmdNext 
         Caption         =   "��һ��(&N)"
         Height          =   350
         Left            =   30
         TabIndex        =   6
         Top             =   1920
         Width           =   1100
      End
      Begin VB.TextBox txtLocate 
         Height          =   320
         Left            =   30
         TabIndex        =   5
         ToolTipText     =   "������һ��F3��س�����λ�����F4"
         Top             =   1470
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   0
         TabIndex        =   3
         Top             =   690
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lblLocate 
         Caption         =   "����(&F)"
         Height          =   255
         Left            =   30
         TabIndex        =   4
         Top             =   1230
         Width           =   1095
      End
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5847
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0442
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0896
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0CEA
            Key             =   "Book"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0E44
            Key             =   "BookOpen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":0F9E
            Key             =   "bm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTreeSel.frx":1538
            Key             =   "item"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTreeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Dim mstrID As String
Dim mstr�ϼ�ID As String
Dim mstr�ϼ����� As String
Dim mstr�ϼ����� As String
Dim mstrԭ���� As String
Dim mblnSecceed As Boolean
Dim mblnNode As Boolean
Dim mblnRoot As Boolean '����ѡ������
Private mIntStart As Integer               '��¼��ѯ�Ŀ�ʼλ��
Private mIntEnd As Integer                  '��¼���λ��

Dim mstrCaption As String

Dim mblnCheckChild As Boolean

Private Sub cmdCancel_Click()
    mblnSecceed = False
    Unload Me
End Sub

Private Function FullChild(ByVal NodeSour As Node, ByVal NodeFind As Node) As Boolean
'���Ҫ���ҵ��Ǹ������ǲ��������Ӷ���
'����ָ���Ķ���ݹ�����Ӷ���
Dim i As Long
Dim objNode As Node
Dim blnReturn As Boolean
    
    If NodeSour.Key = NodeFind.Key Then
        FullChild = True
        Exit Function
    Else
        If Not NodeSour.Child Is Nothing Then
            i = NodeSour.Child.FirstSibling.Index
            Set objNode = NodeSour.Child
            While i <= NodeSour.Child.LastSibling.Index
                If objNode.Key = NodeFind.Key Then
                    FullChild = True
                    Exit Function
                Else
                    blnReturn = FullChild(objNode, NodeFind)
                    If blnReturn = True Then
                        FullChild = True
                        Exit Function
                    End If
                    Set objNode = objNode.Next
                    If Not objNode Is Nothing Then
                        i = objNode.Index
                    Else
                        Exit Function
                    End If
                End If
            Wend
        End If
    End If
End Function


Private Sub cmdNext_Click()
    Call txtLocate_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim nod As Node
    Dim i As Integer
    Dim str���� As String
    
    
    Set nod = tvw.SelectedItem
    
    If tvw.SelectedItem.Key = "Root" Then
        tvw.SelectedItem.Expanded = True
        tvw.SelectedItem.EnsureVisible
        If mblnRoot = False Then
            MsgBox "��ѡ���ӷ��ࡣ", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If Not (Trim(mstr�ϼ�ID) = "" Or Trim(mstr�ϼ�ID) = "0") Then
        If FullChild(tvw.Nodes.Item("C" & mstr�ϼ�ID), tvw.SelectedItem) And mblnCheckChild = True Then
            If IsNumeric(mstrID) Then      'ֻ�в��������Ĳż��
                If CLng(mstrID) > 0 Then
                    MsgBox "�˽ڵ㲻����Ҫ������ѡ��", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    mblnSecceed = True
    With tvw.SelectedItem
        If .Key = "Root" Then
            If mblnRoot = False Then
                MsgBox "��ѡ���ӷ��ࡣ", vbInformation, gstrSysName
                Exit Sub
            End If
            mstr�ϼ�ID = ""
            mstr�ϼ����� = "��"
            mstr�ϼ����� = ""
        ElseIf .ForeColor = &H8000000C Then
            MsgBox "�޸ò��ŵ�Ȩ�ޣ�", vbInformation, gstrSysName
            mstr�ϼ�ID = ""
            mstr�ϼ����� = "��"
            mstr�ϼ����� = ""
            Exit Sub
        Else
            i = InStr(.Text, "��")
            mstr�ϼ�ID = Mid(.Key, 2)
            mstr�ϼ����� = Mid(.Text, i + 1)
            mstr�ϼ����� = Mid(.Text, 2, i - 2)
        End If
    End With
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = mstrCaption
    RestoreWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tvw.Top = 100
    tvw.Left = 100
    tvw.Height = ScaleHeight - 200
    tvw.Width = picOpt.Left - tvw.Left - 200
    
End Sub

Public Function ShowTree(ByVal strSQL As String, str�ϼ�ID As String, str�ϼ����� As String, str�ϼ����� As String, _
                        strID As String, ByVal strCaption As String, _
                        ByVal strRoot As String, Optional blnRoot As Boolean = True, Optional strԭ���� As String, _
                        Optional ByVal IconIndex As Long = 0, _
                        Optional ByVal SelectIconIndex As Long = 0, _
                        Optional ByVal ExpIconIndex As Long = 0, _
                        Optional ByVal blnChild As Boolean = True) As Boolean
'����:����SQL�����ʾ������Ŀ,��ѡ��ĳ��ĩ����Ŀ
'����:strSql        SQL���
'     str�ϼ�ID     ������ѡ����Ŀ���ϼ�ID
'     str�ϼ�����   ������ѡ����Ŀ���ϼ�����
'     str�ϼ�����   ������ѡ����Ŀ���ϼ�����
'     strID         ������ѡ����Ŀ��ID
'     strRoot       �����ı���
'     strICO        ͼ����Դ������
'     strCaption    ���ڵı���
'     IconIndex     ͼ������
'     SelectIconIndex   ѡ�����ͼ������
'     ExpIconIndex  ��չͼ������
'     blnChild      ����ӽ��
'����:����ѡ�񷵻�True,���򷵻�False.
On Error GoTo errHandle
    Dim rs���� As New ADODB.Recordset
    Dim objNode As Node, bln���� As Boolean, i As Long
    
    
    mblnRoot = blnRoot
    mstrCaption = strCaption
    mblnCheckChild = blnChild
    
    Call zlDatabase.OpenRecordset(rs����, strSQL, Me.Caption)
'    For i = 0 To rs����.Fields.Count - 1
'        If rs����.Fields(i).Name = "����" Then
'            bln���� = True
'            Exit For
'        End If
'    Next
    
    tvw.Nodes.Clear
    tvw.Nodes.Add , , "Root", strRoot, "Root", "Root"
    tvw.Nodes("Root").Sorted = True
    Do Until rs����.EOF
        
        If IsNull(rs����("�ϼ�id")) Then
            Set objNode = tvw.Nodes.Add("Root", tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), IIF(IconIndex > 0 And IconIndex < 7, IconIndex, "Write"), IIF(SelectIconIndex > 0 And SelectIconIndex < 7, SelectIconIndex, "Write"))
        Else
            Set objNode = tvw.Nodes.Add("C" & rs����("�ϼ�id"), tvwChild, "C" & rs����("id"), "��" & rs����("����") & "��" & rs����("����"), IIF(IconIndex > 0 And IconIndex < 7, IconIndex, "Write"), IIF(SelectIconIndex > 0 And SelectIconIndex < 7, SelectIconIndex, "Write"))
        End If
        objNode.Tag = Nvl(rs����!����)
'        If bln���� Then objNode.Tag = rs����!����
        If SelectIconIndex > 0 And SelectIconIndex < 7 Then
            objNode.ExpandedImage = SelectIconIndex
        End If
        objNode.Sorted = True
        rs����.MoveNext
    Loop
    If str�ϼ�ID = "0" Then str�ϼ�ID = ""
    If str�ϼ�ID = "" Then
        tvw.Nodes("Root").Selected = True
        tvw.Nodes("Root").Expanded = True
    Else
        tvw.Nodes("C" & str�ϼ�ID).Selected = True
        tvw.Nodes("C" & str�ϼ�ID).EnsureVisible
    End If
    
    mstrID = strID
    mstr�ϼ�ID = str�ϼ�ID
    mstrԭ���� = strԭ����
    Me.Show vbModal
    ShowTree = mblnSecceed
    '�ɹ��˲ŷ���ֵ
    If mblnSecceed = True Then
        str�ϼ�ID = mstr�ϼ�ID
        str�ϼ����� = mstr�ϼ�����
        str�ϼ����� = mstr�ϼ�����
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    mIntEnd = 0
    mIntStart = 0
End Sub

Private Sub tvw_DblClick()
    If mblnNode Then cmdOK_Click
End Sub

Private Sub tvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If mblnNode Then cmdOK_Click
    End If
End Sub

Private Sub tvw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnNode = False
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    mblnNode = True
    mIntStart = tvw.SelectedItem.Index
End Sub

Public Function ShowTreePrivs(ByVal lngOperationID As Long, str�ϼ�ID As String, str�ϼ����� As String, str�ϼ����� As String) As Boolean
'����:װ���������ŵ�tvwMain_S
    Dim nodTmp As Node
    Dim rsDeptID As ADODB.Recordset
    Dim strTemp As String
    strTemp = "Write"
    
    On Error GoTo errHandle
    gstrSQL = "Select Max(Level) as ��,A.ID,A.�ϼ�ID,A.����,'��'||A.����||'��' ����,Upper(a.����) as ���� " & _
              "From ���ű� A Start With ID IN(Select ����ID From ������Ա Where ��ԱID=[1]) Connect by Prior �ϼ�ID=ID " & _
              "Group by A.ID,A.�ϼ�ID,A.����,A.����,a.���� " & _
              "Order by A.����,�� Desc"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvw
        .LineStyle = tvwRootLines
        .Sorted = True
        .Nodes.Clear
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!�ϼ�ID), 0, rsDeptID!�ϼ�ID) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!�ϼ�ID, tvwChild, "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.Tag = rsDeptID!����
            nodTmp.ForeColor = &H8000000C
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    End With
    '�����ӽ��
    gstrSQL = "Select ID,�ϼ�ID,'��'||����||'��' ����,����,Upper(����) as ���� " & _
              "From ���ű� A " & _
              "Start With ID IN(Select ����ID From ������Ա Where ��ԱID=[1]) Connect by Prior ID=�ϼ�ID"
    Set rsDeptID = zlDatabase.OpenSQLRecord(gstrSQL, Caption, glngUserId)
    With tvw
        Do While Not rsDeptID.EOF
            If IIF(IsNull(rsDeptID!�ϼ�ID), 0, rsDeptID!�ϼ�ID) = 0 Then
                If .Nodes.Count > 0 Then
                    If FindKey("C" & rsDeptID!ID) = False Then
                        Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                    Else
                        Set nodTmp = .Nodes("C" & rsDeptID!ID)
                    End If
                Else
                    Set nodTmp = .Nodes.Add(, , "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                End If
            Else
                If FindKey("C" & rsDeptID!ID) = False Then
                    Set nodTmp = .Nodes.Add("C" & rsDeptID!�ϼ�ID, tvwChild, "C" & rsDeptID!ID, rsDeptID!���� & rsDeptID!����, strTemp, strTemp)
                Else
                    Set nodTmp = .Nodes("C" & rsDeptID!ID)
                End If
            End If
            nodTmp.Tag = rsDeptID!����
            nodTmp.ForeColor = vbBlack
            rsDeptID.MoveNext
        Loop
        rsDeptID.Close
    
        If .Nodes.Count > 0 Then .Nodes(1).Selected = True
    
    End With
    Me.Show vbModal
    ShowTreePrivs = mblnSecceed
    If mblnSecceed = True Then
        str�ϼ�ID = mstr�ϼ�ID
        str�ϼ����� = mstr�ϼ�����
        str�ϼ����� = mstr�ϼ�����
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FindKey(ByVal strKey As String) As Boolean
    Dim nodTmp As Node
    For Each nodTmp In tvw.Nodes
        If nodTmp.Key = strKey Then
            FindKey = True
            Exit Function
        End If
    Next
End Function

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1

        If lngStart >= tvw.Nodes.Count Then lngStart = 1
        
        For i = tvw.Nodes.Count To 1 Step -1    '�������һ����λ��
            If UCase(tvw.Nodes(i).Text) Like "*" & UCase(txtLocate.Text) & "*" Or UCase(tvw.Nodes(i).Tag) Like "*" & UCase(txtLocate.Text) & "*" Then
                mIntEnd = i
                Exit For
            End If
        Next
        
        If lngStart - 1 = mIntEnd Then  '��������һ�����ѯ��һ��
            lngStart = 1
        End If
        If mIntStart < Val(lblLocate.Tag) And mIntStart <> 0 Then '����ѡ���˲�ѯ��λ��
            lngStart = mIntStart + 1
            mIntStart = 0
        End If
        For i = lngStart To tvw.Nodes.Count
            If tvw.Nodes(i).Text Like "*" & txtLocate.Text & "*" Or tvw.Nodes(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call tvw.Nodes(i).EnsureVisible
                tvw.Nodes(i).Selected = True
                lblLocate.Tag = i
                tvw.SetFocus
                Exit For
            End If
            If i = tvw.Nodes.Count Then
                MsgBox "û�в�ѯ�������������Ϣ�����������룡", vbInformation, gstrSysName
                txtLocate.Text = ""
                txtLocate.SetFocus
            End If
        Next
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub
