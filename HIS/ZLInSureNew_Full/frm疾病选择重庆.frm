VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm����ѡ������ 
   AutoRedraw      =   -1  'True
   Caption         =   "����ѡ��"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   Icon            =   "frm����ѡ������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9930
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   9930
      TabIndex        =   1
      Top             =   4725
      Width           =   9930
      Begin VB.TextBox txt����֢ 
         Height          =   300
         Left            =   1110
         MaxLength       =   80
         TabIndex        =   3
         Top             =   150
         Width           =   8415
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "ɸѡ(&S)"
         Height          =   350
         Left            =   3750
         TabIndex        =   11
         Top             =   540
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   1110
         TabIndex        =   5
         Top             =   570
         Width           =   2625
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8430
         TabIndex        =   7
         Top             =   540
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   7170
         TabIndex        =   6
         Top             =   540
         Width           =   1100
      End
      Begin VB.Label lbl����֢ 
         AutoSize        =   -1  'True
         Caption         =   "����֢(&U)"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   630
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3915
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
            Picture         =   "frm����ѡ������.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm����ѡ������.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   3930
      Left            =   330
      TabIndex        =   0
      Top             =   450
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   6932
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.TabStrip tab���� 
      Height          =   4125
      Left            =   210
      TabIndex        =   8
      Top             =   150
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7276
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ͨ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   3960
      Left            =   6840
      TabIndex        =   10
      Top             =   420
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   6985
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Label lblSelect 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ�б�(&L)"
      Height          =   180
      Left            =   6870
      TabIndex        =   9
      Top             =   180
      Width           =   990
   End
End
Attribute VB_Name = "frm����ѡ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String
Private mstrName As String
Private mstr����֢ As String
Private mblnOK As Boolean
Private mcnYB As New ADODB.Connection   'ҽ��ǰ�÷���������

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lvwDetail.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '����ѡ����Ŀ����
    mstrCode = lvwDetail.SelectedItem.Text
    mstrName = lvwDetail.SelectedItem.SubItems(1)
    mstr����֢ = Trim(txt����֢.Text)
    If Trim(mstr����֢) = "" Then
        MsgBox "����֢����Ϊ�գ�", vbInformation, gstrSysName
        txt����֢.SetFocus
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetCode(ByVal varList As Variant, strCode As String, str���� As String, str���� As String, str����֢ As String) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ�������������������������Ϊѡ������
'���أ��ɹ�����True
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim strServer As String, strUser As String, strPass As String
    Dim lst As ListItem, lngIndex As Long
    
    mblnOK = False
    
    MousePointer = vbHourglass
    On Error GoTo ErrH
    
    If strCode = "����" Or strCode = "��Ժ" Then
        '���ȶ���������������
        gstrSQL = "Select ������,����ֵ From ���ղ��� Where ����=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_������)
        Do Until rsTemp.EOF
            strTemp = IIf(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ"))
            Select Case rsTemp("������")
                Case "ҽ��������"
                    strServer = strTemp
                Case "ҽ���û���"
                    strUser = strTemp
                Case "ҽ���û�����"
                    strPass = strTemp
            End Select
            rsTemp.MoveNext
        Loop
        If OraDataOpen(mcnYB, strServer, strUser, strPass) = False Then
            MousePointer = vbDefault
            Exit Function
        End If
    
        If rsTemp.State = adStateOpen Then rsTemp.Close
        If strCode = "����" Then
            tab����.Visible = False
            rsTemp.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML where bzfl in (2,3) Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
        Else
            '���Կ���Tab
            tab����.Visible = True
            tab����.Tag = "��Ժ"
            '20031231:���:���񲡵ĵ�����5��
            rsTemp.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML where bzfl in(1,2) Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
        End If
        If rsTemp.EOF = True Then
            MousePointer = vbDefault
            MsgBox "δ��ҽ��ǰ�÷������ж�����ز��֡�", vbInformation, gstrSysName
            Exit Function
        End If
        
        lvwDetail.ListItems.Clear
        Do Until rsTemp.EOF
            Set lst = lvwDetail.ListItems.Add(, , rsTemp("����"), "Detail", "Detail")
            lst.SubItems(1) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            lst.SubItems(2) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
            If lst.Text = str���� Then
                lst.Selected = True
                lst.EnsureVisible
            End If
            rsTemp.MoveNext
        Loop
    Else
        '���ⲡ
        '���Ƚ��ִ���ԭ
        strTemp = ""
        For lngIndex = 1 To UBound(varList)
            strTemp = strTemp & varList(lngIndex) & "|"
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        
        varList = Split(strTemp, "$")
        lvwDetail.ListItems.Clear
        For lngIndex = 0 To UBound(varList)
            strTemp = varList(lngIndex)
            If InStr(strTemp, "|") > 0 Then
                Set lst = lvwDetail.ListItems.Add(, , Split(strTemp, "|")(0), "Detail", "Detail")
                lst.SubItems(1) = Split(strTemp, "|")(1)
                lst.SubItems(2) = zlCommFun.SpellCode(Split(strTemp, "|")(1))
            End If
        Next
        
        If lvwDetail.ListItems.Count = 0 Then
            MousePointer = vbDefault
            MsgBox "�ò���������ͨ�������ⲡ�֡�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    
    MousePointer = vbDefault
    
    mstrCode = str����
    mstr����֢ = str����֢
    frm����ѡ������.Show vbModal
    '����ֵ
    If mblnOK = True Then
        str���� = mstrCode
        str���� = mstrName
        str����֢ = mstr����֢
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Function

Private Sub cmdSelect_Click()
    Dim itmDetail As ListItem, itmSelect As ListItem
    Dim strFind As String
    
    lvwSelect.ListItems.Clear
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    
    strFind = "*" & strFind & "*"
    For Each itmDetail In lvwDetail.ListItems
        If itmDetail.Text Like strFind Or itmDetail.SubItems(1) Like strFind Or itmDetail.SubItems(2) Like strFind Then
            Set itmSelect = lvwSelect.ListItems.Add(, , itmDetail.Text, "Detail", "Detail")
            itmSelect.SubItems(1) = itmDetail.SubItems(1)
        End If
    Next
End Sub

Private Sub Form_Load()
    
    Me.txt����֢ = mstr����֢
End Sub

Private Sub Form_Resize()
    Dim ctlTemp As Control
    
    If tab����.Tag = "��Ժ" Then
        Set ctlTemp = tab����
    Else
        Set ctlTemp = lvwDetail
    End If
    ctlTemp.Top = 60
    ctlTemp.Left = ScaleLeft
    ctlTemp.Height = Me.ScaleHeight - lvwDetail.Top - picCmd.Height
    
    lblSelect.Top = ctlTemp.Top
    lvwSelect.Top = lblSelect.Top + lblSelect.Height + 60
    lvwSelect.Left = ScaleWidth - lvwSelect.Width
    lblSelect.Left = lvwSelect.Left
    lvwSelect.Height = ctlTemp.Height - lvwSelect.Top
    
    On Error Resume Next
    ctlTemp.Width = lvwSelect.Left - 45 - ctlTemp.Left
    If tab����.Tag = "��Ժ" Then
        lvwDetail.Top = tab����.ClientTop
        lvwDetail.Left = tab����.ClientLeft
        lvwDetail.Width = tab����.ClientWidth
        lvwDetail.Height = tab����.ClientHeight
    End If
    
    txt����֢.Width = picCmd.Width - txt����֢.Left - (picCmd.Width - cmdCancel.Left - cmdCancel.Width)
End Sub

Private Sub lvwSelect_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwSelect, ColumnHeader.Index)
End Sub

Private Sub LvwSelect_DblClick()
    If lvwSelect.SelectedItem Is Nothing Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    '����ѡ����Ŀ����
    mstrCode = lvwSelect.SelectedItem.Text
    mblnOK = True
    Unload Me
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwDetail_DblClick()
    cmdOK_Click
End Sub

Private Sub lvwdetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvwDetail.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvwDetail.SortOrder = lvwDescending
    Else
        lvwDetail.SortOrder = lvwAscending
    End If
    lvwDetail.Sorted = True
    intIdx = ColumnHeader.Index
    
    If Not lvwDetail.SelectedItem Is Nothing Then lvwDetail.SelectedItem.EnsureVisible
End Sub

Private Sub tab����_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    
    lvwDetail.ListItems.Clear
    If mcnYB.State = adStateClosed Then Exit Sub
    
    MousePointer = vbHourglass
    On Error GoTo errHandle
    
    If tab����.SelectedItem.Caption = "��ͨ��" Then
        '20031231:���:���񲡵ĵ�����5��
        rsTemp.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML where bzfl=1 Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
    ElseIf tab����.SelectedItem.Caption = "������" Then
        rsTemp.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML where bzfl=4 and nvl(tjm,' ')<>'��' Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
    Else
        rsTemp.Open "select BZBM ����,BZMC ����,ZJM ����  from BZML where bzfl=5 Order by BZBM", mcnYB, adOpenStatic, adLockReadOnly
    End If
    
    Do Until rsTemp.EOF
        Set lst = lvwDetail.ListItems.Add(, , rsTemp("����"), "Detail", "Detail")
        lst.SubItems(1) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        lst.SubItems(2) = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        
        rsTemp.MoveNext
    Loop
    
    lvwDetail.SetFocus
    MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Sub

Private Sub txtFind_Change()
'���ܣ������û���������ݲ���ƥ�������
    Dim lst As ListItem, lngIndex As Long, lngSubItems As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    If lvwDetail.ListItems.Count = 0 Then Exit Sub
    
    Set lst = lvwDetail.FindItem(strFind, lvwText, , lvwPartial)
    If Not lst Is Nothing Then
        lst.Selected = True
        lst.EnsureVisible
    Else
        '���ı�������������ƥ��
        lngSubItems = lvwDetail.ColumnHeaders.Count - 1
        For Each lst In lvwDetail.ListItems
            For lngIndex = 1 To lngSubItems
                If lst.SubItems(lngIndex) Like strFind & "*" Then
                    lst.Selected = True
                    lst.EnsureVisible
                    Exit Sub
                End If
            Next
            
        Next
    End If
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub
