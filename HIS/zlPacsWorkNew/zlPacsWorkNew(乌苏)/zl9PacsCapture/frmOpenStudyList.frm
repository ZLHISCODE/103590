VERSION 5.00
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#70.0#0"; "zl9PacsControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmOpenStudyList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�򿪼��"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenStudyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin zl9PacsCapture.TranControl tcFrmQuery 
      Height          =   5085
      Left            =   3405
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   8969
      Begin VB.PictureBox picQuery 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4590
         Left            =   390
         ScaleHeight     =   4560
         ScaleWidth      =   5040
         TabIndex        =   10
         Top             =   120
         Width           =   5070
         Begin VB.CommandButton cmdNotOk 
            Caption         =   "ȡ ��(&Q)"
            Height          =   420
            Left            =   3300
            TabIndex        =   15
            Top             =   3765
            Width           =   1215
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ȷ ��(&O)"
            Height          =   420
            Left            =   1485
            TabIndex        =   14
            Top             =   3765
            Width           =   1215
         End
         Begin VB.TextBox txtQueryValue 
            Height          =   390
            Left            =   1515
            TabIndex        =   13
            Top             =   2985
            Width           =   3060
         End
         Begin VB.ComboBox cbxQueryType 
            Height          =   330
            ItemData        =   "frmOpenStudyList.frx":000C
            Left            =   1530
            List            =   "frmOpenStudyList.frx":0028
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2505
            Width           =   3045
         End
         Begin VB.TextBox txtName 
            Height          =   390
            Left            =   1530
            TabIndex        =   11
            Top             =   420
            Width           =   3060
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   390
            Left            =   1530
            TabIndex        =   16
            Top             =   1485
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   41535.6989930556
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   390
            Left            =   1530
            TabIndex        =   17
            Top             =   945
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   41534.6979166667
         End
         Begin VB.Label labName 
            Caption         =   "�� ѯ ֵ"
            Height          =   270
            Index           =   4
            Left            =   435
            TabIndex        =   22
            Top             =   3060
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "��ѯ�ű�"
            Height          =   270
            Index           =   3
            Left            =   435
            TabIndex        =   21
            Top             =   2565
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "��������"
            Height          =   270
            Index           =   2
            Left            =   405
            TabIndex        =   20
            Top             =   1575
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "��ʼ����"
            Height          =   270
            Index           =   1
            Left            =   405
            TabIndex        =   19
            Top             =   1020
            Width           =   870
         End
         Begin VB.Label labName 
            Caption         =   "��    ��"
            Height          =   270
            Index           =   0
            Left            =   405
            TabIndex        =   18
            Top             =   465
            Width           =   870
         End
      End
   End
   Begin VB.PictureBox picPanel 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   6315
      Width           =   12090
      Begin VB.PictureBox picInf 
         BorderStyle     =   0  'None
         Height          =   1065
         Left            =   45
         ScaleHeight     =   1065
         ScaleWidth      =   6600
         TabIndex        =   5
         Top             =   45
         Visible         =   0   'False
         Width           =   6600
         Begin VB.Label labAdviceInf 
            Height          =   645
            Left            =   1455
            TabIndex        =   8
            Top             =   360
            Width           =   5040
            WordWrap        =   -1  'True
         End
         Begin VB.Label labMoneyState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   42
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   855
            Left            =   15
            TabIndex        =   7
            Top             =   105
            Width           =   870
         End
         Begin VB.Label labAdviceContext 
            Caption         =   "ҽ�����ݣ�"
            Height          =   255
            Left            =   1035
            TabIndex        =   6
            Top             =   105
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ ��(&C)"
         Height          =   975
         Left            =   10950
         Picture         =   "frmOpenStudyList.frx":007C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   105
         Width           =   1080
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   975
         Left            =   9885
         Picture         =   "frmOpenStudyList.frx":0570
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   105
         Width           =   1080
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "�� ��(&F)"
         Height          =   975
         Left            =   8820
         Picture         =   "frmOpenStudyList.frx":15B2
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   105
         Width           =   1080
      End
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   510
      Top             =   120
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
            Picture         =   "frmOpenStudyList.frx":1AAB
            Key             =   "סԺ"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":2385
            Key             =   "����"
            Object.Tag             =   "2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwStudy 
      Height          =   6240
      Left            =   45
      TabIndex        =   3
      Top             =   15
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11007
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imglist"
      SmallIcons      =   "Imglist"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "�Ա�"
         Text            =   "�Ա�"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "����"
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "��ʶ��"
         Text            =   "��ʶ��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Ӱ�����"
         Text            =   "Ӱ�����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "ҽ������"
         Text            =   "ҽ������"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "���״̬"
         Text            =   "���״̬"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "���˿���"
         Text            =   "���˿���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "����ʱ��"
         Text            =   "����ʱ��"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Key             =   "������"
         Text            =   "������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Key             =   "������"
         Text            =   "������"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmOpenStudyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mlngModule As Long
Public blnOk As Boolean


Public Function ShowStudyWindow(ByRef lngAdviceId As Long, lngSendNo As Long, ByRef lngStudyState As Long, objOwner As Object) As Boolean
'��ʾ��鴰��
    blnOk = False
    
    Me.Show 1, objOwner
    
    If Me.blnOk Then
        lngAdviceId = Nvl(Me.lvwStudy.SelectedItem.Tag)
        lngSendNo = Nvl(Me.lvwStudy.SelectedItem.ListSubItems(1).Tag)
        lngStudyState = Nvl(Me.lvwStudy.SelectedItem.ListSubItems(2).Tag)
    End If
    
    ShowStudyWindow = blnOk
End Function

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    blnOk = False
    Call Me.Hide
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdFind_Click()
On Error GoTo errHandle
    Call ShowQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub ShowQueryWindow()
    tcFrmQuery.Left = 0
    tcFrmQuery.Top = 0
    tcFrmQuery.Width = Me.ScaleWidth
    tcFrmQuery.Height = Me.ScaleHeight
    
    picQuery.Left = (tcFrmQuery.Width - picQuery.Width) / 2
    picQuery.Top = (tcFrmQuery.Height - picQuery.Height) / 2
    
    dtpStart.value = Now - 7
    dtpEnd.value = Now
    cbxQueryType.ListIndex = 0
    
    tcFrmQuery.Visible = True
    tcFrmQuery.Translucence
End Sub

Private Sub CloseQueryWindow()
    tcFrmQuery.Visible = False
End Sub

Private Sub cmdNotOk_Click()
On Error GoTo errHandle
    Call CloseQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdOk_Click()
On Error GoTo errHandle
    Dim strFilter As String
    Dim strQueryType As String
    
    If Trim(txtQueryValue.Text) <> "" Then
        Select Case cbxQueryType.Text
            Case "�� �� ��"
                strQueryType = "c.����"
            Case "�� �� ��"
                strQueryType = "d.�����"
            Case "ס Ժ ��"
                strQueryType = "d.סԺ��"
            Case "�� �� ��"
                strQueryType = "d.������"
            Case "���￨��"
                strQueryType = "d.���￨��"
            Case "IC �� ��"
                strQueryType = "d.IC����"
            Case "ҽ �� ��"
                strQueryType = "d.ҽ����"
        End Select
        
        strFilter = strQueryType & "='" & txtQueryValue.Text & "'"
    Else
        strFilter = " a.���� like '" & txtName.Text & "%' and b.����ʱ�� between " & To_Date(dtpStart.value) & " and " & To_Date(dtpEnd.value)
    End If
    
    Call LoadStudyData(strFilter)
    
    Call CloseQueryWindow
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    If lvwStudy.ListItems.Count <= 0 Then
        Call MsgboxCus("û�пɽ��вɼ��ļ�����ݡ�", vbOKOnly, G_STR_HINT_TITLE)
        Exit Sub
    End If
    
    If lvwStudy.SelectedItem Is Nothing Then
        Call MsgboxCus("��ѡ����Ҫ���вɼ��ļ�����ݡ�", vbOKOnly, gstrSysName)
        Exit Sub
    End If

    blnOk = True
    Call Me.Hide
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOKOnly, G_STR_HINT_TITLE
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '�������ö�
    
    Call zlCL_RestoreWinState(Me, App.ProductName)
    
    Call LoadStudyData

End Sub

Private Function GetColumnIndex(ByVal strColumnCaption As String) As Long
    Dim i As Long
    
    For i = 1 To lvwStudy.ColumnHeaders.Count
        If UCase(lvwStudy.ColumnHeaders(i).Text) = UCase(strColumnCaption) Then Exit For
    Next i
    
    GetColumnIndex = i - 1
End Function

Private Sub LoadStudyData(Optional strFilter As String = "")
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strCurFilter As String
    
    strCurFilter = strFilter
    If Trim(strCurFilter) = "" Then
        strCurFilter = "b.ִ�й��� in(2,3) and b.�״�ʱ�� between sysdate - 3 and sysdate"
    End If
    
    strSQL = "select /*+ Rule*/ a.id,b.���ͺ�, a.����, a.�Ա�, a.����, a.������Դ, e.���� as ���˿���, a.ҽ������, " & _
                    "Decode(a.������Դ,3,a.����ҽ��,b.������) ������, b.����ʱ�� as ����ʱ��, c.������, b.�״�ʱ�� as ����ʱ��, " & _
                    "Decode(a.������Դ,2,d.סԺ��,d.�����) ��ʶ��, c.Ӱ�����,c.����, nvl(b.ִ�й���,0) as ������ " & _
            "from ����ҽ����¼ a, ����ҽ������ b, Ӱ�����¼ c, ������Ϣ d, ���ű� e " & _
            "where a.ID=b.ҽ��id and b.ҽ��id=c.ҽ��Id(+) and a.����id=d.����id and a.���˿���id=e.id and a.���ID is null and a.ִ�п���ID=" & glngDepartId & IIf(strCurFilter <> "", " and ", "") & strCurFilter
            
                
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, Me.Caption)
    
    lvwStudy.ListItems.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        Call SetListItemData(rsData)
        Call rsData.MoveNext
    Wend
    
End Sub

Private Sub SetListItemData(rsCurData As ADODB.Recordset)
    Dim objNewItem As ListItem
    Dim objNewSubItem As ListSubItem
 
    Set objNewItem = lvwStudy.ListItems.Add
    
    objNewItem.Text = Nvl(rsCurData!����)
    objNewItem.Icon = 2
    objNewItem.SmallIcon = 2

    objNewItem.Tag = Nvl(rsCurData!ID)

    
    objNewItem.SubItems(GetColumnIndex("�Ա�")) = Nvl(rsCurData!�Ա�)
    objNewItem.SubItems(GetColumnIndex("����")) = Nvl(rsCurData!����)
    
    objNewItem.SubItems(GetColumnIndex("����")) = Nvl(rsCurData!����)
    
    objNewItem.SubItems(GetColumnIndex("��ʶ��")) = Nvl(rsCurData!��ʶ��)
    If Nvl(rsCurData!������Դ) = 2 Then
        objNewItem.ListSubItems(GetColumnIndex("��ʶ��")).ReportIcon = 1
    End If
    
    objNewItem.SubItems(GetColumnIndex("Ӱ�����")) = Nvl(rsCurData!Ӱ�����)
    objNewItem.SubItems(GetColumnIndex("ҽ������")) = Nvl(rsCurData!ҽ������)
    objNewItem.SubItems(GetColumnIndex("���״̬")) = Decode(Nvl(rsCurData!������), -1, "�Ѳ���", 0, "�ѵǼ�", 1, "�ѵǼ�", 2, "�ѱ���", 3, "�Ѽ��", 4, "�����", 5, "�����", "�����")
    objNewItem.SubItems(GetColumnIndex("���˿���")) = Nvl(rsCurData!���˿���)
    objNewItem.SubItems(GetColumnIndex("����ʱ��")) = Format(Nvl(rsCurData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
    objNewItem.SubItems(GetColumnIndex("����ʱ��")) = Format(Nvl(rsCurData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
    objNewItem.SubItems(GetColumnIndex("������")) = Nvl(rsCurData!������)
    objNewItem.SubItems(GetColumnIndex("������")) = Nvl(rsCurData!������)

    objNewItem.ListSubItems(1).Tag = Nvl(rsCurData!���ͺ�)
    objNewItem.ListSubItems(2).Tag = Nvl(rsCurData!������)
End Sub


Private Sub Form_Resize()
On Error GoTo errHandle
    lvwStudy.Left = 120
    lvwStudy.Top = 120
    lvwStudy.Height = Me.ScaleHeight - picPanel.Height - 120
    lvwStudy.Width = Me.ScaleWidth - 240

    cmdCancel.Left = picPanel.Width - cmdCancel.Width - 120
    cmdSure.Left = cmdCancel.Left - cmdSure.Width + 10
    cmdFind.Left = cmdSure.Left - cmdFind.Width + 20
    
    Exit Sub
errHandle:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call zlCL_SaveWinState(Me, App.ProductName)
End Sub


Private Sub lvwStudy_DblClick()
On Error GoTo errHandle
    If lvwStudy.SelectedItem Is Nothing Then Exit Sub
    
    Call cmdSure_Click
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOK, G_STR_HINT_TITLE
End Sub

Private Sub lvwStudy_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo errHandle
    Dim lngCurAdviceId As Long
    
    lngCurAdviceId = Item.Tag
    
    Call ConfigAdviceInf(lngCurAdviceId, Item.ListSubItems(GetColumnIndex("ҽ������")).Text)
Exit Sub
errHandle:
    MsgboxCus err.Description, vbOK, G_STR_HINT_TITLE
End Sub


Private Sub ConfigAdviceInf(ByVal lngAdviceId As Long, ByVal strAdviceContext As String)
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngMoneyState As Long
    
    strSQL = "select a.id as ҽ��ID, nvl(a.���Id, 0) as ���ID, b.�Ʒ�״̬,b.��¼���� " & _
            " from ����ҽ����¼ a, ����ҽ������ b " & _
            " where a.Id=b.ҽ��ID and (a.id=[1] or a.���ID=[1])"
    Set rsData = zlCL_GetDBObj.OpenSQLRecord(strSQL, Me.Caption, lngAdviceId)
    
    picInf.Visible = False
        
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngMoneyState = GetMoneyState(rsData)
    Select Case lngMoneyState
        Case 0
            labMoneyState.Caption = "Ƿ"
            labMoneyState.ForeColor = &H80FF&
        Case 1
            labMoneyState.Caption = "��"
            labMoneyState.ForeColor = &HC000&
        Case 2
            labMoneyState.Caption = "��"
            labMoneyState.ForeColor = &HFF0000
        Case 3
    End Select
    
    labAdviceInf.Caption = strAdviceContext
    
    picInf.Visible = True
End Sub



Private Function GetMoneyState(rsData As ADODB.Recordset) As Long
    '�ж��Ƿ��Ѿ��շ�
    '"����ҽ������.��¼����"--- 1���շѵģ�2�Ǽ��ʵġ�

    'ͨ��"����ҽ������.�Ʒ�״̬"ֱ���ж�,ԭ��ֵ��-1-����Ʒ�;0-δ�Ʒ�;1-�ѼƷѣ����ڼ��ʵ�������������ʵ���������ԭ��ֵ���䡣
    '�����շѵ��ķ��ͼ�¼����������״̬��2-�����շѣ�3-ȫ���շ�

    'û�ж�Ӧ���õ�ҽ�������������һ����"-1-����Ʒ�"����û�������շѶ��գ�һ����"0-δ�Ʒ�"������Ȼ�������շѶ��գ�������Ϊ���ͺ��ֹ��Ʒѣ�����ҽ������ȥ���ɡ�
    '"1-�ѼƷ�"���Ƿ���ʱ�����˷��õġ��������˷��õ��ݲ���ʾ�շ��ˣ����ɿ����Ǽ��ʻ��۵������շѻ��۵��������շѻ��۵��Ͷ�����״̬��
    '"2-�����շ�"��ʾ�����շѺͲ����˷ѵ����������û�յ��ꡣ

    '���շ���ʾ״̬�����շѣ��޷��ã�δ�շѣ�
    'δ�շ�----
    '1����ҽ�����շѵ��ģ���������������δ�շ�
    '   (1)��һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (1,2)��δ�շ� ------����¼����=1 and �Ʒ�״̬ in (1,2)��
    '���շѣ�
    '1����ҽ���Ǽ��˵����շ�-------����¼����=2��
    '2����ҽ�����շѵ��ģ����������������շ�
    '   (1)�ų�δ�շѺ���һ����ҽ���Ͳ�λҽ���� �Ʒ�״̬ =3 ���շ�-----����¼����=1 and �Ʒ�״̬ = 3��
    '�޷���
    '1����ҽ�����շѵ��ģ����������������޷���
    '   (1)������ҽ���Ͳ�λҽ���� �Ʒ�״̬ in (-1,0)���޷��� ------����¼����=1 and �Ʒ�״̬ in (-1,0)��


    ' intCharged  '0--δ�շѣ�1--���շѣ�2--�޷���
    Dim lngTempCharged As Long

    lngTempCharged = 2 '�޷���
    
    rsData.Filter = "���Id = 0"

    If Nvl(rsData!��¼����, 2) = 2 Then
        'סԺ�ǼǵĲ��ˣ����û�мƷѣ����Ϊ�޷���
        If Nvl(rsData!�Ʒ�״̬, -1) = 0 Then
            lngTempCharged = 2
        Else
            lngTempCharged = 1  '���շ�
        End If
    Else
        If Nvl(rsData!�Ʒ�״̬, -1) = 1 Or Nvl(rsData!�Ʒ�״̬, -1) = 2 Then
            lngTempCharged = 0      'δ�շ�
        Else        '��ҽ���ļƷ�״̬�� -1,0,3  ��3--���շѣ�-1��0--�޷��ã�
            '��ѯ��ҽ��δ�Ʒѻ����Ѿ��շ��ˣ���Ҫ�鲿λҽ�����շ����������ҽ�����Ѿ��շѣ��������շ�

            '��������������շѵģ��ȼ�¼�����շ�
            If Nvl(rsData!�Ʒ�״̬, -1) = 3 Then
                lngTempCharged = 1      '���շ�
            End If

            rsData.Filter = "���ID <> 0 "
            Do While rsData.EOF = False
                If Nvl(rsData!�Ʒ�״̬, -1) = 1 Or Nvl(rsData!�Ʒ�״̬, -1) = 2 Then
                    lngTempCharged = 0      'δ�շ�

                    Exit Do
                ElseIf Nvl(rsData!�Ʒ�״̬, -1) = 3 Then
                    lngTempCharged = 1      '���շ�
                End If

                rsData.MoveNext
            Loop

        End If
    End If

    GetMoneyState = lngTempCharged
End Function





















