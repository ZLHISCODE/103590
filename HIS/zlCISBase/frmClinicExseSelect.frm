VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicExseSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ����"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   Icon            =   "frmClinicExseSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2700
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleMode       =   0  'User
      ScaleWidth      =   33.75
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   45
   End
   Begin VB.CheckBox ChkDown 
      Caption         =   "��ʾ�¼�Ŀ¼����"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   2025
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   8160
      TabIndex        =   3
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   6840
      TabIndex        =   2
      Top             =   4440
      Width           =   1100
   End
   Begin MSComctlLib.ListView LivMain 
      Height          =   4035
      Left            =   3360
      TabIndex        =   0
      Top             =   330
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img16"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "���㵥λ"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "�ۼ�"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.TreeView LvwMain 
      Height          =   4035
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   7117
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2700
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicExseSelect.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��һ����Ŀ��Ȼ����ȷ��"
      Height          =   180
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   2520
   End
End
Attribute VB_Name = "frmClinicExseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MouseStartX As Single                       '�ƶ�ǰ����λ��
Dim NowIndex As Long                            '��ǰλ��
Private mstr������� As String

Sub LoadTree()
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select level as ����,8 ����,Id, �ϼ�id, ����, ����" & _
             " From �շѷ���Ŀ¼ " & _
             " Start With �ϼ�id Is Null " & _
             " Connect By Prior Id = �ϼ�id " & _
             " Union " & _
             " Select level as ����,����,Id, �ϼ�id, ����, ���� " & _
             " From ���Ʒ���Ŀ¼ " & _
             " Where ���� in(1,2,3,7) " & _
             " Start With �ϼ�id Is Null " & _
             " Connect By Prior Id = �ϼ�id " & _
             " Order By ����,����,���� "

    zlDatabase.OpenRecordset rsTmp, gstrSql, "ѡ����"
    
    '�����
    LvwMain.Nodes.Clear
    LvwMain.Nodes.Add , , "Root", "�����շ���Ŀ", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C1", "[1]����ҩ", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C2", "[2]�гɷ�", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C3", "[3]�в�ҩ", 1, 1
    LvwMain.Nodes.Add "Root", tvwChild, "C7", "[7]��������", 1, 1
    
    With rsTmp
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                If !���� <> 8 Then
                    LvwMain.Nodes.Add "C" & Val(!����), tvwChild, "C" & Val(!����) & Val(!ID), "[" & !���� & "]" & !����, 1, 1
                Else
                    LvwMain.Nodes.Add "Root", tvwChild, "C" & Val(!����) & Val(!ID), "[" & !���� & "]" & !����, 1, 1
                End If
            Else
                LvwMain.Nodes.Add "C" & Val(!����) & Val(!�ϼ�ID), tvwChild, "C" & Val(!����) & Val(!ID), "[" & !���� & "]" & !����, 1, 1
            End If
            .MoveNext
        Loop
    End With

    rsTmp.Close
    
    Dim nod As Node
    On Error Resume Next
    Set nod = LvwMain.Nodes(strKey)
    If Err <> 0 Then
        Set nod = LvwMain.Nodes("Root").Child
        nod.Selected = True
        nod.Expanded = True
        LvwMain_NodeClick nod
        NowIndex = nod.Index
    Else
        Err.Clear
        nod.Selected = True
        nod.Expanded = True
        nod.EnsureVisible
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ChkDown_Click()
    Dim nod As Node
    Set nod = Me.LvwMain.Nodes(Me.LvwMain.SelectedItem.Index)
    LvwMain_NodeClick nod
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    If Me.LivMain.ListItems.Count > 0 Then
        gstrSql = "select * from �շ���ĿĿ¼ where id = " & Mid(Me.LivMain.SelectedItem.Key, 2)
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Mid(Me.LivMain.SelectedItem.Key, 2)))
        Set frmClinicExse.rsSelect = rsTmp
    End If
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    LoadTree
End Sub

Private Sub Form_Resize()
    
    'LvwMain
    Me.LvwMain.Width = Me.picSplit.Left
    
    'picSplit
    Me.picSplit.Top = Me.LvwMain.Top
    Me.picSplit.Height = Me.LvwMain.Height
    
    'Livmain
    Me.LivMain.Top = Me.LvwMain.Top
    Me.LivMain.Left = Me.picSplit.Left + Me.picSplit.Width
    Me.LivMain.Height = Me.LvwMain.Height
    Me.LivMain.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    
End Sub
Private Sub LivMain_DblClick()
    cmdOK_Click
End Sub

Private Sub LvwMain_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rsTmp As New ADODB.Recordset
    Dim ItmX As ListItem
    Dim str��� As String
    Dim str���� As String
    Dim strTemp As String
    Dim i As Integer
    
    If Node.Key = "Root" Then Exit Sub
    On Error GoTo errHandle
    str��� = Mid(Node.Key, 2, 1)
    str���� = Mid(Node.Key, 3)
    
    
    If str��� <> "8" Then
        For i = 0 To UBound(Split(mstr�������, ","))
            strTemp = IIf(strTemp = "", "(A.�������=" & Split(mstr�������, ",")(i), strTemp & " or A.�������=" & Split(mstr�������, ",")(i))
        Next
        strTemp = strTemp & ")"
        If Len(Node.Key) <= 2 Then Exit Sub
        If Me.ChkDown.Value = 0 Then
            gstrSql = "Select A.Id,C.����id,A.����, A.����, A.���,  A.����, A.���㵥λ," & _
                    " Decode(Nvl(A.�Ƿ���,0),0,ltrim(rtrim(to_char(nvl(D.�ּ�,0),'9999999990.0000'))),'ʱ��') As �ۼ� " & _
                    " From �շ���ĿĿ¼ A," & IIf(str��� <> "7", "ҩƷ���", "��������") & " B,������ĿĿ¼ C,�շѼ�Ŀ D " & _
                    " Where A.ID=D.�շ�ϸĿID(+) And a.ID = b." & IIf(str��� <> "7", "ҩƷid", "����id") & " And b." & IIf(str��� <> "7", "ҩ��id", "����id") & " = C.ID And C.����id = [1] " & _
                    " and (A.����ʱ�� is null or A.����ʱ�� = to_date('3000-01-01', 'YYYY-MM-DD')) " & _
                    " And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)  And D.�۸�ȼ� Is Null " & _
                    " And " & strTemp & _
                    " Order By A.����"
        Else
            gstrSql = "Select a.Id, c.����id, a.����, A.����, A.���, A.����, a.���㵥λ, " & _
                    " Decode(Nvl(A.�Ƿ���,0),0,ltrim(rtrim(to_char(nvl(e.�ּ�,0),'9999999990.0000'))),'ʱ��') As �ۼ� " & _
                    " From �շ���ĿĿ¼ a," & IIf(str��� <> "7", "ҩƷ���", "��������") & " B,������ĿĿ¼ c,�շѼ�Ŀ e, " & _
                    " (Select * From ���Ʒ���Ŀ¼ " & _
                    " Start With �ϼ�id = [1] " & _
                    " Connect By Prior Id = �ϼ�id " & _
                    " Union " & _
                    " Select * From ���Ʒ���Ŀ¼ Where Id = [1]) d " & _
                    " Where A.ID=e.�շ�ϸĿID(+) And a.ID = b." & IIf(str��� <> "7", "ҩƷid", "����id") & " And b." & IIf(str��� <> "7", "ҩ��id", "����id") & " = c.ID And c.����id = d.ID " & _
                    " and (A.����ʱ�� is null or A.����ʱ�� = to_date('3000-01-01', 'YYYY-MM-DD')) " & _
                    " And e.ִ������ <= SYSDATE AND (e.��ֹ���� > SYSDATE OR e.��ֹ���� IS NULL)   And e.�۸�ȼ� Is Null " & _
                    " And " & strTemp & _
                    " Order By a.���� "
        End If
        

    Else
        For i = 0 To UBound(Split(mstr�������, ","))
            strTemp = IIf(strTemp = "", "(I.�������=" & Split(mstr�������, ",")(i), strTemp & " or I.�������=" & Split(mstr�������, ",")(i))
        Next
        strTemp = strTemp & ")"
        If Me.ChkDown.Value = 0 Then
            gstrSql = " Select 1 As ĩ��,I.ID,����ID,I.����,I.����,I.���,I.����, I.���㵥λ," & _
                    " Decode(Nvl(I.�Ƿ���,0),0,ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.0000'))),ltrim(rtrim(to_char(Sum(nvl(D.ȱʡ�۸�,0)),'9999999990.0000')))) As �ۼ� " & _
                    " from �շ���ĿĿ¼ I,�շѼ�Ŀ D " & _
                    " where I.ID=D.�շ�ϸĿID(+) And I.��� not in ('1','J')" & _
                    " and ����ID = [1] " & _
                    " and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
                    " And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)  And D.�۸�ȼ� Is Null " & _
                    " And " & strTemp & _
                    " Group By i.Id, ����id, i.����, i.����, i.���, i.����, i.���㵥λ, i.�Ƿ��� " & _
                    " Order By ���� "
        Else
            gstrSql = "select b.id , b.���� , b.���� ,b.���,b.����, b.���㵥λ,b.�ۼ� " & _
                    " From " & _
                    "     (select * from �շѷ���Ŀ¼ " & _
                    "        start with �ϼ�id = [1] " & _
                    "        connect by prior id = �ϼ�id " & _
                    "      Union " & _
                    "      select * from �շѷ���Ŀ¼ " & _
                    "        where id = [1] )  a , " & _
                    " (Select 1 As ĩ��,I.ID,����ID,I.����,I.����,I.���,I.����, I.���㵥λ, " & _
                    " Decode(Nvl(I.�Ƿ���,0),0,ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.0000'))),ltrim(rtrim(to_char(Sum(nvl(D.ȱʡ�۸�,0)),'9999999990.0000')))) As �ۼ� " & _
                    " from �շ���ĿĿ¼ I,�շѼ�Ŀ D " & _
                    " Where I.ID=D.�շ�ϸĿID(+) And " & _
                    " I.��� not in ('1','J') and " & _
                    " (I.����ʱ�� is null or I.����ʱ�� = to_date('3000-01-01', 'YYYY-MM-DD')) And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)  And D.�۸�ȼ� Is Null  " & _
                    " And " & strTemp & _
                    " Group By i.Id, ����id, i.����, i.����, i.���, i.����, i.���㵥λ, i.�Ƿ��� ) b " & _
                    " Where b.����id = a.ID  Order By b.���� "
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ѡ����", str����, mstr�������, gstrPriceClass)
    
    '���
    Me.LivMain.ListItems.Clear
    
    Me.MousePointer = 11
     
    
    Do Until rsTmp.EOF
        Set ItmX = Me.LivMain.ListItems.Add(, "A" & rsTmp("ID"), rsTmp("����"), 1, 1)
        ItmX.SubItems(1) = Nvl(rsTmp("����"))
        ItmX.SubItems(2) = Nvl(rsTmp("���"))
        ItmX.SubItems(3) = Nvl(rsTmp("����"))
        ItmX.SubItems(4) = Nvl(rsTmp("���㵥λ"))
        ItmX.SubItems(5) = Nvl(rsTmp("�ۼ�"))
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    
    Me.MousePointer = 1
    
    NowIndex = Node.Index
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MouseStartX = X
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MoveTmp As Single
    '��ʱ���η��������
    On Error Resume Next
    If Button = 1 Then
        
        '�õ��ƶ����λ��
        MoveTmp = Me.picSplit.Left + X - MouseStartX
        
        '����������С���ʱ�˳�
        If MoveTmp <= 2000 Or Me.ScaleWidth - MoveTmp <= 2000 Then Exit Sub
        
        'picSplit
        picSplit.Left = MoveTmp
        
        'LvwMain
        Me.LvwMain.Width = Me.picSplit.Left
        
        'LivMain
        Me.LivMain.Left = Me.picSplit.Left + Me.picSplit.Width
        Me.LivMain.Width = Me.ScaleWidth - Me.picSplit.Left - Me.picSplit.Width
    End If
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal str������� As String)
    mstr������� = str�������
    
    Me.Show 1, frmParent
End Sub
