VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H80000005&
   Caption         =   "�˵�����滮"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmMenu.frx":0000
   ScaleHeight     =   6900
   ScaleWidth      =   6960
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView tvwTemp 
      Height          =   1215
      Left            =   60
      TabIndex        =   25
      Top             =   1620
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   5940
      Top             =   5850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "����(&I)"
      Height          =   350
      Left            =   3990
      TabIndex        =   5
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "����(&X)"
      Height          =   350
      Left            =   5160
      TabIndex        =   6
      Top             =   1980
      Width           =   1100
   End
   Begin VB.TextBox txtShort 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1860
      MaxLength       =   30
      TabIndex        =   20
      Top             =   6180
      Width           =   1905
   End
   Begin MSComctlLib.ImageCombo icIcon 
      Height          =   315
      Left            =   4740
      TabIndex        =   22
      Top             =   5820
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      ImageList       =   "ils32"
   End
   Begin VB.CommandButton cmdDrop 
      Caption         =   "ɾ��(&E)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2100
      TabIndex        =   4
      Top             =   1980
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7050
      Top             =   6300
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
            Picture         =   "frmMenu.frx":04F9
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtExplain 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1860
      MaxLength       =   250
      TabIndex        =   24
      Top             =   6540
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1860
      MaxLength       =   30
      TabIndex        =   18
      Top             =   5820
      Width           =   1905
   End
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   5855
      Picture         =   "frmMenu.frx":228B
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "���ӽڵ�"
      Top             =   3000
      Width           =   405
   End
   Begin VB.CommandButton cmdEdit 
      DisabledPicture =   "frmMenu.frx":238D
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   5855
      Picture         =   "frmMenu.frx":272A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "˳�򿿺�"
      Top             =   5145
      Width           =   405
   End
   Begin VB.CommandButton cmdEdit 
      DisabledPicture =   "frmMenu.frx":2CB4
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   5855
      Picture         =   "frmMenu.frx":3053
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "˳��ǰ"
      Top             =   4725
      Width           =   405
   End
   Begin VB.CommandButton cmdEdit 
      DisabledPicture =   "frmMenu.frx":35DD
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5855
      Picture         =   "frmMenu.frx":3977
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "�½�һ��"
      Top             =   4320
      Width           =   405
   End
   Begin MSComctlLib.TreeView tvwMenu 
      Height          =   2685
      Left            =   990
      TabIndex        =   10
      Top             =   3015
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4736
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ils16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   7020
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3F01
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "����(&A)"
      Height          =   350
      Left            =   930
      TabIndex        =   3
      Top             =   1980
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3960
      TabIndex        =   7
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "��ԭ(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5160
      TabIndex        =   8
      Top             =   2580
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwMenu 
      Height          =   1005
      Left            =   965
      TabIndex        =   2
      Top             =   900
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1773
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_�˵���"
         Object.Tag             =   "�˵���"
         Text            =   "�˵���"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdEdit 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   5855
      Picture         =   "frmMenu.frx":5C93
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "ɾ���ڵ�"
      Top             =   3405
      Width           =   405
   End
   Begin VB.CommandButton cmdEdit 
      DisabledPicture =   "frmMenu.frx":5D95
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   5855
      Picture         =   "frmMenu.frx":612A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "����һ��"
      Top             =   3900
      Width           =   405
   End
   Begin VB.Line linSplit 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   930
      X2              =   8190
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Line linSplit 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   930
      X2              =   8190
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label lblIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͼ��(&I)"
      Height          =   180
      Left            =   4050
      TabIndex        =   21
      Top             =   5880
      Width           =   630
   End
   Begin VB.Label lblShort 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�̱���(&H)"
      Height          =   180
      Left            =   960
      TabIndex        =   19
      Top             =   6240
      Width           =   810
   End
   Begin VB.Label lblExplain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&X)"
      Height          =   180
      Left            =   960
      TabIndex        =   23
      Top             =   6600
      Width           =   630
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&T)"
      Height          =   180
      Left            =   960
      TabIndex        =   17
      Top             =   5880
      Width           =   630
   End
   Begin VB.Label lblConstruct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˵�����(&M)"
      Height          =   180
      Left            =   960
      TabIndex        =   9
      Top             =   2730
      Width           =   990
   End
   Begin VB.Label lblMenu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˵���(&G)"
      Height          =   180
      Left            =   960
      TabIndex        =   1
      Top             =   630
      Width           =   810
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmMenu.frx":66B4
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�˵�����滮"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   1440
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum const�༭��ʽ
    con�� = 0
    con�� = 1
    con�� = 2
    con�� = 3
    
    con���� = 4
    conɾ�� = 5
End Enum

Dim mblnModify As Boolean
Dim mstrMenuName As String

Private Sub cmdEdit_Click(Index As Integer)
    Dim nodNew As Node
    Dim nodTemp As Node, nod As Node
    
    LockWindowUpdate tvwMenu.hwnd
    Select Case Index
        Case con�� '����
            '����Ϊ�丸�ڵ㴴��һ������һģһ�����ֵܽڵ�
            Set nodNew = tvwMenu.Nodes.Add(tvwMenu.SelectedItem.Parent.Index, tvwNext, , _
                tvwMenu.SelectedItem.Text, tvwMenu.SelectedItem.Image, tvwMenu.SelectedItem.SelectedImage)
            nodNew.Tag = tvwMenu.SelectedItem.Tag
            'Ȼ������ĺ���ȫ���Ƶ��½ڵ���
            Set nodTemp = tvwMenu.SelectedItem.Child
            Do Until nodTemp Is Nothing
                Set nod = nodTemp
                Set nodTemp = nodTemp.Next
                Set nod.Parent = nodNew
            Loop
            nodNew.Expanded = tvwMenu.SelectedItem.Expanded
            tvwMenu.Nodes.Remove tvwMenu.SelectedItem.Index
        Case con�� '����
            '����Ϊ��ǰһ�ڵ㴴��һ������һģһ�����ӽڵ�
            Set nodNew = tvwMenu.Nodes.Add(tvwMenu.SelectedItem.Previous.Index, tvwChild, , _
                tvwMenu.SelectedItem.Text, tvwMenu.SelectedItem.Image, tvwMenu.SelectedItem.SelectedImage)
            nodNew.Tag = tvwMenu.SelectedItem.Tag
            'Ȼ������ĺ���ȫ���Ƶ��½ڵ���
            Set nodTemp = tvwMenu.SelectedItem.Child
            Do Until nodTemp Is Nothing
                Set nod = nodTemp
                Set nodTemp = nodTemp.Next
                Set nod.Parent = nodNew
            Loop
            nodNew.Expanded = tvwMenu.SelectedItem.Expanded
            tvwMenu.Nodes.Remove tvwMenu.SelectedItem.Index
        Case con�� '����
            '����Ϊ��ǰһ�ڵ㴴��һ������һģһ���ĸ��ڵ�
            Set nodNew = tvwMenu.Nodes.Add(tvwMenu.SelectedItem.Previous.Index, tvwPrevious, , _
                tvwMenu.SelectedItem.Text, tvwMenu.SelectedItem.Image, tvwMenu.SelectedItem.SelectedImage)
            nodNew.Tag = tvwMenu.SelectedItem.Tag
            'Ȼ������ĺ���ȫ���Ƶ��½ڵ���
            Set nodTemp = tvwMenu.SelectedItem.Child
            Do Until nodTemp Is Nothing
                Set nod = nodTemp
                Set nodTemp = nodTemp.Next
                Set nod.Parent = nodNew
            Loop
            nodNew.Expanded = tvwMenu.SelectedItem.Expanded
            tvwMenu.Nodes.Remove tvwMenu.SelectedItem.Index
        Case con�� '����
            '����Ϊ���һ�ڵ㴴��һ������һģһ���ĵܵܽڵ�
            Set nodNew = tvwMenu.Nodes.Add(tvwMenu.SelectedItem.Next.Index, tvwNext, , _
                tvwMenu.SelectedItem.Text, tvwMenu.SelectedItem.Image, tvwMenu.SelectedItem.SelectedImage)
            nodNew.Tag = tvwMenu.SelectedItem.Tag
            'Ȼ������ĺ���ȫ���Ƶ��½ڵ���
            Set nodTemp = tvwMenu.SelectedItem.Child
            Do Until nodTemp Is Nothing
                Set nod = nodTemp
                Set nodTemp = nodTemp.Next
                Set nod.Parent = nodNew
            Loop
            nodNew.Expanded = tvwMenu.SelectedItem.Expanded
            tvwMenu.Nodes.Remove tvwMenu.SelectedItem.Index
        Case con���� '����
            '����Ϊ�䴴��һ�����ӽڵ�
            Set nodNew = tvwMenu.Nodes.Add(tvwMenu.SelectedItem.Index, tvwChild, , "�����ڵ�", "K99", "K99")
            nodNew.Tag = "'''�����ڵ�'"
        Case conɾ�� 'ɾ��
            If tvwMenu.SelectedItem.Previous Is Nothing Then
                Set nodNew = tvwMenu.SelectedItem.Parent
            Else
                Set nodNew = tvwMenu.SelectedItem.Previous
            End If
            tvwMenu.Nodes.Remove tvwMenu.SelectedItem.Index
    End Select
    LockWindowUpdate 0
    nodNew.Selected = True
    nodNew.EnsureVisible
    mblnModify = True
    Call SetEnable
    If Index = 4 Then
        txtName.SetFocus
    Else
        tvwMenu.SetFocus
    End If
End Sub

Private Sub cmdExp_Click()
    'û��װϵͳ��ֻ�й������ݾͿ���������
    If lvwMenu.SelectedItem Is Nothing Then Exit Sub
    
    '--------����ļ���
    cdlFile.CancelError = True
    cdlFile.Filter = "�˵���ϵ�ļ�(*.ZLM)|*.ZLM"
    cdlFile.Flags = cdlOFNOverwritePrompt
    cdlFile.FileName = lvwMenu.SelectedItem.Text & ".ZLM"
    
    On Error Resume Next
    cdlFile.ShowSave
    If Err <> 0 Then Exit Sub
    
    '--------��ʽ����
    Dim rsTemp As New ADODB.Recordset
    Dim strModule As String
    Dim strIcon As String
    Dim nod As MSComctlLib.Node
    Dim objSys As New Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Dim arrֵ As Variant
    Dim int���� As Integer   '��ǰ����
    Dim strInsert As String
    
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_menu_tree", lvwMenu.SelectedItem.Text)
    '���ȵ��������б�
    tvwTemp.Nodes.Clear
    tvwTemp.Nodes.Add , , "Root", lvwMenu.SelectedItem.Text, "Root", "Root"
    Do Until rsTemp.EOF
        strModule = IIf(IsNull(rsTemp("ģ��")), "", rsTemp("ģ��"))
        
        If IsNull(rsTemp("ͼ��")) Or rsTemp("ͼ��") = 0 Then
            strIcon = IIf(strModule = "", "K99", "K100")
        Else
            strIcon = "K" & rsTemp("ͼ��")
        End If
        
        If IsNull(rsTemp("�ϼ�ID")) Then
            Set nod = tvwTemp.Nodes.Add("Root", tvwChild, "C" & rsTemp("ID"), rsTemp("����"), strIcon, strIcon)
        Else
            Set nod = tvwTemp.Nodes.Add("C" & rsTemp("�ϼ�ID"), tvwChild, "C" & rsTemp("ID"), rsTemp("����"), strIcon, strIcon)
        End If
        nod.Tag = strModule & "'" & IIf(IsNull(rsTemp("���")), "", rsTemp("���")) & _
                    "'" & IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��")) & _
                    "'" & IIf(IsNull(rsTemp("�̱���")), "", rsTemp("�̱���")) & _
                    "'" & IIf(IsNull(rsTemp("ϵͳ")), "", rsTemp("ϵͳ"))
        rsTemp.MoveNext
    Loop
    '�ٴ�����ļ�
    Set objFile = objSys.CreateTextFile(cdlFile.FileName)
    
    Set nod = tvwTemp.Nodes("Root").Child
    Do Until nod Is Nothing
        arrֵ = Split(nod.Tag, "'")
        strInsert = "Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values([���],zlMenus_id.nextval," & _
                        "Null,'" & nod.Text & "','" & arrֵ(3) & "','" & arrֵ(1) & "'," & _
                        Mid(nod.Image, 2) & ",'" & arrֵ(2) & "'," & IIf(arrֵ(4) = "", "Null", arrֵ(4)) & "," & IIf(arrֵ(0) = "", "Null", arrֵ(0)) & ");"
        objFile.WriteLine (strInsert)
        
        Call ExportNode(objFile, nod, 1)
        Set nod = nod.Next
    Loop
    
    objFile.Close
    MsgBox "�˵���ϵ�ļ�����ɹ���", vbInformation, gstrSysName
End Sub

Private Function ExportNode(objFile As Scripting.TextStream, ByVal nod As Node, ByVal int���� As Integer) As Integer
'���ܣ��ݹ���ýڵ㵼��
'������objFile   ����ļ�
'      nod       �ϼ��ڵ�
'      int����   '��ǰ�ڵ�ļ���
'���أ��Ѿ������Ľڵ���
    Dim arrֵ As Variant
    Dim strInsert As String, int��� As Integer, intCount As Long
    
    int��� = 0
    Set nod = nod.Child
    Do Until nod Is Nothing
        int��� = int��� + 1
        arrֵ = Split(nod.Tag, "'")
        strInsert = "Insert Into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Values([���],zlMenus_id.nextval," & _
                        Space(int���� * 2) & "zlMenus_id.nextval-" & int��� & ",'" & nod.Text & "','" & arrֵ(3) & "','" & arrֵ(1) & "'," & _
                        Mid(nod.Image, 2) & ",'" & arrֵ(2) & "'," & IIf(arrֵ(4) = "", "Null", arrֵ(4)) & "," & IIf(arrֵ(0) = "", "Null", arrֵ(0)) & ");"
        objFile.WriteLine (strInsert)
        
        intCount = ExportNode(objFile, nod, int���� + 1)
        int��� = int��� + intCount
        Set nod = nod.Next
    Loop
    
    ExportNode = int���
End Function

Private Sub cmdImp_Click()
    'û��װϵͳ��ֻ�й������ݾͿ���������
    If MsgBox("��������ڵ�ϵͳ�����뵼��ʱ����ͬ�����ܵ����ʧ�ܡ�������ֹ��޸��ļ���ʵ�֡�" & vbCrLf & "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    
    '--------����ļ���
    Dim strName As String
    
    cdlFile.CancelError = True
    cdlFile.Filter = "�˵���ϵ�ļ�(*.ZLM)|*.ZLM"
    cdlFile.Flags = cdlOFNFileMustExist
    
    On Error Resume Next
    cdlFile.ShowOpen
    If Err <> 0 Then Exit Sub
    strName = Left(cdlFile.FileTitle, Len(cdlFile.FileTitle) - 4)
    If StrIsValid(strName, 30) = False Then
        Exit Sub
    End If
    If strName = "ȱʡ" Then
        MsgBox "�µ���Ĳ˵���ϵ���Ʋ���Ϊ��ȱʡ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '--------��ʽ����
    Dim objSys As New Scripting.FileSystemObject
    Dim objFile As Scripting.TextStream
    Dim strInsert As String
    
    gcnOracle.BeginTrans
    On Error GoTo ErrHandle
    
    Set objFile = objSys.OpenTextFile(cdlFile.FileName)
    
    Do Until objFile.AtEndOfStream
        strInsert = objFile.ReadLine
        gstrSQL = Trim(Replace(strInsert, "[���]", "'" & strName & "'"))
        If Right(gstrSQL, 1) = ";" Then gstrSQL = Left(gstrSQL, Len(gstrSQL) - 1)
        gcnOracle.Execute gstrSQL
    Loop
    gcnOracle.CommitTrans
    objFile.Close
    lvwMenu.ListItems.Add , , strName, "Root"
    MsgBox "�˵���ϵ��" & strName & "������ɹ���", vbInformation, gstrSysName
    Exit Sub
ErrHandle:
    gcnOracle.RollbackTrans
    MsgBox "�˵���ϵ��" & strName & "������ʧ�ܡ�", vbInformation, gstrSysName
End Sub

Private Sub cmdRestore_Click()
'��ԭ�˵���ϵ
    Call FillMenu
End Sub

Private Sub cmdSave_Click()
'����˵���ϵ
    If SaveMenu(tvwMenu.Nodes("Root").Text) = True Then
        mblnModify = False
        SetEnable
    End If
End Sub

Private Sub cmdNew_Click()
    Dim strMenuName As String
    Dim rsTemp As New ADODB.Recordset
    
    strMenuName = frmNameEdit.GetName(name�˵�)
    If strMenuName = "" Then Exit Sub
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_Menu_Group", strMenuName)
    If rsTemp("����") > 0 Then
        MsgBox "����Ϊ��" & strMenuName & "���Ĳ˵����Ѿ����ڣ�����ʧ�ܡ�", vbExclamation, gstrSysName
        Exit Sub
    End If
    DoEvents
    If strMenuName = "" Then Exit Sub
    If SaveMenu(strMenuName) = True Then
        lvwMenu.ListItems.Add , , strMenuName, "Root"
    End If
    
End Sub

Private Sub cmdDrop_Click()
    Dim intIndex As Integer
    Dim strRemarks As String

    If MsgBox("��ȷʵҪɾ����" & lvwMenu.SelectedItem.Text & "���˵��飿", vbYesNo Or vbQuestion Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '��֤��ݲ��������˵��
        If Not CheckAuditStatus("0403", "ɾ��", strRemarks) Then Exit Sub
    On Error Resume Next
    gcnOracle.Execute "delete from zlmenus where ���='" & lvwMenu.SelectedItem.Text & "'"
    If Err <> 0 Then
        Err.Clear
        MsgBox "ɾ�������쳣������", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    '������Ҫ������־
    Call SaveAuditLog(3, "ɾ��", "ɾ���˵��顰" & lvwMenu.SelectedItem.Text & "��", strRemarks)
    intIndex = lvwMenu.SelectedItem.Index
    lvwMenu.ListItems.Remove intIndex
    If lvwMenu.ListItems.Count > 0 Then
        If intIndex > lvwMenu.ListItems.Count Then intIndex = lvwMenu.ListItems.Count
    
        lvwMenu.ListItems(intIndex).Selected = True
        Call FillMenu
    Else
        Call SetEnable
    End If
    
End Sub

Private Sub Form_Deactivate()
    If mblnModify = True Then
        If MsgBox("�Ѿ��޸��˲˵���Ĺ��ɣ���������棬�����Զ���ԭ��" & vbCr & "�Ƿ񱣴棿", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            Call cmdSave_Click
        Else
            Call cmdRestore_Click
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnModify = True Then
        If MsgBox("�Ѿ��޸��˲˵���Ĺ��ɣ���������棬�����Զ���ԭ��" & vbCr & "�Ƿ񱣴棿", vbQuestion Or vbYesNo Or vbDefaultButton2) = vbYes Then
            Call cmdSave_Click
        End If
    End If
    mblnModify = False
    mstrMenuName = ""
End Sub

Private Sub Form_Load()
'��ɳ�ʼ��
    Call InitIcon
    Call FillMenuName
    Call SetEnable
End Sub

Private Sub InitIcon()
'��ʼ��ͼ���װ��
    Dim i As Integer
    
    For i = 99 To 240
        ils16.ListImages.Add , "K" & i, LoadResPicture(i, vbResIcon)
        ils32.ListImages.Add , "K" & i, LoadResPicture(i, vbResIcon)
        icIcon.ComboItems.Add , "K" & i, , "K" & i, "K" & i
    Next
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    Dim i As Long
    
    '��С�ĸ߶�
    sngTemp = IIf(ScaleHeight > 6000, ScaleHeight, 6000)
    '�������ϼ��㶥��
    txtExplain.Top = sngTemp - txtExplain.Height - 200
    lblExplain.Top = txtExplain.Top + 60
    txtShort.Top = txtExplain.Top - txtShort.Height - 60
    lblShort.Top = txtShort.Top + 60
    txtName.Top = txtShort.Top - txtName.Height - 60
    lblName.Top = txtName.Top + 60
    icIcon.Top = txtName.Top
    lblIcon.Top = lblName.Top
    tvwMenu.Height = txtName.Top - tvwMenu.Top - 60
    
    '��С�Ŀ��
    sngTemp = IIf(ScaleWidth > 6000, ScaleWidth, 6000)
    lvwMenu.Width = sngTemp - lvwMenu.Left - 200
    linSplit(0).X2 = linSplit(0).X1 + lvwMenu.Width
    linSplit(1).X2 = linSplit(0).X2
    
    cmdEdit(0).Left = sngTemp - cmdEdit(0).Width - 200
    For i = 1 To 5
        cmdEdit(i).Left = cmdEdit(0).Left
    Next
    
    tvwMenu.Width = cmdEdit(0).Left - tvwMenu.Left - 200
    
    icIcon.Left = tvwMenu.Left + tvwMenu.Width - icIcon.Width
    lblIcon.Left = icIcon.Left - lblIcon.Width - 30
    
    txtName.Width = lblIcon.Left - txtName.Left - 200
    txtShort.Width = lblIcon.Left - txtShort.Left - 200
    txtExplain.Width = tvwMenu.Left + tvwMenu.Width - txtExplain.Left - 200
    
    cmdRestore.Left = lvwMenu.Left + lvwMenu.Width - cmdRestore.Width
    cmdSave.Left = cmdRestore.Left - cmdSave.Width - 30
    
    cmdExp.Left = cmdRestore.Left
    cmdImp.Left = cmdSave.Left
End Sub

Private Sub FillMenuName()
'װ��˵��������
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_Menu_Group", "")
    Do Until rsTemp.EOF
        lvwMenu.ListItems.Add , , rsTemp("���"), "Root"
        rsTemp.MoveNext
    Loop
    If lvwMenu.ListItems.Count > 0 Then lvwMenu.ListItems(1).Selected = True
    Call FillMenu
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillMenu()
'װ��һ������˵��Ĺ���
    Dim rsTemp As New ADODB.Recordset
    Dim strMenu As String
    Dim strModule As String
    Dim strIcon As String
    Dim nod As MSComctlLib.Node
    
    On Error GoTo ErrHandle
    tvwMenu.Nodes.Clear
    If lvwMenu.SelectedItem Is Nothing Then
        cmdNew.Enabled = False
        Exit Sub
    Else
        cmdNew.Enabled = True
    End If
    strMenu = lvwMenu.SelectedItem.Text
    tvwMenu.Nodes.Add , , "Root", strMenu, "Root", "Root"
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Popedom.Get_menu_tree", strMenu)
        
    Do Until rsTemp.EOF
        strModule = IIf(IsNull(rsTemp("ģ��")), "", rsTemp("ģ��"))
        
        If IsNull(rsTemp("ͼ��")) Or rsTemp("ͼ��") = 0 Then
            strIcon = IIf(strModule = "", "K99", "K100")
        Else
            strIcon = "K" & rsTemp("ͼ��")
        End If
        
        If IsNull(rsTemp("�ϼ�ID")) Then
            Set nod = tvwMenu.Nodes.Add("Root", tvwChild, "C" & rsTemp("ID"), rsTemp("����"), strIcon, strIcon)
        Else
            Set nod = tvwMenu.Nodes.Add("C" & rsTemp("�ϼ�ID"), tvwChild, "C" & rsTemp("ID"), rsTemp("����"), strIcon, strIcon)
        End If
        nod.Tag = strModule & "'" & IIf(IsNull(rsTemp("���")), "", rsTemp("���")) & _
                    "'" & IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��")) & _
                    "'" & IIf(IsNull(rsTemp("�̱���")), "", rsTemp("�̱���")) & _
                    "'" & IIf(IsNull(rsTemp("ϵͳ")), "", rsTemp("ϵͳ"))
        rsTemp.MoveNext
    Loop
    tvwMenu.Nodes(1).Selected = True
    tvwMenu.SelectedItem.Expanded = True
    mblnModify = False
    Call SetEnable
    Exit Sub
ErrHandle:
    If MsgBox("װ��˵�ʱ�������´���" & vbCrLf & vbCrLf & _
                Err.Description & vbCrLf & vbCrLf & "��Ҫ����һ����", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Sub

Private Function SaveMenu(ByVal str��� As String) As Boolean
'���ݵ�ǰ�˵����ɲ����˵���ϵ
    On Error GoTo ErrHandle
    
    gcnOracle.BeginTrans
    MousePointer = 11
    '����ɾ�����е�����
    gcnOracle.Execute "delete from zlmenus where ���='" & str��� & "'"
    '������
    SaveMenuItem tvwMenu.Nodes("Root"), "", str���
    
    gcnOracle.CommitTrans
    MousePointer = 0
    SaveMenu = True
    Exit Function
ErrHandle:
    MsgBox "����ʧ�ܡ�", vbExclamation, gstrSysName
    gcnOracle.RollbackTrans
    MousePointer = 0
    SaveMenu = False
End Function

Private Sub SaveMenuItem(nod As Node, ByVal str�ϼ�ID As String, ByVal str��� As String)
    Dim nodTemp As Node
    Dim strFormat As String
    Dim strID As String
    Dim varStr() As String
    
    
    On Error GoTo 0
    Set nodTemp = nod.Child
    Do Until nodTemp Is Nothing
        strID = GetNextId()
        varStr = Split(nodTemp.Tag, "'")
        gcnOracle.Execute "insert into zlmenus (���,ID,�ϼ�ID,����,���,˵��,ϵͳ,ģ��,�̱���,ͼ��) values('" & _
            str��� & "','" & strID & "','" & str�ϼ�ID & "','" & nodTemp.Text & "','" & _
            varStr(1) & "','" & varStr(2) & "','" & varStr(4) & "','" & varStr(0) & "','" & varStr(3) & "'," & Mid(nodTemp.Image, 2) & ")"
        
        '�ݹ����
        Call SaveMenuItem(nodTemp, strID, str���)
        Set nodTemp = nodTemp.Next
    Loop
    
    
End Sub

Private Function GetNextId() As Long
    '-------------------------------------------------------------
    '���ܣ���ȡָ�����ΨһID��
    '������strTable
    '      ���ڵı���
    '���أ���ǰ���ΨһID��
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand
    With rsTemp
        .Open "SELECT zlmenus_id.NextVal FROM DUAL", gcnOracle
        GetNextId = .Fields(0).Value
        .Close
    End With
    Exit Function
    
errHand:
    GetNextId = Null
    Err = 0
End Function

Private Sub SetEnable()
'���ø�����ť��Enable����
    Dim strMenu As String
    Dim blnLvw As Boolean
    Dim blnTvw As Boolean, nod As Node
    Dim i As Integer
    On Error GoTo ErrHandle
    If lvwMenu.SelectedItem Is Nothing Then
        strMenu = ""
    Else
        strMenu = lvwMenu.SelectedItem.Text
    End If
    
    cmdNew.Enabled = strMenu <> ""
    cmdImp.Enabled = cmdNew.Enabled
    cmdExp.Enabled = cmdNew.Enabled
    
    blnLvw = Not (strMenu = "" Or strMenu = "ȱʡ")
    
    If tvwMenu.SelectedItem Is Nothing Then
        blnTvw = False
    Else
        blnTvw = True
        '��ʾ�����˵��
        If tvwMenu.SelectedItem.Image = "Root" Then
            txtName.Text = ""
            txtShort.Text = ""
            txtExplain.Text = ""
            Set icIcon.SelectedItem = Nothing
            
        Else
            Dim varStr() As String
            txtName.Text = tvwMenu.SelectedItem.Text
            varStr = Split(tvwMenu.SelectedItem.Tag, "'")
            txtExplain.Text = varStr(2)
            txtShort.Text = varStr(3)
            Set icIcon.SelectedItem = icIcon.ComboItems(tvwMenu.SelectedItem.Image)
        End If
    End If
        

    cmdDrop.Enabled = blnLvw
    cmdSave.Enabled = blnLvw And mblnModify
    cmdRestore.Enabled = blnLvw And mblnModify
    txtExplain.Enabled = blnLvw
    txtName.Enabled = blnLvw
    txtShort.Enabled = blnLvw
    icIcon.Enabled = blnLvw
    For i = 0 To 5
        cmdEdit(i).Enabled = blnLvw And blnTvw
    Next
    
    If blnTvw = False Then Exit Sub
    Set nod = tvwMenu.SelectedItem
    If blnLvw = True Then
        '����˵���������
        If nod.Image = "Root" Then
            For i = 0 To 5
                cmdEdit(i).Enabled = False
            Next
            cmdEdit(4).Enabled = True
            txtExplain.Enabled = False
            txtName.Enabled = False
            txtShort.Enabled = False
            icIcon.Enabled = False
        Else
            If Split(nod.Tag, "'")(0) <> "" Then txtName.Enabled = False
            
            cmdEdit(con��).Enabled = nod.Parent.Image <> "Root"
            
            If Not nod.Parent.Parent Is Nothing Then
                If nod.Parent.Parent.Image = "Root" And Split(nod.Tag, "'")(0) <> "" Then '��ʾ��ģ���
                    cmdEdit(con��).Enabled = False
                End If
            
            End If
            If nod.Previous Is Nothing Then
                cmdEdit(con��).Enabled = False
            Else
                cmdEdit(con��).Enabled = (Split(nod.Previous.Tag, "'")(0) = "")
            End If
            cmdEdit(con��).Enabled = Not (nod.FirstSibling Is nod)
            cmdEdit(con��).Enabled = Not (nod.LastSibling Is nod)
            cmdEdit(con����).Enabled = Split(nod.Tag, "'")(0) = ""
            cmdEdit(conɾ��).Enabled = Not (nod.Parent.Image = "Root" And nod.Root.Children = 1)
        End If
    End If
    Exit Sub
ErrHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub icIcon_Click()
    If icIcon.SelectedItem Is Nothing Then Exit Sub
        
    tvwMenu.SelectedItem.Image = icIcon.SelectedItem.Key
    tvwMenu.SelectedItem.SelectedImage = icIcon.SelectedItem.Key

    mblnModify = True
    cmdSave.Enabled = True
    cmdRestore.Enabled = True
End Sub

Private Sub lvwMenu_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    lvwMenu.Drag 0
End Sub

Private Sub lvwMenu_KeyUp(KeyCode As Integer, Shift As Integer)
    Call SetEnable
End Sub

Private Sub lvwMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If lvwMenu.HitTest(X, Y) Is Nothing Then Exit Sub
        lvwMenu.Drag 1
    End If
    If lvwMenu.SelectedItem.Text <> mstrMenuName Then
        mstrMenuName = lvwMenu.SelectedItem.Text
        Call SetEnable
    End If
End Sub

Private Sub lvwMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'����ItemClick�¼���ԭ�������Ǹ��¼���������    (�����ס����)
    If lvwMenu.SelectedItem Is Nothing Then Exit Sub
    If lvwMenu.SelectedItem.Text <> mstrMenuName Then
        If mblnModify = True Then
            If MsgBox("�˵��顰" & mstrMenuName & "���Ѿ����޸ģ��ڸ���֮ǰ�Ƿ���Ҫ���棿", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
                SaveMenu mstrMenuName
            End If
        End If
        mstrMenuName = lvwMenu.SelectedItem.Text
        Call FillMenu
    End If
End Sub

Private Sub tvwMenu_Collapse(ByVal Node As MSComctlLib.Node)
    Call tvwMenu_NodeClick(Node)
End Sub

Private Sub tvwMenu_NodeClick(ByVal Node As MSComctlLib.Node)
    Call SetEnable
End Sub

Private Sub txtName_Change()
    If ActiveControl Is txtName Then
        mblnModify = True
        If txtShort.Text = "" Then txtShort.Text = txtName.Text
        cmdSave.Enabled = True
        cmdRestore.Enabled = True
    End If
End Sub

Private Sub txtName_GotFocus()
    SelAll txtName
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim nod As Node
    Dim strName As String
    
    If tvwMenu.SelectedItem Is Nothing Then Exit Sub
    If tvwMenu.SelectedItem.Key = "Root" Then Exit Sub
    
    strName = Trim(txtName.Text)
    If strName = "" Then
        MsgBox "���ⲻ��Ϊ�ա�", vbExclamation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    If StrIsValid(strName, txtName.MaxLength) = False Then
        Cancel = True
        Exit Sub
    End If
    Set nod = tvwMenu.SelectedItem.FirstSibling
    Do Until nod Is Nothing
        If nod.Text = strName Then
            If Not nod Is tvwMenu.SelectedItem Then
                MsgBox "ͬ���˵����Ѿ�����ͬ����Ĳ˵����ˡ�", vbExclamation, gstrSysName
                Cancel = True
                Exit Sub
            End If
        End If
        Set nod = nod.Next
    Loop
    tvwMenu.SelectedItem.Text = strName
End Sub

Private Sub txtShort_Change()
    If ActiveControl Is txtShort Then
        mblnModify = True
        cmdSave.Enabled = True
        cmdRestore.Enabled = True
    End If
End Sub

Private Sub txtShort_GotFocus()
    SelAll txtShort
End Sub

Private Sub txtShort_Validate(Cancel As Boolean)
    Dim strShort As String
    Dim varStr() As String
    
    If tvwMenu.SelectedItem Is Nothing Then Exit Sub
    If tvwMenu.SelectedItem.Key = "Root" Then Exit Sub
    
    strShort = Trim(txtShort.Text)
    If strShort = "" Then
        MsgBox "�̱��ⲻ��Ϊ�ա�", vbExclamation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    If StrIsValid(strShort, txtShort.MaxLength) = False Then
        Cancel = True
        Exit Sub
    End If
    
    varStr = Split(tvwMenu.SelectedItem.Tag, "'")
    tvwMenu.SelectedItem.Tag = varStr(0) & "'" & varStr(1) & "'" & varStr(2) & "'" & strShort & "'" & varStr(4)

End Sub

Private Sub txtExplain_Change()
    If ActiveControl Is txtExplain Then
        mblnModify = True
        cmdSave.Enabled = True
        cmdRestore.Enabled = True
    End If
End Sub

Private Sub txtExplain_GotFocus()
    SelAll txtExplain
End Sub

Private Sub txtExplain_Validate(Cancel As Boolean)
    Dim strExplain As String
    Dim varStr() As String
    
    strExplain = Trim(txtExplain.Text)
    If StrIsValid(strExplain, txtExplain.MaxLength) = False Then
        Cancel = True
        Exit Sub
    End If
    
    varStr = Split(tvwMenu.SelectedItem.Tag, "'")
    tvwMenu.SelectedItem.Tag = varStr(0) & "'" & varStr(1) & "'" & strExplain & "'" & varStr(3) & "'" & varStr(4)

End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

End Sub

