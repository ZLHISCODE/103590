VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMgrUserGrant 
   Caption         =   "��������Ȩ"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   Icon            =   "frmMgrUserGrant.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   8790
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   3930
      Picture         =   "frmMgrUserGrant.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   4500
      Picture         =   "frmMgrUserGrant.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwGranted 
      Height          =   3800
      Left            =   5040
      TabIndex        =   9
      Top             =   960
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   6694
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwNoGrant 
      Height          =   3800
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   6694
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "�����û�(&F)"
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   65
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4800
      TabIndex        =   1
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6150
      TabIndex        =   3
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7425
      TabIndex        =   4
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -360
      TabIndex        =   0
      Top             =   525
      Width           =   10110
   End
   Begin MSComctlLib.ImageList Img16 
      Left            =   3975
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1A5E
            Key             =   "�Զ�����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":82C0
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EB22
            Key             =   "������־����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":15384
            Key             =   "������ʱ����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1BBE6
            Key             =   "ϵͳװж����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":22448
            Key             =   "����ת��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":28CAA
            Key             =   "�û�ע�����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":2F50C
            Key             =   "ϵͳ��Ǩ����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":35D6E
            Key             =   "ϵͳ��������"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":3C5D0
            Key             =   "������־����"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":42E32
            Key             =   "������־����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":49694
            Key             =   "ϵͳ����ѡ��"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":4FEF6
            Key             =   "�������޸�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":56758
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":5CFBA
            Key             =   "վ���ļ��ռ�"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6381C
            Key             =   "������Ч����"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6A07E
            Key             =   "��̨��ҵ����"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":708E0
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":77142
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":7D9A4
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":84206
            Key             =   "���ݵ���"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":8AA68
            Key             =   "����״̬���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":912CA
            Key             =   "�û���װ�ű�"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":97B2C
            Key             =   "վ�㲿������"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":9E38E
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":A4BF0
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":AB452
            Key             =   "�û���Ȩ����"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B1CB4
            Key             =   "��ɫ��Ȩ����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B8516
            Key             =   "�˵�����滮"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":BED78
            Key             =   "�ͻ������п���"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C55DA
            Key             =   "Ȩ�޹���"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C5EB4
            Key             =   "װж����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C678E
            Key             =   "���ݹ���"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7068
            Key             =   "���й���"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7602
            Key             =   "ר���"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7EDC
            Key             =   "DBA����"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":CE73E
            Key             =   "�ռ����"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":D4FA0
            Key             =   "SQL����"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":DB802
            Key             =   "�Ự����"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":E2064
            Key             =   "�������"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":E88C6
            Key             =   "SQL����"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EF128
            Key             =   "���ݿ�����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFunc 
      Height          =   1710
      Left            =   225
      TabIndex        =   12
      Top             =   4875
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   3016
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "˵��"
         Object.Width           =   12347
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ȱʡ"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�ԡ����Ʊ򡱽�����Ȩ����"
      Height          =   180
      Left            =   960
      TabIndex        =   7
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgOne 
      Height          =   480
      Left            =   300
      Picture         =   "frmMgrUserGrant.frx":F598A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblModul 
      AutoSize        =   -1  'True
      Caption         =   "����Ȩģ��(&A)"
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   660
      Width           =   1170
   End
   Begin VB.Label lblGranted 
      AutoSize        =   -1  'True
      Caption         =   "����Ȩģ��(&G)"
      Height          =   180
      Left            =   4935
      TabIndex        =   5
      Top             =   660
      Width           =   1170
   End
End
Attribute VB_Name = "frmMgrUserGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrProg As String
Private mstrAccount As String 'Ϊ�ձ�ʾ���û���Ȩ
Private mblnOk As Boolean
Private mrsProgFuncs As ADODB.Recordset
Private mblnIsChange As Boolean '��¼�����Ƿ������޸�

Private Enum LvwFuncList
    LFL_���� = 0
    LFL_˵�� = 1
    LFL_ȱʡ = 2
End Enum

Public Function GrantToProg(ByVal strAccount As String, ByVal strUser As String, ByVal strProg As String) As Boolean
    mstrUser = strUser
    mstrAccount = strAccount
    mstrProg = strProg
    mblnOk = False
    Me.Show 1
    GrantToProg = mblnOk
End Function

Private Sub cmdCancel_Click()
    If mblnIsChange Then
        If MsgBox("����Ա����Ȩ����Ϣ�ѱ����ģ�ȷ��Ҫ�������Ĳ��˳���", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbOK Then
    Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click()
    Call FindPersonnel
End Sub

Private Sub MoveProg(objMoveIn As TreeView, objMoveOut As TreeView)
    Dim i As Long, y As Long
    Dim strDel As String, Node As Node
    
    For i = objMoveOut.Nodes.Count To 1 Step -1
        err = 0
        On Error Resume Next
        If objMoveOut.Nodes(i).Checked And Not objMoveOut.Nodes(i).Parent Is Nothing Then
            mblnIsChange = True
            If err = 0 Then
                err = 0
                If objMoveIn.Nodes(objMoveOut.Nodes(i).Parent.Key).Key <> "" Then
                    If err <> 0 Then
                        '��������
                        Set Node = objMoveIn.Nodes.Add(, , objMoveOut.Nodes(i).Parent.Key, objMoveOut.Nodes(i).Parent.Text, objMoveOut.Nodes(i).Parent.Image, objMoveOut.Nodes(i).Parent.SelectedImage)
                        Node.Expanded = objMoveOut.Nodes(i).Parent.Expanded
                        Node.Checked = objMoveOut.Nodes(i).Parent.Checked
                        Node.ForeColor = objMoveOut.Nodes(i).Parent.ForeColor
                    End If
                     '��������
                    Set Node = objMoveIn.Nodes.Add(objMoveOut.Nodes(i).Parent.Key, tvwChild, objMoveOut.Nodes(i).Key, objMoveOut.Nodes(i).Text, objMoveOut.Nodes(i).Image, objMoveOut.Nodes(i).SelectedImage)
                    Node.Expanded = objMoveOut.Nodes(i).Expanded
                    Node.Checked = objMoveOut.Nodes(i).Checked
                    Node.ForeColor = objMoveOut.Nodes(i).ForeColor
                    'ɾ������
                    If objMoveOut.Nodes(i).Parent.Children = 1 Then
                        objMoveOut.Nodes.Remove objMoveOut.Nodes(i).Parent.Index
                    Else
                        objMoveOut.Nodes.Remove i
                    End If
                    
                End If
                On Error GoTo 0
            End If
        End If
    Next
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        Call MoveProg(tvwGranted, tvwNoGrant)
    ElseIf Index = 1 Then
        Call MoveProg(tvwNoGrant, tvwGranted)
    End If
End Sub

Private Sub cmdOK_Click()
'���ܣ���Ȩ
    Dim i As Integer, j As Integer
    Dim strProg As String, strFunc As String, strKey As String
    Dim StrJiami() As Byte
    Dim strPwText As String
    Dim rsTemp As New ADODB.Recordset
    
    If mstrAccount = "" Then
        MsgBox "���Ȳ�����Ҫ��Ȩ���û���", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    
    '��װ�����ַ���
    For i = 1 To tvwGranted.Nodes.Count
        If Not tvwGranted.Nodes(i).Parent Is Nothing Then
            strKey = Mid(tvwGranted.Nodes(i).Key, 2)
            mrsProgFuncs.Filter = "��� = '" & strKey & "' And Ȩ�� = 1"
            strFunc = ""
            Do While Not mrsProgFuncs.EOF
                strFunc = strFunc & "|" & mrsProgFuncs!����
                mrsProgFuncs.MoveNext
            Loop
            If strFunc <> "" Then
                strFunc = ":" & "����" & "|" & Mid(strFunc, 2)
            Else
                strFunc = ":" & "����"
            End If
            strProg = strProg & "," & strKey & strFunc
        End If
    Next
    strProg = Mid(strProg, 2)
    '���ܼ���
    If strProg <> "" Then
        Call DES_Encode(StrConv(strProg, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("��λ����", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    On Error GoTo errHandle
    gstrSQL = "Select 1 From zlMgrGrant Where �û���='" & mstrAccount & "'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount > 0 Then
        If strPwText = "" Then
            gstrSQL = "Delete zlMgrGrant Where �û���='" & mstrAccount & "'"
        Else
            gstrSQL = "Update zlMgrGrant Set ����='" & strPwText & "' Where �û���='" & mstrAccount & "'"
        End If
    Else
        gstrSQL = "Insert into zlMgrGrant(�û���,����) values('" & mstrAccount & "','" & strPwText & "')"
    End If
    gcnOracle.Execute gstrSQL
    '���¹���Ա�˻���Ϣ
    rsTemp.Close
    'δ��Ȩ���򲻸��¹���Ա��Ϣ
    If Not gstrPassword Like "δ��Ȩ�ĳ���:*" Then
        gstrSQL = "Select 1 From zlRegInfo where ��Ŀ='����Ա'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
         If rsTemp.RecordCount > 0 Then
            gstrSQL = "Update zlRegInfo Set ����='" & gstrUserName & "' Where ��Ŀ='����Ա'"
        Else
            gstrSQL = "Insert into zlRegInfo(��Ŀ,����) values('����Ա','" & gstrUserName & "')"
        End If
        gcnOracle.Execute gstrSQL
        '��֤��
        strPwText = ""
        ReDim Preserve StrJiami(0)
        If gstrPassword <> "" Then
            Call DES_Encode(StrConv(gstrPassword, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("��λ����", False, 0))
            strPwText = FuncByteTo16Code(StrJiami)
        End If
        rsTemp.Close
        gstrSQL = "Select 1 From zlRegInfo where ��Ŀ='��֤��'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
         If rsTemp.RecordCount > 0 Then
            gstrSQL = "Update zlRegInfo Set ����='" & strPwText & "' Where ��Ŀ='��֤��'"
        Else
            gstrSQL = "Insert into zlRegInfo(��Ŀ,����) values('��֤��','" & strPwText & "')"
        End If
        gcnOracle.Execute gstrSQL
    End If
    mblnOk = True
    Unload Me
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Load()
    If mstrAccount = "" Then
        lblNote.Caption = "���������û�������Ա��������롣"
        txtFind.Visible = True
        cmdFind.Visible = True
    Else
        lblNote.Caption = "���ڶ�""" & mstrUser & """���й�������Ȩ��"
        txtFind.Visible = False
        cmdFind.Visible = False
    End If
    Call InitProgFuncData
    Call FillProg
End Sub

'��ʼ��ģ�鹦����Ϣ��һ����¼����
Private Sub InitProgFuncData()
    Dim rsTemp As ADODB.Recordset
    Dim strProg As String
    Dim arrProg() As String
    Dim arrFunc() As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errh
    '��ѯ�����п�����Ȩ��ģ�鼰����
    gstrSQL = "Select a.���, a.����, a.�ϼ�, b.����, b.ȱʡ, b.����, 0 Ȩ��, b.˵��" & vbNewLine & _
            "From Zlsvrtools a, Zlsvrfuncs b" & vbNewLine & _
            "Where a.��� = b.���(+) And a.��� <> '0404'" & vbNewLine & _
            "Order By a.���, b.����"
    Set mrsProgFuncs = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption))
    '����Ȩ���ֶ�
    '����mstrProg���Ӷ���ȡѡ����Ա��ӵ�е�ģ�鼰���ܵ���Ȩ���
    arrProg = Split(mstrProg, ",")
    For i = 0 To UBound(arrProg)
        strProg = Split(arrProg(i), ":")(0)
        Call RecUpdate(mrsProgFuncs, "��� = '" & strProg & "' And ���� = '����'", "Ȩ��", 1)
        arrFunc = Split(Split(arrProg(i) & ":", ":")(1), "|")
        If UBound(arrFunc) = -1 Then
            '����ԭ�����û���ֻ��ģ���������Ȩ����δ�Թ��ܽ�����Ȩ���ʹ����ַ����϶�Ϊ�գ���ʱĬ�Ϲ�ѡ���й���
            Call RecUpdate(mrsProgFuncs, "��� = '" & strProg & "'", "Ȩ��", 1)
        Else
        For j = 0 To UBound(arrFunc)
            Call RecUpdate(mrsProgFuncs, "��� = '" & strProg & "' And ���� = '" & arrFunc(j) & "'", "Ȩ��", 1)
        Next
        End If
    Next
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub FillProg()
'���ܣ���书��
    Dim strProg As String, Node As Node
    Dim i As Long
    
    On Error GoTo errHandle
    '��ʾ���û����еĽ�ɫ
    mrsProgFuncs.Filter = "���� = '����' Or ���� = Null"

    Do Until mrsProgFuncs.EOF
        With IIf(mrsProgFuncs!Ȩ�� = 0, tvwNoGrant, tvwGranted)
            '�ϼ����߶���
            If IsNull(mrsProgFuncs("�ϼ�")) Then
                Set Node = tvwNoGrant.Nodes.Add(, , "D" & mrsProgFuncs("���"), "��" & mrsProgFuncs("���") & "��" & mrsProgFuncs("����"))
                tvwNoGrant.Nodes("D" & mrsProgFuncs("���")).Sorted = True
                tvwNoGrant.Nodes("D" & mrsProgFuncs("���")).Expanded = True
                tvwNoGrant.Nodes("D" & mrsProgFuncs("���")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!���� & "").Index
                err.Clear: On Error GoTo errHandle
                Set Node = tvwGranted.Nodes.Add(, , "D" & mrsProgFuncs("���"), "��" & mrsProgFuncs("���") & "��" & mrsProgFuncs("����"))
                tvwGranted.Nodes("D" & mrsProgFuncs("���")).Sorted = True
                tvwGranted.Nodes("D" & mrsProgFuncs("���")).Expanded = True
                tvwGranted.Nodes("D" & mrsProgFuncs("���")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!���� & "").Index
                err.Clear: On Error GoTo errHandle
            Else
                Set Node = .Nodes.Add("D" & mrsProgFuncs("�ϼ�"), tvwChild, "C" & mrsProgFuncs("���"), mrsProgFuncs("����"))
                .Nodes("C" & mrsProgFuncs("���")).Sorted = True
                Node.Checked = False
                mrsProgFuncs.Update "Ȩ��", 0
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(mrsProgFuncs!���� & "").Index
                err.Clear: On Error GoTo errHandle
            End If
        End With
        mrsProgFuncs.MoveNext
    Loop
    'ɾ��û������ķ���
    For i = tvwNoGrant.Nodes.Count To 1 Step -1
        If tvwNoGrant.Nodes(i).Children = 0 And tvwNoGrant.Nodes(i).Parent Is Nothing Then
            tvwNoGrant.Nodes.Remove i
        End If
    Next
    For i = tvwGranted.Nodes.Count To 1 Step -1
        If tvwGranted.Nodes(i).Children = 0 And tvwGranted.Nodes(i).Parent Is Nothing Then
            tvwGranted.Nodes.Remove i
        End If
    Next
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

'���ģ���Ӧ�Ĺ��ܣ��������ܲ�����ʾ��ΪĬ������
Private Sub FillFunction(ByVal strPorgNo As String)
    Dim lst As ListItem
    
    On Error GoTo errh
    lvwFunc.ListItems.Clear
    mrsProgFuncs.Filter = "��� = '" & strPorgNo & "' And ���� <> '����'"
    '��书����Ϣ����ѡ���
    With mrsProgFuncs
        Do While Not .EOF
            Set lst = lvwFunc.ListItems.Add(, !��� & "_" & !����, !����)
            lst.SubItems(LFL_˵��) = !˵�� & ""
            lst.SubItems(LFL_ȱʡ) = !ȱʡ & ""
            lst.Checked = IIf(!Ȩ�� = 1, True, False)
            .MoveNext
        Loop
    End With
    Exit Sub
errh:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        If mblnIsChange Then
            If MsgBox("����Ա����Ȩ����Ϣ�ѱ����ģ�ȷ��Ҫ�������Ĳ��˳���", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                Cancel = 1
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraLine.Width = Me.Width
    cmdOK.Move Me.Width - cmdCancel.Width - cmdOK.Width - 400, Me.Height - cmdOK.Height - 650
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 100, cmdOK.Top
    lvwFunc.Width = Me.Width - 600
    lvwFunc.Top = cmdOK.Top - lvwFunc.Height - 100
    If lvwFunc.ListItems.Count > 5 Then
        lvwFunc.ColumnHeaders(2).Width = lvwFunc.Width - lvwFunc.ColumnHeaders(1).Width - 250
    Else
        lvwFunc.ColumnHeaders(2).Width = lvwFunc.Width - lvwFunc.ColumnHeaders(1).Width
    End If
    tvwNoGrant.Width = Me.Width \ 2 - 885
    tvwGranted.Width = Me.Width \ 2 - 885
    tvwNoGrant.Height = lvwFunc.Top - tvwNoGrant.Top - 100
    tvwGranted.Height = tvwNoGrant.Height
    tvwGranted.Left = tvwNoGrant.Left + tvwNoGrant.Width + 1185
    
    cmdMove(1).Left = tvwNoGrant.Left + tvwNoGrant.Width + 150
    cmdMove(0).Left = cmdMove(1).Left + cmdMove(1).Width + 150
    lblGranted.Left = tvwGranted.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsProgFuncs = Nothing
            mblnIsChange = False
End Sub

Private Sub lvwFunc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    '���޸���һ�����¼�¼��
    mblnIsChange = True
    Call RecUpdate(mrsProgFuncs, "��� = '" & Split(Item.Key, "_")(0) & "' And ���� = '" & Item.Text & "'", "Ȩ��", IIf(Item.Checked, 1, 0))
End Sub

Private Sub tvwGranted_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Call tvwGranted_NodeClick(Node)
    Call NodeCheckMode(Node, tvwGranted)
End Sub

Private Sub tvwGranted_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillFunction(Mid(Node.Key, 2))
    If Node = tvwGranted.SelectedItem Then
        lvwFunc.Enabled = True
        lvwFunc.BackColor = &H80000005
    End If
End Sub

Private Sub tvwNoGrant_NodeCheck(ByVal Node As MSComctlLib.Node)
    Node.Selected = True
    Call tvwNoGrant_NodeClick(Node)
    Call NodeCheckMode(Node, tvwNoGrant)
    If Node = tvwNoGrant.SelectedItem Then
        If Node.Checked = False Then
           lvwFunc.Enabled = False
           lvwFunc.BackColor = &H8000000F
        Else
           lvwFunc.Enabled = True
           lvwFunc.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub NodeCheckMode(ByRef Node As MSComctlLib.Node, ByRef objtvwThis As TreeView)
'���ܣ�������ѡ�и��ڵ㣬�Զ�ѡ�������ӽڵ㣬ѡ�������ӽڵ㣬���ڵ�Ҳѡ��
    Dim i As Long
    Dim blnIsNothing As Boolean
    
    LockWindowUpdate objtvwThis.hwnd
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Key Then
                    objtvwThis.Nodes(i).Checked = Node.Checked
                End If
            End If
        Next
    Else
        For i = Node.Parent.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Parent.Key Then
                    If Not objtvwThis.Nodes(i).Checked = Node.Checked Then blnIsNothing = True
                End If
            End If
        Next
        '����ѡ�ı����ǵ�ǰѡ������lvwFunc�е�Ϊȱʡ����Ҳ��ѡ��
        If Node = objtvwThis.SelectedItem Then
            For i = 1 To lvwFunc.ListItems.Count
                If Node.Checked = False Then
                    lvwFunc.ListItems.Item(i).Checked = False
                    Call lvwFunc_ItemCheck(lvwFunc.ListItems.Item(i))
                ElseIf lvwFunc.ListItems.Item(i).SubItems(LFL_ȱʡ) = "1" Then
                    lvwFunc.ListItems.Item(i).Checked = True
                    Call lvwFunc_ItemCheck(lvwFunc.ListItems.Item(i))
                End If
            Next
        End If
        If blnIsNothing Then
            Node.Parent.Checked = False
        Else
            Node.Parent.Checked = Node.Checked
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub tvwNoGrant_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Parent Is Nothing Then Exit Sub
    Call FillFunction(Mid(Node.Key, 2))
    If Node.Checked = False Then
        lvwFunc.Enabled = False
        lvwFunc.BackColor = &H8000000F
    Else
        lvwFunc.Enabled = True
        lvwFunc.BackColor = &H80000005
    End If
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindPersonnel
    End If
End Sub

Private Sub FindPersonnel()
'���ܣ�������Ա
    Dim rsTemp As New Recordset
    Dim objPoint As POINTAPI
    
    If txtFind.Text = "" Then Exit Sub
    gstrSQL = "Select b.�û���, c.����, c.����, d.���� As ��������" & vbNewLine & _
            "From  Zlmgrgrant A,�ϻ���Ա�� B, ��Ա�� C, ���ű� D, ������Ա E" & vbNewLine & _
            "Where a.�û���(+) = b.�û��� And b.��Աid = c.Id And c.Id = e.��Աid And d.Id = e.����id And A.�û��� is null And e.ȱʡ = 1 And B.�û��� <> '" & gstrUserName & "'" & _
            " And(b.�û��� like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���� Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���� Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.���=' & UCase(Trim(txtFind.Text)) & ')" & _
            " Order By c.����"
    Set rsTemp = New ADODB.Recordset
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ҵ��û������ڣ������Ѿ�ӵ����Ȩ�ޣ����顣", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        Exit Sub
    End If
    Call ClientToScreen(txtFind.hwnd, objPoint)
    
    If frmSelectList.ShowSelect(Nothing, rsTemp, "�û���,900,0,1;����,900,0,1;����,650,0,0;��������,1500,0,1", objPoint.x * 15 - 30, objPoint.y * 15 + cmdFind.Height - 30, txtFind.Width + cmdFind.Width + 1300, 3000, "", "������Ա", , , True) = False Then
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        rsTemp.Filter = 0
        Exit Sub
    Else
        txtFind.Text = rsTemp!���� & ""
        mstrAccount = rsTemp!�û��� & ""
        mstrUser = rsTemp!���� & ""
    End If
End Sub
