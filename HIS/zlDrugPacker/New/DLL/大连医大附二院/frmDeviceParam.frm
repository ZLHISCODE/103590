VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDeviceParam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�豸Ӧ�ò�����Ϣ"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5970
   Icon            =   "frmDeviceParam.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   5970
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboDevice 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   780
      Width           =   4935
   End
   Begin VB.Frame fraLine1 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   5900
   End
   Begin TabDlg.SSTab sstDeviceParam 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ӧ�ó���(&0)"
      TabPicture(0)   =   "frmDeviceParam.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDevice(10)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LvwҩƷ����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ListView LvwҩƷ���� 
         Height          =   4065
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   7170
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgLvwSel"
         SmallIcons      =   "imgLvwSel"
         ColHdrIcons     =   "imgLvwSel"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   4680
      TabIndex        =   1
      Top             =   6600
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   3480
      TabIndex        =   0
      Top             =   6600
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�豸"
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDeviceParam.frx":0326
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "���÷�ҩ�豸��ʹ�û�����Ӧ�ò�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmDeviceParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�豸ID As Long
Private mlngҩ��ID As Long

Public Sub ShowMeByDevice(ByVal frmOwner As Form, ByVal lng�豸id As Long)
    mlng�豸ID = lng�豸id
    
    Call GetDevice(0, lng�豸id)
    
    Me.Show vbModal, frmOwner
    
    Exit Sub
End Sub

Public Sub ShowMeByStock(ByVal frmOwner As Form, ByVal lngҩ��ID As Long)
    mlngҩ��ID = lngҩ��ID
    
    Call GetDevice(1, mlngҩ��ID)
    
    Me.Show vbModal, frmOwner
    
    Exit Sub
End Sub

Private Sub IniData(ByVal lngҩ��ID As Long)
    Dim rsData As ADODB.Recordset
    Dim byt���� As Byte     '1-����,2-סԺ,3-�����סԺ
    
    On Error GoTo errHandle
    
    '�������
    gstrSQL = "Select ������� From ��������˵�� " & _
                  "Where ����id = [1] And ������� in (1,2,3) " & _
                  "Order By ������� "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "IniData", lngҩ��ID)
    Do While rsData.EOF = False
        If byt���� = 0 Then
            byt���� = NVL(rsData!�������, 0)
        ElseIf byt���� = 3 Then
            Exit Do
        Else
            Select Case NVL(rsData!�������, 0)
                Case 1  '����
                    If byt���� = 2 Then
                        byt���� = 3
                        Exit Do
                    End If
                Case 2  'סԺ
                    If byt���� = 1 Then
                        byt���� = 3
                        Exit Do
                    End If
                Case 3  '�����סԺ
                    byt���� = 3
                    Exit Do
            End Select
        End If
        
        rsData.MoveNext
    Loop
    
'    If byt���� = 1 Then
'        optObject(0).Value = True
'        optObject(0).Enabled = True
'        optObject(1).Value = False
'        optObject(1).Enabled = False
'     ElseIf byt���� = 2 Then
'        optObject(0).Value = False
'        optObject(0).Enabled = False
'        optObject(1).Value = True
'        optObject(1).Enabled = True
'     ElseIf byt���� = 3 Then
'        optObject(0).Value = True
'        optObject(0).Enabled = True
'        optObject(1).Value = True
'        optObject(1).Enabled = True
'    End If
        
'    'ҵ�����
'    With cboDispense
'        .Clear
'        .AddItem "�����շ�", 0
'        .AddItem "������ҩ��ҩ����", 1
'        .AddItem "������ҩ��ҩ����", 2
'    End With
'
'    With cboSend
'        .Clear
'        .AddItem "����Ӧ", 0
'        .AddItem "ҩƷ������ҩ����", 1
'    End With
    
'    '�ϴ�����
'    chkBillType(0).Value = 1
'    chkBillType(1).Value = 1
'    chkBillType(2).Value = 1
    
    gstrSQL = "Select Distinct J.����||'-'||J.���� ����" & _
         " From ����ִ�п��� A,ҩƷ���� B,ҩƷ���� J " & _
         " Where A.������ĿID=B.ҩ��ID And B.ҩƷ����=J.����" & _
         " And A.ִ�п���ID=[1]" & _
         " Order By j.���� || '-' || j.���� "
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "IniData", lngҩ��ID)
    
    With LvwҩƷ����
        .ListItems.Clear
        .ListItems.Add , "_" & .ListItems.Count + 1, "����ҩƷ����" ', 1, 1
        .ListItems(.ListItems.Count).Checked = True
        Do While Not rsData.EOF
            .ListItems.Add , "_" & .ListItems.Count + 1, Mid(rsData!����, InStr(1, rsData!����, "-") + 1) ', 1, 1
            .ListItems(.ListItems.Count).Checked = True
            rsData.MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub GetDeviceParam(ByVal lng�豸id As Long)
    Dim rsData As ADODB.Recordset
    Dim n As Integer
    
    On Error GoTo errHandle
    
    gstrSQL = "Select a.����id, a.�豸id, Nvl(a.����ֵ, b.ȱʡֵ) As ����ֵ, b.������, b.������, b.����˵�� " & vbNewLine & _
        " From ҩ���豸���� A, �Զ���ҩ���� B" & vbNewLine & _
        " Where a.����id(+) = b.Id and a.�豸id(+)=[1] "

    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDeviceParam", lng�豸id)
    
    Do While Not rsData.EOF
        Select Case rsData!������
'            Case "�������"
'                If rsData!����ֵ = 1 Then
'                    optObject(0).Value = True
'                Else
'                    optObject(1).Value = True
'                End If
'            Case "Ԥ��ҩ��Ӧ"
'                cboDispense.ListIndex = Val(rsData!����ֵ) - 1
'            Case "������Ӧ"
'                cboSend.ListIndex = Val(rsData!����ֵ)
'            Case "��������"
'                chkBillType(0).Value = Val(Mid(rsData!����ֵ, 1, 1))
'                chkBillType(1).Value = Val(Mid(rsData!����ֵ, 2, 1))
'                chkBillType(2).Value = Val(Mid(rsData!����ֵ, 3, 1))
            Case "ҩƷ����"
                With LvwҩƷ����
                    If .ListItems.Count = 0 Then
                        Exit Sub
                    End If
                    
                    For n = 1 To .ListItems.Count
                        .ListItems(n).Checked = False
                        If NVL(rsData!����ֵ) = "����" Then
                            .ListItems(n).Checked = True
                        Else
                            If InStr(1, "," & NVL(rsData!����ֵ) & ",", "," & .ListItems(n).Text & ",") > 0 Then
                                .ListItems(n).Checked = True
                            End If
                        End If
                    Next
                End With
        End Select
                
        rsData.MoveNext
    Loop
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub


Private Sub cboDevice_Click()
'    If mlng�豸ID <> Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(0)) And mlngҩ��ID <> Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(1)) Then
'        mlng�豸ID = Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(0))
'        mlngҩ��ID = Val(Split(cboDevice.ItemData(cboDevice.ListIndex), "|")(1))
'
'        Call IniData(mlngҩ��ID)
'        DoEvents
'        Call GetDeviceParam(mlng�豸ID)
'    End If

    If mlng�豸ID <> cboDevice.ItemData(cboDevice.ListIndex) Then
        mlng�豸ID = cboDevice.ItemData(cboDevice.ListIndex)
        Call IniData(mlngҩ��ID)
        DoEvents
        Call GetDeviceParam(mlng�豸ID)
    End If
    
End Sub


Private Sub cmdSave_Click()
    Dim str���� As String
    Dim n As Integer
    
    On Error GoTo errHandle
    
    gobjConn.BeginTrans
    
    '�������ŷֱ𱣴����
'    '�������
'    gstrSQL = "Zl_ҩ���豸����_Update("
'    gstrSQL = gstrSQL & 1 & ","
'    gstrSQL = gstrSQL & mlng�豸ID & ","
'    gstrSQL = gstrSQL & IIf(optObject(0).Value = True, 1, 2)
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'    '��ҩ����
'    gstrSQL = "Zl_ҩ���豸����_Update("
'    gstrSQL = gstrSQL & 2 & ","
'    gstrSQL = gstrSQL & mlng�豸ID & ","
'    gstrSQL = gstrSQL & cboDispense.ListIndex + 1
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'    '���͹���
'    gstrSQL = "Zl_ҩ���豸����_Update("
'    gstrSQL = gstrSQL & 3 & ","
'    gstrSQL = gstrSQL & mlng�豸ID & ","
'    gstrSQL = gstrSQL & cboSend.ListIndex
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
'
'
'    '��������
'    gstrSQL = "Zl_ҩ���豸����_Update("
'    gstrSQL = gstrSQL & 4 & ","
'    gstrSQL = gstrSQL & mlng�豸ID & ","
'    gstrSQL = gstrSQL & chkBillType(0).Value & chkBillType(1).Value & chkBillType(2).Value
'    gstrSQL = gstrSQL & ")"
'    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "CmdSave_Click")
     
    '����
    If LvwҩƷ����.ListItems(1).Checked Then
        str���� = "����"
    Else
        For n = 1 To LvwҩƷ����.ListItems.Count
            If LvwҩƷ����.ListItems(n).Checked Then
                str���� = IIf(str���� = "", "", str���� & ",") & LvwҩƷ����.ListItems(n).Text
            End If
        Next
    End If
    gstrSQL = "Zl_ҩ���豸����_Update("
    gstrSQL = gstrSQL & 1 & ","
    gstrSQL = gstrSQL & mlng�豸ID & ","
    gstrSQL = gstrSQL & IIf(str���� = "����", "null", "'" & str���� & "'")
    gstrSQL = gstrSQL & ")"
    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "cmdSave_Click")
    
    gobjConn.CommitTrans
    
    MsgBox "�����ѱ��棡", vbInformation, GSTR_INTERFACE_NAME
    
    Exit Sub
errHandle:
    gobjConn.RollbackTrans
    gobjComLib.ErrCenter
    gstrMessage = Err.Description
End Sub

Private Sub LvwҩƷ����_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim n As Integer
    Dim blnAllChecked As Boolean
    
    With LvwҩƷ����
        For n = 1 To .ListItems.Count
            .ListItems(n).Selected = False
        Next
        Item.Selected = True
        If Item.Text = "����ҩƷ����" Then
            If Item.Checked Then
                blnAllChecked = True
            End If
                
            For n = 1 To .ListItems.Count
                .ListItems(n).Checked = blnAllChecked
            Next
        Else
            If Item.Checked = False Then
                .ListItems(1).Checked = False
            End If
        End If
    End With
End Sub

Private Sub GetDevice(ByVal bytType As Byte, ByVal lng��ʶid As Long)
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errHandle
     
    If bytType = 0 Then
        '���豸IDȡ�豸��Ϣ
        gstrSQL = "Select a.Id As �豸id, a.ʹ�ò���id As ҩ��id, '��' || a.���� || '��' || a.���� || '(' || a.�ͺ� || ')' || ' - ' || b.���� As ���� " & _
            " From ҩ����ҩ�豸 A, ���ű� B " & _
            " Where a.ʹ�ò���id = b.Id And a.ID = [1] "
    Else
        '��ҩ��IDȡ�豸��Ϣ
        gstrSQL = "Select a.Id As �豸id, a.ʹ�ò���id As ҩ��id, '��' || a.���� || '��' || a.���� || '(' || a.�ͺ� || ')' || ' - ' || b.���� As ���� " & _
            " From ҩ����ҩ�豸 A, ���ű� B " & _
            " Where a.ʹ�ò���id = b.Id And a.ʹ�ò���id = [1] "
    End If
    Set rsData = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "GetDevice", lng��ʶid)
    
    cboDevice.Clear
    Do While rsData.EOF = False
        cboDevice.AddItem rsData!����
        cboDevice.ItemData(cboDevice.NewIndex) = rsData!�豸ID '"" & rsData!�豸ID & "|" & rsData!ҩ��id
        rsData.MoveNext
    Loop
    
    If bytType = 1 And cboDevice.ListCount > 0 Then
        cboDevice.ListIndex = 0
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
