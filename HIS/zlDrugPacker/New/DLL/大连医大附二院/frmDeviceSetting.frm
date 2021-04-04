VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeviceSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ϴ���������"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   Icon            =   "frmDeviceSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7215
   StartUpPosition =   1  '����������
   Begin VB.Frame fraOther 
      Caption         =   "����"
      Height          =   3855
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox chkBillType 
         Caption         =   "���ʵ�"
         Height          =   180
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkBillType 
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.CheckBox chkBillType 
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin MSComctlLib.TreeView tvwDrugType 
         Height          =   2580
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4551
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ͣ�"
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblDevice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ͣ�"
         Height          =   180
         Index           =   12
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   360
      Left            =   4680
      TabIndex        =   1
      Top             =   4080
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5880
      TabIndex        =   2
      Top             =   4080
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwDevices 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgEnabled 
      Left            =   360
      Top             =   3840
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
            Picture         =   "frmDeviceSetting.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeviceSetting.frx":0794
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDeviceSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_ROOT = "ROOT"
Private Const STR_ALL = "ȫ��"

Public Sub ShowMe(ByVal lngDeptID As Long)
    '���ע���豸
    Call Init
    Call FullData(1, lngDeptID)
    
    If lvwDevices.ListItems.Count = 0 Then
        MsgBox "��δע��ҩ���Զ����豸��", vbInformation, GSTR_INTERFACE_NAME
        Unload Me
        Exit Sub
    End If
    
    Call FullData(2)
    
    Call lvwDevices_Click
    
    Show vbModal, gfrmOwner
    
End Sub

Private Sub chkBillType_Click(Index As Integer)
    Dim i As Integer
    
    tvwDrugType.Enabled = False
    If chkBillType(Index).Value = 1 Then
        tvwDrugType.Enabled = True
    Else
        For i = 0 To chkBillType.Count - 1
            If chkBillType(i).Value = 1 Then
                tvwDrugType.Enabled = True
                Exit For
            End If
        Next
    End If
    
    If tvwDrugType.Enabled = False Then
        '�����ѡ����
        tvwDrugType.Nodes(STR_ROOT).Checked = False
        tvwDrugType_NodeCheck tvwDrugType.Nodes(STR_ROOT)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim lngDeviceID As Long
    Dim strBillType As String, strDrugType As String
    
    '���浱ǰ���ý�lvwDevices
    With lvwDevices.SelectedItem
        .SubItems(6) = GetBillTypeStr()
        .SubItems(7) = GetDrugTypeStr()
    End With
    
    '����
    On Error GoTo errHandle
    gobjConn.BeginTrans
    For i = 1 To lvwDevices.ListItems.Count
        With lvwDevices.ListItems(i)
            lngDeviceID = Val(Mid(.Key, 3))
            strBillType = Trim(.SubItems(6))
            strDrugType = Trim(.SubItems(7))
            
            If lngDeviceID > 0 Then
                gstrSQL = "zl_ҩ��ע���豸_Setting("
                gstrSQL = gstrSQL & lngDeviceID & ","
                gstrSQL = gstrSQL & IIf(strBillType = "", "null", "'" & strBillType & "'") & ","
                gstrSQL = gstrSQL & IIf(strDrugType = "", "null", "'" & strDrugType & "'")
                gstrSQL = gstrSQL & ")"
                Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "����ҩ���豸�ϴ�����")
            End If
            
        End With
    Next
    gobjConn.CommitTrans
    
    Unload Me
    Exit Sub
    
errHandle:
    gobjConn.RollbackTrans
    gobjComLib.ErrCenter
    gstrMessage = Err.Description
End Sub

Private Sub Form_Load()
    '
End Sub

Private Sub Init()
    With Me.lvwDevices
        .ColumnHeaders.Add , , "����", 1000
        .ColumnHeaders.Add , , "����", 1500
        .ColumnHeaders.Add , , "�ͺ�", 1500
        .ColumnHeaders.Add , , "������", 1000
        .ColumnHeaders.Add , , "ʹ�ò���", 2000
        
        .ColumnHeaders.Add , , "�������", 0
'        .ColumnHeaders.Add , , "��ҩҵ��", 0
'        .ColumnHeaders.Add , , "����ҵ��", 0
        .ColumnHeaders.Add , , "��������", 0
        .ColumnHeaders.Add , , "ҩƷ����", 0
        .View = lvwReport
        .Icons = Me.imgEnabled
        .SmallIcons = Me.imgEnabled
    End With
End Sub

Private Sub FullData(ByVal bytType As Byte, Optional ByVal lngID As Long)
'���ܣ���������
'������
'  bytType��1-lvwDevices��2-tvwDrugType
'  lngID��
'    bytType=1��bytID��ʾ����ID
'    bytType=2��bytID����

    Dim rsTmp As ADODB.Recordset
    Dim itmX As ListItem
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    If bytType = 1 Then
        lvwDevices.ListItems.Clear
        gstrSQL = "Select a.Id, a.����, a.����, a.�ͺ�, a.����, Max(b.����) ������, Max(c.����) ʹ�ò���, " & _
                  "    Max(Decode(d.������, 1, d.����ֵ, Null)) �������, " & _
                  "    Max(Decode(d.������, 4, d.����ֵ, Null)) ��������, " & _
                  "    Max(Decode(d.������, 5, d.����ֵ, Null)) ҩƷ����  " & _
                  "From ҩ��ע���豸 A, ҩ���豸���� B, ���ű� C, " & _
                  "    (Select b.�豸id, b.����ֵ, a.������ From Zlparameters A, ҩ���豸���� B Where a.Id = b.����id) D " & _
                  "Where a.����id = b.Id And a.����id = c.Id And a.Id = d.�豸id(+) and a.����id = [1] " & _
                  "Group By a.Id, a.����, a.����, a.�ͺ�, a.���� " & _
                  "Order By ʹ�ò���, ������, ���� "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��ע���豸��Ϣ", lngID)
        Do While rsTmp.EOF = False
            intIndex = IIf(gobjComLib.zlCommFun.NVL(rsTmp!����, 0) = 0, 1, 2)
            Set itmX = lvwDevices.ListItems.Add(, "D_" & rsTmp!ID, rsTmp!����, intIndex, intIndex)
            itmX.SubItems(1) = rsTmp!����
            itmX.SubItems(2) = gobjComLib.zlCommFun.NVL(rsTmp!�ͺ�)
            itmX.SubItems(3) = rsTmp!������
            itmX.SubItems(4) = rsTmp!ʹ�ò���
            itmX.SubItems(5) = gobjComLib.zlCommFun.NVL(rsTmp!�������)
'            itmX.SubItems(6) = gobjComLib.zlCommFun.Nvl(rsTmp!��ҩҵ��)
'            itmX.SubItems(7) = gobjComLib.zlCommFun.Nvl(rsTmp!����ҵ��)
            itmX.SubItems(6) = gobjComLib.zlCommFun.NVL(rsTmp!��������)
            itmX.SubItems(7) = gobjComLib.zlCommFun.NVL(rsTmp!ҩƷ����)
            
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    Else
        tvwDrugType.Nodes.Clear
        tvwDrugType.Nodes.Add , , STR_ROOT, STR_ALL
        tvwDrugType.Nodes(STR_ROOT).Expanded = True
        
        gstrSQL = "Select ���� From ҩƷ���� Order By ���� "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩƷ����")
        Do While rsTmp.EOF = False
            With tvwDrugType.Nodes
                .Add STR_ROOT, tvwChild, rsTmp!����, rsTmp!����
            End With
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub lvwDevices_Click()
    Dim strType As String
    Dim arrType As Variant
    Dim strKey As String, strBill As String
    Dim intService As Integer
    Dim i As Integer, j As Integer
    
    If Val(lvwDevices.Tag) <> Val(Mid(lvwDevices.SelectedItem.Key, 3)) Then
        '���豸ѡ��ļ��ͣ�������lvwDevices�ġ�ҩƷ���͡�����
        If Val(lvwDevices.Tag) > 0 Then
            With lvwDevices.ListItems("D_" & Val(lvwDevices.Tag))
                .SubItems(6) = GetBillTypeStr()
                .SubItems(7) = GetDrugTypeStr()
            End With
        End If
        
        '��¼���豸
        lvwDevices.Tag = Val(Mid(lvwDevices.SelectedItem.Key, 3))
        
    Else
        Exit Sub
    End If
    
    strKey = lvwDevices.SelectedItem.Key
    
    '�������
    intService = Val(lvwDevices.SelectedItem.SubItems(5))
    
    'chkBillType
    If intService = 1 Then
        '����
        For i = 0 To chkBillType.Count - 1
            chkBillType(i).Value = 0
            chkBillType(i).Enabled = False
        Next
    Else
        'סԺ
        For i = 0 To chkBillType.Count - 1
            chkBillType(i).Enabled = True
        Next
        strBill = Trim(lvwDevices.SelectedItem.SubItems(6))     '��������
        If strBill = "" Then
            '���е��ݲ���ѡ
            For i = 0 To chkBillType.Count - 1
                chkBillType(i).Value = 0
            Next
        Else
            chkBillType(0).Value = IIf(InStr(";" & strBill & ";", ";1;") > 0, 1, 0)
            chkBillType(1).Value = IIf(InStr(";" & strBill & ";", ";2;") > 0, 1, 0)
            chkBillType(2).Value = IIf(InStr(";" & strBill & ";", ";3;") > 0, 1, 0)
        End If
        tvwDrugType.Enabled = strBill <> ""
    End If
    
    'tvwDrugType
    If tvwDrugType.Enabled Then
        strType = lvwDevices.SelectedItem.SubItems(7)               '��ѡ��ҩƷ����
        arrType = Split(strType, GSTR_SEPARAT)
        
        tvwDrugType.Nodes(STR_ROOT).Checked = False
        tvwDrugType_NodeCheck tvwDrugType.Nodes(STR_ROOT)
        If strType <> "" Then
            For i = LBound(arrType) To UBound(arrType)
                If i = LBound(arrType) And arrType(i) = STR_ALL Then
                    tvwDrugType.Nodes(STR_ROOT).Checked = True
                    tvwDrugType_NodeCheck tvwDrugType.Nodes(STR_ROOT)
                    Exit For
                Else
                    'ͬ����ѡ�ļ���
                    For j = 2 To tvwDrugType.Nodes.Count
                        If tvwDrugType.Nodes(j).Text = arrType(i) Then
                            tvwDrugType.Nodes(j).Checked = True
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
    End If
    
End Sub

Private Sub tvwDrugType_NodeCheck(ByVal Node As MSComctlLib.Node)
'˵����tvwDrugType�ؼ�����ѡ�Ĵ���
    Dim i As Integer
    Dim blnAll As Boolean
    
    If Node.Parent Is Nothing Then
        For i = 1 To tvwDrugType.Nodes.Count
            tvwDrugType.Nodes(i).Checked = Node.Checked
        Next
    Else
        blnAll = True
        For i = 1 To tvwDrugType.Nodes.Count
            If tvwDrugType.Nodes(i).Checked = False And tvwDrugType.Nodes(i).Key <> STR_ROOT Then
                blnAll = False
                Exit For
            End If
        Next
        Node.Parent.Checked = blnAll
    End If
End Sub

Private Function GetDrugTypeStr() As String
'���ܣ���ȡ��ǰtvwDrugType�ļ���ѡ���ַ���
    Dim i As Integer
    Dim strTmp As String
    
    With tvwDrugType
        If .Nodes(STR_ROOT).Checked Then
            strTmp = STR_ALL
        Else
            For i = 1 To tvwDrugType.Nodes.Count
                If tvwDrugType.Nodes(i).Checked Then
                    strTmp = strTmp & tvwDrugType.Nodes(i).Text
                    strTmp = strTmp & GSTR_SEPARAT
                End If
            Next
            If strTmp <> "" Then
                strTmp = Left(strTmp, Len(strTmp) - 1)
            End If
        End If
        GetDrugTypeStr = strTmp
    End With
    
End Function

Private Function GetBillTypeStr() As String
'���ܣ���ȡ��ǰ���������ַ���
    Dim strBill As String

    If chkBillType(0).Value = 1 Then
        strBill = strBill & "1" & GSTR_SEPARAT_CHILD
    End If
    If chkBillType(1).Value = 1 Then
        strBill = strBill & "2" & GSTR_SEPARAT_CHILD
    End If
    If chkBillType(2).Value = 1 Then
        strBill = strBill & "3" & GSTR_SEPARAT_CHILD
    End If
    If strBill <> "" Then strBill = Left(strBill, Len(strBill) - 1)
    GetBillTypeStr = strBill
End Function

