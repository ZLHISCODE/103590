VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "frmSetting"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8910
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgEnabled 
      Left            =   3720
      Top             =   4800
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
            Picture         =   "frmSetting.frx":15162
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetting.frx":154B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraMain 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdDeviceDel 
         Caption         =   "ɾ��(&D)"
         Height          =   360
         Left            =   7650
         TabIndex        =   10
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdDeviceEdit 
         Caption         =   "�޸�(&E)"
         Height          =   360
         Left            =   6690
         TabIndex        =   9
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdDeviceCreate 
         Caption         =   "�½�(&C)"
         Height          =   360
         Left            =   5730
         TabIndex        =   8
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdLinkDel 
         Caption         =   "ɾ��(&D)"
         Height          =   360
         Left            =   2040
         TabIndex        =   5
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdLinkEdit 
         Caption         =   "�޸�(&E)"
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdLinkCreate 
         Caption         =   "�½�(&C)"
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwDevices 
         Height          =   3330
         Left            =   3000
         TabIndex        =   7
         Top             =   465
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5874
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
      Begin VB.ListBox lstLink 
         Height          =   3300
         ItemData        =   "frmSetting.frx":15806
         Left            =   120
         List            =   "frmSetting.frx":15808
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblDevices 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�豸�б�"
         Height          =   180
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����б�"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   360
      Left            =   7520
      TabIndex        =   11
      Top             =   4800
      Width           =   1110
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceCreate_Click()
    If lstLink.ListIndex < 0 Then lstLink.ListIndex = 0

    frmDeviceReg.ShowMe Me, 0, lstLink.ItemData(lstLink.ListIndex)
    Call FullData(2, lstLink.ItemData(lstLink.ListIndex))
    
    If lvwDevices.ListItems.Count > 0 Then lvwDevices.ListItems(lvwDevices.ListItems.Count).Selected = True
    
End Sub

Private Sub cmdDeviceDel_Click()
    Dim strTmp As String
    Dim intItem As Integer
    
    If lvwDevices.SelectedItem Is Nothing Then Exit Sub
    
    strTmp = lvwDevices.SelectedItem.Text
    If MsgBox("�Ƿ�ɾ����" & strTmp & "���豸��", vbInformation + vbYesNo + vbDefaultButton2, GSTR_INTERFACE_NAME) = vbNo Then Exit Sub
    
    intItem = lvwDevices.SelectedItem.Index
    strTmp = Mid(lvwDevices.SelectedItem.Key, 3)
    gstrSQL = "ZL_ҩ��ע���豸_DELETE(" & strTmp & ")"
    On Error GoTo errHandle
    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "ɾ��ҩ��ע���豸")
    On Error GoTo 0
    
    Call FullData(2, lstLink.ItemData(lstLink.ListIndex))
    
    If intItem > 1 Then lvwDevices.ListItems(intItem - 1).Selected = True
    
    Exit Sub

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cmdDeviceEdit_Click()
    If lvwDevices.SelectedItem Is Nothing Then Exit Sub
    
    Dim intItem As Integer
    
    intItem = lvwDevices.SelectedItem.Index
    frmDeviceReg.ShowMe Me, 1, Val(lvwDevices.SelectedItem.Tag)
    Call FullData(2, lstLink.ItemData(lstLink.ListIndex))
    
    If lvwDevices.ListItems.Count = 0 Then Exit Sub
    
    If intItem > lvwDevices.ListItems.Count Then
        lvwDevices.ListItems(lvwDevices.ListItems.Count).Selected = True
    Else
        lvwDevices.ListItems(intItem).Selected = True
    End If
End Sub

Private Sub cmdLinkCreate_Click()
    frmLink.ShowMe Me
    Call FullData(1)
    
    If lstLink.ListCount > 0 Then lstLink.ListIndex = lstLink.ListCount - 1

End Sub

Private Sub cmdLinkDel_Click()
    Dim strTmp As String
    Dim lngID As Long
    Dim intItem As Integer
    
    If lstLink.ListIndex < 0 Then Exit Sub
    
    strTmp = lstLink.List(lstLink.ListIndex)
    lngID = lstLink.ItemData(lstLink.ListIndex)
    
    If MsgBox("�Ƿ�ɾ����" & strTmp & "�����ӣ�", vbInformation + vbYesNo + vbDefaultButton2, GSTR_INTERFACE_NAME) = vbNo Then Exit Sub
    
    intItem = lstLink.ListIndex
    gstrSQL = "ZL_ҩ���豸����_DELETE(" & lngID & ")"
    On Error GoTo errHandle
    Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "ɾ��ҩ���豸����")
    On Error GoTo 0
    
    Call FullData(1)
    
    If intItem > 0 Then lstLink.ListIndex = intItem - 1
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub cmdLinkEdit_Click()
    If lstLink.ListIndex < 0 Then Exit Sub
    
    Dim intItem As Integer
    
    intItem = lstLink.ListIndex
    frmLink.ShowMe Me, lstLink.ItemData(lstLink.ListIndex)
    Call FullData(1)
    
    lstLink.ListIndex = intItem
    
End Sub

Private Sub Form_Load()
    Call Init
    Call FullData(1)
    If lstLink.ListCount > 0 Then lstLink.ListIndex = 0
End Sub

Private Sub lstLink_Click()
    cmdLinkEdit.Enabled = lstLink.ListCount > 0
    cmdLinkDel.Enabled = cmdLinkEdit.Enabled
    Call FullData(2, lstLink.ItemData(lstLink.ListIndex))
End Sub

Private Sub FullData(ByVal bytType As Byte, Optional ByVal lngID As Long)
'���ܣ��������
'������
'   bytType���������ͣ�bytType=1��ʾ���ӣ�bytType=2��ʾ�豸
'   lngID������ID
    
    Dim rsTmp As ADODB.Recordset

    On Error GoTo errHandle
    If bytType = 1 Then
    
        lstLink.Clear
        gstrSQL = "Select ID, ���� From ҩ���豸���� Order by ���� "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡ�Զ���ϵͳ����������")
        Do While Not rsTmp.EOF
            lstLink.AddItem rsTmp!����
            lstLink.ItemData(lstLink.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
    ElseIf bytType = 2 Then
    
        Dim itmX As ListItem
        Dim intIndex As Integer
    
        lvwDevices.ListItems.Clear
        gstrSQL = "Select a.ID, a.����, a.����, a.�ͺ�, a.������, a.����, b.���� ʹ�ò��� " & _
                  "From ҩ��ע���豸 A, ���ű� B " & _
                  "Where a.����id = b.id and a.����ID = [1] " & _
                  "Order by a.���� "
        Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ���豸", lngID)
        Do While Not rsTmp.EOF
            intIndex = IIf(gobjComLib.zlCommFun.Nvl(rsTmp!����, 0) = 0, 1, 2)
            Set itmX = lvwDevices.ListItems.Add(, "D_" & rsTmp!ID, rsTmp!����, intIndex, intIndex)
            itmX.Tag = rsTmp!ID
            itmX.SubItems(1) = rsTmp!����
            itmX.SubItems(2) = gobjComLib.zlCommFun.Nvl(rsTmp!�ͺ�)
            itmX.SubItems(3) = gobjComLib.zlCommFun.Nvl(rsTmp!������)
            itmX.SubItems(4) = rsTmp!ʹ�ò���
            itmX.Checked = intIndex = 2
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
    End If
    
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub Init()
    
    With lvwDevices
        .ColumnHeaders.Add , , "����", 1000
        .ColumnHeaders.Add , , "����", 1500
        .ColumnHeaders.Add , , "�ͺ�", 1500
        .ColumnHeaders.Add , , "������", 2000
        .ColumnHeaders.Add , , "ʹ�ò���", 2000
        .View = lvwReport
        .Icons = Me.imgEnabled
        .SmallIcons = Me.imgEnabled
    End With

End Sub

Private Sub lvwDevices_ItemCheck(ByVal Item As MSComctlLib.ListItem)
MsgBox "1"
    Item.Icon = IIf(Item.Checked, 2, 1)
MsgBox "2"
    Item.SmallIcon = IIf(Item.Checked, 2, 1)
MsgBox "3"
End Sub
