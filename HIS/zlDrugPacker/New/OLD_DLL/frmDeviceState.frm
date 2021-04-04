VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeviceState 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�豸����"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frmDeviceState.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6735
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   360
      Left            =   4200
      TabIndex        =   1
      Top             =   2760
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      Top             =   2760
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwDevices 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgEnabled 
      Left            =   360
      Top             =   2400
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
            Picture         =   "frmDeviceState.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeviceState.frx":2AF4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDeviceState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDeptID As Long

Public Sub ShowMe(ByVal lngDeptID As Long)
    mlngDeptID = lngDeptID
    
    Call Init
    Call FullData
    
    If lvwDevices.ListItems.Count = 0 Then
        MsgBox "��δע��ҩ���Զ����豸��", vbInformation, GSTR_INTERFACE_NAME
        Unload Me
        Exit Sub
    End If
    
    Show vbModal, gfrmOwner
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim itmX As ListItem
    
    On Error GoTo errHandle
    gobjConn.BeginTrans
    For i = 1 To lvwDevices.ListItems.Count
        Set itmX = lvwDevices.ListItems(i)
        gstrSQL = "Zl_ҩ��ע���豸_Switch(" & _
                  Mid(itmX.Key, 3) & "," & _
                  IIf(itmX.Checked, "1", "0") & _
                  ")"
        Call gobjComLib.zldatabase.ExecuteProcedure(gstrSQL, "�����豸����״̬")
    Next
    gobjConn.CommitTrans
    
    Unload Me
    Exit Sub
    
errHandle:
    gobjConn.CommitTrans
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
        .View = lvwReport
        .Icons = Me.imgEnabled
        .SmallIcons = Me.imgEnabled
    End With
End Sub

Private Sub FullData()
    Dim rsTmp As ADODB.Recordset
    Dim itmX As ListItem
    Dim intIndex As Integer
    
    gstrSQL = "Select a.Id, a.����, a.����, a.�ͺ�, a.����, b.���� ������, c.���� ʹ�ò��� " & _
              "From ҩ��ע���豸 A, ҩ���豸���� B, ���ű� C " & _
              "Where a.����id = b.Id And a.����id = c.Id "
    If mlngDeptID <> 0 Then
        gstrSQL = gstrSQL & "and a.����id = [1] "
    End If
    gstrSQL = gstrSQL & "Order By c.����, b.����, a.���� "
    
    lvwDevices.ListItems.Clear
    
    On Error GoTo errHandle
    Set rsTmp = gobjComLib.zldatabase.OpenSQLRecord(gstrSQL, "��ȡҩ��ע���豸��Ϣ", mlngDeptID)
    Do While rsTmp.EOF = False
        intIndex = IIf(gobjComLib.zlCommFun.Nvl(rsTmp!����, 0) = 0, 1, 2)
        Set itmX = lvwDevices.ListItems.Add(, "D_" & rsTmp!ID, rsTmp!����, intIndex, intIndex)
        itmX.SubItems(1) = rsTmp!����
        itmX.SubItems(2) = gobjComLib.zlCommFun.Nvl(rsTmp!�ͺ�)
        itmX.SubItems(3) = rsTmp!������
        itmX.SubItems(4) = rsTmp!ʹ�ò���
        itmX.Checked = intIndex = 2
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    Exit Sub
    
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub

Private Sub lvwDevices_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Icon = IIf(Item.Checked, 2, 1)
    Item.SmallIcon = IIf(Item.Checked, 2, 1)
End Sub
