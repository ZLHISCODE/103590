VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPACSImageDeviceSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӱ���豸����"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   FillColor       =   &H00FF0000&
   Icon            =   "frmPACSImageDeviceSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7170
      TabIndex        =   2
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6030
      TabIndex        =   1
      Top             =   5340
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwImageDevice 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   8705
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�豸��"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�豸��"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP��ַ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�˿ں�"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "����AE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Զ���豸AE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ĭ���豸"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˫��ѡ��һ���豸��ΪĬ���豸��"
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   2700
   End
End
Attribute VB_Name = "frmPACSImageDeviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrDeviceName As String
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ĭ��Ӱ���豸", mstrDeviceName
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objItem As ListItem
    
    strSQL = "select �豸�� , �豸��, IP��ַ,�˿ں�,����AE,�豸AE from Ӱ���豸Ŀ¼ where ���� = 4 "
    
    On Error GoTo ErrOther
    
    '��������Ӱ���豸
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Me.lvwImageDevice.ListItems.Clear
    Do Until rsTmp.EOF
        With Me.lvwImageDevice.ListItems
            Set objItem = .Add(, "_" & rsTmp("�豸��"), rsTmp("�豸��"))
            objItem.SubItems(1) = rsTmp("�豸��")
            objItem.SubItems(2) = rsTmp("IP��ַ")
            objItem.SubItems(3) = Nvl(rsTmp("�˿ں�"))
            objItem.SubItems(4) = Nvl(rsTmp("����AE"))
            objItem.SubItems(5) = Nvl(rsTmp("�豸AE"))
        End With
        rsTmp.MoveNext
    Loop
    
    '�����豸Ĭ��ֵ
    mstrDeviceName = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "Ĭ��Ӱ���豸", "")
    If mstrDeviceName <> "" Then
        Set objItem = Me.lvwImageDevice.FindItem(Mid(mstrDeviceName, 2))
        If Not objItem Is Nothing Then
            objItem.SubItems(6) = "��"
        Else
'            If Me.lvwImageDevice.ListItems.Count > 0 Then Me.lvwImageDevice.ListItems(1).SubItems(6) = "��"
        End If
    Else
'        If Me.lvwImageDevice.ListItems.Count > 0 Then Me.lvwImageDevice.ListItems(1).SubItems(6) = "��"
    End If
    
    Exit Sub
    
ErrOther:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Sub subSelectDefault()
    '------------------------------------------------
    '���ܣ�ѡ��һ��Ĭ���豸
    '������
    '���أ���
    '�ϼ���������̣�lvwImageDevice_DblClick��
    '�¼���������̣���
    '���õ��ⲿ������mstrDeviceName
    '�����ˣ�����
    '------------------------------------------------
    Dim i As Integer
    
    If Me.lvwImageDevice.SelectedItem Is Nothing Then Exit Sub
    
    With Me.lvwImageDevice
        For i = 1 To .ListItems.Count
            .ListItems(i).SubItems(6) = ""
        Next
    End With
    Me.lvwImageDevice.SelectedItem.SubItems(6) = "��"
    mstrDeviceName = Me.lvwImageDevice.SelectedItem.Key
End Sub

Private Sub lvwImageDevice_DblClick()
    subSelectDefault
End Sub
