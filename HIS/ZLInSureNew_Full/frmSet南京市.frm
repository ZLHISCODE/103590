VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSet�Ͼ��� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frmSet�Ͼ���.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   1290
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   3990
      Top             =   0
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
            Picture         =   "frmSet�Ͼ���.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw����ҽ�� 
      Height          =   3435
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "���"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "�Ż����"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   150
      Picture         =   "frmSet�Ͼ���.frx":128E
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ����������ҽ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   930
      TabIndex        =   3
      Top             =   180
      Width           =   1620
   End
End
Attribute VB_Name = "frmSet�Ͼ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer                '��ǰ��ҽ��������ʾ���嵥��
Private mblnStart As Boolean                '�Ƿ�����
Private mstrSelect As String                '֧�ֵ�ҽ��
Private mblnOK As Boolean

Public Function ��������(ByVal intinsure As Integer) As Boolean
    mintInsure = intinsure
    mblnOK = False
    Me.Show 1
    �������� = mblnOK
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim intDO As Integer, intCOUNT As Integer
    
    intCOUNT = lvw����ҽ��.ListItems.Count
    
    '��֯�ɴ�
    For intDO = 1 To intCOUNT
        If lvw����ҽ��.ListItems(intDO).Checked Then mstrSelect = mstrSelect & "," & Mid(lvw����ҽ��.ListItems(intDO).Key, 3) & ";" & lvw����ҽ��.ListItems(intDO).SubItems(3)
    Next
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����ҽ���ӿ�", mstrSelect)
    
    mblnOK = True
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim arrSelect
    Dim str�Ż���� As String
    Dim lvwItem As ListItem
    Dim intDO As Integer, intCOUNT As Integer
    Dim rsTemp As New ADODB.Recordset
    mstrSelect = GetSetting("ZLSOFT", "����ȫ��", "����ҽ���ӿ�", "")
    
    '˵����ѡ�񱾵�֧�ֵ�����
    gstrSQL = " Select A.���,A.����,A.˵��,Nvl(A.���,0) AS ���" & _
              " From ������� A " & _
              " Where Nvl(�Ƿ��ֹ,0)=0 and ҽ������ Is Not NULL And ���<>[1]" & _
              " Order By A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ѱ�װ������ʹ�õ�ҽ���ӿ�", mintInsure)
    If rsTemp.RecordCount = 0 Then
        MsgBox "����û�а�װ�κ�ҽ���ӿڣ����޷�Ϊ����ѡ��֧�ֵ�ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Me.lvw����ҽ��.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw����ҽ��.ListItems.Add(, "K_" & !���, !����, , 1)
            lvwItem.SubItems(1) = !���
            lvwItem.SubItems(2) = IIf(!��� = 1, "��", "��")
            lvwItem.SubItems(3) = "��ͨ"    '0-��ͨ;1-����;2-����;3-�����
            lvwItem.SubItems(4) = Nvl(!˵��)
            
            lvwItem.Tag = !���
            .MoveNext
        Loop
        lvw����ҽ��.ListItems(1).Selected = True
    End With
    
    '��ʾ����֧�ֵ�ҽ��
    On Error Resume Next
    arrSelect = Split(mstrSelect, ",")
    intCOUNT = UBound(arrSelect)
    For intDO = 0 To intCOUNT
        lvw����ҽ��.ListItems("K_" & Split(arrSelect(intDO), ";")(0)).Checked = True
        lvw����ҽ��.ListItems("K_" & Split(arrSelect(intDO), ";")(0)).SubItems(3) = Split(arrSelect(intDO), ";")(1)
    Next
    
    mstrSelect = ""
    mblnStart = True
End Sub

Private Sub lvw����ҽ��_DblClick()
    Dim str�Ż���� As String
    str�Ż���� = lvw����ҽ��.SelectedItem.SubItems(3)
    Select Case str�Ż����
    Case "��ͨ"
        str�Ż���� = "����"
    Case "����"
        str�Ż���� = "����"
    Case "����"
        str�Ż���� = "�����"
    Case Else
        str�Ż���� = "��ͨ"
    End Select
    lvw����ҽ��.SelectedItem.SubItems(3) = str�Ż����
End Sub
