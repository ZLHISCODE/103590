VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmѡ��ǰҽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��ǰҽ��"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmѡ��ǰҽ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2820
      TabIndex        =   1
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   2
      Top             =   2580
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
            Picture         =   "frmѡ��ǰҽ��.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw����ҽ�� 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmѡ��ǰҽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStart As Boolean
Private mint���� As Integer
Private mstrSelect As String
'�ӱ���֧�ֵ�ҽ���ӿ��н���ѡ�����������շ���סԺ�Ǽǵȴ��ڳ���

Private Sub cmdȡ��_Click()
    mint���� = 0
    Unload Me
    Exit Sub
End Sub

Private Sub cmdȷ��_Click()
    If lvw����ҽ��.SelectedItem Is Nothing Then Exit Sub
    mint���� = Val(lvw����ҽ��.SelectedItem)
    
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
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    mblnStart = False
    mstrSelect = GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", "")
    If mstrSelect = "" Then Exit Sub
    
    '˵����ѡ�񱾵�֧�ֵ�����
    gstrSQL = " Select A.���,A.����,A.˵��" & _
              " From ������� A,Table(Cast(f_num2list([1]) As Zltools.t_Numlist)) B" & _
              " Where A.��� = B.Column_Value And A.ҽ���� Is NULL" & _
              " UNION " & _
              " Select DISTINCT A.���,A.����,A.˵��" & _
              " From ������� B,������� A,Table(Cast(f_num2list([1]) As Zltools.t_Numlist)) C" & _
              " Where B.ҽ����=A.ҽ������ And B.��� = C.Column_Value" & _
              " Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����֧�ֵ�����", mstrSelect)
    If rsTemp.RecordCount = 0 Then
        MsgBox "��δ���ñ���֧�ֵ�ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lvw����ҽ��.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw����ҽ��.ListItems.Add(, "K_" & !���, !���, , 1)
            lvwItem.SubItems(1) = !����
            .MoveNext
        Loop
        lvw����ҽ��.ListItems(1).Selected = True
    End With
    
    mblnStart = True
End Sub

Public Function ShowSelect() As Integer
    '��ʾ����
    Dim rtn         As Long
    rtn = SetWindowPos(Me.hwnd, -1, CurrentX, CurrentY, 0, 0, 3)
    Me.Show 1
    ShowSelect = mint����
End Function

Private Sub lvw����ҽ��_DblClick()
    Call cmdȷ��_Click
End Sub

Private Sub lvw����ҽ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call cmdȷ��_Click
End Sub
