VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendItemChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀѡ��"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   Icon            =   "frmTendItemChoose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   2640
      Top             =   1710
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
            Picture         =   "frmTendItemChoose.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   2
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3210
      TabIndex        =   1
      Top             =   3480
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5900
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLvw"
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��Ŀ���"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "��Ŀ����"
         Object.Width           =   4410
      EndProperty
   End
End
Attribute VB_Name = "frmTendItemChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng��Ŀ��� As Long
Private mstr��Ŀ���� As String
Private mstrSelItems As String
Private mrsItems As New ADODB.Recordset

Public Function ShowSelect(ByVal strSelItems As String, ByVal byt����ȼ� As Integer, ByVal intӤ�� As Integer, ByVal lng����ID As Long) As String
    On Error Resume Next
    Dim lvwItem As ListItem
    
    mblnOK = False
    mstrSelItems = strSelItems
    '����ǰ�Ĺ�����ȡ��Ŀ�嵥��¼��
    gstrSQL = " Select B.��Ŀ���,B.��Ŀ���� " & _
             " From �����¼��Ŀ B" & _
             " Where B.Ӧ�÷�ʽ<>0 " & IIf(byt����ȼ� = -1, "", " And B.����ȼ�>=[1]") & IIf(intӤ�� = -1, "", " And B.���ò��� IN (0,[2])") & _
             " And (B.���ÿ���=1 Or (B.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=B.��Ŀ��� And D.����id=[3])))" & _
             " Order by B.��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���п��õĻ�����Ŀ", byt����ȼ�, IIf(intӤ�� = 0, 1, 2), lng����ID)
    If mrsItems.RecordCount = 0 Then
        MsgBox "û�пɹ���ӵ���Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    '����ѡ�����Ŀ����ؼ���
    lvwItems.ListItems.Clear
    With mrsItems
        Do While Not .EOF
            If InStr(1, mstrSelItems, "," & !��Ŀ��� & ",") = 0 Then
                Set lvwItem = lvwItems.ListItems.Add(, "K" & lvwItems.ListItems.Count, !��Ŀ���, , 1)
                lvwItem.SubItems(1) = !��Ŀ����
            End If
            .MoveNext
        Loop
    End With
    If lvwItems.ListItems.Count = 0 Then
        MsgBox "û�пɹ���ӵ���Ŀ��", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    Me.Show 1
    If mblnOK Then ShowSelect = mlng��Ŀ��� & "|" & mstr��Ŀ����
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    mlng��Ŀ��� = lvwItems.SelectedItem
    mstr��Ŀ���� = lvwItems.SelectedItem.SubItems(1)
    mblnOK = True
    Unload Me
End Sub

Private Sub lvwItems_DblClick()
    Call lvwItems_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    If KeyCode = vbKeyReturn Then Call cmdȷ��_Click
End Sub
