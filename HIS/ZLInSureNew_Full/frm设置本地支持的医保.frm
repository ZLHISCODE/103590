VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���ñ���֧�ֵ�ҽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ñ���֧�ֵ�ҽ��"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frm���ñ���֧�ֵ�ҽ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   4230
      Top             =   180
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
            Picture         =   "frm���ñ���֧�ֵ�ҽ��.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5640
      TabIndex        =   3
      Top             =   1470
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5640
      TabIndex        =   2
      Top             =   960
      Width           =   1100
   End
   Begin MSComctlLib.ListView lvw����ҽ�� 
      Height          =   3435
      Left            =   240
      TabIndex        =   0
      Top             =   780
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
      NumItems        =   4
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
         Text            =   "˵��"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ѡ�񱾵�֧�ֵ�ҽ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1170
      TabIndex        =   1
      Top             =   360
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   390
      Picture         =   "frm���ñ���֧�ֵ�ҽ��.frx":24F4
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frm���ñ���֧�ֵ�ҽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStart As Boolean                '�Ƿ�����
Private mstrSelect As String                '֧�ֵ�ҽ��

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim intDO As Integer, intCOUNT As Integer
    Dim int��� As Integer          'ĳԺ������ͬʱʹ�ö�����ҽ������
    
    intCOUNT = lvw����ҽ��.ListItems.Count
    '����ж��ٸ����ҽ������
    For intDO = 1 To intCOUNT
        If lvw����ҽ��.ListItems(intDO).Checked Then
            If lvw����ҽ��.ListItems(intDO).Tag = 1 Then
                int��� = int��� + 1
            End If
        End If
    Next
    If int��� > 1 Then
        MsgBox "Ŀǰ��֧��ͬʱʹ�ö�����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��֯�ɴ�
    For intDO = 1 To intCOUNT
        If lvw����ҽ��.ListItems(intDO).Checked Then mstrSelect = mstrSelect & "," & Mid(lvw����ҽ��.ListItems(intDO).Key, 3)
    Next
    If mstrSelect <> "" Then mstrSelect = Mid(mstrSelect, 2)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", mstrSelect)
    
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
    Dim lvwItem As ListItem
    Dim intDO As Integer, intCOUNT As Integer
    Dim rsTemp As New ADODB.Recordset
    mstrSelect = GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", "")
    
    '˵����ѡ�񱾵�֧�ֵ�����
    gstrSQL = " Select A.���,A.����,A.˵��,Nvl(A.���,0) AS ���" & _
              " From ������� A Where Nvl(�Ƿ��ֹ,0)=0" & _
              " Order By A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ѱ�װ������ʹ�õ�ҽ���ӿ�")
    If rsTemp.RecordCount = 0 Then
        MsgBox "����û�а�װ�κ�ҽ���ӿڣ����޷�Ϊ����ѡ��֧�ֵ�ҽ����", vbInformation, gstrSysName
        Exit Sub
    Else
'        rsTemp.Filter = "����=1"
'        If rsTemp.RecordCount = 0 Then
'            MsgBox "����û������ҽ���ӿڣ����޷�Ϊ����ѡ��֧�ֵ�ҽ����", vbInformation, gstrSysName
'            rsTemp.Filter = 0
'            Exit Sub
'        End If
    End If
    
    Me.lvw����ҽ��.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set lvwItem = lvw����ҽ��.ListItems.Add(, "K_" & !���, !����, , 1)
            lvwItem.SubItems(1) = !���
            lvwItem.SubItems(2) = IIf(!��� = 1, "��", "��")
            lvwItem.SubItems(3) = Nvl(!˵��)
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
        lvw����ҽ��.ListItems("K_" & arrSelect(intDO)).Checked = True
    Next
    
    mstrSelect = ""
    mblnStart = True
End Sub
