VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmByBillShow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ݷ�����ʾ"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   Icon            =   "frmByBillShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList ImgPublic 
      Left            =   7020
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lvw�����б� 
      Height          =   4125
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "��������"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "�˳�(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6480
      TabIndex        =   1
      Top             =   4200
      Width           =   1100
   End
   Begin MSComctlLib.ListView Lvw�������б� 
      Height          =   4125
      Left            =   2490
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblInfo 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4300
      Width           =   6015
   End
   Begin VB.Image ImgLeftRight 
      Height          =   3675
      Left            =   2460
      MousePointer    =   9  'Size W E
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "frmByBillShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BlnStartUp As Boolean                   '�����ɹ����
Private strSQL As String
Private RecClass As New ADODB.Recordset         'ҩƷ���ݷ���



Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    BlnStartUp = False
    
    If DependOnCheck = False Then Exit Sub
    If LoadInIcon = False Then Exit Sub
    LoadInTvw
    Call RestoreWinState(Me, App.ProductName)
    
    BlnStartUp = True
End Sub

Private Function DependOnCheck() As Boolean
    DependOnCheck = False
    '--�������ݼ��--
    
    On Error GoTo errHandle
    With RecClass
'        If .State = 1 Then .Close
        strSQL = "Select ����,����,����,˵�� From ҩƷ���ݷ��� Order by ����"
        
'        Call SQLTest(App.Title, Me.Caption, strSQL)
        Set RecClass = zldatabase.OpenSQLRecord(strSQL, "DependOnCheck")
'        Call SQLTest
        
        If .EOF Then
            MsgBox "ҩƷ���ݷ������ݲ�ȫ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadInIcon() As Boolean
    '--Ϊ���ؼ�װ��ͼ��--
    On Error Resume Next
    Err = 0
    LoadInIcon = False
    
    '--�б�Lvw��������--
    With ImgPublic
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add , , LoadResPicture("BILL1", vbResIcon)
        .ListImages.Add , , LoadResPicture("CLASS", vbResIcon)
    End With
    With Lvw�����б�
        Set .SmallIcons = ImgPublic
    End With
    With Lvw�������б�
        Set .SmallIcons = ImgPublic
    End With
    
    If Err <> 0 Then
        MsgBox "�����Դ�ļ���ʧ�����������������ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    LoadInIcon = True
End Function

Private Function LoadInTvw()
    '--�����ݷ���װ�����Ϳؼ�--
    
    Dim ItemThis As ListItem
    With RecClass
        Do While Not .EOF
            Set ItemThis = Lvw�����б�.ListItems.Add(, "K_" & !����, !����, , 1)
            ItemThis.Tag = !����
            
            .MoveNext
        Loop
    End With
    
    With Lvw�����б�
        .ListItems(1).Selected = True
        .SelectedItem.Selected = True
    End With
    Lvw�����б�_ItemClick Lvw�����б�.SelectedItem
End Function

Private Sub ImgLeftRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight
        If .Left + X < 2000 Then Exit Sub
        If .Left + X > Me.ScaleWidth - 3500 Then Exit Sub
        
        .Move .Left + X
    End With
    
    With Me.Lvw�����б�
        .Width = ImgLeftRight.Left
    End With
    
    With Me.Lvw�������б�
        .Left = ImgLeftRight.Left + ImgLeftRight.Width
        .Width = Me.ScaleWidth - .Left
    End With
End Sub

Private Sub Lvw�����б�_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw�����б�
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw�����б�_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--��ָ�����ݰ�����ҩƷ���������--
    Dim StrInfo As String
    
    On Error GoTo errHandle
    strSQL = "Select ����,����,Decode(ϵ��,1,'���','����') as ϵ�� From ҩƷ������ Where ID IN " & _
             " (Select ���ID From ҩƷ�������� Where ����=[1]) Order by ���� "
    Set RecClass = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(Lvw�����б�.SelectedItem.Key, 3)))
        
    With RecClass
        '��ʾָ�����ݵ�˵����Ϣ
        Select Case Lvw�����б�.SelectedItem.Tag
        Case "1"
            StrInfo = "�õ���ֻ����һ��������"
        Case "2"
            StrInfo = "�õ���ֻ����һ�ֳ������"
        Case "3"
            StrInfo = "�õ���ֻ����һ��������һ�ֳ������"
        Case "4"
            StrInfo = "�õ����������������"
        Case "5"
            StrInfo = "�õ���������ֳ������"
        End Select
        lblInfo.Caption = StrInfo
    End With
    
    LoadInLvw
    Exit Sub
errHandle:
    If ERRCENTER() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LoadInLvw()
    '��������д��
    Dim ItemThis As ListItem
    
    Lvw�������б�.ListItems.Clear
    With RecClass
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Set ItemThis = Lvw�������б�.ListItems.Add(, , !����, , 2)
            ItemThis.SubItems(1) = !����
            ItemThis.SubItems(2) = !ϵ��
            .MoveNext
        Loop
    End With
End Function

Private Sub Lvw�������б�_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Lvw�������б�
        .Sorted = False
        
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIF(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub
