VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBillShowInOutClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ݷ�����ʾ"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Height          =   3765
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6641
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
      Top             =   3840
      Width           =   1100
   End
   Begin MSComctlLib.ListView Lvw�������б� 
      Height          =   3765
      Left            =   2490
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6641
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   4245
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8916
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgLeftRight 
      Height          =   3675
      Left            =   2460
      MousePointer    =   9  'Size W E
      Top             =   60
      Width           =   45
   End
End
Attribute VB_Name = "frmBillShowInOutClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BlnStartUp As Boolean                   '�����ɹ����
Private strSQL As String
Private RecClass As New ADODB.Recordset         '���ϵ��ݷ���



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
    
    With RecClass
        If .State = 1 Then .Close
        strSQL = "Select ����,����,����,˵�� From ҩƷ���ݷ��� where ����>=30 Order by ����"
        
        Call SQLTest(App.Title, Me.Caption, strSQL)
        .Open strSQL, gcnOracle
        Call SQLTest
        
        If .EOF Then
            MsgBox "���ϵ��ݷ������ݲ�ȫ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    DependOnCheck = True
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

Private Sub ImgLeftRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight
        If .Left + x < 2000 Then Exit Sub
        If .Left + x > Me.ScaleWidth - 3500 Then Exit Sub
        
        .Move .Left + x
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
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

Private Sub Lvw�����б�_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '--��ָ�����ݰ����Ĳ������������--
    Dim StrInfo As String
    
    With RecClass
        If .State = 1 Then .Close
        strSQL = "Select ����,����,Decode(ϵ��,1,'���','����') as ϵ�� From ҩƷ������ Where ID IN " & _
                 " (Select ���ID From ҩƷ�������� Where ����=" & Mid(Lvw�����б�.SelectedItem.Key, 3) & ") Order by ���� "
        
        Call SQLTest(App.Title, Me.Caption, strSQL)
        .Open strSQL, gcnOracle
        Call SQLTest
        
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
        stbThis.Panels(2) = StrInfo
    End With
    
    LoadInLvw
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
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub
