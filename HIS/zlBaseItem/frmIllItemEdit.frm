VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIllItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������༭"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmIllItemEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -45
      TabIndex        =   16
      Top             =   5370
      Width           =   5865
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4485
      TabIndex        =   15
      Top             =   5535
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3255
      TabIndex        =   14
      Top             =   5535
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   2355
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
            Picture         =   "frmIllItemEdit.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIllItemEdit.frx":0326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   0
      Left            =   180
      TabIndex        =   18
      Top             =   555
      Visible         =   0   'False
      Width           =   5280
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "ͳ����"
         Top             =   1815
         Width           =   1680
      End
      Begin VB.CheckBox Chk���� 
         Caption         =   "�Ƿ�¼�������Ϣ"
         Height          =   240
         Left            =   1005
         TabIndex        =   11
         Top             =   2205
         Width           =   1815
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   270
         Left            =   4890
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4320
         Width           =   255
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1815
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1005
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "����"
         Top             =   135
         Width           =   1395
      End
      Begin VB.ComboBox cmb��Ч 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1410
         Width           =   1695
      End
      Begin VB.ComboBox cmb�Ա� 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1410
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4305
         Width           =   3885
      End
      Begin VB.TextBox txtEdit 
         Height          =   1710
         Index           =   4
         Left            =   1005
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Tag             =   "˵��"
         Top             =   2535
         Width           =   4155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1005
         MaxLength       =   150
         TabIndex        =   2
         Tag             =   "����"
         Top             =   555
         Width           =   4155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "ƴ����"
         Top             =   975
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   3210
         MaxLength       =   20
         TabIndex        =   4
         Tag             =   "�����"
         Top             =   975
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   1
         Tag             =   "����"
         Top             =   135
         Width           =   1680
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ͳ����(&A)"
         Height          =   180
         Index           =   9
         Left            =   2670
         TabIndex        =   9
         Top             =   1875
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&L)"
         Height          =   180
         Index           =   8
         Left            =   0
         TabIndex        =   7
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&J)"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   29
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "˵��(&E)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   28
         Top             =   2535
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   1
         Left            =   2790
         TabIndex        =   26
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�����(&S)"
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   25
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������Ч(&F)"
         Height          =   180
         Index           =   7
         Left            =   2490
         TabIndex        =   24
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&T)"
         Height          =   180
         Index           =   5
         Left            =   330
         TabIndex        =   22
         Top             =   4365
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(ƴ��)"
         Height          =   180
         Left            =   2385
         TabIndex        =   21
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(���)"
         Height          =   180
         Left            =   4620
         TabIndex        =   20
         Top             =   1035
         Width           =   540
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4620
      Index           =   1
      Left            =   180
      TabIndex        =   30
      Top             =   585
      Visible         =   0   'False
      Width           =   5280
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         TabIndex        =   41
         ToolTipText     =   "������һ����F3��س�����λ�����F4"
         Top             =   75
         Width           =   1905
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ����"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3690
         TabIndex        =   34
         Top             =   4230
         Width           =   1575
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ����ͬ����Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   33
         Top             =   4230
         Width           =   1620
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ��Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   32
         Top             =   4230
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.CheckBox chkSelectAll 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   675
      End
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   3645
         Left            =   0
         TabIndex        =   35
         Top             =   495
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   6429
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   960
         TabIndex        =   42
         Top             =   135
         Width           =   360
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   4650
      Index           =   2
      Left            =   165
      TabIndex        =   36
      Top             =   585
      Visible         =   0   'False
      Width           =   5280
      Begin VSFlex8Ctl.VSFlexGrid vs���� 
         Height          =   4275
         Left            =   30
         TabIndex        =   40
         Top             =   15
         Width           =   5250
         _cx             =   9260
         _cy             =   7541
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmIllItemEdit.frx":0640
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ����"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   3690
         TabIndex        =   39
         Top             =   4395
         Width           =   1575
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ����ͬ����Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1860
         TabIndex        =   38
         Top             =   4395
         Width           =   1620
      End
      Begin VB.OptionButton opt����Ӧ�÷�Χ 
         Appearance      =   0  'Flat
         Caption         =   "Ӧ���ڵ�ǰ��Ŀ"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   37
         Top             =   4380
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5145
      Left            =   75
      TabIndex        =   17
      Top             =   180
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   9075
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ŀ(&1)"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ӧ����(&2)"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��Ӧ����(&3)"
            Key             =   "K3"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIllItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String             '��ǰ�༭����ĿID
Dim mstr����ID As String     '��ǰ�༭�ķ�����ĿID
Dim mstr������� As String
Dim mstrԭ����   As String       '����ü�����ԭʼ���룬�����ж��Ƿ�Ҫ���µõ����
Dim mlng���     As Long

Dim mblnChange As Boolean  '���޸�
Private mlng���볤�� As Long

Private Const mconInt���� As Integer = 0
Private Const mconInt���� As Integer = 1
Private Const mconInt���� As Integer = 2
Private Const mconIntƴ���� As Integer = 3
Private Const mconInt˵�� As Integer = 4
Private Const mconInt���� As Integer = 5
Private Const mconInt����� As Integer = 6
Private Const mconIntͳ���� As Integer = 7

Private Sub IniDept()
    Dim rsTemp As ADODB.Recordset
    Dim lngId As Long
    
    lngId = Val(mstrID)
    
    On Error GoTo ErrHandle
    gstrSQL = " Select A.���� || '-' || A.���� ����, A.ID, Nvl(B.����id, 0) ����id , A.����" & _
            " From ���ű� A, (Select ����id From ����������� Where ����id = [1]) B " & _
            " Where A.ID = B.����id(+) And " & _
            " A.ID In (Select ����id From ��������˵�� Where �������� In ('�ٴ�', '���', '����', '����', '����', 'Ӫ��')) And " & _
            " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
            " Order By A.���� || '-' || A.���� "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ٴ���ҽ���ಿ��", lngId)
    
    With rsTemp
        If .EOF Then
            MsgBox "û�������ٴ���ҽ���ಿ�ţ������Ź���", vbInformation, gstrSysName
            Exit Sub
        End If
        Me.Lvw����.ListItems.Clear
        Do While Not .EOF
            Me.Lvw����.ListItems.Add , "_" & !ID, !����, 1, 1
            Me.Lvw����.ListItems("_" & !ID).Tag = NVL(!����)
            If !����ID > 0 Then
                Me.Lvw����.ListItems("_" & !ID).Checked = True
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Init��������(Optional ByVal str�������� As String)
    '----------------------------------------------------------------------------
    '����:��ʼ��������������
    '����:str��������-ָ��ָ������������
    '����:
    '����:���˺�
    '����:2007/08/14
    '----------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select ���� From �������� order by ����"
    
    Err = 0: On Error GoTo ErrHand:
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "��ȡ��������")
    With rsTemp
        Me.cbo��������.AddItem ""
        If str�������� = "" Then
            cbo��������.ListIndex = cbo��������.NewIndex
        End If
        Do While Not .EOF
            cbo��������.AddItem !����
            If str�������� = NVL(!����) Then
                cbo��������.ListIndex = cbo��������.NewIndex
            End If
            .MoveNext
        Loop
        If str�������� <> "" And cbo��������.ListIndex < 0 Then
            cbo��������.AddItem str��������
            cbo��������.ListIndex = cbo��������.NewIndex
        End If
        If cbo��������.ListIndex < 0 Then cbo��������.ListIndex = 0
        cbo��������.Tag = cbo��������.Text
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub TabShow(ByVal i As Integer)
    tabMain.Tabs(i).Selected = True
    tabMain_Click
End Sub

Private Sub cbo��������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkSelectAll_Click()
    Dim n As Integer
    Dim BlnSelect As Boolean
    
    If chkSelectAll.value = 2 Then Exit Sub
    
    BlnSelect = (chkSelectAll.value = 1)
    
    With Lvw����
        For n = 1 To .ListItems.Count
            .ListItems(n).Checked = BlnSelect
        Next
    End With
End Sub


Private Sub Chk����_Click()
    mblnChange = True
End Sub

Private Sub Chk����_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub Chk����_LostFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub cmb�Ա�_Click()
    mblnChange = True
End Sub

Private Sub cmb��Ч_Click()
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    '���ڿ����޸Ķ���ڵ㣬����ֻ��ǿ��ˢ��
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(mconInt����).Text = ""
    txtEdit(mconInt����).Text = ""
    txtEdit(mconInt����).Text = ""
    txtEdit(mconIntƴ����).Text = ""
    txtEdit(mconInt˵��).Text = ""
    cmb��Ч.ListIndex = 0
    cmb�Ա�.ListIndex = 0
    
    Call TabShow(1)
    txtEdit(mconInt����).SetFocus
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
'����:��������������������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    
    If mstr����ID = "1" Then
        If MsgBox("�ڸ���Ŀ�����Ӽ���������ϵͳ�Դ�����ļ�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End If
    For i = 0 To 7
        If i <> 5 Then
            If zlCommFun.StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength, , txtEdit(i).Tag) = False Then
                Call TabShow(1)
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    
'    If zlCommFun.ActualLen(txtEdit(mconIntͳ����).Text) > txtEdit(mconIntͳ����).MaxLength Then
'        MsgBox "ͳ���벻�ܳ���" & txtEdit(mconIntͳ����).MaxLength & "���ַ���" & txtEdit(mconIntͳ����).MaxLength / 2 & "������,����!"
'        Call TabShow(1)
'        txtEdit(mconIntͳ����).SetFocus
'        zlControl.TxtSelAll txtEdit(mconIntͳ����)
'        Exit Function
'    End If
    
    txtEdit(mconInt����).Text = UCase(Trim(txtEdit(mconInt����).Text))
    txtEdit(mconInt����).Text = UCase(Trim(txtEdit(mconInt����).Text))
    txtEdit(mconInt����).Text = Trim(txtEdit(mconInt����).Text)
    
    If Len(txtEdit(mconInt����).Text) = 0 Then
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        Call TabShow(1)
        txtEdit(mconInt����).SetFocus
        Exit Function
    End If
    
    If InStr(txtEdit(mconInt����).Text, "+") > 0 And InStr(txtEdit(mconInt����).Text, "*") = 0 Then
        MsgBox "���ڸ��봦�����Ǻű��롣", vbExclamation, gstrSysName
        Call TabShow(1)
        zlControl.TxtSelAll txtEdit(mconInt����)
        txtEdit(mconInt����).SetFocus
        Exit Function
    End If
    
    If Len(Trim(txtEdit(mconInt����).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        Call TabShow(1)
        txtEdit(mconInt����).Text = ""
        txtEdit(mconInt����).SetFocus
        Exit Function
    End If
    
    'ר����Լ���������ص㣬��Ӳ�ԵĹ涨
    If mstr������� = "D" Or mstr������� = "Y" Or mstr������� = "M" Then
        '���������ĸ
        If ������("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. ", txtEdit(mconInt����).Text) = False Then
            Call TabShow(1)
            txtEdit(mconInt����).SetFocus
            zlControl.TxtSelAll txtEdit(mconInt����)
            Exit Function
        End If
        If ������("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. ", txtEdit(mconInt����).Text) = False Then
            Call TabShow(1)
            txtEdit(mconInt����).SetFocus
            zlControl.TxtSelAll txtEdit(mconInt����)
            Exit Function
        End If
        
        '�������ĸ
        Select Case mstr�������
            Case "D"
                If InStr("ABCDEFGHIJKLMNOPQRSTUZ", Left(txtEdit(mconInt����).Text, 1)) = 0 Then
                    MsgBox "���������ĸ����", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
            Case "Y"
                If InStr("VWXY", Left(txtEdit(mconInt����).Text, 1)) = 0 Then
                    MsgBox "�ⲿԭ����������ĸֻΪVWXY����֮һ��", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
            Case "M"
                If "M" <> Left(txtEdit(mconInt����).Text, 1) Then
                    MsgBox "��̬ѧ���������ĸֻ��M��", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
        End Select
        '��������ĸ
        Select Case mstr�������
            Case "D", "Y"
                If ������("0123456789+-. ", Mid(txtEdit(mconInt����).Text, 2), True) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
                If ������("0123456789+-*/. ", Mid(txtEdit(mconInt����).Text, 2), True) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
            Case "M"
                If ������("0123456789/ ", Mid(txtEdit(mconInt����).Text, 2)) = False Then
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
                i = InStr(txtEdit(mconInt����), "/")
                If i = 0 Then
                    MsgBox "�����������̬���롣", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
                strTemp = Mid(txtEdit(mconInt����).Text, i + 1)
                If Len(strTemp) <> 1 Then
                    MsgBox "������̬�������", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
                If InStr("012369", strTemp) = 0 Then
                    MsgBox "������̬�������", vbInformation, gstrSysName
                    Call TabShow(1)
                    txtEdit(mconInt����).SetFocus
                    zlControl.TxtSelAll txtEdit(mconInt����)
                    Exit Function
                End If
                
        End Select
    End If
    
'    '������ķ�Χ
'    Dim arrValues As Variant, blnMatch As Boolean
'    Dim lngCount As Long, lngPos As Long
'    Dim strBegin As String, strEnd As String
'
'    If Trim(txtEdit(mconInt����).Tag) = "" Then
'        '�������û�����ñ��뷶Χ����ô���Բ�����
'        blnMatch = True
'    Else
'        arrValues = Split(txtEdit(mconInt����).Tag, ",")
'        For lngCount = LBound(arrValues) To UBound(arrValues)
'            lngPos = InStr(arrValues(lngCount), "-")
'            If lngPos = 0 Then
'                'û�ж̺��ߣ�ֻ���Ըô���ͷ
'                strBegin = Trim(arrValues(lngCount))
'                If txtEdit(mconInt����).Text Like (strBegin & "*") Then
'                    blnMatch = True
'                    Exit For
'                End If
'            Else
'                'ȡֵ��Χ
'                strBegin = Trim(Mid(arrValues(lngCount), 1, lngPos - 1))
'                strEnd = Trim(Mid(arrValues(lngCount), lngPos + 1))
'
'                strTemp = Mid(txtEdit(mconInt����).Text, 1, Len(strBegin))
'                If strTemp >= strBegin And strTemp <= strEnd Then
'                    blnMatch = True
'                    Exit For
'                End If
'            End If
'        Next
'    End If
'
'    If blnMatch = False Then
'        MsgBox "������󣬵�ǰ�����µı��뷶Χ�ǣ�" & vbCrLf & vbCrLf & txtEdit(mconInt����).Tag, vbInformation, gstrSysName
'        Call TabShow(1)
'        txtEdit(mconInt����).SetFocus
'        zlControl.TxtSelAll txtEdit(mconInt����)
'        Exit Function
'    End If
    
    '���
    With vs����
        If vs����.Tag <> "DEL" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("����"))) <> "" And Val(.Cell(flexcpData, i, .ColIndex("����"))) = 0 Then
                    MsgBox "������Ĳ��ֲ���ȷ,����������!", vbInformation + vbDefaultButton1, gstrSysName
                    TabShow (3)
                    .Row = i
                    If .RowIsVisible(i) = False Then
                        .TopRow = i
                    End If
                    .SetFocus
                    Exit Function
                End If
            Next
        End If
    End With
    IsValid = True
End Function

Private Function ������(ByVal strPartten As String, ByVal strCheck As String, Optional blnX As Boolean = False) As Boolean
'����:blnX �Ƿ�֧�ֵ�4����ĸΪX����G01.X55*
    Dim i As Long
    Dim blnIsValid As Boolean
    
    blnIsValid = True
    
    For i = 1 To Len(strCheck)
        If InStr(strPartten, Mid(strCheck, i, 1)) = 0 Then
            If Not (blnX = True And i = 4 And Mid(strCheck, 4, 1) = "X") Then
'                '�ų���4����ĸΪX�����
'                MsgBox "���ڱ���������Ϸ���ĸ��", vbInformation, gstrSysName
'                Exit Function
                blnIsValid = False
                Exit For
            End If
        End If
    Next
    
    If Not blnIsValid Then
        If MsgBox("�ڱ�������к��зǷ���ĸ��ȷ�ϱ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) <> vbYes Then
            ������ = False
            Exit Function
        End If
    End If
    
    
    ������ = True
    
    
End Function

Private Function Save��Ŀ() As Boolean
'����:����༭�����ݵ�����������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim lngId As Long
    Dim lng��� As Long
    Dim lng����id As Long
    Dim bln���� As Boolean
    Dim str�������� As String
    
    Dim lst As ListItem
    Dim strDeptId As String
    Dim n As Integer
    Dim intӦ�÷�Χ As Integer
    Dim int����Ӧ�÷�Χ As Integer
    Dim str���� As String '����|����id,����1|����id1 ....
    On Error GoTo ErrHandle
    
    
    '�����жϱ�������
    lng��� = mlng���
    Err = 0: On Error GoTo ErrHand:
    
    If mstrԭ���� <> txtEdit(mconInt����).Text Then
    
        gstrSQL = "select max(���) as ����,count(���) as ���� from ��������Ŀ¼ " & _
                   " Where ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, txtEdit(mconInt����).Text)
        
        If rsTemp("����") > 0 Then
            If MsgBox("����" & txtEdit(mconInt����).Text & "�Ѵ��ڣ��Ƿ�Ҫ������һ����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
            
            lng��� = rsTemp("����") + 1
        End If
        rsTemp.Close
    End If
    
    For n = 1 To Lvw����.ListItems.Count
        If Lvw����.ListItems(n).Checked = True Then
            strDeptId = IIF(strDeptId = "", Mid(Lvw����.ListItems(n).Key, 2), strDeptId & "," & Mid(Lvw����.ListItems(n).Key, 2))
        End If
    Next
    
    str���� = ""
    With vs����
        For n = 1 To .Rows - 1
            If .RowData(n) <> 0 Then
                If Val(.Cell(flexcpData, n, .ColIndex("����"))) <> 0 Then
                    str���� = str���� & "," & .RowData(n) & "|" & Val(.Cell(flexcpData, n, .ColIndex("����")))
                End If
            End If
        Next
    End With
    If str���� <> "" Then str���� = Mid(str����, 2)
    
    For n = 0 To optӦ�÷�Χ.UBound
        If optӦ�÷�Χ(n).value = True Then
            intӦ�÷�Χ = n
            Exit For
        End If
    Next
    
    For n = 0 To opt����Ӧ�÷�Χ.UBound
        If opt����Ӧ�÷�Χ(n).value = True Then
            int����Ӧ�÷�Χ = n
            Exit For
        End If
    Next
    
    lng����id = Val(mstr����ID)
    If cbo��������.Enabled And cbo��������.ListIndex > 0 Then
        str�������� = cbo��������.Text
    Else
        str�������� = ""
    End If
    If mstrID = "" Then       '����һ����¼
        
        '����û�з���ı��뻹Ӧ�õõ������ID
        If cmd����.Visible = False And mstr����ID = "" Then
            gstrSQL = "select ID from ����������� where ���=[1] and rownum<2"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr�������)
            
            If rsTemp.EOF Then
                
                lng����id = zlDatabase.GetNextId("�����������")
                
                gstrSQL = "ZL_�����������_INSERT(" & lng����id & ",NULL,1,'" & Mid(frmIllManage.cmbType.Text, 3) & "','','" & mstr������� & "',1)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                '�õ�����ID
                lng����id = rsTemp("ID")
            End If
        End If
                
        lngId = zlDatabase.GetNextId("��������Ŀ¼")
        
        'Zl_��������Ŀ¼_Insert
        gstrSQL = "Zl_��������Ŀ¼_Insert("
        '  Id_In       In ��������Ŀ¼.ID%Type,
        gstrSQL = gstrSQL & "" & lngId & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ���_In     In ��������Ŀ¼.���%Type,
        gstrSQL = gstrSQL & "" & lng��� & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ͳ����_In   In ��������Ŀ¼.ͳ����%Type,
        gstrSQL = gstrSQL & "" & IIF(Trim(txtEdit(mconIntͳ����).Text) = "", "NULL", "'" & Trim(txtEdit(mconIntͳ����).Text) & "'") & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconIntƴ����).Text & "',"
        '  ˵��_In     In ��������Ŀ¼.˵��%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt˵��).Text & "',"
        '  �Ա�����_In In ��������Ŀ¼.�Ա�����%Type,
        gstrSQL = gstrSQL & "'" & cmb�Ա�.Text & "',"
        '  ��Ч����_In In ��������Ŀ¼.��Ч����%Type,
        gstrSQL = gstrSQL & "'" & cmb��Ч.Text & "',"
        '  ���_In     In ��������Ŀ¼.���%Type,
        gstrSQL = gstrSQL & "'" & mstr������� & "',"
        '  ��������_In In ��������Ŀ¼.��������%Type,
        gstrSQL = gstrSQL & "" & IIF(str�������� = "", "NULL", "'" & str�������� & "'") & ","
        '  ����id_In   In ��������Ŀ¼.����id%Type,
        gstrSQL = gstrSQL & "" & lng����id & ","
        '  ����_In     In ��������Ŀ¼.����%Type := Null,
        gstrSQL = gstrSQL & "'" & Chk����.value & "',"
        '  �����_In   In ��������Ŀ¼.�����%Type := Null,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt�����).Text & "',"
        '  ����_In     In Varchar2, --����ID��:����ID1,����ID2,����ID3...
        gstrSQL = gstrSQL & "'" & strDeptId & "',"
        '  Ӧ��_In     In Number := 0 --Ӧ�÷�Χ:0-Ӧ���ڵ�ǰ��Ŀ;1-Ӧ����ͬ����Ŀ;2-Ӧ���ڵ�ǰ����
        gstrSQL = gstrSQL & "" & intӦ�÷�Χ & ")"
    Else    '�޸�
        ' Zl_��������Ŀ¼_Update
        gstrSQL = "Zl_��������Ŀ¼_Update("
        '  Id_In       In ��������Ŀ¼.ID%Type,
        gstrSQL = gstrSQL & "" & mstrID & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ���_In     In ��������Ŀ¼.���%Type,
        gstrSQL = gstrSQL & "" & lng��� & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ͳ����_In   In ��������Ŀ¼.ͳ����%Type,
        gstrSQL = gstrSQL & "" & IIF(Trim(txtEdit(mconIntͳ����).Text) = "", "NULL", "'" & Trim(txtEdit(mconIntͳ����).Text) & "'") & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt����).Text & "',"
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconIntƴ����).Text & "',"
        '  ˵��_In     In ��������Ŀ¼.˵��%Type,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt˵��).Text & "',"
        '  �Ա�����_In In ��������Ŀ¼.�Ա�����%Type,
        gstrSQL = gstrSQL & "'" & cmb�Ա�.Text & "',"
        '  ��Ч����_In In ��������Ŀ¼.��Ч����%Type,
        gstrSQL = gstrSQL & "'" & cmb��Ч.Text & "',"
        '  ���_In     In ��������Ŀ¼.���%Type,
        gstrSQL = gstrSQL & "'" & mstr������� & "',"
        '  ��������_In In ��������Ŀ¼.��������%Type,
        gstrSQL = gstrSQL & "" & IIF(str�������� = "", "NULL", "'" & str�������� & "'") & ","
        '  ����id_In   In ��������Ŀ¼.����id%Type,
        gstrSQL = gstrSQL & "" & lng����id & ","
        '  ����_In     In ��������Ŀ¼.����%Type,
        gstrSQL = gstrSQL & "'" & Chk����.value & "',"
        '  �����_In   In ��������Ŀ¼.�����%Type := Null,
        gstrSQL = gstrSQL & "'" & txtEdit(mconInt�����).Text & "',"
        '  ����_In     In Varchar2, --����ID��:����ID1,����ID2,����ID3...
        gstrSQL = gstrSQL & "'" & strDeptId & "',"
        '  Ӧ��_In     In Number := 0 --Ӧ�÷�Χ:0-Ӧ���ڵ�ǰ��Ŀ;1-Ӧ����ͬ����Ŀ;2-Ӧ���ڵ�ǰ����
        gstrSQL = gstrSQL & "" & intӦ�÷�Χ & ")"
    End If
    
    gcnOracle.BeginTrans
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ' Zl_�������ֶ�Ӧ_Update
    gstrSQL = "Zl_�������ֶ�Ӧ_Update("
    
    If mstrID = "" Then
        '  Id_In     In ��������Ŀ¼.ID%Type,
        gstrSQL = gstrSQL & "" & lngId & ","
    Else
        '  Id_In     In ��������Ŀ¼.ID%Type,
        gstrSQL = gstrSQL & "" & mstrID & ","
    End If
    '  ���_In   In ��������Ŀ¼.���%Type,
    gstrSQL = gstrSQL & "'" & mstr������� & "',"

    '  ����id_In In ��������Ŀ¼.����id%Type,
    gstrSQL = gstrSQL & "" & lng����id & ","

    '  ����_In   In Varchar2 := Null, --����id��,����1|����id1,����2|����id2.....
    gstrSQL = gstrSQL & "'" & str���� & "',"
    
    '  Ӧ��_In   In Number := 0 --���ֵ�Ӧ�÷�Χ:0-Ӧ���ڵ�ǰ��Ŀ;1-Ӧ����ͬ����Ŀ;2-Ӧ���ڵ�ǰ����
    
    gstrSQL = gstrSQL & "" & int����Ӧ�÷�Χ & ")"
    
    If vs����.Tag <> "DEL" Then
        '��Ҫ�ǲ�������ʱ,��Ӧ������Щ����
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    gcnOracle.CommitTrans
    
     Err = 0: On Error GoTo ErrHandle
    '������������ݽ��и���
    '�����жϸ����Ƿ�Ӧ���ڵ�ǰ�б�����ʾ
    With frmIllManage
        If .tvwMain_S.SelectedItem Is Nothing Then
            bln���� = True
        Else
            If .mnuViewAll.Checked = True Then
                '��ʾ�����µ�����
                bln���� = IsChild(mstr����ID, .tvwMain_S.SelectedItem)
            Else
                bln���� = (mstr����ID = Mid(.tvwMain_S.SelectedItem.Key, 2))
            End If
        End If
    End With
    With frmIllManage.lvwMain
        If mstrID = "" Then
            If bln���� = True Then
                Set lst = .ListItems.Add(, "K" & lngId, " ", "Item", "Item")
                Call ShowItem(lst)
                lst.Selected = True
                DoEvents
                lst.EnsureVisible
            End If
        Else
            If bln���� = True Then
                Call ShowItem(.SelectedItem)
            Else
                'ɾ��
                Dim intIndex As Long
                intIndex = .SelectedItem.Index
                .ListItems.Remove .SelectedItem.Key
                If .ListItems.Count > 0 Then
                    intIndex = IIF(.ListItems.Count > intIndex, intIndex, .ListItems.Count)
                    .ListItems(intIndex).Selected = True
                    .ListItems(intIndex).EnsureVisible
                End If
            End If
        End If
    End With
    Call frmIllManage.SetMenu
    
    '�´β����ٴ����ݿ�����ȡ��
    mstr����ID = lng����id
    Save��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Function IsChild(ByVal strKey As String, ByVal nod As Node) As Boolean
'�ж�ĳ���ؼ����Ƿ�����һ���ڵ�����ӽڵ�
    Dim nodTemp As Node
    
    If strKey = Mid(nod.Key, 2) Then
        IsChild = True
        Exit Function
    End If
    Set nodTemp = nod.Child
    Do Until nodTemp Is Nothing
        If IsChild(strKey, nodTemp) = True Then
            IsChild = True
            Exit Function
        End If
        Set nodTemp = nodTemp.Next
    Loop
End Function

Private Sub ShowItem(lst As ListItem)
'������ʾĳһ��,����ˢ��
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol  As Long
    Dim varValue As Variant
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "" & _
    "   Select A.ID,A.����,����,A.����,A.���� As ƴ����,A.�����,A.˵��,A.��������,A.ͳ����,A.�Ա�����," & _
    "       A.��Ч���� as ������Ч,decode(A.����,'1','¼��') ������Ϣ,to_char(A.����ʱ��,'yyyy-mm-dd') as  ����ʱ��, " & _
    "          to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��" & _
    "   From ��������Ŀ¼ A  Where A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(lst.Key, 2)))
                
    '����ListView�����������ݿ�ȡ��
    lst.Text = rsTemp("����")
    With frmIllManage.lvwMain
        For lngCol = 2 To .ColumnHeaders.Count
            varValue = rsTemp(.ColumnHeaders(lngCol).Text).value
            lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function �����༭(ByVal bln���� As Boolean, ByVal str������Ŀ As String, ByVal str������ĿID As String, _
    ByVal str������� As String, Optional ByVal strID As String = "") As Boolean
'����:��������õı����������ڽ���ͨѶ�ĳ���
'����:str������Ŀ     ���������������
'     str������ĿID   �����������ID
'     str�������     ������������
'     strID           ���������ĵ�ID
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim intSys As Integer
    Dim j As Integer
    
    mstr������� = str�������
    
    cmb�Ա�.AddItem ""
    cmb�Ա�.AddItem "��"
    cmb�Ա�.AddItem "Ů"
    
    '����26069 By lesfeng 2009-11-16 ������������ ͬʱע��Ȩ�޷���
    On Error GoTo ErrHandle
    intSys = 0
    j = 0
    gstrSQL = "SELECT ����� FROM zlsystems WHERE ���=300" ' Int((glngSys)
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!�����) Then
            intSys = rsTmp!�����
        End If
    End If
    rsTmp.Close
    If Int(intSys / 100) = Int(glngSys / 100) Then
        gstrSQL = "SELECT ����,����,ȱʡ��־ From ���ƽ�� "
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If Not rsTmp.EOF Then
            While Not rsTmp.EOF
                cmb��Ч.AddItem IIF(IsNull(rsTmp!����), "", rsTmp!����)
                If IsNull(rsTmp!����) Then cmb��Ч.ListIndex = cmb��Ч.NewIndex
    '            If IIF(IsNull(rsTmp!ȱʡ��־), 0, rsTmp!ȱʡ��־) = 1 Then
    '                cmb��Ч.ListIndex = cmb��Ч.NewIndex
    '            End If
                j = 1
                rsTmp.MoveNext
            Wend
        End If
        rsTmp.Close
    End If
    
    If j = 0 Then
        cmb��Ч.AddItem ""
        cmb��Ч.AddItem "����"
        cmb��Ч.AddItem "��ת"
        cmb��Ч.AddItem "δ��"
        cmb��Ч.AddItem "����"
        cmb��Ч.AddItem "��Ч"
        cmb��Ч.AddItem "����"
    End If
    
    If bln���� = False Then
        'û����������ѡ��
        lblEdit(5).Visible = False
        txtEdit(mconInt����).Visible = False
        cmd����.Visible = False

'        Frame1.Top = Frame1.Top - 450
'        cmdOk.Top = cmdOk.Top - 450
'        cmdCancel.Top = cmdCancel.Top - 450
'        Height = Height - 450
    End If
    
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If strID <> "" Then
        
        gstrSQL = "select A.ID,A.����,A.���,A.����,A.����,A.���� As ƴ����,A.�����,A.�Ա�����,A.��Ч����,A.����,A.��������,A.ͳ����,A.˵�� " & _
                ",B.��� as ͳ�����,B.ID as ����ID,B.���� as ��������,B.���뷶Χ" & _
                " from ��������Ŀ¼ A,����������� B " & _
                " where B.ID(+)=A.����ID and A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        
        txtEdit(mconInt����).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        mstrԭ���� = txtEdit(mconInt����).Text
        mlng��� = NVL(rsTemp("���"), 0)
        txtEdit(mconInt����).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(mconInt����).Text = Trim(rsTemp("����"))
        txtEdit(mconIntƴ����).Text = IIF(IsNull(rsTemp("ƴ����")), "", rsTemp("ƴ����"))
        txtEdit(mconInt�����).Text = IIF(IsNull(rsTemp("�����")), "", rsTemp("�����"))
        txtEdit(mconInt˵��).Text = IIF(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
        txtEdit(mconIntͳ����).Text = NVL(rsTemp!ͳ����)
        
        If Not IsNull(rsTemp("�Ա�����")) Then
            cmb�Ա�.Text = rsTemp("�Ա�����")
        End If
        If Not IsNull(rsTemp("��Ч����")) Then
            cmb��Ч.Text = rsTemp("��Ч����")
        End If
        mstr����ID = IIF(IsNull(rsTemp("����ID")), "", rsTemp("����ID"))
        
        If IsNull(rsTemp("��������")) Then
            txtEdit(mconInt����).Text = "��"
        Else
            txtEdit(mconInt����).Text = "��" & rsTemp("ͳ�����") & "��" & Trim(rsTemp("��������"))
        End If
        txtEdit(mconInt����).Tag = IIF(IsNull(rsTemp("���뷶Χ")), "", rsTemp("���뷶Χ"))
        Chk����.value = IIF(NVL(rsTemp("����"), 0) = 0, 0, 1)
        Call Init��������(NVL(rsTemp!��������))
    Else
        If mstr������� = "M" Then txtEdit(mconInt����).Text = "M"
        
        mstr����ID = str������ĿID
        txtEdit(mconInt����).Text = str������Ŀ
        
        If bln���� = True Then
            gstrSQL = "select A.���뷶Χ " & _
                    " from ����������� A " & _
                    " where A.ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(str������ĿID))
            
            txtEdit(mconInt����).Tag = IIF(IsNull(rsTemp("���뷶Χ")), "", rsTemp("���뷶Χ"))
        End If
        mlng��� = 1
        Call Init��������("")
    End If
    Call GetDefineSize
    Call IniDept
    
    '���˺�:2007/08/15���뼲���벡�ֵĶ��չ�ϵ
    Call Init����(Val(strID))
    
    If UCase(mstr�������) <> "S" Then
        '���˺�:2007/08/14:ֻ��"S"���͵Ĳ��ܱ༭��������
        cbo��������.Enabled = False
    End If
    mblnChange = False
    frmIllItemEdit.Show vbModal
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd����_Click()
    Dim blnRe As Boolean
    Dim str���� As String
    Dim str����ID As String
    Dim str���뷶Χ As String
    
    str����ID = mstr����ID
    str���� = txtEdit(mconInt����).Text
    blnRe = frmClassSel.ShowTree(str����ID, str����, str���뷶Χ, mstr�������, "", False)
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        mstr����ID = str����ID
        txtEdit(mconInt����).Text = str����
        txtEdit(mconInt����).Tag = str���뷶Χ
        mblnChange = True
    End If
End Sub

Private Sub Form_Activate()
    Call tabMain_Click
    txtEdit(mconInt����).SetFocus
    Lvw����.MultiSelect = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    For i = fra.LBound To fra.UBound
        fra(i).Visible = False
    Next
    i = tabMain.SelectedItem.Index - 1
    fra(i).Visible = True
End Sub


Private Sub txtEdit_Change(Index As Integer)
    If Index = 2 Then
        txtEdit(mconIntƴ����).Text = zlStr.GetCodeByVB(txtEdit(mconInt����).Text)
        txtEdit(mconInt�����).Text = zlStr.GetCodeByORCL(txtEdit(mconInt����).Text, False, mlng���볤��)
    ElseIf Index = 3 Then
        txtEdit(mconIntƴ����).Text = UCase(txtEdit(mconIntƴ����).Text)
    ElseIf Index = 6 Then
        txtEdit(mconInt�����).Text = UCase(txtEdit(mconInt�����).Text)
    End If
    mblnChange = True
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select ����� From ��������Ŀ¼ Where Rownum = 0 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("�����").DefinedSize
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 2 Or Index = 4 Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
'���ڶ����ı�����ò�Ҫ�ӿո�
    If Index = 4 And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If Index = 0 Or Index = 1 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        'ֻ��ȡ��Щ��ĸ
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789/*-+. " & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub Init����(ByVal lng����ID As Long)
    '-------------------------------------------------------------------
    '����:��ʼ����
    '����:lng����ID-����ID(0��ʾ,����)
    '����:���˺�
    '����:2007/08/15
    '-------------------------------------------------------------------
    Dim rs���� As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim i As Long
    Err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "Select TABLE_NAME from table_privileges where table_name='�������'"
    zlDatabase.OpenRecordset rs����, gstrSQL, Me.Caption
    If rs����.RecordCount = 0 Then
        '��ʾû�м�¼,�������ö�Ӧ��ϵ
        tabMain.Tabs.Remove "K3"
        vs����.Tag = "DEL"
        Exit Sub
    End If
    vs����.Tag = ""
    gstrSQL = "" & _
        "   Select b.����,b.����ID,c.����||'-'||c.���� as ���� " & _
        "   From �������ֶ�Ӧ b,���ղ��� c " & _
        "   where  b.����id=c.id and b.����id=[1]"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    
    gstrSQL = "Select ���,���� From ������� where ҽԺ���� is not null"
    Call zlDatabase.OpenRecordset(rs����, gstrSQL, Me.Caption)
    With vs����
        If rs����.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
                .RowData(1) = 0
            Next
            .Editable = flexEDNone
            Exit Sub
        End If
        .Rows = rs����.RecordCount + 1
        i = 1
        Do While Not rs����.EOF
            .RowData(i) = Val(NVL(rs����!���))
            .TextMatrix(i, .ColIndex("����")) = NVL(rs����!����)
            rs����.Filter = "����=" & Val(NVL(rs����!���))
            If rs����.EOF = True Then
                .TextMatrix(i, .ColIndex("����")) = ""
                .Cell(flexcpData, i, .ColIndex("����")) = ""
            Else
                .TextMatrix(i, .ColIndex("����")) = NVL(rs����!����)
                .Cell(flexcpData, i, .ColIndex("����")) = NVL(rs����!����ID)
            End If
            i = i + 1
             rs����.MoveNext
        Loop
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("����")) = "..."
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub
 

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart >= Lvw����.ListItems.Count Then lngStart = 1
    
        For i = lngStart To Lvw����.ListItems.Count
            If Lvw����.ListItems(i).Text Like "*" & txtLocate.Text & "*" Or Lvw����.ListItems(i).Tag Like "*" & UCase(txtLocate.Text) & "*" Then
                Call Lvw����.ListItems(i).EnsureVisible
                Lvw����.ListItems(i).Selected = True
                lblLocate.Tag = i
                Lvw����.SetFocus
                Exit For
            End If
        Next
    End If
End Sub


Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs����
        Select Case Col
        Case .ColIndex("����")
             .ColComboList(.ColIndex("����")) = "..."
        End Select
    End With
End Sub

Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs����
        Select Case Col
        Case .ColIndex("����")
             Cancel = True
        End Select
    End With

End Sub

Private Sub vs����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case vs����.ColIndex("����")
        'ѡ����
        Call Select����(vs����.RowData(Row), "")
    Case Else
    End Select
End Sub
Private Function Select����(ByVal lng���� As Long, ByVal strKey As String)
    '---------------------------------------------------------------------------------
    '����:ѡ��ָ������Ĳ���
    '����:lng����-����
    '����:ѡ��ɹ�,����ture,���򷵻�False
    '����:���˺�
    '����:2007/08/15
    '---------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strLeft As String
    Dim blnCancel As Boolean
    
    Dim vRect As RECT

    strLeft = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    Err = 0: On Error GoTo ErrHand:
    
    Dim sngX As Single, sngY As Single
    Call CalcPosition(sngX, sngY, vs����)
     
     If strKey <> "" Then
        strKey = strLeft & strKey & "%"
        gstrSQL = "" & _
          "   Select Id, ����, ����, ����, decode('0','��ͨ��','1','���Բ�','2','���ֲ�','') As ���, ����ⶥ��, �ⶥ�߽�� " & _
          "    From ���ղ��� " & _
          "    Where ���� = [1] And (���� Like [2] Or ���� Like [2] Or ���� Like [3]) " & _
          "    Order by ����"
    Else
        gstrSQL = "" & _
          "   Select Id, ����, ����, ����, decode('0','��ͨ��','1','���Բ�','2','���ֲ�','') As ���, ����ⶥ��, �ⶥ�߽�� " & _
          "    From ���ղ��� " & _
          "    Where ���� = [1]" & _
          "    Order by ����"
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "���ղ���ѡ��", False, "", "", False, False, True, sngX, sngY - vs����.CellHeight, vs����.CellHeight, blnCancel, False, False, lng����, strKey, CStr(UCase(strKey)))
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "������ָ���Ĳ���,����!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs����
        .TextMatrix(.Row, .ColIndex("����")) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        If strKey <> "" Then
            .EditText = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        End If
        .Cell(flexcpData, .Row, .ColIndex("����")) = NVL(rsTemp!ID)
    End With
    Select���� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs����_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vs����_DblClick()
    With vs����
        Select Case .Col
        Case .ColIndex("����")
             .ColComboList(.ColIndex("����")) = ""
        End Select
    End With
End Sub

Private Sub vs����_EnterCell()
    vs����.ColComboList(vs����.ColIndex("����")) = "..."
End Sub

Private Sub vs����_GotFocus()
    With vs����
        .BackColorSel = &H8000000D
        .GridColor = &H0&
        .GridColorFixed = &H0&
    End With
End Sub

Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If vs����.Col = vs����.ColIndex("����") And KeyCode <> vbKeyReturn Then
       vs����.ColComboList(vs����.ColIndex("����")) = ""
    End If
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
        Dim strKey As String
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        With vs����
            Select Case Col
            Case .ColIndex("����")
                .Cell(flexcpData, Row, Col) = ""
                If KeyCode <> vbKeyReturn Then Exit Sub
                
                strKey = Trim(vs����.EditText)
                strKey = Replace(strKey, Chr(vbKeyReturn), "")
                strKey = Replace(strKey, Chr(10), "")
                
                If strKey = "" Then Exit Sub
                
                If Select����(.RowData(.Row), strKey) = False Then
                    'ѡ��ʧ��
                    
                End If
                
                .Col = 1
                .SetFocus
            End Select
        End With
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vs����_LostFocus()
    With vs����
        .BackColorSel = &H8000000C
        .GridColor = &H808080
        .GridColorFixed = &H808080
    End With
End Sub
Private Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

