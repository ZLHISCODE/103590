VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSpecChar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�������ѡ��"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "frmSpecChar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraCard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   0
      Left            =   180
      TabIndex        =   28
      Tag             =   $"frmSpecChar.frx":014A
      Top             =   525
      Width           =   6645
      Begin VB.Frame fraLineHYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2520
         Left            =   3195
         TabIndex        =   32
         Top             =   570
         Width           =   30
      End
      Begin VB.Frame fraLineHYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   510
         TabIndex        =   31
         Top             =   2745
         Width           =   5505
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshHY 
         Height          =   675
         Left            =   510
         TabIndex        =   1
         Top             =   2415
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ĥ  ��"
         Height          =   1245
         Index           =   7
         Left            =   5760
         TabIndex        =   44
         Top             =   1065
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ĥ  ��"
         Height          =   1245
         Index           =   6
         Left            =   5412
         TabIndex        =   43
         Top             =   1065
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  һ  ĥ  ��"
         Height          =   1245
         Index           =   5
         Left            =   5070
         TabIndex        =   42
         Top             =   1065
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ǰ  ĥ  ��"
         Height          =   1605
         Index           =   4
         Left            =   4728
         TabIndex        =   41
         Top             =   705
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  һ  ǰ  ĥ  ��"
         Height          =   1605
         Index           =   3
         Left            =   4386
         TabIndex        =   40
         Top             =   705
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��      ��"
         Height          =   885
         Index           =   2
         Left            =   4044
         TabIndex        =   39
         Top             =   1425
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ��"
         Height          =   885
         Index           =   1
         Left            =   3702
         TabIndex        =   38
         Top             =   1425
         Width           =   165
      End
      Begin VB.Label lblHY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ��"
         Height          =   885
         Index           =   0
         Left            =   3360
         TabIndex        =   37
         Top             =   1425
         Width           =   165
      End
      Begin VB.Label lblHY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   10
         Left            =   3015
         TabIndex        =   36
         Top             =   345
         Width           =   360
      End
      Begin VB.Label lblHY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   11
         Left            =   3015
         TabIndex        =   35
         Top             =   3195
         Width           =   360
      End
      Begin VB.Label lblHY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ҳ�"
         Height          =   180
         Index           =   9
         Left            =   6075
         TabIndex        =   34
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label lblHY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   8
         Left            =   105
         TabIndex        =   33
         Top             =   2670
         Width           =   360
      End
   End
   Begin VB.Frame fraCard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   2
      Left            =   180
      TabIndex        =   29
      Tag             =   "�¾�ʷ"
      Top             =   525
      Width           =   6645
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3720
         TabIndex        =   6
         Text            =   "�վ�����(��ĩ��ͣ������)"
         ToolTipText     =   "�վ�����(��ĩ��ͣ������)"
         Top             =   1800
         Width           =   2220
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2070
         TabIndex        =   5
         Text            =   "�����������"
         ToolTipText     =   "�����������"
         Top             =   2010
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2070
         TabIndex        =   4
         Text            =   "ÿ���о�����"
         ToolTipText     =   "ÿ���о�����"
         Top             =   1605
         Width           =   1170
      End
      Begin VB.TextBox txtYJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   735
         TabIndex        =   3
         Text            =   "��������"
         ToolTipText     =   "��������"
         Top             =   1800
         Width           =   915
      End
      Begin VB.Line Line1 
         X1              =   1995
         X2              =   3315
         Y1              =   1928
         Y2              =   1928
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         TabIndex        =   48
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1710
         TabIndex        =   47
         Top             =   1815
         Width           =   240
      End
   End
   Begin VB.Frame fraCard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   3
      Left            =   180
      TabIndex        =   30
      Tag             =   "����ѡ��"
      Top             =   525
      Width           =   6645
      Begin MSComctlLib.ImageList img32 
         Left            =   255
         Top             =   1575
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   13811126
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecChar.frx":0157
               Key             =   "Select"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSpecChar.frx":0A31
               Key             =   "Item"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���(&A)"
         Height          =   350
         Left            =   3150
         TabIndex        =   9
         Top             =   3855
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdModi 
         Caption         =   "����(&R)"
         Height          =   350
         Left            =   4275
         TabIndex        =   10
         Top             =   3855
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   5400
         TabIndex        =   11
         Top             =   3855
         Visible         =   0   'False
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshChar 
         Height          =   4245
         Left            =   1065
         TabIndex        =   8
         Top             =   0
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   7488
         _Version        =   393216
         BackColor       =   16777215
         FixedRows       =   0
         FixedCols       =   0
         BackColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ListView lvwType 
         DragIcon        =   "frmSpecChar.frx":130B
         Height          =   4245
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   7488
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "img32"
         ForeColor       =   0
         BackColor       =   13752539
         Appearance      =   1
         MouseIcon       =   "frmSpecChar.frx":145D
         NumItems        =   0
      End
   End
   Begin VB.Frame fraCard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Index           =   1
      Left            =   180
      TabIndex        =   27
      Tag             =   "������ע"
      Top             =   525
      Width           =   6645
      Begin VB.Frame fraLineRYH 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   30
         Left            =   1140
         TabIndex        =   46
         Top             =   2790
         Width           =   4065
      End
      Begin VB.Frame fraLineRYV 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2520
         Left            =   3180
         TabIndex        =   45
         Top             =   615
         Width           =   30
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRY 
         Height          =   675
         Left            =   1140
         TabIndex        =   2
         Top             =   2460
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   1191
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorBkg    =   16777215
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   0
         BorderStyle     =   0
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   16
      End
      Begin VB.Label lblRY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   5
         Left            =   705
         TabIndex        =   23
         Top             =   2715
         Width           =   360
      End
      Begin VB.Label lblRY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ҳ�"
         Height          =   180
         Index           =   6
         Left            =   5295
         TabIndex        =   24
         Top             =   2715
         Width           =   360
      End
      Begin VB.Label lblRY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   8
         Left            =   3000
         TabIndex        =   26
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lblRY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Index           =   7
         Left            =   3000
         TabIndex        =   25
         Top             =   390
         Width           =   360
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSpecChar.frx":15BF
         Height          =   1245
         Index           =   0
         Left            =   3375
         TabIndex        =   18
         Top             =   1110
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSpecChar.frx":15D5
         Height          =   1245
         Index           =   1
         Left            =   3690
         TabIndex        =   19
         Top             =   1110
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ��"
         Height          =   885
         Index           =   2
         Left            =   4035
         TabIndex        =   20
         Top             =   1470
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  һ  ��  ĥ  ��"
         Height          =   1605
         Index           =   3
         Left            =   4365
         TabIndex        =   21
         Top             =   750
         Width           =   165
      End
      Begin VB.Label lblRY 
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��  ��  ĥ  ��"
         Height          =   1605
         Index           =   4
         Left            =   4710
         TabIndex        =   22
         Top             =   750
         Width           =   165
      End
   End
   Begin VB.CommandButton cmdDesign 
      Caption         =   "���(&D)"
      Height          =   350
      Left            =   7155
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "����(&M)"
      Height          =   350
      Left            =   7155
      TabIndex        =   16
      Top             =   3420
      Visible         =   0   'False
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabCard 
      Height          =   4770
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   8414
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedWidth   =   2290
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   $"frmSpecChar.frx":15ED
            Key             =   $"frmSpecChar.frx":15FE
            Object.Tag             =   $"frmSpecChar.frx":160B
            Object.ToolTipText     =   $"frmSpecChar.frx":1618
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������ע(&2)"
            Key             =   "������ע"
            Object.Tag             =   "������ע"
            Object.ToolTipText     =   "������ע"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�¾�ʷ(&3)"
            Key             =   "�¾�ʷ"
            Object.Tag             =   "�¾�ʷ"
            Object.ToolTipText     =   "�¾�ʷ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ѡ��(&4)"
            Key             =   "����ѡ��"
            Object.Tag             =   "����ѡ��"
            Object.ToolTipText     =   "����ѡ��"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7155
      TabIndex        =   17
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   350
      Left            =   7155
      TabIndex        =   14
      Top             =   1290
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7155
      TabIndex        =   13
      Top             =   840
      Width           =   1100
   End
   Begin VB.TextBox txtChar 
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   12
      Top             =   4920
      Width           =   6810
   End
End
Attribute VB_Name = "frmSpecChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�¾�ʷ������ʾ
Private Const YJ���� = "��������������������"
Private Const YJ��ĸ = "���˪�����ū۫񬩬�"
Private Const YJ����1 = _
        "�ͪϪѪӪժת٪۪ݪ�" & _
        "�����������" & _
        "��������������������" & _
        "��������������������" & _
        "�ǫɫ˫ͫϫѫӫի׫�" & _
        "�ݫ߫��������" & _
        "�������������������" & _
        "��������������������" & _
        "���ìŬǬɬˬͬϬѬ�"
Private Const YJ����2 = _
        "�����������ªĪƪȪ�" & _
        "�ΪЪҪԪ֪تڪܪު�" & _
        "������������" & _
        "��������������������" & _
        "�����������������«�" & _
        "�ȫʫ̫ΫЫҫԫ֫ث�" & _
        "�ޫ���������" & _
        "��������������������" & _
        "��������������������" & _
        "�¬ĬƬȬʬ̬άЬҬ�"
'������ע�ַ�
Private Const RY���� = "��������������������������������������������������"
Private Const RYС���� = "����������"
Private Const RYС��ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY����� = "����������"
Private Const RY���ĸ = "����������"
Private Const RY�ҷ��� = "����������"
Private Const RY�ҷ�ĸ = "����������"
'������ע�ַ�
Private Const HY���� = "��������������������������������������������������������������������������������������������������������������������������������"
Private Const HYС���� = "����������������"
Private Const HYС��ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY����� = "����������������"
Private Const HY���ĸ = "����������������"
Private Const HY�ҷ��� = "����������������"
Private Const HY�ҷ�ĸ = "����������������"

Public mstrChar As String '��:��ѡ��ķ��ż�
Private Const M_FLAGCOLOR = &HC0E0FF
Private Const SW_SHOWNORMAL = 1
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAdd_Click()
    If txtChar.Text = "" Then
        MsgBox "û�ж����ַ���", vbInformation, gstrSysName
        txtChar.SetFocus: Exit Sub
    End If
    mshChar.Redraw = False
    If mshChar.Rows = 0 Then
        mshChar.Rows = 1
        mshChar.Cols = 2
        mshChar.ColWidth(0) = (mshChar.Width - 60 - 225) * 0.2
        mshChar.ColWidth(1) = (mshChar.Width - 60 - 225) * 0.8
        mshChar.ColAlignment(0) = 1
        mshChar.ColAlignment(1) = 1
    Else
        mshChar.Rows = mshChar.Rows + 1
    End If
    mshChar.Row = mshChar.Rows - 1
    mshChar.Col = 0
    mshChar.CellFontSize = 9
    mshChar.CellAlignment = 4
    mshChar.Text = mshChar.Rows
    
    mshChar.Col = 1
    mshChar.Text = txtChar.Text
    
    mshChar.Col = 0: mshChar.ColSel = mshChar.Cols - 1
    mshChar.TopRow = mshChar.Rows - 1
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Char" & mshChar.Rows - 1, txtChar.Text
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Count", mshChar.Rows
    mshChar.Redraw = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer, k As Integer
    
    If mshChar.Rows = 0 Then
        MsgBox "û�п�ɾ�����ַ���", vbInformation, gstrSysName
        txtChar.SetFocus: Exit Sub
    End If
    
    mshChar.Redraw = False
    
    For i = 0 To mshChar.Rows - 1
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Char" & i
    Next
        
    k = mshChar.Row
    For i = mshChar.Row + 1 To mshChar.Rows - 1
        mshChar.TextMatrix(i, 0) = Val(mshChar.TextMatrix(i, 0)) - 1
    Next
    If k = 0 Then
        mshChar.Rows = 0
    Else
        mshChar.RemoveItem k
    End If
    
    If mshChar.Rows > 0 Then
        If k <= mshChar.Rows - 1 Then
            mshChar.Row = k
        Else
            mshChar.Row = mshChar.Rows - 1
        End If
        mshChar.Col = 0: mshChar.ColSel = mshChar.Cols - 1
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Count", mshChar.Rows
    For i = 0 To mshChar.Rows - 1
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Char" & i, mshChar.TextMatrix(i, 1)
    Next
    
    mshChar.Redraw = True
End Sub

Private Sub cmdDesign_Click()
    Dim strFile As String
    Dim strPath As String * 200
    Dim lngLen As Long
    
    If zlCommFun.IsWindowsNT Then
        strFile = "EUDCEdit.exe"
    Else
        lngLen = GetWindowsDirectory(strPath, 200)
        strFile = Left(strPath, 1) & ":\Program Files\Accessories\EUDCEdit.exe"
        If Dir(strFile) = "" Then
            strFile = "�ù��������ϵͳ��δ��װ��ǰ��װ״̬����ȷ��" & vbCrLf & _
                "�����ʹ��""���/ɾ�������Windows��װ��������������ֳ���""�����°�װ�ù��ܡ�"
            MsgBox strFile, vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    'ShellExecute hwnd, "open", strFile, "", "", SW_SHOWNORMAL
    Shell strFile, vbNormalFocus
    Me.Refresh
    If Err.Number <> 0 Then
        MsgBox "�Բ�����Ĳ���ϵͳ��֧�ָ���ܣ�", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdModi_Click()
    If txtChar.Text = "" Then
        MsgBox "û�ж����ַ���", vbInformation, gstrSysName
        txtChar.SetFocus: Exit Sub
    End If
    mshChar.TextMatrix(mshChar.Row, 1) = txtChar.Text
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Char" & mshChar.Row, txtChar.Text
End Sub

Private Sub cmdMore_Click()
    On Error Resume Next
    Shell "charmap.exe", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "�Բ�����Ĳ���ϵͳ��֧�ָ���ܣ�", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdOK_Click()
    mstrChar = txtChar.Text
    gblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim strSQL As String, i As Integer
    
    gblnOK = False
        
    '������ע
    mshHY.Rows = 2: mshHY.Cols = 16
    mshHY.Height = mshHY.RowHeightMin * mshHY.Rows - 30
    mshHY.Width = mshHY.RowHeightMin * mshHY.Cols - 90
    mshHY.Left = (mshHY.Container.Width - mshHY.Width) / 2
    For i = 0 To mshHY.Cols - 1
        mshHY.ColWidth(i) = mshHY.RowHeight(0)
        mshHY.ColAlignment(i) = 4
        If i + 1 <= 8 Then
            mshHY.TextMatrix(0, i) = 8 - ((i + 1) Mod 9) + 1
            mshHY.TextMatrix(1, i) = 8 - ((i + 1) Mod 9) + 1
        Else
            mshHY.TextMatrix(0, i) = (i - 7) Mod 9
            mshHY.TextMatrix(1, i) = (i - 7) Mod 9
        End If
    Next
    fraLineHYH.Left = mshHY.Left
    fraLineHYH.Top = mshHY.Top + (mshHY.Height - fraLineHYH.Height) / 2
    fraLineHYH.Width = mshHY.Width
    fraLineHYV.Height = mshHY.Height * 3.7
    fraLineHYV.Top = mshHY.Top + mshHY.Height - fraLineHYV.Height + 30
    fraLineHYV.Left = mshHY.Left + (mshHY.Width - fraLineHYV.Width) / 2
    
    For i = 0 To 7
        lblHY(i).Top = mshHY.Top - 75 - lblHY(i).Height
        lblHY(i).Left = fraLineHYV.Left + (mshHY.ColWidth(0) - lblHY(i).Width) / 2 + mshHY.ColWidth(0) * i
    Next
    lblHY(8).Top = fraLineHYH.Top - lblHY(8).Height / 2
    lblHY(8).Left = fraLineHYH.Left - lblHY(8).Width - 60
    lblHY(9).Top = lblHY(8).Top
    lblHY(9).Left = fraLineHYH.Left + fraLineHYH.Width + 60
    lblHY(10).Left = fraLineHYV.Left - lblHY(10).Width / 2
    lblHY(10).Top = fraLineHYV.Top - lblHY(10).Height - 30
    lblHY(11).Left = lblHY(10).Left
    lblHY(11).Top = mshHY.Top + mshHY.Height + 60
    mshHY.Row = 0: mshHY.Col = 8
    
    '������ע
    mshRY.Rows = 2: mshRY.Cols = 10
    mshRY.Height = mshRY.RowHeightMin * mshRY.Rows - 30
    mshRY.Width = mshRY.RowHeightMin * mshRY.Cols - 90
    mshRY.Left = (mshRY.Container.Width - mshRY.Width) / 2
    
    mshRY.TextMatrix(0, 0) = "��"
    mshRY.TextMatrix(0, 1) = "��"
    mshRY.TextMatrix(0, 2) = "��"
    mshRY.TextMatrix(0, 3) = "��"
    mshRY.TextMatrix(0, 4) = "��"
    For i = 0 To mshRY.Cols - 1
        mshRY.ColWidth(i) = mshRY.RowHeight(0)
        mshRY.ColAlignment(i) = 4
        
        If i >= 5 Then mshRY.TextMatrix(0, i) = mshRY.TextMatrix(0, mshRY.Cols - i - 1)
        mshRY.TextMatrix(1, i) = mshRY.TextMatrix(0, i)
    Next
    
    fraLineRYH.Left = mshRY.Left
    fraLineRYH.Top = mshRY.Top + (mshRY.Height - fraLineRYH.Height) / 2
    fraLineRYH.Width = mshRY.Width
    fraLineRYV.Height = mshRY.Height * 3.7
    fraLineRYV.Top = mshRY.Top + mshRY.Height - fraLineRYV.Height + 30
    fraLineRYV.Left = mshRY.Left + (mshRY.Width - fraLineRYV.Width) / 2
    
    For i = 0 To 4
        lblRY(i).Top = mshRY.Top - 75 - lblRY(i).Height
        lblRY(i).Left = fraLineRYV.Left + (mshRY.ColWidth(0) - lblRY(i).Width) / 2 + mshRY.ColWidth(0) * i
    Next
    lblRY(5).Top = fraLineRYH.Top - lblRY(5).Height / 2
    lblRY(5).Left = fraLineRYH.Left - lblRY(5).Width - 60
    lblRY(6).Top = lblRY(5).Top
    lblRY(6).Left = fraLineRYH.Left + fraLineRYH.Width + 60
    lblRY(7).Left = fraLineRYV.Left - lblRY(7).Width / 2
    lblRY(7).Top = fraLineRYV.Top - lblRY(7).Height - 30
    lblRY(8).Left = lblRY(7).Left
    lblRY(8).Top = mshRY.Top + mshRY.Height + 60
    mshRY.Row = 0: mshRY.Col = 5
    
    '����ѡ��
    SetWindowLong lvwType.hWnd, GWL_STYLE, GetWindowLong(lvwType.hWnd, GWL_STYLE) Or &H2000
    
    strSQL = "Select Distinct ��� From ������� Order by Decode(���,'����',1,0),���"
    rsTmp.CursorLocation = adUseClient
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    rsTmp.Open strSQL, gcnOracle, adOpenKeyset
    Call SQLTest
    
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwType.ListItems.Add(, , rsTmp!���, "Item")
        rsTmp.MoveNext
    Next
    Set objItem = lvwType.ListItems.Add(, , "�Զ���", 1)
    objItem.ForeColor = vbBlue
    
    img32.ListImages.Add , "Overlay", img32.Overlay("Item", "Select")
    Call ArrayIcons(lvwType)
    Call lvwType_ItemClick(lvwType.SelectedItem)
    
    '�Ƴ��¾�ʷ��Ƭ,���θù���
    tabCard.Tabs.Remove "�¾�ʷ"
    
    '���
    Call tabCard_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub fraCard_DblClick(Index As Integer)
    Dim strTmp As String
    
    Select Case Index
        Case 0
            strTmp = MakeToothString(mshHY, 8)
            If strTmp <> "" Then
                txtChar.Text = strTmp
                If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
            End If
        Case 1
            strTmp = MakeToothString(mshRY, 5)
            If strTmp <> "" Then
                txtChar.Text = strTmp
                If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
            End If
        Case 2
            strTmp = MakeYJString
            If strTmp <> "" Then
                txtChar.Text = strTmp
                If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
            End If
    End Select
End Sub

Private Sub lvwType_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    
    For i = 1 To lvwType.ListItems.Count
        If i = Item.Index Then
            Item.Icon = "Overlay"
        ElseIf lvwType.ListItems(i).Icon <> i Then
            lvwType.ListItems(i).Icon = "Item"
        End If
    Next
    
    '--
    Call LoadSpecChar(Item.Text)
    
    cmdAdd.Visible = Item.Text = "�Զ���"
    cmdModi.Visible = cmdAdd.Visible
    cmdDel.Visible = cmdAdd.Visible

    mshChar.Height = lvwType.Height - IIf(cmdAdd.Visible, 500, 0)
    
    If Item.Text = "�Զ���" Then
        mshChar.SelectionMode = flexSelectionByRow
    Else
        mshChar.SelectionMode = flexSelectionFree
    End If
End Sub

Private Sub lvwType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objItem As ListItem, blnOver As Boolean
    
    Set objItem = lvwType.HitTest(X, Y)
    If objItem Is Nothing Then
        lvwType.MousePointer = ccDefault
    Else
        lvwType.MousePointer = ccCustom
    End If
    If Button = 1 Then lvwType.Drag 1
End Sub

Private Function LoadSpecChar(strType As String) As Boolean
'���ܣ���ȡָ�����������ַ���
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer, strSQL As String
    Dim intCnt As Integer
    
    Screen.MousePointer = 11
    mshChar.Redraw = False
    
    mshChar.Clear
    mshChar.ClearStructure
    mshChar.Rows = 0: mshChar.Cols = 0
    If strType = "�Զ���" Then
        intCnt = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Count", "0"))
        If intCnt > 0 Then
            mshChar.Cols = 2
            mshChar.Rows = intCnt
            mshChar.ColWidth(0) = 500
            mshChar.ColWidth(1) = mshChar.Width - mshChar.ColWidth(0) - 300
            mshChar.ColAlignment(0) = 1
            mshChar.ColAlignment(1) = 1
            For i = 0 To intCnt - 1
                mshChar.RowHeight(i) = 525
                mshChar.Row = i
                mshChar.Col = 0
                mshChar.CellAlignment = 4: mshChar.CellFontSize = 9
                mshChar.Text = i + 1
                
                mshChar.Col = 1
                mshChar.Text = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\SpecChar\", "Char" & i, "")
            Next
        End If
    Else
        strSQL = "Select * From ������� Where ���='" & strType & "' Order by ����"
        rsTmp.CursorLocation = adUseClient
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        
        mshChar.Cols = 10
        mshChar.Rows = CInt(rsTmp.RecordCount / mshChar.Cols + 0.5)
        mshChar.FixedCols = 0: mshChar.FixedRows = 0
        
        For i = 0 To mshChar.Cols - 1
            mshChar.ColWidth(i) = (mshChar.Width - 300) / mshChar.Cols
        Next
        mshChar.RowHeight(0) = mshChar.ColWidth(0)
        
        mshChar.Row = 0: mshChar.Col = 0
        For i = 1 To rsTmp.RecordCount
            mshChar.CellAlignment = 4
            mshChar.Text = rsTmp!�ַ�
            If mshChar.Col + 1 > mshChar.Cols - 1 Then
                mshChar.Col = 0
                If mshChar.Row + 1 <= mshChar.Rows - 1 Then
                    mshChar.Row = mshChar.Row + 1
                    mshChar.RowHeight(mshChar.Row) = mshChar.ColWidth(0)
                Else
                    mshChar.Row = 0
                End If
            Else
                mshChar.Col = mshChar.Col + 1
            End If
            rsTmp.MoveNext
        Next
    End If
    If mshChar.Rows > 0 Then mshChar.Row = 0
    If mshChar.Cols > 0 Then
        mshChar.Col = 0
        If lvwType.SelectedItem.Text = "�Զ���" Then mshChar.ColSel = mshChar.Cols - 1
    End If
    mshChar.Redraw = True
    Screen.MousePointer = 0
    LoadSpecChar = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub mshChar_DblClick()
    Call mshChar_KeyPress(13)
End Sub

Private Sub mshChar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And mshChar.Text <> "" Then
        KeyAscii = 0
        If lvwType.SelectedItem.Text = "�Զ���" Then
            txtChar.SelText = mshChar.TextMatrix(mshChar.Row, mshChar.Cols - 1)
        Else
            txtChar.SelText = mshChar.Text
        End If
    End If
End Sub

Private Sub mshChar_SelChange()
    mshChar.RowSel = mshChar.Row
    If lvwType.SelectedItem.Text <> "�Զ���" Then mshChar.ColSel = mshChar.Col
End Sub

Private Sub mshHY_Click()
    If mshHY.CellBackColor = vbWhite Then
        mshHY.CellBackColor = M_FLAGCOLOR
    Else
        mshHY.CellBackColor = vbWhite
    End If
    txtChar.Text = MakeToothString(mshHY, 8)
    If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
End Sub

Private Sub mshHY_EnterCell()
    mshHY.CellFontBold = True
    mshHY.CellFontUnderline = True
    mshHY.CellForeColor = vbBlue
End Sub

Private Sub mshHY_GotFocus()
    mshHY_EnterCell
End Sub

Private Sub mshHY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then mshHY_Click
End Sub

Private Sub mshHY_LeaveCell()
    mshHY.CellFontBold = False
    mshHY.CellFontUnderline = False
    mshHY.CellForeColor = mshHY.ForeColor
End Sub

Private Sub mshHY_LostFocus()
    mshHY_LeaveCell
End Sub

Private Sub mshRY_Click()
    If mshRY.CellBackColor = vbWhite Then
        mshRY.CellBackColor = M_FLAGCOLOR
    Else
        mshRY.CellBackColor = vbWhite
    End If
    txtChar.Text = MakeToothString(mshRY, 5)
    If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
End Sub

Private Sub mshRY_EnterCell()
    mshRY.CellFontBold = True
    mshRY.CellFontUnderline = True
    mshRY.CellForeColor = vbBlue
End Sub

Private Sub mshRY_GotFocus()
    mshRY_EnterCell
End Sub

Private Sub mshRY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then mshRY_Click
End Sub

Private Sub mshRY_LeaveCell()
    mshRY.CellFontBold = False
    mshRY.CellFontUnderline = False
    mshRY.CellForeColor = mshRY.ForeColor
End Sub

Private Sub mshRY_LostFocus()
    mshRY_LeaveCell
End Sub

Private Sub tabCard_Click()
    Dim i As Integer
    
    For i = 0 To fraCard.UBound
        fraCard(i).Visible = fraCard(i).Tag = tabCard.SelectedItem.Key
    Next
    cmdMore.Visible = tabCard.SelectedItem.Index = tabCard.Tabs.Count
    cmdDesign.Visible = cmdMore.Visible
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtChar_Change()
    cmdOK.Enabled = txtChar.Text <> ""
End Sub

Private Sub txtChar_KeyPress(KeyAscii As Integer)
    If InStr("'%?&", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtYJ_Change(Index As Integer)
    If Visible Then
        txtChar.Text = MakeYJString
        If txtChar.SelLength = 0 Then txtChar.SelStart = Len(txtChar.Text)
    End If
End Sub

Private Sub txtYJ_DblClick(Index As Integer)
    txtYJ_Change Index
End Sub

Private Sub txtYJ_GotFocus(Index As Integer)
    If txtYJ(Index).Text = txtYJ(Index).ToolTipText Then
        'txtYJ(Index).Text = ""
    End If
    zlControl.TxtSelAll txtYJ(Index)
End Sub

Private Sub txtYJ_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtYJ_LostFocus(Index As Integer)
    If Index = 3 Then
        If Not (IsNumeric(txtYJ(Index).Text) Or IsDate(txtYJ(Index).Text)) Then
            txtYJ(Index).Text = txtYJ(Index).ToolTipText
        End If
    Else
        If Not IsNumeric(txtYJ(Index).Text) Then
            txtYJ(Index).Text = txtYJ(Index).ToolTipText
        End If
    End If
End Sub

Private Function MakeToothString(objMSH As MSHFlexGrid, bytCount As Byte) As String
'���ܣ����ݺ�����ע��������ʾ������ע�������ַ�����
'������objMSH=������������ע���
'      bytCount=����������
    Dim intRow As Integer, intCol As Integer
    Dim byt���� As Byte, byt��ĸ As Byte
    Dim i As Integer, j As Integer, strTmp As String
    Dim A As String, B As String, C As String, D As String 'A=����,B=����,C=����,D=����
        
    Dim YC���� As String
    Dim YCС���� As String, YCС��ĸ As String
    Dim YC����� As String, YC���ĸ As String
    Dim YC����� As String, YC���ĸ As String
    Dim YC�ҷ��� As String, YC�ҷ�ĸ As String
        
    If objMSH.Name = "mshHY" Then
        YC���� = HY����
        YCС���� = HYС����: YCС��ĸ = HYС��ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC����� = HY�����: YC���ĸ = HY���ĸ
        YC�ҷ��� = HY�ҷ���: YC�ҷ�ĸ = HY�ҷ�ĸ
    Else
        YC���� = RY����
        YCС���� = RYС����: YCС��ĸ = RYС��ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC����� = RY�����: YC���ĸ = RY���ĸ
        YC�ҷ��� = RY�ҷ���: YC�ҷ�ĸ = RY�ҷ�ĸ
    End If
            
    '��ABCD�ĸ�����ı�ע���,�����Ŀ�ʼ��ݺ�,��"37"
    objMSH.Redraw = False
    intRow = objMSH.Row: intCol = objMSH.Col
    
    objMSH.Row = 0
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then A = A & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then B = B & i - bytCount
    Next
    
    objMSH.Row = 1
    For i = bytCount To 1 Step -1
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then C = C & bytCount + 1 - i
    Next
    For i = bytCount + 1 To bytCount * 2
        objMSH.Col = i - 1
        If objMSH.CellBackColor = M_FLAGCOLOR Then D = D & i - bytCount
    Next
    
    objMSH.Row = intRow: objMSH.Col = intCol
    objMSH.Redraw = True
    
    '���ݲ�ͬ�ĸ��������������ע�����ַ���
    If A <> "" And B = "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = Len(A) To 1 Step -1
            If i = 1 Then
                strTmp = strTmp & Mid(YC�����, CByte(Mid(A, i, 1)), 1)
            Else
                strTmp = strTmp & Mid(YC�����, CByte(Mid(A, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And B <> "" And C = "" And D = "" Then
        'ֻ�����ϱ�ע
        For i = 1 To Len(B)
            If i = 1 Then
                strTmp = strTmp & Mid(YC�ҷ���, CByte(Mid(B, i, 1)), 1)
            Else
                strTmp = strTmp & Mid(YC�����, CByte(Mid(B, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And B = "" And C <> "" And D = "" Then
        'ֻ�����±�ע
        For i = Len(C) To 1 Step -1
            If i = 1 Then
                strTmp = strTmp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            Else
                strTmp = strTmp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
            End If
        Next
    ElseIf A = "" And B = "" And C = "" And D <> "" Then
        'ֻ�����±�ע
        For i = 1 To Len(D)
            If i = 1 Then
                strTmp = strTmp & Mid(YC�ҷ�ĸ, CByte(Mid(D, i, 1)), 1)
            Else
                strTmp = strTmp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
            End If
        Next
    ElseIf A <> "" And B <> "" And C = "" And D = "" Then
        'ֻ���������б�ע
        For i = Len(A) To 1 Step -1
            strTmp = strTmp & Mid(YC�����, CByte(Mid(A, i, 1)), 1)
        Next
        strTmp = strTmp & "��"
        For i = 1 To Len(B)
            strTmp = strTmp & Mid(YC�����, CByte(Mid(B, i, 1)), 1)
        Next
    ElseIf A = "" And B = "" And C <> "" And D <> "" Then
        'ֻ���������б�ע
        For i = Len(C) To 1 Step -1
            strTmp = strTmp & Mid(YC���ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTmp = strTmp & "��"
        For i = 1 To Len(D)
            strTmp = strTmp & Mid(YC���ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf A <> "" And B = "" And C = "" And D <> "" Then
        'ֻ�����������б�ע
        For i = Len(A) To 1 Step -1
            strTmp = strTmp & Mid(YCС����, CByte(Mid(A, i, 1)), 1)
        Next
        strTmp = strTmp & "��"
        For i = 1 To Len(D)
            strTmp = strTmp & Mid(YCС��ĸ, CByte(Mid(D, i, 1)), 1)
        Next
    ElseIf A = "" And B <> "" And C <> "" And D = "" Then
        'ֻ�����������б�ע
        For i = Len(C) To 1 Step -1
            strTmp = strTmp & Mid(YCС��ĸ, CByte(Mid(C, i, 1)), 1)
        Next
        strTmp = strTmp & "��"
        For i = 1 To Len(B)
            strTmp = strTmp & Mid(YCС����, CByte(Mid(B, i, 1)), 1)
        Next
    ElseIf Not (A = "" And B = "" And C = "" And D = "") Then
        '���¶��б�ע
        If A = "" And C = "" Then strTmp = "��"
        
        '����߷�����
        i = 1: j = 1 'i��ӦA,j��ӦC
        Do While i <= Len(A) Or j <= Len(C)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(A) Then byt���� = Mid(A, i, 1)
            If j <= Len(C) Then byt��ĸ = Mid(C, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTmp = strTmp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTmp = strTmp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTmp = strTmp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
        strTmp = StrReverse(strTmp)
        
        '���ӷ�
        If (A <> "" Or C <> "") And (B <> "" Or D <> "") Then
            strTmp = strTmp & "��"
        ElseIf B = "" And D = "" Then
            strTmp = strTmp & "��"
        End If
        
        '���ұ߷�����
        i = 1: j = 1 'i��ӦB,j��ӦD
        Do While i <= Len(B) Or j <= Len(D)
            byt���� = 0: byt��ĸ = 0
            If i <= Len(B) Then byt���� = Mid(B, i, 1)
            If j <= Len(D) Then byt��ĸ = Mid(D, j, 1)
            '���ݷ��ӷ�ĸ��һ�������������
            If byt���� <> 0 And byt��ĸ <> 0 Then
                strTmp = strTmp & Mid(YC����, (byt��ĸ - 1) * bytCount + byt����, 1)
            ElseIf byt���� <> 0 And byt��ĸ = 0 Then
                strTmp = strTmp & Mid(YCС����, byt����, 1)
            ElseIf byt���� = 0 And byt��ĸ <> 0 Then
                strTmp = strTmp & Mid(YCС��ĸ, byt��ĸ, 1)
            End If
            i = i + 1: j = j + 1
        Loop
    End If
    
    MakeToothString = strTmp
End Function

Private Function MakeYJString() As String
'���ܣ������¾�ʷ��д���������������ַ���ע��
    Dim str���� As String, str��ĸ As String
    Dim strTmp As String
    
    If Not (IsNumeric(txtYJ(1).Text) And IsNumeric(txtYJ(2).Text)) Then Exit Function
    
    '��������֣��������Ҷ���
    '------------------------
    str���� = Right(Format(Int(txtYJ(1).Text), "00"), 2)
    str��ĸ = Right(Format(Int(txtYJ(2).Text), "00"), 2)
    
    '��10λ���ַ�
    If Val(Left(str��ĸ, 1)) <> 0 Or Val(Left(str����, 1)) <> 0 Then
        If Val(Left(str��ĸ, 1)) <> 0 And Val(Left(str����, 1)) <> 0 Then
            strTmp = Mid(YJ����1, (Val(Left(str��ĸ, 1)) - 1) * 10 + Val(Left(str����, 1)) + 1, 1)
        ElseIf Val(Left(str����, 1)) = 0 Then
            strTmp = Mid(YJ��ĸ, Val(Left(str��ĸ, 1)) + 1, 1)
        ElseIf Val(Left(str��ĸ, 1)) = 0 Then
            strTmp = Mid(YJ����, Val(Left(str����, 1)) + 1, 1)
        End If
    End If
        
    '���λ���ַ�
    strTmp = strTmp & Mid(YJ����2, Val(Right(str��ĸ, 1)) * 10 + Val(Right(str����, 1)) + 1, 1)
        
    '��������ַ�
    If IsNumeric(txtYJ(0).Text) Then
        strTmp = txtYJ(0).Text & " ��" & strTmp
    End If
    If IsNumeric(txtYJ(3).Text) Or IsDate(txtYJ(3).Text) Then
        strTmp = strTmp & "�� " & txtYJ(3).Text
    End If
    
    MakeYJString = strTmp
End Function
