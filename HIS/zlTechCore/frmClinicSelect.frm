VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "������Ŀѡ����"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   Icon            =   "frmClinicSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9120
   Begin VB.CheckBox chkStock 
      Caption         =   "��ʾҩƷ���(&S)"
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   5658
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Frame fraInfo 
      Height          =   435
      Left            =   30
      TabIndex        =   15
      Top             =   -75
      Width           =   9075
      Begin VB.CheckBox chkSub 
         Caption         =   "�����¼���Ŀ(&T)"
         Height          =   195
         Left            =   7380
         TabIndex        =   10
         Top             =   165
         Width           =   1650
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   180
         Left            =   225
         TabIndex        =   16
         Top             =   165
         Width           =   270
      End
   End
   Begin VB.Frame fraStat 
      Height          =   5160
      Left            =   15
      TabIndex        =   18
      Top             =   285
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdSelClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   1425
         TabIndex        =   7
         ToolTipText     =   "Ctrl+R"
         Top             =   3705
         Width           =   1100
      End
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   1425
         TabIndex        =   6
         ToolTipText     =   "Ctrl+A"
         Top             =   3345
         Width           =   1100
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "ȫ����ʾ"
         Height          =   195
         Left            =   1050
         TabIndex        =   4
         Top             =   1965
         Width           =   1020
      End
      Begin VB.CommandButton cmdStat 
         Caption         =   "ͳ��(&S)"
         Height          =   350
         Left            =   1425
         TabIndex        =   5
         Top             =   2835
         Width           =   1100
      End
      Begin VB.TextBox txtCount 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1050
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "100"
         Top             =   1620
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   1185
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   67174403
         CurrentDate     =   38434
      End
      Begin VB.ComboBox cboDate 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmClinicSelect.frx":058A
         Left            =   1050
         List            =   "frmClinicSelect.frx":05A6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   825
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "���ͳ�Ƶ�ʱ�䷶Χ�ϳ����ٶȿ��ܻ�����������ĵȴ���"
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   165
         TabIndex        =   23
         Top             =   2340
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2100
         TabIndex        =   22
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��ʾ��ǰ"
         Height          =   180
         Left            =   285
         TabIndex        =   21
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ͳ��ʱ��"
         Height          =   180
         Left            =   285
         TabIndex        =   20
         Top             =   885
         Width           =   720
      End
      Begin VB.Label lblStatTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "�Զ�ͳ��""XXXXXX""������õ�������Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   270
         Width           =   2400
      End
   End
   Begin MSComctlLib.ImageList imgOften 
      Left            =   1110
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":060A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":0D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":13FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":1AF8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrOften 
      Height          =   450
      Left            =   495
      TabIndex        =   17
      Top             =   5505
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   794
      ButtonWidth     =   1561
      ButtonHeight    =   794
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgOften"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "Often"
            Description     =   "����"
            Object.ToolTipText     =   "��ʾ������Ŀ(F2)"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "ͳ��"
            Key             =   "Stat"
            Description     =   "ͳ��"
            Object.ToolTipText     =   "ͳ�Ƴ�����Ŀ"
            Object.Tag             =   "ͳ��"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "New"
            Description     =   "����"
            Object.ToolTipText     =   "���볣����Ŀ(F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "�Ƴ�"
            Key             =   "Del"
            Description     =   "�Ƴ�"
            Object.ToolTipText     =   "�Ƴ�������Ŀ(Del)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4665
      Left            =   2790
      TabIndex        =   8
      Top             =   375
      Width           =   6300
      _cx             =   11112
      _cy             =   8229
      Appearance      =   1
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicSelect.frx":21F2
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   930
         Top             =   900
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":227F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClinicSelect.frx":2759
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   2745
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   555
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   13
      Top             =   5580
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6240
      TabIndex        =   12
      Top             =   5580
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   600
      Left            =   2835
      TabIndex        =   9
      Top             =   4815
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   1058
      TabWidthStyle   =   2
      TabFixedWidth   =   1623
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      Placement       =   1
      ImageList       =   "img16"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��(0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�г�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�в�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":2C33
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":31CD
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3767
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":3D01
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":429B
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicSelect.frx":4835
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   4995
      Left            =   15
      TabIndex        =   0
      Top             =   390
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   8811
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Shape Shp 
      Height          =   405
      Left            =   4800
      Top             =   5550
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   10000
      Y1              =   5445
      Y2              =   5445
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmClinicSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrPrivs As String
Private mrsItem As ADODB.Recordset
Private mint��Ч As Integer
Private mstr�Ա� As String
Private mstr���� As String
Private mobjTXT As Object
Private mint��Χ As Integer '1-����,2-סԺ
Private mlng����ID As Long

Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean

Private mbln�۸� As Boolean
Private mbln���� As Boolean
Private mint���� As Integer
Private mstrLike As String
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long
Private mlng��ҩ�� As Long

Public Function ShowSelect(frmParent As Object, ByVal strPrivs As String, ByVal int��Ч As Integer, ByVal str�Ա� As String, _
    Optional ByVal str���� As String, Optional objTXT As Object, Optional ByVal int��Χ As Integer = 2, _
    Optional ByVal lng����ID As Long) As ADODB.Recordset
'���ܣ���ʾ������Ŀѡ����
'������int��Ч=ҽ����Ч
'      str�Ա�=�����Ա�
'      str����=����ƥ�������,���û����Ϊѡ������ʽ,����Ϊ�б�ʽ
'      objTXT=�����б�λ�������
'      blnCancel(O):�Ƿ�ȡ��
'      int��Χ=1-����,2-סԺ
'      lng����ID=ѡ����ʱ(str����="")����������࿪ʼ��ʾ
'���أ����û������,��ȡ��,�򷵻�Nothing������Ϊһ������������Ŀ���ݵļ�¼
    mstrPrivs = strPrivs
    mint��Ч = int��Ч
    mstr�Ա� = str�Ա�
    mstr���� = str����
    Set mobjTXT = objTXT
    mint��Χ = int��Χ
    mlng����ID = lng����ID
    
    mstrSaveTag = mint��Χ & IIF(mstr���� <> "", 1, 0) & IIF(gblnҩƷ�������ҽ�� Or mint��Ч = 1, 1, 0)
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        Set ShowSelect = mrsItem
    Else
        Set ShowSelect = Nothing
    End If
End Function

Private Sub cboDate_Click()
    Dim curDate As Date
    
    If cboDate.ListIndex = cboDate.ListCount - 1 Then
        dtpDate.Enabled = True
        dtpDate.SetFocus
    Else
        dtpDate.Enabled = False
        curDate = zlDatabase.Currentdate
        Select Case cboDate.ListIndex
            Case 0 'һ��
                dtpDate.Value = DateAdd("ww", -1, curDate)
            Case 1 '����(15��)
                dtpDate.Value = DateAdd("d", -15, curDate)
            Case 2 'һ��
                dtpDate.Value = DateAdd("m", -1, curDate)
            Case 3 '����
                dtpDate.Value = DateAdd("m", -2, curDate)
            Case 4 '����
                dtpDate.Value = DateAdd("m", -3, curDate)
            Case 5 '����
                dtpDate.Value = DateAdd("m", -6, curDate)
            Case 6 'һ��
                dtpDate.Value = DateAdd("yyyy", -1, curDate)
        End Select
    End If
End Sub

Private Sub chkAll_Click()
    txtCount.Enabled = chkAll.Value = 0
End Sub

Private Sub chkStock_Click()
    If Visible Then
        Call FillList
        vsItem.SetFocus
    End If
End Sub

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    Set mrsItem = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdStat_Click()
    If chkAll.Value = 0 Then
        If Val(txtCount.Text) <= 0 Then
            MsgBox "��������ȷ����ʾ������", vbInformation, gstrSysName
            txtCount.SetFocus: Exit Sub
        End If
    End If
    
    Call FillStat(True)
    vsItem.SetFocus
End Sub

Private Sub cmdSelALL_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        If Val(vsItem.TextMatrix(i, 1)) <> 0 Then
            vsItem.TextMatrix(i, 2) = 1
        End If
    Next
End Sub

Private Sub cmdSelClear_Click()
    Dim i As Long
    
    For i = 1 To vsItem.Rows - 1
        vsItem.TextMatrix(i, 2) = 0
    Next
End Sub

Private Sub Form_Activate()
    If Not tvw_s.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf Shift = vbAltMask Then
        If Between(KeyCode, vbKey0, vbKey9) Then
            lngIdx = KeyCode - vbKey0 + 1
        End If
        If tabClass.SelectedItem.Index <> lngIdx And Between(lngIdx, 1, tabClass.Tabs.Count) Then
            tabClass.Tabs(lngIdx).Selected = True
        End If
    ElseIf Shift = vbCtrlMask Then
        If KeyCode = vbKeyA Then
            If fraStat.Visible Then cmdSelALL_Click
        ElseIf KeyCode = vbKeyR Then
            If fraStat.Visible Then cmdSelClear_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If tbrOften.Buttons("Often").Visible And tbrOften.Buttons("Often").Enabled Then
            If tbrOften.Buttons("Often").Value = tbrPressed Then
                tbrOften.Buttons("Often").Value = tbrUnpressed
            Else
                tbrOften.Buttons("Often").Value = tbrPressed
            End If
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If tbrOften.Buttons("New").Visible And tbrOften.Buttons("New").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("New"))
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If tbrOften.Buttons("Del").Visible And tbrOften.Buttons("Del").Enabled Then
            Call tbrOften_ButtonClick(tbrOften.Buttons("Del"))
        End If
    End If
End Sub

Private Function ExistOftenItem(Optional ByVal lng����ID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng����ID <> 0 Then
        strSQL = "Select ID From ���Ʒ���Ŀ¼ Start With ID=[2] Connect by Prior ID=�ϼ�ID"
        strSQL = "Select Count(A.������ĿID) as Num From ���Ƹ�����Ŀ A,������ĿĿ¼ B Where A.������ĿID=B.ID And B.����ID IN(" & strSQL & ") And A.��ԱID=[1]"
    Else
        strSQL = "Select Count(*) as Num From ���Ƹ�����Ŀ Where ��ԱID=[1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, UserInfo.ID, lng����ID)
    ExistOftenItem = Nvl(rsTmp!Num, 0) > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Dim blnDo As Boolean
    Dim str������ĿIDs As String, str�շ�ϸĿIDs As String
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)

    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    Set mrsItem = Nothing
    
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "") '����ƥ�䷽ʽ
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint��Χ = 1, "����", "סԺ") & "ȱʡ��ҩ��", 0))
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint��Χ = 1, "����", "סԺ") & "ȱʡ��ҩ��", 0))
    mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint��Χ = 1, "����", "סԺ") & "ȱʡ��ҩ��", 0))
    
    'ѡ�����е�����
    mbln���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ��Ŀ����", 1)) <> 0 '�Ƿ���ʾ����
    mbln�۸� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾҩƷ�۸�", 1)) <> 0 '�Ƿ���ʾҩƷ�۸�
    chkStock.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾҩƷ���", 0)) '�Ƿ���ʾҩƷ���
    
    lblStatTitle.Caption = Replace(lblStatTitle.Caption, "XXXXXX", UserInfo.����)
    cboDate.ListIndex = 0
    Call SetOftenToolBar(mstr���� = "")
    
    If mstr���� = "" Then
        tvw_s.Visible = True
        
        '��ȡ���ʧ��,����ʾ,��ȡ���˳�
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '�����,��ʾ,��ȡ���˳�
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "û����������������,���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
        
        '����и�����Ŀ,ȱʡת��������Ŀ
        If ExistOftenItem(mlng����ID) Then
            tbrOften.Buttons("Often").Value = tbrPressed
            Call tbrOften_ButtonClick(tbrOften.Buttons("Often"))
        End If
    Else
        fraInfo.Visible = False
        tvw_s.Visible = False
        fraLR.Visible = False
        chkSub.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Line1(0).Visible = False
        Line1(1).Visible = False
        Shp.Visible = True

        '�����ƥ��ĸ�����Ŀ,������ʾ������Ŀ
        If ExistOftenItem Then
            tbrOften.Buttons("Often").Value = tbrPressed
            Call SwitchToOften(False, False)
            Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
            
            '���û�����л�����
            If Not cmdOK.Enabled Then
                tbrOften.Buttons("Often").Value = tbrUnpressed
                Call SwitchToOften(False, False)
                Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
            End If
        Else
            '���ƥ������
            Call FillList(True, str������ĿIDs, str�շ�ϸĿIDs)
        End If
        
        If cmdOK.Enabled And vsItem.Rows = vsItem.FixedRows + 1 Then
            'ֻ��һ����Ŀʱ,ֱ�ӷ���
            If tbrOften.Buttons("Often").Value = tbrUnpressed Then
                mblnOK = True: Unload Me: Exit Sub
            Else
                blnDo = True '������Ŀƥ��ʱʼ����ʾ
            End If
        End If
        
        If (cmdOK.Enabled And vsItem.Rows > vsItem.FixedRows + 1) Or blnDo Then
            '������ͬһ����Ŀʱ,ֱ�ӷ���:�������շ�ϸĿID
            If mstr���� <> "" Then
                If UBound(Split(str������ĿIDs, ",")) = 1 _
                    And UBound(Split(str�շ�ϸĿIDs, ",")) <= 1 Then
                    '������Ŀƥ��ʱʼ����ʾ
                    If tbrOften.Buttons("Often").Value = tbrUnpressed Then
                        mblnOK = True: Unload Me: Exit Sub
                    End If
                End If
            End If
        
            vsItem.Appearance = ccFlat
            vsItem.BorderStyle = ccFixedSingle
                        
            Call SetFormSize
            Call Form_Resize
        Else
            '������,��ʾ,��ȡ���˳�
            MsgBox "û���ҵ�ƥ���������Ŀ��", vbInformation, gstrSysName
            Set mrsItem = Nothing
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub SetFormSize()
    Dim vRect As RECT, i As Long
    Dim lngUpH As Long, lngDnH As Long
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long

    Call FormSetCaption(Me, False, False)
    Call GetWindowRect(mobjTXT.Hwnd, vRect) '�����λ��
    
    '���ô���ߴ��λ��
    '������
    Me.Left = vRect.Left * Screen.TwipsPerPixelX
    lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 60 '+3D�߿�
    For i = 0 To vsItem.Cols - 1
        lngColW = lngColW + IIF(vsItem.ColHidden(i), 0, vsItem.ColWidth(i))
    Next
    If Me.Left + lngColW + lngScrW > Screen.Width - lngScrW Then
        Me.Width = Screen.Width - lngScrW - Me.Left
    Else
        Me.Width = lngColW + lngScrW
    End If
    
    '����߶�
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '��Ļ���ø߶�
    lngUpH = vRect.Top * Screen.TwipsPerPixelY '������ø߶�
    lngDnH = lngScrH - vRect.Bottom * Screen.TwipsPerPixelY '������ø߶�
    Me.Height = vsItem.Rows * vsItem.RowHeight(0) + tbrOften.Height + 45 '395 '+���Ƭ�߶�
    If Me.Height < 1500 Then Me.Height = 2000 '������С�߶�
    If Me.Height > lngUpH And Me.Height > lngDnH Then
        Me.Height = IIF(lngUpH < lngDnH, lngDnH, lngUpH)
    End If
    If Me.Height > lngScrH / 2 Then Me.Height = lngScrH / 2 '�������߶�
    If Me.Height <= lngDnH Then
        Me.Top = vRect.Bottom * Screen.TwipsPerPixelY
    ElseIf Me.Height <= lngUpH Then
        Me.Top = vRect.Top * Screen.TwipsPerPixelY - Me.Height
    End If
End Sub
    
Private Sub SetOftenToolBar(ByVal blnCaption As Boolean)
'���ܣ����ù������Ƿ���ʾ�ı�
    Dim lngW As Long, i As Long
    
    For i = 1 To tbrOften.Buttons.Count
        tbrOften.Buttons(i).Caption = IIF(blnCaption, tbrOften.Buttons(i).Description, "")
    Next
    If blnCaption Then
        tbrOften.TextAlignment = tbrTextAlignRight
    Else
        tbrOften.TextAlignment = tbrTextAlignBottom
    End If
    
    For i = 1 To tbrOften.Buttons.Count
        If tbrOften.Buttons(i).Visible Then
            lngW = lngW + tbrOften.Buttons(i).Width
        End If
    Next
    tbrOften.Width = lngW
End Sub

Private Sub Form_Resize()
    Dim lngLeft As Long
    
    On Error Resume Next
    
    If mstr���� = "" Then
        fraInfo.Left = 0
        fraInfo.Width = Me.ScaleWidth
        chkSub.Left = fraInfo.Width - IIF(chkSub.Visible, chkSub.Width, 0) - 45
        lblInfo.Width = IIF(chkSub.Visible, chkSub.Left, fraInfo.Width) - lblInfo.Left - 45
        
        lngLeft = IIF(tvw_s.Visible, tvw_s.Width, 0) + IIF(fraStat.Visible, fraStat.Width, 0) + IIF(fraLR.Visible, fraLR.Width, 0)
        
        If tvw_s.Visible Then
            tvw_s.Left = 0
            tvw_s.Top = fraInfo.Top + fraInfo.Height + 15
            tvw_s.Height = Me.ScaleHeight - tvw_s.Top - 615
        End If
        If fraStat.Visible Then
            fraStat.Left = 0
            fraStat.Top = fraInfo.Top + fraInfo.Height - 90
            fraStat.Height = Me.ScaleHeight - fraStat.Top - 615
        End If
        If fraLR.Visible Then
            fraLR.Top = tvw_s.Top
            fraLR.Left = tvw_s.Left + tvw_s.Width
            fraLR.Height = tvw_s.Height
        End If
        
        vsItem.Top = fraInfo.Top + fraInfo.Height + 15
        vsItem.Left = lngLeft
        vsItem.Width = Me.ScaleWidth - lngLeft
        vsItem.Height = Me.ScaleHeight - vsItem.Top - 615 - IIF(tabClass.Visible, 350, 0)
        
        If tabClass.Visible Then
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
            tabClass.Left = vsItem.Left + 30
            tabClass.Width = vsItem.Width - 60
        End If
        
        Line1(0).X1 = 0: Line1(0).X2 = Me.ScaleWidth
        Line1(0).Y1 = tvw_s.Top + vsItem.Height + IIF(tabClass.Visible, 350, 0) + 75: Line1(0).Y2 = Line1(0).Y1
        
        Line1(1).X1 = Line1(0).X1: Line1(1).X2 = Line1(0).X2
        Line1(1).Y1 = Line1(0).Y1 - 15: Line1(1).Y2 = Line1(1).Y1
        
        cmdOK.Top = Line1(1).Y1 + 120
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.8 < 4000 Then
            cmdCancel.Left = 4000
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.8
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 15
        
        tbrOften.Top = cmdOK.Top + (cmdOK.Height - tbrOften.Height) / 2
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        'vsItem.Height = Me.ScaleHeight - IIf(tabClass.Tabs.Count > 1, 380, 0)
        vsItem.Height = Me.ScaleHeight - tbrOften.Height - 15
        
        tbrOften.Left = Me.ScaleWidth - tbrOften.Width - 15
        tbrOften.Top = vsItem.Top + vsItem.Height
        
        If chkStock.Visible Then
            chkStock.Left = tbrOften.Left - chkStock.Width - 30
            chkStock.Top = tbrOften.Top + (tbrOften.Height - chkStock.Height) / 2
        End If
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - tbrOften.Width - IIF(chkStock.Visible, chkStock.Width, 0) - 120
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
        End If
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ѡ�����е�����
    If chkStock.Visible Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾҩƷ���", chkStock.Value '�Ƿ���ʾҩƷ���
    End If

    If tbrOften.Buttons("Often").Value = tbrPressed Then
        Call SaveColPosition("Often")
        Call SaveColWidth("Often")
    Else
        Call SaveColPosition
        Call SaveColWidth
    End If
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or vsItem.Width - x < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tvw_s.Width = tvw_s.Width + x
        vsItem.Left = vsItem.Left + x
        vsItem.Width = vsItem.Width - x
        tabClass.Left = tabClass.Left + x
        tabClass.Width = tabClass.Width - x
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    On Error GoTo errH
    
    If mlng����ID <> 0 Then
        strSQL = _
            " Select 1 as ��,����,ID,�ϼ�ID,����,���� From ���Ʒ���Ŀ¼ Where ID=[1]" & _
            " Union ALL " & _
            " Select Level+1 as ��,����,ID,�ϼ�ID,����,����" & _
            " From ���Ʒ���Ŀ¼ Where ����<>7 Start With �ϼ�ID=[1] Connect by Prior ID=�ϼ�ID" & _
            " Order by ��,����"
    Else
        strSQL = _
            " Select 0 as ��,����,-���� as ID,-NULL as �ϼ�ID,NULL as ����," & _
            " ����||'.'||Decode(����,1,'����ҩ',2,'�г�ҩ',3,'�в�ҩ',4,'��ҩ�䷽',5,'������Ŀ',6,'��������','7','��������') as ����" & _
            " From ���Ʒ���Ŀ¼ Where ����<>7 Group by ����"
        strSQL = strSQL & " Union ALL " & _
            " Select Level as ��,����,ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,����,����" & _
            " From ���Ʒ���Ŀ¼ Where ����<>7 Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
            " Order by ��,����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, "Close")
        End If
        objNode.Tag = rsTmp!���� '��ŷ�������
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tabClass_Click()
    If Not mblnClick Then Exit Sub
    
    If fraStat.Visible Then
        Call FillStat
    Else
        Call FillList
    End If
    vsItem.SetFocus
End Sub

Private Sub SwitchToOften(Optional ByVal blnFill As Boolean = True, Optional ByVal blnSaveColPos As Boolean = True)
'���ܣ��ڳ�����Ŀ��ѡ����Ŀ����֮���л�
'������blnFill=�л�֮���Ƿ�����ˢ���嵥
    Dim blnNoStat As Boolean
    
    '�䷽�ͳ����޷�ͳ�Ƴ���
    If mlng����ID <> 0 And Not tvw_s.SelectedItem Is Nothing Then
        If InStr(",4,6,", Val(tvw_s.SelectedItem.Tag)) > 0 Then
            blnNoStat = True
        End If
    End If
    tbrOften.Buttons("Stat").Visible = tbrOften.Buttons("Often").Value = tbrPressed And mstr���� = "" And Not blnNoStat
    
    tbrOften.Buttons("New").Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
    tbrOften.Buttons("Del").Visible = tbrOften.Buttons("Often").Value = tbrPressed
    If mstr���� = "" Then
        chkSub.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        tvw_s.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        fraLR.Visible = tbrOften.Buttons("Often").Value = tbrUnpressed
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").Value = tbrPressed Then
                Call SaveColPosition(tvw_s.SelectedItem.Tag)
                Call SaveColWidth(tvw_s.SelectedItem.Tag)
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    Else
        If blnSaveColPos Then
            If tbrOften.Buttons("Often").Value = tbrPressed Then
                Call SaveColPosition
                Call SaveColWidth
            Else
                Call SaveColPosition("Often")
                Call SaveColWidth("Often")
            End If
        End If
    End If
    Call SetOftenToolBar(mstr���� = "")
    Call Form_Resize
    
    If blnFill Then Call FillList(True)
End Sub

Private Sub SwitchToState(Optional ByVal blnFill As Boolean = True)
'���ܣ��ڳ�����Ŀ���棬����ͳ�ƺͳ��ý���֮���л�
'������blnFill=�л�֮���Ƿ�����ˢ���嵥
    If tbrOften.Buttons("Stat").Value = tbrUnpressed Then
        fraStat.Visible = False
        tbrOften.Buttons("Del").Visible = True
        tbrOften.Buttons("New").Visible = False
        Call Form_Resize
        If blnFill Then Call FillList(True)
        If Visible Then vsItem.SetFocus
    Else
        fraStat.Visible = True
        lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ"
        tbrOften.Buttons("Del").Visible = False
        tbrOften.Buttons("New").Visible = True
        Call Form_Resize
        vsItem.FixedRows = 0: vsItem.Rows = 0
        vsItem.FixedCols = 0: vsItem.Cols = 0
        If Visible Then cboDate.SetFocus
    End If
End Sub

Private Sub tbrOften_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Often" Then
        If Visible And mstr���� <> "" Then
            LockWindowUpdate Me.Hwnd
            Call SwitchToOften
            Call SetFormSize
            Call Form_Resize
            LockWindowUpdate 0
        Else
            '�л���ѡ�����ʱ�ȹر�ͳ�ƽ���
            If tbrOften.Buttons("Stat").Value = tbrPressed Then
                tbrOften.Buttons("Stat").Value = tbrUnpressed
                Call SwitchToState(False)
                Call SwitchToOften(, False)
            Else
                Call SwitchToOften
            End If
        End If
    ElseIf Button.Key = "Stat" Then
        Call SwitchToState
    ElseIf Button.Key = "New" Then
        Call NewOftenNew
    ElseIf Button.Key = "Del" Then
        Call OftenItemDel
    End If
End Sub

Private Sub NewOftenNew()
'���ܣ�����ǰ������Ŀ������˳�����Ŀ
    Dim arrSQL As Variant, i As Long
    Dim lngCol��Ŀ As Long, lngCol���� As Long
    
    arrSQL = Array()
    If Not fraStat.Visible Then
        If mrsItem.EOF Then Exit Sub
        
        ReDim arrSQL(0)
        arrSQL(0) = "ZL_���Ƹ�����Ŀ_Insert(" & UserInfo.ID & "," & mrsItem!������ĿID & ")"
    Else
        lngCol��Ŀ = GetCol("������ĿID")
        lngCol���� = GetCol("����")
        With vsItem
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) <> 0 And Val(.TextMatrix(i, lngCol��Ŀ)) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_���Ƹ�����Ŀ_Insert(" & UserInfo.ID & "," & Val(.TextMatrix(i, lngCol��Ŀ)) & "," & Val(.TextMatrix(i, lngCol����)) & ")"
                End If
            Next
        End With
        If UBound(arrSQL) < 0 Then
            MsgBox "������ѡ��һ��Ҫ����ĳ�����Ŀ��", vbInformation, gstrSysName
            vsItem.SetFocus: Exit Sub
        Else
            If MsgBox("�㵱ǰѡ���� " & UBound(arrSQL) + 1 & " ����Ŀ��Ҫ����Щ��Ŀ����Ϊ��ĸ��˳�����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
        
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    Screen.MousePointer = 0
    
    If Not fraStat.Visible Then
        MsgBox "��Ŀ""" & mrsItem!���� & """�Ѿ�������ĸ��˳�����Ŀ��", vbInformation, gstrSysName
    Else
        MsgBox "��ѡ�����Ŀ�Ѿ�������ĸ��˳�����Ŀ��", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCol(ByVal strName As String) As Long
    Dim i As Long
    For i = 1 To vsItem.Cols - 1
        If vsItem.TextMatrix(0, i) = strName Then
            GetCol = i: Exit Function
        End If
    Next
End Function

Private Sub OftenItemDel()
'���ܣ�����ǰ����������Ŀ�Ƴ�
    Dim strSQL As String, lngRow As Long
    
    If mrsItem.EOF Then Exit Sub
    If MsgBox("ȷʵҪ��""" & mrsItem!���� & """����ĸ�����Ŀ���Ƴ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    lngRow = vsItem.Row
    
    strSQL = "ZL_���Ƹ�����Ŀ_Delete(" & UserInfo.ID & "," & mrsItem!������ĿID & ")"
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With vsItem
        If lngRow = .FixedRows And .Rows = .FixedRows + 1 Then
            vsItem.Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
        Else
            .RemoveItem lngRow
            If lngRow <= .Rows - 1 Then
                .Row = lngRow
            Else
                .Row = .Rows - 1
            End If
        End If
        Call .ShowCell(.Row, .Col)
        Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstrPreNode Then Exit Sub
    '���ı�ʱ,���浱ǰ˳��(������)
    If Visible Then
        Call SaveColPosition(tvw_s.Nodes(mstrPreNode).Tag)
        Call SaveColWidth(tvw_s.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = Node.Key
    
    Call FillList(True)
End Sub

Private Function GetTreePath(ByVal objNode As Node) As String
'���ܣ���ȡ����·����
    Dim tmpNode As Node, strTmp As String
    Set tmpNode = objNode
    Do While Not tmpNode Is Nothing
        strTmp = IIF(InStr(tmpNode.Text, "[") > 0, NeedName(tmpNode.Text), Mid(tmpNode.Text, 3)) & "\" & strTmp
        Set tmpNode = tmpNode.Parent
    Loop
    GetTreePath = strTmp
End Function

Private Sub txtCount_GotFocus()
    Call zlControl.TxtSelAll(txtCount)
End Sub

Private Sub txtCount_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <> OldRow Then
        If NewRow >= vsItem.FixedRows Then
            mrsItem.Filter = "KeyID=" & Val(vsItem.TextMatrix(NewRow, 1))
            'ͳ�Ƴ�����Ŀ����ֱ��ѡ��,��Ϊû�й�Ȩ��,�ٳ�����
            cmdOK.Enabled = mrsItem.RecordCount = 1 And Not fraStat.Visible
        Else
            cmdOK.Enabled = False
        End If
        cmdOK.Visible = Not fraStat.Visible And mstr���� = ""
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    If Order = 0 Then Exit Sub
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '��Ϊ������˳��ı�,���Ա���ԭʼ�к�
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '�̶�
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    If vsItem.ColDataType(Col) = flexDTBoolean Then
        Order = 0
    Else
        'ǿ�Ʊ����а��ַ�������
        If vsItem.TextMatrix(0, Col) = "����" Then
            If Order = 1 Then Order = 7
            If Order = 2 Then Order = 8
        End If
    End If
End Sub

Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) = flexDTBoolean Then Cancel = True
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) <> flexDTBoolean Then
        Cancel = True
    ElseIf Val(vsItem.TextMatrix(Row, 1)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        If cmdOK.Enabled Then
            Call vsItem_KeyPress(13)
        ElseIf fraStat.Visible Then
            Call vsItem_KeyPress(32)
        End If
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then
            cmdOK_Click
        ElseIf fraStat.Visible Then
            If vsItem.Row + 1 <= vsItem.Rows - 1 Then
                vsItem.Row = vsItem.Row + 1
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    ElseIf KeyAscii = 32 Then
        KeyAscii = 0
        If fraStat.Visible Then
            If Val(vsItem.TextMatrix(vsItem.Row, 1)) <> 0 Then
                If Val(vsItem.TextMatrix(vsItem.Row, 2)) = 0 Then
                    vsItem.TextMatrix(vsItem.Row, 2) = 1
                Else
                    vsItem.TextMatrix(vsItem.Row, 2) = 0
                End If
            End If
        End If
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub

Private Sub SaveColPosition(Optional ByVal strType As String)
'���ܣ�������˳��:�к�,˳��|...
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", 1)) = 0 Then Exit Sub
    If tbrOften.Buttons("Stat").Value = tbrPressed Or fraStat.Visible Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr���� = "" And strType = "" And tvw_s.Visible And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub RestoreColPosition()
'���ܣ��ָ���˳��
'˵����Ӧ����������֮ǰ
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", 1)) = 0 Then Exit Sub
    
    With vsItem
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '�̶�
        strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,�ı������к�Ҳ�ı�
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .Cols - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .Cols - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'���ܣ������п��
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", 1)) = 0 Then Exit Sub
    If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'���ܣ��ָ��п��
'˵����Ӧ���ڻָ�����֮��
    Dim strType As String
    
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", 1)) = 0 Then Exit Sub
    
    If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColSort()
'���ܣ�������
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", 1)) <> 0 Then
            If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
            If tbrOften.Buttons("Often").Value = tbrPressed Then strType = "Often" '�̶�
            strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '��Ϊ���ܵ�����˳��,���Բ�����ʵ��������
                For i = 0 To .Cols - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .Cols - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(1).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(2).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function FillList(Optional ByVal blnClass As Boolean, _
    Optional str������ĿIDs As String, Optional str�շ�ϸĿIDs As String) As Boolean
'���ܣ����ݵ�ǰ��������װ��������ĿĿ¼
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, i As Long, j As Long
    Dim arrClass As Variant, strClass As String
    Dim strSub As String, str�������� As String
    Dim str�Ա� As String, strStock As String
    Dim strInput As String, lngҩ��ID As Long
    Dim blnLoad As Boolean, objTab As MSComctlLib.Tab
    Dim str��Χ As String, strҩƷ As String
    Dim blnOften As Boolean, blnStock As Boolean
    Dim str������� As String
    
    Dim lng����ID As Long, int���� As Integer, str��� As String

    str������ĿIDs = "": str�շ�ϸĿIDs = ""
    Set objNode = tvw_s.SelectedItem '����ΪNothing
    blnOften = tbrOften.Buttons("Often").Value = tbrPressed '�Ƿ���ʾ������Ŀ
    
    '�Ƿ���ʾ���ѡ��
    blnStock = mstr���� <> "" And tabClass.SelectedItem.Index = 1 _
        And Not blnOften And (gblnҩƷ�������ҽ�� Or mint��Ч = 1) _
        And Not (mlng��ҩ�� = 0 And mlng��ҩ�� = 0 And mlng��ҩ�� = 0)
    chkStock.Visible = blnStock
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '�����������ֶ�����
    '------------------------------------------------------------------------
    '������Ŀ�������Ա�
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    Else
        str�Ա� = "0"
    End If
    
    '������Ŀ�Ĳ�������
    str�������� = "Decode(A.���," & _
        "'H',Decode(A.��������,'1','����ȼ�','������')," & _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨','4','��ҩ�÷�',Null)," & _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ','7','����','8','����','9','����','10','��Σ','11','����',NULL)," & _
        "A.��������)"
    
    If mstr���� = "" Then
        int���� = Val(objNode.Tag): lng����ID = Val(Mid(objNode.Key, 2))
        If Not blnOften Then
            '�����еķ���ID
            If chkSub.Value = 1 Then
                '��ʾ�¼�����Ŀ
                If Val(Mid(objNode.Key, 2)) < 0 Then
                    strSub = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Where ����=[1])"
                Else
                    strSub = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Start With ID=[3] Connect by Prior ID=�ϼ�ID)"
                End If
            Else
                strSub = " And A.����ID=[3]"
            End If
            
            '�����е�����ȷ�����
            If Val(objNode.Tag) = 5 Then
                strSub = strSub & " And A.��� Not IN('4','5','6','7','8','9')"
            Else
                If Val(objNode.Tag) = 1 Then
                    str��� = "5"
                ElseIf Val(objNode.Tag) = 2 Then
                    str��� = "6"
                ElseIf Val(objNode.Tag) = 3 Then
                    str��� = "7"
                ElseIf Val(objNode.Tag) = 4 Then
                    str��� = "8"
                ElseIf Val(objNode.Tag) = 6 Then
                    str��� = "9"
                ElseIf Val(objNode.Tag) = 7 Then
                    str��� = "4"
                End If
                If str��� <> "" Then
                    strSub = strSub & " And A.���=[2]"
                End If
            End If
        ElseIf mlng����ID <> 0 Then 'ͨ��������ȷ���ķ�����������г�����Ŀ
            strSub = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Start With ID=[4] Connect by Prior ID=�ϼ�ID)"
        
            '�����е�����ȷ�����
            If Val(objNode.Tag) = 5 Then
                strSub = strSub & " And A.��� Not IN('4','5','6','7','8','9')"
            Else
                If Val(objNode.Tag) = 1 Then
                    str��� = "5"
                ElseIf Val(objNode.Tag) = 2 Then
                    str��� = "6"
                ElseIf Val(objNode.Tag) = 3 Then
                    str��� = "7"
                ElseIf Val(objNode.Tag) = 4 Then
                    str��� = "8"
                ElseIf Val(objNode.Tag) = 6 Then
                    str��� = "9"
                ElseIf Val(objNode.Tag) = 7 Then
                    str��� = "4"
                End If
                If str��� <> "" Then
                    strSub = strSub & " And A.���=[2]"
                End If
            End If
        Else
            '��ʾ���з���,����еĸ��˳�����Ŀ
        End If
    Else
        '����ƥ��:�޷�ȷ�����༰���,��������Ŀ��ƥ��
        If Len(mstr����) < 2 Then mstrLike = "" '�Ż�
        strInput = " And A.���<>'4' And (A.���� Like [5] And B.����=[7]" & _
            " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
    End If
    '���Ƭȷ�����
    If tabClass.SelectedItem.Key <> "" Then
        str��� = Mid(tabClass.SelectedItem.Key, 2)
        strSub = strSub & " And A.���=[2]"
    End If
    
    '��ȡ����
    '------------------------------------------------------------------------
    If Not (gblnҩƷ�������ҽ�� Or mint��Ч = 1) Then
        '����ҩƷȨ��
        strҩƷ = ""
        If InStr(mstrPrivs, "�´�����ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And C.�������<>'����ҩ'"
        End If
        If InStr(mstrPrivs, "�´ﶾ��ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And C.�������<>'����ҩ'"
        End If
        If InStr(mstrPrivs, "�´����ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And C.��ֵ���� Not IN('����','����')"
        End If
        
        'ҩƷ������Ŀ����:��������ҩƷ����ʱ�Ŷ�ȡ
        blnLoad = False
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            If mstr���� <> "" Then
                strInput = " And A.���<>'4' And (A.���� Like [5] And B.����=[7]" & _
                    " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
                If IsNumeric(mstr����) Then
                    '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���<>'4' And (A.���� Like [5] And B.����=[7] Or B.���� Like [6] And B.����=3)"
                ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                    'X1.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr����) Then
                    '��������,��ֻƥ������
                    strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                End If
                
                strSQL = _
                    " Select " & IIF(Not mbln����, "Distinct", "") & _
                        " A.��� As ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID," & _
                        " D.���� As ���,A.����,B.����," & IIF(mbln����, "B.����,", "") & _
                        " A.���㵥λ,A.�걾��λ,C.ҩƷ����," & str�������� & " As ��Ŀ����," & _
                        " C.����ְ�� as ����ְ��ID" & IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ҩƷ���� C,������Ŀ��� D,������Ŀ���� B,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.ID=B.������ĿID And A.ID=C.ҩ��ID And A.���=D.���� And A.��� IN ('5','6','7')" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strInput & strSub & strҩƷ & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            Else
                strSQL = _
                    " Select " & _
                        " A.��� As ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID," & _
                        " D.���� As ���,A.����,A.����,A.���㵥λ,A.�걾��λ," & _
                        " C.ҩƷ����," & str�������� & " As ��Ŀ����,C.����ְ�� as ����ְ��ID" & _
                        IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ҩƷ���� C,������Ŀ��� D,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.ID=C.ҩ��ID And A.���=D.���� And A.��� IN ('5','6','7')" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strSub & strҩƷ & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            End If
        End If
        
        '��ҩƷ������Ŀ����:���಻��ҩƷ����ʱ���ض�ȡ
        blnLoad = False
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) = 0
        End If
        If blnLoad Then
            If mstr���� <> "" Then
                strInput = " And A.���<>'4' And (A.���� Like [5] Or B.���� Like [6] Or B.���� Like [6]) And B.����=[7]"
                If IsNumeric(mstr����) Then
                    '1X.����ȫ������ʱֻƥ�����
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���<>'4' And A.���� Like [5] And B.����=[7]"
                ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                    'X1.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr����) Then
                    '��������,��ֻƥ������
                    strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                End If
            
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln����, "Distinct", "") & _
                        " A.��� As ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID," & _
                        " D.���� As ���,A.����,B.����," & IIF(mbln����, "B.����,", "") & _
                        " A.���㵥λ,A.�걾��λ,NULL as ҩƷ����," & str�������� & " As ��Ŀ����," & _
                        " Null as ����ְ��ID" & IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ������Ŀ��� D,������Ŀ���� B,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.ID=B.������ĿID And A.���=D.���� And A.��� Not IN('5','6','7')" & _
                        " And (A.���<>'9' Or A.��ԱID=[11] Or A.��ԱID is Null)" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strInput & strSub & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            Else
                '�����η�����ѡ��ķ�ҩƷ������Ŀ,��������ҩƷͬʱ��ʾ,��˿��Լ���һЩ�ֶ�
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & _
                        " A.��� As ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID,D.���� As ���," & _
                        " A.����,A.����,A.���㵥λ,A.�걾��λ" & IIF(blnOften, ",Null as ҩƷ����", "") & "," & _
                        str�������� & " As ��Ŀ����,Null as ����ְ��ID" & IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ������Ŀ��� D,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.���=D.���� And A.��� Not IN('5','6','7')" & _
                        " And (A.���<>'9' Or A.��ԱID=[11] Or A.��ԱID is Null)" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strSub & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            End If
        End If
    Else
        'ҩƷ���,ĳһ��ҩ��δָ��ʱ,����������¼
        If blnOften Then
            strStock = "" 'û�з���,�޷�ȷ��ҩƷ���,��������ķ�ʽ����
        Else
            If mstr���� = "" Then
                '���ݷ������ȷ��ҩƷ���
                If Val(objNode.Tag) = 1 Then
                    lngҩ��ID = mlng��ҩ��
                ElseIf Val(objNode.Tag) = 2 Then
                    lngҩ��ID = mlng��ҩ��
                ElseIf Val(objNode.Tag) = 3 Then
                    lngҩ��ID = mlng��ҩ��
                End If
            Else
                'û�з���,�޷�ȷ��ҩƷ���
                If Mid(tabClass.SelectedItem.Key, 2) = "5" Then
                    lngҩ��ID = mlng��ҩ��
                ElseIf Mid(tabClass.SelectedItem.Key, 2) = "6" Then
                    lngҩ��ID = mlng��ҩ��
                ElseIf Mid(tabClass.SelectedItem.Key, 2) = "7" Then
                    lngҩ��ID = mlng��ҩ��
                End If
            End If
            If lngҩ��ID <> 0 Then
                strStock = _
                    "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ��� From ҩƷ��� A" & _
                    " Where A.���� = 1 And A.�ⷿID=[12]" & _
                    " And (Nvl(A.����, 0) = 0 Or A.Ч�� Is Null Or A.Ч�� > Trunc(Sysdate))" & _
                    " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
            ElseIf blnStock And chkStock.Value = 1 Then
                strStock = _
                    "Select A.ҩƷID,Sum(Nvl(A.��������,0)) as ���" & _
                    " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
                    " Where A.���� = 1 And (Nvl(A.����,0)=0 Or A.Ч�� Is Null Or A.Ч��>Trunc(Sysdate))" & _
                        " And A.�ⷿID=Decode(B.���,'5',[13],'6',[14],'7',[15],Null)" & _
                        " And A.ҩƷID=B.ID And B.��� IN('5','6','7')" & _
                    " Group by A.ҩƷID Having Sum(Nvl(A.��������,0))<>0"
                'strStock = "" '�Ż�
            End If
        End If
        
        '����ҩƷȨ��
        strҩƷ = ""
        If InStr(mstrPrivs, "�´�����ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And D.�������<>'����ҩ'"
        End If
        If InStr(mstrPrivs, "�´ﶾ��ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And D.�������<>'����ҩ'"
        End If
        If InStr(mstrPrivs, "�´����ҩ��") = 0 Then
            strҩƷ = strҩƷ & " And D.��ֵ���� Not IN('����','����')"
        End If
        
        'ҩƷ��񲿷�:��������ҩƷ����ʱ�Ŷ�ȡ
        blnLoad = False
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) > 0
        End If
        If blnLoad Then
            str��Χ = IIF(mint��Χ = 1, "����", "סԺ")
                
            '�Ƿ�����п��:ָ��ҩ��ʱ����ϵͳ�����Ƿ�Ҫ���ƿ��
            str������� = " And A.ID=X.ҩƷID(+)"
            If gblnStock Then
                str������� = str������� & " And (" & _
                    " A.���='5' And ([13]=0 Or X.ҩƷID Is Not Null)" & _
                    " Or A.���='6' And ([14]=0 Or X.ҩƷID Is Not Null)" & _
                    " Or A.���='7' And ([15]=0 Or X.ҩƷID Is Not Null)" & _
                    ")"
            End If
                
            If mstr���� <> "" Then
                strInput = " And A.���<>'4' And (A.���� Like [5] And B.����=[7]" & _
                    " Or B.���� Like [6] And B.����=[7] Or B.���� Like [6] And B.���� IN([7],3))"
                If IsNumeric(mstr����) Then
                    '1X.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���<>'4' And (A.���� Like [5] And B.����=[7] Or B.���� Like [6] And B.����=3)"
                ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                    'X1.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr����) Then
                    '��������,��ֻƥ������
                    strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                End If
            
                '���Ƹ��������ƥ����ʾ
                strSQL = "Select " & IIF(Not mbln����, "Distinct", "") & _
                    " A.ID,A.���,A.����,B.����," & IIF(mbln����, "B.����,", "") & _
                    " A.���,A.����,A.��������,A.˵��,A.�Ƿ���" & IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From �շ���Ŀ���� B,�շ���ĿĿ¼ A" & IIF(blnOften, ",ҩƷ��� Q,���Ƹ�����Ŀ R", "") & _
                    " Where A.ID=B.�շ�ϸĿID And A.��� IN ('5','6','7')" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And A.������� IN([8],3)" & strInput & _
                    IIF(blnOften, " And Q.ҩƷID=A.ID And R.������ĿID=Q.ҩ��ID And R.��ԱID=[11]", "")
                If mbln�۸� Then
                    strSQL = "Select A.ID,A.���,A.����,A.����," & IIF(mbln����, "A.����,", "") & _
                        " A.���,A.����,A.��������,A.˵��,Sum(Decode(A.�Ƿ���,1,NULL,B.�ּ�)) as �۸�" & IIF(blnOften, ",A.Ƶ��ID", "") & _
                        " From �շѼ�Ŀ B,(" & strSQL & ") A" & _
                        " Where A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Group by A.ID,A.���,A.����,A.����," & IIF(mbln����, "A.����,", "") & _
                        "A.���,A.����,A.��������,A.˵��" & IIF(blnOften, ",A.Ƶ��ID", "")
                ElseIf mstrLike = "" And strStock <> "" Then
                    '���������ü�������ʱ(����ƥ��),�����(+)����(ҩƷ���),����ҪGroup Byһ��(���)
                    '��Group by ��Distinct ͬʱ����ʱ(Not mbln����)��Oracle��ֻѡ�����Group by
                    strSQL = strSQL & " Group By A.ID,A.���,A.����,B.����," & IIF(mbln����, "B.����,", "") & _
                        " A.���,A.����,A.��������,A.˵��,A.�Ƿ���" & IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "")
                End If
                
                strSQL = _
                    " Select " & _
                        " A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                        " F.���� AS ���,A.����,A.����," & IIF(mbln����, "A.����,", "") & _
                        " E.���㵥λ,A.���,A.����,D.ҩƷ����,NULL AS ��Ŀ����,A.��������,A.˵��,D.����ְ�� as ����ְ��ID" & _
                        IIF(mbln�۸�, ",Decode(A.�۸�,NULL,NULL,LTrim(To_Char(A.�۸�*C." & str��Χ & "��װ,'999990.0000'))||'/'||C." & str��Χ & "��λ) as �۸�", "") & _
                        IIF(strStock <> "", ",Decode(X.���,NULL,NULL,LTrim(To_Char(X.���/C." & str��Χ & "��װ,'999990.0000'))||C." & str��Χ & "��λ) as ���", "") & _
                        IIF(blnOften, ",A.Ƶ��ID", "") & _
                    " From ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,(" & strSQL & ") A" & _
                        IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                    " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID" & _
                        " And A.���=F.���� And E.��� IN('5','6','7')" & _
                        " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                        " And E.������� IN([8],3) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                        IIF(strStock <> "", str�������, "") & strҩƷ & Replace(strSub, "A.", "E.")
            Else
                '��ҩ���Ƹ��ݲ���������ʾ
                If mbln�۸� Then
                    strSQL = _
                        " Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                            " F.���� AS ���,A.����,Nvl(G.����,A.����) as ����,E.���㵥λ,A.���,A.����,D.ҩƷ����,NULL AS ��Ŀ����,A.��������,A.˵��,D.����ְ�� as ����ְ��ID," & _
                            " Decode(A.�Ƿ���,1,NULL,LTrim(To_Char(Sum(B.�ּ�)*C." & str��Χ & "��װ,'999990.0000'))||'/'||C." & str��Χ & "��λ) as �۸�" & _
                            IIF(strStock <> "", ",Decode(X.���,NULL,NULL,LTrim(To_Char(X.���/C." & str��Χ & "��װ,'999990.0000'))||C." & str��Χ & "��λ) as ���", "") & _
                            IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                        " From �շѼ�Ŀ B,�շ���ĿĿ¼ A,ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,�շ���Ŀ���� G" & _
                            IIF(blnOften, ",���Ƹ�����Ŀ R", "") & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID" & _
                            " And A.���=F.���� And E.��� IN('5','6','7')" & _
                            " And A.ID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
                            " And E.������� IN([8],3) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                            " And A.��� IN ('5','6','7') And A.������� IN([8],3)" & _
                            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                            " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                            IIF(blnOften, " And R.������ĿID=E.ID And R.��ԱID=[11]", "") & _
                            IIF(strStock <> "", str�������, "") & strҩƷ & Replace(strSub, "A.", "E.") & _
                            " And A.ID=B.�շ�ϸĿID And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " Group by A.���,E.ID,A.ID,F.����,A.����,Nvl(G.����,A.����),E.���㵥λ,A.���,A.����,D.ҩƷ����,A.��������,A.˵��,D.����ְ��,A.�Ƿ���," & _
                            "C." & str��Χ & "��װ,C." & str��Χ & "��λ" & IIF(strStock <> "", ",X.���", "") & IIF(blnOften, ",Nvl(R.Ƶ��,1)", "")
                Else
                    strSQL = _
                        " Select A.��� AS ���ID,E.ID as ������ĿID,A.ID as �շ�ϸĿID," & _
                            " F.���� AS ���,A.����,Nvl(G.����,A.����) as ����,E.���㵥λ,A.���,A.����,D.ҩƷ����,NULL AS ��Ŀ����,A.��������,A.˵��,D.����ְ�� as ����ְ��ID" & _
                            IIF(strStock <> "", ",Decode(X.���,NULL,NULL,LTrim(To_Char(X.���/C." & str��Χ & "��װ,'999990.0000'))||C." & str��Χ & "��λ) as ���", "") & _
                            IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                        " From �շ���ĿĿ¼ A,ҩƷ��� C,ҩƷ���� D,������ĿĿ¼ E,�շ���Ŀ��� F,�շ���Ŀ���� G" & _
                            IIF(blnOften, ",���Ƹ�����Ŀ R", "") & IIF(strStock <> "", ",(" & strStock & ") X", "") & _
                        " Where A.ID=C.ҩƷID And C.ҩ��ID=D.ҩ��ID And D.ҩ��ID=E.ID" & _
                            " And A.���=F.���� And E.��� IN('5','6','7')" & _
                            " And A.ID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
                            " And E.������� IN([8],3) And Nvl(E.ִ��Ƶ��,0) IN(0,[9])" & _
                            " And A.��� IN ('5','6','7') And A.������� IN([8],3)" & _
                            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                            " And (E.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or E.����ʱ�� IS NULL)" & _
                            IIF(blnOften, " And R.������ĿID=E.ID And R.��ԱID=[11]", "") & _
                            IIF(strStock <> "", str�������, "") & strҩƷ & Replace(strSub, "A.", "E.")
                End If
            End If
        End If
        
        '��ҩƷ������Ŀ����:���಻��ҩƷ����ʱ���ض�ȡ
        blnLoad = False
        If mstr���� <> "" Or (blnOften And mlng����ID = 0) Then
            blnLoad = True
        Else
            blnLoad = InStr(",1,2,3,", Val(objNode.Tag)) = 0
        End If
        If blnLoad Then
            If mstr���� <> "" Then
                strInput = " And A.���<>'4' And (A.���� Like [5] Or B.���� Like [6] Or B.���� Like [6]) And B.����=[7]"
                If IsNumeric(mstr����) Then
                    '1X.����ȫ������ʱֻƥ�����
                    If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���<>'4' And A.���� Like [5] And B.����=[7]"
                ElseIf zlCommFun.IsCharAlpha(mstr����) Then
                    'X1.����ȫ����ĸʱֻƥ�����
                    If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                ElseIf zlCommFun.IsCharChinese(mstr����) Then
                    '��������,��ֻƥ������
                    strInput = " And A.���<>'4' And B.���� Like [6] And B.����=[7]"
                End If
            
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & IIF(Not mbln����, "Distinct", "") & _
                        " A.��� As ���ID,A.ID as ������ĿID,-NULL as �շ�ϸĿID," & _
                        " D.���� As ���,A.����,B.����," & IIF(mbln����, "B.����,", "") & _
                        " A.���㵥λ,A.�걾��λ as ���,NULL AS ����,NULL as ҩƷ����," & _
                        str�������� & " As ��Ŀ����,NULL as ��������,Null as ˵��,Null as ����ְ��ID" & _
                        IIF(mbln�۸�, ",NULL as �۸�", "") & IIF(strStock <> "", ",NULL As ���", "") & _
                        IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ������Ŀ��� D,������Ŀ���� B,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.ID=B.������ĿID And A.���=D.���� And A.��� Not IN('5','6','7')" & _
                        " And (A.���<>'9' Or A.��ԱID=[11] Or A.��ԱID is Null)" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strInput & strSub & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            Else
                '�����η�����ѡ��ķ�ҩƷ������Ŀ,��������ҩƷͬʱ��ʾ,��˿��Լ���һЩ�ֶ�
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                    " Select " & _
                        " A.��� as ���ID,A.ID as ������ĿID,-NULL as �շ�ϸĿID,D.���� as ���," & _
                        " A.����,A.����,A.���㵥λ,A.�걾��λ as ���" & IIF(blnOften, ",Null as ����,Null as ҩƷ����", "") & "," & _
                        str�������� & " as ��Ŀ����" & IIF(blnOften, ",Null as ��������,Null as ˵��,Null as ����ְ��ID" & _
                        IIF(mbln�۸�, ",NULL as �۸�", "") & IIF(strStock <> "", ",Null as ���", ""), "") & _
                        IIF(blnOften, ",Nvl(R.Ƶ��,1) as Ƶ��ID", "") & _
                    " From ������Ŀ��� D,������ĿĿ¼ A" & IIF(blnOften, ",���Ƹ�����Ŀ R", "") & _
                    " Where A.���=D.���� And A.��� Not IN('5','6','7')" & _
                        " And (A.���<>'9' Or A.��ԱID=[11] Or A.��ԱID is Null)" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                        " And A.������� IN([8],3) And Nvl(A.����Ӧ��,0)=1 And Instr([10],','||Nvl(A.�����Ա�,0)||',')>0" & _
                        " And Nvl(A.ִ��Ƶ��,0) IN(0,[9])" & strSub & _
                        IIF(blnOften, " And R.������ĿID=A.ID And R.��ԱID=[11]", "")
            End If
        End If
    End If
    If blnOften Then
        strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by Ƶ��ID Desc,Decode(���ID,'4','I',���ID),���,����" '�����,���ڻ���֮��
    Else
        strSQL = "Select Rownum as KeyID,A.* From (" & strSQL & ") A Order by Decode(���ID,'4','I',���ID),���,����" '�����,���ڻ���֮��
    End If
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int����, str���, lng����ID, mlng����ID, _
        UCase(mstr����) & "%", mstrLike & UCase(mstr����) & "%", mint���� + 1, mint��Χ, _
        IIF(mint��Ч = 0, 2, 1), "," & str�Ա� & ",", UserInfo.ID, lngҩ��ID, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��)
    
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        If InStr(",���,�۸�,", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '�ָ���˳��:Ӧ����������֮ǰ
    Call RestoreColPosition
    Call RestoreColWidth
    '������:������,�Ա���洦���к�
    Call RestoreColSort
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '�ռ����Ƭ��Ϣ
        If InStr(strClass & ",", "," & mrsItem!���ID & mrsItem!��� & ",") = 0 Then
            strClass = strClass & "," & mrsItem!���ID & mrsItem!���
        End If
        
        '�ռ���ĿID:ֻ�ռ����2��
        If mstr���� <> "" Then
            If UBound(Split(str������ĿIDs, ",")) < 2 Then
                If InStr(str������ĿIDs & ",", "," & mrsItem!������ĿID & ",") = 0 Then
                    str������ĿIDs = str������ĿIDs & "," & mrsItem!������ĿID
                End If
            End If
            If UBound(Split(str�շ�ϸĿIDs, ",")) < 2 Then
                If Not IsNull(mrsItem!�շ�ϸĿID) Then
                    If InStr(str�շ�ϸĿIDs & ",", "," & mrsItem!�շ�ϸĿID & ",") = 0 Then
                        str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & mrsItem!�շ�ϸĿID
                    End If
                End If
            End If
        End If
        mrsItem.MoveNext
    Next
    
    '�������࿨Ƭ:�ж���ʱ����Ŀ���϶�ʱ
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '��Alt��ݼ������޷�����
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    If blnOften Then
        lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ������ " & mrsItem.RecordCount & " ����Ŀ"
    Else
        lblInfo.Caption = "��ǰѡ��" & GetTreePath(tvw_s.SelectedItem) & tabClass.SelectedItem.Tag & "������ " & mrsItem.RecordCount & " ����Ŀ"
    End If
    
    vsItem.FrozenCols = 0
    vsItem.Editable = flexEDNone
    vsItem.SheetBorder = vsItem.BackColor
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillList = True
    Exit Function
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function

Private Function FillStat(Optional ByVal blnClass As Boolean) As Boolean
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
    Dim objNode As Node, objItem As ListItem
    Dim strSQL As String, i As Long, j As Long
    Dim arrClass As Variant, strClass As String
    Dim str���� As String, str��� As String
    Dim str�������� As String, strѡ����� As String
    Dim objTab As MSComctlLib.Tab

    Set objNode = tvw_s.SelectedItem '����ΪNothing
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Editable = flexEDKbdMouse
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '�����������ֶ�����
    '------------------------------------------------------------------------
    '������Ŀ�Ĳ�������
    str�������� = "Decode(A.���," & _
        "'H',Decode(A.��������,'1','����ȼ�','������')," & _
        "'E',Decode(A.��������,'1','��������','2','��ҩ;��','3','��ҩ�巨','4','��ҩ�÷�',Null)," & _
        "'Z',Decode(A.��������,'1','����','2','סԺ','3','ת��','4','����','5','��Ժ','6','תԺ',NULL)," & _
        "A.��������)"
    
    'ֻͳ�Ƹ÷�������ĳ�����Ŀ
    If mlng����ID <> 0 Then
        str���� = " And A.����ID IN(Select ID From ���Ʒ���Ŀ¼ Start With ID=[1] Connect by Prior ID=�ϼ�ID)"
        
        '�����е�����ȷ�����
        If Val(objNode.Tag) = 5 Then
            str��� = str��� & " And A.������� Not IN('4','5','6','7','8','9')"
        Else
            If Val(objNode.Tag) = 1 Then
                strѡ����� = "5"
            ElseIf Val(objNode.Tag) = 2 Then
                strѡ����� = "6"
            ElseIf Val(objNode.Tag) = 3 Then
                strѡ����� = "7"
            ElseIf Val(objNode.Tag) = 4 Then
                strѡ����� = "8"
            ElseIf Val(objNode.Tag) = 6 Then
                strѡ����� = "9"
            ElseIf Val(objNode.Tag) = 7 Then
                strѡ����� = "4"
            End If
            If strѡ����� <> "" Then
                str��� = str��� & " And A.�������=[2]"
            End If
        End If
    End If

    '���Ƭȷ�����
    If tabClass.SelectedItem.Key <> "" Then
        strѡ����� = Mid(tabClass.SelectedItem.Key, 2)
        str��� = str��� & " And A.�������=[2]"
    End If
    
    '��ȡ����:û�����Ƴ���/����Ӧ��,�������,�Ա�Χ,ҩƷȨ��;��ҩȱʡ���ܵ���Ӧ��
    '------------------------------------------------------------------------------
    strSQL = _
        " Select A.������ĿID,Count(A.������ĿID) As ����" & _
        " From ����ҽ����¼ A,����ҽ��״̬ B" & _
        " Where A.ID=B.ҽ��ID And B.��������=1" & str��� & _
        "   And B.����ʱ��>=[3] And B.������Ա=[4]" & _
        " Group By A.������ĿID"
    If MovedByDate(dtpDate.Value) Then
        strSQL = strSQL & " Union ALL " & _
            " Select A.������ĿID,Count(A.������ĿID) As ����" & _
            " From H����ҽ����¼ A,H����ҽ��״̬ B" & _
            " Where A.ID=B.ҽ��ID And B.��������=1" & str��� & _
            "   And B.����ʱ��>=[3] And B.������Ա=[4]" & _
            " Group By A.������ĿID"
    End If
    
    If InStr(str���, "='5'") > 0 Or InStr(str���, "='6'") > 0 Or InStr(str���, "='7'") > 0 Then
        'ҩƷ������Ŀ����
        strSQL = _
            " Select A.��� as ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID," & _
            " D.���� as ���,A.����,A.����,A.���㵥λ,C.ҩƷ����,B.����" & _
            " From ������Ŀ��� D,ҩƷ���� C,������ĿĿ¼ A,(" & strSQL & ") B" & _
            " Where A.���=D.���� And A.ID=C.ҩ��ID And A.ID=B.������ĿID" & _
            "   And Not (A.���='E' And Nvl(A.��������,'0')<>'0')" & str���� & _
            "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            "   And Nvl(A.�������,0)<>0 And (Nvl(A.����Ӧ��,0)=1 Or A.���='7')"
        If chkAll.Value = 0 Then
            strSQL = _
                "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,ҩƷ����,-1*���� as ���� From (" & strSQL & ")" & _
                " Group by -1*����,���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,ҩƷ����"
            strSQL = "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,ҩƷ����,Abs(����) as ���� From (" & strSQL & ") Where Rownum<=[5]"
        End If
    Else
        '��ҩƷ���ݻ����в���
        strSQL = _
            " Select A.��� as ���ID,A.ID as ������ĿID,NULL as �շ�ϸĿID," & _
            " D.���� as ���,A.����,A.����,A.���㵥λ,A.�걾��λ," & str�������� & " As ��Ŀ����,B.����" & _
            " From ������Ŀ��� D,������ĿĿ¼ A,(" & strSQL & ") B" & _
            " Where A.���=D.���� And A.ID=B.������ĿID" & str���� & _
            "   And Not (A.���='E' And Nvl(A.��������,'0')<>'0')" & _
            "   And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            "   And Nvl(A.�������,0)<>0 And (Nvl(A.����Ӧ��,0)=1 Or A.���='7')"
        If chkAll.Value = 0 Then
            strSQL = _
                "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,�걾��λ,��Ŀ����,-1*���� as ���� From (" & strSQL & ")" & _
                " Group by -1*����,���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,�걾��λ,��Ŀ����"
            strSQL = "Select ���ID,������ĿID,�շ�ϸĿID,���,����,����,���㵥λ,�걾��λ,��Ŀ����,Abs(����) as ���� From (" & strSQL & ") Where Rownum<=[5]"
        End If
    End If
    strSQL = "Select Rownum as KeyID,Null as ѡ��,A.* From (" & strSQL & ") A Order by ���� Desc,����"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    'Set mrsItem = New ADODB.Recordset
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, strѡ�����, CDate(Format(dtpDate.Value, "yyyy-MM-dd")), UserInfo.����, Val(txtCount.Text))
    
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '����ͳ�Ƴ�����Ŀʱ����Ϊ��0��0��
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColAlignment(i) = 1
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '�ռ����Ƭ��Ϣ
        If InStr(strClass & ",", "," & mrsItem!���ID & mrsItem!��� & ",") = 0 Then
            strClass = strClass & "," & mrsItem!���ID & mrsItem!���
        End If
        mrsItem.MoveNext
    Next
    
    '�������࿨Ƭ:�ж���ʱ����Ŀ���϶�ʱ
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '��Alt��ݼ������޷�����
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    lblInfo.Caption = "��ǰѡ��""" & UserInfo.���� & """�ĸ��˳�����Ŀ������ " & mrsItem.RecordCount & " ����Ŀ"
    
    vsItem.TextMatrix(0, 2) = ""
    vsItem.ColWidth(2) = 300
    vsItem.ColDataType(2) = flexDTBoolean
    vsItem.FrozenCols = 2
    vsItem.Editable = flexEDKbdMouse
    vsItem.SheetBorder = vbBlack
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    Screen.MousePointer = 0
    FillStat = True
    Exit Function
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function
