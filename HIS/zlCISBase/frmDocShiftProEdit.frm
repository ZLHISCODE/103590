VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmDocShiftProEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ҽ�����Ӱಡ����Ŀ-����"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8880
   Icon            =   "frmDocShiftProEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   8880
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSource 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3840
      Width           =   7935
      Begin VB.CheckBox chkOnlyRead 
         Caption         =   "��Ŀֻ��"
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   0
         Width           =   1095
      End
      Begin VB.ComboBox cboResource 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblResourse 
         AutoSize        =   -1  'True
         Caption         =   "��ȡ��Դ"
         Height          =   180
         Left            =   360
         TabIndex        =   36
         Top             =   45
         Width           =   720
      End
   End
   Begin VB.Frame fraMed 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   6720
      Width           =   8535
      Begin VB.ComboBox cboMed 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1200
         TabIndex        =   38
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtDiagn 
         Height          =   300
         Left            =   5040
         TabIndex        =   12
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblMed 
         AutoSize        =   -1  'True
         Caption         =   "������������"
         Height          =   180
         Left            =   0
         TabIndex        =   34
         Top             =   45
         Width           =   1080
      End
      Begin VB.Label lblDiagn 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Left            =   4440
         TabIndex        =   33
         Top             =   45
         Width           =   540
      End
   End
   Begin VB.Frame fra3 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   30
      Top             =   7200
      Width           =   8655
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6960
         TabIndex        =   16
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "������˳�(&S)"
         Height          =   350
         Left            =   5520
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "��������������������������ظ���Ŀ"
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   0
         Width           =   3855
      End
      Begin VB.Frame fraLine 
         Height          =   50
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   10935
      End
      Begin VB.CommandButton cmdSaveNew 
         Caption         =   "�����������Ŀ(&N)"
         Height          =   350
         Left            =   3600
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame fraSQL 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   8655
      Begin XtremeSyntaxEdit.SyntaxEdit SynSQL 
         Height          =   1935
         Left            =   1200
         TabIndex        =   10
         Top             =   0
         Width           =   7095
         _Version        =   983043
         _ExtentX        =   12515
         _ExtentY        =   3413
         _StockProps     =   84
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         EnableSyntaxColorization=   -1  'True
         ShowLineNumbers =   0   'False
         ShowSelectionMargin=   0   'False
         ShowScrollBarVert=   -1  'True
         ShowScrollBarHorz=   -1  'True
         EnableVirtualSpace=   0   'False
         EnableAutoIndent=   -1  'True
         ShowWhiteSpace  =   0   'False
         ShowCollapsibleNodes=   -1  'True
         AutoCompleteWndWidth=   160
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "��֤(&C)"
         Height          =   350
         Left            =   7200
         TabIndex        =   11
         Top             =   2040
         Width           =   1100
      End
      Begin VB.Label lblSQL 
         Caption         =   "ֻ������һ���ı����͵��ֶΣ���ʹ��[����ID],[��ҳID],[��ʼʱ��],[����ʱ��]��Ϊ����"
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   2040
         Width           =   5775
      End
      Begin VB.Label lblInSQL 
         AutoSize        =   -1  'True
         Caption         =   "��ȡSQL"
         Height          =   180
         Left            =   450
         TabIndex        =   29
         Top             =   0
         Width           =   630
      End
   End
   Begin MSComCtl2.UpDown UpDownRow 
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      Max             =   1000
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDescript 
      Height          =   300
      Left            =   1320
      TabIndex        =   2
      Top             =   1000
      Width           =   4815
   End
   Begin VB.TextBox txtRow 
      Height          =   300
      Left            =   5400
      TabIndex        =   6
      Text            =   "1"
      Top             =   1845
      Width           =   495
   End
   Begin VB.ComboBox cboPrintFormat 
      Height          =   300
      Left            =   5400
      TabIndex        =   4
      Text            =   "cboPrintFormat"
      Top             =   1425
      Width           =   3015
   End
   Begin VB.ComboBox cboPrintType 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1425
      Width           =   2415
   End
   Begin VB.ComboBox cboPrintForm 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1845
      Width           =   2415
   End
   Begin VB.TextBox txtPrjName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.ComboBox cblPrjType 
      Height          =   300
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfRange 
      Height          =   1455
      Left            =   1320
      TabIndex        =   7
      Top             =   2280
      Width           =   7095
      _cx             =   12515
      _cy             =   2566
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDocShiftProEdit.frx":5C02
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
      Editable        =   2
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   8280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":5C92
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":622C
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":67C6
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":D028
            Key             =   "add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":1388A
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":1429C
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":14CAE
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocShiftProEdit.frx":156C0
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNames 
      AutoSize        =   -1  'True
      Caption         =   "��Ѫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1440
      TabIndex        =   27
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "�������ͼ��"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label lblPrjName 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ����"
      Height          =   180
      Left            =   480
      TabIndex        =   25
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblPrjType 
      AutoSize        =   -1  'True
      Caption         =   "��Ŀ���"
      Height          =   180
      Left            =   4560
      TabIndex        =   24
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblPrintForm 
      AutoSize        =   -1  'True
      Caption         =   "������ʽ"
      Height          =   180
      Left            =   480
      TabIndex        =   23
      Top             =   1905
      Width           =   720
   End
   Begin VB.Label lblPrintType 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   480
      TabIndex        =   22
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label lblPrintFormat 
      AutoSize        =   -1  'True
      Caption         =   "�����ʽ"
      Height          =   180
      Left            =   4560
      TabIndex        =   21
      Top             =   1470
      Width           =   720
   End
   Begin VB.Label lblPrintRange 
      AutoSize        =   -1  'True
      Caption         =   "����ֵ��"
      Height          =   180
      Left            =   480
      TabIndex        =   20
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblProntRow 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   4560
      TabIndex        =   19
      Top             =   1905
      Width           =   720
   End
   Begin VB.Label lblDes 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   480
      TabIndex        =   17
      Top             =   1050
      Width           =   720
   End
End
Attribute VB_Name = "frmDocShiftProEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrSName As String
Private mstrPatiPrj As String
Private mlngNum As Long
Private mblnOK As Boolean

Public Function ShowMe(ByVal bytType As Byte, ByVal strSName As String, ByVal strPatiPrj As String) As Boolean
'bytType:1-������2-�޸�
    
    mbytType = bytType
    mstrSName = strSName
    mstrPatiPrj = strPatiPrj
    mblnOK = False
    
    Me.Show 1
    ShowMe = mblnOK
End Function

Private Sub LoadData()
'���ؽ�������
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim varTemp As Variant
    Dim i As Long
    
    Select Case mbytType
        Case 1
            Me.Caption = "ҽ�����Ӱಡ����Ŀ-����"
            cblPrjType.ListIndex = 0
            cboPrintForm.ListIndex = 0
            cboPrintType.ListIndex = 0
            cboResource.ListIndex = 0
            cmdSave.Caption = "������˳�(&S)"
            cmdSaveNew.Visible = True
            lblNames.Caption = mstrSName
        Case 2
            Me.Caption = "ҽ�����Ӱಡ����Ŀ-�޸�"
            cmdSave.Caption = "����(&S)"
            cmdSaveNew.Visible = False
            Set rsTemp = GetPatiTypeInfo(mstrSName, mstrPatiPrj)
            lblNames.Caption = mstrSName
            txtPrjName.Text = mstrPatiPrj
            If rsTemp.RecordCount = 1 Then
                mlngNum = rsTemp!���
                For i = 0 To cblPrjType.ListCount - 1
                    If cblPrjType.List(i) = rsTemp!��Ŀ��� Then
                        cblPrjType.ListIndex = i
                    End If
                Next
                txtDescript.Text = rsTemp!�������� & ""
                cboPrintForm.ListIndex = Val(rsTemp!������ʽ) - 1
                cboPrintType.ListIndex = Val(rsTemp!��������)
                For i = 0 To cboPrintFormat.ListCount - 1
                    If cboPrintFormat.List(i) = rsTemp!�����ʽ Then
                        cboPrintFormat.ListIndex = i
                    End If
                Next
                strTemp = rsTemp!����ֵ�� & ""
                If Val(rsTemp!������ʽ) <> 1 Then
                    varTemp = Split(strTemp, ",")
                    With vsfRange
                        .Rows = 1
                        .Rows = UBound(varTemp) + 2
                        For i = 0 To UBound(varTemp)
                            If Mid(varTemp(i), 1, 1) = "*" Then
                                .Cell(flexcpChecked, i + 1, .ColIndex("�ı���")) = flexChecked
                                .TextMatrix(i + 1, .ColIndex("ֵ��")) = Mid(varTemp(i), 2)
                            Else
                                .TextMatrix(i + 1, .ColIndex("ֵ��")) = varTemp(i)
                            End If
                        Next
                    End With
                End If
                For i = 0 To cboResource.ListCount - 1
                    If cboResource.List(i) = rsTemp!��ȡ��Դ & "" Then
                        SynSQL.Text = rsTemp!��ȡSQL & ""
                        cboResource.ListIndex = i
                    End If
                Next
                strTemp = rsTemp!��ȡ���� & ""
                If Val(rsTemp!��ȡ��Դ & "") = 4 Then
                    varTemp = Split(strTemp, ":")
                    cboMed.Text = varTemp(0)
                    If UBound(varTemp) > 0 Then
                        txtDiagn.Text = varTemp(1)
                    End If
                End If
                chkOnlyRead.Value = Val(rsTemp!�Ƿ�ֻ�� & "")
            End If
    End Select
End Sub

Private Sub cblPrjType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboMed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
End Sub

Private Sub cboPrintForm_Click()
    Call AdjustLocation
End Sub

Private Sub cboPrintForm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboPrintFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
End Sub

Private Sub cboPrintType_Click()
    Dim blnDate As Boolean
    
    blnDate = Val(cboPrintType.Text) = 1
    cboPrintFormat.Visible = blnDate
    lblPrintFormat.Visible = blnDate
    If blnDate Then cboPrintFormat.ListIndex = 0
    If cboPrintFormat.Visible = False Then cboPrintFormat.Text = ""
End Sub

Private Sub cboPrintType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cboResource_Click()
    Call AdjustLocation
End Sub

Private Sub cboResource_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub
Private Sub chkHidden_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnlyRead_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    If CheckSQL Then
        MsgBox "��֤�ɹ���", vbInformation, Me.Caption
    End If
End Sub

Private Sub cmdSave_Click()
    
    If CehckData = False Then Exit Sub
    If SaveData Then
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub cmdSaveNew_Click()

    If CehckData = False Then Exit Sub
    If SaveData Then
        mblnOK = True
        Unload Me
        Call ShowMe(mbytType, mstrSName, mstrPatiPrj)
    End If
End Sub

Private Sub Form_Load()

    With cblPrjType
        .AddItem "S"
        .AddItem "B"
        .AddItem "A"
        .AddItem "R"
    End With
    With cboPrintForm
        .AddItem "1-�����"
        .AddItem "2-����ѡ��"
        .AddItem "3-����ѡ��"
    End With
    With cboPrintType
        .AddItem "0-�ı�"
        .AddItem "1-����"
        .AddItem "2-����"
    End With
    With cboPrintFormat
        .AddItem "YYYY-MM-DD HH:mm:ss"
        .AddItem "YYYY-MM-DD HH:mm"
        .AddItem "YYYY-MM-DD"
        .AddItem "HH:mm:ss"
        .AddItem "HH:mm"
    End With
    With cboResource
        .AddItem "0-�ֹ�����"
        .AddItem "1-�������"
        .AddItem "2-��������"
        .AddItem "3-��Ѫ���"
        .AddItem "4-��������"
        .AddItem "99-SQL��ȡ"
    End With
    With SynSQL
        '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
        .SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
        .SyntaxScheme = GetSqlColor
    End With
    Call GetEmr
    Call LoadData
End Sub

Private Function SaveData() As Boolean
    Dim strTemp As String, strNames As String
    Dim i As Long, lngTemp As Long
    
    On Error GoTo errH
    gstrSql = "Zl_ҽ�����Ӱಡ����Ŀ_Edit(" & mbytType & ",'" & lblNames.Caption & _
        "','" & txtPrjName.Text & "','" & mstrPatiPrj & "'," & IIf(mbytType = 1, 0, mlngNum) & ",'" & cblPrjType.Text & _
        "'," & Val(cboPrintForm.Text) & "," & Val(cboPrintType.Text) & ",'" & cboPrintFormat.Text & "'"
    lngTemp = Val(cboPrintForm.Text)
    If lngTemp = 3 Or lngTemp = 2 Then
        With vsfRange
            For i = 1 To .Rows - 1
                strTemp = .TextMatrix(i, .ColIndex("ֵ��"))
                If strTemp <> "" Then
                    If .Cell(flexcpChecked, i, .ColIndex("�ı���")) = flexChecked Then
                        strTemp = "*" & strTemp
                    End If
                    strNames = strNames & "," & strTemp
                End If
            Next
            strNames = Mid(strNames, 2)
        End With
    Else
        strNames = ""
    End If
    gstrSql = gstrSql & ",'" & strNames & "'," & txtRow.Text & "," & Val(cboResource.Text)
    strTemp = ""
    lngTemp = Val(cboResource.Text)
    If lngTemp = 4 Then
        gstrSql = gstrSql & ",'" & cboMed.Text & ":" & txtDiagn.Text & "',''"
    ElseIf lngTemp = 99 Then
        strTemp = SynSQL.Text
        strTemp = "'" & Replace(strTemp, "'", "''") & "'"
        gstrSql = gstrSql & ",''," & strTemp
    Else
        gstrSql = gstrSql & ",'',''"
    End If
    gstrSql = gstrSql & ",'" & txtDescript.Text & "'," & chkOnlyRead.Value & "," & chkHidden.Value & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    SaveData = True
    Call frmDocShiftBase.RefreshPrj(mbytType)
    Exit Function
errH:
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function CehckData() As Boolean
'��������ǰ�Ļ������
    Dim lngTemp As Long
    
    If txtPrjName.Text = "" Then
        MsgBox "��Ŀ���Ʋ���Ϊ�գ����飡", vbInformation, Me.Caption
        Call zlcontrol.ControlSetFocus(txtPrjName)
        Exit Function
    ElseIf zlstr.ActualLen(txtPrjName.Text) > 20 Then
        MsgBox "��Ŀ���Ʋ��ܳ���10�����֣����飡", vbInformation, Me.Caption
        Call zlcontrol.ControlSetFocus(txtPrjName)
        Exit Function
    End If
    
    If zlstr.ActualLen(txtDescript.Text) > 20 Then
        MsgBox "�������ֲ��ܳ���10�����֣����飡", vbInformation, Me.Caption
        Call zlcontrol.ControlSetFocus(txtDescript)
        Exit Function
    End If
    
    If vsfRange.Visible Then
        If vsfRange.Rows < 4 Then
            MsgBox "������߶���ѡ��ʱ��Ӧ�������������ϵ�ֵ��", vbInformation, Me.Caption
            Call zlcontrol.ControlSetFocus(vsfRange)
            Exit Function
        End If
    End If
    lngTemp = Val(cboResource.Text)
    If lngTemp = 99 Then
        If CheckSQL = False Then Exit Function
    ElseIf lngTemp = 4 Then
        If cboMed.Text = "" Then
            MsgBox "�����������Ʋ���Ϊ�գ����飡", vbInformation, Me.Caption
            Call zlcontrol.ControlSetFocus(cboMed)
            Exit Function
        End If
        If txtDiagn.Text = "" Then
            MsgBox "���������Ϊ�գ����飡", vbInformation, Me.Caption
            Call zlcontrol.ControlSetFocus(txtDiagn)
            Exit Function
        End If
        If zlstr.ActualLen(cboMed.Text) + zlstr.ActualLen(txtDiagn.Text) > 100 Then
            MsgBox "�������ݲ��ܳ���50�����֣����飡", vbInformation, Me.Caption
            Call zlcontrol.ControlSetFocus(txtDiagn)
            Exit Function
        End If
    End If
    CehckData = True
End Function

Private Function CheckSQL() As Boolean
'���SQL����ȷ��
    Dim rsTemp As ADODB.Recordset
    
    gstrSql = Trim(SynSQL.Text)
    If gstrSql = "" Then
        MsgBox "SQL����Ϊ�գ����飡", vbInformation, Me.Caption
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    ElseIf zlstr.ActualLen(gstrSql) > 4000 Then
        MsgBox "��ȡSQL���ܳ���4000�ַ������飡", vbInformation, "��֤SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    gstrSql = Replace(gstrSql, "[��ҳID]", "0")
    gstrSql = Replace(gstrSql, "[����ID]", "0")
    gstrSql = Replace(gstrSql, "[��ʼʱ��]", "sysdate")
    gstrSql = Replace(gstrSql, "[����ʱ��]", "sysdate")
    On Error Resume Next
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��֤SQL")
    If err.Number <> 0 Then
        MsgBox "SQL��д����ȷ�����飡" & vbNewLine & err.Description, vbInformation, "��֤SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    If rsTemp.Fields.Count > 1 Then
        MsgBox "��ȡSQLֻ�ܷ���һ���ı����͵��ֶΣ����飡", vbInformation, "��֤SQL"
        Call zlcontrol.ControlSetFocus(SynSQL)
        Exit Function
    End If
    CheckSQL = True
End Function


Private Sub AdjustLocation()
'����ѡ��Ĳ�ͬ����������ؼ���λ��
    Dim strTemp As String
    Dim lngType As Long, lngSource As Long
        
    fraSQL.Visible = False
    fraMed.Visible = False
    lngType = Val(cboPrintForm.Text)
    If lngType = 1 Then '1��ʾ�����
        lblPrintRange.Visible = False
        vsfRange.Visible = False 'ֻ��ѡ����ʱ����ʾֵ���������Ҫ
        fraSource.Visible = True 'ֻ���������Ҫ��ȡ��Դ��ѡ�����Ҫ
        fraSource.Move 120, cboPrintForm.Top + cboPrintForm.Height + 120
    Else '2��ʾ��ѡ�3��ʾ��ѡ��
        lblPrintRange.Visible = True
        vsfRange.Visible = True
        fraSource.Visible = False
        vsfRange.Move cboPrintForm.Left, cboPrintForm.Top + cboPrintForm.Height + 120
    End If
    lngSource = Val(cboResource.Text)
    If lngSource = 99 Then 'SQL��ȡ
        fraSQL.Visible = True 'SQL��ȡʱ������ʾSQL��
        fraMed.Visible = False
        If lngType = 1 Then
            fraSQL.Move 120, fraSource.Top + fraSource.Height + 120
        Else
            fraSQL.Move 120, vsfRange.Top + vsfRange.Height + 120
        End If
        fra3.Move 120, fraSQL.Top + fraSQL.Height + 120
    ElseIf lngSource = 4 Then '��������
        fraSQL.Visible = False 'ֻ�в�������ʱ����ʾ�������ݵ�
        fraMed.Visible = True
        If lngType = 1 Then
            fraMed.Move 120, fraSource.Top + fraSource.Height + 120
        Else
            fraMed.Move 120, vsfRange.Top + vsfRange.Height + 120
        End If
        fra3.Move 120, fraMed.Top + fraMed.Height + 120
    Else
        If lngType = 1 Then
            fra3.Move 120, fraSource.Top + fraSource.Height + 120
        Else
            fra3.Move 120, vsfRange.Top + vsfRange.Height + 120
        End If
    End If
    chkOnlyRead.Visible = IIf(lngSource = 0, False, True)
    chkOnlyRead.Value = 0
    Me.Height = fra3.Top + fra3.Height
End Sub

Private Sub synSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = 2 Then
        SynSQL.Paste
    ElseIf KeyCode = vbKeyZ And Shift = 2 Then
        SynSQL.Undo
    ElseIf KeyCode = vbKeyY And Shift = 2 Then
        SynSQL.Redo
    ElseIf KeyCode = vbKeyC And Shift = 2 Then
        SynSQL.Copy
    ElseIf KeyCode = vbKeyA And Shift = 2 Then
        SynSQL.SelectAll
    End If
End Sub

Private Sub txtDescript_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
End Sub

Private Sub txtDiagn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
End Sub

Private Sub txtPrjName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcommfun.PressKey(vbKeyTab)
    KeyAscii = IIf(KeyAscii = Asc("'"), 0, KeyAscii)
    KeyAscii = IIf(KeyAscii = Asc(";"), 0, KeyAscii)
End Sub

Private Sub txtRow_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlcontrol.ControlSetFocus(cboResource)
    KeyAscii = IIf(InStr("0123456789" & Chr(8), Chr(KeyAscii)), KeyAscii, 0)
End Sub

Private Sub UpDownRow_Change()
    txtRow.Text = UpDownRow.Value
End Sub

Private Sub vsfRange_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '������������Զ����ӻ��������
    With vsfRange
        If .Row = .Rows - 1 Then
            If .TextMatrix(.Row, .ColIndex("ֵ��")) <> "" Then
                .Rows = .Rows + 1
            End If
        End If
        
        If .TextMatrix(.Row, .ColIndex("ֵ��")) <> "" Then
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
            End If
        Else
            If .Row > 1 Then
                .RemoveItem .Row
            End If
        End If
    End With
End Sub

Private Sub vsfRange_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfRange
        If NewRow = 1 Then
            If .Rows = 2 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
        End If
    End With
End Sub

Private Sub vsfRange_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vsfRange
        If Not (Col = .ColIndex("ֵ��") Or Col = .ColIndex("�ı���")) Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfRange_Click()
    Dim lngRow As Long, lngCheck As Long
    Dim strRange As String
    
    With vsfRange
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row - 1
            End If
        ElseIf .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row + 1
            End If
        End If
        If lngRow = 0 Then Exit Sub
        lngCheck = .Cell(flexcpChecked, .Row, .ColIndex("�ı���"))
        strRange = .TextMatrix(.Row, .ColIndex("ֵ��"))
        .Cell(flexcpChecked, .Row, .ColIndex("�ı���")) = .Cell(flexcpChecked, lngRow, .ColIndex("�ı���"))
        .TextMatrix(.Row, .ColIndex("ֵ��")) = .TextMatrix(lngRow, .ColIndex("ֵ��"))
        .Cell(flexcpChecked, lngRow, .ColIndex("�ı���")) = lngCheck
        .TextMatrix(lngRow, .ColIndex("ֵ��")) = strRange
        .Row = lngRow
    End With
    
End Sub

Private Sub GetEmr()
'���ò����ӿڻ�ȡ������������
    Dim objEMR As Object
    Dim rsTemp As ADODB.Recordset
    
    Set objEMR = gfrmMain.mobjEMR
    If objEMR Is Nothing Then Exit Sub
    If Not objEMR.IsInited Or objEMR.IsOffline Then Exit Sub
    On Error Resume Next
    Set rsTemp = objEMR.GetDictItemsByTitle("�����ļ�����")
    If err.Number <> 0 Then Exit Sub
    Call zlcontrol.CboAddData(cboMed, rsTemp, True)
End Sub
