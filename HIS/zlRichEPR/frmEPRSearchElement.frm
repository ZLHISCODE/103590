VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRSearchElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ҫ�صļ�������"
   ClientHeight    =   4935
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7515
   Icon            =   "frmEPRSearchElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton optAsk 
      Caption         =   "������һ����(&2)"
      Height          =   180
      Index           =   1
      Left            =   4305
      TabIndex        =   18
      Top             =   780
      Width           =   1665
   End
   Begin VB.OptionButton optAsk 
      Caption         =   "����ȫ������(&1)"
      Height          =   180
      Index           =   0
      Left            =   2565
      TabIndex        =   17
      Top             =   780
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.CommandButton cmdAppend 
      Caption         =   "���(&A)"
      Height          =   350
      Left            =   6195
      TabIndex        =   10
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6195
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   915
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2925
      Left            =   -5610
      TabIndex        =   15
      Top             =   915
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame fraDefine 
      Caption         =   "������������:"
      Height          =   1050
      Left            =   90
      TabIndex        =   3
      Top             =   3855
      Width           =   7335
      Begin VB.ComboBox cboValue 
         Height          =   300
         Left            =   3255
         TabIndex        =   9
         Top             =   525
         Width           =   3960
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   525
         Width           =   1950
      End
      Begin VB.ComboBox cboFormula 
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   525
         Width           =   1140
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀ(&I):"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblFormula 
         AutoSize        =   -1  'True
         Caption         =   "����(&F):"
         Height          =   180
         Left            =   2085
         TabIndex        =   6
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "ֵ(&V):"
         Height          =   180
         Left            =   3255
         TabIndex        =   8
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "�Ƴ�(&R)"
      Height          =   350
      Left            =   6195
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3090
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6195
      TabIndex        =   14
      Top             =   465
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6195
      TabIndex        =   13
      Top             =   90
      Width           =   1200
   End
   Begin VB.Frame fraCodex 
      Height          =   30
      Left            =   75
      TabIndex        =   12
      Top             =   630
      Width           =   5910
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -135
      Top             =   4755
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
            Picture         =   "frmEPRSearchElement.frx":058A
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgThis 
      Height          =   2745
      Left            =   120
      TabIndex        =   2
      Top             =   1050
      Width           =   5835
      _cx             =   10292
      _cy             =   4842
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmEPRSearchElement.frx":09DC
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lblConditions 
      AutoSize        =   -1  'True
      Caption         =   "���������б�(&L):"
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   795
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "���ò����а����ġ��̶�����Ҫ�ء������������Ա㾫ȷ�ؼ�����ϣ���Ĳ�����¼��"
      Height          =   360
      Left            =   885
      TabIndex        =   0
      Top             =   120
      Width           =   5040
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEPRSearchElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gstrMatch As String

Const conCol��ĿID As Integer = 0
Const conCol��Ŀ�� As Integer = 1
Const conCol���� As Integer = 2
Const conCol��ϵʽ As Integer = 3
Const conCol����ֵ As Integer = 4

'�������
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByRef strTerms As String) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������ frmParent-������
    '       strTerms-Ҫ������
    '���أ�ȷ�����ķ���True��ȡ������False
    '---------------------------------------------------
Dim aryTerm() As String, aryField() As String
Dim lngCount As Long
    
    If strTerms <> "" Then
        If Val(Left(strTerms, 1)) = 0 Then Me.optAsk(1).Value = True
        aryTerm = Split(Mid(strTerms, 3), "|")
        With Me.vfgThis
            .Redraw = flexRDNone
            For lngCount = 0 To UBound(aryTerm)
                .Rows = .Rows + 1
                aryField = Split(aryTerm(lngCount), ";")
                .TextMatrix(.Rows - 1, conCol��ĿID) = Val(aryField(conCol��ĿID))
                .TextMatrix(.Rows - 1, conCol��Ŀ��) = aryField(conCol��Ŀ��)
                .TextMatrix(.Rows - 1, conCol����) = Val(aryField(conCol����))
                .TextMatrix(.Rows - 1, conCol��ϵʽ) = aryField(conCol��ϵʽ)
                .TextMatrix(.Rows - 1, conCol����ֵ) = aryField(conCol����ֵ)
            Next
            .Redraw = flexRDDirect
            If .Rows > .FixedRows Then .Row = .Rows - 1
        End With
    End If
    
    Me.Show vbModal, frmParent
    
    If mblnOK Then
        With Me.vfgThis
            strTerms = ""
            For lngCount = .FixedRows To .Rows - 1
                strTerms = strTerms & _
                    "|" & .TextMatrix(lngCount, conCol��ĿID) & _
                    ";" & .TextMatrix(lngCount, conCol��Ŀ��) & _
                    ";" & .TextMatrix(lngCount, conCol����) & _
                    ";" & .TextMatrix(lngCount, conCol��ϵʽ) & _
                    ";" & .TextMatrix(lngCount, conCol����ֵ)
            Next
        End With
        If strTerms <> "" Then strTerms = IIf(Me.optAsk(0).Value, 1, 0) & strTerms
    End If
    ShowMe = mblnOK: Unload Me
End Function

Private Sub cboFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboValue_Change()
    ValidControlText cboValue
End Sub

Private Sub cboValue_GotFocus()
    Me.cboValue.SelStart = 0: Me.cboValue.SelLength = 100
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdAppend_Click()
Dim strTemp As String
    'ϸ������ȷ�Լ��
    If Trim(Me.txtItem.Tag) <> Trim(Me.txtItem.Text) Or Trim(Me.txtItem.Text) = "" Then
        MsgBox "δָ����ȷ������ϸ����Ŀ��", vbExclamation, gstrSysName
        Me.txtItem.SetFocus: Exit Sub
    End If
    If Trim(Me.cboFormula.Text) = "" Then
        MsgBox "δָ����ȷ������ϸ���ϵʽ��", vbExclamation, gstrSysName
        Me.cboFormula.SetFocus: Exit Sub
    End If
    If Me.cboValue.Enabled Then
        If Trim(Me.cboValue.Text) = "" Then
            MsgBox "δָ��������ϸ������ֵ��", vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
        strTemp = zlVerifyForm
        If strTemp <> "" Then
            MsgBox strTemp, vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
    End If
    
    '��ϸ����ӵ������
    With Me.vfgThis
        .Redraw = flexRDNone
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, conCol��ĿID) = Val(Me.lblItem.Tag)
        .TextMatrix(.Rows - 1, conCol��Ŀ��) = Trim(Me.txtItem.Text)
        .TextMatrix(.Rows - 1, conCol����) = Val(Me.lblFormula.Tag)
        .TextMatrix(.Rows - 1, conCol��ϵʽ) = Trim(Me.cboFormula.Text)
        .TextMatrix(.Rows - 1, conCol����ֵ) = Trim(Me.cboValue.Text)
        .Row = .Rows - 1
        .Col = conCol��Ŀ��
        .Redraw = flexRDDirect
    End With

    
    '���ϸ����ؼ����ݣ��Ա㶨���µ�ϸ��
    Me.lblItem.Tag = ""
    Me.txtItem.Text = ""
    Me.txtItem.Tag = ""
    Me.lblValue.Tag = ""
    Me.cboValue.Text = ""
    Me.lblFormula.Tag = ""
    Me.cboFormula.Clear
    Me.vfgThis.SetFocus
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    mblnOK = True: Me.Hide
End Sub

Private Sub cmdRemove_Click()
Dim rsTemp As New ADODB.Recordset
    If Val(Me.vfgThis.TextMatrix(Me.vfgThis.Row, conCol��ĿID)) = 0 Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select i.Id, i.����, i.������, i.Ӣ����, i.����, i.����, i.С��, i.��λ, Decode(i.�滻��, 2, '', i.��ֵ��) As ��ֵ��" & _
            " From ����������Ŀ i" & _
            " Where i.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Me.vfgThis.TextMatrix(Me.vfgThis.Row, conCol��ĿID)))
    With rsTemp
        Me.lblItem.Tag = !ID
        Me.txtItem.Text = !������
        Me.txtItem.Tag = !������
        Me.lblFormula.Tag = IIf(IsNull(!����), 0, !����)
        Me.lblValue.Tag = "" & !��ֵ��
        Me.cboValue.Tag = "" & !��λ
        Call zlAdjustForm
    End With
    
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        Me.cboFormula.Text = .TextMatrix(.Row, conCol��ϵʽ)
        Me.cboValue.Text = .TextMatrix(.Row, conCol����ֵ)
        .Rows = .Rows - 1: .Row = .Rows - 1
    End With
    Me.txtItem.SetFocus
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwList.Visible Then
        Me.lvwList.Visible = False
        Me.txtItem.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim lngCount As Long
    With Me.vfgThis
        .Redraw = flexRDNone
        .Rows = 1: .Cols = 5
        For lngCount = 0 To .Cols - 1
            .ColAlignment(lngCount) = 1
        Next
        .TextMatrix(0, conCol��ĿID) = "��ĿID"
        .TextMatrix(0, conCol��Ŀ��) = "��Ŀ��"
        .TextMatrix(0, conCol����) = "����"
        .TextMatrix(0, conCol��ϵʽ) = "��ϵʽ"
        .TextMatrix(0, conCol����ֵ) = "����ֵ"
        
        .ColWidth(conCol��ĿID) = 0
        .ColWidth(conCol��Ŀ��) = 1600
        .ColWidth(conCol����) = 0
        .ColWidth(conCol��ϵʽ) = 900
        .ColWidth(conCol����ֵ) = .Width - .ColWidth(conCol��Ŀ��) - .ColWidth(conCol��ϵʽ) - 250
        .Redraw = flexRDDirect
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "������", "������", 1800
        .Add , "����", "����", 1000
        .Add , "����", "����", 600
        .Add , "��ֵ��", "��ֵ��", 4000
    End With
    Me.lvwList.ColumnHeaders("����").Position = 1
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = Split(.SelectedItem.Tag, ",")(0)
        Me.txtItem.Tag = Split(.SelectedItem.Tag, ",")(0)
        Me.lblFormula.Tag = Split(.SelectedItem.Tag, ",")(1)
        Me.lblValue.Tag = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("��ֵ��").Index - 1)
        Me.cboValue.Tag = Split(.SelectedItem.Tag, ",")(2)
        Call zlAdjustForm
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub txtItem_Change()
    ValidControlText txtItem
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    If InStr(" ~!@#$^&*()+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select i.Id, i.����, i.������, i.Ӣ����, i.����, i.����, i.С��, i.��λ, Decode(i.�滻��, 2, '', i.��ֵ��) As ��ֵ��" & _
            " From ����������Ŀ i" & _
            " Where i.���� In (0, 1) And (i.���� Like [1] || '%' Or i.������ Like '" & gstrMatch & "'|| [1] ||'%' Or Upper(i.Ӣ����) Like '" & gstrMatch & "'|| [1] ||'%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtItem.Text))
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�ָ������Ҫ�أ�", vbExclamation, gstrSysName
            Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
            Me.txtItem.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.lblItem.Tag = !ID
            Me.txtItem.Text = !������
            Me.txtItem.Tag = !������
            Me.lblFormula.Tag = IIf(IsNull(!����), 0, !����)
            Me.lblValue.Tag = IIf(IsNull(!��ֵ��), "", !��ֵ��)
            Me.cboValue.Tag = IIf(IsNull(!��λ), "", !��λ)
            Call zlAdjustForm
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !������ & IIf(IsNull(!Ӣ����), "", "(" & !Ӣ���� & ")"), "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = !����
            Select Case IIf(IsNull(!����), 0, !����)
            Case 0
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "��ֵ"
            Case 1
                objItem.SubItems(Me.lvwList.ColumnHeaders("����").Index - 1) = "����"
            End Select
            objItem.SubItems(Me.lvwList.ColumnHeaders("��ֵ��").Index - 1) = IIf(IsNull(!��ֵ��), "", !��ֵ��)
            objItem.Tag = !������ & "," & IIf(IsNull(!����), 0, !����) & "," & IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        With Me.lvwList
            .ListItems(1).Selected = True
            .Left = Me.fraDefine.Left + Me.txtItem.Left
            .Top = Me.fraDefine.Top + Me.txtItem.Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgThis_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub zlAdjustForm()
Dim lngCount As Long
    '-------------------------------------------------
    '�����������ʽ�Ŀ�ѡ��Χ
    '��Σ� ������Me.lblFormula.Tag�е���ֵ���ͣ�Me.lblValue.Tag�е���ֵ��
    '-------------------------------------------------
    Dim aryValue() As String
    Me.cboValue.Clear
    Me.cboValue.Enabled = False
    Me.cboFormula.Clear
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '��ֵ
        If Me.cboValue.Tag = "" Then
            Me.lblValue.Caption = "ֵ(&V):(��ֵ��)"
        Else
            Me.lblValue.Caption = "ֵ(&V):(��ֵ�� ��λ:" & Me.cboValue.Tag & ")"
        End If
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "С��"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 1  '����
        Me.lblValue.Caption = "ֵ(&V):(������)"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.AddItem "����"
        Me.cboFormula.AddItem "������"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case Else
    End Select
    
    aryValue = Split(Me.lblValue.Tag, ";")
    For lngCount = LBound(aryValue) To UBound(aryValue)
        Me.cboValue.AddItem aryValue(lngCount)
    Next
End Sub

Private Function zlVerifyForm() As String
    '-------------------------------------------------
    '�ж��������ʽ��ֵ�������ȷ��
    '��Σ�������Me.lblFormula.Tag�е���ֵ����
    '       Me.lblValue.Tag�е���ֵ��
    '       Me.lblFormula.text�еĹ�ϵʽ
    '       Me.lblValue.text�е�����
    '���Σ���ȷ����""�����򷵻ش�����Ϣ
    '-------------------------------------------------
Dim aryValue() As String
Dim lngCount As Long
    zlVerifyForm = ""
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '��ֵ
        Select Case Me.cboFormula.Text
        Case "����", "������", "����", "С��", "����", "����"
            Me.cboValue.Text = Val(Me.cboValue.Text)
        Case "����"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "����ֵδ�������ڡ�Ҫ�����ֵ1,ֵ2����ʽ��֯��д��": Exit Function
            End If
            Me.cboValue.Text = Val(aryValue(0)) & "," & Val(aryValue(1))
        Case "����", "������"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "�����Ϊ������ֵ��û��Ҫ���á����ڡ��򡰲����ڡ��Ĺ�ϵʽ��": Exit Function
            End If
            Me.cboValue.Text = ""
            For lngCount = LBound(aryValue) To UBound(aryValue)
                Me.cboValue.Text = Me.cboValue.Text & "," & Val(aryValue(lngCount))
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 1  '����
        Select Case Me.cboFormula.Text
        Case "����", "������", "����", "������"
        Case "����", "������"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "�����Ϊ������ֵ��û��Ҫ���á����ڡ��򡰲����ڡ��Ĺ�ϵʽ��": Exit Function
            End If
        End Select
    End Select
End Function

