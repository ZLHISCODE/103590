VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelValue 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器定义"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmSelValue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6555
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   6555
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4635
      Width           =   6555
      Begin VB.CommandButton cmdPreview 
         Caption         =   "预览(&E)"
         Height          =   350
         Left            =   390
         TabIndex        =   5
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "验证明细(&V)"
         Height          =   350
         Left            =   1605
         TabIndex        =   6
         Top             =   600
         Width           =   1320
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3885
         TabIndex        =   7
         Top             =   600
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5130
         TabIndex        =   8
         Top             =   600
         Width           =   1100
      End
      Begin VB.TextBox txtDefShow 
         Height          =   300
         Left            =   1215
         MaxLength       =   255
         TabIndex        =   3
         Top             =   60
         Width           =   1665
      End
      Begin VB.TextBox txtDefBand 
         Height          =   300
         Left            =   3900
         MaxLength       =   255
         TabIndex        =   4
         Top             =   60
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省显示"
         Height          =   180
         Left            =   435
         TabIndex        =   14
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缺省绑定"
         Height          =   180
         Left            =   3135
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   15
         X2              =   19200
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   19200
         Y1              =   465
         Y2              =   465
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshField 
      Height          =   1650
      Left            =   180
      TabIndex        =   2
      Top             =   2865
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   2910
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      RowHeightMin    =   250
      BackColorSel    =   10251637
      BackColorBkg    =   16777215
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      MouseIcon       =   "frmSelValue.frx":014A
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   4560
      Picture         =   "frmSelValue.frx":0464
      ScaleHeight     =   1350
      ScaleWidth      =   1785
      TabIndex        =   10
      Top             =   2865
      Width           =   1785
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "如果分类数据中存在ID及上级ID字段，则在选择器中会自动以树形结构显示。"
         Height          =   735
         Left            =   60
         TabIndex        =   11
         Top             =   525
         Width           =   1770
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   180
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6150
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2805
      Width           =   6150
   End
   Begin VB.TextBox txtSQL 
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   495
      Width           =   6165
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   4560
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   8043
      TabWidthStyle   =   2
      TabFixedWidth   =   2291
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "明细数据(&1)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "分类数据(&2)"
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "frmSelValue.frx":12A6
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6015
      Top             =   3615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelValue.frx":1408
            Key             =   "VarChar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelValue.frx":151A
            Key             =   "Numeric"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelValue.frx":162C
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelValue.frx":173E
            Key             =   "Date"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSelValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入/出：SQL及字段描述
Public mstrSQLList As String
Public mstrSQLTree As String
Public mstrFLDList As String
Public mstrFLDTree As String
Public mstrObj As String
Public mstrDef As String '缺省值

'入：
Public mbytDataType As Byte   '参数数据类型
Public mstrParName As String '参数名称
Public mlngSys As Long
Public mstrOwner As String

Private mstrObjList As String
Private mstrObjTree As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim strFields As String, mstrObject As String
    Dim strSQL As String, strR As String
    Dim i As Integer, strFldPre As String
    Dim blnDo As Boolean
    
    strSQL = RemoveNote(txtSQL.Text)
    
    If TrimChar(strSQL) = "" Then
        If tbs.SelectedItem.Index = 1 Then
            MsgBox "请先输入SQL语句！", vbInformation, App.Title
            txtSQL.SetFocus: Exit Sub
        Else
            mstrSQLTree = ""
            mstrFLDTree = ""
            mstrObjTree = ""
            '去除明细表数据中的关联字段
            mstrFLDList = Replace(mstrFLDList, "&R", "")
            Call ClearGrid
            Call SetEnable
            Exit Sub
        End If
    End If
    
    strSQL = TrimChar(Replace(Replace(strSQL, "[*]", ""), "[系统]", mlngSys))
    
    'SQL对象所有者权限检查
    '取对象
    mstrObject = SQLObject(strSQL)
    If mstrObject = "" Then
        MsgBox "不能分析SQL语句所查询的数据对象,请检查是否正确书写！", vbInformation, App.Title
        txtSQL.SetFocus: Exit Sub
    End If
    
    '是否有权限
    strR = CheckObjectPriv(mstrObject, mstrOwner)
    If strR <> "" Then
        MsgBox "当前用户不具有下列对象或没有权限访问这些对象:" & vbCrLf & vbCrLf & strR, vbInformation, App.Title
        txtSQL.SetFocus: Exit Sub
    End If
    
    '取所有者
    mstrObject = ObjectOwner(mstrObject, mstrOwner, Me)
    If mstrObject = "取消" Then Exit Sub '取消操作
    
    strSQL = SQLOwner(strSQL, mstrObject)
    
    Screen.MousePointer = 11
    
    strFields = CheckSQL(strSQL, strR)
    
    Screen.MousePointer = 0
    Me.Refresh
    
    If strFields = "" Then
        MsgBox "SQL语句校验失败！" & vbCrLf & vbCrLf & _
            "错误 " & strR & vbCrLf & vbCrLf & _
            "请检查是否正确书写！", vbInformation, App.Title
    Else
        For i = 0 To UBound(Split(strFields, "|"))
            '不允许写入的字段类型
            If CLng(Split(Split(strFields, "|")(i), ",")(1)) = adLongVarBinary Then
                MsgBox "该数据中存在二进制类型的字段项目,选择器不能处理这种项目,请修改！", vbInformation, App.Title
                Exit Sub
            End If
        Next
        
        If tbs.SelectedItem.Index = 1 Then
            mstrSQLList = tbs.SelectedItem.Tag
            mstrObjList = mstrObject
            strFldPre = mstrFLDList
            mstrFLDList = ""
            For i = 0 To UBound(Split(strFields, "|"))
                strR = GetScript(strFldPre, CStr(Split(Split(strFields, "|")(i), ",")(0)))
                If UCase(Split(Split(strFields, "|")(i), ",")(0)) Like "*ID" Then
                    mstrFLDList = mstrFLDList & "|" & Split(strFields, "|")(i) & IIf(strR = "", ",", strR)
                Else
                    mstrFLDList = mstrFLDList & "|" & Split(strFields, "|")(i) & IIf(strR = "", ",&S", strR)
                End If
            Next
            mstrFLDList = Mid(mstrFLDList, 2)
        Else
            mstrSQLTree = tbs.SelectedItem.Tag
            mstrObjTree = mstrObject
            strFldPre = mstrFLDTree
            mstrFLDTree = ""
            For i = 0 To UBound(Split(strFields, "|"))
                strR = GetScript(strFldPre, CStr(Split(Split(strFields, "|")(i), ",")(0)))
                '只能有一个字段供数据显示(缺省为第一个)
                If Not blnDo Then
                    If UCase(Split(Split(strFields, "|")(i), ",")(0)) Like "*ID" Then
                        mstrFLDTree = mstrFLDTree & "|" & Split(strFields, "|")(i) & IIf(strR = "", ",", strR)
                        If IIf(strR = "", ",", strR) Like "*&S*" Then blnDo = True
                    Else
                        mstrFLDTree = mstrFLDTree & "|" & Split(strFields, "|")(i) & IIf(strR = "", ",&S", strR)
                        If IIf(strR = "", ",&S", strR) Like "*&S*" Then blnDo = True
                    End If
                Else
                    mstrFLDTree = mstrFLDTree & "|" & Split(strFields, "|")(i) & ","
                End If
            Next
            mstrFLDTree = Mid(mstrFLDTree, 2)
        End If
        Call InitGrid
        Call SetEnable
    End If
End Sub

Private Sub cmdOK_Click()
    If Not CheckValid Then Exit Sub
    
    '取对象
    mstrObj = mstrObjList & "|" & mstrObjTree
    
    '取缺省值
    If txtDefShow.Text <> "" Then
        mstrDef = txtDefShow.Text & "|" & txtDefBand.Text
    Else
        mstrDef = ""
    End If
    
    gblnOK = True
    Hide
End Sub

Private Sub cmdPreview_Click()
    Dim blnOK As Boolean
    
    If Not CheckValid Then Exit Sub
    
    blnOK = gblnOK
    
    frmSelect.mstrSQLList = Replace(Replace(mstrSQLList, "[*]", ""), "[系统]", mlngSys)
    frmSelect.mstrSQLTree = Replace(Replace(mstrSQLTree, "[*]", ""), "[系统]", mlngSys)
    frmSelect.mstrFLDList = mstrFLDList
    frmSelect.mstrFLDTree = mstrFLDTree
    frmSelect.mstrParName = mstrParName
    frmSelect.mbytDataType = mbytDataType
    
    frmSelect.mlngSeekHwnd = cmdPreview.hwnd
    
    On Error Resume Next
    Err.Clear
    
    frmSelect.Show 1, Me
    If gblnOK Then Unload frmSelect
    
    gblnOK = blnOK
End Sub

Private Sub Form_Load()
    gblnOK = False
    
    RestoreWinState Me, App.ProductName
    
    Caption = Caption & " - 参数:" & mstrParName

    tbs.Tabs(1).Tag = mstrSQLList
    tbs.Tabs(2).Tag = mstrSQLTree
    
    txtSQL.Text = tbs.SelectedItem.Tag
    
    If mstrObj <> "" And UBound(Split(mstrObj, "|")) = 1 Then
        mstrObjList = Split(mstrObj, "|")(0)
        mstrObjTree = Split(mstrObj, "|")(1)
    End If
    
    Call SetEnable
    
    If mstrDef <> "" Then
        txtDefShow.Text = Split(mstrDef, "|")(0)
        txtDefBand.Text = Split(mstrDef, "|")(1)
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Dim sngScale As Single
    
    tbs.Width = Me.ScaleWidth - 160
    tbs.Height = Me.ScaleHeight - 1150
    
    txtSQL.Width = tbs.Width - 230
    pic.Width = txtSQL.Width
    If tbs.SelectedItem.Index = 1 Then
        mshField.Width = txtSQL.Width
    Else
        mshField.Width = txtSQL.Width - picInfo.Width - 100
        picInfo.Left = mshField.Left + mshField.Width + 100
    End If
    
    sngScale = txtSQL.Height / (txtSQL.Height + mshField.Height)
    If sngScale >= 1 Or sngScale <= 0 Then sngScale = 0.5
    
    txtSQL.Height = (tbs.Height - 550 - pic.Height) * sngScale
    pic.Top = txtSQL.Top + txtSQL.Height
    mshField.Top = pic.Top + pic.Height
    mshField.Height = (tbs.Height - 550 - pic.Height) * (1 - sngScale)
    picInfo.Top = mshField.Top + 100
    picInfo.Height = mshField.Height - 100
    
    If Me.ScaleWidth - cmdCancel.Width * 1.3 >= 4300 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnDo As Boolean, lngW As Long, i As Integer
    
    '处理Check类型单元的鼠标
    If mshField.MouseRow >= 1 And mshField.MouseRow <= mshField.Rows - 1 Then
        If mshField.MouseCol >= 2 And mshField.MouseCol <= mshField.Cols - 1 Then
            If mshField.TextMatrix(mshField.MouseRow, 1) <> "" Then
                If Y <= (mshField.Rows - mshField.TopRow + 1) * mshField.RowHeight(0) Then
                    For i = 0 To mshField.Cols - 1
                        lngW = lngW + mshField.ColWidth(i)
                    Next
                    If X <= lngW Then blnDo = True
                End If
            End If
        End If
    End If
    If blnDo Then
        mshField.MousePointer = 99
    Else
        mshField.MousePointer = 0
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If txtSQL.Height + Y < 500 Or mshField.Height - Y < picInfo.Height Then Exit Sub
        pic.Top = pic.Top + Y
        txtSQL.Height = txtSQL.Height + Y
        mshField.Top = mshField.Top + Y
        mshField.Height = mshField.Height - Y
        picInfo.Top = picInfo.Top + Y
        picInfo.Height = picInfo.Height - Y
    End If
End Sub

Private Sub mshField_DBLClick()
'功能：字段编辑
    Dim intRow As Integer, intCol As Integer
    Dim i As Integer, blnDo As Boolean
    
    mshField.Redraw = False
    
    If mshField.MousePointer = 99 Then
        intRow = mshField.MouseRow
        intCol = mshField.MouseCol
        If tbs.SelectedItem.Index = 1 Then
            Select Case intCol
                Case 2 '选择项目
                    If mshField.TextMatrix(intRow, intCol) = "" Then
                        mshField.TextMatrix(intRow, intCol) = "√"
                        Call SetField(mstrFLDList, mshField.TextMatrix(intRow, 1), intCol, True)
                    Else
                        '至少要有一个选择显示项
                        For i = 1 To mshField.Rows - 1
                            If mshField.TextMatrix(i, intCol) <> "" And i <> intRow Then blnDo = True: Exit For
                        Next
                        If blnDo Then
                            mshField.TextMatrix(intRow, intCol) = ""
                            Call SetField(mstrFLDList, mshField.TextMatrix(intRow, 1), intCol, False)
                        End If
                    End If
                Case 3, 4, 5 '显示项目,绑定项目,关联项目
                    If intCol = 5 And mstrFLDTree = "" Then Exit Sub
                    
                    If intCol = 5 Or intCol = 4 Then '不允许关联的或绑定的字段类型
                        Select Case mshField.RowData(intRow)
                            Case adChar, adVarChar, adNumeric, adVarNumeric, adDBTimeStamp
                            Case Else
                                Exit Sub
                        End Select
                    End If
                    
                    '这些项目只能有一个选中
                    If mshField.TextMatrix(intRow, intCol) = "" Then
                        For i = 1 To mshField.Rows - 1
                            If i <> intRow And mshField.TextMatrix(i, intCol) <> "" Then
                                mshField.TextMatrix(i, intCol) = ""
                                Call SetField(mstrFLDList, mshField.TextMatrix(i, 1), intCol, False)
                            End If
                        Next
                        mshField.TextMatrix(intRow, intCol) = "√"
                        Call SetField(mstrFLDList, mshField.TextMatrix(intRow, 1), intCol, True)
                    End If
            End Select
        Else
            Select Case intCol
                Case 2, 5 '选择项目,关联项目
                    If intCol = 5 And mstrFLDList = "" Then Exit Sub
                    
                    If intCol = 5 Then '不允许关联的字段类型
                        Select Case mshField.RowData(intRow)
                            Case adChar, adVarChar, adNumeric, adVarNumeric, adDBTimeStamp
                            Case Else
                                Exit Sub
                        End Select
                    End If
                    
                    '这些项目只能有一个选中
                    If mshField.TextMatrix(intRow, intCol) = "" Then
                        For i = 1 To mshField.Rows - 1
                            If i <> intRow And mshField.TextMatrix(i, intCol) <> "" Then
                                mshField.TextMatrix(i, intCol) = ""
                                Call SetField(mstrFLDTree, mshField.TextMatrix(i, 1), intCol, False)
                            End If
                        Next
                        mshField.TextMatrix(intRow, intCol) = "√"
                        Call SetField(mstrFLDTree, mshField.TextMatrix(intRow, 1), intCol, True)
                    End If
            End Select
        End If
    End If
    
    mshField.Redraw = True
End Sub

Private Sub InitGrid()
    Dim i As Integer, strFld As String
    
    mshField.Redraw = False
    
    mshField.Clear
    mshField.Rows = 2
    mshField.Cols = 6
    
    For i = 0 To mshField.Cols - 1
        mshField.ColAlignmentFixed(i) = 4
        If i = 1 Then
            mshField.ColAlignment(i) = 1
        Else
            mshField.ColAlignment(i) = 4
        End If
    Next
    
    mshField.TextMatrix(0, 0) = ""
    mshField.TextMatrix(0, 1) = "项目名称"
    mshField.TextMatrix(0, 2) = "选择项目" '&S
    mshField.TextMatrix(0, 3) = "显示项目" '&D
    mshField.TextMatrix(0, 4) = "绑定项目" '&B
    mshField.TextMatrix(0, 5) = "关联项目" '&R
    
    mshField.ColWidth(0) = 300
    mshField.ColWidth(1) = 1500
    mshField.ColWidth(2) = 950
    
    If tbs.SelectedItem.Index = 1 Then
        mshField.ColWidth(3) = 950
        mshField.ColWidth(4) = 950
    Else
        mshField.ColWidth(3) = 0
        mshField.ColWidth(4) = 0
    End If
    mshField.ColWidth(5) = 950
    
    '根据变量显示字段
    If tbs.SelectedItem.Index = 1 Then
        strFld = mstrFLDList
    Else
        strFld = mstrFLDTree
    End If
    
    For i = 0 To UBound(Split(strFld, "|"))
        If i > 0 Then mshField.Rows = mshField.Rows + 1
        mshField.TextMatrix(i + 1, 1) = Split(Split(strFld, "|")(i), ",")(0) '字段名
        mshField.RowData(i + 1) = CLng(Split(Split(strFld, "|")(i), ",")(1)) '字段类型
        mshField.Row = i + 1: mshField.Col = 0
        Select Case mshField.RowData(i + 1)
            Case adNumeric, adVarNumeric
                Set mshField.CellPicture = img16.ListImages("Numeric").Picture
            Case adChar, adVarChar, adLongVarChar
                Set mshField.CellPicture = img16.ListImages("VarChar").Picture
            Case adDBTimeStamp
                Set mshField.CellPicture = img16.ListImages("Date").Picture
            Case Else
                Set mshField.CellPicture = img16.ListImages("Other").Picture
        End Select
        mshField.CellPictureAlignment = 4
        mshField.CellBackColor = vbWhite
        '√
        If Split(Split(strFld, "|")(i), ",")(2) Like "*&S*" Then mshField.TextMatrix(i + 1, 2) = "√"
        If Split(Split(strFld, "|")(i), ",")(2) Like "*&D*" Then mshField.TextMatrix(i + 1, 3) = "√"
        If Split(Split(strFld, "|")(i), ",")(2) Like "*&B*" Then mshField.TextMatrix(i + 1, 4) = "√"
        If Split(Split(strFld, "|")(i), ",")(2) Like "*&R*" Then mshField.TextMatrix(i + 1, 5) = "√"
    Next
    
    mshField.Row = 1: mshField.Col = 1: mshField.ColSel = mshField.Cols - 1
    mshField.Redraw = True
End Sub

Private Sub ClearGrid()
    Dim i As Integer
    
    mshField.Redraw = False
    
    mshField.Clear
    mshField.Rows = 2
    mshField.Cols = 6
    
    For i = 0 To mshField.Cols - 1
        mshField.ColAlignmentFixed(i) = 4
        If i = 1 Then
            mshField.ColAlignment(i) = 1
        Else
            mshField.ColAlignment(i) = 4
        End If
    Next
    
    mshField.TextMatrix(0, 0) = ""
    mshField.TextMatrix(0, 1) = "项目名称"
    mshField.TextMatrix(0, 2) = "选择项目"
    mshField.TextMatrix(0, 3) = "显示项目"
    mshField.TextMatrix(0, 4) = "绑定项目"
    mshField.TextMatrix(0, 5) = "关联项目"
    mshField.ColWidth(0) = 300
    mshField.ColWidth(1) = 1500
    mshField.ColWidth(2) = 950
    If tbs.SelectedItem.Index = 1 Then
        mshField.ColWidth(3) = 950
        mshField.ColWidth(4) = 950
    Else
        mshField.ColWidth(3) = 0
        mshField.ColWidth(4) = 0
    End If
    mshField.ColWidth(5) = 950
    
    mshField.Row = 1: mshField.Col = 1: mshField.ColSel = mshField.Cols - 1
    mshField.Redraw = True
End Sub

Private Sub tbs_Click()
    txtSQL.Text = tbs.SelectedItem.Tag
    If tbs.SelectedItem.Index = 1 Then
        picInfo.Visible = False
        cmdCheck.Caption = "验证明细(&V)"
        mshField.Width = 6165
    Else
        picInfo.Visible = True
        cmdCheck.Caption = "验证分类(&V)"
        mshField.Width = 4265
    End If
    Call SetEnable
    Form_Resize
    txtSQL.SetFocus
End Sub

Private Sub txtDefBand_GotFocus()
    SelAll txtDefBand
End Sub

Private Sub txtDefBand_KeyPress(KeyAscii As Integer)
    If InStr("'`~!@#$^&{}"";\|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDefShow_GotFocus()
    SelAll txtDefShow
End Sub

Private Sub txtDefShow_KeyPress(KeyAscii As Integer)
    If InStr("'`~!@#$^&{}"";\|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 2 Then SelAll txtSQL
End Sub

Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$^&{}"";:\", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSQL_Change()
    tbs.SelectedItem.Tag = txtSQL.Text
    If TrimChar(txtSQL.Text) = "" And tbs.SelectedItem.Index = 2 Then
        mstrSQLTree = ""
        mstrFLDTree = ""
        mstrObjTree = ""
        '去除明细表数据中的关联字段
        mstrFLDList = Replace(mstrFLDList, "&R", "")
        Call ClearGrid
    End If
    Call SetEnable
End Sub

Private Sub SetEnable()
    If tbs.SelectedItem.Index = 1 Then
        If UCase(tbs.Tabs(1).Tag) = UCase(mstrSQLList) And tbs.Tabs(1).Tag <> "" Then
            txtSQL.BackColor = vbWhite
            Call InitGrid
        Else
            txtSQL.BackColor = Me.BackColor
            Call ClearGrid
            cmdOK.Enabled = False
            cmdPreview.Enabled = False
        End If
    Else
        If UCase(tbs.Tabs(2).Tag) = UCase(mstrSQLTree) Or tbs.Tabs(2).Tag = "" Then
            txtSQL.BackColor = vbWhite
            If tbs.Tabs(2).Tag = "" Then
                Call ClearGrid
            Else
                Call InitGrid
            End If
        Else
            txtSQL.BackColor = Me.BackColor
            Call ClearGrid
            cmdOK.Enabled = False
            cmdPreview.Enabled = False
        End If
    End If
    
    If UCase(tbs.Tabs(1).Tag) = UCase(mstrSQLList) And UCase(tbs.Tabs(2).Tag) = UCase(mstrSQLTree) And mstrSQLList <> "" Then
        cmdOK.Enabled = True
        cmdPreview.Enabled = True
    End If
End Sub

Private Sub SetField(strFLDs As String, strFiledName As String, intType As Integer, bln As Boolean)
'功能：设置某个字段的描述(SDBR)
'参数：strFlds=字段描述串
'      strFiledName=字段名
'      intType=描述类型(2=S,3=D,4=B,5=R)
'      bln=描述是否有效
'返回：strFlds=修改后的字段描述串
    Dim i As Integer, strTmp As String
    Dim strModi As String, strScript As String
    
    strScript = Switch(intType = 2, "&S", intType = 3, "&D", intType = 4, "&B", intType = 5, "&R")
    
    For i = 0 To UBound(Split(strFLDs, "|"))
        strTmp = Split(Split(strFLDs, "|")(i), ",")(2)
        If Split(Split(strFLDs, "|")(i), ",")(0) = strFiledName Then
            If bln Then
                If InStr(strTmp, strScript) = 0 Then
                    strTmp = strTmp & strScript
                End If
            Else
                strTmp = Replace(strTmp, strScript, "")
            End If
        End If
        strModi = strModi & "|" & _
            Split(Split(strFLDs, "|")(i), ",")(0) & "," & _
            Split(Split(strFLDs, "|")(i), ",")(1) & "," & strTmp
    Next
    strFLDs = Mid(strModi, 2)
End Sub

Private Function CheckValid() As Boolean
'功能：检查选择器定义的合法性
    Dim i As Integer, lngList As Long, lngTree As Long
    
    '检查选择字段
    If InStr(mstrFLDList, "&S") = 0 Then
        MsgBox "在明细数据中没有设置供选择的字段项目！", vbInformation, App.Title
        Exit Function
    End If
    If mstrFLDTree <> "" And InStr(mstrFLDTree, "&S") = 0 Then
        MsgBox "在分类数据中没有设置供选择的字段项目！", vbInformation, App.Title
        Exit Function
    End If
    
    '检查关联字段
    If mstrFLDList <> "" And mstrFLDTree <> "" Then
        '1.是否都设置了
        If InStr(mstrFLDList, "&R") = 0 And InStr(mstrFLDTree, "&R") = 0 Then
            MsgBox "在明细数据和分类数据之间还没有设置互相关联的项目！", vbInformation, App.Title
            Exit Function
        ElseIf InStr(mstrFLDList, "&R") = 0 Then
            MsgBox "在明细数据中还没有设置与分类数据相关联的项目！", vbInformation, App.Title
            Exit Function
        ElseIf InStr(mstrFLDTree, "&R") = 0 Then
            MsgBox "在分类数据中还没有设置与明细数据相关联的项目！", vbInformation, App.Title
            Exit Function
        End If
    
        '2.类型是否相同
        For i = 0 To UBound(Split(mstrFLDList, "|"))
            If InStr(Split(mstrFLDList, "|")(i), "&R") > 0 Then
                lngList = CLng(Split(Split(mstrFLDList, "|")(i), ",")(1))
                Exit For
            End If
        Next
        For i = 0 To UBound(Split(mstrFLDTree, "|"))
            If InStr(Split(mstrFLDTree, "|")(i), "&R") > 0 Then
                lngTree = CLng(Split(Split(mstrFLDTree, "|")(i), ",")(1))
                Exit For
            End If
        Next
        Select Case lngList
            Case adNumeric, adVarNumeric
                If lngTree <> adNumeric And lngTree <> adVarNumeric Then
                    MsgBox "明细数据与分类数据互相关联的项目类型不一致！", vbInformation, App.Title
                    Exit Function
                End If
            Case Else
                If lngList <> lngTree Then
                    MsgBox "明细数据与分类数据互相关联的项目类型不一致！", vbInformation, App.Title
                    Exit Function
                End If
        End Select
    End If
    '检查绑定字段
    '1.是否设置了
    If InStr(mstrFLDList, "&B") = 0 Then
        MsgBox "在明细数据中还没有设置绑定项目！", vbInformation, App.Title
        Exit Function
    End If
    '2.与参数类型是否相同
    For i = 0 To UBound(Split(mstrFLDList, "|"))
        If InStr(Split(mstrFLDList, "|")(i), "&B") > 0 Then
            lngList = CLng(Split(Split(mstrFLDList, "|")(i), ",")(1))
            Exit For
        End If
    Next
    Select Case mbytDataType
        Case 1 '数字型
            If lngList <> adNumeric And lngList <> adVarNumeric Then
                If MsgBox("绑定项目的数据类型与参数定义的数据类型不一致，应该为数字型，是否忽略？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function
            End If
        Case 2 '日期型
            If lngList <> adDBTimeStamp Then
                If MsgBox("绑定项目的数据类型与参数定义的数据类型不一致，应该为日期型，是否忽略？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then Exit Function
            End If
    End Select
    
    '检查显示字段是否设置
    If InStr(mstrFLDList, "&D") = 0 Then
        MsgBox "在明细数据中还没有设置显示项目！", vbInformation, App.Title
        Exit Function
    End If
    
    '检查缺省值相关
    If txtDefBand.Text <> "" Then
        If Trim(txtDefShow.Text) = "" Then
            MsgBox "定义了缺省绑定值则必须同时定义缺省显示内容！", vbInformation, App.Title
            txtDefShow.SetFocus: Exit Function
        End If
    End If
    If Trim(txtDefShow.Text) <> "" And txtDefBand.Text <> "" Then
        Select Case mbytDataType
            Case 1 '数字型
                If Not IsNumeric(txtDefBand.Text) Then
                    MsgBox "缺省绑定值应该为数字类型！", vbInformation, App.Title
                    txtDefBand.SetFocus: Exit Function
                End If
            Case 2 '日期型
                If Not IsDate(txtDefBand.Text) Then
                    MsgBox "缺省绑定值应该为日期类型！", vbInformation, App.Title
                    txtDefBand.SetFocus: Exit Function
                End If
            Case 0 '字符型
'                If txtDefBand.Text = "" Then
'                    MsgBox "字符类型参数必须定义缺省的绑定值！", vbInformation, App.Title
'                    txtDefBand.SetFocus: Exit Function
'                End If
            Case 3 '无类型
'                If txtDefBand.Text = "" Then
'                    If MsgBox("没有定义缺省的绑定值,你确认正确吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
'                        txtDefBand.SetFocus: Exit Function
'                    End If
'                End If
        End Select
    End If
    
    CheckValid = True
End Function

Private Function GetScript(strFld As String, strFiledName As String) As String
'功能：取指定字段的描述
'参数：strFiledName=字段名,strFld=字段描述串
'返回：""=没有这个字段
'说明：返回的描述中已经包括前分隔符","
    Dim i As Integer
    For i = 0 To UBound(Split(strFld, "|"))
        If Split(Split(strFld, "|")(i), ",")(0) = strFiledName Then
            GetScript = "," & Split(Split(strFld, "|")(i), ",")(2)
            Exit Function
        End If
    Next
End Function
