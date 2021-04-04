VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm多病种选择_自贡 
   Caption         =   "多病种选择"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   Icon            =   "frm多病种选择_自贡.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6315
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   90
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   6825
      TabIndex        =   8
      Top             =   3690
      Width           =   6825
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9090
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5790
      Width           =   9090
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   6
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5265
         TabIndex        =   5
         Top             =   105
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   780
         MaxLength       =   6
         TabIndex        =   4
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   630
      End
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   9090
      TabIndex        =   1
      Top             =   0
      Width           =   9090
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   2430
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3195
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   5636
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1065
      Top             =   3600
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
            Picture         =   "frm多病种选择_自贡.frx":1CFA
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   3810
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm多病种选择_自贡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'参数
Private mintCol_Key As Integer
Private mblnOK As Boolean
Private mblnShow As Boolean
Private mrsSel As ADODB.Recordset
Private mrsReturn As New ADODB.Recordset
Private mstrKey  As String
Private mcnObject As ADODB.Connection
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub

Public Function ShowSelect(rsSelect As ADODB.Recordset, ByVal strKey As String, _
    Optional ByVal strTitle As String, Optional ByVal strNote As String, _
    Optional ByVal rsSelected As ADODB.Recordset, Optional ByVal blnShow As Boolean = True, _
    Optional ByVal cnObject As ADODB.Connection) As Boolean
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     rsSelect=选择的记录集
'     strKey=主关键字段
'     strTitle=选择器类型命名
'     strNote=选择说明
'     blnMutilSelect=多选标志（如果允许多选，则固定为第一列是选择列）
    Dim lngIndex As Long
    Dim strValue As String, strFilter As String
    Dim lngRow As Long, intCol As Integer
    Dim arrSelected
    
    Set mrsSel = rsSelect
    mstrKey = strKey
    mblnOK = False
    mblnShow = blnShow
    
    If rsSelect.RecordCount = 0 Then
        MsgBox "没有可选择的数据", vbInformation, gstrSysName
        Exit Function
    End If
    
    '构造列头
    mshSelect.Clear
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then
            mintCol_Key = intCol
            Exit For
        End If
    Next
    
    '如果传入有已选择的内容，则显示（肯定只有多选的才存在）
    Set mshSelect.DataSource = mrsSel
    If mshSelect.Rows = 1 Then mshSelect.Rows = 2
    If Not mblnShow Then
        mshSelect.Rows = 2
        For intCol = 0 To mshSelect.Cols - 1
            mshSelect.TextMatrix(1, intCol) = ""
        Next
    End If
    
    If Not rsSelected Is Nothing Then
        Set mshSelected.DataSource = rsSelected
        If mshSelected.Rows = 1 Then
            mshSelected.Rows = 2
        Else
            mshSelected.Rows = mshSelected.Rows + 1
        End If
    End If
    
    '需要还原
    Call zlControl.MshSetColWidth(mshSelect, Me)
    mshSelect.ColWidth(mintCol_Key) = 0
    
    mshSelect.Row = 1
    mshSelect.RowSel = 1
    mshSelect.Col = 0
    mshSelect.ColSel = mshSelect.Cols - 1
    
    mshSelected.Row = 1
    mshSelected.RowSel = 1
    mshSelected.Col = 0
    mshSelected.ColSel = mshSelected.Cols - 1
    
    If cnObject Is Nothing Then
        Set mcnObject = gcnOracle
    Else
        Set mcnObject = cnObject
    End If
    
    frm多病种选择_自贡.Show vbModal
    If mblnOK Then Set rsSelect = mrsReturn
    ShowSelect = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intCol As Integer
    Dim lngRow As Long
    Dim strFilter As String
    Dim strFields As String, strValues As String
    
    '初始化记录集
    strFields = ""
    For intCol = 0 To mrsSel.Fields.Count - 1
        strFields = strFields & "|" & mrsSel.Fields(intCol).Name & "," & adLongVarChar & "," & mrsSel.Fields(intCol).DefinedSize
    Next
    strFields = Mid(strFields, 2)
    Call Record_Init(mrsReturn, strFields)
    
    '根据传入记录集产生对应的记录集并返回
    strFields = ""
    For intCol = 0 To mrsSel.Fields.Count - 1
        strFields = strFields & "|" & mrsSel.Fields(intCol).Name
    Next
    strFields = Mid(strFields, 2)
    
    With mshSelected
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, mintCol_Key)) <> "" Then
                strValues = ""
                For intCol = 0 To .Cols - 1
                    strValues = strValues & "|" & .TextMatrix(lngRow, intCol)
                Next
                strValues = Mid(strValues, 2)
                
                Call Record_Add(mrsReturn, strFields, strValues)
            End If
        Next
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mshSelect.Rows = 2 And mblnShow Then cmdOK_Click
End Sub

Private Sub Form_Resize()
    Dim intCol As Integer
    On Error Resume Next
    picSplit.Width = Me.ScaleWidth
    
    mshSelect.Top = picInfo.Height
    mshSelect.Left = 0
    mshSelect.Width = Me.ScaleWidth
    mshSelect.Height = picSplit.Top - mshSelect.Top
    
    mshSelected.Top = picSplit.Top + picSplit.Height
    mshSelected.Left = 0
    mshSelected.Width = Me.ScaleWidth
    mshSelected.Height = Me.ScaleHeight - picInfo.Height - mshSelected.Top - 200
    
    '设置列宽度
    For intCol = 0 To mshSelect.Cols - 1
        mshSelected.ColWidth(intCol) = mshSelect.ColWidth(intCol)
    Next
    
    'If Me.ScaleWidth - cmdCancel.Width * 1.3 >= cmdHelp.Left + cmdHelp.Width * 2 + 300 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.1
    'End If
End Sub

Private Sub lvw_DblClick()
    Call cmdOK_Click
End Sub

Private Sub mshSelect_DblClick()
    Dim intCol As Integer, lngRow As Long
    
    '先检查是否已经选择
    For lngRow = 1 To mshSelected.Rows - 1
        If mshSelected.TextMatrix(lngRow, mintCol_Key) = mshSelect.TextMatrix(mshSelect.Row, mintCol_Key) Then
            MsgBox "已经选择了该病种，不能重复选择！", vbInformation, gstrSysName
            Exit Sub
        End If
    Next
    
    '另入到选择器中
    For intCol = 0 To mshSelect.Cols - 1
        mshSelected.TextMatrix(mshSelected.Rows - 1, intCol) = mshSelect.TextMatrix(mshSelect.Row, intCol)
    Next
    mshSelected.Rows = mshSelected.Rows + 1
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mshSelect_DblClick
End Sub

Private Sub mshSelected_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Trim(mshSelected.TextMatrix(mshSelected.Row, mintCol_Key)) <> "" Then
            mshSelected.RemoveItem mshSelected.Row
        End If
    End If
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Move picSplit.Left, picSplit.Top + Y
    Call Form_Resize
End Sub

Private Function EmptyContent() As Boolean
    Dim intCol As Integer
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then Exit For
    Next
    If intCol > mrsSel.Fields.Count - 1 Then intCol = 0
    
    With mshSelect
        If .Rows - 1 = 1 And Val(.TextMatrix(1, intCol)) = 0 Then EmptyContent = True
    End With
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intTYPE As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intTYPE = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intTYPE
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intTYPE, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：根据用户输入的内容查找匹配的内容
'注意，为了提高用户输入的速度，特增加了按简码匹配，请传入记录集时，一定要有简码字段
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Dim strFind As String
    Dim strSql As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    strFind = strFind & "%"
    
    strSql = mrsSel.Source
    mrsSel.Close
    mrsSel.CursorLocation = adUseClient
    mrsSel.Open strSql & " And (SickNum Like '%" & strFind & "' Or SickName Like '%" & strFind & "' Or SickSpell Like '%" & strFind & "')", mcnObject
    
    If mrsSel.RecordCount = 0 Then
        Exit Sub
    Else
        Set mshSelect.DataSource = mrsSel
    End If
    
    Call zlControl.MshSetColWidth(mshSelect, Me)
    mshSelect.ColWidth(mintCol_Key) = 0
End Sub
