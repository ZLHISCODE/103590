VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmListSel 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmListSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3195
      Left            =   150
      TabIndex        =   1
      Top             =   480
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
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   6840
      TabIndex        =   7
      Top             =   0
      Width           =   6840
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   120
         Width           =   2430
      End
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
      ScaleWidth      =   6840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6840
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   780
         MaxLength       =   6
         TabIndex        =   6
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5265
         TabIndex        =   4
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4035
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "查找(&F)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   630
      End
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
            Picture         =   "frmListSel.frx":014A
            Key             =   "Item"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'参数
Private mint险类 As Integer
Private mblnOK As Boolean
Private mblnMutilSelect As Boolean
Private mrsSel As ADODB.Recordset
Private mrsReturn As New ADODB.Recordset
Private mstrKey  As String
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub

Public Function ShowSelect(ByVal intInsure As Integer, rsSelect As ADODB.Recordset, ByVal strKey As String, _
    Optional ByVal strTitle As String, Optional ByVal strNote As String, Optional ByVal blnMutilSelect As Boolean = False) As Boolean
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     rsSelect=选择的记录集
'     strKey=主关键字段
'     strTitle=选择器类型命名
'     strNote=选择说明
'     blnMutilSelect=多选标志（如果允许多选，则固定为第一列是选择列）
    Dim lngIndex As Long
    Dim strValue As String
    Dim lngRow As Long, intCol As Integer
    Set mrsSel = rsSelect
    mstrKey = strKey
    mblnMutilSelect = blnMutilSelect
    mint险类 = intInsure
    mblnOK = False
    
    If rsSelect.RecordCount = 0 Then
        MsgBox "没有可选择的数据", vbInformation, gstrSysName
        Exit Function
    End If
    
    '构造列头
    mshSelect.Clear
    mshSelect.TextMatrix(0, 0) = "Key"
    If mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
        mshSelect.Cols = 11
        mshSelect.TextMatrix(0, 1) = "名称"
        mshSelect.TextMatrix(0, 2) = "规格"
        mshSelect.TextMatrix(0, 3) = "单位"
        mshSelect.TextMatrix(0, 4) = "大类"
        mshSelect.TextMatrix(0, 5) = "发票项目"
        mshSelect.TextMatrix(0, 6) = "化学名"
        mshSelect.TextMatrix(0, 7) = "商品名"
        mshSelect.TextMatrix(0, 8) = "厂家"
        mshSelect.TextMatrix(0, 9) = "剂型"
        mshSelect.TextMatrix(0, 10) = "简码"
        
    End If
    
    '装入数据
    If mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市 Then
        lngRow = 1
        rsSelect.MoveFirst
        Do Until rsSelect.EOF
            If lngRow > 1 Then mshSelect.Rows = mshSelect.Rows + 1
            mshSelect.TextMatrix(lngRow, 0) = rsSelect.Fields(strKey).Value
            strValue = IIf(IsNull(rsSelect!附注), "", rsSelect!附注)
            mshSelect.TextMatrix(lngRow, 1) = IIf(IsNull(rsSelect!名称), "", rsSelect!名称)
            mshSelect.TextMatrix(lngRow, 2) = Split(strValue, "|")(0)
            mshSelect.TextMatrix(lngRow, 3) = Split(strValue, "|")(1)
            mshSelect.TextMatrix(lngRow, 4) = IIf(IsNull(rsSelect!大类), "", rsSelect!大类)
            mshSelect.TextMatrix(lngRow, 5) = Split(strValue, "|")(2)
            If UBound(Split(strValue, "|")) >= 4 Then
                mshSelect.TextMatrix(lngRow, 6) = Split(strValue, "|")(4)
            End If
            If UBound(Split(strValue, "|")) >= 5 Then
                mshSelect.TextMatrix(lngRow, 7) = Split(strValue, "|")(5)
            End If
            If UBound(Split(strValue, "|")) >= 6 Then
                mshSelect.TextMatrix(lngRow, 8) = Split(strValue, "|")(6)
            End If
            If UBound(Split(strValue, "|")) >= 7 Then
                mshSelect.TextMatrix(lngRow, 9) = Split(strValue, "|")(7)
            End If
            mshSelect.TextMatrix(lngRow, 10) = IIf(IsNull(rsSelect!简码), "", rsSelect!简码)
            
            lngRow = lngRow + 1
            rsSelect.MoveNext
        Loop
    Else
        Set mshSelect.DataSource = rsSelect
    End If
    
    Call zlControl.MshSetColWidth(mshSelect, Me)
    If Not (mint险类 = TYPE_福建巨龙 Or mint险类 = TYPE_福建省 Or mint险类 = TYPE_福州市 Or mint险类 = TYPE_南平市) Then
        For intCol = 0 To mrsSel.Fields.Count - 1
            If mrsSel.Fields(intCol).Name = mstrKey And Not (mstrKey Like "*编码*") Then
                mshSelect.ColWidth(intCol) = 0
                Exit For
            End If
        Next
    End If
    mshSelect.Row = 1
    mshSelect.RowSel = 1
    mshSelect.Col = 0
    mshSelect.ColSel = mshSelect.Cols - 1
    
    Me.Caption = strTitle
    Me.lblInfo = strNote
    frmListSel.Show vbModal
    If mblnMutilSelect And mblnOK Then Set rsSelect = mrsReturn
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
    If EmptyContent Then Exit Sub
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then Exit For
    Next
    If intCol > mrsSel.Fields.Count - 1 Then intCol = 0
    If mblnMutilSelect = False Then
        If mrsSel.Fields(mstrKey).Type = adVarChar Or mrsSel.Fields(mstrKey).Type = adChar Or mrsSel.Fields(mstrKey).Type = adWChar Or mrsSel.Fields(mstrKey).Type = adVarWChar Or mrsSel.Fields(mstrKey).Type = adLongVarChar Or mrsSel.Fields(mstrKey).Type = adLongVarWChar Then
            strFilter = mstrKey & "='" & mshSelect.TextMatrix(mshSelect.Row, intCol) & "'"
        Else
            strFilter = mstrKey & "=" & Val(mshSelect.TextMatrix(mshSelect.Row, intCol))
        End If
        mrsSel.Filter = strFilter
    Else
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
        
        With mshSelect
            For lngRow = 1 To .Rows - 1
                If Trim(mshSelect.TextMatrix(lngRow, 0)) = "√" Then
                    mrsSel.MoveFirst
                    mrsSel.Move lngRow - 1
                    
                    strValues = ""
                    For intCol = 0 To mrsSel.Fields.Count - 1
                        strValues = strValues & "|" & mrsSel.Fields(intCol).Value
                    Next
                    strValues = Mid(strValues, 2)
                    
                    Call Record_Add(mrsReturn, strFields, strValues)
                End If
            Next
        End With
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
'    If mshSelect.Rows = 2 Then cmdOK_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    mshSelect.Top = picInfo.Height
    mshSelect.Left = 0
    mshSelect.Width = Me.ScaleWidth
    mshSelect.Height = Me.ScaleHeight - picInfo.Height - picCmd.Height
    
    'If Me.ScaleWidth - cmdCancel.Width * 1.3 >= cmdHelp.Left + cmdHelp.Width * 2 + 300 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.1
    'End If
End Sub

Private Sub lvw_DblClick()
    Call cmdOK_Click
End Sub

Private Sub mshSelect_DblClick()
    If mblnMutilSelect = False Then
        Call cmdOK_Click
    Else
        If mshSelect.TextMatrix(mshSelect.Row, 0) = "" Then
            mshSelect.TextMatrix(mshSelect.Row, 0) = "√"
        Else
            mshSelect.TextMatrix(mshSelect.Row, 0) = ""
        End If
    End If
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mshSelect_DblClick
End Sub

Private Sub txtFind_Change()
'功能：根据用户输入的内容查找匹配的内容
    Dim lngIndex As Long, lngRow As Long, lngCol As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    strFind = strFind & "*"
    If EmptyContent Then Exit Sub
    
    With mshSelect
        For lngRow = 1 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                If .TextMatrix(lngRow, lngCol) Like strFind Then
                    .Row = lngRow
                    .RowSel = lngRow
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End If
            Next
        Next
    End With
End Sub

Private Function EmptyContent() As Boolean
    Dim intCol As Integer
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then Exit For
    Next
    If intCol > mrsSel.Fields.Count - 1 Then intCol = 0
    
    '调试重庆医保银海版 204-04-07
    With mshSelect
        If .Rows - 1 = 1 Then
            If mrsSel.Fields(mstrKey).Type = adVarChar Or mrsSel.Fields(mstrKey).Type = adChar Or mrsSel.Fields(mstrKey).Type = adWChar Or mrsSel.Fields(mstrKey).Type = adVarWChar Or mrsSel.Fields(mstrKey).Type = adLongVarChar Or mrsSel.Fields(mstrKey).Type = adLongVarWChar Then
                If .TextMatrix(1, intCol) = "" Then EmptyContent = True
            Else
                If Val(.TextMatrix(1, intCol)) = 0 Then EmptyContent = True
            End If
        End If
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
    Dim strFieldName As String, intType As Integer, lngLength As Long
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
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
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
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
End Sub
