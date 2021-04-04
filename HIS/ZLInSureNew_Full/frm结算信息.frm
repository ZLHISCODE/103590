VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm结算信息 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算信息"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frm结算信息.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3090
      TabIndex        =   6
      Top             =   1410
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   5325
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3105
      Width           =   5325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   2
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2865
      TabIndex        =   1
      Top             =   3255
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshBill 
      Height          =   2100
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   930
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3704
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      FixedCols       =   2
      BackColorSel    =   4194304
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lbl 
      Caption         =   "以下为医保病人本次结算的相关信息。"
      Height          =   225
      Index           =   0
      Left            =   810
      TabIndex        =   5
      Top             =   450
      Width           =   4965
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   105
      Picture         =   "frm结算信息.frx":000C
      Stretch         =   -1  'True
      Top             =   165
      Width           =   525
   End
End
Attribute VB_Name = "frm结算信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng结帐ID As Long
Private mblnOK As Boolean
Private mblnYesNo As Boolean
Private mstr结算方式 As String
Private mdbl总费 As Double
Private mblnChange As Boolean '是否更改过值
Private mbytType    As Byte     'mbytType-0门诊挂号,1住院
Private mblnLoad    As Boolean
Private Sub cmdCancel_Click()
    mblnOK = False
    If mblnYesNo = False Then
        mblnOK = True
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    mblnOK = True
    Dim str结算方式 As String
'    If mstr结算方式 <> "" And mblnChange = True Then
'            With mshBill
'                For i = 1 To .Rows - 1
'                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i, 1)) <> "现金" Then
'                        str结算方式 = str结算方式 & "||" & .TextMatrix(i, 1) & " |" & .TextMatrix(i, 2)
'                    End If
'                Next
'            End With
'            If str结算方式 <> "" Then
'                '更新预交记录
'                str结算方式 = Mid(str结算方式, 3)
'                If mbytType = 0 Then
'                    gstrSQL = "zl_病人结算记录_Update(" & mlng结帐ID & ",'" & str结算方式 & "',0)"
'                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
'                Else
'                    gstrSQL = "zl_病人结算记录_Update(" & mlng结帐ID & ",'" & str结算方式 & "',1)"
'                    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新预交记录")
'                End If
'            End If
'    End If
    Unload Me
End Sub

'Modified by 朱玉宝 20031218 地区：福州 新增窗体
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Dim i As Long
    
    
    DebugTool "进入加载窗体"
    strArr = Split(mstr结算方式, "|")
    gstrSQL = "Select Decode(A.记录性质,1,'冲预交',11,'冲预交',A.结算方式) 结算方式,Nvl(A.冲预交,0) 金额 " & _
                " From 病人预交记录 A,保险帐户 B " & _
                " Where A.病人ID=B.病人ID And A.结帐ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取本次交易结算信息", mlng结帐ID)
    DebugTool "进入加载窗体-打开记录集"
    
    With mshBill
        .Clear
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "结算限额"
        .TextMatrix(0, 1) = "结算方式"
        .TextMatrix(0, 2) = "金额"
            
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 1200
        
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
        
    mdbl总费 = 0
    With rsTemp
        Do While Not .EOF
            For i = 0 To UBound(strArr)
                mshBill.RowData(.AbsolutePosition) = 0
                If Split(strArr(i), ":")(0) = Nvl(!结算方式) Then
                    If InStr(1, strArr(i), ":") <> 0 Then
                        mshBill.TextMatrix(.AbsolutePosition, 0) = Val(Split(strArr(i), ":")(1))
                        mshBill.RowData(.AbsolutePosition) = 1
                    End If
                    Exit For
                End If
            Next
            mshBill.TextMatrix(.AbsolutePosition, 1) = !结算方式
            mshBill.TextMatrix(.AbsolutePosition, 2) = Format(!金额, "#####0.00;-#####0.00; ;")
            mdbl总费 = mdbl总费 + Nvl(!金额, 0)
            If mshBill.Rows - 1 = .AbsolutePosition Then mshBill.Rows = mshBill.Rows + 1
            .MoveNext
        Loop
        If Trim(mshBill.TextMatrix(mshBill.Rows - 1, 0)) = "" Then mshBill.Rows = mshBill.Rows - 1
    End With
    DebugTool "进入加载窗体完成"
End Sub

Public Function ShowME(Optional ByVal lng结帐ID As Long = 0, Optional blnYesNo As Boolean = False, Optional str结算方式 As String = "", Optional bytType As Byte = 0) As Boolean
    'blnYesNO:代表是否提供确定和取消选项.
    'str结算方式-对某项结算方式进行更改,格式：结算方式:限制额|结算方式:限制额,如:个人帐户:20。表示个人帐户可以更改，但不能超过20元
    'bytType-0-门诊、挂号,1住院
    
    mlng结帐ID = lng结帐ID
    mblnYesNo = blnYesNo
    mstr结算方式 = str结算方式
    mbytType = bytType
    Me.cmdOK.Visible = blnYesNo
    If blnYesNo = False Then
        Me.cmdCancel.Caption = "确定(&O)"
    End If
    DebugTool "已经进入结算方式showme"
    frm结算信息.Show 1
    DebugTool "完成结算方式showme"
    ShowME = mblnOK
End Function

Private Sub MshBill_DblClick()
    '进行更正相关的数据
    With mshBill
        If .RowData(.Row) = 0 Then txtEdit.Visible = False: Exit Sub
        .COL = 2
        mblnLoad = True
        txtEdit.Left = .Left + .CellLeft + 15
        txtEdit.Top = .Top + .CellTop + 15
        txtEdit.Height = .CellHeight - 30
        txtEdit.Width = .CellWidth - 30
        txtEdit.Visible = True
        txtEdit.Tag = .TextMatrix(.Row, 0)
        txtEdit.Text = .TextMatrix(.Row, 2)
        txtEdit.SetFocus
    End With
End Sub

Private Sub txtEdit_Change()
    If mblnLoad Then Exit Sub
    mblnChange = True
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnLoad = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m金额式
End Sub

Private Sub txtEdit_LostFocus()
    
    txtEdit.Text = Format(Abs(Val(txtEdit.Text)), "#####0.00;-####0.00; ;")
    If Val(txtEdit.Tag) < Val(txtEdit.Text) Then
        ShowMsgbox "输入的值不能大于" & Format(Val(txtEdit.Tag), "#####0.00;-####0.00; ;")
        Exit Sub
    End If
    If mdbl总费 < Val(txtEdit.Text) Then
        ShowMsgbox "输入的值不能大于总费用" & Format(mdbl总费, "#####0.00;-####0.00; ;")
        Exit Sub
    End If
    mshBill.TextMatrix(mshBill.Row, 2) = txtEdit.Text
    Call 整理数据
End Sub
Private Sub 整理数据()
    Dim intRow As Integer
    Dim dblTemp As Double
    Dim i As Integer
    
    intRow = 0
    With mshBill
        dblTemp = mdbl总费
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 1)) = "现金" Then
                intRow = i
            Else
                dblTemp = dblTemp - Val(.TextMatrix(i, 2))
            End If
        Next
        If intRow <> 0 Then
            .TextMatrix(intRow, 2) = Format(dblTemp, "####0.00;-####0.00; ;")
        End If
    End With
End Sub

