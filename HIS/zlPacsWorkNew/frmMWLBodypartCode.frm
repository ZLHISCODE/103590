VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMWLBodypartCode 
   Caption         =   "Worklist部位对码设置"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   Icon            =   "frmMWLBodypartCode.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除"
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6000
      TabIndex        =   3
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdImportParts 
      Caption         =   "导入PACS部位"
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   350
      Left            =   7560
      TabIndex        =   0
      Top             =   5880
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsListBodyParts 
      Height          =   5655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9495
      _cx             =   16748
      _cy             =   9975
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   200
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin VB.Menu menuPopup 
      Caption         =   "选择类型"
      Visible         =   0   'False
      Begin VB.Menu menuType 
         Caption         =   "空"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMWLBodypartCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng服务ID As Long
Private mfrmParent As Form

Private Enum ColReturn
    col序号 = 0
    ColID
    Col服务ID
    ColPACS部位名称
    Col设备部位名称
    Col设备部位代码
End Enum

Public Sub zlSohwMe(frmParent As Form, lng服务ID As Long)
    mlng服务ID = lng服务ID
    Set mfrmParent = frmParent
    
    Call InitList
    Call FillList
    Me.Show , mfrmParent
End Sub

Private Sub InitList()
'初始化部位对码列表

    With vsListBodyParts
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Rows = 1
        .Cols = 6
        
        .ColWidth(col序号) = 500
        .ColWidth(ColID) = 0
        .ColWidth(Col服务ID) = 0
        .ColWidth(ColPACS部位名称) = 3000
        .ColWidth(Col设备部位名称) = 3000
        .ColWidth(Col设备部位代码) = 3000
        
        .TextMatrix(0, col序号) = "序号"
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col服务ID) = "服务ID"
        .TextMatrix(0, ColPACS部位名称) = "PACS部位名称"
        .TextMatrix(0, Col设备部位名称) = "设备部位名称"
        .TextMatrix(0, Col设备部位代码) = "设备部位代码"
        
        .ColAlignment(col序号) = flexAlignLeftCenter
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col服务ID) = flexAlignLeftCenter
        .ColAlignment(ColPACS部位名称) = flexAlignLeftCenter
        .ColAlignment(Col设备部位名称) = flexAlignLeftCenter
        .ColAlignment(Col设备部位代码) = flexAlignLeftCenter
        
        .Editable = flexEDKbdMouse
    
    End With
End Sub

Private Sub FillList()
'填充部位对码表
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "Select id,服务ID,PACS部位名称,设备部位名称,设备部位代码 From 影像MWL部位对码 Where 服务ID =[1] order by id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取部位对码", mlng服务ID)
    With vsListBodyParts
        .Rows = rsTemp.RecordCount + 1
        While rsTemp.EOF = False
            .TextMatrix(rsTemp.AbsolutePosition, col序号) = rsTemp.AbsolutePosition
            .TextMatrix(rsTemp.AbsolutePosition, ColID) = rsTemp!ID
            .TextMatrix(rsTemp.AbsolutePosition, Col服务ID) = Nvl(rsTemp!服务ID)
            .TextMatrix(rsTemp.AbsolutePosition, ColPACS部位名称) = Nvl(rsTemp!PACS部位名称)
            .TextMatrix(rsTemp.AbsolutePosition, Col设备部位名称) = Nvl(rsTemp!设备部位名称)
            .TextMatrix(rsTemp.AbsolutePosition, Col设备部位代码) = Nvl(rsTemp!设备部位代码)
            rsTemp.MoveNext
        Wend
    End With
    cmdSave.Enabled = False
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    '删除指定行
    Dim strSQL As String
    Dim lngResult As Long
    
    On Error GoTo err
    
    cmdDelete.Enabled = False
    
    '先判断是否需要提示保存
    If cmdSave.Enabled = True Then
        lngResult = MsgBoxD(mfrmParent, "有数据被修改没有保存，是否需要保存？", vbYesNoCancel, "提示信息")
        If lngResult = vbYes Then
            Call SaveDate
        ElseIf lngResult = vbCancel Then
            cmdDelete.Enabled = True
            Exit Sub
        End If
    End If
    
    '如果没有选中任何行，则不动作
    '如果选中行没有ID，说明没有保存到数据库中，直接在列表中删除，如果有ID ，则在数据库中删除，再在表中删除
    
    If Val(vsListBodyParts.TextMatrix(vsListBodyParts.Row, ColID)) <> 0 Then
        strSQL = "ZL_影像MWL部位对码_删除(" & vsListBodyParts.TextMatrix(vsListBodyParts.Row, ColID) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "删除部位对码")
    End If
    
    '重新装载数据
    Call FillList
    
    cmdDelete.Enabled = True
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    If cmdSave.Enabled = True Then
        If MsgBoxD(mfrmParent, "部位对码有改动，确认要放弃改动退出吗？", vbYesNo, "提示信息") = vbNo Then
            Exit Sub
        End If
    End If
    '关闭窗口
    Unload Me
End Sub

Private Sub cmdImportParts_Click()
'导入PACS的检查部位
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim iCount As Integer
    Dim i As Integer
    
    '先卸载原来的菜单
    On Error Resume Next
    iCount = menuType.Count
    If iCount > 1 Then
        menuType(0).Visible = True
        For i = 1 To iCount - 1
            Unload menuType(i)
        Next i
    End If
    
    On Error GoTo err
    '先查询部位的类型，选择类型后导入该类型的全部部位
    strSQL = "Select distinct 类型 from 诊疗检查部位"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查部位类型")
    
    While rsTemp.EOF = False
        iCount = menuType.Count
        Load menuType(iCount)
        menuType(iCount).Caption = Nvl(rsTemp!类型)
        rsTemp.MoveNext
    Wend
    
    If menuType.Count = 1 Then
        MsgBoxD mfrmParent, "诊疗检查部位表中没有部位信息。请先到“诊疗部位管理”模块设置部位。"
    Else
        menuType(0).Visible = False
        PopupMenu menuPopup
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    '保存对码部位
    Call SaveDate
    cmdSave.Enabled = False
    '重新装载
    Call FillList
End Sub

Private Sub SaveDate()
    '保存对码部位
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo err
    
    For i = 1 To vsListBodyParts.Rows - 1
        '有内容才保存
        If vsListBodyParts.TextMatrix(i, ColPACS部位名称) <> "" Then
            
            strSQL = "Zl_影像MWL部位对码_更新(" & _
                    IIf(vsListBodyParts.TextMatrix(i, ColID) = "", "NULL", vsListBodyParts.TextMatrix(i, ColID)) & _
                    "," & mlng服务ID & ",'" & vsListBodyParts.TextMatrix(i, ColPACS部位名称) & "','" & _
                    vsListBodyParts.TextMatrix(i, Col设备部位名称) & "','" & _
                    vsListBodyParts.TextMatrix(i, Col设备部位代码) & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, "保存对码部位")
        End If
    Next i
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    '调整窗口控件的位置
    vsListBodyParts.Left = 0
    vsListBodyParts.Top = 0
    vsListBodyParts.Width = Me.ScaleWidth
    vsListBodyParts.Height = Me.ScaleHeight - cmdExit.Height - 200
    
    cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 100
    cmdDelete.Top = cmdExit.Top
    cmdSave.Top = cmdExit.Top
    cmdImportParts.Top = cmdExit.Top
    
    cmdExit.Left = Me.ScaleWidth - cmdExit.Width - 100
    cmdSave.Left = cmdExit.Left - cmdSave.Width - 200
    cmdDelete.Left = cmdSave.Left - cmdDelete.Width - 200
End Sub

Private Sub menuType_Click(Index As Integer)
'    部位方法的解析说明
'    例子：“0互斥方法,0附加可选方法,1附加可选方法2;1互斥方法2   0正侧位;0切线位;1颅底位;1共用方法”
'
'    1、“TAB” 区分互斥基础方法和共用方法
'    2、“;”区分每一个基础方法
'    3、“,”区分基础方法中的附加方法
'    4、方法前的1位数字代表是否造影
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str方法 As String
    Dim str名称 As String
    Dim arrItem() As String     '保存基础方法组
    Dim arrChild() As String    '保存附加方法组
    Dim lngItem As Long         '基础方法计数器
    Dim lngChild As Long        '附加方法计数器
    Dim strTemp As String
    
    
    On Error GoTo err
    
    If menuType(Index).Caption <> "" Then
        '根据类型查出部位方法的数据集
        strSQL = "Select 名称,方法 From 诊疗检查部位 Where 类型=[1] order by  分组"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取部位方法", CStr(menuType(Index).Caption))
        
        '解析和填写部位方法
        While rsTemp.EOF = False
            str名称 = Nvl(rsTemp!名称)
            str方法 = Nvl(rsTemp!方法)
            If UBound(Split(str方法, vbTab)) >= 0 Then  '=0表示没有互斥方法；>0表示有互斥方法
                arrItem() = Split(Split(str方法, vbTab)(0), ";")    '得到每一个基础方法
                For lngItem = 0 To UBound(arrItem)
                    strTemp = Mid(arrItem(lngItem), 2)
                    If InStr(1, strTemp, ",") > 0 Then  '如果有“，”号，表示包含附加方法，需要进一步解析
                        arrChild = Split(strTemp, ",")
                        strTemp = ""
                        Call AddOneBodypart(str名称 & arrChild(0))
                        For lngChild = 1 To UBound(arrChild)
                            Call AddOneBodypart(str名称 & Mid(arrChild(lngChild), 2))
                        Next lngChild
                    Else
                        Call AddOneBodypart(str名称 & strTemp)
                    End If
                Next lngItem
            End If
            If UBound(Split(str方法, vbTab)) > 0 Then   '>0表示有互斥方法,后面就跟着共用方法，处理共用方法
                arrItem() = Split(Split(str方法, vbTab)(1), ";")
                For lngItem = 0 To UBound(arrItem)
                    Call AddOneBodypart(str名称 & Mid(arrItem(lngItem), 2))
                Next lngItem
            End If
            rsTemp.MoveNext
        Wend
        cmdSave.Enabled = True
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddOneBodypart(strBodypartName As String)
'向列表中添加一个部位方法
'参数： strBodypartName ---部位方法名称
    Dim i As Integer
    
    On Error GoTo err
    
    '首先检查是否有同名的，如果有，就不添加
    For i = 1 To vsListBodyParts.Rows - 1
        If vsListBodyParts.TextMatrix(i, ColPACS部位名称) = strBodypartName Then
            Exit Sub
        End If
    Next i
    
    '添加部位方法
    With vsListBodyParts
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, col序号) = .Rows - 1
        .TextMatrix(.Rows - 1, ColPACS部位名称) = strBodypartName
    End With
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsListBodyParts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    cmdSave.Enabled = True
End Sub

Private Sub vsListBodyParts_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    On Error Resume Next
    
    '回车转移到下一个编辑框
    If KeyAscii = vbKeyReturn Then
        If Col = ColPACS部位名称 Then
            vsListBodyParts.Selecte Row, Col设备部位名称
        ElseIf Col = Col设备部位名称 Then
            vsListBodyParts.Selecte Row, Col设备部位代码
        ElseIf Col = Col设备部位代码 Then   '回车就增加一个新行,并转到下一个编辑框
            If vsListBodyParts.Row = vsListBodyParts.Rows - 1 Then
                vsListBodyParts.TextMatrix(vsListBodyParts.Row, col序号) = vsListBodyParts.Rows - 1
                vsListBodyParts.Rows = vsListBodyParts.Rows + 1
                vsListBodyParts.Select Row + 1, ColPACS部位名称
            End If
        End If
'        vsListBodyParts.EditCell
    End If
End Sub
