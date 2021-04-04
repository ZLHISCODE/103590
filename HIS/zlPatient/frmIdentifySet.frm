VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmIdentifySet 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "认证接口配置"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   Icon            =   "frmIdentifySet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8790
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame frmIdentify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "三方认证接口"
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "取消"
         Height          =   350
         Left            =   7320
         TabIndex        =   3
         Top             =   2760
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         Caption         =   "确定"
         Height          =   350
         Left            =   6120
         TabIndex        =   2
         Top             =   2760
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInterface 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8295
         _cx             =   14631
         _cy             =   4048
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   325
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
   End
   Begin VB.Image imgDelete 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      Picture         =   "frmIdentifySet.frx":6852
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAdd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   495
      Picture         =   "frmIdentifySet.frx":7254
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmIdentifySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mrsSecdInfo As ADODB.Recordset
Private mrsIneterface As New ADODB.Recordset
Public Enum Cert_Interface
    COL_ID = 0
    COL_编号
    COL_接口名
    COL_部件名
    COL_说明
    COL_是否启用
    COL_Add
    COL_Del
End Enum
Private Enum Change_State
    CS_删除行 = -1
    CS_未改变 = 0
    CS_更新行 = 1
    CS_替换行 = 2
    CS_新增行 = 3
End Enum

Private Sub InitVsfGridHeader()
'功能：初始化列表
    Dim strHeader As String
    strHeader = "ID;编号;接口名,2000,1;部件名,2000,1;说明,2800,1;是否启用,900,4;,270,4;,270,4"
    Call grid.Init(vsfInterface, strHeader, , , 1)
    With vsfInterface
        .ColDataType(.ColIndex("是否启用")) = flexDTBoolean
        .TextMatrix(.FixedRows, COL_是否启用) = 0
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveInterface Then
        Unload Me
    End If
End Sub

Private Function SaveInterface() As Boolean
'功能：保存认证接口配置信息
    Dim arrSQL() As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    
    arrSQL = Array()
    On Error GoTo errH
    Call CachCertInterface(arrSQL)
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    SaveInterface = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    Call InitVsfGridHeader
    Call InitBaseInfo
    Call LoadInterface
End Sub

Public Function ShowMe(frmParent As Object) As Boolean
    Set mfrmParent = frmParent
    If Not mfrmParent Is Nothing Then
        Me.Show , mfrmParent
    End If
End Function

Private Function LoadInterface() As Boolean
'加载三方接口信息
    On Error GoTo errH
    Set mrsIneterface = LoadCertInterface(1)
    If Not mrsIneterface.EOF Then
        Call LoadCachInterface(mrsIneterface)
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadCachInterface(ByVal rsTmp As ADODB.Recordset)
'功能：将证件信息加载并缓存

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim strsInfo As String, strsMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim lngTmp As Long
    Dim lngsTmp As Long
    Dim strType As String
    Dim rsImg As New ADODB.Recordset
    Dim strFile As String
    Dim objFile As New FileSystemObject
    
    On Error GoTo errH
    
     '删除之前的缓存
    mrsSecdInfo.Filter = "控件名='vsfInterface'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If

    lngTmp = 1
    With vsfInterface
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_ID) = "" & rsTmp!ID
            .TextMatrix(lngRow, COL_编号) = "" & rsTmp!编号
            .TextMatrix(lngRow, COL_接口名) = "" & rsTmp!接口名
            .TextMatrix(lngRow, COL_部件名) = "" & rsTmp!部件名
            .TextMatrix(lngRow, COL_说明) = "" & rsTmp!说明
            .TextMatrix(lngRow, COL_是否启用) = IIf(Val("" & rsTmp!是否启用) = 1, -1, 0)
            .Cell(flexcpPicture, lngRow, COL_Add, lngRow, COL_Add) = imgAdd
            .Cell(flexcpPictureAlignment, lngRow, COL_Add, lngRow, COL_Add) = 4
            
            .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
            .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
            
            .RowData(lngRow) = Val(rsTmp!ID & "")
                
            strMainInfo = rsTmp!ID & "|" & rsTmp!编号 & "|" & rsTmp!接口名 & "|" & rsTmp!部件名 & "|" & rsTmp!说明 & "|" & Val("" & rsTmp!是否启用)
            strInfo = strMainInfo
            mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值"), Array(lngTmp, Val(rsTmp!ID & ""), "vsfInterface", strInfo, strMainInfo)
            lngTmp = lngTmp + 1
            rsTmp.MoveNext
        Next
        .Row = 1: .Col = COL_接口名
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitBaseInfo()
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '本记录集的主键
        .Fields.Append "序号", adInteger                                              '标识信息，引用主记录集
        .Fields.Append "控件名", adVarChar, 100                                       '展示信息的控件名称
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable                        '行号或控件数组Index
        .Fields.Append "页码", adInteger                                              '信息所在的页码
        .Fields.Append "原ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable      '信息在加载时的值
        .Fields.Append "主信息原值", adVarChar, 2000, adFldIsNullable    '信息的主要部分，标识一个信息是否被彻底改变，信息在加载时的值
        .Fields.Append "现ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable      '信息在检查时的值
        .Fields.Append "主信息现值", adVarChar, 2000, adFldIsNullable    '信息在检查时的值
        .Fields.Append "改变状态", adInteger                             '信息改变程度0-未改变，1-次级信息改变，2-主信息改变,3-新增,-1，删除
        .Fields.Append "ID", adBigInt, , adFldIsNullable                 '信息行在数据库中的ID,一般表格类控件使用
        .Fields.Append "Tag", adVarChar, 2000                            '存储额外数据
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Private Function CachCertInterface(ByRef arrSQL As Variant) As Boolean
'功能：将证件信息缓存
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim strDels As String
    Dim strAll As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    
    On Error GoTo errH
    With vsfInterface
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_接口名) <> "" And .TextMatrix(i, COL_部件名) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_ID)) & "|" & .TextMatrix(i, COL_编号) & "|" & .TextMatrix(i, COL_接口名) & "|" & .TextMatrix(i, COL_部件名) & "|" & .TextMatrix(i, COL_说明) & "|" & IIf(Nvl(.TextMatrix(i, COL_是否启用), -1) = -1, 1, 0)
                strInfo = strMainInfo
                .RowData(i) = .TextMatrix(i, COL_ID)
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '相同过每记录
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, COL_接口名) = &HC0C0FF
                    Call .ShowCell(i, COL_接口名)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '收集所有用于判断是否有重复行
                End If
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "控件名='vsfInterface' and 序号=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!序号 = lngTmp
                   mrsSecdInfo!控件名 = "vsfInterface"
               End If
               mrsSecdInfo!现ID = Val(.RowData(i))
               mrsSecdInfo!信息现值 = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!主信息现值 = IIf(strMainInfo = "", Null, strMainInfo)
               mrsSecdInfo!IndexEx = i
               mrsSecdInfo.Update
               lngTmp = lngTmp + 1

               mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "控件名='vsfInterface'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
        

        '主信息改变行需要调用删除方法
        mrsSecdInfo.Filter = "(改变状态=" & CS_删除行 & " And 控件名='vsfInterface')" ' OR (改变状态=" & CS_替换行 & " And 控件名='vsfCert')"
        Do While Not mrsSecdInfo.EOF
            strDels = "" & mrsSecdInfo!原ID
            If strDels <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_实名认证接口_Delete(" & Val(strDels) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop

        '主信息改变以及新增行需要调用插入过程        '次级信息改变，调用更新过程
        mrsSecdInfo.Filter = "控件名='vsfInterface' And 改变状态>" & CS_未改变
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mrsSecdInfo!改变状态 = CS_新增行 Then
                arrSQL(UBound(arrSQL)) = "Zl_实名认证接口_Insert(" & "'" & .TextMatrix(lngRow, COL_接口名) & "','" & .TextMatrix(lngRow, COL_部件名) & "','" & .TextMatrix(lngRow, COL_说明) & "'," & IIf(Nvl(.TextMatrix(lngRow, COL_是否启用), -1) = -1, 1, 0) & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_实名认证接口_Update(" & Val(.TextMatrix(lngRow, COL_ID)) & ",'" & .TextMatrix(lngRow, COL_编号) & "','" & .TextMatrix(lngRow, COL_接口名) & "','" & _
                        .TextMatrix(lngRow, COL_部件名) & "','" & .TextMatrix(lngRow, COL_说明) & "'," & IIf(Nvl(.TextMatrix(lngRow, COL_是否启用), -1) = -1, 1, 0) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop
    End With
    CachCertInterface = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmParent = Nothing
    Set mrsSecdInfo = Nothing
    Set mrsIneterface = Nothing
End Sub

Private Sub vsfInterface_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewCol As Long, lngNewRow As Long
    
    lngNewCol = NewCol
    lngNewRow = NewRow
    If lngNewCol = -1 Then Exit Sub
    With vsfInterface
        If lngNewCol = COL_Del Or lngNewCol = COL_Add Then
             .ComboList = "..."
             .FocusRect = flexFocusNone
             Set .CellButtonPicture = IIf(lngNewCol = COL_Del, imgDelete, imgAdd)
        Else
            .ComboList = ""
        End If
        If lngNewRow >= .FixedRows Then
            '显示图片
            If lngNewCol <> COL_Add And .TextMatrix(lngNewRow, COL_接口名) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '下一行诊断为空则不能新增行
                    If .TextMatrix(lngNewRow + 1, COL_接口名) = "" Then
                         Set .Cell(flexcpPicture, lngNewRow, COL_Add) = imgAdd
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, COL_Add) = imgAdd
                End If
            End If
            '显示图片
            If lngNewCol <> COL_Del Then Set .Cell(flexcpPicture, lngNewRow, COL_Del) = imgDelete
        End If
    End With
    
End Sub

Private Sub vsfInterface_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = COL_Add Or lngCol = COL_Del Then Cancel = True
End Sub

Private Sub vsfInterface_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long, lngCount As Long
    Dim i As Long, j As Long
    Dim blnAdd As Boolean
    
    lngCol = Col
    lngRow = Row
    With vsfInterface
        Select Case lngCol
            Case COL_Add
                For i = .Rows - 1 To .FixedRows Step -1
                    If Trim(.TextMatrix(.Rows - 1, COL_接口名)) <> "" And .RowHidden(.Rows - 1) = False Then
                        blnAdd = True
                        Exit For
                    ElseIf Trim(.TextMatrix(.Rows - 1, COL_接口名)) = "" And .RowHidden(.Rows - 1) = False Then
                        Exit For
                    End If
                Next
                If blnAdd Then
                     lngRow = .Rows: .AddItem "", lngRow
                     .TextMatrix(lngRow, COL_是否启用) = 0
                     .Row = lngRow: .Col = COL_接口名
                     .ShowCell .Row, COL_接口名
                End If
                blnAdd = False
            Case COL_Del
                If Trim(.TextMatrix(lngRow, COL_接口名)) <> "" Then
                    If MsgBox("确定要删除接口名为【" & .TextMatrix(lngRow, COL_接口名) & "】的证件信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then
                            If .Rows - 1 = .FixedRows Then
                                For i = COL_ID To COL_是否启用
                                    .TextMatrix(lngRow, i) = ""
                                    .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                                Next
                            ElseIf .Rows - 1 > .FixedRows Then
                                .RemoveItem lngRow
                                .AddItem "", lngRow
                                .RowHidden(lngRow) = True
                            End If
                        Else
                             For i = .FixedRows To .Rows - 1
                                If .TextMatrix(i, COL_接口名) <> "" Then
                                    lngCount = lngCount + 1
                                End If
                            Next
                            If lngCount = .FixedRows Then
                                For j = COL_ID To COL_是否启用
                                    .TextMatrix(lngRow, j) = ""
                                    .Cell(flexcpData, lngRow, j, lngRow, j) = ""
                                Next
                            End If
                        End If
                    Else
                        .Row = lngRow: .Col = COL_接口名
                        .ShowCell .Row, .Col
                    End If
                Else
                    If .Rows - 1 = .FixedRows Or lngRow = .FixedRows Then
                        Exit Sub
                    Else
                        For i = .FixedRows To .Rows - 1
                            If .TextMatrix(i, COL_接口名) <> "" Then
                                lngCount = lngCount + 1
                            End If
                        Next
                        If lngCount <> 0 Then
                            .RemoveItem lngRow
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfInterface_Click()
    Dim lngRow As Long, lngCol As Long
    
    With vsfInterface
        lngRow = .Row
        lngCol = .Col
        If (lngCol = COL_Add Or lngCol = COL_Del) And lngRow >= .FixedRows Then
            If lngCol = COL_Add Then
                If .TextMatrix(lngRow, COL_接口名) = "" Then Exit Sub
            End If
            .Select lngRow, lngCol
            Call vsfInterface_CellButtonClick(lngRow, lngCol)
        End If
    End With
End Sub

Private Sub vsfInterface_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
      
    lngRow = vsfInterface.Row
    lngCol = vsfInterface.Col
    With vsfInterface
        If lngCol = COL_部件名 Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_部件名)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        ElseIf lngCol = COL_接口名 Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_接口名)) >= 50 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        ElseIf lngCol = COL_说明 Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_说明)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub


Private Sub vsfInterface_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim lngCol As Long
      
    lngRow = Row
    lngCol = Col
    With vsfInterface
        If lngCol = COL_部件名 Then
            .TextMatrix(lngRow, COL_部件名) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_部件名)) >= 100 Then
                MsgBox "部件名的字符个数不能大于100个字符或者50个汉字！", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_接口名 Then
            .TextMatrix(lngRow, COL_接口名) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_接口名)) >= 50 Then
                MsgBox "接口名的字符个数不能大于50个字符或者25个汉字！", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_说明 Then
            .TextMatrix(lngRow, COL_说明) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_说明)) >= 100 Then
                MsgBox "说明的字符个数不能大于100个字符或者50个汉字！", vbInformation, gstrSysName
                Cancel = True
            End If
        End If
    End With
End Sub


