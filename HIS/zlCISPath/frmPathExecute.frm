VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPathExecute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "执行路径项目"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9165
   Icon            =   "frmPathExecute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9165
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9165
      TabIndex        =   4
      Top             =   0
      Width           =   9165
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   90
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathExecute.frx":6852
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "请填写路径项目的执行结果和执行说明，执行结果可能做为路径评估的依据。"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6240
      Width           =   9165
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7800
         TabIndex        =   3
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5385
      Left            =   50
      TabIndex        =   1
      Top             =   860
      Width           =   9020
      _cx             =   15910
      _cy             =   9499
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathExecute.frx":6FA7
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPathExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun         As Long             '0-批量执行,1-单个执行,2-批量取消执行
Private mblnOK          As Boolean

Private mPP             As TYPE_PATH_Pati
Private mPati           As TYPE_Pati
Private mint场合        As Integer          'int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Private mlng路径执行ID  As Long             '单个执行时才传入
Private mblnNurse       As Boolean          'mlngFun=2时,=False 允许批量取消执行者为护士的项目,=True 只允许批量取消生成者是护士且执行者是护士的项目
Private mcol            As Collection
Private mrsItem         As ADODB.Recordset
Private mfrmParent      As Object
Private mintMode        As Integer

Private Enum E执行结果
    E已经执行 = 1
    E尚未执行 = 2
    E取消执行 = 3
    E部分执行 = 4
    E提前执行 = 5
    E延后执行 = 6
End Enum

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal lng路径执行ID As Long, ByVal int场合 As Long, Optional ByVal blnNurse As Boolean = False, Optional ByVal intMode As Integer) As Boolean
    
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mPati = t_pati
    mPP = t_pp
    mlng路径执行ID = lng路径执行ID
    mint场合 = int场合
    mblnNurse = blnNurse
    mintMode = intMode
    
    If intMode = 1 Then
        Set mrsItem = GetItemOut
    Else
        Set mrsItem = GetItem
    End If
    If mrsItem.RecordCount = 0 Then
        MsgBox "该病人没有需要由你执行的项目。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetItem() As ADODB.Recordset
'功能：读取待执行的记录
    Dim strSql As String
    Dim strUType As String, strIF As String
    Dim strcol As String
    
    If mint场合 = 0 Then
        strUType = " And nvl(Nvl(b.执行者, a.执行者),1) = 1"
    ElseIf mint场合 = 1 Then
        If mlngFun = 2 And mblnNurse Then
            strUType = " And nvl(Nvl(b.生成者, a.生成者),1) = 2 And nvl(Nvl(b.执行者, a.执行者),2) = 2"
        Else
            strUType = " And nvl(Nvl(b.执行者, a.执行者),2) = 2"
        End If
    End If
    
    If mlngFun = 0 Then
        strIF = "a.路径记录id = [1] And a.阶段id = [2] And a.天数 = [3] And Nvl(a.生成时间性质,0)<>2 And a.执行时间 Is Null"
    ElseIf mlngFun = 2 Then
        strIF = "a.路径记录id = [1] And a.阶段id = [2] And a.天数 = [3]  And a.执行时间 Is Not Null"
    Else
        strIF = "a.id = [4] And Nvl(a.生成时间性质,0)<>2 And a.执行时间 Is Null"
    End If
    If mlngFun = 2 Then
        strcol = "a.执行结果,a.执行说明,"
    Else
        strcol = "Nvl(b.项目结果, a.项目结果) As 执行结果,"
    End If
    'Distinct是因为一个项目可能定义了多个医嘱内容
    strSql = "Select Distinct a.分类,a.ID, c.序号 类别顺序, Nvl(b.项目序号, a.项目序号) 项目顺序, Nvl(b.项目内容, a.项目内容) As 项目内容, " & strcol & " Nvl(b.图标id, a.图标id) 图标id," & _
            "Decode(d.医嘱内容ID,Null,0,1) as 医嘱项目,Decode(e.病人医嘱ID,Null,0,1) as 有医嘱" & vbNewLine & _
            "From 病人路径执行 A,临床路径项目 B,临床路径分类 C,临床路径医嘱 D,病人路径医嘱 E,病人合并路径 F" & vbNewLine & _
            "Where " & strIF & " And f.首要路径记录id(+) = a.路径记录id  And (f.路径id = c.路径id And f.版本号 = c.版本号  or c.路径id = [5] And c.版本号 = [6])" & _
            " And  a.分类 = c.名称 And NVL(c.分支id,0)=NVL(b.分支ID,0) And b.ID = d.路径项目ID(+) And a.ID = E.路径执行ID(+)" & vbNewLine & _
            "And a.项目id = b.Id(+)" & strUType & vbNewLine & _
            "Order by c.序号,Nvl(b.项目序号,a.项目序号)"
    On Error GoTo errH
    Set GetItem = zlDatabase.OpenSQLRecord(strSql, "读取待执行的项目", mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数, mlng路径执行ID, mPP.路径ID, mPP.版本号)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemOut() As ADODB.Recordset
'功能：读取待执行的记录
    Dim strSql As String
    Dim strIF As String
    Dim strcol As String
    
    If mlngFun = 0 Then
        strIF = "a.路径记录id = [1] And a.阶段id = [2] And a.天数 = [3] And a.执行时间 Is Null"
    ElseIf mlngFun = 2 Then
        strIF = "a.路径记录id = [1] And a.阶段id = [2] And a.天数 = [3] And a.执行时间 Is Not Null"
    Else
        strIF = "a.id = [4] And a.执行时间 Is Null"
    End If
    If mlngFun = 2 Then
        strcol = "a.执行结果,a.执行说明,"
    Else
        strcol = "Nvl(b.项目结果, a.项目结果) As 执行结果,"
    End If
    
    strSql = " Select Distinct a.分类, a.Id, c.序号 类别顺序, Nvl(b.项目序号, a.项目序号) 项目顺序, Nvl(b.项目内容, a.项目内容) As 项目内容," & strcol & "Nvl(b.图标id, a.图标id) 图标id," & vbNewLine & _
             "                Decode(d.医嘱内容id, Null, 0, 1) As 医嘱项目, Decode(e.病人医嘱id, Null, 0, 1) As 有医嘱" & vbNewLine & _
             " From 病人门诊路径执行 A, 门诊路径项目 B, 门诊路径分类 C, 门诊路径医嘱 D, 病人门诊路径医嘱 E" & vbNewLine & _
             " Where  " & strIF & " And (c.路径id = [5] And c.版本号 = [6]) And a.分类 = c.名称  And b.Id = d.路径项目id(+) And" & vbNewLine & _
             "      a.Id = e.路径执行id(+) And a.项目id = b.Id(+)" & vbNewLine & _
             " Order By c.序号, Nvl(b.项目序号, a.项目序号)"

    On Error GoTo errH
    Set GetItemOut = zlDatabase.OpenSQLRecord(strSql, "读取待执行的项目", mPP.病人路径ID, mPP.当前阶段ID, mPP.当前天数, mlng路径执行ID, mPP.路径ID, mPP.版本号)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strIDs As String '路径执行ID
        
    If mlngFun = 0 Or mlngFun = 1 Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, mcol("执行结果"))) = "" And CStr(.Cell(flexcpData, i, mcol("执行结果"))) <> "" Then
                    MsgBox "请选择一个执行结果。", vbInformation, gstrSysName
                    .SetFocus
                    .Select i, mcol("执行结果")
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If mlngFun = 0 Or mlngFun = 2 Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    strIDs = strIDs & "," & .TextMatrix(i, mcol("ID"))
                End If
            Next
            If strIDs = "" Then
                MsgBox "请至少选择一个路径项目（打勾）。", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    
    If mintMode = 1 Then
        If SaveItemOut = False Then Exit Sub
    Else
        If SaveItem = False Then Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Function SaveItem() As Boolean
'功能:保存路径项目执行的数据
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSQL As String, strTotal As String, strThis As String, i As Long
    Dim strDate As String
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsItem
        If mlngFun = 2 Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    strSQL = "Zl_病人路径执行_Delete(" & .TextMatrix(i, mcol("ID")) & ")"
                    colSQL.Add strSQL, "C" & colSQL.count + 1
                End If
            Next
        Else
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    strThis = .TextMatrix(i, mcol("ID")) & "|" & .TextMatrix(i, mcol("执行结果")) & "|" & _
                        IIf(Trim(.TextMatrix(i, mcol("执行说明"))) = "", " ", Trim(.TextMatrix(i, mcol("执行说明")))) & "||"
                        
                    If LenB(strTotal & strThis) > 4000 Then
                        strSQL = "Zl_病人路径执行_Update('" & UserInfo.姓名 & "'," & strDate & ",'" & strTotal & "')"
                        colSQL.Add strSQL, "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                End If
            Next
            If strTotal <> "" Then
                strSQL = "Zl_病人路径执行_Update('" & UserInfo.姓名 & "'," & strDate & ",'" & strTotal & "')"
                colSQL.Add strSQL, "C" & colSQL.count + 1
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False

    SaveItem = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveItemOut() As Boolean
'功能:保存路径项目执行的数据
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, strTotal As String, strThis As String, i As Long
    Dim strDate As String
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsItem
        If mlngFun = 2 Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    strSql = "Zl_病人门诊路径执行_Delete(" & .TextMatrix(i, mcol("ID")) & ")"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                End If
            Next
        Else
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("选择")) = 1 Then
                    strThis = .TextMatrix(i, mcol("ID")) & "|" & .TextMatrix(i, mcol("执行结果")) & "|" & _
                        IIf(Trim(.TextMatrix(i, mcol("执行说明"))) = "", " ", Trim(.TextMatrix(i, mcol("执行说明")))) & "||"
                        
                    If LenB(strTotal & strThis) > 4000 Then
                        strSql = "Zl_病人门诊路径执行_Update('" & UserInfo.姓名 & "'," & strDate & ",'" & strTotal & "')"
                        colSQL.Add strSql, "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                End If
            Next
            If strTotal <> "" Then
                strSql = "Zl_病人门诊路径执行_Update('" & UserInfo.姓名 & "'," & strDate & ",'" & strTotal & "')"
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False

    SaveItemOut = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mlngFun <> 2 Then
        If vsItem.Visible And vsItem.Rows > vsItem.FixedRows Then
            vsItem.SetFocus: vsItem.Col = mcol("执行结果")
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, lngW As Long
            
    Call InitItem
    Call LoadItem
    
    vsItem.Top = picInfo.Top + picInfo.Height
    If mlngFun = 1 Then
        '只有一行'没有"选择"列
        For i = 0 To vsItem.Cols - 1
            lngW = lngW + vsItem.ColWidth(i)
        Next
        Me.Width = lngW + 500
        
        vsItem.Width = Me.ScaleWidth - 100
        vsItem.Height = 2000
        Me.Height = picInfo.Height + vsItem.Height + picBottom.Height
        
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Else
        vsItem.Height = Me.ScaleHeight - picInfo.Height - picBottom.Height
        If mlngFun = 2 Then
            Me.Caption = "批量取消执行"
            lblNote.Caption = "请选择要取消执行的路径项目。"
        End If
    End If
    lblTip.Caption = "当前日期:" & mPP.当前日期 & "(第" & mPP.当前天数 & "天)"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
End Sub

Private Sub LoadItem()
'功能：加载待执行的路径项目
    Dim i As Long, j As Long
    Dim str执行结果 As String, str缺省结果 As String, str新执行结果 As String
    Dim arrtmp As Variant
    
    With vsItem
        .Redraw = flexRDNone
        .Rows = .FixedRows + mrsItem.RecordCount
        .MergeCol(0) = True
        For i = 1 To mrsItem.RecordCount
            .TextMatrix(i, mcol("分类")) = mrsItem!分类
            Call .Select(i, mcol("分类"))
            .CellAlignment = flexAlignCenterCenter
            
            .TextMatrix(i, mcol("项目内容")) = mrsItem!项目内容
            
            .Cell(flexcpChecked, i, mcol("选择")) = 1
                        
            If mlngFun = 2 Then
                .TextMatrix(i, mcol("执行结果")) = "" & mrsItem!执行结果
                .TextMatrix(i, mcol("执行说明")) = "" & mrsItem!执行说明
            Else
                If Not IsNull(mrsItem!执行结果) Then
                    str执行结果 = CStr(Split("" & mrsItem!执行结果, vbTab)(0))
                    str缺省结果 = Split("" & mrsItem!执行结果, vbTab)(1)
                End If
                
                If mrsItem!医嘱项目 = 1 And mrsItem!有医嘱 = 0 Then
                '选择生成时未生成的项目，执行结果不能为已经执行
                '缺省结果中没有执行性质
                    If InStr(str执行结果, "|") > 0 Then
                        j = InStr(str执行结果, str缺省结果)
                        If j > 0 Then
                            j = j + Len(str缺省结果) + 1
                            If Val(Mid(str执行结果, j, 1)) = E执行结果.E已经执行 Then str缺省结果 = ""
                        End If
                    End If
                    
                    '可选列表中不显示已经执行的结果记录
                    If InStr(str执行结果, "|") > 0 Then
                        str新执行结果 = ""
                        arrtmp = Split(str执行结果, ",")
                        For j = 0 To UBound(arrtmp)
                            If Val(Split(arrtmp(j), "|")(1)) <> E执行结果.E已经执行 Then
                                str新执行结果 = str新执行结果 & "," & arrtmp(j)
                            End If
                        Next
                        str执行结果 = Mid(str新执行结果, 2)
                    End If
                End If
                .TextMatrix(i, mcol("执行结果")) = str缺省结果
                .Cell(flexcpData, i, mcol("执行结果")) = str执行结果
            End If
            .TextMatrix(i, mcol("ID")) = Val(mrsItem!ID)
            
            If Not IsNull(mrsItem!图标ID) Then
                Call .Select(i, mcol("项目内容"))
                .CellPictureAlignment = flexPicAlignRightCenter 'flexPicAlignLeftCenter
                .CellPicture = GetPathIcon(mrsItem!图标ID)
            End If
            
            mrsItem.MoveNext
        Next
        
        .Redraw = True
        .AutoSize .FixedCols, .Cols - 1, , 45 '在要Draw之后才生效
    End With
End Sub

Private Sub InitItem()
'功能: 初始化路径项目表头
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    If mlngFun = 1 Then
        strcol = "分类,1200,1;项目内容,2600,1;选择;执行结果,900,4;执行说明,2600,1;ID"
    Else
        strcol = "分类,1200,1;项目内容,3100,1;选择,500,4;执行结果,900,4;执行说明,3000,1;ID"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        
        For i = 0 To UBound(arrHead)
            mcol.Add i, Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .ColDataType(mcol("选择")) = flexDTBoolean
        .Redraw = True
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible And mlngFun <> 2 Then
        If NewCol = mcol("执行结果") Then
            Dim strTmp As String, arrtmp As Variant, i As Long, lngP As Long, blnDo As Boolean
            With vsItem
                blnDo = False
                strTmp = .Cell(flexcpData, NewRow, mcol("执行结果"))    '例：病历已书写|1,病历未书写|2
                arrtmp = Split(strTmp, ",")
                For i = 0 To UBound(arrtmp)
                    lngP = InStr(arrtmp(i), "|")
                    If lngP > 0 Then
                        blnDo = True
                        arrtmp(i) = Mid(arrtmp(i), 1, lngP - 1)
                    End If
                Next
                If blnDo Then
                    strTmp = Join(arrtmp, "|")
                Else
                    strTmp = Replace(strTmp, ",", "|")
                End If
                .ColComboList(NewCol) = strTmp
            End With
        End If
    End If
End Sub

Private Sub vsItem_DblClick()
'功能：批量选择
    With vsItem
        If .MouseRow = .FixedRows - 1 And .Col = mcol("选择") Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, mcol("选择")) = IIf(.Cell(flexcpChecked, i, mcol("选择")) = 1, 2, 1)
            Next
        End If
    End With
End Sub

Private Sub vsItem_GotFocus()
    vsItem.ForeColorSel = vbWhite
    vsItem.BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell
    End If
End Sub

Private Sub vsItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsItem.MouseCol = mcol("选择") Then
        vsItem.ToolTipText = "双击列名全部选择或取消"
    Else
        vsItem.ToolTipText = ""
    End If
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mcol("项目内容") Then Cancel = True
    If mlngFun = 2 Then
        If Col <> mcol("选择") Then Cancel = True
    End If
End Sub

Private Sub ResultEnterNextCell()
    With vsItem
        If .Col < mcol("执行说明") Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = IIf(.ColHidden(mcol("选择")), mcol("执行结果"), mcol("选择"))
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

