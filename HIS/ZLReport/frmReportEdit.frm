VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReportEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   Icon            =   "frmReportEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   8535
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraGroups 
      Caption         =   "所属报表组"
      Height          =   3555
      Left            =   4200
      TabIndex        =   9
      Top             =   30
      Width           =   4215
      Begin VSFlex8Ctl.VSFlexGrid vsfGroups 
         Height          =   3135
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3975
         _cx             =   1989548899
         _cy             =   1989547418
         Appearance      =   1
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7200
      TabIndex        =   12
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   11
      Top             =   3720
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   4050
      Begin VB.ComboBox cboClass 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3000
      End
      Begin VB.TextBox txt说明 
         BackColor       =   &H00FFFFFF&
         Height          =   1920
         Left            =   855
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1485
         Width           =   3000
      End
      Begin VB.TextBox txt名称 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   855
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1065
         Width           =   3000
      End
      Begin VB.TextBox txt编号 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   20
         TabIndex        =   4
         Top             =   645
         Width           =   1500
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分类"
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "说明"
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   1530
         Width           =   360
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   705
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmReportEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_GROUPS_COLS As String = _
    "选择,,3,450,B|编号,,3,1500|名称,,3,1500|ID,,0,n"
    
Private mbytMode As Byte            '0-报表；1-报表组；2-报表类；3-报表组子报表
Private mlngSys As Long
Private mlngReportID As Long
Private mlngGroupID As Long
Private mlngClassID As Long
Private mstr名称 As String
Private mstrOld名称 As String
Private mstr编码 As String
Private mstr说明 As String
Private mstrOld说明 As String
Private mblnOK As Boolean
Private mlngModule As Long

Private WithEvents mvsfGroups As clsVSFlexGridEx
Attribute mvsfGroups.VB_VarHelpID = -1

Public Function ShowMe(ByVal frmParent As Object, ByVal lngSys As Long _
    , ByVal bytMode As Byte, ByVal lngModule As Long _
    , Optional ByRef lngGroupID As Long, Optional ByRef lngReportID As Long _
    , Optional ByRef str名称 As String, Optional ByRef str编码 As String _
    , Optional ByRef str说明 As String) As Boolean
    
    mlngSys = lngSys
    mbytMode = bytMode
    mlngModule = lngModule
    mlngReportID = lngReportID
    mlngGroupID = lngGroupID
    mstr名称 = str名称: mstrOld名称 = str名称
    mstr编码 = str编码
    mstr说明 = str说明: mstrOld说明 = str说明
    
    If bytMode = 2 Then
        mlngClassID = lngReportID
    ElseIf bytMode = 1 Then
        mlngClassID = GetClassID(lngGroupID, True)
    Else
        mlngClassID = GetClassID(lngReportID)
    End If
    
    Set mvsfGroups = New clsVSFlexGridEx
    
    Me.Show vbModal, frmParent
    str名称 = mstr名称
    str编码 = mstr编码
    str说明 = mstr说明
    ShowMe = mblnOK
End Function

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    If mbytMode = Val("2-报表类") Then
        If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsCheck As New ADODB.Recordset
    Dim strSQL As String, strOldCode As String, strOldName As String, strOld说明 As String

    Dim intOrder As Integer
    Dim arrSQL() As Variant
    Dim i As Long, lngClassID As Long, lngTemp As Long, lngProgID As Long
    Dim blnTrans As Boolean
    
    If UCase(Me.ActiveControl.name) <> UCase("cmdOK") Then
        If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    End If
    
    arrSQL = Array()
    If Not CheckFormInput(Me) Then Exit Sub
    
    If mbytMode = Val("2-报表类") Then
        If Trim(txt名称.Text) = "所有" Then
            MsgBox "“所有”为分类保留关键字，请修改！", vbInformation, App.Title
            txt名称.SetFocus: Exit Sub
        End If
    Else
        If Trim(txt编号.Text) = "" Then
            MsgBox "请输入报表" & IIF(mbytMode = 1, "组", "") & "的编号！", vbInformation, App.Title
            txt编号.SetFocus: Exit Sub
        End If
    End If
    
    If Trim(txt名称.Text) = "" Then
        Select Case mbytMode
        Case 2
            MsgBox "请输入“报表类”的名称！", vbInformation, App.Title
        Case 1
            MsgBox "请输入“报表组”的名称！", vbInformation, App.Title
        Case Else
            MsgBox "请输入“报表”的名称！", vbInformation, App.Title
        End Select
        txt名称.SetFocus
        Exit Sub
    Else
        txt名称.Text = ConvertSBC(txt名称.Text)
    End If
    
    If Not CheckLen(txt编号, 20, "编号") Then Exit Sub
    If Not CheckLen(txt名称, 30, "名称") Then Exit Sub
    If Not CheckLen(txt说明, 255, "说明") Then Exit Sub
    
    On Error GoTo hErr
    
    '检查
    If mbytMode = Val("2-报表类") Then
        '报表类
        If mlngClassID = 0 Then
            strSQL = "Select 名称 From zlRPTClasses Where Upper(名称) = [1]"
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, UCase(txt名称.Text))
            If rsCheck.RecordCount > 0 Then
                MsgBox "新增分类“" & txt名称.Text & "”重复！", vbInformation, App.Title
                txt名称.SetFocus
                Exit Sub
            End If
            rsCheck.Close
        End If
    Else
        '编号不能重复(报表及报表组)
        If CheckExist("zlReports", "编号", txt编号.Text, mlngReportID) Then
            MsgBox "该编号已经被使用,请重新输入！", vbInformation, App.Title
            txt编号.SetFocus: Exit Sub
        End If
        If CheckExist("zlRPTGroups", "编号", txt编号.Text, mlngGroupID) Then
            MsgBox "该编号已经被使用,请重新输入！", vbInformation, App.Title
            txt编号.SetFocus: Exit Sub
        End If
        
        If mlngGroupID <> 0 And mbytMode <> Val("1-报表组") Then
            strSQL = _
                "Select 1 From zlRPTSubs A,zlReports B " & vbCrLf & _
                "Where B.名称=[1] And A.报表ID=B.ID And A.组ID=[2]" & _
                IIF(mlngReportID = 0, "", " And 报表ID<>[3]")
            Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, txt名称.Text, mlngGroupID, mlngReportID)
            If Not rsCheck.EOF Then
                MsgBox "该报表组中已经包含相同名称的报表！", vbInformation, App.Title
                txt名称.SetFocus: Exit Sub
            End If
        End If
    End If
    
    strOldCode = mstr编码: strOldName = mstrOld名称: strOld说明 = mstrOld说明
    mstr名称 = txt名称.Text: mstr编码 = txt编号.Text: mstr说明 = txt说明.Text
    If cboClass.ListIndex >= 0 Then
        lngClassID = cboClass.ItemData(cboClass.ListIndex)
    End If
    
    '保存
    Select Case mbytMode
    Case Val("2-报表类")
        If mlngClassID = 0 Then
            '新增
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTClasses(ID,上级ID,名称,说明) " & vbCrLf & _
                "Values " & vbCrLf & _
                "(zlRPTClasses_ID.nextval " & vbCrLf & _
                "," & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                ",'" & mstr名称 & "'" & vbCrLf & _
                "," & IIF(mstr说明 = "", "Null", "'" & mstr说明 & "'") & _
                ")"
        ElseIf Not (strOldName = mstr名称 And strOld说明 = mstr说明 And lngClassID = mlngGroupID) Then
            '修改
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Update zlRPTClasses " & vbCrLf & _
                "Set 上级ID = " & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                "   ,名称 = '" & mstr名称 & "'" & vbCrLf & _
                "   ,说明 = " & IIF(mstr说明 = "", "Null", "'" & mstr说明 & "'") & vbCrLf & _
                "Where ID = " & mlngClassID
        End If
    Case Val("1-报表组")
        If mlngGroupID = 0 Then
            mlngGroupID = GetNextID("zlRPTGroups")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTGroups(ID,分类ID,编号,名称,说明) " & vbCrLf & _
                "Values(" & mlngGroupID & _
                IIF(lngClassID = 0, ",Null", "," & lngClassID) & _
                ",'" & mstr编码 & "','" & mstr名称 & "','" & mstr说明 & "')"
        ElseIf Not (strOldName = mstr名称 And strOld说明 = mstr说明 And lngClassID = mlngClassID) Then
            '说明与名称发生变化
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Update zlRPTGroups " & vbCrLf & _
                "Set 编号='" & mstr编码 & "',名称='" & mstr名称 & "',说明='" & mstr说明 & "' " & vbCrLf & _
                "   ,分类ID=" & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                "Where ID=" & mlngGroupID
            '发布到导航台菜单的报表标题
            If mlngModule <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlPrograms " & vbCrLf & _
                    "Set 标题='" & mstr名称 & "',说明='" & mstr说明 & "' " & vbCrLf & _
                    "Where 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlMenus " & vbCrLf & _
                    "Set 标题='" & mstr名称 & "',短标题='" & mstr名称 & "',说明='" & mstr说明 & "' " & vbCrLf & _
                    "Where ID=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
            End If
        End If
    Case Else
        '默认-报表
        If mlngReportID = 0 Then
            '新增
            If mlngSys <> 0 Then mlngSys = 0
            mlngReportID = GetNextID("zlReports")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlReports(ID,分类ID,编号,名称,说明,系统,修改时间,密码) " & vbCrLf & _
                "Values(" & mlngReportID & _
                "," & IIF(lngClassID = 0, "null", lngClassID) & _
                ",'" & mstr编码 & "','" & mstr名称 & "','" & mstr说明 & "'," & vbCrLf & _
                IIF(mlngSys = 0, "NULL", mlngSys) & ",Sysdate," & AdjustStr(GetPass(mstr编码, mstr名称)) & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = _
                "Insert Into zlRPTFmts(报表ID,序号,说明,W,H,纸张,纸向,动态纸张,图样) " & vbCrLf & _
                "Values(" & mlngReportID & ",1,'" & mstr名称 & "1'," & INIT_WIDTH & "," & INIT_HEIGHT & ",9,1,0,0)"

            '所属报表组
            If fraGroups.Visible Then
                For i = 1 To vsfGroups.Rows - 1
                    If Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("选择"))) <> 0 Then
                        lngTemp = Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ID")))
                        strSQL = "Select Count(1) Rec From zlRPTSubs Where 组ID=[1]"
                        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, lngTemp)
                        If Not rsCheck.EOF Then
                            intOrder = Nvl(rsCheck!Rec, 0) + 1
                        Else
                            intOrder = 1
                        End If
                        rsCheck.Close
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlRPTSubs(组ID,报表ID,序号) " & vbCrLf & _
                            "Values(" & lngTemp & "," & mlngReportID & "," & intOrder & ")"
                        If mlngModule <> 0 Then
                            '插入权限记录
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Insert Into zlProgFuncs(系统,序号,功能,说明) " & vbCrLf & _
                                "Values(" & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                "," & mlngModule & _
                                ",'" & mstr名称 & "'" & _
                                ",'" & mstr说明 & "')"
                        End If
                    End If
                Next
            End If
        ElseIf Not (strOldCode = mstr编码 And strOldName = mstr名称 And strOld说明 = mstr说明 _
                        And lngClassID = mlngClassID And Val(vsfGroups.Tag) = 0) Then
            '修改
            If Not (strOldCode = mstr编码 And strOldName = mstr名称 And strOld说明 = mstr说明 _
                        And lngClassID = mlngClassID) Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = _
                    "Update zlReports " & vbCrLf & _
                    "Set 编号='" & mstr编码 & "',名称='" & mstr名称 & "',说明='" & mstr说明 & "'" & vbCrLf & _
                    "   ,密码=" & AdjustStr(GetPass(mstr编码, mstr名称)) & vbCrLf & _
                    "   ,分类ID=" & IIF(lngClassID = 0, "Null", lngClassID) & vbCrLf & _
                    "Where ID=" & mlngReportID
                If mlngModule <> 0 Then '发布到导航台菜单的报表标题
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "Update zlPrograms " & vbCrLf & _
                        "Set 标题='" & mstr名称 & "',说明='" & mstr说明 & "' " & vbCrLf & _
                        "Where Upper(部件)='ZL9REPORT' And 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = _
                        "Update zlMenus " & vbCrLf & _
                        "Set 标题='" & mstr名称 & "',短标题='" & mstr名称 & "',说明='" & mstr说明 & "' " & vbCrLf & _
                        "Where 模块=" & mlngModule & " And Nvl(系统,0)=" & mlngSys & vbCrLf & _
                        "    And Exists(Select 标题 From zlPrograms " & vbCrLf & _
                        "               Where Upper(部件)='ZL9REPORT' And 序号=" & mlngModule & " And Nvl(系统,0)=" & mlngSys & ")"
                End If
                
                '发布到导航台的报表组子表的功能名
                strSQL = _
                    "Select Distinct Nvl(B.系统, 0) 系统, B.程序id 序号, a.组Id " & vbCrLf & _
                    "From Zlrptsubs a, Zlrptgroups b, Zlprograms c" & vbCrLf & _
                    "Where A.组id = B.Id And A.报表id = [1]  And Nvl(B.系统, 0) = Nvl(C.系统, 0) " & vbCrLf & _
                    "    And B.程序id = C.序号 And Upper(C.部件) = 'ZL9REPORT'"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
                Do While Not rsCheck.EOF
                    If strOldName <> mstr名称 Then  '报表名称发生变化
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新子报表名称
                        arrSQL(UBound(arrSQL)) = _
                            "Update zlRPTSubs " & vbNewLine & _
                            "Set 功能 = '" & mstr名称 & "' " & vbNewLine & _
                            "Where 组Id = " & Nvl(rsCheck!组ID) & _
                            "    And 报表Id = " & mlngReportID & " And 功能 = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into Zlprogfuncs" & vbNewLine & _
                            "  (系统, 序号, 功能, 排列, 说明, 缺省值)" & vbNewLine & _
                            "  Select A.系统, A.序号, '" & mstr名称 & "', A.排列, '" & mstr说明 & "', A.缺省值" & vbNewLine & _
                            "  From Zlprogfuncs a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "      And A.功能 = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能授权信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlrolegrant" & vbNewLine & _
                            "  (系统,序号,角色,功能)" & vbNewLine & _
                            "  Select A.系统,A.序号,A.角色, '" & mstr名称 & "' " & vbNewLine & _
                            "  From zlrolegrant a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "     And A.功能 = '" & strOldName & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能对象权限信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlprogprivs" & vbNewLine & _
                            "  (系统,序号,功能,对象,所有者,权限)" & vbNewLine & _
                            "  Select A.系统,A.序号,'" & mstr名称 & "',A.对象,A.所有者,A.权限" & vbNewLine & _
                            "  From zlprogprivs a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "      And A.功能 = '" & strOldName & "'" & _
                            "      And Not Exists(Select 1 From zlProgPrivs " & vbCr & _
                            "                     Where Nvl(系统,0)=Nvl(a.系统,0) And 序号=a.序号 And 功能='基本' " & vbCr & _
                            "                         And 对象=a.对象 And 所有者=a.所有者 And 权限=a.权限)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '删除原始功能名，由于存在级联删除关系
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From Zlprogfuncs a " & vbNewLine & _
                            "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "    And A.功能 = '" & strOldName & "'"
                        '系统、序号、功能 任意一个存在Null，级联删除将失效
                        If Nvl(rsCheck!系统, 0) = 0 Or Nvl(rsCheck!序号, 0) = 0 Then
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete From zlProgPrivs A " & vbNewLine & _
                                "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                                "    And A.功能 = '" & strOldName & "'"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete From zlRoleGrant A " & vbNewLine & _
                                "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                                "    And A.功能 = '" & strOldName & "'"
                        End If
                    Else '报表名称未发生变化,只许更新功能说明
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新功能说明
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlprogfuncs A" & vbNewLine & _
                            "  Set A.说明='" & mstr说明 & "'" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "      And A.功能 = '" & mstr名称 & "'"
                    End If
                    rsCheck.MoveNext
                Loop
                
                '发布到模块的报表功能名
                strSQL = _
                    "Select Nvl(B.系统, 0) 系统, B.程序id 序号, B.功能 " & vbNewLine & _
                    "From Zlrptputs b, Zlprograms c, Zlprogfuncs d " & vbNewLine & _
                    "Where B.报表id =[1] And Nvl(B.系统, 0) = Nvl(C.系统, 0) And B.程序id = C.序号 " & vbNewLine & _
                    "    And Upper(C.部件) <> 'ZL9REPORT' And Nvl(C.系统, 0) = Nvl(D.系统, 0) And C.序号 = D.序号 " & vbNewLine & _
                    "    And D.功能 = B.功能"
                Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
                Do While Not rsCheck.EOF
                    If strOldName <> mstr名称 And mlngSys = 0 Then   '非系统报表名称发生变化，则自动更新功能名称
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新Zlrptputs
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlrptputs Set 功能 = '" & mstr名称 & "' " & vbNewLine & _
                            "Where 报表id = " & mlngReportID & " And Nvl(系统, 0) = " & rsCheck!系统 & " And 程序id = " & rsCheck!序号
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into Zlprogfuncs" & vbNewLine & _
                            "  (系统, 序号, 功能, 排列, 说明, 缺省值)" & vbNewLine & _
                            "  Select A.系统, A.序号, '" & mstr名称 & "', A.排列, '" & mstr说明 & "', A.缺省值" & vbNewLine & _
                            "  From Zlprogfuncs a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "     And A.功能 = '" & rsCheck!功能 & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能授权信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlrolegrant" & vbNewLine & _
                            "  (系统,序号,角色,功能)" & vbNewLine & _
                            "  Select A.系统,A.序号,A.角色, '" & mstr名称 & "' " & vbNewLine & _
                            "  From zlrolegrant a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "      And A.功能 = '" & rsCheck!功能 & "'"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '复制一份原始功能对象权限信息
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlprogprivs" & vbNewLine & _
                            "  (系统,序号,功能,对象,所有者,权限)" & vbNewLine & _
                            "  Select A.系统,A.序号,'" & mstr名称 & "',A.对象,A.所有者,A.权限" & vbNewLine & _
                            "  From zlprogprivs a" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "      And A.功能 = '" & rsCheck!功能 & "'" & _
                            "      And Not Exists(Select 1 From zlProgPrivs " & vbCr & _
                            "                     Where 系统=a.系统 And 序号=a.序号 And 功能='基本' " & vbCr & _
                            "                         And 对象=a.对象 And 所有者=a.所有者 And 权限=a.权限)"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '删除原始功能名，由于存在级联删除关系
                        arrSQL(UBound(arrSQL)) = _
                            "Delete From Zlprogfuncs a " & vbNewLine & _
                            "Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "    And A.功能 = '" & rsCheck!功能 & "'"
                    Else '非系统报表说明变化或者固定报表变更，则只更新功能说明
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1) '更新功能说明
                        arrSQL(UBound(arrSQL)) = _
                            "Update Zlprogfuncs A" & vbNewLine & _
                            "  Set A.说明='" & mstr说明 & "'" & vbNewLine & _
                            "  Where Nvl(A.系统, 0) = " & rsCheck!系统 & " And A.序号 = " & rsCheck!序号 & vbNewLine & _
                            "     And A.功能 = '" & rsCheck!功能 & "'"
                    End If
                    rsCheck.MoveNext
                Loop
            End If
            
            '所属报表组
            If fraGroups.Visible And Val(vsfGroups.Tag) = Val("1-已操作vsfGroup") Then
                For i = 1 To vsfGroups.Rows - 1
                    '获取报表组ID
                    lngTemp = Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("ID")))
                    If Val(vsfGroups.TextMatrix(i, vsfGroups.ColIndex("选择"))) = 0 Then
                        '移出报表组
                        If mlngModule <> 0 Then
                            '已发布的自定义报表
                            '清除组报表的所有权限记录
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Delete zlProgFuncs " & _
                                "Where 系统 is null And 序号 = " & mlngModule
                        Else
                            '未发布的自定义报表
                            lngProgID = ReportGroupIssue(lngTemp)
                            If lngProgID <> 0 Then
                                '组报表有发布，清除子报表的权限记录
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = _
                                    "Delete zlProgFuncs " & _
                                    "Where 系统 is null And 序号 = " & lngProgID & _
                                    "    And 功能 = '" & mstr名称 & "'"
                            End If
                        End If
                        
                        '清除当前报表的子报表记录
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Delete zlRPTSubs " & vbCrLf & _
                            "Where 组ID = " & lngTemp & " And 报表ID =" & mlngReportID
                    Else
                        '移入报表组
                        
                        '获取子报表的序号
                        strSQL = "Select Count(1) Rec From zlRPTSubs Where 组ID=[1]"
                        Set rsCheck = OpenSQLRecord(strSQL, Me.Caption, lngTemp)
                        If Not rsCheck.EOF Then
                            intOrder = Nvl(rsCheck!Rec, 0) + 1
                        Else
                            intOrder = 1
                        End If
                        rsCheck.Close
                        
                        '插入子报表记录
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "Insert Into zlRPTSubs(组ID,报表ID,序号,功能) " & vbCrLf & _
                            "Select " & lngTemp & "," & mlngReportID & "," & intOrder & ",'" & mstr名称 & "' " & vbCrLf & _
                            "From Dual " & vbCr & _
                            "Where Not Exists(Select 1 From zlRPTSubs " & vbCr & _
                            "                 Where 组ID = " & lngTemp & " And 报表ID = " & mlngReportID & ")"
                        
                        If mlngModule <> 0 Then
                            '已发布的自定义报表
                            '插入权限记录
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = _
                                "Insert Into zlProgFuncs(系统,序号,功能,说明) " & vbCrLf & _
                                "Select " & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                "," & mlngModule & _
                                ",'" & mstr名称 & "'" & _
                                ",'" & mstr说明 & "'" & vbCrLf & _
                                "From Dual Where Not Exists(Select 1 From zlProgFuncs " & vbCrLf & _
                                "                           Where 系统 Is Null And 序号 = " & mlngModule & ")"
                        Else
                            '未发布的自定义报表
                            lngProgID = ReportGroupIssue(lngTemp)
                            If lngProgID <> 0 Then
                                '组报表有发布，子报表缺省组报表的程序ID（插入子报表的权限记录）
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = _
                                    "Insert Into zlProgFuncs(系统,序号,功能,说明) " & vbCrLf & _
                                    "Select " & IIF(mlngSys = 0, "NULL", mlngSys) & _
                                    "," & lngProgID & _
                                    ",'" & mstr名称 & "'" & _
                                    ",'" & mstr说明 & "'" & vbCrLf & _
                                    "From Dual " & vbCr & _
                                    "Where Not Exists(Select 1 From zlProgFuncs " & vbCrLf & _
                                    "                 Where 系统 Is Null And 序号 = " & lngProgID & _
                                    "                     And 功能 = '" & mstr名称 & "')"
                            End If
                        End If
                    End If
                Next
            End If
            
        End If
    End Select
    
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            gcnOracle.Execute arrSQL(i)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    '清除缓存
    Set grsReport = Nothing
    mblnOK = True
    Unload Me
    Exit Sub
    
hErr:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume

    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mlngClassID <> 0 Or mlngGroupID <> 0 Or mlngReportID <> 0 Then txt名称.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnViewGroups As Boolean
    
    mblnOK = False
    
    '获取报表分类信息
    If mbytMode = Val("2-报表类") Then
        Call mdlPublic.InitClass(cboClass, mlngClassID, mlngClassID)
    Else
        Call mdlPublic.InitClass(cboClass, mlngClassID)
    End If
    
    txt编号.Text = mstr编码
    txt名称.Text = mstr名称
    txt说明.Text = mstr说明
    
    Select Case mbytMode
    Case Val("2-报表类")
        If mlngClassID = 0 Then
            Caption = "新增报表分类"
        Else
            Caption = "修改报表分类"
        End If
        
        '上级分类
        For i = 0 To cboClass.ListCount - 1
            If mlngGroupID = Val(cboClass.ItemData(i)) Then
                cboClass.ListIndex = i
                Exit For
            End If
        Next
        
        cboClass.Enabled = mlngSys = 0
        
        lblClass.Left = 30
        lblClass.Caption = "上级分类"
        lblName.Top = lblCode.Top
        lblDesc.Top = lblName.Top + 420
        txt名称.Top = txt编号.Top
        txt说明.Top = txt编号.Top + 420
        txt说明.Height = txt说明.Height + 420
        
        txt编号.Text = ""
        txt编号.Visible = False
        lblCode.Visible = False
    Case Val("1-报表组")
        If mlngGroupID = 0 Then
            Caption = "新增报表组"
            txt编号.Text = GetNextNO(True)
        Else
            Caption = "修改报表组"
        End If
        cboClass.Enabled = mlngSys = 0
    Case Val("3-子报表")
        Caption = "修改子报表"
        cboClass.Enabled = False
        blnViewGroups = mlngSys = 0
    Case Else
        cboClass.Enabled = mlngSys = 0
        blnViewGroups = mlngSys = 0
        If mlngReportID = 0 Then
            Caption = "新增报表"
            txt编号.Text = GetNextNO(False)
        Else
            Caption = "修改报表"
        End If
    End Select
    If mlngSys > 0 Then txt编号.Enabled = False
    
    If blnViewGroups Then
        On Error GoTo hErr
        
        strSQL = _
            "Select Decode(Nvl(b.组id, 0), 0, 0, 1) 选择, a.编号, a.名称, a.Id " & vbNewLine & _
            "From zlRPTGroups A, zlRPTSubs B " & vbNewLine & _
            "Where a.Id = b.组id(+) And Nvl(a.系统, 0) = 0 And b.报表Id(+) = [1] "
        Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取共享系统的报表组信息", mlngReportID)
        
        With mvsfGroups
            .AppTemplate EM_Verify, vsfGroups, MSTR_GROUPS_COLS, "编号|名称"
            .Init True
            .Recordset = rsTemp
            .Repaint RT_Rows
        End With
        
        rsTemp.Close
    Else
        Me.Width = 4290
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120
    End If
    vsfGroups.Visible = blnViewGroups
    fraGroups.Visible = blnViewGroups
    
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mvsfGroups = Nothing
End Sub

Private Sub txt编号_GotFocus()
    SelAll txt编号
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    If InStr(1, "~!@#$%^&*()=+[]{}'"";,<>/?\", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt名称_GotFocus()
    SelAll txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf InStr(GSTR_SBC, Chr(KeyAscii)) > 0 Then
        KeyAscii = Asc(Mid(GSTR_DBC, InStr(GSTR_SBC, Chr(KeyAscii)), 1))
    End If
End Sub

Private Sub txt名称_Validate(Cancel As Boolean)
    If txt名称.Text <> "" Then
        txt名称.Text = ConvertSBC(txt名称.Text)
    End If
End Sub

Private Sub txt说明_GotFocus()
    SelAll txt说明
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If InStr(1, "~^&'"";,", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function GetClassID(ByVal lngID As Long, Optional ByVal blnGroup As Boolean = False) As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    GetClassID = 0
    
    On Error GoTo hErr
    
    If blnGroup Then
        strSQL = "Select 分类Id From zlRPTGroups Where Id = [1]"
    Else
        strSQL = "Select 分类Id From zlReports Where Id = [1]"
    End If
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "获取报表分类ID", lngID)
    If rsTemp.EOF = False Then
        GetClassID = Nvl(rsTemp!分类id, 0)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function

Private Sub vsfGroups_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfGroups.Tag = "1"     '1表示已操作，后续代码通过该标志更新报表组数据
End Sub

Private Function ReportGroupIssue(ByVal lngID As Long) As Long
'功能：判断报表组是否已发布
'参数：
'  lngID：报表组ID
'返回：等于0未发布；大于0已发布（即：程序ID）

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo hErr
    ReportGroupIssue = 0
    
    strSQL = _
        "Select 程序ID From zlRPTGroups Where ID = [1] And 发布时间 Is Not Null"
    Set rsTemp = mdlPublic.OpenSQLRecord(strSQL, "判断报表组是否已发布", lngID)
    If rsTemp.RecordCount = 1 Then
        ReportGroupIssue = mdlPublic.Nvl(rsTemp!程序id, 0)
    End If
    rsTemp.Close
    
    Exit Function
    
hErr:
    If ErrCenter = 1 Then Resume
End Function
