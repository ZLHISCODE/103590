VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmKssReasonSelect 
   BorderStyle     =   0  'None
   Caption         =   "抗菌用药理由"
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3020
      Left            =   0
      ScaleHeight     =   2985
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   0
         TabIndex        =   5
         Top             =   2520
         Width           =   6255
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3960
         TabIndex        =   3
         Top             =   2595
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   135
         TabIndex        =   2
         Top             =   2595
         Width           =   1100
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5080
         TabIndex        =   1
         Top             =   2595
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vsgMain 
         Height          =   2535
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   6255
         _cx             =   11033
         _cy             =   4471
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssReasonSelect.frx":0000
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
End
Attribute VB_Name = "frmKssReasonSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mstrName As String   '返回的用药理由名称
Private mstrFind As String
Private mlngleft As Long
Private mlngTop As Long
Private mintType As Integer
Private Enum COL医嘱常用原因
    col编码 = 0
    col名称 = 1
    col简码 = 2
End Enum

Public Function ShowMe(frmParent As Object, ByVal strFind As String, ByRef blnCancle As Boolean, ByVal lngLeft As Long, ByVal lngTop As Long, ByVal intType As Integer) As String
'返回：用药理由名称
'参数：strFind -为空则查找所有，否则根据strFind查找简码，编码，名称
'      intType 1-抗菌用药理由，2-常用嘱托，3-超量说明
    mstrFind = strFind
    mlngleft = lngLeft
    mlngTop = lngTop
    mintType = intType
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    blnCancle = Not mblnOK
    If mblnOK Then
        ShowMe = mstrName
    Else
        ShowMe = ""
    End If
End Function

Private Sub cmdDelete_Click()
    Dim strSQL As String
    
    If vsgMain.Row < 1 Or vsgMain.Row = vsgMain.Rows - 1 Then Exit Sub
    
    If mintType = 1 Or mintType = 3 Then
        strSQL = "zl_医嘱常用原因_Update(1,'" & vsgMain.TextMatrix(vsgMain.Row, col编码) & "')"
    ElseIf mintType = 4 Then
        strSQL = "zl_常用就诊摘要_Update(1,'" & vsgMain.TextMatrix(vsgMain.Row, col编码) & "')"
    Else
        strSQL = "zl_常用嘱托_Insert(Null,Null,'" & UserInfo.姓名 & "','" & vsgMain.TextMatrix(vsgMain.Row, col编码) & "')"
    End If
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    vsgMain.RemoveItem vsgMain.Row
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Call vsgMain_DblClick
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
     vsgMain.SetFocus
End Sub

Private Sub Form_Load()
    Dim strTmp As String, strSQL As String
    Dim rsTmp As Recordset, i As Long
    
    mstrName = ""
    mblnOK = False
    If mstrFind <> "" Then
        If IsNumeric(mstrFind) Then
            strTmp = " Where (编码=LPAD([1]," & IIF(mintType = 1, "4", "5") & ",'0') Or 名称 Like [2]) "
        Else
            strTmp = " Where (简码 Like [2] Or 名称 Like [2]) "
        End If
    End If
    If mintType = 1 Then
        strSQL = "Select 编码,名称,简码 From 医嘱常用原因" & strTmp & IIF(strTmp = "", " Where ", " And ") & " nvl(性质,0)=0 order by to_number(编码)"
    ElseIf mintType = 3 Then
        strSQL = "Select 编码,名称,简码 From 医嘱常用原因" & strTmp & IIF(strTmp = "", " Where ", " And ") & " 人员=[3] And nvl(性质,0)=1 order by to_number(编码)"
    ElseIf mintType = 4 Then
        strSQL = "Select 编码,名称,简码 From 常用就诊摘要" & strTmp & IIF(strTmp = "", " Where ", " And ") & " (人员id=[4] or 人员id is null) order by to_number(编码)"
    Else
        strSQL = "Select 编码,名称,简码 From 常用嘱托" & strTmp & IIF(strTmp = "", " Where ", " And ") & " (人员=[3] or 人员 is null) order by to_number(编码)"
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrFind, "%" & UCase(mstrFind) & "%", UserInfo.姓名, UserInfo.ID)
    
    vsgMain.Rows = 1: vsgMain.AddItem ""
    Me.Left = mlngleft
    Me.Top = mlngTop
    If Not rsTmp.EOF Then
        If rsTmp.RecordCount = 1 Then
            '只有一个记录直接返回
            mblnOK = True
            mstrName = rsTmp!名称 & ""
            Unload Me
        Else
            With vsgMain
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, col编码) = Nvl(rsTmp!编码)
                    .TextMatrix(i, col名称) = Nvl(rsTmp!名称)
                    .TextMatrix(i, col简码) = Nvl(rsTmp!简码)
                    rsTmp.MoveNext
                    .AddItem ""
                Next
                vsgMain.Cell(flexcpBackColor, vsgMain.Rows - 1, col名称) = &HFFEADA
                vsgMain.Row = 1
            End With
        End If
    Else
        Unload Me
        mblnOK = True
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vsgMain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = vsgMain.Rows - 1 And NewCol = col名称 Then
        vsgMain.FocusRect = flexFocusHeavy
        vsgMain.Editable = flexEDKbdMouse
    Else
        vsgMain.FocusRect = flexFocusNone
        vsgMain.Editable = flexEDNone
    End If
End Sub

Private Sub vsgMain_DblClick()
    If vsgMain.Row < 1 Or vsgMain.Row = vsgMain.Rows - 1 Then Exit Sub
    mblnOK = True
    mstrName = vsgMain.TextMatrix(vsgMain.Row, col名称)
    Unload Me
End Sub

Private Sub vsgMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call vsgMain_DblClick
End Sub

Private Sub vsgMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strSQL As String, rsTmp As Recordset
    Dim strSpellCode As String
    Dim strTmp As String
    
    If Row = vsgMain.Rows - 1 And Col = col名称 Then
        If vsgMain.EditText = "" Then Exit Sub
        If mintType = 1 Or mintType = 3 Or mintType = 4 Then
            If mintType = 1 Then
                strTmp = "用药理由"
            ElseIf mintType = 3 Then
                strTmp = "超量说明"
            ElseIf mintType = 4 Then
                strTmp = "就诊摘要"
            End If
            
            If zlCommFun.ActualLen(vsgMain.EditText) > 1000 Then
                MsgBox "输入内容不过超过 500 个汉字或 1000 个字符。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            
            If mintType = 4 Then
                strSQL = "Select 1 From 常用就诊摘要 Where 名称=[1] And 人员ID=[3]"
            ElseIf mintType = 3 Then
                strSql = "Select 1 From 医嘱常用原因 Where 名称=[1] And 性质=1 And 人员=[2]"
            Else
                strSQL = "Select 1 From 医嘱常用原因 Where 名称=[1]"
            End If
            
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vsgMain.EditText, UserInfo.姓名, UserInfo.ID)
            '如果已经有了，提示用户是否继续。
            If rsTmp.RecordCount > 0 Then
                MsgBox "已经存在相同的" & strTmp & "。", vbInformation, Me.Caption
                Cancel = True: Exit Sub
            End If
            
            If mintType = 4 Then
                strSQL = "Select LPad(To_Char(Max(To_Number(编码)) + 1), 4, '0') as 编码 From 常用就诊摘要"
            Else
                strSQL = "Select LPad(To_Char(Max(To_Number(编码)) + 1), 4, '0') as 编码 From 医嘱常用原因"
            End If
            
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTmp.RecordCount < 1 Then Exit Sub
            strSpellCode = Mid(zlCommFun.SpellCode(vsgMain.EditText), 1, 10)
            
            If mintType = 4 Then
                strSQL = "zl_常用就诊摘要_Update(0,'" & rsTmp!编码 & "" & "','" & vsgMain.EditText & "','" & strSpellCode & "'," & UserInfo.ID & ")"
            Else
                strSQL = "zl_医嘱常用原因_Update(0,'" & rsTmp!编码 & "" & "','" & vsgMain.EditText & "','" & strSpellCode & "'" & IIF(mintType = 3, ",1,'" & UserInfo.姓名 & "'", "") & ")"
            End If
            
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
        Else
            On Error GoTo errH
            If zlCommFun.ActualLen(vsgMain.EditText) > 100 Then
                MsgBox "输入内容不过超过 50 个汉字或 100 个字符。", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            strSQL = "Select 1 From 常用嘱托 Where 名称=[1] And (人员=[2] Or 人员 is null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(vsgMain.EditText, "'", "''"), UserInfo.姓名)
            If rsTmp.RecordCount > 0 Then
                MsgBox "该嘱托内容已经在常用嘱托中。", vbInformation, Me.Caption
                Cancel = True: Exit Sub
                Exit Sub
            End If
            
            
            strSpellCode = zlCommFun.zlGetSymbol(vsgMain.EditText, CByte(0))
            strSQL = "zl_常用嘱托_Insert('" & Replace(vsgMain.EditText, "'", "''") & "','" & strSpellCode & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            '补上编码
            strSQL = "Select 编码 From 常用嘱托 Where 名称=[1] And (人员=[2] Or 人员 is null)"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(vsgMain.EditText, "'", "''"), UserInfo.姓名)
        End If
        vsgMain.Editable = flexEDNone
        If rsTmp.RecordCount > 0 Then
            vsgMain.TextMatrix(Row, col编码) = rsTmp!编码
            vsgMain.TextMatrix(Row, col简码) = strSpellCode
        End If
        vsgMain.Cell(flexcpBackColor, Row, col名称) = &H80000005
        vsgMain.AddItem ""
        vsgMain.Cell(flexcpBackColor, vsgMain.Rows - 1, col名称) = &HFFEADA
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
