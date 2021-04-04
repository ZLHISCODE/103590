VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCheckIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自助签到系统"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   Icon            =   "frmCheckIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12180
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame frmLineBruchButtom 
      Height          =   120
      Left            =   0
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
      Width           =   12165
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   20
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer TimerAuto 
      Interval        =   60000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtCard 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   330
      Width           =   4095
   End
   Begin VB.Frame frmLineTop 
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11805
   End
   Begin VB.PictureBox picBrush 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6975
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   12015
      Begin VB.Frame frmLineBrushTop 
         Height          =   120
         Left            =   -120
         TabIndex        =   10
         Top             =   720
         Width           =   12045
      End
      Begin VB.CommandButton cmdCheckIn 
         Caption         =   "签到"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   9
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCheckInAll 
         Caption         =   "全部签到"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   8
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "下一张处方"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "上一张处方"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4155
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   11895
         _cx             =   20981
         _cy             =   7329
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCheckIn.frx":030A
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
         ExplorerBar     =   7
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   8640
         TabIndex        =   19
         Top             =   6360
         Width           =   3225
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "请点击这里签到-->"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   6405
         Width           =   3465
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   16
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   15
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblRecipeComment 
         AutoSize        =   -1  'True
         Caption         =   "总计2张处方未签到"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   8040
         TabIndex        =   14
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "发票号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "处方号："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   12
         Top             =   1140
         Width           =   1740
      End
   End
   Begin VB.PictureBox picUnBrush 
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7215
      ScaleWidth      =   11895
      TabIndex        =   3
      Top             =   1080
      Width           =   11895
      Begin VB.Label lblCommen 
         AutoSize        =   -1  'True
         Caption         =   "欢迎使用自助签到系统"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   3240
         TabIndex        =   4
         Top             =   6480
         Width           =   4050
      End
      Begin VB.Image imgHos 
         Height          =   6075
         Left            =   120
         Picture         =   "frmCheckIn.frx":03F2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   15015
      End
   End
   Begin VB.Label lblCard 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "请刷就诊卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1665
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRecipeList As ADODB.Recordset
Dim mstrCardPasswordRule As String          '卡密文规则
Private mlng药房id As Long
Private mBlnBegin As Boolean

Private Sub cmdBack_Click()
    picBrush.Visible = False
    picUnBrush.Visible = True
    Me.cmdBack.Visible = False
    Me.frmLineBruchButtom.Visible = False
    Me.txtCard.Text = ""
End Sub

Private Sub cmdCheckIn_Click()
    Dim strSQL As String
    
    On Error GoTo errRow
    strSQL = "Zl_未发药品记录_配药确认("
    strSQL = strSQL & "'" & mrsRecipeList!no & "',"
    strSQL = strSQL & mrsRecipeList!单据 & ","
    strSQL = strSQL & mlng药房id & ","
    strSQL = strSQL & "1,"
    strSQL = strSQL & "'auto')"
    Call zldatabase.ExecuteProcedure(strSQL, "cmdCheckIn_Click")
    
    If mrsRecipeList.RecordCount = 1 Then
        Me.picBrush.Visible = False
        Me.picUnBrush.Visible = True
        Me.lblCommen.Caption = "签到成功！"
        Me.txtCard.Text = ""
        Set mrsRecipeList = Nothing
    Else
        Me.lblMsg.Caption = "处方[" & mrsRecipeList!no & "]签到成功"

        Call mrsRecipeList.Delete(adAffectCurrent)
        Me.lblRecipeComment.Caption = " 总计" & mrsRecipeList.RecordCount & "张处方未签到"
        
        If mrsRecipeList.RecordCount = 1 Then
            Me.cmdNext.Visible = False
            Me.cmdPrevious.Visible = False
            Me.cmdCheckInAll.Visible = False
        End If
        
        If Me.cmdNext.Enabled Then
            Call cmdNext_Click
            Me.cmdPrevious.Enabled = False
        Else
            Call cmdPrevious_Click
            Me.cmdNext.Enabled = False
        End If
    End If
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCheckInAll_Click()
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errRow
    With mrsRecipeList
        .MoveFirst
        For i = 1 To .RecordCount
            strSQL = "Zl_未发药品记录_配药确认("
            strSQL = strSQL & "'" & !no & "',"
            strSQL = strSQL & !单据 & ","
            strSQL = strSQL & mlng药房id & ","
            strSQL = strSQL & "1,"
            strSQL = strSQL & "'auto')"
            Call zldatabase.ExecuteProcedure(strSQL, "cmdCheckInAll_Click")
            
            .MoveNext
        Next
    End With
    
    Me.picBrush.Visible = False
    Me.picUnBrush.Visible = True
    Me.lblCommen.Caption = "签到成功！"
    Me.txtCard.Text = ""
    Me.cmdBack.Visible = False
    Set mrsRecipeList = Nothing
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdNext_Click()
    Me.cmdPrevious.Enabled = True
    
    With mrsRecipeList
        If Not .EOF Then .MoveNext
        If Not .EOF Then
            Me.lblBill.Caption = "发票号：" & !号码
            Me.lblNo.Caption = "处方号：" & !no
            
            Call GetDetail(!no)
            
            .MoveNext
            If .EOF Then
                Me.cmdNext.Enabled = False
            End If
            .MovePrevious
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    Me.cmdNext.Enabled = True
     
    With mrsRecipeList
        If Not .BOF Then .MovePrevious
        If Not .BOF Then
            Me.lblBill.Caption = "发票号：" & !号码
            Me.lblNo.Caption = "处方号：" & !no
            Call GetDetail(!no)
            
            .MovePrevious
            If .BOF Then
                Me.cmdPrevious.Enabled = False
            End If
            .MoveNext
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mBlnBegin And KeyAscii = 13 Then
        Exit Sub
    Else
        If mBlnBegin Then
            txtCard.Text = ""
            mBlnBegin = False
        End If
        Call txtCard_KeyPress(KeyAscii)
    End If
    
End Sub

Private Sub Form_Load()
    Me.Caption = "自助签到系统" & "(" & gstrStockName & ")"
        
End Sub

Private Sub IniRecord()
    '初始化数据集
    Set mrsRecipeList = New ADODB.Recordset
    
    With mrsRecipeList
        If .State = 1 Then .Close
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "号码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "no", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "卡号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单据", adDouble, 2, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function GetList(ByVal strText As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo errRow
    
    IniRecord
    strSQL = "Select a.单据,a.no,b.姓名, b.性别, b.年龄, d.号码, e.卡号" & vbNewLine & _
            "From 未发药品记录 a, 病人信息 b, 票据打印内容 c, 票据使用明细 d, 病人医疗卡信息 e" & vbNewLine & _
            "Where a.No = c.No And a.病人id = b.病人id And a.病人id = e.病人id And c.Id = d.打印id And c.数据性质 = 1 And d.票种 = 1 And e.状态 = 0 And" & vbNewLine & _
            "      Nvl(排队状态, 0) = 0 And e.卡号 = [1] and a.库房id=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "GetList", strText, mlng药房id)
    
    With rsTemp
        Do While Not .EOF
            mrsRecipeList.AddNew
            mrsRecipeList!姓名 = !姓名
            mrsRecipeList!单据 = !单据
            mrsRecipeList!no = !no
            mrsRecipeList!性别 = !性别
            mrsRecipeList!年龄 = !年龄
            mrsRecipeList!号码 = !号码
            mrsRecipeList!卡号 = !卡号
            mrsRecipeList.Update
            
            .MoveNext
        Loop
    End With
    
    If rsTemp.RecordCount > 0 Then
        mrsRecipeList.MoveFirst
        With mrsRecipeList
            lblRecipeComment.Caption = "总计" & .RecordCount & "张处方未签到"
            If .RecordCount > 1 Then
                Me.cmdCheckInAll.Visible = True
                Me.cmdNext.Visible = True
                Me.cmdPrevious.Visible = True
                cmdPrevious.Enabled = False
            Else
                Me.cmdCheckInAll.Visible = False
                Me.cmdNext.Visible = False
                Me.cmdPrevious.Visible = False
            End If
        
            If Not .EOF Then
                lblName.Caption = !姓名
                lblSex.Caption = !性别
                lblAge.Caption = !年龄
                lblBill.Caption = "发票号：" & !号码
                lblNo.Caption = "处方号：" & !no
                
                GetList = !no
                Exit Function
            
            End If
        End With
    Else
        Me.lblCommen.Caption = "您没有需要签到的处方！"
        TimerAuto.Enabled = True
        GetList = ""
    End If
    Exit Function
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDetail(ByVal strNO As String)
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim dblMoney As Double
    Dim i As Integer
    
    On Error GoTo errRow
    strSQL = "Select a.No, a.序号, b.名称 药品名称, b.规格, b.计算单位 单位, a.零售价 单价, a.实际数量 数量, c.实收金额 金额" & vbNewLine & _
            "From 药品收发记录 a, 收费项目目录 b, 门诊费用记录 c" & vbNewLine & _
            "Where a.药品id = b.Id And a.费用id = c.Id And (Mod(a.记录状态, 3) = 0 Or a.记录状态 = 1) and a.no=[1] and a.库房id=[2]" & vbNewLine & _
            "Order By a.序号"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "GetDetail", strNO, mlng药房id)
    
    If Not rsTemp.EOF Then
        With Me.vsfList
            .Rows = rsTemp.RecordCount + 2
            
            For i = 1 To rsTemp.RecordCount
                .TextMatrix(i, .ColIndex("序号")) = rsTemp!序号
                .TextMatrix(i, .ColIndex("药品名称")) = rsTemp!药品名称
                .TextMatrix(i, .ColIndex("规格")) = rsTemp!规格
                .TextMatrix(i, .ColIndex("单位")) = rsTemp!单位
                .TextMatrix(i, .ColIndex("单价")) = Format(rsTemp!单价, "0.00#")
                .TextMatrix(i, .ColIndex("数量")) = rsTemp!数量
                .TextMatrix(i, .ColIndex("金额")) = Format(rsTemp!金额, "0.00#")
                dblMoney = Val(rsTemp!金额) + dblMoney
                rsTemp.MoveNext
            Next
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "金额小计：" & dblMoney
            .MergeCells = flexMergeFree
            .MergeRow(.Rows - 1) = True
            
        End With
    End If
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lblCard.Left = (Me.ScaleWidth - (Me.lblCard.Width + Me.txtCard.Width)) / 2
    Me.txtCard.Left = Me.lblCard.Left + Me.lblCard.Width + 150
    Me.cmdBack.Left = Me.ScaleWidth - Me.cmdBack.Width - 200
    
    frmLineTop.Width = Me.ScaleWidth
    picUnBrush.Move picUnBrush.Left, picUnBrush.Top, Me.ScaleWidth - 200, Me.ScaleHeight - frmLineTop.Top - 500
    imgHos.Move imgHos.Left, imgHos.Top, picUnBrush.Width - 100, picUnBrush.Height - 1200
    Me.lblCommen.Move (Me.ScaleWidth - Me.lblCommen.Width) / 2, (picUnBrush.Height - imgHos.Height) / 2 + imgHos.Height
    
    picBrush.Move (Me.ScaleWidth - Me.picBrush.Width) / 2, picBrush.Top, picBrush.Width, Me.ScaleHeight - frmLineTop.Top + 200
    lblRecipeComment.Left = picBrush.Width - lblRecipeComment.Width - 200
    frmLineBrushTop.Width = picBrush.Width - 50
    Me.cmdPrevious.Left = picBrush.Width - cmdPrevious.Width - 100
    Me.cmdNext.Left = Me.cmdPrevious.Left - Me.cmdNext.Width - 100
    
    vsfList.Move vsfList.Left, vsfList.Top, frmLineBrushTop.Width, picUnBrush.Height - frmLineBrushTop.Top - 2200
    
    lblTo.Top = (picBrush.Height - vsfList.Height - vsfList.Top - lblTo.Height) / 2 + vsfList.Height + vsfList.Top - 200
    Me.cmdCheckIn.Top = lblTo.Top - 100
    Me.cmdCheckInAll.Top = lblTo.Top - 100
    lblMsg.Left = picBrush.Width - lblMsg.Width - 200
    lblMsg.Top = lblTo.Top + 200
    
    frmLineBruchButtom.Move frmLineTop.Left, Me.picBrush.Top + Me.cmdCheckIn.Top - 300, frmLineTop.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRecipeList = Nothing
End Sub

Private Sub picUnBrush_KeyPress(KeyAscii As Integer)
    If InStr(":：;；?？''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    txtCard.Text = txtCard.Text & Chr(KeyAscii)
    If Len(txtCard.Text) = txtCard.MaxLength - 1 And KeyAscii <> 8 Then
        txtCard.Text = txtCard.Text & Chr(KeyAscii)
        txtCard.SelStart = Len(txtCard.Text)
        KeyAscii = 0
    End If
End Sub

Private Sub TimerAuto_Timer()
    Me.lblCommen.Caption = "欢迎使用自助签到系统"
    TimerAuto.Enabled = False
    Me.txtCard.Text = ""
    picUnBrush.Visible = True
    picBrush.Visible = False
    Me.cmdBack.Visible = False
End Sub

Private Sub txtCard_GotFocus()
    txtCard.SelStart = 0
    txtCard.SelLength = Len(txtCard.Text)
End Sub

Public Sub ShowMe(ByVal lng药房id As Long, ByVal strType As String)
    mlng药房id = lng药房id
    
    Me.lblCard.Caption = "请刷" & strType
    Me.Show
End Sub


Private Sub txtCard_KeyPress(KeyAscii As Integer)
     Dim strNO As String
    
    If KeyAscii = 13 And Not mBlnBegin Then
        If Me.txtCard.Text <> "" Then
            TimerAuto.Enabled = True
            strNO = GetList(txtCard.Text)
            Call SetPass(txtCard.Text)
            If strNO <> "" Then
                Call GetDetail(strNO)
                picBrush.Visible = True
                picUnBrush.Visible = False
                Me.cmdBack.Visible = True
                Me.frmLineBruchButtom.Visible = True
            Else
                picBrush.Visible = False
                picUnBrush.Visible = True
                Me.cmdBack.Visible = False
                Me.frmLineBruchButtom.Visible = False
            End If
        End If
        
        txtCard.SelStart = 0
        txtCard.SelLength = Len(txtCard.Text)
        mBlnBegin = True
    Else
        If InStr(":：;；?？''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub SetPass(ByVal strText As String)
    Dim count As Integer
    Dim intX As Integer
    Dim intY As Integer
    
    txtCard.Tag = strText
    intX = (Len(strText) - 3) / 2
    
    If intX < 2 Then
        txtCard.Text = Mid(txtCard.Text, 1, 1) & String(3, "*") & Mid(txtCard.Text, 5)
    Else
        txtCard.Text = Mid(txtCard.Text, 1, intX) & String(3, "*") & Mid(txtCard.Text, intX + 4)
    End If
   
End Sub
