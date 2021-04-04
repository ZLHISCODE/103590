VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLabMainSendSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发往仪器"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frmLabMainSendSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7455
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtSend 
      Height          =   300
      Left            =   2745
      TabIndex        =   12
      ToolTipText     =   "填写一次发磅的标本个数，不填表示全部发送"
      Top             =   4560
      Width           =   480
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "发送(&S)"
      Height          =   350
      Left            =   2628
      TabIndex        =   8
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   6165
      TabIndex        =   7
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   270
      TabIndex        =   6
      ToolTipText     =   "Ctrl+A"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   1449
      TabIndex        =   5
      ToolTipText     =   "Ctrl+R"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "自动编号"
      Height          =   350
      Left            =   4986
      TabIndex        =   4
      ToolTipText     =   "请填上盘号，杯号后，再用此功能．"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CheckBox Chk已发送 
      Caption         =   "显示已发送"
      Height          =   180
      Left            =   390
      TabIndex        =   3
      Top             =   4590
      Width           =   1260
   End
   Begin VB.TextBox txt盘号 
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   4545
      Width           =   765
   End
   Begin VB.TextBox txt杯号 
      Height          =   300
      Left            =   6195
      TabIndex        =   1
      Top             =   4560
      Width           =   765
   End
   Begin VB.CommandButton cmdDele 
      Caption         =   "清除编号"
      Height          =   350
      Left            =   3807
      TabIndex        =   0
      ToolTipText     =   "清除所有未发送编号"
      Top             =   4935
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgSample 
      Height          =   4455
      Left            =   150
      TabIndex        =   9
      Top             =   75
      Width           =   7125
      _cx             =   12568
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      ForeColorSel    =   -2147483632
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin MSComctlLib.StatusBar stbInfo 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   5385
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13097
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtTmp 
      Height          =   255
      Left            =   4020
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmLabMainSendSample.frx":000C
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "本次发送       个"
      Height          =   180
      Left            =   1980
      TabIndex        =   13
      Top             =   4620
      Width           =   1530
   End
   Begin VB.Image img已发 
      Height          =   240
      Left            =   2940
      Picture         =   "frmLabMainSendSample.frx":0091
      Top             =   4875
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNull 
      Height          =   255
      Left            =   3975
      Top             =   4905
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl盘号 
      Caption         =   "盘号"
      Height          =   180
      Left            =   3930
      TabIndex        =   11
      Top             =   4605
      Width           =   435
   End
   Begin VB.Label lbl起始杯号 
      Caption         =   "起始杯号"
      Height          =   180
      Left            =   5415
      TabIndex        =   10
      Top             =   4590
      Width           =   765
   End
End
Attribute VB_Name = "frmLabMainSendSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrId() As String
Private mfrmMain As frmLabMain
Private Enum mCol
    ID = 0: 序号: 已发送: 选择: 标本号: 盘号: 杯号: 仪器id: 紧急: 核收时间
End Enum
Private mlngSelect As Long '已选择标本数
Private mlngSend As Long    '已发送的
Private mlngNoSend As Long  '未发送的


Public Sub ShowME(ByRef strIDList() As String, ByVal frmMain As frmLabMain)
    mstrId = strIDList
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain

End Sub

Private Sub RefreshData()
        Dim lngCount As Long
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim str杯号 As String, strIDs As String
        Dim iRow As Integer, lngSeq As Long
        On Error GoTo errHandle
100     txt杯号 = 1
102     txt盘号 = 0
104     mlngSelect = 0
106     mlngNoSend = 0
108     mlngSend = 0
        lngSeq = 0
110     cmdSend.Enabled = False
112     With vfgSample

114         .Rows = 2: .Cols = 10: .FixedRows = 1: .FixedCols = 0
116         .Clear

118         .TextMatrix(0, mCol.ID) = "id":         .ColWidth(mCol.ID) = 0
            .TextMatrix(0, mCol.序号) = "序号":     .ColWidth(mCol.序号) = 600
120         .TextMatrix(0, mCol.选择) = " ":        .ColWidth(mCol.选择) = 300
122         .TextMatrix(0, mCol.已发送) = " ": .ColWidth(mCol.已发送) = 300

124         .TextMatrix(0, mCol.标本号) = "标本号": .ColWidth(mCol.标本号) = 1200
126         .TextMatrix(0, mCol.盘号) = "盘号":     .ColWidth(mCol.盘号) = 1200
128         .TextMatrix(0, mCol.杯号) = "杯号":     .ColWidth(mCol.杯号) = 1200

130         .TextMatrix(0, mCol.仪器id) = "仪器id": .ColWidth(mCol.仪器id) = 0
132         .TextMatrix(0, mCol.紧急) = "紧急": .ColWidth(mCol.紧急) = 0
134         .TextMatrix(0, mCol.核收时间) = "核收时间": .ColWidth(mCol.核收时间) = 1800
136         .Editable = flexEDKbdMouse

138         For lngCount = 0 To .Cols - 1
140             .FixedAlignment(lngCount) = flexAlignCenterCenter
142             If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
            Next

144         For iRow = LBound(mstrId) To UBound(mstrId)
146             strIDs = mstrId(iRow)
148             If strIDs <> "" Then
150                 If Left$(strIDs, 1) = "," Then strIDs = Mid$(strIDs, 2)
                
152                 strSQL = "Select /*+ Rule */ A.Rowid,A.id,A.标本序号,A.杯号,A.是否传送,A.仪器id,A.紧急,A.核收时间" & vbNewLine & _
                            "From 检验标本记录 A, Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)) B" & vbNewLine & _
                            "Where A.ID = B.Column_Value Order by  Lpad(A.标本序号,9,'0'),A.核收时间 "
154                 Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
156                 Do Until rsTmp.EOF
158                     .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
160                     .TextMatrix(.Rows - 1, mCol.标本号) = Trim("" & rsTmp!标本序号)
        
162                     str杯号 = Trim("" & rsTmp!杯号)
164                     If InStr(str杯号, ",") > 0 Then
166                         .TextMatrix(.Rows - 1, mCol.盘号) = Split(str杯号, ",")(0)
168                         .TextMatrix(.Rows - 1, mCol.杯号) = Split(str杯号, ",")(1)
170                         If Val(Split(str杯号, ",")(1)) > Val(txt杯号) Then txt杯号 = Split(str杯号, ",")(1)
172                         If Split(str杯号, ",")(0) <> Trim(txt盘号) And Trim(Split(str杯号, ",")(0)) <> "" Then txt盘号 = Split(str杯号, ",")(0)
                        End If
174                     If Val("" & rsTmp!是否传送) = 1 Then
176                         .Cell(flexcpPicture, .Rows - 1, mCol.已发送) = img已发.Picture
178                         .Cell(flexcpPictureAlignment, .Rows - 1, mCol.已发送) = flexPicAlignLeftCenter
180                         If Chk已发送.Value = 1 Then
182                             .RowHidden(.Rows - 1) = False
184                             mlngSend = mlngSend + 1 '已发送计数
                                lngSeq = lngSeq + 1
                                .TextMatrix(.Rows - 1, mCol.序号) = lngSeq
                            Else
186                             .RowHidden(.Rows - 1) = True
                            End If
                        
                        Else
                            lngSeq = lngSeq + 1
                            .TextMatrix(.Rows - 1, mCol.序号) = lngSeq
188                         .TextMatrix(.Rows - 1, mCol.已发送) = ""
190                         .Cell(flexcpPicture, .Rows - 1, mCol.已发送) = imgNull.Picture
192                         .Cell(flexcpPictureAlignment, .Rows - 1, mCol.已发送) = flexPicAlignLeftCenter
194                         mlngNoSend = mlngNoSend + 1 ' 未发送计数
                        End If
        
196                     .TextMatrix(.Rows - 1, mCol.仪器id) = Val("" & rsTmp!仪器id)
198                     .TextMatrix(.Rows - 1, mCol.紧急) = Val("" & rsTmp!紧急)
200                     .TextMatrix(.Rows - 1, mCol.核收时间) = Format("" & rsTmp!核收时间, "yyyy-MM-dd hh:mm:ss")
202                     .Rows = .Rows + 1
204                     rsTmp.MoveNext
                    Loop
                    
206                 If .Rows > .FixedRows Then .Rows = .Rows - 1
                End If
            Next
        End With
208     Call refreshStb
210     If vfgSample.Rows > vfgSample.FixedRows Then vfgSample.Cell(flexcpChecked, vfgSample.FixedRows, mCol.选择, vfgSample.Rows - 1, mCol.选择) = flexUnchecked     '全部设为未选

        Exit Sub
errHandle:
'        WriteToLog "SendSample.refreshData," & CStr(Erl()) & "行," & Err.Description
212     If ErrCenter() = 1 Then
214         Resume
        End If
End Sub
Private Sub refreshStb()
    stbInfo.Panels(1).Text = "未发送标本:" & mlngNoSend & " 已发送标本:" & mlngSend & " 已选择:" & mlngSelect
    If mlngSelect > 0 Then cmdSend.Enabled = True
    
End Sub
Private Sub Chk已发送_Click()
    Dim lngRow As Long
    Dim lngSeq As Long
    mlngSend = 0
    lngSeq = 0
    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            If Chk已发送.Value = 1 Then
                .RowHidden(lngRow) = False
                lngSeq = lngSeq + 1
                .TextMatrix(lngRow, mCol.序号) = lngSeq
                If .Cell(flexcpPicture, lngRow, mCol.已发送) <> imgNull.Picture Then
                    mlngSend = mlngSend + 1
                End If
            Else
                If .Cell(flexcpPicture, lngRow, mCol.已发送) <> imgNull.Picture Then
                    .RowHidden(lngRow) = True
                Else
                    lngSeq = lngSeq + 1
                    .TextMatrix(lngRow, mCol.序号) = lngSeq
                End If
            End If
        Next
        
        refreshStb
    End With
End Sub

Private Sub cmdAuto_Click()
    Dim str盘号 As String, lng杯号 As Long
    Dim blnAdd As Boolean, lng开始行 As Long
    Dim lngRow As Long
    rtxtTmp.Text = ""
    With vfgSample
        str盘号 = Trim(txt盘号)
        lng杯号 = Val(txt杯号)
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.已发送) = imgNull.Picture Then
                If Trim(.TextMatrix(lngRow, mCol.盘号)) <> "" And Trim(.TextMatrix(lngRow, mCol.杯号)) <> "" Then
                    rtxtTmp.Text = rtxtTmp.Text & "|" & .TextMatrix(lngRow, mCol.盘号) & "," & .TextMatrix(lngRow, mCol.杯号)
                End If
            Else
                If Trim(.TextMatrix(lngRow, mCol.盘号)) <> "" And Trim(.TextMatrix(lngRow, mCol.杯号)) <> "" Then
                    rtxtTmp.Text = rtxtTmp.Text & "|" & .TextMatrix(lngRow, mCol.盘号) & "," & .TextMatrix(lngRow, mCol.杯号)
                End If
            End If
        Next

        lng开始行 = .FixedRows

        For lngRow = lng开始行 To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.已发送) = imgNull.Picture Then
                If Trim(.TextMatrix(lngRow, mCol.盘号)) = "" And Trim(.TextMatrix(lngRow, mCol.杯号)) = "" Then
                    blnAdd = False
                    Do While Not blnAdd
                        If InStr(rtxtTmp.Text & "|", "|" & str盘号 & "," & lng杯号 & "|") <= 0 Then
                            .TextMatrix(lngRow, mCol.盘号) = str盘号
                            .TextMatrix(lngRow, mCol.杯号) = lng杯号
                            Call vfgSample_AfterEdit(lngRow, mCol.杯号)
                            blnAdd = True
                            rtxtTmp.Text = rtxtTmp.Text & "|" & str盘号 & "," & lng杯号
                        Else
                            lng杯号 = lng杯号 + 1
                        End If
                    Loop
                End If
            End If
        Next

    End With
End Sub

Private Sub cmdDele_Click()
    Dim lngRow As Long

    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.已发送) = imgNull.Picture Then
               .TextMatrix(lngRow, mCol.盘号) = ""
               .TextMatrix(lngRow, mCol.杯号) = ""
            End If
        Next
    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim lngRow As Long
    Dim str核收时间 As String, lng急诊 As Long, str标本号 As String, lng仪器id As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngSendMax As Long '发送个数
    
    
    cmdAll.Enabled = False
    cmdAuto.Enabled = False
    cmdClear.Enabled = False
    cmdDele.Enabled = False
    cmdExit.Enabled = False
    cmdSend.Enabled = False
    Chk已发送.Enabled = False
    
    If mlngSelect <= 0 Then
        
        Exit Sub
    End If
    lngSendMax = Val(txtSend.Text)
    If lngSendMax < 0 Then lngSendMax = 0
    If lngSendMax > mlngSelect Then lngSendMax = 0
    
    With vfgSample
        Call .Select(.Row, .COL)
'        WriteToLog "---> 本次发送开始 --->"
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngRow, mCol.选择) = flexChecked Then
                str核收时间 = Format(CDate(.TextMatrix(lngRow, mCol.核收时间)), "yyyy-MM-dd")
                lng急诊 = Val(.TextMatrix(lngRow, mCol.紧急))
                str标本号 = Val(.TextMatrix(lngRow, mCol.标本号))
                lng仪器id = Val(.TextMatrix(lngRow, mCol.仪器id))

                SendSample mfrmMain.WinsockC, mfrmMain.WinsockC.LocalIP, lng仪器id, str核收时间, str标本号, "", False, lng急诊
                strSQL = "Select 是否传送,样本条码,姓名,核收时间,杯号 From 检验标本记录 Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, mCol.ID)))
                
                Do Until rsTmp.EOF
'                    WriteToLog "发送：" & str标本号 & " " & rsTmp!姓名 & " " & rsTmp!样本条码 & " " & rsTmp!杯号 & " " & Format(rsTmp!核收时间, "yyyy-MM-dd HH:mm:ss")
                    If Val("" & rsTmp!是否传送) = 1 Then
                        .Cell(flexcpPicture, lngRow, mCol.已发送) = img已发.Picture
                        .Cell(flexcpPictureAlignment, lngRow, mCol.已发送) = flexPicAlignLeftCenter
                    Else
                        .Cell(flexcpPicture, lngRow, mCol.已发送) = imgNull.Picture
                        .Cell(flexcpPictureAlignment, lngRow, mCol.已发送) = flexPicAlignLeftCenter
                    End If
                    rsTmp.MoveNext
                Loop
            End If
        Next
'        WriteToLog "<--- 本次发送结束 <---"
    End With

    cmdAll.Enabled = True
    cmdAuto.Enabled = True
    cmdClear.Enabled = True
    cmdDele.Enabled = True
    cmdExit.Enabled = True
    cmdSend.Enabled = True
    Chk已发送.Enabled = True
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub cmdAll_Click()
    Dim lngRow As Long
    Dim lngCount As Long
    mlngSelect = 0
    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            'If Not (Trim(.TextMatrix(lngRow, mCol.盘号)) = "" Or Trim(.TextMatrix(lngRow, mCol.杯号)) = "") Then
            If .RowHidden(lngRow) = False Then
                .Cell(flexcpChecked, lngRow, mCol.选择) = flexChecked
                mlngSelect = mlngSelect + 1
            End If
            'End If
        Next
    
        refreshStb
    End With
End Sub

Private Sub cmdClear_Click()
    vfgSample.Cell(flexcpChecked, 1, mCol.选择, vfgSample.Rows - 1, mCol.选择) = flexUnchecked
    mlngSelect = 0
    refreshStb
End Sub

Private Sub vfgSample_AfterEdit(ByVal Row As Long, ByVal COL As Long)
    Dim strSQL As String
    Dim str杯号 As String, str盘号 As String

    With vfgSample
        str杯号 = Trim(.TextMatrix(Row, mCol.杯号))
        str盘号 = Trim(.TextMatrix(Row, mCol.盘号))
        If Not (str杯号 = "" And str盘号 = "") Then
            strSQL = "ZL_检验标本记录_杯号(" & Val(.TextMatrix(Row, mCol.ID)) & ",'" & str盘号 & "," & str杯号 & "')"
        Else
            strSQL = "ZL_检验标本记录_杯号(" & Val(.TextMatrix(Row, mCol.ID)) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
End Sub

Private Sub vfgSample_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If InStr("," & mCol.杯号 & "," & mCol.盘号 & ",", "," & COL & ",") <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub vfgSample_Click()
    With vfgSample
        If .MouseCol = mCol.选择 Then
            'If Trim(.TextMatrix(.Row, mCol.杯号)) = "" Or Trim(.TextMatrix(.Row, mCol.盘号)) = "" Then Exit Sub
            mlngSelect = mlngSelect + IIf(.Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked, 1, -1)
            .Cell(flexcpChecked, .Row, mCol.选择) = IIf(.Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked, flexChecked, flexUnchecked)
            
        End If
        
        Call refreshStb
    End With
End Sub

Private Sub vfgSample_EnterCell()
    With vfgSample
        Dim blnCancle As Boolean
        Call vfgSample_BeforeEdit(.Row, .COL, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
End Sub

Private Sub vfgSample_LeaveCell()
    With vfgSample
        Dim blnCancle As Boolean
        Call vfgSample_BeforeEdit(.Row, .COL, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgSample_ValidateEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Dim intRow As Integer
    Dim str杯号 As String
    Dim str盘号  As String
    With vfgSample
        If IsNumeric(.EditText) = False And .EditText <> "" Then Cancel = True
        Select Case COL

            Case mCol.杯号

                str盘号 = Trim(.TextMatrix(Row, mCol.盘号))
                str杯号 = Trim(.EditText)

            Case mCol.盘号

                str盘号 = Trim(.EditText)
                str杯号 = Trim(.TextMatrix(Row, mCol.杯号))
        End Select

        If str盘号 <> "" And str杯号 <> "" Then
            For intRow = .FixedRows To .Rows - 1
                If intRow <> Row Then
                    If Trim(.TextMatrix(intRow, mCol.盘号)) <> "" And Trim(.TextMatrix(intRow, mCol.杯号)) <> "" Then
                        If Trim(.TextMatrix(intRow, mCol.盘号)) = str盘号 And Trim(.TextMatrix(intRow, mCol.杯号)) = str杯号 Then
                            Cancel = True
                            .Cell(flexcpChecked, Row, mCol.选择) = flexUnchecked
                       End If
                    End If
                End If
            Next
        End If

    End With

End Sub






