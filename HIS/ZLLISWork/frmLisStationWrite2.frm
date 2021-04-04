VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLisStationWrite2 
   BorderStyle     =   0  'None
   Caption         =   "细菌报告填写"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8700
   Icon            =   "frmLisStationWrite2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2685
      Left            =   6570
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "结果名称"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "结果内容"
         Object.Width           =   2540
      EndProperty
   End
   Begin zl9LisWork.VsfGrid vsf 
      Height          =   1845
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   3254
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7095
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationWrite2.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationWrite2.frx":1C94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   2250
      Left            =   240
      TabIndex        =   5
      Top             =   1950
      Width           =   6270
      Begin VB.CheckBox chkLast 
         Caption         =   "上次结果"
         Height          =   180
         Left            =   3840
         TabIndex        =   10
         Top             =   210
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   45
         TabIndex        =   6
         Top             =   120
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   609
         ButtonWidth     =   3043
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "选用抗生素(&G)"
               Key             =   "选用抗生素"
               Object.ToolTipText     =   "选用抗生素"
               Object.Tag             =   "选用抗生素(&G)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "无药敏检验(&N)"
               Key             =   "无药敏检验"
               Object.ToolTipText     =   "无药敏检验"
               Object.Tag             =   "无药敏检验(&N)"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin zl9LisWork.VsfGrid vsfDetail 
         Height          =   1695
         Left            =   30
         TabIndex        =   1
         Top             =   480
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2990
      End
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5625
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt诊断 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   6630
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4200
      Width           =   1875
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   510
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   4605
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   450
      Locked          =   -1  'True
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5130
      Width           =   4665
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   5430
      Top             =   870
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl诊断 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "诊断信息"
      Height          =   360
      Left            =   6180
      TabIndex        =   12
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "检验备注"
      Height          =   345
      Left            =   75
      TabIndex        =   8
      Top             =   5010
      Width           =   375
   End
   Begin VB.Label lblResult 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结果评语"
      Height          =   345
      Left            =   75
      TabIndex        =   9
      Top             =   4530
      Width           =   375
   End
End
Attribute VB_Name = "frmLisStationWrite2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Private mlngKey As Long    '标本ID
Private mDeviceID As Long
Private mstrType As String '检验类型
Private mblnEdit As Boolean '是否允许编辑
Private mbytRedoNumber As Long '重做次数
Private mblSelectHistory As Boolean '是否选择历史
Private mlngHistoryID As Long       '选择的历史ID

Private WithEvents mfrmRequest As frmLabRequest                     '核收登记窗体
Attribute mfrmRequest.VB_VarHelpID = -1

Private mrsSave As New ADODB.Recordset
Private mblnChangeEdit As Boolean, mlngItemID As Long '微生物项目ID


Private Enum mCol
    细菌名称 = 1
    菌落计数
    培养描述
    耐药机制
    上次菌落计数
    抗生素名称 = 1
    药敏方法
    检验结果
    结果标志
    上次结果
    上次标志
End Enum
Private mlng抗生素分组id As Long

Public Event StartEdit(Cancel As Boolean)

Private Sub WriteRecord(ByVal lngRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    '1.先删除原来的记录,可能是已经删除
    On Error GoTo ErrHand
    
    If Vsf.Rows > 0 And lngRow = 0 Then
        lngRow = 1
    End If
    
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(lngRow))
    
'    On Error Resume Next
    
    Call DeleteRecord(mrsSave)
    
    '2.再添加现在的记录
    For mlngLoop = 1 To vsfDetail.Rows - 1
        If Val(vsfDetail.RowData(mlngLoop)) > 0 Then
            mrsSave.AddNew
            mrsSave("Key").Value = Val(Vsf.RowData(lngRow))
            mrsSave("ID").Value = Val(vsfDetail.RowData(mlngLoop))
            mrsSave("Group").Value = mlng抗生素分组id
            mrsSave("抗生素名称").Value = vsfDetail.TextMatrix(mlngLoop, mCol.抗生素名称)
            mrsSave("检验结果").Value = vsfDetail.TextMatrix(mlngLoop, mCol.检验结果)
            mrsSave("结果标志").Value = vsfDetail.TextMatrix(mlngLoop, mCol.结果标志)
            mrsSave("药敏方法").Value = vsfDetail.TextMatrix(mlngLoop, mCol.药敏方法)
        End If
    Next
    
ErrHand:
    
End Sub

Private Sub ReadRecord(ByVal lngRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(lngRow))
    If mrsSave.RecordCount > 0 Then
        mrsSave.MoveFirst
        
        
        '1.再填写先保存的抗生素细目
        For mlngLoop = 1 To mrsSave.RecordCount
            
            vsfDetail.Rows = mlngLoop + 1
            
            vsfDetail.RowData(mlngLoop) = Val(mrsSave("ID").Value)
            mlng抗生素分组id = mrsSave("Group").Value
            vsfDetail.TextMatrix(mlngLoop, 0) = mlngLoop
            vsfDetail.TextMatrix(mlngLoop, mCol.抗生素名称) = mrsSave("抗生素名称").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.检验结果) = mrsSave("检验结果").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.结果标志) = mrsSave("结果标志").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.药敏方法) = mrsSave("药敏方法").Value

            Select Case UCase(Left(vsfDetail.TextMatrix(mlngLoop, mCol.结果标志), 1))
            Case "R"
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.红色
            Case "I"
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.兰色
            Case Else
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.黑色
            End Select

            mrsSave.MoveNext
        Next
        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
        '写入上次结果
        If chkLast.Value = 1 Then Call LoadLastValue
        
    End If
End Sub

Private Sub LoadDefaultGroup(ByVal lngKey As Long)
    '--------------------------------------------------------------------------------------------------------
    '功能:产生当前细菌缺省对应的抗生素分组的抗生素项目
    '参数:lngKey            细菌id
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSql As String
    
    On Error GoTo ErrHand
        
    Call ClearGrid(vsfDetail)

    mstrSql = "SELECT '' AS 序号,A.ID,A.中文名 AS 抗生素名称,'' AS 检验结果," & _
            "Decode(Upper(E.默认药敏),'R','R-耐药','I','I-中介','S','S-敏感',E.默认药敏) AS 结果类型,B.抗生素分组ID,'' As 药敏方法,'' as 上次结果,'' as 上次类型 " & _
            "FROM 检验用抗生素 A,检验抗生素用药 C,检验抗生素组 D,检验细菌抗生素 B,检验细菌 E " & _
            "WHERE A.ID=C.抗生素ID AND C.抗生素分组ID=D.ID AND D.ID=B.抗生素分组ID And B.细菌ID=E.ID AND E.ID= [1] Order By Decode(B.缺省标志,1,1,0) Desc,A.编码"
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        rs.filter = "抗生素分组ID=" & rs("抗生素分组ID")
        mlng抗生素分组id = rs("抗生素分组ID")
        vsfDetail.TextMatrix(0, 0) = "序号"
        Call FillGrid(vsfDetail, rs)
        vsfDetail.TextMatrix(0, 0) = ""
        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function ShowOpenList(objVsf As Object, Optional strText As String, Optional blnWhere As Boolean = True, Optional ByVal bytMode As Byte = 1) As Byte
    '--------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的细菌目录
    '返回:出错返回2;成功返回1;取消返回0
    '--------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    ShowOpenList = 2
    
    Select Case bytMode
    Case 1
        strSQL = "SELECT A.ID,A.编码,A.中文名,A.英文名,A.简码,B.中文名称 AS 类型 " & _
                "FROM 检验细菌 A,检验细菌类型 B " & _
                "WHERE A.类型ID=B.ID " & _
                    "AND (A.编码 Like [3] OR A.中文名 Like [2] OR Upper(A.简码) Like [3])"
    Case 2
        strSQL = "SELECT B.ID," & _
                           "NULL + 0 AS 上级id," & _
                           "0 AS 末级," & _
                           "'' AS 编码," & _
                           "'[' || B.编码 || ']' || B.名称 AS 名称," & _
                           "'' AS 英文名," & _
                           "'' AS 简码 " & _
                      "FROM 检验细菌抗生素 A, 检验抗生素组 B " & _
                     "WHERE A.抗生素分组ID = B.ID And A.细菌ID = [1] AND (B.编码 LIKE [2] OR B.名称 LIKE [2] B.简码 LIKE [2])"
                     
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Vsf.RowData(Vsf.Row)), "%" & strText & "%")
        
        If rs.BOF Then
            ShowOpenList = 0
            Exit Function
        End If
        
        If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
        strSQL = "SELECT B.ID," & _
                           "NULL + 0 AS 上级id," & _
                           "0 AS 末级," & _
                           "'' AS 编码," & _
                           "'[' || B.编码 || ']' || B.名称 AS 名称," & _
                           "'' AS 英文名," & _
                           "'' AS 简码 " & _
                      "FROM 检验细菌抗生素 A, 检验抗生素组 B " & _
                     "WHERE A.抗生素分组ID = B.ID And A.细菌ID = [1] AND (B.编码 LIKE [3] OR B.名称 LIKE [2] B.简码 LIKE [3] )" & _
                    "Union All " & _
                      "SELECT ROWNUM AS ID," & _
                             "B.抗生素分组ID AS 上级id," & _
                             "1 AS 末级," & _
                             "A.编码," & _
                             "A.中文名 AS 名称," & _
                             "A.英文名," & _
                             "A.简码 " & _
                        "FROM 检验用抗生素 A, 检验抗生素用药 B " & _
                       "Where A.ID = B.抗生素ID Order By A.编码"
    Case 3
        strSQL = "SELECT ROWNUM AS ID,A.编码,A.简码,A.名称,A.说明 " & _
                "FROM 检验培养文字 A " & _
                "WHERE (A.编码 Like [3] OR A.简码 Like [3] )"
    Case 4
        strSQL = "select ID,编码,中文名,英文名 from 检验用抗生素 where (编码 like [3] or 中文名 like [3] or 英文名 like [3]) "
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Vsf.RowData(Vsf.Row)), "%" & strText & "%", "%" & UCase(strText) & "%")
    
    
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
    Call CalcPosition(sglX, sglY, objVsf)
    
    Select Case bytMode
    Case 1
        strLvw = "编码,900,0,1;中文名,1800,0,0;英文名,900,0,0;简码,900,0,0;类型,900,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 5400, 4500, Me.Name & "\检验细菌选择", "请从下表中选择一个细菌项目") Then
            GoTo Over
        End If
    Case 3
        strLvw = "编码,900,0,1;简码,900,0,1;名称,1800,0,0;说明,1800,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 4500, 4500, Me.Name & "\检验培养文字选择", "请从下表中选择一个培养文字") Then
            GoTo Over
        End If
    Case 4
        strLvw = "编码,900,0,1;中文名,1800,0,0;英文名,1800,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 4500, 4500, Me.Name & "\检验抗生素选择", "请从下表中选择一个抗生素") Then
            GoTo Over
        End If
    End Select
        
    Exit Function
    
Over:

    Select Case bytMode
    Case 1
        If CheckHave(zlCommFun.Nvl(rs("ID").Value)) Then
            MsgBox "选择的项目“" & zlCommFun.Nvl(rs("中文名").Value) & "”以前已经选择！", vbInformation, gstrSysName
            Exit Function
        End If
        objVsf.EditText = zlCommFun.Nvl(rs("中文名").Value)
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("中文名").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("中文名").Value)
        objVsf.RowData(objVsf.Row) = zlCommFun.Nvl(rs("ID").Value)
    Case 3
        objVsf.EditText = zlCommFun.Nvl(rs("说明").Value)
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("说明").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("说明").Value)
    Case 4
        objVsf.RowData(objVsf.Row) = Nvl(rs("ID"))
'        objVsf.EditText = Nvl(rs("中文名").Value)
        objVsf.TextMatrix(objVsf.Row, mCol.抗生素名称) = Nvl(rs("中文名"))
    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ShowOpenTree(objVsf As Object, Optional ByVal bytMode As Byte = 1) As Byte
    '-----------------------------------------------------------------------------------------
    '功能:打开树型+列表结构的诊疗项目数据
    '返回:出错返回2;成功返回1;取消返回0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim sglX As Single
    Dim sglY As Single
    
    On Error GoTo ErrHand
    
    ShowOpenTree = 2
    
    Select Case bytMode
    Case 1
        strSQL = "SELECT A.ID,NULL+0 AS 上级id,0 AS 末级,A.编码,'['||A.编码||']'||A.中文名称 AS 名称,'' AS 英文名,'' AS 简码,'' AS 类型 " & _
                "FROM 检验细菌类型 A " & _
                "UNION ALL " & _
                "SELECT A.ID,A.类型ID AS 上级id,1 AS 末级,A.编码,A.中文名 AS 名称,A.英文名,A.简码,B.中文名称 AS 类型 " & _
                "FROM 检验细菌 A,检验细菌类型 B " & _
                "WHERE A.类型ID=B.ID "
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF Then
        ShowOpenTree = 0
        Exit Function
    End If

    Call CalcPosition(sglX, sglY, objVsf)
    Select Case bytMode
    Case 1
        
        strLvw = "编码,1200,0,1;名称,1800,0,0;英文名,900,0,0;简码,900,0,0"
        If frmSelectExplorer.ShowSelect(Me, rs, sglX, sglY, 5400, 2400, _
                                    objVsf.CellHeight, "项目树型选择", strLvw, "请选择一个检验项目") Then
                                    
            If CheckHave(zlCommFun.Nvl(rs("ID").Value)) Then
                MsgBox "选择的项目“" & zlCommFun.Nvl(rs("名称").Value) & "”以前已经选择！", vbInformation, gstrSysName
                Exit Function
            End If
            GoTo Over
        End If
    End Select
    
    Exit Function
    
Over:

    Select Case bytMode
    Case 1
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("名称").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("名称").Value)
        objVsf.RowData(objVsf.Row) = zlCommFun.Nvl(rs("ID").Value)
    End Select
    
    ShowOpenTree = 1
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    For mlngLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(mlngLoop)) = lngKey And Vsf.Row <> mlngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function ReadData() As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSql As String
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    
    mstrSql = "SELECT A.报告结果,A.检验人,A.检验时间,A.审核人,A.审核时间,A.检验备注,A.备注,a.初审人,a.初审时间 FROM 检验标本记录 A WHERE A.ID= [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If Not rs.EOF Then
        mbytRedoNumber = Nvl(rs("报告结果"), 0)
        Me.txtComment = Nvl(rs("检验备注"))
        Me.txtResult = Nvl(rs("备注"))
        
        With sbrInfo
            .Panels(1).Text = "报告人：" & Nvl(rs("检验人"))
            .Panels(2).Text = "报告时间：" & IIf(IsNull(rs("检验时间")), "", Format(rs("检验时间"), "yyyy-MM-dd hh:mm"))
            If Nvl(rs("审核人")) <> "" Then
                .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
            Else
                If Nvl(rs("初审人")) <> "" Then
                    .Panels(3).Text = "初审人：" & Nvl(rs("初审人"))
                    .Panels(4).Text = "初审时间：" & IIf(IsNull(rs("初审时间")), "", Format(rs("初审时间"), "yyyy-MM-dd hh:mm"))
                Else
                    .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                    .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                End If
            End If
        End With
    Else
        mstrSql = "SELECT A.报告结果,A.检验人,A.检验时间,A.审核人,A.审核时间,A.检验备注,A.备注,a.初审人,a.初审时间 FROM 检验标本记录 A WHERE A.ID= [1] "
        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
        If Not rs.EOF Then
            mbytRedoNumber = Nvl(rs("报告结果"), 0)
            Me.txtComment = Nvl(rs("检验备注"))
            Me.txtResult = Nvl(rs("备注"))
            
            With sbrInfo
                .Panels(1).Text = "报告人：" & Nvl(rs("检验人"))
                .Panels(2).Text = "报告时间：" & IIf(IsNull(rs("检验时间")), "", Format(rs("检验时间"), "yyyy-MM-dd hh:mm"))
                If Nvl(rs("审核人")) <> "" Then
                    .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                    .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                Else
                    If Nvl(rs("初审人")) <> "" Then
                        .Panels(3).Text = "初审人：" & Nvl(rs("初审人"))
                        .Panels(4).Text = "初审时间：" & IIf(IsNull(rs("初审时间")), "", Format(rs("初审时间"), "yyyy-MM-dd hh:mm"))
                    Else
                        .Panels(3).Text = "审核人：" & Nvl(rs("审核人"))
                        .Panels(4).Text = "审核时间：" & IIf(IsNull(rs("审核时间")), "", Format(rs("审核时间"), "yyyy-MM-dd hh:mm"))
                    End If
                End If
            End With
        Else
            mbytRedoNumber = 0
            Me.txtComment = ""
            Me.txtResult = ""
            
            With sbrInfo
                .Panels(1).Text = "报告人："
                .Panels(2).Text = "报告时间："
                .Panels(3).Text = "审核人："
                .Panels(4).Text = "审核时间："
            End With
        End If
    End If
    
    mstrSql = "SELECT C.报告项目ID FROM 检验标本记录 A,检验申请项目 B,检验报告项目 C " & _
                    "WHERE A.ID=B.标本ID And B.诊疗项目ID=C.诊疗项目ID " & _
                        "AND A.ID= [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If rs.BOF = False Then
        mlngItemID = Nvl(rs("报告项目ID"), 0)
    Else
        mlngItemID = 0
    End If
    
    mstrSql = "SELECT B.ID,D.报告结果,B.中文名 AS 检验项目," & _
                    "A.检验结果 AS 检验结果,A.培养描述 as 结果描述,'' as 上次结果,a.耐药机制 " & _
                    "FROM 检验普通结果 A,检验细菌 B,检验标本记录 D " & _
                    "WHERE A.细菌id = B.ID " & _
                        "AND A.记录类型 = [1] " & _
                        "AND D.ID=A.检验标本ID " & _
                        "AND D.ID= [2] Order by B.编码"
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mbytRedoNumber, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If rs.BOF = False Then
        Vsf.TextMatrix(0, 0) = "#"
        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
    Else
        ResetVsf Vsf
        ResetVsf vsfDetail
    End If
    
    '1.先删除原来的记录
    mrsSave.filter = ""
    Call DeleteRecord(mrsSave)
    
    mstrSql = "SELECT C.细菌ID AS Key,B.ID,B.中文名 AS 抗生素名称, A.结果 AS 检验结果,c.药敏组ID, " & _
            "DECODE(A.结果类型,'R','R-耐药','I','I-中介','S','S-敏感',A.结果类型) AS 结果类型, " & _
            "DECODE(A.药敏方法,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As 药敏方法 " & _
             "FROM 检验药敏结果 A, 检验用抗生素 B,检验普通结果 C " & _
            "Where A.抗生素ID = B.ID And C.ID=A.细菌结果ID AND C.记录类型=A.记录类型 AND C.检验标本id= [1] AND C.记录类型= [2] Order By B.编码"
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey), mbytRedoNumber)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            mrsSave.AddNew
            mrsSave("Key").Value = zlCommFun.Nvl(rs("Key"), 0)
            mrsSave("Group").Value = zlCommFun.Nvl(rs("药敏组ID"), 0)
            mrsSave("ID").Value = zlCommFun.Nvl(rs("ID"), 0)
            mrsSave("抗生素名称").Value = zlCommFun.Nvl(rs("抗生素名称"))
            mrsSave("检验结果").Value = zlCommFun.Nvl(rs("检验结果"))
            mrsSave("结果标志").Value = zlCommFun.Nvl(rs("结果类型"))
            mrsSave("药敏方法").Value = zlCommFun.Nvl(rs("药敏方法"))
            
            rs.MoveNext
        Loop
    End If
    
    
    
    Call vsf_AfterRowColChange(0, 1, 1, 1)
    
    
    If mblSelectHistory = True Then mblnChangeEdit = True
    
    mblSelectHistory = False
    
    '写入上次结果
    If chkLast.Value = 1 Then Call LoadLastValue
    
     '写入诊断信息
    Me.txt诊断.Text = ""
    gstrSql = "Select b.医嘱id, b.项目, b.排列, b.内容" & vbNewLine & _
                "From 检验标本记录 a, 病人医嘱附件 b" & vbNewLine & _
                "Where a.医嘱id = b.医嘱id and a.ID = [1] " & vbNewLine & _
                "Order By 医嘱id, 排列"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("项目")) & ":" & Replace(Nvl(rs("内容")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txt诊断.Text = strTmp
    
    ReadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：显示数据
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    mlngKey = lngKey
    
'    SetEditState False
    mlngHistoryID = 0
    Call Form_Resize
    '初始仪器列表
    If ReadData = False Then Exit Function
    
    zlRefresh = True
End Function

Public Function ZlEditStart(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：编辑数据
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------
    SetEditState True
    '初始仪器列表
    If mlngKey <> lngKey Then
        mlngKey = lngKey
        If ReadData = False Then Exit Function
    End If
    mblnChangeEdit = False
    ZlEditStart = True
    With Vsf
        If Val(.RowData(.Row)) > 0 Then
            
        Else
            .Col = mCol.细菌名称
        End If
        .EditMode(.Col) = 1
        .SetFocus
        
        ShowValue .Col - 1
    End With
    With Vsf
        .EditMode(mCol.检验结果) = 1
        .EditMode(mCol.药敏方法) = 1
        .EditMode(mCol.检验结果) = 1
    End With
End Function

Public Function ZlSave() As Boolean
    '先保存一下当前细菌的抗生素
    Call WriteRecord(Vsf.Row)
    
    If SaveData() = False Then Exit Function

    ZlSave = True
End Function

Public Function ZlCancel() As Boolean
    '提示是否保存
    SetEditState False
    
    ZlCancel = True
End Function

Public Function ZlClearForm() As Boolean
    With sbrInfo
        .Panels(1).Text = "报告人："
        .Panels(2).Text = "报告时间："
        .Panels(3).Text = "审核人："
        .Panels(4).Text = "审核时间："
    End With

    Me.txtComment = ""
    Me.txtResult = ""
    ResetVsf Vsf
    ResetVsf vsfDetail
End Function

Private Sub SetEditState(ByVal blnEdit As Boolean)
    mblnEdit = blnEdit
'    Vsf.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    tbr.Enabled = blnEdit
'    vsfDetail.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    
    txtComment.Locked = Not blnEdit
    txtResult.Locked = Not blnEdit
    Me.lvwSelect.Visible = blnEdit
    Call Form_Resize
End Sub

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim bytResultFlag As Byte
    Dim lngKey As Long
    Dim strSQL() As String
    Dim mlngLoop As Long
    Dim blnNoResult As Boolean
    Dim str结果标志 As String
    Dim lngGroup As Long

    If Not mblnChangeEdit Then SaveData = True: Exit Function

    On Error GoTo ErrHand
    ReDim strSQL(1 To 1)

    '读取检验时间
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")

    strSQL(ReDimArray(strSQL)) = "ZL_检验药敏结果_DELETE(" & mlngKey & "," & mbytRedoNumber & ")"

    blnNoResult = True
    For mlngLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(mlngLoop)) > 0 Then
            blnNoResult = False
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(mlngLoop))
            If mrsSave.RecordCount > 0 Then
                lngGroup = mrsSave("Group").Value
            Else
                lngGroup = 0
            End If
            lngKey = zlDatabase.GetNextId("检验普通结果")
            strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_报告填写2(" & lngKey & "," & _
                mlngKey & ",NULL,'" & _
                Vsf.TextMatrix(mlngLoop, mCol.菌落计数) & "'," & _
                "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                mbytRedoNumber & ",NULL,1," & Val(Vsf.RowData(mlngLoop)) & ",NULL,'" & Vsf.TextMatrix(mlngLoop, mCol.培养描述) & "'," & _
                "NULL,NULL,'" & txtComment & "','" & txtResult & "',0,'" & Vsf.TextMatrix(mlngLoop, mCol.耐药机制) & "','" & UserInfo.姓名 & "'," & _
                IIf(lngGroup = 0, "NULL", lngGroup) & ")"
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(mlngLoop))
            If mrsSave.RecordCount > 0 Then
                mrsSave.MoveFirst
                Do While Not mrsSave.EOF
                    If Len(Trim(mrsSave("检验结果").Value)) > 0 Or Len(Trim(mrsSave("结果标志").Value)) > 0 Then
                        If mrsSave("结果标志").Value = "R-耐药" Or mrsSave("结果标志").Value = "I-中介" Or mrsSave("结果标志").Value = "S-敏感" Then
                            str结果标志 = Left(mrsSave("结果标志").Value, 1)
                        Else
                            str结果标志 = mrsSave("结果标志").Value
                        End If
                        strSQL(ReDimArray(strSQL)) = "ZL_检验药敏结果_INSERT(" & lngKey & "," & mrsSave("ID").Value & _
                            ",'" & UserInfo.姓名 & "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                            mrsSave("检验结果").Value & "'," & IIf(mrsSave("结果标志").Value <> "", "'" & str结果标志 & "'", "NULL") & "," & _
                            mbytRedoNumber & ",NULL," & IIf(IsNull(mrsSave("药敏方法")), "NULL", Val(Left(mrsSave("药敏方法"), 1))) & ")"
                    End If
                    mrsSave.MoveNext
                Loop
            End If
        End If
    Next
    If blnNoResult Then
        strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_报告填写2(" & 0 & "," & _
            mlngKey & ",NULL,''," & _
            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
            mbytRedoNumber & ",NULL,1,NULL,NULL,NULL,NULL,NULL,'" & txtComment & "','" & txtResult & "',0,Null,'" & UserInfo.姓名 & "')"
    End If

    blnTran = True

    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    If Signature(mlngKey, gstrDBUser, "报告") = False Then
        Exit Function
    End If
    

    SaveData = True

    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    mblSelectHistory = True
    mlngHistoryID = Control.ID
    ReadData
    
    '刷新病人信息显示窗体
    On Error Resume Next
'    mfrmRequest.zlRefresh m
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    If mlngKey = Control.ID Then
        Control.Caption = Control.Caption & "(当前)"
    End If
    
    If mlngHistoryID = Control.ID Then
        Control.Checked = True
    End If
End Sub

Private Sub chkLast_Click()
    vsfDetail.Body.ColWidth(mCol.上次结果) = IIf(chkLast.Value, 1300, 0)
    vsfDetail.Body.ColWidth(mCol.上次标志) = IIf(chkLast.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.上次菌落计数) = IIf(chkLast.Value, 1350, 0)
    If chkLast.Value Then LoadLastValue
    
End Sub

Private Sub Form_Load()
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    With Vsf
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "检验项目", 2400, 1, "...", 1
        .NewColumn "检验结果", 1350, 1, , 1, 100
        .NewColumn "结果描述", 2000, 1, , 1, 100
        .NewColumn "耐药机制", 2000, 1, , 1, 50
        .NewColumn "上次结果", 1350, 1, , 1, 100
        .FixedCols = 0
    End With
        
    With vsfDetail
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "抗生素名称", 2400, 1, "...", 1
        .NewColumn "药敏方法", 850, 1, " |1-MIC|2-DISK|3-K-B", 1
        .NewColumn "检验结果", 1300, 1, , 1, 20
        .NewColumn "结果类型", 1000, 1, " |R-耐药|I-中介|S-敏感|ESBL|BLAC|SDD|R*", 1
        .NewColumn "上次结果", 1300, 1, , 0, 20
        .NewColumn "上次类型", 1000, 1, , 0, 20
        .FixedCols = 0
    End With
    
    Set mrsSave = New ADODB.Recordset
    With mrsSave
        
        .Fields.Append "Key", adVarChar, 18
        .Fields.Append "Group", adVarChar, 18
        .Fields.Append "ID", adVarChar, 18
        .Fields.Append "抗生素名称", adVarChar, 50
        .Fields.Append "检验结果", adVarChar, 50
        .Fields.Append "结果标志", adVarChar, 50
        .Fields.Append "药敏方法", adVarChar, 50
        .Open
        
    End With
    lvwSelect.Tag = 1 '默认选择细菌结果
    Set mfrmRequest = frmLabRequest                          '核收登记窗体
    
    SetEditState False
    
    chkLast.Value = Val(zlDatabase.GetPara("frmLisStationWrite2_查看上次结果", 100, 1208, 0))
    vsfDetail.Body.ColWidth(mCol.上次结果) = IIf(chkLast.Value, 1300, 0)
    vsfDetail.Body.ColWidth(mCol.上次标志) = IIf(chkLast.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.上次菌落计数) = IIf(chkLast.Value, 1350, 0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtComment.Visible = zlDatabase.GetPara("显示检验备注", 100, 1208, True)
    lblComment.Visible = txtComment.Visible
    
    With txtComment
        .Left = Me.lblComment.Width
        .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        .Width = Me.ScaleWidth - .Left - Me.lbl诊断.Width - Me.txt诊断.Width
    End With
    With lblComment
        .Left = 0
        .Top = txtComment.Top + (Me.txtComment.Height - .Height) / 2
    End With
    
    With txtResult
        .Left = txtComment.Left
        If txtComment.Visible Then
            .Top = txtComment.Top - .Height - 30
        Else
            .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        End If
        .Width = Me.ScaleWidth - .Left - Me.lbl诊断.Width - Me.txt诊断.Width
    End With
    With lblResult
        .Left = 0
        .Top = txtResult.Top + (Me.txtResult.Height - .Height) / 2
    End With
    
    With Me.lbl诊断
        .Top = Me.lblResult.Top
        .Left = Me.txtResult.Left + Me.txtResult.Width
    End With
    
    With Me.txt诊断
        .Top = Me.txtResult.Top
        .Left = Me.lbl诊断.Left + Me.lbl诊断.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - Me.sbrInfo.Height - 30
    End With
    
    
    With lvwSelect
        .Left = Me.ScaleWidth - .Width - 30
        .Top = 0
        .Height = txtResult.Top - 30 - .Top
    End With
    
    With Vsf
        .Left = -15
        .Top = 0
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) - 30 - .Left
        .Height = txtResult / 2 - 30
    End With
    
    With fra
        .Left = -30
        .Top = Vsf.Top + Vsf.Height
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) + 30 - .Left
        .Height = txtResult.Top - .Top - 30
    End With
    
    With vsfDetail
        .Left = 15
        .Width = fra.Width - 75 - .Left
        .Height = fra.Height - .Top - 45
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("frmLisStationWrite2_查看上次结果", Me.chkLast.Value, 100, 1208)
    mblnEdit = False
End Sub

Private Sub lvwSelect_DblClick()
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    If Not mblnEdit Then Exit Sub
    
    Select Case Val(lvwSelect.Tag)
        Case 1 '选择结果
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.菌落计数) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 2 '培养描述
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.培养描述) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 3 '细菌耐药机制
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.耐药机制) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 4 '选择评语
            Me.txtResult.SelText = lvwSelect.SelectedItem.SubItems(1)
            mblnChangeEdit = True
        Case 5 '选择备注
            Me.txtComment.SelText = lvwSelect.SelectedItem.SubItems(1)
            mblnChangeEdit = True
    End Select
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim rs As New ADODB.Recordset, rsTmp  As New ADODB.Recordset, mstrSql As String
    Dim objPoint As POINTAPI
    
    Select Case Button.Key
        Case "选用抗生素"
            mstrSql = "SELECT A.ID,A.编码,A.名称,C.默认药敏 " & _
                "FROM 检验抗生素组 A,检验细菌抗生素 B,检验细菌 C " & _
                " WHERE A.ID=B.抗生素分组ID AND B.细菌ID=" & Val(Vsf.RowData(Vsf.Row)) & _
                " and B.细菌ID = C.ID " & _
                " GROUP BY A.ID,A.编码,A.名称,C.默认药敏"
                
            Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption)
            If rs.BOF = False Then
                
                Call ClientToScreen(tbr.hWnd, objPoint)
                If frmSelectList.ShowSelect(Me, rs, "编码,900,0,1;名称,2400,0,0", objPoint.X * 15, objPoint.Y * 15 + tbr.Height, 3300, 2400, Me.Name & "\抗生素组选择", "请从下表中选择一个抗生素组项目") Then
                    
                    Call ClearGrid(vsfDetail)
                    
'                    mstrSQL = "SELECT '' AS 序号,A.ID,A.中文名 AS 抗生素名称,'' AS 检验结果,'' AS 结果类型,'' As 药敏方法, " & _
                                "'' as 上次结果,'' as 上次类型 " & _
                                "FROM 检验用抗生素 A,检验抗生素用药 C " & _
                                "WHERE A.ID=C.抗生素ID AND C.抗生素分组ID= [1] Order By A.编码"
                    mstrSql = "SELECT '' AS 序号,A.ID,A.中文名 AS 抗生素名称,'' AS 检验结果," & _
                    " Decode(A.药敏方法, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') AS 药敏方法," & _
                    " Decode('" & Nvl(rs("默认药敏")) & "', 'R', 'R-耐药', 'I', 'I-中介', 'S', 'S-敏感', '" & Nvl(rs("默认药敏")) & "') As 结果类型, " & _
                                "'' as 上次结果,'' as 上次类型 " & _
                                "FROM 检验用抗生素 A,检验抗生素用药 C " & _
                                "WHERE A.ID=C.抗生素ID AND C.抗生素分组ID= [1] Order By A.编码"
                    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, CLng(zlCommFun.Nvl(rs("ID"), 0)))
                    mlng抗生素分组id = CLng(zlCommFun.Nvl(rs("ID"), 0))
    
                    If rsTmp.BOF = False Then
                        vsfDetail.TextMatrix(0, 0) = "序号"
                        Call FillGrid(vsfDetail, rsTmp)
                        vsfDetail.TextMatrix(0, 0) = ""
                        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
                    End If
                    
                    mblnChangeEdit = True
                End If
                gintSelectFocus = 5
                vsfDetail.SetFocus
            Else
                ShowSimpleMsg "没有抗生素分组数据！"
            End If
        Case "无药敏检验"
            Call ClearGrid(vsfDetail)

            mblnChangeEdit = True
    End Select
End Sub

Private Sub txtComment_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtComment_GotFocus()
    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 5
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        mblnChangeEdit = True
    End If
End Sub

Private Sub txtResult_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtResult_GotFocus()
    With txtResult
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 4
End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        mblnChangeEdit = True
    End If
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
    If mrsSave.RecordCount > 0 Then
        Call ReadRecord(Row)
    End If
    mblnChangeEdit = True
    '重新排列序号
    RenumVsf Vsf, 0
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case mCol.菌落计数
            If Left(Vsf.TextMatrix(Row, mCol.菌落计数), 1) = "/" Then
                If LoadModel(Mid(Vsf.TextMatrix(Row, mCol.菌落计数), 2)) Then
                    mblnChangeEdit = True
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error GoTo errH
    If OldRow <> NewRow Then
        '行变化后恢复
        Call ClearGrid(vsfDetail)
        
        mrsSave.filter = ""
        mrsSave.filter = "Key=" & Val(Vsf.RowData(NewRow))
        If mrsSave.RecordCount = 0 Then
'            Call LoadDefaultGroup(Val(vsf.RowData(NewRow)))
        Else
'            Call LoadDefaultGroup(Val(vsf.RowData(NewRow)))
            Call ReadRecord(NewRow)
        End If
    End If
    If OldCol <> NewCol And mblnEdit Then
        ShowValue NewCol - 1
    End If
    If Val(Vsf.RowData(NewRow)) = 0 And mblnEdit Then
        Vsf.Col = mCol.细菌名称
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then Cancel = True: Exit Sub
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
    Call DeleteRecord(mrsSave)
    Call ClearGrid(vsfDetail)
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If Not mblnEdit Then Cancel = True: Exit Sub
    If Val(Vsf.RowData(Row)) = 0 Then
        Col = mCol.细菌名称
        Cancel = True
    End If
End Sub

Private Sub Vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow And OldRow < Vsf.Rows Then
        '行变化前保存
        Call WriteRecord(OldRow)
    End If
    If OldCol <> NewCol And mblnEdit Then
        Vsf.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.菌落计数, mCol.培养描述, mCol.细菌名称, mCol.耐药机制
                Vsf.EditMode(NewCol) = 1
            Case Else
                Vsf.EditMode(mCol.菌落计数) = 1
        End Select
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim blnUpdate As Boolean

    
    If Not mblnEdit Then Exit Sub
    
    On Error GoTo errH
    
    If Col = 1 Then
        bytResult = ShowOpenTree(Vsf, 1)
    Else
        bytResult = ShowOpenList(Vsf, "", False, 3)
    End If
    gintSelectFocus = 4: lvwSelect.SetFocus
    Vsf.SetFocus
    
    Select Case bytResult
    Case 0
        '没有匹配的项目
        MsgBox "没有找到相匹配的结果！", vbInformation, gstrSysName
    Case 1
        '选取了一个项目
        If Col = 1 Then
            If MsgBox("你已选择了新的细菌,是否需要清空当前药敏结果?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                blnUpdate = True
                Call ClearGrid(vsfDetail)
            End If
            
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
            If mrsSave.RecordCount = 0 And blnUpdate = True Then
                Call LoadDefaultGroup(Val(Vsf.RowData(Row)))
            End If
            Vsf.Col = mCol.菌落计数: gintSelectFocus = 4: Vsf.EditMode(mCol.细菌名称) = 0: Vsf.EditMode(mCol.菌落计数) = 1
            
            mblnChangeEdit = True
        End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_GotFocus()
'    If mblnEdit Then ShowValue 1
    If mblnEdit Then ShowValue Me.lvwSelect.Tag
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strSvrText As String, intRet As Integer
    
    If KeyCode = vbKeyReturn Then
        
        If InStr(Vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Cancel = True
            Exit Sub
        End If
            
        Select Case Col
            Case mCol.细菌名称
                intRet = ShowOpenList(Vsf, Vsf.EditText, True, 1)
                gintSelectFocus = 4: lvwSelect.SetFocus
                Vsf.SetFocus
                Select Case intRet
                    Case 0
                        '没有匹配的项目
                        Cancel = True
                        Vsf.Cell(flexcpData, Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.EditText = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.TextMatrix(Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                            
                        MsgBox "没有找到相匹配的结果！", vbInformation, gstrSysName
                    Case 1
                        '选取了一个项目
                        Call ClearGrid(vsfDetail)
                                        
                        mrsSave.filter = ""
                        mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
                        If mrsSave.RecordCount = 0 Then Call LoadDefaultGroup(Val(Vsf.RowData(Row)))
                        Vsf.Col = mCol.菌落计数: gintSelectFocus = 4: Vsf.EditMode(mCol.细菌名称) = 0: Vsf.EditMode(mCol.菌落计数) = 1
                        
                        mblnChangeEdit = True
                        Cancel = True
                    Case 2
                        '取消了本次选择
                        Cancel = True
                        Vsf.Cell(flexcpData, Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.TextMatrix(Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                End Select
                
        End Select
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If Vsf.RowData(Row) = 0 And Row > 1 And KeyAscii = vbKeyReturn Then
        Vsf.Row = Row - 1: KeyAscii = 0
        vsfDetail.Col = mCol.结果标志
        vsfDetail.SetFocus
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Exit Sub
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    mblnChangeEdit = True
End Sub

Private Sub Vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim rsTmp As New ADODB.Recordset
    Dim mFindMode As Integer        '按病人ID或病人姓名方式查找 0=病人id 1=病人姓名
    
    
    Debug.Print Button & " " & Now
    If Button <> vbRightButton Then Exit Sub
    
    On Error GoTo errH
    
    mFindMode = zlDatabase.GetPara("历史病人识别", 100, 1208, 0)
    
    gstrSql = "Select 核收时间, A.ID" & vbNewLine & _
              "From 检验标本记录 A, (Select 病人id, 姓名 From 检验标本记录 Where ID = [1]) B" & vbNewLine & _
              "Where a.微生物标本 = 1 and  " & IIf(mFindMode = 0, " A.病人id = B.病人id ", " a.姓名 =b.姓名 ") & vbNewLine & _
              "Order By 核收时间 Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngKey)
    
    If rsTmp.RecordCount > 1 Then
        Set cbrPopupBar = Me.cbrthis.Add("弹出菜单", xtpBarPopup)
        Do Until rsTmp.EOF
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, rsTmp("ID"), "检验时间:" & rsTmp("核收时间"))
            rsTmp.MoveNext
        Loop
        cbrPopupBar.ShowPopup
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then RaiseEvent StartEdit(Cancel)
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lng细菌ID As Long, lng抗生素ID As Long, lngRow As Long
    Dim str药敏方法 As String, str检验结果 As String
    Dim str结果标志 As String
    Dim strSQL As String
                
    If Col = mCol.结果标志 Then
        Select Case UCase(Left(vsfDetail.TextMatrix(Row, mCol.结果标志), 1))
        Case "R"
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.红色
        Case "I"
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.兰色
        Case Else
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.黑色
        End Select
    ElseIf Col = mCol.药敏方法 Then
        With vsfDetail
            For lngRow = .FixedRows To .Rows - 1
                If lngRow <> Row Then
                    If .TextMatrix(Row, Col) <> .TextMatrix(lngRow, Col) Then
                        .TextMatrix(lngRow, Col) = .TextMatrix(Row, Col)
                    End If
                End If
            Next
        End With
    ElseIf Col = mCol.检验结果 Then
        With vsfDetail
            
            lng细菌ID = Val(Vsf.RowData(Vsf.Row))
            lng抗生素ID = Val(.RowData(.Row))
            str药敏方法 = .TextMatrix(.Row, mCol.药敏方法)
            str检验结果 = .TextMatrix(.Row, mCol.检验结果)
            str结果标志 = .TextMatrix(.Row, mCol.结果标志)
            
            .TextMatrix(Row, mCol.结果标志) = Eval结果标志(lng细菌ID, mlng抗生素分组id, lng抗生素ID, str药敏方法, str检验结果)
            
            Select Case UCase(Left(.TextMatrix(.Row, mCol.结果标志), 1))
            Case "R"
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.红色
            Case "I"
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.兰色
            Case Else
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.黑色
            End Select
            
        End With
    ElseIf Col = mCol.抗生素名称 Then
        With vsfDetail
            ShowOpenList vsfDetail, IIf(.EditText <> "", .EditText, .TextMatrix(.Row, mCol.抗生素名称)), True, 4
            
            WriteRecord Me.Vsf.Row
            Me.lvwSelect.SetFocus
            Me.vsfDetail.SetFocus
            gintSelectFocus = 5
        End With
    End If

    mblnChangeEdit = True
End Sub
Private Function Eval结果标志(ByVal lng细菌ID As Long, ByVal lng抗生素分组ID As Long, ByVal lng抗生素ID As Long, ByVal str药敏方法 As String, ByVal str检验结果 As String) As String
    '计算结果标志
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim dblTmp  As Double, varTmp As Variant, intLoop As Integer
    Dim str结果 As String
    
    If Val(str药敏方法) = 0 Then
        Eval结果标志 = ""
        Exit Function
    End If
    
    If Trim(str检验结果) = "" Then
        Eval结果标志 = ""
        Exit Function
    End If
    
    If InStr(Trim(str检验结果), ">") > 0 Or InStr(str检验结果, "〉") > 0 Then
        Eval结果标志 = "R-耐药"
        Exit Function
    End If
    
    If InStr(Trim(str检验结果), "<") > 0 Or InStr(str检验结果, "〈") > 0 Then
        Eval结果标志 = "S-敏感"
        Exit Function
    End If
    
    strSQL = "Select 判断方式, 参考低值, 参考高值, 低值结果,高值结果,中间结果" & vbNewLine & _
            "From 检验细菌抗生素参考" & vbNewLine & _
            "Where 细菌id = [1] And 抗生素分组id = [2] And 抗生素id = [3] And 药敏方法 = [4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng细菌ID, lng抗生素分组ID, lng抗生素ID, Val(str药敏方法))
    Do Until rsTmp.EOF
        
        If InStr(str检验结果, "/") > 0 Then
            varTmp = Split(str检验结果, "/")
            For intLoop = LBound(varTmp) To UBound(varTmp)
                If dblTmp < Val(varTmp(intLoop)) Then
                    dblTmp = Val(varTmp(intLoop))
                End If
            Next
        Else
            dblTmp = Val(str检验结果)
        End If
        
        If rsTmp!判断方式 = 1 Then
            If dblTmp >= Val("" & rsTmp!参考低值) And dblTmp <= Val("" & rsTmp!参考高值) Then
                str结果 = Trim("" & rsTmp!中间结果)
                If str结果 = "" Then str结果 = "I-中介"
                Eval结果标志 = str结果
            ElseIf dblTmp > Val("" & rsTmp!参考高值) Then
                str结果 = Trim("" & rsTmp!高值结果)
                If str结果 = "" Then str结果 = "R-耐药"
                Eval结果标志 = str结果
            ElseIf dblTmp < Val("" & rsTmp!参考低值) Then
                str结果 = Trim("" & rsTmp!低值结果)
                If str结果 = "" Then str结果 = "S-敏感"
                Eval结果标志 = str结果
            End If
        Else
            If dblTmp > Val("" & rsTmp!参考低值) And dblTmp < Val("" & rsTmp!参考高值) Then
                str结果 = Trim("" & rsTmp!中间结果)
                If str结果 = "" Then str结果 = "I-中介"
                Eval结果标志 = str结果
            ElseIf dblTmp >= Val("" & rsTmp!参考高值) Then
                str结果 = Trim("" & rsTmp!高值结果)
                If str结果 = "" Then str结果 = "R-耐药"
                Eval结果标志 = str结果
            ElseIf dblTmp <= Val("" & rsTmp!参考低值) Then
                str结果 = Trim("" & rsTmp!低值结果)
                If str结果 = "" Then str结果 = "S-敏感"
                Eval结果标志 = str结果
            End If
        End If
        
        rsTmp.MoveNext
    Loop
End Function

Private Sub vsfDetail_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnEdit = True Then
        WriteRecord Me.Vsf.Row
    Else
        Cancel = True
    End If
    
End Sub

Private Sub vsfDetail_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.抗生素名称
    Cancel = True
End Sub

Private Sub vsfDetail_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldCol <> NewCol And mblnEdit Then
'        vsfDetail.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.药敏方法, mCol.检验结果, mCol.结果标志, mCol.抗生素名称
                vsfDetail.EditMode(NewCol) = 1
            Case Else
                vsfDetail.EditMode(mCol.结果标志) = 1
        End Select
    End If
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfDetail
        ShowOpenList vsfDetail, .TextMatrix(Row, mCol.抗生素名称), True, 4
        WriteRecord Me.Vsf.Row
    End With
    gintSelectFocus = 5
    
    Me.lvwSelect.SetFocus
    Me.vsfDetail.SetFocus
End Sub

Private Sub vsfDetail_GotFocus()
'    lvwSelect.ListItems.Clear
End Sub

Private Sub vsfDetail_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SelectNextRow(Row, Col)
        Cancel = True  '禁止自定义控件的KeyPress处理
    End If
End Sub
Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SelectNextRow(Row, Col)
        Exit Sub
    End If
    If Chr(KeyAscii) = "'" Then KeyAscii = 0

    mblnChangeEdit = True
End Sub

Private Sub vsfDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then
        RaiseEvent StartEdit(Cancel)
    End If
End Sub

Private Sub SelectNextRow(ByVal Row As Long, ByVal Col As Long)
    '跳转到下一单元格
   
    With vsfDetail
        If .Col + 1 <= mCol.结果标志 Then
            .Body.Select .Row, .Col + 1
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Body.Select .Row + 1, mCol.抗生素名称
            Else
                If Trim(.RowData(.Row)) <> "" Then
                    .Rows = .Rows + 1
                    .Body.Select .Row + 1, mCol.抗生素名称
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowValue(ByVal intType As Integer)
    'intType：1－检验结果、2－培养描述、3－耐药机制、4-评语、5-备注
    Dim rs As ADODB.Recordset
    Dim strSQL As String, strValue As String, aValues() As String, i As Long
    Dim ListItem As ListItem
    
    On Error GoTo errH
    
    Select Case intType
        Case 1
            strSQL = "SELECT ROWNUM AS ID,编码,名称,名称 As 取值 FROM 检验结果描述 A " & _
                " WHERE 分类=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "微生物")
        Case 2
            strSQL = "SELECT Rownum As ID,A.编码,A.简码,A.名称,A.说明 As 取值 FROM 检验培养文字 A"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Case 3
            strSQL = "select Rownum as ID,A.编码,A.简码,A.名称,A.名称 as 取值 from 细菌耐药机制 A"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Case 4
            strSQL = "SELECT Rownum As ID,A.编码,A.简码,A.名称,A.说明 As 取值 FROM 检验评语文字 A " & _
                "WHERE A.分类 Is Null Or A.分类=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "微生物")
        Case 5
            strSQL = "SELECT Rownum As ID,A.编码,A.简码,A.名称,A.说明 As 取值 FROM 检验备注文字 A " & _
                "WHERE A.分类 Is Null Or A.分类=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "微生物")
    End Select
    
    With lvwSelect
        .ListItems.Clear
        .Tag = intType
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                Set ListItem = .ListItems.Add(, "_" & rs("ID"), Nvl(rs("名称")))
                ListItem.SubItems(1) = Nvl(rs("取值"))
                rs.MoveNext
            Loop
        End If
    End With
    
    '取微生物的取值序列
    If intType = 1 Then
        strSQL = "SELECT 取值序列 FROM 检验项目 WHERE 诊治项目ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngItemID)
        If rs.EOF Then
            strValue = "-|±|+|++|+++|++++"
        Else
            strValue = Nvl(rs("取值序列"), "-|±|+|++|+++|++++")
            strValue = Replace(strValue, ";", "|")
        End If
        aValues = Split(strValue, "|")
        With lvwSelect
            For i = 0 To UBound(aValues)
                Set ListItem = .ListItems.Add(, "V" & i, aValues(i))
                ListItem.SubItems(1) = aValues(i)
            Next
        End With
    End If
    Me.lvwSelect.ColumnHeaders(1).Width = Me.lvwSelect.Width
    Me.lvwSelect.ColumnHeaders(2).Width = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadLastValue()
    '功能 写入最后近一次的微生物检验结果
    Dim rsTmp As New ADODB.Recordset
    Dim lngloop As Long
    Dim intFindPatientType As Integer   '0=按ID查找 1=按姓名查找
    Dim strPatientName As String        '病人姓名
    
    On Error GoTo errH
    
    intFindPatientType = zlDatabase.GetPara("历史病人识别", 100, 1208, 0)
    
    If intFindPatientType <> 0 Then
        gstrSql = "select 姓名 from 检验标本记录 where id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
        If rsTmp.EOF = False Then strPatientName = Nvl(rsTmp("姓名"))
    End If
    
    gstrSql = "Select B.细菌id As Key,b.检验结果 as 细菌结果, C.抗生素id As ID,C.结果 As 检验结果," & vbNewLine & _
                "       Decode(C.结果类型, 'R', 'R-耐药', 'I', 'I-中介', 'S', 'S-敏感',C.结果类型) As 结果类型," & vbNewLine & _
                "       Decode(C.药敏方法, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') As 药敏方法" & vbNewLine & _
                "From (Select ID from (Select b.Id  From 检验标本记录 a , 检验标本记录 b" & vbNewLine & _
                "Where " & IIf(intFindPatientType = 0, " b.病人ID = a.病人ID ", _
                                " b.病人ID in (select 病人ID from 病人信息 where 姓名 = [2] )") & vbNewLine & _
                "And b.Id < [1] And a.Id = [1] " & vbNewLine & _
                "Order By b.Id Desc)   ) a ," & vbNewLine & _
                " 检验普通结果　B, 检验药敏结果 C" & vbNewLine & _
                "Where A.ID = B.检验标本id And B.ID = C.细菌结果id(+) and b.细菌id is not null order by a.id desc  "

                    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, strPatientName)
    
    If rsTmp.EOF Or Me.vsfDetail.Rows = 1 Then Exit Sub       '没有记录时退出
    
    With vsfDetail
        For lngloop = 1 To .Rows - 1
            If .RowData(lngloop) <> "" Then
                rsTmp.filter = "ID=" & .RowData(lngloop)
                If rsTmp.EOF = False Then
                    .TextMatrix(lngloop, mCol.上次结果) = Nvl(rsTmp("检验结果"))

                    .TextMatrix(lngloop, mCol.上次标志) = Nvl(rsTmp("结果类型"))
                    Select Case UCase(Left(.TextMatrix(lngloop, mCol.上次标志), 1))
                    Case "R"
                        .Cell(flexcpForeColor, lngloop, mCol.上次结果, lngloop, .Cols - 1) = COLOR.红色
                    Case "I"
                        .Cell(flexcpForeColor, lngloop, mCol.上次结果, lngloop, .Cols - 1) = COLOR.兰色
                    Case Else
                        .Cell(flexcpForeColor, lngloop, mCol.上次结果, lngloop, .Cols - 1) = COLOR.黑色
                    End Select

                End If
            End If
        Next
    End With
    
    With Vsf
        For lngloop = 1 To .Rows - 1
            If .RowData(lngloop) <> "" Then
                rsTmp.filter = "Key=" & .RowData(lngloop)
                If rsTmp.EOF = False Then
                    .TextMatrix(lngloop, mCol.上次菌落计数) = Nvl(rsTmp("细菌结果"))
                End If
            End If
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Resize()
    '供主窗体调用
    Call Form_Resize
End Sub
Private Function LoadModel(ByVal strCode As String) As Boolean
'调入报告模板(微生物项目)
'strCode：模板编码或简码
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngCurrRow As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    On Error GoTo errH
    
    LoadModel = False
    strSQL = "Select B.ID, B.中文名 As 检验项目, A.检验结果 As 本次结果, A.培养描述" & vbNewLine & _
            "From 检验模板内容 A, 检验细菌 B, 检验模板目录 D" & vbNewLine & _
            "Where A.细菌id = B.ID And D.ID = A.模板id And A.细菌id Is Not Null And (D.编码 = [1] Or D.简码 = [1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode, mDeviceID)
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsTmp("ID"))))
            If lngCurrRow > 0 Then
                If Val(Vsf.RowData(lngCurrRow)) = Nvl(rsTmp("ID")) Then
                    Vsf.TextMatrix(lngCurrRow, mCol.菌落计数) = Nvl(rsTmp("本次结果"))
                    Vsf.TextMatrix(lngCurrRow, mCol.培养描述) = Nvl(rsTmp("培养描述"))
                End If
            End If
            rsTmp.MoveNext
        Loop
        LoadModel = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function FindRepeatLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '功能:查找RowData等于strSeekID的行
    '参数:
    '返回:行号或-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer, intCol As Integer
    FindRepeatLine = -1


    For i = 1 To objMsf.Rows - 1
        If Val(Me.Vsf.RowData(i)) = strSeekID Then
            FindRepeatLine = i
            Exit For
        End If

    Next

    If i <= objMsf.Rows - 1 Then FindRepeatLine = i
End Function

