VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugQualityCard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品质量管理"
   ClientHeight    =   5220
   ClientLeft      =   3825
   ClientTop       =   3465
   ClientWidth     =   8160
   Icon            =   "frmDrugQualityCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6960
      TabIndex        =   38
      Top             =   4320
      Width           =   1100
   End
   Begin TabDlg.SSTab sstQuality 
      Height          =   4935
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "短损信息(&D)"
      TabPicture(0)   =   "frmDrugQualityCard.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fra短损信息"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "处理信息(&V)"
      TabPicture(1)   =   "frmDrugQualityCard.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraExecute"
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra短损信息 
         Height          =   4305
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   6195
         Begin VB.TextBox txt原产地 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   4545
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1200
            Width           =   1320
         End
         Begin VB.TextBox txt批号 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            MaxLength       =   11
            TabIndex        =   9
            Top             =   1590
            Width           =   4665
         End
         Begin VB.TextBox txt产地 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   300
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1215
            Width           =   2505
         End
         Begin VB.CommandButton cmdProvider 
            Caption         =   "…"
            Height          =   300
            Left            =   5610
            TabIndex        =   20
            Top             =   3165
            Width           =   270
         End
         Begin VB.TextBox txtProvider 
            Height          =   300
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   19
            Top             =   3145
            Width           =   4425
         End
         Begin VB.ComboBox cbo短损说明 
            Height          =   300
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2760
            Width           =   2025
         End
         Begin VB.TextBox TxtName 
            Height          =   300
            Left            =   1215
            MaxLength       =   30
            TabIndex        =   3
            Top             =   825
            Width           =   3345
         End
         Begin VB.TextBox TxtNumber 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1200
            TabIndex        =   15
            Top             =   2735
            Width           =   1440
         End
         Begin VB.CommandButton CmdDrugSelect 
            Caption         =   "…"
            Height          =   300
            Left            =   4560
            TabIndex        =   4
            Top             =   825
            Width           =   270
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   420
            Width           =   4665
         End
         Begin MSComCtl2.DTPicker dtp短损日期 
            Height          =   285
            Left            =   3840
            TabIndex        =   22
            Top             =   3563
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   170328067
            CurrentDate     =   36489
         End
         Begin VB.Label lbl原产地 
            AutoSize        =   -1  'True
            Caption         =   "原产地"
            Height          =   180
            Left            =   3840
            TabIndex        =   50
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label lbl单位 
            AutoSize        =   -1  'True
            Caption         =   "盒"
            Height          =   180
            Index           =   2
            Left            =   2730
            TabIndex        =   49
            Top             =   2805
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbl单位 
            AutoSize        =   -1  'True
            Caption         =   "/盒"
            Height          =   180
            Index           =   1
            Left            =   2880
            TabIndex        =   48
            Top             =   2370
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label lbl单位 
            AutoSize        =   -1  'True
            Caption         =   "/盒"
            Height          =   180
            Index           =   0
            Left            =   2880
            TabIndex        =   47
            Top             =   2010
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label txtSale 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4200
            TabIndex        =   13
            Top             =   2325
            Width           =   1680
         End
         Begin VB.Label txtCost 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   4200
            TabIndex        =   11
            Top             =   1965
            Width           =   1680
         End
         Begin VB.Label txtSalePrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   12
            Top             =   2325
            Width           =   1680
         End
         Begin VB.Label txtCostPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   10
            Top             =   1965
            Width           =   1680
         End
         Begin VB.Label lblSale 
            AutoSize        =   -1  'True
            Caption         =   "销售金额"
            Height          =   180
            Left            =   3435
            TabIndex        =   46
            Top             =   2370
            Width           =   720
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
            Caption         =   "成本金额"
            Height          =   180
            Left            =   3435
            TabIndex        =   45
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label lblSalePrice 
            AutoSize        =   -1  'True
            Caption         =   "零售价"
            Height          =   180
            Left            =   375
            TabIndex        =   44
            Top             =   2370
            Width           =   540
         End
         Begin VB.Label lblCostPrice 
            AutoSize        =   -1  'True
            Caption         =   "成本价"
            Height          =   180
            Left            =   375
            TabIndex        =   43
            Top             =   2010
            Width           =   540
         End
         Begin VB.Label lbl批号 
            AutoSize        =   -1  'True
            Caption         =   "批号"
            Height          =   180
            Left            =   375
            TabIndex        =   8
            Top             =   1635
            Width           =   360
         End
         Begin VB.Label lbl产地 
            AutoSize        =   -1  'True
            Caption         =   "生产商"
            Height          =   180
            Left            =   375
            TabIndex        =   5
            Top             =   1275
            Width           =   540
         End
         Begin VB.Label txt登记人 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   1200
            TabIndex        =   37
            Top             =   3555
            Width           =   1440
         End
         Begin VB.Label txt单位 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   5385
            TabIndex        =   36
            Top             =   825
            Width           =   495
         End
         Begin VB.Label Lbl计量单位 
            AutoSize        =   -1  'True
            Caption         =   "单位"
            Height          =   180
            Left            =   4950
            TabIndex        =   35
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Lbldate 
            AutoSize        =   -1  'True
            Caption         =   "短损日期"
            Height          =   180
            Left            =   3075
            TabIndex        =   21
            Top             =   3615
            Width           =   720
         End
         Begin VB.Label Lbl药品来源 
            AutoSize        =   -1  'True
            Caption         =   "供药单位"
            Height          =   180
            Left            =   375
            TabIndex        =   18
            Top             =   3210
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "短损药品"
            Height          =   180
            Left            =   375
            TabIndex        =   2
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "登记人"
            Height          =   180
            Left            =   375
            TabIndex        =   34
            Top             =   3615
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "短损说明"
            Height          =   180
            Left            =   3075
            TabIndex        =   16
            Top             =   2805
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "短损数量"
            Height          =   180
            Left            =   375
            TabIndex        =   14
            Top             =   2805
            Width           =   720
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "库房"
            Height          =   180
            Left            =   375
            TabIndex        =   0
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame fraExecute 
         Height          =   3585
         Left            =   -74760
         TabIndex        =   32
         Top             =   840
         Width           =   6195
         Begin VB.ComboBox cbo外调单位 
            Enabled         =   0   'False
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2700
            Visible         =   0   'False
            Width           =   3795
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   2160
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.ComboBox cbo处理人 
            Height          =   300
            Left            =   2175
            TabIndex        =   28
            Top             =   1635
            Width           =   2535
         End
         Begin VB.ComboBox cbo处理办法 
            Height          =   300
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   540
            Width           =   2535
         End
         Begin MSComCtl2.DTPicker dtp处理日期 
            Height          =   285
            Left            =   2175
            TabIndex        =   26
            Top             =   1095
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   170328067
            CurrentDate     =   36489
         End
         Begin VB.Label lbl外调单位 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "外调单位(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1080
            TabIndex        =   42
            Top             =   2760
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "入出类别(&T)"
            Height          =   180
            Left            =   1080
            TabIndex        =   40
            Top             =   2220
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "处理日期(&Q)"
            Height          =   180
            Left            =   1080
            TabIndex        =   25
            Top             =   1140
            Width           =   990
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "处理人(&E)"
            Height          =   180
            Left            =   1260
            TabIndex        =   27
            Top             =   1695
            Width           =   810
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "处理办法(&M)"
            Height          =   180
            Left            =   1080
            TabIndex        =   23
            Top             =   600
            Width           =   990
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   30
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   29
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "frmDrugQualityCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑模式 As Integer             '1:登记;2:修改;3:处理,4:查看
Private mlng记录ID As Long
Private mblnSuccess As Boolean
Private mblnChange As Boolean
Private mfrmMain As Form
Private mblnHaveRecord As Boolean           '对于已经删除的记录，用此变量来判断，如果没有删除，默认为TRUE，否则为FALSE

Private mlng库房ID As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mint审核时减库存 As Integer         '审核时是否同步处理库存（相当与同时实现其他出库功能）：0－不处理库存；1－要同步处理库存
Private mint库存检查 As Integer             '表示药品出库时是否进行库存检查：0-不检查;1-检查，不足提醒；2-检查，不足禁止

'从参数表中取药品价格、数量、金额小数位数（计算精度）
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数
Private mstrPrivs As String                 '操作员权限

Private Sub CheckDependOn()
    '数据依赖检查
    '在审核时同步处理库存的模式下要检查其他入库单据的相关依赖
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If mint审核时减库存 = 0 Then Exit Sub
    
    gstrSQL = "SELECT b.Id,b.名称 " _
            & "FROM 药品单据性质 A, 药品入出类别 B " _
            & "Where A.类别id = B.ID AND A.单据 = 11 "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "药品其他出库入出类别")
    
    If rsTemp.EOF Then
        MsgBox "未设置其他出库的入出类别，不能同步处理库存！", vbExclamation, gstrSysName
        mint审核时减库存 = 0
        Exit Sub
    End If
    
    With cboType
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    lblType.Visible = True
    cboType.Visible = True
    lbl外调单位.Visible = True
    cbo外调单位.Visible = True
    cbo外调单位.Enabled = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function GetProviderNameById(ByVal lngProviderId As Long) As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 名称 From 供应商 Where ID = [1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取供应商名称", lngProviderId)
    
    If Not rsTemp.EOF Then
        GetProviderNameById = rsTemp!名称
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCard(ByVal FrmMain As Form, ByVal int编辑模式 As Integer, _
        ByVal lng记录id As Long, Optional strPrivs As String) As Boolean
    Dim rsParrel As New Recordset
    
    mblnSuccess = False
    mblnChange = False
    mint编辑模式 = int编辑模式
    mblnHaveRecord = True
    mlng记录ID = lng记录id
    mstrPrivs = strPrivs
    
    On Error GoTo errHandle
    If int编辑模式 > 1 Then
        gstrSQL = "select nvl(处理人,'0') from 药品质量记录 where id=[1]"
        Set rsParrel = zlDataBase.OpenSQLRecord(gstrSQL, "[读取处理人]", lng记录id)
        
        If rsParrel.EOF Then
            MsgBox "该药品质量记录已被其他人删除，请检查！", vbOKOnly, gstrSysName
            Exit Function
        ElseIf rsParrel.Fields(0) <> "0" And InStr(1, 23, mint编辑模式) <> 0 Then
            MsgBox "该药品质量记录已被其他人处理，请检查！", vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    Set mfrmMain = FrmMain
    
    Select Case int编辑模式
        Case 1, 2
            sstQuality.TabEnabled(1) = False
            fraExecute.Enabled = False
            sstQuality.Tab = 0
            Me.Caption = "药品短损登记"
        Case 3
            sstQuality.Tab = 1
            Me.Caption = "药品短损处理"
        Case 4
            Fra短损信息.Enabled = False
            fraExecute.Enabled = False
            Me.Caption = "药品短损查看"
            cmdOk.Enabled = False
    End Select
    mblnChange = False
    Me.Show vbModal, FrmMain
    ShowCard = mblnSuccess
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function VerifyData() As Boolean
    VerifyData = False
    
    If Val(TxtName.Tag) = 0 Then
        MsgBox "短损药品必须设置", vbInformation, gstrSysName
        Me.TxtName.SetFocus
        Exit Function
    End If
    If TxtNumber = "" Then
        MsgBox "短损数量应该输入!", vbInformation, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If Val(TxtNumber) = 0 Then
        MsgBox "短损数量必须大于零!", vbInformation, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If Val(Me.TxtNumber) >= 10 ^ 11 - 1 Then
        MsgBox "数量必须大于0小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
        Me.TxtNumber.SetFocus
        Exit Function
    End If
    
    If mint编辑模式 = 3 Then
        If cboType.Text = "药品外调" Then
            If cbo外调单位.Text = "" Then
                MsgBox "外调单位不能为空，请输入外调单位!", vbInformation, gstrSysName
                cbo外调单位.SetFocus
                Exit Function
            End If
        End If
        
        If cboType.Text = "药品外销" Then
            If cbo外调单位.Text = "" Then
                MsgBox "外销单位不能为空，请输入外销单位!", vbInformation, gstrSysName
                cbo外调单位.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If mint编辑模式 = 3 Then
        If Val(txtProvider.Tag) = 0 Then
            MsgBox "供药单位必须设置", vbInformation, "提示"
            Me.sstQuality.Tab = 0
            Me.txtProvider.SetFocus
            Exit Function
        End If
        If cbo处理人.Text = "" Then
            MsgBox "处理人必须设置", vbInformation, "提示"
            Me.sstQuality.Tab = 1
            Me.cbo处理人.SetFocus
            Exit Function
        End If
        
    End If
        
    VerifyData = True
End Function

Private Sub cboStock_Click()
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    Dim rsList As ADODB.Recordset
    
    On Error GoTo errHandle
    
    str库房性质 = ""
    gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str库房性质 = str库房性质 & "," & rsList!工作性质
        rsList.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
    If bln中药库房 Then
        lbl原产地.Visible = True
        txt原产地.Visible = True
        txt产地.Width = 2505
    Else
        lbl原产地.Visible = False
        txt原产地.Visible = False
        txt产地.Width = 4665
    End If

    mlng库房ID = 0
    If mlng库房ID <> cboStock.ItemData(cboStock.ListIndex) Then
        mlng库房ID = cboStock.ItemData(cboStock.ListIndex)
        Call GetDrugDigit(mlng库房ID, "药品质量管理", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        mint库存检查 = MediWork_GetCheckStockRule(mlng库房ID)
        Call ReleaseSelectorRS
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cboType_click()
    cbo外调单位.Clear
    If cboType.Text = "药品外调" Or cboType.Text = "药品外销" Then
        cbo外调单位.Enabled = True
    End If
    
    If cboType.Text = "药品外销" Then
        lbl外调单位.Caption = "外销单位"
    End If
    
    If cboType.Text = "药品其他出库" Then
        cbo外调单位.Enabled = False
    End If
End Sub

Private Sub cbo处理办法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cbo处理人_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    Dim i As Integer, intIdx As Integer
    Dim strText As String
    Dim rsdepart As New Recordset
    
    On Error GoTo errHandle
    With cbo处理人
        strText = .Text
        If strText = "" Then
            .ListIndex = -1
        Else
            intIdx = -1
            For i = 0 To .ListCount - 1
                If InStr(.List(i), UCase(strText)) > 0 Then
                    If intIdx = -1 Then .ListIndex = i
                    intIdx = i
                End If
            Next
            If intIdx = -1 Then
                gstrSQL = "Select id,姓名 from 人员表 " & _
                          "Where (站点 = [2] Or 站点 is Null) And (简码 like [1] or 编号 like [1]) " & _
                          "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
                Set rsdepart = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[读取处理人]", UCase(strText) & "%", gstrNodeNo)
                
                If Not rsdepart.EOF Then
                    Do While Not rsdepart.EOF
                        For i = 0 To .ListCount - 1
                            If InStr(.List(i), rsdepart.Fields(1)) > 0 Then
                                If intIdx = -1 Then .ListIndex = i
                                intIdx = i
                            End If
                        Next
                        rsdepart.MoveNext
                    Loop
                End If
            End If
        End If
        If Trim(.Text) = "" Then
            MsgBox "对不起，必须输入一个处理人!", vbExclamation + vbOKOnly, gstrSysName
            .SetFocus
            Exit Sub
        End If
        
        If .ListIndex = -1 Then
            MsgBox "对不起，没有找到你输入的人员，请重输！", vbExclamation + vbOKOnly, gstrSysName
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        Else
            If intIdx <> .ListIndex Then SendKeys "{F4}": Exit Sub
        End If
    End With
    OS.PressKey (vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo短损说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub cbo外调单位_DropDown()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If cboType.Text = "药品外调" Then
        gstrSQL = "Select 编码||'-'||名称 AS 外调单位 From 药品外调单位 Order By 编码"
    ElseIf cboType.Text = "药品外销" Then
        gstrSQL = "Select 编码||'-'||名称 AS 外调单位 From 药品外销单位 Order By 编码"
    End If
    
    'Call zldatabase.OpenRecordset(rsTemp, gstrSQL, "读取外调单位")
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-读取外调单位")
    With cbo外调单位
        .Clear
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!外调单位
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDrugSelect_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "药品质量管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
'    Set RecReturn = Frm药品选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), False)
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , False, , , , False, mstrPrivs)
    
    If RecReturn.RecordCount > 0 Then
        TxtName.Tag = RecReturn!药品id
        If gint药品名称显示 = 1 Then
            TxtName.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            TxtName.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txt单位 = Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位)
        txt单位.Tag = Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装)
        txt产地.Text = IIf(IsNull(RecReturn!产地), "", RecReturn!产地)
        txt原产地.Text = IIf(IsNull(RecReturn!原产地), "", RecReturn!原产地)
        txt批号.Text = IIf(IsNull(RecReturn!批号), "", RecReturn!批号)
        txt批号.Tag = IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)
        txtProvider.Tag = RecReturn!上次供应商ID
        
'        If IsNull(RecReturn!成本价) Then
'            txtCostPrice.Caption = ""
'            txtCost.Caption = ""
'        Else
'            txtCostPrice.Caption = zlStr.FormatEx(RecReturn!成本价, mintCostDigit, , True)
'            txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber) * Val(txt单位.Tag), mintMoneyDigit, , True)
'        End If
'        If IsNull(RecReturn!售价) Then
'            txtSalePrice.Caption = ""
'            txtSale.Caption = ""
'        Else
'            txtSalePrice.Caption = zlStr.FormatEx(RecReturn!售价, mintPriceDigit, , True)
'            txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber) * Val(txt单位.Tag), mintMoneyDigit, , True)
'        End If
        
        lbl单位(0).Visible = True
        lbl单位(1).Visible = True
        lbl单位(2).Visible = True
        lbl单位(0).Caption = "/" & txt单位.Caption
        lbl单位(1).Caption = "/" & txt单位.Caption
        lbl单位(2).Caption = txt单位.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(Get成本价(RecReturn!药品id, mlng库房ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)), gtype_UserDrugDigits.Digit_成本价, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt单位.Tag), mintCostDigit, , True)
        txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If RecReturn!时价 = 1 Then
            txtSalePrice.Tag = zlStr.FormatEx(Get零售价(RecReturn!药品id, mlng库房ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), 1), gtype_UserDrugDigits.Digit_零售价, , True)
        Else
            txtSalePrice.Tag = zlStr.FormatEx(RecReturn!售价, gtype_UserDrugDigits.Digit_零售价, , True)
        End If
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt单位.Tag), mintPriceDigit, , True)
        txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        
        If Val(txtProvider.Tag) <> 0 Then
            txtProvider.Text = GetProviderNameById(Val(txtProvider.Tag))
        End If
        
        TxtNumber.SetFocus
    End If
End Sub


Private Function SaveCard() As Boolean
    Dim dblTmp As Double
    On Error GoTo errHandle
    SaveCard = False
    
    If mint编辑模式 = 2 Then
        gstrSQL = "zl_药品质量管理_delete(" & mlng记录ID & ")"
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gstrSQL = "zl_药品质量管理_INSERT("
    '库房ID
    gstrSQL = gstrSQL & cboStock.ItemData(cboStock.ListIndex)
    '药品ID
    gstrSQL = gstrSQL & "," & Val(TxtName.Tag)
    '毁损原因
    gstrSQL = gstrSQL & ",'" & cbo短损说明.Text & "'"
    '毁损数量
    gstrSQL = gstrSQL & "," & FormatEx(Val(TxtNumber.Text) * Val(txt单位.Tag), mintNumberDigit)
    '登记人
    gstrSQL = gstrSQL & ",'" & txt登记人 & "'"
    '登记时间
    gstrSQL = gstrSQL & ",to_date('" & Format(dtp短损日期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')"
    '产地
    gstrSQL = gstrSQL & ",'" & txt产地.Text & "'"
    '批号
    gstrSQL = gstrSQL & ",'" & txt批号.Text & "'"
    '批次
    gstrSQL = gstrSQL & "," & Val(txt批号.Tag)
    '供药单位ID
    gstrSQL = gstrSQL & "," & IIf(Val(txtProvider.Tag) = 0, "NULL", txtProvider.Tag)
    '成本单价
    gstrSQL = gstrSQL & "," & IIf(Val(txtCostPrice.Tag) = 0, "null", Val(txtCostPrice.Tag))
    '成本金额
    dblTmp = zlStr.FormatEx(Val(txtCost.Caption), mintMoneyDigit, , True)
    gstrSQL = gstrSQL & "," & IIf(dblTmp = 0, "null", dblTmp)
    '销售单价
    gstrSQL = gstrSQL & "," & IIf(Val(txtSalePrice.Tag) = 0, "null", Val(txtSalePrice.Tag))
    '销售金额
    dblTmp = zlStr.FormatEx(Val(txtSale.Caption), mintMoneyDigit, , True)
    gstrSQL = gstrSQL & "," & IIf(dblTmp = 0, "null", dblTmp)
    '说明
    gstrSQL = gstrSQL & ",NULL"
    gstrSQL = gstrSQL & ")"

    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
    mblnSuccess = True
    mblnChange = False
    SaveCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveHandle() As Boolean
    Dim lng入出类别id As Long
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim dblOutPrice As Double   '外调价
    Dim strOutUnit As String    '外调单位
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str批准文号 As String
    Dim blnTran As Boolean
    
    Dim rsTemp As New Recordset
    
    On Error GoTo errHandle
    SaveHandle = False
    
    If mint审核时减库存 = 1 Then
        lng入出类别id = cboType.ItemData(cboType.ListIndex)
        chrNo = Sys.GetNextNo(28, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        lngSerial = 1
        lngStockid = mlng库房ID
        lngDrugID = Val(TxtName.Tag)
        lngBatchID = Val(txt批号.Tag)
        
'        dblQuantity = zlStr.FormatEx(Val(TxtNumber.Text) * Val(txt单位.Tag), gtype_UserSaleDigits.Digit_数量)
        dblQuantity = Val(TxtNumber.Tag)    '处理时用原始数量
        
        '库存检查
        If CheckDrugStock(lngStockid, lngDrugID, lngBatchID, dblQuantity, Val(txt单位.Tag)) = False Then
            Exit Function
        End If
        
        gstrSQL = "Select Nvl(A.实际数量,0) 实际数量, Nvl(A.实际金额,0) 实际金额, Nvl(A.实际差价,0) 实际差价, A.效期, A.批准文号, " & _
            " Nvl(B.是否变价, 0) 是否变价, C.现价, Nvl(D.管理费比例, 0) 比例,Nvl(A.批次,0) As 批次,Nvl(A.零售价,0) As 零售价 " & _
            " From 药品库存 A, 收费项目目录 B, 收费价目 C, 药品规格 D " & _
            " Where A.药品id = B.ID And A.药品id = C.收费细目id And A.药品id = D.药品id And A.性质 = 1 And " & _
            " (C.终止日期 Is Null Or Sysdate Between C.执行日期 And Nvl(C.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) And " & _
            " A.库房id = [1] And A.药品id = [2] And Nvl(A.批次, 0) = [3]" & _
            GetPriceClassString("C")
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取价格信息", lngStockid, lngDrugID, lngBatchID)
        
        If rsTemp.RecordCount = 0 Then
            gstrSQL = "Select Nvl(是否变价, 0) 是否变价, Nvl(d.管理费比例, 0) 比例, d.上次批准文号 批准文号, '' 效期" & vbNewLine & _
                "From 收费项目目录 B, 药品规格 D" & vbNewLine & _
                "Where ID = [2] And b.Id = d.药品id"
            
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "取价格信息", lngStockid, lngDrugID, lngBatchID)
        End If
        
'        If rsTemp!是否变价 = 0 Then
'            dblSalePrice = rsTemp!现价
'        Else
'            dblSalePrice = IIf(rsTemp!批次 = 0, rsTemp!实际金额 / rsTemp!实际数量, IIf(rsTemp!零售价 = 0, rsTemp!实际金额 / rsTemp!实际数量, rsTemp!零售价))
'        End If
        
        dblSalePrice = Get售价(rsTemp!是否变价 = 1, lngDrugID, lngStockid, lngBatchID)
        dblSaleMoney = zlStr.FormatEx(dblSalePrice * dblQuantity, mintMoneyDigit)
        
        dblPurchasePrice = Get成本价(lngDrugID, lngStockid, lngBatchID)
        dblPurchaseMoney = zlStr.FormatEx(dblPurchasePrice * dblQuantity, mintMoneyDigit)
        
        dblMistakePrice = zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit)
        
        If cboType.Text = "药品外调" Then
            dblOutPrice = zlStr.FormatEx((1 + rsTemp!比例 / 100) * dblPurchasePrice, gtype_UserSaleDigits.Digit_成本价)
            If Not cbo外调单位.Text = "" Then
                strOutUnit = Mid(cbo外调单位.Text, 1, InStr(1, cbo外调单位.Text, "-") - 1)
            End If
        End If
        
        strBooker = UserInfo.用户姓名
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strProducingArea = IIf(IsNull(txt产地.Text), "", txt产地.Text)
        strBatchNo = IIf(IsNull(txt批号.Text), "", txt批号.Text)
        datTimeLimit = IIf(IsNull(rsTemp!效期), "", rsTemp!效期)
        strBrief = cbo处理办法.Text & "(质量管理自动减库存)"
        str批准文号 = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
        
        rsTemp.Close
        
        '零差价管理：检查是否存在不满足零差价的药品
        If gtype_UserSysParms.P275_零差价管理模式 = 2 Then
            If IsPriceAdjustMod(lngDrugID) = True Then
                If CheckPriceAdjust(lngDrugID, lngStockid, lngBatchID) = False Then
                    MsgBox "该药品已启用零差价管理，但库存记录中售价和成本价不一致，不能进行业务，请检查！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        gcnOracle.BeginTrans
        blnTran = True
    End If
    
    gstrSQL = "zl_药品质量管理_UPDATE("
    '记录ID
    gstrSQL = gstrSQL & mlng记录ID
    '供药单位ID
    gstrSQL = gstrSQL & "," & Val(txtProvider.Tag)
    '解决办法
    gstrSQL = gstrSQL & ",'" & cbo处理办法.Text & "'"
    '处理人
    gstrSQL = gstrSQL & ",'" & cbo处理人.Text & "'"
    '处理时间
    gstrSQL = gstrSQL & ",to_date('" & Format(dtp处理日期.Value, "yyyy-mm-dd") & "','yyyy-mm-dd')"
'    '成本单价
'    gstrSQL = gstrSQL & "," & txtCostPrice.Caption
'    '成本金额
'    gstrSQL = gstrSQL & "," & Val(txtCostPrice.Caption) * Val(TxtNumber.Text) * Val(txt单位.Tag)
'    '销售单价
'    gstrSQL = gstrSQL & "," & txtSalePrice.Caption
'    '销售金额
'    gstrSQL = gstrSQL & "," & Val(txtSalePrice.Caption) * Val(TxtNumber.Text) * Val(txt单位.Tag)
    '出库单NO
    gstrSQL = gstrSQL & "," & IIf(chrNo = "", "Null", "'" & chrNo & "'")
    gstrSQL = gstrSQL & ")"

    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
   
    If mint审核时减库存 = 1 Then
        gstrSQL = "zl_药品其他出库_INSERT("
        '入出类别ID
        gstrSQL = gstrSQL & lng入出类别id
        'NO
        gstrSQL = gstrSQL & ",'" & chrNo & "'"
        '序号
        gstrSQL = gstrSQL & "," & lngSerial
        '库房ID
        gstrSQL = gstrSQL & "," & lngStockid
        '药品ID
        gstrSQL = gstrSQL & "," & lngDrugID
        '批次
        gstrSQL = gstrSQL & "," & lngBatchID
        '填写数量
        gstrSQL = gstrSQL & "," & dblQuantity
        '成本价
        gstrSQL = gstrSQL & "," & dblPurchasePrice
        '成本金额
        gstrSQL = gstrSQL & "," & dblPurchaseMoney
        '售价
        gstrSQL = gstrSQL & "," & dblSalePrice
        '售价金额
        gstrSQL = gstrSQL & "," & dblSaleMoney
        '差价
        gstrSQL = gstrSQL & "," & dblMistakePrice
        '外调价
        gstrSQL = gstrSQL & "," & dblOutPrice
        '外调单位
        gstrSQL = gstrSQL & ",'" & strOutUnit & "'"
        '填制人
        gstrSQL = gstrSQL & ",'" & strBooker & "'"
        '填制日期
        gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
        '产地
        gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
        '批号
        gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
        '效期
        gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
        '摘要
        gstrSQL = gstrSQL & ",'" & strBrief & "'"
        '批准文号
        gstrSQL = gstrSQL & ",'" & str批准文号 & "'"
        '增值税率
        gstrSQL = gstrSQL & ",''"
        '原产地
        gstrSQL = gstrSQL & ",''"
        '修改人
        gstrSQL = gstrSQL & ",''"
        '修改日期
        gstrSQL = gstrSQL & ",''"
        gstrSQL = gstrSQL & ")"

        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        gstrSQL = "zl_药品其他出库_Verify("
        '序号
        gstrSQL = gstrSQL & lngSerial
        'NO
        gstrSQL = gstrSQL & ",'" & chrNo & "'"
        '库房ID
        gstrSQL = gstrSQL & "," & lngStockid
        '药品ID
        gstrSQL = gstrSQL & "," & lngDrugID
        '批次
        gstrSQL = gstrSQL & "," & lngBatchID
        '实际数量
        gstrSQL = gstrSQL & "," & dblQuantity
        '成本价
        gstrSQL = gstrSQL & "," & dblPurchasePrice
        '成本金额
        gstrSQL = gstrSQL & "," & dblPurchaseMoney
        '零售金额
        gstrSQL = gstrSQL & "," & dblSaleMoney
        '差价
        gstrSQL = gstrSQL & "," & dblMistakePrice
        '审核人
        gstrSQL = gstrSQL & ",'" & strBooker & "'"
        '审核日期
        gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
        gstrSQL = gstrSQL & ")"
 
        Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        gcnOracle.CommitTrans
        blnTran = False
    End If
    
    mblnSuccess = True
    mblnChange = False
    SaveHandle = True
    Exit Function
errHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckDrugStock(ByVal lng库房ID As Long, ByVal lng药品ID As Long, ByVal lng批次 As Long, ByVal Dbl药品数量 As Double, ByVal dbl包装系数 As Double) As Boolean
    Dim blnMsg As Boolean
    Dim rsCheck As New ADODB.Recordset
    Dim Dbl数量 As Double
    
    On Error GoTo errHandle
    
    If mint库存检查 = 0 Then CheckDrugStock = True: Exit Function
       
    gstrSQL = "Select Nvl(可用数量,0) 可用数量,Nvl(实际数量,0) 实际数量 " & _
              "From 药品库存 Where 库房ID=[1] And Nvl(批次,0)=[3] And 性质=1 And 药品ID=[2] "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查库存]", lng库房ID, lng药品ID, lng批次)
    
    If mint编辑模式 = 1 Or mint编辑模式 = 2 Then
        '填单和修改时用可用数量检查
        Dbl数量 = rsCheck!可用数量
    ElseIf mint编辑模式 = 3 Then
        '审核时取实际数量来检查
        Dbl数量 = rsCheck!实际数量
    End If
       
    If Dbl药品数量 > Dbl数量 Then
        blnMsg = True
    End If
    
    If blnMsg Then
        If mint库存检查 = 1 Then        '不足提醒
            If MsgBox("出库数量大于现有的库存数量" & Dbl数量 / dbl包装系数 & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else                            '不足禁止
            MsgBox "出库数量大于现有的库存数量" & Dbl数量 / dbl包装系数 & "，不能出库！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    CheckDrugStock = True

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    
    If Not VerifyData Then Exit Sub
    
    Select Case mint编辑模式
        Case 1
            mblnSuccess = SaveCard
            If mblnSuccess = True Then
                TxtName.Text = ""
                TxtName.Tag = ""
                txt单位 = ""
                TxtNumber.Text = ""
                txtProvider.Text = ""
                txtProvider.Tag = ""
                If cboStock.Enabled = True Then
                    cboStock.SetFocus
                Else
                    TxtName.SetFocus
                End If
            End If
                
        Case 2
            mblnSuccess = SaveCard
            If mblnSuccess = True Then
                Unload Me
                Exit Sub
            End If
            
            
        Case 3
            mblnSuccess = SaveHandle
            If mblnSuccess = True Then
                Unload Me
                Exit Sub
            End If
    End Select
    
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' " & _
              "  And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
'    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-供应商", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "供应商", True, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    
    Me.txtProvider.Tag = rsProvider!id
    Me.txtProvider = rsProvider!名称
    
    If mint编辑模式 = 3 Then
        cmdOk.SetFocus
    Else
        dtp短损日期.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtp处理日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub dtp短损日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOk.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    mblnChange = True
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    With cbo短损说明
        .Clear
        gstrSQL = "select 名称 from 毁损发生原因 "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-毁损发生原因")
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        If .ListCount > 1 Then
            .ListIndex = 0
        End If
        
    End With
    
    With cbo处理办法
        gstrSQL = "select 名称 from 毁损解决办法"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-毁损解决办法")
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        If .ListCount > 1 Then
            .ListIndex = 0
        End If
    End With
    
    With cbo处理人
'        gstrSQL = "select id, 姓名 from 人员表 " & _
'                  "Where (站点 = [1] Or 站点 is Null) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        gstrSQL = "Select distinct a.Id, a.姓名 " & vbNewLine & _
                  "From 人员表 A, 部门人员 B, 部门人员 C " & vbNewLine & _
                  "Where a.Id = b.人员id And b.部门id = c.部门id And c.人员id = [2] And (a.站点 = [1] Or a.站点 Is Null) And " & vbNewLine & _
                  "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & vbNewLine & _
                  "Order By a.姓名 "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-人员信息", gstrNodeNo, UserInfo.用户ID)
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!姓名
            .ItemData(.NewIndex) = rsTmp!id
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
        .Text = UserInfo.用户姓名
        
    End With
    
    With mfrmMain.cboStock
        cboStock.Clear
        
        For intIndex = 0 To .ListCount - 1
            cboStock.AddItem .List(intIndex)
            cboStock.ItemData(cboStock.NewIndex) = .ItemData(intIndex)
        Next
        cboStock.ListIndex = .ListIndex
        cboStock.Enabled = .Enabled
    End With
    
    mlng库房ID = cboStock.ItemData(cboStock.ListIndex)
    Call GetDrugDigit(mlng库房ID, "药品质量管理", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    dtp处理日期.Value = Format(Sys.Currentdate, "yyyy-mm-dd")
    dtp短损日期.Value = Format(Sys.Currentdate, "yyyy-mm-dd")
    txt登记人 = UserInfo.用户姓名
    
    mint审核时减库存 = Val(zlDataBase.GetPara("审核时减少库存", glngSys, 模块号.质量管理))
    Call CheckDependOn
    
    If mint编辑模式 > 1 Then
        initCard
    End If
    
    
    If cboType.Text = "药品其他出库" Then
        cbo外调单位.Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim intIndex As Integer
    Dim intBit As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strsql As String
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    Dim rsList As ADODB.Recordset
    
    intBit = IIf(gint药品名称显示 = 2, 1, 0)
    '考虑单价、金额、数量的精度，额外取数据
    On Error GoTo errHandle
    str库房性质 = ""
    gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsList.EOF
        str库房性质 = str库房性质 & "," & rsList!工作性质
        rsList.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True 'bln中药库房为真时，8列及其之后的列号都+1
    
    strsql = "select 成本单价, 成本金额, 销售单价, 销售金额, 毁损数量 " _
           & "from 药品质量记录 where id=[1]"
    Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, mlng记录ID)
    If rsTmp.EOF Then Exit Sub
    With frmDrugQualityList.vsfList
        Me.Tag = .TextMatrix(.Row, 0)
        TxtName.Text = .TextMatrix(.Row, 1)
        TxtName.Tag = .TextMatrix(.Row, 2 + intBit)
        txt批号.Tag = IIf(IsNull(.TextMatrix(.Row, 3 + intBit)), 0, .TextMatrix(.Row, 3 + intBit))
        txtProvider.Tag = IIf(IsNull(.TextMatrix(.Row, 4 + intBit)), 0, .TextMatrix(.Row, 4 + intBit))
        txtProvider.Text = .TextMatrix(.Row, 16 + intBit + IIf(bln中药库房, 1, 0))
        txt产地.Text = IIf(IsNull(.TextMatrix(.Row, 6 + intBit)), "", .TextMatrix(.Row, 6 + intBit))
        
        txt批号.Text = IIf(IsNull(.TextMatrix(.Row, 7 + intBit)), "", .TextMatrix(.Row, 7 + intBit))
    
        If bln中药库房 Then intBit = intBit + 1: txt原产地.Text = IIf(IsNull(.TextMatrix(.Row, 6 + intBit)), "", .TextMatrix(.Row, 6 + intBit))
    
        txt单位 = .TextMatrix(.Row, 12 + intBit)
        TxtNumber.Text = zlStr.FormatEx(rsTmp!毁损数量 / .TextMatrix(.Row, 14 + intBit), IIf(mint编辑模式 = 4, mintNumberDigit, mintNumberDigit), , True)
        TxtNumber.Tag = Nvl(rsTmp!毁损数量, 0)  '记录原始数量
        txt单位.Tag = .TextMatrix(.Row, 14 + intBit)        '比例系数
        
        lbl单位(0).Visible = True
        lbl单位(1).Visible = True
        lbl单位(2).Visible = True
        lbl单位(0).Caption = "/" & txt单位.Caption
        lbl单位(1).Caption = "/" & txt单位.Caption
        lbl单位(2).Caption = txt单位.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(rsTmp!成本单价, gtype_UserDrugDigits.Digit_成本价, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt单位.Tag), mintCostDigit, , True)
        If IsNull(rsTmp!成本金额) Then
            txtCost.Caption = ""
        Else
            txtCost.Caption = zlStr.FormatEx(rsTmp!成本金额, mintMoneyDigit, , True)
        End If
        
        txtSalePrice.Tag = zlStr.FormatEx(rsTmp!销售单价, gtype_UserDrugDigits.Digit_零售价, , True)
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt单位.Tag), mintPriceDigit, , True)
        If IsNull(rsTmp!销售金额) Then
            txtSale.Caption = ""
        Else
            txtSale.Caption = zlStr.FormatEx(rsTmp!销售金额, mintMoneyDigit, , True)
        End If
        
        For intIndex = 0 To cbo短损说明.ListCount - 1
            If cbo短损说明.List(intIndex) = .TextMatrix(.Row, 15 + intBit) Then
                cbo短损说明.ListIndex = intIndex
                Exit For
            End If
        Next
        
        txt登记人 = .TextMatrix(.Row, 17 + intBit)
        dtp短损日期.Value = .TextMatrix(.Row, 18 + intBit)
        If IIf(IsNull(.TextMatrix(.Row, 19 + intBit)), "", .TextMatrix(.Row, 19 + intBit)) <> "" Then
            For intIndex = 0 To cbo处理办法.ListCount - 1
                If cbo处理办法.List(intIndex) = .TextMatrix(.Row, 19 + intBit) Then
                    cbo处理办法.ListIndex = intIndex
                    Exit For
                End If
            Next
            cbo处理人.Text = .TextMatrix(.Row, 20 + intBit)
            dtp处理日期.Value = .TextMatrix(.Row, 21 + intBit)
        End If
        
    End With
    
    If mint编辑模式 = 3 Then
        TxtName.Enabled = False
        CmdDrugSelect.Enabled = False
        txt产地.Enabled = False
        txt原产地.Enabled = False
        
        txt批号.Enabled = False
        TxtNumber.Enabled = False
        cbo短损说明.Enabled = False
        dtp短损日期.Enabled = False
        cboStock.Enabled = False
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If mint编辑模式 = 4 Then Call ReleaseSelectorRS: Exit Sub
    If mblnChange = True Then
        If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Call ReleaseSelectorRS
End Sub


Private Sub txtName_GotFocus()
    Me.TxtName.SelStart = 0
    Me.TxtName.SelLength = Len(Me.TxtName.Text)
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Dim strTmp As String
    
    Me.TxtName.Text = Trim(Me.TxtName.Text)
    If Len(LTrim(RTrim(TxtName))) = 0 Then Exit Sub
    strTmp = UCase(TxtName)
    Dim RecReturn As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
            
    If Mid(strTmp, 1, 1) = "[" Then
        If InStr(2, strTmp, "]") <> 0 Then
            strTmp = Mid(strTmp, 2, InStr(2, strTmp, "]") - 2)
        Else
            strTmp = Mid(strTmp, 2)
        End If
    End If
        
    sngLeft = Me.Left + sstQuality.Left + Fra短损信息.Left + TxtName.Left
    sngTop = Me.Top + Me.Height - Me.ScaleHeight + sstQuality.Top + Fra短损信息.Top + TxtName.Top + TxtName.Height
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - TxtName.Height - 4530
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "药品质量管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strTmp, sngLeft, sngTop, False)
    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strTmp, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , False, , , , False, mstrPrivs)
    
    If RecReturn.RecordCount = 1 Then
        TxtName.Tag = RecReturn!药品id
        If gint药品名称显示 = 1 Then
            TxtName.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            TxtName.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txt单位 = Choose(mintUnit, RecReturn!售价单位, RecReturn!门诊单位, RecReturn!住院单位, RecReturn!药库单位)
        txt单位.Tag = Choose(mintUnit, 1, RecReturn!门诊包装, RecReturn!住院包装, RecReturn!药库包装)
        txt产地.Text = IIf(IsNull(RecReturn!产地), "", RecReturn!产地)
        txt原产地.Text = IIf(IsNull(RecReturn!原产地), "", RecReturn!原产地)
        txt批号.Text = IIf(IsNull(RecReturn!批号), "", RecReturn!批号)
        txt批号.Tag = IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)
        txtProvider.Tag = RecReturn!上次供应商ID
'        If IsNull(RecReturn!成本价) Then
'            txtCostPrice.Caption = ""
'            txtCost.Caption = ""
'        Else
'            txtCostPrice.Caption = zlStr.FormatEx(RecReturn!成本价, mintCostDigit, , True)
'            txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber) * Val(txt单位.Tag), mintMoneyDigit, , True)
'        End If
'        If IsNull(RecReturn!售价) Then
'            txtSalePrice.Caption = ""
'            txtSale.Caption = ""
'        Else
'            txtSalePrice.Caption = zlStr.FormatEx(RecReturn!售价, mintPriceDigit, , True)
'            txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber) * Val(txt单位.Tag), mintMoneyDigit, , True)
'        End If
        
        lbl单位(0).Visible = True
        lbl单位(1).Visible = True
        lbl单位(2).Visible = True
        lbl单位(0).Caption = "/" & txt单位.Caption
        lbl单位(1).Caption = "/" & txt单位.Caption
        lbl单位(2).Caption = txt单位.Caption
        
        txtCostPrice.Tag = zlStr.FormatEx(Get成本价(RecReturn!药品id, mlng库房ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次)), gtype_UserDrugDigits.Digit_成本价, , True)
        txtCostPrice.Caption = zlStr.FormatEx(Val(txtCostPrice.Tag) * Val(txt单位.Tag), mintCostDigit, , True)
        txtCost.Caption = zlStr.FormatEx(Val(txtCostPrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If RecReturn!时价 = 1 Then
            txtSalePrice.Tag = zlStr.FormatEx(Get零售价(RecReturn!药品id, mlng库房ID, IIf(IsNull(RecReturn!批次), 0, RecReturn!批次), 1), gtype_UserDrugDigits.Digit_零售价, , True)
        Else
            txtSalePrice.Tag = zlStr.FormatEx(RecReturn!售价, gtype_UserDrugDigits.Digit_零售价, , True)
        End If
        txtSalePrice.Caption = zlStr.FormatEx(Val(txtSalePrice.Tag) * Val(txt单位.Tag), mintPriceDigit, , True)
        txtSale.Caption = zlStr.FormatEx(Val(txtSalePrice.Caption) * Val(Me.TxtNumber), mintMoneyDigit, , True)
        
        If Val(txtProvider.Tag) <> 0 Then
            txtProvider.Text = GetProviderNameById(Val(txtProvider.Tag))
        End If
        
        TxtNumber.SetFocus
        
        
    End If
End Sub


Private Sub TxtNumber_GotFocus()
    TxtNumber.SelStart = 0
    TxtNumber.SelLength = Len(TxtNumber)
End Sub

Private Sub TxtNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(1, "1234567890." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If TxtNumber.SelLength = Len(TxtNumber.Text) Then Exit Sub
            If Len(Mid(TxtNumber, InStr(1, TxtNumber.Text, ".") + 1)) >= mintNumberDigit And TxtNumber.Text Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End If
End Sub


Private Sub TxtNumber_Validate(Cancel As Boolean)
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim lng批次 As Long
    Dim lng数量 As Long
    
    If Trim(TxtNumber.Text) <> "" Then
        If Not IsNumeric(TxtNumber.Text) Then
            MsgBox "对不起，短损数量必须是数字型，请检查！", vbExclamation, gstrSysName
            Cancel = True
        Else
            If txtCostPrice.Caption <> "" Then
                txtCost.Caption = zlStr.FormatEx(zlStr.FormatEx(Val(txtCostPrice.Caption), mintCostDigit) * zlStr.FormatEx(Val(TxtNumber.Text), mintNumberDigit), mintMoneyDigit, , True)
            End If
            If txtSalePrice.Caption <> "" Then
                txtSale.Caption = zlStr.FormatEx(zlStr.FormatEx(txtSalePrice.Caption, mintPriceDigit) * zlStr.FormatEx(Val(TxtNumber.Text), mintNumberDigit), mintMoneyDigit, , True)
            End If
        End If
        
        TxtNumber.Text = zlStr.FormatEx(TxtNumber.Text, mintNumberDigit, , True)
        lng数量 = Val(TxtNumber.Text) * Val(txt单位.Tag)
        lng批次 = txt批号.Tag
        lng药品ID = TxtName.Tag
        lng库房ID = cboStock.ItemData(cboStock.ListIndex)
        If CheckDrugStock(lng库房ID, lng药品ID, lng批次, lng数量, Val(txt单位.Tag)) = False Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0
    txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strProviderText As String
    Dim adoProvider As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtProvider.hWnd)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mint编辑模式 = 4 Then Exit Sub
     
    On Error GoTo errHandle
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) " & _
                  "  And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1]) "
'        Set adoProvider = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        Set adoProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "药品供应商", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If blnCancel Then txtProvider.SetFocus: Exit Sub
         
        If adoProvider.State = 0 Then
            MsgBox "没有你输入的供药单位，请重输！", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = 0
            Exit Sub
        End If
        
        .Text = adoProvider!名称
        .Tag = adoProvider!id

        adoProvider.Close
    
    End With
    If mint编辑模式 = 3 Then
        cmdOk.SetFocus
    Else
        dtp短损日期.SetFocus
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt产地_GotFocus()
    OS.OpenIme True
    txt产地.SelStart = 0
    txt产地.SelLength = Len(txt产地.Text)
End Sub

Private Sub txt产地_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txt产地_LostFocus()
    OS.OpenIme
End Sub

Private Sub txt产地_Validate(Cancel As Boolean)
    If Trim(txt产地.Text) <> "" Then
        If LenB(StrConv(txt产地.Text, vbFromUnicode)) > 30 Then
            MsgBox "生产商超长，最多能输入15个汉字或30个字符!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
Private Sub txt原产地_GotFocus()
    OS.OpenIme True
    txt原产地.SelStart = 0
    txt原产地.SelLength = Len(txt原产地.Text)
End Sub

Private Sub txt原产地_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txt原产地_LostFocus()
    OS.OpenIme
End Sub

Private Sub txt原产地_Validate(Cancel As Boolean)
    If Trim(txt原产地.Text) <> "" Then
        If LenB(StrConv(txt原产地.Text, vbFromUnicode)) > 30 Then
            MsgBox "原产地超长，最多能输入15个汉字或30个字符!", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
Private Sub txt批号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
