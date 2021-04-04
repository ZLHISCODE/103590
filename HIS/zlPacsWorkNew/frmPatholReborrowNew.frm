VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{84865D89-6B2D-42E2-98C7-18F4206945F5}#2.1#0"; "zl9PacsControl.ocx"
Begin VB.Form frmPatholReborrowNew 
   Caption         =   "借阅登记"
   ClientHeight    =   8235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14055
   Icon            =   "frmPatholReborrowNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14055
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture0 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7935
      ScaleWidth      =   9735
      TabIndex        =   25
      Top             =   120
      Width           =   9735
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   100
         Left            =   0
         TabIndex        =   38
         Top             =   3975
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   185
         MousePointer    =   7
         SplitWidth      =   100
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   3000
         Con2MinSize     =   2000
         Control1Name    =   "Picture1"
         Control2Name    =   "Picture2"
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3860
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   9735
         TabIndex        =   35
         Top             =   4075
         Width           =   9735
         Begin VB.CommandButton cmdCancelLend 
            Caption         =   "撤销借出(&R)"
            Height          =   400
            Left            =   8160
            TabIndex        =   36
            Top             =   2880
            Width           =   1215
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialEnreged 
            Height          =   2775
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4895
            DefaultCols     =   ""
            KeyName         =   "≡"
            BackColor       =   12648447
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3975
         ScaleWidth      =   9735
         TabIndex        =   26
         Top             =   0
         Width           =   9735
         Begin VB.Frame framNameQuery 
            Height          =   615
            Left            =   2280
            TabIndex        =   44
            Top             =   0
            Width           =   6135
            Begin VB.TextBox txtPatholName 
               Height          =   300
               Left            =   720
               TabIndex        =   45
               Top             =   200
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dtpStartDate 
               Height          =   300
               Left            =   3120
               TabIndex        =   46
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   57475075
               CurrentDate     =   40899
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   300
               Left            =   4680
               TabIndex        =   47
               Top             =   200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   0   'False
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   57475075
               CurrentDate     =   40899
            End
            Begin VB.Label Label13 
               Caption         =   "姓 名："
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label14 
               Caption         =   "报到日期："
               Height          =   255
               Left            =   2280
               TabIndex        =   49
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "到"
               Height          =   255
               Left            =   4470
               TabIndex        =   48
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.CommandButton cmdLend 
            Caption         =   "借出材料(&L)"
            Height          =   400
            Left            =   8160
            TabIndex        =   32
            Top             =   3480
            Width           =   1215
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "特检材料"
            Height          =   180
            Index           =   2
            Left            =   2160
            TabIndex        =   31
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "切片材料"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CheckBox chkMaterial 
            Caption         =   "蜡块材料"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton cmdQuery 
            Caption         =   "材料查询(&Q)"
            Height          =   400
            Left            =   8520
            TabIndex        =   28
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox txtPatholNum 
            Height          =   330
            Left            =   720
            TabIndex        =   27
            Top             =   200
            Width           =   1455
         End
         Begin zl9PACSWork.ucFlexGrid ufgMaterialEnreg 
            Height          =   2655
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4683
            DefaultCols     =   ""
            GridRows        =   201
            IsCopyAdoMode   =   0   'False
            IsEjectConfig   =   -1  'True
            HeadFontCharset =   134
            HeadFontWeight  =   400
            DataFontCharset =   134
            DataFontWeight  =   400
         End
         Begin VB.Label Label12 
            Caption         =   "病理号："
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8175
      Left            =   10080
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin RichTextLib.RichTextBox rtfReason 
         Height          =   1695
         Left            =   1080
         TabIndex        =   43
         Top             =   4080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2990
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPatholReborrowNew.frx":000C
      End
      Begin VB.CheckBox chkBorrowType 
         Caption         =   "内部借阅"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   7080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取 消(&C)"
         Height          =   400
         Left            =   2520
         TabIndex        =   23
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   400
         Left            =   1320
         TabIndex        =   22
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox txtMemo 
         Height          =   300
         Left            =   1080
         TabIndex        =   21
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox txtEnregPeople 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   6480
         Width           =   2415
      End
      Begin VB.TextBox txtAddress 
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txtMobilePhone 
         Height          =   300
         Left            =   1080
         TabIndex        =   14
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox txtBorrowDays 
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Text            =   "30"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtMoney 
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtCardNum 
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cbxCardType 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpBorrowDate 
         Height          =   300
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   57475075
         CurrentDate     =   40898
      End
      Begin VB.TextBox txtBorrowPeople 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   42
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   41
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3600
         TabIndex        =   40
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   39
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label11 
         Caption         =   "备注说明："
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   6060
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "登 记 人："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   6540
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "借阅原因："
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4140
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "联系地址："
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "联系电话："
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "借阅天数："
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "借阅押金："
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "证件号码："
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "证件类型："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "借阅日期："
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "借 阅 人："
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPatholReborrowNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const DebugState = False


Private mufgParentBorrowGrid As ucFlexGrid

Private mlngBorrowId As Long
Private mblnIsUpdate As Boolean
Private mblnIsEnter As Boolean

Public blnIsOk As Boolean



Public Sub ShowNewBorrowWindow(ufgParentGrid As ucFlexGrid, owner As Object)
'显示新增借阅窗口
    mblnIsUpdate = False
    blnIsOk = False
    mlngBorrowId = -1
    
    Set mufgParentBorrowGrid = ufgParentGrid
    
    txtEnregPeople.Text = UserInfo.姓名
    
    Call Me.Show(1, owner)
End Sub


Public Sub ShowUpdateBorrowWindow(ufgParentGrid As ucFlexGrid, owner As Object)
'显示新增借阅窗口

    
    mblnIsUpdate = True
    blnIsOk = False
    mlngBorrowId = ufgParentGrid.KeyValue(ufgParentGrid.SelectionRow)
    
    Set mufgParentBorrowGrid = ufgParentGrid
    
    Call ConfigUpdateData
    Call LoadBorrowMaterialDetail
    
    Call Me.Show(1, owner)
End Sub


Private Sub ConfigUpdateData()
'配置更新数据
    Dim blnFind As Boolean
    
    If mlngBorrowId <= 0 Then Exit Sub
    
    With mufgParentBorrowGrid
        txtBorrowPeople.Text = .Text(.SelectionRow, gstrPatholCol_借阅人)
        dtpBorrowDate.value = CDate(.Text(.SelectionRow, gstrPatholCol_借阅日期))
        
        cbxCardType.ListIndex = .GetFieldDataValue(gstrPatholCol_证件类型, .Text(.SelectionRow, gstrPatholCol_证件类型), blnFind)
        
        txtCardNum.Text = .Text(.SelectionRow, gstrPatholCol_证件号码)
        txtMoney.Text = .Text(.SelectionRow, gstrPatholCol_押金)
        txtBorrowDays.Text = .Text(.SelectionRow, gstrPatholCol_借阅天数)
        txtMobilePhone.Text = .Text(.SelectionRow, gstrPatholCol_联系电话)
        txtAddress.Text = .Text(.SelectionRow, gstrPatholCol_联系地址)
        rtfReason.Text = .Text(.SelectionRow, gstrPatholCol_借阅原因)
        txtMemo.Text = .Text(.SelectionRow, gstrPatholCol_备注)
        txtEnregPeople.Text = .Text(.SelectionRow, gstrPatholCol_登记人)
        chkBorrowType.value = IIf(.Text(.SelectionRow, gstrPatholCol_借阅类型) = "内部借阅", True, False)
    End With
End Sub


Private Sub LoadBorrowMaterialDetail()
'载入已借阅材料明细
    Dim strSQL As String
    
    If mlngBorrowId <= 0 Then Exit Sub
    

    strSQL = " select b.id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '蜡块' as 材料类别," & _
            " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量,e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
            " Where c.病理医嘱id = d.病理医嘱id And c.材块id = b.材块id and e.id=b.档案ID And a.归档id = b.ID And b.资料来源 = 1 And a.借阅id = [1] " & _
        " Union All " & _
            " select b.id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, '切片' as 材料类别, " & _
            " decode(o.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量,e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理制片信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
            " Where c.病理医嘱id = d.病理医嘱id And o.病理医嘱id = c.病理医嘱id " & _
            " and o.id = b.制片id and e.id=b.档案ID and a.归档id=b.id and b.资料来源=2 and a.借阅id=[1] " & _
        " Union All " & _
            " select b.id, d.检查类型,d.病理号,c.序号,c.标本名称,c.取材位置, " & _
            " decode(o.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
            " decode(o.特检细目,0,decode(o.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || q.抗体名称 || decode(o.制作类型,-1,'-补',0,'','-重' || o.制作类型) || ')' as 材料明细, " & _
            " nvl(a.借阅数量, 0) as 借阅数量, e.档案名称, e.详细地址, " & _
            " '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置 " & _
            " from 病理检查信息 d, 病理取材信息 c, 病理抗体信息 q, 病理特检信息 o, 病理档案信息 e, 病理归档信息 b, 病理借阅关联 a " & _
            " Where c.病理医嘱id = d.病理医嘱id And q.抗体ID = o.抗体ID And o.病理医嘱id = c.病理医嘱id " & _
            " and o.id = b.特检id and e.id=b.档案ID and a.归档id=b.id and b.资料来源=3 and a.借阅id=[1] "

    Set ufgMaterialEnreged.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBorrowId)
    Call ufgMaterialEnreged.RefreshData
End Sub


Private Sub FilterQueryMaterialData()
'过滤查询出的材料数据
    Dim strFilter As String
    
    strFilter = ""
    
    If ufgMaterialEnreg.DataGrid.Rows < 2 Then Exit Sub
    
    If chkMaterial(0).value <> 0 Then
        strFilter = "材料类别='蜡块'"
    End If
    
    If chkMaterial(1).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "材料类别='切片'"
    End If
    
    If chkMaterial(2).value <> 0 Then
        If strFilter <> "" Then strFilter = strFilter & " or "
        strFilter = strFilter & "材料类别='免疫' or 材料类别='分子' or 材料类别='特染'"
    End If
    
    ufgMaterialEnreg.AdoData.Filter = strFilter
    
    Call ufgMaterialEnreg.RefreshData
    
End Sub


Private Sub chkMaterial_Click(Index As Integer)
On Error GoTo ErrHandle
    Call FilterQueryMaterialData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function LendMaterial() As String
'借出材料

    Dim lngNewRow As Long
    Dim i As Long
    Dim strLog As String
    
    
    strLog = ""
    
    For i = 1 To ufgMaterialEnreg.GridRows - 1
    
        If ufgMaterialEnreg.GetRowCheck(i) Then
            If ufgMaterialEnreged.FindRowIndex(ufgMaterialEnreg.Text(i, gstrPatholCol_ID), gstrPatholCol_ID, True) < 1 Then
                If Val(ufgMaterialEnreg.Text(i, gstrPatholCol_需借数量)) > Val(ufgMaterialEnreg.Text(i, gstrPatholCol_可借数量)) Then
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    
                    strLog = strLog & "所选材料 [病理号:" & ufgMaterialEnreg.Text(i, gstrPatholCol_病理号) & " 材块号:" & ufgMaterialEnreg.Text(i, gstrPatholCol_材块号) & _
                                    " 材料明细:" & ufgMaterialEnreg.Text(i, gstrPatholCol_材料明细) & "] 需借数量不能大于可借数量，未能进行借出操作。"
                Else
                
                    lngNewRow = ufgMaterialEnreged.NewRow
                
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_ID) = ufgMaterialEnreg.Text(i, gstrPatholCol_ID)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_检查类型) = ufgMaterialEnreg.Text(i, gstrPatholCol_检查类型)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_病理号) = ufgMaterialEnreg.Text(i, gstrPatholCol_病理号)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_材块号) = ufgMaterialEnreg.Text(i, gstrPatholCol_材块号)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_标本名称) = ufgMaterialEnreg.Text(i, gstrPatholCol_标本名称)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_取材位置) = ufgMaterialEnreg.Text(i, gstrPatholCol_取材位置)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_材料明细) = ufgMaterialEnreg.Text(i, gstrPatholCol_材料明细)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_材料类别) = ufgMaterialEnreg.Text(i, gstrPatholCol_材料类别)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_借阅数量) = ufgMaterialEnreg.Text(i, gstrPatholCol_需借数量)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_所属档案) = ufgMaterialEnreg.Text(i, gstrPatholCol_所属档案)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_存放位置) = ufgMaterialEnreg.Text(i, gstrPatholCol_存放位置)
                    ufgMaterialEnreged.Text(lngNewRow, gstrPatholCol_详细地址) = ufgMaterialEnreg.Text(i, gstrPatholCol_详细地址)
                
                    Call ufgMaterialEnreged.SetRowCheck(lngNewRow, False)
                End If

            Else
                    If strLog <> "" Then strLog = strLog & vbCrLf
                    strLog = strLog & "所选材料 [病理号:" & ufgMaterialEnreg.Text(i, gstrPatholCol_病理号) & " 材块号:" & ufgMaterialEnreg.Text(i, gstrPatholCol_材块号) & _
                                    " 材料明细:" & ufgMaterialEnreg.Text(i, gstrPatholCol_材料明细) & "] 已在借出材料列表中，不能再次进行借出操作。"
            End If
        End If
    Next i
    
    Call ufgMaterialEnreged.LocateRow(lngNewRow)
    
    LendMaterial = strLog
End Function


Private Sub CancelLend()
'撤销材料借出
    Dim i As Long
    
    For i = ufgMaterialEnreged.GridRows - 1 To 1 Step -1
        If ufgMaterialEnreged.GetRowCheck(i) Then
            Call ufgMaterialEnreged.RemoveRow(i)
        End If
    Next i
End Sub


Private Sub cmdCancel_Click()
On Error GoTo ErrHandle
    blnIsOk = False
    
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancelLend_Click()
'撤销材料借出
On Error GoTo ErrHandle
    If Not ufgMaterialEnreged.IsCheckedRow Then
        Call MsgBoxD(Me, "请勾选需要撤销借出的材料。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call CancelLend
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdLend_Click()
'借出材料
On Error GoTo ErrHandle
    Dim strInf As String
    
    If Not ufgMaterialEnreg.IsCheckedRow Then
        Call MsgBoxD(Me, "请勾选需要借出的材料记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strInf = LendMaterial
    
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
         '如果是提示借出数量问题，则弹出提示后将焦点定位到需借数量上，方便用户修改
        If InStr(strInf, "大于") > 0 Then
            '将需借数量默认成1
            ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_需借数量) = 1
            Call ufgDataGridSetFocus(ufgMaterialEnreg, ufgMaterialEnreg.SelectionRow, ufgMaterialEnreg.GetColIndexWithRowCheck)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdQuery_Click()
On Error GoTo ErrHandle
    Call QueryPatholMaterialData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtPatholNum_KeyPress(KeyAscii As Integer)
'回车快捷查询
On Error GoTo ErrHandle

    If KeyAscii = 13 Then
         '调用查询方法
         Call QueryPatholMaterialData
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtPatholName_KeyPress(KeyAscii As Integer)
'回车执行查询
On Error GoTo ErrHandle

    If KeyAscii = 13 Then
         '调用查询方法
         Call QueryPatholMaterialData
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub QueryPatholMaterialData()
'查询病理材料
    Dim strSQL As String
    Dim strFilter As String
    Dim strLinkTable As String
    
    
    strFilter = " and d.报到时间 between [1] and [2] "
    
    strLinkTable = ""
    
    If txtPatholNum.Text <> "" Then
        strFilter = " and d.病理号=[3] "
    Else
        If txtPatholName.Text <> "" Then
            'strLinkTable = "(select id from 病人医嘱记录 a, 病人信息 b where a.病人ID=b.病人ID and a.相关ID is null and b.姓名" & IIf(InStr(txtPatholName.Text, "%") > 1, " like [4]", "=[4]") & ") h "
            'strFilter = strFilter & " and d.医嘱ID=h.ID "
'            strFilter = strFilter & " " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.姓名  like [4]", " and h.姓名 =[4]")
        End If
    End If
        
    
    '统计遗失的材料数量(查询借阅数量时，只能统计未归还的借阅数量，部分归还和进行遗失处理的记录，将进行遗失处理并在遗失数量中体现出)
    strLinkTable = IIf(strLinkTable <> "", strLinkTable & ",", "") & _
                    " (select nvl(sum(遗失数量),0) as 遗失数量, 归档ID " & _
                    " from 病理遗失信息 a, 病理归档信息 b, 病理检查信息 d Where a.归档ID = b.ID And b.病理医嘱id = d.病理医嘱id " & _
                    Replace(strFilter, "and d.医嘱ID=h.ID", "") & " group by 归档ID ) x, " & _
                    " (select (nvl(sum(借阅数量), 0) - nvl(sum(归还数量), 0)) as 已借数量, 归档ID " & _
                    " from 病理借阅关联 a, 病理归档信息 b, 病理检查信息 d where a.归档ID = b.ID And b.病理医嘱id = d.病理医嘱id and a.归还状态=0" & _
                    Replace(strFilter, "and d.医嘱ID=h.ID", "") & " group by 归档ID" & ") y"
    
    
    
    strSQL = "select /*+ Rule*/ * from (select d.检查类型, d.病理号, h.姓名, a.id, c.序号, c.标本名称, c.取材位置, '蜡块' as 材料类别, " & _
            " case when (c.蜡块数 - nvl(x.遗失数量,0) - nvl(y.已借数量, 0)) <= 0 then 0 else 1 end as 需借数量," & _
            " case when c.申请ID is null then '常规取材' else '补取材' end as 材料明细, " & _
            " (c.蜡块数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0) ) as 可借数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态, a.借阅状态," & _
            " f.档案名称, '房间:' || f.所属房间 || ' 柜号:' || f.所属柜号 || ' 抽屉:' || f.所属抽屉 as 存放位置, f.详细地址 " & _
            " from 病理归档信息 a, 病理取材信息 c, 病理检查信息 d, 病理档案信息 f, 病人医嘱记录 h, " & strLinkTable & _
            " where a.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and d.医嘱ID=h.id and h.相关ID is null and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.档案id=f.id and f.档案状态=1 " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.姓名  like [4]", IIf(txtPatholName.Text = "", "", " and h.姓名 =[4]")) & strFilter & _
        " Union All " & _
            " select d.检查类型, d.病理号,h.姓名, a.id, c.序号, c.标本名称, c.取材位置, '切片' as 材料类别, " & _
            " case when (b.制片数 - nvl(x.遗失数量,0) - nvl(y.已借数量, 0)) <= 0 then 0 else 1 end as 需借数量, " & _
            " decode(b.制片方式,0,'正常',1,'重切',2,'深切',3,'连切',4,'白片',5,'重染',6,'薄片','其他') as 材料明细, " & _
            " (b.制片数 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 可借数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态, a.借阅状态, " & _
            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
            " from 病理归档信息 a, 病理制片信息 b, 病理取材信息 c, 病理检查信息 d, 病理档案信息 e, 病人医嘱记录 h," & strLinkTable & _
            " where a.制片id=b.id and b.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and d.医嘱ID=h.id and h.相关ID is null and a.id = x.归档ID(+) and a.id=y.归档id(+) and a.档案id=e.id  and e.档案状态=1 " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.姓名  like [4]", IIf(txtPatholName.Text = "", "", " and h.姓名 =[4]")) & strFilter & _
        " Union All " & _
            " select d.检查类型, d.病理号,h.姓名, a.id, c.序号, c.标本名称, c.取材位置, decode(b.特检类型,0, '免疫',1,'特染',2,'分子') as 材料类别, " & _
            " case when (1 - nvl(x.遗失数量,0) - nvl(y.已借数量, 0)) <= 0 then 0 else 1 end as 需借数量, " & _
            " decode(b.特检细目,0,decode(b.特检类型,0, '免疫',1,'特染',2,'分子'),1,'鉴别',2,'多耐药',3,'荧光',4,'普通') || '(' || f.抗体名称 || decode(b.制作类型,-1,'-补',0,'','-重' || b.制作类型) || ')' as 材料明细, " & _
            " (1 - nvl(x.遗失数量,0)  - nvl(y.已借数量, 0)) as 可借数量, nvl(x.遗失数量,0) as 遗失数量, nvl(y.已借数量,0) as 已借数量, a.存放状态, a.借阅状态, " & _
            " e.档案名称, '房间:' || e.所属房间 || ' 柜号:' || e.所属柜号 || ' 抽屉:' || e.所属抽屉 as 存放位置, e.详细地址 " & _
            " from 病理归档信息 a, 病理特检信息 b, 病理取材信息 c, 病理检查信息 d, 病理档案信息 e, 病理抗体信息 f, 病人医嘱记录 h, " & strLinkTable & _
            " where a.特检id=b.id and b.材块id=c.材块id and c.病理医嘱id=d.病理医嘱id and d.医嘱ID=h.id and h.相关ID is null and a.id = x.归档ID(+) and a.id=y.归档id(+) " & _
            " and a.档案id=e.id  and e.档案状态=1 and b.抗体ID=f.抗体ID " & IIf(InStr(txtPatholName.Text, "%") > 1, " and h.姓名  like [4]", IIf(txtPatholName.Text = "", "", " and h.姓名 =[4]")) & strFilter & _
        ") order by 可借数量 desc,材料类别, 序号,材料明细,存放状态"


    Set ufgMaterialEnreg.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                    CDate(Format(dtpStartDate.value, "yyyy-mm-dd 00:00:00")), _
                                                    CDate(Format(dtpEndDate.value, "yyyy-mm-dd 23:59:59")), _
                                                    txtPatholNum.Text, _
                                                    txtPatholName.Text)
                                                    
    Call ufgMaterialEnreg.RefreshData
                                                          

    If ufgMaterialEnreg.AdoData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "未查询到相关数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
End Sub

Private Function CheckDataIsValid() As String
'检查数据是否有效，有效则返回空字符串
    If ufgMaterialEnreged.ShowingDataRowCount <= 0 Then
        CheckDataIsValid = "没有选取可借阅的材料。"
        Call ufgMaterialEnreged.SetFocus
        
        Exit Function
    End If
    
    If Trim(txtBorrowPeople.Text) = "" Then
        CheckDataIsValid = "借阅人不能为空。"
        Call txtBorrowPeople.SetFocus
        
        Exit Function
    End If
    
    If Trim(txtCardNum.Text) = "" Then
        CheckDataIsValid = "证件号码不能为空。"
        Call txtCardNum.SetFocus
        
        Exit Function
    End If
    
    
    If Val(txtBorrowDays.Text) <= 0 Then
        CheckDataIsValid = "借阅天数不能小于或等于0。"
        Call txtBorrowDays.SetFocus
        
        Exit Function
    End If
    
    If Trim(rtfReason.Text) = "" Then
        CheckDataIsValid = "借阅原因不能为空。"
        Call rtfReason.SetFocus
        
        Exit Function
    End If
End Function


Private Sub NewBorrow()
'新增借阅
    Dim i As Integer
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngNewBorrowId As Long
    Dim lngNewRecordIndex As Long
    
    strSQL = "select Zl_病理借阅_新增借阅([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12]) as 返回值 from dual"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                            txtBorrowPeople.Text, _
                                            CDate(dtpBorrowDate.value), _
                                            cbxCardType.ListIndex, _
                                            txtCardNum.Text, _
                                            Val(txtMoney.Text), _
                                            Val(txtBorrowDays.Text), _
                                            txtMobilePhone.Text, _
                                            txtAddress.Text, _
                                            rtfReason.Text, _
                                            UserInfo.姓名, _
                                            txtMemo.Text, _
                                            IIf(chkBorrowType.value <> 0, 0, 1) _
                                            )
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "NewBorrow", "未成功获取新增后的借阅ID,本次操作失败。")
        Exit Sub
    End If
    
    lngNewBorrowId = Val(Nvl(rsData!返回值))
    
    Call gcnOracle.BeginTrans
    
On Error GoTo errTrans
    For i = 1 To ufgMaterialEnreged.GridRows - 1
        Call zlDatabase.ExecuteProcedure("Zl_病理借阅_新增材料(" & lngNewBorrowId & "," & _
                                                                ufgMaterialEnreged.Text(i, gstrPatholCol_ID) & "," & _
                                                                Val(ufgMaterialEnreged.Text(i, gstrPatholCol_借阅数量)) & ")", _
                                                                Me.Caption)
    Next i
    
    Call gcnOracle.CommitTrans
    
    
    
    With mufgParentBorrowGrid
        lngNewRecordIndex = .NewRow
        
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_ID, lngNewBorrowId, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅号, lngNewBorrowId, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅人, txtBorrowPeople.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅日期, Format(dtpBorrowDate.value, "yyyy-mm-dd"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_归还日期, Format(dtpBorrowDate.value + Val(txtBorrowDays.Text), "yyyy-mm-dd"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_证件类型, cbxCardType.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_证件号码, txtCardNum.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_押金, Val(txtMoney.Text), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅天数, Val(txtBorrowDays.Text), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_联系电话, txtMobilePhone.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_联系地址, txtAddress.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅原因, rtfReason.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_备注, txtMemo.Text, True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_借阅类型, IIf(chkBorrowType.value <> 0, "内部借阅", "外部借阅"), True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_归还状态, "未归还", True)
        Call .SyncText(lngNewRecordIndex, gstrPatholCol_确认状态, "未确认", True)
        
        Call .LocateRow(lngNewRecordIndex)
        
    End With
    
    Exit Sub
errTrans:
    Call gcnOracle.RollbackTrans
End Sub


Private Sub UpdateBorrow()
'更新借阅
    Dim i As Integer
        
    Call gcnOracle.BeginTrans
    
On Error GoTo errTrans

    '更新借阅记录
    Call zlDatabase.ExecuteProcedure("Zl_病理借阅_更新借阅(" & _
                                            mlngBorrowId & ",'" & _
                                            txtBorrowPeople.Text & "'," & _
                                            zlStr.To_Date(dtpBorrowDate.value) & "," & _
                                            cbxCardType.ListIndex & ",'" & _
                                            txtCardNum.Text & "'," & _
                                            Val(txtMoney.Text) & "," & _
                                            Val(txtBorrowDays.Text) & ",'" & _
                                            txtMobilePhone.Text & "','" & _
                                            txtAddress.Text & "','" & _
                                            rtfReason.Text & "','" & _
                                            txtEnregPeople.Text & "','" & _
                                            txtMemo.Text & "'," & _
                                            IIf(chkBorrowType.value <> 0, 0, 1) & ")", Me.Caption)

    '删除所有借阅材料
    Call zlDatabase.ExecuteProcedure("Zl_病理借阅_清除材料(" & mlngBorrowId & ")", Me.Caption)

    For i = 1 To ufgMaterialEnreged.GridRows - 1
        Call zlDatabase.ExecuteProcedure("Zl_病理借阅_新增材料(" & mlngBorrowId & "," & _
                                                                ufgMaterialEnreged.Text(i, gstrPatholCol_ID) & "," & _
                                                                Val(ufgMaterialEnreged.Text(i, gstrPatholCol_借阅数量)) & ")", _
                                                                Me.Caption)
    Next i
    
    Call gcnOracle.CommitTrans
    
    
    With mufgParentBorrowGrid
        Call .SyncText(.SelectionRow, gstrPatholCol_借阅人, txtBorrowPeople.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_借阅日期, dtpBorrowDate.value)
        Call .SyncText(.SelectionRow, gstrPatholCol_证件类型, cbxCardType.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_证件号码, txtCardNum.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_押金, Val(txtMoney.Text))
        Call .SyncText(.SelectionRow, gstrPatholCol_借阅天数, Val(txtBorrowDays.Text))
        Call .SyncText(.SelectionRow, gstrPatholCol_联系电话, txtMobilePhone.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_联系地址, txtAddress.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_借阅原因, rtfReason.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_备注, txtMemo.Text)
        Call .SyncText(.SelectionRow, gstrPatholCol_借阅类型, IIf(chkBorrowType.value <> 0, "内部借阅", "外部借阅"))
        Call .SyncText(.SelectionRow, gstrPatholCol_归还状态, "未归还")
        Call .SyncText(.SelectionRow, gstrPatholCol_确认状态, "未确认")
        
'        Call .LocateRow(.SelectRowIndex)
        
    End With
    
    Exit Sub
errTrans:
    Call gcnOracle.RollbackTrans
    Call err.Raise(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

Private Sub cmdSure_Click()
'确认材料借阅
On Error GoTo ErrHandle
    Dim strInf As String
    
    strInf = CheckDataIsValid()
    If strInf <> "" Then
        Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If Not mblnIsUpdate Then
        Call NewBorrow
    Else
        Call UpdateBorrow
    End If
    
    blnIsOk = True
    
    Call Me.Hide
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
'    #If DebugState = True Then
'        Call InitDebugObject(1294, Me, "zlhis", "HIS")
'    #End If
    
    dtpBorrowDate.value = zlDatabase.Currentdate
    
    dtpStartDate.value = Format(DateAdd("m", -6, dtpBorrowDate.value), "yyyy-mm-dd")
    dtpEndDate.value = Format(dtpBorrowDate.value, "yyyy-mm-dd")

    Call LoadCardType
    
    Call InitMaterialList
    Call InitMaterialEnregedList

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadCardType()
'0-身份证,1-学生证,2-军官证,3-驾驶证,4-护照,5-社保卡,6-残疾证,7-其他
    Call cbxCardType.AddItem("0-身份证")
    Call cbxCardType.AddItem("1-学生证")
    Call cbxCardType.AddItem("2-军官证")
    Call cbxCardType.AddItem("3-驾驶证")
    Call cbxCardType.AddItem("4-护照")
    Call cbxCardType.AddItem("5-社保卡")
    Call cbxCardType.AddItem("6-残疾证")
    Call cbxCardType.AddItem("7-其他")
    
    cbxCardType.ListIndex = 0
End Sub


Private Sub InitMaterialList()
    '设置行数
    ufgMaterialEnreg.GridRows = glngStandardRowCount
    '设置行高
    ufgMaterialEnreg.RowHeightMin = glngStandardRowHeight
    
    '初始化材料查询列表
    ufgMaterialEnreg.IsKeepRows = False
    ufgMaterialEnreg.DefaultColNames = gstrMaterialBorrowEnregCols
    ufgMaterialEnreg.ColNames = gstrMaterialBorrowEnregCols
    ufgMaterialEnreg.ColConvertFormat = gstrMaterialBorrowEnregConvertFormat
End Sub



Private Sub InitMaterialEnregedList()
    '设置行数
    ufgMaterialEnreged.GridRows = glngStandardRowCount
    '设置行高
    ufgMaterialEnreged.RowHeightMin = glngStandardRowHeight

    '初始化材料查询列表
    ufgMaterialEnreged.IsKeepRows = False
    ufgMaterialEnreged.DefaultColNames = gstrMaterialBorrowEnregedCols
    ufgMaterialEnreged.ColNames = gstrMaterialBorrowEnregedCols
    ufgMaterialEnreged.ColConvertFormat = gstrMaterialBorrowEnregConvertFormat
End Sub

Private Sub Form_Resize()
On Error Resume Next
        
    Picture0.Left = 120
    Picture0.Top = 120
    Picture0.Width = Me.ScaleWidth - Frame2.Width - 360
    Picture0.Height = Me.ScaleHeight - 240
    
    Frame2.Top = 0
    Frame2.Left = Me.ScaleWidth - Frame2.Width - 120
    Frame2.Height = Me.ScaleHeight - 120
    
    Call ucSplitter1.RePaint
err.Clear
End Sub

Private Sub Picture1_Resize()
On Error Resume Next
    ufgMaterialEnreg.Left = 0
    ufgMaterialEnreg.Top = framNameQuery.Top + framNameQuery.Height + 120
    ufgMaterialEnreg.Width = Picture1.ScaleWidth
    ufgMaterialEnreg.Height = Picture1.ScaleHeight - cmdLend.Height - framNameQuery.Height - 360
    
    cmdLend.Top = ufgMaterialEnreg.Top + ufgMaterialEnreg.Height + 120
    cmdLend.Left = Picture1.ScaleWidth - cmdLend.Width
    
    chkMaterial(0).Top = cmdLend.Top + 60
    chkMaterial(1).Top = cmdLend.Top + 60
    chkMaterial(2).Top = cmdLend.Top + 60
    
    
err.Clear
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    ufgMaterialEnreged.Top = 0
    ufgMaterialEnreged.Left = 0
    ufgMaterialEnreged.Width = Picture2.ScaleWidth
    ufgMaterialEnreged.Height = Picture2.ScaleHeight - cmdCancelLend.Height - 120
    
    cmdCancelLend.Top = ufgMaterialEnreged.Height + 120
    cmdCancelLend.Left = Picture2.ScaleWidth - cmdCancelLend.Width
err.Clear
End Sub

Private Sub txtPatholName_Change()
On Error Resume Next
    dtpStartDate.Enabled = IIf(txtPatholName.Text = "", False, True)
    dtpEndDate.Enabled = IIf(txtPatholName.Text = "", False, True)
    
    err.Clear
End Sub

Private Sub ufgMaterialEnreg_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInf As String

    If mblnIsEnter Then
        If Not ufgMaterialEnreg.IsCheckedRow Then
            Call MsgBoxD(Me, "请勾选需要借出的材料记录。", vbOKOnly, Me.Caption)
            Exit Sub
        End If

        strInf = LendMaterial

        If strInf <> "" Then
           Call MsgBoxD(Me, strInf, vbOKOnly, Me.Caption)
           '如果是提示借出数量问题，则弹出提示后将焦点定位到需借数量上，方便用户修改
           If InStr(strInf, "大于") > 0 Then
                '将需借数量默认成1
                 ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_需借数量) = 1
                 Call ufgDataGridSetFocus(ufgMaterialEnreg, Row, Col - 1)
           End If
        End If
        
        mblnIsEnter = False
    End If
End Sub

Private Sub ufgMaterialEnreg_OnCheckChanged(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
        '调用单元格得到焦点方法
        Call ufgDataGridSetFocus(ufgMaterialEnreg, Row, Col)
    err.Clear
End Sub

Private Sub ufgDataGridSetFocus(ufgData As ucFlexGrid, ByVal Row As Long, ByVal Col As Long)
'使某的单元格得到焦点，变成正在编辑状态
    ufgData.DataGrid.SetFocus
    If Col = ufgData.GetColIndexWithRowCheck Then
        If ufgData.GetRowCheck(Row) Then
            Call ufgData.DataGrid.Select(Row, Col + 1)
            Call ufgData.DataGrid.ShowCell(Row, Col + 1)
            Call ufgData.DataGrid.EditCell
        End If
    End If
End Sub

Private Sub ufgMaterialEnreg_OnColsNameReSet()
On Error GoTo ErrHandle

    If ufgMaterialEnreg.DataGrid.Rows > 1 Then Call QueryPatholMaterialData

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgMaterialEnreg_OnKeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     '判断是否按下的是回车键，将结果保存到模块变量中
     mblnIsEnter = IIf(KeyAscii = 13, True, False)
End Sub



Private Sub ufgMaterialEnreg_OnKeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    If KeyCode = 13 Then
'        MsgBox "KeyUpEdit事件"
'    End If
End Sub

Private Sub ufgMaterialEnreg_OnNewRow(ByVal Row As Long)
On Error Resume Next
    If Val(ufgMaterialEnreg.Text(Row, gstrPatholCol_可借数量)) <= 0 Then
        Call ufgMaterialEnreg.DisableCheck(Row, ufgMaterialEnreg.GetColIndexWithRowCheck)
    End If
    
    err.Clear
End Sub

Private Sub ufgMaterialEnreg_OnSelChange()
On Error Resume Next
    Dim lngFindRow As Long
    
    If Not ufgMaterialEnreg.IsSelectionRow Then Exit Sub
    
    lngFindRow = ufgMaterialEnreged.FindRowIndex(ufgMaterialEnreg.Text(ufgMaterialEnreg.SelectionRow, gstrPatholCol_ID), gstrPatholCol_ID, True)
    
    If lngFindRow >= 1 Then
        Call ufgMaterialEnreged.LocateRow(lngFindRow)
    End If
err.Clear
End Sub

Private Sub ufgMaterialEnreged_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrHandle

    Call LoadBorrowMaterialDetail
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

