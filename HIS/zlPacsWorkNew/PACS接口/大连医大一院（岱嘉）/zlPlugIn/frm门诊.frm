VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm门诊 
   BorderStyle     =   0  'None
   ClientHeight    =   735
   ClientLeft      =   6300
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt详细 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   75
      Width           =   1965
   End
   Begin VB.TextBox txt余额 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2835
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   750
   End
   Begin VB.TextBox txt总费用 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   750
   End
   Begin VB.TextBox txt人员类别 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   420
      Width           =   5745
      _cx             =   10134
      _cy             =   503
      Appearance      =   2
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm门诊.frx":0000
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
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      X1              =   -120
      X2              =   8425
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   -120
      X2              =   8425
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "详"
      Height          =   180
      Left            =   3690
      TabIndex        =   7
      Top             =   75
      Width           =   180
   End
   Begin VB.Line Line4 
      DrawMode        =   1  'Blackness
      X1              =   3885
      X2              =   5855
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line3 
      DrawMode        =   1  'Blackness
      X1              =   2835
      X2              =   3645
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "余额"
      Height          =   180
      Left            =   2415
      TabIndex        =   5
      Top             =   75
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "费用"
      Height          =   180
      Left            =   1170
      TabIndex        =   3
      Top             =   75
      Width           =   360
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   1560
      X2              =   2370
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   375
      X2              =   1150
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lab人员类别 
      AutoSize        =   -1  'True
      Caption         =   "人员"
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   360
   End
End
Attribute VB_Name = "frm门诊"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiID          As Long
Private mvarRecId           As Variant
Private mvarKeyId           As Variant
Private mstrReserve         As String
Private mintRecord          As Long

Const col单病种 = &HFF&
Const col普通病 = vbBlack
Const col慢性病 = &HFF0000
Const col特种病 = &HFF00FF

Private Type typ_病种信息
    str编码                 As String
    str类别                 As String
    str名称                 As String
    str说明                 As String
    color                   As Long
End Type
Private var病种()           As typ_病种信息

Const con离休可报费用       As Double = 8000
Dim rsTmp                   As ADODB.Recordset

Public Property Let PatiID(ByVal vNewValue As Long)
    mlngPatiID = vNewValue
End Property

Public Property Let RecId(ByVal vNewValue As Variant)
    mvarRecId = vNewValue
End Property

Public Property Let KeyId(ByVal vNewValue As Variant)
    mvarKeyId = vNewValue
End Property

Public Property Let Reserve(ByVal vNewValue As String)
    mstrReserve = vNewValue
End Property

Public Sub RefreshData()
    Dim rtn                 As Long
    Dim rsSum               As ADODB.Recordset
    Dim dbl门诊总费用       As Double
    Dim dbl住院总费用       As Double
    
    DoEvents
    Me.Show
    rtn = SetWindowPos(Me.hWnd, -1, CurrentX, CurrentY, 0, 0, 3)
    '读取人员类别
    gstrSql = "select A.在职,B.名称 from 保险帐户 A ,保险人群 B where A.在职=B.序号 AND A.险类=B.险类 And A.病人ID=[1]"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    If ChkRsState(rsTmp) Then
        txt人员类别.Text = ""
        txt总费用.Text = ""
        txt详细.Text = ""
        txt余额.Text = ""
        txt总费用.Visible = False
        txt详细.Visible = False
        txt余额.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
        Me.Hide
        DoEvents
    Else
        txt人员类别.Text = rsTmp!名称
        txt总费用.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        txt详细.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        txt余额.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Label2.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Label3.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Label1.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Line2.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Line3.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        Line4.Visible = (rsTmp!在职 = "3" Or rsTmp!在职 = "5")
        If Not (rsTmp!在职 = "3" Or rsTmp!在职 = "5") Then
            txt人员类别.Text = rsTmp!名称
        Else
            '读取门诊总费用
            gstrSql = "select nvl(sum(累计统筹报销), 0) as 金额" & vbCrLf & _
                      "  From 保险结算记录" & vbCrLf & _
                      " Where 性质 = [1]" & vbCrLf & _
                      "   And 病人ID in" & vbCrLf & _
                      "       (Select 病人ID" & vbCrLf & _
                      "          From 医保病人关联表" & vbCrLf & _
                      "         where 医保号 in" & vbCrLf & _
                      "               (Select 医保号 from 医保病人关联表 where 病人ID = [2]))"
            Set rsSum = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, 1, mlngPatiID)
            dbl门诊总费用 = rsSum!金额
            '读取住院总费用
            Set rsSum = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, 2, mlngPatiID)
            dbl住院总费用 = rsSum!金额
            '总费用
            txt总费用.Text = Format(dbl门诊总费用 + dbl住院总费用, "0.00")
            txt详细.Text = "  门：" & Format(dbl门诊总费用, "0") & ";住：" & Format(dbl住院总费用, "0")
            txt余额.Text = Format(con离休可报费用 - dbl门诊总费用 - dbl住院总费用, "0.00")
        End If
    End If
    '读取病种信息
'    cmb病种.Clear
    gstrSql = "SELECT B.编码,DECODE(B.类别,1,'慢性病',2,'特种病',3,'单病种','普通病') AS 类别,B.名称,C.说明  FROM 大连_特病人员 A,保险病种 B,疾病编码目录 C" & vbCrLf & _
              "WHERE A.取消人 is Null And A.病种ID=B.ID AND A.险类 = B.险类 AND B.编码 = C.编码(+) AND A.医保号=(Select 医保号 from 医保病人关联表 where 病人ID = [1])"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    Set vsfDetail.DataSource = rsTmp
    If ChkRsState(rsTmp) Then
        Me.Height = 370
    Else
        vsfDetail.Height = rsTmp.RecordCount * 265 + 20
        Me.Height = vsfDetail.Top + vsfDetail.Height + 10
        
'        ReDim var病种(rsTmp.RecordCount - 1) As typ_病种信息
'        Do While Not rsTmp.EOF
'            var病种(rsTmp.Bookmark - 1).color = Decode(rsTmp!类别, "慢性病", col慢性病, "特种病", col特种病, "单病种", col单病种, col普通病)
'            var病种(rsTmp.Bookmark - 1).str编码 = "" & rsTmp!编码
'            var病种(rsTmp.Bookmark - 1).str类别 = "" & rsTmp!类别
'            var病种(rsTmp.Bookmark - 1).str名称 = "" & rsTmp!名称
'            var病种(rsTmp.Bookmark - 1).str说明 = "" & rsTmp!说明
'            cmb病种.AddItem var病种(rsTmp.Bookmark - 1).str编码
'            cmb病种.ItemData((rsTmp.Bookmark - 1)) = rsTmp.Bookmark - 1
'            rsTmp.MoveNext
'        Loop
'        cmb病种.ListIndex = 0
'        cmb病种.Enabled = rsTmp.RecordCount > 1
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = 0
    Me.Left = 6300
End Sub

Private Sub cmb病种_Click()
'    txt类别.ForeColor = var病种(cmb病种.ListIndex).color
'    txt类别.Text = var病种(cmb病种.ListIndex).str类别
'    txt名称.Text = var病种(cmb病种.ListIndex).str名称
'    txt说明.Text = var病种(cmb病种.ListIndex).str说明
End Sub

