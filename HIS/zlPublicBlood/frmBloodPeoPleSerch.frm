VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmBloodPeoPleSerch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "人员查询"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "frmBloodPeoPleSerch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   300
      Left            =   3180
      TabIndex        =   3
      Top             =   2625
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   300
      Left            =   4290
      TabIndex        =   2
      Top             =   2610
      Width           =   1000
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBloodPeoPleSerch.frx":08CA
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFSerch 
      Height          =   1890
      Left            =   60
      TabIndex        =   1
      Top             =   555
      Width           =   5310
      _cx             =   9366
      _cy             =   3334
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
Attribute VB_Name = "frmBloodPeoPleSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String    '
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mRsPeople As ADODB.Recordset

Private Sub CMDcancel_Click()
    mstrKey = ""
    Unload Me
End Sub

Private Sub CMDok_Click()
    Dim lngi As Long
    mstrKey = ""
    For lngi = 1 To VSFSerch.Rows - 1
        If Abs(Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("选择")))) = 1 And Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("病人id"))) > 0 Then
            mstrKey = Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("病人id"))) & "-" & Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("主页ID"))) & "-" & Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("类型")))
        End If
    Next
    If mstrKey = "" Then
        MsgBox "请选择要添加的病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    Unload Me
End Sub

Public Function SerchPeople(frmMain As Object, lngModule As Long) As String
    Dim strCardKind As String
    mstrKey = ""
    '初始化Patidentify控件
    Call CreateSquareCardObject(frmMain, 2200, lngModule)
    strCardKind = "姓|姓名|0|0|0|0|0|0;住|住院号|0|0|0|0|0|0;门|门诊号|0|0|0|0|0|0;就|就诊卡|0|0|8|0|0|0;身|二代身份证|0|0|0|0|0|0;IC|IC卡|1|0|0|0|0|0"
    If Not gobjCardSquare Is Nothing Then
        strCardKind = gobjCardSquare.zlGetIDKindStr(strCardKind)
    End If
    '这个对象传入Nothing,传入主窗体，主窗体关闭时会触发active事件（应该是多次刷多次调用该方法的问题）
    Call PatiIdentify.zlInit(Nothing, 2200, , gcnOracle, gstrDBUser, gobjCardSquare, strCardKind)
    PatiIdentify.AutoSize = True
'    PI1.ShowPropertySet = True
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    '初始化表格
    Call initvsf
    
    Me.Show 1, frmMain
    SerchPeople = mstrKey
End Function

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '功能：读取卡号、查找数据
    Dim strSQL As String
    Dim lngPatiID As Long
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人id
    End If
    If lngPatiID = 0 Then
        Call FindPeople(PatiIdentify.Text)
    Else
            strSQL = _
                " Select Distinct b.病人id, b.主页id, b.姓名, b.性别, b.年龄, b.出院病床 As 床号, b.住院号, '' As 挂号单,0 类型" & vbNewLine & _
                " From 血液收发记录 d, 血液配血记录 c, 病案主页 b, 病人信息 a" & vbNewLine & _
                " Where d.配发id = c.Id And Mod(d.记录状态, 3) = 1 And d.审核人 Is Not Null And c.病人id = b.病人id And c.主页id = b.主页id And" & vbNewLine & _
                "      b.病人id = a.病人id And b.主页id = a.主页id And a.在院 = 1 And a.病人ID=[1]" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct b.病人id, a.Id As 主页id, a.姓名, a.性别, a.年龄, '' As 床号, 0 As 住院号, a.No As 挂号单,1 类型" & vbNewLine & _
                " From 血液收发记录 d, 血液配血记录 c, 病人医嘱记录 b, 病人挂号记录 a" & vbNewLine & _
                " Where d.配发id = c.Id And Mod(d.记录状态, 3) = 1 And d.审核人 Is Not Null And c.申请id = b.Id And b.病人id = a.病人id And b.挂号单 = a.No And" & vbNewLine & _
                "      b.诊疗类别 = 'K' And a.执行状态 = 2 And a.记录性质 = 1 And a.记录状态 = 1 And a.病人id = [1]"
        Set mRsPeople = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", lngPatiID)
        Call mclsVsf.LoadGrid(mRsPeople)
    End If
    If mRsPeople.EOF = True Then
        VSFSerch.TextMatrix(1, VSFSerch.ColIndex("选择")) = -1 '默认首行数据初始为选中状态
    End If
End Sub


Private Sub initvsf()
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, VSFSerch, True, True)
        Call .ClearColumn
        Call .AppendColumn("选择", 400, flexAlignLeftCenter, flexDTBoolean, "", "选择", True)
        Call .AppendColumn("病人id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("主页id", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
        Call .AppendColumn("姓名", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("性别", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("年龄", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("床号", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("住院号", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("挂号单", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("类型", 800, flexAlignLeftCenter, flexDTLong, "", , False, False, False, True)
        
        .AppendRows = False
        .SysHidden(.ColIndex("病人id")) = True
        .SysHidden(.ColIndex("主页id")) = True

        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
        
    End With
End Sub

Private Sub FindPeople(strfind As String)
    Dim strSQL As String
    If strfind = "" Then Exit Sub
    strSQL = _
        " Select Distinct b.病人id, b.主页id, b.姓名, b.性别, b.年龄, b.出院病床 As 床号, b.住院号, '' As 挂号单,0 类型" & vbNewLine & _
        " From 血液收发记录 d, 血液配血记录 c, 病案主页 b, 病人信息 a" & vbNewLine & _
        " Where d.配发id = c.Id And Mod(d.记录状态, 3) = 1 And d.审核人 Is Not Null And c.病人id = b.病人id And c.主页id = b.主页id And" & vbNewLine & _
        "      b.病人id = a.病人id And b.主页id = a.主页id And a.在院 = 1 And a.姓名 Like [1]" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select Distinct b.病人id, a.Id As 主页id, a.姓名, a.性别, a.年龄, '' As 床号, 0 As 住院号, a.No As 挂号单,1 类型" & vbNewLine & _
        " From 血液收发记录 d, 血液配血记录 c, 病人医嘱记录 b, 病人挂号记录 a, 病人信息 e" & vbNewLine & _
        " Where d.配发id = c.Id And Mod(d.记录状态, 3) = 1 And d.审核人 Is Not Null And c.申请id = b.Id And b.病人id = a.病人id And b.挂号单 = a.No And" & vbNewLine & _
        "      b.诊疗类别 = 'K' And a.执行状态 = 2 And a.记录性质 = 1 And a.记录状态 = 1 And a.病人id = e.病人id And e.姓名 Like [1]"
    Set mRsPeople = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", strfind & "%")
    Call mclsVsf.LoadGrid(mRsPeople)
End Sub

Private Sub VSFSerch_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSFSerch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngi As Long
    If Col <> VSFSerch.ColIndex("选择") Then Cancel = True: Exit Sub
    If Val(VSFSerch.TextMatrix(Row, VSFSerch.ColIndex("病人id"))) = 0 Then Cancel = True: Exit Sub
    For lngi = VSFSerch.FixedRows To VSFSerch.Rows - 1
        If lngi <> Row Then
            VSFSerch.TextMatrix(lngi, Col) = 0
        End If
    Next
End Sub

Private Sub VSFSerch_DblClick()
    Dim lngRow As Long
    If VSFSerch.Row >= VSFSerch.FixedRows And VSFSerch.Row < VSFSerch.Rows Then
        If Val(VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("病人id"))) = 0 Then Exit Sub
        For lngRow = VSFSerch.FixedRows To VSFSerch.Rows - 1
            If lngRow <> VSFSerch.Row Then
                VSFSerch.TextMatrix(lngRow, VSFSerch.ColIndex("选择")) = 0
            End If
        Next
        If Abs(Val(VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("选择")))) = 0 Then
            VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("选择")) = 1
        Else
            VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("选择")) = 0
        End If
    End If
End Sub
