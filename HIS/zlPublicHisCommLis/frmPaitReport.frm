VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPaitReport 
   Caption         =   "病人报告查看"
   ClientHeight    =   11160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   Icon            =   "frmPaitReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11160
   ScaleWidth      =   16755
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   270
      ScaleHeight     =   8265
      ScaleWidth      =   15615
      TabIndex        =   16
      Top             =   1950
      Width           =   15645
      Begin VB.PictureBox picPaitList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5925
         Left            =   150
         ScaleHeight     =   5895
         ScaleWidth      =   3375
         TabIndex        =   22
         Top             =   390
         Width           =   3405
         Begin VSFlex8Ctl.VSFlexGrid vsfPaitList 
            Height          =   3525
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   3705
            _cx             =   6535
            _cy             =   6218
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
      Begin VB.Frame fraWE 
         BorderStyle     =   0  'None
         Height          =   3645
         Left            =   3810
         MousePointer    =   9  'Size W E
         TabIndex        =   20
         Top             =   870
         Width           =   105
      End
      Begin VB.PictureBox picPaitReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7935
         Left            =   4050
         ScaleHeight     =   7905
         ScaleWidth      =   11295
         TabIndex        =   17
         Top             =   180
         Width           =   11325
         Begin VSFlex8Ctl.VSFlexGrid vsfScroll 
            Height          =   6315
            Left            =   180
            TabIndex        =   18
            Top             =   810
            Width           =   8745
            _cx             =   15425
            _cy             =   11139
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
            ForeColor       =   -2147483643
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483643
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483643
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483632
            FloodColor      =   -2147483643
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   -1  'True
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
            BackColorFrozen =   -2147483643
            ForeColorFrozen =   -2147483643
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.PictureBox picScroll 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   6015
               Left            =   420
               ScaleHeight     =   6015
               ScaleWidth      =   8055
               TabIndex        =   19
               Top             =   90
               Visible         =   0   'False
               Width           =   8055
               Begin zl9LisInsideComm.uclReport uclSampleReport 
                  Height          =   5145
                  Index           =   0
                  Left            =   60
                  TabIndex        =   21
                  Top             =   60
                  Width           =   7965
                  _extentx        =   14049
                  _extenty        =   10451
               End
            End
         End
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1785
      ScaleWidth      =   16455
      TabIndex        =   0
      Top             =   120
      Width           =   16485
      Begin VB.PictureBox picIDKIND 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   285
         ScaleWidth      =   705
         TabIndex        =   26
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtPaitKey 
         Height          =   315
         Left            =   960
         TabIndex        =   25
         ToolTipText     =   "“－”开头头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
         Top             =   210
         Width           =   5955
      End
      Begin VB.CheckBox chkVerifyDate 
         Height          =   255
         Left            =   7530
         TabIndex        =   5
         Top             =   960
         Width           =   300
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   945
         TabIndex        =   4
         Top             =   570
         Width           =   2970
      End
      Begin VB.ComboBox cbodor 
         Height          =   300
         Left            =   4830
         TabIndex        =   3
         Top             =   570
         Width           =   2910
      End
      Begin VB.ComboBox cboDiseases 
         Height          =   300
         ItemData        =   "frmPaitReport.frx":6852
         Left            =   945
         List            =   "frmPaitReport.frx":685F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1305
         Width           =   1485
      End
      Begin VB.TextBox txtRptCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   3570
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0"
         Top             =   1335
         Width           =   510
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Left            =   2490
         TabIndex        =   6
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpVS 
         Height          =   300
         Left            =   4830
         TabIndex        =   7
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpVE 
         Height          =   300
         Left            =   6225
         TabIndex        =   8
         Top             =   930
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpS 
         Height          =   300
         Left            =   945
         TabIndex        =   9
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.UpDown upd 
         Height          =   330
         Left            =   4170
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1230
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         OrigLeft        =   3480
         OrigTop         =   420
         OrigRight       =   3735
         OrigBottom      =   690
         Max             =   99
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label 病历号 
         AutoSize        =   -1  'True
         Caption         =   "病 历 号"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "报告日期"
         Height          =   240
         Left            =   4080
         TabIndex        =   15
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "申请日期"
         Height          =   240
         Left            =   60
         TabIndex        =   14
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblDor 
         AutoSize        =   -1  'True
         Caption         =   "开单医生"
         Height          =   180
         Left            =   4020
         TabIndex        =   13
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "申请科室↓"
         Height          =   180
         Left            =   60
         TabIndex        =   12
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "传 染 病"
         Height          =   180
         Left            =   60
         TabIndex        =   11
         Top             =   1365
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   3570
         X2              =   4140
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "查看病人最近          份报告"
         Height          =   225
         Left            =   2520
         TabIndex        =   10
         Top             =   1350
         Width           =   2925
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPaitReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'动态设置是否显示窗体边框
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const const_PicRectBackColour As Long = &HE0E0E0

'打印PDF
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private mblnDoctorShow As Boolean                       '是否是医生站调用
Private mstrPrivs As String                             '操作员权限
Private mlngPatientID As Long                           '病人ID
Private mlngPatientPage As Long                         '主页ID
Private mstrPatientGH As String                         '挂号单
Private mstrThirdReport As String                       '三方报告
Private WithEvents mobjIDKind As VBControlExtender      'IDKind对象
Attribute mobjIDKind.VB_VarHelpID = -1

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-26
'功    能:  加载窗体
'           objfrm              调用对象
'           intShowType         调用来源，1=技师站调用，可以显示和打印未审核的报告，2=医生站调用，此时只显示已经审核的报告，
'           lngPatientID        病人ID
'           strPrivs            模块权限
'           lngDept             打开当前模块的科室
'           lngDeptDistrict     打开当前模块的病区
'           intPatientType      病人来源
'           lngPatientPage      主页ID
'           blnShowBorder       是否显示窗体
'           blnFindData         打开窗体时是否默认加载数据
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function showMe(objFrm As Object, Optional lngPatientID As Long, Optional strPrivs As String, Optional lngDept As Long, Optional lngDeptDistrict As Long, _
                       Optional intPatientType As Integer, Optional lngPatientPage As Long, Optional strErr As String, Optional ByVal blnShowBorder As Boolean, _
                       Optional ByRef objOutFrm As Object, Optional blnDoctorShow As Boolean = True, Optional ByVal blnFindData As Boolean = True) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

          '获取权限
1         On Error GoTo showMe_Error

2         mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 2001)
3         mstrPrivs = strPrivs & ";" & mstrPrivs
4         mblnDoctorShow = blnDoctorShow
5         mlngPatientPage = lngPatientPage


6         If lngPatientID <> 0 Then txtPaitKey.Text = lngPatientID

7         If lngDeptDistrict > 0 Then
              '住院查看检验报告
8             lblDept.Caption = "申请病区↓"
9             Call InitDepts(1)
10            Call GetDeptDor
11            If cboDept.ListCount > 0 Then
12                CboFind cboDept, lngDeptDistrict
13            End If
14            If cbodor.ListCount > 0 Then
15                CboFind cbodor, UserInfo.ID
16            End If

              '查询住院病人入出院时间
17            strSQL = "select 入院日期,nvl(出院日期,sysdate) 出院日期 from 病案主页 where 病人id=[1] and 主页ID=[2]"
18            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病案主页", lngPatientID, lngPatientPage)
19            If Not rsTmp.EOF Then
20                dtpS.value = CDate(Format(rsTmp("入院日期") & "", "yyyy/mm/dd hh:mm:ss"))
21                dtpE.value = CDate(Format(rsTmp("出院日期") & "", "yyyy/mm/dd hh:mm:ss"))
22            End If
23        Else
              '门诊查看检验报告
24            lblDept.Caption = "申请科室↓"
25            Call InitDepts(0)
26            Call GetDeptDor
27            If cboDept.ListCount > 0 Then
28                CboFind cboDept, lngDept
29            End If
30        End If

31        If blnShowBorder Then
32            Me.Show  '如果不显示窗体的边框，则表示该窗体为嵌入式调用，不是调用show方法
33        Else
34            Call YSystemMenu(Me.hWnd)
35        End If
          
          '默认加载数据
36        If blnFindData Then
37            Call GetDeptPaits
38        End If

39        Set objOutFrm = Me

40        showMe = True


41        Exit Function
showMe_Error:
42        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(showMe)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
43        Err.Clear
End Function

Public Function BHShowMe(lngMain As Long, Optional strErr As String) As Boolean
    On Error GoTo errH
    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)
    

    gobjLiscomlib.ShowChildWindow Me.hWnd, lngMain
    BHShowMe = True
        

    Exit Function
errH:
    strErr = "出错函数(ShowMe),出错信息:" & Err.Number & " " & Err.Description
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/25
'功    能:调用API动态设置窗体的border
'入    参:
'           new_Hwnd    窗体的句柄
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
      '功能：初始化住院临床科室
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strDeptIDs As String, lngPreDept As Long


1         On Error GoTo InitDepts_Error

2         If mblnDoctorShow Then
3             If intDeptView = 0 Then
                  '按科室读取显示
                  '包含门急诊观察室的病人还没有上床，不加只显床上有病人的科室的限制
4                 If InStr(";" & mstrPrivs & ";", ";全院病人;") > 0 Then
5                     strDeptIDs = GetUser科室IDs
6                     strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,a.简码" & _
                             " From 部门表 A,部门性质说明 B" & _
                             " Where B.部门ID=A.ID And B.工作性质='临床'" & _
                             " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                             " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                             " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                             " Order by A.编码"
7                 Else
                      '求有权限的科室：本身所在科室+所属病区包含的科室
8                     strSQL = _
                    " Select A.ID,A.编码,A.名称,a.简码,Nvl(C.缺省,0) as 缺省" & _
                             " From 部门表 A,部门性质说明 B,部门人员 C" & _
                             " Where B.部门ID=A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                             " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                             " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                             " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                             " And B.工作性质='临床'"
9                     strSQL = strSQL & " Union " & _
                             " Select C.ID,C.编码,C.名称,C.简码,Nvl(A.缺省,0) As 缺省" & _
                             " From 部门人员 A,病区科室对应 B,部门表 C" & _
                             " Where A.部门ID=B.病区ID And B.科室ID=C.ID And A.人员ID=[1]" & _
                             " And Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.病区ID)" & _
                             " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.病区ID)" & _
                             " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                             " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
10                    If InStr(";" & mstrPrivs & ";", ";ICU病人;") > 0 Then
11                        strSQL = strSQL & " Union " & _
                                 " Select A.ID,A.编码,A.名称,a.简码,0 As 缺省" & _
                                 " From 部门表 A" & _
                                 " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                                 " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='临床')" & _
                                 " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                                 " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
12                    End If
13                    strSQL = "Select ID,编码,名称,简码,Max(缺省) As 缺省 From (" & strSQL & ") Group By ID,编码,名称,简码 Order by 编码"
14                End If
15            Else
                  '按病区读取显示
16                If InStr(";" & mstrPrivs & ";", ";全院病人;") > 0 Then
17                    strDeptIDs = GetUser病区IDs
18                    strSQL = _
                    " Select Distinct A.ID,A.编码,A.名称,a.简码" & _
                             " From 部门表 A,部门性质说明 B " & _
                             " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                             " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                             " Order by A.编码"
19                Else
                      '求有权病区：直接所在病区+所在科室所属病区
20                    strSQL = _
                    " Select A.ID,A.编码,A.名称,a.简码,Nvl(C.缺省,0) as 缺省" & _
                             " From 部门表 A,部门性质说明 B,部门人员 C" & _
                             " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                             " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                             " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
21                    strSQL = strSQL & " Union " & _
                             " Select C.ID,C.编码,C.名称,C.简码,Nvl(A.缺省,0) as 缺省" & _
                             " From 部门人员 A,病区科室对应 B,部门表 C" & _
                             " Where A.部门ID=B.科室ID And B.病区ID=C.ID And A.人员ID=[1]" & _
                             " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.科室ID)" & _
                             " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.科室ID)" & _
                             " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)"
22                    If InStr(";" & mstrPrivs & ";", ";ICU病人;") > 0 Then
23                        strSQL = strSQL & " Union " & _
                                 " Select A.ID,A.编码,A.名称,a.简码,0 As 缺省" & _
                                 " From 部门表 A" & _
                                 " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                                 " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='护理')" & _
                                 " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                                 " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
24                    End If
25                    strSQL = "Select ID,编码,名称,简码,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称,简码 Order by 编码"
26                End If
27            End If
28        Else
29            strSQL = "Select Distinct a.id, a.编码, a.名称, a.简码 From 部门表 A, 部门性质说明 B" & _
                     " Where a.Id = b.部门id And a.撤档时间 Is Not Null And a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd') And" & _
                     " (b.工作性质 = '临床' Or b.工作性质 = '治疗' Or b.工作性质 = '护理' Or b.工作性质 = '检验') order by a.编码"
30        End If

31        cboDept.Clear
32        If InStr(";" & mstrPrivs & ";", ";所有科室;") > 0 Then cboDept.AddItem "00-所有科室"
33        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, UserInfo.ID)

34        For i = 1 To rsTmp.RecordCount
35            cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称 & "[" & rsTmp!简码 & "]"
36            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
37            rsTmp.MoveNext
38        Next
39        If rsTmp.RecordCount > 0 Then
40            cboDept.ListIndex = 0
41        End If
42        InitDepts = True


43        Exit Function
InitDepts_Error:
44        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "执行(InitDepts)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
45        Err.Clear

End Function

Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean, Optional strErr As String) As String
      '功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
      '参数：是否取所属病区下的科室
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser科室IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          '没有强制限制临床,可能医技科室用
7         If blnNew Then
8             strSQL = "Select 1 as 类别,部门ID From 部门人员 Where 人员ID=[1] Union" & _
                     " Select Distinct 2 as 类别,B.科室ID From 部门人员 A,病区科室对应 B" & _
                     " Where A.部门ID=B.病区ID And A.人员ID=[1]"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", UserInfo.ID)
10        End If
11        If bln病区 = False Then
12            rsTmp.Filter = "类别 = 1"
13        Else
14            rsTmp.Filter = ""
15        End If

16        For i = 1 To rsTmp.RecordCount
17            If InStr("," & GetUser科室IDs & ",", "," & rsTmp!部门ID & ",") = 0 Then
18                GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
19            End If
20            rsTmp.MoveNext
21        Next
22        GetUser科室IDs = Mid(GetUser科室IDs, 2)



23        Exit Function
GetUser科室IDs_Error:
24        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "执行(GetUser科室IDs)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear

End Function

Public Function GetUser病区IDs(Optional strErr As String) As String
      '功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser病区IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
7         If blnNew Then
8             strSQL = _
              "Select Distinct 病区ID From (" & _
                     " Select A.部门ID as 病区ID" & _
                     " From 部门性质说明 A,部门人员 B" & _
                     " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
                     " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
                     " Union" & _
                     " Select A.病区ID From 病区科室对应 A,部门人员 B" & _
                     " Where A.科室ID=B.部门ID And B.人员ID=[1])"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", UserInfo.ID)
10        ElseIf rsTmp.RecordCount > 0 Then
11            rsTmp.MoveFirst
12        End If
13        For i = 1 To rsTmp.RecordCount
14            GetUser病区IDs = GetUser病区IDs & "," & rsTmp!病区ID
15            rsTmp.MoveNext
16        Next

17        GetUser病区IDs = Mid(GetUser病区IDs, 2)



18        Exit Function
GetUser病区IDs_Error:
19        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "执行(GetUser病区IDs)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
20        Err.Clear

End Function

Private Sub CboFind(objcbo As ComboBox, lngID As Long)
    '功能           找到cbo对应的id
    Dim intloop As Integer
    With objcbo
        For intloop = 0 To .ListCount - 1
            If .ItemData(intloop) = lngID Then
                .ListIndex = intloop
                Exit Sub
            End If
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub cboDept_Click()
    '获取科室医生
    Call GetDeptDor(cboDept.ItemData(cboDept.ListIndex))
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-26
'功    能:  获取选中科室病人
'入    参:
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Sub GetDeptPaits()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsPait As ADODB.Recordset
          Dim strPatientIDs As String
          Dim strArr() As String
          Dim i As Integer
          Dim j As Integer
          Dim lngPaitID As Long



1         On Error GoTo GetDeptPaits_Error

2         If Trim(txtPaitKey.Text) <> "" Then
              '刷卡通过病人ID查找，通过ID查找时忽略其他条件
3             If Mid(Trim(txtPaitKey.Text), 1, 1) = "-" Then
4                 If IsNumeric(Mid(Trim(txtPaitKey.Text), 2)) Then lngPaitID = Mid(Trim(txtPaitKey.Text), 2)
5             Else
6                 If IsNumeric(Trim(txtPaitKey.Text)) Then lngPaitID = Trim(txtPaitKey.Text)
7             End If

8             strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.HIS病人ID Order By a.申请时间 Desc) 序号, a.HIS病人ID, a.姓名," & vbCrLf & _
                     "                 Decode(a.性别, '1', '男', '2', '女', '9', '未知', '不区分') 性别, a.年龄," & vbCrLf & _
                     "                 Nvl(a.病历号, Decode(a.病人来源, 1, a.门诊号, 2, a.住院号, 3, a.病历号, 4, a.健康号, Decode(a.挂号单, Null, a.收费单号, a.挂号单))) 病历号,a.挂号单" & vbCrLf & _
                     "          From 检验报告记录 A" & vbCrLf & _
                     "          Where Nvl(a.是否质控标本, 0) = 0 and a.HIS病人ID =[1]) Where 序号 = 1"
9             Set rsPait = ComOpenSQL(Sel_Lis_DB, strSQL, "病人列表", lngPaitID)
10        Else
              '在新版中去查找病人
11            strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.HIS病人ID Order By a.申请时间 Desc) 序号, a.HIS病人ID, a.姓名," & vbCrLf & _
                     "                 Decode(a.性别, '1', '男', '2', '女', '9', '未知', '不区分') 性别, a.年龄," & vbCrLf & _
                     "                 Nvl(a.病历号, Decode(a.病人来源, 1, a.门诊号, 2, a.住院号, 3, a.病历号, 4, a.健康号, Decode(a.挂号单, Null, a.收费单号, a.挂号单))) 病历号,a.挂号单" & vbCrLf & _
                     "          From 检验报告记录 A" & vbCrLf & _
                     "          Where Nvl(a.是否质控标本, 0) = 0 and a.HIS病人ID is not null And a.申请时间 Between [1] And [2] "


12            If Trim(Me.cboDept.Text) <> "00-所有科室" Then
13                strSQL = strSQL & " and (a.申请科室=[3] or a.申请科室 is null)"
14            End If

15            If Trim(Me.cbodor.Text) <> "00-所有" Then
16                strSQL = strSQL & " and (a.申请人=[4] or a.申请人 is null)"
17            End If

18            Select Case Trim(cboDiseases.Text)
              Case "所有"

19            Case "传染病"
20                strSQL = strSQL & " and nvl(a.是否传染病,0)=1"
21            Case "非传染病"
22                strSQL = strSQL & " and nvl(a.是否传染病,0)=0"
23            End Select

24            If chkVerifyDate.value = 1 Then
25                strSQL = strSQL & " and a.审核时间 between [5] and [6]"
26            End If

27            strSQL = strSQL & " ) Where 序号 = 1"

28            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人列表", CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.Text), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
29            If rsPait Is Nothing Then
30                Set rsPait = gobjLiscomlib.CopyNewRec(rsTmp, True)
31            End If
32            Do While Not rsTmp.EOF
33                With rsPait
34                    .Filter = "HIS病人ID=" & rsTmp("HIS病人ID")
35                    If .RecordCount <= 0 Then
36                        .AddNew
37                        For j = 0 To .Fields.Count - 1
38                            .Fields(j).value = rsTmp.Fields(j).value
39                        Next
40                    End If
41                End With
42                rsTmp.MoveNext
43            Loop

              '在老版中去查找病人
44            strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.病人ID Order By a.申请时间 Desc) 序号, a.病人ID his病人ID, a.姓名," & vbCrLf & _
                     "                 Decode(a.性别, '1', '男', '2', '女', '9', '未知', '不区分') 性别, a.年龄, Decode(a.病人来源, 1, a.门诊号, 2, a.住院号, a.挂号单) 病历号,a.挂号单" & vbCrLf & _
                     "          From 检验标本记录 A" & vbCrLf & _
                     "          Where Nvl(a.是否质控品, 0) = 0 and a.病人ID is not null And a.申请时间 Between [1] And [2] "

45            If Trim(Me.cboDept.Text) <> "00-所有科室" Then
46                strSQL = strSQL & " and (a.申请科室ID=[3] or a.申请科室ID is null)"
47            End If

48            If Trim(Me.cbodor.Text) <> "00-所有" Then
49                strSQL = strSQL & " and (a.申请人=[4] or a.申请人 is null)"
50            End If

51            Select Case Trim(cboDiseases.Text)
              Case "传染病"
52                strSQL = strSQL & " and 0=1"
53            End Select

54            If chkVerifyDate.value = 1 Then
55                strSQL = strSQL & " and a.审核时间 between [5] and [6]"
56            End If

57            strSQL = strSQL & " ) Where 序号 = 1"

58            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人列表", CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.ItemData(cboDept.ListIndex)), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
59            If rsPait Is Nothing Then
60                Set rsPait = gobjLiscomlib.CopyNewRec(rsTmp, True)
61            End If
62            Do While Not rsTmp.EOF
63                With rsPait
64                    .Filter = "HIS病人ID=" & rsTmp("HIS病人ID")
65                    If .RecordCount <= 0 Then
66                        .AddNew
67                        For j = 0 To .Fields.Count - 1
68                            .Fields(j).value = rsTmp.Fields(j).value
69                        Next
70                    End If
71                End With
72                rsTmp.MoveNext
73            Loop

74            If Not rsPait Is Nothing Then
75                rsPait.Filter = ""
76                If rsPait.RecordCount > 0 Then rsPait.MoveFirst
77            End If
78        End If
79        Call gobjLiscomlib.SetDataToVSF(vsfPaitList, rsPait)    '将病人信息加载到列表

80        With vsfPaitList
81            .ColHidden(.ColIndex("HIS病人ID")) = True
82            .ColHidden(.ColIndex("序号")) = True
83            .ColHidden(.ColIndex("挂号单")) = True

              '默认选中第一个
84            If .Rows > 1 Then
85                Call vsfPaitList_AfterRowColChange(0, 0, 1, 0)
86            End If
87        End With

88        With Me.txtPaitKey
89            .SelStart = 0
90            .SelLength = Len(.Text)
91            .SetFocus
92        End With


93        Exit Sub
GetDeptPaits_Error:
94        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(GetDeptPaits)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
95        Err.Clear

End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim strFind As String
    Dim lngS As Long
    Dim strTxt As String


    If KeyAscii = vbKeyReturn Then
        With cboDept
            strFind = UCase(Trim(.Text))
            '按编码查找
            If IsNumeric(strFind) Then
                For i = 0 To .ListCount - 1
                    If .List(i) Like strFind & "-*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            Else
                '按简码查找
                For i = 0 To .ListCount - 1
                    lngS = InStr(.List(i), "[")
                    If lngS > 0 Then
                        strTxt = Mid(.List(i), lngS)
                    End If
                    If UCase(strTxt) = "[" & strFind & "]" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
                '按名称查找
                For i = 0 To .ListCount - 1
                    If .List(i) Like "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub cbodor_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim strFind As String
    Dim lngS As Long
    Dim strTxt As String
    If KeyAscii = vbKeyReturn Then
        With cbodor
            strFind = UCase(Trim(.Text))
            '按编码查找
            If IsNumeric(strFind) Then
                For i = 0 To .ListCount - 1
                    If .List(i) Like strFind & "-*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            Else
                '按简码查找
                For i = 0 To .ListCount - 1
                    lngS = InStr(.List(i), "[")
                    If lngS > 0 Then
                        strTxt = Mid(.List(i), lngS)
                    End If
                    If UCase(strTxt) = "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
                '按名称查找
                For i = 0 To .ListCount - 1
                    If .List(i) Like "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ConMenu_Browse_Find        '查找
        Call GetDeptPaits
    Case ConMenu_Browse_Print       '打印未打印报告
        Call PrintPaitReport(2, mlngPatientID, False)
    Case ConMenu_Browse_PrintAll    '打印所有报告
        Call PrintPaitReport(2, mlngPatientID, True)
    Case ConMenu_Browse_PrintView   '预览为打印报告
        Call PrintPaitReport(1, mlngPatientID, False)
    Case ConMenu_Browse_PrintViewAll    '预览所有
        Call PrintPaitReport(1, mlngPatientID, True)
    Case ConMenu_pop_Dept
        lblDept.Caption = "申请科室↓"
        InitDepts 0
    Case ConMenu_pop_DeptDistrict
        lblDept.Caption = "申请病区↓"
        InitDepts 1
    Case ConMenu_Appfor_ClincHelp   '诊疗参考
        Call ShowClincHelp
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99    '外挂功能执行
'        Call ExePlugIn(Control.Parameter, mlngKey)
    Case ConMenu_Browse_Exit       '退出
        Unload Me
    End Select
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-30
'功    能:  连续打印病人多份报告
'入    参:
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Sub PrintPaitReport(ByVal bytType As Byte, ByVal lngPaitID As Long, ByVal blnPrintAll As Boolean)
          Dim objLisPrint As Object
          Dim objForm As Object
          
1         On Error GoTo PrintPaitReport_Error

2         If objLisPrint Is Nothing Then Set objLisPrint = CreateObject("zlPublicLIS.clsLis")
          '先打印新版报告
3         If Not objLisPrint Is Nothing Then
4             Call objLisPrint.Init(gcnHisOracle)
5         End If

          '加载打印窗体
6         Set objForm = objLisPrint.GetForm()

         '调用打印对应的病人报告
7         Call objLisPrint.PrintLisReport(objForm, lngPaitID, mstrPatientGH, mlngPatientPage, 2, bytType, mblnDoctorShow, blnPrintAll)

8         Set objLisPrint = Nothing
          
          '打印三方微生物报告
9         Call beginPrint


10        Exit Sub
PrintPaitReport_Error:
11        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(PrintPaitReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
12        Err.Clear
End Sub

'----------三方微生物报告处理
Private Sub beginPrint()
    Dim strFileSource As String
    Dim lng报告ID As String
    Dim strArr() As String
    Dim i As Integer
    
    If mstrThirdReport <> "" Then
        If Left(mstrThirdReport, 4) = "<SP>" Then mstrThirdReport = Mid(mstrThirdReport, 5)
    Else
        Exit Sub
    End If
    strArr = Split(mstrThirdReport, "<SP>")
    For i = 0 To UBound(strArr)
        strFileSource = GetLisRptFile(strArr(i))
        lng报告ID = Split(strArr(i), ";")(0)
        Call FunFastPrint(strFileSource, lng报告ID)
    Next

End Sub

Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'功能：API调用快速打印PDF文件
'参数：strFile 文件路径
    Dim RetVal As Long
    Dim strSQL As String
    Dim ShExInfo As SHELLEXECUTEINFO
    
     On Error GoTo errH
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fMask = &H40
        .hWnd = 0
        .lpVerb = "print"
        .lpFile = strFile
        .lpParameters = ""
        .lpDirectory = vbNullChar
        .nShow = 2
    End With
    RetVal = ShellExecuteEx(ShExInfo)
    If RetVal = 0 Then
        Exit Sub
    End If
'    strSQL = "Zl_医嘱报告内容_Print(" & lngRptID & ",0)"
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
   Exit Sub
errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With picFilter
        .Left = Left
        .Top = Top
        .Width = Right - Left
    End With
    With picMain
        .Left = Left
        .Top = picFilter.Top + picFilter.Height
        .Width = Right - Left
        .Height = Bottom - .Top
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    mobjIDKind.object.ActiveFastKey
End Sub

Private Sub Form_Load()
      '功能创建工具条
          Dim cbrControl As CommandBarControl
          Dim cbrToolBar As CommandBar
          '-----------------------------------------------------
1         On Error GoTo Form_Load_Error

2         CommandBarsGlobalSettings.App = App
3         CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
4         CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
5         Me.cbrMain.VisualTheme = xtpThemeOffice2003
6         Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
7         With Me.cbrMain.Options
8             .ShowExpandButtonAlways = False
9             .ToolBarAccelTips = True
10            .AlwaysShowFullMenus = False
11            .IconsWithShadow = True    '放在VisualTheme后有效
12            .UseDisabledIcons = True
13            .LargeIcons = True
14            .SetIconSize True, 24, 24
15            .SetIconSize False, 16, 16
16        End With
17        Me.cbrMain.EnableCustomization False

          '-----------------------------------------------------
          '菜单定义
18        Me.cbrMain.ActiveMenuBar.Title = "菜单"
19        Me.cbrMain.ActiveMenuBar.Visible = False
20        Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
21        cbrToolBar.ShowTextBelowIcons = False
22        cbrToolBar.EnableDocking xtpFlagStretched
23        With cbrToolBar.Controls

24            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "查找(&F5)"): cbrControl.BeginGroup = True
25            Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "打印未打印报告(&F2)")
26            cbrControl.Style = xtpButtonIconAndCaption
27            With cbrControl.CommandBar.Controls
28                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintAll, "打印所有  ")
29                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "打印设置  ")
30                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "重置打印  ")
31                cbrControl.Visible = InStr(mstrPrivs, "重置自助机报告打印次数") > 0
32            End With
33            Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_PrintView, "预览未打印报告")
34            cbrControl.Style = xtpButtonIconAndCaption
35            With cbrControl.CommandBar.Controls
36                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintViewAll, "预览所有  ")
37            End With
38            Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "诊疗参考")
39            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "退出"): cbrControl.BeginGroup = True
40        End With

          '创建插件按钮
41        Call CreatePlugInButton(cbrToolBar)

42        For Each cbrControl In cbrToolBar.Controls
43            If cbrControl.Type = xtpControlButton Then
44                cbrControl.Style = xtpButtonIconAndCaption
45            End If
46        Next

          '快键绑定
47        With Me.cbrMain.KeyBindings
48            .Add 0, VK_F2, ConMenu_Browse_Print
49            .Add 0, VK_F5, ConMenu_Browse_Find
50        End With


          '初始化IDKind
51        If mobjIDKind Is Nothing Then
52            Set mobjIDKind = NewControl(Me, "zlLisControl.ucLisIDKind", "ucLisIDKind", picIDKIND)
53            If mobjIDKind Is Nothing Then
54                Me.picIDKIND.Visible = False
55            End If
56            picIDKIND.BorderStyle = 0
57        End If

58        dtpE.value = gobjLiscomlib.comcurrdate
59        dtpS.value = dtpE.value - 7

60        Call gobjLiscomlib.vfgSetting(0, Me.vsfPaitList, "姓名,2000,1;性别,800,1;年龄,800,1;病历号,1000,1")

61        txtRptCount.Text = Val(ComGetPara(Sel_Lis_DB, "检验报告查看份数", 2500, 2500, "7"))

          '是否显示传染病筛选框
62        cboDiseases.Enabled = InStr(mstrPrivs, "查看传染病报告") > 0
63        Me.cboDiseases.ListIndex = 2

64        txtPaitKey.TabIndex = 0
65        cboDept.TabIndex = 1
66        cbodor.TabIndex = 2
67        dtpS.TabIndex = 3
68        dtpE.TabIndex = 4
69        dtpVS.TabIndex = 5
70        dtpVE.TabIndex = 6
71        cboDiseases.TabIndex = 7
72        txtRptCount.TabIndex = 8


73        Exit Sub
Form_Load_Error:
74        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(Form_Load)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
75        Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mstrPrivs = ""
    mstrThirdReport = ""
    
    For i = 1 To uclSampleReport.Count - 1
        Unload uclSampleReport(i)
    Next
    Set mobjIDKind = Nothing
    
    
    Call ComSetPara(Sel_Lis_DB, "检验报告查看份数", Val(txtRptCount.Text), 2500, 2500)
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-26
'功    能:  根据选择的病人获取病人的检验报告
'入    参:
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Sub GetPaitReport(ByVal lngPaitID As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsRpt As ADODB.Recordset
          Dim lngUclCount As Long
          Dim i As Integer
          Dim j As Integer
          Dim strWhere As String
          Dim strDept As String
          Dim strDor As String
          Dim lngS As Long
          Dim lngE As Long

          '新版报告
1         On Error GoTo GetPaitReport_Error

2         gobjLiscomlib.ShowFlash "正在加载报告,请稍候...", Me

          '先卸载上一次的控件
3         picScroll.Visible = False
4         For i = 1 To Me.uclSampleReport.Count - 1
5             Call uclSampleReport(i).UnloadCrl   '卸载控件中使用的对象
6             Unload uclSampleReport(i)
7         Next


8         strSQL = "select * from (select a.ID,a.微生物,0 结果次数,a.阳性报告,25 版本,a.诊断,a.备注,a.申请时间,a.核收时间 from 检验报告记录 A where a.HIS病人ID=[1] and  a.申请时间 between [3] and [4] [条件] and a.审核人 is not null order by a.申请时间 desc) where rownum<=[2]  "


9         If Trim(Me.cboDept.Text) <> "00-所有科室" Then
10            strDept = Me.cboDept.Text
11            lngS = InStr(strDept, "-") + 1
12            lngE = InStr(strDept, "[")
13            strDept = Mid(strDept, lngS, lngE - lngS)
14            strWhere = strWhere & " and (a.申请科室=[5] or a.申请科室 is null)"
15        End If

16        If Trim(Me.cbodor.Text) <> "00-所有" Then
17            strDor = Me.cbodor.Text
18            lngS = InStr(strDor, "-") + 1
19            lngE = InStr(strDor, "[")
20            strDor = Mid(strDor, lngS, lngE - lngS)
21            strWhere = strWhere & " and (a.申请人=[6] or a.申请人 is null)"
22        End If

23        Select Case Trim(cboDiseases.Text)
          Case "传染病"
24            strWhere = strWhere & " and nvl(a.是否传染病,0)=1"
25        Case "非传染病"
26            strWhere = strWhere & " and nvl(a.是否传染病,0)=0"
27        End Select

28        If chkVerifyDate.value = 1 Then
29            strSQL = strSQL & " and a.审核时间 between [7] and [8]"
30        End If

31        strSQL = Replace(strSQL, "[条件]", strWhere)
32        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验报告记录", lngPaitID, Val(txtRptCount.Text), CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), strDept, strDor, CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
33        If rsRpt Is Nothing Then
34            Set rsRpt = gobjLiscomlib.CopyNewRec(rsTmp, True)
35        End If
36        Do While Not rsTmp.EOF
37            With rsRpt
38                .AddNew
39                For j = 0 To .Fields.Count - 1
40                    .Fields(j).value = rsTmp.Fields(j).value
41                Next
42            End With
43            rsTmp.MoveNext
44        Loop

          '老版报告
45        strSQL = "select * from (select a.ID,a.微生物标本 微生物,a.报告结果 结果次数,1 阳性报告,10 版本,'' 诊断,'' 备注,a.申请时间,a.核收时间  from 检验标本记录 A where a.病人ID=[1] and  a.申请时间 between [3] and [4] [条件] and a.审核人 is not null order by a.申请时间 desc) where rownum<=[2]"

46        strWhere = ""
47        If Trim(Me.cboDept.Text) <> "00-所有科室" Then
48            strWhere = " and (a.申请科室ID=[5] or a.申请科室ID is null)"
49        End If

50        If Trim(Me.cbodor.Text) <> "00-所有" Then
51            strWhere = strWhere & " and (a.申请人=[6] or a.申请人 is null)"
52        End If


53        If chkVerifyDate.value = 1 Then
54            strWhere = strWhere & " and a.审核时间 between [7] and [8]"
55        End If

56        Select Case Trim(cboDiseases.Text)
          Case "传染病"
57            strWhere = strWhere & " and 0=1"
58        End Select

59        strSQL = Replace(strSQL, "[条件]", strWhere)

60        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验报告记录", lngPaitID, Val(txtRptCount.Text), CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.ItemData(cboDept.ListIndex)), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))

61        If rsRpt Is Nothing Then
62            Set rsRpt = gobjLiscomlib.CopyNewRec(rsTmp, True)
63        End If
64        Do While Not rsTmp.EOF
65            With rsRpt
66                .AddNew
67                For j = 0 To .Fields.Count - 1
68                    .Fields(j).value = rsTmp.Fields(j).value
69                Next
70            End With
71            rsTmp.MoveNext
72        Loop
73        rsRpt.Filter = ""
74        If rsRpt.RecordCount > 0 Then
75            rsRpt.MoveFirst
76            picScroll.Visible = True
77        End If
78        rsRpt.Sort = "申请时间 desc"
79        mstrThirdReport = ""
80        Do While Not rsRpt.EOF
81            If lngUclCount >= Val(txtRptCount.Text) Then Exit Do
82            Call ShowPaitReport(Me, mblnDoctorShow, lngPaitID, Val(rsRpt("ID") & ""), Val(rsRpt("版本") & ""), Val(rsRpt("微生物") & ""), Val(rsRpt("阳性报告") & ""), rsRpt("诊断") & "", rsRpt("备注") & "", Val(rsRpt("结果次数") & ""), lngUclCount, CDate(Format(rsRpt("核收时间") & "", "yyyy/mm/dd hh:mm:ss")))
83            rsRpt.MoveNext
84        Loop


          '设置滚动条
85        Me.vsfScroll.Rows = picScroll.Height / 225

86        gobjLiscomlib.StopFlash

87        Exit Sub
GetPaitReport_Error:
88        gobjLiscomlib.StopFlash
89        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "执行(GetPaitReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
90        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-07-27
'功    能:  显示报告
'入    参:
'           objFrm          调用窗体
'           mblnDoctorShow  是否是医生站调用
'           lngPaintID      病人ID
'           lngSampleID     标本ID
'           intVersion      报告版本，25=新版LIS，10=老版LIS
'           intSampleType   是否是微生物报告，0=普通报告，1=微生物报告
'           intPositive     报告类型，1=药敏报告，3=PDF报告，其他=阴性报告
'           strDiagnosis    诊断
'           strResult       备注
'           intCount        老版LIS结果次数
'           dteSampleTime   标本核收时间

'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Public Function ShowPaitReport(objFrm As Object, ByVal blnDoctorShow As Boolean, ByVal lngPaintID As Long, ByVal lngSampleID As Long, ByVal intVersion As Long, _
                                ByVal intSampleType As Integer, Optional ByVal intPositive As Integer, _
                                Optional ByVal strDiagnosis As String, Optional ByVal strResult As String, _
                                Optional ByVal intCount As Integer, Optional ByRef lngUclCount As Long, Optional ByVal dteSampleTime As Date) As Long
          Dim lngHeight As Long
          Dim strThirdReport As String

1         On Error GoTo ShowPaitReport_Error

2         If lngUclCount = 0 Then
              '加载报告
3             lngHeight = uclSampleReport(lngUclCount).GetSampleReport(Me, blnDoctorShow, lngPaintID, lngSampleID, intVersion, intSampleType, intPositive, strDiagnosis, strResult, intCount, dteSampleTime, mstrPrivs, strThirdReport)
4         Else
5             Load uclSampleReport(lngUclCount)
6             lngHeight = uclSampleReport(lngUclCount).GetSampleReport(Me, blnDoctorShow, lngPaintID, lngSampleID, intVersion, intSampleType, intPositive, strDiagnosis, strResult, intCount, dteSampleTime, mstrPrivs, strThirdReport)
7         End If
8         If strThirdReport <> "" Then
9             mstrThirdReport = mstrThirdReport & "<SP>" & strThirdReport
10        End If
          '加自定义报告控件放在vsf中，以便随之滚动
          
11        With uclSampleReport(lngUclCount)
12            If lngUclCount = 0 Then
13                .Left = 0
14                .Top = 0
15                .Width = picScroll.Width
16                .Height = lngHeight
17                Me.picScroll.Height = .Height
18            Else
19                .Left = 0
20                .Top = uclSampleReport(lngUclCount - 1).Top + uclSampleReport(lngUclCount - 1).Height + 200
21                .Width = picScroll.Width
22                .Height = lngHeight
23                .Visible = True
24                Me.picScroll.Height = Me.picScroll.Height + .Height + 200
25            End If
26        End With

27        lngUclCount = lngUclCount + 1


28        Exit Function
ShowPaitReport_Error:
29        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(ShowPaitReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
30        Err.Clear

End Function

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picPaitList
        Rightcoll.Add Me.picPaitReport
        Call SplitWE(LeftColl, Me.fraWE, Rightcoll, X, 1000)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub

Private Sub lblDept_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbrMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "申请科室")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "申请病区")
    End With
    vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
    vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen picFilter.hWnd, vPoint

    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub

Private Sub mobjIDKind_ObjectEvent(Info As EventInfo)
    Select Case Info
        Case "ReadCard"
                txtPaitKey.Text = IIf(Info.EventParameters(1).value = 0, Info.EventParameters(0).value, Info.EventParameters(1).value)
                Call GetDeptPaits
                txtPaitKey.SelStart = 0
                txtPaitKey.SelLength = Len(txtPaitKey.Text)
    End Select
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With picPaitList
        .Left = 0
        .Top = 0
        .Height = Me.picMain.Height
    End With
    With fraWE
        .Left = picPaitList.Left + picPaitList.Width
        .Top = 0
        .Height = Me.picMain.Height
    End With
    With picPaitReport
        .Left = fraWE.Left + fraWE.Width
        .Top = 0
        .Width = Me.picMain.Width - .Left
        .Height = Me.picMain.Height
    End With
End Sub

Private Sub picPaitList_Resize()
    On Error Resume Next
    With vsfPaitList
        .Left = 0
        .Top = 0
        .Width = Me.picPaitList.Width
        .Height = Me.picPaitList.Height
    End With
End Sub

Private Sub picPaitReport_Resize()
    On Error Resume Next
    With vsfScroll
        .Left = 0
        .Top = 0
        .Width = Me.picPaitReport.Width
        .Height = Me.picPaitReport.Height
    End With
    With picScroll
        .Left = 0
        .Top = -vsfScroll.TopRow * vsfScroll.RowHeight(0)
        .Width = Me.picPaitReport.Width - 300
    End With
End Sub

Private Sub picScroll_Resize()
    Dim i As Integer
    
    For i = 0 To uclSampleReport.Count - 1
        With uclSampleReport(i)
            .Left = 0
            .Width = Me.picScroll.Width
        End With
    Next
End Sub

Private Sub txtPaitKey_GotFocus()
    With Me.txtPaitKey
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtPaitKey_KeyPress(KeyAscii As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strFind As String
          Dim strWhere As String


1         On Error GoTo txtPaitKey_KeyPress_Error

2         If KeyAscii = vbKeyReturn Then
              '通过输入的内容查找病人ID
3             strFind = Trim(txtPaitKey.Text)
4             If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then    '病人ID
5                 strWhere = " and a.HIS病人ID = [2] "
6                 strFind = Mid(strFind, 2)
7             ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then    '住院号
8                 strWhere = " and a.住院号 = [1] "
9                 strFind = Mid(strFind, 2)
10            ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then    '门诊号
11                strWhere = " and a.门诊号 = [1] "
12                strFind = Mid(strFind, 2)
13            ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then    '挂号单
14                strWhere = " and a.挂号单 = [1] "
15            ElseIf Left(strFind, 1) = "/" Then    '收费单据号
16                strWhere = " and a.收费单号 = [1] "
17            End If
18            strSQL = "      Select His病人id" & vbNewLine & _
                     "       From 检验申请组合" & vbNewLine & _
                     "       Where His病人id = [2] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His病人id" & vbNewLine & _
                     "       From 检验申请组合" & vbNewLine & _
                     "       Where 住院号 = [1] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His病人id" & vbNewLine & _
                     "       From 检验申请组合" & vbNewLine & _
                     "       Where 门诊号 = [1] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His病人id" & vbNewLine & _
                     "       From 检验申请组合" & vbNewLine & _
                     "       Where 挂号单 = [1]" & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His病人id" & vbNewLine & _
                     "       From 检验申请组合" & vbNewLine & _
                     "       Where 样本条码 = [1]"
19            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "病人ID", strFind, Val(strFind))
20            If Not rsTmp.EOF Then
21                txtPaitKey.Text = rsTmp("HIS病人ID")
22                Call GetDeptPaits   '再通过ID查找病人
23            End If
24        End If


25        Exit Sub
txtPaitKey_KeyPress_Error:
26        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(txtPaitKey_KeyPress)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear
End Sub

Private Sub txtRptCount_Change()
    upd.value = Val(txtRptCount.Text)
End Sub

Private Sub txtRptCount_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub upd_DownClick()
    txtRptCount.Text = upd.value
End Sub

Private Sub upd_UpClick()
    txtRptCount.Text = upd.value
End Sub

Private Sub vsfPaitList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    With Me.vsfPaitList
        If NewRow < 1 Then Exit Sub
        If .ColIndex("HIS病人ID") < 0 Or .ColIndex("挂号单") < 0 Then Exit Sub
        mlngPatientID = Val(.TextMatrix(NewRow, .ColIndex("HIS病人ID")))
        mstrPatientGH = .TextMatrix(NewRow, .ColIndex("挂号单"))
        If mstrPatientGH = "0" Then mstrPatientGH = ""
        Call GetPaitReport(Val(.TextMatrix(NewRow, .ColIndex("HIS病人ID"))))
    End With
End Sub

Private Sub vsfScroll_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Me.picScroll.Top = -vsfScroll.TopRow * vsfScroll.RowHeight(0)
End Sub


Private Function GetDeptDor(Optional ByVal lngDeptID As Long) As ADODB.Recordset
      '功能           传入科室或病区返回对应的医生记录集
      '参数
      '               lngDeptID 科室ID或病区ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo GetDeptDor_Error

2         strSQL = "Select distinct b.id, b.编号,b.姓名,b.简码" & vbNewLine & _
                   "From 部门人员 A, 人员表 B, 部门表 C,人员性质说明 D" & vbNewLine & _
                   "Where A.人员id = B.Id And A.部门id = C.Id And b.id=D.人员ID And (C.撤档时间 Is Null Or C.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) "
3         If lngDeptID <> 0 Then
4             strSQL = strSQL & "and c.id = [1] "
5         End If
          
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID)
7         With cbodor
8             .Clear
9             .AddItem "00-所有"
10            .ItemData(.NewIndex) = 0
11            Do Until rsTmp.EOF
12                .AddItem rsTmp!编号 & "-" & rsTmp!姓名 & "[" & rsTmp!简码 & "]"
13                .ItemData(.NewIndex) = rsTmp!ID
14                rsTmp.MoveNext
15            Loop
16            If .ListCount > 0 Then .ListIndex = 0

17        End With


18        Exit Function
GetDeptDor_Error:
19        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(GetDeptDor)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
20        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-19
'功    能:  显示诊疗参考
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim lngSampleID As Long
          Dim lngVer As Long

1         On Error GoTo ShowClincHelp_Error

'2         With Me.vsfLeft
'3             If .Row < 1 Then
'4                 MsgBox "请选中一份报告", vbInformation, gSysInfo.AppName
'5                 Exit Sub
'6             End If
'7             If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
'8                 MsgBox "请选中一份报告", vbInformation, gSysInfo.AppName
'9                 Exit Sub
'10            End If
'11            lngSampleID = Val(.TextMatrix(.Row, .ColIndex("ID")))
'12            lngVer = Val(.TextMatrix(.Row, .ColIndex("版本")))
'13        End With
'
'14        Call funShowClincHelp(Me, lngSampleID, lngVer)


15        Exit Sub
ShowClincHelp_Error:
16        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "执行(ShowClincHelp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear

End Sub


