VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAppRequestFilter 
   BorderStyle     =   0  'None
   Caption         =   "记录过滤"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdOK 
      Caption         =   "过滤(&O)"
      Height          =   350
      Left            =   2505
      TabIndex        =   9
      Top             =   3795
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3660
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.CheckBox chkShowSet 
         Caption         =   "显示已处理记录"
         Height          =   375
         Left            =   420
         TabIndex        =   16
         Top             =   255
         Width           =   1665
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "按处理时间查找"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   1365
         Width           =   1665
      End
      Begin VB.ComboBox cbo复诊方式 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3045
         Width           =   2085
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "按登记时间查找"
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.ComboBox cbo处理人 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2640
         Width           =   2085
      End
      Begin VB.ComboBox cbo登记人 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2235
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Index           =   0
         Left            =   2355
         TabIndex        =   2
         Top             =   1005
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42338
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   1005
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42328
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Index           =   1
         Left            =   2355
         TabIndex        =   12
         Top             =   1770
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42338
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   1770
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42328
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2115
         TabIndex        =   15
         Top             =   1830
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复诊方式"
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   3105
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "处理人"
         Height          =   180
         Left            =   465
         TabIndex        =   7
         Top             =   2700
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   2115
         TabIndex        =   5
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   465
         TabIndex        =   4
         Top             =   2295
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmAppRequestFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mfrmParent As Object

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Public Sub SetForm(frmParent As Object)
    Set mfrmParent = frmParent
End Sub

Private Sub chkShowSet_Click()
    If chkShowSet.Value = 1 Then
        cbo处理人.Enabled = True
        chkDate(1).Enabled = True
        dtpBegin(1).Enabled = True
        dtpEnd(1).Enabled = True
    Else
        cbo处理人.Enabled = False
        chkDate(1).Enabled = False
        chkDate(1).Value = False
        dtpBegin(1).Enabled = False
        dtpEnd(1).Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    Call mfrmParent.RefreshRecord
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Function zlGetFullFieldsTable(Optional strTableName As String = "门诊费用记录", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取一张数据表中的字段.类似于Select Id,....
    '入参：bytHistory-0-不包含历史数据,1-仅包含历史数据,2-两都都包含( select * from tablename Union select * from Htablename)
    '      strWhere-条件
    '      blnSubTable-是否子表
    '      strAliasName-别名
    '出参：
    '返回：select ID ... From tableName Union ALL
    '编制：刘兴洪
    '日期：2010-03-10 11:19:11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '无
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '仅历史
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '两者都包含
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
    
End Function

Private Function GetPersonnel(str性质 As String, Optional blnBaseInfo As Boolean) As ADODB.Recordset
'功能：读取指定性质的人员列表
    Dim strSQL As String
    On Error GoTo errH
    
    If str性质 <> "" Then
        If blnBaseInfo Then
            strSQL = "Select a.id,a.编号,a.简码,a.姓名 From 人员表 a,人员性质说明 b" & _
            " Where a.ID = b.人员ID And b.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by a.简码"
        Else
            strSQL = "Select a.Id, a.编号, a.姓名, a.简码, a.身份证号, a.出生日期, a.性别, a.民族, a.工作日期, a.办公室电话, a.电子邮件, a.执业类别, a.执业范围, " & _
                    "a.管理职务, a.专业技术职务, a.聘任技术职务, a.学历, a.所学专业, a.留学时间, a.留学渠道, a.接受培训, a.科研课题, a.个人简介, a.建档时间, " & _
                    "a.撤档时间, a.撤档原因, a.别名, a.站点 From 人员表 a,人员性质说明 b" & _
            " Where a.ID = b.人员ID And b.人员性质=[1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by a.简码"
        End If
    Else
        If blnBaseInfo Then
            strSQL = "Select id,编号,简码,姓名 From 人员表 A" & _
            " Where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by 简码"
        Else
            strSQL = zlGetFullFieldsTable("人员表", 0, "", False) & _
            " Where (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by 简码"
        End If
    End If
    Set GetPersonnel = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, str性质)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：加载基础数据
    '编制：刘尔旋
    '日期：2016-01-11
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset

    Set rsTemp = GetPersonnel("", True)

    cbo登记人.Clear
    cbo登记人.AddItem "所有登记人-"
    cbo登记人.ListIndex = 0
    If rsTemp.RecordCount > 0 Then
        Call rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            cbo登记人.AddItem rsTemp!简码 & "-" & rsTemp!姓名
            If Nvl(rsTemp!姓名) = UserInfo.姓名 Then cbo登记人.ListIndex = cbo登记人.NewIndex
            rsTemp.MoveNext
        Next
    End If
    
    cbo处理人.Clear
    cbo处理人.AddItem "所有处理人-"
    cbo处理人.ListIndex = 0
    If rsTemp.RecordCount > 0 Then
        Call rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            cbo处理人.AddItem rsTemp!简码 & "-" & rsTemp!姓名
            If Nvl(rsTemp!姓名) = UserInfo.姓名 Then cbo处理人.ListIndex = cbo处理人.NewIndex
            rsTemp.MoveNext
        Next
    End If
    
    cbo复诊方式.Clear
    cbo复诊方式.AddItem "所有方式-"
    cbo复诊方式.ListIndex = 0
    cbo复诊方式.AddItem "1-按疗程复诊"
    cbo复诊方式.AddItem "2-按月复诊"
    cbo复诊方式.AddItem "3-按周复诊"
    cbo复诊方式.AddItem "4-按天复诊"
    
    dtpBegin(0).Value = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd 00:00:00")
    dtpEnd(0).Value = Format(gobjDatabase.CurrentDate + 1, "yyyy-mm-dd 23:59:59")
    dtpBegin(1).Value = Format(gobjDatabase.CurrentDate - 7, "yyyy-mm-dd 00:00:00")
    dtpEnd(1).Value = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd 23:59:59")
    
    LoadData = True
End Function

