VERSION 5.00
Begin VB.Form frm住院 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   750
   ClientLeft      =   6300
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   937.5
   ScaleMode       =   0  'User
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt人员类别 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   465
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txt总费用 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   60
      Width           =   750
   End
   Begin VB.TextBox txt余额 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   750
   End
   Begin VB.TextBox txt详细 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   1965
   End
   Begin VB.TextBox txt说明 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   435
      Width           =   1965
   End
   Begin VB.TextBox txt名称 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "全髋关节置换术（双侧）"
      Top             =   435
      Width           =   2115
   End
   Begin VB.TextBox txt类别 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "单病种"
      Top             =   450
      Width           =   555
   End
   Begin VB.Label lab手术 
      AutoSize        =   -1  'True
      Caption         =   "丙类"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Left            =   255
      TabIndex        =   12
      Top             =   450
      Width           =   450
   End
   Begin VB.Label lab人员类别 
      AutoSize        =   -1  'True
      Caption         =   "人员"
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   360
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   465
      X2              =   1240
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   1650
      X2              =   2460
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "费用"
      Height          =   180
      Left            =   1260
      TabIndex        =   10
      Top             =   60
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "余额"
      Height          =   180
      Left            =   2505
      TabIndex        =   9
      Top             =   60
      Width           =   360
   End
   Begin VB.Line Line3 
      DrawMode        =   1  'Blackness
      X1              =   2925
      X2              =   3735
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Line Line4 
      DrawMode        =   1  'Blackness
      X1              =   3975
      X2              =   5945
      Y1              =   356.25
      Y2              =   356.25
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "详"
      Height          =   180
      Left            =   3780
      TabIndex        =   8
      Top             =   60
      Width           =   180
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   -15
      X2              =   8530
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      X1              =   -15
      X2              =   8530
      Y1              =   468.75
      Y2              =   468.75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "详"
      Height          =   180
      Left            =   3780
      TabIndex        =   7
      Top             =   495
      Width           =   180
   End
   Begin VB.Line lin病种 
      BorderColor     =   &H00000000&
      X1              =   1650
      X2              =   3795
      Y1              =   843.75
      Y2              =   843.75
   End
   Begin VB.Line Line8 
      X1              =   3975
      X2              =   6120
      Y1              =   825
      Y2              =   825
   End
End
Attribute VB_Name = "frm住院"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiID          As Long
Private mvarRecId           As Variant
Private mvarKeyId           As Variant
Private mstrReserve         As String

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
Private var病种             As typ_病种信息

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
    Dim dbl总费用           As Double
    Dim dbl特殊材料总费用   As Double
    Dim dbl限制总额         As Double
    Dim lng病种ID           As Long
    Dim intInsure           As Integer
    
    DoEvents
    Me.Show
    rtn = SetWindowPos(Me.hWnd, -1, CurrentX, CurrentY, 0, 0, 3)
    
    '检测是否丙类手术
    gstrSql = "select 1 from 病人手麻记录 A,大连_丙类手术 B where A.诊疗项目ID = B.手术ID And A.病人ID=[1] AND A.主页ID=[2]"
    lab手术.Visible = Not ChkRsState(gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)))
    '读取人员类别
    gstrSql = "select A.险类,A.在职,B.名称 from 保险帐户 A ,保险人群 B where A.在职=B.序号 AND A.险类=B.险类 And A.病人ID=[1]"
    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID)
    If ChkRsState(rsTmp) Then
        txt人员类别.Text = ""
        Me.Hide
    Else
        intInsure = rsTmp!险类
        txt人员类别.Text = rsTmp!名称
        If ChkRsState(rsTmp) Then
            Me.Height = 370
        Else
            '读取病种信息
            gstrSql = "select C.ID,B.编码,DECODE(C.类别,1,'慢性病',2,'特种病',3,'单病种','普通病') AS 类别,B.名称,B.说明" & vbCrLf & _
                  "from (Select 诊断描述 From 病人诊断记录 where ID IN (select Max(ID) as ID from 病人诊断记录 where 诊断类型=1 AND 病人ID = [1] And 主页ID = [2] group by 病人ID,主页ID )) A,疾病编码目录 B,保险病种 C" & vbCrLf & _
                  "where zl_split(zl_split(A.诊断描述,')',0),'(',1)=B.编码 AND B.编码=C.编码(+)"
            Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
            
            If ChkRsState(rsTmp) Then
                Me.Height = 370
                '已使用费用
                gstrSql = "select nvl(sum(实收金额),0) as 金额 from 住院费用记录 where 病人ID = [1] And 主页ID = [2]"
                dbl总费用 = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)).Fields(0)
                '读取科室限额
                gstrSql = "select  nvl(限制金额,0) from 大连_科室限额 where 科室ID in (select nvl(出院科室ID,入院科室ID) " & _
                          "from 病案主页 where 病人ID=[1] And 主页ID=[2])"
                Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
                If ChkRsState(rsTmp) Then
                    dbl限制总额 = 0
                Else
                    dbl限制总额 = rsTmp.Fields(0)
                End If
                '总费用
                txt总费用.Text = Format(dbl总费用, "0.00")
                txt详细.Text = "  科：" & Format(dbl限制总额, "0")
                txt余额.Text = Format(dbl限制总额 + dbl特殊材料总费用 - dbl总费用, "0.00")
                txt余额.ForeColor = IIf(Val(txt余额.Text) < 0, col单病种, col慢性病)
            Else
                Me.Height = 730
                var病种.color = Decode(rsTmp!类别, "慢性病", col慢性病, "特种病", col特种病, "单病种", col单病种, col普通病)
                var病种.str编码 = "" & rsTmp!编码
                var病种.str类别 = "" & rsTmp!类别
                var病种.str名称 = "" & rsTmp!名称
                var病种.str说明 = "" & rsTmp!说明
                txt类别.ForeColor = var病种.color
                txt类别.Text = var病种.str类别
                txt名称.Text = var病种.str名称
                txt说明.Text = var病种.str说明
                '病种限额
                lng病种ID = Val("" & rsTmp!ID)
                '检测当前病种是否有限额
                gstrSql = "select nvl(限制金额,0) as 金额 from 大连_病种限额 where 险类=[1] And 病种ID=[2]"
                Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, intInsure, lng病种ID)
                '已使用费用
                gstrSql = "select nvl(sum(实收金额),0) as 金额 from 住院费用记录 where 病人ID = [1] And 主页ID = [2]"
                dbl总费用 = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId)).Fields(0)
                If ChkRsState(rsTmp) Then
                    '读取科室限额
                    gstrSql = "select  nvl(限制金额,0) from 大连_科室限额 where 科室ID in (select nvl(出院科室ID,入院科室ID) " & _
                              "from 病案主页 where 病人ID=[1] And 主页ID=[2])"
                    Set rsTmp = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngPatiID, Val(mvarRecId))
                    If ChkRsState(rsTmp) Then
                        dbl限制总额 = 0
                    Else
                        dbl限制总额 = rsTmp.Fields(0)
                    End If
                    '总费用
                    txt总费用.Text = Format(dbl总费用, "0.00")
                    txt详细.Text = "  科：" & Format(dbl限制总额, "0")
                    txt余额.Text = Format(dbl限制总额 + dbl特殊材料总费用 - dbl总费用, "0.00")
                    txt余额.ForeColor = IIf(Val(txt余额.Text) < 0, col单病种, col慢性病)
                Else
                    '读取病种限额
                    dbl限制总额 = rsTmp!金额
                    gstrSql = "select nvl(sum(限制金额),0) as 金额 from 住院费用记录 A,大连_病种材料 B " & _
                              "where A.收费细目ID = B.收费ID And  病人ID = [2] And 主页ID = [3]  And 险类=[1] And B.病种ID=[4]"
                    dbl特殊材料总费用 = gDatabase.OpenSQLRecord(gstrSql, Me.Caption, intInsure, mlngPatiID, Val(mvarRecId), lng病种ID).Fields(0)
                    '总费用
                    txt总费用.Text = Format(dbl总费用, "0.00")
                    txt详细.Text = "  单：" & Format(dbl限制总额, "0") & ";特：" & Format(dbl特殊材料总费用, "0")
                    txt余额.Text = Format(dbl限制总额 - dbl总费用, "0.00")
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    Me.Top = 0
    Me.Left = 6300
End Sub
 
