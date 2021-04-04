VERSION 5.00
Begin VB.Form Frm医保对码_资阳 
   Caption         =   "医保项目对码"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7125
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd确定 
      Appearance      =   0  'Flat
      Caption         =   "关    闭"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt标志 
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1200
      TabIndex        =   25
      Top             =   3080
      Width           =   1095
   End
   Begin VB.CommandButton cmd查询 
      Appearance      =   0  'Flat
      Caption         =   "查   询"
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmd申报 
      Appearance      =   0  'Flat
      Caption         =   "申    报"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmd费别 
      Height          =   300
      ItemData        =   "Frm医保对码_资阳.frx":0000
      Left            =   4560
      List            =   "Frm医保对码_资阳.frx":000D
      TabIndex        =   21
      Top             =   2595
      Width           =   1095
   End
   Begin VB.ComboBox cmd类别 
      Height          =   300
      ItemData        =   "Frm医保对码_资阳.frx":0023
      Left            =   4560
      List            =   "Frm医保对码_资阳.frx":0030
      TabIndex        =   20
      Top             =   160
      Width           =   1095
   End
   Begin VB.TextBox txt费用项目 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2595
      Width           =   2175
   End
   Begin VB.TextBox txt规格 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2115
      Width           =   2175
   End
   Begin VB.TextBox txt单价 
      Height          =   270
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1635
      Width           =   2175
   End
   Begin VB.TextBox txt单位 
      Height          =   270
      Left            =   1200
      TabIndex        =   16
      Top             =   1635
      Width           =   2175
   End
   Begin VB.TextBox txt别名 
      Height          =   270
      Left            =   4560
      TabIndex        =   15
      Top             =   1155
      Width           =   2175
   End
   Begin VB.TextBox txt简码 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1155
      Width           =   2175
   End
   Begin VB.TextBox txt英文名称 
      Height          =   270
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox txt中文名称 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox txt编号 
      Height          =   270
      Left            =   1200
      TabIndex        =   11
      Top             =   200
      Width           =   2175
   End
   Begin VB.Label lbl标志 
      Caption         =   "启用标志"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lbl费用类别 
      AutoSize        =   -1  'True
      Caption         =   "费用类别"
      Height          =   180
      Left            =   3720
      TabIndex        =   10
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lbl费用项目 
      AutoSize        =   -1  'True
      Caption         =   "费用项目"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lbl规格 
      AutoSize        =   -1  'True
      Caption         =   "规    格"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lbl单价 
      AutoSize        =   -1  'True
      Caption         =   "单    价"
      Height          =   180
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl单位 
      AutoSize        =   -1  'True
      Caption         =   "售价单位"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl别名 
      AutoSize        =   -1  'True
      Caption         =   "别    名"
      Height          =   180
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label lbl简码 
      AutoSize        =   -1  'True
      Caption         =   "简    码"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label Lbl英文名称 
      AutoSize        =   -1  'True
      Caption         =   "英文名称"
      Height          =   180
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Lbl中文名称 
      AutoSize        =   -1  'True
      Caption         =   "中文名称"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lbl类别 
      AutoSize        =   -1  'True
      Caption         =   "类    别"
      Height          =   180
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lbl编号 
      AutoSize        =   -1  'True
      Caption         =   "编    号"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Frm医保对码_资阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '入出参数,医保项目DetailCode
Private mrsDetail As ADODB.Recordset, mrsTMP As ADODB.Recordset
Private mblnOK As Boolean
Private mint中心 As Integer
Private mint险类 As Integer
Private mintID As Integer

Public Function GetCode(strCode As String, ByVal int中心 As Integer, ByVal int险类 As Integer) As Boolean
'功能：获得一个收费项目的医保编码
'参数：strCode 既作为输入参数，又输出
'返回：成功返回True
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, mrsTMP As ADODB.Recordset
    
    mblnOK = False
    mint中心 = int中心
    
    On Error GoTo ErrH
    
    Set mrsTMP = New ADODB.Recordset
    mrsTMP.CursorLocation = adUseClient
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient

    mint险类 = int险类
    strSQL = "Select * from 医保支付项目 Where 险类=[1] And 中心=[2] and 项目编码=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, mint中心, strCode)
    If rsTmp.EOF Then
        gstrSQL = "select A.ID,A.编码,decode(A.类别,'J','服务','5','药品','6','药品','7','药品','诊疗') as 类别," & _
                  "A.名称 As 中文名称,'' as 英文名称, " & _
                  "zlspellcode(A.名称) as 简码,substrb(A.名称,1,40) as 别名,substrb(A.计算单位,1,20) as 计算单位, " & _
                  "B.现价,substrb(substr(A.规格,1,instr(A.规格,'┆')-1),1,20) as 规格, " & _
                  "D.名称 as 费用项目,A.费用类型 as 费用类别,'未申报' as 标志 " & _
                  "from 收费细目 A,收费价目 B,收入项目 D " & _
                  "where A.ID=B.收费细目ID and B.收入项目ID=D.ID And " & _
                  "nvl(B.终止日期,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                  "A.编码=[1]"
        Else
        gstrSQL = "select A.ID,E.项目编码 as 编码,F.名称 as 类别," & _
                  "A.名称 As 中文名称,'' as 英文名称, " & _
                  "zlspellcode(A.名称) as 简码,substrb(A.名称,1,40) as 别名,substrb(A.计算单位,1,20) as 计算单位, " & _
                  "B.现价,substrb(substr(A.规格,1,instr(A.规格,'┆')-1),1,20) as 规格, " & _
                  "D.名称 as 费用项目,A.费用类型 as 费用类别,decode(nvl(E.是否医保,0),1,'启用','未启用') as 标志 " & _
                  "from 收费细目 A,收费价目 B,收入项目 D,医保支付项目 E,保险支付大类 F " & _
                  "where A.ID=B.收费细目ID and B.收入项目ID=D.ID And " & _
                  "nvl(B.终止日期,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                  "A.编码=[1] and A.ID=E.收费细目ID And E.险类=F.险类 And E.大类ID=F.ID and F.险类=[2] And E.中心=" & mint中心
    End If
    Set mrsTMP = zlDatabase.OpenSQLRecord(gstrSQL, "保险项目选择", strCode, mint险类)
    If Not mrsTMP.EOF Then
        mrsTMP.MoveFirst
        mintID = mrsTMP!ID
        cmd类别.Text = mrsTMP!类别
        txt编号.Text = mrsTMP!编码
        txt中文名称.Text = mrsTMP!中文名称
        txt简码.Text = mrsTMP!简码
        txt别名.Text = IIf(IsNull(mrsTMP!别名), "", mrsTMP!别名)
        txt单位.Text = IIf(IsNull(mrsTMP!计算单位), "", mrsTMP!计算单位)
        txt单价.Text = mrsTMP!现价
        txt规格.Text = IIf(IsNull(mrsTMP!规格), "", mrsTMP!规格)
        txt费用项目.Text = IIf(IsNull(mrsTMP!费用项目), "", mrsTMP!费用项目)
        cmd费别.Text = IIf(IsNull(mrsTMP!费用类别), "", mrsTMP!费用类别)
        txt标志.Text = mrsTMP!标志
    End If
    
    Frm医保对码_资阳.Show 1
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmd查询_Click()
   Dim StrInput As String, strOutput As String
   Dim strTmpArr As Variant, strArr As Variant
   Dim rsTmp As New ADODB.Recordset
   Dim int大类id As Integer
   
    
   
    StrInput = vbTab & g病人身份_广元旺苍.机构编码
    StrInput = StrInput & vbTab & txt编号.Text
    
    If 业务请求_广元旺苍(提取项目_资阳, StrInput, strOutput) = False Then Exit Sub
    
    strArr = Split(strOutput, "@$")
    strTmpArr = Split(strArr(0), "||")
    txt标志.Text = strTmpArr(1)

    If cmd费别.Text <> strTmpArr(4) Or cmd类别.Text <> strTmpArr(2) Then
       If MsgBox("该项目在医保中心的费用类别与本地的不一致，是否更新?", vbOKCancel) = vbOK Then
            cmd费别.Text = strTmpArr(4)
            cmd类别.Text = strTmpArr(2)
            
            '更新费用类别
            '$IF HIS9.19
            #If gverControl = 0 Then
                gstrSQL = "ZL_收费细目_UPDATE_资阳(" & mintID & ",'" & cmd费别.Text & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类别")
            #Else
            '$ELSE  HIS+
                gstrSQL = "ZL_收费项目目录_UPDATE_资阳(" & mintID & ",'" & cmd费别.Text & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新费用类别")
            #End If
        End If
    End If
    
    gstrSQL = "select nvl(ID,0) as ID from 保险支付大类 where 险类=[1] And 名称=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, CStr(cmd类别.Text))
    int大类id = rsTmp!ID
    
    gstrSQL = "ZL_医保支付项目_Modify(" & mintID & "," & mint险类 & "," & mint中心 & "," & _
              int大类id & ",'" & txt编号.Text & "','" & txt别名.Text & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "'," & IIf(txt标志.Text = "启用", 1, 0) & ")"
    ExecuteProcedure_广元旺苍 "保存医保支付项目"
    
End Sub

Private Sub cmd取消_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd确定_Click()
    '返回选择项目编码
    
    mstrCode = txt编号.Text
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd申报_Click()
   Dim StrInput As String, strOutput As String
   Dim rsTmp As New ADODB.Recordset
   Dim int大类id As Integer
     
    StrInput = vbTab & g病人身份_广元旺苍.机构编码
    StrInput = StrInput & vbTab & txt编号.Text & "||"
    StrInput = StrInput & cmd类别.Text & "||"
    StrInput = StrInput & txt中文名称.Text & "||"
    StrInput = StrInput & txt英文名称.Text & "||"
    StrInput = StrInput & txt简码.Text & "||"
    StrInput = StrInput & txt别名.Text & "||"
    StrInput = StrInput & txt单位.Text & "||"
    StrInput = StrInput & txt单价.Text & "||"
    StrInput = StrInput & txt规格.Text & "||"
    StrInput = StrInput & txt费用项目.Text & "||"
    StrInput = StrInput & cmd费别.Text
    
    StrInput = StrInput & vbTab & gstrUserName
    StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "YYYY-M-DD")
    
    If 业务请求_广元旺苍(申报项目_资阳, StrInput, strOutput) = False Then Exit Sub
    
    gstrSQL = "select nvl(ID,0) as ID from 保险支付大类 where 险类=[1] And 名称=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, CStr(cmd类别.Text))
    int大类id = rsTmp!ID
    
    gstrSQL = "ZL_医保支付项目_Modify(" & mintID & "," & mint险类 & "," & mint中心 & "," & _
               int大类id & ",'" & txt编号.Text & "','" & txt别名.Text & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "',0)"
    ExecuteProcedure_广元旺苍 "保存医保支付项目"
    
    MsgBox "该项目已经成功传输到医保中心，请通知医保中心审核！"
End Sub


Private Sub txt编号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, mrsTMP As ADODB.Recordset
    If KeyCode = vbKeyReturn Then
   
        mblnOK = False
        
        Set mrsTMP = New ADODB.Recordset
        mrsTMP.CursorLocation = adUseClient
        Set rsTmp = New ADODB.Recordset
        rsTmp.CursorLocation = adUseClient
    
        strSQL = "Select * from 医保支付项目 Where 险类=[1] And 中心=[2] and 项目编码=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "查询支付大类", mint险类, mint中心, CStr(txt编号.Text))
        If rsTmp.EOF Then
            gstrSQL = "select A.ID,A.编码,decode(A.类别,'J','服务','1','服务','5','药品','6','药品','7','药品','诊疗') as 类别," & _
                      "A.名称 As 中文名称,'' as 英文名称, " & _
                      "zlspellcode(A.名称) as 简码,substrb(A.名称,1,40) as 别名,substr(A.计算单位,1,20) as 计算单位, " & _
                      "B.现价,substr(substr(A.规格,1,instr(A.规格,'┆')-1),1,20) as 规格, " & _
                      "D.名称 as 费用项目,A.费用类型 as 费用类别 ,'未申报' as 标志 " & _
                      "from 收费细目 A,收费价目 B,收入项目 D " & _
                      "where A.ID=B.收费细目ID and B.收入项目ID=D.ID And " & _
                      "nvl(B.终止日期,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                      "A.编码=[1]"
            Else
            gstrSQL = "select A.ID,E.项目编码 as 编码,F.名称 as 类别," & _
                      "A.名称 As 中文名称,'' as 英文名称, " & _
                      "zlspellcode(A.名称) as 简码,substrb(A.名称,1,40) as 别名,substr(A.计算单位,1,20) as 计算单位, " & _
                      "B.现价,substr(substr(A.规格,1,instr(A.规格,'┆')-1),1,20) as 规格, " & _
                      "D.名称 as 费用项目,A.费用类型 as 费用类别,decode(nvl(E.是否医保,0),1,'启用','未启用') as 标志 " & _
                      "from 收费细目 A,收费价目 B,收费别名 C,收入项目 D,医保支付项目 E,保险支付大类 F " & _
                      "where A.ID=B.收费细目ID and B.收入项目ID=D.ID And " & _
                      "nvl(B.终止日期,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                      "A.编码=[1] and A.ID=E.收费细目ID And E.险类=F.险类 And E.大类ID=F.ID and F.险类=[2] And E.中心=[3]"
        End If
        Set mrsTMP = zlDatabase.OpenSQLRecord(gstrSQL, "", CStr(txt编号.Text), mint险类, mint中心)
            
        If Not mrsTMP.EOF Then
            mrsTMP.MoveFirst
            mintID = mrsTMP!ID
            cmd类别.Text = mrsTMP!类别
            txt编号.Text = mrsTMP!编码
            txt中文名称.Text = mrsTMP!中文名称
            txt简码.Text = mrsTMP!简码
            txt别名.Text = IIf(IsNull(mrsTMP!别名), "", mrsTMP!别名)
            txt单位.Text = IIf(IsNull(mrsTMP!计算单位), "", mrsTMP!计算单位)
            txt单价.Text = mrsTMP!现价
            txt规格.Text = IIf(IsNull(mrsTMP!规格), "", mrsTMP!规格)
            txt费用项目.Text = IIf(IsNull(mrsTMP!费用项目), "", mrsTMP!费用项目)
            cmd费别.Text = IIf(IsNull(mrsTMP!费用类别), "", mrsTMP!费用类别)
            txt标志.Text = mrsTMP!标志
        End If
    End If
End Sub
