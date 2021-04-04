VERSION 5.00
Begin VB.Form frmClinicBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "项目诊疗单据"
   ClientHeight    =   3192
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   5448
   Icon            =   "frmClinicBill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   5448
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboTest 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "体检采用(&M)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   675
      Width           =   1290
   End
   Begin VB.CheckBox chkIn 
      Caption         =   "住院采用(&I)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   4
      Top             =   1380
      Width           =   1290
   End
   Begin VB.CheckBox chkOut 
      Caption         =   "门诊采用(&T)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   2
      Top             =   1035
      Width           =   1290
   End
   Begin VB.ComboBox cboIn 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1335
      Width           =   3255
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -180
      TabIndex        =   13
      Top             =   525
      Width           =   6615
   End
   Begin VB.OptionButton optScope 
      Caption         =   "用于本类别的项目"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   7
      Top             =   2115
      Width           =   5610
   End
   Begin VB.Frame fraBottom 
      Height          =   30
      Left            =   -165
      TabIndex        =   12
      Top             =   2490
      Width           =   6585
   End
   Begin VB.OptionButton optScope 
      Caption         =   "用于本项目"
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   6
      Top             =   1800
      Value           =   -1  'True
      Width           =   5610
   End
   Begin VB.ComboBox cboOut 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4170
      TabIndex        =   9
      Top             =   2655
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      Picture         =   "frmClinicBill.frx":058A
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2655
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3060
      TabIndex        =   8
      Top             =   2655
      Width           =   1100
   End
   Begin VB.Image imgNote 
      Height          =   384
      Left            =   156
      Picture         =   "frmClinicBill.frx":06D4
      Top             =   12
      Width           =   384
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    设置诊疗项目对应的诊疗单据，以便在医嘱发送执行过程中，采用符合项目特性的单据，满足诊疗过程需要。"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   645
      TabIndex        =   10
      Top             =   75
      Width           =   4680
   End
End
Attribute VB_Name = "frmClinicBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前项目：由me.optScope(0).tag保存，由上级程序通过ShowMe函数传入
'---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim strTemp As String
Dim intCount As Integer

Public Sub ShowMe(ByVal frmParent As Object, Optional ByVal lng项目id As Long)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Dim str诊疗类别 As String
    Dim intControl As Integer       '用来控制初始化时复选框的勾选。0-不勾选;1-勾选
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.类别,I.编码,I.名称,I.分类id,nvl(I.服务对象,0) as 服务对象,K.编码 as 类别码,K.名称 as 类别名" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.id=[1] and I.类别=K.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng项目id)
    
    With rsTemp
        If .BOF Or .EOF Then Unload Me: Exit Sub
        str诊疗类别 = !类别
        Me.optScope(0).Tag = !ID: Me.optScope(0).Caption = "&1、应用于本项目(" & !编码 & "-" & !名称 & ")"
        Me.optScope(1).Tag = !类别码: Me.optScope(1).Caption = "&2、应用于所有“" & !类别名 & "”类项目"
        
        If !服务对象 = 1 Or !服务对象 = 3 Then Me.chkOut.Enabled = True
        If !服务对象 = 2 Or !服务对象 = 3 Then Me.chkIn.Enabled = True
        If !服务对象 = 4 Then Me.chkTest.Enabled = True
    End With
    
    gstrSql = "select ID,编码,名称" & _
            " from 诊疗分类目录" & _
            " start with id=[1] " & _
            " connect by prior 上级id=id" & _
            " order by level"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTemp!分类id))
        
    With rsTemp
        Do While Not .EOF
            Load Me.optScope(.AbsolutePosition + 1)
            Me.optScope(.AbsolutePosition + 1).Tag = !ID
            Me.optScope(.AbsolutePosition + 1).Caption = "&" & .AbsolutePosition + 2 & "、应用于“[" & !编码 & "]" & !名称 & "”类项目"
            Me.optScope(.AbsolutePosition + 1).Left = Me.optScope(0).Left
            Me.optScope(.AbsolutePosition + 1).Top = Me.optScope(.AbsolutePosition).Top + Me.optScope(1).Top - Me.optScope(0).Top
            Me.optScope(.AbsolutePosition + 1).Visible = True
            .MoveNext
        Loop
    End With
    
    If Me.chkOut.Enabled Then
        gstrSql = "select 病历文件id from 病历单据应用 where 应用场合=1 and 诊疗项目id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng项目id)
        
        
        If Not rsTemp.EOF Then
            Me.cboOut.Tag = rsTemp!病历文件id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,编号,名称 from 病历文件列表 where 种类=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                Me.chkOut.Value = 0: Me.chkOut.Enabled = False
            ElseIf intControl = 0 Then
                Me.chkOut.Value = 0: Me.cboOut.Enabled = False
            Else
                Me.chkOut.Value = 1: Me.cboOut.Enabled = True
            End If
            Me.cboOut.ListIndex = -1
            Do While Not .EOF
                Me.cboOut.AddItem !编号 & "-" & !名称
                Me.cboOut.ItemData(Me.cboOut.NewIndex) = !ID
                If !ID = Val(Me.cboOut.Tag) Then
                    Me.cboOut.ListIndex = Me.cboOut.NewIndex
                End If
                .MoveNext
            Loop
       
        End With
    End If
    If cboOut.ListIndex = -1 Then
        '药品设置默认的项目：西药对应西药处方签，中草药对应中药处方签
        '如果病历文件列表中涉及药品的内容或顺序做了改变，下面代码可能也要做相应调整
        If str诊疗类别 = "5" Or str诊疗类别 = "6" Then
            cboOut.ListIndex = 0
        ElseIf str诊疗类别 = "7" Then
            cboOut.ListIndex = 1
        Else
            cboOut.Enabled = False: chkOut.Value = 0
        End If
    End If
    chkOut.Tag = cboOut.ListIndex
    
    If Me.chkIn.Enabled Then
        gstrSql = "select 病历文件id from 病历单据应用 where 应用场合=2 and 诊疗项目id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng项目id)
        
        If Not rsTemp.EOF Then
            Me.cboIn.Tag = rsTemp!病历文件id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,编号,名称 from 病历文件列表 where 种类=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                Me.chkIn.Value = 0: Me.chkIn.Enabled = False
            ElseIf intControl = 0 Then
                Me.chkIn.Value = 0: Me.cboIn.Enabled = False
            Else
                Me.chkIn.Value = 1: Me.cboIn.Enabled = True
            End If
            Me.cboIn.ListIndex = -1
            Do While Not .EOF
                Me.cboIn.AddItem !编号 & "-" & !名称
                Me.cboIn.ItemData(Me.cboIn.NewIndex) = !ID
                If !ID = Val(Me.cboIn.Tag) Then
                    Me.cboIn.ListIndex = Me.cboIn.NewIndex
                End If
                .MoveNext
            Loop
            
        End With
    End If
    If Me.cboIn.ListIndex = -1 Then
        '药品设置默认的项目：西药对应西药处方签，中草药对应中药处方签
        '如果病历文件列表中涉及药品的内容或顺序做了改变，下面代码可能也要做相应调整
        If str诊疗类别 = "5" Or str诊疗类别 = "6" Then
            cboIn.ListIndex = 0
        ElseIf str诊疗类别 = "7" Then
            cboIn.ListIndex = 1
        Else
            cboIn.Enabled = False: chkIn.Value = 0
        End If
    End If
    chkIn.Tag = cboIn.ListIndex
    
    If chkTest.Enabled Then
        gstrSql = "select 病历文件id from 病历单据应用 where 应用场合=4 and 诊疗项目id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng项目id)
        
        If Not rsTemp.EOF Then
            Me.cboTest.Tag = rsTemp!病历文件id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,编号,名称 from 病历文件列表 where 种类=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                chkTest.Value = 0: chkTest.Enabled = False
            ElseIf intControl = 0 Then
                chkTest.Value = 0: cboTest.Enabled = False
            Else
                chkTest.Value = 1: cboTest.Enabled = True
            End If
            cboTest.ListIndex = -1
            Do While Not .EOF
                cboTest.AddItem !编号 & "-" & !名称
                cboTest.ItemData(cboTest.NewIndex) = !ID
                If !ID = Val(cboTest.Tag) Then
                    cboTest.ListIndex = cboTest.NewIndex
                End If
                .MoveNext
            Loop
        End With
    End If
    chkTest.Tag = cboTest.ListIndex
    If cboTest.ListIndex = -1 Then
        cboTest.Enabled = False: chkTest.Value = 0
    End If
    
    Me.optScope(0).Value = True
    Me.fraBottom.Top = Me.optScope(Me.optScope.Count - 1).Top + 300
    Me.cmdHelp.Top = Me.fraBottom.Top + 150
    Me.cmdOK.Top = Me.cmdHelp.Top: Me.cmdCancel.Top = Me.cmdHelp.Top
    Me.Height = Me.cmdHelp.Top + Me.cmdHelp.Height + 500
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub cboIn_Click()
    chkIn.Tag = Me.cboIn.ListIndex
End Sub

Private Sub cboOut_Click()
    chkOut.Tag = cboOut.ListIndex
End Sub

Private Sub cboTest_Click()
    chkTest.Tag = cboTest.ListIndex
End Sub

Private Sub chkIn_Click()
    If Me.chkIn.Value = 1 Then
        Me.cboIn.Enabled = True
        If Me.cboIn.ListCount > 0 Then
            Me.cboIn.ListIndex = Val(chkIn.Tag)
        Else
            Me.cboIn.ListIndex = -1
        End If
    Else
        Me.cboIn.Enabled = False
    End If
End Sub

Private Sub chkOut_Click()
    If Me.chkOut.Value = 1 Then
        Me.cboOut.Enabled = True
        If Me.cboOut.ListCount > 0 Then
            Me.cboOut.ListIndex = Val(chkOut.Tag)
        Else
            Me.cboOut.ListIndex = -1
        End If
    Else
        Me.cboOut.Enabled = False
    End If
End Sub

Private Sub chkTest_Click()
    If Me.chkTest.Value = 1 Then
        Me.cboTest.Enabled = True
        If cboTest.ListCount > 0 Then
            Me.cboTest.ListIndex = Val(chkTest.Tag)
        Else
            Me.cboTest.ListIndex = -1
        End If
    Else
        Me.cboTest.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If optScope(0).Value = False Then
        For i = 1 To optScope.UBound
            If optScope(i).Value = True Then
                If MsgBox("该药品诊疗单据应用范围为“" & optScope(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
        
    gstrSql = "zl_诊疗单据应用_Update("
    
    If Me.cboOut.Enabled = False Or Me.cboOut.ListIndex = -1 Then
        gstrSql = gstrSql & "null"
    Else
        gstrSql = gstrSql & Me.cboOut.ItemData(Me.cboOut.ListIndex)
    End If
    
    If Me.cboIn.Enabled = False Or Me.cboIn.ListIndex = -1 Then
        gstrSql = gstrSql & ",null"
    Else
        gstrSql = gstrSql & "," & Me.cboIn.ItemData(Me.cboIn.ListIndex)
    End If
    
    If Me.optScope(0).Value = True Then
        gstrSql = gstrSql & ",0,'" & Me.optScope(0).Tag & "'"
    ElseIf Me.optScope(1).Value = True Then
        gstrSql = gstrSql & ",1,'" & Me.optScope(1).Tag & "'"
    Else
        For intCount = 2 To Me.optScope.Count - 1
            If Me.optScope(intCount).Value = True Then
                gstrSql = gstrSql & ",2,'" & Me.optScope(intCount).Tag & "'"
                Exit For
            End If
        Next
    End If
    
    If cboTest.Enabled = False Or cboTest.ListIndex = -1 Then
        gstrSql = gstrSql & ",null)"
    Else
        gstrSql = gstrSql & "," & cboTest.ItemData(cboTest.ListIndex) & ")"
    End If
    
    err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub optScope_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optScope.UBound
        If i = Index Then
            optScope(i).FontBold = True
        Else
            optScope(i).FontBold = False
        End If
    Next
End Sub
