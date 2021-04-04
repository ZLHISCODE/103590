VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatholAntibody_AntiUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "抗体维护"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "frmPatholAntibody_AntiUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   33
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.TextBox txtAlredyCount 
      Height          =   300
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtUseCount 
      Height          =   300
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   1
      Top             =   915
      Width           =   1695
   End
   Begin VB.CheckBox chkContinue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "确认后继续添加"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox txtApplySituation 
      Height          =   300
      Left            =   4200
      TabIndex        =   9
      Top             =   2915
      Width           =   2025
   End
   Begin VB.ComboBox cbxActionObject 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   2915
      Width           =   2025
   End
   Begin VB.ComboBox cbxLieracType 
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   2415
      Width           =   2025
   End
   Begin VB.ComboBox cbxCloneType 
      Height          =   300
      ItemData        =   "frmPatholAntibody_AntiUpdate.frx":179A
      Left            =   1080
      List            =   "frmPatholAntibody_AntiUpdate.frx":179C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2415
      Width           =   2025
   End
   Begin VB.TextBox txtMemo 
      Height          =   780
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3915
      Width           =   5145
   End
   Begin VB.CommandButton cmdNewAntibody_Cancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   5040
      TabIndex        =   14
      Top             =   5235
      Width           =   1215
   End
   Begin VB.CommandButton cmdNewAntibody_Sure 
      Caption         =   "确 定(&S)"
      Height          =   400
      Left            =   3720
      TabIndex        =   13
      Top             =   5235
      Width           =   1215
   End
   Begin VB.TextBox txtAntibodyName 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   915
      Width           =   1815
   End
   Begin VB.ComboBox cbxValidCount 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1915
      Width           =   1665
   End
   Begin VB.TextBox txtRegisterDoctor 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   10
      Top             =   3415
      Width           =   2025
   End
   Begin MSComCtl2.DTPicker dtpMadeDate 
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   1415
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpOverdueDate 
      Height          =   300
      Left            =   4200
      TabIndex        =   5
      Top             =   1915
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin MSComCtl2.DTPicker dtpRegisterTime 
      Height          =   300
      Left            =   4200
      TabIndex        =   11
      Top             =   3415
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   99745795
      CurrentDate     =   40646.4399652778
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "应用情况："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   31
      Top             =   2935
      Width           =   900
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "作用对象："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   2935
      Width           =   900
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "理化性质："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   29
      Top             =   2445
      Width           =   900
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "克 隆 性："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   28
      Top             =   2445
      Width           =   900
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "月"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2880
      TabIndex        =   27
      Top             =   1960
      Width           =   180
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用人份："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   26
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "抗体名称："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  备  注："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   3960
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6360
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "    请正确录入抗体的相关信息，有红色星号标记的为必录数据，在添加过程中，如有数据不符合系统要求的，系统将进行提示。"
      Height          =   495
      Left            =   840
      TabIndex        =   23
      Top             =   195
      Width           =   5535
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   120
      Picture         =   "frmPatholAntibody_AntiUpdate.frx":179E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "已用人份："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   1465
      Width           =   900
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "生产日期："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   21
      Top             =   1465
      Width           =   900
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "有 效 期："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1955
      Width           =   900
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "过期日期："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   19
      Top             =   1955
      Width           =   900
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "登 记 人："
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   900
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "登记时间："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   17
      Top             =   3425
      Width           =   900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmPatholAntibody_AntiUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentGrid As ucFlexGrid
Private mblnIsSucceed As Boolean
Private mblnIsUpdate As Boolean

Property Get IsSucceed() As Boolean
    IsSucceed = mblnIsSucceed
End Property

Property Get IsUpdate() As Boolean
    IsUpdate = mblnIsUpdate
End Property

Property Let IsUpdate(value As Boolean)
    mblnIsUpdate = value
End Property

Public Function ShowAddAntibodyWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'显示抗体新增窗口
    Dim curDate As Date
    
    ShowAddAntibodyWindow = False
    
    Set mufgParentGrid = ufgParentGrid
    
    Me.Caption = "新增抗体"
    mblnIsUpdate = False
    
    curDate = zlDatabase.Currentdate
    
    dtpMadeDate.value = curDate
    dtpOverdueDate.value = curDate
    dtpRegisterTime.value = curDate
    txtRegisterDoctor.Text = UserInfo.姓名
    
    Call CloseProcessHint
    
    chkContinue.value = False
    chkContinue.Visible = True
    
    Call Me.Show(1, owner)
    
    ShowAddAntibodyWindow = mblnIsSucceed

End Function


Public Function ShowUpdateAntibodyWindow(ufgParentGrid As ucFlexGrid, owner As Form) As Boolean
'显示抗体更新窗口
    ShowUpdateAntibodyWindow = False
    
    Set mufgParentGrid = ufgParentGrid
        
    Me.Caption = "更新抗体"
    mblnIsUpdate = True
        
    Call CloseProcessHint
    
    Call ConfigUpdateFace
    
    chkContinue.value = False
    chkContinue.Visible = False

    
    Call Me.Show(1, owner)
    
    ShowUpdateAntibodyWindow = mblnIsSucceed
End Function


Private Function GetCloneTypeIndex(ByVal strCloneValue As String) As Long
'取得当前克隆性质
    GetCloneTypeIndex = 0
    
    If strCloneValue = "多克隆" Then
        GetCloneTypeIndex = 1
    End If
End Function

Public Sub ConfigUpdateFace()
    
    With mufgParentGrid
        txtAntibodyName.Text = .Text(.SelectionRow, gstrAntibody_抗体名称)
        txtUseCount.Text = .Text(.SelectionRow, gstrAntibody_使用人份)
        txtAlredyCount.Text = .Text(.SelectionRow, gstrAntibody_已用人份)
        dtpMadeDate.value = .Text(.SelectionRow, gstrAntibody_生产日期)
        cbxValidCount.Text = Val(.Text(.SelectionRow, gstrAntibody_有效期))
        dtpOverdueDate.value = .Text(.SelectionRow, gstrAntibody_过期日期)
        cbxCloneType.ListIndex = GetCloneTypeIndex(.Text(.SelectionRow, gstrAntibody_克隆性))
        cbxLieracType.Text = .Text(.SelectionRow, gstrAntibody_理化性质)
        cbxActionObject.Text = .Text(.SelectionRow, gstrAntibody_作用对象)
        txtApplySituation.Text = .Text(.SelectionRow, gstrAntibody_应用情况)
        txtRegisterDoctor.Text = .Text(.SelectionRow, gstrAntibody_登记人)
        dtpRegisterTime.value = .Text(.SelectionRow, gstrAntibody_登记时间)
        txtMemo.Text = .Text(.SelectionRow, gstrAntibody_备注)
    End With
    
    '判断该抗体是否已被使用过，如果已被使用，则某些信息不能进行更新
    Dim strSQL As String
    Dim rsUsed As ADODB.Recordset
    
    
    strSQL = "select 1 from 病理特检信息 where 抗体ID=[1]"
    Set rsUsed = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow))
    
    If rsUsed.RecordCount <= 0 Then Exit Sub
    
    txtAntibodyName.Enabled = False
    txtAntibodyName.BackColor = &HE0E0E0
    
    cbxCloneType.Enabled = False
    cbxCloneType.BackColor = &HE0E0E0
    
    cbxLieracType.Enabled = False
    cbxLieracType.BackColor = &HE0E0E0
    
    cbxActionObject.Enabled = False
    cbxActionObject.BackColor = &HE0E0E0
    
End Sub

Private Sub LoadValidDate()
'载入有效期
    cbxValidCount.Clear
    
    Call cbxValidCount.AddItem("3")
    Call cbxValidCount.AddItem("6")
    Call cbxValidCount.AddItem("9")
    Call cbxValidCount.AddItem("12")
    Call cbxValidCount.AddItem("18")
    Call cbxValidCount.AddItem("24")
    Call cbxValidCount.AddItem("36")
End Sub


Private Sub LoadCloneType()
'载入克隆类型
    cbxCloneType.Clear
    
    Call cbxCloneType.AddItem("0-单克隆（浓缩型）")
    Call cbxCloneType.AddItem("1-单克隆（即用型）")
    Call cbxCloneType.AddItem("2-多克隆（浓缩型）")
    Call cbxCloneType.AddItem("3-多克隆（即用型）")
    
    cbxCloneType.ListIndex = 0
End Sub


Private Sub LoadLieracType()
'载入理化性质
    cbxLieracType.Clear
    
    Call cbxLieracType.AddItem("IgM")
    Call cbxLieracType.AddItem("IgG")
    Call cbxLieracType.AddItem("IgA")
    Call cbxLieracType.AddItem("IgE")
    Call cbxLieracType.AddItem("IgD")
End Sub


Private Sub LoadActionObject()
'载入作用对象
    cbxActionObject.Clear
    
    Call cbxActionObject.AddItem("抗毒素")
    Call cbxActionObject.AddItem("抗菌抗体")
    Call cbxActionObject.AddItem("抗病毒抗体")
    Call cbxActionObject.AddItem("亲细胞抗体")
End Sub


Private Function CheckAntibodyDataIsValid() As String
    CheckAntibodyDataIsValid = ""
    
    '检查标本名称是否为空
    If Trim(txtAntibodyName.Text) = "" Then
        CheckAntibodyDataIsValid = "抗体名称不能为空。"
        
        Call txtAntibodyName.SetFocus
        Exit Function
    End If
    
    '检查标本数量是否正确录入
    If Trim(txtUseCount.Text) = "" Or Val(txtUseCount.Text) = 0 Then
        CheckAntibodyDataIsValid = "使用人份输入无效，请输入有效数字。"
        
        Call txtUseCount.SetFocus
        Exit Function
    End If
    
    If dtpOverdueDate.value <= dtpMadeDate.value Then
        CheckAntibodyDataIsValid = "过期日期必须大于生产日期。"
        
        Call dtpOverdueDate.SetFocus
        Exit Function
    End If
    
    
    
    '检查抗体名称是否重复

    Dim i As Integer
    For i = 1 To mufgParentGrid.GridRows - 1
        If Not mufgParentGrid.RowState(i) = TDataRowState.Del Then
            If Not mblnIsUpdate Then
                If mufgParentGrid.Text(i, gstrAntibody_抗体名称) = txtAntibodyName.Text Then
                    CheckAntibodyDataIsValid = "抗体名称重复。"
                
                    Call txtAntibodyName.SetFocus
                    Exit Function
                End If
            Else
                If Not mufgParentGrid.SelectionRow = i Then
                    If mufgParentGrid.Text(i, gstrAntibody_抗体名称) = txtAntibodyName.Text Then
                        CheckAntibodyDataIsValid = "抗体名称重复。"
                    
                        Call txtAntibodyName.SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    
End Function



Private Function NewAntibody(ByRef lngAntibodyId As Long) As String
'在数据库中新增抗体记录
On Error GoTo ErrHandle

    Dim strSQL As String
    Dim rsReture As ADODB.Recordset
    
    NewAntibody = ""
    
'    strSQL = "zl_病理抗体_新增('" & txtAntibodyName.Text & "'," & Val(txtUseCount.Text) & "," & Val(txtAlredyCount.Text) & "," & _
'                                To_Date(dtpMadeDate.value) & "," & Val(cbxValidCount.Text) & "," & To_Date(dtpOverdueDate.value) & "," & _
'                                Val(cbxCloneType.Text) & ",'" & cbxActionObject.Text & "','" & cbxLieracType.Text & "','" & _
'                                txtApplySituation.Text & "','" & txtRegisterDoctor.Text & "'," & To_Date(dtpRegisterTime.value) & ",'" & _
'                                txtMemo.Text & "')"
'
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    
                                
    strSQL = "select zl_病理抗体_新增([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13]) as 返回值 from dual"
                                
    Set rsReture = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                txtAntibodyName.Text, _
                                Val(txtUseCount.Text), _
                                Val(txtAlredyCount.Text), _
                                CDate(dtpMadeDate.value), _
                                Val(cbxValidCount.Text), _
                                CDate(dtpOverdueDate.value), _
                                Val(cbxCloneType.Text), _
                                cbxActionObject.Text, _
                                cbxLieracType.Text, _
                                txtApplySituation.Text, _
                                txtRegisterDoctor.Text, _
                                CDate(dtpRegisterTime.value), _
                                txtMemo.Text)
                                
    If rsReture.RecordCount > 0 Then lngAntibodyId = rsReture!返回值
    
Exit Function
ErrHandle:
    NewAntibody = err.Description
End Function


Private Function AddAntibodyToList(lngNewAntibodyId As Long) As String
'添加抗体记录到显示列表
On Error GoTo ErrHandle
    AddAntibodyToList = ""
    
    Dim lngNewRecordIndex As Long
    
    AddAntibodyToList = ""
    
    With mufgParentGrid
        lngNewRecordIndex = .NewRow
        
        .Text(lngNewRecordIndex, gstrAntibody_抗体ID) = lngNewAntibodyId
        .Text(lngNewRecordIndex, gstrAntibody_抗体名称) = txtAntibodyName.Text
        .Text(lngNewRecordIndex, gstrAntibody_使用人份) = Val(txtUseCount.Text)
        .Text(lngNewRecordIndex, gstrAntibody_已用人份) = Val(txtAlredyCount.Text)
        .Text(lngNewRecordIndex, gstrAntibody_生产日期) = Format(dtpMadeDate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrAntibody_有效期) = Val(cbxValidCount.Text) & "月"
        .Text(lngNewRecordIndex, gstrAntibody_过期日期) = Format(dtpOverdueDate.value, gstrDateFormat)
        .Text(lngNewRecordIndex, gstrAntibody_克隆性) = Trim(zlStr.SubB(cbxCloneType.Text, InStr(1, cbxCloneType.Text, "-") + 1, 50))
        .Text(lngNewRecordIndex, gstrAntibody_作用对象) = cbxActionObject.Text
        .Text(lngNewRecordIndex, gstrAntibody_理化性质) = cbxLieracType.Text
        .Text(lngNewRecordIndex, gstrAntibody_应用情况) = txtApplySituation.Text
        .Text(lngNewRecordIndex, gstrAntibody_使用状态) = "使用中"
        .Text(lngNewRecordIndex, gstrAntibody_登记人) = txtRegisterDoctor.Text
        .Text(lngNewRecordIndex, gstrAntibody_登记时间) = dtpRegisterTime.value
        .Text(lngNewRecordIndex, gstrAntibody_备注) = txtMemo.Text
    
    End With
     
    
Exit Function
ErrHandle:
    AddAntibodyToList = err.Description
End Function


Private Function UpdateAntibodyInfToList()
'更新抗体列表中的抗体信息
On Error GoTo ErrHandle
    UpdateAntibodyInfToList = ""
    
    With mufgParentGrid
        .Text(.SelectionRow, gstrAntibody_抗体名称) = txtAntibodyName.Text
        .Text(.SelectionRow, gstrAntibody_使用人份) = Val(txtUseCount.Text)
        .Text(.SelectionRow, gstrAntibody_已用人份) = Val(txtAlredyCount.Text)
        .Text(.SelectionRow, gstrAntibody_生产日期) = Format(dtpMadeDate.value, gstrDateFormat)
        .Text(.SelectionRow, gstrAntibody_有效期) = Val(cbxValidCount.Text) & "月"
        .Text(.SelectionRow, gstrAntibody_过期日期) = Format(dtpOverdueDate.value, gstrDateFormat)
        .Text(.SelectionRow, gstrAntibody_克隆性) = Trim(zlStr.SubB(cbxCloneType.Text, InStr(1, cbxCloneType.Text, "-") + 1, 50))
        .Text(.SelectionRow, gstrAntibody_作用对象) = cbxActionObject.Text
        .Text(.SelectionRow, gstrAntibody_理化性质) = cbxLieracType.Text
        .Text(.SelectionRow, gstrAntibody_应用情况) = txtApplySituation.Text
        .Text(.SelectionRow, gstrAntibody_使用状态) = "使用中"
        .Text(.SelectionRow, gstrAntibody_登记人) = txtRegisterDoctor.Text
        .Text(.SelectionRow, gstrAntibody_登记时间) = dtpRegisterTime.value
        .Text(.SelectionRow, gstrAntibody_备注) = txtMemo.Text
    End With
Exit Function
ErrHandle:
    UpdateAntibodyInfToList = err.Description
End Function



Private Function UpdateAntibody() As String
'更新数据库中的抗体数据
On Error GoTo ErrHandle

    Dim strSQL As String
    Dim lngCurAntibodyId As Long
    
    UpdateAntibody = ""
    
    lngCurAntibodyId = mufgParentGrid.KeyValue(mufgParentGrid.SelectionRow)
    
    strSQL = "zl_病理抗体_更新(" & lngCurAntibodyId & ",'" & txtAntibodyName.Text & "'," & Val(txtUseCount.Text) & "," & Val(txtAlredyCount.Text) & "," & _
                                zlStr.To_Date(dtpMadeDate.value) & "," & Val(cbxValidCount.Text) & "," & zlStr.To_Date(dtpOverdueDate.value) & "," & _
                                Val(cbxCloneType.Text) & ",'" & cbxActionObject.Text & "','" & cbxLieracType.Text & "','" & _
                                txtApplySituation.Text & "','" & txtRegisterDoctor.Text & "'," & zlStr.To_Date(dtpRegisterTime.value) & ",'" & _
                                txtMemo.Text & "')"
                                
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
Exit Function
ErrHandle:
    UpdateAntibody = err.Description
End Function





Private Sub cbxValidCount_LostFocus()
On Error Resume Next
    If Val(cbxValidCount.Text) > 0 Then
        dtpOverdueDate.value = DateAdd("m", Val(cbxValidCount.Text), dtpMadeDate.value)
    End If
End Sub

Private Sub cmdNewAntibody_Cancel_Click()
'    mblnIsSucceed = False '当确认后，如果不退出添加界面，则取消的时候，不能赋值为假
    Call Me.Hide
End Sub

Private Sub cmdNewAntibody_Sure_Click()
On Error GoTo ErrHandle
    Dim strErr As String
    Dim lngNewAntibodyId As Long
    
    mblnIsSucceed = False
    
    
    strErr = CheckAntibodyDataIsValid()
    If Trim(strErr) <> "" Then
        Call ShowProcessHint(strErr)
        Exit Sub
    End If
    
    If Not mblnIsUpdate Then
        '新增抗体记录
        strErr = NewAntibody(lngNewAntibodyId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = AddAntibodyToList(lngNewAntibodyId)
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        Call mufgParentGrid.LocateRow(mufgParentGrid.GridRows - 1)
    Else
        '更新抗体记录
        strErr = UpdateAntibody()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
        
        strErr = UpdateAntibodyInfToList()
        If Trim(strErr) <> "" Then
            Call ShowProcessHint(strErr)
            Exit Sub
        End If
    End If
    
    mblnIsSucceed = True
    
    If Not CBool(chkContinue.value) Then
        Call Unload(Me)
    End If
    
    Call CloseProcessHint
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpOverdueDate_LostFocus()
On Error Resume Next
    cbxValidCount.Text = DateDiff("m", dtpMadeDate.value, dtpOverdueDate.value)
End Sub

Private Sub Form_Initialize()
    mblnIsSucceed = False
    mblnIsUpdate = False
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call LoadValidDate
    Call LoadCloneType
    Call LoadLieracType
    Call LoadActionObject
    
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'显示处理信息
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub CloseProcessHint()
'关闭处理提示
    picShow.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub txtUseCount_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtAlredyCount_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If Not ((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
