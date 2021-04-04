VERSION 5.00
Begin VB.Form frmPatholSpecimenCfg 
   Caption         =   "病理标本设置"
   ClientHeight    =   5940
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10455
   Icon            =   "frmPatholSpecimenCfg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10455
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   10095
      TabIndex        =   2
      Top             =   5280
      Width           =   10095
      Begin VB.CommandButton cmdExit 
         Caption         =   "退  出(&E)"
         Height          =   400
         Left            =   8880
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cbxSpecimenPart 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删  除(&D)"
         Height          =   400
         Left            =   6240
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保  存(&S)"
         Height          =   400
         Left            =   7560
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label labSpecimenPart 
         Caption         =   "标本部位："
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   195
         Width           =   975
      End
   End
   Begin VB.Frame framSpecimenCfg 
      Caption         =   "病理检查标本记录"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7858
         DefaultCols     =   ""
         GridRows        =   201
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholSpecimenCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub AdjustFace()
'调整界面布局
    framSpecimenCfg.Left = 120
    framSpecimenCfg.Top = 120
    framSpecimenCfg.Width = Me.Width - 360
    framSpecimenCfg.Height = Me.Height - picControl.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSpecimenCfg.Width - 240
    ufgData.Height = framSpecimenCfg.Height - 360
    
    
    picControl.Left = 120
    picControl.Top = Me.Height - picControl.Height - 620
    picControl.Width = Me.Width - 360
    
    
    cbxSpecimenPart.Top = 0
    
    labSpecimenPart.Left = 0
    labSpecimenPart.Top = cbxSpecimenPart.Top + 30
    
    cbxSpecimenPart.Left = labSpecimenPart.Left + labSpecimenPart.Width + 60
    
    cmdExit.Left = picControl.Width - cmdSave.Width
    cmdExit.Top = 0
    
    cmdSave.Left = cmdExit.Left - cmdSave.Width - 120
    cmdSave.Top = 0
    
    cmdDel.Left = cmdSave.Left - cmdDel.Width - 120
    cmdDel.Top = 0
    
End Sub



Private Sub InitStudySpecimenList()
    '设置行数  因用户需求定制为500行
    ufgData.GridRows = 501
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.DefaultColNames = gstrSpecimenModuleCols
    
    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
    '初始化病理检查标本列表
    ufgData.ColNames = gstrSpecimenModuleCols
    ufgData.ColConvertFormat = gstrSpecimenModuleConvertFormat
End Sub


Private Sub cbxSpecimenPart_Click()
'过滤套餐信息
On Error GoTo ErrHandle
    Dim strSQL As String
    
    If cbxSpecimenPart.Text = "" Then
        strSQL = "select ID,标本名称,标本部位,标本类型,默认标本量,默认制片数,简码,备注 from 病理检查标本 order by 标本部位,标本名称"
    Else
        strSQL = "select ID,标本名称,标本部位,标本类型,默认标本量,默认制片数,简码,备注 from 病理检查标本 where 标本部位=[1] order by 标本部位,标本名称"
    End If
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbxSpecimenPart.Text)
    
    Call ufgData.RefreshData
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdDel_Click()
'删除病理检查标本
On Error GoTo ErrHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行删除的病理检查标本。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除该病理检查标本数据吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgData.DelCurRow
    
    '保存删除的数据
    Call SaveStudySpeciments(True)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandle
    Call Me.Hide
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cmdSave_Click()
On Error GoTo ErrHandle
    Dim blnValid As Boolean
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到病理检查列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '保存套餐信息
    Call SaveStudySpeciments
    
    Call ConfigInput
    
    Call MsgBoxD(Me, "数据已保存成功。", vbOKOnly, Me.Caption)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitStudySpecimenList
    Call LoadStudySpecimenMoudleData
    
    Call ConfigInput
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ConfigInput()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strSpecimenParts As String
    
    '读取已经存在的标本部位
    strSQL = "select distinct(标本部位) as 标本部位 from 病理检查标本 order by 标本部位"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    ufgData.ComboxListFormat(ufgData.GetColIndex(gstrSpecimenModule_标本部位)) = ""
    cbxSpecimenPart.Clear
     
    If rsData.RecordCount > 0 Then

        Call cbxSpecimenPart.AddItem("")
        
        strSpecimenParts = "|"
        
        While Not rsData.EOF
            If Nvl(rsData!标本部位) <> "" Then
                
                If strSpecimenParts <> "|" Then strSpecimenParts = strSpecimenParts & "|"
                
                strSpecimenParts = strSpecimenParts & Nvl(rsData!标本部位)
                Call cbxSpecimenPart.AddItem(Nvl(rsData!标本部位))
            
            End If
            rsData.MoveNext
        Wend
        
        If strSpecimenParts = "|" Then Exit Sub
        ufgData.ComboxListFormat(ufgData.GetColIndex(gstrSpecimenModule_标本部位)) = strSpecimenParts
    End If
End Sub


Private Sub LoadStudySpecimenMoudleData()
'载入病理检查标本模板数据
    Dim strSQL As String
    
    strSQL = "select ID,标本名称,标本部位,标本类型,默认标本量,默认制片数,简码,备注 from 病理检查标本 order by 标本部位,标本名称"
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Call ufgData.RefreshData
End Sub


Private Sub SaveStudySpeciments(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'blnIsSaveOnlyDel:是否保存仅删除的数据

'保存病理检查标本数据
    Dim i As Long
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    For i = 1 To ufgData.GridRows - 1
        Select Case ufgData.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                
                strSQL = "select ZL_病理标本配置_新增([1],[2],[3],[4],[5],[6],[7]) as 返回值 from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                                        ufgData.Text(i, gstrSpecimenModule_标本名称), _
                                                        ufgData.Text(i, gstrSpecimenModule_标本部位), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_标本类型)), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_默认标本量)), _
                                                        Val(ufgData.Text(i, gstrSpecimenModule_默认制片数)), _
                                                        ufgData.Text(i, gstrSpecimenModule_简码), _
                                                        ufgData.Text(i, gstrSpecimenModule_备注))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveStudySpeciments", "未成功获取新增后的病理检查标本ID,处理失败。")
                    Exit Sub
                End If
                
                
                Call ufgData.SyncText(i, gstrSpecimenModule_ID, rsData!返回值)
				
				ufgData.RowState(i) = TDataRowState.Normal
                                                        
            Case TDataRowState.Del
                strSQL = "ZL_病理标本配置_删除(" & Val(ufgData.KeyValue(i)) & ")"
                
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
				
				ufgData.RowState(i) = TDataRowState.Normal
                
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                strSQL = "ZL_病理标本配置_更新(" & Val(ufgData.KeyValue(i)) & ",'" & _
                                                ufgData.Text(i, gstrSpecimenModule_标本名称) & "','" & _
                                                ufgData.Text(i, gstrSpecimenModule_标本部位) & "'," & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_标本类型)) & ",'" & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_默认标本量)) & "'," & _
                                                Val(ufgData.Text(i, gstrSpecimenModule_默认制片数)) & ",'" & _
                                                ufgData.Text(i, gstrSpecimenModule_简码) & "','" & _
                                                ufgData.Text(i, gstrSpecimenModule_备注) & "')"
                                                
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
				
				ufgData.RowState(i) = TDataRowState.Normal

        End Select
        
    Next i
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
        
    
    '如果未录入标本名称，则显示淡红色
    iCol = ufgData.GetColIndex(gstrSpecimenModule_标本名称)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrSpecimenModule_标本名称) = "", ufgData.ErrCellColor, ufgData.BackColor)
           
    
    
    '如果默认制片数小于1，则显示淡红色
    iCol = ufgData.GetColIndex(gstrSpecimenModule_默认制片数)
    
    ufgData.CellColor(Row, iCol) = IIf(Val(ufgData.Text(Row, gstrSpecimenModule_默认制片数)) < 1, ufgData.ErrCellColor, ufgData.BackColor)
    
    
    '设置检查标本简码
    If ufgData.Text(Row, gstrSpecimenModule_标本名称) <> "" Then
        If ufgData.Text(Row, gstrSpecimenModule_简码) = "" Then ufgData.Text(Row, gstrSpecimenModule_简码) = zlCommFun.SpellCode(ufgData.Text(Row, gstrSpecimenModule_标本名称))
    End If
End Sub



Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strComboboxText As String
    
    If Row > 0 Then
        '自动填写标本部位
        If ufgData.Text(Row, gstrSpecimenModule_标本部位) = "" Then
            If Row - 1 > 0 Then
                If ufgData.Text(Row - 1, gstrSpecimenModule_标本部位) <> "" Then
                    ufgData.Text(Row, gstrSpecimenModule_标本部位) = ufgData.Text(Row - 1, gstrSpecimenModule_标本部位)
                End If
            End If
            
            If ufgData.Text(Row, gstrSpecimenModule_标本部位) = "" Then
                strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrSpecimenModule_标本部位))
                
                If strComboboxText <> "" Then
                    If InStr(strComboboxText, ";") > 0 Then
                        strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, ";") - 1)
                    End If
                    ufgData.Text(Row, gstrSpecimenModule_标本部位) = Mid(strComboboxText, InStr(strComboboxText, "#") + 1, 255)
                    
                End If
            End If
        End If
        
        '自动填写标本类型
        If ufgData.Text(Row, gstrSpecimenModule_标本类型) = "" Then
                If Row - 1 > 0 Then
                    If ufgData.Text(Row - 1, gstrSpecimenModule_标本类型) <> "" Then
                        ufgData.Text(Row, gstrSpecimenModule_标本类型) = ufgData.Text(Row - 1, gstrSpecimenModule_标本类型)
                    End If
                End If
                
                If ufgData.Text(Row, gstrSpecimenModule_标本类型) = "" Then
                    strComboboxText = ufgData.DataGrid.ColComboList(ufgData.GetColIndex(gstrSpecimenModule_标本类型))
                    
                    If strComboboxText <> "" Then
                        If InStr(strComboboxText, "|") > 0 Then
                            strComboboxText = Mid(strComboboxText, 1, InStr(strComboboxText, "|") - 1)
                        End If
                        ufgData.Text(Row, gstrSpecimenModule_标本类型) = strComboboxText
                        
                    End If
                End If
        End If
                
        '设置默认制片数
        If Val(ufgData.Text(Row, gstrSpecimenModule_默认制片数)) <= 0 Then ufgData.Text(Row, gstrSpecimenModule_默认制片数) = "1"
        
    End If
End Sub
