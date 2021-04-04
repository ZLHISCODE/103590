VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "剂型选择"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmSelModel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   4830
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwModel 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4600
      _ExtentX        =   8123
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblAll 
      AutoSize        =   -1  'True
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4680
      Width           =   390
   End
   Begin VB.Label lblUnAll 
      AutoSize        =   -1  'True
      Caption         =   "全不选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4680
      Width           =   585
   End
End
Attribute VB_Name = "frmSelModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strDrugModel As String
Public mcnHIS As New ADODB.Connection

Public Property Get DrugModel() As String
    DrugModel = strDrugModel
End Property

'Public Property Get ResultModal() As Boolean
'    ResultModal = blnResultModal
'End Property

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Byte
    Dim strTmp As String, strInsert As String
    Dim rsTmp As New ADODB.Recordset
    Dim cmdInsert As New ADODB.Command
    Dim blnBegin As Boolean
    
    For i = 1 To lvwModel.ListItems.Count
        If lvwModel.ListItems(i).Checked = True Then
            strTmp = strTmp & "'" & Trim(lvwModel.ListItems(i).Text) & "',"
        End If
    Next
    If strTmp = "" Then Exit Sub
    strDrugModel = Left(strTmp, Len(strTmp) - 1)
    
    '更新药品基础信息

    Screen.MousePointer = vbHourglass
    DoEvents
    '生成脚本
    strTmp = "Select A.*, B.编码 厂商编码 " _
           & "From (Select A.药品id, B.编码, B.名称, B.规格, D.药品剂型 药品类型, A.剂量系数 剂量, E.计算单位 剂量单位," _
           & "        A.住院包装 包装数量, A.住院单位 包装单位, Nvl(A.上次产地, B.产地) 生产厂商," _
           & "        zlTools.zlSpellCode(B.名称) 拼音码" _
           & "      From 药品规格 A, 收费项目目录 B, 药品特性 D, 诊疗项目目录 E" _
           & "      Where A.药品id = B.ID And A.药名id = D.药名id And A.药名id = E.ID And B.类别 In ('5', '6', '7') And" _
           & "        Nvl(B.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')) = To_Date('3000-1-1', 'yyyy-mm-dd') And" _
           & "        D.药品剂型 In (" & frmSelModel.DrugModel & ")) A, 药品生产商 B " _
           & "Where A.生产厂商 = B.名称(+) order by cast(a.编码 as int) "
    If Me.mcnHIS.State = adStateClosed Then Me.mcnHIS.Open
    rsTmp.Open strTmp, Me.mcnHIS
    If rsTmp.EOF Then
        Screen.MousePointer = vbDefault
        rsTmp.Close
        Exit Sub
    End If
    
    On Error GoTo errHand
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        gcnOutside.BeginTrans
        blnBegin = True
        strInsert = "delete dbo.atf_his_druginfo where drug_code='" & rsTmp!编码 & "' and drugname='" & rsTmp!名称 & "' " & Chr(13) _
                  & "insert into dbo.atf_his_druginfo (drug_code,drugname,specification,drug_type," _
                  & "dosage,dos_unit,pack_amount,pack_name,manufactory,py_code,manu_no) " & Chr(13)
        strTmp = "select '" & rsTmp!编码 & "'," _
               & "'" & rsTmp!名称 & "'," _
               & "'" & rsTmp!规格 & "'," _
               & "'" & rsTmp!药品类型 & "'," _
               & CDbl(rsTmp!剂量) & "," _
               & "'" & rsTmp!剂量单位 & "'," _
               & CDbl(rsTmp!包装数量) & "," _
               & "'" & rsTmp!包装单位 & "'," _
               & "'" & IIf(IsNull(rsTmp!生产厂商), "", rsTmp!生产厂商) & "'," _
               & "'" & rsTmp!拼音码 & "'," _
               & "'" & IIf(IsNull(rsTmp!厂商编码), "", rsTmp!厂商编码) & "' "
        strInsert = strInsert & strTmp & Chr(13)
        '更新数据
        With cmdInsert
            .ActiveConnection = gcnOutside
            .CommandText = strInsert
            .Execute
        End With
        If blnBegin Then
            gcnOutside.CommitTrans
        End If
        blnBegin = False
        rsTmp.MoveNext
    Loop
    rsTmp.Close

    Screen.MousePointer = vbDefault
    MsgBox "传送成功！", vbInformation, GSTR_SYSNAME
    Exit Sub

errHand:
    gcnOutside.RollbackTrans
    blnBegin = False
    Call OutPutLog("上传药品剂型异常：" & Err.Description)
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    lvwModel.View = lvwReport
'    Call InitLvwModel
End Sub

Public Sub InitLvwModel() '(ByVal cnZLHIS As ADODB.Connection)
    Dim rsTmp As New ADODB.Recordset
    Dim itmTmp As ListItem
    'Set rsTmp = zlDatabase.OpenSQLRecord("select distinct 药品剂型 from 药品特性", Me.Caption)
    If Me.mcnHIS.State = adStateClosed Then Me.mcnHIS.Open
    rsTmp.Open "select distinct 药品剂型 from 药品特性", Me.mcnHIS
    If Not rsTmp.EOF Then
        With lvwModel
            .ListItems.Clear
            .ColumnHeaders.Add 1, "K_Choose", "选择剂型", lvwModel.Width - 400
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                Set itmTmp = .ListItems.Add(, , rsTmp!药品剂型)
                rsTmp.MoveNext
            Loop
        End With
    End If
    rsTmp.Close
End Sub

Private Sub lblAll_Click()
    Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        lvwModel.ListItems(i).Checked = True
    Next
End Sub

Private Sub lblUnAll_Click()
    Dim i As Integer
    For i = 1 To lvwModel.ListItems.Count
        lvwModel.ListItems(i).Checked = False
    Next
End Sub

