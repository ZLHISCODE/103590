VERSION 5.00
Begin VB.Form frmBzcxxz铜仁 
   Caption         =   "病种重新选择"
   ClientHeight    =   1365
   ClientLeft      =   4200
   ClientTop       =   4155
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4710
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txt病种 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "该病人病种已丢失请重新选择!"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmBzcxxz铜仁"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mintTimes As Integer
Private mblnOK As Boolean
Private mint场合 As Integer
Private mlng病种ID As Long
Private mstr病种编码 As String

Private Sub cmdCancel_Click()
txt病种.Tag = ""
mlng病种ID = 0
Unload Me
End Sub


Private Sub cmdOK_Click()
    mlng病种ID = Val(txt病种.Tag)
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd病种_Click()
    On Error GoTo errHandle
    Dim rs病种 As ADODB.Recordset
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病',0,'普通病') as 类别 " & _
            " From 保险病种 A where A.险类=[1] And A.类别 IN ([2])"
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "重新选择病种", TYPE_铜仁, CStr(IIf(mint场合 = 0, "0,1,2", "0")))
    
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_铜仁, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
            txt病种.Text = rs病种("名称")
            txt病种.Tag = rs病种("ID")
            mstr病种编码 = rs病种("编码")
        End If
    End If
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetPatient(ByVal int场合 As Integer, ByVal bln修改密码 As Boolean, 病种ID As Long) As Boolean
    Me.Show vbModal
    If mblnOK = True Then
        病种ID = mlng病种ID
    End If
    GetPatient = mblnOK
End Function

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    Dim str前   As String
    
    On Error GoTo errHandle
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
        txt病种.Tag = ""
    If txt病种.Text = "" Or txt病种.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    str前 = IIf(gstrMatchMethod = 0, "%", "")
    strText = txt病种.Text
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特殊病',0,'普通病') 类别 " & _
             "   FROM 保险病种 A WHERE A.险类=[1] And A.类别 IN ([2]) And (" & _
             "   A.编码 like '" & str前 & "' || [3] || '%' or A.名称 like '" & str前 & "' || [3] || '%' or A.简码 like '" & str前 & "' || [3] || '%')"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_铜仁, IIf(mint场合 = 0, "0,1,2", "0"), strText)
    
    If rsTmp.RecordCount > 0 Then
        '出现选择器
        If rsTmp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_铜仁, rsTmp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Text = rsTmp("名称")
        txt病种.Tag = rsTmp("ID")
        mstr病种编码 = rsTmp("编码")
       ' SendKeys "{TAB}"
       cmdOK.SetFocus
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
