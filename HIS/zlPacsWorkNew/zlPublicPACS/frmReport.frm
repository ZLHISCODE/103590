VERSION 5.00
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   0  'None
   Caption         =   "报告图像"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRich 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   120
      ScaleHeight     =   3150
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   240
      Width           =   4830
      Begin zlRichEditor.Editor edtThis 
         Height          =   2580
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4551
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function zlInitModule(ByVal lngRecordId As Long) As Long
    Call zlRefresh(lngRecordId)
    
    zlInitModule = Me.hWnd
End Function

Private Sub zlRefresh(ByVal lngRecordId As Long, Optional ByVal blnMoved As Boolean)
'功能：刷新病历显示内容；
'参数：lngRecordId：电子病历记录ID
    Dim rs As New ADODB.Recordset
    Dim collFile As New Collection, lngLen1 As Long, lngLen2 As Long, i As Integer, lngFileID As Long, strIDs As String, lngStart As Long, StrKey As String
    
    On Error GoTo errHand
    If lngRecordId = 0 Then Exit Sub
    
    If SetRichDocsPos(lngRecordId) Then Exit Sub
        
    '共享文档加载
    gstrSQL = "Select Count(C.Id) As 数目, c.病人ID,c.主页ID, c.文件id, c.创建时间" & vbNewLine & _
            "From 病历文件列表 F, 病历文件列表 B, 电子病历记录 C" & vbNewLine & _
            "Where f.种类 = b.种类 And f.页面 = b.页面 And b.Id = c.文件id And c.Id = [1]" & vbNewLine & _
            "Group By c.病人ID,c.主页ID, c.文件id, c.创建时间"
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    
    If rs.RecordCount <= 0 Then Exit Sub
    
    lngFileID = rs!文件ID
    edtThis.Freeze
    edtThis.ReadOnly = False
    edtThis.ForceEdit = True
    edtThis.InProcessing = True
    edtThis.Tag = "LoadFile"
    edtThis.NewDoc
    
    '读取RTF文件
    Call ReadRTF(edtThis, lngRecordId, True, blnMoved)
    
    If lngRecordId > 0 Then
        '设置页面格式
        Dim mEPRFileInfo As Object
        
        Set mEPRFileInfo = CreateObject("zlRichEPR.cEPRFileDefineInfo")
        
        gstrSQL = "Select a.格式 From 病历页面格式 a, 病历文件列表 b" & _
                " Where b.id=[1] And a.种类 = b.种类 And a.编号 = b.页面"
        Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
        If Not rs.EOF Then
            mEPRFileInfo.格式 = gobjComLib.zlCommFun.Nvl(rs("格式").Value)
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
            Me.edtThis.ResetWYSIWYG
        End If
        Set mEPRFileInfo = Nothing
    End If
    
    edtThis.SelStart = 0
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    Call SetRichDocsPos(lngRecordId)
    
    gobjComLib.zlCommFun.StopFlash
    Exit Sub
errHand:
    gobjComLib.zlCommFun.StopFlash
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
    
    On Error Resume Next

    edtThis.SelStart = 0
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    err.Clear
End Sub

Private Function SetRichDocsPos(ByVal lngRecordId As Long) As Boolean
    '通过ID先定位，无法定位时再加载
    Dim lngKSS As Long, lngKSE As Long, lngKES As Long, lngKEE As Long, blnNeed As Boolean, lngKey As Long, lngLen As Long, i As Integer
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen
        If FindNextKey(edtThis, i, "F", lngKey, lngKSS, lngKSE, lngKES, lngKEE, blnNeed) Then
            If edtThis.Range(lngKSE, lngKES).Text = lngRecordId Then
                edtThis.Range(lngKEE + 1, lngKEE + 1).Selected
                SetRichDocsPos = True
                Exit Function
            End If
            i = lngKEE
        Else
            Exit Function
        End If
    Next
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.picRich.Left = 0
    Me.picRich.Top = 0
    Me.picRich.Width = Me.ScaleWidth
    Me.picRich.Height = Me.ScaleHeight
End Sub

Private Sub picRich_Resize()
    On Error Resume Next
    
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub
