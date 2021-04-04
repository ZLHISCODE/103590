VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSfyInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "FrmSfyInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   360
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出"
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   6240
      Width           =   1100
   End
   Begin MSComctlLib.ListView LviewInfo 
      Height          =   5220
      Left            =   240
      TabIndex        =   2
      Top             =   576
      Width           =   8088
      _ExtentX        =   14261
      _ExtentY        =   9208
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   10398
      MultiRow        =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5520
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
End
Attribute VB_Name = "FrmSfyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public img As DicomImage

Private Sub cmdExport_Click()
    
    Dim strFileName As String
    Dim i As Integer
    Dim strInfo As String
    Dim iType As Integer
    
    dlgSave.Filter = "(*.TXT)|*.TXT"
    dlgSave.ShowSave
    strFileName = dlgSave.Filename
    If strFileName <> "" Then
        
        On Error GoTo err
        
        If Me.TabStrip1.SelectedItem.Index = 2 Then
            iType = 1
        Else
            iType = 0
        End If
        Open strFileName For Output As #1
        For i = 1 To LviewInfo.ListItems.Count
            If iType = 1 Then
                strInfo = LviewInfo.ListItems(i).Text & " " & left(LviewInfo.ListItems(i).SubItems(1) & Space(30), 30) & " : " & LviewInfo.ListItems(i).SubItems(2) & " " & LviewInfo.ListItems(i).SubItems(3) & vbCrLf
            Else
                strInfo = LviewInfo.ListItems(i).Text & " " & left(LviewInfo.ListItems(i).SubItems(1) & Space(30), 30) & " : " & LviewInfo.ListItems(i).SubItems(2)
            End If
            Print #1, strInfo
        Next i
        Close #1
    End If
    Exit Sub
err:
    Close #1
    
End Sub

'退出
Private Sub cmdOK_Click()
    Unload Me
End Sub
'初使化表头
Sub InitializeListHead()
    Dim i As Integer
    i = Me.LviewInfo.width / 5
    If Me.TabStrip1.SelectedItem.Index = 2 Then
        With Me.LviewInfo.ColumnHeaders
            .Clear
            .Add , "A" & 1, "组-元素", i
            .Add , "A" & 2, "英文名", i * 1.5
            .Add , "A" & 3, "中文名", i
            .Add , "A" & 4, "病人信息", i * 2
        End With
    Else
        With Me.LviewInfo.ColumnHeaders
            .Clear
            .Add , "A" & 1, "组-元素", i
            .Add , "A" & 2, "英文名", i * 1.5
            .Add , "A" & 3, "值", i * 3
        End With
    End If
End Sub
'读入信息
Sub LoadInfo(ImgView As DicomImage, StrFiltrate As String)
    Dim RS As New ADODB.Recordset
    Dim Lvt As ListItem
    Dim LngBRInfo As Long
    Dim StrTmp As Variant
    Dim strSQL As String
    Dim i As Integer
    
'    Select Case StrFiltrate
'        Case "全部信息"
'            strSQL = "select ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算 from 图像信息表 "
'        Case "常用信息"
'            strSQL = "select ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算 from 图像信息表 "
'    End Select
    If blLocalRun = True Then
        strSQL = "select ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算 from 图像信息表 "
        RS.Open strSQL, cnAccess
    Else
        strSQL = "select ID,开始地址,结束地址,英文名称,中文名称,中文简称,英文简称,常用, 被选用,位置,角内序号,使用计算 from 影像图像信息表 "
        Set RS = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    Me.LviewInfo.view = lvwReport
    Me.LviewInfo.ListItems.Clear
    '读入数据
    Do While Not RS.EOF
        Set Lvt = Me.LviewInfo.ListItems.Add(, "BR" & RS("id"), "(" & RS("开始地址") & "," & RS("结束地址") & ")")
        Lvt.SubItems(1) = RS("英文名称")
        Lvt.SubItems(2) = RS("中文名称")
        If IsNull(ImgView.Attributes(Val("&H" & RS("开始地址")), "&H" & RS("结束地址"))) = False Then
            StrTmp = ImgView.Attributes(Val("&H" & RS("开始地址")), Val("&H" & RS("结束地址")))
            If IsArray(StrTmp) = True Then
                For i = 1 To UBound(StrTmp)
                    If Lvt.SubItems(3) <> "" Then
                        Lvt.SubItems(3) = Lvt.SubItems(3) & ";" & StrTmp(i)
                    Else
                        Lvt.SubItems(3) = StrTmp(i)
                    End If
                Next
            Else
                Lvt.SubItems(3) = StrTmp
            End If
        End If
        RS.MoveNext
    Loop
    RS.Close

End Sub

Private Sub Form_Load()
    '初使化Tab
    Me.TabStrip1.Tabs(1).Caption = "全部信息"
    Me.TabStrip1.Tabs.Add , "B", "常用信息"
    If img.Attributes(&H8, &H60).Exists Then
        If UCase(img.Attributes(&H8, &H60).Value) = "PR" Then
            Me.TabStrip1.Tabs.Add , "C", "PR信息"
        End If
    End If
    
    '写入信息
    InitializeListHead
'    LoadInfo img, Me.TabStrip1.SelectedItem
     AppendAttributes "", "", img.Attributes
End Sub
'选中当前信息
Private Sub TabStrip1_Click()
    If Me.TabStrip1.SelectedItem.Index = 2 Then
        InitializeListHead
        LoadInfo img, Me.TabStrip1.SelectedItem
    ElseIf Me.TabStrip1.SelectedItem.Index = 3 Then
        
        InitializeListHead
        Me.LviewInfo.ListItems.Clear
        If Not img.PresentationState Is Nothing Then
            AppendAttributes "", "", img.PresentationState.Attributes
        End If
    Else
        InitializeListHead
        Me.LviewInfo.ListItems.Clear
        AppendAttributes "", "", img.Attributes
    End If
End Sub
Sub AppendAttributes(ByRef list, prefix, ByRef ob As Object)
    Dim at As DicomAttribute
    Dim s As DicomDataSets
    Dim i As Integer
    Dim v As Variant
    Dim objItem As ListItem
    Static j As Integer
    Dim tmpStr As String
    For Each at In ob
        list = list & prefix & "(" & hex4(at.Group) & "," & hex4(at.Element) & ") : "
        Set objItem = Me.LviewInfo.ListItems.Add(, "A" & j, prefix & "(" & hex4(at.Group) & "," & hex4(at.Element) & ") : ")
        list = list & left(at.Description & Space(30), 30) & ": "
        objItem.SubItems(1) = at.Description
        If (at.Group = &H7FE0) Then ' pixel data
            list = list & "Pixel data" & vbCrLf
        ElseIf (VarType(at.Value) = 9) Then ' i.e. a sequence
            Set s = at.Value
            list = list & "Sequence of " & s.Count & " items:" & vbCrLf
            For i = 1 To s.Count
                'list = list & prefix & ">---------------" & vbCrLf
                j = j + 1
                AppendAttributes list, prefix & ">", s(i).Attributes
            Next
            'list = list & prefix & ">---------------" & vbCrLf
        Else
            v = at.Value ' could be variant or array
            If (VarType(v) > 8192) Then ' i.e. an array
                list = ""
                list = list & "Multiple values :"
                If UBound(v, 1) > 32 Then
                    list = list & "Array of " & UBound(v, 1) & " elements"
                Else
                    For i = LBound(v, 1) To UBound(v, 1)
                        list = list & v(i)
                        If i <> UBound(v, 1) Then list = list & " : "
                    Next
                End If
                list = list & vbCrLf
                objItem.SubItems(2) = list
            Else
                list = list & v & vbCrLf
                objItem.SubItems(2) = IIf(IsNull(v), "", v)
            End If
        End If
        j = j + 1
    Next
End Sub

Function hex4(ByVal v As Long) As String
    hex4 = Right("000" & Hex(v), 4)
End Function


