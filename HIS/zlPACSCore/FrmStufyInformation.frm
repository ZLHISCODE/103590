VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmStufyInformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人信息"
   ClientHeight    =   6768
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6768
   ScaleWidth      =   8580
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView LviewInfo 
      Height          =   5220
      Left            =   240
      TabIndex        =   2
      Top             =   576
      Width           =   8088
      _ExtentX        =   14266
      _ExtentY        =   9208
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
      _ExtentX        =   14732
      _ExtentY        =   10393
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
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3648
      TabIndex        =   0
      Top             =   6228
      Width           =   1100
   End
End
Attribute VB_Name = "FrmstufyInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public img As DicomImage

'退出
Private Sub CmdOK_Click()
    Unload Me
End Sub
'初使化表头
Sub InitializeListHead()
    Dim I As Integer
    I = Me.LviewInfo.width / 5
    With Me.LviewInfo.ColumnHeaders
        .Clear
        .Add , "A" & 1, "英文名称", I * 2
        .Add , "A" & 2, "中文名称", I
        .Add , "A" & 3, "病人信息", I * 2
    End With
End Sub
'读入信息
Sub LoadInfo(ImgView As DicomImage, StrFiltrate As String)
    Dim RS As New ADODB.Recordset
    Dim Lvt As ListItem
    Dim LngBRInfo As Long
    Dim StrTmp As Variant
    Dim StrSQL As String
    Select Case StrFiltrate
        Case "全部信息"
            StrSQL = "select * from 图像信息表 "
        Case "常用信息"
            StrSQL = "select * from 图像信息表 where 常用 <> 0 "
    End Select
        
    RS.Open StrSQL, cnAccess
    Me.LviewInfo.View = lvwReport
    Me.LviewInfo.ListItems.Clear
    '读入数据
    Do While Not RS.EOF
        Set Lvt = Me.LviewInfo.ListItems.Add(, "BR" & RS("id"), RS("英文名称"))
        Lvt.SubItems(1) = RS("中文名称")
        If IsNull(ImgView.Attributes(Val("&H" & RS("开始地址")), "&H" & RS("结束地址"))) = False Then
            StrTmp = ImgView.Attributes(Val("&H" & RS("开始地址")), Val("&H" & RS("结束地址")))
            If IsArray(StrTmp) = True Then
                For I = 1 To UBound(StrTmp)
                    If Lvt.SubItems(2) <> "" Then
                        Lvt.SubItems(2) = Lvt.SubItems(2) & ";" & StrTmp(I)
                    Else
                        Lvt.SubItems(2) = StrTmp(I)
                    End If
                Next
            Else
                Lvt.SubItems(2) = StrTmp
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
    '写入信息
    InitializeListHead
    LoadInfo img, Me.TabStrip1.SelectedItem
End Sub
'选中当前信息
Private Sub TabStrip1_Click()
    LoadInfo img, Me.TabStrip1.SelectedItem
End Sub
