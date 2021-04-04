VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccoutChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "帐套选择"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "FrmAccoutChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView LvwSelect 
      Height          =   1005
      Index           =   0
      Left            =   1380
      TabIndex        =   2
      Top             =   -600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1773
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "Img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   4860
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccoutChoose.frx":062A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   1
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2910
      TabIndex        =   0
      Top             =   1860
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "FrmAccoutChoose.frx":0C64
      Top             =   60
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "    发现你同时具有以下系统多个帐套的权限，请选择本次操作的帐套："
      Height          =   405
      Left            =   990
      TabIndex        =   4
      Top             =   60
      Width           =   4455
   End
   Begin VB.Label LblNote 
      AutoSize        =   -1  'True
      Caption         =   "医院信息系统"
      Height          =   180
      Index           =   0
      Left            =   1350
      TabIndex        =   3
      Top             =   -780
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "FrmAccoutChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsSystems As ADODB.Recordset
Private mstrSQL As String
Private mstrCodes As String
Private mstrComponent As String
Private mlngCur As Long
Private mintCurTab As Integer
Private mblnMutil As Boolean
Private mblnMutilSys As Boolean

Public BlnSelect As Boolean

Private Sub Cmd取消_Click()
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    '产生SQL
    Dim lvwThis As Control, LvwItem As ListItem
    
    For Each lvwThis In Me.Controls
        If TypeName(lvwThis) = "ListView" Then
            If lvwThis.Index <> 0 Then
                mstrSQL = mstrSQL & IIf(mstrSQL = "", "", ",") & "'" & lvwThis.SelectedItem.Tag & "'"
            Else
                For Each LvwItem In lvwThis.ListItems
                    mstrSQL = mstrSQL & IIf(mstrSQL = "", "", ",") & "'" & LvwItem.Tag & "'"
                Next
            End If
        End If
    Next
    
    BlnSelect = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If BlnSelect = False Then
        Dim LngStyle As Long
        LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hwnd, 0 '先隐藏
        ShowWindow Me.hwnd, 1 '再显示
    End If
End Sub

Private Sub Form_Load()
    Me.Hide
    mblnMutilSys = False
    BlnSelect = False
    
    mstrComponent = GetSetting("ZLSOFT", "注册信息", "本机部件", "")

    '打开记录集，如果无多帐套，则退出
    Set mrsSystems = gclsBusiness.GetAccount(mstrComponent)
    With mrsSystems
        mintCurTab = 0
        mstrCodes = ""
        
        Do While Not .EOF
            '检测该系统是否有多帐套,否则插入Index=0的Listview;是则增加Listview,并插
            mblnMutil = False
            mlngCur = .AbsolutePosition
            If mstrCodes <> !名称 Then
                mstrCodes = !名称
                .Filter = "名称='" & mstrCodes & "'"
                mblnMutil = (.RecordCount > 1)
                If mblnMutilSys = False Then mblnMutilSys = mblnMutil
                
                If mblnMutil Then
                    mintCurTab = mintCurTab + 1
                    Load LvwSelect(mintCurTab)
                    With LvwSelect(mintCurTab)
                        .ListItems.Clear
                        .Left = LvwSelect(mintCurTab - 1).Left
                        .Top = LvwSelect(mintCurTab - 1).Top + 1400
                        .Width = LvwSelect(mintCurTab - 1).Width
                        .Height = LvwSelect(mintCurTab - 1).Height
                        .Visible = True
                    End With
                    Load LblNote(mintCurTab)
                    With LblNote(mintCurTab)
                        .Left = LblNote(mintCurTab - 1).Left
                        .Top = LblNote(mintCurTab - 1).Top + 1400
                        .Width = LblNote(mintCurTab - 1).Width
                        .Height = LblNote(mintCurTab - 1).Height
                        .Visible = True
                        .Caption = mstrCodes
                    End With
                    
                    '插入记录
                    Do While Not .EOF
                        LvwSelect(mintCurTab).ListItems.Add , "K_" & LvwSelect(mintCurTab).ListItems.Count + 1, mstrCodes & IIf(Val(!编号) = 0, "", "（" & Val(!编号) & "）"), 1
                        LvwSelect(mintCurTab).ListItems("K_" & LvwSelect(mintCurTab).ListItems.Count).Tag = !系统
                        .MoveNext
                    Loop
                Else
                    '插入记录到LvwSelect(0)
                    LvwSelect(0).ListItems.Add , "K_" & LvwSelect(0).ListItems.Count + 1, mstrCodes & IIf(Val(!编号) = 0, "", "（" & Val(!编号) & "）"), 1
                    LvwSelect(0).ListItems("K_" & LvwSelect(0).ListItems.Count).Tag = !系统
                End If
            End If
                
            .Filter = 0
            .MoveFirst
            .Move mlngCur - 1
            .MoveNext
        Loop
        
        With Cmd确定
            .Top = LvwSelect(mintCurTab).Top + LvwSelect(mintCurTab).Height + 150
        End With
        Cmd取消.Top = Cmd确定.Top
        
        Me.Height = Me.Cmd确定.Top + Me.Cmd确定.Height + 550
    End With
    
    mstrSQL = ""
    If mblnMutilSys = False Then Cmd确定_Click
End Sub

Public Function Show_me() As String
    On Error Resume Next
    
    Me.Show 1
    Show_me = mstrSQL
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mrsSystems = Nothing
End Sub

Private Sub LvwSelect_DblClick(Index As Integer)
    LvwSelect_KeyDown Index, vbKeyReturn, 0
End Sub

Private Sub LvwSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < LvwSelect.Count - 1 Then
            LvwSelect(Index + 1).SetFocus
        Else
            Cmd确定.SetFocus
        End If
    End If
End Sub
