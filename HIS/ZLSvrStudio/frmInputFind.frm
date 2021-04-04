VERSION 5.00
Begin VB.Form frmInputFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查找"
   ClientHeight    =   2385
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5595
   Icon            =   "frmInputFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4290
      TabIndex        =   3
      Top             =   1935
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找下一条(&N)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2715
      TabIndex        =   2
      Top             =   1935
      Width           =   1530
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1935
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   60
      Left            =   0
      TabIndex        =   4
      Top             =   1785
      Width           =   5565
   End
   Begin VB.ComboBox cboSource 
      Height          =   300
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   930
      Width           =   3435
   End
   Begin VB.Label lblComment 
      Caption         =   "    输入希望查找的基本字词的字词、拼音码或五笔码。如存在多条，可依序查找下一条，直到找到你希望查找的项目。"
      Height          =   525
      Left            =   885
      TabIndex        =   7
      Top             =   135
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "(共查找到10条，当前为第1条)"
      Height          =   180
      Left            =   870
      TabIndex        =   6
      Top             =   1455
      Width           =   2430
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frmInputFind.frx":058A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "查找内容(&F)"
      Height          =   180
      Left            =   885
      TabIndex        =   0
      Top             =   1005
      Width           =   990
   End
End
Attribute VB_Name = "frmInputFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsFind As New ADODB.Recordset
Private strFind As String
Private intCount As Integer
Private mfrmMain As Object
Private mbytMode As Byte

Public Function ShowFind(ByVal frmMain As Object, Optional ByVal bytMode As Byte = 1) As Boolean
    
    Set mfrmMain = frmMain
    mbytMode = bytMode
    
    Me.Show 1, frmMain
    
End Function

Private Sub cboSource_Click()
    If Trim(cboSource.Text) <> "" Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub cboSource_GotFocus()
    cboSource.SelStart = 0
    cboSource.SelLength = 100
End Sub

Private Sub cboSource_KeyPress(KeyAscii As Integer)
    
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cboSource_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(cboSource.Text) <> "" Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strTemp As String
    Dim strKey As String
    
    If Trim(cboSource.Text) = "" Then
        MsgBox "请输入查找的内容", vbExclamation, gstrSysName
        cboSource.SetFocus
        Exit Sub
    End If
    
    If InStr(cboSource.Text, "'") > 0 Then
        MsgBox "输入的内容有非法字符 '", vbExclamation, gstrSysName
        cboSource.SetFocus
        Exit Sub
    End If
    
    strTemp = ""
    For intCount = 0 To cboSource.ListCount
        strTemp = strTemp & ";" & cboSource.List(intCount)
    Next
    
    If InStr(1, strTemp, ";" & Trim(cboSource.Text)) = 0 Then
        cboSource.AddItem Trim(cboSource.Text), 0
    End If
    
    gstrSQL = "select 字词,输入码,输入法 " & _
            " from zlwordbasic " & _
            " where 是否字=" & mbytMode & " and (输入码 like '%" & Trim(cboSource.Text) & "%'" & _
                    "or 字词 like '%" & Trim(cboSource.Text) & "%')  ORDER BY 输入法,输入码"
    
    Err = 0
    On Error GoTo errHand
    
    With rsFind
        If strFind <> gstrSQL Or .State <> adStateOpen Then
            If .State = adStateOpen Then .Close
            rsFind.Open gstrSQL, gcnOracle
            
            If .EOF Then
                MsgBox "不存在查找的内容！", vbExclamation, gstrSysName
                
                .Close
                cmdFind.Enabled = False
                lblNote.Caption = ""
                cboSource.SetFocus
                
                Exit Sub
            End If
            strFind = gstrSQL
        Else
            .MoveNext
            If .EOF Then
                MsgBox "已查找到最后一条项目！", vbExclamation, gstrSysName
                
                .Close
                cboSource.Text = ""
                cmdFind.Enabled = False
                lblNote.Caption = ""
                cboSource.SetFocus
                
                Exit Sub
            End If
        End If
        
        lblNote.Caption = "(共查找到" & .RecordCount & "条，当前为第" & .AbsolutePosition & "条)"
        
        If rsFind("输入法").Value = 1 Then
            strKey = "P" & Left(rsFind("输入码").Value, 1)
        Else
            strKey = "W" & Left(rsFind("输入码").Value, 1)
        End If
        
        Call mfrmMain.LocationItem(strKey, rsFind("字词").Value, rsFind("输入码").Value)
        
    End With
    
    Exit Sub
    
errHand:
    MsgBox "查找基本字词失败！" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name    '
End Sub

Private Sub Form_Load()
    Dim lngLoop As Long
    Dim strSectoin  As String
    Dim str条件  As String
    
    cboSource.Clear
    
    strSectoin = "实用工具\输入码工具\" & gstrUserName & "\历史查找"
    
    For lngLoop = 0 To CLng(Val(GetSetting("ZLSOFT", strSectoin, "查找项数", "0")))
        str条件 = GetSetting("ZLSOFT", strSectoin, "查找内容" & lngLoop, "")
        If str条件 <> "" Then cboSource.AddItem str条件
    Next
    
    strFind = ""
    lblNote.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngLoop As Long
    Dim strSectoin As String
    
    '进行了常用条件的更改
    strSectoin = "实用工具\输入码工具\" & gstrUserName & "\历史查找"
    
    On Error Resume Next '如果没有该键值，就会出错
    DeleteSetting "ZLSOFT", strSectoin '删除以前的条件
    On Error GoTo 0
    
    Call SaveSetting("ZLSOFT", strSectoin, "查找项数", IIf(cboSource.ListCount > 10, 10, cboSource.ListCount))
    
    For lngLoop = 0 To cboSource.ListCount
        
        If lngLoop > 10 Then Exit For
        
        Call SaveSetting("ZLSOFT", strSectoin, "查找内容" & lngLoop, cboSource.List(lngLoop))

    Next
    
End Sub
