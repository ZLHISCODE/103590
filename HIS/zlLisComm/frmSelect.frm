VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择仪器"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   2
      Top             =   2820
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   780
      TabIndex        =   1
      Top             =   2775
      Width           =   1100
   End
   Begin VB.ListBox lst可加入仪器 
      Height          =   2580
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   5955
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean

Public Function Select仪器() As Boolean
    mblnOK = False
    Me.Show vbModal
    Select仪器 = mblnOK
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim lngID As Long
    Dim i As Integer
    Dim blnAdd As Boolean
    If lst可加入仪器.ListIndex >= 0 Then
        lngID = lst可加入仪器.ItemData(lst可加入仪器.ListIndex)
        blnAdd = False
        For i = LBound(g仪器) To UBound(g仪器)
            If g仪器(i).ID <= 0 Then
                blnAdd = True
                Exit For
            End If
        Next
        If Not blnAdd Then
            ReDim Preserve g仪器(UBound(g仪器) + 1)
            i = UBound(g仪器)
            blnAdd = True
        End If
        If blnAdd Then
            g仪器(i).ID = lngID
            g仪器(i).COM口 = 0
            g仪器(i).类型 = 0
            g仪器(i).波特率 = 9600
            g仪器(i).数据位 = 8
            g仪器(i).停止位 = 1
            g仪器(i).校验位 = "N"
            g仪器(i).握手 = 0
            g仪器(i).字符模式 = 0
            g仪器(i).IP = "127.0.0.1"
            g仪器(i).IP端口 = "6666"
            g仪器(i).主机 = 0
            g仪器(i).SaveAsID = 0
            g仪器(i).自动应答 = "0"
            g仪器(i).可发已核标本 = "1"
            g仪器(i).通讯目录 = App.Path & "\Dev_" & lngID
            g仪器(i).自动审核人 = ""
            g仪器(i).自动计算质控 = 0
            g仪器(i).另存为通道码 = 0
            mblnOK = True
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer, lngCount As Long
    Dim blnAdd As Boolean
    
    Set rsTmp = GetDevices
    lst可加入仪器.Clear
    
    If rsTmp Is Nothing Then Exit Sub
    lngCount = 0
    Do Until rsTmp.EOF
        lngCount = lngCount + 1
        blnAdd = True
        For i = LBound(g仪器) To UBound(g仪器)
            If g仪器(i).ID = rsTmp!ID Then
                blnAdd = False
                Exit For
            End If
        Next
        '控制仪器数量
        If gstr仪器数量 <> "" Then
            If lngCount > Val(gstr仪器数量) Then
                blnAdd = False
            End If
        End If
        
        If blnAdd Then
            lst可加入仪器.AddItem "(" & rsTmp!编码 & ")" & rsTmp!名称
            lst可加入仪器.ItemData(lst可加入仪器.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Loop
End Sub
