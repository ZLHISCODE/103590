VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmListener 
   Caption         =   "中联HL7侦听服务"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   6615
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      Height          =   350
      Left            =   1680
      TabIndex        =   9
      Text            =   "1024"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   350
      Left            =   5880
      TabIndex        =   7
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtLogPath 
      Height          =   350
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "返回ACK响应"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "发送"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "连接Listener"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtResult 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "侦听"
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   1100
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   2002
      LocalPort       =   2001
   End
   Begin VB.Label Label2 
      Caption         =   "侦听端口号"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "LOG日志文件"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "frmListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim strData As String
Dim strACK As String

Private Sub cmdBrowse_Click()
    Me.dlgBrowse.Filter = "Log File(*.log)|*.log|All(*.*)|*.*"
    Me.dlgBrowse.ShowOpen
    txtLogPath = Me.dlgBrowse.FileName
End Sub

Private Sub cmdListen_Click()
    On Error GoTo ListenErr
    If Me.cmdListen.Caption = "侦听" Then
        Me.sckServer(0).Close
        'Me.Winsock1.LocalPort = Val(txtPort.Text)
        Me.sckServer(0).Bind Val(txtPort.Text)
        Me.sckServer(0).Listen
        If Me.sckServer(0).State = sckListening Then
            subAddLog " 开始侦听端口: " & Me.sckServer(0).LocalPort
        Else
            subAddLog " 侦听端口 " & Me.sckServer(0).LocalPort & " 出错"
        End If
        Me.cmdListen.Caption = "停止"
    Else
        If Me.sckServer(0).State = sckListening Then
            subAddLog " 正在停止侦听"
            Me.sckServer(0).Close
            If Me.sckServer.Count = 2 Then
                Me.sckServer(1).Close
                Unload Me.sckServer(1)
            End If
            If Me.sckServer(0).State = sckClosed Then subAddLog " 成功停止侦听"
        Else
            subAddLog " 侦听出错"
        End If
        Me.cmdListen.Caption = "侦听"
    End If
    Exit Sub
ListenErr:
    If err.Number = 10048 Then
        subAddLog " 端口已被占用"
    End If
End Sub


Private Sub Command2_Click()
    Me.Winsock2.Connect "127.0.0.1", 8088
End Sub

Private Sub Command3_Click()
    Me.Winsock2.SendData strData
    Dim strDa As String
    strDa = Chr$(11) & "MSH" '|^~\&|MESA_ADT|XYZ_ADMITTING|MESA_IS|XYZ_HOSPITAL|||ADT^A04|101102|P|2.3.1||||||||" & vbCrLf _
                            '& "EVN||200004211000||||200004210950 " & vbCrLf _
                            '& "PID|||583020^^^ADT1||WHITE^CHARLES||19980704|M||AI|7616 STANFORD AVE^^ST. LOUIS^MO^63130|||||||20-98-1701||||||||||||" & vbCrLf _
                            '& "PV1||E||||||5101^NELL^FREDERICK^P^^DR|||||||||||V1002^^^ADT1|||||||||||||||||||||||||200004210950||||||||" & vbCrLf
    Me.Winsock2.SendData strDa
End Sub

Private Sub Form_Load()
    Me.txtLogPath.Text = App.Path & "\接收日志.log"
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.sckServer(0).State <> sckClosed Then Me.sckServer(0).Close
End Sub


Private Sub sckServer_Close(Index As Integer)

    subAddLog " 关闭和 " & Me.sckServer(1).RemoteHostIP & " 之间的连接"
    
    Me.sckServer(1).Close
    Unload Me.sckServer(1)
End Sub

Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    If Me.sckServer.Count = 1 Then
        Load sckServer(1)
        sckServer(1).Accept requestID
    End If
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strMSGID As String
    
    Me.sckServer(Index).GetData strData, vbString
    subAddLog "接收到数据：" & vbCrLf & strData & vbCrLf
    subAddLog "接收到的数据内容为：" & vbCrLf & funcParseData(strData)
    If Me.Check1.Value = 1 Then
        strMSGID = GetMSGID(strData)
        strACK = Chr(11) & "MSH|^~\&|||||200702021000||ACK|" & strMSGID & "|P|2.4" & vbCrLf & _
                "MSA|AA|" & strMSGID & Chr(28) & Chr(13)
        Me.sckServer(Index).SendData strACK
        subAddLog "返回响应：" & vbCrLf & strACK & vbCrLf
    End If
End Sub


Private Sub subAddLog(strLog As String)

    Dim lngFileHandle As Long '文件句柄
    Dim fsObject As New Scripting.FileSystemObject
    Dim strShowText As String
    
    strShowText = vbCrLf & Date & " " & Time & strLog
    Me.txtResult.Text = Me.txtResult.Text & strShowText
    Me.txtResult.SelStart = Len(Me.txtResult.Text)
    On Error GoTo err
    If fsObject.FileExists(Me.txtLogPath.Text) = True Then
        lngFileHandle = FreeFile() '取得句柄
        Open Me.txtLogPath.Text For Append As lngFileHandle    '打开文件
        Print #lngFileHandle, strShowText    '写入文本
        Close lngFileHandle    '关闭文件
    End If
    Exit Sub
err:
    
End Sub

Private Function funcParseData(strData) As String
    Dim strSegment() As String
    Dim strField() As String
    Dim i As Integer
    
    If strData <> "" Then
        strSegment = Split(strData, Chr(13))
        For i = 0 To UBound(strSegment) - 1
        
            strField = Split(strSegment(i), "|")
            If Trim(strField(0)) = Chr(11) & "MSH" Then
                If UBound(strField) >= 9 Then
                    funcParseData = funcParseData & vbCrLf & "MSH-10，消息控制ID  = " & strField(9)
                End If
            End If
        
            If strField(0) = "PID" Or strField(0) = vbLf & "PID" Then
                If UBound(strField) >= 8 Then
                    funcParseData = funcParseData & vbCrLf & "PID-5，Patient Name = " & strField(5)
                    funcParseData = funcParseData & vbCrLf & "PID-8，Patient Sex = " & strField(8)
                End If
            End If
    
            If strField(0) = "OBR" Or strField(0) = vbLf & "OBR" Then
                If UBound(strField) >= 4 Then
                    funcParseData = funcParseData & vbCrLf & "OBR-4，医嘱内容 = " & strField(4)
                End If
            End If
            
            If strField(0) = "PV1" Or strField(0) = vbLf & "PV1" Then
                If UBound(strField) >= 19 Then
                    funcParseData = funcParseData & vbCrLf & "PV1-2，患者类别 = " & strField(2)
                    funcParseData = funcParseData & vbCrLf & "PV1-7，开嘱医生 = " & strField(7)
                    funcParseData = funcParseData & vbCrLf & "PV1-19，Visit Number = " & strField(19)
                End If
                If UBound(strField) > 44 Then
                    funcParseData = funcParseData & vbCrLf & "PV1-44，入院时间 = " & strField(44)
                End If
                
                
            End If
            
            If strField(0) = "ORC" Or strField(0) = vbLf & "ORC" Then
                If UBound(strField) >= 1 Then
                    funcParseData = funcParseData & vbCrLf & "ORC-1，医嘱控制码 = " & strField(1)
                    funcParseData = funcParseData & vbCrLf & "ORC-2，医嘱ID = " & strField(2)
                End If
            End If
    
            If strField(0) = "OBX" Or strField(0) = vbLf & "OBX" Then
                If UBound(strField) >= 3 Then
                    funcParseData = funcParseData & vbCrLf & "OBX-2，值类型 = " & strField(2)
                    funcParseData = funcParseData & vbCrLf & "OBX-3，观察标识符 = " & strField(3)
                End If
            End If
        Next i
    End If
End Function


Private Function GetMSGID(strData As String) As String
    Dim strSegment() As String
    Dim strField() As String
    Dim i As Integer
    
    'strSegment = Split(strData, Chr(13))
    'For i = 0 To UBound(strSegment) - 1
    If strData <> "" Then
        strField = Split(strData, "|")
        If Trim(strField(0)) = Chr(11) & "MSH" Then
            If UBound(strField) > 10 Then
                GetMSGID = strField(9)
            End If
        End If
    End If
   ' End If
End Function

