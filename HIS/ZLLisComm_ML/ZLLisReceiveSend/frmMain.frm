VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   Caption         =   "ͨѶ"
   ClientHeight    =   7755
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   12195
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   12195
   StartUpPosition =   2  '��Ļ����
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picIcon 
      Height          =   285
      Left            =   645
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame fraWE 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2835
      Left            =   7545
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   15
      Width           =   45
   End
   Begin VB.PictureBox picA 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   210
      Picture         =   "frmMain.frx":29F2
      ScaleHeight     =   540
      ScaleWidth      =   585
      TabIndex        =   4
      Top             =   2265
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picB 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   195
      Picture         =   "frmMain.frx":53E4
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   3
      Top             =   3045
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picC 
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   135
      Picture         =   "frmMain.frx":7DD6
      ScaleHeight     =   570
      ScaleWidth      =   540
      TabIndex        =   2
      Top             =   3645
      Visible         =   0   'False
      Width           =   540
   End
   Begin RichTextLib.RichTextBox rtxtLogHex 
      Height          =   5280
      Left            =   1365
      TabIndex        =   0
      Top             =   1320
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   9313
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":A7C8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer TimAutoAnswer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   315
      Top             =   6585
   End
   Begin MSCommLib.MSComm COM 
      Left            =   135
      Top             =   5535
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock WSK 
      Left            =   225
      Top             =   5085
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtxtLogTxt 
      Height          =   5145
      Left            =   7830
      TabIndex        =   1
      Top             =   1395
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   9075
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":A865
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer TimSendCmd 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   4830
   End
   Begin RichTextLib.RichTextBox rtxLogDebug 
      Height          =   3735
      Left            =   1710
      TabIndex        =   7
      Top             =   120
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   6588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmMain.frx":A902
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "����(&D)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuH2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&E)"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Popu"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "��С(&I)"
      End
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "�ָ�(&R)"
      End
      Begin VB.Menu mnuH1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "�ر�(&C)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrBuffer As String
Private mstrStat  As String '״̬, �գ���ʼ״̬ "����"����һ��ͨѶ�������� "����"����һ��ͨѶ�Ƿ�������
Private mLasterTime As Date '�ϴ�ͨѶʱ��
Private mstrTimeSendCmd As String '��ʱ����ָ��

Private mblnListen As Boolean '���ӶϿ����Ƿ���������,��������ģʽ�£��Է��Ͽ����Ӻ����¼���

'Public Event CustEvent(ByVal strMessage As String)

Private mintSendStep As Integer '˫��ͨѶ�ã�һ���ʾͨѶ����
Private mstrResponse As String  '˫��ͨѶ�ã���ʾ���յ�����Ϣ
Private mLastSampleInfo As String '˫��ͨѶ�ã����͵�������Ϣ
Private mblnUndo As Boolean       '˫��ͨѶ�ã�Ҫ���Ĳ���
Private mintType As Integer       '˫��ͨѶ�ã��Ƿ���걾

Private mstrTip As String
Private mblnAutoAnswer As Boolean '�Ƿ����ö�ʱӦ��

Private Sub COM_OnComm()
    '�ڽ�������ʱ�������ö�ʱ��
    Dim byt_Bit() As Byte '-���ն���������
    Dim strData As String
    Dim i As Integer, str�� As String

    
    Select Case COM.CommEvent
        Case comEvSend ' 1 �����¼���
        Case comEvReceive '2 �����¼���
            
            TimSendCmd.Enabled = False
            TimAutoAnswer.Enabled = False
            strData = ""
            
            If mstrStat = "" Or mstrStat = "����" Then
                str�� = IIf(mstrStat = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " ��)")
                Call addLog("TITLE", "<== Receive " & Format(Now, "yyyy-MM-dd HH:mm:ss") & " " & str��)
                mLasterTime = Now
                mstrStat = "����"
            End If
            
            If COM.InputMode = comInputModeText Then
                strData = COM.Input
                mstrBuffer = mstrBuffer & strData
                If Len(strData) > 0 Then Call addLog("TXT", strData)
            Else
                byt_Bit = COM.Input
                For i = LBound(byt_Bit) To UBound(byt_Bit)
                    strData = strData & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
                Next
                mstrBuffer = mstrBuffer & strData
                If Len(strData) > 0 Then Call addLog("HEX", strData)
                
            End If
            '�������յ�������
            mstrBuffer = Decode(mstrBuffer)
            
    
            TimSendCmd.Enabled = True
            If mblnAutoAnswer = True Then TimAutoAnswer.Enabled = True
            
        Case comEvCTS '3 clear-to-send �߱仯��
        
        Case comEvDSR '4 data-set ready �߱仯��
        
        Case comEvCD '5 carrier detect �߱仯��
        
        Case comEvRing '6 �����⡣
        
        Case comEvEOF '7 �ļ�������
    End Select

    'If App.EXEName <> "VB6.EXE" Then Call ChangeICon

End Sub

Private Sub addDebug(ByVal strType As String, ByVal strData As String)
    Dim strFilename As String
    On Error GoTo errH
    
    Me.rtxLogDebug = Me.rtxLogDebug.Text & strData
    Me.rtxLogDebug.SelStart = Len(Me.rtxLogDebug.Text)
        
    If UBound(Split(Me.rtxLogDebug.Text, vbNewLine)) > 1000 Then
        If mnuDebug.Checked Then
            strFilename = App.Path & "\������־_" & Format(Now, "yyyyMMddHHMM") & ".log"
            If gFileObject.FileExists(strFilename) Then gFileObject.DeleteFile strFilename
            Call rtxLogDebug.SaveFile(strFilename, rtfText)
        End If
        Me.rtxLogDebug.Text = ""
        
    End If
    
    Exit Sub
errH:
    WriteErrLog "addDebug " & strType, strData, Err.Description
End Sub
Private Sub addLog(ByVal strType As String, ByVal strData As String)
        '���յ���������ʾ������ؼ�
        'strTYPE �� HEX �����������HEX��ʽ������,TXT�����������TXT��ʽ������ ,TITLE-�����������������Ϣ
        '
        Dim strTxt As String, lngCount As Long, strBit As String
        Dim strHex As String
        Dim txtStream As TextStream, strFilename As String
        Dim fileTmp As File
    
   
        On Error GoTo errH
    
100     If strType = "HEX" Then
            '--- TXT
102         strTxt = ""
104         For lngCount = 1 To Len(strData) / 3
106             strBit = Chr("&H" & Mid(strData, 2, 2))
108             strTxt = strTxt & strBit
            Next
110         Me.rtxtLogTxt.Text = Me.rtxtLogTxt.Text & strTxt
112         Me.rtxtLogTxt.SelStart = Len(Me.rtxtLogTxt.Text)
            '--- HEX
114         strHex = Replace(strData, ",", " ")
116         Call ShowHex(strHex)
        
        
118     ElseIf strType = "TXT" Then
            '--- txt
120         strTxt = strData
122         Me.rtxtLogTxt.Text = Me.rtxtLogTxt.Text & strTxt
124         Me.rtxtLogTxt.SelStart = Len(Me.rtxtLogTxt.Text)
        
            '-- HEX
126         strHex = ""
128         For lngCount = 1 To Len(strData)
130             strBit = Mid(strData, lngCount, 1)
132             strBit = Hex(Asc(strBit))
134             If Len(strBit) = 1 Then
136                 strBit = "0" & strBit
138             ElseIf Len(strBit) = 4 Then
140                 strBit = Mid(strBit, 1, 2) & " " & Mid(strBit, 3, 2)
                End If
142             strHex = strHex & " " & strBit
            Next
144         Call ShowHex(strHex)
        
146     ElseIf strType = "TITLE" Then
    
148         If Trim(Me.rtxtLogHex.Text) <> "" Then strData = vbNewLine & strData
150         strData = strData & vbNewLine
        
152         Me.rtxtLogHex.Text = Me.rtxtLogHex.Text & strData
154         Me.rtxtLogHex.SelStart = Len(Me.rtxtLogHex.Text)
        
156         Me.rtxtLogTxt.Text = Me.rtxtLogTxt.Text & strData
158         Me.rtxtLogTxt.SelStart = Len(Me.rtxtLogTxt.Text)
        End If
    
160     If mstrStat = "����" And strType <> "TITLE" Then
            '������յ���ԭʼ����
162         strFilename = gstrRAWDIR & "\" & Format(Now, "yyyyMMdd") & ".txt"
164         Set txtStream = gFileObject.OpenTextFile(strFilename, ForAppending, True, TristateFalse)
166         txtStream.Write strData
168         txtStream.Close
170         Set txtStream = Nothing
        
172         Set fileTmp = gFileObject.GetFile(strFilename)
174         If fileTmp.Size > 3072 Then
176             fileTmp.Delete True
            End If
178         Set fileTmp = Nothing
180     ElseIf mstrStat = "����" And strType <> "TITLE" Then
182         strFilename = gstrSendDir & "\" & Format(Now, "yyyyMMdd") & ".txt"
        
184         Set txtStream = gFileObject.OpenTextFile(strFilename, ForAppending, True, TristateFalse)
186         txtStream.Write strData
188         txtStream.Close
190         Set txtStream = Nothing
192         Set fileTmp = gFileObject.GetFile(strFilename)
194         If fileTmp.Size > 3072 Then
196             fileTmp.Delete True
            End If
198         Set fileTmp = Nothing
        End If
    
200     If UBound(Split(Me.rtxtLogHex.Text, vbNewLine)) > 1000 Then
202         If mnuDebug.Checked Then Call SaveLog
204         Me.rtxtLogHex.Text = ""
206         Me.rtxtLogTxt.Text = ""
        End If
    
        Exit Sub
errH:
208     WriteErrLog "AddLog " & strType, strData, CStr(Erl()) & "��" & Err.Description
End Sub

Private Sub ShowHex(ByVal strInHex As String)
        Dim strHex As String, strChar As String, strLine As String
        Dim varTmp As Variant, strEndLine As String, strTmp As String, strFirst As String, lngLEN As Long
        Dim lngĩ�г��� As Long, i As Integer, blnAddCR As Boolean
    
        On Error GoTo hErr
100     strHex = strInHex
102     blnAddCR = False
104     If InStr(Me.rtxtLogHex.Text, vbNewLine) > 0 Then
106         varTmp = Split(Me.rtxtLogHex.Text, vbNewLine)
108         strEndLine = varTmp(UBound(varTmp))
        Else
110         strEndLine = Me.rtxtLogHex.Text
        End If
        
112     lngĩ�г��� = Len(strEndLine)
    
114     If strEndLine <> "" Then
116         For i = 0 To Len(strEndLine) / 3 - 1
118             If (i * 3 + 1) > (16 * 3) Then Exit For
120             strTmp = Mid(strEndLine, i * 3 + 1, 1)
122             If strTmp <> " " Then
                    '����ʮ�����Ƶ���
124                 strEndLine = ""
126                 lngĩ�г��� = 0
128                 blnAddCR = True
                    Exit For
                End If
            Next
        End If
    
130     If Len(strEndLine) >= 16 * 3 Then
132         strEndLine = " " & Trim(Mid(strEndLine, 1, 16 * 3))
        End If
    
134     If Len(strEndLine) >= 16 * 3 Then
136         If Len(strHex) >= 16 * 3 Then
138             strTmp = strHex
140             strHex = FormatHexLine(strTmp)
142             If Mid(strHex, 1, 2) = vbNewLine Then strHex = Mid(strHex, 3)
            End If
        Else
144         lngLEN = 16 * 3 - Len(strEndLine)
146         strFirst = strEndLine & Mid(strHex, 1, lngLEN)
            '����
148         strLine = strFirst
150         strFirst = FormatHexLine(strLine)
152         If Mid(strFirst, 1, 2) = vbNewLine Then strFirst = Mid(strFirst, 3)
            'ʣ�ಿ��
154         strLine = Mid(strHex, lngLEN + 1)
156         strHex = strFirst & FormatHexLine(strLine)
        End If
158     If lngĩ�г��� > 0 Then
160         Me.rtxtLogHex.Text = Mid(Me.rtxtLogHex.Text, 1, Len(Me.rtxtLogHex.Text) - lngĩ�г���) & IIf(blnAddCR, vbNewLine, "") & strHex
        Else
162         Me.rtxtLogHex.Text = Me.rtxtLogHex.Text & strHex
        End If
164     Me.rtxtLogHex.SelStart = Len(Me.rtxtLogHex.Text)
        Exit Sub
hErr:
166     WriteErrLog "ShowHex " & strInHex, Me.rtxtLogHex.Text, CStr(Erl()) & "��" & Err.Description
End Sub

Private Function FormatHexLine(ByVal strHexCode As String) As String
        Dim strHex As String
        Dim strTmp As String
        Dim strLine As String
        Dim strChar As String
        Dim i As Integer, byteChar As Byte
        On Error GoTo hErr
100     strHex = strHexCode
        'If Len(strHex) >= 16 * 3 Then
102         strTmp = strHex
104         strHex = ""
106         Do While Len(strTmp) >= 16 * 3
108             strLine = Mid(strTmp, 1, 16 * 3)
110             strChar = ""
112             For i = 0 To Len(strLine) / 3 - 1
114                 byteChar = CByte("&H" & Trim(Mid(strLine, 3 * i + 1, 3)))
116                 If byteChar >= 33 And byteChar <= 125 Then
118                     strChar = strChar & Chr(byteChar)
                    Else
120                     strChar = strChar & "."
                    End If
                Next
122             strHex = strHex & vbNewLine & strLine & "   " & strChar
124             strTmp = Mid(strTmp, 16 * 3 + 1)
            Loop
            'ĩβ
126         If strTmp <> "" Then
128             strLine = Mid(strTmp, 1, 16 * 3)
130             strChar = ""
132             For i = 0 To Len(strLine) / 3 - 1
134                 byteChar = CByte("&H" & Trim(Mid(strLine, 3 * i + 1, 3)))
136                 If byteChar >= 33 And byteChar <= 125 Then
138                     strChar = strChar & Chr(byteChar)
                    Else
140                     strChar = strChar & "."
                    End If
                Next
142             If i < 16 Then
144                 strLine = strLine & Space((16 - i) * 3)
                End If
146             strHex = strHex & vbNewLine & strLine & "   " & strChar
            End If
        'End If
148     FormatHexLine = strHex
        Exit Function
hErr:
150     WriteErrLog "FormatHexLine ", strHexCode, CStr(Erl()) & "��" & Err.Description
End Function
Private Sub ChangeICon()
    If picA.Tag = "" Then
        ModifyIcon Me.picIcon.hwnd, Me.picB.Image, mstrTip
        picA.Tag = "B"
    Else
        ModifyIcon Me.picIcon.hwnd, Me.picC.Image, mstrTip
        picA.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    mstrStat = ""
    mnuDebug.Checked = False
    Call OpenPort
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub Form_Resize()
    Dim lngTop As Long
    On Error Resume Next
    
    Select Case WindowState
    
        Case vbMinimized
            mnuTrayMinimize.Enabled = False
            mnuTrayRestore.Enabled = True
            Me.Hide
            AddIcon picIcon.hwnd, Me.Icon, mstrTip
        Case Else
            mnuTrayMinimize.Enabled = True
            mnuTrayRestore.Enabled = False
            RemoveIcon picIcon.hwnd
    End Select

    lngTop = 10
    If mnuDebug.Checked Then
        Me.rtxLogDebug.Top = 10
        Me.rtxLogDebug.Left = Me.ScaleLeft + 10
        Me.rtxLogDebug.Width = Me.ScaleWidth - 10
        'me.rtxLogDebug.Height
        lngTop = Me.rtxLogDebug.Top + Me.rtxLogDebug.Height + 10
    End If
    With Me.rtxtLogHex
        .Left = 10
        .Top = lngTop
        '.Width = Me.ScaleWidth - 80
        .Height = Me.ScaleHeight - .Top - 80

    End With
    With Me.fraWE
        .Top = lngTop
        .Left = Me.ScaleLeft + Me.rtxtLogHex.Width + 10
        .Height = Me.ScaleHeight - .Top - 80
        
    End With
    With Me.rtxtLogTxt
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Top = lngTop
        .Width = Me.ScaleWidth - Me.rtxtLogHex.Width - Me.fraWE.Width - 80
        .Height = Me.ScaleHeight - .Top - 80
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ClosePort
    '���ͨѶ�����ļ�
    If gstrLockFile <> "" Then
        If gFileObject.FileExists(gstrLockFile) Then gFileObject.DeleteFile gstrLockFile
    End If
    If Not gobjLisDev Is Nothing Then Set gobjLisDev = Nothing
    
    RemoveIcon picIcon.hwnd
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.rtxtLogHex.Width = Me.rtxtLogHex.Width + x
         
        Me.fraWE.Left = Me.fraWE.Left + x
        Me.rtxtLogTxt.Left = Me.fraWE.Left + Me.fraWE.Width
        Me.rtxtLogTxt.Width = Me.rtxtLogTxt.Width - x
    End If
End Sub

Private Sub mnuDebug_Click()
    mnuDebug.Checked = Not mnuDebug.Checked
    Call Form_Resize
End Sub

Private Sub mnuExit_Click()
    Call mnuTrayClose_Click
End Sub


Private Function Decode(ByVal StrInput As String) As String
        '����ͨѶ���򣬽���ԭʼ���ݣ�����������浽ResultĿ¼
        '
        Dim strResult As String, strReserved As String, strCmd As String
        Dim lngDataID As Long
        Dim blnGetSample As Boolean
        Dim aSampleInfo() As String, aSamples() As String, i As Long
        Dim strResponse As String, blnSuccess As Boolean
        Dim strSampleNO As String, aTmp() As String, strBarcode As String
        Dim strSendData As String
        Dim blnClearData As Boolean, lngIndex As Long
        Dim lngFileNo As Long, strFilename As String
        Dim strTmp As String
        Dim varResult As Variant
        On Error GoTo ErrHandle
    
        '����ԭʼ�Ľ�������
    
100     strCmd = ""
102     strReserved = ""
104     strResult = ""
    
106     If gobjLisDev Is Nothing Then Exit Function
    
        On Error Resume Next
    
    
         '-----Beging ˫��ͨѶ
108     mstrResponse = StrInput   '˫��ͨѶ�õı������������������,��������������ϵ�˫��ӿ���,������˫��ӿڲ���
    
   
110     If mintSendStep > 0 Then  '˫��ͨѶ�ڼ�
112         blnSuccess = False
            
            Call addDebug("", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & "��ӿڷ���ת������" & vbNewLine & mLastSampleInfo)
114         strSendData = gobjLisDev.SendSample(mLastSampleInfo, mintSendStep, blnSuccess, mstrResponse, mblnUndo, mintType)
116         If strSendData <> "" Then Call SendCmd(strSendData)
                                
118         If Not blnSuccess Then
120             mintSendStep = 0 '�������ʧ�ܣ���ȡ������
            Else
122             mstrResponse = ""
            End If
        
124         If mintSendStep = 0 Then

126             strSendData = ""
                '�ϴζ�����������Ϣ�����ڿյ�ʱ�򣬲�ɾ���ļ�,
                '  ����������Ϊ����UF����������͵Ľӿ���˫��ͨѶʱ������֧�ֶಽӦ��
128             If mLastSampleInfo <> "" Then strSendData = ReadSendDirFile(gstrSendDir & "\SendSample", True)
130             mLastSampleInfo = ""
132             mblnUndo = False
134             mintType = 0
136             If strSendData <> "" Then
                    '����ɾ���ˡ�
138                 If UBound(Split(strSendData, ";")) >= 2 And UBound(Split(strSendData, "|")) >= 10 Then
140                     mblnUndo = Val(Split(strSendData, ";")(1)) = 1
142                     mintType = Val(Split(strSendData, ";")(2))
144                     mLastSampleInfo = Split(strSendData, ";")(0)
                    End If
                End If
146             strSendData = ""
            End If
148         Decode = ""
            Exit Function
        End If
        '----- End ˫��ͨѶ
    
150     Call gobjLisDev.Analyse(StrInput, strResult, strReserved, strCmd)
    
152     Decode = strReserved
    
154     If Err.Number <> 0 Then
156         Call addDebug("", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " Decode Analyse Err:��" & CStr(Erl()) & "��," & Err.Description & ",����ӿڣ�Analyse�������д���")
        End If
        Dim strBack As String
        On Error GoTo ErrHandle
    
        If strResult <> "" Then
            varResult = Split(strResult, "||")
            For i = 0 To UBound(varResult)
                '----- ���ͽ������Ӧ��ָ��
                If UBound(Split(varResult(i), "|")) < 5 Then
158                 If Len(strCmd) > 0 Then
160                     aTmp = Split(strCmd, "|")
162                     If UBound(aTmp) > 0 Then
164                         strBack = Mid(strCmd, 3)
166                         If Val(aTmp(0)) = 1 Then '���������ȡ�걾��Ϣ�����Ǽ�����
                            
168                             If Len(strBack) > 0 Then
                                    '������Ϣ
170                                 Call SendCmd(strBack)
                                End If
                            
172                             If Len(varResult(i)) > 0 Then '����˫��ͨѶ����
174                                 strTmp = Format(Now, "YYYY-MM-DD") & "|^^0"
176                                 If varResult(i) = strTmp Then
                                        'UFϵͳ����ϼ�ʱ��˫��ͨѶ����:
                                        '    ����˫��������ָ�����Ҫ�ֳ�2�����͡�
                                        '    �򷵻ص�ֵ�������룬�걾�ţ����ܴ�HIS�еõ�������Ϣ��
                                        '    ��������ֻ�ܵ��ýӿڵ�SendSample����,���ĵ�������ϢΪ��.
                                        '    ��ʱ�����������������ӿ������
178                                     blnSuccess = False
180                                     strTmp = ""
182                                     strSendData = gobjLisDev.SendSample(strTmp, mintSendStep, blnSuccess, mstrResponse, mblnUndo, mintType)
184                                     If strSendData <> "" Then Call SendCmd(strSendData)
                                                            
186                                     If Not blnSuccess Then
188                                         mintSendStep = 0
                                        Else
190                                         mstrResponse = ""
                                        End If
                                    Else
192                                     strFilename = Dir(gstrResultDIR & "\IQ" & Format(Now, "yyyyMMdd") & "_*.txt")
194                                     If strFilename <> "" Then
196                                         lngIndex = Val(Split(strFilename, "_")(1))
                                        End If
                                        Do
198                                         lngIndex = lngIndex + 1
200                                         strFilename = gstrResultDIR & "\IQ" & Format(Now, "yyyyMMdd") & "_" & lngIndex & ".txt"
202                                         If Not gFileObject.FileExists(strFilename) Then
204                                             lngFileNo = FreeFile
206                                             Open strFilename For Binary Access Read Write Lock Read Write As lngFileNo
208                                             Put lngFileNo, , CStr(varResult(i))
210                                             Close lngFileNo
                                                addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & "��LIS������ѯ����" & vbNewLine & varResult(i)
                                                Exit Do
                                            End If
                                        Loop
                                    End If
                                End If
'                                Exit Function
212                         ElseIf Val(aTmp(0)) = 0 Then
214                             If Len(strBack) > 0 Then
                                    '������Ϣ
216                                 Call SendCmd(strBack)
                                End If
                            End If
                        Else
218                         Call SendCmd(strCmd)
                        End If
                    End If
                Else
220                 If Len(varResult(i)) > 0 Then
                        '���ؼ������󣬷��͵�Ӧ��ָ��
                        If Len(strCmd) > 0 Then
                            aTmp = Split(strCmd, "|")
                            If UBound(aTmp) > 0 Then
                                strCmd = Mid(strCmd, 3)
                                If Val(aTmp(0)) = 0 Then
                                    If Len(strCmd) > 0 Then
                                        '������Ϣ
                                        Call SendCmd(strCmd)
                                    End If
                                End If
                            Else
                                Call SendCmd(strCmd)
                            End If
                        End If
                        
                        '����������
222                     addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ������" & vbNewLine & varResult(i)
224                     strFilename = Dir(gstrResultDIR & "\RE" & Format(Now, "yyyyMMdd") & "_*.txt")
226                     If strFilename <> "" Then
228                         lngIndex = Val(Split(strFilename, "_")(1))
                        End If
                        Do
230                         lngIndex = lngIndex + 1
232                         strFilename = gstrResultDIR & "\RE" & Format(Now, "yyyyMMdd") & "_" & lngIndex & ".txt"
234                         If Not gFileObject.FileExists(strFilename) Then
236                             lngFileNo = FreeFile
238                             Open strFilename For Binary Access Read Write Lock Read Write As lngFileNo
240                             Put lngFileNo, , Replace(varResult(i), Chr(&HD) & Chr(&HA), "CHR(10) CHR(13)")
242                             Close lngFileNo
                                Exit Do
                            End If
                        Loop
244                     addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " ���������浽�ļ�" & vbNewLine & strFilename
                    End If
                End If
            Next
        Else
            If Len(strCmd) > 0 Then
                aTmp = Split(strCmd, "|")
                If UBound(aTmp) > 0 Then
                    strBack = Mid(strCmd, 3)
                    If Val(aTmp(0)) = 1 Then '���������ȡ�걾��Ϣ�����Ǽ�����
                        If Len(strBack) > 0 Then
                            '������Ϣ
                            Call SendCmd(strBack)
                        End If
                    ElseIf Val(aTmp(0)) = 0 Then
                        If Len(strBack) > 0 Then
                            '������Ϣ
                            Call SendCmd(strBack)
                        End If
                    End If
                Else
                    Call SendCmd(strCmd)
                End If
            End If
        End If
        Exit Function
ErrHandle:
246     Call addLog("TITLE", "Decode Err :��" & CStr(Erl()) & "��," & Err.Description)
End Function

Private Sub SaveBuffer()
    '���� �������ݵ�����Ӳ��
    Dim lngFileNo As Long
    Dim strFilename As String
    Dim lngCount As Long
    Dim bytData() As Byte
    Dim strCode As String
    Dim blnOpen As Boolean
    Dim lngBits As Long, lngloop As Long
    
    On Error GoTo errH
    If mstrBuffer = "" Then Exit Sub
    strCode = mstrBuffer
    blnOpen = False
    
    lngCount = 1
    
    Do
        strFilename = gstrRAWDIR & "\Buff" & Format(lngCount, "000") & ".TXT"
        If gFileObject.FileExists(strFilename) = False Then
            If g��������.�ַ�ģʽ = 0 Then
                lngFileNo = FreeFile
                Open strFilename For Binary Access Read Write Lock Read Write As lngFileNo
                blnOpen = True
                Put lngFileNo, , strCode
                Close lngFileNo
                blnOpen = False
                mstrBuffer = ""
            Else
                lngBits = Len(strCode) / 3
                ReDim bytData(lngBits - 1)
                For lngloop = 1 To lngBits
                    bytData(lngloop - 1) = Val("&H" & Mid(Left(strCode, 3), 2))
                    strCode = Mid(strCode, 4)
                Next
                lngFileNo = FreeFile
                Open strFilename For Binary Access Read Write Lock Read Write As lngFileNo
                blnOpen = True
                Put lngFileNo, , bytData()
                Close lngFileNo
                blnOpen = False
            End If
            mstrBuffer = ""
            Exit Do
        End If
        lngCount = lngCount + 1
        If lngCount > 999 Then Exit Do
    Loop
    Exit Sub
errH:
    If blnOpen = True Then Close lngFileNo
End Sub

Private Sub SendFile(ByVal strFile As String, Optional ByVal blnDelete As Boolean = False)
    '����һ���ļ��е�����
    'strFile  : �ļ���
    'blnDelete: ���ͺ��Ƿ�ɾ�����ļ�
    
    Dim lngFileSize As Long, lngCurSize As Long, lngModSize As Long, intChunkSize As Integer, lngBlocks As Long
    Dim lngFileNo As Long, lngCount As Long
    Dim bytData() As Byte, strSendData As String
    Dim str��  As String
    
    Dim blnFileOpen As Boolean
    On Error GoTo errH
    
    If mstrStat = "" Or mstrStat = "����" Then
        str�� = IIf(mstrStat = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " ��)")
        Call addLog("TITLE", "Send " & Format(Now, "yyyy-MM-dd HH:mm:ss") & " " & str��)
        mLasterTime = Now
        mstrStat = "����"
    End If
    blnFileOpen = False
    
    If strFile = "" Then Exit Sub
    If gFileObject.FileExists(strFile) = False Then Exit Sub
    
    lngFileNo = FreeFile
    Open strFile For Binary Access Read Lock Write As lngFileNo
    blnFileOpen = True
    lngFileSize = LOF(lngFileNo)
    intChunkSize = 512
    If lngFileSize > 0 Then
        lngModSize = lngFileSize Mod intChunkSize
        lngBlocks = lngFileSize \ intChunkSize - IIf(lngModSize = 0, 1, 0)
        For lngCount = 0 To lngBlocks
            If lngCount = lngFileSize \ intChunkSize Then
                lngCurSize = lngModSize
            Else
                lngCurSize = intChunkSize
            End If
            ReDim bytData(lngCurSize - 1) As Byte
            
            Get lngFileNo, , bytData
            
            If g��������.���� = 0 Then
                COM.Output = bytData()
            Else
                WSK.SendData bytData()
            End If
        Next
    End If
    Close lngFileNo
    If blnDelete = True Then Call gFileObject.DeleteFile(strFile)
    Exit Sub
errH:
    If blnFileOpen Then Close lngFileNo
    
End Sub

Private Sub OpenPort()
    '��ͨѶ
    Dim strTmp As String
    
    On Error GoTo errH
    Call ReadSet
    Set gobjLisDev = Nothing
    Set gobjLisDev = CreateObject(g��������.ͨѶ����)
    
    mblnListen = False
    With g��������
        Me.Caption = "ͨѶ����(Ver " & App.Major & "." & App.Minor & "." & App.Revision & ")-" & .ͨѶ����
        If .���� = 0 Then
            If COM.PortOpen = True Then COM.PortOpen = False
            COM.CommPort = .COM�˿�
            COM.Settings = .������ & "," & .У��λ & "," & .����λ & "," & .ֹͣλ
            COM.InputMode = .�ַ�ģʽ
            COM.Handshaking = Val(.����)
            COM.RTSEnable = True
            COM.RThreshold = 1        'ÿ����һ���ַ�����on_comm�¼�
            COM.InBufferCount = 0     '������ջ���
            COM.InputLen = 0          'ʹ��inputʱ,��ȡ���ջ�������ȫ��������
            COM.InBufferSize = .�����С
            COM.PortOpen = True
            
            'Call addLog("TITLE", "Open COM" & COM.CommPort & " " & COM.Settings & " " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
            addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �򿪶˿ڣ�" & COM.CommPort & "���˿����ã�" & COM.Settings & ",ͨѶ����" & .ͨѶ����
            mstrTip = "COM" & COM.CommPort & " " & COM.Settings
        Else
            If Not WSK.State = sckOpen Then
                WSK.Close
                WSK.Tag = .�ַ�ģʽ    '�����ģʽ
                If .���� = 1 Then
                    WSK.Protocol = sckTCPProtocol
                    WSK.Bind .IP�˿�, .IP
                    WSK.Listen
                    mblnListen = True
                    'Call addLog("TITLE", WSK.LocalIP & ":" & WSK.LocalPort & " Listen " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
                    addDebug "", Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����������" & WSK.LocalIP & ":" & WSK.LocalPort & ",ͨѶ����" & .ͨѶ����
                    mstrTip = WSK.LocalIP & ":" & WSK.LocalPort & " Listen"
                 Else
                    WSK.Protocol = sckTCPProtocol  '����ͨѶЭ��
                    WSK.RemoteHost = .IP     'Զ��IP
                    WSK.RemotePort = .IP�˿�
                    WSK.Connect  '����
                    addDebug "", Format(Now, "yyyy-MM-dd HH:mm:ss") & " ����������" & .IP & ":" & .IP�˿� & ",ͨѶ����" & .ͨѶ����
                    
                    mstrTip = WSK.LocalIP & ":" & WSK.LocalPort & " ->" & WSK.RemoteHost & ":" & WSK.RemotePort
                End If
            End If
        End If
    End With
    
    '���豸���Ϳ�ʼ��������
     
    If Not gobjLisDev Is Nothing Then
        On Error Resume Next
        strTmp = gobjLisDev.GetStartCmd
        If Err.Number <> 0 Then Call addLog("TITLE", "OpenPort GetStartCmd :" & Err.Description)
        If Len(strTmp) > 0 And Err.Number = 0 Then Call SendCmd(strTmp)
    End If
    
    '˫��ͨѶ��ʱ������趨
    TimSendCmd.Interval = g��������.ͨѶ���� * 1000
    TimSendCmd.Enabled = True
    
    '��ʱ����ָ��
    TimAutoAnswer.Enabled = False: TimAutoAnswer.Interval = 0
    mblnAutoAnswer = False
    If Val(g��������.�Զ�Ӧ��) > 0.1 And Val(g��������.�Զ�Ӧ��) < 600 Then
        TimAutoAnswer.Interval = Val(g��������.�Զ�Ӧ��) * 1000
        TimAutoAnswer.Enabled = True
        mblnAutoAnswer = True
    End If
    
    Exit Sub
errH:
    If InStr(Err.Description, "ActiveX �������ܴ�������") > 0 Then
        strTmp = "ͨѶ�������Ʋ���ȷ��ӿڲ������޴�ͨѶ����"
    Else
        strTmp = ""
    End If
    Call addLog("TITLE", "OpenPort Err :" & Err.Description & IIf(strTmp = "", "", vbNewLine & strTmp))
End Sub

Private Sub ClosePort()
    '�ر�ͨѶ
    Dim strCmd As String
    If Not gobjLisDev Is Nothing Then
        On Error Resume Next
        strCmd = gobjLisDev.GetEndCmd
        If Err.Number <> 0 Then Call addLog("TITLE", "ClosePort GetEndCmd Err:" & Err.Description)
        If strCmd <> "" And Err.Number = 0 Then Call SendCmd(strCmd)
        If Err.Number <> 0 Then Err.Clear
    End If
    On Error GoTo errH
    
    If g��������.���� = 0 Then
      If COM.PortOpen = True Then COM.PortOpen = False
      'Call addLog("TITLE", "Close COM" & COM.CommPort & " " & COM.Settings & " " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
      addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �˿ڹرգ�" & COM.CommPort & "���˿����ã�" & COM.Settings
    Else
        mblnListen = False
        WSK.Close
        'Call addLog("TITLE", "Close Connect" & WSK.RemoteHost & " " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
        addDebug "", vbNewLine & Format(Now, "yyyy-MM-dd HH:mm:ss") & " �ر����ӣ�" & WSK.RemoteHost
    End If
    Exit Sub
errH:
    Call addLog("TITLE", "ClosePort Err:" & Err.Description)
End Sub

Private Sub SendCmd(ByVal strSendCmd As String, Optional intErr As Integer = 0)
    '������Ϣ
    'interr= 0ʱ�ŷ��ͣ�Ϊ1ʱ�����͵�����
    Dim bitByte() As Byte
    Dim lngBits As Long, lngloop As Long
    Dim strCode As String
    Dim ReturnBin As Boolean
    Dim blnErr As Boolean, str�� As String
    Dim lngCount As Long    '����10��
    On Error GoTo errH
    
    lngCount = 0
    If mstrStat = "" Or mstrStat = "����" Then
        str�� = IIf(mstrStat = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " ��)")
        Call addLog("TITLE", "==> Send " & Format(Now, "yyyy-MM-dd HH:mm:ss") & " " & str��)
        mLasterTime = Now
        mstrStat = "����"
    End If
    If strSendCmd = "" Then Exit Sub
    '���ݽ���ģʽȷ������ģʽ
    '0-�ַ�ģʽ 1-�ַ�����
    ReturnBin = g��������.�ַ�ģʽ = 1
    
    If ReturnBin Then
        '���������ݣ�תΪ�ַ�����
        strCode = strSendCmd
        lngBits = Len(strCode) / 3
        If lngBits > 0 Then
            ReDim bitByte(lngBits - 1)
            For lngloop = 1 To lngBits
                bitByte(lngloop - 1) = Val("&H" & Mid(Left(strCode, 3), 2))
                strCode = Mid(strCode, 4)
            Next
        Else
            blnErr = True
            Call addLog("TITLE", "SendCmd Err: ���Ƕ����Ƹ�ʽ������ ")
        End If
    End If
    
    If g��������.���� = 0 Then
        If intErr = 0 Then
            If ReturnBin Then
                If blnErr = False Then
                    COM.Output = bitByte
                    addLog "HEX", strSendCmd
                End If
            Else
                COM.Output = strSendCmd
                addLog "TXT", strSendCmd
            End If
        End If
    Else
        If intErr = 0 Then
            If ReturnBin Then
                If blnErr = False Then
                    Do While WSK.State <> sckConnected And lngCount < 10000
                        lngCount = lngCount + 1
                        DoEvents
                    Loop
                    Call WSK.SendData(bitByte)    '���ַ�����
                    addLog "HEX", strSendCmd
                End If
            Else
                Do While WSK.State <> sckConnected And lngCount < 10000
                    lngCount = lngCount + 1
                    DoEvents
                Loop
                Call WSK.SendData(strSendCmd) '���ı�
                addLog "TXT", strSendCmd
            End If
        End If
    End If
    
    Exit Sub
errH:
    Call addLog("TITLE", "SendCmd  " & strSendCmd & " Err:" & Err.Description)
End Sub

Private Sub mnuFileSave_Click()
    Call SaveLog
    MsgBox "����ɹ�!"
End Sub

Private Sub SaveLog()
    
    Dim strFilename As String
    
    On Error Resume Next
    
    strFilename = App.Path & "\ͨѶ��־_HEX" & Format(Now, "yyyyMMddHHMM") & ".log"
    If gFileObject.FileExists(strFilename) Then gFileObject.DeleteFile strFilename
    Call rtxtLogHex.SaveFile(strFilename, rtfText)
    
    strFilename = App.Path & "\ͨѶ��־_TXT" & Format(Now, "yyyyMMddHHMM") & ".log"
    If gFileObject.FileExists(strFilename) Then gFileObject.DeleteFile strFilename
    Call rtxtLogTxt.SaveFile(strFilename, rtfText)


End Sub
Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '--------------------------------------------------------------------------------------------------
    '����:  ����ͼ��ĸ��ִ����¼�
    '--------------------------------------------------------------------------------------------------
    On Error Resume Next
    Select Case Button '
        Case vbLeftButton
            Me.Show
            Me.WindowState = vbNormal
        Case vbRightButton
            ModifyIcon picIcon.hwnd, Me.Icon, , False
            Me.PopupMenu Me.mnuTray
            ModifyIcon picIcon.hwnd, Me.Icon
    End Select '
End Sub

Private Sub TimAutoAnswer_Timer()
    Dim strCmd As String
    On Error GoTo errH
    If g��������.���� = 0 Then
        If COM.InBufferCount > 0 Then Exit Sub
    Else
        If WSK.BytesReceived > 0 Then Exit Sub
    End If
    If gobjLisDev Is Nothing Then Exit Sub
    
    strCmd = gobjLisDev.GetAnswerCmd
    If strCmd <> "" Then Call SendCmd(strCmd)
    Exit Sub
errH:
    TimAutoAnswer.Enabled = False
    mblnAutoAnswer = False
    Call addLog("TITLE", "TimAutoAnswer Err:" & Err.Description)
End Sub

Private Sub TimSendCmd_Timer()

    Dim strFile  As String
    Dim strSendCmd As String
    Dim strSampleInfo As String
    Dim blnUndo As Boolean, iType As Integer, blnSuccess As Boolean
    Dim strTmp As String
    Dim strSaveDataLog  As String
    On Error GoTo errH
    
    ' ����д����յ����ݣ����˳�
    If g��������.���� = 0 Then
        If COM.InBufferCount > 0 Then Exit Sub
    Else
        If WSK.BytesReceived > 0 Then Exit Sub
    End If
    
    'SetTrayIcon Me.picA
    
    If mintSendStep > 0 Then Exit Sub      '���ڷ��ͣ��˳�
    
    '��ȡӦ����Ϣ������
    If gobjLisDev Is Nothing Then Exit Sub
    
    strFile = Dir(gstrSendDir & "\SendSample*.txt")

    If strFile <> "" Then
        If mLastSampleInfo = "" Then
            strSampleInfo = ReadSendDirFile(gstrSendDir & "\SendSample", False)  '��ɾ���ļ�,��Ҫ���������ӿڷ��صĲ�����ȷ���Ƿ�ɾ��
        Else
            strSampleInfo = mLastSampleInfo & ";" & IIf(mblnUndo, "1", "0") & ";" & mintType
        End If

        If UBound(Split(strSampleInfo, ";")) >= 2 And UBound(Split(strSampleInfo, "|")) >= 10 Then
            blnUndo = Val(Split(strSampleInfo, ";")(1)) = 1
            iType = Val(Split(strSampleInfo, ";")(2))
            strSampleInfo = Split(strSampleInfo, ";")(0)
            blnSuccess = False
                  
            strSendCmd = gobjLisDev.SendSample(strSampleInfo, mintSendStep, blnSuccess, mstrResponse, blnUndo, iType)
            
            If strSendCmd <> "" Then Call SendCmd(strSendCmd)
            If blnSuccess = True Then
                mstrResponse = ""
            Else
                mintSendStep = 0
            End If
            
            If mintSendStep = 0 Then
                strSampleInfo = ReadSendDirFile(gstrSendDir & "\SendSample", True)
                mLastSampleInfo = ""
                mblnUndo = False
                mintType = 0
            Else
                mLastSampleInfo = strSampleInfo
                mblnUndo = blnUndo
                mintType = iType
            End If

        ElseIf strSampleInfo <> "" Then
        
            Call addLog("TITLE", "TimSendCmd Err: Ӧ����Ϣ��ʽ���� " & strSampleInfo)
            '��ǰ�걾�ĸ�ʽ�����ڴ˴�ɾ���ļ����ú���ı걾��ִ�С�
            strTmp = ReadSendDirFile(gstrSendDir & "\SendSample", False)
            If strTmp = "" Then
                strTmp = ReadSendDirFile(gstrSendDir & "\SendSample", True)
            End If

            strSampleInfo = ""
        Else
            '���ļ���ֱ��ɾ��
            Call ReadSendDirFile(gstrSendDir & "\SendSample", True)
        End If
        
        strFile = Dir(gstrSendDir & "\SendSample*.txt")

    End If
    
    If g��������.���� = 0 Then
        If COM.InBufferCount > 0 Then Exit Sub
    Else
        If WSK.BytesReceived > 0 Then Exit Sub
    End If
    
    '�����ӿ�
    If Dir(gstrSendDir & "\ResetExe.txt") <> "" Then
        Call gFileObject.DeleteFile(gstrSendDir & "\ResetExe.txt")
        Call OpenPort
    End If
    If Dir(gstrSendDir & "\CloseExe.txt") <> "" Then
        Call gFileObject.CopyFile(gstrSendDir & "\CloseExe.txt", gstrSendDir & "\CloseEnd.txt")
        Call gFileObject.DeleteFile(gstrSendDir & "\CloseExe.txt")
        If gFileObject.FileExists(gstrLockFile) Then gFileObject.DeleteFile gstrLockFile
        Call Shell(App.Path & "\" & App.EXEName)
        End
    End If
    If Dir(gstrSendDir & "\SaveDataLog*.txt") <> "" Then
        strSaveDataLog = ReadSendDirFile(gstrSendDir & "\SaveDataLog", True)
        If strSaveDataLog <> "" Then addDebug "", vbNewLine & strSaveDataLog
    End If
    
    Exit Sub

errH:
    Call addLog("TITLE", "TimSendCmd Err:��" & CStr(Erl()) & "��" & Err.Description)
End Sub

Private Function ReadSendDirFile(ByVal strFileType As String, ByVal blnDelete As Boolean) As String

        '�������ļ�
        Dim strFilename As String
        Dim objStream As TextStream
 
        Dim strLine As String, lngCount As Long
    
        On Error GoTo errH
    
100     strFilename = Dir(strFileType & "_*.txt")
102     If strFilename <> "" Then
104         Do While lngCount < 1000
106             lngCount = lngCount + 1
108             strFilename = strFileType & "_" & Format(lngCount, "000") & ".txt"
110             If gFileObject.FileExists(strFilename) Then Exit Do
            Loop
        
112         If gFileObject.FileExists(strFilename) Then
        
114             Set objStream = gFileObject.OpenTextFile(strFilename, ForReading)
116             strLine = ""
                Do
118                 If objStream.AtEndOfStream Then Exit Do
120                 If strFileType = gstrSendDir & "\SaveDataLog" Then
122                     strLine = strLine & IIf(strLine = "", "", vbNewLine) & objStream.ReadLine
                    Else
124                     strLine = strLine & objStream.ReadLine
                    End If
                Loop
126             objStream.Close
128             Set objStream = Nothing
            
130             If gFileObject.FileExists(strFilename) And blnDelete = True Then
132                 Call gFileObject.DeleteFile(strFilename)
                Else
                    '��ȡ����֮����մ��ļ�,�����ظ�����
134                 Set objStream = gFileObject.CreateTextFile(strFilename, True)
136                 objStream.Close
138                 Set objStream = Nothing
                
                End If
140             ReadSendDirFile = strLine
            End If
        End If
        Exit Function
errH:
142     Call addLog("TITLE", "ReadSendDirFile Err:��" & CStr(Erl()) & "��," & Err.Description)
End Function

Private Sub WSK_Connect()
    Call addLog("TITLE", WSK.LocalIP & "<-->" & WSK.RemoteHostIP & " Connected " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
End Sub

Private Sub WSK_ConnectionRequest(ByVal requestID As Long)
    If WSK.State <> sckClosed Then WSK.Close
    WSK.Accept requestID
    
    Call addLog("TITLE", "Accept " & WSK.RemoteHostIP & ":" & WSK.RemotePort & " " & Format(Now, "yyyy-MM-dd HH:mm:ss"))
End Sub

Private Sub WSK_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If g��������.���� = 1 And mblnListen = True Then
        WSK.Close
        WSK.Listen
    End If
End Sub

Private Sub WSK_Close()
    If g��������.���� = 1 And mblnListen = True Then
        WSK.Close
        WSK.Listen
    End If
End Sub

Private Sub WSK_DataArrival(ByVal bytesTotal As Long)
    Dim byt_Bit() As Byte '-���ն���������
    Dim strData As String, str�� As String
    Dim i As Integer
    
    '��ͣ ��ʱ��

    TimSendCmd.Enabled = False
    TimAutoAnswer.Enabled = False
    
    If mstrStat = "" Or mstrStat = "����" Then
        str�� = IIf(mstrStat = "", "", "(+ " & DateDiff("s", mLasterTime, Now) & " ��)")
        Call addLog("TITLE", "<== Receive " & Format(Now, "yyyy-MM-dd HH:mm:ss") & " " & str��)
        mLasterTime = Now
        mstrStat = "����"
    End If
    
    If Val(g��������.�ַ�ģʽ) = 0 Then
        strData = ""
        WSK.GetData strData
        mstrBuffer = mstrBuffer & strData
        If Len(strData) > 0 Then Call addLog("TXT", strData)
    Else
        WSK.GetData byt_Bit, vbArray + vbByte
        strData = ""
        For i = 0 To UBound(byt_Bit)
            strData = strData & "," & IIf(Len(Hex(byt_Bit(i))) = 1, "0" & Hex(byt_Bit(i)), Hex(byt_Bit(i)))
        Next
        mstrBuffer = mstrBuffer & strData
        If Len(strData) > 0 Then Call addLog("HEX", strData)
    End If
    '�������յ�������
    mstrBuffer = Decode(mstrBuffer)
   ' If App.PrevInstance <> "VB6" Then Call ChangeICon
    
    TimSendCmd.Enabled = True
    If mblnAutoAnswer = True Then TimAutoAnswer.Enabled = True
    
End Sub

Private Sub mnuTrayClose_Click()
    If MsgBox("�رպ󽫲��ܽ����������ݣ���ȷ���Ƿ��˳���", vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
        Unload Me
        End
    End If
End Sub

'���̲˵�����
'Private Sub mnuTrayMaximize_Click()
'    WindowState = vbMaximized
'End Sub

Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
    Me.Hide
End Sub

'Private Sub mnuTrayMove_Click()
'    SendMessage hwnd, WM_SYSCOMMAND, _
'        SC_MOVE, 0&
'End Sub

Private Sub mnuTrayRestore_Click()
    WindowState = vbNormal
    Me.Show
End Sub

'Private Sub mnuTraySize_Click()
'    SendMessage hwnd, WM_SYSCOMMAND, _
'        SC_SIZE, 0&
'End Sub



