VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4065
   ClientLeft      =   285
   ClientTop       =   1155
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "frmDisplay.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   6000
   StartUpPosition =   2  '��Ļ����
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      _cx             =   1085
      _cy             =   661
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "true"
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5520
      Top             =   3480
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� �� �� ͣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Tag             =   "�� �� �� �� ͣ"
      Top             =   1800
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label lblPrefix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblPrefix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ԥ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblPrefix 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblCash 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "98.50Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblThanks 
      BackStyle       =   0  'Transparent
      Caption         =   "�����뵱�����,лл!"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Label lblDrugWindow 
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ʋ���������ҩ�����Ŵ���ȡҩ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Tag             =   "�����Ʋ���&Windowsȡҩ"
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Label lblCash 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1.50Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblCash 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100.00Ԫ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblWaiting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� �� !"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label lblPatient 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   3480
   End
   Begin VB.Label lblFree 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� �� �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Tag             =   "�� �� �� �� ��"
      Top             =   1800
      Width           =   5775
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnTest As Boolean
Public mblnLoad As Boolean

Private Type cp
  wp As Single
  hp As Single
  tp As Single
  lp As Single
End Type
Private ap() As cp
Private Const MFONTSIZE = 14.25
Private mlngFHeight As Long
Private mstrPicfile As String
Private mblnRightClick As Boolean



Private Sub GetPlaceData()
  Dim i As Integer
  For i = 0 To Controls.Count - 1
    If TypeName(Controls(i)) = "Label" Or TypeName(Controls(i)) = "ShockwaveFlash" Then
    With ap(i)
      .wp = Controls(i).Width / Me.ScaleWidth
      .hp = Controls(i).Height / Me.ScaleHeight
      .lp = Controls(i).Left / Me.ScaleWidth
      .tp = Controls(i).Top / Me.ScaleHeight
    End With
    End If
  Next
End Sub

Private Sub Form_DblClick()
    If mblnRightClick Then
        mblnRightClick = False
         '��������ͣ
        Timer1.Enabled = False
        lblPause.Visible = Not lblPause.Visible
        lblFree.Visible = Not lblPause.Visible
        
        Call ShowMain(False)
        Call ShowFee(False)
        Call ShowFlash
    Else
        If Me.WindowState <> 2 Then
            Me.WindowState = 2
        Else
            Me.WindowState = 0
            If mblnTest Then Unload Me
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�������岻��Ӧ�����¼�
'    If KeyCode = vbKeyF4 Then
'        If Timer1.Enabled Then Timer1.Enabled = False
'        'Ctrl+F4�����ڿ���
'        lblFree.Visible = Shift = 2
'        lblPause.Visible = Not lblFree.Visible
'
'        Call ShowMain(False)
'        Call ShowFee(False)
'        Call ShowFlash
'    End If
End Sub

Private Sub Form_Load()
    Dim arrInfo As Variant
    Dim strPicfile As String, strSwffile As String
    
    mblnLoad = True
    mblnRightClick = False
    
    '����ԭʼ��С,����Ҫ���ڻָ�ע����¼��λ��֮ǰ,��ΪҪ��resize
    ReDim ap(0 To Controls.Count - 1)
    Call GetPlaceData
    mlngFHeight = Me.Height
    
    arrInfo = Split(GetSetting("ZLSOFT", "����ȫ��", "˫����ʾ��λ��", ""), ",")
    If UBound(arrInfo) = 3 Then
        Me.Top = arrInfo(0)
        Me.Left = arrInfo(1)
        Me.Width = arrInfo(2)
        Me.Height = arrInfo(3)
    End If
        
    mstrPicfile = GetSetting("ZLSOFT", "����ȫ��", "����ͼƬ", "")
    If mstrPicfile <> "" Then
        Me.Picture = LoadPicture(mstrPicfile)
    End If
    
    swf.Movie = GetSetting("ZLSOFT", "����ȫ��", "SWF�ļ�", "")
    Call ShowFlash  '��ʼ��ʾ
    
    If Val(GetSetting("ZLSOFT", "����ȫ��", "�е�����Ϣ", "0")) = 1 Then
        lblThanks.Caption = GetSetting("ZLSOFT", "����ȫ��", "������Ϣ", "")
        lblThanks.Tag = lblThanks.Caption
    Else
        lblThanks.Caption = ""
        lblThanks.Tag = ""
    End If
    
    If mblnTest Then
        lblWaiting.Visible = False
        lblFree.Visible = False
        lblPause.Visible = False
        MsgBox "˫�������,�ٴ�˫���˳�!", vbInformation, Me.Caption
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(frmDisplay.hWnd)
    ElseIf Button = 2 Then
       mblnRightClick = True
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Integer, lngFontSize As Long, lngW As Long, lngH As Long
    
    On Error Resume Next
    
    If Me.Picture <> 0 Then
        Me.PaintPicture Me.Picture, Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    End If
        
    lngFontSize = MFONTSIZE * Me.ScaleHeight / mlngFHeight
    lngW = Me.ScaleWidth
    lngH = Me.ScaleHeight
    For i = 0 To Controls.Count - 1
        If TypeName(Controls(i)) = "Label" Then
            Controls(i).Move ap(i).lp * lngW, ap(i).tp * lngH, ap(i).wp * lngW, ap(i).hp * lngH
            
            Select Case Controls(i).Name
                Case "lblPatient", "lblWaiting", "lblFree", "lblPause"
                    Controls(i).FontSize = lngFontSize * 1.8
                Case "lblPrefix", "lblCash"
                    Controls(i).FontSize = lngFontSize * 1.3
                Case Else
                     Controls(i).FontSize = lngFontSize
            End Select
     End If
    Next
    swf.Left = Me.ScaleLeft
    swf.Top = Me.ScaleTop
    swf.Width = Me.ScaleWidth
    swf.Height = Me.ScaleHeight - lblPause.Height
    If swf.Visible Then
        lblPause.Top = swf.Top + swf.Height
        lblFree.Top = lblPause.Top
    End If
    
    SaveSetting "ZLSOFT", "����ȫ��", "˫����ʾ��λ��", Me.Top & "," & Me.Left & "," & Me.Width & "," & Me.Height
End Sub


Private Sub Timer1_Timer()
'��ʾҩ����,1���Ӻ󴥷� , Interval = 60000
    Timer1.Enabled = False
    
    Call ShowMain(False)
    Call ShowFee(False)
    lblFree.Visible = True
    lblPause.Visible = False
    Call ShowFlash
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'ֻ��������LEDʱ���Բ����Ż����,ֱ��ʹ��ʱ,�����˳���ʽ�رմ��岻�����
    mblnTest = False
    mblnLoad = False
End Sub


Private Sub lblPause_DblClick()
    Call Form_DblClick
End Sub

Private Sub lblPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblFree_DblClick()
    Call Form_DblClick
End Sub

Private Sub lblFree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub



Private Sub lblCash_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveObj(frmDisplay.hWnd)
End Sub
Private Sub lblDrugWindow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Call MoveObj(frmDisplay.hWnd)
End Sub

Private Sub lblDrugWindow_DblClick()
    Call Form_DblClick
End Sub

Private Sub lblPatient_DblClick()
    Call Form_DblClick
End Sub

Private Sub lblCash_DblClick(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub lblPrefix_DblClick(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub lblThanks_DblClick()
    Call Form_DblClick
End Sub

Private Sub lblWaiting_DblClick()
    Call Form_DblClick
End Sub

Private Sub ShowFlash()
'���к���ͣһ���ǻ����
    If swf.Movie <> "" Then
        If lblFree.Visible Then '����
            If lblFree.Visible Then lblFree.Top = Me.Height - lblFree.Height
            swf.Visible = lblFree.Visible
            Call PlayFlash(swf, swf.Visible, Me.Height - lblFree.Height)
            
        Else    '��ͣ���ٿ���
            If lblPause.Visible Then lblPause.Top = Me.Height - lblPause.Height
            swf.Visible = lblPause.Visible
            Call PlayFlash(swf, swf.Visible, Me.Height - lblPause.Height)
        End If
    End If
End Sub


'----�ⲿ����-------------------------------

Public Sub Check_Update_BkPic()
    Dim strPicfile As String
    
    strPicfile = GetSetting("ZLSOFT", "����ȫ��", "����ͼƬ", "")
    If mstrPicfile <> strPicfile Then
        mstrPicfile = strPicfile
        If mstrPicfile <> "" Then
            Me.Picture = LoadPicture(mstrPicfile)
        Else
            Me.Picture = Nothing
        End If
        
        Call Form_Resize
    End If
End Sub

Public Sub ShowMain(bln As Boolean)
    
    If bln And Timer1.Enabled Then Timer1.Enabled = False
    If bln And lblPause.Visible Then
        lblPause.Visible = False
        Call ShowFlash
    End If
    If bln And lblFree.Visible Then
        lblFree.Visible = False
        Call ShowFlash
    End If
    
    lblPatient.Visible = bln
    lblPatient.Caption = ""
    lblWaiting.Visible = bln
End Sub

Public Sub ShowFee(bln As Boolean)
    Dim i As Long
    
    If bln And Timer1.Enabled Then Timer1.Enabled = False
    If bln And lblPause.Visible Then
        lblPause.Visible = False
        Call ShowFlash
    End If
    If bln And lblFree.Visible Then
        lblFree.Visible = False
        Call ShowFlash
    End If
    If bln And lblWaiting.Visible Then lblWaiting.Visible = False
        
    For i = 0 To 2
        lblPrefix(i).Visible = bln
        lblCash(i).Visible = bln
        lblCash(i).Caption = ""
    Next
    
    lblDrugWindow.Visible = bln
    lblDrugWindow.Caption = ""
    
    lblThanks.Visible = bln
    lblThanks.Caption = ""
End Sub
