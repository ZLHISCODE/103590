VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmUnSendDrug 
   BorderStyle     =   0  'None
   Caption         =   "δ��ҩƷ"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmUnSendDrug.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timReset 
      Left            =   3720
      Top             =   1560
   End
   Begin VB.Timer timRefresh 
      Enabled         =   0   'False
      Left            =   3720
      Top             =   960
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfViewer 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2535
      _cx             =   4471
      _cy             =   3201
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "��ȡҩ���ˣ�"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label lblFormNO 
      AutoSize        =   -1  'True
      Caption         =   "��ҩ���ںţ�"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmUnSendDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------
'ESC���˳�����F12�����÷�ҩ���ڣ����˫���ָ���ʾ����
'--------------------------------------------------------

'��ʾ���˵ĸ���
Const INT_PATIENTS = 20
'���λ��
Const INT_INTERVAL = 200
'��ʾ������Ϣ��ȣ���λ����
Const INT_COLWIDTH = 3.5
Const HWND_NOTOPMOST = -2
Const SWP_NOACTIVATE = &H10

'��ʾ��Ļ
Public mbytScreen As Byte
'���ں����塢�ֺ�
Public mstrFormFont As String
Public mintFormSize As Integer
'��ȡҩ���塢�ֺ�
Public mstrDrugFont As String
Public mintDrugSize As Integer
'��ʾ�������塢�ֺ�
Public mstrPatientFont As String
Public mintPatientSize As Integer
'һ����ʾN��������Ϣ
Public mintCols As Integer
'������ɫ
Public mlngBackColorA As Long, mlngBackColorB As Long, mlngBackColorC As Long
'ǰ����ɫ
Public mlngForeColorA As Long, mlngForeColorB As Long, mlngForeColorC As Long

Public mlngStockID As Long

Private mobjMonitors As New cMonitors

Private Sub InitFont(ByVal intA As Integer, ByVal intB As Integer, ByVal intC As Integer)
    With lblFormNO
        .Font.Name = mstrFormFont
        .Font.Size = intA
        .BackColor = TransColor(False, mlngBackColorA)
        .ForeColor = TransColor(True, mlngForeColorA)
    End With
    With lblPatient
        .Top = lblFormNO.Top + lblFormNO.Height + INT_INTERVAL * 5
        .Left = lblFormNO.Left
        .Font.Name = mstrDrugFont
        .Font.Size = intB
        .BackColor = TransColor(False, mlngBackColorB)
        .ForeColor = TransColor(True, mlngForeColorB)
    End With
    With vsfViewer
        .Top = lblPatient.Top + lblPatient.Height + INT_INTERVAL
        .Left = lblFormNO.Left
        .Font.Name = mstrPatientFont
        .Font.Size = intC
    End With
End Sub

Private Sub InitVSFViewer()
    With vsfViewer
        .Enabled = False
        .BackColor = TransColor(False, mlngBackColorC)
        .ForeColor = TransColor(True, mlngForeColorC)
        .BackColorBkg = TransColor(False, mlngBackColorC)
        .SheetBorder = TransColor(False, mlngBackColorC) 'TransColor(True, mlngForeColorC)
        If .Rows > 0 Then .CellBackColor = TransColor(False, mlngBackColorC)
        If .Rows > 0 Then .CellForeColor = TransColor(True, mlngForeColorC)
        .GridLineWidth = 0
        .GridColor = TransColor(False, mlngBackColorC)
        .FixedCols = 0
        .FixedRows = 0
        .FloodColor = TransColor(True, mlngForeColorC)
        .BackColorSel = TransColor(False, mlngBackColorC)
        .ForeColorSel = TransColor(True, mlngForeColorC)
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarNone
        .ColWidthMax = INT_COLWIDTH * .Font.Size * 22
        .Cols = mintCols
        If .Rows > 0 Then .Row = 1
    End With
End Sub

Private Sub Form_DblClick()
    If mbytScreen = 0 Then
        Top = 0
        Left = 0
        Width = Screen.Width
        Height = Screen.Height
    Else
        Dim objMonitor As cMonitor
        Set objMonitor = mobjMonitors.Monitor(2)
        Top = 0
        Left = Screen.Width
        Width = objMonitor.Width * Screen.TwipsPerPixelX
        Height = objMonitor.Height * Screen.TwipsPerPixelY
    End If
    timReset.Interval = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Me.Tag = "1"
        frmSetWindow.Show vbModal, frmUserLogin
    ElseIf KeyCode = vbKeyF11 Then
        frmInterface.Show vbModal, frmUserLogin
        If mbytScreen <> 0 Then
            Dim objMonitor As cMonitor
            Set objMonitor = mobjMonitors.Monitor(1)
            
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Select Case Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ҩ������ʾ", Key:="����", Default:=""))
        Case 0: mstrFormFont = "����"
        Case 1: mstrFormFont = "����"
        Case Else: mstrFormFont = "����_GB2312"
    End Select
    Select Case Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ȡҩ������ʾ", Key:="����", Default:=""))
        Case 0: mstrDrugFont = "����"
        Case 1: mstrDrugFont = "����"
        Case Else:   mstrDrugFont = "����_GB2312"
    End Select
    Select Case Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="����", Default:=""))
        Case 0: mstrPatientFont = "����"
        Case 1: mstrPatientFont = "����"
        Case Else:   mstrPatientFont = "����_GB2312"
    End Select
    
    mintFormSize = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ҩ������ʾ", Key:="�ֺ�", Default:=""))
    If mintFormSize = 0 Then mintFormSize = 60
    mintDrugSize = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ȡҩ������ʾ", Key:="�ֺ�", Default:=""))
    If mintDrugSize = 0 Then mintDrugSize = 60
    mintPatientSize = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="�ֺ�", Default:=""))
    If mintPatientSize = 0 Then mintPatientSize = 72
    
    mlngForeColorA = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ҩ������ʾ", Key:="����ǰ��ɫ", Default:=""))
    mlngForeColorB = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ȡҩ������ʾ", Key:="����ǰ��ɫ", Default:=""))
    mlngForeColorC = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="����ǰ��ɫ", Default:=""))
    
    mlngBackColorA = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ҩ������ʾ", Key:="���屳��ɫ", Default:=""))
    mlngBackColorB = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\��ȡҩ������ʾ", Key:="���屳��ɫ", Default:=""))
    mlngBackColorC = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="���屳��ɫ", Default:=""))
    
    mintCols = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="����", Default:=""))
    If mintCols = 0 Then mintCols = 4
    
    mbytScreen = Val(GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ", Key:="��ʾ��Ļ", Default:=""))
    If mobjMonitors.MonitorCount = 1 Then mbytScreen = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturnValue As Long
    If Button = 1 Then
        timReset.Interval = 60000   'һ���Ӻ�ָ�ȫ������
        Call ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    End If
End Sub

Private Sub Form_Resize()
    With vsfViewer
        If mbytScreen = 0 Then
            .Width = Me.ScaleWidth - .Left * 2
            .Height = Me.ScaleHeight - .Top - INT_INTERVAL
        Else
            Dim objMonitor As cMonitor
            Set objMonitor = mobjMonitors.Monitor(2)
            .Width = objMonitor.Width * Screen.TwipsPerPixelX - .Left * 2
            .Height = objMonitor.Width * Screen.TwipsPerPixelY - .Top - INT_INTERVAL
        End If
    End With
End Sub

Private Sub FillData()
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim j As Integer
    With rsTmp
        strTmp = "select ���� from (" & _
                 "select distinct ����,to_char(��������,'yyyymmdd') from δ��ҩƷ��¼ " & _
                 "where ���� is not null and ��ҩ�� is not null " & _
                 "    and �ⷿid=" & mlngStockID & IIf(Val(lblFormNO.Tag) = 0, "", " and ��ҩ����='" & Me.lblFormNO.Tag & "'") & _
                 "    and ��������>(sysdate-24/24)" & _
                 " order by to_char(��������,'yyyymmdd')) a where rownum<=" & INT_PATIENTS
'        strTmp = "select distinct ����id,����,to_char(��������,'yyyymmdd') from δ��ҩƷ��¼ " & _
'                 "where ���� is not null and ����id is not null and nvl(δ����,0)>0 " & _
'                 " order by to_char(��������,'yyyymmdd') desc "
        .Open strTmp, gcnOracle
        vsfViewer.Clear
        vsfViewer.Rows = .RecordCount
        vsfViewer.Cols = mintCols
        vsfViewer.Redraw = flexRDNone
        i = 0: j = 0
        Do While Not .EOF
            vsfViewer.TextMatrix(i, j) = Trim(!����)
            If j >= mintCols - 1 Then
                vsfViewer.RowHeight(i) = vsfViewer.Font.Size * 26
                i = i + 1: j = 0
            Else
                j = j + 1
            End If
            .MoveNext
        Loop
        .Close
        vsfViewer.Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSetWindow
    Unload frmUserLogin
End Sub

Private Sub timRefresh_Timer()
    On Error Resume Next
    Call FillData
End Sub

Public Sub Entry(ByVal lngStockID As Long, ByVal strFormNO As String)
    Dim strVal As String
    Dim objMonitor As cMonitor
    
    lblFormNO.Tag = strFormNO
    lblFormNO.Caption = "��ҩ���ںţ�" & lblFormNO.Tag
    mlngStockID = lngStockID

    If mbytScreen = 0 Then
        Top = 0
        Left = 0
        Width = Screen.Width
        Height = Screen.Height
    Else
        Set objMonitor = mobjMonitors.Monitor(2)
        Top = 0
        Left = Screen.Width
        Width = objMonitor.Width * Screen.TwipsPerPixelX
        Height = objMonitor.Height * Screen.TwipsPerPixelY
    End If

    ForeColor = TransColor(True, mlngForeColorA)
    BackColor = TransColor(False, mlngBackColorA)
    BorderStyle = 0
    
    InitFont mintFormSize, mintDrugSize, mintPatientSize
    InitVSFViewer
    FillData
    With timRefresh
        strVal = GetSetting(appName:="ZLSOFT", Section:="δ��ҩ������ʾ\����������ʾ", Key:="ˢ������", Default:="")
        .Enabled = True
        .Interval = IIf(Trim(strVal) = "", 30000, Val(strVal) * 1000)
    End With

End Sub

Private Sub timReset_Timer()
    Call Form_DblClick
    timReset.Interval = 0
End Sub

