VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatientFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˲���"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7755
   Icon            =   "frmPatientFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6240
      TabIndex        =   26
      Top             =   510
      Width           =   1395
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "��ʼ����(&F)"
      Height          =   350
      Left            =   6240
      TabIndex        =   25
      Top             =   90
      Width           =   1395
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "ѡ��(&S)"
      Height          =   350
      Left            =   6255
      TabIndex        =   24
      Top             =   2025
      Width           =   1395
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6240
      TabIndex        =   23
      Top             =   930
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   5400
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1275
         TabIndex        =   8
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1275
         TabIndex        =   7
         Top             =   585
         Width           =   1485
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1275
         TabIndex        =   6
         Top             =   1650
         Width           =   3510
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   3690
         TabIndex        =   5
         Top             =   585
         Width           =   1110
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   3690
         TabIndex        =   4
         Top             =   945
         Width           =   1110
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   3690
         TabIndex        =   3
         Top             =   1305
         Width           =   1110
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   945
         Width           =   1485
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   3675
         TabIndex        =   1
         Top             =   240
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   1275
         TabIndex        =   9
         Top             =   1305
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114819075
         CurrentDate     =   38329
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         Top             =   2010
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114819075
         CurrentDate     =   37401
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1275
         TabIndex        =   11
         Top             =   2010
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   114819075
         CurrentDate     =   37401
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����I&D"
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   22
         Top             =   300
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����(&M)"
         Height          =   180
         Index           =   1
         Left            =   2835
         TabIndex        =   21
         Top             =   645
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��(&Z)"
         Height          =   180
         Index           =   2
         Left            =   2835
         TabIndex        =   20
         Top             =   1020
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&E)"
         Height          =   180
         Index           =   3
         Left            =   3015
         TabIndex        =   19
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   4
         Left            =   585
         TabIndex        =   18
         Top             =   645
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Index           =   5
         Left            =   585
         TabIndex        =   17
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&I)"
         Height          =   180
         Index           =   6
         Left            =   225
         TabIndex        =   16
         Top             =   1725
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   7
         Left            =   225
         TabIndex        =   15
         Top             =   1380
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   8
         Left            =   3015
         TabIndex        =   14
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&T)"
         Height          =   180
         Left            =   2790
         TabIndex        =   13
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴξ���(&L)"
         Height          =   180
         Left            =   225
         TabIndex        =   12
         Top             =   2070
         Width           =   990
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2955
      Left            =   45
      TabIndex        =   27
      Top             =   2445
      Width           =   7665
      _cx             =   13520
      _cy             =   5212
      Appearance      =   1
      BorderStyle     =   1
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   5400
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13626
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPatientFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long

'�������Զ�����̻���************************************************************************************************
Public Function ShowFind(ByVal frmMain As Object, ByRef lngKey As Long) As Boolean
    
    mblnStartUp = True
    mblnOK = False
        
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
'    Call AppendSapceRows(vsf, lnX, lnY)
    
    Me.Show 1, frmMain
    
    lngKey = mlngKey
    
    ShowFind = mblnOK
    
End Function

Private Function InitData() As Boolean
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    Dim strVsf As String
    
    mlngKey = 0
    
    dtp(2).MaxDate = zlDatabase.Currentdate
    dtp(1).MaxDate = dtp(2).MaxDate

    dtp(1).Value = Format(zlDatabase.Currentdate - 30, dtp(1).CustomFormat)
    dtp(2).Value = Format(zlDatabase.Currentdate, dtp(2).CustomFormat)
    
    dtp(1).Value = Null
    dtp(2).Value = Null
    
    dtp(0).MaxDate = zlDatabase.Currentdate
    dtp(0).Value = DateAdd("yyyy", -25, zlDatabase.Currentdate)
    dtp(0).Value = Null
    
    On Error GoTo errH
    
    strSQL = "Select * From �Ա�"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo(0).AddItem ""
    Do While Not rs.EOF
        cbo(0).AddItem rs!����
        rs.MoveNext
    Loop
    cbo(0).ListIndex = 0
    
    InitData = True
        
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        KeyAscii = 0
                
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    '���ܣ����ݵ�ǰ�����������Ҳ���
    
    Dim strSQL As String
    Dim lngLoop As Integer
    Dim rs As New ADODB.Recordset
    
    cmdFind.Enabled = False
    
    If Trim(txt(0).Text) <> "" Then
        strSQL = strSQL & " And ����ID=" & Val(txt(0).Text)
    End If
    If Trim(txt(3).Text) <> "" Then
        strSQL = strSQL & " And �����=" & IIf(IsNumeric(txt(3).Text), txt(3).Text, "0")
    End If
    
    If Trim(txt(4).Text) <> "" Then
        strSQL = strSQL & " And סԺ��=" & IIf(IsNumeric(txt(4).Text), txt(4).Text, "0")
    End If
    
    If Trim(txt(5).Text) <> "" Then
        strSQL = strSQL & " And ��ǰ����=" & Val(txt(5).Text)
    End If
    
    If Trim(txt(1).Text) <> "" Then
        strSQL = strSQL & " And Upper(����) Like '%" & UCase(txt(1).Text) & "%'"
    End If
    If cbo(0).Text <> "" Then
        strSQL = strSQL & " And �Ա�='" & cbo(0).Text & "'"
    End If
    If Trim(txt(6).Text) <> "" Then
        strSQL = strSQL & " And ����='" & txt(6).Text & "'"
    End If
    If Not IsNull(dtp(0).Value) Then
        strSQL = strSQL & " And ��������=To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    End If
    If Trim(txt(2).Text) <> "" Then
        strSQL = strSQL & " And ���֤��='" & txt(2).Text & "'"
    End If
            
    If Not IsNull(dtp(1).Value) And Not IsNull(dtp(2).Value) Then
        If dtp(2).Value <= dtp(1).Value Then
            MsgBox "�ϴξ���Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, ParamInfo.ϵͳ����
            dtp(2).SetFocus
            cmdFind.Enabled = True
            Exit Sub
        End If
        strSQL = strSQL & " And ����ʱ�� Between To_Date('" & Format(dtp(1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " And To_Date('" & Format(dtp(2).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    ElseIf Not IsNull(dtp(1).Value) Then
        strSQL = strSQL & " And ����ʱ�� Between To_Date('" & Format(dtp(1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " And To_Date('" & Format(dtp(1).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    If strSQL = "" Then
        MsgBox "����������һ������������", vbInformation, ParamInfo.ϵͳ����
        txt(1).SetFocus
        cmdFind.Enabled = True
        Exit Sub
    End If
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    strSQL = _
        " Select " & _
        " ����ID,�����,�ѱ�,סԺ��,��ǰ����,����,�Ա�,����,To_Char(��������,'YYYY-MM-DD') as ��������," & _
        " ���֤��,�����ص�,��ͥ��ַ,������λ,���,ְҵ,ѧ��,To_Char(����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��" & _
        " From ������Ϣ" & _
        " Where ͣ��ʱ�� is NULL " & strSQL & _
        " Order by ����ID"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rs.BOF = False Then
        
        stbThis.Panels(1).Text = " ���˲��ҽ��:�� " & rs.RecordCount & " �����������Ĳ���"
        
        Set vsf.DataSource = rs
        vsf.Cols = vsf.Cols + 1
        
        vsf.TextMatrix(0, 0) = "����ID": vsf.ColWidth(0) = 750: vsf.ColAlignment(0) = 1
        vsf.TextMatrix(0, 1) = "�����": vsf.ColWidth(1) = 750: vsf.ColAlignment(1) = 1
        vsf.TextMatrix(0, 2) = "�ѱ�": vsf.ColWidth(2) = 850: vsf.ColAlignment(2) = 1
        vsf.TextMatrix(0, 3) = "סԺ��": vsf.ColWidth(3) = 850: vsf.ColAlignment(3) = 1
        vsf.TextMatrix(0, 4) = "��ǰ����": vsf.ColWidth(4) = 850: vsf.ColAlignment(4) = 1
        vsf.TextMatrix(0, 5) = "����": vsf.ColWidth(5) = 700: vsf.ColAlignment(5) = 1
        vsf.TextMatrix(0, 6) = "�Ա�": vsf.ColWidth(6) = 500: vsf.ColAlignment(6) = 4
        vsf.TextMatrix(0, 7) = "����": vsf.ColWidth(7) = 500: vsf.ColAlignment(7) = 1
        vsf.TextMatrix(0, 8) = "��������": vsf.ColWidth(8) = 1000: vsf.ColAlignment(8) = 4
        vsf.TextMatrix(0, 9) = "���֤��": vsf.ColWidth(9) = 1600: vsf.ColAlignment(9) = 1
        vsf.TextMatrix(0, 10) = "�����ص�": vsf.ColWidth(10) = 2000: vsf.ColAlignment(10) = 1
        vsf.TextMatrix(0, 11) = "��ͥ��ַ": vsf.ColWidth(11) = 2000: vsf.ColAlignment(11) = 1
        vsf.TextMatrix(0, 12) = "������λ": vsf.ColWidth(12) = 2000: vsf.ColAlignment(12) = 1
        vsf.TextMatrix(0, 13) = "���": vsf.ColWidth(13) = 1000: vsf.ColAlignment(13) = 1
        vsf.TextMatrix(0, 14) = "ְҵ": vsf.ColWidth(14) = 1000: vsf.ColAlignment(14) = 1
        vsf.TextMatrix(0, 15) = "ѧ��": vsf.ColWidth(15) = 500: vsf.ColAlignment(15) = 1
        vsf.TextMatrix(0, 16) = "�ϴξ���ʱ��": vsf.ColWidth(16) = 1600: vsf.ColAlignment(16) = 4
    Else
        
        stbThis.Panels(1).Text = " û�в��ҵ����������Ĳ���"
        
        vsf.Clear
        vsf.Cols = 2
        vsf.Rows = 2
        vsf.FixedCols = 0
        vsf.FixedRows = 1
        
    End If
    vsf.Row = 1: vsf.Col = 0: vsf.ColSel = vsf.Cols - 1
    vsf.TopRow = 1
    
    Screen.MousePointer = 0
    vsf.SetFocus
    
'    Call AppendSapceRows(vsf, lnX, lnY)
    
    cmdFind.Enabled = True
    
    Exit Sub
    
errH:
    
'    Call AppendSapceRows(vsf, lnX, lnY)
    
    cmdFind.Enabled = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSelect_Click()
    If Val(vsf.TextMatrix(vsf.Row, 0)) = 0 Then
        MsgBox "û�в�����Ϣ����ѡ��", vbInformation, ParamInfo.ϵͳ����
        Exit Sub
    End If
    mlngKey = Val(vsf.TextMatrix(vsf.Row, 0))
    
    mblnOK = True
    
    Unload Me
End Sub


Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
    
        KeyCode = 0
                
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 1
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
    
        KeyAscii = 0
                
        zlCommFun.PressKey vbKeyTab
        
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1
        zlCommFun.OpenIme False
    End Select
    

    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_DblClick()
    
    Call cmdSelect_Click
    
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsf_DblClick
End Sub


