VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegist 
   BackColor       =   &H80000005&
   Caption         =   "�û�ע�����"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   Picture         =   "frmRegist.frx":0000
   ScaleHeight     =   7425
   ScaleWidth      =   8025
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSpecReg 
      Caption         =   "�鿴������Ȩ��(&O)"
      Height          =   345
      Left            =   5040
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1095
      TabIndex        =   14
      Tag             =   "����"
      Top             =   3300
      Width           =   4170
   End
   Begin VSFlex8Ctl.VSFlexGrid vsInfo 
      Height          =   1275
      Left            =   735
      TabIndex        =   0
      Top             =   555
      Width           =   6360
      _cx             =   1980181458
      _cy             =   1980172489
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegist.frx":04F9
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin MSComctlLib.ProgressBar pgbRegist 
      Height          =   75
      Left            =   375
      TabIndex        =   1
      Top             =   2415
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "У��(&V)"
      Height          =   350
      Left            =   3945
      TabIndex        =   4
      Top             =   2550
      Width           =   1100
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   6885
      TabIndex        =   9
      Top             =   3345
      Width           =   675
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   6105
      TabIndex        =   8
      Top             =   3345
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ϵͳ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5325
      TabIndex        =   7
      Top             =   3345
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "��ԭ(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6300
      TabIndex        =   6
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ��(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5115
      TabIndex        =   5
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdRegist 
      Caption         =   "����ע��(&R)��"
      Height          =   350
      Left            =   375
      TabIndex        =   2
      Top             =   2550
      Width           =   1440
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgFunc 
      Height          =   3555
      Left            =   375
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3705
      Width           =   7170
      _cx             =   12647
      _cy             =   6271
      Appearance      =   0
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
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
   Begin VB.Label lblFind 
      BackStyle       =   0  'Transparent
      Caption         =   "����(&Z)"
      Height          =   255
      Left            =   375
      TabIndex        =   13
      Top             =   3345
      Width           =   690
   End
   Begin VB.Label lblRegist 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ע�ᣬ���Ե�..."
      Height          =   210
      Left            =   1905
      TabIndex        =   3
      Top             =   2655
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û�ע�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   12
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   150
      Picture         =   "frmRegist.frx":05FB
      Top             =   570
      Width           =   480
   End
   Begin VB.Label lblRegFunc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ѱ�װϵͳ��Ӧ����Ȩ��"
      Height          =   180
      Left            =   375
      TabIndex        =   11
      Top             =   3060
      Width           =   1980
   End
End
Attribute VB_Name = "frmRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conIdent = 4  '����͹�����Ȩ��¼�������ո�

Dim strSQL As String
Dim lngCount As Long
'---------------------------------------------
Dim mstrRegCode As String      '��ʱ������ע����
Dim mblnIsCancel As Boolean
Dim mintIndex As Integer
Dim mintCount As Integer        '��λʱ��¼��һ�ζ�λ��λ��


Private Sub cmdApply_Click()
    Dim blnAudit As Boolean
    Dim strRegError As String
    
    err = 0: On Error GoTo errHand
     
    Me.MousePointer = vbHourglass
    
    gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
    
    Me.Tag = ""
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    
    '�ٴε�����֤���Ա�֤��Ϣ��ȷ��
    strRegError = gobjRegister.zlRegCheck(False)
    Me.MousePointer = vbDefault
    
    If strRegError = "" Then
        SaveSetting "ZLSOFT", "ע����Ϣ", "��λ����", gobjRegister.zlRegInfo("��λ����", , -1)
        MsgBox "ע����Ȩ��Ϣ�Ѿ�Ӧ�ã�", vbInformation, gstrSysName
    Else
        MsgBox strRegError, vbExclamation, gstrSysName
    End If
    Exit Sub
errHand:
    MsgBox "Ӧ��ʧ�ܣ������ļ�����ȷ�ԣ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Call zlRefGrant
    Me.Tag = ""
    Me.vfgFunc.SetFocus
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    MsgBox "ע����Ȩ��Ϣ�ѻ�ԭ��", vbInformation, gstrSysName
End Sub

Private Sub cmdRegist_Click()
    Dim strFile As String, strRegError As String
    Dim rsFile As New ADODB.Recordset
    Dim blnApplyEnabled As Boolean, blnCancelEnabled As Boolean, blnVerifyEnabled As Boolean
    Dim i As Integer, blnNotPrompt As Boolean
    
    With frmMDIMain.DlgMain
        .FileName = ""
        .DialogTitle = "ѡ��ע����Ȩ�ļ�"
        .Filter = "(ע����Ȩ�ļ�)|*.zcr"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        strFile = .FileName
    End With
        
    Me.cmdRegist.Enabled = False
    
    '��¼��ťԭ����enabled����
    blnApplyEnabled = Me.cmdApply.Enabled
    blnCancelEnabled = Me.cmdCancel.Enabled
    blnVerifyEnabled = Me.cmdVerify.Enabled
    
    '���ð�ť
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdVerify.Enabled = False
    For i = 0 To optGrade.UBound
        If Not optGrade(i).value Then optGrade(i).Enabled = False
    Next
    For i = 0 To optGrade.UBound
        If optGrade(i).value Then optGrade(i).Enabled = False
    Next
    
    err = 0: On Error GoTo errHand
        
    lblRegist.Visible = True
    Me.MousePointer = vbHourglass
    
    If gobjRegister.zlRegBuild(strFile, pgbRegist) = False Then
        lblRegist.Visible = False
        Me.MousePointer = vbDefault
        
        blnNotPrompt = True
        GoTo errHand
    End If
    
    lblRegist.Visible = False
    Me.MousePointer = vbDefault
    
    Me.cmdRegist.Enabled = True
    
    '��ԭ��ť��enabled����
    Me.cmdApply.Enabled = blnApplyEnabled
    Me.cmdCancel.Enabled = blnCancelEnabled
    Me.cmdVerify.Enabled = blnVerifyEnabled
    '���ÿؼ�
    optGrade(0).Enabled = True
    optGrade(1).Enabled = True
    optGrade(2).Enabled = True
    
    strRegError = gobjRegister.zlRegCheck(True)
    If strRegError = "" Then
        Call zlRefGrant(True)
        Me.Tag = "�޸�"
        cmdApply.Enabled = True
        cmdCancel.Enabled = True
    Else
        Call zlRefGrant
        MsgBox strRegError & vbCrLf & "ϵͳ�Ѿ��Զ���ԭ��", vbExclamation, gstrSysName
        Me.Tag = ""
        Me.cmdApply.Enabled = False
        Me.cmdCancel.Enabled = False
    End If
    Me.vfgFunc.SetFocus
    Exit Sub

errHand:
    Me.cmdRegist.Enabled = True
    Me.cmdApply.Enabled = False
    
    '��ԭ��ť��enabled����
    Me.cmdCancel.Enabled = blnCancelEnabled
    Me.cmdVerify.Enabled = blnVerifyEnabled
    '���ÿؼ�
    optGrade(0).Enabled = True
    optGrade(1).Enabled = True
    optGrade(2).Enabled = True
    
    If Not blnNotPrompt Then MsgBox "ע����Ȩ�ļ�ʱ���ִ������飡" & vbNewLine & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub cmdSpecReg_Click()
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, strSQL As String
    Dim blnFirst As Boolean
    
    On Error GoTo errHandle
    
    Set objPopup = gcbsMain.Add("Popup", xtpBarPopup)
    
    strSQL = "Select Item, Prog, Text From Table(Cast(zltools.f_Reg_Info(" & IIf(cmdApply.Enabled, 1, 0) & ") As zlTools.t_Reg_Rowset))"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    
    With objPopup.Controls
        rsTemp.Filter = "Item='�ƶ���ʿվ��Ȩ����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 1, "��ʽ", 2, "����", 3, "����"))
        End If
        rsTemp.Filter = "Item='�ƶ���ʿվ��Ȩ����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 0, "������", rsTemp!Text & "��"))
        End If
        rsTemp.Filter = "Item='�ƶ���ʿվ�豸����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 0, "������", rsTemp!Text & "̨"))
        End If
        rsTemp.Filter = "Item='�ƶ�ҽ��վ��Ȩ����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 1, "��ʽ", 2, "����", 3, "����"))
        End If
        rsTemp.Filter = "Item='�ƶ�ҽ��վ��Ȩ����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 0, "������", rsTemp!Text & "��"))
        End If
        rsTemp.Filter = "Item='�ƶ�ҽ��վ�豸����'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & Decode(Val(Nvl(rsTemp!Text)), 0, "������", rsTemp!Text & "̨"))
        End If
    End With
    
    rsTemp.Filter = "Prog=-1"
    If Not rsTemp.EOF Then
        blnFirst = True
        With objPopup.Controls
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "��" & rsTemp!Text)
                If blnFirst Then objControl.BeginGroup = True
                blnFirst = False
                rsTemp.MoveNext
            Loop
        End With
    End If
        
    If objPopup.Controls.Count > 0 Then
        GetWindowRect Me.hwnd, vRect
        objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + cmdSpecReg.Left, vRect.Top * Screen.TwipsPerPixelY + cmdSpecReg.Top + cmdSpecReg.Height
    Else
        MsgBox "�������ض���Ȩ��Ŀ���ݡ�", vbInformation, Me.Caption
    End If
    
    Exit Sub
errHandle:
    MsgBox err.Number & ":" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub cmdVerify_Click()
    Dim strRegError As String
    Me.MousePointer = vbHourglass
    strRegError = gobjRegister.zlRegCheck(IIf(Me.Tag = "�޸�", True, False))
    Me.MousePointer = vbDefault
    If strRegError = "" Then
        MsgBox "��ǰע����Ȩ�ļ���ȷ��Ч��", vbInformation, gstrSysName
    Else
        MsgBox strRegError, vbExclamation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    Call zlRefGrant
End Sub

Private Sub Form_Deactivate()
    If Tag = "�޸�" Then
        If MsgBox("�Ѿ��޸���ע����Ϣ����������棬�����Զ���ԭ��" & vbCr & "�Ƿ񱣴棿", vbQuestion + vbYesNo) = vbYes Then
            Call cmdApply_Click
        Else
            Call cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()
    '�������ʼ��
    txtFind.Text = "�������Ż�ؼ���"
    txtFind.ForeColor = vbGrayText
    mintCount = -1
    
    mblnIsCancel = False
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.vfgFunc.Height = Me.ScaleHeight - Me.vfgFunc.Top - 150
    
End Sub

Private Sub optGrade_Click(Index As Integer)
    With Me.vfgFunc
        .Redraw = flexRDNone
        For lngCount = .FixedRows To .Rows - 1
            Select Case Index
            Case 0
                .RowHidden(lngCount) = (Val(.TextMatrix(lngCount, 2)) > -2)
            Case 1
                .RowHidden(lngCount) = (Val(.TextMatrix(lngCount, 2)) > -1)
            Case 2
                .RowHidden(lngCount) = False
            End Select
        Next
        .Redraw = flexRDDirect
    End With
End Sub
    
'--------------------------------------------------
'���ܣ������ݿ���ļ�ˢ����Ȩ��Ϣ
'������blnTemp-�Ƿ��ļ�ˢ��
'--------------------------------------------------
Private Sub zlRefGrant(Optional blnTemp As Boolean)
    Dim rsTemp As New ADODB.Recordset
    Dim intKind As Integer, intLimit As Integer, intStation As Integer
    Dim strUnitName As String, i As Integer
    
    On Error GoTo errHandle
    '��Ȩ��Ϣ
    With vsInfo
        strUnitName = gobjRegister.zlRegInfo("��λ����", blnTemp, -1)
        .TextMatrix(0, 1) = Replace(strUnitName, ";", vbCrLf)
        If strUnitName <> "" Then
            i = UBound(Split(strUnitName, ";")) + 1
        Else
            i = 1
        End If
        .Height = .rowHeight(1) * (5.5 + i - 1)
        .rowHeight(0) = .rowHeight(1) * i
        .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 2, 0, 2) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 1, 0, 1) = flexAlignLeftTop
        .Cell(flexcpAlignment, 0, 3, 0, 3) = flexAlignLeftTop
        cmdSpecReg.Top = .Top + .Height + 30
        
        Select Case Val(gobjRegister.zlRegInfo("��Ȩ����", blnTemp))
            Case 1: intKind = 1: .TextMatrix(1, 1) = "��ʽ�汾"
            Case 2: intKind = 2: .TextMatrix(1, 1) = "���ð汾"
            Case Else: intKind = 3: .TextMatrix(1, 1) = "���԰汾"
        End Select
        
        .TextMatrix(2, 1) = "������"
        If intKind <> 1 Then
            intLimit = Val(gobjRegister.zlRegInfo("ʹ������", blnTemp))
            If intLimit > 0 Then .TextMatrix(2, 1) = "����" & intLimit & "��"
        End If
        
        intStation = Val(gobjRegister.zlRegInfo("��Ȩվ��", blnTemp))
        If intStation = 0 Then
            .TextMatrix(3, 1) = "������"
        Else
            .TextMatrix(3, 1) = "������" & intStation & "վ��"
        End If
        .TextMatrix(4, 1) = gobjRegister.zlRegInfo("��Ȩ����", blnTemp)
        
        'PACS/LIS��Ȩ
        .TextMatrix(0, 3) = gobjRegister.zlRegInfo("Ӱ��DICOM�豸����", blnTemp)
        If .TextMatrix(0, 3) = "" Then .TextMatrix(0, 3) = "������"
    
        .TextMatrix(1, 3) = gobjRegister.zlRegInfo("Ӱ����Ƶ�豸����", blnTemp)
        If .TextMatrix(1, 3) = "" Then .TextMatrix(1, 3) = "������"
    
        .TextMatrix(2, 3) = gobjRegister.zlRegInfo("Ӱ��Ƭ��ӡ������", blnTemp)
        If .TextMatrix(2, 3) = "" Then .TextMatrix(2, 3) = "������"
    
        .TextMatrix(3, 3) = gobjRegister.zlRegInfo("Ӱ���Ƭվ����", blnTemp)
        If .TextMatrix(3, 3) = "" Then .TextMatrix(3, 3) = "������"
    
        .TextMatrix(4, 3) = gobjRegister.zlRegInfo("������������", blnTemp)
        If .TextMatrix(4, 3) = "" Then .TextMatrix(4, 3) = "������"
    
    End With
    
    '��Ȩ����
    If blnTemp Then
        strSQL = "Select Distinct r.ϵͳ, 0 As ���, -2 As ����, r.ϵͳ || '-' || u.���� As ����" & _
                " From zlRegFile r, zlSystems u, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s" & _
                " Where r.ϵͳ = Trunc(u.��� / 100) And u.��� = s.��� And r.��Ŀ = '��Ȩ����' And r.���� = '����'" & _
                " Union All" & _
                " Select Distinct r.ϵͳ, r.���, -1 As ����, '" & Space(conIdent) & "' || r.��� || '-' || p.���� As ����" & _
                " From zlRegFile r, zlPrograms p, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s,zlRPTGroups g" & _
                " Where r.ϵͳ = Trunc(p.ϵͳ / 100) And r.��� = p.��� And p.ϵͳ = s.��� And r.��Ŀ = '��Ȩ����'" & _
                "   And p.ϵͳ=g.ϵͳ(+) And p.���=g.����ID(+) And (r.���� = '����' Or g.����ID is Not Null)" & _
                " Union All" & _
                " Select r.ϵͳ, r.���, Nvl(f.����, 0) As ����, '" & Space(conIdent * 2) & "' || f.���� As ����" & _
                " From zlRegFile r, zlProgfuncs f, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s" & _
                " Where r.ϵͳ = Trunc(f.ϵͳ / 100) And r.��� = f.��� And r.���� = f.���� And f.ϵͳ = s.��� And r.��Ŀ = '��Ȩ����' And r.���� <> '����'" & _
                " Order By ϵͳ, ���, ����"
    Else
        strSQL = "Select Distinct r.ϵͳ, 0 As ���, -2 As ����, r.ϵͳ || '-' || u.���� As ����" & _
                " From zlRegFunc r, zlSystems u, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s" & _
                " Where r.ϵͳ = Trunc(u.��� / 100) And u.��� = s.��� And r.���� = '����'" & _
                " Union All" & _
                " Select Distinct r.ϵͳ, r.���, -1 As ����, '" & Space(conIdent) & "' || r.��� || '-' || p.���� As ����" & _
                " From zlRegFunc r, zlPrograms p, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s,zlRPTGroups g" & _
                " Where r.ϵͳ = Trunc(p.ϵͳ / 100) And r.��� = p.��� And p.ϵͳ = s.���" & _
                "   And p.ϵͳ=g.ϵͳ(+) And p.���=g.����ID(+) And (r.���� = '����' Or g.����ID is Not Null)" & _
                " Union All" & _
                " Select r.ϵͳ, r.���, Nvl(f.����, 0) As ����, '" & Space(conIdent * 2) & "' || f.���� As ����" & _
                " From zlRegFunc r, zlProgfuncs f, (Select Min(���) As ��� From zlSystems Group By Trunc(��� / 100)) s" & _
                " Where r.ϵͳ = Trunc(f.ϵͳ / 100) And r.��� = f.��� And r.���� = f.���� And f.ϵͳ = s.��� And r.���� <> '����'" & _
                " Order By ϵͳ, ���, ����"
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    
    With Me.vfgFunc
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(1) = 0: .ColHidden(1) = True
        .ColWidth(2) = 0: .ColHidden(2) = True
    End With
    Me.optGrade(1).value = True
    Call optGrade_Click(1)
    Exit Sub
errHandle:
    MsgBox "ˢ����Ȩ��Ϣʱ���ִ�����ʾ����Ȩ��Ϣ���ܲ���ȷ��" & vbNewLine & err.Description, vbExclamation, Me.Caption
End Sub

'--------------------------------------------------
'�������߹淶�������ṩ�ĺ���
'--------------------------------------------------
Public Function SupportPrint() As Boolean
    '���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
    '�������ڵ��ã�ʵ�־���Ĵ�ӡ����
    '���û�пɴ�ӡ�ģ�������һ���յĽӿ�
    
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = "�û�ע����Ϣ"
    
    objRow.Add vsInfo.TextMatrix(0, 0) & vsInfo.TextMatrix(0, 1)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(date, "yyyy��MM��dd��")
    Set objPrint.Body = Me.vfgFunc
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub txtFind_Change()
    mintCount = -1
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" And txtFind.ForeColor = vbGrayText Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    Else
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim blnFindTag As Boolean
    
    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        txtFind.Text = Replace(txtFind.Text, " ", "")
        With vfgFunc
            blnFindTag = False
            For intRow = mintCount + 1 To vfgFunc.Rows - 1
                If .RowHidden(intRow) = False And InStr(.TextMatrix(intRow, 3), txtFind.Text) > 0 Then blnFindTag = True: Exit For
            Next
            If blnFindTag Then .Row = intRow: .ShowCell intRow, 3: mintCount = intRow
            If intRow = .Rows Then
                If mintCount = -1 Then
                    Call MsgBox("δ�ҵ��롰" & txtFind.Text & "��ƥ�����Ŀ�������������Ż�ؼ��֡�", vbInformation, gstrSysName)
                    txtFind.Text = "": txtFind.SetFocus
                Else
                    mintCount = -1
                End If
            End If
        End With
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "�������Ż�ؼ���"
        txtFind.ForeColor = vbGrayText
    End If
End Sub


