VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLeaveMedi 
   BorderStyle     =   0  'None
   Caption         =   "ҩƷ�Ĵ����"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   Icon            =   "frmLeaveMedi.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTow 
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   345
      ScaleHeight     =   2580
      ScaleWidth      =   7005
      TabIndex        =   5
      Top             =   660
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid vsListUsed 
         Height          =   3255
         Left            =   150
         TabIndex        =   6
         Top             =   165
         Width           =   7800
         _cx             =   13758
         _cy             =   5741
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLeaveMedi.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   End
   Begin VB.PictureBox picOne 
      BorderStyle     =   0  'None
      Height          =   2580
      Left            =   960
      ScaleHeight     =   2580
      ScaleWidth      =   7005
      TabIndex        =   3
      Top             =   3300
      Width           =   7005
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   3255
         Left            =   -435
         TabIndex        =   4
         Top             =   195
         Width           =   7800
         _cx             =   13758
         _cy             =   5741
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmLeaveMedi.frx":68ED
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   3045
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   5371
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2145
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeaveMedi.frx":6988
            Key             =   "mx"
            Object.Tag             =   "mx"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeaveMedi.frx":D1EA
            Key             =   "use"
            Object.Tag             =   "use"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picNS 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   1875
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6615
      TabIndex        =   1
      Top             =   2955
      Width           =   6615
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMaster 
      Height          =   2445
      Left            =   270
      TabIndex        =   0
      Top             =   450
      Width           =   7800
      _cx             =   13758
      _cy             =   4313
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
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmLeaveMedi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const subMenu_Add = 101 '����
Private Const subMenu_Modify = 102 '�޸�
Private Const subMenu_Delete = 103 'ɾ��
Private Const subMenu_Post = 104 'ʹ�õǼ�
Private Const subMenu_Repertory = 105 '����ѯ
Private Const subMenu_AccountBook = 106 '���̨��

Private mlng����ID As Long
Private mlng����ID As Long
Private mdateBeging As Date
Private mdateEnd As Date
Private mobjMasters As MediMasters
Private mstr���� As String
Private mstr���� As String
Private mstr�Ա� As String
Private mstr���� As String
Private mstr�Һŵ� As String

Private mblnEditList As Boolean '�༭״̬
Private mcbsMain As CommandBars

Public Property Let ����(ByVal vData As String)
    mstr���� = vData
End Property

Public Property Let �Ա�(ByVal vData As String)
    mstr�Ա� = vData
End Property

Public Property Let ����(ByVal vData As String)
    mstr���� = vData
End Property

Public Property Let ����(ByVal vData As String)
    mstr���� = vData
End Property

Public Property Let �Һŵ�(ByVal vData As String)
    mstr�Һŵ� = vData
End Property

Public Property Let ����ID(ByVal vData As Long)
    mlng����ID = vData
End Property

Public Property Let ����ID(ByVal vData As Long)
    mlng����ID = vData
End Property

Public Property Let dateBeging(ByVal vData As Date)
    mdateBeging = vData
End Property

Public Property Let DateEnd(ByVal vData As Date)
    mdateEnd = vData
End Property

Private Sub cbsSub_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    'Call dkpSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
'    On Error Resume Next
'    With vsMaster
'        .Left = 0
'        .Top = lngTop
'        .Width = lngRight
'    End With
''    With picNS
''        .Left = 0
''        .Top = vsMaster.Top + vsMaster.Height
''        .Width = lngRight
''    End With
''    With tabRecord
''        .Left = 0
''        .Top = picNS.Top + picNS.Height
''    End With
'    With vsList
'        .Left = 0
'        .Top = tabRecord.Top + tabRecord.Height + 15
'        .Width = lngRight
'        .Height = lngBottom - lngTop - vsMaster.Height - picNS.Height - tabRecord.Height - 60
'    End With

    vsList.Top = lngTop
    vsList.Left = lngLeft
    vsList.Height = (lngBottom - lngTop) / 2
    vsList.Width = lngRight - lngLeft

    vsMaster.Top = lngTop
    vsMaster.Left = lngLeft
    vsMaster.Height = (lngBottom - lngTop) / 2
    vsMaster.Width = lngRight - lngLeft

    vsListUsed.Top = lngTop
    vsListUsed.Left = lngLeft
    vsListUsed.Height = (lngBottom - lngTop) / 2
    vsListUsed.Width = lngRight - lngLeft
End Sub

Private Sub dkpSub_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    'If Action = PaneActionFloating Then
'    dkpSub.PanelPaintManager.Position = xtpTabPositionBottom
    'End If
End Sub

Private Sub picNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsMaster.Height + Y < 1000 Or tbcSub.Height - Y < 800 Then Exit Sub
        picNS.Top = picNS.Top + Y
        vsMaster.Height = vsMaster.Height + Y
        
    
        tbcSub.Top = tbcSub.Top + Y
        tbcSub.Height = tbcSub.Height - Y
         
    End If
End Sub
Private Sub Form_Load()
    Call Fill_Master
End Sub

Public Sub ShowLeaveMedi(ByVal lng����ID As Long, Optional ByVal lng����ID As Long)
    mlng����ID = lng����ID
    mlng����ID = lng����ID
End Sub

Public Sub Fill_Master()
    Dim strHead As String, objMaster As MediMaster
    strHead = "NO,1000,1;����,900,1;�Ա�,450,4;����,450,1;����Ա,900,1;�Ǽ�ʱ��,1500,1;�ϼ�,900,7;ʹ�����,900,1;ժҪ,1000,1;����ʱ��,0,1;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsMaster)
    Set mobjMasters = New MediMasters
    Call mobjMasters.GetMediMasters(mdateBeging, mdateEnd, mlng����ID, mlng����ID)
    
    If mobjMasters Is Nothing Then Exit Sub
    
    For Each objMaster In mobjMasters
        With vsMaster
            .TextMatrix(.Rows - 1, .ColIndex("NO")) = objMaster.NO
            .TextMatrix(.Rows - 1, .ColIndex("����")) = objMaster.����
            .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = objMaster.�Ա�
            .TextMatrix(.Rows - 1, .ColIndex("����")) = objMaster.����
            .TextMatrix(.Rows - 1, .ColIndex("����Ա")) = objMaster.����Ա
            .TextMatrix(.Rows - 1, .ColIndex("�Ǽ�ʱ��")) = Format(objMaster.�Ǽ�ʱ��, "yy-MM-dd hh:mm")
            .TextMatrix(.Rows - 1, .ColIndex("�ϼ�")) = Format(objMaster.�ϼ�, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("ʹ�����")) = objMaster.ʹ�����
            .TextMatrix(.Rows - 1, .ColIndex("ժҪ")) = objMaster.ժҪ
            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = IIf(objMaster.����ʱ�� = CDate(0), "", Format(objMaster.����ʱ��, "yy-MM-dd hh:mm"))
            .TextMatrix(.Rows - 1, .ColIndex("Key")) = objMaster.NO & "_" & IIf(objMaster.����ʱ�� = CDate(0), "0", Format(objMaster.����ʱ��, "yyMMddhhmmss"))
            .Rows = .Rows + 1
        End With
    Next
    If vsMaster.Rows > 2 Then
        Call vsMaster.RemoveItem(vsMaster.Rows - 1)
    End If
    vsMaster.Editable = flexEDNone
    vsMaster_RowColChange
    'vsMaster.Select vsMaster.Rows - (Rows - 1), 1
End Sub

Private Sub initVsList()
    Dim strHead As String
    
    strHead = "ҩƷ��Դ,650,1;ҩƷ���������,2800,1;���,1800,1;��;,550,1;����,750,7;��������,600,7;���㵥λ,450,4;����,750,7;���,1000,7;ԭʹ������,0,7;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsList)
    
    strHead = "ʹ��ʱ��,1200,1;ҩƷ��Դ,650,1;ҩƷ���������,2800,1;���,1800,1;��;,550,1;ʹ������,600,7;���㵥λ,450,4;����,750,7;���,1000,7;������,650,1;ժҪ,1000,1;Key,0,1"
    Call SetVsFlexGridHead(strHead, vsListUsed)
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With vsMaster
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Width = Me.ScaleWidth
    End With
    With picNS
        .Left = Me.ScaleLeft
        .Top = vsMaster.Top + vsMaster.Height
        .Width = Me.ScaleWidth
    End With
    With tbcSub
        .Left = Me.ScaleLeft
        .Top = picNS.Top + picNS.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsMaster.Height - picNS.Height - 60
    End With
    With vsList
        
    End With
End Sub

Private Sub picOne_Resize()
    With vsList
        .Top = picOne.Top
        .Width = picOne.Width
        .Height = picOne.Height
        .Left = picOne.Left
    End With
End Sub



Private Sub picTow_Resize()
    With vsListUsed
        .Top = picTow.Top
        .Width = picTow.Width
        .Height = picTow.Height
        .Left = picTow.Left
    End With
End Sub



Private Sub vsMaster_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsMaster_RowColChange()
    
    With vsMaster
    If .ColIndex("Key") > 0 Then
        Call Fill_List(.TextMatrix(.Row, .ColIndex("Key")))
    Else
        Call Fill_List("")
    End If
    End With
End Sub

Private Sub Fill_List(ByVal strNo As String)

    Dim strHead As String, objBIll As MediBill, str��Դ As String, str��; As String
    Dim objMaster As MediMaster, i As Integer
    Call initVsList

        If strNo = "" Or strNo = "NO" Then Exit Sub
        If mobjMasters Is Nothing Then Exit Sub
        Set objMaster = mobjMasters.Item(strNo)
        With vsList
            For i = 1 To objMaster.BillCount
                Set objBIll = objMaster.BillItem(i)
                If objBIll.���ϵ�� = 1 Then
                    Select Case objBIll.ִ�з���
                    Case 1
                        str��; = "��Һ"
                    Case 2
                        str��; = "ע��"
                    Case 3
                        str��; = "Ƥ��"
                    Case Else
                        str��; = "����"
                    End Select
                    If objBIll.ҩƷID = 0 And objBIll.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objBIll.ҩƷID <> 0 And objBIll.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objBIll.ҩƷID <> 0 And objBIll.ҽ��ID <> 0 Then
                        str��Դ = "ҽ��"
                    Else
                        str��Դ = "����"
                    End If
                    
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ��Դ")) = str��Դ
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������")) = objBIll.ҩƷ����
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = objBIll.���
                    .TextMatrix(.Rows - 1, .ColIndex("��;")) = str��;
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objBIll.����
                    .TextMatrix(.Rows - 1, .ColIndex("��������")) = objBIll.��������
                    .TextMatrix(.Rows - 1, .ColIndex("���㵥λ")) = objBIll.���㵥λ
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(objBIll.����, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(objBIll.���, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("ԭʹ������")) = objBIll.��������
                    .TextMatrix(.Rows - 1, .ColIndex("Key")) = objBIll.��� & "_" & objBIll.���ϵ�� & "_" & Format(objBIll.�Ǽ�ʱ��, "yyMMddhhmmss")
                    .Rows = .Rows + 1
                End If
            Next
            If .Rows > 2 Then
                .RemoveItem (.Rows - 1)
            End If
        End With
        

        '------ʹ�ü�¼
        
        If strNo = "" Or strNo = "NO" Then Exit Sub
        If mobjMasters Is Nothing Then Exit Sub
        Set objMaster = mobjMasters.Item(strNo)
        With vsListUsed
            For i = 1 To objMaster.BillCount
                Set objBIll = objMaster.BillItem(i)
                If objBIll.���ϵ�� = -1 Then
                    Select Case objBIll.ִ�з���
                    Case 1
                        str��; = "��Һ"
                    Case 2
                        str��; = "ע��"
                    Case 3
                        str��; = "Ƥ��"
                    Case Else
                        str��; = "����"
                    End Select
                    If objBIll.ҩƷID = 0 And objBIll.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objBIll.ҩƷID <> 0 And objBIll.ҽ��ID = 0 Then
                        str��Դ = "Ŀ¼��"
                    ElseIf objBIll.ҩƷID <> 0 And objBIll.ҽ��ID <> 0 Then
                        str��Դ = "ҽ��"
                    Else
                        str��Դ = "����"
                    End If
                    .TextMatrix(.Rows - 1, .ColIndex("ʹ��ʱ��")) = Format(objBIll.�Ǽ�ʱ��, "yy-MM-dd hh:mm")
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ��Դ")) = str��Դ
                    .TextMatrix(.Rows - 1, .ColIndex("ҩƷ���������")) = objBIll.ҩƷ����
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = objBIll.���
                    .TextMatrix(.Rows - 1, .ColIndex("��;")) = str��;
                    .TextMatrix(.Rows - 1, .ColIndex("ʹ������")) = objBIll.����
                    .TextMatrix(.Rows - 1, .ColIndex("���㵥λ")) = objBIll.���㵥λ
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(objBIll.����, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("���")) = Format(objBIll.���, "0.00")
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = objBIll.������
                    .TextMatrix(.Rows - 1, .ColIndex("ժҪ")) = objBIll.ʹ��ժҪ
                    .TextMatrix(.Rows - 1, .ColIndex("Key")) = objBIll.��� & "_" & objBIll.���ϵ�� & "_" & Format(objBIll.�Ǽ�ʱ��, "yyMMddhhmmss")
                    .Rows = .Rows + 1
                End If
            Next
            If .Rows > 2 Then
                .RemoveItem (.Rows - 1)
            End If
        End With
End Sub

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object)
    '������Ҫ���ʼ���������ϵĲ˵�
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '�ݴ�ҩƷ�Ĳ˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set mcbsMain = cbsMain
    Set mcbsMain.Icons = zlCommFun.GetPubIcons
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�Ĵ�ҩƷ(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Add, "����(&A)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "ʹ�õǼ�(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "�����Ǽ�(&U)")
    End With
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
'    If objMenu Is Nothing Then
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
'        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
'    End If
'    With objMenu.CommandBar.Controls
'        '���������ǰ��,�������
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Repertory, "����ѯ(&R)", 1)
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_AccountBook, "���̨��(&A)", 2)
'    End With
    
    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Set objBar = cbsMain(2)
    For Each objControl In objBar.Controls '�����ǰ������һ��Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            Set objControl = objBar.Controls(objControl.Index - 1): Exit For
        End If
    Next
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Add, "����", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Modify, "�޸�", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Delete, "ɾ��", objControl.Index + 1)
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "ʹ�õǼ�", objControl.Index + 1): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Leave_UndoPost, "�����Ǽ�", objControl.Index + 1)
    End With
'    cbsSub.ActiveMenuBar.Visible = False
    
    
    '-- vsLIst
'    Dim objPane As Pane, objPaneBase As Pane
'
'    Me.dkpSub.Options.UseSplitterTracker = False 'ʵʱ�϶�
'    Me.dkpSub.Options.ThemedFloatingFrames = True
'    Me.dkpSub.VisualTheme = ThemeOffice2003
'    Me.dkpSub.Options.HideClient = True
'
'    If dkpSub.FindPane(1) Is Nothing Then
'        Set objPaneBase = Me.dkpSub.CreatePane(1, 700, 200, DockTopOf, Nothing)
'        objPaneBase.Title = "�ݴ�ҩƷ"
'        objPaneBase.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If
'
'    If dkpSub.FindPane(2) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(2, 700, 200, DockBottomOf, objPaneBase)
'        objPaneBase.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If
    
    
    With tbcSub
        If tbcSub.ItemCount <= 0 Then
            With .PaintManager
                .Appearance = xtpTabAppearanceExcel
                .ClientFrame = xtpTabFrameSingleLine
                
                .Position = xtpTabPositionBottom 'ѡ��ڵײ�
                .BoldSelected = True
                .OneNoteColors = True
                .ShowIcons = True
            End With
    
            .InsertItem(0, "�ݴ���ϸ", picOne.hwnd, 0).Tag = "�ݴ���ϸ"
            .InsertItem(1, "ʹ�ü�¼", picTow.hwnd, 0).Tag = "ʹ�ü�¼"
            .Item(0).Selected = True
        End If
    End With
'    If dkpSub.FindPane(2) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(2, 700, 200, DockBottomOf, objPaneBase)
'        objPane.Title = "�ݴ���ϸ"
'        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'
'        vsList.Visible = False
'    End If
'
'    If dkpSub.FindPane(3) Is Nothing Then
'        Set objPane = Me.dkpSub.CreatePane(3, 700, 200, DockBottomOf, objPaneBase)
'        objPane.Title = "ʹ����ϸ"
'        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
'    End If


   '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    
    'Me.dkpSub.SetCommandBars cbsSub
    
End Sub

Public Sub zlRefresh()
    Call Fill_Master
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
    '
    Select Case Control.ID
        Case conMenu_Edit_Leave_Add
        '����
            Control.Enabled = mlng����ID <> 0 And (Not mblnEditList) And InStr(gstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0
        Case conMenu_Edit_Leave_Modify
        '�޸�
             
            If vsMaster.Row > 0 And vsMaster.ColIndex("NO") < vsMaster.Cols Then
                Control.Enabled = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("NO")) <> "" And (Not mblnEditList) _
                                  And vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("ʹ�����")) = "δ��" _
                                  And InStr(gstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0
            Else
                Control.Enabled = False
            End If
            If vsList.Editable = flexEDKbd Then Control.Enabled = False
        Case conMenu_Edit_Leave_Delete
        'ɾ��
            
            If vsMaster.Row > 0 And vsMaster.ColIndex("NO") < vsMaster.Cols Then
                Control.Enabled = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("NO")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Leave_Post
        '�Ǽ�
            If vsList.Row > 0 And vsList.ColIndex("Key") < vsList.Cols Then
                Control.Enabled = vsList.TextMatrix(vsList.Row, vsList.ColIndex("Key")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Leave_UndoPost
        '�����Ǽ�
            
            If tbcSub.Item(1).Selected Then
                If vsListUsed.Row > 0 And vsListUsed.ColIndex("Key") < vsListUsed.Cols Then
                    Control.Enabled = vsListUsed.TextMatrix(vsListUsed.Row, vsListUsed.ColIndex("Key")) <> "" And (Not mblnEditList) And InStr(gstrPrivs, ";" & "ҩƷ�Ĵ�" & ";") <> 0
                End If
            Else
                Control.Enabled = False
            End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl, ByVal frmMain As frmTransfusion)
    '��������ñ������ִ�й���
    Dim lngRow As Long, strMasterKey As String, lngMastRow As Long
    Select Case Control.ID
        Case conMenu_Edit_Leave_Add '����
            If mlng����ID <> 0 Then
                Set frmLeaveMediMana.pMediMaster = New MediMaster
                frmLeaveMediMana.pintType = 1
                With frmLeaveMediMana.pMediMaster
                    .�������� = mstr����
                    .���� = mstr����
                    .�Ա� = mstr�Ա�
                    .���� = mstr����
                    .����ID = mlng����ID
                    .����ID = mlng����ID
                    .�Һŵ� = mstr�Һŵ�
                    .����Ա = UserInfo.����
                    .�Ǽ�ʱ�� = zlDatabase.Currentdate
                End With
                frmLeaveMediMana.Show vbModal, Me
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                Call frmMain.ˢ��(lngRow)
                vsMaster.Row = lngMastRow
                vsMaster_RowColChange
            End If
        Case conMenu_Edit_Leave_Modify '�޸�
            With vsMaster
            If .TextMatrix(.Row, .ColIndex("Key")) <> "" Then
                frmLeaveMediMana.pintType = 2
                Set frmLeaveMediMana.pMediMaster = mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key")))
                With frmLeaveMediMana.pMediMaster
                    .�Һŵ� = mstr�Һŵ�
                    .����ID = mlng����ID
                    .����ID = mlng����ID
                    .�������� = mstr����
                    .����Ա = UserInfo.����
                End With
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                frmLeaveMediMana.Show vbModal, Me
                Call frmMain.ˢ��(lngRow)
                If vsMaster.Rows > lngMastRow Then
                    vsMaster.Row = lngMastRow
                End If
                vsMaster_RowColChange

            End If
            End With
        Case conMenu_Edit_Leave_Delete 'ɾ��
            With vsMaster
                lngRow = GetMainCurRowIndex(frmMain)
                Call mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key"))).DeleteBill(0)
                Call frmMain.ˢ��(lngRow)
            End With
        Case conMenu_Edit_Leave_Post 'ʹ�õǼ�
            With vsMaster
            If .TextMatrix(.Row, .ColIndex("Key")) <> "" Then
                frmLeaveMediMana.pintType = 3
                Set frmLeaveMediMana.pMediMaster = mobjMasters.Item(.TextMatrix(.Row, .ColIndex("Key")))
                With frmLeaveMediMana.pMediMaster
                    .�Һŵ� = mstr�Һŵ�
                    .����ID = mlng����ID
                    .����ID = mlng����ID
                    .�������� = mstr����
                    .����Ա = UserInfo.����
                End With
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                frmLeaveMediMana.Show vbModal, Me
                Call frmMain.ˢ��(lngRow)
                If vsMaster.Rows > lngMastRow Then
                    vsMaster.Row = lngMastRow
                End If
                vsMaster_RowColChange
            End If
            End With
        Case conMenu_Edit_Leave_UndoPost '����ʹ�õǼ�
            strMasterKey = vsMaster.TextMatrix(vsMaster.Row, vsMaster.ColIndex("Key"))
            If strMasterKey <> "" Then
                lngRow = GetMainCurRowIndex(frmMain)
                lngMastRow = vsMaster.Row
                With vsListUsed
                    If .TextMatrix(.Row, .ColIndex("ҩƷ��Դ")) <> "ҽ��" Then
                        If MsgBox("��ȷ�ϣ��Ƿ�����" & .TextMatrix(.Row, .ColIndex("ҩƷ���������")) & "��������Ϊ ��" & .TextMatrix(.Row, .ColIndex("ʹ������")) & "����ʹ�ü�¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                            Call mobjMasters.Item(strMasterKey).UndoUse(.TextMatrix(.Row, .ColIndex("Key")))
                            Call frmMain.ˢ��(lngRow)
                            If vsMaster.Rows > lngMastRow Then
                                vsMaster.Row = lngMastRow
                            End If
                            vsMaster_RowColChange
                        End If
                    Else
                        MsgBox "ҽ����¼�����ڴ˳���ʹ�ã�", vbInformation, gstrSysName
                    End If
                End With
            End If
        Case conMenu_Edit_Leave_Repertory '����ѯ
        Case conMenu_Edit_Leave_AccountBook '̨��
        
    End Select
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    '����������ñ�����ĵ����˵�
End Sub

Private Function GetMainCurRowIndex(ByVal frmMain As frmTransfusion) As Long
    'ȡ������ĵ�ǰ������RPT�ؼ��е�ѡ���е�index
    Dim objRpt As ReportControl
    
    If frmMain.tbcList.Selected.Tag = "δ�ӵ�" Then
        Set objRpt = frmMain.rptQueue0
    ElseIf frmMain.tbcList.Selected.Tag = "����Һ" Then
        Set objRpt = frmMain.rptQueue1
    ElseIf frmMain.tbcList.Selected.Tag = "������" Then
        Set objRpt = frmMain.rptQueue5
    ElseIf frmMain.tbcList.Selected.Tag = "��ִ��" Then
        Set objRpt = frmMain.rptQueue6
    ElseIf frmMain.tbcList.Selected.Tag = "ִ����" Then
        Set objRpt = frmMain.rptQueue7
    ElseIf frmMain.tbcList.Selected.Tag = "�ѽ���" Then
        Set objRpt = frmMain.rptPati
    End If
    GetMainCurRowIndex = objRpt.SelectedRows(0).Index
End Function

