VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMachineManage 
   Caption         =   "ҩƷ�豸����"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmMachineManage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSecond 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   2415
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfSecond 
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1935
         _cx             =   3413
         _cy             =   1720
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
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
         Cols            =   10
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
   End
   Begin VB.PictureBox picPrimary 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   2415
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   2415
      Begin VSFlex8Ctl.VSFlexGrid vsfPrimary 
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1935
         _cx             =   3413
         _cy             =   1720
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
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
         Cols            =   10
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
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6570
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMachineManage.frx":038A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   360
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMachineManage.frx":0C1C
      Left            =   840
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMachineManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_BILL As String = _
        "���,,3,2000|ID,,0,0|����,,3,3000|�ӿ�����,,3,1500|��������,,0,0,d|ͣ������,,3,1000,d|��ע,,3,3500"
Private Const MSTR_DETAIL As String = _
        "�ⷿ����,,3,2000|�ⷿ����,,3,3000|ҩƷ����,,3,7000"

Private mblnShow As Boolean                     '��ʾ״̬��Load�¼���Ĺ��̴���
Private mfrmOwner As Form
Private WithEvents mclsPrimary As clsVSFlexGridEx
Attribute mclsPrimary.VB_VarHelpID = -1
Private WithEvents mclsSecond As clsVSFlexGridEx
Attribute mclsSecond.VB_VarHelpID = -1

Public Sub ShowMe(ByVal frmOwner As Form)
    Set mfrmOwner = frmOwner
    Show , frmOwner
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
    Dim objPrint As Object
    
    Select Case Control.ID
    Case enuMenus.����
        If frmMachineEdit.ShowMe(Me, enuEditState.����) Then
            Call FuncCall(enuMenus.ˢ��)
        End If
    
    Case enuMenus.�޸�
        If frmMachineEdit.ShowMe(Me, enuEditState.�޸�, Val(vsfPrimary.TextMatrix(vsfPrimary.Row, vsfPrimary.ColIndex("ID")))) Then
            Call FuncCall(enuMenus.ˢ��)
        End If
    
    Case enuMenus.ɾ��
        Call DeleteINF
        
    Case enuMenus.����
        Call ChangeState(1)
        Call FuncCall(enuMenus.ˢ��)
    
    Case enuMenus.ͣ��
        Call ChangeState(0)
        Call FuncCall(enuMenus.ˢ��)
    
    Case enuMenus.ˢ��
        Screen.MousePointer = vbHourglass
        Call FillData(0)
        Call FillData(1)
        Screen.MousePointer = vbDefault
        
    Case enuMenus.�˳�
        Unload Me
    
    Case enuMenus.��ӡ����
        If gobjZLPrint Is Nothing Then Exit Sub
        Call gobjZLPrint.zlPrintSet
    
    Case enuMenus.��ӡԤ��, enuMenus.��ӡ, enuMenus.���Excel
        If TypeName(Me.ActiveControl) = "VSFlexGrid" Then
            If gobjZLPrint Is Nothing Then Exit Sub
            
            Set objPrint = CreateObject("zl9PrintMode.zlPrint1Grd")
            If UCase(Me.ActiveControl.Name) = "VSFPRIMARY" Then
                Set objPrint.Body = vsfPrimary
            Else
                Set objPrint.Body = vsfSecond
            End If
            
            On Error GoTo hErr
            If Control.ID = enuMenus.��ӡԤ�� Then
                gobjZLPrint.zlPrintOrView1Grd objPrint, 0
            ElseIf Control.ID = enuMenus.��ӡ Then
                gobjZLPrint.zlPrintOrView1Grd objPrint, 1
            Else
                gobjZLPrint.zlPrintOrView1Grd objPrint, 3
            End If
            On Error GoTo 0
        End If
        
    Case enuMenus.��׼��ť
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
        
    Case enuMenus.�ı���ǩ
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
        
    Case enuMenus.��ͼ��
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
        
    Case enuMenus.״̬��
        stbMain.Visible = Not Control.Checked
        cbsMain.RecalcLayout
    
    Case enuMenus.��������
        Call gobjComLib.ShowHelp(App.ProductName, Me.hwnd, Me.Name)
        
    Case enuMenus.������ҳ
        Call gobjComLib.zlHomePage(Me.hwnd)
        
    Case enuMenus.������̳
        Call gobjComLib.zlWebForum(Me.hwnd)
        
    Case enuMenus.���ͷ���
        Call gobjComLib.zlMailTo(Me.hwnd)
    End Select
    Exit Sub
    
hErr:
    Call gobjComLib.ErrCenter
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbMain.Visible Then
        Bottom = stbMain.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTemp As String
    Dim objControl As CommandBarControl
    
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.ID
    Case enuMenus.�޸�
        Control.Enabled = vsfPrimary.Rows > 1
    Case enuMenus.ɾ��
        Control.Enabled = vsfPrimary.Rows > 1
    Case enuMenus.����
        With vsfPrimary
            Control.Enabled = .Rows > 1 And _
                            (Trim(.TextMatrix(.Row, .ColIndex("ͣ������"))) <> "" Or _
                             Trim(.TextMatrix(.Row, .ColIndex("��������"))) = "")
        End With
    Case enuMenus.ͣ��
        With vsfPrimary
            Control.Enabled = .Rows > 1 And _
                            Not (Trim(.TextMatrix(.Row, .ColIndex("ͣ������"))) <> "" Or _
                                 Trim(.TextMatrix(.Row, .ColIndex("��������"))) = "")
        End With
    Case enuMenus.��ʾ
        Control.Enabled = Me.Visible = False
    Case enuMenus.����
        Control.Enabled = Me.Visible
    Case enuMenus.��׼��ť
        Control.Checked = Me.cbsMain(2).Visible
    Case enuMenus.�ı���ǩ
        Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
    Case enuMenus.��ͼ��
        Control.Checked = cbsMain.Options.LargeIcons
    Case enuMenus.״̬��
        Control.Checked = Me.stbMain.Visible
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPrimary.hwnd
    Case 2
        Item.Handle = picSecond.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnShow Then
        Screen.MousePointer = vbHourglass
        Me.Visible = False
        
        Call InitDockPane
        Call InitCommandbars
        Call InitVSF
        
        gobjComLib.RestoreWinState Me, App.EXEName
        
        Call FillData
        Call FillData(1)
    
        cbsMain.RecalcLayout
        mblnShow = False
        
        Me.Visible = True
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    '��ʼ������
    Set mclsPrimary = New clsVSFlexGridEx
    Set mclsSecond = New clsVSFlexGridEx
    
    mblnShow = True         '���з����
End Sub

Private Sub InitDockPane()
    Dim panTop As Pane, panBottom As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
        
        Set panTop = .CreatePane(1, 0, Me.ScaleY(Me.Height, vbTwips, vbPixels) \ 2, DockTopOf)
        With panTop
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = "�ӿ���Ϣ"
        End With
        
        Set panBottom = .CreatePane(2, 0, Me.ScaleY(Me.Height, vbTwips, vbPixels) \ 2, DockBottomOf)
        With panBottom
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = "�ⷿ�������Ϣ"
        End With
    End With
End Sub

Private Sub InitCommandbars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        .EnableCustomization False
        Set .Icons = mfrmOwner.imgMain.Icons
        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
'    picLine01_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
'    picLine02.BackColor = picLine01_S.BackColor
    
    '�ļ�
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�ļ�, "�ļ�(&F)", -1, False)
    With cbpTmp
        .ID = enuMenus.�ļ�
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡ����, "��ӡ����(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡԤ��, "��ӡԤ��(&V)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡ, "��ӡ")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.���Excel, "�����&Excel...")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�")
        cbcTmp.BeginGroup = True
    End With
    
    '����
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�༭, "�༭(&E)", -1, False)
    With cbpTmp
        .ID = enuMenus.����
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����, "����")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�޸�, "�޸�")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ɾ��, "ɾ��")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����, "�ӿ�����(&S)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ͣ��, "�ӿ�ֹͣ(&P)")
    End With
    
    '�鿴
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�鿴, "�鿴(&V)", -1, False)
    With cbpTmp
        .ID = enuMenus.�鿴
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.������, "������(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��׼��ť, "��׼��ť(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ı���ǩ, "�ı���ǩ(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ͼ��, "��ͼ��(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.״̬��, "״̬��(&S)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��")
        cbcTmp.BeginGroup = True
    End With
    
    '����
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.����, "����(&H)", -1, False)
    With cbpTmp
        .ID = enuMenus.����
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB�ϵ�����, "&WEB�ϵ�����")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������ҳ, "������ҳ(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̳, "������̳(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.���ͷ���, "���ͷ���(&K)")
    End With
    
    '�˵���Ŀ����
    With cbsMain.KeyBindings
        .Add 8, vbKeyP, enuMenus.��ӡ
        .Add 8, vbKeyX, enuMenus.�˳�
        .Add 8, vbKeyA, enuMenus.����
        .Add 8, vbKeyE, enuMenus.�޸�
        .Add 8, vbKeyD, enuMenus.ɾ��
        .Add 0, vbKeyF1, enuMenus.��������
        .Add 0, vbKeyF5, enuMenus.ˢ��
    End With
    
    '���幤����
    Set cbrTmp = cbsMain.Add("������", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.����, "����")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�޸�, "�޸�")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ɾ��, "ɾ��")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.����, "�ӿ�����")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ͣ��, "�ӿ�ͣ��")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�")
        cbcTmp.BeginGroup = True
    End With
    
    '��ͼ�꣬���ı��İ�ť���
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.Type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub InitVSF()
    With mclsPrimary
        .Bunding = vsfPrimary
        .Init
        .Head = MSTR_BILL
        .ColsReadonly = ""
        .Editable = EM_Display
        .Repaint RT_Columns
    End With
    With vsfPrimary
        .RowHeight(0) = 350
        .ExplorerBar = flexExSort
    End With
    
    With mclsSecond
        .Bunding = vsfSecond
        .Init
        .Head = MSTR_DETAIL
        .ColsReadonly = ""
        .Editable = EM_Display
        .Repaint RT_Columns
    End With
    With vsfSecond
        .RowHeight(0) = 350
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Width < 6000 Then Width = 6000
    If Height < 4000 Then Height = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gobjComLib.SaveWinState Me, App.EXEName
    
    Set mclsSecond = Nothing
    Set mclsPrimary = Nothing
End Sub

Private Sub mclsPrimary_EventFillData(ByVal Row As Long, ByVal Col As Long)
    Dim arrType As Variant
    Dim intIndex As Integer
    
    If Col = mclsPrimary.Bunding.ColIndex("�ӿ�����") And Row > 0 Then
        arrType = Split(GSTR_TYPE, "|")
        intIndex = Val(mclsPrimary.Bunding.TextMatrix(Row, Col))
        If intIndex > 0 Then
            On Error GoTo hErr
            mclsPrimary.Bunding.TextMatrix(Row, Col) = arrType(intIndex - 1)    '��д�ӿ����͵�Ԫ���磺5-YUYAMA
        End If
    End If
    
    Exit Sub
    
hErr:
    MsgBox "�豸���Ͳ���ȷ��", vbInformation, GSTR_MSG
End Sub

Private Sub picPrimary_Resize()
    On Error Resume Next
    With vsfPrimary
        .Top = 0
        .Left = 0
        .Width = picPrimary.ScaleWidth
        .Height = picPrimary.ScaleHeight
    End With
End Sub

Private Sub picSecond_Resize()
    On Error Resume Next
    With vsfSecond
        .Top = 0
        .Left = 0
        .Width = picSecond.ScaleWidth
        .Height = picSecond.ScaleHeight
    End With
End Sub

Private Sub FuncCall(ByVal lngMenuID As Long)
    Dim objControl As CommandBarControl
    
    Set objControl = cbsMain.ActiveMenuBar.FindControl(, lngMenuID, , True)
    If Not objControl Is Nothing Then
        If objControl.Enabled And objControl.Visible Then Call cbsMain_Execute(objControl)
    End If
End Sub

Private Sub vsfPrimary_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible = False Then Exit Sub
    If OldRow <> NewRow Then
        Call FillData(1)
    End If
End Sub

Private Sub vsfPrimary_DblClick()
    Call FuncCall(enuMenus.�޸�)
End Sub

Private Sub DeleteINF()
'���ܣ�ɾ��ע��Ľӿڼ�¼

    Dim strName As String
    Dim lngID As Long
    Dim rsSQL As ADODB.Recordset
    
    With vsfPrimary
        lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        If lngID <= 0 Then
            MsgBox "��ǰ�����쳣��", vbInformation, GSTR_MSG
            Exit Sub
        End If
    End With
    
    On Error GoTo hErr
    
    gstrSQL = "Select '��'|| ��� ||'��'||���� ���� From ҩƷ�豸�ӿ� Where ID = [1] "
    Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ�豸�ӿ�", lngID)
    If rsSQL.EOF Then
        rsSQL.Close
        MsgBox "��ǰ�ӿ�δ�ҵ���", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    strName = rsSQL!����
    rsSQL.Close

    If MsgBox(mdlMain.FormatString("ȷ��Ҫɾ����[1]���ӿڣ�", strName), vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        gstrSQL = mdlMain.FormatString("ZL_ҩƷ�豸�ӿ�_DELETE([1])", lngID)
        Call gobjComLib.zlDatabase.ExecuteProcedure(gstrSQL, "")
    End If
    
    Call FuncCall(enuMenus.ˢ��)
    Exit Sub

hErr:
    If gstrSQL Like "Select *" Then
        If gobjComLib.ErrCenter = 1 Then Resume
    Else
        Call gobjComLib.ErrCenter
    End If
End Sub

Private Sub FillData(Optional ByVal bytType As Byte)
'���ܣ��������������
'������
'  bytType��0-������1-������

    Dim rsSQL As ADODB.Recordset
    Dim lngID As Long
    
    On Error GoTo hErr
    
    If bytType = 0 Then
        gstrSQL = "Select ID, ���, ����, ���� �ӿ�����, ��������, ͣ������, ��ע " & vbNewLine & _
                  "From ҩƷ�豸�ӿ� " & vbNewLine & _
                  "Order By ID"
        Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "����ҩƷ�豸�ӿ�")
        mclsPrimary.Recordset = rsSQL
        mclsPrimary.Repaint RT_Rows
        rsSQL.Close
    Else
        With vsfPrimary
            lngID = Val(.TextMatrix(.Row, .ColIndex("ID")))
        End With
        
        gstrSQL = _
            "Select �ⷿid, �ⷿ����, �ⷿ����, f_List2str(Cast(Collect(�������� Order By ���ͱ���) As t_Strlist), '��') ҩƷ����" & vbNewLine & _
            "From (Select a.���� ���ͱ���, a.���� ��������, d.�ⷿid, b.���� �ⷿ����, b.���� �ⷿ����" & vbNewLine & _
            "      From ҩƷ���� A, ���ű� B, ҩƷ�豸�ӿ� C," & vbNewLine & _
            "        Xmltable('//root/bm' Passing c.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���ͱ��� Varchar2(20) Path 'jxbm') D" & vbNewLine & _
            "      Where d.�ⷿid = b.Id(+) And d.���ͱ��� = a.����(+) And c.Id = [1] )" & vbNewLine & _
            "Group By �ⷿid, �ⷿ����, �ⷿ����"
        Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "�ӿڵĿⷿ�����", lngID)
        mclsSecond.Recordset = rsSQL
        mclsSecond.Repaint RT_Rows
        rsSQL.Close
    End If
    
    Exit Sub
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub ChangeState(ByVal bytType As Byte)
'���ܣ��޸Ľӿڵ�״̬
'������
'  bytType��1-���ã�0-ͣ��

    Dim lngID As Long
    
    On Error GoTo hErr
    
    lngID = Val(vsfPrimary.TextMatrix(vsfPrimary.Row, vsfPrimary.ColIndex("ID")))
    
    gstrSQL = mdlMain.FormatString("ZL_ҩƷ�豸�ӿ�_STATE([1], [2])", lngID, bytType)
    Call gobjComLib.zlDatabase.ExecuteProcedure(gstrSQL, "")
    
    Exit Sub
    
hErr:
    Call gobjComLib.ErrCenter
End Sub
