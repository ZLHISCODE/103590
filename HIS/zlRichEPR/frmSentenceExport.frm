VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmSentenceExport 
   AutoRedraw      =   -1  'True
   Caption         =   "�ʾ䵼��"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmSentenceExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   9615
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picNowList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   480
      ScaleHeight     =   3735
      ScaleWidth      =   2775
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmdYes 
         Caption         =   " ȷ ��"
         Height          =   380
         Left            =   840
         TabIndex        =   6
         Top             =   3300
         Width           =   1100
      End
      Begin MSComctlLib.TreeView tvwNowList 
         Height          =   3255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5741
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imgClass"
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3600
      ScaleHeight     =   1335
      ScaleWidth      =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   2160
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   975
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   1575
         _cx             =   2778
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
         BackColorBkg    =   16777215
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
         ScrollBars      =   2
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
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1931
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgClass"
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6915
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
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
   Begin MSComctlLib.ImageList imgClass 
      Left            =   2160
      Top             =   120
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
            Picture         =   "frmSentenceExport.frx":6852
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":6DEC
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7386
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7920
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":7EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceExport.frx":8454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmSentenceExport.frx":8D2E
   End
   Begin XtremeCommandBars.ImageManager imgTools 
      Left            =   600
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSentenceExport.frx":8DCB
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSentenceExport.frx":1087B
      Left            =   1080
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSentenceExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mCol
    Range = 0: Choose: Num: pName: Depart: personnel: clasID: ID: Class
End Enum
Public Event zlRefParentTree()      'ˢ�´ʾ��б�
Private mlngFileID As Long, mlngWordId As Long
Private mtion As Collection         '�洢����Ҫ������ߵ����Ĵʾ�id
Private mblnDifference As Boolean   '��ʾ���뵼����Ϊ���ʾ������Ϊ�ٱ�ʾ����
Private msrtXmlPathName As String   '��¼xml�ļ���·��
Private mstrHigher As String        '�����ϼ���������
Private objBar As CommandBar        '������
Private mColClass As Collection     '��¼����

Private oDoc As DOMDocument         'xml�ĵ�
Private oRoot  As IXMLDOMElement    '���ڵ�
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTag As String
    Select Case Control.ID
        Case 10:
            Call ImportOrExport
        Case 11:
            Call ImportOrExport
        Case 15:
            On Error GoTo ErrHandle
            With dlgThis
                On Error Resume Next
                .DialogTitle = "���ļ�"
                .Filter = "*.ZIP|*.zip"
                .flags = &H80000 + &H1000 + &H200000 + &H800
                .CancelError = True
                .InitDir = "C:\APPSOFT"
                .ShowOpen
            If Err.Number = 32755 Then Err.Clear: Exit Sub
            '���н�ѹ������
            Dim strFilePath As String '��ʱ�ļ�����ѹ�����xml�ļ����رմ���ʱɾ������Ϊ��������ʱ��Ҳ��Ҫ��Ϊ��Լʱ�䣩
            strFilePath = zlFilesUnZip(.Filename)
            Set oDoc = New DOMDocument
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile strFilePath, True
            '����������κ�Ԫ�أ����˳�
            If oDoc.documentElement Is Nothing Then
                MsgBox "���ļ����ݸ�ʽ����ȷ���ѱ��𻵣�", vbInformation, gstrSysName: Exit Sub
            End If
            If oDoc Is Nothing Then
                Set oDoc = New DOMDocument
            End If
            msrtXmlPathName = strFilePath
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile msrtXmlPathName, True
            Set oRoot = oDoc.selectSingleNode("Document")       'oRoot��Ϊ���ڵ�
            '����������κ�Ԫ�أ����˳�
            If Not oDoc.documentElement Is Nothing Then
                Call zlXmlTree
            End If
            End With
        Case 16:
            Unload Me
        Case 17:
            CheckAllOrClearAll (True)
        Case 18:
            CheckAllOrClearAll (False)
        Case 22:
            strTag = Me.tvwClass.SelectedItem.Tag
            Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & vbCrLf & Split(strTag, vbCrLf)(3)
            Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 0)
        Case 23, 24:
            Me.tvwClass.SelectedItem.Checked = Not Me.tvwClass.SelectedItem.Checked
        Case 26:
            strTag = Me.tvwClass.SelectedItem.Tag
            Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Split(strTag, vbCrLf)(2) & vbCrLf
            Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 0)
    End Select
    Exit Sub
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "ȡ������", vbInformation, "��ʾ"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub cbsThis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    '�����νṹ���ж�λ
    If CommandBar.Title = "�ϼ�����" Then
        If Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) <> "" Then
            Me.tvwNowList.Nodes("_" & Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3)).Selected = True
        Else
            If Me.tvwNowList.Nodes.Count > 0 Then Me.tvwNowList.Nodes(1).Selected = True
        End If
        Me.tvwNowList.Tag = 1
        
    ElseIf CommandBar.Title = "��ӵ�ָ������" Then
        If Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) <> "" Then
            Me.tvwNowList.Nodes("_" & Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2)).Selected = True
        Else
           If Me.tvwNowList.Nodes.Count > 0 Then Me.tvwNowList.Nodes(1).Selected = True
        End If
        Me.tvwNowList.Tag = 0
        
    End If
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    Select Case Control.ID
        Case 2:
            Control.Enabled = Me.tvwClass.Nodes.Count > 0
        Case 10
            Control.Visible = mblnDifference
        Case 11
            Control.Visible = Not mblnDifference
        Case 15
            Control.Visible = Not mblnDifference
        Case 21:
            If Me.tvwNowList.Nodes.Count > 0 Then
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) = "" And Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) = ""
            Else
                Control.Enabled = False
            End If
        Case 22:
            If Me.tvwClass.SelectedItem.Parent Is Nothing Then
                Control.Enabled = False
            Else
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) <> ""
            End If
        Case 23:
            Control.Enabled = Not Me.tvwClass.SelectedItem.Checked
        Case 24:
            Control.Enabled = Me.tvwClass.SelectedItem.Checked
        Case 25:
            If Me.tvwNowList.Nodes.Count > 0 Then
                Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) = "" And Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(2) = ""
            Else
                Control.Enabled = False
            End If
        Case 26:
            Control.Enabled = Split(Me.tvwClass.SelectedItem.Tag, vbCrLf)(3) <> ""
    End Select
End Sub
Private Sub ImportOrExport()
    On Error GoTo ErrHandle
    Dim strFileXML As String
    If mblnDifference Then
        With dlgThis
                On Error Resume Next
                .DialogTitle = "�����ļ�"
                .Filter = "*.ZIP|*.zip"
                .flags = &H200000 + &H2000 + &H2 + &H800
                .CancelError = True
                .InitDir = "C:\APPSOFT"
                .Filename = "Sentence.zip"
                .ShowSave
                If Err.Number = 32755 Then Err.Clear: Exit Sub
                zlCommfun.ShowFlash "���Ժ����ڵ���..."
                Screen.MousePointer = vbHourglass
                strFileXML = "Sentence.xml"
                Call ToXml(strFileXML)
                '����ѹ������
                Call zlFilesZip(strFileXML, .Filename)
                zlCommfun.StopFlash
                Screen.MousePointer = vbDefault
                MsgBox "�����ɹ�,�ļ���ַ��" & .Filename, vbOKOnly, "��ʾ"
        End With
    Else
        Dim i As Integer
        For i = 1 To Me.tvwClass.Nodes.Count
            If Me.tvwClass.Nodes(i).Checked = True Then Exit For
        Next
        If i = Me.tvwClass.Nodes.Count + 1 Then
            MsgBox "û��ѡ��Ҫ����ķ���", vbOKOnly, "��ʾ": Exit Sub
        Else
            zlCommfun.ShowFlash "���Ժ����ڵ���..."
            Screen.MousePointer = vbHourglass
            
             Call ImportXMLFile
            
            zlCommfun.StopFlash
            Screen.MousePointer = vbDefault
            MsgBox "����ɹ�!", vbOKOnly, "��ʾ"
            RaiseEvent zlRefParentTree
        End If
    End If
    Unload Me
    Exit Sub
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "ȡ������", vbInformation, "��ʾ"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Sub

Private Sub cmdYes_Click()
    Dim strTag As String
    strTag = Me.tvwClass.SelectedItem.Tag
    If Me.tvwNowList.Tag = "" Then
    ElseIf Me.tvwNowList.Tag = 1 Then
        Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Split(strTag, vbCrLf)(3) & vbCrLf & Mid(Me.tvwNowList.SelectedItem.Key, 2)
        Me.tvwClass.SelectedItem.ForeColor = RGB(128, 0, 128)
    ElseIf Me.tvwNowList.Tag = 0 Then
        Me.tvwClass.SelectedItem.Tag = Split(strTag, vbCrLf)(0) & vbCrLf & Split(strTag, vbCrLf)(1) & vbCrLf & Mid(Me.tvwNowList.SelectedItem.Key, 2) & vbCrLf & Split(strTag, vbCrLf)(3)
        Me.tvwClass.SelectedItem.Text = Me.tvwNowList.SelectedItem.Text
        Me.tvwClass.SelectedItem.ForeColor = RGB(0, 0, 255)
    End If
    Me.tvwNowList.Tag = ""
    Me.cbsThis.ClosePopups
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = Me.tvwClass.hwnd
        Case 2
            Item.Handle = Me.picList.hwnd
        Case 3
            Item.Handle = Me.rtbText.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim cbpPopup As CommandBarPopup
    Dim cbpNew As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbpCustom As CommandBarControlCustom
    Set objBar = Me.cbsThis.Add("Tools", xtpBarTop)
    objBar.ContextMenuPresent = False           '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False           '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagStretched
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = Me.imgTools.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True                 '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbsThis.EnableCustomization False
    Me.cbsThis.ActiveMenuBar.Visible = False
    With objBar.Controls
        Set cbrControl = .Add(xtpControlButton, 15, "��"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 10, "����"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 11, "����"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 17, "ȫѡ"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, 18, "ȫ��"): cbrControl.Style = xtpButtonIconAndCaption
        Set cbpNew = .Add(xtpControlPopup, 21, "ָ������"): cbpNew.CommandBar.Title = "��ӵ�ָ������"
        Set cbpCustom = cbpNew.CommandBar.Controls.Add(xtpControlCustom, 211, "��ӵ�ָ�������б�"): cbpCustom.Handle = Me.picNowList.hwnd
        cbpNew.ID = 21: cbpNew.Visible = False: cbpNew.BeginGroup = True
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, 22, "ȡ��ָ��"): cbrControl.Style = xtpButtonIconAndCaption: cbrControl.Visible = False
        Set cbpNew = .Add(xtpControlPopup, 25, "ָ���ϼ�"): cbpNew.CommandBar.Title = "�ϼ�����"
        Set cbpCustom = cbpNew.CommandBar.Controls.Add(xtpControlCustom, 251, "�ϼ������б�"): cbpCustom.Handle = Me.picNowList.hwnd
        cbpNew.ID = 25: cbpNew.Visible = False: cbpNew.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, 26, "ȡ���ϼ�"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption: cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, 16, "�˳�"): cbrControl.BeginGroup = True: cbrControl.Style = xtpButtonIconAndCaption
    End With
    
       '���ô��岼��
    dkpMan.SetCommandBars Me.cbsThis
    dkpMan.Options.ThemedFloatingFrames = True
    Dim panThis As Pane, panSub As Pane, panOper As Pane
    
    Set panThis = dkpMan.CreatePane(1, 400, 800, DockLeftOf, Nothing)
    panThis.Title = "�ʾ����"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    
    Set panThis = dkpMan.CreatePane(2, 1100, 300, DockRightOf, panThis)
    panThis.Title = "�ʾ��б�"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    

    Set panSub = dkpMan.CreatePane(3, 1100, 500, DockBottomOf, panThis)
    panSub.Title = "�ʾ�����"
    panSub.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    With vsfList
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, mCol.Choose) = "ѡ��"
        .TextMatrix(0, mCol.Class) = "����"
        .TextMatrix(0, mCol.Depart) = "����"
        .TextMatrix(0, mCol.Num) = "���"
        .TextMatrix(0, mCol.pName) = "����"
        .TextMatrix(0, mCol.personnel) = "��Ա"
        .FixedCols = 0
        .ColDataType(mCol.Choose) = flexDTBoolean
        .SelectionMode = flexSelectionByRow
    End With
    
End Sub
Public Function ShowMe(blnDifference As Boolean, frmParent As Object) As Boolean
'---------------------------------------------------------
    '��ʾ�������뵼��������
    'blnDifference ���Ϊ���ʾ������Ϊ����Ϊ����
    'frmParent �ϼ����壬����ģ̬����ʾ
    '����ֵ ��Ϊ�ɹ� ��Ϊʧ��
'---------------------------------------------------------
    mblnDifference = blnDifference
    On Error GoTo ErrHandle
    If blnDifference Then '����
         Me.Caption = "�����ʾ䵼��"
         If zlRefTree = -1 Then
            MsgBox "���ݿⲻ���ڴʾ������Ϣ", vbOKOnly, "��ʾ"
            Exit Function
         End If
         Me.cbsThis.ActiveMenuBar.Visible = False
         objBar.Visible = True
         
    Else '����
        Me.Caption = "�����ʾ䵼��"
        On Error Resume Next
        With dlgThis
            .DialogTitle = "���ļ�"
            .Filter = "*.ZIP|*.zip"
            .flags = &H80000 + &H1000 + &H200000 + &H800
            .CancelError = True
            .InitDir = "C:\APPSOFT"
            .ShowOpen
            If Err.Number = 32755 Then Err.Clear: Exit Function
            '���н�ѹ������
            Dim strFilePath As String '��ʱ�ļ�����ѹ�����xml�ļ����رմ���ʱɾ������Ϊ��������ʱ��Ҳ��Ҫ��Ϊ��Լʱ�䣩
            strFilePath = zlFilesUnZip(.Filename)
            Set oDoc = New DOMDocument
            oDoc.Load strFilePath
            If gobjFSO.FileExists(strFilePath) Then gobjFSO.DeleteFile strFilePath, True
            '����������κ�Ԫ�أ����˳�
            If oDoc.documentElement Is Nothing Then
                MsgBox "���ļ����ݸ�ʽ����ȷ���ѱ��𻵣�", vbInformation, gstrSysName: Exit Function
            End If
            Set oRoot = oDoc.selectSingleNode("Document")       'oRoot��Ϊ���ڵ�
            Call zlXmlTree
            Call picNowList_Resize
        End With
    End If
    ShowMe = True
    Set mtion = New Collection
    Me.Show 0, frmParent
    Exit Function
ErrHandle:
    If Err.Number = 32755 Then
        MsgBox "ȡ������", vbInformation, "��ʾ"
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    ShowMe = False
End Function
Private Function zlRefTree(Optional lngID As Long) As Long
    '���ܣ�ˢ��װ��ָ������Ĳ����ļ��嵥������λ��ָ�����ļ���
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    
    gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��, ��Χ" & vbNewLine & _
            "From �����ʾ����" & vbNewLine & _
            "Start With �ϼ�id Is Null" & vbNewLine & _
            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
            "Order By Level, ����"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rsTemp.RecordCount < 1 Then
        zlRefTree = -1
        Exit Function
    End If
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !���� & "-" & !����, "close")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, !���� & "-" & !����, "close")
                End If
            objNode.Tag = !˵�� & vbCrLf & !��Χ: objNode.Sorted = True: objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    If Me.tvwClass.Nodes.Count > 0 Then
            If lngID <> 0 Then
                Me.tvwClass.Nodes("_" & lngID).Selected = True
            Else
                Me.tvwClass.Nodes(1).Selected = True
            End If
            If Me.tvwClass.SelectedItem.Children > 0 Then Me.tvwClass.SelectedItem.Expanded = True
            Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
    Else
        Call tvwClass_NodeClick(Nothing)
    End If
    zlRefTree = Me.tvwClass.Nodes.Count
    Exit Function

Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefTree = Me.tvwClass.Nodes.Count
End Function
Private Function zlXmlTree() As Integer
    Dim oNode As IXMLDOMNode            '���ڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim i As Long
    '�ж��Ƿ��Ǵʾ�xml
    If oRoot.getAttribute("EditType") <> "MedicalWords" Then
        MsgBox "��xml�ļ����ǲ����ʾ䵼����xml,�޷��Ӵ�xml���벡���ʾ�", vbOKOnly, "��ʾ"
        Exit Function
    End If
    '��ȡ������Ϣ
    On Error Resume Next
    mstrHigher = ""
    Me.tvwClass.Nodes.Clear
    Call zlTree(oRoot)
    zlXmlTree = tvwClass.Nodes.Count
    
    '��ʼ��Ŀ���Ĳ����ʾ�
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As MSComctlLib.Node
    gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��, ��Χ" & vbNewLine & _
            "From �����ʾ����" & vbNewLine & _
            "Start With �ϼ�id Is Null" & vbNewLine & _
            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
            "Order By Level, ����"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.tvwNowList.Nodes.Clear
        Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    Set objNode = Me.tvwNowList.Nodes.Add(, , "_" & !ID, !���� & "-" & !����, "close")
                Else
                    Set objNode = Me.tvwNowList.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, !���� & "-" & !����, "close")
                End If
                objNode.Expanded = True
                objNode.Sorted = True: objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub zlTree(oNode As IXMLDOMNode, Optional strHigher As String)
'----------------------------------------------------------------------------------------------------------
'��xml�ļ��еĽڵ�󶨵������б���
'----------------------------------------------------------------------------------------------------------
    Dim objNode As MSComctlLib.Node
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim i As Long, rsCount As ADODB.Recordset
    Dim j As Long, introw As Long
    
    For i = 0 To oNode.selectNodes("Class").Length - 1
        Set oSubNode1 = oNode.selectNodes("Class")(i)
        If Not oSubNode1 Is Nothing Then
            If GetNodeValue(oSubNode1, "�ϼ�id", "") = "" Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & GetNodeValue(oSubNode1, "ID", ""), GetNodeValue(oSubNode1, "����", "") & "-" & GetNodeValue(oSubNode1, "����", ""), "close")
                 objNode.Expanded = True
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & GetNodeValue(oSubNode1, "�ϼ�id", ""), tvwChild, "_" & GetNodeValue(oSubNode1, "ID", ""), GetNodeValue(oSubNode1, "����", "") & "-" & GetNodeValue(oSubNode1, "����", ""), "close")
            End If
            objNode.Tag = GetNodeValue(oSubNode1, "˵��", "") & vbCrLf & GetNodeValue(oSubNode1, "��Χ", "")
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
    
            If GetNodeValue(oSubNode1, "�ϼ�id", "") = "" Then
                gstrSQL = "select a.id from �����ʾ���� a where  a.����=[1]"
            Else
                gstrSQL = "select a.id from �����ʾ���� a,�����ʾ���� b where  a.����=[1] and  a.�ϼ�id=b.id and b.����=[2]"
            End If
            Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode1, "����", ""), strHigher)
            '�жϷ�����Ŀ������Ƿ����
            If rsCount.RecordCount < 1 Then
                objNode.ForeColor = RGB(255, 20, 147)
                objNode.Tag = objNode.Tag & vbCrLf & vbCrLf
            Else
                objNode.Tag = objNode.Tag & vbCrLf & rsCount("id") & vbCrLf
            End If
            If oSubNode1.selectNodes("Class").Length > 0 Then
                Call zlTree(oSubNode1, GetNodeValue(oSubNode1, "����", ""))
            End If
        End If
    Next
End Sub

Private Function zlSubRefList(lngFileID As Long, Optional ByVal blnCheck As Boolean) As Long
    '******************************************************************************************************************
    '���ܣ�ˢ��װ���嵥������λ��ָ���ļ�¼��
    '������
    '���أ�
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim i As Integer
    '���ѡ�е���ͬһ�����࣬��ˢ��
    If lngFileID = mlngFileID Then Exit Function
    mlngFileID = lngFileID
    '------------------------------------------------------------------------------------------------------------------
       gstrSQL = "Select L.ID, L.����id, C.���� || '-' || C.���� As ����, L.���, L.����, L.ͨ�ü�, D.���� As ����, P.���� As ��Ա" & vbNewLine & _
                "From �����ʾ���� C, �����ʾ�ʾ�� L, ���ű� D, ��Ա�� P" & vbNewLine & _
                "Where C.ID = L.����id And L.����id = D.ID And L.��Աid = P.ID And L.����id = [1] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "ȫԺ��������") <> 0 Then
    
    ElseIf InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
     End If
    
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    '--------------------------------------------------------------------------------------------------------------
    With Me.vsfList
        .Rows = 1
        .ScrollBars = flexScrollBarNone 'Ϊ��ֹ�������������Ȳ���ʾ������
        On Error Resume Next
        Do While Not rsTemp.EOF
            .AddItem ""
            i = i + 1
            Select Case rsTemp!ͨ�ü�
                    Case 0:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                    Case 1:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-����"
                    Case 2:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-����"
                    Case Else:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                End Select
                .Cell(flexcpAlignment, i, mCol.Range) = flexAlignCenterCenter
            .TextMatrix(i, mCol.ID) = rsTemp!ID
            .TextMatrix(i, mCol.clasID) = rsTemp!����id
            .TextMatrix(i, mCol.Class) = rsTemp!����
            .TextMatrix(i, mCol.Num) = rsTemp!���
            .TextMatrix(i, mCol.pName) = rsTemp!����
            .TextMatrix(i, mCol.Depart) = rsTemp!����
            .TextMatrix(i, mCol.personnel) = rsTemp!��Ա
            If blnCheck Then
                '�ж��Ƿ��ǲ�ѡ�񲻵���ʾ�
                If mtion("_" & rsTemp!ID) <> rsTemp!ID & "" Then
                    .TextMatrix(i, mCol.Choose) = 1
                Else
                    .TextMatrix(i, mCol.Choose) = 0
                End If
            End If
            rsTemp.MoveNext
        Loop
        .ScrollBars = flexScrollBarVertical
    End With
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.vsfList.Rows
End Function
Private Sub zlSubRefText(lngWordId As Long)
    Dim rsTemp As ADODB.Recordset, lngStart As Long, strText As String
    If lngWordId = mlngWordId Then Exit Sub
    mlngWordId = lngWordId
    rtbText.Text = ""
    Err = 0: On Error GoTo Errhand
    gstrSQL = "Select ��������, �����ı�, Ҫ������, Ҫ�ص�λ From �����ʾ���� Where �ʾ�id = [1] Order By ���д���"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        Do While Not .EOF
            lngStart = Len(Me.rtbText.Text)
            Me.rtbText.SelStart = lngStart
            Me.rtbText.SelLength = 0
            Select Case !��������
            Case 0 '��������
                strText = IIf(IsNull(!�����ı�), " ", !�����ı�)
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                strText = IIf(IsNull(!�����ı�), "{" & !Ҫ������ & "}" & !Ҫ�ص�λ, "{" & !�����ı� & "}")
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            .MoveNext
        Loop
        Me.rtbText.SelStart = 0
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Me.vsfList.Move 0, 0, Me.picList.Width, Me.vsfList.Height
    Me.rtbText.Move 0, Me.vsfList.Height, Me.ScaleWidth, Me.rtbText.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zlCommfun.StopFlash
    Set mtion = Nothing
    Set oDoc = Nothing
    Set oRoot = Nothing
    If gobjFSO.FileExists(msrtXmlPathName) Then gobjFSO.DeleteFile msrtXmlPathName, True
    msrtXmlPathName = ""
End Sub
Private Sub piclist_Resize()
    With vsfList
        .Move 0, 0, picList.Width, picList.Height
        .ColWidth(mCol.Choose) = 450
        .ColWidth(mCol.Range) = 300
        .ColWidth(mCol.Class) = 0
        .ColWidth(mCol.Num) = (picList.Width - 750) / 4 - 300
        .ColWidth(mCol.pName) = (picList.Width - 750) / 4 + 300
        .ColWidth(mCol.Depart) = (picList.Width - 750) / 4
        .ColWidth(mCol.personnel) = (picList.Width - 750) / 4
        .ColWidth(mCol.clasID) = 0
        .ColWidth(mCol.ID) = 0
    End With
End Sub
Private Sub ToXml(strFilePath As String)

    Dim oNode As IXMLDOMNode
    Dim Node As MSComctlLib.Node
    Dim j As Integer
    'XML�ĵ�
    Set oDoc = New DOMDocument
    'ע��
    oDoc.appendChild oDoc.createComment(gstrSysName & "  " & _
        "����Ա:" & gstrUserName & "������:" & gstrDeptName & "��ʱ��:" & _
        Format(Now(), "YYYY��MM��DD��"))
    '���ڵ�
    Set oRoot = oDoc.createElement("Document")
    Set oDoc.documentElement = oRoot    '����Ϊ���ڵ�
    Call oRoot.setAttribute("EditType", "MedicalWords")
    Set mColClass = New Collection

    On Error Resume Next
    For j = 1 To tvwClass.Nodes.Count
        Set Node = tvwClass.Nodes(j)
        If Node.Checked Then
            If mColClass(Node.Parent.Key) Is Nothing Then
                Set oNode = oRoot
            Else
               Set oNode = mColClass(Node.Parent.Key)
            End If
            Call CreateChild(oNode, 1, Node)
        End If
    Next
    
    Dim pi As IXMLDOMProcessingInstruction
    Set pi = oDoc.createProcessingInstruction("xml", "version='1.0' encoding='gb2312'")
    Call oDoc.insertBefore(pi, oDoc.childNodes(0))
    'ֱ�ӱ�����ļ�����
    oDoc.Save strFilePath
    
    Set mtion = New Collection
    Set mColClass = Nothing
    Set oDoc = Nothing
End Sub
Private Sub CreateChild(Parent As IXMLDOMNode, intNodeId As Integer, Node As MSComctlLib.Node)
'----------------------------------------------------------------------------------------------
'����˵��
'Parent ���׽ڵ�
'intNodeId �ȼ�
'Node ѡ�еĽڵ�
'----------------------------------------------------------------------------------------------
     
    Dim intSenId As Double
    Dim rsSentence As ADODB.Recordset, rsContent As ADODB.Recordset, rsCondition As ADODB.Recordset
    Dim oNode As IXMLDOMNode            '���ڵ�
    Dim oSubNode1 As IXMLDOMNode        '�ӽڵ�
    Dim oSubNode2 As IXMLDOMNode         '�ڵ�
    Dim oSubNode3 As IXMLDOMNode         '�ڵ�
    Dim intNodeId1 As Integer, intNodeId2 As Integer, intNodeId3 As Integer
        'If Node.Checked Then
            '�ʾ������Ϣ
            intNodeId1 = intNodeId + 1
            intNodeId2 = intNodeId + 2
            intNodeId3 = intNodeId + 3
            Set oNode = CreateNode(intNodeId1, Parent, "Class", NODE_ELEMENT, "")
            '�ѵ�����Ľڵ㱣��
            On Error Resume Next
            If mColClass(Node.Key) Is Nothing Then
                mColClass.Add oNode, Node.Key
            End If
            
            
            CreateNode 1, oNode, "ID", , Mid(Node.Key, 2)
            If Node.Parent Is Nothing Or Parent.nodeName = "Document" Then
                CreateNode intNodeId1, oNode, "�ϼ�id", , ""
            Else
                CreateNode intNodeId1, oNode, "�ϼ�id", , Mid(Node.Parent.Key, 2)
            End If
            CreateNode intNodeId1, oNode, "����", , Split(Node.Text, "-")(0)
            CreateNode intNodeId1, oNode, "����", , Split(Node.Text, "-")(1)
            CreateNode intNodeId1, oNode, "˵��", , Split(Node.Tag, vbCrLf)(0)
            CreateNode intNodeId1, oNode, "��Χ", , Split(Node.Tag, vbCrLf)(1)
                 '�ʾ�ʾ��
            Set rsSentence = GetSentence(Mid(Node.Key, 2))
            
            
            Do While Not rsSentence.EOF
                intSenId = NVL(rsSentence!ID)
                On Error Resume Next
                If mtion("_" & intSenId) <> intSenId & "" Then
                    Set oSubNode1 = CreateNode(intNodeId1, oNode, "Sentence", NODE_ELEMENT, "")
                    CreateNode intNodeId2, oSubNode1, "ID", , rsSentence!ID
                    CreateNode intNodeId2, oSubNode1, "����id", , rsSentence!����id
                    CreateNode intNodeId2, oSubNode1, "���", , rsSentence!���
                    CreateNode intNodeId2, oSubNode1, "����", , NVL(rsSentence!����)
                    CreateNode intNodeId2, oSubNode1, "ͨ�ü�", , NVL(rsSentence!ͨ�ü�)
                    CreateNode intNodeId2, oSubNode1, "����id", , NVL(rsSentence!����ID)
                    CreateNode intNodeId2, oSubNode1, "��Աid", , NVL(rsSentence!��ԱID)
                    
                    
                    '��ȡ�������ʾ�����
                    gstrSQL = "select t.�ʾ�id,t.������,t.����ֵ from �����ʾ����� t where t.�ʾ�id=[1]"
                    Set rsCondition = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, intSenId)
                    
                    Do While Not rsCondition.EOF
                        Set oSubNode3 = CreateNode(intNodeId2, oSubNode1, "Condition", NODE_ELEMENT, "")
                        CreateNode intNodeId3, oSubNode3, "�ʾ�id", , rsCondition!�ʾ�id
                        CreateNode intNodeId3, oSubNode3, "������", , rsCondition!������
                        CreateNode intNodeId3, oSubNode3, "����ֵ", , rsCondition!����ֵ
                        rsCondition.MoveNext
                    Loop
                    '��ȡ����Ӧ�ʾ������
                    gstrSQL = "select t.�ʾ�id,t.���д���,t.��������, t.�����ı�,t.����Ҫ��id,t.�滻��,t.Ҫ������," & _
                                "t.Ҫ������,t.Ҫ�س���,t.Ҫ��С��,t.Ҫ�ص�λ,t.Ҫ�ر�ʾ,t.Ҫ��ֵ��,t.������̬,t.�������� " & _
                                " From �����ʾ���� t Where �ʾ�id = [1] Order By t.���д���"

                    Set rsContent = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, intSenId)
                    Do While Not rsContent.EOF
                        Set oSubNode2 = CreateNode(intNodeId2, oSubNode1, "Content", NODE_ELEMENT, "")
                        CreateNode intNodeId3, oSubNode2, "�ʾ�id", , rsContent!�ʾ�id
                        CreateNode intNodeId3, oSubNode2, "���д���", , rsContent!���д���
                        CreateNode intNodeId3, oSubNode2, "��������", , rsContent!��������
                        CreateNode intNodeId3, oSubNode2, "�����ı�", , NVL(rsContent!�����ı�)
                        CreateNode intNodeId3, oSubNode2, "����Ҫ��id", , NVL(rsContent!����Ҫ��ID)
                        CreateNode intNodeId3, oSubNode2, "�滻��", , NVL(rsContent!�滻��)
                        CreateNode intNodeId3, oSubNode2, "Ҫ������", , NVL(rsContent!Ҫ������)
                        CreateNode intNodeId3, oSubNode2, "Ҫ�س���", , NVL(rsContent!Ҫ�س���)
                        CreateNode intNodeId3, oSubNode2, "Ҫ��С��", , NVL(rsContent!Ҫ��С��)
                        CreateNode intNodeId3, oSubNode2, "Ҫ�ص�λ", , NVL(rsContent!Ҫ�ص�λ)
                        CreateNode intNodeId3, oSubNode2, "Ҫ�ر�ʾ", , NVL(rsContent!Ҫ�ر�ʾ)
                        CreateNode intNodeId3, oSubNode2, "Ҫ��ֵ��", , NVL(rsContent!Ҫ��ֵ��)
                        CreateNode intNodeId3, oSubNode2, "������̬", , NVL(rsContent!������̬)
                        CreateNode intNodeId3, oSubNode2, "��������", , NVL(rsContent!��������)
                        rsContent.MoveNext
                    Loop
                End If
            rsSentence.MoveNext
        Loop
    'End If
    '�ӷ���
'    If Node.Children > 0 Then
'        For j = 1 To tvwClass.Nodes.Count
'            If Not tvwClass.Nodes(j).Parent Is Nothing Then
'                If tvwClass.Nodes(j).Parent.Key = Node.Key Then
'                    If oNode Is Nothing Then
'                        Call CreateChild(Parent, intNodeId, tvwClass.Nodes(j))
'                    Else
'                        Call CreateChild(oNode, intNodeId, tvwClass.Nodes(j))
'                    End If
'                End If
'            End If
'        Next
'    End If
End Sub
Private Function GetSentence(lngFileID As Long) As ADODB.Recordset
'------------------------------------------------------------------------------------------------
'���ã����ڻ�ȡ�����ʾ��������
'����������id
'------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------
       gstrSQL = "select L.ID, L.����id, L.���, L.����, L.ͨ�ü�,l.����id,l.��Աid " & vbNewLine & _
                "From �����ʾ���� C, �����ʾ�ʾ�� L, ���ű� D, ��Ա�� P" & vbNewLine & _
                "Where C.ID = L.����id And L.����id = D.ID And L.��Աid = P.ID And L.����id = [1] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "ȫԺ��������") <> 0 Then
    
    ElseIf InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� In (1, 2) And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User))"

    Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.ͨ�ü�, 0) = 0 Or" & vbNewLine & _
                "      L.ͨ�ü� = 1 And" & vbNewLine & _
                "      L.����id In (Select R.����id From ������Ա R, �ϻ���Ա�� U Where R.��Աid = U.��Աid And U.�û��� = User) Or" & vbNewLine & _
                "      L.ͨ�ü� = 2 And L.��Աid In (Select U.��Աid From �ϻ���Ա�� U Where U.�û��� = User))"
     End If
    
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Set GetSentence = rsTemp
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub ImportXMLFile()
'-------------------------------------------------------------------------------------------------
'���ã����ڵ���ʾ��ʱ������ݴ�����Ҫ�ǶԷ���Ĵ���
'-------------------------------------------------------------------------------------------------
    Dim objNode As MSComctlLib.Node
    Dim j As Long, i As Long, strHigherId As String
    Dim lngItemID As Long, strMaxNum As String
    Dim rsMaxNum As ADODB.Recordset
    Dim oNode As IXMLDOMNode
    On Error GoTo Errhand
    For j = 1 To Me.tvwClass.Nodes.Count
        Set objNode = tvwClass.Nodes(j)
            '�ж��Ƿ���÷���
            If objNode.Checked = True Then
            Set oNode = oRoot.getElementsByTagName("Class[ID=" & Mid(objNode.Key, 2) & "]")(0)
                Debug.Print oNode.selectSingleNode("����").Text
                '�жϷ����Ƿ����ָ������
                If Split(objNode.Tag, vbCrLf)(2) <> "" Then
                    '�� ����÷���
                    lngItemID = Split(objNode.Tag, vbCrLf)(2)
                Else '�� �������࣬����
                    '�ж��Ƿ�ָ���ϼ�����
                    If Split(objNode.Tag, vbCrLf)(3) = "" Then
                        '�ж��Ƿ��Ǹ��ڵ�
                        If GetNodeValue(oNode, "�ϼ�id", "") = "" Then
                            lngItemID = zldatabase.GetNextId("�����ʾ����")
                            gstrSQL = "select max(to_number(����)) as ���� from �����ʾ���� t where t.�ϼ�id is null"
                            strHigherId = "null"
                        Else
                            '�ж���ǰ���������ϼ������Ƿ�ָ��
                            If Split(objNode.Parent.Tag, vbCrLf)(2) <> "" Then
                                strHigherId = Split(objNode.Parent.Tag, vbCrLf)(2)
                                lngItemID = zldatabase.GetNextId("�����ʾ����")
                                gstrSQL = "select ���� from �����ʾ���� where ����=(select max(to_number(����)) as ���� from �����ʾ���� t where t.�ϼ�id=[1])"
                            Else
                                '�����ݿ��в�ѯ�ϼ������id
                                Dim strSQL As String, rsLevel As ADODB.Recordset
                                strSQL = "Select ID, �ϼ�id, ����, ����, ˵��, ��Χ " & vbNewLine & _
                                            "From �����ʾ���� where ����=[1] and level=[2] " & vbNewLine & _
                                            "Start With �ϼ�id Is Null" & vbNewLine & _
                                            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
                                            "Order By ID desc"
                                Set rsLevel = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(objNode.Parent.Text, InStr(objNode.Parent.Text, "-") + 1), UBound(Split(objNode.FullPath, "\")))
                                Debug.Print Mid(objNode.Parent.Text, InStr(objNode.Parent.Text, "-") + 1) & "|" & UBound(Split(objNode.FullPath, "\"))
                                If rsLevel.RecordCount > 0 Then
                                    strHigherId = rsLevel("ID")
                                    strMaxNum = rsLevel("����")
                                    gstrSQL = "select ���� from �����ʾ���� where ����=(select max(to_number(����)) as ���� from �����ʾ���� t where t.�ϼ�id=[1])"
                                Else
                                    strHigherId = "null"
                                    gstrSQL = "select max(to_number(����)) as ���� from �����ʾ���� t where t.�ϼ�id is null"
                                End If
                                lngItemID = zldatabase.GetNextId("�����ʾ����")
                            End If
                        End If
                    Else
                        '��ȡ�ʾ�id���ϼ�����id
                        lngItemID = zldatabase.GetNextId("�����ʾ����")
                        strHigherId = Split(objNode.Tag, vbCrLf)(3)
                        gstrSQL = "select max(����) as ���� from �����ʾ���� t where t.id=[1]"
                        Set rsMaxNum = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strHigherId)
                        If rsMaxNum.RecordCount > 0 Then strMaxNum = NVL(rsMaxNum!����)
                        gstrSQL = "select max(to_number(����)) as ���� from �����ʾ���� t where t.�ϼ�id=[1]"
                    End If
                    
                    Set rsMaxNum = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strHigherId)
                    
                    If rsMaxNum.RecordCount > 0 Then
                        strMaxNum = NVL(rsMaxNum!����)
                        If Mid(strMaxNum, 1, 1) = 0 Then
                            strMaxNum = "0" & strMaxNum + 1
                        Else
                            strMaxNum = Val(strMaxNum) + 1
                        End If
                    Else
                        strMaxNum = strMaxNum & "01"
                    End If
                    If Val(strMaxNum) < 10 Then strMaxNum = "0" & Val(strMaxNum)
                    gstrSQL = "Zl_�����ʾ����_Edit(1," & lngItemID & "," & IIf(strHigherId = "", "Null", strHigherId) & ",'" & strMaxNum & "'," & _
                        " '" & GetNodeValue(oNode, "����", "") & "','" & GetNodeValue(oNode, "˵��", "") & "','" & GetNodeValue(oNode, "��Χ", "") & "')"
                    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
                Call ImportXmlSentence(oNode, lngItemID)
            End If
    Next
    Exit Sub
Errhand:
    MsgBox Err.Description
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ImportXmlSentence(oNode As IXMLDOMNode, lngItemID As Long)
'-------------------------------------------------------------------------------------------------
'����:��xml�ļ��е���ʾ�ʾ����Ŀ���
'����˵����
'onode xml����ڵ�
'lngItemID ����id
'-------------------------------------------------------------------------------------------------
    Dim oSubNode1 As IXMLDOMNode, oSubNode2 As IXMLDOMNode, oSubNode3 As IXMLDOMNode      '�ӽڵ�
    Dim i As Long, j As Long, k As Long, ArraySQL() As String, lngCount As Long
    Dim lngWordId As Long '�����ʾ�id
    Dim blnTran As Boolean
    Dim rsMaxNumber As ADODB.Recordset, strNumber As String
    
    On Error GoTo Errhand
    '��ȡ����Ӧ���νṹ�еķ���
    For i = 0 To oNode.selectNodes("Sentence").Length - 1
        Set oSubNode1 = oNode.selectNodes("Sentence")(i)
        On Error Resume Next
        If mtion("_" & GetNodeValue(oSubNode1, "ID", "")) <> GetNodeValue(oSubNode1, "ID", "") Then
            gstrSQL = "select max(���) as ��� from �����ʾ�ʾ��  where ����id=[1]"
            Set rsMaxNumber = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngItemID)
            If rsMaxNumber.RecordCount > 0 Then
                Dim intLen As Integer
                strNumber = NVL(rsMaxNumber("���"))
                If strNumber = "" Then
                    strNumber = "00001"
                Else
                    intLen = Len(strNumber)
                    strNumber = strNumber + 1
                    For j = 1 To intLen - Len(strNumber)
                        strNumber = "0" & strNumber
                    Next
                End If
            Else
                strNumber = "00001"
            End If
            '��Ӳ����ʾ�ʾ��
            lngWordId = zldatabase.GetNextId("�����ʾ�ʾ��")
            gstrSQL = lngWordId & "," & lngItemID & ",'" & strNumber & "','" & GetNodeValue(oSubNode1, "����", "") & "'"
            Select Case GetNodeValue(oSubNode1, "ͨ�ü�", "")
                Case 0:
                    gstrSQL = gstrSQL & ",0"
                Case 1:
                    gstrSQL = gstrSQL & ",1"
                Case 2:
                    gstrSQL = gstrSQL & ",2"
                Case Else:
                    gstrSQL = gstrSQL & ",null"
            End Select
            gstrSQL = gstrSQL & "," & glngDeptId & "," & glngUserId
            gstrSQL = "Zl_�����ʾ�ʾ��_Edit(1," & gstrSQL & ")"
            
            '��ȡSQL�������
            ReDim ArraySQL(1 To 2) As String
            ArraySQL(1) = gstrSQL
            'ǰ�ڴ���
            ArraySQL(2) = "Zl_�����ʾ����_Beforesave(" & lngWordId & ")"
            
            For j = 0 To oSubNode1.selectNodes("Content").Length - 1
                Set oSubNode2 = oSubNode1.selectNodes("Content")(j)
                '���ݵ��������������֣�1��2��Ҫ��
                If GetNodeValue(oSubNode2, "��������", "") = 0 Then
                    Dim strIn As String, lngLen As Long, inti As Integer, strSub As String
                    strIn = GetNodeValue(oSubNode2, "�����ı�", "")
                    strIn = Replace(strIn, "'", "' || chr(39) || '")
                    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '����strIn�ǲ�������vbCrlf�ġ�
                    lngLen = Len(strIn)
                    
                    '����4000Ϊ��ֶδ洢��
                    inti = 0
                    Do While (inti * 2000 + 1 <= lngLen)
                        lngCount = UBound(ArraySQL) + 1
                        ReDim Preserve ArraySQL(1 To lngCount) As String
                    
                        strSub = Mid(strIn, inti * 2000 + 1, 2000)
                    
                        gstrSQL = "Zl_�����ʾ����_Insert(" & lngWordId & "," & GetNodeValue(oSubNode2, "���д���", "") & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
                        
                        ArraySQL(lngCount) = gstrSQL
                       
                        inti = inti + 1
                    Loop
                Else
                    lngCount = UBound(ArraySQL) + 1
                    ReDim Preserve ArraySQL(1 To lngCount) As String
                    Dim rsCount As ADODB.Recordset, Treatmentid As String
                    If GetNodeValue(oSubNode2, "��������", "") = 2 And GetNodeValue(oSubNode2, "����Ҫ��ID", "") <> "" Then
                        gstrSQL = "select id from ����������Ŀ where id=[1]"
                        Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode2, "����Ҫ��ID", ""))
                        If rsCount.RecordCount > 0 Then
                            Treatmentid = GetNodeValue(oSubNode2, "����Ҫ��ID", "")
                        Else
                            Treatmentid = "null"
                        End If
                    Else
                        Treatmentid = "null"
                    End If
                    gstrSQL = "Zl_�����ʾ����_Insert(" & lngWordId & _
                                "," & GetNodeValue(oSubNode2, "���д���", "") & _
                                ",1,'" & GetNodeValue(oSubNode2, "�����ı�", "") & _
                                "','" & GetNodeValue(oSubNode2, "Ҫ������", "") & "'," & _
                                Treatmentid & "," & GetNodeValue(oSubNode2, "�滻��", "") & _
                                "," & IIf(GetNodeValue(oSubNode2, "Ҫ������", "") = "", "Null", GetNodeValue(oSubNode2, "Ҫ������", "")) & _
                                "," & GetNodeValue(oSubNode2, "Ҫ�س���", "") & _
                                "," & GetNodeValue(oSubNode2, "Ҫ��С��", "") & _
                                ",'" & GetNodeValue(oSubNode2, "Ҫ�ص�λ", "") & _
                                "'," & GetNodeValue(oSubNode2, "Ҫ�ر�ʾ", "") & _
                                ",'" & GetNodeValue(oSubNode2, "Ҫ��ֵ��", "") & _
                                "'," & GetNodeValue(oSubNode2, "������̬", "") & _
                                ",'" & GetNodeValue(oSubNode2, "��������", "") & "')"
                    ArraySQL(lngCount) = gstrSQL
                End If
            Next
                
            '���ڴ���
            lngCount = UBound(ArraySQL) + 1
            ReDim Preserve ArraySQL(1 To lngCount) As String
            gstrSQL = "Zl_�����ʾ����_Aftersave(" & lngWordId & ")"
            ArraySQL(lngCount) = gstrSQL
            
                    '���ò����ʾ�����
            For j = 0 To oSubNode1.selectNodes("Condition").Length - 1
                Set oSubNode3 = oSubNode1.selectNodes("Condition")(j)
                lngCount = UBound(ArraySQL) + 1
                ReDim Preserve ArraySQL(1 To lngCount) As String
                
                gstrSQL = "Zl_�����ʾ�����_Edit(" & lngWordId & ",'" & GetNodeValue(oSubNode3, "������", "") & "','" & GetNodeValue(oSubNode3, "����ֵ", "") & "')"
                
                ArraySQL(lngCount) = gstrSQL
            Next
            
            'ִ�б������
            Err = 0: On Error GoTo Errhand
            gcnOracle.BeginTrans
            blnTran = True
            For k = 1 To UBound(ArraySQL)
                gstrSQL = ArraySQL(k)
                Call zldatabase.ExecuteProcedure(gstrSQL, "cEPRDocument")
            Next
            gcnOracle.CommitTrans
            blnTran = False
        End If
    Next
    Exit Sub
Errhand:
    If InStr(1, Err.Description, "�����ʾ�����(���ơ��ϼ�ID)�����ظ�") > 0 Then MsgBox "�����ʾ�����(���ơ��ϼ�ID)�����ظ���", vbInformation, gstrSysName
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picNowList_Resize()
    With picNowList
        Me.tvwNowList.Move 0, 0, .Width, .Height - Me.cmdYes.Height - 100
    End With
End Sub
Private Sub tvwClass_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�����Ҽ��˵�
    If Button = 2 And mblnDifference = False And Me.tvwClass.Nodes.Count > 0 Then
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        Dim cbpPopup As CommandBarPopup
        Dim cbpCustom As CommandBarControlCustom
    
        Set Popup = Me.cbsThis.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            Set cbpPopup = .Add(xtpControlPopup, 21, "��ӵ�ָ������(&R)")
                cbpPopup.CommandBar.Title = "��ӵ�ָ������"
            Set cbpCustom = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 211, "��ӵ�ָ�������б�")
                cbpCustom.Handle = Me.picNowList.hwnd
                cbpPopup.ID = 21
            Set Control = .Add(xtpControlButton, 22, "ȡ��ָ������(&A)")
            Set Control = .Add(xtpControlButton, 23, "����˷���(&D)")
            Control.BeginGroup = True
            Set Control = .Add(xtpControlButton, 24, "������˷�����(&L)")
            Set cbpPopup = .Add(xtpControlPopup, 25, "ָ���ϼ�����(&S)")
                cbpPopup.CommandBar.Title = "�ϼ�����"
            cbpPopup.BeginGroup = True
            Set cbpCustom = cbpPopup.CommandBar.Controls.Add(xtpControlCustom, 251, "�ϼ������б�")
                cbpCustom.Handle = Me.picNowList.hwnd
            cbpPopup.ID = 25
            Set Control = .Add(xtpControlButton, 26, "ȡ���ϼ�����(&S)")
        End With
        Popup.ShowPopup
    End If
   
End Sub

Private Sub tvwClass_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Call CheckTvw(Node)
    Call CheckParentsNodes(Node)
    If Mid(Node.Key, 2) <> Me.vsfList.Tag Then Exit Sub
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mCol.Choose) = Node.Checked
            '��mtion���в���
            On Error Resume Next
            If Node.Checked = False Then
                mtion.Add .TextMatrix(i, mCol.ID), "_" & .TextMatrix(i, mCol.ID)
            Else
                mtion.Remove "_" & .TextMatrix(i, mCol.ID)
            End If
        Next
    End With
End Sub
Private Sub CheckTvw(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    If Node.Children < 1 Then Exit Sub
    With tvwClass
        For i = 1 To .Nodes.Count
            If Not .Nodes(i).Parent Is Nothing Then
                If .Nodes(i).Parent.Key = Node.Key Then
                    .Nodes(i).Checked = Node.Checked
                    Call CheckTvw(.Nodes(i))
                End If
            End If
        Next
    End With
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If mblnDifference Then
        Call zlSubRefList(Mid(Node.Key, 2), Node.Checked)
    Else
        Call xmlSubRefList(Mid(Node.Key, 2), Node.Checked, Node.Text)
    End If
    Me.vsfList.Tag = Mid(Node.Key, 2) '��Ǵ�ʱ�б����������Ǹ����ڵ�
    If Me.vsfList.Rows > 1 Then
        Me.stbThis.Panels(2).Text = "�÷�������" & Me.vsfList.Rows - 1 & "���ʾ�"
    Else
        Me.stbThis.Panels(2).Text = ""
    End If
End Sub
Private Sub CheckParentsNodes(ByVal oNode As Node)
 Do While (Not oNode.Parent Is Nothing)
    oNode.Parent.Checked = True
    Set oNode = oNode.Parent
 Loop
End Sub
Private Sub xmlSubRefList(lngFileID As Long, blnCheck As Boolean, strClassName As String)
    Dim strPath As String

    If oRoot Is Nothing Then Exit Sub
    strPath = "Class[ID=" & lngFileID & "]"
    Call xmlSenceList(oRoot.getElementsByTagName(strPath)(0), lngFileID, blnCheck, strClassName)
End Sub
Private Sub xmlSenceList(Node As IXMLDOMNode, lngFileID As Long, blnCheck As Boolean, strClassName As String)
    Dim oSubNode2 As IXMLDOMNode        '�ӽڵ�
    Dim i As Long, j As Long
    Dim NodeList As IXMLDOMNodeList
    Dim k As Long

        With vsfList
            .ScrollBars = flexScrollBarNone 'Ϊ��ֹ�������ݵ�ʱ��������������Ȳ���ʾ������
            .Rows = 1
            For j = 0 To Node.selectNodes("Sentence").Length - 1
                Set oSubNode2 = Node.selectNodes("Sentence")(j)
                .AddItem ""
                k = k + 1
                Select Case GetNodeValue(oSubNode2, "ͨ�ü�", "")
                    Case 0:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                    Case 1:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-����"
                    Case 2:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-����"
                    Case Else:
                        .TextMatrix(k, 5) = ""
                        .Cell(flexcpPicture, k, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-ȫԺ"
                End Select
                .Cell(flexcpAlignment, k, mCol.Range) = flexAlignCenterCenter
                .TextMatrix(k, mCol.ID) = GetNodeValue(oSubNode2, "ID", "")
                .TextMatrix(k, mCol.clasID) = GetNodeValue(oSubNode2, "����id", "")
                .TextMatrix(k, mCol.Class) = GetNodeValue(oSubNode2, "����", "")
                .TextMatrix(k, mCol.Num) = GetNodeValue(oSubNode2, "���", "")
                .TextMatrix(k, mCol.pName) = GetNodeValue(oSubNode2, "����", "")
                If blnCheck Then
                    On Error Resume Next '�ڴ˴���������쳣��ʾ��mtion��û�м�¼
                    If mtion("_" & GetNodeValue(oSubNode2, "ID", "")) <> GetNodeValue(oSubNode2, "ID", "") Then
                        .TextMatrix(k, mCol.Choose) = 1
                    Else
                        .TextMatrix(k, mCol.Choose) = 0
                    End If
                End If
                '��ȡ��������
                Dim oSubNode3 As IXMLDOMNode        '�ӽڵ�
                Dim M As Long, lngStart As Long, strText As String
                For M = 0 To oSubNode2.selectNodes("Content").Length - 1
                    Set oSubNode3 = oSubNode2.selectNodes("Content")(M)
                    lngStart = Len(Me.rtbText.Text)
                    Me.rtbText.SelStart = lngStart
                    Me.rtbText.SelLength = 0
                    Select Case GetNodeValue(oSubNode3, "��������", "")
                    Case 0 '��������
                        strText = GetNodeValue(oSubNode3, "�����ı�", "")
                        With Me.rtbText
                            .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                            .SelUnderline = False
                        End With
                    Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                        If GetNodeValue(oSubNode3, "�����ı�", "") = "" Then
                            strText = "{" & GetNodeValue(oSubNode3, "Ҫ������", "") & "}" & GetNodeValue(oSubNode3, "Ҫ�ص�λ", "")
                        Else
                            strText = "{" & GetNodeValue(oSubNode3, "�����ı�", "") & "}"
                        End If
                        With Me.rtbText
                            .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                            .SelUnderline = True
                        End With
                    End Select
                Next
                Me.rtbText.SelStart = 0
                .Cell(flexcpData, k, mCol.Range) = Me.rtbText.Text
                Me.rtbText.Text = ""
                
                Dim rsCount As ADODB.Recordset
                gstrSQL = "select a.id from �����ʾ�ʾ�� a,�����ʾ���� b where a.����id=b.id and a.����=[1] and b.����=[2]"
                Set rsCount = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetNodeValue(oSubNode2, "����", ""), (Split(strClassName, "-")(1)))
                '����ڶ�Ӧ�����²����Ÿôʾ䣬���ɫ��ʾ
                If rsCount.RecordCount < 1 Then
                    .Cell(flexcpForeColor, k, mCol.Range, k, mCol.Class) = RGB(255, 20, 147)
                End If
            Next
            .ScrollBars = flexScrollBarVertical
        End With
        Exit Sub
End Sub
Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim i As Integer
        With vsfList
            If .Rows <= 1 Then
                Exit Sub
            End If
            If (.ColWidth(mCol.Range) < x And x < .ColWidth(mCol.Range) + .ColWidth(mCol.Choose)) And Button = 1 Then
                '��ѡ�ߴʾ��ʱ��ͬʱ��δѡ�еĴʾ��¼��mtion��
                If .TextMatrix(.Row, mCol.Choose) = "1" Then
                    .TextMatrix(.Row, mCol.Choose) = "0"
                    mtion.Add .TextMatrix(.Row, mCol.ID), "_" & .TextMatrix(.Row, mCol.ID)
                Else
                    .TextMatrix(.Row, mCol.Choose) = "1"
                    For i = 1 To tvwClass.Nodes.Count
                        If Mid(tvwClass.Nodes(i).Key, 2) = .TextMatrix(.Row, mCol.clasID) Then
                            If tvwClass.Nodes(i).Checked = False Then tvwClass.Nodes(i).Checked = True
                            Exit For
                        End If
                    Next
                    On Error GoTo Errhand
                    mtion.Remove "_" & .TextMatrix(.Row, mCol.ID)
                End If
            End If
            If Button = 1 Then
                If mblnDifference Then
                    If .TextMatrix(.Row, mCol.ID) <> "" Then Call zlSubRefText(.TextMatrix(.Row, mCol.ID))
                Else
                    Me.rtbText.Text = Me.vsfList.Cell(flexcpData, vsfList.Row, mCol.Range)
                End If
            End If
        End With
        Exit Sub
Errhand:
        For i = 1 To vsfList.Rows - 1
            If vsfList.TextMatrix(i, mCol.Choose) = "" Then
                mtion.Add vsfList.TextMatrix(i, mCol.ID), "_" & vsfList.TextMatrix(i, mCol.ID)
            Else
               ' mtion.Remove "_" & vsfList.TextMatrix(i, mCol.ID)
            End If
        Next
End Sub
'################################################################################################################
'## ���ܣ�  ����һ��XML�ڵ㲢��ֵ
'##
'## ������  TabNumber   :   �������������ʾ�ж��ٸ�Tab�Ʊ���������Ķ���
'##         Parent      :   ���ڵ�
'##         Node_Type   :   �ڵ����ͣ�Ŀǰ֧�� NODE_ELEMENT ��NODE_CDATA_SECTION ��NODE_COMMENT ��NODE_ATTRIBUTE�ȣ�
'##         Node_Name   :   �ڵ�����
'##         Node_Value  :   �ڵ��ı�
'################################################################################################################
Private Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal node_name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "")
    Dim New_Node As IXMLDOMNode
    
    '�ַ�����ֵ���ã���Ӱ�����ݣ���ֻӰ���Ķ����۶�
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '�����ı��ڵ�
    '�����½ڵ�
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, node_name, "")
    '�����ı�ֵ
    New_Node.Text = Node_Value
    '��ӵ����ڵ�
    Parent.appendChild New_Node
    '���ĩβ�س�����Ӱ�����ݣ���ֻӰ���Ķ����۶�
    'Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf)   '�����ı��ڵ�
    Set CreateNode = New_Node
End Function

'################################################################################################################
'## ���ܣ�  ��ȡһ���ڵ��ֵ
'##
'## ������  CurNode         :   ��ǰ�ڵ����
'##         SubNodeName     :   �ӽڵ�����
'##         DefaultValue    :   Ĭ��ֵ
'################################################################################################################
Private Function GetNodeValue(ByVal CurNode As IXMLDOMNode, _
    ByVal SubNodeName As String, _
    Optional ByVal DefaultValue As String = "") As String
    
    On Error Resume Next
    Dim NodeTMP As IXMLDOMNode
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
    
    If InStr(GetNodeValue, vbCr) > 0 And InStr(GetNodeValue, vbLf) = 0 Then 'ֻ�лس����޻��з�
        GetNodeValue = Replace(GetNodeValue, vbCr, vbCrLf)
    ElseIf InStr(GetNodeValue, vbLf) > 0 And InStr(GetNodeValue, vbCr) = 0 Then 'ֻ�л��з��޻س���
        GetNodeValue = Replace(GetNodeValue, vbLf, vbCrLf)
    End If
End Function

'################################################################################################################
'## ���ܣ�  ȫѡ/ȫ��
'################################################################################################################
Private Function CheckAllOrClearAll(ByVal blnOn As Boolean)
    Dim oNode As Node, i As Long
    For Each oNode In Me.tvwClass.Nodes
        oNode.Checked = blnOn
        Call CheckTvw(oNode)
    Next
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, mCol.Choose) = blnOn
            '��mtion���в���
            On Error Resume Next
            If Not blnOn Then
                mtion.Add .TextMatrix(i, mCol.ID), "_" & .TextMatrix(i, mCol.ID)
            Else
                mtion.Remove "_" & .TextMatrix(i, mCol.ID)
            End If
        Next
    End With
End Function

