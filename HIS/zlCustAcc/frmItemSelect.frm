VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "�շ���Ŀѡ����"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmItemSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10185
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5190
      Left            =   3255
      TabIndex        =   1
      Top             =   435
      Width           =   6900
      _cx             =   12171
      _cy             =   9155
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemSelect.frx":058A
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   1215
         Top             =   1050
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSelect.frx":0617
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSelect.frx":0AF1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   480
      Left            =   30
      TabIndex        =   7
      Top             =   -75
      Width           =   10155
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   210
         Left            =   225
         TabIndex        =   8
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.CheckBox chkSub 
      Caption         =   "��ʾ�����¼���Ŀ(&S)"
      Height          =   210
      Left            =   555
      TabIndex        =   5
      Top             =   6210
      Width           =   2295
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3195
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   480
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   380
      Left            =   7365
      TabIndex        =   4
      Top             =   6180
      Width           =   1250
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6105
      TabIndex        =   3
      Top             =   6180
      Width           =   1250
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   540
      Left            =   3315
      TabIndex        =   2
      Top             =   5460
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   953
      TabWidthStyle   =   2
      TabFixedWidth   =   1764
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��(0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�г�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�в�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":0FCB
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":1565
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":1AFF
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2099
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2633
            Key             =   "��ҩ"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2BCD
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5565
      Left            =   15
      TabIndex        =   0
      Top             =   435
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   9816
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
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
   Begin VB.Shape Shp 
      Height          =   405
      Left            =   3195
      Top             =   6105
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   12000
      Y1              =   6045
      Y2              =   6045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   6060
      Y2              =   6060
   End
End
Attribute VB_Name = "frmItemSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mint������Դ As Integer
Private mint���� As Integer
Private mstr��� As String
Private mstr���� As String
Private mlngHwnd As Long
Private mstr��׼��Ŀ As String

Private mrsItem As ADODB.Recordset
Private mlng��ĿID As Long
Private mblnOK As Boolean
Private mstrLike As String

Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean
Private mstrPriceGrade As String

Public Function ShowSelect(frmParent As Object, ByVal strPrivs As String, _
    ByVal int������Դ As Integer, ByVal int���� As Integer, ByVal str��� As String, Optional ByVal str���� As String, _
    Optional ByVal lngHwnd As Long, Optional ByVal str��׼��Ŀ As String, Optional ByVal strPriceGrade As String) As Long
'���ܣ���ʾ�շ���Ŀѡ����
'������int������Դ=ָ������Դ,1-����,2-סԺ,<=0-������
'      str���="'5','D','Z'..",��ʾ����ѡ���ǰȷ��Ҫ��������,Ϊ�ձ�ʾ�������
'      str����=����ƥ�������,���û����Ϊѡ������ʽ,����Ϊ�б�ʽ
'      lngHwnd=�����б�λ�������ľ��
'      str��׼��Ŀ=����ҽ������
'���أ����û������(����ʾ),��ȡ��,�򷵻�0�������շ���ĿID
    mstrPrivs = strPrivs
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    mint������Դ = int������Դ
    mint���� = int����
    mstr��� = str���
    mstr���� = str����
    mlngHwnd = lngHwnd
    mstr��׼��Ŀ = str��׼��Ŀ
    mstrPriceGrade = strPriceGrade
    
    mstrSaveTag = IIf(mstr���� <> "", 1, 0)
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        ShowSelect = mlng��ĿID
    End If
End Function

Private Sub SaveColWidth(Optional ByVal strType As String)
'���ܣ������п��
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'���ܣ��ָ��п��
'˵����Ӧ���ڻָ�����֮��
    Dim strType As String
    
    If Not gblnMyStyle Then Exit Sub
    
    If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlng��ĿID = Val(vsItem.TextMatrix(vsItem.Row, 1))
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    If Not tvw_s.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf Shift = vbAltMask Then
        If Between(KeyCode, vbKey0, vbKey9) Then
            lngIdx = KeyCode - vbKey0 + 1
        End If
        If tabClass.SelectedItem.Index <> lngIdx And Between(lngIdx, 1, tabClass.Tabs.Count) Then
            tabClass.Tabs(lngIdx).Selected = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long
    Dim vRect As RECT, strIDs As String, i As Long
    Dim lngUpH As Long, lngDnH As Long
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    
    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    mlng��ĿID = 0
    mstrLike = gstrLike
    
    If mstr���� = "" Then
        '��ȡ���ʧ��,����ʾ,��ȡ���˳�
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '�����,��ʾ,��ȡ���˳�
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "û����������շ���Ŀ���,���ȵ��շ���Ŀ���������á�", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    Else
        fraInfo.Visible = False
        tvw_s.Visible = False
        fraLR.Visible = False
        chkSub.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Line1(0).Visible = False
        Line1(1).Visible = False
        Shp.Visible = True

        '���ƥ������
        Call FillList(True, strIDs)
        
        If mrsItem Is Nothing Then
            Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount = 1 Then
            'ֻ��һ����Ŀʱ,ֱ�ӷ���
            mlng��ĿID = Val(vsItem.TextMatrix(vsItem.Row, 1))
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount > 0 Then
            '������ͬһ����Ŀʱ,ֱ�ӷ���
            If mstr���� <> "" Then
                If UBound(Split(strIDs, ",")) = 1 Then
                    mlng��ĿID = Val(vsItem.TextMatrix(vsItem.Row, 1))
                    mblnOK = True: Unload Me: Exit Sub
                End If
            End If
            
            vsItem.Appearance = flexFlat
            Call zlControl.FormSetCaption(Me, False, False)
            Call GetWindowRect(mlngHwnd, vRect) '�����λ��
            vRect.Left = vRect.Left - 2
            vRect.Top = vRect.Top - 4
            vRect.Bottom = vRect.Bottom + 4
            
            '���ô���ߴ��λ��
            '������
            Me.Left = vRect.Left * Screen.TwipsPerPixelX
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX + 60 '+3D�߿�
            For i = 0 To vsItem.Cols - 1
                lngColW = lngColW + IIf(vsItem.ColHidden(i), 0, vsItem.ColWidth(i))
            Next
            If Me.Left + lngColW + lngScrW > Screen.Width - lngScrW Then
                Me.Width = Screen.Width - lngScrW - Me.Left
            Else
                Me.Width = lngColW + lngScrW
            End If
            
            '����߶�
            lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * Screen.TwipsPerPixelY '��Ļ���ø߶�
            lngUpH = vRect.Top * Screen.TwipsPerPixelY '������ø߶�
            lngDnH = lngScrH - vRect.Bottom * Screen.TwipsPerPixelY '������ø߶�
            Me.Height = vsItem.Rows * vsItem.RowHeight(0) + 375 '+���Ƭ�߶�
            If Me.Height < 1500 Then Me.Height = 1500 '������С�߶�
            If Me.Height > lngUpH And Me.Height > lngDnH Then
                Me.Height = IIf(lngUpH < lngDnH, lngDnH, lngUpH)
            End If
            If Me.Height > lngScrH / 2 Then Me.Height = lngScrH / 2 '�������߶�
            If Me.Height <= lngDnH Then
                Me.Top = vRect.Bottom * Screen.TwipsPerPixelY
            ElseIf Me.Height <= lngUpH Then
                Me.Top = vRect.Top * Screen.TwipsPerPixelY - Me.Height
            End If
            
            Call Form_Resize
        Else
            '������,��ʾ,��ȡ���˳�
            MsgBox "û���ҵ�������������շ���Ŀ��", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mstr���� = "" Then
        fraInfo.Left = 0
        fraInfo.Width = Me.ScaleWidth
        
        tvw_s.Left = 0
        tvw_s.Top = fraInfo.Top + fraInfo.Height + 15
        tvw_s.Height = Me.ScaleHeight - tvw_s.Top - 690
        
        fraLR.Top = tvw_s.Top
        fraLR.Left = tvw_s.Left + tvw_s.Width
        fraLR.Height = tvw_s.Height
        
        vsItem.Top = tvw_s.Top
        vsItem.Left = fraLR.Left + fraLR.Width
        vsItem.Width = Me.ScaleWidth - tvw_s.Width - fraLR.Width
        vsItem.Height = tvw_s.Height - IIf(tabClass.Visible, 350, 0)
        
        If tabClass.Visible Then
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
            tabClass.Left = vsItem.Left + 30
            tabClass.Width = vsItem.Width - 60
        End If
        
        Line1(0).X1 = 0: Line1(0).X2 = Me.ScaleWidth
        Line1(0).Y1 = tvw_s.Top + tvw_s.Height + 60: Line1(0).Y2 = Line1(0).Y1
        
        Line1(1).X1 = Line1(0).X1: Line1(1).X2 = Line1(0).X2
        Line1(1).Y1 = Line1(0).Y1 - 15: Line1(1).Y2 = Line1(1).Y1
        
        cmdOK.Top = Line1(1).Y1 + 135
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.5 < 4100 Then
            cmdCancel.Left = 4100
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.5
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
        
        chkSub.Top = cmdOK.Top + (cmdOK.Height - chkSub.Height) / 2
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        vsItem.Height = Me.ScaleHeight - IIf(tabClass.Tabs.Count > 1, 375, 0)
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - 120
            tabClass.Top = Me.ScaleHeight - tabClass.Height - 30
        End If
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
    Call SaveColPosition
    Call SaveColWidth
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        tabClass.Left = tabClass.Left + X
        tabClass.Width = tabClass.Width - X
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    '��ҩƷ�շ���Ŀ
    strSQL = _
        " Select Level as ��,0 as ����,ID,�ϼ�ID,'['||����||']'||���� as ����" & _
        " From �շѷ���Ŀ¼ Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Order by ��,����"
    
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Name)
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, "Close")
        End If
        objNode.Tag = rsTmp!���� '��ŷ�������:0-��ҩƷ,1-����ҩ,2-�г�ҩ,3-�в�ҩ
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        'tvw_s.Nodes(1).Selected = True
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        cmdOK.Enabled = Val(vsItem.TextMatrix(NewRow, 1)) <> 0
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '��Ϊ������˳��ı�,���Ա���ԭʼ�к�
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    'ǿ�Ʊ����а��ַ�������
    If vsItem.TextMatrix(0, Col) = "����" Then
        If Order = 1 Then Order = 7
        If Order = 2 Then Order = 8
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        Call vsItem_KeyPress(13)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then cmdOK_Click
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub

Private Sub tabClass_Click()
    If Not mblnClick Then Exit Sub
    Call FillList
    vsItem.SetFocus
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstrPreNode Then Exit Sub
    '���ı�ʱ,���浱ǰ˳��(������)
    If Visible Then
        Call SaveColPosition(tvw_s.Nodes(mstrPreNode).Tag)
        Call SaveColWidth(tvw_s.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = Node.Key
        
    Call FillList(True)
End Sub

Private Function GetTreePath(ByVal objNode As Node) As String
'���ܣ���ȡ����·����
    Dim tmpNode As Node, strTmp As String
    Set tmpNode = objNode
    Do While Not tmpNode Is Nothing
        strTmp = zlStr.NeedName(Replace(tmpNode.Text, Chr(13), "")) & "\" & strTmp
        Set tmpNode = tmpNode.Parent
    Loop
    GetTreePath = strTmp
End Function

Private Sub SaveColPosition(Optional ByVal strType As String)
'���ܣ�������˳��:�к�,˳��|...
'˵����Ӧ����SaveWinState֮ǰ,���ڲ�ʹ�ø��Ի�ʱ��ע������
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr���� = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub RestoreColPosition()
'���ܣ��ָ���˳��
'˵����Ӧ����������֮ǰ
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
        strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,�ı������к�Ҳ�ı�
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .Cols - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .Cols - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub RestoreColSort()
'���ܣ�������
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Not gblnMyStyle Then
            If mstr���� = "" Then strType = tvw_s.SelectedItem.Tag
            strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '��Ϊ���ܵ�����˳��,���Բ�����ʵ��������
                For i = 0 To .Cols - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .Cols - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(1).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(2).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function FillList(Optional ByVal blnClass As Boolean, Optional strIDs As String) As Boolean
'���ܣ����ݵ�ǰ��������װ��������ĿĿ¼
'������blnClass=�Ƿ��ؽ����࿨(Ӧ��������Ŀ�ı�ʱ���ؽ�)
'      strIDs=��ȡ����ĿID��,�����ж�����ʱ�Ƿ������ͬ��ͬһ���շ���Ŀ
    Dim objTab As MSComctlLib.Tab
    Dim objNode As Node, objItem As ListItem
    Dim arrClass As Variant, strClass As String, strInput As String
    Dim str��� As String, str����ID As String
    Dim strMain As String, strSQL As String, strItem As String, strSQLItem As String
    Dim i As Long, j As Long
    Dim lng�շѷ���ID As Long
    Dim strWherePriceGrade As String
    
    strIDs = ""
    Set objNode = tvw_s.SelectedItem '����ƥ��ʱ,ΪNothing
    
    '�����Ŀ�嵥�����࿨Ƭ
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
    Me.Refresh
    
    '�����������ֶ�����
    '------------------------------------------------------------------------
    If mstr���� = "" Then
        lng�շѷ���ID = Val(Mid(objNode.Key, 2))
        '�����еķ���ID
        If chkSub.Value = 1 Then
            '��ʾ�¼�����Ŀ
            '�շѷ���Ŀ¼
            str����ID = " And A.����ID IN(Select ID From �շѷ���Ŀ¼ Start With ID=[1] Connect by Prior ID=�ϼ�ID)"
        Else
            '�շѷ���Ŀ¼
            str����ID = " And A.����ID=[1]"
        End If
    Else
        '����ƥ��
        If Len(mstr����) < 2 Then mstrLike = "" '�Ż�
    End If
    
    '���Ƭȷ�����
    If tabClass.SelectedItem.Key <> "" Then
        str��� = Mid(tabClass.SelectedItem.Key, 2)
    End If
    
    
    strSQLItem = " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
    
    '��ȡ����
    '------------------------------------------------------------------------
    '�����շ���Ŀ����
    If str��� = "" Then
        If mstr��� = "" Then
            strItem = " And A.��� Not IN('4','5','6','7','1')"
        Else
            strItem = " And Instr([3],A.���)>0"
        End If
    Else
        strItem = " And A.���=[2]"
    End If
    If mstr���� <> "" Then
        strInput = " And (A.���� Like [4] Or B.���� Like [5] Or B.���� Like [5]) And B.����=[6]"
        If IsNumeric(mstr����) Then                         '10,11.����ȫ������ʱֻƥ�����
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���� Like [4] And B.����=[6]"
        ElseIf zlCommFun.IsCharAlpha(mstr����) Then         '01,11.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.���� Like [5] And B.����=[6]"
        ElseIf zlCommFun.IsCharChinese(mstr����) Then
            strInput = " And B.���� Like [5] And B.����=[6]"
        End If
        
        If mint���� = 0 Then
            strMain = _
                " Select Distinct A.ID,A.���,A.����,B.����,B.����," & _
                " A.���㵥λ,A.���,A.����,A.��������,Null ҽ������,A.˵��,A.�Ƿ���" & _
                " From �շ���Ŀ���� B,�շ���ĿĿ¼ A" & _
                " Where A.ID=B.�շ�ϸĿID And A.�������" & IIf(mint������Դ > 0, " IN([7],3)", "<>0") & strSQLItem & _
                strItem & str����ID & mstr��׼��Ŀ & strInput
        Else
            strMain = _
                " Select Distinct A.ID,A.���,A.����,B.����,B.����," & _
                " A.���㵥λ,A.���,A.����,A.��������,D.���� ҽ������,A.˵��,A.�Ƿ���" & _
                " From �շ���Ŀ���� B,�շ���ĿĿ¼ A,����֧����Ŀ C,����֧������ D" & _
                " Where A.ID=B.�շ�ϸĿID And A.�������" & IIf(mint������Դ > 0, " IN([7],3)", "<>0") & strSQLItem & _
                " And A.ID=C.�շ�ϸĿID(+) And C.����(+)=[8] And C.����ID=D.ID(+)" & vbNewLine & _
                strItem & str����ID & mstr��׼��Ŀ & strInput
        End If
    Else
        If mint���� = 0 Then
            strMain = _
                " Select A.ID,A.���,A.����,A.����," & _
                " A.���㵥λ,A.���,A.����,A.��������,Null ҽ������,A.˵��,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A" & _
                " Where A.�������" & IIf(mint������Դ > 0, " IN([7],3)", "<>0") & strSQLItem & _
                strItem & str����ID & mstr��׼��Ŀ
        Else
            strMain = _
                " Select A.ID,A.���,A.����,A.����," & _
                " A.���㵥λ,A.���,A.����,A.��������,D.���� ҽ������,A.˵��,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A,����֧����Ŀ C,����֧������ D" & _
                " Where A.�������" & IIf(mint������Դ > 0, " IN([7],3)", "<>0") & strSQLItem & _
                " And A.ID=C.�շ�ϸĿID(+) And C.����(+)=[8] And C.����ID=D.ID(+)" & vbNewLine & _
                strItem & str����ID & mstr��׼��Ŀ
        End If
    End If
    
    If mstrPriceGrade <> "" Then
        strWherePriceGrade = _
            "      And (c.�۸�ȼ� = [9]" & vbNewLine & _
            "          Or (c.�۸�ȼ� Is Null" & vbNewLine & _
            "              And Not Exists(Select 1" & vbNewLine & _
            "                             From �շѼ�Ŀ" & vbNewLine & _
            "                             Where c.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [9]" & vbNewLine & _
            "                                   And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And c.�۸�ȼ� Is Null"
    End If
    strSQL = _
        " Select A.ID,A.��� as ���ID,B.��� as ˳��ID,B.���� as ���,A.����," & _
        " A.����," & IIf(mstr���� <> "", "A.����,", "") & _
        " A.���㵥λ as ��λ,A.���,A.����,A.��������,A.ҽ������,A.˵��," & _
        " Decode(A.�Ƿ���,1,'���',LTrim(To_Char(Sum(C.�ּ�),'999999" & gstrFeePrecisionFmt & "'))) as ����" & _
        " From �շѼ�Ŀ C,(" & strMain & ") A,�շ���Ŀ��� B" & _
        " Where A.���=B.���� And A.ID=C.�շ�ϸĿID" & _
        "   And Sysdate Between C.ִ������+0 and Nvl(C.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))" & _
            strWherePriceGrade & vbNewLine & _
        " Group by A.ID,A.���,B.���,B.����,A.����,A.����," & IIf(mstr���� <> "", "A.����,", "") & _
        " A.���,A.����,A.��������,A.ҽ������,A.˵��,A.�Ƿ���,A.���㵥λ" & _
        " Order by ˳��ID,����"
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, _
        lng�շѷ���ID, str���, mstr���, UCase(mstr����) & "%", mstrLike & UCase(mstr����) & "%", _
        gbytCode + 1, mint������Դ, mint����, mstrPriceGrade)
    
    '������
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '�����Ե���
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        If InStr("����,���", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '��¼ԭʼ�к�,���ڴ�����˳��
    Next
    
    '�ָ���˳��:Ӧ����������֮ǰ
    Call RestoreColPosition
    Call RestoreColWidth
    '������:������,�Ա���洦���к�
    Call RestoreColSort
    
    '��Ƭ������ݼ���
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '�ռ����Ƭ��Ϣ
        If InStr(strClass & ",", "," & mrsItem!���ID & mrsItem!��� & ",") = 0 Then
            strClass = strClass & "," & mrsItem!���ID & mrsItem!���
        End If
        
        '�ռ���ĿID:ֻ�ռ����2��
        If mstr���� <> "" Then
            If UBound(Split(strIDs, ",")) < 2 Then
                If InStr(strIDs & ",", "," & mrsItem!ID & ",") = 0 Then
                    strIDs = strIDs & "," & mrsItem!ID
                End If
            End If
        End If
        mrsItem.MoveNext
    Next
    
    '�������࿨Ƭ:�ж���ʱ����Ŀ���϶�ʱ
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '��Alt��ݼ������޷�����
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '�к��п��
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst
    lblInfo.Caption = "��ǰѡ��" & GetTreePath(tvw_s.SelectedItem) & tabClass.SelectedItem.Tag & "������ " & mrsItem.RecordCount & " ����Ŀ"
        
    Screen.MousePointer = 0
    Exit Function
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function
