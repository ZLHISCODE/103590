VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicShortCut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6120
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   2535
   ControlBox      =   0   'False
   Icon            =   "frmClinicShortCut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmClinicShortCut"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraScope 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   2055
      Begin VB.OptionButton optScope 
         Caption         =   "ȫԺ"
         ForeColor       =   &H000040C0&
         Height          =   255
         Index           =   2
         Left            =   1400
         TabIndex        =   13
         Top             =   0
         Width           =   700
      End
      Begin VB.OptionButton optScope 
         Caption         =   "����"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   1
         Left            =   700
         TabIndex        =   12
         Top             =   0
         Width           =   680
      End
      Begin VB.OptionButton optScope 
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   680
      End
   End
   Begin XtremeSuiteControls.TabControl tbcScheme 
      Height          =   1065
      Left            =   570
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   1365
      _Version        =   589884
      _ExtentX        =   2408
      _ExtentY        =   1879
      _StockProps     =   64
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   450
      ScaleHeight     =   240
      ScaleWidth      =   1890
      TabIndex        =   1
      Top             =   390
      Width           =   1890
      Begin VB.Label lblClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1665
         TabIndex        =   4
         Top             =   30
         Width           =   210
      End
      Begin VB.Label lblMenu 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "��"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   1425
         TabIndex        =   3
         Top             =   30
         Width           =   180
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   75
         TabIndex        =   2
         Top             =   30
         Width           =   390
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   4635
      Left            =   330
      TabIndex        =   0
      Top             =   870
      Width           =   1815
      _cx             =   3201
      _cy             =   8176
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
      MousePointer    =   54
      BackColor       =   12648384
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   0
      BackColorBkg    =   12648384
      BackColorAlternate=   12648384
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   15659506
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClinicShortCut.frx":000C
      ScrollTrack     =   -1  'True
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
      Ellipsis        =   1
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   195
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   45
      Index           =   3
      Left            =   315
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   5925
      Width           =   2000
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   45
      Index           =   2
      Left            =   330
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   105
      Width           =   2000
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   5835
      Index           =   1
      Left            =   2385
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   135
      Width           =   45
   End
   Begin VB.Label lblBdr 
      BackColor       =   &H00808080&
      Height          =   5835
      Index           =   0
      Left            =   60
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   75
      Width           =   45
   End
End
Attribute VB_Name = "frmClinicShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event ItemClick(ByVal ���� As Integer, ByVal ����ID As Long)

Private mrslist As ADODB.Recordset

Private mint��Χ As Integer
Private mstr����ID As String
Private mfrmParent As Object
Private mobjPop As CommandBar
Private mintType As Integer
Private mblnShow As Boolean
Private mlngPreRow As Long
Private mblnFirst As Boolean
Private mblnNoChange As Boolean
Private Const mlngMinH = 1000
Private Const mlngMinW = 200

Public Sub ShowMe(frmParent As Object, ByVal int���� As Integer, ByVal int��Χ As Integer, ByVal lng����ID As Long, lng����id As Long, Optional ByVal blnBySave As Boolean, Optional ByVal lngҽ������ID As Long)
'������blnBySave=�Ƿ�����ϴ������ʾ��������ʾ
'������int����=���ó��ϣ�0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      int��Χ=������Դ��1-����,2-סԺ
    Dim blnShow As Boolean
    
    Set mfrmParent = frmParent
    mint��Χ = int��Χ
    
    '�������ÿ���
    If int���� = 2 Then
        mstr����ID = "," & lngҽ������ID & ","
    Else
        mstr����ID = IIF(lng����ID <> 0, "," & lng����ID, "") & "," & lng����id & ","
    End If
    
    If blnBySave Then
        blnShow = Val(zlDatabase.GetPara("��ʾ����������", glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))) <> 0
    Else
        blnShow = Not mblnShow
    End If
    
    mblnFirst = True
    
    If blnShow Then
        mblnShow = True
        Me.Show , frmParent
        Call ClearListSel
    Else
        If Not mrslist Is Nothing Then Me.Hide '�����˲�����
        mblnShow = False
    End If
    
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Public Sub ShowShortCut(ByVal intType As Integer)
'���ܣ��л��˵�����ʾ
'������intType=1-9,��Ӧ��Ӧ˳��Ĳ˵�
    Dim objControl As CommandBarControl
    
    If Not lblMenu.Enabled Then Exit Sub
    
    If mblnShow Then
        Set objControl = mobjPop.FindControl(, intType)
        If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
    End If
End Sub

Public Sub SaveShowState()
'���ܣ����������ʾ���
    '����ҽ���´�ģ��Ϊ׼
    Call zlDatabase.SetPara("��ʾ����������", IIF(mblnShow, 1, 0), glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))
End Sub

Private Function GetSchemeClass(ByVal intType As Integer) As ADODB.Recordset
'���ܣ���ȡ���׷����ļ�¼��
    Dim strSql As String
 
    On Error GoTo errH
    If intType = 0 Then '����
        strSql = " And c.��ԱID=[2]"
    ElseIf intType = 1 Then '����
        strSql = " And Exists(Select 1 From �������ÿ��� Where ��ĿID=c.ID And Instr([3],','||����ID||',')>0)"
    ElseIf intType = 2 Then 'ȫԺ
        strSql = " And c.��ԱID is NULL And Not Exists(Select 1 From �������ÿ��� Where ��ĿID=c.ID)"
    End If
    strSql = "Select Decode(a.�ϼ�id, Null, '', '  ') || a.���� as ���׷���, a.Id" & vbNewLine & _
            "From ���Ʒ���Ŀ¼ A" & vbNewLine & _
            "Where a.���� = 6 And Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From ������ĿĿ¼ C" & vbNewLine & _
            "       Where c.��� = '9' And " & IIF(mint��Χ = 3, "Nvl(c.�������,0)<>0", "c.������� IN([1],3)") & " And c.����id = a.Id" & strSql & vbNewLine & _
            "             And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null)" & vbNewLine & _
            "             And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null))" & _
            "Order By ����"
    Set GetSchemeClass = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mint��Χ, UserInfo.ID, mstr����ID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadSchemeClass()
'���ܣ����س��׷���Ŀ¼
    Dim rsTmp As ADODB.Recordset, i As Integer, objItem As TabControlItem
    
    mblnNoChange = True
    
    tbcScheme.RemoveAll
    Set rsTmp = GetSchemeClass(Val("" & tbcScheme.Tag))
    For i = 1 To rsTmp.RecordCount
        Set objItem = tbcScheme.InsertItem(i - 1, rsTmp!���׷���, vsList.Hwnd, 0)
        objItem.Tag = Val(rsTmp!ID)
        rsTmp.MoveNext
    Next
    If rsTmp.RecordCount = 0 Then
        Set objItem = tbcScheme.InsertItem(0, "�޳��׷���Ŀ¼", vsList.Hwnd, 0)
        objItem.Tag = -1
    End If
    mblnNoChange = False
    
    If tbcScheme.ItemCount > 0 Then
        If tbcScheme.ItemCount > 1 Then
            tbcScheme.Item(1).Selected = True   'ǿ�ƽ�����������û��ʾ����
        End If
        tbcScheme.Item(0).Selected = True
        Call tbcScheme_SelectedChanged(tbcScheme.Item(0))
    End If
    
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control Is Nothing Then Exit Sub
    
    lblTitle.Caption = Control.Caption
    If InStr(lblTitle.Caption, "(") > 0 Then
        lblTitle.Caption = Split(lblTitle.Caption, "(")(0)
    End If
      
    If Control.ID = 8 Then
    
        fraScope.Visible = True
        tbcScheme.Enabled = True
        tbcScheme.Visible = True
        Call LoadSchemeClass
                        
    ElseIf tbcScheme.Visible Then
        '����ѡ��ĳ�����ʾ��Χ������ҽ���´�ģ��Ϊ׼
        
        tbcScheme.RemoveAll
        tbcScheme.Visible = False
        fraScope.Visible = False
        
        '����TAB�󶨻���
        SetParent vsList.Hwnd, Me.Hwnd
        Me.Width = Me.Width + 30
        Me.Width = Me.Width - 30
    End If
    
    Call FillList(Control.ID)
   
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control Is Nothing Then Exit Sub
    
    Control.Checked = Control.ID = mintType
End Sub

Private Sub Form_Activate()
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
        
    If mintType = 8 Then
        If tbcScheme.ItemCount > 1 Then
            tbcScheme.Item(1).Selected = True   'ǿ�ƽ�����������û��ʾ����
            tbcScheme.Item(0).Selected = True
        End If
        Me.Width = Me.Width + 30
        Me.Width = Me.Width - 30
    Else
        SetParent vsList.Hwnd, Me.Hwnd '���ӵķ�vsList��ʾ������
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("`") Then
        KeyAscii = 0
        Me.Hide
        mblnShow = False
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim objControl As CommandBarControl
    Dim intType As Integer
    Dim strPos As String, lngH As Long, lngW As Long
    
    Call zlControl.FormSetCaption(Me, False, False)

    '��ʼ��������
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    
    Set mobjPop = cbsMain.Add("�����˵�", xtpBarPopup)
    With mobjPop.Controls
        Set objControl = .Add(xtpControlButton, 1, "��ҩĿ¼(&1)")
        Set objControl = .Add(xtpControlButton, 2, "��ҩĿ¼(&2)")
        Set objControl = .Add(xtpControlButton, 3, "��ҩĿ¼(&3)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, 4, "�䷽Ŀ¼(&4)")
        Set objControl = .Add(xtpControlButton, 5, "����Ŀ¼(&5)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, 6, "��������(&6)")
        Set objControl = .Add(xtpControlButton, 7, "����Ŀ¼(&7)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, 8, "���׷���(&8)")
    End With
    With cbsMain.KeyBindings
        .Add FALT, vbKey1, 1
        .Add FALT, vbKey2, 2
        .Add FALT, vbKey3, 3
        .Add FALT, vbKey4, 4
        .Add FALT, vbKey5, 5
        .Add FALT, vbKey6, 6
        .Add FALT, vbKey7, 7
        .Add FALT, vbKey8, 8
    End With
    
    '��ʼ��ѡ�
    With tbcScheme
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameNone
            
            .Color = xtpTabColorOffice2003
            .OneNoteColors = True
            
            .Layout = xtpTabLayoutAutoSize
            .BoldSelected = True
            .HotTracking = True
        End With
    End With
    
    vsList.BackColor = cbsMain.GetSpecialColor(XPCOLOR_3DFACE)
    vsList.BackColorBkg = cbsMain.GetSpecialColor(XPCOLOR_3DFACE)
    fraScope.BackColor = vsList.BackColor
    '-------------------------------------------------------------------
    mblnNoChange = False
    fraScope.Visible = False
    On Error GoTo errH
    
    '��Ϊ�����˳��׷��������Ͱ��˵�˳����һ��
    strSql = "Select ID,Decode(����,6,7,7,6,����) as ����,����,���� From ���Ʒ���Ŀ¼" & _
        " Where ���� IN(1,2,3,4,5,6,7) And �ϼ�ID Is Null And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by ����,����"
    Set mrslist = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrslist, strSql, Me.Caption)
    

    '0-���ˣ�1-���ң�2-ȫԺ
    intType = Val(zlDatabase.GetPara("��������׷�Χ", glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�), , _
            Array(fraScope, optScope(0), optScope(1), optScope(2))))
    If intType > optScope.UBound Then intType = 0
    optScope(intType).value = True
    tbcScheme.Tag = intType
    
    intType = 0
    mintType = Val(zlDatabase.GetPara("��������������", glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�), , , , intType)) + 1
    Set objControl = mobjPop.FindControl(, mintType)
    Call cbsMain_Execute(objControl)
    If (intType = 3 Or intType = 15) Then
        lblMenu.Enabled = False
    End If
    
    '�ָ�����ߴ�,Ҫ�Ȼָ��ߴ�(ԭ���Ŀ�������õĿ�ȴ��ڲ��죩
    strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutSize" & mintType, "")
    If strPos <> "" Then
        lngW = Val(Split(strPos, ",")(0))
        lngH = Val(Split(strPos, ",")(1))
    Else
        lngW = 2200
        lngH = 6000
    End If
    If lngW < 2200 Then lngW = 2200
    If lngH < 3000 Then lngH = 3000
    
    Me.Height = lngH: Me.Width = lngW
    strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutPostion", "375,-100")
    Me.Top = mfrmParent.Top + Val(Split(strPos, ",")(0))
    Me.Left = mfrmParent.Left + mfrmParent.Width + Val(Split(strPos, ",")(1)) - Me.Width
        
    If (Me.Left + Me.Width > Screen.Width) Or (Me.Left < 0) Or (Me.Top + Me.Height > Screen.Height) Or (Me.Top < 0) Then
        Me.Left = mfrmParent.Left + mfrmParent.Width - 100 - Me.Width
        Me.Top = mfrmParent.Top + 375
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetSchemeList() As ADODB.Recordset
'���ܣ���ȡ���׷����ļ�¼��
    Dim strSql As String, lngSchemeClass As Long
 
    On Error GoTo errH
    If optScope(0).value Then '����
        strSql = " And ��ԱID=[2]"
    ElseIf optScope(1).value Then '����
        strSql = " And Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID And Instr([3],','||����ID||',')>0)"
    Else 'ȫԺ
        strSql = " And ��ԱID is NULL And Not Exists(Select 1 From �������ÿ��� Where ��ĿID=A.ID)"
    End If
    
    If Not tbcScheme.Selected Is Nothing Then
        lngSchemeClass = Val(tbcScheme.Selected.Tag)
        strSql = strSql & " And A.����ID = [4]"
    End If
    strSql = "Select ID,8 as ����,����,���� From ������ĿĿ¼ A" & _
        " Where ���='9' And " & IIF(mint��Χ = 3, "Nvl(A.�������,0)<>0", "A.������� IN([1],3)") & strSql & _
        " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by ����"
    Set GetSchemeList = zlDatabase.OpenSQLRecord(strSql, "��ȡ���׷���", mint��Χ, UserInfo.ID, mstr����ID, lngSchemeClass)

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FillList(ByVal intType As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
'    Dim strPos As String, lngH As Long, lngW As Long
'    Dim lngTbcH As Long
    
    mlngPreRow = -1
    mintType = intType
    If intType = 8 Then
        '��ʾ�ĳ��׷�����Χ
        Set rsTmp = GetSchemeList
    Else
        Set rsTmp = mrslist
        rsTmp.Filter = "����=" & intType
    End If
    
    With vsList
        .Redraw = flexRDNone
        
        .Rows = 0
        .Cols = 1
        .Rows = rsTmp.RecordCount
        If .Rows = 0 Then .Rows = 1
        
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i - 1, 0) = " " & rsTmp!����
            .Cell(flexcpData, i - 1, 0) = CLng(rsTmp!ID)
            rsTmp.MoveNext
        Next
        .Redraw = flexRDDirect
    End With
        
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetCloseButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'������intState=0-����,1-����,2-����
    If intState = 0 Then
        lblClose.BackColor = picTitle.BackColor
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 0
    ElseIf intState = 1 Then
        lblClose.BackColor = vsList.BackColorSel
        lblClose.ForeColor = vbBlack
        lblClose.BorderStyle = 1
    ElseIf intState = 2 Then
        lblClose.BackColor = 11899525
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 1
    End If
    
    If blnSize Then
        lblClose.Width = 210
        lblClose.Height = 195
        lblClose.Left = picTitle.Width - lblClose.Width - 15
        lblClose.Top = (picTitle.Height - lblClose.Height) / 2
    End If
End Sub

Private Sub SetMenuButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'������intState=0-����,1-����,2-����
    If intState = 0 Then
        lblMenu.BackColor = picTitle.BackColor
        lblMenu.ForeColor = vbWhite
        lblMenu.BorderStyle = 0
    ElseIf intState = 1 Then
        lblMenu.BackColor = vsList.BackColorSel
        lblMenu.ForeColor = vbBlack
        lblMenu.BorderStyle = 1
    ElseIf intState = 2 Then
        lblMenu.BackColor = 11899525
        lblMenu.ForeColor = vbWhite
        lblMenu.BorderStyle = 1
    End If
    
    If blnSize Then
        lblMenu.Width = 210
        lblMenu.Height = 195
        lblMenu.Left = picTitle.Width - lblMenu.Width - lblClose.Width - 30
        lblMenu.Top = (picTitle.Height - lblMenu.Height) / 2
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
    Call SetMenuButton(0)
    Call ClearListSel
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    lblBdr(0).Left = 0: lblBdr(0).Top = 0: lblBdr(0).Height = Me.ScaleHeight
    lblBdr(1).Left = Me.ScaleWidth - lblBdr(1).Width: lblBdr(1).Top = 0: lblBdr(1).Height = Me.ScaleHeight
    lblBdr(2).Left = 0: lblBdr(2).Top = 0: lblBdr(2).Width = Me.ScaleWidth
    lblBdr(3).Left = 0: lblBdr(3).Top = Me.ScaleHeight - lblBdr(3).Height: lblBdr(3).Width = Me.ScaleWidth
    
    picTitle.Left = lblBdr(0).Width + 30
    picTitle.Top = lblBdr(2).Height + 30
    picTitle.Width = Me.Width - picTitle.Left * 2
    
    If tbcScheme.Visible Then
        fraScope.Left = picTitle.Left
        fraScope.Top = picTitle.Top + picTitle.Height + 30
        
        tbcScheme.Left = picTitle.Left
        tbcScheme.Top = fraScope.Top + fraScope.Height + 30
        tbcScheme.Height = Me.ScaleHeight - tbcScheme.Top - (lblBdr(2).Height + 30)
        tbcScheme.Width = Me.ScaleWidth - (lblBdr(0).Width + 30) * 2
        fraScope.Width = tbcScheme.Width
    Else
        vsList.Left = picTitle.Left
        vsList.Top = picTitle.Top + picTitle.Height + 30
        vsList.Height = Me.ScaleHeight - vsList.Top - (lblBdr(2).Height + 30)
        vsList.Width = Me.ScaleWidth - (lblBdr(0).Width + 30) * 2
    End If
    
    Call SetCloseButton(0, True)
    Call SetMenuButton(0, True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTop As Long, lngRight As Long
    
    '����ҽ���´�ģ��Ϊ׼
    Call zlDatabase.SetPara("��������������", mintType - 1, glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))
    
    If mfrmParent.WindowState = 0 Then
        '������������������Ͻǵ�λ��
        lngTop = Me.Top - mfrmParent.Top
        lngRight = Me.Left + Me.Width - (mfrmParent.Left + mfrmParent.Width)
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutPostion", lngTop & "," & lngRight
        
        '���洰��ߴ�
        If mintType <> 0 Then
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "ShortCutSize" & mintType, Me.Width & "," & Me.Height
        End If
        
        '����ѡ��ĳ�����ʾ��Χ
        If mintType = 8 Then
            '����ҽ���´�ģ��Ϊ׼
            Call zlDatabase.SetPara("��������׷�Χ", Val("" & tbcScheme.Tag), glngSys, IIF(mint��Χ = 1, p����ҽ���´�, pסԺҽ���´�))
        End If
    End If
    
    mblnShow = False
    Set mrslist = Nothing
End Sub

Private Sub lblBdr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next
        If Index = 0 Then
            If Me.Width - x > mfrmParent.Width / 2 Then Exit Sub
            If Me.Width - x < mlngMinW Then Exit Sub
            Me.Left = Me.Left + x
            Me.Width = Me.Width - x
        ElseIf Index = 1 Then
            If Me.Width + x > mfrmParent.Width / 2 Then Exit Sub
            If Me.Width + x < mlngMinW Then Exit Sub
            Me.Width = Me.Width + x
        ElseIf Index = 2 Then
            If Me.Height - Y < mlngMinH Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 3 Then
            If Me.Height + Y < mlngMinH Then Exit Sub
            Me.Height = Me.Height + Y
        End If
        Call Form_Resize
    End If
End Sub

Private Sub lblBdr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call SetCloseButton(2)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetMenuButton(0)
    
    If x >= 0 And Y >= 0 And x <= lblClose.Width And Y <= lblClose.Height Then
        If Button = 1 Then
            Call SetCloseButton(2)
        Else
            Call SetCloseButton(1)
        End If
    Else
        Call SetCloseButton(1)
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And Y >= 0 And x <= lblClose.Width And Y <= lblClose.Height Then
        Me.Hide
        mblnShow = False
        Call SetCloseButton(0)
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
    End If
End Sub

Private Sub lblMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call SetMenuButton(2)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub lblMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
    
    If x >= 0 And Y >= 0 And x <= lblMenu.Width And Y <= lblMenu.Height Then
        If Button = 1 Then
            Call SetMenuButton(2)
        Else
            Call SetMenuButton(1)
        End If
    Else
        Call SetMenuButton(1)
    End If
End Sub

Private Sub lblMenu_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim vPoint As PointAPI
    
    If x >= 0 And Y >= 0 And x <= lblMenu.Width And Y <= lblMenu.Height Then
        Call SetCloseButton(0)
        
        vPoint.x = lblMenu.Left / Screen.TwipsPerPixelX
        vPoint.Y = (lblMenu.Top + lblMenu.Height) / Screen.TwipsPerPixelY
        ClientToScreen picTitle.Hwnd, vPoint
        mobjPop.ShowPopup , vPoint.x * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
    Else
        Call SetMenuButton(0)
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub optScope_Click(Index As Integer)
    tbcScheme.Tag = Index
    If Not Me.Visible Then Exit Sub
    
    Call LoadSchemeClass
    Call Form_Resize
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
        If mfrmParent.Visible Then
            mfrmParent.SetFocus
        End If
    End If
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
    Call SetMenuButton(0)
End Sub

Private Sub tbcScheme_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNoChange Then Exit Sub
    Call FillList(mintType)
    
    If Me.Visible Then
        If vsList.Visible And vsList.Enabled Then vsList.SetFocus             '���ӵķ�vsList��ʾ������
    End If
    Me.Width = Me.Width + 30    '���ӵ�Item�ᱻvsList����
    Me.Width = Me.Width - 30
    
    If mfrmParent.Visible And mfrmParent.Enabled Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long, lng����ID As Long
    
    With vsList
        lngRow = .MouseRow
        If lngRow >= 0 Then
            lng����ID = .Cell(flexcpData, lngRow, 0)
            If lng����ID <> 0 Then
                Call ClearListSel
                RaiseEvent ItemClick(mintType, lng����ID)
            End If
        End If
    End With
    
    If mfrmParent.Visible And mfrmParent.Enabled Then
        mfrmParent.SetFocus
    End If
End Sub

Private Sub vsList_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long
    
    Call SetCloseButton(0)
    Call SetMenuButton(0)
    
    With vsList
        lngRow = .MouseRow
        If lngRow >= 0 And mlngPreRow <> lngRow Then
            Call ClearListSel
            .Cell(flexcpForeColor, lngRow, 0) = .ForeColorSel
            .Cell(flexcpBackColor, lngRow, 0) = .BackColorSel
            .CellBorderRange lngRow, 0, lngRow, 0, 1, 1, 1, 1, 1, 1, 1
            
            mlngPreRow = lngRow
            
            .ToolTipText = .Cell(flexcpText, lngRow, 0)
        End If
    End With
End Sub

Private Sub ClearListSel()
    With vsList
        If mlngPreRow >= 0 Then
            .Cell(flexcpForeColor, mlngPreRow, 0) = .ForeColor
            .Cell(flexcpBackColor, mlngPreRow, 0) = .BackColor
            .CellBorderRange mlngPreRow, 0, mlngPreRow, 0, 0, 0, 0, 0, 0, 0, 0
            mlngPreRow = -1
        End If
    End With
End Sub

