VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMemo 
   Caption         =   "���˱�ע�༭"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMemo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   8295
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picUserInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   8295
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      Begin VB.Label lblUserInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   120
         Picture         =   "frmMemo.frx":6852
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5430
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   635
      SimpleText      =   $"frmMemo.frx":851C
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMemo.frx":8563
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9551
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin VSFlex8Ctl.VSFlexGrid vsfMemo 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
      _cx             =   3625
      _cy             =   3625
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
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
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2640
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":8DF7
            Key             =   "������־"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":9191
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":F9F3
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16255
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":167EF
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16B89
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":16F23
            Key             =   "����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":172BD
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":17657
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":179F1
            Key             =   "������"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":1E253
            Key             =   "������"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":24AB5
            Key             =   "������"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":2B317
            Key             =   "������"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":31B79
            Key             =   "�Ѿܾ�"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":383DB
            Key             =   "������"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":3EC3D
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":4549F
            Key             =   "������"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":4BD01
            Key             =   "�ѽ���"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":52563
            Key             =   "����"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":528FD
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":52C97
            Key             =   "����_�̶�"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":594F9
            Key             =   "��Ŀ"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":5FD5B
            Key             =   "���"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":665BD
            Key             =   "����"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":6CE1F
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMemo.frx":6D831
            Key             =   "�¼�"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   1440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnAllowEdit As Boolean
Private mblnDataChanged As Boolean
Private mbln���� As Boolean
Private mrsPatiInfo As ADODB.Recordset

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Public Event AfterDataChanged()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer, objControl As CommandBarControl
    
    On Err GoTo errHandle:
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hWnd)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '�˳�
        Unload Me
        
    Case conMenu_Edit_Modify  '�༭
        mclsVsf.AllowEdit = True
        Call vsfMemo.Select(vsfMemo.Row, vsfMemo.ColIndex("��ע��Ϣ"))
    Case conMenu_Edit_Save '����
        If Not SaveData Then Exit Sub
        mclsVsf.AllowEdit = False
        DataChanged = False
        Call LoadData
    Case conMenu_Edit_Untread  'ȡ��
        If MsgBox("���Ѿ��Ըò��˱�ע��Ϣ�����޸ģ��Ƿ񱣴棿", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            DataChanged = False
            mclsVsf.AllowEdit = False
            Call LoadData
        Else
            If Not SaveData Then Exit Sub
            mclsVsf.AllowEdit = False
            DataChanged = False
            Call LoadData
        End If
    Case conMenu_Edit_Delete 'ɾ��
        If mblnAllowEdit = False Then Exit Sub
        If Not CheckWOver(vsfMemo.Row) Then
            MsgBox "������ɾ���Ǳ�����ɵ���Ŀ��", vbInformation, gstrSysName
            Exit Sub
        End If
        If MsgBox("��ȷ��Ҫɾ����ѡ��ע��Ϣ��", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        mclsVsf.AllowEdit = True
        gstrSQL = "ZL_���˱�ע��Ϣ_Delete(" & Val(vsfMemo.TextMatrix(vsfMemo.Row, vsfMemo.ColIndex("ID"))) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call mclsVsf.DeleteRow(vsfMemo.Row)
        mclsVsf.AllowEdit = False
    Case conMenu_View_Refresh
        DataChanged = False
        mclsVsf.AllowEdit = False
        Call LoadData
    End Select
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function SaveData() As Boolean
    '��������
    Dim i As Integer, iRow As Integer
    Dim blnTrans As Boolean, intTmp As Integer
    Dim strSQL() As String, str���� As String, str���� As String
    Dim lngColor As Long '-214748363
    On Error GoTo errHandle
    
    With vsfMemo
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("��ע��Ϣ"))) <> "" Then
                If zlCommFun.ActualLen(Trim(.TextMatrix(i, .ColIndex("��ע��Ϣ")))) > 200 Then
                    MsgBox "��ע��Ϣ������������� 100 �����ֻ� 200 ���ַ���", vbInformation, gstrSysName
                    vsfMemo.Row = i
                    vsfMemo.SetFocus: Exit Function
                End If
                If .TextMatrix(i, .ColIndex("���ı�־")) = "1" And Val(.TextMatrix(i, .ColIndex("ID"))) > 0 Then
                    ReDim Preserve strSQL(intTmp)
                    strSQL(UBound(strSQL)) = "Zl_���˱�ע��Ϣ_Update(" & Val(.TextMatrix(i, .ColIndex("ID"))) & "," & mlng����ID & "," & mlng��ҳID & ",'" & _
                            Trim(.TextMatrix(i, .ColIndex("��ע��Ϣ"))) & "',to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("�Ǽ�ʱ��"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("�Ǽ���"))) & "',1," & Val(.TextMatrix(i, .ColIndex("�Ƿ����"))) & ",to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("���ʱ��"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("�����"))) & "')"
                    intTmp = intTmp + 1
                ElseIf .TextMatrix(i, .ColIndex("���ı�־")) = "1" And Val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
                    ReDim Preserve strSQL(intTmp)
                    strSQL(UBound(strSQL)) = "Zl_���˱�ע��Ϣ_Update(" & zlDatabase.GetNextId("���˱�ע��Ϣ") & "," & mlng����ID & "," & mlng��ҳID & ",'" & _
                            Trim(.TextMatrix(i, .ColIndex("��ע��Ϣ"))) & "',to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("�Ǽ�ʱ��"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("�Ǽ���"))) & "',0," & Val(.TextMatrix(i, .ColIndex("�Ƿ����"))) & ",to_Date('" & _
                            Trim(.TextMatrix(i, .ColIndex("���ʱ��"))) & "','YYYY-MM-DD HH24:MI:SS'), '" & _
                            Trim(.TextMatrix(i, .ColIndex("�����"))) & "')"
                    intTmp = intTmp + 1
                End If
                .TextMatrix(i, .ColIndex("���ı�־")) = "0"
            End If
        Next
    End With
    
    If intTmp > 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(strSQL) To UBound(strSQL)
            Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    With Me.picUserInfo
         .Left = lngLeft + 30: .Top = lngTop + 30
         .width = lngRight - lngLeft - 30
         '.Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
    End With
    With Me.vsfMemo
'        If mblnAllowEdit = True Then
'            .Left = lngLeft + 30: .Top = lngTop + 30
'            .Width = lngRight - lngLeft - 30
'            .Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
'        Else
            .Left = lngLeft + 30: .Top = picUserInfo.Height + lngTop + 30
            .width = lngRight - lngLeft - 30
            .Height = lngBottom - .Top - 30 ' - IIf(mblnAllowEdit, stbThis.Height, 0)
'        End If
    End With
    
    mclsVsf.AppendRows = True
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
        
    '-------------------------------
    Case conMenu_Edit_Save '����
        Control.Enabled = DataChanged
    Case conMenu_Edit_Untread 'ȡ��
        Control.Enabled = DataChanged
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    mblnAllowEdit = True
    mbln���� = True

    '�˵�������
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call InitCommandBar
    Call InitVSFlexGrid
    Call LoadData
    'û�б༭Ȩ���ǲ�����༭
    If InStr(mstrPrivs, "���˱�ע�༭") = 0 Then mblnAllowEdit = False
    '��Ժ���˲�����༭
    If Not IsNull(mrsPatiInfo!��Ժ����) And mbln���� Then mblnAllowEdit = False
    Me.lblUserInfo = "������" & mrsPatiInfo!���� & "     " & "�Ա�" & mrsPatiInfo!�Ա� & "   " & "���䣺" & mrsPatiInfo!���� & "   " & "סԺ�ţ�" & mrsPatiInfo!סԺ��
    If Not mblnAllowEdit Then
        Me.Caption = "���˱�ע��Ϣ(��ǰ�û���" & UserInfo.���� & ")"
        'Me.picUserInfo.Visible = True
        'Me.lblUserInfo = "������" & mrsPatiInfo!���� & "     " & "�Ա�" & mrsPatiInfo!�Ա� & "   " & "���䣺" & mrsPatiInfo!���� & "   " & "סԺ�ţ�" & mrsPatiInfo!סԺ��
        Me.cbsMain.ActiveMenuBar.Visible = False
        For i = 2 To cbsMain.Count
            cbsMain(i).Visible = False
        Next
        stbThis.Visible = False
        Me.cbsMain.RecalcLayout
    End If
End Sub

Private Sub InitVSFlexGrid()
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, vsfMemo, True, True, ils16)
        Call .ClearColumn
        
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[���]", False)
        Call .AppendColumn("���ı�־", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        Call .AppendColumn("����ID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        Call .AppendColumn("��ҳID", 0, flexAlignLeftCenter, flexDTLong, , , True, , , True)
        
        Call .AppendColumn("��ע��Ϣ", 5000, flexAlignLeftCenter, flexDTString, , "����", True, , , , True)
        Call .AppendColumn("�Ǽ�ʱ��", 2000, flexAlignLeftCenter, flexDTString, , , True)
        Call .AppendColumn("�Ǽ���", 800, flexAlignLeftCenter, flexDTString, , , True)
        '51338,������,2012-09-04,����Ƿ���ɡ����ʱ�䡢�����
        Call .AppendColumn("�Ƿ����", 1000, flexAlignCenterCenter, flexDTBoolean, , "�Ƿ����", False, True)
        Call .AppendColumn("���ʱ��", 2000, flexAlignLeftCenter, flexDTString, , , True)
        Call .AppendColumn("�����", 800, flexAlignLeftCenter, flexDTString, , , True)
            
        If InStr(mstrPrivs, "���˱�ע�༭") And mblnAllowEdit Then
            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("��ע��Ϣ"), True, vbVsfEditText, , 200)
            Call .InitializeEditColumn(.ColIndex("�Ƿ����"), True, vbVsfEditCheck, , 1)
        End If
        .SysHidden(.ColIndex("���ı�־")) = True
        .IndicatorMode = 2
        .AppendRows = True
        .AllowEdit = False
    End With
End Sub

Private Sub LoadData()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Dim lngColor As Long
    
    On Error GoTo errHandle
    ' ��ȡ�����Ƿ����
    strSQL = " Select Nvl(sum(�������),0) ������� From ������� Where ����ID=[1] And ����=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If rsTmp.RecordCount > 0 Then
        mbln���� = Not CBool(Val("" & rsTmp!�������))
    End If
    
    mclsVsf.ClearGrid
            
    zlCommFun.ShowFlash "����װ�����ݣ����Ե�..."
    
    LockWindowUpdate vsfMemo.hWnd
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    
    '51338,������,2012-09-04,����Ƿ���ɡ����ʱ�䡢�����
    strSQL = "Select id, ����id, ��ҳid, ����, �Ǽ�ʱ��, �Ǽ���,�Ƿ����,���ʱ��,����� From ���˱�ע��Ϣ Where ����id = [1] And ��ҳid = [2] Order By Id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.BOF = False Then
        Call mclsVsf.LoadGrid(rsTmp)
    End If
    For lngColor = vsfMemo.FixedRows To vsfMemo.Rows - 1
        If IsDate(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("�Ǽ�ʱ��"))) = True Then vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("�Ǽ�ʱ��")) = Format(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("�Ǽ�ʱ��")), "YYYY-MM-DD HH:mm:ss")
        If IsDate(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("���ʱ��"))) = True Then vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("���ʱ��")) = Format(vsfMemo.TextMatrix(lngColor, vsfMemo.ColIndex("���ʱ��")), "YYYY-MM-DD HH:mm:ss")
    Next
    LockWindowUpdate 0
    
    zlCommFun.StopFlash
    DataChanged = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub InitCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons

    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False) '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")

        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�༭(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False '����
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)") '����
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)") '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False) '����
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)") '����
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName) '����
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False '����
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False '����
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True '����
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����

        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�༭"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
    End With
    For Each objControl In objBar.Controls
        objControl.Style = xtpButtonIconAndCaption
    Next

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyDelete, conMenu_Edit_Delete 'ɾ��
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub SetOver(ByVal lngRow As Long)
    '51338,������,2012-09-04,����Ƿ���ɡ����ʱ�䡢�����
    If vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("�Ƿ����")) = "-1" Then
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("���ʱ��")) = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("�����")) = UserInfo.����
    Else
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("���ʱ��")) = ""
        vsfMemo.TextMatrix(lngRow, vsfMemo.ColIndex("�����")) = ""
    End If
End Sub

Private Sub mclsVSF_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVSF_AfterNewRow(ByVal Row As Long, Col As Long)
    If mclsVsf.AllowEdit = False Then Exit Sub
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("���ı�־")) = "1"
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�Ǽ�ʱ��")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�Ǽ���")) = UserInfo.����
    Call SetOver(Row)
End Sub

Private Sub mclsVSF_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errHandle
    If Not CheckWOver(vsfMemo.Row) Then
        MsgBox "������ɾ���Ǳ�����ɵ���Ŀ��", vbInformation, gstrSysName
        Cancel = True
        Exit Sub
    End If
    gstrSQL = "ZL_���˱�ע��Ϣ_Delete(" & Val(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("ID"))) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mclsVSF_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("��ע��Ϣ"))) = ""
End Sub

Private Sub vsfMemo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("���ı�־")) = "1"
    If Col = vsfMemo.ColIndex("�Ƿ����") Then Call SetOver(Row)
    DataChanged = True
End Sub

Private Sub vsfMemo_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfMemo_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_AfterSort(ByVal Col As Long, Order As Integer)
    Call mclsVsf.RestoreRow(mclsVsf.SaveKey)
    vsfMemo.ShowCell vsfMemo.Row, vsfMemo.Col
End Sub

Private Sub vsfMemo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsfMemo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow = NewRow Then Exit Sub
    vsfMemo.ForeColorSel = vsfMemo.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsfMemo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfMemo_ChangeEdit()
    With vsfMemo
        Select Case .Col
        Case .ColIndex("��ע��Ϣ")
            .TextMatrix(.Row, .Col) = .EditText
            .TextMatrix(.Row, .ColIndex("���ı�־")) = "1"
            .TextMatrix(.Row, .ColIndex("�Ǽ�ʱ��")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(.Row, .ColIndex("�Ǽ���")) = UserInfo.����
            Call SetOver(.Row)
            DataChanged = True
        End Select
    End With
End Sub

Private Sub vsfMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfMemo_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytRet As Byte
    Dim strDoctor As String
    Dim bln������Ա As Boolean
    
    With vsfMemo
        If KeyCode = vbKeyReturn Then
            
            If InStr(.EditText, "'") > 0 Then
                KeyCode = 0
                .EditText = ""
                Exit Sub
            End If
            
            strText = Trim(.EditText)
                                
            Select Case Col
            Case .ColIndex("��ע��Ϣ")
                    DataChanged = True
                    .TextMatrix(Row, .ColIndex("���ı�־")) = "1"
                    .TextMatrix(Row, .ColIndex("�Ǽ�ʱ��")) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
                    .TextMatrix(Row, .ColIndex("�Ǽ���")) = UserInfo.����
                    Call SetOver(Row)
'                End If
            Case Else
                Call mclsVsf.LocationNextCell
            End Select
            
            mclsVsf.LocationNextCell
            
            mclsVsf.SetFocus , , True
        End If
    End With
End Sub

Private Sub vsfMemo_KeyPress(KeyAscii As Integer)
    '�༭����
    If InStr("'", Chr(KeyAscii)) > 0 Then Exit Sub
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsfMemo_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfMemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsfMemo.MouseRow, vsfMemo.MouseCol)
    End Select
End Sub

Private Sub vsfMemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�����˵�����
        Call gclsBase.SendLMouseButton(vsfMemo.hWnd, X, Y)
        If mclsVsf.MoveColumn = False Then
            RaiseEvent MouseUp(Button, Shift, X, Y)
        End If
    End Select
End Sub

Private Sub vsfMemo_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsfMemo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '51338,������,2012-09-04,����Ƿ���ɡ����ʱ�䡢�����
    '�Ѿ������ɵ���Ŀֻ�б��˲��ܽ����޸�
    If Not CheckWOver(Row) And mclsVsf.AllowEdit = True Then Cancel = True: Exit Sub
    '�༭����
    If Col = vsfMemo.ColIndex("��ע��Ϣ") Then vsfMemo.EditMaxLength = 200
    If Col = vsfMemo.ColIndex("�Ƿ����") And Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("��ע��Ϣ"))) = "" Then Cancel = True: Exit Sub
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfMemo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.ValidateEdit(Col, Cancel)
    If Cancel Then Exit Sub
End Sub

Public Function ShowMe(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    
    ShowMe = gblnOK
End Function

Private Function CheckWOver(ByVal Row As Long)
'�����Ѿ���ɵ���Ŀ����Ƿ��Ǳ������
    If Val(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�Ƿ����"))) = -1 Then '�����Ѿ����
        If Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�����"))) = "" Then Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�����"))) = UserInfo.����
        If Trim(vsfMemo.TextMatrix(Row, vsfMemo.ColIndex("�����"))) <> Trim(UserInfo.����) Then Exit Function
    End If
    CheckWOver = True
End Function
