VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmQCReport 
   Caption         =   "�ʿر���"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmQCReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3390
      Left            =   4410
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3390
      ScaleWidth      =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1650
      Width           =   45
   End
   Begin VB.PictureBox picCalc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5145
      Left            =   6150
      ScaleHeight     =   5145
      ScaleWidth      =   4785
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   4785
      Begin VSFlex8Ctl.VSFlexGrid vfgWord 
         Height          =   4935
         Left            =   4635
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Width           =   2175
         _cx             =   3836
         _cy             =   8705
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   1
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
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgReport 
         Height          =   4935
         Left            =   -60
         TabIndex        =   4
         Top             =   420
         Width           =   4605
         _cx             =   8123
         _cy             =   8705
         Appearance      =   2
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
         BackColorSel    =   16635590
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
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
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
   Begin VB.PictureBox picReport 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1824
      Left            =   450
      ScaleHeight     =   1830
      ScaleWidth      =   3060
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1860
      Width           =   3060
      Begin VSFlex8Ctl.VSFlexGrid vfgReportEdit 
         Height          =   672
         Left            =   60
         TabIndex        =   1
         Top             =   252
         Width           =   1656
         _cx             =   2921
         _cy             =   1185
         Appearance      =   2
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   120
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmQCReport.frx":6852
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmQCReport.frx":D0B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmQCReport.frx":13916
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin XtremeCommandBars.ImageManager ImageLib 
      Left            =   4140
      Top             =   870
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmQCReport.frx":1A178
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQCReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintEdit As Integer
Private Enum mColR  '�ʿر������
    ID = 0: ������Ŀid: ���: ����: �걾��: ��Ŀ: ���: �ʿ�Ʒ: ˮƽ
End Enum

Private Enum mRow
    ��� = 0: ����: ��ʾ: ԭ��: ��ʩ: ����: ����: �鵵
End Enum


Public Sub ShowME(strResList As String, lngItemID, strFromDate As String, strToDate As String, frmParent As Form)
    '���ܣ�������ʾ���ԣ�ˢ�³��ʿؼ�¼��ͼ�κͱ���

    Call zlRefReport(strResList, lngItemID, strFromDate, strToDate)
    Me.Show vbModal, frmParent
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long
    Select Case Control.ID
        Case conMenu_Edit_ItemEdit          '�༭

            vfgReport.Editable = flexEDKbdMouse
            vfgReportEdit.Enabled = False
            vfgWord.Enabled = True
            cbsMain.ActiveMenuBar.Controls.Item(1).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(2).Enabled = True
            cbsMain.ActiveMenuBar.Controls.Item(3).Enabled = True
            cbsMain.ActiveMenuBar.Controls.Item(4).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(5).Enabled = False
        Case conMenu_Edit_ItemUndo          'ȡ��
        
            vfgReport.Editable = flexEDNone
 
            vfgReportEdit.Enabled = True
            vfgWord.Enabled = False
            
            Call zlRefresh(Val(Me.vfgReportEdit.TextMatrix(vfgReportEdit.RowSel, mColR.ID)))
        Case conMenu_Edit_ItemSave          '����
        
            vfgReportEdit.Enabled = True
            vfgWord.Enabled = False
            
            SaveQCReport
            
            Call zlRefresh(Val(Me.vfgReportEdit.TextMatrix(vfgReportEdit.RowSel, mColR.ID)))
        Case conMenu_Verify_AuditingLogin   '�鵵
            Call GetArchive(0)
        Case conMenu_Verify_LogOut          'ȡ���鵵
            Call GetArchive(1)
        Case conMenu_Edit_Exit              '�˳�
            Unload Me
    End Select
End Sub

Private Sub SaveQCReport()
    Dim lngResult As Long
    Dim lng_BiaoJi As Long
    Dim str_GuiZe As String
    Dim str_TiShi As String
    Dim lng_ItemId As Long
    Dim str_YuanYing As String
    Dim str_CuoShi As String
    Dim str_JieLun As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    With Me.vfgReport
        If .EditWindow <> 0 Then .TextMatrix(.Row, 1) = .EditText
        .TextMatrix(mRow.ԭ��, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.ԭ��, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        .TextMatrix(mRow.��ʩ, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.��ʩ, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        .TextMatrix(mRow.����, 1) = Trim(Replace(Replace(Replace(Replace(.TextMatrix(mRow.����, 1), vbCrLf, ""), vbCr, ""), vbLf, ""), "'", ""))
        If .TextMatrix(mRow.ԭ��, 1) = "" And .TextMatrix(mRow.��ʩ, 1) = "" And .TextMatrix(mRow.����, 1) = "" Then
            If MsgBox("��û����д�κα������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        If LenB(StrConv(.TextMatrix(mRow.ԭ��, 1), vbFromUnicode)) > 500 Then
            MsgBox "ԭ�򳬳������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            Exit Sub
        End If
        If LenB(StrConv(.TextMatrix(mRow.��ʩ, 1), vbFromUnicode)) > 500 Then
            MsgBox "��ʩ���������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            Exit Sub
        End If
        If LenB(StrConv(.TextMatrix(mRow.����, 1), vbFromUnicode)) > 500 Then
            MsgBox "���۳��������500���ַ���ȳ����֣���", vbInformation, gstrSysName
            Exit Sub
        End If
        strSQL = ",'" & .TextMatrix(mRow.ԭ��, 1) & "'"
        strSQL = strSQL & ",'" & .TextMatrix(mRow.��ʩ, 1) & "'"
        strSQL = strSQL & ",'" & .TextMatrix(mRow.����, 1) & "'"
    End With
    
    
    lngResult = Me.vfgReportEdit.TextMatrix(vfgReportEdit.RowSel, mColR.ID)
    If Me.vfgReport.TextMatrix(mRow.���, 1) = "�ڿأ�" Then
        lng_BiaoJi = 0
    ElseIf Me.vfgReport.TextMatrix(mRow.���, 1) = "���棡" Then
        lng_BiaoJi = 1
    Else
        lng_BiaoJi = 2
    End If
    str_GuiZe = Me.vfgReport.TextMatrix(mRow.����, 1)
    str_TiShi = Me.vfgReport.TextMatrix(mRow.��ʾ, 1)
    lng_ItemId = Me.vfgReportEdit.TextMatrix(vfgReportEdit.RowSel, mColR.������Ŀid)
    
    strSQL = "Zl_�����ʿر���_Update(" & lngResult & "," & lng_BiaoJi & ",'" & str_GuiZe & "','" & str_TiShi & "'," & lng_ItemId & _
             strSQL & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub GetArchive(ByVal intType As Integer)
    Dim strSQL As String
    
    On Error GoTo errH
    If intType = 0 Then
        
        strSQL = "Zl_�����ʿر���_Archive(" & Val(Me.vfgReportEdit.TextMatrix(Me.vfgReportEdit.RowSel, mColR.ID)) & ",0)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption

    Else
        
        strSQL = "Zl_�����ʿر���_Archive(" & Val(Me.vfgReportEdit.TextMatrix(Me.vfgReportEdit.RowSel, mColR.ID)) & ",1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption

    End If

    Call zlRefresh(Val(Me.vfgReportEdit.TextMatrix(vfgReportEdit.RowSel, mColR.ID)))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Resize()

    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    On Error Resume Next

    With picReport
        .Left = lngLeft + 45
        .Top = lngTop + 45
        .Width = (lngRight - lngLeft) * 0.4
        .Height = lngBottom - lngTop - 135
    End With
    With pic
        .Left = (lngRight - lngLeft) * 0.4 + 45
        .Top = lngTop + 45
        .Height = picReport.Height
    End With
    With picCalc
        .Left = lngLeft + picReport.Width + pic.Width
        .Top = lngTop + 45
        .Width = lngRight - lngLeft - picReport.Width - pic.Width
         .Height = lngBottom - lngTop - 135
    End With
    
End Sub

Private Sub picReport_Resize()
    With vfgReportEdit
        .Left = 10
        .Top = 10
        .Width = picReport.Width - 20
        .Height = picReport.Height - 20
    End With
End Sub

Private Sub picCalc_Resize()
    With vfgReport
        .Left = 45
        .Top = 45
        .Width = picCalc.Width - 45
        .Height = picCalc.Height * 0.5 - 20
    End With
    
    With vfgWord
        .Left = 45
        .Top = vfgReport.Height + 20
        .Width = picCalc.Width - 45
        .Height = picCalc.Height * 0.5 - 40
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo hErr

    '-- ������
    Dim Menus As New Collection, strSQL As String
    
    Menus.Add conMenu_Edit_ItemEdit & ",�༭,False"
    Menus.Add conMenu_Edit_ItemUndo & ",ȡ��,True"
    Menus.Add conMenu_Edit_ItemSave & ",����,False"
    
    Menus.Add conMenu_Verify_AuditingLogin & ",�鵵,False"
    Menus.Add conMenu_Verify_LogOut & ",ȡ���鵵,False"
    
    Menus.Add conMenu_Edit_Exit & ",�˳�����,True"
    Call CbsButtonInit(cbsMain, Menus, True, xtpBarTop)
    
    vfgReportEdit.Enabled = True
    vfgWord.Enabled = False
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picReport.Width + X < 1000 Or picCalc.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        picReport.Width = picReport.Width + X
        picCalc.Left = picCalc.Left + X
        picCalc.Width = picCalc.Width - X
        Me.Refresh
    End If
End Sub

Private Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
    '�����������˵�
    'cbsMain :����������
    'Buttons :�˵�����,ÿ��Ԫ�صĸ�ʽΪ �˵�id,����,�Ƿ����
    'blnLargeIcons :�Ƿ��ͼ��
    'Position      :�˵�λ��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim strButton As Variant
    Dim varButton As Variant

    Call CbsSetting(cbsMain)
    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.ActiveMenuBar
    cbsMain.Options.LargeIcons = blnLargeIcons  'Сͼ��
    objBar.Position = Position   '�������ڶ���

    For Each strButton In Buttons
        varButton = Split(strButton, ",")
        With objBar.Controls
            Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '����
            objControl.Style = xtpButtonIconAndCaption
            If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '����
        End With
    Next
    cbsMain.RecalcLayout
End Sub

Private Function CbsSetting(ByRef cbsMain As CommandBars)
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
    With cbsMain.Options
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
    Set cbsMain.Icons = ImageLib.Icons
    cbsMain.ActiveMenuBar.ContextMenuPresent = False    '��ֹ�Ҽ�ѡ�񹤾�����ȡ��
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap  '��ֹ�ƶ�������
End Function


Private Sub vfgReportEdit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zlRefresh(Val(Me.vfgReportEdit.TextMatrix(NewRow, mColR.ID)))
End Sub


Private Sub setListFormat()
    '���ܣ���ʼ�����òο�ֵ�б�
    '������ blnKeepData-�Ƿ������ݣ���ֻ���������ø�ʽ
    With Me.vfgReport
        .Redraw = flexRDNone
        .Clear
        .Rows = 8: .FixedRows = 0: .Cols = 2: .FixedCols = 1
        .TextMatrix(mRow.���, 0) = "���"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.��ʾ, 0) = "��ʾ"
        .TextMatrix(mRow.ԭ��, 0) = "ԭ��"
        .TextMatrix(mRow.��ʩ, 0) = "��ʩ"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.����, 0) = "����"
        .TextMatrix(mRow.�鵵, 0) = "�鵵"
        .ColWidth(0) = 500
        .Redraw = flexRDDirect
    End With
End Sub

Private Function zlRefresh(lngID As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    '�����ǰ����ʾ
    Call setListFormat
    If lngID = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ������Ϣ
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select ���, ����, ��ʾ, ԭ��, ��ʩ, ����, ������, ����ʱ��, �鵵��, �鵵ʱ�� From �����ʿر��� Where ���id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If rsTemp.RecordCount > 0 Then
        With Me.vfgReport
            .Redraw = flexRDNone
            Select Case Val("" & rsTemp!���)
            Case 0: .TextMatrix(mRow.���, 1) = "�ڿأ�"
            Case 1: .TextMatrix(mRow.���, 1) = "���棡"
            Case 2: .TextMatrix(mRow.���, 1) = "ʧ�أ�"
            End Select
            .TextMatrix(mRow.����, 1) = "" & rsTemp!����
            .TextMatrix(mRow.��ʾ, 1) = "" & rsTemp!��ʾ
            .TextMatrix(mRow.ԭ��, 1) = "" & rsTemp!ԭ��
            .TextMatrix(mRow.��ʩ, 1) = "" & rsTemp!��ʩ
            .TextMatrix(mRow.����, 1) = "" & rsTemp!����
            .TextMatrix(mRow.����, 1) = rsTemp!������ & IIf(IsNull(rsTemp!������), "", ", ") & Format(rsTemp!����ʱ��, "yyyy��MM��dd�� hh:mm")
            .TextMatrix(mRow.�鵵, 1) = rsTemp!�鵵�� & IIf(IsNull(rsTemp!�鵵��), "", ", ") & Format(rsTemp!�鵵ʱ��, "yyyy��MM��dd�� hh:mm")
            .Redraw = flexRDDirect
            If rsTemp("�鵵��") & "" <> "" Then
                cbsMain.ActiveMenuBar.Controls.Item(1).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(2).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(3).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(4).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(5).Enabled = True
                
            Else
                cbsMain.ActiveMenuBar.Controls.Item(1).Enabled = True
                cbsMain.ActiveMenuBar.Controls.Item(2).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(3).Enabled = False
                cbsMain.ActiveMenuBar.Controls.Item(4).Enabled = True
                cbsMain.ActiveMenuBar.Controls.Item(5).Enabled = False
            End If
            Call .AutoSize(1)
        End With
    Else
        vfgReport.TextMatrix(mRow.���, 1) = "�ڿأ�"
        vfgReport.TextMatrix(mRow.����, 1) = "δΥ���ʿع���"
        vfgReport.TextMatrix(mRow.��ʾ, 1) = "δʧ��"
        cbsMain.ActiveMenuBar.Controls.Item(1).Enabled = True
        cbsMain.ActiveMenuBar.Controls.Item(2).Enabled = False
        cbsMain.ActiveMenuBar.Controls.Item(3).Enabled = False
        cbsMain.ActiveMenuBar.Controls.Item(4).Enabled = False
        cbsMain.ActiveMenuBar.Controls.Item(5).Enabled = False
         
    End If
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function



Public Sub zlRefReport(strResList As String, lngItemID, strFromDate As String, strToDate As String)
    '���ܣ�ˢ���ʿر���
    '������ strResList  ��ǰѡ����ʿ�Ʒid�����Զ��ŷָ�
    '       lngItemId   ��ǰ��Ŀid
    '       strFromDate ��ʼ����
    '       strToDate   ��������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    Dim lngCount As Long
    Dim strSQL As String
    Err = 0: On Error GoTo ErrHand
    '��ȡʧ�ر���
    strSQL = "Select R.ID,R.������Ŀid, Nvl(T.���, 0) As ���, Q.����ʱ�� As ����, Q.�걾��� As �걾��,D.������ ||'/'||Ӣ���� as ��Ŀ, Zl_lis_ToNumber(Q.�ʿ�Ʒid,R.������Ŀid,R.������,R.id) As ���," & vbNewLine & _
            "       M.���� || ', ' || M.���� As �ʿ�Ʒ, M.ˮƽ, Q.������" & vbNewLine & _
            "From �����ʿؼ�¼ Q, �����ʿ�Ʒ M, ������ͨ��� R, �����ʿر��� T,����������Ŀ D" & vbNewLine & _
            "Where Q.�ʿ�Ʒid = M.ID And Q.�걾id = R.����걾id And R.ID = T.���id(+) And Nvl(R.���ý��,0)=0 And /*Nvl(R.�Ƿ����, 0) = 1 And*/ " & vbNewLine & _
            "      Instr(',' || [1] || ',', ',' || Q.�ʿ�Ʒid || ',') > 0 And R.������Ŀid + 0 = D.ID  And R.������Ŀid + 0 = [2]  And" & vbNewLine & _
            "      R.������ is not null and  (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By Q.����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strResList, lngItemID, strFromDate, strToDate)
    With Me.vfgReportEdit
        .Redraw = flexRDNone
        
        .Clear
        
        Set .DataSource = rsTemp
        Call .AutoSize(mColR.���, .Cols - 1)
        .ColWidth(mColR.ID) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.������Ŀid) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.���) = 280: .TextMatrix(0, mColR.���) = ""
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            Select Case .TextMatrix(lngCount, mColR.���)
                Case 0: Set .Cell(flexcpPicture, lngCount, mColR.���) = Me.imgList.ListImages(2).Picture
                Case 1: Set .Cell(flexcpPicture, lngCount, mColR.���) = Me.imgList.ListImages(1).Picture
                Case 2: Set .Cell(flexcpPicture, lngCount, mColR.���) = Me.imgList.ListImages(3).Picture
            End Select
            .TextMatrix(lngCount, mColR.���) = ""
            If Left(.TextMatrix(lngCount, mColR.���), 1) = "." Then .TextMatrix(lngCount, mColR.���) = "0" & .TextMatrix(lngCount, mColR.���)
        Next
        .Redraw = flexRDDirect
        If .Rows > 1 Then
            Call .Select(1, 1)
        Else
            cbsMain.ActiveMenuBar.Controls.Item(1).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(2).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(3).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(4).Enabled = False
            cbsMain.ActiveMenuBar.Controls.Item(5).Enabled = False
        End If
    End With
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim strGroup As String
    Dim strSQL As String
    Select Case NewRow
    Case mRow.ԭ��: strGroup = "ԭ��"
    Case mRow.��ʩ: strGroup = "��ʩ"
    Case mRow.����: strGroup = "����"
    Case Else: Me.vfgWord.Rows = Me.vfgWord.FixedRows: Exit Sub
    End Select
    
    If OldRow = NewRow Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    strSQL = "Select ���� As ""��ѡ�ʾ�:"" From �ʿر���ʾ� Where ���� Is Null Or ���� = [1] Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strGroup)
    Set Me.vfgWord.DataSource = rsTemp
    Call Me.vfgWord.AutoSize(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgReport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Row
    Case mRow.ԭ��, mRow.��ʩ, mRow.����: Cancel = False
    Case Else: Cancel = True
    End Select
End Sub

Private Sub vfgReport_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub vfgWord_DblClick()
    With Me.vfgReport
        If Me.vfgWord.Row < Me.vfgWord.FixedRows Then Exit Sub
        If Trim(.TextMatrix(.Row, 1)) = "" Then
            .TextMatrix(.Row, 1) = Me.vfgWord.Text
        Else
            .TextMatrix(.Row, 1) = Trim(.TextMatrix(.Row, 1)) & "��" & Me.vfgWord.Text
        End If
        Call .AutoSize(1)
    End With
End Sub
