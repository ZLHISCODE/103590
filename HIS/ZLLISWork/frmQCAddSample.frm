VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmQCAddSample 
   Caption         =   "�ʿر걾"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12765
   Icon            =   "frmQCAddSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   12765
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txt������ 
      Height          =   300
      Left            =   6570
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "������������ʱ��������ţ�ȱʡȡ�����ʿ�Ʒ�����еĶ�Ӧ�걾��"
      Top             =   585
      Width           =   1770
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   585
      Width           =   990
   End
   Begin VB.ComboBox cbo�ʿ�Ʒ 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4980
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   3000
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgSample 
      Height          =   4500
      Left            =   270
      TabIndex        =   0
      Top             =   1410
      Width           =   10785
      _cx             =   19024
      _cy             =   7937
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
      ForeColorSel    =   -2147483632
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cbo���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   3200
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   300
      Left            =   1830
      TabIndex        =   3
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   528351235
      CurrentDate     =   39590
      MaxDate         =   401769
      MinDate         =   2
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   3780
      TabIndex        =   4
      Top             =   585
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   528351235
      CurrentDate     =   39590
      MaxDate         =   401769
      MinDate         =   2
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   180
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmQCAddSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private marrData() As String  '����ԭʼ����
Private mstrPriv As String

Private Sub SetVsFlexGridHead(ByVal strHead As String, ByRef vsGrid As VSFlexGrid)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    
    With vsGrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
         
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
        Next
        
        '�̶������־���
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 350
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Private Sub initCbsThis(cbsMain As CommandBars)
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
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)  '����
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")  '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")

        'Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True '����
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&P)")
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
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����

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

    '���������⴦��
    '-----------------------------------------------------
'    ���˵��Ҳ�Ĳ��� �����￨�Ų��ң�֧��ˢ��
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Dept, "����")
        objControl.ID = conMenu_View_Dept
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Dept + 1, "")
        objCustom.Handle = cbo����.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
        Set objControl = .Add(xtpControlLabel, conMenu_View_FindType, "�ʿ�Ʒ")
        objControl.ID = conMenu_View_FindType
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = cbo�ʿ�Ʒ.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ") '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��") '����

        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "����"):
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")

        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True '����
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�") '����
        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 1, "����")
        objControl.ID = conMenu_EditPopup + 1
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 2, "")
        objCustom.Handle = dtpStart.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 3, "��")
        objControl.ID = conMenu_EditPopup + 3
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 4, "")
        objCustom.Handle = dtpEnd.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 5, "�������")
        objControl.ID = conMenu_EditPopup + 5
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 6, "")
        objCustom.Handle = cbo�������.hWnd
        objCustom.Flags = xtpFlagRightAlign
        
        Set objControl = .Add(xtpControlLabel, conMenu_EditPopup + 7, "�ʿ�������")
        objControl.ID = conMenu_EditPopup + 7
        objControl.Flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_EditPopup + 8, "")
        objCustom.Handle = txt������.hWnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings

        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ

        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '��ӡ����
        .AddHiddenCommand conMenu_File_Excel            '�����Excel
    End With

    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
'    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub
Private Sub reSetHead(ByVal dateStart As Date, ByVal dateEnd As Date)
    '��ʼ��vfgSample�ؼ�����
    Dim strHead As String, lng���� As Long
    Dim i As Integer
    
    lng���� = DateDiff("d", dateStart, dateEnd)
    For i = 0 To lng����
        strHead = strHead & ";" & Format(dateStart + i, "yyyy-MM-dd") & ",1300,7"
    Next
    strHead = "������Ŀ,1200,1" & strHead & ";��ĿID,0,1;����,0,1;����,0,1"
    Call SetVsFlexGridHead(strHead, vfgSample)

End Sub

Private Sub RefreshData()
    Dim lng�ʿ�ID As Long, int������� As Integer
    Dim dateStart As Date, dateEnd As Date
    Dim i As Integer
    
    Dim strsql As String, strTmpSQL As String
    
    Dim rsTmp As ADODB.Recordset
    dateStart = Format(dtpStart.Value, "yyyy-MM-dd")
    dateEnd = Format(dtpEnd.Value, "yyyy-MM-dd")
    If dateStart > dateEnd Then
        MsgBox "��ʼ���ڲ��ܴ��ڽ�������!", vbQuestion, Me.Caption
        Exit Sub
    End If
    Call reSetHead(dateStart, dateEnd)
    ReDim marrData(vfgSample.Rows, vfgSample.Cols)
    
    If cbo�ʿ�Ʒ.ListIndex < 0 Then Exit Sub
    
    lng�ʿ�ID = cbo�ʿ�Ʒ.ItemData(cbo�ʿ�Ʒ.ListIndex)
    If lng�ʿ�ID <= 0 Then Exit Sub


    strsql = "Select ��ʼ����,�������� From �����ʿ�Ʒ Where  id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng�ʿ�ID, dateEnd)
    Do Until rsTmp.EOF
        If dateStart < rsTmp!��ʼ���� Or dateEnd > rsTmp!�������� Then
            MsgBox "�����趨�����ʿ�ƷЧ���ڣ�������ָ�����ڣ�", vbInformation, Me.Caption
            Exit Sub
        End If
        rsTmp.MoveNext
    Loop
    
    int������� = cbo�������.List(cbo�������.ListIndex)
    '------------- ������
    Dim intRow As Integer, intFindRow As Integer
    
    On Error GoTo ErrHandle
    strsql = "Select �걾�� From �����ʿ�Ʒ where id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng�ʿ�ID)
    Do Until rsTmp.EOF
        If "" & rsTmp!�걾�� <> "" Then
            txt������ = "" & rsTmp!�걾��
        End If
        rsTmp.MoveNext
    Loop
    
    '--- �ӿհ�����Ŀ
        
    
    strsql = "Select A.�ʿ�Ʒid, A.��Ŀid, A.ȡֵ����, A.����ֵ, E.�������, F.����, F.������, E.��д" & vbNewLine & _
            "From �����ʿ�Ʒ��Ŀ A, ������Ŀ E, ����������Ŀ F" & vbNewLine & _
            "Where A.��Ŀid = E.������Ŀid And A.��Ŀid = F.ID And A.�ʿ�Ʒid = [1]" & vbNewLine & _
            "Order By F.����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng�ʿ�ID)
    With vfgSample
        .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = "�걾��"
        .Rows = .Rows + 1
        Do Until rsTmp.EOF
           
            .TextMatrix(.Rows - 1, .ColIndex("��ĿID")) = "" & rsTmp!��Ŀid
            .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = "" & rsTmp!������ & " " & rsTmp!��д
            .TextMatrix(.Rows - 1, .ColIndex("����")) = "" & rsTmp!�������
            .TextMatrix(.Rows - 1, .ColIndex("����")) = "" & rsTmp!ȡֵ����
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Trim(.TextMatrix(.Rows - 1, 0)) = "" Then .Rows = .Rows - 1
        
        ReDim marrData(vfgSample.Rows, vfgSample.Cols)
    End With
    
    '---- ȡ��������
    strsql = "Select A.*, E.�������, F.����, F.������, E.��д, D.������, T.���" & vbNewLine & _
            "From (Select A.�ʿ�Ʒid, A.��Ŀid,C.�걾���, B.�걾id, B.����ʱ��, A.ȡֵ����, A.����ֵ" & vbNewLine & _
            "       From �����ʿ�Ʒ��Ŀ A, �����ʿؼ�¼ B,����걾��¼ C" & vbNewLine & _
            "       Where B.�걾id=C.ID And A.�ʿ�Ʒid = B.�ʿ�Ʒid And A.�ʿ�Ʒid = [1] And" & vbNewLine & _
            "             B.����ʱ�� Between [2] And [3] And B.���Դ���=[4]) A," & vbNewLine & _
            "     ������ͨ��� D, ������Ŀ E, ����������Ŀ F,�����ʿر��� T" & vbNewLine & _
            "Where D.ID=T.���ID(+) And A.�걾id = D.����걾id And A.��Ŀid = D.������Ŀid And A.��Ŀid = E.������Ŀid And A.��Ŀid = F.ID" & vbNewLine & _
            "Order By A.����ʱ��, F.����"
    dateEnd = dateEnd + 1
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng�ʿ�ID, CDate(dateStart), CDate(dateEnd), int�������)
    
    With vfgSample
        'intRow = .FixedRows
        Do Until rsTmp.EOF
            intFindRow = .FindRow("" & rsTmp!��Ŀid, .FixedRows, .ColIndex("��ĿID"))
            If intFindRow > 0 Then
                intRow = intFindRow
            Else
                intRow = .Rows - 1
            End If
            
            .TextMatrix(intRow, .ColIndex("��ĿID")) = "" & rsTmp!��Ŀid
            .TextMatrix(intRow, .ColIndex("������Ŀ")) = "" & rsTmp!������ & " " & rsTmp!��д
            
            For i = 1 To .Cols - 1
                If Format("" & rsTmp!����ʱ��, "yyyy-MM-dd") = .TextMatrix(0, i) Then
                    .TextMatrix(intRow, i) = "" & rsTmp!������
                    marrData(intRow, i) = "" & rsTmp!������ & "|" & rsTmp!�걾ID
                    If Val("" & rsTmp!���) = 2 Then 'ʧ��(��)
                        .Cell(flexcpForeColor, intRow, i) = vbRed
                    ElseIf Val("" & rsTmp!���) = 0 Then '����
                        .Cell(flexcpForeColor, intRow, i) = .ForeColor
                    Else  '����(���)
                        .Cell(flexcpForeColor, intRow, i) = vbMagenta
                    End If
                    If Val("" & rsTmp!�걾ID) > 0 Then
                        .TextMatrix(.FixedRows, i) = "" & rsTmp!�걾���
                        marrData(.FixedRows, i) = "" & rsTmp!�걾���
                    End If
                    Exit For
                End If
            Next
            If Not (intFindRow = intRow And intFindRow > 0) Then
                intRow = intRow + 1
                .Rows = .Rows + 1
            End If
            

            rsTmp.MoveNext
        Loop
        
        '��ȱʡ�걾��
        For i = 1 To .Cols - 1
            If IsDate(.TextMatrix(0, i)) Then
                If Val(.TextMatrix(.FixedRows, i)) = 0 And Val(txt������) <> 0 Then
                    .TextMatrix(.FixedRows, i) = Val(txt������)
                End If
            End If
        Next
        If Trim(.TextMatrix(.Rows - 1, 0)) = "" Then .Rows = .Rows - 1
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize .ColIndex("������Ŀ")
        '.FrozenCols = 1
        .AllowUserFreezing = flexFreezeColumns
        
        .Editable = flexEDNone
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub

Private Sub Load����()
    Dim strsql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If InStr(1, mstrPriv, "���п���") > 0 Then
        strsql = " Select Distinct  a.id,a.���� , a.����  From �������� a ,���ű� b,�����ʿ�Ʒ c " & _
                  "Where a.ʹ��С��ID = b.ID and a.id = c.����id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName)
        
    Else
        strsql = " Select Distinct a.id,a.���� , a.����  From ������Ա D,�������� a ,���ű� b , �����ʿ�Ʒ c " & _
                  " Where a.ʹ��С��ID = b.ID and a.ʹ��С��id=D.����id and D.��Աid = [1]  " & _
                  " and a.id = c.����Id "
        Set rsTemp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, UserInfo.ID)
    End If
    
    cbo����.Clear
    Do Until rsTemp.EOF
        cbo����.AddItem "" & rsTemp!���� & " " & rsTemp!����
        cbo����.ItemData(cbo����.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    If cbo����.ListCount > 0 Then cbo����.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SaveData()
    '��������
    Dim intRow As Integer, intCol As Integer
    Dim strData As String, strOLDdata As String
    Dim lng��ĿID As Long, str���� As String
    Dim lng�걾ID As Long, bln��ʾ�걾�� As Boolean
    Dim strNewItem As String '��������������Ŀ
    Dim str�걾�� As String
    
    bln��ʾ�걾�� = False
    
    If Val(cbo�������.List(cbo�������.ListIndex)) > 1 Then
        MsgBox "ע�⣺���û����д��" & Val(cbo�������.List(cbo�������.ListIndex)) - 1 & "�ε�����,��ô�������ݵĴ������ǵ�" & Val(cbo�������.List(cbo�������.ListIndex)) - 1 & "�Σ�", vbInformation, Me.Caption
    End If
    
    For intCol = vfgSample.ColIndex("������Ŀ") + 1 To vfgSample.ColIndex("��ĿID") - 1
        strNewItem = ""
        str���� = vfgSample.TextMatrix(0, intCol)
        lng�걾ID = 0
        str�걾�� = Val(vfgSample.TextMatrix(1, intCol))
        If str�걾�� = "0" Then str�걾�� = Val(txt������)
        
        For intRow = 2 To vfgSample.Rows - 1
            strData = vfgSample.TextMatrix(intRow, intCol)
            strOLDdata = marrData(intRow, intCol)
            
            If InStr(strOLDdata, "|") > 0 Then
                lng�걾ID = Split(strOLDdata, "|")(1)
            End If
            
            If strOLDdata <> "" Then
                If strData <> Split(strOLDdata, "|")(0) Then
                    'Ҫ����
                    If InStr(strOLDdata, "|") > 0 Then
                        '��ԭʼ��¼
                        
                        lng��ĿID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("��ĿID")))
                        strNewItem = strNewItem & "|" & lng��ĿID & "^" & strData
                    Else
                        '����
                        lng��ĿID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("��ĿID")))
                        strNewItem = strNewItem & "|" & lng��ĿID & "^" & strData
                        
                        If Val(str�걾��) = 0 Then bln��ʾ�걾�� = True
                    End If
                End If
            Else
                If strData <> "" Then
                    '����
                    lng��ĿID = Val(vfgSample.TextMatrix(intRow, vfgSample.ColIndex("��ĿID")))
                    strNewItem = strNewItem & "|" & lng��ĿID & "^" & strData
                    
                    If Val(str�걾��) = 0 Then bln��ʾ�걾�� = True
                End If
            End If
        Next
        If bln��ʾ�걾�� Then
            MsgBox "����д�걾�ţ�", vbInformation, Me.Caption
            Exit Sub
        End If
        If strNewItem <> "" Then
            strNewItem = Mid(strNewItem, 2)
            Call Edit_Sample(lng�걾ID, str����, strNewItem, str�걾��)
        End If
    Next
    
    Call RefreshData
End Sub

Private Sub Edit_Sample(ByVal lngID_in As Long, _
                        ByVal str���� As String, ByVal strItemRecords As String, ByVal str�걾�� As String)
    '�����ʿر걾
    Dim lngID As Long       '�걾id
    Dim lngDeviceID As Long '����id
    Dim strSampleNO As String '�걾��
    Dim lngQCID As Long '�ʿ�ƷID
    
    Dim blnTrans As Boolean '�Ƿ�ʼ����
    On Error GoTo ErrHandle
    
    If lngID_in = 0 Then
        lngID = zlDatabase.GetNextId("����걾��¼")
    Else
        lngID = lngID_in
    End If
    
    strSampleNO = str�걾��
    lngDeviceID = cbo����.ItemData(cbo����.ListIndex)
    lngQCID = cbo�ʿ�Ʒ.ItemData(cbo�ʿ�Ʒ.ListIndex)
    
'    gcnOracle.BeginTrans
'    blnTrans = True
    If lngID_in = 0 Then
        gstrSql = "ZL_����걾��¼_INSERT(" & lngID & ",NULL,'" & _
            strSampleNO & "',NULL,NULL," & lngDeviceID & ",NULL," & _
            "To_Date('" & str���� & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
            "To_Date('" & str���� & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.���� & "'," & _
            "Null,To_Date('" & str���� & "','yyyy-mm-dd hh24:mi:ss'),'" & gstrUserName & "','0',Null,0,Null)"
        zlDatabase.ExecuteProcedure gstrSql, "���������ʱ��¼"
    End If
    
    gstrSql = "ZL_������ͨ���_BATCHUPDATE(" & lngID & "," & _
        lngDeviceID & ",Null,Null,Null,'" & strItemRecords & "')"
    zlDatabase.ExecuteProcedure gstrSql, "����������"
    
    gstrSql = "ZL_�����ʿؼ�¼_EDIT(1," & lngID & "," & lngQCID & ")"
    zlDatabase.ExecuteProcedure gstrSql, "����Ϊ�ʿ�Ʒ"
    
'    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
ErrHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

'------ �����ǿؼ�����

Private Sub cbo�������_Click()

    Call RefreshData
    If Val(cbo�������.List(cbo�������.ListIndex)) > 1 Then
        txt������ = Val(txt������.Text) + Val(cbo�������.List(cbo�������.ListIndex)) - 1
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Me.vfgSample.Editable = flexEDKbdMouse
    Case conMenu_Edit_Untread
        Call RefreshData
    Case conMenu_Edit_Save
        Call SaveData
    Case conMenu_View_Refresh
        Call RefreshData
    Case conMenu_File_Exit
        Unload Me
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_Modify
        Control.Enabled = Not (Me.vfgSample.Editable = flexEDKbdMouse)
    Case conMenu_Edit_Save, conMenu_Edit_Untread
        Control.Enabled = (Me.vfgSample.Editable = flexEDKbdMouse)
    End Select

End Sub

Private Sub dtpEnd_Change()
Call RefreshData
End Sub

Private Sub dtpStart_Change()
Call RefreshData
End Sub

Private Sub Form_Load()
    
    Call initCbsThis(cbsThis)
    
    '�����ڼ��������
    dtpStart = Now - 5
    dtpEnd = Now
    Call reSetHead(dtpStart.Value, dtpEnd.Value)
    
    cbo�������.Clear
    cbo�������.AddItem "1"
    cbo�������.AddItem "2"
    cbo�������.AddItem "3"
    cbo�������.AddItem "4"
    cbo�������.AddItem "5"
    cbo�������.AddItem "6"
    cbo�������.AddItem "7"
    cbo�������.AddItem "8"
    cbo�������.AddItem "9"
    
    cbo�������.ListIndex = 0
    
    Call Load����
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Call cbsThis_Resize
End Sub

Private Sub cbo�ʿ�Ʒ_Click()
    Call RefreshData
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With vfgSample
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbo����_Click()
    Dim strsql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng����id As Long
    Dim dateStart As Date
    Dim dateEnd As Date
    
    On Error GoTo ErrHandle
    If cbo����.ListIndex < 0 Then Exit Sub
    
    lng����id = cbo����.ItemData(cbo����.ListIndex)
    dateStart = Format(dtpStart.Value, "yyyy-MM-dd")
    dateEnd = Format(dtpEnd.Value, "yyyy-MM-dd")
    strsql = "Select ID,����,����,Ũ��,ˮƽ From �����ʿ�Ʒ Where [2] between ��ʼ���� and �������� and [3] between ��ʼ���� and���������� and ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, lng����id, dateStart, dateEnd)
    cbo�ʿ�Ʒ.Clear
    Do Until rsTmp.EOF
        cbo�ʿ�Ʒ.AddItem "" & rsTmp!���� & " " & rsTmp!���� & " ˮƽ:" & rsTmp!ˮƽ
        cbo�ʿ�Ʒ.ItemData(cbo�ʿ�Ʒ.NewIndex) = rsTmp!ID
        
        rsTmp.MoveNext
    Loop
    If cbo�ʿ�Ʒ.ListCount > 0 Then cbo�ʿ�Ʒ.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(ByVal strPrivate As String, ByVal frmMain As Form)
    mstrPriv = strPrivate
    
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgSample_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    Dim lngCount As Long
    
    If Col = 0 Then Exit Sub
    If Trim(Me.vfgSample.TextMatrix(Row, Col)) = "" Then Exit Sub
    
    strLists = Trim(Me.vfgSample.TextMatrix(Row, vfgSample.ColIndex("����")))
    strValue = Trim(Me.vfgSample.TextMatrix(Row, Col))
    
    If strLists = "" Then
        If Row = 1 Then
            Me.vfgSample.TextMatrix(Row, Col) = CLng(Val(strValue)) '�걾���У�������С��
        Else
            If InStr(strValue, "E+") > 0 And Val(strValue) > 0 Then
                Me.vfgSample.TextMatrix(Row, Col) = strValue
            Else
                Me.vfgSample.TextMatrix(Row, Col) = Format(Val(strValue), "0.00")
            End If
        End If
        Exit Sub
    End If
    For lngCount = 0 To UBound(Split(strLists, ";"))
        If vfgSample = Split(strLists, ";")(lngCount) Then Exit Sub
    Next
    Me.vfgSample.TextMatrix(Row, Col) = ""
    
    strValue = "����ĿΪ�붨����Ŀ�������ȡֵ����(" & strLists & ")Ҫ��"
    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgSample_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
    If Row = 1 Then '�걾���У���ԭ���б걾�ţ����ܸ�
        If marrData(Row, Col) <> "" Then Cancel = True
    End If
End Sub
