VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm������ҩ���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ҩ����"
   ClientHeight    =   7644
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8424
   Icon            =   "frm������ҩ����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7644
   ScaleWidth      =   8424
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraInput 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8175
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "¼��NO�����������￨�����б��в��Ҳ���λ"
         Height          =   180
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   3600
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFWindows 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8175
      _cx             =   14420
      _cy             =   11033
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm������ҩ����.frx":1A72
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
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm������ҩ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngҩ��ID As Long
Private mrsList As Recordset
Private mstrWin As String
Private mdate��ʼ���� As Date
Private mdate�������� As Date
Private mstr���� As String
Private mstrSourceDep As String
Private mstrDeptNode As String
Private mstrCurrentWin As String           '��ǰ����
Private mintShowBill�շ� As Integer
Private mintShowBill���� As Integer
Private mblnOper As Boolean     '�Ƿ�ִ���˴��ڵ���

'�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
Private mstrUserRecipeColor As String
Private Sub InitComman()
'--------------------------------------
'��ʼ��CommandBars1�ؼ�

'--------------------------------------
    With CommandBarsGlobalSettings
        Set .App = App
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����������ϵͳ�Զ�ʶ��
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With Me.cbsMain
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        .Item(1).Delete
        .Icons = frmPublic.imgPublic.Icons
    End With
End Sub
Private Sub InitTool()
'-----------------------------------------------------
'���ù�����
'----------------------------------------------------
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    Set objBar = cbsMain.Add("������1", xtpBarTop)
    objBar.ContextMenuPresent = False '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_Guide, "��������")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_Average, "ƽ������")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        objControl.Visible = (mstrCurrentWin <> "")

        Set objControl = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��")
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        
        Set objControl = .Add(xtpControlButton, mconMenu_Edit_Recipe_OK, "ȷ��")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        
        Set objControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
        objControl.Style = xtpButtonIconAndCaption 'ͬʱ��ʾͼ�������
        objControl.BeginGroup = True
        
    End With
End Sub


Private Sub Init����()
    Dim strsql As String
    Dim rsRecord As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select ����,���� from ��ҩ���� where ҩ��id=[1] and �ϰ��=1"
    If mstrCurrentWin <> "" Then strsql = strsql & " And ����<>[2] "
    Set rsRecord = zldatabase.OpenSQLRecord(strsql, "Init����", mlngҩ��ID, mstrCurrentWin)
    
    If Not (rsRecord Is Nothing) Then
        Do While Not rsRecord.EOF
            mstrWin = mstrWin & rsRecord!���� & "|"
            rsRecord.MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub SetAverageWindows()
    'ƽ�����䷢ҩ���ڣ������ڴ����°�ʱ�������ڴ���ƽ�����䵽�����ϰ�Ĵ���
    '˼·��1�������ݺ����򣬲�������IDͳ�Ʊ������ж�����
    '2��ͳ���м������ϰ�ķ�ҩ���ڣ�������ƽ������
    '3��ͬһ������ȷ�����䵽ͬһ������
    Dim strPatis As String      '�������������μ�¼������Ϣ����������+����id���������ͬ�����޲���id������Ϊͬһ�ˣ�
    Dim arrPaits, arrWins, arrPatiWins
    Dim i As Integer, lngRows As Long
    Dim intWinCount As Integer  '���ϰര������
    Dim strWinPati As String    '���ںͲ��˶�Ӧ��ϵ
    
    mrsList.Filter = ""
    mrsList.Sort = "����,NO"
    If mrsList.RecordCount = 0 Then Exit Sub
    
    '����Ѿ�û���ϰര���˾��˳�
    If mstrWin = "" Then Exit Sub
    
    '��NO������¼����ID�����ظ�
    With mrsList
        Do While Not .EOF
            If InStr(1, "|" & strPatis & "|", "|" & !���� & !����ID & "|") = 0 Then
                strPatis = IIf(strPatis = "", "", strPatis & "|") & !���� & !����ID
            End If
            
            .MoveNext
        Loop
    End With
    
    arrPaits = Array()
    arrPaits = Split(strPatis, "|")
    
    If Right(mstrWin, 1) = "|" Then mstrWin = Mid(mstrWin, 1, Len(mstrWin) - 1)
    arrWins = Array()
    arrWins = Split(mstrWin, "|")
    intWinCount = UBound(arrWins) + 1
    
    'ƽ�����䴰�ڣ���¼���ںͲ���ID��ϵ
    For i = 0 To UBound(arrPaits)
        strWinPati = IIf(strWinPati = "", "", strWinPati & "|") & arrPaits(i) & "," & arrWins(IIf(((i + 1) Mod intWinCount) = 0, intWinCount - 1, (i + 1) Mod intWinCount - 1))
    Next
    
    arrPatiWins = Array()
    arrPatiWins = Split(strWinPati, "|")
    
    '�����´���
    With VSFWindows
        For lngRows = 1 To .rows - 1
            For i = 0 To UBound(arrPatiWins)
                If .TextMatrix(lngRows, .ColIndex("����")) & .TextMatrix(lngRows, .ColIndex("����id")) = Split(arrPatiWins(i), ",")(0) Then
                    .TextMatrix(lngRows, .ColIndex("�´���")) = Split(arrPatiWins(i), ",")(1)
                    Exit For
                End If
            Next
        Next
    End With
    
End Sub

Public Function showMe(ByVal lngҩ��ID As Long, ByVal FrmMain As Form, ByVal date��ʼ���� As Date, ByVal date�������� As Date, _
    ByVal strDeptNode As String, Optional ByVal strCurrentWin As String) As Boolean
    'lngҩ��ID����ǰҩ��
    'FrmMain��������
    'date��ʼ����,date�������ڣ������������ڷ�Χ
    'strDeptNode��ҩ��վ��
    'strCurrentWin����ǰ����
    mlngҩ��ID = lngҩ��ID
    mdate��ʼ���� = date��ʼ����
    mdate�������� = date��������
    mstrDeptNode = strDeptNode
    mstrCurrentWin = strCurrentWin
    
    Call frm������ҩ����.Show(1, FrmMain)
    
    showMe = mblnOper
End Function

Private Sub InitVSFGrid()
    With Me.VSFWindows
        .rows = 1
        
        .ColComboList(.ColIndex("�´���")) = mstrWin
        .ComboList = mstrWin
    End With
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_Refresh  'ִ��ˢ�²���
            Me.VSFWindows.rows = 1
            LoadData
            InitGridData
        Case mconMenu_File_Exit      'ִ���˳�����
            Unload Me
        Case mconMenu_Edit_Recipe_Guide  'ִ������������
            Edit_Recipe_Guide
        Case mconMenu_Edit_Recipe_OK     'ִ��ȷ������
            Call Edit_Recipe_OK
            Me.VSFWindows.rows = 1
            LoadData
            InitGridData
        Case mconMenu_Edit_Recipe_Average   'ƽ������
            Call SetAverageWindows
    End Select
End Sub
Private Sub Edit_Recipe_Guide()
    Dim strConWin As String
    Dim str���� As String
    Dim str�´��� As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    frm����������.showMe mlngҩ��ID, strConWin, str����, str�´���, mstrCurrentWin
    
    For i = -1 To UBound(Split(strConWin, ","))
        If i = UBound(Split(strConWin, ",")) - 1 Then Exit For
        
        For j = -1 To UBound(Split(str����, ","))
            If j = UBound(Split(str����, ",")) - 1 Then Exit For
            
            If UBound(Split(strConWin, ",")) <> -1 Or UBound(Split(str����, ",")) <> -1 Then
                For k = 1 To Me.VSFWindows.rows - 1
                    If UBound(Split(strConWin, ",")) = -1 Then
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("���")) = Split(str����, ",")(j + 1) Then
                            Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("�´���")) = str�´���
                        End If
                    ElseIf UBound(Split(str����, ",")) = -1 Then
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("�ִ���")) = Split(strConWin, ",")(i + 1) Then
                            Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("�´���")) = str�´���
                        End If
                    Else
                        If Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("�ִ���")) = Split(strConWin, ",")(i + 1) And Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("���")) = Split(str����, ",")(j + 1) Then
                             Me.VSFWindows.TextMatrix(k, VSFWindows.ColIndex("�´���")) = str�´���
                        End If
                    End If
                Next
            End If
        Next
    Next
End Sub
Private Sub Edit_Recipe_OK()
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    
    With VSFWindows
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("�ִ���")) <> .TextMatrix(i, .ColIndex("�´���")) And .TextMatrix(i, .ColIndex("�´���")) <> "" Then
                gstrSQL = "zl_δ��ҩƷ��¼_���䷢ҩ����("
                gstrSQL = gstrSQL & "'" & .TextMatrix(i, .ColIndex("NO")) & "'"
                gstrSQL = gstrSQL & "," & .TextMatrix(i, .ColIndex("����"))
                gstrSQL = gstrSQL & "," & mlngҩ��ID
                gstrSQL = gstrSQL & ",'" & .TextMatrix(i, .ColIndex("�´���")) & "')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
    End With
    
    On Error GoTo errHandle
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "Edit_Recipe_OK")
    Next
    gcnOracle.CommitTrans
    
    mblnOper = True
    
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnOper = False
    
    mstrUserRecipeColor = zldatabase.GetPara("������ɫ", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    InitComman
    
    InitTool
    
    Init����
    
    InitVSFGrid
    
    LoadData
    
    InitGridData
End Sub

Private Sub LoadData()
    Dim strsql As String
    Dim str���� As String
    Dim strSub1 As String
    Dim strSub2 As String
    Dim strסԺ As String
    
    On Error GoTo errHandle
    gstrSQL = "Select /*+ Rule*/ '' As ��ɫ, �������� ,'' As ѡ�� ,'0' As ��־,����,����,���շ�,��ҩ��,NO,����,��ҩ����," & _
            " ���￨��,�����,���֤��,IC����,����ID,ҽ����,סԺ��," & _
            " �����־,��¼����,Decode(A.�շ����, '7', '7', '0') As �շ����,���� " & _
            " From ("
            
    str���� = " Select A.��ҩ����,A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.NO,A.����,C.���۽��,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,d.ʵ�ս��, Nvl(A.��������,Nvl(C.ע��֤��,0)) As ��������,D.�����־,D.��¼����,D.�շ���� " & _
            " From " & _
            " (Select distinct A.��ҩ����,B.���￨��,B.�����,B.���֤��,B.IC����,B.ҽ����,B.סԺ��,A.���ȼ�,A.��������,Decode(Nvl(A.���շ�,0),1,'','(δ)')||Decode(A.����,8,'�շ�',9,'����') ����,A.����,A.���շ�,'' ��ҩ��,A.No,A.����,To_Char(A.��������,'yyyy-MM-dd hh24:mi:ss') ����,1 �ɲ���,' ' ˵��,B.����ID, A.��������,a.�Է�����id " & _
            " From δ��ҩƷ��¼ A,������Ϣ B " & _
            " Where 1=1 "

    '��Ҫ����
    str���� = str���� & " And (A.�ⷿID=[1] Or A.�ⷿID Is NULL) And A.�������� Between [2] And [3] "
    
    
    str���� = str���� & " And A.����ID=B.����ID(+)"
    
    
    str���� = str���� & " And A.���� IN(8,9)"
    
    If mstrCurrentWin <> "" Then str���� = str���� & " And A.��ҩ����=[5] "
        
    str���� = str���� & ") A,ҩƷ�շ���¼ C, ������ü�¼ D, ���ű� B " & _
              " Where C.����id = D.ID And nvl(c.��ҩ��ʽ,-999)<>-1 and A.����=C.���� And A.NO=C.NO And C.����� Is NULL " & _
              " And Nvl(D.����״̬,0)<>1 And (C.�ⷿid=[1] Or C.�ⷿid Is null)  And a.�Է�����id = b.Id "
    
    If mstrDeptNode <> "" Then
        str���� = str���� & " And (b.վ�� = [4] Or b.վ�� Is Null) "
    End If
    
    strסԺ = Replace(str����, "������ü�¼", "סԺ���ü�¼")
    strסԺ = Replace(strסԺ, "And Nvl(D.����״̬,0)<>1", "")
    
    '���ﻮ�ۼ��������
    gstrSQL = gstrSQL & str���� & "Union All " & strסԺ
    
    
    '''''Group By
    gstrSQL = gstrSQL & ") A GROUP BY A.��ҩ����,A.���ȼ�,A.����,A.����,A.���շ�,A.��ҩ��,A.No,A.����,A.����,A.�ɲ���,A.˵��,A.���￨��,A.�����,A.���֤��,A.IC����,A.����ID,A.ҽ����,A.סԺ��,A.��������,A.�����־,A.��¼����,Decode(A.�շ����, '7', '7', '0'),A.���� "
    
    '''''Order By
    gstrSQL = gstrSQL & " Order By A.����,A.����,A.����id,A.No"
    
    Set mrsList = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            mlngҩ��ID, _
            mdate��ʼ����, _
            zldatabase.Currentdate, _
            mstrDeptNode, _
            mstrCurrentWin)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitGridData()
    Dim i As Integer
    
    With Me.VSFWindows
        .Redraw = flexRDNone
        
        .rows = .rows + mrsList.RecordCount
        For i = 1 To mrsList.RecordCount
            If mrsList!�������� = 1 Then
                .TextMatrix(i, .ColIndex("����")) = "����"
            ElseIf mrsList!�������� = 2 Then
                .TextMatrix(i, .ColIndex("����")) = "����"
            ElseIf mrsList!�������� = 3 Then
                .TextMatrix(i, .ColIndex("����")) = "����"
            ElseIf mrsList!�������� = 4 Then
                .TextMatrix(i, .ColIndex("����")) = "��һ"
            ElseIf mrsList!�������� = 5 Then
                .TextMatrix(i, .ColIndex("����")) = "����"
            Else
                .TextMatrix(i, .ColIndex("����")) = "��ͨ"
            End If
            
            .TextMatrix(i, .ColIndex("���")) = IIf(IsNull(mrsList!����), "", mrsList!����)
            .TextMatrix(i, .ColIndex("NO")) = mrsList!NO
            .TextMatrix(i, .ColIndex("����")) = Format(mrsList!����, "YYYY-MM-DD HH:MM")
            .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(mrsList!����), "", mrsList!����)
            .TextMatrix(i, .ColIndex("�ִ���")) = zlStr.Nvl(mrsList!��ҩ����)
            .TextMatrix(i, .ColIndex("����")) = mrsList!����
            .TextMatrix(i, .ColIndex("���￨��")) = zlStr.Nvl(mrsList!���￨��)
            .TextMatrix(i, .ColIndex("����ID")) = zlStr.Nvl(mrsList!����ID)
            
            .Cell(flexcpBackColor, i, .ColIndex("����"), i, .ColIndex("����")) = Val(Split(mstrUserRecipeColor, ";")(Val(mrsList!��������)))
            
            mrsList.MoveNext
        Next
        
        .Cell(flexcpFontBold, 0, .ColIndex("�´���"), 0, .ColIndex("�´���")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("�´���"), 0, .ColIndex("�´���")) = vbBlue
        
        If .rows > 1 Then
            .Cell(flexcpBackColor, 1, .ColIndex("�´���"), .rows - 1, .ColIndex("�´���")) = &HFFE3C8    '&HFFEDDD     '&HFFC0C0
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.fraInput.Move 100, 490, Me.ScaleWidth - 200, Me.fraInput.Height
    Me.VSFWindows.Move 100, Me.fraInput.Top + Me.fraInput.Height, Me.ScaleWidth - 200, Me.ScaleHeight - (Me.fraInput.Top + Me.fraInput.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrWin = ""
End Sub

Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim i As Integer
    
    If KeyCode <> 13 Then Exit Sub
    strText = Trim(Me.txtMsg.Text)
    
    For i = 1 To Me.VSFWindows.rows - 1
        If InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("NO")), strText) <> 0 Or InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("����")), strText) Or InStr(1, Me.VSFWindows.TextMatrix(i, VSFWindows.ColIndex("���￨��")), strText) Then
            Me.VSFWindows.Row = i
            Exit Sub
        End If
    Next
End Sub

Private Sub VSFWindows_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> VSFWindows.ColIndex("�´���") Then
        Cancel = True
    End If
End Sub


