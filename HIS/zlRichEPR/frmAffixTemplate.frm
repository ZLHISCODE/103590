VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmAffixTemplate 
   Caption         =   "����ģ������"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12555
   Icon            =   "frmAffixTemplate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   12555
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vfgAffix 
      Height          =   6000
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   3870
      _cx             =   6826
      _cy             =   10583
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
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769985
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   285
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6945
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAffixTemplate.frx":06EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17066
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vfgTemplate 
      Height          =   5970
      Left            =   3990
      TabIndex        =   2
      Top             =   855
      Width           =   8520
      _cx             =   15028
      _cy             =   10530
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
      BackColorSel    =   16769985
      ForeColorSel    =   0
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   285
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
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
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
      BackColorFrozen =   -2147483630
      ForeColorFrozen =   -2147483630
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   1110
      Top             =   225
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAffixTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�˵�����ö�ٶ���
Private Enum TMenuType
    mtFile = 1
    mtSave
    mtCancel
    mtImport
    mtExport
    mtQuit
    
    mtEdit
    mtNew
    mtDel
    mtClearCount
End Enum


'��״̬
Private Enum TRowState
    rsNormal = 0
    rsNew
    rsDel
    rsModify
End Enum


'����ģ���б���ж���
Private Enum TAffixTemplateCol
    atcState = 0
    atcTitle = 1
    atcCount = 2
    atcContext = 3
End Enum

Private mrsAffixTemplate As ADODB.Recordset

Private mblnModifyState As Boolean      '�޸�״̬
Private mlngRequestPageId As Long       '��������Id

Private mstrStartEditText As String

Public Sub ShowAffixConfig(ByVal lngRequestPageId As Long, objOwner As Object)
    mlngRequestPageId = lngRequestPageId
    Me.Show 1, objOwner
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo ErrHandle
    zlMailTo hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo ErrHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo ErrHandle
    zlHomePage hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).STYLE
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.STYLE = intStyle
        Next
    Next
    
    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub CancelModify()
'�����޸�
    vfgTemplate.Row = vfgTemplate.Row
    
    If MsgBox("�Ƿ����������޸ģ�", vbYesNo, Me.Caption) = vbNo Then Exit Sub
     
    Call LoadAffixTemplate(mlngRequestPageId, vfgAffix.Text)
     
    mblnModifyState = False
End Sub


Private Sub NewTemplate()
'����ģ��
    vfgTemplate.Rows = vfgTemplate.Rows + 1
    vfgTemplate.TextMatrix(vfgTemplate.Rows - 1, atcCount) = 0
    
    vfgTemplate.Col = 1
    vfgTemplate.Row = vfgTemplate.Rows - 1
    
    '����������״̬
    Call SetRowDataState(vfgTemplate.Row, rsNew)
    
    vfgTemplate.EditCell
    
    mblnModifyState = True
End Sub

Private Sub DelTemplate()
'ɾ��ģ��
    Dim lngNextRow As Long
    Dim i As Long
    
    If vfgTemplate.Rows <= 1 Then Exit Sub
    
    lngNextRow = -1
    
    If vfgTemplate.Row < vfgTemplate.Rows - 1 Then
        For i = vfgTemplate.Row + 1 To vfgTemplate.Rows - 1
            If Not vfgTemplate.RowHidden(i) Then
                lngNextRow = i
                Exit For
            End If
        Next i
    End If
    
    If lngNextRow = -1 Or vfgTemplate.Row = vfgTemplate.Rows - 1 Then
        For i = vfgTemplate.Rows - 1 To 1 Step -1
            If Not vfgTemplate.RowHidden(i) And i <> vfgTemplate.Row Then
                lngNextRow = i
                Exit For
            End If
        Next i
    End If
    


     vfgTemplate.RowHidden(vfgTemplate.Row) = True
     
     '������״̬
     If vfgTemplate.Cell(flexcpData, vfgTemplate.Row, 1) <> "" Then
        Call SetRowDataState(vfgTemplate.Row, rsDel)
     Else
        Call SetRowDataState(vfgTemplate.Row, rsNormal)
     End If
     
     If lngNextRow > -1 Then vfgTemplate.Row = lngNextRow

    mblnModifyState = True
End Sub

Private Sub SetRowDataState(ByVal lngRow As Long, ByVal rsState As TRowState)
'����������״̬
    vfgTemplate.Cell(flexcpData, lngRow, atcState) = rsState
End Sub

Private Function GetRowDataState(ByVal lngRow As Long) As TRowState
'��ȡ������״̬
    GetRowDataState = Val(vfgTemplate.Cell(flexcpData, lngRow, atcState))
End Function


Private Function VerifyDataInputIsOk() As Boolean
'��֤���������Ƿ���ȷ
    Dim i As Long
    
    VerifyDataInputIsOk = False
    
    For i = 1 To vfgTemplate.Rows - 1
        If Not vfgTemplate.RowHidden(i) Then
            If vfgTemplate.TextMatrix(i, atcTitle) = "" Then
                MsgBox "���ⲻ��Ϊ�ա�", vbOKOnly, Me.Caption
                
                Call vfgTemplate.ShowCell(i, atcTitle)
                Call vfgTemplate.Select(i, atcTitle)
                Call vfgTemplate.EditCell
                
                Exit Function
            End If
            
            If Len(vfgTemplate.TextMatrix(i, atcCount)) > 8 Then
                MsgBox "��ֵλ�����ܳ���8λ��", vbOKOnly, Me.Caption
                
                Call vfgTemplate.ShowCell(i, atcCount)
                Call vfgTemplate.Select(i, atcCount)
                Call vfgTemplate.EditCell
                
                Exit Function
            End If
        End If
    Next i
    
    VerifyDataInputIsOk = True
End Function

Private Function SaveTemplate() As Boolean
'����ģ��
    Dim i As Long
    Dim arySql() As String
    Dim rsRowState As TRowState
    
    vfgTemplate.Row = vfgTemplate.Row
    
    SaveTemplate = False
    
    If Not VerifyDataInputIsOk Then Exit Function
    
    ReDim Preserve arySql(1)
    
    arySql(0) = ""
    
    For i = 1 To vfgTemplate.Rows - 1
        rsRowState = GetRowDataState(i)
        
        Select Case rsRowState
            Case TRowState.rsNew
                If vfgTemplate.TextMatrix(i, atcTitle) <> "" Then
                    ReDim Preserve arySql(UBound(arySql) + 1)
                    arySql(UBound(arySql)) = "zl_��������ģ��_Insert(" & mlngRequestPageId & ",'" & _
                                                                        vfgAffix.Text & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcTitle) & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcContext) & "'," & _
                                                                        Val(vfgTemplate.TextMatrix(i, atcCount)) & ")"
                End If

            Case TRowState.rsDel
                ReDim Preserve arySql(UBound(arySql) + 1)
                arySql(UBound(arySql)) = "zl_��������ģ��_Del(" & vfgTemplate.Cell(flexcpData, i, atcTitle) & ")"
                
            Case TRowState.rsModify
                If vfgTemplate.TextMatrix(i, atcTitle) <> "" Then
                    ReDim Preserve arySql(UBound(arySql) + 1)
                    arySql(UBound(arySql)) = "zl_��������ģ��_Update(" & vfgTemplate.Cell(flexcpData, i, atcTitle) & ",'" & _
                                                                        vfgTemplate.TextMatrix(i, atcTitle) & "','" & _
                                                                        vfgTemplate.TextMatrix(i, atcContext) & "'," & _
                                                                        Val(vfgTemplate.TextMatrix(i, atcCount)) & ")"
                End If
                
        End Select
    Next i
    
    
On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    '�������ݱ���
    For i = LBound(arySql) To UBound(arySql)
        If arySql(i) <> "" Then
            Call zlDatabase.ExecuteProcedure(arySql(i), "���渽��ģ��")
        End If
    Next i
    
    gcnOracle.CommitTrans
    
    '�������븽��ģ������
    Call LoadTemplateToDataSet(mlngRequestPageId)
    
    mblnModifyState = False
    SaveTemplate = True
Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Sub ClearUseCount()
'���ʹ�ô���
    Dim i As Long
    
    For i = 1 To vfgTemplate.Rows - 1
        If Val(vfgTemplate.TextMatrix(i, atcCount)) <> 0 Then
            vfgTemplate.TextMatrix(i, atcCount) = 0
            Call SetRowDataState(i, rsModify)
            mblnModifyState = True
        End If
    Next i
End Sub


Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Select Case Control.ID
    
        Case TMenuType.mtCancel
            Call CancelModify       '�����޸�
            
        Case TMenuType.mtNew
            Call NewTemplate        '����ģ��
            
        Case TMenuType.mtDel
            Call DelTemplate        'ɾ��ģ��
                        
        Case TMenuType.mtSave
            Call SaveTemplate       '����ģ��
            
        Case TMenuType.mtClearCount
            Call ClearUseCount      '���ʹ�ô���
                            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
'        Case TMenuType.mtImport
'            '���뷽��......
'
'        Case TMenuType.mtExport
'            '����ģ��......
            
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(Control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(Control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(Control)
            
'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
'����ģ�����
    vfgAffix.Left = Left
    vfgAffix.Top = Top
    vfgAffix.Height = Bottom - IIf(stbThis.Visible, stbThis.Height, 0) - Top
    
    
    vfgTemplate.Left = vfgAffix.Width + 80
    vfgTemplate.Top = Top
    vfgTemplate.Width = ScaleWidth - vfgAffix.Width - 80
    vfgTemplate.Height = vfgAffix.Height
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case TMenuType.mtCancel, TMenuType.mtSave
            Control.Enabled = mblnModifyState
        Case TMenuType.mtNew, TMenuType.mtDel, TMenuType.mtClearCount
            Control.Enabled = vfgAffix.Rows > 1
    End Select
End Sub

Private Sub LoadRequestAffix(ByVal lngRequestPageId As Long)
'�������븽��
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select ��Ŀ from �������ݸ��� where ֻ��=0 and �ļ�Id=[1] order by ����"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRequestPageId)
    
    vfgAffix.Rows = 1
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    While Not rsData.EOF
        vfgAffix.Rows = vfgAffix.Rows + 1
        vfgAffix.TextMatrix(vfgAffix.Rows - 1, 0) = NVL(rsData!��Ŀ)
        vfgAffix.Cell(flexcpAlignment, vfgAffix.Rows - 1) = flexAlignLeftCenter
    
        Call rsData.MoveNext
    Wend
    
    vfgAffix.Row = 1
End Sub

Private Sub LoadAffixTemplate(ByVal lngRequestPageId As Long, ByVal strProjectName As String)
'���븽��ģ��

    vfgTemplate.Rows = 1
    
    If mrsAffixTemplate Is Nothing Then Exit Sub
    
    mrsAffixTemplate.Filter = "�����ļ�Id=" & lngRequestPageId & " and ���ݸ���='" & strProjectName & "'"
    
    If mrsAffixTemplate.RecordCount <= 0 Then Exit Sub
    
    While Not mrsAffixTemplate.EOF
        vfgTemplate.Rows = vfgTemplate.Rows + 1
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcTitle) = NVL(mrsAffixTemplate!ģ�����)
        vfgTemplate.Cell(flexcpData, vfgTemplate.Rows - 1, atcTitle) = NVL(mrsAffixTemplate!ID)
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcCount) = NVL(mrsAffixTemplate!ʹ�ô���)
        
        vfgTemplate.Cell(flexcpText, vfgTemplate.Rows - 1, atcContext) = NVL(mrsAffixTemplate!ģ������)
        
        Call SetRowDataState(vfgTemplate.Rows - 1, TRowState.rsNormal)
        
        Call mrsAffixTemplate.MoveNext
    Wend
    
    vfgTemplate.Row = 1
End Sub


Private Sub LoadTemplateToDataSet(ByVal lngRequestPageId As Long)
'���븽��ģ�嵽���ݼ�
    Dim strSQL As String
    
    strSQL = "select Id,�����ļ�Id,���ݸ���,ģ�����,ģ������,ʹ�ô��� from ��������ģ�� where �����ļ�Id=[1] order by ���ݸ���,ģ�����"
    
    Set mrsAffixTemplate = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ģ��", lngRequestPageId)
End Sub

Private Sub Form_Load()
'###############���õ�������################
'    InitDebugObject 1290, Me, "zlhis", "HIS"
'    mlngRequestPageId = 118
'###########################################

    Call RestoreWinState(Me, App.ProductName)

    mblnModifyState = False

    
    Call InitFaceList
    Call InitCommandBars
    
    Call LoadTemplateToDataSet(mlngRequestPageId)
    Call LoadRequestAffix(mlngRequestPageId)
End Sub


Public Sub InitDebugObject(ByVal lngModuleNum As Long, ByVal frmMain As Object, ByVal strUser As String, ByVal strPwd As String)
'��ʼ������״̬�µ��������
    Set gcnOracle = New ADODB.Connection
    
    Call OraDataOpen("", strUser, strPwd)
    
    glngSys = 100
    gstrPrivs = ";PACS�����ӡ;PACS����ɾ��;PACS������д;PACS�������Ʊ���;PACS�����޶�;PACS���˱���;�ɼ���������;��������;�洢����;��������;����;��鱨��;���Ǽ�;������;��ɫͨ��;�Ŷӽк�;���ͼ��;ȡ������;ȡ��������;ɾ����ʱӰ��;��Ƶ�ɼ�;���;���п���;ͼ�����;δ�ɷѱ���;�ļ�����;�ޱ������;Ӱ���ʿ�;������������;Excel���;"
    glngModul = lngModuleNum
    
    
    
    Call InitCommon(gcnOracle)
    
    Call RegCheck
End Sub

Private Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '���������Ϣ
            strError = Err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    gstrDBUser = UCase(strUserName)
    SetDbUser gstrDBUser
    
    OraDataOpen = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    OraDataOpen = False
    Err = 0
End Function


Private Sub InitFaceList()
'��ʼ�������б�����
    vfgAffix.TextMatrix(0, 0) = "���븽��"
    
    
    vfgTemplate.ColWidth(atcState) = 120
    
    vfgTemplate.TextMatrix(0, atcTitle) = "����"
    vfgTemplate.ColWidth(atcTitle) = 1600
    vfgTemplate.ColAlignment(atcTitle) = flexAlignLeftCenter
    
    vfgTemplate.TextMatrix(0, atcCount) = "����"
    vfgTemplate.ColWidth(2) = 520
    vfgTemplate.ColAlignment(atcCount) = flexAlignLeftCenter
    
    vfgTemplate.TextMatrix(0, atcContext) = "����"
    vfgTemplate.ColAlignment(atcContext) = flexAlignLeftCenter
End Sub




Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�
    End With
    
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                             '�Ƿ������Զ�������
        Set .Icons = zlCommFun.GetPubIcons                     '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "�ļ�(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��(&C)"): cbrControl.IconId = 3565
'    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtImport, "����(&I)"): cbrControl.IconId = 0: cbrControl.BeginGroup = True
'    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtExport, "����(&E)"): cbrControl.IconId = 0
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "�༭(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtNew, "����(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "ɾ��(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtClearCount, "���ʹ�ô���(&F)"): cbrControl.IconId = 3002: cbrControl.BeginGroup = True
        
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����", "���淽��"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��", "ȡ���޸�"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtNew, "����", "��������"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "ɾ��", "ɾ������"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�", "�˳�"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
End Sub




Public Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(T)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
            
                With cbrControl.CommandBar '�����˵�
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
        End With
    End If

    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    If Not (objHelpMenu Is Nothing) Then
        Set cbrMenuBar = objHelpMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "��������(M)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 901
                
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(W)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
                
                With cbrControl.CommandBar
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(0)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(1)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 1
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(2)")
                        cbrPopControl.Category = strMenuTag
                        cbrPopControl.IconId = 9022
                End With
                
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
                cbrControl.Category = strMenuTag
                cbrControl.IconId = 1
        End With
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    vfgTemplate.Row = vfgTemplate.Row
    
    If Not mblnModifyState Then Exit Sub
    
    lngResult = MsgBox("��ǰ���븽�" & vfgAffix.TextMatrix(vfgAffix.Row, vfgAffix.Col) & "����ģ�������ѱ��޸ģ��Ƿ񱣴棿", vbYesNoCancel, Me.Caption)
    
    Select Case lngResult
        Case vbNo
            mblnModifyState = False
            Exit Sub
        Case vbCancel
            Cancel = True
        Case vbYes
            If Not SaveTemplate Then Cancel = True
    End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgAffix_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If Not mblnModifyState Then Exit Sub
    
    lngResult = MsgBox("��ǰ���븽�" & vfgAffix.TextMatrix(OldRowSel, OldColSel) & "����ģ�������ѱ��޸ģ��Ƿ񱣴棿", vbYesNoCancel, Me.Caption)
    
    Select Case lngResult
        Case vbNo
            mblnModifyState = False
            Exit Sub
        Case vbCancel
            Cancel = True
        Case vbYes
            If Not SaveTemplate Then Cancel = True
    End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vfgAffix_SelChange()
On Error GoTo ErrHandle
    
    Call LoadAffixTemplate(mlngRequestPageId, vfgAffix.Text)
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vfgTemplate_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    Dim rsRowState As TRowState
    
    '���û�����κθı䣬�򲻽���༭״̬
    If vfgTemplate.TextMatrix(Row, Col) = mstrStartEditText Then Exit Sub
    
    rsRowState = GetRowDataState(Row)
    If rsRowState <> rsNew And rsRowState <> rsDel Then
        Call SetRowDataState(Row, rsModify)
    End If

    mblnModifyState = True
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub vfgTemplate_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = atcCount Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii = vbKeyEscape Then Exit Sub
        If KeyAscii = vbKeyDelete Then Exit Sub
        If KeyAscii = vbKeyBack Then Exit Sub
            
        KeyAscii = 0
        
    End If
End Sub

Private Sub vfgTemplate_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        Call EditNextCell(vfgTemplate.Row)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub EditNextCell(ByVal lngRow As Long, Optional ByVal blnAutoNextRow As Boolean = True)
'�༭��һ��
    Dim iRow As Long
    Dim iCol As Long
            
    If vfgTemplate.Editable = flexEDNone Then Exit Sub
    
    Do While vfgTemplate.Col + 1 < vfgTemplate.Cols
        If Not vfgTemplate.ColHidden(vfgTemplate.Col + 1) Then
            Exit Do
        Else
            Call vfgTemplate.Select(lngRow, vfgTemplate.Col + 1)
        End If
    Loop
    
nextCell:
    
    If vfgTemplate.Col + 1 >= vfgTemplate.Cols Then
 
        iRow = GetNextRowIndex(lngRow)
        
        If iRow > 0 Then
            For iCol = 1 To vfgTemplate.Cols - 1
                If Not vfgTemplate.ColHidden(iCol) Then Exit For
            Next iCol
            
            If iRow < vfgTemplate.Rows Then
                If iCol = vfgTemplate.Cols Then iCol = vfgTemplate.Cols - 1
                
                Call vfgTemplate.Select(iRow, iCol)
                Call vfgTemplate.ShowCell(iRow, iCol)
            End If
        End If
        
        Call vfgTemplate.EditCell
 
        Exit Sub
    End If
    
    
    Call vfgTemplate.Select(lngRow, vfgTemplate.Col + 1)
        
    Call vfgTemplate.EditCell
End Sub

Public Function GetNextRowIndex(ByVal lngRow As Long) As Long
'ȡ����һ�е�����
    Dim i As Long
    
    GetNextRowIndex = -1
    
    For i = lngRow + 1 To vfgTemplate.Rows - 1
        If Not vfgTemplate.RowHidden(i) Then
            GetNextRowIndex = i
            Exit Function
        End If
    Next i
    
    If GetNextRowIndex = -1 Then
        i = lngRow - 1
        Do While i > 0
            If Not vfgTemplate.RowHidden(i) Then
                GetNextRowIndex = i
                Exit Function
            End If
            
            i = i - 1
        Loop
    End If
End Function

Private Sub vfgTemplate_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error Resume Next
    mstrStartEditText = vfgTemplate.TextMatrix(Row, Col)
End Sub

