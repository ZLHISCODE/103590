VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBloodInstantRptPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ѫִ�е���ӡ"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7830
   Icon            =   "frmBloodInstantRptPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid VSFPrint 
      Height          =   1560
      Left            =   1755
      TabIndex        =   0
      Top             =   390
      Width           =   2790
      _cx             =   4921
      _cy             =   2752
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
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
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
      RowHeightMin    =   270
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBloodInstantRptPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean
Private mlngAdviceId As Long
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

Public Function ShowMe(ByVal objfrm As Object, ByVal lngAdviceid As Long) As Boolean
    
    mblnOk = False
    mlngAdviceId = lngAdviceid
    Call InitCommandBar
    If LoadVsfPrint = False Then Exit Function
    If Not objfrm Is Nothing Then
        Me.Show 1, objfrm
    Else
        Me.Show 1
    End If
    ShowMe = mblnOk
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Select Case Control.id
        Case conMenu_Edit_SelAll, conMenu_Edit_ClsAll
            For i = VSFPrint.FixedRows To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(i, VSFPrint.ColIndex("ѡ��")) = IIf(Control.id = conMenu_Edit_SelAll, 1, 0)
            Next
        Case conMenu_File_PrintSet
            Call Rptprint(0)
        Case conMenu_File_Preview
            Call Rptprint(1)
        Case conMenu_File_Print
            Call Rptprint(2)
        Case conMenu_File_Exit
            mblnOk = False
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim Rmain As RECT
    
    On Error Resume Next
    Call cbsMain.GetClientRect(Rmain.Left, Rmain.Top, Rmain.Right, Rmain.Bottom)
    With VSFPrint
        .Left = Rmain.Left
        .Top = Rmain.Top
        .Width = Rmain.Right - Rmain.Left
        .Height = Rmain.Bottom - Rmain.Top
    End With
End Sub

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��Commandbar
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ������
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    Set cbsMain.Icons = gobjCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap Or xtpFlagAlignBottom
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_SelAll, "ȫѡ", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_ClsAll, "ȫ��", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_PrintSet, "��ӡ����", True, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ", False, , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�", True, , xtpButtonIconAndCaption)

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_SelAll            'ȫѡ
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_ClsAll         'ȫ��
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadVsfPrint() As Boolean
'���ܣ���ʼ����񣬲����ر������
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, VSFPrint, True, True)
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTBoolean, "", "ѡ��", False)
        Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)  '�շ�ID
        Call .AppendColumn("״̬", 810, flexAlignLeftCenter, flexDTString, , "ѪҺ״̬") '����ִ��״̬
        Call .AppendColumn("Ѫ�����", 1200, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("ѪҺ����", 1800, flexAlignLeftCenter, flexDTString)
        Call .AppendColumn("���", 810, flexAlignLeftCenter, flexDTString, , "ѪҺ���")
        Call .AppendColumn("ABO", 810, flexAlignLeftCenter, flexDTString, , "ABO", True)
        Call .AppendColumn("Rh(D)", 600, flexAlignLeftCenter, flexDTString, , "RH", True)
        
        'ִ����Ϣ
        Call .AppendColumn("ִ����", 1200, flexAlignLeftCenter, flexDTString, , "��ʼִ����")
        Call .AppendColumn("��ʼʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
        Call .AppendColumn("������", 1200, flexAlignLeftCenter, flexDTString, , "����ִ����")
        Call .AppendColumn("����ʱ��", 1500, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm")
       
        .AppendRows = False
        Call .InitializeEdit(True, False, False)
        Call .InitializeEditColumn(.ColIndex("ѡ��"), True, vbVsfEditCheck)
    End With
    
    strSQL = _
        "Select Id, Max(Ѫ�����) Ѫ�����, Max(Abo) Abo, Max(Rh) Rh, Max(ѪҺ����) ѪҺ����, Max(ѪҺ���) ѪҺ���, Max(ѪҺ״̬) ѪҺ״̬, Max(��ʼʱ��) ��ʼʱ��," & vbNewLine & _
        "       Max(��ʼִ����) ��ʼִ����, Max(����ʱ��) ����ʱ��, Max(����ִ����) ����ִ����" & vbNewLine & _
        "From (Select a.Id, a.Ѫ�����, a.Abo, a.Rh, e.���� As ѪҺ����, e.��� ѪҺ���," & vbNewLine & _
        "              Decode(Nvl(f.ִ��״̬, 0), 1, '����ִ��', 2, '���ִ��', 3, 'ִֹͣ��') ѪҺ״̬, Decode(g.��¼����, 1, g.ִ��ʱ��) ��ʼʱ��," & vbNewLine & _
        "              Decode(g.��¼����, 1, g.ִ����) ��ʼִ����, Decode(g.��¼����, 3, g.ִ��ʱ��) ����ʱ��, Decode(g.��¼����, 3, g.ִ����) ����ִ����" & vbNewLine & _
        "       From �շ���ĿĿ¼ e, ѪҺ�շ���¼ a, ѪҺִ�м�¼ g, ѪҺ���ͼ�¼ f, ѪҺ��Ѫ��¼ b" & vbNewLine & _
        "       Where e.Id = a.ѪҺid And Nvl(a.��д����, 0) <> 0 And Mod(a.��¼״̬, 3) = 1 And a.Id = f.�շ�id And g.�շ�id = f.�շ�id And" & vbNewLine & _
        "             f.�䷢id = b.Id And b.����id = [1])" & vbNewLine & _
        "Group By Id" & vbNewLine & _
        "Order By ��ʼʱ��"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "�ѷ�ѪҺ��Ϣ��ȡ", mlngAdviceId)
    If rsTemp.EOF Then
        MsgBox "��ҽ����δ������Ѫִ������Ǽǣ���ǼǺ��ٽ��д˲�����", vbInformation, gstrSysName
        Exit Function
    End If
    Call mclsVsf.LoadGrid(rsTemp, "", True)
    LoadVsfPrint = True
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub Rptprint(ByVal bytMode As Byte)
    Dim i As Integer
    Dim strIDs As String
    Dim strRptName As String
    
    strRptName = "ZL22_BILL_9005_1" 'ZL22_BILL_1938
    Select Case bytMode
        Case 0  '��ӡ����
            Call ReportPrintSet(gcnOracle, 2200, strRptName, Me)
        Case 1, 2 'Ԥ��  ��ӡ
            With VSFPrint
                For i = .FixedRows To .Rows - 1
                    If Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    End If
                Next
            End With
            strIDs = Mid(strIDs, 2)
            ReportOpen gcnOracle, 2200, strRptName, Me, "ҽ��id=" & mlngAdviceId, "�շ�ID=" & strIDs, bytMode
            mblnOk = True
    End Select
End Sub

Private Sub VSFPrint_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSFPrint_DblClick()
    With VSFPrint
        If .Row >= .FixedRows And .Col >= .FixedCols Then
            If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 Then
                .TextMatrix(.Row, .ColIndex("ѡ��")) = IIf(Abs(Val(.TextMatrix(.Row, .ColIndex("ѡ��")))) = 1, 0, 1)
            End If
        End If
    End With
End Sub

Private Sub VSFPrint_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not Col = VSFPrint.ColIndex("ѡ��") Then
        Cancel = True
    End If
End Sub
