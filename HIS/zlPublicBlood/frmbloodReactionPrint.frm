VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmbloodReactionPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ����"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "frmbloodReactionPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VSFlex8Ctl.VSFlexGrid VSFPrint 
      Height          =   2295
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   4215
      _cx             =   7435
      _cy             =   4048
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
      AllowUserResizing=   1
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Bindings        =   "frmbloodReactionPrint.frx":000C
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmbloodReactionPrint.frx":0020
   End
End
Attribute VB_Name = "frmbloodReactionPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlng����ID As Long
Private mlng������Դ As Long          '1-����  2-סԺ
Private mlng��ҳid As String
Private mRsFY As ADODB.Recordset      '������Ϣ��¼��
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mstr�շ�id As String
Private mstr�Һŵ� As String
Private mlng�׶� As Long              '0-����  1-סԺ  2-��Ѫ��
Private marrFilter                    '�ֽ�mstrFilter��õ��Ĺ�����������
Private mstrFilter As String          '���������ַ���
Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ��Commandbar
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ������
    
    Call CommandBarInit(cbsMain)
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    '------------------------------------------------------------------------------------------------------------------
    Set cbsMain.Icons = ImageManager.Icons
    cbsMain.Options.LargeIcons = False
    
    Set objBar = cbsMain.Add("������", xtpBarTop)
        objBar.ContextMenuPresent = False
        objBar.ShowTextBelowIcons = False
        objBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap Or xtpFlagAlignBottom
        
        Set objControl = NewToolBar(objBar, xtpControlButton, 1, "ȫѡ", True, , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 2, "ȫ��", True, , xtpButtonIconAndCaption)
        Set objControl = NewToolBar(objBar, xtpControlButton, 4, "ȷ��", True, , xtpButtonIconAndCaption)
        mobjStateInfo.Flags = xtpFlagRightAlign

    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings

        .Add FCONTROL, vbKeyA, 1            'ȫѡ
        .Add FSHIFT, vbKeyDelete, 2         'ȫ��
        
    End With
    
    InitCommandBar = True
    Exit Function
ErrHand:
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long
    Dim lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim StrSqlSAD As String
    Dim strOPT As String
    Dim CanbeTransfer As Boolean
    Dim blnHD As Boolean
    blnHD = True
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        
            Case "��ʼ���"
                '��ʼvsf���
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSFPrint, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("�շ�id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("", 400, flexAlignLeftCenter, flexDTBoolean, "", , True)
                    Call .AppendColumn("Ѫ�����", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��Ѫ��Ŀ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("������", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("״̬", 1000, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("��Ѫʷ", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��Ѫ����", 0, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��Ӧʱ��", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("��¼��", 900, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("��¼ʱ��", 1200, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ȷ����", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("ȷ��ʱ��", 1200, flexAlignLeftCenter, flexDTString, "", , True)

                    .AppendRows = False
                    .SysHidden(.ColIndex("�շ�id")) = True
                    .SysHidden(.ColIndex("��Ѫʷ")) = True
                    .SysHidden(.ColIndex("��Ѫ����")) = True
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    
                End With
                
            Case "��ȡ��Ӧ��¼"
                Dim lngFilter
                Dim ArrTime
                Dim str�Ƿ��� As String
                Dim lng����id As Long
                Dim str��ʼʱ�� As String
                Dim str����ʱ�� As String
                
                If mstrFilter <> "" Then
                    lngFilter = marrFilter(3) '�ύ״̬
                    str�Ƿ��� = marrFilter(2) '��¼��
                    lng����id = marrFilter(0) '����id
                    str��ʼʱ�� = Split(marrFilter(1), "'")(0)
                    str����ʱ�� = Split(marrFilter(1), "'")(1)
                Else
                    lngFilter = 0
                    str�Ƿ��� = ""
                    lng����id = 0
                    str��ʼʱ�� = Now
                    str����ʱ�� = Now
                End If
                '��ȡ���˵���Ѫ��Ӧ���ݣ���Ҫ�Ǵ���Ѫ��Ӧ��¼�л�ȡ
                strSqlFY = " select distinct d.�շ�id, c.Ѫ�����, Decode(d.��Ѫʷ, 0, '��', '��') As ��Ѫʷ, d.��Ѫ����, d.��Ѫ��Ŀ, d.������, " & _
                           " Decode(d.״̬, 0, 'δ�ύ', 1, 'ҽ�����ύ', '��Ѫ�����ύ') As ״̬, to_char(d.��Ӧʱ��,'yyyy-mm-dd HH24:mi:ss') as ��Ӧʱ��, d.��¼��, d.��¼ʱ��, d.ȷ����,to_char(d.ȷ��ʱ��,'yyyy-mm-dd HH24:mi:ss') as ȷ��ʱ�� " & _
                           " from ����ҽ����¼ a,ѪҺ��Ѫ��¼ b,ѪҺ�շ���¼ c,��Ѫ��Ӧ��¼ d " & _
                           " where d.�շ�id=c.id and c.�䷢id=b.id and mod(c.��¼״̬,3)=1 and c.�˶��� is not null and b.����id=a.id " & _
                           " and a.����id=[1] "
                
                If mlng������Դ = 2 Then 'סԺ����
                    If lngFilter = 0 And mlng�׶� = 1 Then 'ȫ������,ҽ���׶�
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] "
                    ElseIf lngFilter = 0 And mlng�׶� = 2 Then 'ȫ������,��Ѫ�ƽ׶�
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬ <>0 "
                    ElseIf lngFilter = 1 And mlng�׶� = 1 Then 'δ�ύ����,ҽ��
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬=0 "
                    ElseIf lngFilter = 1 And mlng�׶� = 2 Then 'δ�ύ����,��Ѫ��
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬=1 "
                    ElseIf lngFilter = 2 And mlng�׶� = 1 Then '���ύ���ݣ�ҽ��
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬ <>0 "
                    ElseIf lngFilter = 2 And mlng�׶� = 2 Then '���ύ���ݣ���Ѫ��
                        strSqlFY = strSqlFY & "and a.��ҳid=[2] and d.״̬=2 "
                    End If
                Else
                    If lngFilter = 0 And mlng�׶� = 1 Then 'ȫ������,ҽ���׶�
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] "
                    ElseIf lngFilter = 0 And mlng�׶� = 2 Then 'ȫ������,��Ѫ�ƽ׶�
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬ <>0 "
                    ElseIf lngFilter = 1 And mlng�׶� = 1 Then 'δ�ύ����,ҽ��
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬=0 "
                    ElseIf lngFilter = 1 And mlng�׶� = 2 Then 'δ�ύ����,��Ѫ��
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬=1 "
                    ElseIf lngFilter = 2 And mlng�׶� = 1 Then '���ύ���ݣ�ҽ��
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬ <>0 "
                    ElseIf lngFilter = 2 And mlng�׶� = 2 Then '���ύ���ݣ���Ѫ��
                        strSqlFY = strSqlFY & "and a.�Һŵ�=[7] and d.״̬=2 "
                    End If
                End If
                
                If marrFilter(2) <> "" Then
                    strSqlFY = strSqlFY & " and d.��¼��=[3] "
                End If
    
                If Val(marrFilter(0)) <> -1 And mstrFilter <> "" Then '���û�й�������ʱVal(mArrFilter(0))=0
                    strSqlFY = strSqlFY & " and c.�Է�����id =[4] "
                End If
                
                If mlng�׶� = 2 And mstrFilter <> "" Then  '��Ѫ��Ҫ��ʱ����˷�Ӧ��¼
                    strSqlFY = strSqlFY & " and d.��Ӧʱ�� Between [5] and [6] order by ��Ӧʱ��"
                Else
                    strSqlFY = strSqlFY & " order by ��Ӧʱ�� "
                End If
                
                Set mRsFY = gobjDatabase.OpenSQLRecord(strSqlFY, "������Ѫ��Ӧ��¼", mlng����ID, mlng��ҳid, str�Ƿ���, lng����id, CDate(str��ʼʱ��), CDate(str����ʱ��), mstr�Һŵ�)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    ExecuteCommand = False
End Function

Public Function BloodPrintList(lng����ID As Long, lng������Դ As Long, lng��ҳid As Long, strFilter As String, lng�׶� As Long, Optional strSelBloodid As String = "") As String
    '���ܣ���ʼ����ӡ����
    '������lng����ID-����id
    '      lng������Դ-1-���� 2-סԺ
    '      lng��ҳid-��ҳid
    '      strFilter-���������ַ���
    '      lng�׶�-0-���� 1-סԺ 2-��Ѫ��
    Dim strSQL As String
    Dim rsSql As ADODB.Recordset
    Dim lngi As Long
    mstrFilter = strFilter
    If strFilter <> "" Then
        marrFilter = Split(strFilter, "|")
    Else
        ReDim marrFilter(0 To 3)
    End If
    
    mlng����ID = lng����ID
    mlng������Դ = lng������Դ
    mlng��ҳid = lng��ҳid
    mstr�շ�id = ""
'    mlngFilter = lngFilter
    mlng�׶� = lng�׶�
    strSQL = " select no from ���˹Һż�¼ where id=[1]"
    Set rsSql = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", mlng��ҳid)
    If rsSql.RecordCount > 0 Then
        mstr�Һŵ� = rsSql.Fields("no")
    End If
    InitCommandBar
    Call ExecuteCommand("��ʼ���")
    Call ExecuteCommand("��ȡ��Ӧ��¼")
    Call mclsVsf.LoadGrid(mRsFY)
    
    If strSelBloodid <> "" Then '��λ��ѡ�е���Ѫ��Ӧ��¼
        For lngi = 1 To VSFPrint.Rows - 1
            If mclsVsf.TextMatrix(lngi, mclsVsf.ColIndex("�շ�id")) = Val(strSelBloodid) Then
                VSFPrint.TextMatrix(lngi, 1) = -1
                Exit For
            End If
        Next
    End If
    Me.Show (1)
    BloodPrintList = mstr�շ�id
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngi As Long
    Dim blnOk As Boolean
    blnOk = False
    Select Case Control.id
        Case 1
            For lngi = 1 To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(lngi, 1) = -1
            Next
        Case 2
            For lngi = 1 To VSFPrint.Rows - 1
                VSFPrint.TextMatrix(lngi, 1) = 0
            Next
        Case 4
            mstr�շ�id = ""
            For lngi = 1 To VSFPrint.Rows - 1
                If Val(VSFPrint.TextMatrix(lngi, 1)) = -1 Then
                    mstr�շ�id = mstr�շ�id & VSFPrint.TextMatrix(lngi, VSFPrint.ColIndex("�շ�id")) & ";"
                    blnOk = True
                End If
            Next
            If mstr�շ�id <> "" Then mstr�շ�id = Left(mstr�շ�id, Len(mstr�շ�id) - 1)
            If blnOk = True Then
                Unload Me
            Else
                MsgBox "δѡ��Ӧ��¼��", vbInformation, gstrSysName
            End If
            
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long
    On Error GoTo Errorhand
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    '���������ؼ�Resize����
    VSFPrint.Move lngLeft, lngTop + 50, lngRight - lngLeft, lngBottom - lngTop
Errorhand:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set mRsFY = Nothing
End Sub

Private Sub VSFPrint_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub

