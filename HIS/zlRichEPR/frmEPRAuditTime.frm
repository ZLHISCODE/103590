VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEPRAuditTime 
   BorderStyle     =   0  'None
   Caption         =   "�������ݼ��"
   ClientHeight    =   3840
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   6975
   Icon            =   "frmEPRAuditTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   0
      Left            =   465
      ScaleHeight     =   2535
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   270
      Width           =   5145
      Begin VSFlex8Ctl.VSFlexGrid vfgAudit 
         Height          =   1560
         Left            =   375
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3960
         _cx             =   6985
         _cy             =   2752
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
         BackColorSel    =   16772055
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
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
End
Attribute VB_Name = "frmEPRAuditTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ��־ = 0: ����: ����ID: ��ҳID: �¼�Ե��: Ӧд����: Ӧд����: ����: ����ʱ��: Ҫ��ʱ��: ���ʱ��: ��ɼ�¼id: ��ǰʱ��: ��ע˵��
End Enum

Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mlngMoual As Long
Private mclsAudit As clsVsf
Private mintType As Integer
Public Event AfterDocumentChanged(ByVal lngEPRKey As Long)
Public Event SelectVfgRow(ByVal strPatiInfo As String)
Public Event GotFocus()

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
'    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    
End Function

Public Sub zlClearData()
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mclsAudit.ClearGrid
End Sub

Public Sub zlPrintData(ByVal bytMode As Byte, Optional ByVal strPatiInfo As String)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow


    Set objPrint.Body = vfgAudit
    objPrint.Title.Text = "����ʱ�޼���¼"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add(strPatiInfo)
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""

End Sub

Public Function GetCurrentEPRKey() As Long
    With vfgAudit
        GetCurrentEPRKey = Val(.TextMatrix(.Row, mCol.��ɼ�¼id))
    End With
End Function

Public Function zlRefreshData(ByVal lngPatientKey As Long, ByVal lngPatientPageKey As Long, ByVal intKind As Integer, _
    Optional ByVal lngDeptId As Long, Optional ByVal intType As Integer = 1, Optional ByVal intState As Integer = 0, _
    Optional ByVal dtEndTime As Date) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������lngDeptId�ⲿ����ID��intType��1-��ǰ���ˣ�2-�ҵĲ��ˣ�3-���Ʋ���
    '���أ�
    '******************************************************************************************************************
    
    Dim lngPatiID As Long, lngPageId As Long
    Dim lngBalance As Long
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    
    If lngPatientKey = 0 Then Exit Function
    
    mintKind = intKind
    mintType = intType
    '��ȡʱ�޼������
    Call ExecuteCommand("ʱ�޼��", lngPatientKey, lngPatientPageKey)
     Select Case intType
            Case 1 '��ǰ����
            gstrSQL = "Select 0 As ���,'' as ����,a.����id ,a.��ҳid ,To_Char(�¼�ʱ��, 'yyyy-mm-dd hh24:mi ') || �¼� As �¼�Ե��, ������� || '-' || �������� As Ӧд����,b.���� As Ӧд����," & _
            "        Decode(Ψһ, 1, '��д', '��' || ���ں� || '����д') As ����, ��ʼʱ�� ����ʱ��, ����ʱ�� Ҫ��ʱ��, ���ʱ��, ��ɼ�¼id, Sysdate As ��ǰʱ��, Null As ��ע˵��" & _
            " From ���Ӳ���ʱ�� a,���ű� b" & _
            " Where a.����id = [1] And a.��ҳid = [2] And (a.�������� = [3] Or a.�������� in (5,6) And [3]<>4) And a.����ʱ�� - Sysdate < 2 And a.����id=b.ID" & _
            " Order By a.�¼�ʱ��,a.�������,A.��ʼʱ��"
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientKey, lngPatientPageKey, intKind)
            Case 2 '�ҵĲ���
            gstrSQL = "Select 0 As ���,d.����,a.����id ,a.��ҳid ,To_Char(�¼�ʱ��, 'yyyy-mm-dd hh24:mi ') || �¼� As �¼�Ե��, ������� || '-' || �������� As Ӧд����,b.���� As Ӧд����," & _
            "        Decode(Ψһ, 1, '��д', '��' || ���ں� || '����д') As ����, ��ʼʱ�� ����ʱ��, ����ʱ�� Ҫ��ʱ��, ���ʱ��, ��ɼ�¼id, Sysdate As ��ǰʱ��, Null As ��ע˵��" & _
            " From ���Ӳ���ʱ�� a,���ű� b,������ҳ c,������Ϣ D" & _
            " Where  a.��ҳid = [1] And (a.�������� = [2] Or a.�������� in (5,6) And [2]<>4) And a.����ʱ�� - Sysdate < 2 And a.����id=b.ID and C.����ID=a.����id  " & _
            " and c.��ҳid=a.��ҳid And d.����id=c.����id and"
             If intState = 2 Then
                gstrSQL = gstrSQL & " c.��Ժ����>=[4]"
              Else
                gstrSQL = gstrSQL & " D.��Ժ=1"
             End If
             If intKind = 1 Then
             gstrSQL = gstrSQL & " and c.����ҽʦ=[3] " & " Order By a.�¼�ʱ��,a.�������,A.��ʼʱ��"
             ElseIf intKind = 2 Then
             gstrSQL = gstrSQL & " and c.סԺҽʦ=[3] " & " Order By a.�¼�ʱ��,a.�������,A.��ʼʱ��"
             Else
             gstrSQL = gstrSQL & " and c.���λ�ʿ=[3] " & " Order By a.�¼�ʱ��,a.�������,A.��ʼʱ��"
             End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientPageKey, intKind, gstrUserName, dtEndTime)
            Case 3 '���Ʋ���
             gstrSQL = "Select 0 As ���,d.����,a.����id ,a.��ҳid ,To_Char(�¼�ʱ��, 'yyyy-mm-dd hh24:mi ') || �¼� As �¼�Ե��, ������� || '-' || �������� As Ӧд����,b.���� As Ӧд����," & _
            "        Decode(Ψһ, 1, '��д', '��' || ���ں� || '����д') As ����, ��ʼʱ�� ����ʱ��, ����ʱ�� Ҫ��ʱ��, ���ʱ��, ��ɼ�¼id, Sysdate As ��ǰʱ��, Null As ��ע˵��" & _
            " From ���Ӳ���ʱ�� a,���ű� b,������ҳ c,������Ϣ D" & _
            " Where a.��ҳid = [1] And (a.�������� = [2] Or a.�������� in (5,6) And [2]<>4) And a.����ʱ�� - Sysdate < 2 And a.����id=b.ID and"
             If intState = 2 Then
                gstrSQL = gstrSQL & " c.��Ժ����>=[4]"
             Else
                gstrSQL = gstrSQL & " D.��Ժ=1"
             End If
             gstrSQL = gstrSQL & " and C.����ID=a.����id  and c.��ҳid=a.��ҳid And d.����id=c.����id and c.��Ժ����ID=[3] Order By a.�¼�ʱ��,a.�������,A.��ʼʱ��"
             Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatientPageKey, intKind, lngDeptId, dtEndTime)
     End Select
    
    With Me.vfgAudit
        .Clear
        .FixedCols = 0
        Set .DataSource = rs
       
        .MergeCells = flexMergeFree: .MergeCol(mCol.�¼�Ե��) = True: .MergeCol(mCol.Ӧд����) = True: .MergeCol(mCol.����) = True
        .ColWidth(mCol.��־) = 250: .ColWidth(mCol.����ʱ��) = 1800: .ColWidth(mCol.Ҫ��ʱ��) = 1800: .ColWidth(mCol.���ʱ��) = 1800
        .ColWidth(mCol.��ɼ�¼id) = 0: .ColWidth(mCol.��ǰʱ��) = 0: .ColWidth(mCol.��ע˵��) = 2200: .ColWidth(mCol.����ID) = 0
        .ColWidth(mCol.��ҳID) = 0
         If mintType = 1 Then
         .ColWidth(mCol.����) = 0
         Else
         .ColWidth(mCol.����) = 1000
         End If
        .FixedCols = 1
        .TextMatrix(0, mCol.��־) = ""
        .FixedAlignment(mCol.��־) = flexAlignCenterCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.���ʱ��) = "" Then
                If .TextMatrix(lngCount, mCol.��ɼ�¼id) = "" Then
                    .TextMatrix(lngCount, mCol.��ע˵��) = "δ��д"
                Else
                    .TextMatrix(lngCount, mCol.��ע˵��) = "������д"
                End If
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.��ǰʱ��)) - CDate(.TextMatrix(lngCount, mCol.Ҫ��ʱ��))) * 24)
                .TextMatrix(lngCount, mCol.��־) = "��"
                If lngBalance >= 0 Then
                    .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & IIf(lngBalance = 0, "", ",�ѳ���" & lngBalance & "Сʱ")
                    .Cell(flexcpForeColor, lngCount, mCol.��ע˵��, lngCount, mCol.��ע˵��) = RGB(255, 0, 0)
                Else
                    If Abs(lngBalance) < 4 Then
                        .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(128, 128, 0)
                        .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�뾡�����"
                    Else
                        .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(0, 0, 255)
                        .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�밴ʱ���"
                    End If
                End If
            Else
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.���ʱ��)) - CDate(.TextMatrix(lngCount, mCol.Ҫ��ʱ��))) * 24)
                If lngBalance > 0 Then
                    .TextMatrix(lngCount, mCol.��־) = "��"
                    .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.��ע˵��) = "���,������" & lngBalance & "Сʱ"
                    .Cell(flexcpForeColor, lngCount, mCol.��ע˵��, lngCount, mCol.��ע˵��) = RGB(255, 0, 0)
                Else
                    .TextMatrix(lngCount, mCol.��ע˵��) = "�������"
                End If
            End If
            .TextMatrix(lngCount, mCol.����ʱ��) = Format(.TextMatrix(lngCount, mCol.����ʱ��), "yyyy-MM-dd HH:mm")
            .TextMatrix(lngCount, mCol.Ҫ��ʱ��) = Format(.TextMatrix(lngCount, mCol.Ҫ��ʱ��), "yyyy-MM-dd HH:mm")
            .TextMatrix(lngCount, mCol.���ʱ��) = Format(.TextMatrix(lngCount, mCol.���ʱ��), "yyyy-MM-dd HH:mm")
        Next
        .Row = 0
    End With
    
    zlRefreshData = True
    
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
                
        '------------------------------------------------------------------------------------------------------------------
        Set mclsAudit = New clsVsf
        With mclsAudit

            Call .Initialize(Me.Controls, vfgAudit, True, False)
            Call .ClearColumn

            Call .AppendColumn("", 250, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("�¼�Ե��", 1800, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 1800, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ҳID", 0, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("Ӧд����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("Ӧд����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����ʱ��", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("Ҫ��ʱ��", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("���ʱ��", 1800, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("��ɼ�¼id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("��ǰʱ��", 1670, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", "", True)
            Call .AppendColumn("��ע˵��", 900, flexAlignLeftCenter, flexDTString, "", "", True)

        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"


    '------------------------------------------------------------------------------------------------------------------
    Case "ʱ�޼��"
        
        strSQL = "zl_���Ӳ���ʱ��_makeup(" & Val(varParam(0)) & "," & Val(varParam(1)) & "," & IIf(mintKind = 1, 1, 2) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "����ʱ�޼��")
        
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Sub Form_Resize()

    On Error Resume Next
    
    picPane(0).Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsAudit Is Nothing) Then Set mclsAudit = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    vfgAudit.Move 0, 0, picPane(Index).Width, picPane(Index).Height
End Sub

Private Sub vfgAudit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strPatiInfo As String, lngPatiID As Long, lngPageId As Long
    Dim rsTemp As New ADODB.Recordset
    With vfgAudit
        If OldRow <> NewRow And NewRow > 0 Then
            
            RaiseEvent AfterDocumentChanged(Val(.TextMatrix(NewRow, mCol.��ɼ�¼id)))
            lngPatiID = Val(.TextMatrix(NewRow, mCol.����ID))
            lngPageId = Val(.TextMatrix(NewRow, mCol.��ҳID))
            If mintType <> 1 Then
                If mintKind = 1 Then
                    gstrSQL = "Select r.�����, r.No, r.����, r.�Ա�, r.����, r.�Ǽ�ʱ�� From ���˹Һż�¼ r Where r.Id =[1] And r.��¼����=1  and r.��¼״̬=1"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPageId)
                    With rsTemp
                        If .RecordCount <= 0 Then strPatiInfo = "�ò��˲����ڣ����ܴ������ݴ���"
                        strPatiInfo = "�����:" & !����� & "(No:" & !NO & ")    ����:" & !���� & "(" & !�Ա� & ")" & _
                                    "  ����:" & Format(!�Ǽ�ʱ��, "yyyy-MM-dd hh:mm")
                    End With
                Else
                    gstrSQL = "Select b.סԺ��, a.����, a.�Ա�, a.����, b.��Ժ���� As ����, b.��Ժ����" & _
                            " From ������Ϣ a, ������ҳ b" & _
                            " Where a.����id = b.����id And b.����id = [1] And Nvl(b.��ҳid, 0) = [2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId)
                    With rsTemp
                        If .RecordCount <= 0 Then strPatiInfo = "�ò��˲����ڣ����ܸպñ���ݺϲ��ȣ�"
                        strPatiInfo = "סԺ��:" & !סԺ�� & "(��" & lngPageId & "��סԺ)    ����:" & !���� & "(" & !�Ա� & ")" & _
                                    "  ����:" & Format(!��Ժ����, "yyyy-MM-dd hh:mm")
                    End With
                End If
                RaiseEvent SelectVfgRow(strPatiInfo)
            End If
            
            
        End If
    End With
End Sub

Private Sub vfgAudit_GotFocus()
    RaiseEvent GotFocus
End Sub

