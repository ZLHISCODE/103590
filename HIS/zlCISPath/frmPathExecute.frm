VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPathExecute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ִ��·����Ŀ"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9165
   Icon            =   "frmPathExecute.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9165
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9165
      TabIndex        =   4
      Top             =   0
      Width           =   9165
      Begin VB.Label lblTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   90
      End
      Begin VB.Image imgInfo 
         Height          =   720
         Left            =   195
         Picture         =   "frmPathExecute.frx":6852
         Top             =   45
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   10000
         Y1              =   800
         Y2              =   800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   10000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "����д·����Ŀ��ִ�н����ִ��˵����ִ�н��������Ϊ·�����������ݡ�"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6240
      Width           =   9165
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7800
         TabIndex        =   3
         Top             =   240
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5385
      Left            =   50
      TabIndex        =   1
      Top             =   860
      Width           =   9020
      _cx             =   15910
      _cy             =   9499
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
      BackColorFixed  =   15597549
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   32768
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPathExecute.frx":6FA7
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      BackColorFrozen =   14811105
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmPathExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngFun         As Long             '0-����ִ��,1-����ִ��,2-����ȡ��ִ��
Private mblnOK          As Boolean

Private mPP             As TYPE_PATH_Pati
Private mPati           As TYPE_Pati
Private mint����        As Integer          'int����=0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Private mlng·��ִ��ID  As Long             '����ִ��ʱ�Ŵ���
Private mblnNurse       As Boolean          'mlngFun=2ʱ,=False ��������ȡ��ִ����Ϊ��ʿ����Ŀ,=True ֻ��������ȡ���������ǻ�ʿ��ִ�����ǻ�ʿ����Ŀ
Private mcol            As Collection
Private mrsItem         As ADODB.Recordset
Private mfrmParent      As Object
Private mintMode        As Integer

Private Enum Eִ�н��
    E�Ѿ�ִ�� = 1
    E��δִ�� = 2
    Eȡ��ִ�� = 3
    E����ִ�� = 4
    E��ǰִ�� = 5
    E�Ӻ�ִ�� = 6
End Enum

Public Function ShowMe(frmParent As Object, ByVal lngFun As Long, t_pati As TYPE_Pati, t_pp As TYPE_PATH_Pati, _
    ByVal lng·��ִ��ID As Long, ByVal int���� As Long, Optional ByVal blnNurse As Boolean = False, Optional ByVal intMode As Integer) As Boolean
    
    Set mfrmParent = frmParent
    mlngFun = lngFun
    mPati = t_pati
    mPP = t_pp
    mlng·��ִ��ID = lng·��ִ��ID
    mint���� = int����
    mblnNurse = blnNurse
    mintMode = intMode
    
    If intMode = 1 Then
        Set mrsItem = GetItemOut
    Else
        Set mrsItem = GetItem
    End If
    If mrsItem.RecordCount = 0 Then
        MsgBox "�ò���û����Ҫ����ִ�е���Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Function GetItem() As ADODB.Recordset
'���ܣ���ȡ��ִ�еļ�¼
    Dim strSql As String
    Dim strUType As String, strIF As String
    Dim strcol As String
    
    If mint���� = 0 Then
        strUType = " And nvl(Nvl(b.ִ����, a.ִ����),1) = 1"
    ElseIf mint���� = 1 Then
        If mlngFun = 2 And mblnNurse Then
            strUType = " And nvl(Nvl(b.������, a.������),1) = 2 And nvl(Nvl(b.ִ����, a.ִ����),2) = 2"
        Else
            strUType = " And nvl(Nvl(b.ִ����, a.ִ����),2) = 2"
        End If
    End If
    
    If mlngFun = 0 Then
        strIF = "a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And Nvl(a.����ʱ������,0)<>2 And a.ִ��ʱ�� Is Null"
    ElseIf mlngFun = 2 Then
        strIF = "a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3]  And a.ִ��ʱ�� Is Not Null"
    Else
        strIF = "a.id = [4] And Nvl(a.����ʱ������,0)<>2 And a.ִ��ʱ�� Is Null"
    End If
    If mlngFun = 2 Then
        strcol = "a.ִ�н��,a.ִ��˵��,"
    Else
        strcol = "Nvl(b.��Ŀ���, a.��Ŀ���) As ִ�н��,"
    End If
    'Distinct����Ϊһ����Ŀ���ܶ����˶��ҽ������
    strSql = "Select Distinct a.����,a.ID, c.��� ���˳��, Nvl(b.��Ŀ���, a.��Ŀ���) ��Ŀ˳��, Nvl(b.��Ŀ����, a.��Ŀ����) As ��Ŀ����, " & strcol & " Nvl(b.ͼ��id, a.ͼ��id) ͼ��id," & _
            "Decode(d.ҽ������ID,Null,0,1) as ҽ����Ŀ,Decode(e.����ҽ��ID,Null,0,1) as ��ҽ��" & vbNewLine & _
            "From ����·��ִ�� A,�ٴ�·����Ŀ B,�ٴ�·������ C,�ٴ�·��ҽ�� D,����·��ҽ�� E,���˺ϲ�·�� F" & vbNewLine & _
            "Where " & strIF & " And f.��Ҫ·����¼id(+) = a.·����¼id  And (f.·��id = c.·��id And f.�汾�� = c.�汾��  or c.·��id = [5] And c.�汾�� = [6])" & _
            " And  a.���� = c.���� And NVL(c.��֧id,0)=NVL(b.��֧ID,0) And b.ID = d.·����ĿID(+) And a.ID = E.·��ִ��ID(+)" & vbNewLine & _
            "And a.��Ŀid = b.Id(+)" & strUType & vbNewLine & _
            "Order by c.���,Nvl(b.��Ŀ���,a.��Ŀ���)"
    On Error GoTo errH
    Set GetItem = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ִ�е���Ŀ", mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����, mlng·��ִ��ID, mPP.·��ID, mPP.�汾��)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetItemOut() As ADODB.Recordset
'���ܣ���ȡ��ִ�еļ�¼
    Dim strSql As String
    Dim strIF As String
    Dim strcol As String
    
    If mlngFun = 0 Then
        strIF = "a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And a.ִ��ʱ�� Is Null"
    ElseIf mlngFun = 2 Then
        strIF = "a.·����¼id = [1] And a.�׶�id = [2] And a.���� = [3] And a.ִ��ʱ�� Is Not Null"
    Else
        strIF = "a.id = [4] And a.ִ��ʱ�� Is Null"
    End If
    If mlngFun = 2 Then
        strcol = "a.ִ�н��,a.ִ��˵��,"
    Else
        strcol = "Nvl(b.��Ŀ���, a.��Ŀ���) As ִ�н��,"
    End If
    
    strSql = " Select Distinct a.����, a.Id, c.��� ���˳��, Nvl(b.��Ŀ���, a.��Ŀ���) ��Ŀ˳��, Nvl(b.��Ŀ����, a.��Ŀ����) As ��Ŀ����," & strcol & "Nvl(b.ͼ��id, a.ͼ��id) ͼ��id," & vbNewLine & _
             "                Decode(d.ҽ������id, Null, 0, 1) As ҽ����Ŀ, Decode(e.����ҽ��id, Null, 0, 1) As ��ҽ��" & vbNewLine & _
             " From ��������·��ִ�� A, ����·����Ŀ B, ����·������ C, ����·��ҽ�� D, ��������·��ҽ�� E" & vbNewLine & _
             " Where  " & strIF & " And (c.·��id = [5] And c.�汾�� = [6]) And a.���� = c.����  And b.Id = d.·����Ŀid(+) And" & vbNewLine & _
             "      a.Id = e.·��ִ��id(+) And a.��Ŀid = b.Id(+)" & vbNewLine & _
             " Order By c.���, Nvl(b.��Ŀ���, a.��Ŀ���)"

    On Error GoTo errH
    Set GetItemOut = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ִ�е���Ŀ", mPP.����·��ID, mPP.��ǰ�׶�ID, mPP.��ǰ����, mlng·��ִ��ID, mPP.·��ID, mPP.�汾��)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strIDs As String '·��ִ��ID
        
    If mlngFun = 0 Or mlngFun = 1 Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, mcol("ִ�н��"))) = "" And CStr(.Cell(flexcpData, i, mcol("ִ�н��"))) <> "" Then
                    MsgBox "��ѡ��һ��ִ�н����", vbInformation, gstrSysName
                    .SetFocus
                    .Select i, mcol("ִ�н��")
                    .TopRow = i
                    Exit Sub
                End If
            Next
        End With
    End If
    
    If mlngFun = 0 Or mlngFun = 2 Then
        With vsItem
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    strIDs = strIDs & "," & .TextMatrix(i, mcol("ID"))
                End If
            Next
            If strIDs = "" Then
                MsgBox "������ѡ��һ��·����Ŀ���򹴣���", vbInformation, gstrSysName
                Exit Sub
            End If
        End With
    End If
    
    If mintMode = 1 Then
        If SaveItemOut = False Then Exit Sub
    Else
        If SaveItem = False Then Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Function SaveItem() As Boolean
'����:����·����Ŀִ�е�����
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSQL As String, strTotal As String, strThis As String, i As Long
    Dim strDate As String
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsItem
        If mlngFun = 2 Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    strSQL = "Zl_����·��ִ��_Delete(" & .TextMatrix(i, mcol("ID")) & ")"
                    colSQL.Add strSQL, "C" & colSQL.count + 1
                End If
            Next
        Else
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    strThis = .TextMatrix(i, mcol("ID")) & "|" & .TextMatrix(i, mcol("ִ�н��")) & "|" & _
                        IIf(Trim(.TextMatrix(i, mcol("ִ��˵��"))) = "", " ", Trim(.TextMatrix(i, mcol("ִ��˵��")))) & "||"
                        
                    If LenB(strTotal & strThis) > 4000 Then
                        strSQL = "Zl_����·��ִ��_Update('" & UserInfo.���� & "'," & strDate & ",'" & strTotal & "')"
                        colSQL.Add strSQL, "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                End If
            Next
            If strTotal <> "" Then
                strSQL = "Zl_����·��ִ��_Update('" & UserInfo.���� & "'," & strDate & ",'" & strTotal & "')"
                colSQL.Add strSQL, "C" & colSQL.count + 1
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False

    SaveItem = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveItemOut() As Boolean
'����:����·����Ŀִ�е�����
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSql As String, strTotal As String, strThis As String, i As Long
    Dim strDate As String
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsItem
        If mlngFun = 2 Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    strSql = "Zl_��������·��ִ��_Delete(" & .TextMatrix(i, mcol("ID")) & ")"
                    colSQL.Add strSql, "C" & colSQL.count + 1
                End If
            Next
        Else
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, mcol("ѡ��")) = 1 Then
                    strThis = .TextMatrix(i, mcol("ID")) & "|" & .TextMatrix(i, mcol("ִ�н��")) & "|" & _
                        IIf(Trim(.TextMatrix(i, mcol("ִ��˵��"))) = "", " ", Trim(.TextMatrix(i, mcol("ִ��˵��")))) & "||"
                        
                    If LenB(strTotal & strThis) > 4000 Then
                        strSql = "Zl_��������·��ִ��_Update('" & UserInfo.���� & "'," & strDate & ",'" & strTotal & "')"
                        colSQL.Add strSql, "C" & colSQL.count + 1
                        strTotal = strThis
                    Else
                        strTotal = strTotal & strThis
                    End If
                End If
            Next
            If strTotal <> "" Then
                strSql = "Zl_��������·��ִ��_Update('" & UserInfo.���� & "'," & strDate & ",'" & strTotal & "')"
                colSQL.Add strSql, "C" & colSQL.count + 1
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 1 To colSQL.count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False

    SaveItemOut = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mlngFun <> 2 Then
        If vsItem.Visible And vsItem.Rows > vsItem.FixedRows Then
            vsItem.SetFocus: vsItem.Col = mcol("ִ�н��")
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '����������ָ�����������
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long, lngW As Long
            
    Call InitItem
    Call LoadItem
    
    vsItem.Top = picInfo.Top + picInfo.Height
    If mlngFun = 1 Then
        'ֻ��һ��'û��"ѡ��"��
        For i = 0 To vsItem.Cols - 1
            lngW = lngW + vsItem.ColWidth(i)
        Next
        Me.Width = lngW + 500
        
        vsItem.Width = Me.ScaleWidth - 100
        vsItem.Height = 2000
        Me.Height = picInfo.Height + vsItem.Height + picBottom.Height
        
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    Else
        vsItem.Height = Me.ScaleHeight - picInfo.Height - picBottom.Height
        If mlngFun = 2 Then
            Me.Caption = "����ȡ��ִ��"
            lblNote.Caption = "��ѡ��Ҫȡ��ִ�е�·����Ŀ��"
        End If
    End If
    lblTip.Caption = "��ǰ����:" & mPP.��ǰ���� & "(��" & mPP.��ǰ���� & "��)"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
End Sub

Private Sub LoadItem()
'���ܣ����ش�ִ�е�·����Ŀ
    Dim i As Long, j As Long
    Dim strִ�н�� As String, strȱʡ��� As String, str��ִ�н�� As String
    Dim arrtmp As Variant
    
    With vsItem
        .Redraw = flexRDNone
        .Rows = .FixedRows + mrsItem.RecordCount
        .MergeCol(0) = True
        For i = 1 To mrsItem.RecordCount
            .TextMatrix(i, mcol("����")) = mrsItem!����
            Call .Select(i, mcol("����"))
            .CellAlignment = flexAlignCenterCenter
            
            .TextMatrix(i, mcol("��Ŀ����")) = mrsItem!��Ŀ����
            
            .Cell(flexcpChecked, i, mcol("ѡ��")) = 1
                        
            If mlngFun = 2 Then
                .TextMatrix(i, mcol("ִ�н��")) = "" & mrsItem!ִ�н��
                .TextMatrix(i, mcol("ִ��˵��")) = "" & mrsItem!ִ��˵��
            Else
                If Not IsNull(mrsItem!ִ�н��) Then
                    strִ�н�� = CStr(Split("" & mrsItem!ִ�н��, vbTab)(0))
                    strȱʡ��� = Split("" & mrsItem!ִ�н��, vbTab)(1)
                End If
                
                If mrsItem!ҽ����Ŀ = 1 And mrsItem!��ҽ�� = 0 Then
                'ѡ������ʱδ���ɵ���Ŀ��ִ�н������Ϊ�Ѿ�ִ��
                'ȱʡ�����û��ִ������
                    If InStr(strִ�н��, "|") > 0 Then
                        j = InStr(strִ�н��, strȱʡ���)
                        If j > 0 Then
                            j = j + Len(strȱʡ���) + 1
                            If Val(Mid(strִ�н��, j, 1)) = Eִ�н��.E�Ѿ�ִ�� Then strȱʡ��� = ""
                        End If
                    End If
                    
                    '��ѡ�б��в���ʾ�Ѿ�ִ�еĽ����¼
                    If InStr(strִ�н��, "|") > 0 Then
                        str��ִ�н�� = ""
                        arrtmp = Split(strִ�н��, ",")
                        For j = 0 To UBound(arrtmp)
                            If Val(Split(arrtmp(j), "|")(1)) <> Eִ�н��.E�Ѿ�ִ�� Then
                                str��ִ�н�� = str��ִ�н�� & "," & arrtmp(j)
                            End If
                        Next
                        strִ�н�� = Mid(str��ִ�н��, 2)
                    End If
                End If
                .TextMatrix(i, mcol("ִ�н��")) = strȱʡ���
                .Cell(flexcpData, i, mcol("ִ�н��")) = strִ�н��
            End If
            .TextMatrix(i, mcol("ID")) = Val(mrsItem!ID)
            
            If Not IsNull(mrsItem!ͼ��ID) Then
                Call .Select(i, mcol("��Ŀ����"))
                .CellPictureAlignment = flexPicAlignRightCenter 'flexPicAlignLeftCenter
                .CellPicture = GetPathIcon(mrsItem!ͼ��ID)
            End If
            
            mrsItem.MoveNext
        Next
        
        .Redraw = True
        .AutoSize .FixedCols, .Cols - 1, , 45 '��ҪDraw֮�����Ч
    End With
End Sub

Private Sub InitItem()
'����: ��ʼ��·����Ŀ��ͷ
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    If mlngFun = 1 Then
        strcol = "����,1200,1;��Ŀ����,2600,1;ѡ��;ִ�н��,900,4;ִ��˵��,2600,1;ID"
    Else
        strcol = "����,1200,1;��Ŀ����,3100,1;ѡ��,500,4;ִ�н��,900,4;ִ��˵��,3000,1;ID"
    End If
    arrHead = Split(strcol, ";")
    Set mcol = New Collection
   
    With vsItem
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        
        For i = 0 To UBound(arrHead)
            mcol.Add i, Split(arrHead(i), ",")(0)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .ColDataType(mcol("ѡ��")) = flexDTBoolean
        .Redraw = True
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Me.Visible And mlngFun <> 2 Then
        If NewCol = mcol("ִ�н��") Then
            Dim strTmp As String, arrtmp As Variant, i As Long, lngP As Long, blnDo As Boolean
            With vsItem
                blnDo = False
                strTmp = .Cell(flexcpData, NewRow, mcol("ִ�н��"))    '������������д|1,����δ��д|2
                arrtmp = Split(strTmp, ",")
                For i = 0 To UBound(arrtmp)
                    lngP = InStr(arrtmp(i), "|")
                    If lngP > 0 Then
                        blnDo = True
                        arrtmp(i) = Mid(arrtmp(i), 1, lngP - 1)
                    End If
                Next
                If blnDo Then
                    strTmp = Join(arrtmp, "|")
                Else
                    strTmp = Replace(strTmp, ",", "|")
                End If
                .ColComboList(NewCol) = strTmp
            End With
        End If
    End If
End Sub

Private Sub vsItem_DblClick()
'���ܣ�����ѡ��
    With vsItem
        If .MouseRow = .FixedRows - 1 And .Col = mcol("ѡ��") Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpChecked, i, mcol("ѡ��")) = IIf(.Cell(flexcpChecked, i, mcol("ѡ��")) = 1, 2, 1)
            Next
        End If
    End With
End Sub

Private Sub vsItem_GotFocus()
    vsItem.ForeColorSel = vbWhite
    vsItem.BackColorSel = &H8000000D
End Sub

Private Sub vsItem_LostFocus()
    vsItem.ForeColorSel = vbBlack
    vsItem.BackColorSel = vbWhite
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call ResultEnterNextCell
    End If
End Sub

Private Sub vsItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsItem.MouseCol = mcol("ѡ��") Then
        vsItem.ToolTipText = "˫������ȫ��ѡ���ȡ��"
    Else
        vsItem.ToolTipText = ""
    End If
End Sub

Private Sub vsItem_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mcol("��Ŀ����") Then Cancel = True
    If mlngFun = 2 Then
        If Col <> mcol("ѡ��") Then Cancel = True
    End If
End Sub

Private Sub ResultEnterNextCell()
    With vsItem
        If .Col < mcol("ִ��˵��") Then
            .Col = .Col + 1
        ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1: .Col = IIf(.ColHidden(mcol("ѡ��")), mcol("ִ�н��"), mcol("ѡ��"))
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

