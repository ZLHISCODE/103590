VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmItemImport 
   Caption         =   "������Ŀ"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9615
   Icon            =   "frmItemImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9615
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOutput 
      Caption         =   "������&O��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   10
      Top             =   4560
      Width           =   1100
   End
   Begin VB.CheckBox chk��Ӧ�� 
      Caption         =   "�µĳ����Զ�����"
      Height          =   375
      Left            =   8400
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox chkStop 
      Caption         =   "����������ֹ����"
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ����&C��"
      Height          =   350
      Left            =   8400
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "���루&I��"
      Enabled         =   0   'False
      Height          =   350
      Left            =   8400
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "�Ϸ��Լ��"
      Height          =   350
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Width           =   1100
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "��"
      Height          =   300
      Left            =   8040
      TabIndex        =   3
      Top             =   120
      Width           =   280
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFList 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8175
      _cx             =   14420
      _cy             =   8705
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
      BackColorAlternate=   14737632
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemImport.frx":6852
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
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   225
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog dlgOutput 
      Left            =   720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin VB.Label lblFile 
      Caption         =   "�ļ�"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   173
      Width           =   420
   End
   Begin VB.Label lblFileName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmItemImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTRCHARGETYPE As String = "�ϼ�����,����,����"
Private Const mstrCharge As String = "���,����,����,����,�Ƿ���,������Ŀ,�ּ�,��ʶ����,��ʶ����,��ѡ��,���,���㵥λ,�������,��������"
Private Const MSTRSTUFFTYPE As String = "�ϼ�����,����,����"
Private Const MSTRSTUFF As String = "����,Ʒ�ֱ���,Ʒ������,������,���,������,ɢװ��λ,��װ��λ,ɢװ��װ����ϵ��,�Ƿ���,�ɱ���,�ۼ�,������Ŀ,��Դ,�������,��ʶ����,��ʶ����,���Ŀⷿ����,���ϲ��ŷ���,Ч��(��),��׼�ĺ�,��Ʒע���̱�,ע��֤��,���֤��,���֤Ч��,��Ӧ������,��Ӧ�����֤��,��Ӧ�����֤Ч��"
Private Const MSTRMEDICALTYPE As String = "���,�ϼ�����,����,����"
Private Const MSTRMEDICAL As String = "���,����,Ʒ�ֱ���,Ʒ������,������,ҩƷ���,����,����,������λ,�ۼ۵�λ,�ۼۼ�������ϵ��,���ﵥλ,���ﵥλ����ϵ��,סԺ��λ,סԺ��λ����ϵ��,ҩ�ⵥλ,ҩ���װ����ϵ��,�Ƿ���,�ɱ���,�ۼ�,������Ŀ,סԺ�ɷ����,����ɷ����,�������,ҩ�����,ҩ������,Ч��(��),��Ӧ������,��Ӧ�����֤��,��Ӧ�����֤Ч��"
Private Const MINTTITLE As Integer = 2 '������
Private mobjWB As Object
Private mobjWS As Object
Private mobjWSType As Object
Private mobjXLS As Object
Private mRsError As Recordset
Private mintType As Integer
Private mLngCount As Long
Private mLngType As Long
Private mLngSumType As Long
Private mLngSumCount As Long
Private mCollTypeCols As Collection
Private mCollItemCols As Collection
Private mstrIn As String
Private mblnExists As Boolean

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdCheck_Click()
    Dim i As Long
    
    '��ձ���ڵ�����
    Me.VSFList.Rows = 1
    Set mRsError = New Recordset
    
    '������ݼ����ֶ�
    With mRsError
        If .State = 1 Then .Close

        .Fields.Append "Type", adVarChar, 200
        .Fields.Append "Error", adVarChar, 1000
        .Fields.Append "Row", adBigInt
        .Fields.Append "Col", adSmallInt
        .Fields.Append "Page", adVarChar, 20


        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    '������ʽ
   If CheckExcel = False Then Exit Sub

    '��������Ϣ
    CheckType mintType

    '�����Ŀ��Ϣ
    '��鹫����Ŀ
    CheckPub
    '���������Ŀ
    If mintType = 1 Then
        CheckCharge
    ElseIf mintType = 2 Then
        CheckMedi
    Else
        CheckStuff
    End If
    
    '��������
    With mRsError
        If .RecordCount > 0 Then
            VSFList.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To .RecordCount
                Me.VSFList.RowHeight(i) = 300
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("��������")) = mRsError!Type
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("����ԭ��")) = mRsError!Error
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("������")) = mRsError!Row
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("������")) = mRsError!Col
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("���")) = mRsError!Page
                .MoveNext
            Next
            
            Me.cmdCheck.Enabled = True
            Me.cmdImport.Enabled = False
            Me.cmdOutput.Enabled = True
        ElseIf .RecordCount = 0 Then
            Me.cmdCheck.Enabled = False
            Me.cmdImport.Enabled = True
            Me.cmdOutput.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdChoose_Click()
    OpenFile

    Me.cmdCheck.Enabled = Not (Me.lblFileName.Caption = "")
    Me.cmdImport.Enabled = False
End Sub

Private Sub cmdImport_Click()
    Dim strText As String
    Dim i As Long
    
    If cmdCheck.Enabled = True Then
        MsgBox "���Ƚ��С��Ϸ��Լ�顿������", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    mLngType = 0
    mLngCount = 0
    
    Me.prg.Visible = True
    Me.VSFList.Rows = 1
    Set mRsError = New Recordset
    
    '������ݼ����ֶ�
    With mRsError
        If .State = 1 Then .Close

        .Fields.Append "Type", adVarChar, 200
        .Fields.Append "Error", adVarChar, 1000
        .Fields.Append "Row", adBigInt
        .Fields.Append "Col", adSmallInt
        .Fields.Append "Page", adVarChar, 20
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    
    '�������
    SaveType
    '������Ŀ
    If mintType = 1 Then
        SaveData
    ElseIf mintType = 2 Then
        SaveMedi
    Else
        SaveStuff
    End If
    
    '��������
    With mRsError
        If .RecordCount > 0 Then
            VSFList.Rows = .RecordCount + 1
            .MoveFirst
            For i = 1 To .RecordCount
                Me.VSFList.RowHeight(i) = 300
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("��������")) = mRsError!Type
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("����ԭ��")) = mRsError!Error
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("������")) = mRsError!Row
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("������")) = mRsError!Col
                Me.VSFList.TextMatrix(i, VSFList.ColIndex("���")) = mRsError!Page
                If mRsError!Page = "����" Then
                    mobjWSType.Rows(Val(mRsError!Row)).Font.Color = vbRed
                Else
                    mobjWS.Rows(Val(mRsError!Row)).Font.Color = vbRed
                End If
                .MoveNext
            Next
            Me.cmdOutput.Enabled = True
        ElseIf .RecordCount = 0 Then
            Me.cmdCheck.Enabled = False
            Me.cmdOutput.Enabled = False
        End If
    End With
    
    '��ʾ�ɹ�����Ŀ
    With VSFList
        .Rows = .Rows + 1
        strText = "����һ����" & mLngSumType & "�����ݣ��ɹ�����" & mLngType & "�����ݣ�ϸĿһ����" & mLngSumCount & "�����ݣ��ɹ�����" & mLngCount & "�����ݣ�"
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = strText
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
        .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
        
        .MergeCells = flexMergeFree
        .MergeRow(.Rows - 1) = True
    End With
    
    Me.prg.Visible = False
End Sub

Private Sub SaveType()
    Dim strSql As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItemID As Long
    Dim strSqlTemp As String  '�����ѯ��֤��sql���
    Dim strTemp As String
    Dim blnStop As Boolean
    
    
    On Error Resume Next
    blnStop = (Me.chkStop.Value = 1)
    With mobjWSType.UsedRange
        For lngRow = 3 To .Rows.Count
            If mintType = 1 Then
                lngItemID = sys.NextId("�շѷ���Ŀ¼")
                strSql = "ZL_�շѷ���Ŀ¼_INSERT(" & lngItemID & ","
            Else
                lngItemID = sys.NextId("���Ʒ���Ŀ¼")
                strSql = "ZL_���Ʒ���Ŀ¼_INSERT(" & lngItemID & ","
            End If

            For lngCol = 1 To .Columns.Count
                If mintType = 2 And lngCol = 1 Then
                    If .cells(lngRow, lngCol) = "����ҩ" Then
                        strTemp = "1,"
                    ElseIf .cells(lngRow, lngCol) = "�г�ҩ" Then
                        strTemp = "2,"
                    ElseIf .cells(lngRow, lngCol) = "�в�ҩ" Then
                        strTemp = "3,"
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "����ֵֻ��Ϊ����ҩ���г�ҩ���в�ҩ", "����"
                            
                            GoTo ErrHandle
                        End If
                    End If
                ElseIf (mintType = 2 And lngCol = 2) Or (mintType <> 2 And lngCol = 1) Then
                    If .cells(lngRow, lngCol) <> "" Then
                        If GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, IIf(mintType = 2, Val(strTemp), 7), False, True) <> 0 And .cells(lngRow, lngCol) <> "" Then
                            strSql = strSql & GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, IIf(mintType = 2, Val(strTemp), 7), False, True) & ","
                        Else
                            If blnStop Then
                                Exit Sub
                            Else
                                GoTo ErrHandle
                            End If
                        End If
                    Else
                        strSql = strSql & "null,"
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
            Next
            '����
            strSql = strSql & "'" & zlStr.GetCodeByVB(.cells(lngRow, lngCol - 1)) & "',"

            '������Ŀ������
            If mintType = 2 Then
                strSql = strSql & strTemp
            ElseIf mintType = 3 Then
                strSql = strSql & "7,"
            End If

            strSql = strSql & "0)"

            zlDatabase.ExecuteProcedure strSql, "SaveType"
            If Err.Number <> 0 Then
                If blnStop Then
                    MsgBox "�������ݳ���������ֹ", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                Else
                    AddErr "�������", lngRow, lngCol, Err.Description, "����"
                    GoTo ErrHandle
                End If
            End If
        mLngType = mLngType + 1
ErrHandle:
        prg.Value = Int((lngRow - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub

Private Sub SaveData()
    Dim strSql As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngItemID As Long
    Dim strSqlTemp As String  '�����ѯ��֤��sql���
    Dim strTemp As String
    Dim int������Ŀ As Integer
    Dim strID As String
    Dim lng��ĿID As Long
    Dim str���� As String
    Dim blnStop As Boolean
    Dim dateNow As Date
    Dim strNo As String
    Dim str���� As String
    Dim rsTemp As Recordset
    Dim rs��� As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng����id As Long
    
    On Error Resume Next
    blnStop = chkStop.Value
    str���� = ""
    
    '��ȡ��ǰ���
    strSql = "Select ����,���� From �շ���Ŀ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            lngCol = mCollItemCols.Item("���")
'            mobjWS.cells(lngRow, lngCol) = "1111"
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If .cells(lngRow, lngCol) = "�Һ�" Then
                    int������Ŀ = 1
                ElseIf .cells(lngRow, lngCol) = "����" Then
                    int������Ŀ = 3
                Else
                    int������Ŀ = 0
                End If
                
                rsTemp.Filter = "����='" & .cells(lngRow, lngCol) & "'"

                If Not rsTemp.EOF Then
                    strTemp = rsTemp!����
                Else
                    '��𲻴��ڵĴ���
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "��𲻴���", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            strID = sys.NextId("�շ���ĿĿ¼")
            strSql = "zl_�շ�ϸĿ_insert(" & int������Ŀ & "," & strID & ",'" & strTemp & "',"
            
            '����
            lngCol = mCollItemCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '��ʶ����
            lngCol = mCollItemCols.Item("��ʶ����")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '��ʶ����
            lngCol = mCollItemCols.Item("��ʶ����")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '��ѡ��
            lngCol = mCollItemCols.Item("��ѡ��")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '����
            lngCol = mCollItemCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If


            '�ϼ�����
            lngCol = mCollItemCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ϼ�����Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                lng����id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 0)
                If lng����id <> 0 Then
                    strSql = strSql & lng����id & ","
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        GoTo ErrHandle
                    End If
                End If
            End If

            '���
            lngCol = mCollItemCols.Item("���")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & Replace(.cells(lngRow, lngCol), "'", "''") & "',"
            End If

            '˵��
            strSql = strSql & "'',"


            '���㵥λ
            lngCol = mCollItemCols.Item("���㵥λ")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '��������
            lngCol = mCollItemCols.Item("��������")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "'',"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '���ηѱ��Ƿ��ۣ�
            strSql = strSql & "0,"
            
            '�Ƿ���
            lngCol = mCollItemCols.Item("�Ƿ���")
            If .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                strSql = strSql & "0,"
            End If
    
            '�Ӱ�Ӽۣ�ִ�п���
            strSql = strSql & "0,0,"
            
            '�������
            lngCol = mCollItemCols.Item("�������")
            strSql = strSql & Val(.cells(lngRow, lngCol)) & ","

            'ժҪ
            strSql = strSql & "'',"

            '�ּۣ�ԭ��
            lngCol = mCollItemCols.Item("�ּ�")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ּ�Ϊ��", "��ϸ"
                    
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    strSql = strSql & .cells(lngRow, lngCol) & ",0,"
                Else
                    If blnStop Then
                        
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ּ�ֻ��Ϊ����", "��ϸ"
                        
                        GoTo ErrHandle
                    End If
                End If
            End If

            '����
            lngCol = mCollItemCols.Item("����")
            If zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) <> "" Then
                str���� = "1''" & .cells(lngRow, lngCol) & "''1''" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "''"
            End If

            If zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) <> "" Then
                str���� = str���� & "1''" & .cells(lngRow, lngCol) & "''2''" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "''"
            End If

            strSql = strSql & "'" & str���� & "',"

            '¼��������¼��������Χ������ȷ�ϣ�����ȷ�Ϸ�Χ���Զ�����,վ�㣬������Ŀ
            strSql = strSql & "0,0,0,0,0,null,'')"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            'Ϊ���ϸ��²���
            If strTemp = "M" Then
                strSql = "ZL_�շ�ϸĿ_���ϲ���(" & strID & ",'" & Replace(.cells(lngRow, lngCol), "'", "''") & "')"
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            End If

            '�����շѼ�Ŀ��Ϣ
            lng��ĿID = sys.NextId("�շѼ�Ŀ")
            lngCol = mCollItemCols.Item("�ּ�")
            dateNow = sys.Currentdate
            strNo = sys.GetNextNo(9)
            strSql = ""
            strSql = "zl_�շѼ�Ŀ_insert(" & lng��ĿID & ",null," & strID & "," & GetTypeID(.cells(lngRow, mCollItemCols.Item("������Ŀ")), lngRow, mCollItemCols.Item("������Ŀ"), 0, True) & ",0," & .cells(lngRow, lngCol) & ",null,null,'��ʼ����',null,'" & gstrUserName & " ',to_date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss'),1,'" & strNo & "',1)"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "�������ݳ���������ֹ", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        AddErr "�������", lngRow, lngCol, Err.Description, "��ϸ"
                        gcnOracle.RollbackTrans
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
        

    End With
End Sub

Private Sub SaveMedi()
    Dim blnStop As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTemp As String
    Dim strSql As String
    Dim lngҩ��id As Long
    Dim lngҩƷID As Long
    Dim lng����id As Long
    Dim str��Ӧ������ As String
    Dim str��Ӧ�����֤�� As String
    Dim str��Ӧ�����֤Ч�� As String
    Dim intType As Integer
    Dim str���� As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng����id As Long
    Dim intKind As Integer
    
    On Error Resume Next
    blnStop = chkStop.Value
    str���� = ""
    
    '��ȡƷ�ֱ���
    strSql = "Select id,���� From ������ĿĿ¼ Where ��� In ('5','6','7')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    Do While Not rsTemp.EOF
        str���� = str���� & rsTemp!ID & "[" & rsTemp!���� & "],"
        rsTemp.MoveNext
    Loop
    
    '��ȡ��Ӧ��
    strSql = "Select ���� From ��Ӧ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "CheckSupplier")
    
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            '���
            lngCol = mCollItemCols.Item("Ʒ�ֱ���")
            If InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") <= 0 Then
                lngҩ��id = sys.NextId("������ĿĿ¼")
                str���� = str���� & lngҩ��id & "[" & .cells(lngRow, lngCol) & "],"
                lngCol = mCollItemCols.Item("���")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "���Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    If .cells(lngRow, lngCol) = "����ҩ" Then
                        strTemp = "5"
                        intKind = 1
                        strSql = "Zl_��ҩƷ��_Insert('" & strTemp & "',"
                    ElseIf .cells(lngRow, lngCol) = "�г�ҩ" Then
                        strTemp = "6"
                        intKind = 2
                        strSql = "Zl_��ҩƷ��_Insert('" & strTemp & "',"
                    ElseIf .cells(lngRow, lngCol) = "�в�ҩ" Then
                        strTemp = "7"
                        intKind = 3
                        strSql = "Zl_��ҩƷ��_Insert('" & strTemp & "',"
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "����ֵֻ��Ϊ����ҩ���г�ҩ���в�ҩ", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                End If

                '�ϼ�����
                lngCol = mCollItemCols.Item("����")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "�ϼ�����Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    lng����id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind)
                    If lng����id <> 0 Then
                        strSql = strSql & lng����id & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            GoTo ErrHandle
                        End If
                    End If
                End If
    
                'Ʒ��id
                strSql = strSql & lngҩ��id & ","

                '����
                lngCol = mCollItemCols.Item("Ʒ�ֱ���")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "Ʒ�ֱ���Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                '����
                lngCol = mCollItemCols.Item("Ʒ������")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "Ʒ������Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If

                'ƴ������
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "',"
    
                '��ʼ���
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "',"
    
                'Ӣ������
                strSql = strSql & "'',"
    
                '������λ
                lngCol = mCollItemCols.Item("������λ")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "������λΪ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If

                '����
                If strTemp <> "7" Then
                    lngCol = mCollItemCols.Item("����")
                    If .cells(lngRow, lngCol) = "" Then
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
                            GoTo ErrHandle
                        End If
                    Else
                        strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                    End If
                End If
                
                '�������,��ֵ����,��Դ���,��ҩ�ݴ�
                strSql = strSql & "5,1,1,1)"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                lngҩ��id = Mid(Mid(str����, 1, InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") - 1), InStrRev(Mid(str����, 1, InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") - 1), ",") + 1)
            End If
            
            strSql = ""
            '����ҩƷ���
            If strTemp = "5" Then
                strSql = "Zl_��ҩ���_Insert(" & lngҩ��id & ","
            ElseIf strTemp = "6" Then
                strSql = "Zl_��ҩ���_Insert(" & lngҩ��id & ","
            Else
                strSql = "Zl_��ҩ���_Insert(" & lngҩ��id & ","
            End If

            'ҩƷid
            lngҩƷID = sys.NextId("�շ���ĿĿ¼")
            strSql = strSql & lngҩƷID & ","

            '�����룺ҩƷƷ�ֱ����������1
            lngCol = mCollItemCols.Item("������")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "������Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '���
            lngCol = mCollItemCols.Item("ҩƷ���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "ҩƷ���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '����
            lngCol = mCollItemCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '��Ʒ��,ƴ������,��ʼ���,������,��ʶ��,ҩƷ��Դ
            strSql = strSql & "'','','','','','',"

            '��׼�ĺ�
'            lngCol = mCollItemCols.Item("��׼�ĺ�")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "'',"
'            Else
'
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            strSql = strSql & "'',"

            'ע���̱�
'            lngCol = mCollItemCols.Item("ע���̱�")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "'',"
'            Else
'
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            strSql = strSql & "'',"

            '�ۼ۵�λ
            lngCol = mCollItemCols.Item("�ۼ۵�λ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼ۵�λΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

            '�ۼۼ�������ϵ��������ϵ��
            lngCol = mCollItemCols.Item("�ۼۼ�������ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼۼ�������ϵ��Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If

            '���ﵥλ
            lngCol = mCollItemCols.Item("���ﵥλ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "���ﵥλΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else

                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

             '���ﵥλ����ϵ���������װ
            lngCol = mCollItemCols.Item("���ﵥλ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "���ﵥλ����ϵ��Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If
            
            If strTemp <> "7" Then
                'סԺ��λ
                lngCol = mCollItemCols.Item("סԺ��λ")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "סԺ��λΪ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
    
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                 'סԺ��λ����ϵ����סԺ��װ
                lngCol = mCollItemCols.Item("סԺ��λ����ϵ��")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "סԺ��λ����ϵ��Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & .cells(lngRow, lngCol) & ","
                End If
            End If

            'ҩ�ⵥλ
            lngCol = mCollItemCols.Item("ҩ�ⵥλ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "ҩ�ⵥλΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If

             'ҩ�ⵥλ����ϵ����ҩ���װ
            lngCol = mCollItemCols.Item("ҩ���װ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "ҩ���װ����ϵ��Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If

            '���쵥λ,���췧ֵ
            strSql = strSql & "null,null,"

            '�Ƿ���
            lngCol = mCollItemCols.Item("�Ƿ���")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "�Ƿ��۵�ֵֻ��Ϊ�ջ򡮡̡�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If

            'ָ�������ۣ��ɱ���
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ɱ��۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ɱ���ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '����
            strSql = strSql & "100,"

            'ָ�����ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ۼ۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼ�ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            'ָ�������,����ѱ���,ҩ�ۼ���,��������
            strSql = strSql & "13.0435,0,null,null,"

             '�������
            lngCol = mCollItemCols.Item("�������")
            
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 And Val(.cells(lngRow, lngCol)) <= 3 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�������ֻ��Ϊ0-3�����ֻ���Ϊ��", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�������ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            'Gmp��֤,�б�ҩƷ,���ηѱ�,
            strSql = strSql & "0,0,0,"
            
            'סԺ�ɷ����
            lngCol = mCollItemCols.Item("סԺ�ɷ����")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            Else
                strSql = strSql & Mid(.cells(lngRow, lngCol), 1, 1) & ","
            End If

            'ҩ�����
            lngCol = mCollItemCols.Item("ҩ�����")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "ҩ�������ֵֻ��Ϊ�ջ򡮡̡�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If

            'ҩ������
            lngCol = mCollItemCols.Item("ҩ������")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "ҩ��������ֵֻ��Ϊ�ջ򡮡̡�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If

            '���Ч��
            lngCol = mCollItemCols.Item("Ч��(��)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "Ч��(��)��ֵֻ�ܴ���0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "Ч��(��)��ֵֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            Else
                strSql = strSql & "null,"
            End If

            '���������
            strSql = strSql & "100,"

            '�ɱ���
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ɱ��۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "�ɱ���ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '�ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ۼ۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "�ۼ�ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '������Ŀ
            lngCol = mCollItemCols.Item("������Ŀ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "������ĿΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                lng����id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind, True)
                If lng����id <> 0 Then
                    strSql = strSql & GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, intKind, True)
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "������Ŀ������", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            If strTemp = "7" Then
                '��ͬ��λid,˵��,��̬����,��ҩ����,��ѡ��,��ֵ˰��,����ҩ��,��ҩ��̬,վ��,�Ƿ񳣱�,������Ŀ
                strSql = strSql & ",Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,"
            Else
                '��ͬ��λid,˵��,��̬����,��ҩ����,��ѡ��,��ֵ˰��,����ҩ��,վ��,�Ƿ񳣱�,�洢�¶�,�洢����,��ҩ����,�Ƿ�������,����,������Ŀ
                strSql = strSql & ",Null,Null,0,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,"
            End If
            
            '����ɷ����
            lngCol = mCollItemCols.Item("����ɷ����")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0"
            Else
                strSql = strSql & Mid(.cells(lngRow, lngCol), 1, 1)
            End If
            
            
            strSql = strSql & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            '���湩Ӧ��
            lngCol = mCollItemCols.Item("��Ӧ������")
            If .cells(lngRow, lngCol) <> "" Then
                str��Ӧ������ = .cells(lngRow, lngCol)
                
                rsTemp.Filter = "����='" & str��Ӧ������ & "'"
                
                If rsTemp.EOF Then
                    lngCol = mCollItemCols.Item("��Ӧ�����֤��")
                    str��Ӧ�����֤�� = .cells(lngRow, lngCol)
                    
                    lngCol = mCollItemCols.Item("��Ӧ�����֤Ч��")
                    str��Ӧ�����֤Ч�� = .cells(lngRow, lngCol)
                    
                    strSql = ""
                    intType = CheckSupplier(str��Ӧ������, str��Ӧ�����֤��, str��Ӧ�����֤Ч��, lngRow, lngCol, strSql)
                    If intType = 1 Then
                        Exit Sub
                    ElseIf intType = 2 Then
                        GoTo ErrHandle
                    End If
                    
                    If strSql <> "" Then
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = strSql
                    End If
                End If
            End If
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "�������ݳ���������ֹ", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        gcnOracle.RollbackTrans
                        AddErr "�������", lngRow, lngCol, Err.Description, "��ϸ"
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub

Private Function CheckSupplier(ByVal strVal As String, ByVal str���֤�� As String, ByVal str���֤Ч�� As String, ByVal lngRow As Long, ByVal lngCol As Long, ByRef strSql As String) As Integer
'�������Ƿ���ڣ��������Ƿ�����
    Dim rsTemp As Recordset
    Dim lngId As Long
    Dim blnStop As Boolean
    Dim strTemp As String
    

    On Error Resume Next
    blnStop = (Me.chkStop.Value = 1)
    If chk��Ӧ��.Value = 1 Then

        strTemp = "Select max(����) ����,zlSpellCode([1], 10) ���� From ��Ӧ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strTemp, "CheckSupplier", strVal)
        
        lngId = sys.NextId("��Ӧ��")
        strSql = "zl_��Ӧ��_insert("
        'id
        strSql = strSql & lngId & ","
        '�ϼ�id
        strSql = strSql & "null,"
        '����
        If IsNull(rsTemp!����) Then
            strSql = strSql & "'01',"
        Else
            strSql = strSql & "'" & Format(Val(rsTemp!����) + 1, String(Len(Trim(rsTemp!����)), "0")) & "',"
        End If
        '����
        strSql = strSql & "'" & strVal & "',"
        
'            ����
        strSql = strSql & "'" & rsTemp!���� & "',"
'            ��ַ
        strSql = strSql & "null,"
'            �绰
        strSql = strSql & "null,"
'            ��������
        strSql = strSql & "null,"
'            �ʺ�
        strSql = strSql & "null,"
'            ��ϵ��
        strSql = strSql & "null,"
'            ˰��ǼǺ�
        strSql = strSql & "null,"
'            ���֤��
        strSql = strSql & "'" & str���֤�� & "',"
'            ���֤Ч��
        strSql = strSql & IIf(str���֤Ч�� <> "", "to_date('" & str���֤Ч�� & "','YYYY-MM-dd'),", "Null,")
'            ִ�պ�
        strSql = strSql & "null,"
'            ִ��Ч��
        strSql = strSql & "null,"
'            ��Ȩ��
        strSql = strSql & "null,"
'            ��Ȩ��
        strSql = strSql & "null,"
'            ����
        strSql = strSql & IIf(mintType = 2, "'10000',", "'00001',")
        strSql = strSql & "0,0,Null,null,null,null,null,Null,Null,1,0)"
    Else
        If blnStop Then
            CheckSupplier = 1
        Else
            AddErr "��Ӧ�̲�����", lngRow, lngCol, "��ǰ��Ӧ�̲�����", "��ϸ"
            CheckSupplier = 2
        End If
    End If
End Function

Private Function GetTypeID(ByVal strVal As String, ByVal lngRow As Long, ByVal lngCol As Long, ByVal intType As Integer, Optional ByVal blnType As Boolean, Optional ByVal blnTypePage As Boolean) As Long
    Dim strSql As String
    Dim strType As String
    Dim strSecType As String
    Dim rsTemp As Recordset
    Dim Count As Long
    Dim i As Integer
    Dim strTemp As String

    On Error GoTo ErrHandle
    '�������Ƿ�ֻ��һ��
    If InStr(1, strVal, "\") > 1 Then
        strType = Mid(strVal, InStrRev(strVal, "\") + 1)
        strSecType = Mid(strVal, 1, InStrRev(strVal, "\") - 1)
    Else
        strType = strVal
        strSecType = ""
    End If

    If strSecType = "" Then
        '����ֻ��һ�������
        strSql = "Select id,���� " & _
                 "From �շѷ���Ŀ¼ " & _
                 "Where ���� = [1] "
    Else
        strSql = "Select id,���� From �շѷ���Ŀ¼" & vbNewLine & _
                "Where ���� = [1] And �ϼ�id In (Select �ϼ�id From �շѷ���Ŀ¼ Where ���� = [2] "
        Count = 2
        For i = UBound(Split(strSecType, "\")) - 1 To 0 Step -1
            If i < UBound(Split(strSecType, "\")) Then
                Count = Count + 1
                strTemp = "And �ϼ�id In (Select �ϼ�id From �շѷ���Ŀ¼ Where ���� = '" & Split(strSecType, "\")(i) & "'"
                strSql = strSql & strTemp
            End If
        Next
        strSecType = Split(strSecType, "\")(UBound(Split(strSecType, "\")))
        strSql = strSql & String(Count - 1, ")")
    End If

    If mintType <> 1 And blnType = False Then
        strSql = strSql & " and ����=[3]"
        strSql = Replace(strSql, "�շѷ���Ŀ¼", "���Ʒ���Ŀ¼")
    ElseIf blnType = True Then
        strSql = Replace(strSql, "�շѷ���Ŀ¼", "������Ŀ")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "GetTypeID", strType, strSecType, intType)

    If Not blnType Then
        If rsTemp.RecordCount = 1 Then
            GetTypeID = rsTemp!ID
        ElseIf rsTemp.RecordCount = 0 Then
            GetTypeID = 0
            AddErr "�ϼ����಻����", lngRow, lngCol, "�ϼ����ࡾ" & strVal & "��������", IIf(blnTypePage, "����", "��ϸ")
        Else
            GetTypeID = rsTemp!ID
            AddErr "�ϼ����಻Ψһ", lngRow, lngCol, "�ϼ����ࡾ" & strVal & "���ж����Ĭ���ڱ���Ϊ��" & rsTemp!���� & "���ķ�����", IIf(blnTypePage, "����", "��ϸ")
        End If
    Else
        If rsTemp.RecordCount = 1 Then
            GetTypeID = rsTemp!ID
        ElseIf rsTemp.RecordCount = 0 Then
            GetTypeID = 0
            AddErr "������Ŀ������", lngRow, lngCol, "������Ŀ��" & strVal & "��������", "��ϸ"
        Else
            GetTypeID = rsTemp!ID
            AddErr "������Ŀ��Ψһ", lngRow, lngCol, "������Ŀ��" & strVal & "���ж����Ĭ���ڱ���Ϊ��" & rsTemp!���� & "����������Ŀ��", "��ϸ"
        End If
    End If
    Exit Function
ErrHandle:
    '�ռ�������Ϣ
End Function

Private Sub SaveStuff()
    Dim blnStop As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTemp As String
    Dim strSql As String
    Dim lng����ID As Long
    Dim lng����ID As Long
    Dim lng����id As Long
    Dim str��Ӧ������ As String
    Dim str��Ӧ�����֤�� As String
    Dim str��Ӧ�����֤Ч�� As String
    Dim intType As Integer
    Dim str���� As String
    Dim rsTemp As Recordset
    Dim arrSql As Variant
    Dim i As Integer
    Dim lng����id As Long
    
    On Error Resume Next
    blnStop = chkStop.Value
    str���� = ""
    
    '��ȡƷ�ֱ���
    strSql = "Select id,���� From ������ĿĿ¼ Where ��� ='4'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "SaveData")
        
    Do While Not rsTemp.EOF
        str���� = str���� & rsTemp!ID & "[" & rsTemp!���� & "],"
        rsTemp.MoveNext
    Loop
    
    '��ȡ��Ӧ��
    strSql = "Select ���� From ��Ӧ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "CheckSupplier")
    
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            arrSql = Array()
            lngCol = mCollItemCols.Item("Ʒ�ֱ���")
            If InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") <= 0 Then
                lng����ID = sys.NextId("������ĿĿ¼")
                str���� = str���� & lng����ID & "[" & .cells(lngRow, lngCol) & "],"
                '����Ʒ��
                strSql = "zl_����Ʒ��_INSERT("
    
                '�ϼ�����
                lngCol = mCollItemCols.Item("����")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    lng����id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 7)
                    If lng����id <> 0 Then
                        strSql = strSql & lng����id & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            GoTo ErrHandle
                        End If
                    End If
                End If
    
                'Ʒ��id
                strSql = strSql & lng����ID & ","
    
                'Ʒ�ֱ���
                lngCol = mCollItemCols.Item("Ʒ�ֱ���")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "Ʒ�ֱ���Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                '����
                lngCol = mCollItemCols.Item("Ʒ������")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "Ʒ������Ϊ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
                
                '��λ
                lngCol = mCollItemCols.Item("ɢװ��λ")
                If .cells(lngRow, lngCol) = "" Then
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "��ֵ����", lngRow, lngCol, "ɢװ��λΪ��", "��ϸ"
                        GoTo ErrHandle
                    End If
                Else
                    strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
                End If
    
                'ƴ������
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol)) & "',"
    
                '��ʼ���
                strSql = strSql & "'" & zlStr.GetCodeByORCL(.cells(lngRow, lngCol), True) & "',"
    
                'Ӣ������
                strSql = strSql & "'',"
                
                'վ��
                strSql = strSql & "null,"
                
                '�����Ա�
                strSql = strSql & "0,"
                
                '����
                strSql = strSql & "'')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = strSql
            Else
                lng����ID = Mid(Mid(str����, 1, InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") - 1), InStrRev(Mid(str����, 1, InStr(1, str����, "[" & .cells(lngRow, lngCol) & "]") - 1), ",") + 1)
            End If
            
            strSql = ""
            '���Ĺ��
            strSql = "Zl_��������_Insert("
            
            '����id
            strSql = strSql & lng����ID & ","
            
            '����id
            lng����ID = sys.NextId("�շ���ĿĿ¼")
            strSql = strSql & lng����ID & ","
            
            '����:�������ɵĹ���
            lngCol = mCollItemCols.Item("������")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "������Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            
            '���
            lngCol = mCollItemCols.Item("���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '����
            lngCol = mCollItemCols.Item("������")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '��ʶ����
            lngCol = mCollItemCols.Item("��ʶ����")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '��ʶ����
            lngCol = mCollItemCols.Item("��ʶ����")
            If .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "null,"
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '��ѡ��
             strSql = strSql & "null,"
'            lngCol = mCollItemCols.Item("��ѡ��")
'            If .cells(lngRow, lngCol) = "" Then
'                strSQL = strSQL & "null,"
'            Else
'                strSQL = strSQL & "'" & .cells(lngRow, lngCol) & "',"
'            End If
            
            '������Դ
            strSql = strSql & "'',"
            
            '��Դ���
            strSql = strSql & "'',"
            
            'ɢװ��λ
            lngCol = mCollItemCols.Item("ɢװ��λ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "ɢװ��λΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            '��װ��λ
            lngCol = mCollItemCols.Item("��װ��λ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "��װ��λΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            End If
            
            'ɢװ��װ����ϵ��
            lngCol = mCollItemCols.Item("ɢװ��װ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "ɢװ��װ����ϵ��Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                strSql = strSql & .cells(lngRow, lngCol) & ","
            End If
            
            '�Ƿ��ۣ���
            lngCol = mCollItemCols.Item("�Ƿ���")
            If .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            ElseIf .cells(lngRow, lngCol) = "" Then
                strSql = strSql & "0,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "�Ƿ��۵�ֵֻ��Ϊ���̡����߿�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If
            
            'ָ��������
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ɱ��۵�ֵ�������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ɱ��۵�ֵֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '����
            strSql = strSql & "100,"
            
            'ָ�����ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ۼ۵�ֵ�������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼ۵�ֵֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            'ָ�������
            strSql = strSql & "13.0435,"
            
            '��������
            strSql = strSql & "null,"
            
            '�������
            lngCol = mCollItemCols.Item("�������")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 And Val(.cells(lngRow, lngCol)) <= 3 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�������ֻ��Ϊ0-3�����ֻ���Ϊ��", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�������ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '���ηѱ�
            strSql = strSql & "0,"
            
            '���Ŀⷿ����
            lngCol = mCollItemCols.Item("���Ŀⷿ����")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "���Ŀⷿ������ֵֻ��Ϊ�ջ򡮡̡�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If

            '���ϲ��ŷ���
            lngCol = mCollItemCols.Item("���ϲ��ŷ���")
            If .cells(lngRow, lngCol) = "" Then
               strSql = strSql & "0,"
            ElseIf .cells(lngRow, lngCol) = "��" Then
                strSql = strSql & "1,"
            Else
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "ֵ����", lngRow, lngCol, "���ϲ��ŷ�����ֵֻ��Ϊ�ջ򡮡̡�", "��ϸ"
                    GoTo ErrHandle
                End If
            End If

            '���Ч��
            lngCol = mCollItemCols.Item("Ч��(��)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) >= 0 Then
                        strSql = strSql & .cells(lngRow, lngCol) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "Ч��(��)��ֵֻ�ܴ���0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "Ч��(��)��ֵֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            Else
                strSql = strSql & "null,"
            End If
            
            
            '���Ч��
            strSql = strSql & "Null,"
            
            '�޾��Բ���
            strSql = strSql & "0,"
            
            'һ���Բ���
            strSql = strSql & "0,"
            
            'ԭ����
            strSql = strSql & "0,"
            
            '���������
            strSql = strSql & "100,"

            '�ɱ���
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ɱ��۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "�ɱ���ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '��������
            strSql = strSql & "0,"
            
            '�������
            strSql = strSql & "0,"

            '�ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If Val(.cells(lngRow, lngCol)) > 0 Then
                        strSql = strSql & Val(.cells(lngRow, lngCol)) & ","
                    Else
                        If blnStop Then
                            Exit Sub
                        Else
                            AddErr "ֵ����", lngRow, lngCol, "�ۼ۱������0", "��ϸ"
                            GoTo ErrHandle
                        End If
                    End If
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼ�ֻ��Ϊ����", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If

            '������Ŀ
            lngCol = mCollItemCols.Item("������Ŀ")
            If .cells(lngRow, lngCol) = "" Then
                If blnStop Then
                    Exit Sub
                Else
                    AddErr "��ֵ����", lngRow, lngCol, "������ĿΪ��", "��ϸ"
                    GoTo ErrHandle
                End If
            Else
                lng����id = GetTypeID(.cells(lngRow, lngCol), lngRow, lngCol, 7, True)
                If lng����id <> 0 Then
                    strSql = strSql & lng����id & ","
                Else
                    If blnStop Then
                        Exit Sub
                    Else
                        AddErr "ֵ����", lngRow, lngCol, "������Ŀ������", "��ϸ"
                        GoTo ErrHandle
                    End If
                End If
            End If
            
            '��׼�ĺ�
            lngCol = mCollItemCols.Item("��׼�ĺ�")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            
            '��Ʒע���̱�
            lngCol = mCollItemCols.Item("��Ʒע���̱�")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            'ע��֤��
            lngCol = mCollItemCols.Item("ע��֤��")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '���֤��
            lngCol = mCollItemCols.Item("���֤��")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & .cells(lngRow, lngCol) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '���֤��Ч��
            lngCol = mCollItemCols.Item("���֤Ч��")
            If .cells(lngRow, lngCol) <> "" Then
                strSql = strSql & "'" & CDate(.cells(lngRow, lngCol)) & "',"
            Else
                strSql = strSql & "null,"
            End If
            
            '���ʷ���
            strSql = strSql & "'',"
                
            '�洢����
            strSql = strSql & "'',"
            
            '���ٲ���
            strSql = strSql & "0,"
            
            'վ��
            strSql = strSql & "'')"
            'Ʒ��
            'ƴ��
            '���
            '��ֵ˰��
            '˵��
            '��ֵ����
            '�������
            '������Ŀ
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = strSql
            
            '���湩Ӧ��
            lngCol = mCollItemCols.Item("��Ӧ������")
            If .cells(lngRow, lngCol) <> "" Then
                str��Ӧ������ = .cells(lngRow, lngCol)
                rsTemp.Filter = "����='" & str��Ӧ������ & "'"
                
                If rsTemp.EOF Then
                    lngCol = mCollItemCols.Item("��Ӧ�����֤��")
                    str��Ӧ�����֤�� = .cells(lngRow, lngCol)
                    
                    lngCol = mCollItemCols.Item("��Ӧ�����֤Ч��")
                    str��Ӧ�����֤Ч�� = .cells(lngRow, lngCol)
                    
                    strSql = ""
                    intType = CheckSupplier(str��Ӧ������, str��Ӧ�����֤��, str��Ӧ�����֤Ч��, lngRow, lngCol, strSql)
                    If intType = 1 Then
                        Exit Sub
                    ElseIf intType = 2 Then
                        GoTo ErrHandle
                    End If
                    
                    If strSql <> "" Then
                        ReDim Preserve arrSql(UBound(arrSql) + 1)
                        arrSql(UBound(arrSql)) = strSql
                    End If
                End If
            End If
            
            gcnOracle.BeginTrans
            For i = 0 To UBound(arrSql)
                Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SavaData")
                If Err.Number <> 0 Then
                    If blnStop Then
                        gcnOracle.RollbackTrans
                        MsgBox "�������ݳ���������ֹ", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    Else
                        
                        AddErr "�������", lngRow, lngCol, Err.Description, "��ϸ"
                        gcnOracle.RollbackTrans
                        GoTo ErrHandle
                    End If
                    
                End If
            Next
            gcnOracle.CommitTrans
            mLngCount = mLngCount + 1
ErrHandle:
        prg.Value = Int((lngRow + mLngSumType - 2) / (mLngSumType + mLngSumCount) * 100)
        Next
    End With
End Sub


Private Sub cmdOutput_Click()
    On Error GoTo ErrHandle
    Me.VSFList.SaveGrid "C:\APPSOFT\�����ļ�\������Ϣ.xls", flexFileExcel, True
    MsgBox "������Ϣ�����ɹ���", vbInformation + vbCritical, gstrSysName
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    Me.lblFileName.Caption = mstrIn
    dlgOpenFile.FileName = mstrIn

    Me.VSFList.RowHeight(0) = 300
    VSFList.Cell(flexcpFontBold, 0, 0, 0, VSFList.Cols - 1) = True
    Exit Sub
ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub OpenFile()
    dlgOpenFile.Filter = "xlsx|*.xlsx|xls|*.xls"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        lblFileName.Caption = dlgOpenFile.FileName
    End If
    
End Sub

Private Function CheckExcel() As Boolean
'������ʽ
    
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String
    Dim strItem As String
    
    Set mCollTypeCols = New Collection
    Set mCollItemCols = New Collection
    If mintType = 1 Then
        strType = MSTRCHARGETYPE
        strItem = mstrCharge
    ElseIf mintType = 2 Then
        strType = MSTRMEDICALTYPE
        strItem = MSTRMEDICAL
    Else
        strType = MSTRSTUFFTYPE
        strItem = MSTRSTUFF
    End If
    
    On Error GoTo ErrHandle
    Set mobjWB = mobjXLS.Workbooks.Open(Me.lblFileName.Caption)
    mblnExists = True
    
    Set mobjWSType = mobjWB.Sheets(1)
    With mobjWSType.UsedRange
        '��������˳����
        If .Columns.Count <> UBound(Split(strType, ",")) + 1 Then
            MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "�ⲿ�ļ�����������ȷ�����飡", vbInformation, gstrSysName
            CheckExcel = False
            Exit Function
        End If

        For lngCol = 1 To .Columns.Count
            If .cells(MINTTITLE, lngCol) <> Split(strType, ",")(lngCol - 1) Then
                MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "�ⲿ�ļ�������˳����ȷ�����飡", vbInformation, gstrSysName
                CheckExcel = False
                Exit Function
            End If
            Call mCollTypeCols.Add(lngCol, .cells(MINTTITLE, lngCol))
        Next
        
        mLngSumType = .Rows.Count - MINTTITLE
    End With
    
    Set mobjWS = mobjWB.Sheets(2)
    With mobjWS.UsedRange
        '��������˳����
        If .Columns.Count <> UBound(Split(strItem, ",")) + 1 Then
            MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "�ⲿ�ļ�����������ȷ�����飡", vbInformation, gstrSysName
            CheckExcel = False
            Exit Function
        End If

        For lngCol = 1 To .Columns.Count
            If .cells(MINTTITLE, lngCol) <> Split(strItem, ",")(lngCol - 1) Then
                MsgBox "'" & Me.lblFileName.Caption & "'" & vbNewLine & vbNewLine & "�ⲿ�ļ�������˳����ȷ�����飡", vbInformation, gstrSysName
                CheckExcel = False
                Exit Function
            End If
            Call mCollItemCols.Add(lngCol, .cells(MINTTITLE, lngCol))
        Next
        
        mLngSumCount = .Rows.Count - MINTTITLE
    End With
    CheckExcel = True
    Exit Function
ErrHandle:
    MsgBox "����·��[" & Me.lblFileName.Caption & "]�µ��ļ��Ƿ���ڣ�", vbInformation, gstrSysName
    CheckExcel = False
End Function

Private Sub CheckPub()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long

    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
           '������
            If mintType = 1 Then
                lngCol = mCollItemCols.Item("����")
            Else
                lngCol = mCollItemCols.Item("Ʒ�ֱ���")
            End If
            
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��ֵ", "��ϸ"
            Else
                '���������Ƿ��зǷ��ַ�
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "ֵ����", lngRow, lngCol, "�����к��зǷ��ַ�", "��ϸ"
                    End If
                Next
            End If

            '�������
            If mintType = 1 Then
                lngCol = mCollItemCols.Item("����")
            Else
                lngCol = mCollItemCols.Item("Ʒ������")
            End If
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��ֵ", "��ϸ"
            End If

            '���������
            lngCol = mCollItemCols.Item("�������")
            If .cells(lngRow, lngCol) <> "" Then
                If CDbl(.cells(lngRow, lngCol)) > 3 Then
                    AddErr "ֵ����", lngRow, lngCol, "��������ֵ������Χ", "��ϸ"
                End If
            End If

            lngCol = mCollItemCols.Item("������Ŀ")
            '�������Ϊ�գ����ü��
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "������ĿΪ��", "��ϸ"
            End If
        Next
    End With
End Sub

Private Sub CheckCharge()
'�շ���Ŀ���
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strName As String   '�����Ѿ��������ϼ�����
    Dim strNotExit As String  '�����Ѿ�����Ҳ����ڵ��ϼ�����
    Dim strType As String     '����ֱ���ϼ�������
    Dim strSecType As String  '����ڶ����ϼ�������
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long
    '�����Ŀ����
     With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
        '�������Ƿ����
            lngCol = mCollItemCols.Item("���")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "���Ϊ��", "��ϸ"
            End If
            
            
            '�������Ƿ����
            lngCol = mCollItemCols.Item("����")
            '�������Ϊ�գ����ü��
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
            End If

            '����ּ�
            lngCol = mCollItemCols.Item("�ּ�")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ּ�Ϊ��ֵ", "��ϸ"
            Else
                If Not IsNumeric(.cells(lngRow, lngCol)) Then
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ּ۲�������", "��ϸ"
                ElseIf CDbl(.cells(lngRow, lngCol)) < 0 Then
                    AddErr "ֵ����", lngRow, lngCol, "�ּ۲���Ϊ����", "��ϸ"
                End If
            End If

            '���������
            lngCol = mCollItemCols.Item("�������")
            If .cells(lngRow, lngCol) <> "" Then
                If CInt(.cells(lngRow, lngCol)) > 3 Then
                    AddErr "ֵ����", lngRow, lngCol, "��������ֵ������Χ", "��ϸ"
                End If
            End If
        Next
    End With
End Sub

Private Sub CheckMedi()
'ҩƷ��Ŀ���
'���ҩƷ�������ݵĺϷ���
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strName As String   '�����Ѿ��������ϼ�����
    Dim strNotExit As String  '�����Ѿ�����Ҳ����ڵ��ϼ�����
    Dim strType As String     '����ֱ���ϼ�������
    Dim strSecType As String  '����ڶ����ϼ�������
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '���ҩƷ����
    '���ҩƷ��Ŀ¼�͹����Ϣ
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            '��������Ϣ
            lngCol = mCollItemCols.Item("���")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
            ElseIf .cells(lngRow, lngCol) <> "����ҩ" And .cells(lngRow, lngCol) <> "�г�ҩ" And .cells(lngRow, lngCol) <> "�в�ҩ" Then
                AddErr "ֵ����", lngRow, lngCol, "����ֻ��Ϊ����ҩ���г�ҩ�����в�ҩ", "��ϸ"
            End If

            '��������Ϣ
            lngCol = mCollItemCols.Item("����")
            '�������Ϊ�գ����ü��
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
            End If
            
            '������
            lngCol = mCollItemCols.Item("������")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "������Ϊ��ֵ", "��ϸ"
            Else
                '���������Ƿ��зǷ��ַ�
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "ֵ����", lngRow, lngCol, "�������к��зǷ��ַ�", "��ϸ"
                    End If
                Next
            End If
            
            '������
            lngCol = mCollItemCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��ֵ", "��ϸ"
            End If

            '����ۼۼ�������ϵ��
            lngCol = mCollItemCols.Item("�ۼۼ�������ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ۼۼ�������ϵ��Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "�ۼۼ�������ϵ����ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼۼ�������ϵ����ֵ����������", "��ϸ"
                End If
            End If


            '������ﵥλ����ϵ��
            lngCol = mCollItemCols.Item("���ﵥλ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "���ﵥλ����ϵ��Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "���ﵥλ����ϵ����ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "���ﵥλ����ϵ����ֵ����������", "��ϸ"
                End If
            End If

            '���סԺ��λת��ϵ��
            lngCol = mCollItemCols.Item("סԺ��λ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "סԺ��λת��ϵ��Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "סԺ��λת��ϵ����ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "סԺ��λת��ϵ����ֵ����������", "��ϸ"
                End If
            End If

            '���ҩ���װ����ϵ��
            lngCol = mCollItemCols.Item("ҩ���װ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "ҩ���װ����ϵ��Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "ҩ���װ����ϵ����ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "ҩ���װ����ϵ����ֵ����������", "��ϸ"
                End If
            End If

            '���ɱ���
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "�ɱ��۵�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ɱ��۵�ֵ����������", "��ϸ"
                End If
            End If

            '����ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "�ۼ۵�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼ۵�ֵ����������", "��ϸ"
                End If
            End If


            '���Ч��
            lngCol = mCollItemCols.Item("Ч��(��)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "Ч�ڵ�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "Ч�ڵ�ֵ����������", "��ϸ"
                End If
            End If
        Next
    End With
End Sub

Private Sub CheckStuff()
'������Ŀ���
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String     '����ֱ���ϼ�������
    Dim strSecType As String  '����ڶ����ϼ�������
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '���������Ŀ����
    With mobjWS.UsedRange
        For lngRow = 3 To .Rows.Count
            '��������Ϣ
            lngCol = mCollItemCols.Item("����")
            '�������Ϊ�գ����ü��
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "��ϸ"
            End If
            
            '������
            lngCol = mCollItemCols.Item("������")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "������Ϊ��ֵ", "��ϸ"
            Else
                '���������Ƿ��зǷ��ַ�
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "ֵ����", lngRow, lngCol, "�������к��зǷ��ַ�", "��ϸ"
                    End If
                Next
            End If
            
            '�����
            lngCol = mCollItemCols.Item("���")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "���Ϊ��", "��ϸ"
            End If
            
            'ɢװ��λ
            lngCol = mCollItemCols.Item("ɢװ��λ")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "ɢװ��λΪ��", "��ϸ"
            End If
            
            '��װ��λ
            lngCol = mCollItemCols.Item("��װ��λ")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "��װ��λΪ��", "��ϸ"
            End If
            
            '����ϵ��
            lngCol = mCollItemCols.Item("ɢװ��װ����ϵ��")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "ɢװ��װ����ϵ��Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "ɢװ��װ����ϵ����ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "ɢװ��װ����ϵ����ֵ����������", "��ϸ"
                End If
            End If
            
            '�ɱ���
            lngCol = mCollItemCols.Item("�ɱ���")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ɱ���Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "�ɱ��۵�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ɱ��۵�ֵ����������", "��ϸ"
                End If
            End If

            '����ۼ�
            lngCol = mCollItemCols.Item("�ۼ�")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "�ۼ�Ϊ��ֵ", "��ϸ"
            Else
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "�ۼ۵�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "�ۼ۵�ֵ����������", "��ϸ"
                End If
            End If


            '���Ч��
            lngCol = mCollItemCols.Item("Ч��(��)")
            If .cells(lngRow, lngCol) <> "" Then
                If IsNumeric(.cells(lngRow, lngCol)) Then
                    If CDbl(.cells(lngRow, lngCol)) <= 0 Then
                        AddErr "ֵ����", lngRow, lngCol, "Ч�ڵ�ֵ�������0", "��ϸ"
                    End If
                Else
                    AddErr "ֵ���ʹ���", lngRow, lngCol, "Ч�ڵ�ֵ����������", "��ϸ"
                End If
            End If
        Next
    End With
End Sub


Private Sub CheckType(ByVal intType As Integer)
'������Ŀ���
'intType��1-�շ���Ŀ���࣬2-ҩƷ��Ŀ���࣬3-������Ŀ����
    Dim lngCol As Long
    Dim lngRow As Long
    Dim i As Integer
    Dim strType As String     '����ֱ���ϼ�������
    Dim strSecType As String  '����ڶ����ϼ�������
    Dim rsTemp As Recordset
    Dim strSql As String
    Dim lngTemp As Long
    Dim strTemp As String
    Dim Count As Long

    '�����ಿ��
    With mobjWSType.UsedRange
        For lngRow = 3 To .Rows.Count
            If intType = 2 Then
                '��������Ϣ
                lngCol = mCollTypeCols.Item("���")
                If .cells(lngRow, lngCol) = "" Then
                    AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��", "����"
                ElseIf .cells(lngRow, lngCol) <> "����ҩ" And .cells(lngRow, lngCol) <> "�г�ҩ" And .cells(lngRow, lngCol) <> "�в�ҩ" Then
                    AddErr "ֵ����", lngRow, lngCol, "����ֻ��Ϊ����ҩ���г�ҩ�����в�ҩ", "����"
                End If
            End If

            '������
            lngCol = mCollTypeCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��ֵ", "����"
            Else
                '���������Ƿ��зǷ��ַ�
                For i = 1 To Len(.cells(lngRow, lngCol))
                    If InStr(1, "QWERTYUIOPASDFGHJKLZXCVBNM0123456789_", UCase(Mid(.cells(lngRow, lngCol), i, 1))) < 1 Then
                        AddErr "ֵ����", lngRow, lngCol, "�����к��зǷ��ַ�", "����"
                    End If
                Next
            End If

            '�������
            lngCol = mCollTypeCols.Item("����")
            If .cells(lngRow, lngCol) = "" Then
                AddErr "��ֵ����", lngRow, lngCol, "����Ϊ��ֵ", "����"
            End If
        Next
    End With
End Sub


Public Sub ShowMe(ByVal intType As Integer, ByVal frmParent As Form)
    On Error Resume Next
    Set mobjXLS = CreateObject("Excel.Application")
    
    If mobjXLS Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    
    mobjXLS.DisplayAlerts = False
    mintType = intType
    
    If mintType = 1 Then
        mstrIn = "C:\APPSOFT\�����ļ�\�շ�Ŀ¼"
    ElseIf mintType = 2 Then
        mstrIn = "C:\APPSOFT\�����ļ�\ҩƷĿ¼"
    Else
        mstrIn = "C:\APPSOFT\�����ļ�\����Ŀ¼"
    End If
    
    Me.Show 1, frmParent
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Me.cmdCheck.Left = Me.Width - 400 - Me.cmdCheck.Width
    Me.lblFileName.Width = Me.cmdCheck.Left - Me.cmdChoose.Width - 550
    Me.cmdChoose.Left = Me.cmdCheck.Left - 500
    Me.VSFList.Width = Me.cmdChoose.Width + Me.cmdChoose.Left - Me.lblFile.Left
    Me.VSFList.Height = Me.Height - Me.cmdCheck.Height - 900
    Me.cmdCancle.Left = Me.cmdCheck.Left
    Me.cmdCancle.Top = Me.VSFList.Top + Me.VSFList.Height - Me.cmdCancle.Height
    
    Me.cmdOutput.Left = Me.cmdCheck.Left
    Me.cmdOutput.Top = Me.cmdCancle.Top - Me.cmdOutput.Height - 100
    
    Me.cmdImport.Left = Me.cmdCheck.Left
    Me.cmdImport.Top = Me.cmdOutput.Top - Me.cmdImport.Height - 100
    
    Me.chkStop.Left = Me.cmdCheck.Left
    Me.chkStop.Top = Me.cmdImport.Top - Me.chkStop.Height
    
    Me.chk��Ӧ��.Left = Me.cmdCheck.Left
    Me.chk��Ӧ��.Top = Me.chkStop.Top - Me.chk��Ӧ��.Height
    
    Me.prg.Top = Me.Height - Me.prg.Height - 550
    Me.prg.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�ر�Excel����
    If mblnExists And Not mobjWB Is Nothing Then mobjXLS.ActiveWorkbook.SaveAs (Me.lblFileName.Caption)
    Set mobjWB = Nothing
    Set mobjWS = Nothing
    Set mobjWSType = Nothing
    mobjXLS.quit
    Set mobjXLS = Nothing
    Set mRsError = Nothing
    mblnExists = False
    mintType = 0
    mLngCount = 0
    mLngType = 0
    mLngSumType = 0
    mLngSumCount = 0
    
    Set mCollTypeCols = Nothing
    Set mCollItemCols = Nothing
    mstrIn = ""
End Sub

Private Sub AddErr(ByVal strType As String, ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, ByVal strPage As String)
    mRsError.AddNew
    mRsError!Type = strType
    mRsError!Row = lngRow
    mRsError!Col = lngCol
    mRsError!Error = strContent
    mRsError!Page = strPage
    mRsError.Update
End Sub

