VERSION 5.00
Begin VB.Form frmVerifyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����Ŀ�༭"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8505
   Icon            =   "frmVerifyEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   3105
      Left            =   45
      TabIndex        =   33
      Top             =   -45
      Width           =   7110
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   5655
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   165
         Width           =   1395
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   270
         Index           =   2
         Left            =   4545
         TabIndex        =   2
         Top             =   180
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   165
         Width           =   3390
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1620
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1125
         TabIndex        =   15
         Top             =   1635
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   3360
         TabIndex        =   23
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   1125
         TabIndex        =   21
         Top             =   2370
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1125
         TabIndex        =   19
         Top             =   2010
         Width           =   3390
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   3360
         TabIndex        =   12
         Top             =   1275
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1125
         TabIndex        =   10
         Top             =   1275
         Width           =   1155
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1125
         TabIndex        =   8
         Top             =   915
         Width           =   3390
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1125
         TabIndex        =   6
         Top             =   570
         Width           =   1155
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ҫִ�а���(&E)"
         Height          =   210
         Index           =   1
         Left            =   1140
         TabIndex        =   25
         Top             =   2805
         Width           =   1680
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺ(&2)"
         Height          =   210
         Index           =   3
         Left            =   5760
         TabIndex        =   27
         Top             =   1320
         Width           =   1005
      End
      Begin VB.CheckBox chk 
         Caption         =   "����(&1)"
         Height          =   210
         Index           =   2
         Left            =   5760
         TabIndex        =   26
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   0
         TabIndex        =   38
         Top             =   480
         Width           =   7110
      End
      Begin VB.Frame Frame2 
         Height          =   2685
         Left            =   5310
         TabIndex        =   34
         Top             =   420
         Width           =   30
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����Χ"
         Height          =   180
         Index           =   15
         Left            =   5775
         TabIndex        =   39
         Top             =   660
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5445
         Picture         =   "frmVerifyEdit.frx":000C
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&T)"
         Height          =   180
         Index           =   1
         Left            =   4995
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&D)"
         Height          =   180
         Index           =   12
         Left            =   465
         TabIndex        =   0
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lbl 
         Caption         =   "˵���Ƿ��������Ŀ��Ϊ���ﲡ�˻�סԺ���˵����ƴ�ʩӦ�ã����߲���ֱ��Ӧ���ڲ��ˡ�"
         Height          =   1155
         Index           =   26
         Left            =   5745
         TabIndex        =   36
         Top             =   1635
         Width           =   1275
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(ƴ��)"
         Height          =   180
         Index           =   10
         Left            =   2280
         TabIndex        =   22
         Top             =   2445
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(ƴ��)"
         Height          =   180
         Index           =   8
         Left            =   2280
         TabIndex        =   11
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(���)"
         Height          =   180
         Index           =   11
         Left            =   4545
         TabIndex        =   24
         Top             =   2445
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "(���)"
         Height          =   180
         Index           =   9
         Left            =   4545
         TabIndex        =   13
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "���㵥λ(&U)"
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   14
         Top             =   1710
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����Ա�(&X)"
         Height          =   180
         Index           =   6
         Left            =   2340
         TabIndex        =   16
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&F)"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   20
         Top             =   2430
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&A)"
         Height          =   180
         Index           =   4
         Left            =   465
         TabIndex        =   18
         Top             =   2070
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   3
         Left            =   465
         TabIndex        =   9
         Top             =   1365
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   2
         Left            =   465
         TabIndex        =   7
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&B)"
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   5
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7275
      TabIndex        =   30
      Top             =   75
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7275
      TabIndex        =   31
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7275
      TabIndex        =   32
      Top             =   1410
      Width           =   1100
   End
   Begin VB.Frame Frame4 
      Height          =   2715
      Left            =   45
      TabIndex        =   35
      Top             =   2985
      Width           =   7110
      Begin VB.PictureBox vsf 
         Height          =   2505
         Left            =   1440
         ScaleHeight     =   2445
         ScaleWidth      =   5550
         TabIndex        =   29
         Top             =   150
         Width           =   5610
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   105
         Picture         =   "frmVerifyEdit.frx":1D06
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�����Ŀ(&M)"
         Height          =   180
         Index           =   14
         Left            =   435
         TabIndex        =   28
         Top             =   270
         Width           =   990
      End
      Begin VB.Label lbl 
         Caption         =   "�ڿ�����Ŀ�ļ������뵥ʱ����ͬʱ�������õ����������Ŀ��"
         Height          =   1815
         Index           =   13
         Left            =   435
         TabIndex        =   37
         Top             =   600
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmVerifyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mlngUpKey As Long
Private mlngKey As Long
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSQL As String
            
Private Function ShowOpenList(Optional strText As String) As Byte
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    strLvw = "����,1200,0,1;������Ŀ,2700,0,0;Ӣ����д,900,0,0;�걾����,900,0,0"

    ShowOpenList = 2
    
    strTmp = Trim(vsf.TextMatrix(1, 4))
    For mlngLoop = 2 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" And mlngLoop <> vsf.Row Then
            strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
            If strTmp = "" Then
                ShowSimpleMsg "���õļ�����Ŀû�й�ͬ�ı걾���ͣ�"
                Exit Function
            End If
        End If
    Next
    
    strText = UCase(strText)
    
    strSQL = "SELECT C.ID,C.����,C.������ AS ������Ŀ,C.Ӣ���� AS Ӣ������,D.��д AS Ӣ����д,zlGetSample(C.ID) AS �걾����,D.���㹫ʽ " & _
                "FROM ������ĿĿ¼ A,���鱨����Ŀ B,����������Ŀ C,������Ŀ D " & _
                "WHERE A.ID=B.������ĿID AND  D.��Ŀ��� IN (1,3) AND NVL(A.�����Ŀ,0)=0 " & _
                    IIf(strTmp = "", "", "AND C.ID IN (SELECT ��Ŀid FROM ������Ŀ�ο� WHERE INSTR('," & strTmp & ",',','||�걾����||',')>0)") & _
                    "AND B.������ĿID=C.ID AND C.ID=D.������Ŀid AND A.���='C'"
                    
    strSQL = strSQL & " AND (UPPER(A.����) Like '%" & strText & "%' OR UPPER(D.��д) LIKE '%" & strText & "%' OR A.���� Like '%" & strText & "%' OR A.ID IN (SELECT ������Ŀid FROM ������Ŀ���� WHERE (���� Like '%" & strText & "%' OR UPPER(����) Like '%" & strText & "%')))"
        
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    
    If rs.BOF Then
        
        ShowOpenList = 0
        
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then
    
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "ѡ�����Ŀ��" & zlCommFun.Nvl(rs("������Ŀ").value) & "����ǰ�Ѿ�ѡ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("Ӣ������").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("Ӣ����д").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("�걾����").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("���㹫ʽ").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenList = 1
        Exit Function
    End If
    
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY + 30, 4200, 2100, Me.Name & "\������Ŀѡ��", "����±���ѡ��һ����Ŀ") Then
        
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "ѡ�����Ŀ��" & zlCommFun.Nvl(rs("������Ŀ").value) & "����ǰ�Ѿ�ѡ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("������Ŀ").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("Ӣ������").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("Ӣ����д").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("�걾����").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("���㹫ʽ").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenList = 1
        
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
            
            
Private Function ShowOpenTree() As Byte
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo errHand
    
    strLvw = "����,1200,0,1;����,2700,0,0;Ӣ����д,900,0,0;�걾����,900,0,0"

    ShowOpenTree = 2
    
    strTmp = Trim(vsf.TextMatrix(1, 4))
    For mlngLoop = 2 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" And mlngLoop <> vsf.Row Then
            strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
            If strTmp = "" Then
                ShowSimpleMsg "���õļ�����Ŀû�й�ͬ�ı걾���ͣ�"
                Exit Function
            End If
        End If
    Next
    
    
    strSQL = "select * " & _
             "from (Select DISTINCT ID,�ϼ�ID,0 as ĩ��,����,���� ,'' as Ӣ������,'' as Ӣ����д,'' AS �걾����,'' AS ���㹫ʽ, " & _
                                   "DECODE(�ϼ�ID, Null, ID * POWER(10, 20), �ϼ�ID * POWER(10, 20) + ID) As ���� " & _
                     "From ���Ʒ���Ŀ¼ " & _
                    "Where ���� = 5 " & _
                    "Start With ID IN (SELECT DISTINCT ����id FROM ������ĿĿ¼ WHERE ��� = 'C') " & _
                   "Connect by Prior �ϼ�ID = ID " & _
                   "Union All " & _
                     "Select C.ID,A.����id AS �ϼ�ID,1 as ĩ��, A.����,A.����,C.Ӣ���� AS Ӣ������,D.��д AS Ӣ����д,zlGetSample(C.ID) AS �걾����,D.���㹫ʽ, " & _
                            "1 AS ���� " & _
                       "FROM ������ĿĿ¼ A,���鱨����Ŀ B,����������Ŀ C,������Ŀ D " & _
                      "Where A.ID=B.������Ŀid AND B.������Ŀid=C.ID AND C.ID=D.������Ŀid AND D.��Ŀ��� IN (1,3) AND NVL(A.�����Ŀ,0)=0 AND A.��� = 'C' AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) " & _
                            IIf(strTmp = "", "", "AND C.ID IN (SELECT ��Ŀid FROM ������Ŀ�ο� WHERE INSTR('," & strTmp & ",',','||�걾����||',')>0)") & _
                   ") A " & _
            "ORDER BY A.ĩ��, A.����, A.����"
                        
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    
    If rs.BOF Then Exit Function
    
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectExplorer.ShowSelect(Me, _
                            rs, _
                            sglX, _
                            sglY, _
                            5400, _
                            2400, _
                            vsf.CellHeight, _
                            "������Ŀ����ѡ��", _
                            strLvw, _
                            "��ѡ��һ��������Ŀ") Then
                            
        If CheckHave(zlCommFun.Nvl(rs("ID").value)) Then
            MsgBox "ѡ�����Ŀ��" & zlCommFun.Nvl(rs("����").value) & "����ǰ�Ѿ�ѡ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        vsf.EditText = zlCommFun.Nvl(rs("����").value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.Nvl(rs("����").value)
        vsf.TextMatrix(vsf.Row, 2) = zlCommFun.Nvl(rs("Ӣ������").value)
        vsf.TextMatrix(vsf.Row, 3) = zlCommFun.Nvl(rs("Ӣ����д").value)
        vsf.TextMatrix(vsf.Row, 4) = zlCommFun.Nvl(rs("�걾����").value)
        vsf.TextMatrix(vsf.Row, 5) = zlCommFun.Nvl(rs("���㹫ʽ").value)
        vsf.RowData(vsf.Row) = zlCommFun.Nvl(rs("ID").value)
        
        ShowOpenTree = 1
        
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    For mlngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(mlngLoop)) = lngKey And vsf.Row <> mlngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.Nvl(rsData("ID")))
        
        On Error GoTo errHand
        For lngLoop = 0 To objMsf.Cols - 1
        
            On Error Resume Next
            strMask = ""
            strMask = MaskArray(lngLoop)
                                    
            On Error GoTo errHand
            If strMask <> "" Then
                objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
            Else
                objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop)))
            End If
                        
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngUpKey As Long, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    
    mblnStartUp = True
    mblnOK = False
    
    mlngUpKey = lngUpKey
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    If ReadData = False Then
        cmdOK.Tag = ""
        Exit Function
    End If
    
    If mlngKey = 0 Then
        
        If GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\������Ŀ����\", "����", 0) = 0 Then
        
            mstrSQL = "SELECT NVL(MAX(����),'0000000') AS ���� FROM ������ĿĿ¼ WHERE ��� >= 'A'"
            zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
            If mRs.BOF = False Then txt(0).Text = Right(String(10, "0") & Val(mRs("����")) + 1, Len(mRs("����")))
            
        Else
            strTmp = Mid(txt(8).Text, 2, InStr(1, txt(8).Text, "]") - 2)
            
            mstrSQL = "SELECT NVL(MAX(����),'0000000') AS ���� FROM ������ĿĿ¼ WHERE ��� >= 'A' and ���� like '" & strTmp & "%'"
            zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
            If mRs.BOF = False Then txt(0).Text = strTmp & Right(String(10, "0") & Val(mRs("����")) + 1, Len(mRs("����")) - Len(strTmp))
            
        End If
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '1.������볤��
    txt(1).MaxLength = GetMaxLength("������ĿĿ¼", "����")
    txt(0).MaxLength = GetMaxLength("������ĿĿ¼", "����")
    txt(2).MaxLength = GetMaxLength("������Ŀ����", "����")
    txt(3).MaxLength = GetMaxLength("������Ŀ����", "����")
    txt(7).MaxLength = GetMaxLength("������ĿĿ¼", "���㵥λ")
    txt(4).MaxLength = GetMaxLength("������Ŀ����", "����")
    txt(5).MaxLength = GetMaxLength("������Ŀ����", "����")
    txt(6).MaxLength = GetMaxLength("������Ŀ����", "����")
            
    '2.��������
    mstrSQL = "SELECT ����||'-'||����,0 FROM ���Ƽ�������"
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    Call AddComboData(cbo(0), rs)
    If cbo(0).ListCount > 0 Then cbo(0).ListIndex = 0
    
    '3.�����Ա�
    cbo(1).AddItem "1-����"
    cbo(1).ItemData(cbo(1).NewIndex) = 0
    
    cbo(1).AddItem "2-����"
    cbo(1).ItemData(cbo(1).NewIndex) = 1
    
    cbo(1).AddItem "3-Ů��"
    cbo(1).ItemData(cbo(1).NewIndex) = 2
    
    cbo(1).ListIndex = 0
    
    '2.��Ŀ����
    mstrSQL = "SELECT '['||����||']'||���� AS ���� FROM ���Ʒ���Ŀ¼ WHERE ID=" & mlngUpKey
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        txt(8).Text = zlCommFun.Nvl(rs("����").value)
        txt(8).Tag = mlngUpKey
    End If
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "������Ŀ", 3300, 1, "...", 1
        .NewColumn "Ӣ������", 1500, 1
        .NewColumn "Ӣ����д", 810, 1
        .NewColumn "�걾����", 900, 1
        .NewColumn "���㹫ʽ", 0, 1
        .FixedCols = 1
        .ColHidden(2) = True
    End With
        
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    mstrSQL = "SELECT A.*,B.���� AS �ϼ����� " & _
                "FROM ������ĿĿ¼ A,���Ʒ���Ŀ¼ B " & _
                "WHERE A.����id=B.ID(+) AND A.ID=" & mlngKey
                
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.Nvl(rs("����").value)
        txt(1).Text = zlCommFun.Nvl(rs("����").value)
        txt(7).Text = zlCommFun.Nvl(rs("���㵥λ").value)
        txt(8).Text = zlCommFun.Nvl(rs("�ϼ�����").value)
        txt(8).Tag = zlCommFun.Nvl(rs("����id").value)
        
        On Error Resume Next
        cbo(1).ListIndex = zlCommFun.Nvl(rs("�����Ա�").value, 0)
        On Error GoTo errHand
                
        chk(1).value = zlCommFun.Nvl(rs("ִ�а���").value, 0)
        
        zlControl.CboLocate cbo(0), zlCommFun.Nvl(rs("��������").value)
        
        Select Case zlCommFun.Nvl(rs("�������").value, 1)
        Case 1
            chk(2).value = 1
        Case 2
            chk(3).value = 1
        Case 3
            chk(2).value = 1
            chk(3).value = 1
        End Select
                
    End If
    
    mstrSQL = "SELECT ����,����,����,���� " & _
                "FROM ������Ŀ���� A " & _
                "WHERE A.���� IN (1,9) AND A.������Ŀid=" & mlngKey
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If rs("����") = 1 Then
                If rs("����").value = 1 Then
                    txt(2).Text = zlCommFun.Nvl(rs("����").value)
                Else
                    txt(3).Text = zlCommFun.Nvl(rs("����").value)
                End If
            Else
                txt(4).Text = zlCommFun.Nvl(rs("����").value)
                If rs("����").value = 1 Then
                    txt(5).Text = zlCommFun.Nvl(rs("����").value)
                Else
                    txt(6).Text = zlCommFun.Nvl(rs("����").value)
                End If
            End If
            rs.MoveNext
        Loop
    End If
                
    mstrSQL = "SELECT '' AS ���," & _
                      "A.������Ŀid AS ID," & _
                      "C.���� AS ������Ŀ," & _
                      "D.Ӣ���� AS Ӣ������," & _
                      "E.��д AS Ӣ����д,zlGetSample(D.ID) AS �걾����,E.���㹫ʽ " & _
                 "FROM ���鱨����Ŀ A," & _
                      "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = " & mlngKey & ") B," & _
                      "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> " & mlngKey & " AND " & _
                      "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid"
    
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF = False Then
        vsf.TextMatrix(0, 0) = "���"
        Call FillGrid(vsf, rs)
        vsf.TextMatrix(0, 0) = ""
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varTmp As Variant
    Dim lngLeftPos As Long
    Dim lngRightPos As Long
    Dim strTmpID As String
    Dim strID As String
    Dim rs As New ADODB.Recordset
                
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "��Ŀ���벻��Ϊ��ֵ��"
        LocationObj txt(0)
        Exit Function
    End If
    
    strTmp = CheckNumeric(txt(0).Text, txt(0).MaxLength, 0, 1)
    If strTmp <> "" Then
        ShowSimpleMsg "����" & strTmp
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        ShowSimpleMsg "��Ŀ���Ʋ���Ϊ��ֵ��"
        LocationObj txt(1)
        Exit Function
    End If
    
    '������㹫ʽ�Ƿ���ȷ
    For mlngLoop = 1 To vsf.Rows - 1
        strTmpID = strTmpID & "," & Trim(vsf.RowData(mlngLoop))
    Next
    strTmpID = strTmpID & ","
    
    For mlngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(mlngLoop, 5)) <> "" Then
            
            '�Ǽ�����Ŀ,�м��㹫ʽ
            strTmp = Trim(vsf.TextMatrix(mlngLoop, 5))
            
            lngLeftPos = InStr(strTmp, "[")
            lngRightPos = InStr(strTmp, "]")
            Do While lngLeftPos > 0 And lngRightPos > 0
                
                strID = Trim(Mid(strTmp, lngLeftPos + 1, lngRightPos - lngLeftPos - 1))
                If strID <> "" Then
                    '�ҵ�����Ŀ
                    If InStr(strTmpID, "," & strID & ",") = 0 Then
                        'û��������
                        
                        Call zlDatabase.OpenRecordset(rs, "SELECT ������ FROM ����������Ŀ WHERE ID=" & Val(strID), Me.Caption)
                        If rs.BOF = False Then
                            
                            ShowSimpleMsg "���㹫ʽ�С�" & zlCommFun.Nvl(rs("������")) & "����Ŀδ�������ڵ�ǰ��ϼ�����Ŀ�У�"
                            
                            vsf.Row = mlngLoop
                            vsf.Col = 1
                            vsf.ShowCell vsf.Row, vsf.Col
                                
                            Exit Function
                        End If
                        
                        
                        ShowSimpleMsg "��" & Trim(vsf.TextMatrix(mlngLoop, 1)) & "����Ŀ�ļ��㹫ʽ����"
                            
                        vsf.Row = mlngLoop
                        vsf.Col = 1
                        vsf.ShowCell vsf.Row, vsf.Col
                            
                        Exit Function
                        
                    End If
                End If
                
                strTmp = Mid(strTmp, lngRightPos + 1)
                lngLeftPos = InStr(strTmp, "[")
                lngRightPos = InStr(strTmp, "]")
            Loop
            
        End If
    Next
            
    '����걾�����Ƿ���ͬ
    
    If Val(vsf.RowData(1)) > 0 Then
        strTmp = vsf.TextMatrix(1, 4)
        
        For mlngLoop = 2 To vsf.Rows - 1
            If Val(vsf.RowData(mlngLoop)) > 0 And vsf.TextMatrix(mlngLoop, 4) <> "" Then
                
                strTmp = GetCommon(strTmp, Split(vsf.TextMatrix(mlngLoop, 4), ","))
                
                If strTmp = "" Then
                    
                    ShowSimpleMsg "���õļ�����Ŀû�й�ͬ�ı걾���ͣ�"
                    vsf.Row = mlngLoop
                    vsf.Col = 4
                    vsf.ShowCell vsf.Row, vsf.Col
                    
                    Exit Function
                End If
                
            End If
        Next
    End If
    
    
    ValidData = True
    
End Function

Private Function GetCommon(ByVal str��׼ As String, ByVal var��� As Variant) As String
                    
    Dim lngLoop As Long
        
    GetCommon = ""
    
    For lngLoop = 0 To UBound(var���)
        If InStr("," & str��׼ & ",", "," & var���(lngLoop) & ",") > 0 Then
            GetCommon = GetCommon & "," & var���(lngLoop)
        End If
    Next
    
    If GetCommon <> "" Then GetCommon = Mid(GetCommon, 2)
    
End Function


Private Function SaveData(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Dim blnTran As Boolean
    Dim strSQL(1 To 2) As String
    
    On Error GoTo errHand
    
    If mlngKey = 0 Then
        '����
        lngKey = zlDatabase.GetNextId("������ĿĿ¼")
        
        strSQL(1) = "ZL_������Ŀ_INSERT('C'," & Val(txt(8).Tag) & "," & _
                                        lngKey & ",'" & _
                                        txt(0).Text & "','" & _
                                        txt(1).Text & "','" & _
                                        txt(2).Text & "','" & _
                                        txt(3).Text & "','" & _
                                        txt(4).Text & "','" & _
                                        txt(5).Text & "','" & _
                                        txt(6).Text & "','" & _
                                        zlCommFun.GetNeedName(cbo(0).Text) & "',1,1,3,'" & _
                                        txt(7).Text & "'," & _
                                        cbo(1).ItemData(cbo(1).ListIndex) & "," & _
                                        chk(1).value & "," & _
                                        IIf(chk(2).value And chk(3).value, 3, IIf(chk(2).value, 1, IIf(chk(3).value, 2, 0))) & "," & _
                                        "1,NULL,NULL,1,NULL,NULL,'',NULL)"
    Else
        '�޸�
        lngKey = mlngKey

        strSQL(1) = "ZL_������Ŀ_UPDATE('C'," & Val(txt(8).Tag) & "," & _
                                        lngKey & ",'" & _
                                        txt(0).Text & "','" & _
                                        txt(1).Text & "','" & _
                                        txt(2).Text & "','" & _
                                        txt(3).Text & "','" & _
                                        txt(4).Text & "','" & _
                                        txt(5).Text & "','" & _
                                        txt(6).Text & "','" & _
                                        zlCommFun.GetNeedName(cbo(0).Text) & "',1,1,3,'" & _
                                        txt(7).Text & "'," & _
                                        cbo(1).ItemData(cbo(1).ListIndex) & "," & _
                                        chk(1).value & "," & _
                                        IIf(chk(2).value And chk(3).value, 3, IIf(chk(2).value, 1, IIf(chk(3).value, 2, 0))) & "," & _
                                        "1,NULL,NULL,1,NULL,NULL,'',NULL,1)"
    End If
    
    For mlngLoop = 1 To vsf.Rows - 1
        If vsf.RowData(mlngLoop) > 0 Then
            strValue = strValue & "|null^" & Val(vsf.RowData(mlngLoop))
        End If
    Next
    If strValue <> "" Then strValue = Mid(strValue, 2)
    strSQL(2) = "ZL_���鱨����Ŀ_UPDATE(" & lngKey & ",'" & strValue & "')"
    
    blnTran = True
    
    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    SaveData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Sub cbo_Click(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim objPoint As POINTAPI
    Dim rs As New ADODB.Recordset
    
    mstrSQL = "Select ID," & _
                    "�ϼ�ID," & _
                    "0 as ĩ��," & _
                    "'['||����||']'||���� AS ���� " & _
              "From ���Ʒ���Ŀ¼ where ����=5 " & _
              "Start With �ϼ�ID IS NULL " & _
                "Connect by Prior ID=�ϼ�ID "
                
    Call zlDatabase.OpenRecordset(rs, mstrSQL, Me.Caption)
    If rs.BOF Then Exit Sub
    
    Call ClientToScreen(txt(8).hwnd, objPoint)
    If frmSelectTree.ShowSelect(Me, rs, objPoint.x * 15 - 30, objPoint.y * 15 + txt(8).Height - 30, txt(8).Width, 3000, txt(8).Height, txt(8).Tag, "�������ѡ��", "��ѡ��һ������λ��") Then
        txt(8).Text = zlCommFun.Nvl(rs("����").value)
        txt(8).Tag = zlCommFun.Nvl(rs("ID").value)
    End If
    
    txt(8).SetFocus
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub CmdOK_Click()
    Dim lngKey As Long
    Dim strTmp As String
    
    If cmdOK.Tag <> "" Then
        
        If ValidData = False Then Exit Sub
        
        If SaveData(lngKey) = False Then Exit Sub
        
        mblnOK = True
        
        'ˢ���������������ʾ
        Call mfrmMain.EditRefresh(2, lngKey)
        
        If mlngKey = 0 Then
            
            '����ؼ�����
            txt(0).Text = ""
            txt(1).Text = ""
            txt(2).Text = ""
            txt(3).Text = ""
            txt(4).Text = ""
            txt(5).Text = ""
            txt(6).Text = ""
            txt(7).Text = ""
            
            '�����һ��Ŀ������Ŀ
            Call ClearGrid(vsf)
            
            '����ȱʡ����Ŀ����
            If GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\������Ŀ����\", "����", 0) = 0 Then
            
                mstrSQL = "SELECT NVL(MAX(����),'0000000') AS ���� FROM ������ĿĿ¼ WHERE ��� >= 'A'"
                zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
                If mRs.BOF = False Then txt(0).Text = Right(String(10, "0") & Val(mRs("����")) + 1, Len(mRs("����")))
                
            Else
                strTmp = Mid(txt(8).Text, 2, InStr(1, txt(8).Text, "]") - 2)
                
                mstrSQL = "SELECT NVL(MAX(����),'0000000') AS ���� FROM ������ĿĿ¼ WHERE ��� >= 'A' and ���� like '" & strTmp & "%'"
                zlDatabase.OpenRecordset mRs, mstrSQL, Me.Caption
                If mRs.BOF = False Then txt(0).Text = strTmp & Right(String(10, "0") & Val(mRs("����")) + 1, Len(mRs("����")) - Len(strTmp))
                
            End If
            
            '��λ��ѡ����Ŀ����
            LocationObj txt(0)
            
            cmdOK.Tag = ""
            Exit Sub
        End If
        
    End If
    
    cmdOK.Tag = ""
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 7
        Call zlCommFun.OpenIme(True)
    End Select
    
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
        If Index = 8 Then zlCommFun.PressKey vbKeyTab
    Else
        Select Case Index
        Case 2, 3, 5, 6
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 0
            KeyAscii = FilterKeyAscii(KeyAscii, 1)
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 7
        Call zlCommFun.OpenIme(False)
    End Select
    
    If Index = 1 Then
        If InStr(txt(Index).Text, "'") = 0 Then
            txt(2).Text = zlGetSymbol(txt(Index).Text, 0)
            txt(3).Text = zlGetSymbol(txt(Index).Text, 1)
        End If
    End If
    
    If Index = 4 Then
        If InStr(txt(Index).Text, "'") = 0 Then
            txt(5).Text = zlGetSymbol(txt(Index).Text, 0)
            txt(6).Text = zlGetSymbol(txt(Index).Text, 1)
        End If
    End If
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If Val(vsf.RowData(Row)) <= 0 Then
        Col = 1
        Cancel = True
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Select Case ShowOpenTree
    Case 0
        'û��ƥ�����Ŀ
        MsgBox "û���ҵ���ƥ��Ľ�����߱걾���Ͳ�һ�£�", vbInformation, gstrSysName
        
    Case 1
        'ѡȡ��һ����Ŀ
        cmdOK.Tag = "Changed"
    Case 2
        'ȡ���˱���ѡ��
        
    End Select
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                Cancel = True
                Exit Sub
            End If
                        
            Select Case ShowOpenList(vsf.EditText)
            Case 0
                'û��ƥ�����Ŀ
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "û���ҵ���ƥ��Ľ�����߱걾���Ͳ�һ�£�", vbInformation, gstrSysName
                
            Case 1
                'ѡȡ��һ����Ŀ
                cmdOK.Tag = "Changed"
            Case 2
                'ȡ���˱���ѡ��
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End Select
        End If
    Else
        cmdOK.Tag = "Changed"
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        If Col = 1 And Trim(vsf.TextMatrix(Row, Col)) = "" Then
            zlCommFun.PressKey vbKeyTab
            Cancel = True
            KeyAscii = 0
        End If
    End If
End Sub


