VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLisStationWrite2 
   BorderStyle     =   0  'None
   Caption         =   "ϸ��������д"
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8700
   Icon            =   "frmLisStationWrite2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ListView lvwSelect 
      Height          =   2685
      Left            =   6570
      TabIndex        =   4
      Top             =   105
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "�������"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "�������"
         Object.Width           =   2540
      EndProperty
   End
   Begin zl9LisWork.VsfGrid vsf 
      Height          =   1845
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   3254
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7095
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationWrite2.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationWrite2.frx":1C94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   2250
      Left            =   240
      TabIndex        =   5
      Top             =   1950
      Width           =   6270
      Begin VB.CheckBox chkLast 
         Caption         =   "�ϴν��"
         Height          =   180
         Left            =   3840
         TabIndex        =   10
         Top             =   210
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   45
         TabIndex        =   6
         Top             =   120
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   609
         ButtonWidth     =   3043
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ѡ�ÿ�����(&G)"
               Key             =   "ѡ�ÿ�����"
               Object.ToolTipText     =   "ѡ�ÿ�����"
               Object.Tag             =   "ѡ�ÿ�����(&G)"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ҩ������(&N)"
               Key             =   "��ҩ������"
               Object.ToolTipText     =   "��ҩ������"
               Object.Tag             =   "��ҩ������(&N)"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin zl9LisWork.VsfGrid vsfDetail 
         Height          =   1695
         Left            =   30
         TabIndex        =   1
         Top             =   480
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   2990
      End
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5625
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4586
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
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
   Begin VB.TextBox txt��� 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   6630
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4200
      Width           =   1875
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   510
      Locked          =   -1  'True
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   4605
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   450
      Locked          =   -1  'True
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5130
      Width           =   4665
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   5430
      Top             =   870
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lbl��� 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϣ"
      Height          =   360
      Left            =   6180
      TabIndex        =   12
      Top             =   4200
      Width           =   450
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "���鱸ע"
      Height          =   345
      Left            =   75
      TabIndex        =   8
      Top             =   5010
      Width           =   375
   End
   Begin VB.Label lblResult 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   345
      Left            =   75
      TabIndex        =   9
      Top             =   4530
      Width           =   375
   End
End
Attribute VB_Name = "frmLisStationWrite2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mlngKey As Long    '�걾ID
Private mDeviceID As Long
Private mstrType As String '��������
Private mblnEdit As Boolean '�Ƿ�����༭
Private mbytRedoNumber As Long '��������
Private mblSelectHistory As Boolean '�Ƿ�ѡ����ʷ
Private mlngHistoryID As Long       'ѡ�����ʷID

Private WithEvents mfrmRequest As frmLabRequest                     '���յǼǴ���
Attribute mfrmRequest.VB_VarHelpID = -1

Private mrsSave As New ADODB.Recordset
Private mblnChangeEdit As Boolean, mlngItemID As Long '΢������ĿID


Private Enum mCol
    ϸ������ = 1
    �������
    ��������
    ��ҩ����
    �ϴξ������
    ���������� = 1
    ҩ������
    ������
    �����־
    �ϴν��
    �ϴα�־
End Enum
Private mlng�����ط���id As Long

Public Event StartEdit(Cancel As Boolean)

Private Sub WriteRecord(ByVal lngRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    '1.��ɾ��ԭ���ļ�¼,�������Ѿ�ɾ��
    On Error GoTo ErrHand
    
    If Vsf.Rows > 0 And lngRow = 0 Then
        lngRow = 1
    End If
    
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(lngRow))
    
'    On Error Resume Next
    
    Call DeleteRecord(mrsSave)
    
    '2.��������ڵļ�¼
    For mlngLoop = 1 To vsfDetail.Rows - 1
        If Val(vsfDetail.RowData(mlngLoop)) > 0 Then
            mrsSave.AddNew
            mrsSave("Key").Value = Val(Vsf.RowData(lngRow))
            mrsSave("ID").Value = Val(vsfDetail.RowData(mlngLoop))
            mrsSave("Group").Value = mlng�����ط���id
            mrsSave("����������").Value = vsfDetail.TextMatrix(mlngLoop, mCol.����������)
            mrsSave("������").Value = vsfDetail.TextMatrix(mlngLoop, mCol.������)
            mrsSave("�����־").Value = vsfDetail.TextMatrix(mlngLoop, mCol.�����־)
            mrsSave("ҩ������").Value = vsfDetail.TextMatrix(mlngLoop, mCol.ҩ������)
        End If
    Next
    
ErrHand:
    
End Sub

Private Sub ReadRecord(ByVal lngRow As Long)
    '--------------------------------------------------------------------------------------------------------
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(lngRow))
    If mrsSave.RecordCount > 0 Then
        mrsSave.MoveFirst
        
        
        '1.����д�ȱ���Ŀ�����ϸĿ
        For mlngLoop = 1 To mrsSave.RecordCount
            
            vsfDetail.Rows = mlngLoop + 1
            
            vsfDetail.RowData(mlngLoop) = Val(mrsSave("ID").Value)
            mlng�����ط���id = mrsSave("Group").Value
            vsfDetail.TextMatrix(mlngLoop, 0) = mlngLoop
            vsfDetail.TextMatrix(mlngLoop, mCol.����������) = mrsSave("����������").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.������) = mrsSave("������").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.�����־) = mrsSave("�����־").Value
            vsfDetail.TextMatrix(mlngLoop, mCol.ҩ������) = mrsSave("ҩ������").Value

            Select Case UCase(Left(vsfDetail.TextMatrix(mlngLoop, mCol.�����־), 1))
            Case "R"
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.��ɫ
            Case "I"
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.��ɫ
            Case Else
                vsfDetail.Cell(flexcpForeColor, mlngLoop, 0, mlngLoop, vsfDetail.Cols - 2) = COLOR.��ɫ
            End Select

            mrsSave.MoveNext
        Next
        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
        'д���ϴν��
        If chkLast.Value = 1 Then Call LoadLastValue
        
    End If
End Sub

Private Sub LoadDefaultGroup(ByVal lngKey As Long)
    '--------------------------------------------------------------------------------------------------------
    '����:������ǰϸ��ȱʡ��Ӧ�Ŀ����ط���Ŀ�������Ŀ
    '����:lngKey            ϸ��id
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSql As String
    
    On Error GoTo ErrHand
        
    Call ClearGrid(vsfDetail)

    mstrSql = "SELECT '' AS ���,A.ID,A.������ AS ����������,'' AS ������," & _
            "Decode(Upper(E.Ĭ��ҩ��),'R','R-��ҩ','I','I-�н�','S','S-����',E.Ĭ��ҩ��) AS �������,B.�����ط���ID,'' As ҩ������,'' as �ϴν��,'' as �ϴ����� " & _
            "FROM �����ÿ����� A,���鿹������ҩ C,���鿹������ D,����ϸ�������� B,����ϸ�� E " & _
            "WHERE A.ID=C.������ID AND C.�����ط���ID=D.ID AND D.ID=B.�����ط���ID And B.ϸ��ID=E.ID AND E.ID= [1] Order By Decode(B.ȱʡ��־,1,1,0) Desc,A.����"
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        rs.filter = "�����ط���ID=" & rs("�����ط���ID")
        mlng�����ط���id = rs("�����ط���ID")
        vsfDetail.TextMatrix(0, 0) = "���"
        Call FillGrid(vsfDetail, rs)
        vsfDetail.TextMatrix(0, 0) = ""
        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function ShowOpenList(objVsf As Object, Optional strText As String, Optional blnWhere As Boolean = True, Optional ByVal bytMode As Byte = 1) As Byte
    '--------------------------------------------------------------------------------------------------------
    '����:���б�ṹ��ϸ��Ŀ¼
    '����:������2;�ɹ�����1;ȡ������0
    '--------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand
    
    ShowOpenList = 2
    
    Select Case bytMode
    Case 1
        strSQL = "SELECT A.ID,A.����,A.������,A.Ӣ����,A.����,B.�������� AS ���� " & _
                "FROM ����ϸ�� A,����ϸ������ B " & _
                "WHERE A.����ID=B.ID " & _
                    "AND (A.���� Like [3] OR A.������ Like [2] OR Upper(A.����) Like [3])"
    Case 2
        strSQL = "SELECT B.ID," & _
                           "NULL + 0 AS �ϼ�id," & _
                           "0 AS ĩ��," & _
                           "'' AS ����," & _
                           "'[' || B.���� || ']' || B.���� AS ����," & _
                           "'' AS Ӣ����," & _
                           "'' AS ���� " & _
                      "FROM ����ϸ�������� A, ���鿹������ B " & _
                     "WHERE A.�����ط���ID = B.ID And A.ϸ��ID = [1] AND (B.���� LIKE [2] OR B.���� LIKE [2] B.���� LIKE [2])"
                     
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Vsf.RowData(Vsf.Row)), "%" & strText & "%")
        
        If rs.BOF Then
            ShowOpenList = 0
            Exit Function
        End If
        
        If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
        strSQL = "SELECT B.ID," & _
                           "NULL + 0 AS �ϼ�id," & _
                           "0 AS ĩ��," & _
                           "'' AS ����," & _
                           "'[' || B.���� || ']' || B.���� AS ����," & _
                           "'' AS Ӣ����," & _
                           "'' AS ���� " & _
                      "FROM ����ϸ�������� A, ���鿹������ B " & _
                     "WHERE A.�����ط���ID = B.ID And A.ϸ��ID = [1] AND (B.���� LIKE [3] OR B.���� LIKE [2] B.���� LIKE [3] )" & _
                    "Union All " & _
                      "SELECT ROWNUM AS ID," & _
                             "B.�����ط���ID AS �ϼ�id," & _
                             "1 AS ĩ��," & _
                             "A.����," & _
                             "A.������ AS ����," & _
                             "A.Ӣ����," & _
                             "A.���� " & _
                        "FROM �����ÿ����� A, ���鿹������ҩ B " & _
                       "Where A.ID = B.������ID Order By A.����"
    Case 3
        strSQL = "SELECT ROWNUM AS ID,A.����,A.����,A.����,A.˵�� " & _
                "FROM ������������ A " & _
                "WHERE (A.���� Like [3] OR A.���� Like [3] )"
    Case 4
        strSQL = "select ID,����,������,Ӣ���� from �����ÿ����� where (���� like [3] or ������ like [3] or Ӣ���� like [3]) "
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Vsf.RowData(Vsf.Row)), "%" & strText & "%", "%" & UCase(strText) & "%")
    
    
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And blnWhere Then GoTo Over
        
    Call CalcPosition(sglX, sglY, objVsf)
    
    Select Case bytMode
    Case 1
        strLvw = "����,900,0,1;������,1800,0,0;Ӣ����,900,0,0;����,900,0,0;����,900,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 5400, 4500, Me.Name & "\����ϸ��ѡ��", "����±���ѡ��һ��ϸ����Ŀ") Then
            GoTo Over
        End If
    Case 3
        strLvw = "����,900,0,1;����,900,0,1;����,1800,0,0;˵��,1800,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 4500, 4500, Me.Name & "\������������ѡ��", "����±���ѡ��һ����������") Then
            GoTo Over
        End If
    Case 4
        strLvw = "����,900,0,1;������,1800,0,0;Ӣ����,1800,0,0"
        If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 4500, 4500, Me.Name & "\���鿹����ѡ��", "����±���ѡ��һ��������") Then
            GoTo Over
        End If
    End Select
        
    Exit Function
    
Over:

    Select Case bytMode
    Case 1
        If CheckHave(zlCommFun.Nvl(rs("ID").Value)) Then
            MsgBox "ѡ�����Ŀ��" & zlCommFun.Nvl(rs("������").Value) & "����ǰ�Ѿ�ѡ��", vbInformation, gstrSysName
            Exit Function
        End If
        objVsf.EditText = zlCommFun.Nvl(rs("������").Value)
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("������").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("������").Value)
        objVsf.RowData(objVsf.Row) = zlCommFun.Nvl(rs("ID").Value)
    Case 3
        objVsf.EditText = zlCommFun.Nvl(rs("˵��").Value)
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("˵��").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("˵��").Value)
    Case 4
        objVsf.RowData(objVsf.Row) = Nvl(rs("ID"))
'        objVsf.EditText = Nvl(rs("������").Value)
        objVsf.TextMatrix(objVsf.Row, mCol.����������) = Nvl(rs("������"))
    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ShowOpenTree(objVsf As Object, Optional ByVal bytMode As Byte = 1) As Byte
    '-----------------------------------------------------------------------------------------
    '����:������+�б�ṹ��������Ŀ����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim sglX As Single
    Dim sglY As Single
    
    On Error GoTo ErrHand
    
    ShowOpenTree = 2
    
    Select Case bytMode
    Case 1
        strSQL = "SELECT A.ID,NULL+0 AS �ϼ�id,0 AS ĩ��,A.����,'['||A.����||']'||A.�������� AS ����,'' AS Ӣ����,'' AS ����,'' AS ���� " & _
                "FROM ����ϸ������ A " & _
                "UNION ALL " & _
                "SELECT A.ID,A.����ID AS �ϼ�id,1 AS ĩ��,A.����,A.������ AS ����,A.Ӣ����,A.����,B.�������� AS ���� " & _
                "FROM ����ϸ�� A,����ϸ������ B " & _
                "WHERE A.����ID=B.ID "
    End Select
    
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rs.BOF Then
        ShowOpenTree = 0
        Exit Function
    End If

    Call CalcPosition(sglX, sglY, objVsf)
    Select Case bytMode
    Case 1
        
        strLvw = "����,1200,0,1;����,1800,0,0;Ӣ����,900,0,0;����,900,0,0"
        If frmSelectExplorer.ShowSelect(Me, rs, sglX, sglY, 5400, 2400, _
                                    objVsf.CellHeight, "��Ŀ����ѡ��", strLvw, "��ѡ��һ��������Ŀ") Then
                                    
            If CheckHave(zlCommFun.Nvl(rs("ID").Value)) Then
                MsgBox "ѡ�����Ŀ��" & zlCommFun.Nvl(rs("����").Value) & "����ǰ�Ѿ�ѡ��", vbInformation, gstrSysName
                Exit Function
            End If
            GoTo Over
        End If
    End Select
    
    Exit Function
    
Over:

    Select Case bytMode
    Case 1
        objVsf.Cell(flexcpData, objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("����").Value)
        objVsf.TextMatrix(objVsf.Row, objVsf.Col) = zlCommFun.Nvl(rs("����").Value)
        objVsf.RowData(objVsf.Row) = zlCommFun.Nvl(rs("ID").Value)
    End Select
    
    ShowOpenTree = 1
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim mlngLoop As Long
    
    For mlngLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(mlngLoop)) = lngKey And Vsf.Row <> mlngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function ReadData() As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset, mstrSql As String
    Dim strTmp As String
    
    On Error GoTo ErrHand
    
    Vsf.Rows = 2
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
    
    mstrSql = "SELECT A.������,A.������,A.����ʱ��,A.�����,A.���ʱ��,A.���鱸ע,A.��ע,a.������,a.����ʱ�� FROM ����걾��¼ A WHERE A.ID= [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If Not rs.EOF Then
        mbytRedoNumber = Nvl(rs("������"), 0)
        Me.txtComment = Nvl(rs("���鱸ע"))
        Me.txtResult = Nvl(rs("��ע"))
        
        With sbrInfo
            .Panels(1).Text = "�����ˣ�" & Nvl(rs("������"))
            .Panels(2).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
            If Nvl(rs("�����")) <> "" Then
                .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
            Else
                If Nvl(rs("������")) <> "" Then
                    .Panels(3).Text = "�����ˣ�" & Nvl(rs("������"))
                    .Panels(4).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                Else
                    .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                    .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                End If
            End If
        End With
    Else
        mstrSql = "SELECT A.������,A.������,A.����ʱ��,A.�����,A.���ʱ��,A.���鱸ע,A.��ע,a.������,a.����ʱ�� FROM ����걾��¼ A WHERE A.ID= [1] "
        Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
        If Not rs.EOF Then
            mbytRedoNumber = Nvl(rs("������"), 0)
            Me.txtComment = Nvl(rs("���鱸ע"))
            Me.txtResult = Nvl(rs("��ע"))
            
            With sbrInfo
                .Panels(1).Text = "�����ˣ�" & Nvl(rs("������"))
                .Panels(2).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                If Nvl(rs("�����")) <> "" Then
                    .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                    .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                Else
                    If Nvl(rs("������")) <> "" Then
                        .Panels(3).Text = "�����ˣ�" & Nvl(rs("������"))
                        .Panels(4).Text = "����ʱ�䣺" & IIf(IsNull(rs("����ʱ��")), "", Format(rs("����ʱ��"), "yyyy-MM-dd hh:mm"))
                    Else
                        .Panels(3).Text = "����ˣ�" & Nvl(rs("�����"))
                        .Panels(4).Text = "���ʱ�䣺" & IIf(IsNull(rs("���ʱ��")), "", Format(rs("���ʱ��"), "yyyy-MM-dd hh:mm"))
                    End If
                End If
            End With
        Else
            mbytRedoNumber = 0
            Me.txtComment = ""
            Me.txtResult = ""
            
            With sbrInfo
                .Panels(1).Text = "�����ˣ�"
                .Panels(2).Text = "����ʱ�䣺"
                .Panels(3).Text = "����ˣ�"
                .Panels(4).Text = "���ʱ�䣺"
            End With
        End If
    End If
    
    mstrSql = "SELECT C.������ĿID FROM ����걾��¼ A,����������Ŀ B,���鱨����Ŀ C " & _
                    "WHERE A.ID=B.�걾ID And B.������ĿID=C.������ĿID " & _
                        "AND A.ID= [1] "
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If rs.BOF = False Then
        mlngItemID = Nvl(rs("������ĿID"), 0)
    Else
        mlngItemID = 0
    End If
    
    mstrSql = "SELECT B.ID,D.������,B.������ AS ������Ŀ," & _
                    "A.������ AS ������,A.�������� as �������,'' as �ϴν��,a.��ҩ���� " & _
                    "FROM ������ͨ��� A,����ϸ�� B,����걾��¼ D " & _
                    "WHERE A.ϸ��id = B.ID " & _
                        "AND A.��¼���� = [1] " & _
                        "AND D.ID=A.����걾ID " & _
                        "AND D.ID= [2] Order by B.����"
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, mbytRedoNumber, IIf(mblSelectHistory, mlngHistoryID, mlngKey))
    If rs.BOF = False Then
        Vsf.TextMatrix(0, 0) = "#"
        Call FillGrid_UQ(Vsf, rs, Array("", "", "", ""))
        Vsf.TextMatrix(0, 0) = ""
        Vsf.Cell(flexcpBackColor, 1, 0, Vsf.Rows - 1, 0) = &HFDD6C6
    Else
        ResetVsf Vsf
        ResetVsf vsfDetail
    End If
    
    '1.��ɾ��ԭ���ļ�¼
    mrsSave.filter = ""
    Call DeleteRecord(mrsSave)
    
    mstrSql = "SELECT C.ϸ��ID AS Key,B.ID,B.������ AS ����������, A.��� AS ������,c.ҩ����ID, " & _
            "DECODE(A.�������,'R','R-��ҩ','I','I-�н�','S','S-����',A.�������) AS �������, " & _
            "DECODE(A.ҩ������,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As ҩ������ " & _
             "FROM ����ҩ����� A, �����ÿ����� B,������ͨ��� C " & _
            "Where A.������ID = B.ID And C.ID=A.ϸ�����ID AND C.��¼����=A.��¼���� AND C.����걾id= [1] AND C.��¼����= [2] Order By B.����"
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, IIf(mblSelectHistory, mlngHistoryID, mlngKey), mbytRedoNumber)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            mrsSave.AddNew
            mrsSave("Key").Value = zlCommFun.Nvl(rs("Key"), 0)
            mrsSave("Group").Value = zlCommFun.Nvl(rs("ҩ����ID"), 0)
            mrsSave("ID").Value = zlCommFun.Nvl(rs("ID"), 0)
            mrsSave("����������").Value = zlCommFun.Nvl(rs("����������"))
            mrsSave("������").Value = zlCommFun.Nvl(rs("������"))
            mrsSave("�����־").Value = zlCommFun.Nvl(rs("�������"))
            mrsSave("ҩ������").Value = zlCommFun.Nvl(rs("ҩ������"))
            
            rs.MoveNext
        Loop
    End If
    
    
    
    Call vsf_AfterRowColChange(0, 1, 1, 1)
    
    
    If mblSelectHistory = True Then mblnChangeEdit = True
    
    mblSelectHistory = False
    
    'д���ϴν��
    If chkLast.Value = 1 Then Call LoadLastValue
    
     'д�������Ϣ
    Me.txt���.Text = ""
    gstrSql = "Select b.ҽ��id, b.��Ŀ, b.����, b.����" & vbNewLine & _
                "From ����걾��¼ a, ����ҽ������ b" & vbNewLine & _
                "Where a.ҽ��id = b.ҽ��id and a.ID = [1] " & vbNewLine & _
                "Order By ҽ��id, ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
    
    Do Until rs.EOF
        strTmp = strTmp & Nvl(rs("��Ŀ")) & ":" & Replace(Nvl(rs("����")), vbCrLf, vbCrLf & "    ") & vbCrLf
        rs.MoveNext
    Loop
    Me.txt���.Text = strTmp
    
    ReadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlRefresh(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mlngKey = lngKey
    
'    SetEditState False
    mlngHistoryID = 0
    Call Form_Resize
    '��ʼ�����б�
    If ReadData = False Then Exit Function
    
    zlRefresh = True
End Function

Public Function ZlEditStart(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ��༭����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    SetEditState True
    '��ʼ�����б�
    If mlngKey <> lngKey Then
        mlngKey = lngKey
        If ReadData = False Then Exit Function
    End If
    mblnChangeEdit = False
    ZlEditStart = True
    With Vsf
        If Val(.RowData(.Row)) > 0 Then
            
        Else
            .Col = mCol.ϸ������
        End If
        .EditMode(.Col) = 1
        .SetFocus
        
        ShowValue .Col - 1
    End With
    With Vsf
        .EditMode(mCol.������) = 1
        .EditMode(mCol.ҩ������) = 1
        .EditMode(mCol.������) = 1
    End With
End Function

Public Function ZlSave() As Boolean
    '�ȱ���һ�µ�ǰϸ���Ŀ�����
    Call WriteRecord(Vsf.Row)
    
    If SaveData() = False Then Exit Function

    ZlSave = True
End Function

Public Function ZlCancel() As Boolean
    '��ʾ�Ƿ񱣴�
    SetEditState False
    
    ZlCancel = True
End Function

Public Function ZlClearForm() As Boolean
    With sbrInfo
        .Panels(1).Text = "�����ˣ�"
        .Panels(2).Text = "����ʱ�䣺"
        .Panels(3).Text = "����ˣ�"
        .Panels(4).Text = "���ʱ�䣺"
    End With

    Me.txtComment = ""
    Me.txtResult = ""
    ResetVsf Vsf
    ResetVsf vsfDetail
End Function

Private Sub SetEditState(ByVal blnEdit As Boolean)
    mblnEdit = blnEdit
'    Vsf.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    tbr.Enabled = blnEdit
'    vsfDetail.Body.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    
    txtComment.Locked = Not blnEdit
    txtResult.Locked = Not blnEdit
    Me.lvwSelect.Visible = blnEdit
    Call Form_Resize
End Sub

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim bytResultFlag As Byte
    Dim lngKey As Long
    Dim strSQL() As String
    Dim mlngLoop As Long
    Dim blnNoResult As Boolean
    Dim str�����־ As String
    Dim lngGroup As Long

    If Not mblnChangeEdit Then SaveData = True: Exit Function

    On Error GoTo ErrHand
    ReDim strSQL(1 To 1)

    '��ȡ����ʱ��
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")

    strSQL(ReDimArray(strSQL)) = "ZL_����ҩ�����_DELETE(" & mlngKey & "," & mbytRedoNumber & ")"

    blnNoResult = True
    For mlngLoop = 1 To Vsf.Rows - 1
        If Val(Vsf.RowData(mlngLoop)) > 0 Then
            blnNoResult = False
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(mlngLoop))
            If mrsSave.RecordCount > 0 Then
                lngGroup = mrsSave("Group").Value
            Else
                lngGroup = 0
            End If
            lngKey = zlDatabase.GetNextId("������ͨ���")
            strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_������д2(" & lngKey & "," & _
                mlngKey & ",NULL,'" & _
                Vsf.TextMatrix(mlngLoop, mCol.�������) & "'," & _
                "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
                mbytRedoNumber & ",NULL,1," & Val(Vsf.RowData(mlngLoop)) & ",NULL,'" & Vsf.TextMatrix(mlngLoop, mCol.��������) & "'," & _
                "NULL,NULL,'" & txtComment & "','" & txtResult & "',0,'" & Vsf.TextMatrix(mlngLoop, mCol.��ҩ����) & "','" & UserInfo.���� & "'," & _
                IIf(lngGroup = 0, "NULL", lngGroup) & ")"
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(mlngLoop))
            If mrsSave.RecordCount > 0 Then
                mrsSave.MoveFirst
                Do While Not mrsSave.EOF
                    If Len(Trim(mrsSave("������").Value)) > 0 Or Len(Trim(mrsSave("�����־").Value)) > 0 Then
                        If mrsSave("�����־").Value = "R-��ҩ" Or mrsSave("�����־").Value = "I-�н�" Or mrsSave("�����־").Value = "S-����" Then
                            str�����־ = Left(mrsSave("�����־").Value, 1)
                        Else
                            str�����־ = mrsSave("�����־").Value
                        End If
                        strSQL(ReDimArray(strSQL)) = "ZL_����ҩ�����_INSERT(" & lngKey & "," & mrsSave("ID").Value & _
                            ",'" & UserInfo.���� & "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),'" & _
                            mrsSave("������").Value & "'," & IIf(mrsSave("�����־").Value <> "", "'" & str�����־ & "'", "NULL") & "," & _
                            mbytRedoNumber & ",NULL," & IIf(IsNull(mrsSave("ҩ������")), "NULL", Val(Left(mrsSave("ҩ������"), 1))) & ")"
                    End If
                    mrsSave.MoveNext
                Loop
            End If
        End If
    Next
    If blnNoResult Then
        strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_������д2(" & 0 & "," & _
            mlngKey & ",NULL,''," & _
            "TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & _
            mbytRedoNumber & ",NULL,1,NULL,NULL,NULL,NULL,NULL,'" & txtComment & "','" & txtResult & "',0,Null,'" & UserInfo.���� & "')"
    End If

    blnTran = True

    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    If Signature(mlngKey, gstrDBUser, "����") = False Then
        Exit Function
    End If
    

    SaveData = True

    Exit Function
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    mblSelectHistory = True
    mlngHistoryID = Control.ID
    ReadData
    
    'ˢ�²�����Ϣ��ʾ����
    On Error Resume Next
'    mfrmRequest.zlRefresh m
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    If mlngKey = Control.ID Then
        Control.Caption = Control.Caption & "(��ǰ)"
    End If
    
    If mlngHistoryID = Control.ID Then
        Control.Checked = True
    End If
End Sub

Private Sub chkLast_Click()
    vsfDetail.Body.ColWidth(mCol.�ϴν��) = IIf(chkLast.Value, 1300, 0)
    vsfDetail.Body.ColWidth(mCol.�ϴα�־) = IIf(chkLast.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.�ϴξ������) = IIf(chkLast.Value, 1350, 0)
    If chkLast.Value Then LoadLastValue
    
End Sub

Private Sub Form_Load()
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    With Vsf
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "������Ŀ", 2400, 1, "...", 1
        .NewColumn "������", 1350, 1, , 1, 100
        .NewColumn "�������", 2000, 1, , 1, 100
        .NewColumn "��ҩ����", 2000, 1, , 1, 50
        .NewColumn "�ϴν��", 1350, 1, , 1, 100
        .FixedCols = 0
    End With
        
    With vsfDetail
        .Body.BackColor = &H80000005
        .Body.Appearance = flex3DLight
        .Body.BorderStyle = flexBorderFlat
        .Body.BackColorFixed = &HFDD6C6
        .Body.GridLinesFixed = flexGridFlat
        .Body.RowHeightMin = 300
        .Body.Editable = flexEDKbdMouse
        
        .Cols = 0
        .NewColumn "", 300, 7
        .NewColumn "����������", 2400, 1, "...", 1
        .NewColumn "ҩ������", 850, 1, " |1-MIC|2-DISK|3-K-B", 1
        .NewColumn "������", 1300, 1, , 1, 20
        .NewColumn "�������", 1000, 1, " |R-��ҩ|I-�н�|S-����|ESBL|BLAC|SDD|R*", 1
        .NewColumn "�ϴν��", 1300, 1, , 0, 20
        .NewColumn "�ϴ�����", 1000, 1, , 0, 20
        .FixedCols = 0
    End With
    
    Set mrsSave = New ADODB.Recordset
    With mrsSave
        
        .Fields.Append "Key", adVarChar, 18
        .Fields.Append "Group", adVarChar, 18
        .Fields.Append "ID", adVarChar, 18
        .Fields.Append "����������", adVarChar, 50
        .Fields.Append "������", adVarChar, 50
        .Fields.Append "�����־", adVarChar, 50
        .Fields.Append "ҩ������", adVarChar, 50
        .Open
        
    End With
    lvwSelect.Tag = 1 'Ĭ��ѡ��ϸ�����
    Set mfrmRequest = frmLabRequest                          '���յǼǴ���
    
    SetEditState False
    
    chkLast.Value = Val(zlDatabase.GetPara("frmLisStationWrite2_�鿴�ϴν��", 100, 1208, 0))
    vsfDetail.Body.ColWidth(mCol.�ϴν��) = IIf(chkLast.Value, 1300, 0)
    vsfDetail.Body.ColWidth(mCol.�ϴα�־) = IIf(chkLast.Value, 1000, 0)
    Vsf.Body.ColWidth(mCol.�ϴξ������) = IIf(chkLast.Value, 1350, 0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtComment.Visible = zlDatabase.GetPara("��ʾ���鱸ע", 100, 1208, True)
    lblComment.Visible = txtComment.Visible
    
    With txtComment
        .Left = Me.lblComment.Width
        .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        .Width = Me.ScaleWidth - .Left - Me.lbl���.Width - Me.txt���.Width
    End With
    With lblComment
        .Left = 0
        .Top = txtComment.Top + (Me.txtComment.Height - .Height) / 2
    End With
    
    With txtResult
        .Left = txtComment.Left
        If txtComment.Visible Then
            .Top = txtComment.Top - .Height - 30
        Else
            .Top = Me.ScaleHeight - Me.sbrInfo.Height - .Height - 30
        End If
        .Width = Me.ScaleWidth - .Left - Me.lbl���.Width - Me.txt���.Width
    End With
    With lblResult
        .Left = 0
        .Top = txtResult.Top + (Me.txtResult.Height - .Height) / 2
    End With
    
    With Me.lbl���
        .Top = Me.lblResult.Top
        .Left = Me.txtResult.Left + Me.txtResult.Width
    End With
    
    With Me.txt���
        .Top = Me.txtResult.Top
        .Left = Me.lbl���.Left + Me.lbl���.Width
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - Me.sbrInfo.Height - 30
    End With
    
    
    With lvwSelect
        .Left = Me.ScaleWidth - .Width - 30
        .Top = 0
        .Height = txtResult.Top - 30 - .Top
    End With
    
    With Vsf
        .Left = -15
        .Top = 0
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) - 30 - .Left
        .Height = txtResult / 2 - 30
    End With
    
    With fra
        .Left = -30
        .Top = Vsf.Top + Vsf.Height
        .Width = IIf(Me.lvwSelect.Visible, Me.lvwSelect.Left, Me.ScaleWidth) + 30 - .Left
        .Height = txtResult.Top - .Top - 30
    End With
    
    With vsfDetail
        .Left = 15
        .Width = fra.Width - 75 - .Left
        .Height = fra.Height - .Top - 45
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("frmLisStationWrite2_�鿴�ϴν��", Me.chkLast.Value, 100, 1208)
    mblnEdit = False
End Sub

Private Sub lvwSelect_DblClick()
    If lvwSelect.SelectedItem Is Nothing Then Exit Sub
    If Not mblnEdit Then Exit Sub
    
    Select Case Val(lvwSelect.Tag)
        Case 1 'ѡ����
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.�������) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 2 '��������
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.��������) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 3 'ϸ����ҩ����
            If Val(Vsf.RowData(Vsf.Row)) > 0 Then
                Vsf.TextMatrix(Vsf.Row, mCol.��ҩ����) = lvwSelect.SelectedItem.SubItems(1)
                Vsf.SetFocus
                mblnChangeEdit = True
            End If
        Case 4 'ѡ������
            Me.txtResult.SelText = lvwSelect.SelectedItem.SubItems(1)
            mblnChangeEdit = True
        Case 5 'ѡ��ע
            Me.txtComment.SelText = lvwSelect.SelectedItem.SubItems(1)
            mblnChangeEdit = True
    End Select
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim rs As New ADODB.Recordset, rsTmp  As New ADODB.Recordset, mstrSql As String
    Dim objPoint As POINTAPI
    
    Select Case Button.Key
        Case "ѡ�ÿ�����"
            mstrSql = "SELECT A.ID,A.����,A.����,C.Ĭ��ҩ�� " & _
                "FROM ���鿹������ A,����ϸ�������� B,����ϸ�� C " & _
                " WHERE A.ID=B.�����ط���ID AND B.ϸ��ID=" & Val(Vsf.RowData(Vsf.Row)) & _
                " and B.ϸ��ID = C.ID " & _
                " GROUP BY A.ID,A.����,A.����,C.Ĭ��ҩ��"
                
            Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption)
            If rs.BOF = False Then
                
                Call ClientToScreen(tbr.hWnd, objPoint)
                If frmSelectList.ShowSelect(Me, rs, "����,900,0,1;����,2400,0,0", objPoint.X * 15, objPoint.Y * 15 + tbr.Height, 3300, 2400, Me.Name & "\��������ѡ��", "����±���ѡ��һ������������Ŀ") Then
                    
                    Call ClearGrid(vsfDetail)
                    
'                    mstrSQL = "SELECT '' AS ���,A.ID,A.������ AS ����������,'' AS ������,'' AS �������,'' As ҩ������, " & _
                                "'' as �ϴν��,'' as �ϴ����� " & _
                                "FROM �����ÿ����� A,���鿹������ҩ C " & _
                                "WHERE A.ID=C.������ID AND C.�����ط���ID= [1] Order By A.����"
                    mstrSql = "SELECT '' AS ���,A.ID,A.������ AS ����������,'' AS ������," & _
                    " Decode(A.ҩ������, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') AS ҩ������," & _
                    " Decode('" & Nvl(rs("Ĭ��ҩ��")) & "', 'R', 'R-��ҩ', 'I', 'I-�н�', 'S', 'S-����', '" & Nvl(rs("Ĭ��ҩ��")) & "') As �������, " & _
                                "'' as �ϴν��,'' as �ϴ����� " & _
                                "FROM �����ÿ����� A,���鿹������ҩ C " & _
                                "WHERE A.ID=C.������ID AND C.�����ط���ID= [1] Order By A.����"
                    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, CLng(zlCommFun.Nvl(rs("ID"), 0)))
                    mlng�����ط���id = CLng(zlCommFun.Nvl(rs("ID"), 0))
    
                    If rsTmp.BOF = False Then
                        vsfDetail.TextMatrix(0, 0) = "���"
                        Call FillGrid(vsfDetail, rsTmp)
                        vsfDetail.TextMatrix(0, 0) = ""
                        vsfDetail.Cell(flexcpBackColor, 1, 0, vsfDetail.Rows - 1, 0) = &HFDD6C6
                    End If
                    
                    mblnChangeEdit = True
                End If
                gintSelectFocus = 5
                vsfDetail.SetFocus
            Else
                ShowSimpleMsg "û�п����ط������ݣ�"
            End If
        Case "��ҩ������"
            Call ClearGrid(vsfDetail)

            mblnChangeEdit = True
    End Select
End Sub

Private Sub txtComment_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtComment_GotFocus()
    With txtComment
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 5
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        mblnChangeEdit = True
    End If
End Sub

Private Sub txtResult_Change()
    mblnChangeEdit = True
End Sub

Private Sub txtResult_GotFocus()
    With txtResult
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
    If mblnEdit Then ShowValue 4
End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        mblnChangeEdit = True
    End If
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
    If mrsSave.RecordCount > 0 Then
        Call ReadRecord(Row)
    End If
    mblnChangeEdit = True
    '�����������
    RenumVsf Vsf, 0
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case mCol.�������
            If Left(Vsf.TextMatrix(Row, mCol.�������), 1) = "/" Then
                If LoadModel(Mid(Vsf.TextMatrix(Row, mCol.�������), 2)) Then
                    mblnChangeEdit = True
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error GoTo errH
    If OldRow <> NewRow Then
        '�б仯��ָ�
        Call ClearGrid(vsfDetail)
        
        mrsSave.filter = ""
        mrsSave.filter = "Key=" & Val(Vsf.RowData(NewRow))
        If mrsSave.RecordCount = 0 Then
'            Call LoadDefaultGroup(Val(vsf.RowData(NewRow)))
        Else
'            Call LoadDefaultGroup(Val(vsf.RowData(NewRow)))
            Call ReadRecord(NewRow)
        End If
    End If
    If OldCol <> NewCol And mblnEdit Then
        ShowValue NewCol - 1
    End If
    If Val(Vsf.RowData(NewRow)) = 0 And mblnEdit Then
        Vsf.Col = mCol.ϸ������
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then Cancel = True: Exit Sub
    mrsSave.filter = ""
    mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
    Call DeleteRecord(mrsSave)
    Call ClearGrid(vsfDetail)
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If Not mblnEdit Then Cancel = True: Exit Sub
    If Val(Vsf.RowData(Row)) = 0 Then
        Col = mCol.ϸ������
        Cancel = True
    End If
End Sub

Private Sub Vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow And OldRow < Vsf.Rows Then
        '�б仯ǰ����
        Call WriteRecord(OldRow)
    End If
    If OldCol <> NewCol And mblnEdit Then
        Vsf.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.�������, mCol.��������, mCol.ϸ������, mCol.��ҩ����
                Vsf.EditMode(NewCol) = 1
            Case Else
                Vsf.EditMode(mCol.�������) = 1
        End Select
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim blnUpdate As Boolean

    
    If Not mblnEdit Then Exit Sub
    
    On Error GoTo errH
    
    If Col = 1 Then
        bytResult = ShowOpenTree(Vsf, 1)
    Else
        bytResult = ShowOpenList(Vsf, "", False, 3)
    End If
    gintSelectFocus = 4: lvwSelect.SetFocus
    Vsf.SetFocus
    
    Select Case bytResult
    Case 0
        'û��ƥ�����Ŀ
        MsgBox "û���ҵ���ƥ��Ľ����", vbInformation, gstrSysName
    Case 1
        'ѡȡ��һ����Ŀ
        If Col = 1 Then
            If MsgBox("����ѡ�����µ�ϸ��,�Ƿ���Ҫ��յ�ǰҩ�����?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                blnUpdate = True
                Call ClearGrid(vsfDetail)
            End If
            
            mrsSave.filter = ""
            mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
            If mrsSave.RecordCount = 0 And blnUpdate = True Then
                Call LoadDefaultGroup(Val(Vsf.RowData(Row)))
            End If
            Vsf.Col = mCol.�������: gintSelectFocus = 4: Vsf.EditMode(mCol.ϸ������) = 0: Vsf.EditMode(mCol.�������) = 1
            
            mblnChangeEdit = True
        End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf_GotFocus()
'    If mblnEdit Then ShowValue 1
    If mblnEdit Then ShowValue Me.lvwSelect.Tag
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strSvrText As String, intRet As Integer
    
    If KeyCode = vbKeyReturn Then
        
        If InStr(Vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Cancel = True
            Exit Sub
        End If
            
        Select Case Col
            Case mCol.ϸ������
                intRet = ShowOpenList(Vsf, Vsf.EditText, True, 1)
                gintSelectFocus = 4: lvwSelect.SetFocus
                Vsf.SetFocus
                Select Case intRet
                    Case 0
                        'û��ƥ�����Ŀ
                        Cancel = True
                        Vsf.Cell(flexcpData, Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.EditText = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.TextMatrix(Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                            
                        MsgBox "û���ҵ���ƥ��Ľ����", vbInformation, gstrSysName
                    Case 1
                        'ѡȡ��һ����Ŀ
                        Call ClearGrid(vsfDetail)
                                        
                        mrsSave.filter = ""
                        mrsSave.filter = "Key=" & Val(Vsf.RowData(Row))
                        If mrsSave.RecordCount = 0 Then Call LoadDefaultGroup(Val(Vsf.RowData(Row)))
                        Vsf.Col = mCol.�������: gintSelectFocus = 4: Vsf.EditMode(mCol.ϸ������) = 0: Vsf.EditMode(mCol.�������) = 1
                        
                        mblnChangeEdit = True
                        Cancel = True
                    Case 2
                        'ȡ���˱���ѡ��
                        Cancel = True
                        Vsf.Cell(flexcpData, Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                        Vsf.TextMatrix(Row, Col) = Vsf.Cell(flexcpData, Row, Col)
                End Select
                
        End Select
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If Vsf.RowData(Row) = 0 And Row > 1 And KeyAscii = vbKeyReturn Then
        Vsf.Row = Row - 1: KeyAscii = 0
        vsfDetail.Col = mCol.�����־
        vsfDetail.SetFocus
    End If
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Exit Sub
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    mblnChangeEdit = True
End Sub

Private Sub Vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim rsTmp As New ADODB.Recordset
    Dim mFindMode As Integer        '������ID����������ʽ���� 0=����id 1=��������
    
    
    Debug.Print Button & " " & Now
    If Button <> vbRightButton Then Exit Sub
    
    On Error GoTo errH
    
    mFindMode = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
    
    gstrSql = "Select ����ʱ��, A.ID" & vbNewLine & _
              "From ����걾��¼ A, (Select ����id, ���� From ����걾��¼ Where ID = [1]) B" & vbNewLine & _
              "Where a.΢����걾 = 1 and  " & IIf(mFindMode = 0, " A.����id = B.����id ", " a.���� =b.���� ") & vbNewLine & _
              "Order By ����ʱ�� Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngKey)
    
    If rsTmp.RecordCount > 1 Then
        Set cbrPopupBar = Me.cbrthis.Add("�����˵�", xtpBarPopup)
        Do Until rsTmp.EOF
            Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, rsTmp("ID"), "����ʱ��:" & rsTmp("����ʱ��"))
            rsTmp.MoveNext
        Loop
        cbrPopupBar.ShowPopup
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then RaiseEvent StartEdit(Cancel)
End Sub

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngϸ��ID As Long, lng������ID As Long, lngRow As Long
    Dim strҩ������ As String, str������ As String
    Dim str�����־ As String
    Dim strSQL As String
                
    If Col = mCol.�����־ Then
        Select Case UCase(Left(vsfDetail.TextMatrix(Row, mCol.�����־), 1))
        Case "R"
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.��ɫ
        Case "I"
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.��ɫ
        Case Else
            vsfDetail.Cell(flexcpForeColor, Row, 0, Row, vsfDetail.Cols - 2) = COLOR.��ɫ
        End Select
    ElseIf Col = mCol.ҩ������ Then
        With vsfDetail
            For lngRow = .FixedRows To .Rows - 1
                If lngRow <> Row Then
                    If .TextMatrix(Row, Col) <> .TextMatrix(lngRow, Col) Then
                        .TextMatrix(lngRow, Col) = .TextMatrix(Row, Col)
                    End If
                End If
            Next
        End With
    ElseIf Col = mCol.������ Then
        With vsfDetail
            
            lngϸ��ID = Val(Vsf.RowData(Vsf.Row))
            lng������ID = Val(.RowData(.Row))
            strҩ������ = .TextMatrix(.Row, mCol.ҩ������)
            str������ = .TextMatrix(.Row, mCol.������)
            str�����־ = .TextMatrix(.Row, mCol.�����־)
            
            .TextMatrix(Row, mCol.�����־) = Eval�����־(lngϸ��ID, mlng�����ط���id, lng������ID, strҩ������, str������)
            
            Select Case UCase(Left(.TextMatrix(.Row, mCol.�����־), 1))
            Case "R"
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.��ɫ
            Case "I"
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.��ɫ
            Case Else
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 2) = COLOR.��ɫ
            End Select
            
        End With
    ElseIf Col = mCol.���������� Then
        With vsfDetail
            ShowOpenList vsfDetail, IIf(.EditText <> "", .EditText, .TextMatrix(.Row, mCol.����������)), True, 4
            
            WriteRecord Me.Vsf.Row
            Me.lvwSelect.SetFocus
            Me.vsfDetail.SetFocus
            gintSelectFocus = 5
        End With
    End If

    mblnChangeEdit = True
End Sub
Private Function Eval�����־(ByVal lngϸ��ID As Long, ByVal lng�����ط���ID As Long, ByVal lng������ID As Long, ByVal strҩ������ As String, ByVal str������ As String) As String
    '��������־
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim dblTmp  As Double, varTmp As Variant, intLoop As Integer
    Dim str��� As String
    
    If Val(strҩ������) = 0 Then
        Eval�����־ = ""
        Exit Function
    End If
    
    If Trim(str������) = "" Then
        Eval�����־ = ""
        Exit Function
    End If
    
    If InStr(Trim(str������), ">") > 0 Or InStr(str������, "��") > 0 Then
        Eval�����־ = "R-��ҩ"
        Exit Function
    End If
    
    If InStr(Trim(str������), "<") > 0 Or InStr(str������, "��") > 0 Then
        Eval�����־ = "S-����"
        Exit Function
    End If
    
    strSQL = "Select �жϷ�ʽ, �ο���ֵ, �ο���ֵ, ��ֵ���,��ֵ���,�м���" & vbNewLine & _
            "From ����ϸ�������زο�" & vbNewLine & _
            "Where ϸ��id = [1] And �����ط���id = [2] And ������id = [3] And ҩ������ = [4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngϸ��ID, lng�����ط���ID, lng������ID, Val(strҩ������))
    Do Until rsTmp.EOF
        
        If InStr(str������, "/") > 0 Then
            varTmp = Split(str������, "/")
            For intLoop = LBound(varTmp) To UBound(varTmp)
                If dblTmp < Val(varTmp(intLoop)) Then
                    dblTmp = Val(varTmp(intLoop))
                End If
            Next
        Else
            dblTmp = Val(str������)
        End If
        
        If rsTmp!�жϷ�ʽ = 1 Then
            If dblTmp >= Val("" & rsTmp!�ο���ֵ) And dblTmp <= Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!�м���)
                If str��� = "" Then str��� = "I-�н�"
                Eval�����־ = str���
            ElseIf dblTmp > Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!��ֵ���)
                If str��� = "" Then str��� = "R-��ҩ"
                Eval�����־ = str���
            ElseIf dblTmp < Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!��ֵ���)
                If str��� = "" Then str��� = "S-����"
                Eval�����־ = str���
            End If
        Else
            If dblTmp > Val("" & rsTmp!�ο���ֵ) And dblTmp < Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!�м���)
                If str��� = "" Then str��� = "I-�н�"
                Eval�����־ = str���
            ElseIf dblTmp >= Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!��ֵ���)
                If str��� = "" Then str��� = "R-��ҩ"
                Eval�����־ = str���
            ElseIf dblTmp <= Val("" & rsTmp!�ο���ֵ) Then
                str��� = Trim("" & rsTmp!��ֵ���)
                If str��� = "" Then str��� = "S-����"
                Eval�����־ = str���
            End If
        End If
        
        rsTmp.MoveNext
    Loop
End Function

Private Sub vsfDetail_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnEdit = True Then
        WriteRecord Me.Vsf.Row
    Else
        Cancel = True
    End If
    
End Sub

Private Sub vsfDetail_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Col = mCol.����������
    Cancel = True
End Sub

Private Sub vsfDetail_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldCol <> NewCol And mblnEdit Then
'        vsfDetail.EditMode(OldCol) = 0
        Select Case NewCol
            Case mCol.ҩ������, mCol.������, mCol.�����־, mCol.����������
                vsfDetail.EditMode(NewCol) = 1
            Case Else
                vsfDetail.EditMode(mCol.�����־) = 1
        End Select
    End If
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfDetail
        ShowOpenList vsfDetail, .TextMatrix(Row, mCol.����������), True, 4
        WriteRecord Me.Vsf.Row
    End With
    gintSelectFocus = 5
    
    Me.lvwSelect.SetFocus
    Me.vsfDetail.SetFocus
End Sub

Private Sub vsfDetail_GotFocus()
'    lvwSelect.ListItems.Clear
End Sub

Private Sub vsfDetail_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SelectNextRow(Row, Col)
        Cancel = True  '��ֹ�Զ���ؼ���KeyPress����
    End If
End Sub
Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call SelectNextRow(Row, Col)
        Exit Sub
    End If
    If Chr(KeyAscii) = "'" Then KeyAscii = 0

    mblnChangeEdit = True
End Sub

Private Sub vsfDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not mblnEdit Then
        RaiseEvent StartEdit(Cancel)
    End If
End Sub

Private Sub SelectNextRow(ByVal Row As Long, ByVal Col As Long)
    '��ת����һ��Ԫ��
   
    With vsfDetail
        If .Col + 1 <= mCol.�����־ Then
            .Body.Select .Row, .Col + 1
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Body.Select .Row + 1, mCol.����������
            Else
                If Trim(.RowData(.Row)) <> "" Then
                    .Rows = .Rows + 1
                    .Body.Select .Row + 1, mCol.����������
                End If
            End If
        End If
    End With
End Sub

Private Sub ShowValue(ByVal intType As Integer)
    'intType��1����������2������������3����ҩ���ơ�4-���5-��ע
    Dim rs As ADODB.Recordset
    Dim strSQL As String, strValue As String, aValues() As String, i As Long
    Dim ListItem As ListItem
    
    On Error GoTo errH
    
    Select Case intType
        Case 1
            strSQL = "SELECT ROWNUM AS ID,����,����,���� As ȡֵ FROM ���������� A " & _
                " WHERE ����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "΢����")
        Case 2
            strSQL = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ȡֵ FROM ������������ A"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Case 3
            strSQL = "select Rownum as ID,A.����,A.����,A.����,A.���� as ȡֵ from ϸ����ҩ���� A"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Case 4
            strSQL = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ȡֵ FROM ������������ A " & _
                "WHERE A.���� Is Null Or A.����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "΢����")
        Case 5
            strSQL = "SELECT Rownum As ID,A.����,A.����,A.����,A.˵�� As ȡֵ FROM ���鱸ע���� A " & _
                "WHERE A.���� Is Null Or A.����=[1]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "΢����")
    End Select
    
    With lvwSelect
        .ListItems.Clear
        .Tag = intType
        If Not rs Is Nothing Then
            Do While Not rs.EOF
                Set ListItem = .ListItems.Add(, "_" & rs("ID"), Nvl(rs("����")))
                ListItem.SubItems(1) = Nvl(rs("ȡֵ"))
                rs.MoveNext
            Loop
        End If
    End With
    
    'ȡ΢�����ȡֵ����
    If intType = 1 Then
        strSQL = "SELECT ȡֵ���� FROM ������Ŀ WHERE ������ĿID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngItemID)
        If rs.EOF Then
            strValue = "-|��|+|++|+++|++++"
        Else
            strValue = Nvl(rs("ȡֵ����"), "-|��|+|++|+++|++++")
            strValue = Replace(strValue, ";", "|")
        End If
        aValues = Split(strValue, "|")
        With lvwSelect
            For i = 0 To UBound(aValues)
                Set ListItem = .ListItems.Add(, "V" & i, aValues(i))
                ListItem.SubItems(1) = aValues(i)
            Next
        End With
    End If
    Me.lvwSelect.ColumnHeaders(1).Width = Me.lvwSelect.Width
    Me.lvwSelect.ColumnHeaders(2).Width = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadLastValue()
    '���� д������һ�ε�΢���������
    Dim rsTmp As New ADODB.Recordset
    Dim lngloop As Long
    Dim intFindPatientType As Integer   '0=��ID���� 1=����������
    Dim strPatientName As String        '��������
    
    On Error GoTo errH
    
    intFindPatientType = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0)
    
    If intFindPatientType <> 0 Then
        gstrSql = "select ���� from ����걾��¼ where id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
        If rsTmp.EOF = False Then strPatientName = Nvl(rsTmp("����"))
    End If
    
    gstrSql = "Select B.ϸ��id As Key,b.������ as ϸ�����, C.������id As ID,C.��� As ������," & vbNewLine & _
                "       Decode(C.�������, 'R', 'R-��ҩ', 'I', 'I-�н�', 'S', 'S-����',C.�������) As �������," & vbNewLine & _
                "       Decode(C.ҩ������, 1, '1-MIC', 2, '2-DISK', 3, '3-K-B', '') As ҩ������" & vbNewLine & _
                "From (Select ID from (Select b.Id  From ����걾��¼ a , ����걾��¼ b" & vbNewLine & _
                "Where " & IIf(intFindPatientType = 0, " b.����ID = a.����ID ", _
                                " b.����ID in (select ����ID from ������Ϣ where ���� = [2] )") & vbNewLine & _
                "And b.Id < [1] And a.Id = [1] " & vbNewLine & _
                "Order By b.Id Desc)   ) a ," & vbNewLine & _
                " ������ͨ�����B, ����ҩ����� C" & vbNewLine & _
                "Where A.ID = B.����걾id And B.ID = C.ϸ�����id(+) and b.ϸ��id is not null order by a.id desc  "

                    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey, strPatientName)
    
    If rsTmp.EOF Or Me.vsfDetail.Rows = 1 Then Exit Sub       'û�м�¼ʱ�˳�
    
    With vsfDetail
        For lngloop = 1 To .Rows - 1
            If .RowData(lngloop) <> "" Then
                rsTmp.filter = "ID=" & .RowData(lngloop)
                If rsTmp.EOF = False Then
                    .TextMatrix(lngloop, mCol.�ϴν��) = Nvl(rsTmp("������"))

                    .TextMatrix(lngloop, mCol.�ϴα�־) = Nvl(rsTmp("�������"))
                    Select Case UCase(Left(.TextMatrix(lngloop, mCol.�ϴα�־), 1))
                    Case "R"
                        .Cell(flexcpForeColor, lngloop, mCol.�ϴν��, lngloop, .Cols - 1) = COLOR.��ɫ
                    Case "I"
                        .Cell(flexcpForeColor, lngloop, mCol.�ϴν��, lngloop, .Cols - 1) = COLOR.��ɫ
                    Case Else
                        .Cell(flexcpForeColor, lngloop, mCol.�ϴν��, lngloop, .Cols - 1) = COLOR.��ɫ
                    End Select

                End If
            End If
        Next
    End With
    
    With Vsf
        For lngloop = 1 To .Rows - 1
            If .RowData(lngloop) <> "" Then
                rsTmp.filter = "Key=" & .RowData(lngloop)
                If rsTmp.EOF = False Then
                    .TextMatrix(lngloop, mCol.�ϴξ������) = Nvl(rsTmp("ϸ�����"))
                End If
            End If
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Resize()
    '�����������
    Call Form_Resize
End Sub
Private Function LoadModel(ByVal strCode As String) As Boolean
'���뱨��ģ��(΢������Ŀ)
'strCode��ģ���������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngCurrRow As Long
    Dim intColCount As Integer
    Dim intCol As Integer
    
    On Error GoTo errH
    
    LoadModel = False
    strSQL = "Select B.ID, B.������ As ������Ŀ, A.������ As ���ν��, A.��������" & vbNewLine & _
            "From ����ģ������ A, ����ϸ�� B, ����ģ��Ŀ¼ D" & vbNewLine & _
            "Where A.ϸ��id = B.ID And D.ID = A.ģ��id And A.ϸ��id Is Not Null And (D.���� = [1] Or D.���� = [1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode, mDeviceID)
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            lngCurrRow = FindRepeatLine(Vsf, CStr(zlCommFun.Nvl(rsTmp("ID"))))
            If lngCurrRow > 0 Then
                If Val(Vsf.RowData(lngCurrRow)) = Nvl(rsTmp("ID")) Then
                    Vsf.TextMatrix(lngCurrRow, mCol.�������) = Nvl(rsTmp("���ν��"))
                    Vsf.TextMatrix(lngCurrRow, mCol.��������) = Nvl(rsTmp("��������"))
                End If
            End If
            rsTmp.MoveNext
        Loop
        LoadModel = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function FindRepeatLine(ByRef objMsf As Object, ByVal strSeekID As String) As Long
    '-------------------------------------------------------------------------------------------------------------
    '����:����RowData����strSeekID����
    '����:
    '����:�кŻ�-1
    '-------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim intColCount As Integer, intCol As Integer
    FindRepeatLine = -1


    For i = 1 To objMsf.Rows - 1
        If Val(Me.Vsf.RowData(i)) = strSeekID Then
            FindRepeatLine = i
            Exit For
        End If

    Next

    If i <= objMsf.Rows - 1 Then FindRepeatLine = i
End Function

