VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmLabAnalyseData 
   Caption         =   "���������ռ�"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "frmLabAnalyseData.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11760
   StartUpPosition =   1  '����������
   Begin XtremeReportControl.ReportControl rptSource 
      Height          =   2865
      Left            =   60
      TabIndex        =   19
      Top             =   1530
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl rptAnalyse 
      Height          =   2865
      Left            =   6660
      TabIndex        =   20
      Top             =   1530
      Width           =   3975
      _Version        =   589884
      _ExtentX        =   7011
      _ExtentY        =   5054
      _StockProps     =   0
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      MultipleSelection=   0   'False
      ShowItemsInGroups=   -1  'True
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton cmdRightAll 
      Caption         =   ">>>"
      Height          =   435
      Left            =   5970
      TabIndex        =   18
      Top             =   4530
      Width           =   525
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "==>"
      Height          =   435
      Left            =   5970
      TabIndex        =   17
      Top             =   3570
      Width           =   525
   End
   Begin VB.CommandButton cmdLeftAll 
      Caption         =   "<<<"
      Height          =   435
      Left            =   5970
      TabIndex        =   16
      Top             =   2790
      Width           =   525
   End
   Begin VB.CommandButton CmdLeft 
      Caption         =   "<=="
      Height          =   435
      Left            =   5970
      TabIndex        =   15
      Top             =   1980
      Width           =   525
   End
   Begin VB.Frame fraFilter 
      Caption         =   "��ѯ����"
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12135
      Begin VB.CommandButton cmdFind 
         Caption         =   "��ѯ"
         Height          =   345
         Left            =   9270
         TabIndex        =   12
         Top             =   210
         Width           =   1065
      End
      Begin VB.ComboBox cbo����Ŀ�� 
         Height          =   300
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   630
         Width           =   3285
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   630
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Top             =   255
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95748097
         CurrentDate     =   39769
      End
      Begin VB.TextBox txt�걾�� 
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Top             =   240
         Width           =   1785
      End
      Begin VB.TextBox txt���� 
         Height          =   315
         Left            =   930
         TabIndex        =   3
         Top             =   240
         Width           =   1245
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   7440
         TabIndex        =   7
         Top             =   255
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   95748097
         CurrentDate     =   39769
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����Ŀ��"
         Height          =   180
         Left            =   4800
         TabIndex        =   10
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "��    ��"
         Height          =   180
         Left            =   150
         TabIndex        =   8
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��                  ---"
         Height          =   180
         Left            =   4800
         TabIndex        =   5
         Top             =   300
         Width           =   2610
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�걾����"
         Height          =   180
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�걾��"
         Height          =   180
         Left            =   2250
         TabIndex        =   1
         Top             =   300
         Width           =   540
      End
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCaptAnalyse 
      Height          =   315
      Left            =   6690
      TabIndex        =   14
      Top             =   1200
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   6
      Caption         =   "��������"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCapSource 
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   1200
      Width           =   2895
      _Version        =   589884
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   6
      Caption         =   "ԭʼ����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
   End
End
Attribute VB_Name = "frmLabAnalyseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum mCol
    �걾ID
    �걾��
    ����
    �Ա�
    ����
    ����ʱ��
    ��;
End Enum
Private mlngMachine As Long



Private Sub cmdFind_Click()
    Dim astrItem() As String
    Dim lngLoop As Long
    Dim varItem() As String
    Dim strBegingNO As String, strEndNO As String
    Dim strWhere  As String
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim strNumber As String
    
    If DateDiff("d", Me.DTPStart, Me.DTPEnd) > 30 Then
        If MsgBox("����ѡ���ʱ��δ���30�죬���ܵ��²�ѯ���ݹ�����ѯʱ�������" & vbCrLf & "�Ƿ������", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Me.DTPStart.SetFocus
            Exit Sub
        End If
    End If
    
    If Me.cbo����Ŀ��.ListCount = 0 Then
        MsgBox "��û�����÷���Ŀ�ģ��뵽�ֵ�����������ӷ���Ŀ��!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    '==========================================================����ԭʼ����=======================================================================
    gstrSql = " Select a.ID,a.�걾���, a.����, a.����, a.�Ա�, a.����ʱ��, " & vbNewLine & _
                "  Decode(a.����id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                "                 a.�걾���) As �걾����ʾ " & vbNewLine & _
                " From ����걾��¼ a , ���������¼ b " & vbNewLine & _
                " Where a.id = b.�걾ID(+) And (����id = [1] Or Nvl(����id, -1) = [1]) " & vbNewLine & _
                " And ����ʱ�� between [2] and [3] and b.��; is null "
    
    
    '�걾��
    If Trim(txt�걾��) <> "" Then
        txt�걾�� = Replace(Replace(txt�걾��, "��", "~"), "-", "~")
        varItem = Split(Trim(txt�걾��.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt����) <> "" Then
        strWhere = strWhere & " and to_Number(�걾���) between [4] and [5] "
        strBegingNO = TransSampleNO(Val(Me.txt����) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt����) & "-9999")
    End If
    gstrSql = gstrSql & strWhere
    
    Me.rptSource.Records.DeleteAll
    strTmp = Me.cbo����Ŀ��.List(Me.cbo����Ŀ��.ListIndex)
    strTmp = Mid(strTmp, 1, InStr(strTmp, "-") - 1)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo����.ItemData(Me.cbo����.ListIndex)), _
                    CDate(Format(Me.DTPStart, "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:59:59")), Val(strBegingNO), Val(strEndNO))
                        
    Do Until rsTmp.EOF
        
        Set Record = Me.rptSource.Records.Add
            For intLoop = 0 To Me.rptSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.�걾ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mCol.�걾��).Value = Nvl(rsTmp("�걾���"))
            Record.Item(mCol.�걾��).Caption = Nvl(rsTmp("�걾����ʾ"))
            Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
            Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
        rsTmp.MoveNext
    Loop
    Me.rptSource.Populate
    '==========================================================================================================================================
    
    '===========================================================��ѯ��������===================================================================
    gstrSql = " Select a.ID,a.�걾���, a.����, a.����, a.�Ա�, a.����ʱ��,c.����, " & vbNewLine & _
                "  Decode(a.����id, Null," & vbNewLine & _
                "                 To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000')," & vbNewLine & _
                "                 a.�걾���) As �걾����ʾ " & vbNewLine & _
                " From ����걾��¼ a , ���������¼ b, ���������; c " & vbNewLine & _
                " Where a.id = b.�걾ID And b.��;=c.���� and (����id = [1] Or Nvl(����id, -1) = [1]) " & vbNewLine & _
                " And ����ʱ�� between [2] and [3] and b.��; is not null and c.���� = [6] "
    
    
    '�걾��
    If Trim(txt�걾��) <> "" Then
        txt�걾�� = Replace(Replace(txt�걾��, "��", "~"), "-", "~")
        varItem = Split(Trim(txt�걾��.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt����) <> 0, Val(Me.txt����) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(�걾���) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt����) <> "" Then
        strWhere = strWhere & " and to_Number(�걾���) between [4] and [5] "
        strBegingNO = TransSampleNO(Val(Me.txt����) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt����) & "-9999")
    End If
    gstrSql = gstrSql & strWhere
    Me.rptAnalyse.Records.DeleteAll
    strNumber = Me.cbo����Ŀ��.List(Me.cbo����Ŀ��.ListIndex)
    strNumber = Mid(strNumber, 1, InStr(strNumber, "-") - 1)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo����.ItemData(Me.cbo����.ListIndex)), _
                    CDate(Format(Me.DTPStart, "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.DTPEnd, "yyyy-mm-dd 23:59:59")), Val(strBegingNO), Val(strEndNO), strNumber)
                        
    Do Until rsTmp.EOF
        
        Set Record = Me.rptAnalyse.Records.Add
            For intLoop = 0 To Me.rptSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mCol.�걾ID).Value = Nvl(rsTmp("ID"))
            Record.Item(mCol.�걾��).Value = Nvl(rsTmp("�걾���"))
            Record.Item(mCol.�걾��).Caption = Nvl(rsTmp("�걾����ʾ"))
            Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mCol.�Ա�).Value = Nvl(rsTmp("�Ա�"))
            Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
            Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
            Record.Item(mCol.��;).Value = Nvl(rsTmp("����"))
        rsTmp.MoveNext
    Loop
    Me.rptAnalyse.Populate
    '==========================================================================================================================================
    
End Sub

Private Sub cmdLeft_Click()
    Call SaveData(1)
End Sub

Private Sub cmdLeftAll_Click()
    Call SaveData(2)
End Sub

Private Sub cmdRight_Click()
    Call SaveData(3)
End Sub

Private Sub cmdRightAll_Click()
    Call SaveData(4)
End Sub

Private Sub Form_Load()
    Dim Column As ReportColumn
    Dim rsTmp As New ADODB.Recordset

    With Me.rptSource.Columns
        rptSource.AllowColumnRemove = False
        rptSource.ShowItemsInGroups = False
        
        With rptSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mCol.�걾ID, "�걾ID", 75, True): Column.Visible = False
        Set Column = .Add(mCol.�걾��, "�걾��", 75, True)
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 75, True)
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 75, True)
    End With
    
    With Me.rptAnalyse.Columns
        rptAnalyse.AllowColumnRemove = False
        rptAnalyse.ShowItemsInGroups = False
        
        With rptAnalyse.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mCol.�걾ID, "�걾ID", 75, True): Column.Visible = False
        Set Column = .Add(mCol.�걾��, "�걾��", 75, True)
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.�Ա�, "�Ա�", 75, True)
        Set Column = .Add(mCol.����, "����", 75, True)
        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 75, True)
        Set Column = .Add(mCol.��;, "��;", 100, True)
    End With
    
    Me.DTPStart = Now
    Me.DTPEnd = Now
    
    gstrSql = "select Id,����,���� from �������� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo����
        .Clear
        .AddItem "[�ֹ�]"
        .ItemData(.NewIndex) = -1
        Do While Not rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = Nvl(rsTmp("ID"))
            If mlngMachine = Nvl(rsTmp("ID")) Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
    gstrSql = "select ����,���� from ���������; "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo����Ŀ��
        .Clear
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("����")) & "-" & Nvl(rsTmp("����"))
            .ItemData(.NewIndex) = Nvl(rsTmp("����"))
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 And .ListIndex = -1 Then
            .ListIndex = 0
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With Me.fraFilter
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    
    With Me.ShortCapSource
        .Left = 0
        .Width = (Me.ScaleWidth / 2) - Me.CmdLeft.Width - (50 * 2)
    End With
    
    With Me.rptSource
        .Left = 0
        .Width = Me.ShortCapSource.Width
        .Height = Me.ScaleHeight - .Top
    End With
    
    With Me.ShortCaptAnalyse
        .Left = Me.ShortCapSource.Left + Me.ShortCapSource.Width + Me.CmdLeft.Width + 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With Me.rptAnalyse
        .Left = Me.ShortCaptAnalyse.Left
        .Width = Me.ShortCaptAnalyse.Width
        .Height = Me.ScaleHeight - .Top
    End With
    
    With Me.CmdLeft
        .Left = Me.rptSource.Left + Me.rptSource.Width + 50
        .Top = (Me.rptSource.Height / 4 / 2 * 1) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdLeftAll
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 2) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdRight
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 3) + (.Height / 2) + Me.rptSource.Top
    End With
    
    With Me.cmdRightAll
        .Left = Me.CmdLeft.Left
        .Top = (Me.rptSource.Height / 4 / 2 * 4) + (.Height / 2) + Me.rptSource.Top
    End With
End Sub
Public Sub ShowMe(Objfrm As Object, lngMachine As Long)
    mlngMachine = lngMachine
    Me.Show vbModal, Objfrm
End Sub
Public Sub SaveData(EditMode As Integer)
    '����               д�������ɾ������
    '����               EditMode
    '                   1=ɾ����ǰһ����������
    '                   2=ɾ����ǰ���з�������
    '                   3=���뵱ǰһ����������
    '                   4=ɾ����ǰ���з�������
    Dim lngLoop As Long
    Dim strAnalyse As String
    Dim intColCount As Integer
    Dim Record As ReportRecord
    
    Select Case EditMode
        Case 1, 2                                                   'ɾ���������ݼ�¼
            'û��ԭʼ���ݻ�û��ѡ����Ŀ��ʱ�˳�
            If Me.rptAnalyse.Records.Count = 0 Then
                MsgBox "û�����ݿ���ѡ��������ѡ���������в�ѯ!", vbInformation, Me.Caption
                Exit Sub
            End If
            strAnalyse = Me.cbo����Ŀ��.List(Me.cbo����Ŀ��.ListIndex)
            If EditMode = 1 Then
                If Me.rptAnalyse.FocusedRow Is Nothing Then
                    Exit Sub
                End If
                gstrSql = "Zl_���������¼_Edit(2," & Me.rptAnalyse.FocusedRow.Record(mCol.�걾ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                Set Record = Me.rptSource.Records.Add
                For intColCount = 0 To Me.rptSource.Columns.Count - 1
                    Record.AddItem ""
                Next
                Record.Item(mCol.�걾ID).Value = Me.rptAnalyse.FocusedRow.Record(mCol.�걾ID).Value
                Record.Item(mCol.�걾��).Value = Me.rptAnalyse.FocusedRow.Record(mCol.�걾��).Value
                Record.Item(mCol.�걾��).Caption = Me.rptAnalyse.FocusedRow.Record(mCol.�걾��).Caption
                Record.Item(mCol.����).Value = Me.rptAnalyse.FocusedRow.Record(mCol.����).Value
                Record.Item(mCol.�Ա�).Value = Me.rptAnalyse.FocusedRow.Record(mCol.�Ա�).Value
                Record.Item(mCol.����).Value = Me.rptAnalyse.FocusedRow.Record(mCol.����).Value
                Record.Item(mCol.����ʱ��).Value = Me.rptAnalyse.FocusedRow.Record(mCol.����ʱ��).Value
                
                Me.rptAnalyse.Records.RemoveAt (Me.rptAnalyse.FocusedRow.Index)
            Else
                'ɾ������
                If MsgBox("�Ƿ�ȷ��Ҫɾ����ǰ�����µ����з�������?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                    Exit Sub
                End If
                For lngLoop = 0 To Me.rptAnalyse.Records.Count - 1
                    gstrSql = "Zl_���������¼_Edit(2," & Me.rptAnalyse.Records(lngLoop).Item(mCol.�걾ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    
                    Set Record = Me.rptSource.Records.Add
                    For intColCount = 0 To Me.rptSource.Columns.Count - 1
                        Record.AddItem ""
                    Next
                    Record.Item(mCol.�걾ID).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.�걾ID).Value
                    Record.Item(mCol.�걾��).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.�걾��).Value
                    Record.Item(mCol.�걾��).Caption = Me.rptAnalyse.Records(lngLoop).Item(mCol.�걾��).Caption
                    Record.Item(mCol.����).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.����).Value
                    Record.Item(mCol.�Ա�).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.�Ա�).Value
                    Record.Item(mCol.����).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.����).Value
                    Record.Item(mCol.����ʱ��).Value = Me.rptAnalyse.Records(lngLoop).Item(mCol.����ʱ��).Value
                Next
                Me.rptAnalyse.Records.DeleteAll
            End If
            Me.rptAnalyse.Populate
            Me.rptSource.Populate
        Case 3, 4                                                   '�����������
            'û��ԭʼ���ݻ�û��ѡ����Ŀ��ʱ�˳�
            If Me.rptSource.Records.Count = 0 Then
                MsgBox "û�����ݿ���ѡ��������ѡ���������в�ѯ!", vbInformation, Me.Caption
                Exit Sub
            End If
            If Me.cbo����Ŀ��.ListCount = 0 Then
                MsgBox "��ѡ��һ������Ŀ��!", vbInformation, Me.Caption
                Me.cbo����Ŀ��.SetFocus
                Exit Sub
            End If
            strAnalyse = Me.cbo����Ŀ��.List(Me.cbo����Ŀ��.ListIndex)
            If EditMode = 3 Then
                '����д��
                If Me.rptSource.FocusedRow Is Nothing Then
                    Exit Sub
                End If
                gstrSql = "Zl_���������¼_Edit(1," & Me.rptSource.FocusedRow.Record(mCol.�걾ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                
                Set Record = Me.rptAnalyse.Records.Add
                For intColCount = 0 To Me.rptAnalyse.Columns.Count - 1
                    Record.AddItem ""
                Next
                Record.Item(mCol.�걾ID).Value = Me.rptSource.FocusedRow.Record(mCol.�걾ID).Value
                Record.Item(mCol.�걾��).Value = Me.rptSource.FocusedRow.Record(mCol.�걾��).Value
                Record.Item(mCol.�걾��).Caption = Me.rptSource.FocusedRow.Record(mCol.�걾��).Caption
                Record.Item(mCol.����).Value = Me.rptSource.FocusedRow.Record(mCol.����).Value
                Record.Item(mCol.�Ա�).Value = Me.rptSource.FocusedRow.Record(mCol.�Ա�).Value
                Record.Item(mCol.����).Value = Me.rptSource.FocusedRow.Record(mCol.����).Value
                Record.Item(mCol.����ʱ��).Value = Me.rptSource.FocusedRow.Record(mCol.����ʱ��).Value
                Record.Item(mCol.��;).Value = strAnalyse
                Me.rptSource.Records.RemoveAt (Me.rptSource.FocusedRow.Index)
                
            Else
                'д������
                For lngLoop = 0 To Me.rptSource.Records.Count - 1
                    gstrSql = "Zl_���������¼_Edit(1," & Me.rptSource.Records(lngLoop).Item(mCol.�걾ID).Value & ",'" & _
                            Mid(strAnalyse, 1, InStr(strAnalyse, "-") - 1) & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    Set Record = Me.rptAnalyse.Records.Add
                    For intColCount = 0 To Me.rptAnalyse.Columns.Count - 1
                        Record.AddItem ""
                    Next
                    Record.Item(mCol.�걾ID).Value = Me.rptSource.Records(lngLoop).Item(mCol.�걾ID).Value
                    Record.Item(mCol.�걾��).Value = Me.rptSource.Records(lngLoop).Item(mCol.�걾��).Value
                    Record.Item(mCol.�걾��).Caption = Me.rptSource.Records(lngLoop).Item(mCol.�걾��).Caption
                    Record.Item(mCol.����).Value = Me.rptSource.Records(lngLoop).Item(mCol.����).Value
                    Record.Item(mCol.�Ա�).Value = Me.rptSource.Records(lngLoop).Item(mCol.�Ա�).Value
                    Record.Item(mCol.����).Value = Me.rptSource.Records(lngLoop).Item(mCol.����).Value
                    Record.Item(mCol.����ʱ��).Value = Me.rptSource.Records(lngLoop).Item(mCol.����ʱ��).Value
                    Record.Item(mCol.��;).Value = strAnalyse
                Next
                Me.rptSource.Records.DeleteAll
            End If
            Me.rptSource.Populate
            Me.rptAnalyse.Populate
    End Select
End Sub

Private Sub rptAnalyse_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call SaveData(1)
End Sub

Private Sub rptSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call SaveData(3)
End Sub
