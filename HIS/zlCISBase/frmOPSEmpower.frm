VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmOPSEmpower 
   Caption         =   "������Ȩ����"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   Icon            =   "frmOPSEmpower.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12855
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicSQ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5640
      ScaleHeight     =   2775
      ScaleWidth      =   3855
      TabIndex        =   12
      Top             =   4440
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vsSQ 
         Height          =   6420
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   7305
         _cx             =   12885
         _cy             =   11324
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOPSEmpower.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         OwnerDraw       =   1
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
   Begin VB.PictureBox picOPS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   5640
      ScaleHeight     =   2535
      ScaleWidth      =   3855
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vsOPS 
         Height          =   6420
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   7305
         _cx             =   12885
         _cy             =   11324
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOPSEmpower.frx":68ED
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         OwnerDraw       =   1
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5220
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9208
      _StockProps     =   64
   End
   Begin VB.CheckBox chkEdit 
      Caption         =   "����Ȩ"
      Height          =   195
      Left            =   7080
      TabIndex        =   9
      ToolTipText     =   "Ctrl+��ѡ������ѡ��"
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox chkExec 
      Caption         =   "ִ��Ȩ"
      Height          =   195
      Left            =   8160
      TabIndex        =   8
      ToolTipText     =   "Ctrl+��ѡ������ѡ��"
      Top             =   120
      Value           =   1  'Checked
      Width           =   840
   End
   Begin VB.TextBox txtFindItem 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "���Ҳ���(Ctrl+F)"
      Top             =   120
      Width           =   1155
   End
   Begin VB.Frame fraDoctor 
      Caption         =   "ҽ��"
      ForeColor       =   &H000040C0&
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   5295
         Left            =   70
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   9340
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chk����� 
         Caption         =   "ֻ��ʾ����˵�ҽ��"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1050
         Width           =   2175
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   3
         Top             =   667
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   690
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8265
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   635
      SimpleText      =   $"frmOPSEmpower.frx":6988
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOPSEmpower.frx":69CF
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17595
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
   Begin MSComctlLib.ImageList img16 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":7263
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":DAC5
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":14327
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":148C1
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":14E5B
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmOPSEmpower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCodeType As Long         '0-ƴ��,1-���
Private mobjBar As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean
Private mblnNotRef As Boolean
Private mBln��Ȩ��� As Boolean

Private mlngFindNum As Long
Private mlngFindItemNum As Long    '������Ŀ
'���������ʱ������ǩ�����ܣ������жϼ��� And 1 = 0
Private mblnTmp As Boolean
Private Enum Enum_Dor
    COL_��ԱID = 0
    col_ѡ�� = 1
    COL_���� = 2
    col_�����ȼ� = 3
    COL_ƴ������ = 4
    COL_��ʼ��� = 5
    COL_�������� = 6
    COL_��������ID = 7
End Enum

Private Enum Enum_Advice
    col���� = 0
    colִ�� = 1
    col���� = 2
    col�������� = 3
    col�����ȼ� = 4
    COL������ģ = 5
    col������� = 6
    COLվ�� = 7
    COL���� = 8
    COL������ = 9
    COL����ʱ�� = 10
End Enum



Private Sub cboDept_Click()
    Call LoadDoc
End Sub

Private Sub SaveEmpower(ByVal lngType As Long)
'���ܣ���Ȩ
'������lngType��0-�ڿ�����ִ��Ȩ��1-�ڿ���Ȩ��2-��ִ��Ȩ
    Dim strSql As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long
    Dim strDocs As String, lngDoc As Long
    Dim arrSql() As Variant
    Dim strItems As String
    Dim blnTrans As Boolean
    
    For i = 0 To rptDoc.Records.Count - 1
        If rptDoc.Records(i).Tag = "1" Then
            strDocs = strDocs & "," & rptDoc.Records(i)(COL_��ԱID).Value
        End If
    Next
    strDocs = Mid(strDocs, 2)
    On Error GoTo errH
    If strDocs <> "" Then
        If InStr(strDocs, ",") > 0 Then
            '���������Ȩ�������Ƿ��Ѿ��ڹ�Ȩ����ʾ������Ȩ
            strSql = "Select /*+Rule */" & vbNewLine & _
                " f_List2str(Cast(Collect(����) As t_Strlist)) As ����" & vbNewLine & _
                "From (Select Distinct b.����" & vbNewLine & _
                "       From ��Ա����Ȩ�� A, ��Ա�� B" & vbNewLine & _
                "       Where a.��Աid = b.Id " & IIf(lngType > 0, " And A.��¼����=[2]", "") & " and a.��Աid In (Select Column_Value From Table(f_Num2list([1]))))"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDocs, lngType)
            If rsTmp.RecordCount > 0 Then
                If rsTmp!���� & "" <> "" Then
                    If MsgBox("����ҽ���Ѿ���Ȩ���Ƿ�Ҫȡ����Щҽ����Ȩ��������Ȩ��" & vbCrLf & rsTmp!����, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            lngDoc = 0
        Else
            lngDoc = Val(strDocs)
        End If
    Else
        lngDoc = Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value)
        strDocs = lngDoc
    End If
    strSql = _
            "Select ID, �ϼ�id, 0 As ĩ��, ����, ����, Null As ����, Null As �����ȼ�,  Null As ������ģ, Null As �������, Null As վ��," & vbNewLine & _
            "       Null As �ѹ�ѡcheck" & vbNewLine & _
            "From ���Ʒ���Ŀ¼" & vbNewLine & _
            "Where ���� =5 And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "Start With �ϼ�id Is Null" & vbNewLine & _
            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select  B.ID,B.����ID,1 ,b.����, b.����,Upper(E.����) as ����, d.��������, b.��������, decode(B.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������, b.վ��," & IIf(lngType = 0, "Decode(Count(A.��¼����), 2, 1, 0)", "Decode(Max(NVL(A.��¼����,0)),0,0,1) ") & vbNewLine & _
            "From ��Ա����Ȩ�� A, ������ĿĿ¼ B, ������϶��� C, ��������Ŀ¼ D,������Ŀ���� E" & vbNewLine & _
            "Where a.������Ŀid(+) = b.Id And b.Id = c.����id(+) And c.����id = d.Id(+) AND E.������ĿID=B.ID And a.��Աid(+) = [1] and e.����=[2] And e.����=1 And b.���='F' And (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            IIf(cboDept.ItemData(cboDept.ListIndex) <> -1, " and (exists(Select 1 From �������ÿ��� F Where F.��ĿID=b.ID And F.����ID=[3])  Or Not Exists(Select 1 From �������ÿ��� F Where F.��ĿID=b.ID))", "") & _
            IIf(lngType = 0, "", " And A.��¼����(+) = " & lngType) & _
            "Group By b.Id,B.����ID, b.����, b.����, b.��������, b.�������, b.վ��, d.��������,E.����"
    
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "������Ŀ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "����ʾû������ķ���", lngDoc, mlngCodeType + 1, cboDept.ItemData(cboDept.ListIndex))
    arrSql = Array()
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û���������ݿ���ѡ��", vbInformation, gstrSysName
        End If
    Else
        rsTmp.Filter = "�ѹ�ѡcheck=1 And ĩ��=1"
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If Len(strItems & "," & rsTmp!ID) > 4000 Then
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = "Zl_��Ա����Ȩ��_Update('" & strDocs & "','" & Mid(strItems, 2) & "'," & lngType & "," & IIf(UBound(arrSql) = 0, 1, 0) & IIf(mBln��Ȩ���, "", ",1,'" & UserInfo.���� & "'") & ")"
                    strItems = ""
                End If
                strItems = strItems & "," & rsTmp!ID
                
                rsTmp.MoveNext
            Loop
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_��Ա����Ȩ��_Update('" & strDocs & "','" & Mid(strItems, 2) & "'," & lngType & "," & IIf(UBound(arrSql) = 0, 1, 0) & IIf(mBln��Ȩ���, "", ",1,'" & UserInfo.���� & "'") & ")"
        Else
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_��Ա����Ȩ��_Update('" & strDocs & "',''," & lngType & ",1" & IIf(mBln��Ȩ���, "", ",0,1,'" & UserInfo.���� & "'") & ")"
        End If

        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        If tbcSub.Selected.Caption = "������Ŀ" Then
            Call LoadItem
            If Not mBln��Ȩ��� Then
                Call LoadCheck
            End If
        Else
            Call LoadCheck
        End If
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveModify()
'���ܣ�����ɾ��Ȩ��
    Dim i As Long
    Dim arrSql() As Variant
    Dim blnTrans As Boolean
    Dim intType As Integer
    Dim intDelete As Integer
    
    With vsOPS
        arrSql = Array()
        If chkEdit.Value = 1 And chkExec.Value = 0 Then
            intDelete = 3
        ElseIf chkEdit.Value = 0 And chkExec.Value = 1 Then
            intDelete = 4
        Else
            intDelete = 2
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col����) <> .Cell(flexcpData, i, col����) And chkEdit.Value = 1 Or .Cell(flexcpChecked, i, colִ��) <> .Cell(flexcpData, i, colִ��) And chkExec.Value = 1 Then
                If .Cell(flexcpChecked, i, col����) = 1 And .Cell(flexcpChecked, i, colִ��) = 2 Then
                    intType = 1
                ElseIf .Cell(flexcpChecked, i, col����) = 2 And .Cell(flexcpChecked, i, colִ��) = 1 Then
                    intType = 2
                ElseIf .Cell(flexcpChecked, i, col����) = 2 And .Cell(flexcpChecked, i, colִ��) = 2 Then
                    intType = 3
                Else
                    intType = 0
                End If
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = "Zl_��Ա����Ȩ��_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value) & "','" & .RowData(i) & "'," & intType & "," & intDelete & IIf(mBln��Ȩ���, "", ",1,'" & UserInfo.���� & "'") & ")"
            End If
        Next
        On err GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        If tbcSub.Selected.Caption = "������Ŀ" Then
            Call LoadItem
            If Not mBln��Ȩ��� Then
                Call LoadCheck
            End If
        Else
            Call LoadCheck
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveCheck(i As Integer)
'���ܣ�ִ����Ȩ����
'������i=0 ͨ�����롢=1�ܾ�����
    Dim strSql As String
    Dim blnTrans As Boolean

    If i = 1 Then
        strSql = "Zl_��Ա����Ȩ��_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value) & "','0',0,0,3,null,'" & UserInfo.���� & "')"
    Else
        strSql = "Zl_��Ա����Ȩ��_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value) & "','0',0,0,2,null,'" & UserInfo.���� & "')"
    End If
    On err GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(CStr(strSql), Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    Call LoadCheck
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadItem()
'���ܣ�����ҽ��ӵ�е�Ȩ�ޡ�

    Dim rsTmp As Recordset
    Dim strSql As String
    
    If rptDoc.SelectedRows.Count > 0 Then
        If rptDoc.SelectedRows(0).GroupRow = False Then
            strSql = "Select b.Id,Decode(Min(��¼����),1,1,2) As ����Ȩ,Decode(Max(��¼����),2,1,2) As ִ��Ȩ,b.����,b.����,f.����||'-'||f.���� as ��������," & _
                " decode(B.�������,1,'����',2,'סԺ',3,'�����סԺ',4,'���','��ֱ��Ӧ���ڲ���') as �������,b.վ��,d.��������,Upper(E.����) as ����" & _
                " From ��Ա����Ȩ�� A,������ĿĿ¼ B,������϶��� C,��������Ŀ¼ D,������Ŀ���� E,����������ģ F" & _
                " Where a.������Ŀid=b.Id And b.Id=c.����id(+) And c.����id=d.Id(+) AND E.������ĿID=B.ID And b.�������� in (f.����,f.����) And a.��Աid=[1]" & _
                " and e.����=[2] And (B.����ʱ�� Is Null Or B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                IIf(chkEdit.Value = 0, " And a.��¼���� <>1", "") & IIf(chkExec.Value = 0, " And a.��¼���� <>2", "") & _
                " Group By b.Id,b.����,b.����,b.�������,b.վ��,d.��������,E.����,f.����,f.����"
                
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value), mlngCodeType + 1)
            With vsOPS
                .Rows = 1
                Do While Not rsTmp.EOF
                    .AddItem ""
                    .RowData(.Rows - 1) = rsTmp!ID & ""
                    .Cell(flexcpChecked, .Rows - 1, col����) = Val(rsTmp!����Ȩ & "")
                    '����ȡ����ָ�״̬
                    .Cell(flexcpData, .Rows - 1, col����) = Val(rsTmp!����Ȩ & "")
                    .Cell(flexcpChecked, .Rows - 1, colִ��) = Val(rsTmp!ִ��Ȩ & "")
                    .Cell(flexcpData, .Rows - 1, colִ��) = Val(rsTmp!ִ��Ȩ & "")
                    .TextMatrix(.Rows - 1, col����) = rsTmp!����
                    .TextMatrix(.Rows - 1, col��������) = rsTmp!���� & ""
                    .TextMatrix(.Rows - 1, col�����ȼ�) = rsTmp!�������� & ""
                    .TextMatrix(.Rows - 1, COL������ģ) = rsTmp!�������� & ""
                    .TextMatrix(.Rows - 1, col�������) = rsTmp!������� & ""
                    .TextMatrix(.Rows - 1, COLվ��) = rsTmp!վ�� & ""
                    .TextMatrix(.Rows - 1, COL����) = rsTmp!���� & ""
                    rsTmp.MoveNext
                Loop
                
                If .Rows = 1 Then .AddItem ""
                .Cell(flexcpBackColor, 1, col����, .Rows - 1, colִ��) = &HE1FFE1
                If chkEdit.Value = 0 Then
                    .ColHidden(col����) = True
                Else
                    .ColHidden(col����) = False
                End If
                If chkExec.Value = 0 Then
                    .ColHidden(colִ��) = True
                Else
                    .ColHidden(colִ��) = False
                End If
            End With
        Else
            vsOPS.Rows = 1: vsOPS.AddItem ""
        End If
        mlngFindItemNum = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCheck()
'���ܣ����ش���˵�������Ȩ��
    Dim rsTmp As Recordset
    Dim strSql As String
    
    If rptDoc.SelectedRows.Count > 0 Then
        If rptDoc.SelectedRows(0).GroupRow = False Then
            strSql = "Select b.Id, Decode(A.Ȩ��,2, 2,3,2, 1) As ����Ȩ, Decode(A.Ȩ��,1, 2,3,2, 1) As ִ��Ȩ, b.����, b.����," & vbNewLine & _
                        "       f.���� || '-' || f.���� As ��������, Decode(b.�������, 1, '����', 2, 'סԺ', 3, '�����סԺ', 4, '���', '��ֱ��Ӧ���ڲ���') As �������, b.վ��," & vbNewLine & _
                        "       d.��������, Upper(e.����) As ����,a.������,a.����ʱ��,a.���״̬,a.������,a.����ʱ��" & vbNewLine & _
                        "From ��Ա����Ȩ������ A, ������ĿĿ¼ B, ������϶��� C, ��������Ŀ¼ D, ������Ŀ���� E, ����������ģ F" & vbNewLine & _
                        "Where a.������Ŀid = b.Id And A.���״̬ =1 And b.Id = c.����id(+) And c.����id = d.Id(+) And e.������Ŀid = b.Id And b.�������� In (f.����, f.����) And" & vbNewLine & _
                        "      a.��Ȩ��Աid = [1] And e.���� = [2] And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) order by a.����ʱ��,b.����"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptDoc.SelectedRows(0).Record(COL_��ԱID).Value), mlngCodeType + 1)
            If Not rsTmp.EOF Then
                tbcSub.Item(1).Caption = IIf(rsTmp.RecordCount = 0, "�������Ȩ", "�������Ȩ��" & rsTmp.RecordCount & " ��")
            Else
                tbcSub.Item(1).Caption = "�������Ȩ"
            End If
            With vsSQ
                .Rows = 1
                Do While Not rsTmp.EOF
                    .AddItem ""
                    .RowData(.Rows - 1) = rsTmp!ID & ""
                    .Cell(flexcpChecked, .Rows - 1, col����) = Val(rsTmp!����Ȩ & "")
                    '����ȡ����ָ�״̬
                    .Cell(flexcpData, .Rows - 1, col����) = Val(rsTmp!����Ȩ & "")
                    .Cell(flexcpChecked, .Rows - 1, colִ��) = Val(rsTmp!ִ��Ȩ & "")
                    .Cell(flexcpData, .Rows - 1, colִ��) = Val(rsTmp!ִ��Ȩ & "")
                    .TextMatrix(.Rows - 1, col����) = rsTmp!����
                    .TextMatrix(.Rows - 1, col��������) = rsTmp!���� & ""
                    .TextMatrix(.Rows - 1, col�����ȼ�) = rsTmp!�������� & ""
                    .TextMatrix(.Rows - 1, COL������ģ) = rsTmp!�������� & ""
                    .TextMatrix(.Rows - 1, col�������) = rsTmp!������� & ""
                    .TextMatrix(.Rows - 1, COLվ��) = rsTmp!վ�� & ""
                    .TextMatrix(.Rows - 1, COL����) = rsTmp!���� & ""
                    .TextMatrix(.Rows - 1, COL������) = rsTmp!������ & ""
                    .TextMatrix(.Rows - 1, COL����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd HH:mm")
                    rsTmp.MoveNext
                Loop
                
                If .Rows = 1 Then .AddItem ""
            End With
        Else
            vsSQ.Rows = 1: vsSQ.AddItem ""
        End If
        mlngFindItemNum = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadDoc()
'����Ȩ�ޱȲ���Ա�͵�ҽ��
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long, y As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngPrssID As Long, lngSelectRow As Long, lngDept As Long
    
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    If Val(cboDept.ItemData(cboDept.ListIndex)) = -1 Then
        rptDoc.GroupsOrder.DeleteAll
    Else
        If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 And rptDoc.GroupsOrder.Count = 0 Then rptDoc.GroupsOrder.Add rptDoc.Columns(COL_��������)
    End If
    strSql = "Select DISTINCT a.Id, A.�Ա�" & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", ",b.����ID,e.���� as ��������") & ",a.����,a.�����ȼ�, Upper(zlSpellCode(a.����)) As ƴ������, Upper(Zlwbcode(a.����)) As ��ʼ���" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� D,���ű� E" & vbNewLine & _
            "Where a.Id = b.��Աid And e.ID=b.����ID And d.��Աid = a.Id  And d.��Ա���� = 'ҽ��' And " & vbNewLine & _
            "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)  " & vbNewLine & _
            "   " & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.����id=[2]") & IIf(chk�����.Value = 1, " And (Exists(Select 1 From ��Ա����Ȩ������ F Where F.��Ȩ��Աid = A.id And F.���״̬ = 1))", "")
            
    
    On Error GoTo errH
    
    rptDoc.Records.DeleteAll
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)))
    
    With rptDoc
        i = 0
        lngSelectRow = -1
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!�Ա� & "" = "Ů", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!�����ȼ� & "")
            Set objItem = objRecord.AddItem(rsTmp!ƴ������ & "")
            Set objItem = objRecord.AddItem(rsTmp!��ʼ��� & "")
            If Val(cboDept.ItemData(cboDept.ListIndex)) <> -1 Then
                Set objItem = objRecord.AddItem(rsTmp!�������� & "")
                Set objItem = objRecord.AddItem(rsTmp!����ID & "")
            End If

            
            rsTmp.MoveNext
            i = i + 1
        Loop
        .Populate
        If lngPrssID <> 0 Then
            vsOPS.Rows = 1
            vsOPS.AddItem ""
        End If
    End With
    mlngFindNum = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_Untread     'ȡ��
        With vsOPS
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col����) <> .Cell(flexcpData, i, col����) Or .Cell(flexcpChecked, i, colִ��) <> .Cell(flexcpData, i, colִ��) Then
                    .Cell(flexcpChecked, i, col����) = .Cell(flexcpData, i, col����)
                    .Cell(flexcpChecked, i, colִ��) = .Cell(flexcpData, i, colִ��)
                End If
            Next
        End With
        mblnIsUpdate = False
    Case conMenu_Manage_Complete '��Ȩͨ��
        If MsgBox("ȷ��Ҫͨ����" & IIf(rptDoc.SelectedRows(0).Record(COL_����).Value = "", "��", rptDoc.SelectedRows(0).Record(COL_����).Value) & "ҽ������Ȩ���룿", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
        Call SaveCheck(0)
    Case conMenu_Manage_UnArrange '��Ȩ�ܾ�
        If MsgBox("ȷ��Ҫ�ܾ���" & IIf(rptDoc.SelectedRows(0).Record(COL_����).Value = "", "��", rptDoc.SelectedRows(0).Record(COL_����).Value) & "ҽ������Ȩ���룿", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
        Call SaveCheck(1)
    Case conMenu_Edit_Save        '����
        Call SaveModify
        mblnIsUpdate = False
    Case conMenu_Kss_Grant  '�ڿ�����ִ��Ȩ
        Call SaveEmpower(0)
    Case conMenu_Kss_Grant * 100# + 1 '�ڿ���Ȩ
        Call SaveEmpower(1)
    Case conMenu_Kss_Grant * 100# + 2 '��ִ��Ȩ
        Call SaveEmpower(2)
    Case conMenu_View_Find '����
        txtFind.SetFocus '��ʱ��Ҫ��λһ��
        If txtFind.Text <> "" Then
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_FindNext '������һ��
        If Me.ActiveControl.Name = "txtFindItem" Or Me.ActiveControl.Name = "vsOPS" Then
            If txtFindItem.Text = "" Then
                txtFindItem.SetFocus
            Else
                Call txtFindItem_KeyPress(vbKeyReturn)
            End If
        Else
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call txtFind_KeyPress(vbKeyReturn)
            End If
        End If
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_View_Refresh 'ˢ��
        Call LoadItem
        Call LoadCheck
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With fraDoctor
        .Top = lngTop
        .Left = lngLeft + 100
        .Height = lngBottom - lngTop - stbThis.Height
    End With
    rptDoc.Height = fraDoctor.Height - IIf(mBln��Ȩ���, 1540, 1250)
    
    
    With tbcSub
        .Top = fraDoctor.Top
        .Height = fraDoctor.Height
        .Width = Me.Width - fraDoctor.Left - fraDoctor.Width - 400
        picOPS.Width = .Width
        picOPS.Height = .Height - 350
        vsOPS.Top = 0: vsOPS.Left = 0: vsOPS.Width = picOPS.Width: vsOPS.Height = picOPS.Height
        PicSQ.Width = .Width
        PicSQ.Height = .Height - 350
        vsSQ.Top = 0: vsSQ.Left = 0: vsSQ.Width = PicSQ.Width: vsSQ.Height = PicSQ.Height
    End With
    
    
    Me.Refresh
End Sub

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '����Ȩ�����ð�ť�ɼ�״̬
    
    Select Case Control.ID
            Case conMenu_Kss_Grant, conMenu_Kss_Grant * 100# + 1, conMenu_Kss_Grant * 100# + 2, conMenu_Edit_Save, conMenu_Edit_Untread
                Control.Visible = tbcSub.Selected.Caption = "������Ŀ"
            Case conMenu_Manage_Complete, conMenu_Manage_UnArrange
                Control.Visible = tbcSub.Selected.Caption <> "������Ŀ"
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '����Ȩ�����ð�ť�ɼ�״̬
    If mblnIsUpdate Then
        If Control.ID = conMenu_Edit_Untread Or Control.ID = conMenu_Edit_Save Then
            Control.Enabled = True
            If Visible And fraDoctor.Enabled = True Then fraDoctor.Enabled = False
        Else
            Control.Enabled = False
        End If
        Exit Sub
    Else
        Control.Enabled = True
        If Visible And fraDoctor.Enabled = False Then fraDoctor.Enabled = True
    End If
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
        Case conMenu_Edit_Untread, conMenu_Edit_Save
            Control.Enabled = mblnIsUpdate
        Case conMenu_Kss_Grant  '��Ȩ
            blnEnabled = rptDoc.SelectedRows.Count > 0
            If rptDoc.SelectedRows.Count > 0 Then
                blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '������һ��
            Control.Visible = False
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
        Case conMenu_Manage_Complete '���ͨ��
            Control.Enabled = mBln��Ȩ���
            If Control.Enabled = True Then
                blnEnabled = rptDoc.SelectedRows.Count > 0
                If rptDoc.SelectedRows.Count > 0 Then
                    blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
                End If
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Manage_UnArrange '��˲�ͨ��
            Control.Enabled = mBln��Ȩ���
            If Control.Enabled = True Then
                blnEnabled = rptDoc.SelectedRows.Count > 0
                If rptDoc.SelectedRows.Count > 0 Then
                    blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
                End If
                Control.Enabled = blnEnabled
            End If
    End Select
End Sub


Private Sub chkEdit_Click()
    If chkExec.Value = 0 And chkEdit.Value = 0 Then
        mblnNotRef = True
        chkEdit.Value = 1: Exit Sub
        mblnNotRef = False
    End If
    If mblnNotRef = True Then
        mblnNotRef = False
        Exit Sub
    End If
    If tbcSub.Selected.Caption = "������Ŀ" Then
        Call LoadItem
    End If
End Sub

Private Sub chkExec_Click()
    If chkEdit.Value = 0 And chkExec.Value = 0 Then
        mblnNotRef = True
        chkExec.Value = 1: Exit Sub
        mblnNotRef = False
    End If
    If mblnNotRef = True Then
        mblnNotRef = False
        Exit Sub
    End If
    If tbcSub.Selected.Caption = "������Ŀ" Then
        Call LoadItem
    End If
End Sub

Private Sub chk�����_Click()
    Call LoadDoc
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mstrPrivs = GetPrivFunc(glngSys, 1080)
    mBln��Ȩ��� = InStr(mstrPrivs, "��Ȩ���") > 0
    
    mlngModul = 1080
    mlngCodeType = zlDatabase.GetPara("���뷽ʽ")
    mblnIsUpdate = False
    
    rptDoc.Top = IIf(mBln��Ȩ���, 1440, 1150)
    chk�����.Visible = mBln��Ȩ���
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "������Ŀ", picOPS.hwnd, 0).Tag = "������Ŀ"
        .InsertItem(1, "�������Ȩ", PicSQ.hwnd, 0).Tag = "�������Ȩ"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    
    'vsFlexGrid
    '-----------------------------------------------------
    strHead = "����,700,1;ִ��,700,1 ;����,1000,1;��������,2500,1;�����ȼ�,1000,1;������ģ,1000,1;�������,1000,1;Ժ��,800,7;����"
    Call InitTable(vsOPS, strHead)
    vsOPS.Editable = flexEDKbdMouse
    vsOPS.Cell(flexcpPictureAlignment, 0, col����) = flexPicAlignLeftCenter
    vsOPS.Cell(flexcpPictureAlignment, 0, colִ��) = flexPicAlignLeftCenter
    vsOPS.Cell(flexcpAlignment, 0, col����) = flexPicAlignRightCenter
    vsOPS.Cell(flexcpAlignment, 0, colִ��) = flexPicAlignRightCenter
    vsOPS.Cell(flexcpPicture, 0, col����) = img16.ListImages("unCheck").Picture
    vsOPS.Cell(flexcpPicture, 0, colִ��) = img16.ListImages("unCheck").Picture
    vsOPS.ColDataType(col����) = flexDTBoolean
    vsOPS.ColDataType(colִ��) = flexDTBoolean
    
    strHead = "����,700,1;ִ��,700,1 ;����,1000,1;��������,2500,1;�����ȼ�,1000,1;������ģ,1000,1;�������,1000,1;Ժ��,800,7;����;������,700,1;����ʱ��,1700,1"
    Call InitTable(vsSQ, strHead)
    vsSQ.Editable = flexEDNone
    vsSQ.ColDataType(col����) = flexDTBoolean
    vsSQ.ColDataType(colִ��) = flexDTBoolean

    
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    
    Call LoadDept
End Sub

Private Sub LoadDept()
'���ز���Ա��������
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select B.ID,B.����,B.���� " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", ",A.ȱʡ") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", "������Ա A, ") & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";���в���;") > 0, "", " And a.����id = B.Id And A.��ԱID = [1] ") & vbNewLine & _
            "  And C.�������� = '�ٴ�' And C.������� <> 0  And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"

    On Error GoTo errH
    cboDept.Clear
    '���в���
    If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 Then
        cboDept.AddItem "���в���"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '����ȱʡ
        If InStr(";" & mstrPrivs & ";", ";���в���;") = 0 Then
            If rsTmp!ȱʡ = 1 Then
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptDoc
        
        Set objCol = .Columns.Add(COL_��ԱID, "��ԱID", 0, False)
        Set objCol = .Columns.Add(col_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(COL_����, "����", 70, True)
        Set objCol = .Columns.Add(col_�����ȼ�, "�����ȼ�", 80, True)
        Set objCol = .Columns.Add(COL_ƴ������, "ƴ������", 0, False)
        Set objCol = .Columns.Add(COL_��ʼ���, "��ʼ���", 0, False)
        Set objCol = .Columns.Add(COL_��������, "��������", 0, False)
        Set objCol = .Columns.Add(COL_��������ID, "��������ID", 0, False)


        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ��ҽ��..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        If InStr(";" & mstrPrivs & ";", ";���в���;") > 0 Then .GroupsOrder.Add .Columns(COL_��������)
    End With
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Kss_Grant, "�ڿ�����ִ��Ȩ")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "�ڿ�����ִ��Ȩ")
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 1, "�ڿ���Ȩ")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 2, "��ִ��Ȩ")
            objControl.IconId = conMenu_Kss_Grant
        End With
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("������", xtpBarTop)
    With mobjBar.Controls

        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "��Ȩͨ��"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_UnArrange, "�ܾ���Ȩ"): objControl.IconId = 4114
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Kss_Grant, "�ڿ�����ִ��Ȩ")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "�ڿ�����ִ��Ȩ")
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 1, "�ڿ���Ȩ")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 2, "��ִ��Ȩ")
            objControl.IconId = conMenu_Kss_Grant
        End With

        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
    End With
    With cbsMain.ActiveMenuBar.Controls
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Show * 100# + 1, "")
        objCustom.Handle = chkEdit.hwnd
        objCustom.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Show * 100# + 2, "")
        objCustom.Handle = chkExec.hwnd
        objCustom.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlCustom, conMenu_View_Find * 100# + 1, "  ����")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFindItem.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsUpdate = True Then
        If MsgBox("��ǰ���������δ���棬�Ƿ�Ҫ�˳���", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    mlngFindNum = 0
    mlngFindItemNum = 0
End Sub

Private Sub rptDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptDoc.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptDoc_RowDblClick(rptDoc.SelectedRows(0), rptDoc.SelectedRows(0).Record.Item(col_ѡ��))
        End If
    End If
End Sub

Private Sub rptDoc_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptDoc.HitTest(x, y).ht = xtpHitTestHeader Then
            Set objColumn = rptDoc.HitTest(x, y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptDoc.Columns(col_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(col_ѡ��).Icon = img16.ListImages("Check").Index - 1
                            rptDoc.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptDoc.Columns(col_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(col_ѡ��).Icon = -1
                            rptDoc.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptDoc_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(col_ѡ��).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(col_ѡ��).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptDoc.Populate
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '��������Ȩ��������Ŀ
    If tbcSub.Selected.Caption = "������Ŀ" Then
        Call LoadItem
    Else
        Call LoadCheck
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "������Ŀ" Then
        Call LoadItem
    Else
        Call LoadCheck
    End If
End Sub

Private Sub txtFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If zlCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(COL_����).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(mlngCodeType = 0, COL_ƴ������, COL_��ʼ���)).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Then
                        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(COL_����).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "��ǰ����û���ҵ������ҵ�ҽ����", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "�Ѿ������һ��ҽ���ˡ�", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub

Private Sub txtFindItem_Change()
    mlngFindItemNum = 0
End Sub

Private Sub txtFindItem_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFindItem)
    End If
End Sub

Private Sub txtFindItem_KeyPress(KeyAscii As Integer)
    Dim i As Long, int���� As Integer
    Dim strFind As String
    If KeyAscii = vbKeyReturn Then
        With vsOPS
            strFind = UCase(Trim(txtFindItem.Text))
            If zlCommFun.IsCharChinese(txtFindItem.Text) Then
                '���ĵ�ֻ������
                int���� = 1
            ElseIf zlCommFun.IsCharAlpha(txtFindItem.Text) Then
                'Ӣ�Ĳ����ƺͼ���
                int���� = 2
            Else
                '��������Ƽ���ͱ���
                int���� = 3
            End If
            For i = mlngFindItemNum To .Rows - 1
                If int���� = 1 Then
                    If UCase(.TextMatrix(i, col��������)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Then
                        .Row = i
                        .ShowCell i, col��������
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                ElseIf int���� = 2 Then
                    If UCase(.TextMatrix(i, col��������)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, COL����)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Then
                        .Row = i
                        .ShowCell i, col��������
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                Else
                    If UCase(.TextMatrix(i, col��������)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, COL����)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, col����)) = strFind Then
                        .Row = i
                        .ShowCell i, col��������
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                End If
            Next
            If mlngFindItemNum = 0 Then
                MsgBox "��ǰҽ�����߱�����ѯ��������Ŀ������ִ��Ȩ�ޡ�", vbInformation, Me.Caption
            ElseIf mlngFindItemNum <> 0 Then
                MsgBox "�Ѿ����������һ����Ŀ�ˡ�", vbInformation, Me.Caption
                mlngFindItemNum = 0
            End If
        End With
    End If
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col���� Or Col = colִ�� Then
        If mblnIsUpdate = False Then mblnIsUpdate = True
    End If
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mlngFindItemNum <> 0 Then mlngFindItemNum = NewRow
End Sub

Private Sub vsOPS_AfterSort(ByVal Col As Long, Order As Integer)
    mlngFindItemNum = 0
End Sub

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col���� And Col <> colִ�� Then
        Cancel = True
    End If
End Sub

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsOPS.RowData(Row) & "" = "" Then
        Cancel = True
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptDoc.Visible = False Then Exit Sub
    If rptDoc.Records.Count > 0 Then
        If rptDoc.SelectedRows.Count = 0 Then Exit Sub
        If rptDoc.SelectedRows(0).GroupRow Then Exit Sub
        strSubhead = rptDoc.SelectedRows(0).Record(COL_����).Value & IIf(tbcSub.Selected.Caption = "������Ŀ", "����Ȩ���嵥", "�����Ȩ��")
    Else
        Exit Sub
    End If
    
    '���ô�ӡ��������
    Set objPrint.Body = IIf(tbcSub.Selected.Caption = "������Ŀ", Me.vsOPS, Me.vsSQ)
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

