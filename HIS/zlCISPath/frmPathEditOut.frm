VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathEditOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ٴ�·����Ϣ"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "frmPathEditOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtMaxDay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   200
      IMEMode         =   2  'OFF
      Left            =   1500
      MaxLength       =   3
      TabIndex        =   28
      Top             =   6960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������"
      Height          =   195
      Left            =   360
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7545
      Width           =   1020
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2175
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Height          =   75
      Left            =   -105
      TabIndex        =   25
      Top             =   7335
      Width           =   7200
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   225
      TabIndex        =   24
      Top             =   1980
      Width           =   6870
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   210
      TabIndex        =   23
      Top             =   540
      Width           =   6885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5040
      TabIndex        =   27
      Top             =   7470
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3900
      TabIndex        =   26
      Top             =   7470
      Width           =   1100
   End
   Begin VB.OptionButton optӦ�÷�Χ 
      Caption         =   "ָ������"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   3000
      Value           =   -1  'True
      Width           =   1020
   End
   Begin VB.OptionButton optӦ�÷�Χ 
      Caption         =   "ȫԺͨ��"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   2640
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   1200
      Left            =   3330
      TabIndex        =   17
      Top             =   2595
      Width           =   2955
      _cx             =   5212
      _cy             =   2117
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathEditOut.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.TextBox txt˵�� 
      Height          =   510
      Left            =   960
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1430
      Width           =   5325
   End
   Begin VB.ComboBox cbo���䵥λ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5565
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2175
      Width           =   720
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5175
      MaxLength       =   3
      TabIndex        =   12
      Top             =   2175
      Width           =   360
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4635
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2175
      Width           =   360
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   690
      Width           =   2415
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   960
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1050
      Width           =   5295
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4575
      MaxLength       =   5
      TabIndex        =   3
      Top             =   690
      Width           =   1680
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDisease 
      Height          =   2625
      Left            =   300
      TabIndex        =   19
      Top             =   4200
      Width           =   5955
      _cx             =   10504
      _cy             =   4630
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPathEditOut.frx":05BE
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblConfirmDay 
      Caption         =   "�����ʱ�䣺_______ ��"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   7020
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblConfirm 
      Caption         =   "���������ʱ��δ�������򲻼�����·����"
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   7020
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�����ٴ�·���Ļ�����Ϣ�����ö���Ӧ�÷�Χ����Ӧ���ֵ�����"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   720
      TabIndex        =   22
      Top             =   195
      Width           =   5220
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmPathEditOut.frx":0623
      Top             =   50
      Width           =   480
   End
   Begin VB.Label lbl��Ӧ���� 
      AutoSize        =   -1  'True
      Caption         =   "�������ò���(&I)"
      Height          =   180
      Left            =   285
      TabIndex        =   18
      Top             =   3960
      Width           =   1350
   End
   Begin VB.Label lblӦ�÷�Χ 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ�÷�Χ(&S)"
      Height          =   180
      Left            =   285
      TabIndex        =   14
      Top             =   2625
      Width           =   990
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      Caption         =   "˵��(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   6
      Top             =   1480
      Width           =   630
   End
   Begin VB.Label lbl���䷶Χ 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Left            =   5040
      TabIndex        =   21
      Top             =   2235
      Width           =   90
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "��������(&Y)"
      Height          =   180
      Left            =   3495
      TabIndex        =   10
      Top             =   2235
      Width           =   990
   End
   Begin VB.Label lbl�����Ա� 
      AutoSize        =   -1  'True
      Caption         =   "�����Ա�(&X)"
      Height          =   180
      Left            =   285
      TabIndex        =   8
      Top             =   2235
      Width           =   990
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����(&K)"
      Height          =   180
      Left            =   255
      TabIndex        =   0
      Top             =   750
      Width           =   630
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   1110
      Width           =   630
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����(&C)"
      Height          =   180
      Left            =   3840
      TabIndex        =   2
      Top             =   750
      Width           =   630
   End
End
Attribute VB_Name = "frmPathEditOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AfterSave(ByVal ���� As String, ByVal ���� As String)

Private mlng·��ID  As Long
Private mstrPrivs   As String
Private mstr����    As String
Private mblnReturn  As Boolean
Private mblnChange  As Boolean
Private mblnOK      As Boolean

Public Function ShowEdit(frmMain As Object, ByVal strPrivs As String, Optional ByVal lng·��ID As Long, Optional ByVal str���� As String) As Boolean
'���ܣ����������޸��ٴ�·��
'������lng·��ID=�޸�ʱ��������IDֵ������ʱ������
'      str����=����ʱ�����뵱ǰ��ѡ��ķ�����Ϊȱʡ��Ҳ���Բ�����
    mstrPrivs = strPrivs
    mlng·��ID = lng·��ID
    mstr���� = str����
    
    Me.Show 1, frmMain
    ShowEdit = mblnOK
End Function

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_Click()
    If mlng·��ID = 0 Then
        txt����.Text = GetNextCode(cbo����.Text, 1)
    End If
    If vsDept.Enabled Then
        vsDept.Rows = 1
        vsDept.Rows = 2
        Call AddDept
    End If
    mblnChange = True
End Sub

Private Sub cbo����_GotFocus()
    Call zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    If mlng·��ID = 0 And cbo����.ListIndex = -1 Then
        txt����.Text = GetNextCode(cbo����.Text, 1)
    End If
End Sub

Private Sub cbo���䵥λ_Click()
    mblnChange = True
End Sub

Private Sub cbo�����Ա�_Click()
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'�����д��Ϣ���ұ���
    Dim rsTmp As ADODB.Recordset
    Dim str����IDs As String
    Dim str����IDs As String
    Dim strSql As String, i As Long
    Dim strTmp As String, intLimit As Integer

    '1)�����������Ŀ
    If cbo����.Text = "" Then
        MsgBox "����ָ���ٴ�·���ķ��ࡣ", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If txt����.Text = "" Then
        MsgBox "����ָ���ٴ�·���ı��롣", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If txt����.Text = "" Then
        MsgBox "����ָ���ٴ�·�������ơ�", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If txt��������(0).Text <> "" And txt��������(1).Text = "" Or _
       txt��������(0).Text = "" And txt��������(1).Text <> "" Then
        MsgBox "�ٴ�·�������õ�����Ӧ����һ����Χ��", vbInformation, gstrSysName
        If txt��������(0).Text = "" Then
            txt��������(0).SetFocus
        Else
            txt��������(1).SetFocus
        End If
        Exit Sub
    End If

    '2)���볤�ȼ��
    If zlCommFun.ActualLen(cbo����.Text) > 50 Then
        MsgBox "�ٴ�·���ķ�����Ϣ���ֻ���� 25 �����ֻ� 50 ���ַ���", vbInformation, gstrSysName
        cbo����.SetFocus
        Exit Sub
    End If
    If zlCommFun.ActualLen(txt����.Text) > txt����.MaxLength Then
        MsgBox "�ٴ�·�����������ֻ���� " & txt����.MaxLength \ 2 & " �����ֻ� " & txt����.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt����.SetFocus
        Exit Sub
    End If
    If zlCommFun.ActualLen(txt˵��.Text) > txt˵��.MaxLength Then
        MsgBox "�ٴ�·����˵����Ϣ���ֻ���� " & txt˵��.MaxLength \ 2 & " �����ֻ� " & txt˵��.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt˵��.SetFocus
        Exit Sub
    End If

    '3)�������
    If optӦ�÷�Χ(1).Value Then
        With vsDept
            strTmp = ""
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, 0) <> "" Then
                    If InStr(strTmp & ",", "," & .RowData(i) & ",") > 0 Then
                        MsgBox "���ִ���������ͬ�Ŀ��ҡ�", vbInformation, gstrSysName
                        .Row = i: .Col = 0
                        .ShowCell .Row, .Col
                        .SetFocus
                        Exit Sub
                    Else
                        strTmp = strTmp & "," & .RowData(i)
                    End If
                End If
            Next
        End With
    End If
    With vsDisease
        strTmp = ""
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, 2) <> "" Then
                If Val(.TextMatrix(i, 0)) <> 0 Then
                    strSql = "A" & .RowData(i)
                Else
                    strSql = "B" & .RowData(i)
                End If

                If InStr(strTmp & ",", "," & strSql & ",") > 0 Then
                    MsgBox "���ִ���������ͬ�Ĳ��֡�", vbInformation, gstrSysName
                    .Row = i: .Col = 2
                    .ShowCell .Row, .Col
                    .SetFocus
                    Exit Sub
                Else
                    strTmp = strTmp & "," & strSql
                End If
            End If
        Next
    End With

    '4)��������
    If optӦ�÷�Χ(1).Value Then
        For i = 1 To vsDept.Rows - 1
            If vsDept.RowData(i) <> 0 Then
                str����IDs = str����IDs & "," & vsDept.RowData(i)
            End If
        Next
        str����IDs = Mid(str����IDs, 2)
        If str����IDs = "" Then
            MsgBox "����ָ���ٴ�·���Ŀ���Ӧ�÷�Χ��", vbInformation, gstrSysName
            vsDept.SetFocus
            Exit Sub
        End If
    End If
    
    For i = 1 To vsDisease.Rows - 1
        If vsDisease.RowData(i) <> 0 And Val(vsDisease.TextMatrix(i, 0)) <> 0 Then
            str����IDs = str����IDs & "," & vsDisease.RowData(i)
        End If
    Next
    str����IDs = Mid(str����IDs, 2) & ";"

    For i = 1 To vsDisease.Rows - 1
        If vsDisease.RowData(i) <> 0 And Val(vsDisease.TextMatrix(i, 1)) <> 0 Then
            str����IDs = str����IDs & vsDisease.RowData(i) & ","
        End If
    Next

    If str����IDs = ";" Then
        MsgBox "����ָ���ٴ�·������Ӧ�ĵĲ��֡�", vbInformation, gstrSysName
        vsDisease.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtMaxDay.Text) = "" Then
'        MsgBox "�������������ʱ�䡣", vbInformation, gstrSysName
'        txtMaxDay.SetFocus
'        Exit Sub
'    End If

    'ȥ�����ұߵĶ���
    If Right(str����IDs, 1) = "," Then
        str����IDs = Left(str����IDs, Len(str����IDs) - 1)
    End If
    If mlng·��ID = 0 Then
        strSql = "Zl_����·��Ŀ¼_Insert('" & cbo����.Text & "','" & txt����.Text & "','" & txt����.Text & "','" & txt˵��.Text & "'," & _
                 cbo�����Ա�.ListIndex & ",'" & IIf(txt��������(0).Text <> "", txt��������(0).Text & "-" & txt��������(1).Text & cbo���䵥λ.Text, "") & "'," & _
                 IIf(optӦ�÷�Χ(0).Value, 1, 2) & "," & Val(txtMaxDay.Text) & ",'" & str����IDs & "','" & str����IDs & "',Null)"
    Else
        strSql = "Zl_����·��Ŀ¼_Update(" & mlng·��ID & ",'" & cbo����.Text & "','" & txt����.Text & "','" & txt����.Text & "','" & txt˵��.Text & "'," & _
                 cbo�����Ա�.ListIndex & ",'" & IIf(txt��������(0).Text <> "", txt��������(0).Text & "-" & txt��������(1).Text & cbo���䵥λ.Text, "") & "'," & _
                 IIf(optӦ�÷�Χ(0).Value, 1, 2) & "," & Val(txtMaxDay.Text) & ",'" & str����IDs & "','" & str����IDs & "')"
    End If

    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    On Error GoTo 0

    '5)��ɴ���
    mblnOK = True
    RaiseEvent AfterSave(cbo����.Text, txt����.Text)

    '��������
    If mlng·��ID = 0 And chk����.Value = 1 Then
        '��Ȩ�������
        If InStr(mstrPrivs, "������·������") > 0 Then
            intLimit = 0
        ElseIf InStr(mstrPrivs, "30������·��") > 0 Then
            intLimit = 30
        ElseIf InStr(mstrPrivs, "5������·��") > 0 Then
            intLimit = 5
        End If
        If intLimit > 0 Then
            On Error GoTo errH
            strSql = "Select Nvl(Count(*),0) as ���� From ����·��Ŀ¼"
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
            If rsTmp!���� < intLimit Then
                intLimit = 0
            End If
            On Error GoTo 0
        End If
        If intLimit = 0 Then
            txt����.Text = GetNextCode(cbo����.Text, 1)
            txt����.Text = ""
            txt˵��.Text = ""
            mblnChange = False
            txt����.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeName(Me.ActiveControl) <> "VSFlexGrid" Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intIdx As Integer

    On Error GoTo errH

    mblnOK = False
    
    '������Ϣ
    strSql = "Select Distinct ���� From ����·��Ŀ¼ Where ���� is Not NULL Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!����
        rsTmp.MoveNext
    Loop

    '�Ա�
    cbo�����Ա�.AddItem "0-�������Ա�"
    cbo�����Ա�.AddItem "1-����"
    cbo�����Ա�.AddItem "2-Ů��"
    cbo�����Ա�.ListIndex = 0

    '���䵥λ
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0

    'Ȩ������
    optӦ�÷�Χ(1).Value = True    'ȱʡΪָ������
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        optӦ�÷�Χ(0).Enabled = False
    End If

    '�ٴ�·����Ϣ
    '-------------------------------------------------------------------------------------
    If mlng·��ID = 0 Then
        '�����ٴ�·��
        vsDept.Enabled = optӦ�÷�Χ(1).Value
        cbo����.ListIndex = Cbo.FindIndex(cbo����, mstr����)    '��������Call AddDept
    Else
        vsDept.Enabled = optӦ�÷�Χ(1).Value
        chk����.Visible = False

        '�޸��ٴ�·��
        strSql = "Select ����,����,����,˵��,�����Ա�,��������,ͨ��,�����ʱ�� From ����·��Ŀ¼ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)

        cbo����.Text = NVL(rsTmp!����)
        txt����.Text = rsTmp!����
        txt����.Text = rsTmp!����
        txt˵��.Text = NVL(rsTmp!˵��)
        cbo�����Ա�.ListIndex = Val(NVL(rsTmp!�����Ա�, 0))
        txtMaxDay.Text = NVL(rsTmp!�����ʱ��)

        If Not IsNull(rsTmp!��������) Then
            txt��������(0).Text = Split(rsTmp!��������, "-")(0)
            txt��������(1).Text = Val(Split(rsTmp!��������, "-")(1))
            cbo���䵥λ.ListIndex = Cbo.FindIndex(cbo���䵥λ, CStr(Right(Split(rsTmp!��������, "-")(1), 1)))
        End If

        'Ӧ�ÿ��ҷ�Χ
        optӦ�÷�Χ(0).Value = Val(NVL(rsTmp!ͨ��, 1)) = 1
        optӦ�÷�Χ(1).Value = Val(NVL(rsTmp!ͨ��, 1)) = 2
        If Val(NVL(rsTmp!ͨ��, 1)) = 2 Then
            strSql = "Select B.ID,B.����,B.���� From ����·������ A,���ű� B Where A.����ID=B.ID And A.·��ID=[1] Order by B.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
            If Not rsTmp.EOF Then
                vsDept.Rows = vsDept.FixedRows + rsTmp.RecordCount + 1    '��һ����
                For intIdx = 1 To rsTmp.RecordCount
                    vsDept.RowData(intIdx) = Val(rsTmp!ID)
                    vsDept.TextMatrix(intIdx, 0) = rsTmp!���� & "-" & rsTmp!����
                    vsDept.Cell(flexcpData, intIdx, 0) = vsDept.TextMatrix(intIdx, 0)

                    rsTmp.MoveNext
                Next
            End If
        End If
        vsDept.Row = 0: vsDept.Row = 1: vsDept.Col = 0

        '��Ӧ���ַ�Χ
        strSql = " Select A.����ID,B.���� as ��������,B.���� as ��������," & _
                 " A.���ID,C.���� as ��ϱ���,C.���� as �������" & _
                 " From ����·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
                 " Where A.����ID=B.ID(+) And A.���ID=C.ID(+) And A.·��ID=[1] " & _
                 " Order by B.����,C.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
        If Not rsTmp.EOF Then
            vsDisease.Rows = vsDisease.FixedRows + rsTmp.RecordCount + 1    '��һ����
            For intIdx = 1 To rsTmp.RecordCount
                If Not IsNull(rsTmp!����id) Then
                    vsDisease.RowData(intIdx) = Val(rsTmp!����id & "")
                    vsDisease.TextMatrix(intIdx, 0) = -1
                    vsDisease.TextMatrix(intIdx, 1) = 0
                    vsDisease.TextMatrix(intIdx, 2) = "[" & rsTmp!�������� & "]" & rsTmp!��������
                    vsDisease.ColData(2) = "," & rsTmp!�������� & ","
                Else
                    vsDisease.RowData(intIdx) = Val(rsTmp!���id & "")
                    vsDisease.TextMatrix(intIdx, 1) = -1
                    vsDisease.TextMatrix(intIdx, 0) = 0
                    vsDisease.TextMatrix(intIdx, 2) = "[" & rsTmp!��ϱ��� & "]" & rsTmp!�������
                    vsDisease.ColData(2) = "," & rsTmp!��ϱ��� & ","
                End If
                vsDisease.Cell(flexcpData, intIdx, 2) = vsDisease.TextMatrix(intIdx, 0)
                rsTmp.MoveNext
            Next
            vsDisease.TextMatrix(vsDisease.Rows - 1, 0) = vsDisease.TextMatrix(vsDisease.Rows - 2, 0)
            vsDisease.TextMatrix(vsDisease.Rows - 1, 1) = vsDisease.TextMatrix(vsDisease.Rows - 2, 1)
        End If
        vsDisease.Row = 0: vsDisease.Row = 1: vsDisease.Col = 2
    End If
    vsDisease_AfterRowColChange -1, -1, 1, 2
    mblnChange = False
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mblnOK And mlng·��ID <> 0 And mblnChange Then
        If MsgBox("���ٴ�·������Ϣ�ѱ����ģ�ȷʵҪ���������˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
    
    mstrPrivs = ""
    mlng·��ID = 0
    mstr���� = ""
End Sub

Private Sub optӦ�÷�Χ_Click(Index As Integer)
    vsDept.Enabled = optӦ�÷�Χ(1).Value
    If Visible And vsDept.Enabled Then
        vsDept.SetFocus
    Else
        vsDept.Rows = 1
        vsDept.Rows = 2
    End If
    mblnChange = True
End Sub

Private Sub txtMaxDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub ttxtMaxDay_GotFocus()
    Call zlControl.TxtSelAll(txtMaxDay)
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt����_Change()
    mblnChange = True
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt��������_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txt��������_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt��������(Index))
End Sub

Private Sub txt��������_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_Change()
    mblnChange = True
End Sub

Private Sub txt˵��_GotFocus()
    Call zlControl.TxtSelAll(txt˵��)
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsDept_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDept
        If NewCol <> 2 Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    With vsDept
        If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
            '��ǰ��Ա�����ٴ�����
            strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,������Ա B,��������˵�� C" & _
                    " Where A.ID=B.����ID And B.��ԱID=[1]" & _
                    " And A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
        Else
            'ȫԺ�ٴ�����
            strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,��������˵�� C" & _
                    " Where A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
        End If
        vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�ٴ�����", False, "", "", False, False, True, _
            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, UserInfo.ID)
        
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û���ٴ��������ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetDeptInput(Row, rsTmp)
            Call DeptEnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    With vsDept
        If KeyCode = vbKeyF4 Then
            If .Col = 0 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 0) <> "" Then
                If MsgBox("ȷʵҪ������п�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDept_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    With vsDept
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DeptEnterNextCell
        ElseIf .Col = 0 Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsDept_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsDept_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDept.EditSelStart = 0
    vsDept.EditSelLength = zlCommFun.ActualLen(vsDept.EditText)
End Sub

Private Sub vsDept_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsDept
        If Col = 0 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then
                    Call DeptEnterNextCell
                End If
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then
                    Call DeptEnterNextCell
                End If
            Else
                strInput = UCase(.EditText)
                If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
                    '��ǰ��Ա�����ٴ�����
                    strSql = " Select Distinct A.ID,A.����,A.����,A.����" & _
                             " From ���ű� A,������Ա B,��������˵�� C" & _
                             " Where A.ID=B.����ID And B.��ԱID=[3]" & _
                             " And A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                             " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                             " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                             " Order by A.����"
                Else
                    'ȫԺ�ٴ�����
                    strSql = " Select Distinct A.ID,A.����,A.����,A.����" & _
                             " From ���ű� A,��������˵�� C" & _
                             " Where A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                             " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                             " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                             " Order by A.����"
                End If
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, "�ٴ�����", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", UserInfo.ID)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ����ٴ����ҡ�", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetDeptInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then
                        Call DeptEnterNextCell(True)
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub DeptEnterNextCell(Optional ByVal blnNewRow As Boolean)
    Dim i As Long, j As Long
    
    With vsDept
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 0
            .ShowCell .Row, .Col
        Else
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim i As Long
    Dim intCount As Integer

    With vsDept
        For i = 1 To rsInput.RecordCount
            If .FindRow(Val(rsInput!ID)) = -1 Then
                intCount = intCount + i
                If i > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                End If

                .RowData(lngRow) = Val(rsInput!ID)
                .TextMatrix(lngRow, 0) = rsInput!���� & "-" & rsInput!����
                .Cell(flexcpData, lngRow, 0) = .TextMatrix(lngRow, 0)
            End If
            rsInput.MoveNext
        Next

        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 And intCount > 0 Then
            .AddItem "", lngRow + 1
        End If

        mblnChange = True
    End With
End Sub

Private Sub vsDisease_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Or Col = 1 Then
        With vsDisease
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .TextMatrix(Row, IIf(Col = 1, 0, 1)) = 0
                
                If .RowData(Row) <> 0 Then
                    .RowData(Row) = 0
                    .TextMatrix(Row, 2) = ""
                    .Cell(flexcpData, Row, 2) = ""
                    
                    mblnChange = True
                End If
            End If
        End With
    End If
    
    Call vsDisease_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsDisease_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDisease
        If NewCol <> 2 Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            .ComboList = "..."
        End If
    End With
End Sub

Private Sub vsDisease_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = 1 Then
        If Val(vsDisease.TextMatrix(Row, Col)) <> 0 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsDisease_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
End Sub

Private Sub vsDisease_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset

    With vsDisease
        If Val(.TextMatrix(Row, 1)) <> 0 Then
            '���������:һ����Ͽ������ڶ������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", 0, , True, False, .ColData(2))
        Else
            'D-ICD-10��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D,B", 0, Decode(cbo�����Ա�.ListIndex, 1, "��", 2, "Ů"), True, True, .ColData(2))
        End If
        If Not rsTmp Is Nothing Then
            Call SetDiseaseInput(Row, rsTmp)
            Call DiseaseEnterNextCell(True)
        End If
    End With
End Sub

Private Sub vsDisease_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strTemp As String

    With vsDisease
        If KeyCode = vbKeyF4 Then
            If .Col = 2 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, 2) <> "" Then
                If MsgBox("ȷʵҪ�������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strTemp = .TextMatrix(.Row, 2)
                    strTemp = Mid(strTemp, 2, InStr(strTemp, "]") - 2)
                    .ColData(2) = Replace(.ColData(2), "," & strTemp & ",", "")
                    .RemoveItem .Row
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDisease_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDisease_KeyPress(KeyAscii As Integer)
    With vsDisease
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiseaseEnterNextCell
        Else
            If .Col = 2 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDisease_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDisease_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDisease_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDisease.EditSelStart = 0
    vsDisease.EditSelLength = zlCommFun.ActualLen(vsDisease.EditText)
End Sub

Private Sub vsDisease_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str�Ա� As String, strInput As String
    Dim vPoint As POINTAPI, int������� As Integer
    
    With vsDisease
        If Col = 2 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then
                    Call DiseaseEnterNextCell
                End If
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then
                    Call DiseaseEnterNextCell
                End If
            Else
                strInput = UCase(.EditText)
                If Val(.TextMatrix(Row, 1)) <> 0 Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSql = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
                    Else
                        strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                    End If
                    strSql = " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                             " From �������Ŀ¼ A,������ϱ��� B" & _
                             " Where A.ID=B.���ID And A.���=2" & _
                             " And B.����=[4] And (" & strSql & ")" & _
                             " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " Order by A.����"
                Else
                    If cbo�����Ա�.ListIndex = 1 Then
                        str�Ա� = "��"
                    ElseIf cbo�����Ա�.ListIndex = 2 Then
                        str�Ա� = "Ů"
                    End If
                    'D-ICD-10��������
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSql = "���� Like [2]" '���뺺��ʱ,ֻƥ������
                    Else
                        strSql = "���� Like [1] Or ���� Like [2] Or " & IIf(gint���� = 0, "����", "�����") & " Like [2]"
                    End If
                    strSql = " Select ID,ID as ��ĿID,����,����,����," & IIf(gint���� = 0, "����", "����� as ����") & ",˵��" & _
                             " From ��������Ŀ¼ Where ��� In('D','B') And (" & strSql & ")" & _
                             IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                             " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " Order by ����"
                End If
                
                vPoint = zlControl.GetCoordPos(.Hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 0, IIf(Val(.TextMatrix(Row, 1)) <> 0, "��ϱ���", "��������"), _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", gstrLike & strInput & "%", str�Ա�, gint���� + 1)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call SetDiseaseInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then
                        Call DiseaseEnterNextCell(True)
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub DiseaseEnterNextCell(Optional ByVal blnNewRow As Boolean)
    With vsDisease
        If blnNewRow Then
            .Row = .Rows - 1: .Col = 2
            .ShowCell .Row, .Col
        Else
            If .Col + 1 <= .Cols - 1 Then
                .Col = .Col + 1
            ElseIf .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1: .Col = 2
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub SetDiseaseInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim i As Long
    Dim intCount As Integer

    With vsDisease
        For i = 1 To rsInput.RecordCount
            If .FindRow(Val(rsInput!��ĿID)) = -1 Then
                intCount = intCount + 1    '����Ӽ�¼�����ظ��ļ�¼����
                If intCount > 1 Then
                    .AddItem "", lngRow + 1
                    lngRow = lngRow + 1
                    .TextMatrix(lngRow, 0) = .TextMatrix(lngRow - 1, 0)
                    .TextMatrix(lngRow, 1) = .TextMatrix(lngRow - 1, 1)
                End If
                .RowData(lngRow) = Val(rsInput!��ĿID)
                .TextMatrix(lngRow, 2) = "[" & rsInput!���� & "]" & NVL(rsInput!����)
                .Cell(flexcpData, lngRow, 2) = .TextMatrix(lngRow, 2)
                .ColData(2) = .ColData(2) & "," & rsInput!���� & ","
            End If
            rsInput.MoveNext
        Next

        'ʼ�ձ���һ���У�intCount:��һ����Ӽ�¼��û��ʱ����ֹ��ӿ���
        If lngRow = .Rows - 1 And intCount > 0 Then
            .AddItem "", lngRow + 1
            .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
            .TextMatrix(.Rows - 1, 1) = .TextMatrix(.Rows - 2, 1)
        End If

        mblnChange = True
    End With
End Sub

Private Sub AddDept()
'����:ָ������ʱ������·��·���������ƣ��Զ�����ٴ����Ҳ���
    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim i           As Long

    On Error GoTo errH

    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        '��ȫԺ·��Ȩ��
        '·������Ա���ڶ���ٴ����ҵ�������ȸ��ݷ������ƣ��ӹ���Ա�����ٴ��������ҵ������������ͬ�Ŀ��ң��������
        strSql = " Select Distinct A.ID,A.����,A.����,A.����" & _
                 " From ���ű� A,������Ա B,��������˵�� C" & _
                 " Where A.ID=B.����ID And B.��ԱID=[1]" & _
                 " And A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'  " & _
                 " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                 " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                 " Order by A.����"
    Else
        'ȫԺ·��Ȩ��
        '���ݷ������Ʋ��ң��ҵ��ͼ��أ��Ҳ������Զ����أ�����Ա�ֶ�����
        strSql = " Select Distinct A.ID,A.����,A.����,A.����" & _
                 " From ���ű� A,��������˵�� C" & _
                 " Where A.ID=C.����ID And C.������� IN(1,3) And C.��������='�ٴ�'" & _
                 " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                 " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                 " Order by A.����"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    With vsDept
        rsTmp.Filter = "����='" & cbo����.List(cbo����.ListIndex) & "'"
        For i = 1 To rsTmp.RecordCount
            If .FindRow(rsTmp!ID) = -1 Then    '�Ѿ���ӹ�����ֹ���
                .TextMatrix(i, 0) = rsTmp!���� & "-" & rsTmp!����
                .RowData(i) = Val(rsTmp!ID)
                .Rows = .Rows + 1
            End If
            rsTmp.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
