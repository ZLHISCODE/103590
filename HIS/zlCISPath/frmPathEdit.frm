VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPathEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ٴ�·����Ϣ"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmPathEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   735
      Left            =   4800
      TabIndex        =   38
      Top             =   620
      Width           =   2160
      Begin VB.OptionButton opt��Ҫ·�� 
         Caption         =   "��Ҫ·��"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton opt�ϲ�·�� 
         Caption         =   "�ϲ�·��"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.CheckBox chk��ϲ�ͬ����������� 
      Caption         =   "����·��ʱ�������Ժ��ϲ������ò��ַ�Χ�ڣ�����ѡ��������ɡ�"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   6960
      Width           =   5895
   End
   Begin VB.TextBox txtConfirmDay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   200
      IMEMode         =   2  'OFF
      Left            =   1245
      MaxLength       =   2
      TabIndex        =   35
      Text            =   "0"
      Top             =   6645
      Width           =   300
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������"
      Height          =   195
      Left            =   360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7430
      Width           =   1020
   End
   Begin VB.ComboBox cbo�����Ա� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2540
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Height          =   75
      Left            =   -105
      TabIndex        =   33
      Top             =   7215
      Width           =   7200
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   225
      TabIndex        =   32
      Top             =   1980
      Width           =   6870
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   210
      TabIndex        =   31
      Top             =   540
      Width           =   6885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5760
      TabIndex        =   40
      Top             =   7352
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4620
      TabIndex        =   39
      Top             =   7352
      Width           =   1100
   End
   Begin VB.OptionButton optӦ�÷�Χ 
      Caption         =   "ָ������"
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   22
      Top             =   3360
      Value           =   -1  'True
      Width           =   1020
   End
   Begin VB.OptionButton optӦ�÷�Χ 
      Caption         =   "ȫԺͨ��"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   21
      Top             =   3000
      Width           =   1020
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   840
      Left            =   3690
      TabIndex        =   23
      Top             =   2955
      Width           =   3195
      _cx             =   5636
      _cy             =   1482
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
      FormatString    =   $"frmPathEdit.frx":058A
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
      TabIndex        =   9
      Top             =   1430
      Width           =   5920
   End
   Begin VB.ComboBox cbo���䵥λ 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5805
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2540
      Width           =   720
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   5415
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2540
      Width           =   360
   End
   Begin VB.TextBox txt�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4875
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2540
      Width           =   360
   End
   Begin VB.ComboBox cbo���ò��� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4875
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2160
      Width           =   1650
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2160
      Width           =   1740
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   960
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1050
      Width           =   3855
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3615
      MaxLength       =   5
      TabIndex        =   3
      Top             =   690
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDisease 
      Height          =   2385
      Index           =   0
      Left            =   300
      TabIndex        =   25
      Top             =   4200
      Width           =   3315
      _cx             =   5847
      _cy             =   4207
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
      FormatString    =   $"frmPathEdit.frx":05BE
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
   Begin VSFlex8Ctl.VSFlexGrid vsDisease 
      Height          =   2385
      Index           =   1
      Left            =   3690
      TabIndex        =   27
      Top             =   4200
      Width           =   3210
      _cx             =   5662
      _cy             =   4207
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
      FormatString    =   $"frmPathEdit.frx":0623
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
   Begin VB.Label Lbl��ת���ղ��� 
      AutoSize        =   -1  'True
      Caption         =   "��ת���ղ���(&J)"
      Height          =   180
      Left            =   3720
      TabIndex        =   26
      Top             =   3960
      Width           =   1350
   End
   Begin VB.Label lblConfirm 
      Caption         =   "��Ժʱ�䳬��ȷ���������������ٴ�·����"
      Height          =   255
      Left            =   2040
      TabIndex        =   36
      Top             =   6705
      Width           =   3855
   End
   Begin VB.Label lblConfirmDay 
      Caption         =   "ȷ��������___ ��"
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   6705
      Width           =   1455
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�ٴ�·���Ļ�����Ϣ�����ö���Ӧ�÷�Χ����Ӧ���ֵ�����"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   720
      TabIndex        =   30
      Top             =   195
      Width           =   4860
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   150
      Picture         =   "frmPathEdit.frx":0688
      Top             =   50
      Width           =   480
   End
   Begin VB.Label lbl��Ӧ���� 
      AutoSize        =   -1  'True
      Caption         =   "�������ò���(&I)"
      Height          =   180
      Left            =   285
      TabIndex        =   24
      Top             =   3960
      Width           =   1350
   End
   Begin VB.Label lblӦ�÷�Χ 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ�÷�Χ(&S)"
      Height          =   180
      Left            =   285
      TabIndex        =   20
      Top             =   2985
      Width           =   990
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      Caption         =   "˵��(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   8
      Top             =   1480
      Width           =   630
   End
   Begin VB.Label lbl���䷶Χ 
      AutoSize        =   -1  'True
      Caption         =   "-"
      Height          =   180
      Left            =   5280
      TabIndex        =   29
      Top             =   2595
      Width           =   90
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "��������(&Y)"
      Height          =   180
      Left            =   3735
      TabIndex        =   16
      Top             =   2595
      Width           =   990
   End
   Begin VB.Label lbl�����Ա� 
      AutoSize        =   -1  'True
      Caption         =   "�����Ա�(&X)"
      Height          =   180
      Left            =   285
      TabIndex        =   14
      Top             =   2600
      Width           =   990
   End
   Begin VB.Label lbl���ò��� 
      AutoSize        =   -1  'True
      Caption         =   "���ò���(&B)"
      Height          =   180
      Left            =   3735
      TabIndex        =   12
      Top             =   2220
      Width           =   990
   End
   Begin VB.Label lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "��������(&T)"
      Height          =   180
      Left            =   285
      TabIndex        =   10
      Top             =   2220
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
      Left            =   2880
      TabIndex        =   2
      Top             =   750
      Width           =   630
   End
End
Attribute VB_Name = "frmPathEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event AfterSave(ByVal ���� As String, ByVal ���� As String)

Private mstrPrivs As String
Private mlng·��ID As Long
Private mstr���� As String
Private mblnReturn As Boolean
Private mblnChange As Boolean
Private mblnOK As Boolean

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

Private Sub cbo��������_Click()
    mblnChange = True
End Sub

Private Sub cbo����_Change()
    mblnChange = True
End Sub

Private Sub cbo����_Click()
    If mlng·��ID = 0 Then
        txt����.Text = GetNextCode(cbo����.Text)
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
        txt����.Text = GetNextCode(cbo����.Text)
    End If
End Sub

Private Sub cbo���䵥λ_Click()
    mblnChange = True
End Sub

Private Sub cbo���ò���_Click()
    mblnChange = True
End Sub

Private Sub cbo�����Ա�_Click()
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim str����IDs As String
    Dim str����IDs As String
    Dim strSql As String, i As Long
    Dim strTmp As String, intLimit As Integer
    Dim str��ת����IDs As String

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
        cbo����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt����.Text) > txt����.MaxLength Then
        MsgBox "�ٴ�·�����������ֻ���� " & txt����.MaxLength \ 2 & " �����ֻ� " & txt����.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt˵��.Text) > txt˵��.MaxLength Then
        MsgBox "�ٴ�·����˵����Ϣ���ֻ���� " & txt˵��.MaxLength \ 2 & " �����ֻ� " & txt˵��.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txt˵��.SetFocus: Exit Sub
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
                        .SetFocus: Exit Sub
                    Else
                        strTmp = strTmp & "," & .RowData(i)
                    End If
                End If
            Next
        End With
    End If
    With vsDisease(0)
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
                    .SetFocus: Exit Sub
                Else
                    strTmp = strTmp & "," & strSql
                End If
            End If
        Next
    End With
    If vsDisease(1).Enabled Then
        With vsDisease(1)
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
                        .SetFocus: Exit Sub
                    Else
                        strTmp = strTmp & "," & strSql
                    End If
                End If
            Next
        End With
    End If

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
            vsDept.SetFocus: Exit Sub
        End If
    End If

    For i = 1 To vsDisease(0).Rows - 1
        If vsDisease(0).RowData(i) <> 0 And Val(vsDisease(0).TextMatrix(i, 0)) <> 0 Then
            str����IDs = str����IDs & "," & vsDisease(0).RowData(i)
        End If
    Next
    str����IDs = Mid(str����IDs, 2) & ";"

    For i = 1 To vsDisease(0).Rows - 1
        If vsDisease(0).RowData(i) <> 0 And Val(vsDisease(0).TextMatrix(i, 1)) <> 0 Then
            str����IDs = str����IDs & vsDisease(0).RowData(i) & ","
        End If
    Next

    If vsDisease(1).Enabled Then
        For i = 1 To vsDisease(1).Rows - 1
            If vsDisease(1).RowData(i) <> 0 And Val(vsDisease(1).TextMatrix(i, 0)) <> 0 Then
                str��ת����IDs = str��ת����IDs & "," & vsDisease(1).RowData(i)
            End If
        Next

        str��ת����IDs = Mid(str��ת����IDs, 2) & ";"

        For i = 1 To vsDisease(1).Rows - 1
            If vsDisease(1).RowData(i) <> 0 And Val(vsDisease(1).TextMatrix(i, 1)) <> 0 Then
                str��ת����IDs = str��ת����IDs & vsDisease(1).RowData(i) & ","
            End If
        Next
    Else
        str��ת����IDs = ";"
    End If

    If str����IDs = ";" Then
        MsgBox "����ָ���ٴ�·������Ӧ�ĵĲ��֡�", vbInformation, gstrSysName
        vsDisease(0).SetFocus: Exit Sub
    End If
    If str��ת����IDs = ";" Then
        str��ת����IDs = ""
    End If
    'ȥ�����ұߵĶ���
    If Right(str����IDs, 1) = "," Then str����IDs = Left(str����IDs, Len(str����IDs) - 1)
    If Right(str��ת����IDs, 1) = "," Then str��ת����IDs = Left(str��ת����IDs, Len(str��ת����IDs) - 1)
    If mlng·��ID = 0 Then
        strSql = "Zl_�ٴ�·��Ŀ¼_Insert('" & cbo����.Text & "','" & txt����.Text & "','" & txt����.Text & "'," & _
                 "'" & txt˵��.Text & "','" & zlCommFun.GetNeedName(cbo��������.Text) & "','" & IIf(cbo���ò���.ListIndex = 0, "", zlCommFun.GetNeedName(cbo���ò���.Text)) & "'," & _
                 cbo�����Ա�.ListIndex & ",'" & IIf(txt��������(0).Text <> "", txt��������(0).Text & "-" & txt��������(1).Text & cbo���䵥λ.Text, "") & "'," & _
                 IIf(optӦ�÷�Χ(0).Value, 1, 2) & ",'" & str����IDs & "','" & str����IDs & "',Null," & IIf(txtConfirmDay.Enabled, ZVal(Val(txtConfirmDay.Text)), "Null") & ",'" & _
                 str��ת����IDs & "'," & ZVal(chk��ϲ�ͬ�����������.Value) & "," & IIf(opt�ϲ�·��.Value, "1", "0") & ")"
    Else
        strSql = "Zl_�ٴ�·��Ŀ¼_Update(" & mlng·��ID & ",'" & cbo����.Text & "','" & txt����.Text & "','" & txt����.Text & "'," & _
                 "'" & txt˵��.Text & "','" & zlCommFun.GetNeedName(cbo��������.Text) & "','" & IIf(cbo���ò���.ListIndex = 0, "", zlCommFun.GetNeedName(cbo���ò���.Text)) & "'," & _
                 cbo�����Ա�.ListIndex & ",'" & IIf(txt��������(0).Text <> "", txt��������(0).Text & "-" & txt��������(1).Text & cbo���䵥λ.Text, "") & "'," & _
                 IIf(optӦ�÷�Χ(0).Value, 1, 2) & ",'" & str����IDs & "','" & str����IDs & "'," & IIf(txtConfirmDay.Enabled, ZVal(Val(txtConfirmDay.Text)), "Null") & ",'" & _
                 str��ת����IDs & "'," & ZVal(chk��ϲ�ͬ�����������.Value) & "," & IIf(opt�ϲ�·��.Value, "1", "0") & ")"
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
            strSql = "Select Nvl(Count(*),0) as ���� From �ٴ�·��Ŀ¼"
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
            If rsTmp!���� < intLimit Then intLimit = 0
            On Error GoTo 0
        End If
        If intLimit = 0 Then
            txt����.Text = GetNextCode(cbo����.Text)
            txt����.Text = "": txt˵��.Text = ""
            txtConfirmDay.Text = "0"
            mblnChange = False: txt����.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
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

    '�ֵ���Ϣ��ȡ
    '-------------------------------------------------------------------------------------
    '������Ϣ
    strSql = "Select Distinct ���� From �ٴ�·��Ŀ¼ Where ���� is Not NULL Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!����
        rsTmp.MoveNext
    Loop

    '��������
    cbo��������.AddItem ""
    cbo��������.ListIndex = 0
    strSql = "Select ����,����,���� From �ٴ��������� Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo��������.AddItem rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop

    '����
    cbo���ò���.AddItem "0-�����ֲ���"
    cbo���ò���.ListIndex = 0
    strSql = "Select ����,����,���� From ���� Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
    Do While Not rsTmp.EOF
        cbo���ò���.AddItem rsTmp!���� & "-" & rsTmp!����
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
        strSql = "Select ����,����,����,˵��,��������,���ò���,�����Ա�,��������,ͨ��,ȷ������,����·������,���� From �ٴ�·��Ŀ¼ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)

        cbo����.Text = Nvl(rsTmp!����)
        txt����.Text = rsTmp!����
        txt����.Text = rsTmp!����
        txt˵��.Text = Nvl(rsTmp!˵��)
        txtConfirmDay.Text = Val("" & rsTmp!ȷ������)
        chk��ϲ�ͬ�����������.Value = Val(rsTmp!����·������ & "")

        If Not IsNull(rsTmp!��������) Then
            cbo��������.ListIndex = Cbo.FindIndex(cbo��������, CStr(rsTmp!��������))
        End If

        If Not IsNull(rsTmp!���ò���) Then
            cbo���ò���.ListIndex = Cbo.FindIndex(cbo���ò���, CStr(rsTmp!���ò���))
        End If

        cbo�����Ա�.ListIndex = Val(Nvl(rsTmp!�����Ա�, 0))

        If Not IsNull(rsTmp!��������) Then
            txt��������(0).Text = Split(rsTmp!��������, "-")(0)
            txt��������(1).Text = Val(Split(rsTmp!��������, "-")(1))
            cbo���䵥λ.ListIndex = Cbo.FindIndex(cbo���䵥λ, CStr(Right(Split(rsTmp!��������, "-")(1), 1)))
        End If

        If Val(rsTmp!���� & "") = 1 Then
            opt�ϲ�·��.Value = True
        Else
            opt��Ҫ·��.Value = True
        End If

        'Ӧ�ÿ��ҷ�Χ
        optӦ�÷�Χ(0).Value = Val(Nvl(rsTmp!ͨ��, 1)) = 1
        optӦ�÷�Χ(1).Value = Val(Nvl(rsTmp!ͨ��, 1)) = 2
        If Val(Nvl(rsTmp!ͨ��, 1)) = 2 Then
            strSql = "Select B.ID,B.����,B.���� From �ٴ�·������ A,���ű� B Where A.����ID=B.ID And A.·��ID=[1] Order by B.����"
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
        strSql = _
        " Select" & _
                 " A.����ID,B.���� as ��������,B.���� as ��������," & _
                 " A.���ID,C.���� as ��ϱ���,C.���� as �������,Nvl(a.����,0) as ����" & _
                 " From �ٴ�·������ A,��������Ŀ¼ B,�������Ŀ¼ C" & _
                 " Where A.����ID=B.ID(+) And A.���ID=C.ID(+) And A.·��ID=[1] " & _
                 " Order by B.����,C.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
        If Not rsTmp.EOF Then
            rsTmp.Filter = "����=0"
            vsDisease(0).Rows = vsDisease(0).FixedRows + rsTmp.RecordCount + 1    '��һ����
            For intIdx = 1 To rsTmp.RecordCount
                If Not IsNull(rsTmp!����id) Then
                    vsDisease(0).RowData(intIdx) = Val(rsTmp!����id & "")
                    vsDisease(0).TextMatrix(intIdx, 0) = -1
                    vsDisease(0).TextMatrix(intIdx, 1) = 0
                    vsDisease(0).TextMatrix(intIdx, 2) = "[" & rsTmp!�������� & "]" & rsTmp!��������
                    vsDisease(0).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!�������� & ","
                Else
                    vsDisease(0).RowData(intIdx) = Val(rsTmp!���id & "")
                    vsDisease(0).TextMatrix(intIdx, 1) = -1
                    vsDisease(0).TextMatrix(intIdx, 0) = 0
                    vsDisease(0).TextMatrix(intIdx, 2) = "[" & rsTmp!��ϱ��� & "]" & rsTmp!�������
                    vsDisease(0).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!��ϱ��� & ","
                End If
                vsDisease(0).Cell(flexcpData, intIdx, 2) = vsDisease(0).TextMatrix(intIdx, 0)

                rsTmp.MoveNext
            Next
            vsDisease(0).TextMatrix(vsDisease(0).Rows - 1, 0) = vsDisease(0).TextMatrix(vsDisease(0).Rows - 2, 0)
            vsDisease(0).TextMatrix(vsDisease(0).Rows - 1, 1) = vsDisease(0).TextMatrix(vsDisease(0).Rows - 2, 1)

            rsTmp.Filter = "����=1"
            vsDisease(1).Rows = vsDisease(1).FixedRows + rsTmp.RecordCount + 1    '��һ����
            For intIdx = 1 To rsTmp.RecordCount
                If Not IsNull(rsTmp!����id) Then
                    vsDisease(1).RowData(intIdx) = Val(rsTmp!����id & "")
                    vsDisease(1).TextMatrix(intIdx, 0) = -1
                    vsDisease(1).TextMatrix(intIdx, 1) = 0
                    vsDisease(1).TextMatrix(intIdx, 2) = "[" & rsTmp!�������� & "]" & rsTmp!��������
                    vsDisease(1).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!�������� & ","
                Else
                    vsDisease(1).RowData(intIdx) = Val(rsTmp!���id & "")
                    vsDisease(1).TextMatrix(intIdx, 1) = -1
                    vsDisease(1).TextMatrix(intIdx, 0) = 0
                    vsDisease(1).TextMatrix(intIdx, 2) = "[" & rsTmp!��ϱ��� & "]" & rsTmp!�������
                    vsDisease(1).ColData(2) = vsDisease(1).ColData(2) & "," & rsTmp!��ϱ��� & ","
                End If
                vsDisease(1).Cell(flexcpData, intIdx, 0) = vsDisease(1).TextMatrix(intIdx, 0)

                rsTmp.MoveNext
            Next
            If rsTmp.RecordCount = 0 Then
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 0) = -1
            Else
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 0) = vsDisease(1).TextMatrix(vsDisease(1).Rows - 2, 0)
                vsDisease(1).TextMatrix(vsDisease(1).Rows - 1, 1) = vsDisease(1).TextMatrix(vsDisease(1).Rows - 2, 1)
            End If
        End If
        vsDisease(0).Row = 0: vsDisease(0).Row = 1: vsDisease(0).Col = 2
        vsDisease(1).Row = 0: vsDisease(1).Row = 1: vsDisease(1).Col = 2
    End If
    vsDisease_AfterRowColChange 0, -1, -1, 1, 2
    vsDisease_AfterRowColChange 1, -1, -1, 1, 2
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

Private Sub opt�ϲ�·��_Click()
    Call SetFirstPath(False)
End Sub

Private Sub opt��Ҫ·��_Click()
    Call SetFirstPath(True)
End Sub

Private Sub SetFirstPath(ByVal blnVisible As Boolean)
'���ܣ����ѡ����Ҫ·����ϲ�·���������ý���Ŀؼ�
'������blnVisible=true ����Ϊ��Ҫ·��������Ϊ�ϲ�·��
    Lbl��ת���ղ���.Enabled = blnVisible
    vsDisease(1).Enabled = blnVisible
    vsDisease(1).BackColor = IIf(blnVisible, vbWindowBackground, vbButtonFace)
    vsDisease(1).BackColorBkg = IIf(blnVisible, vbWindowBackground, vbButtonFace)
    lblConfirmDay.Enabled = blnVisible
    txtConfirmDay.Enabled = blnVisible
    txtConfirmDay.BackColor = IIf(blnVisible, &HC0E0FF, Me.BackColor)
    lblConfirm.Enabled = blnVisible
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

Private Sub txtConfirmDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtConfirmDay_GotFocus()
    Call zlControl.TxtSelAll(txtConfirmDay)
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
                " And A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
        Else
            'ȫԺ�ٴ�����
            strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                " From ���ű� A,��������˵�� C" & _
                " Where A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
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
                If mblnReturn Then Call DeptEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DeptEnterNextCell
            Else
                strInput = UCase(.EditText)
                If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
                    '��ǰ��Ա�����ٴ�����
                    strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,������Ա B,��������˵�� C" & _
                        " Where A.ID=B.����ID And B.��ԱID=[3]" & _
                        " And A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
                        " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                        " Order by A.����"
                Else
                    'ȫԺ�ٴ�����
                    strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                        " From ���ű� A,��������˵�� C" & _
                        " Where A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
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
                    If mblnReturn Then Call DeptEnterNextCell(True)
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
'���ܣ��������ҩ�������
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

Private Sub vsDisease_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Or Col = 1 Then
        With vsDisease(Index)
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
    
    Call vsDisease_AfterRowColChange(Index, -1, -1, Row, Col)
End Sub

Private Sub vsDisease_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsDisease(Index)
        If NewCol <> 2 Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            .ComboList = "..."
        End If
    End With
End Sub

Private Sub vsDisease_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Or Col = 1 Then
        If Val(vsDisease(Index).TextMatrix(Row, Col)) <> 0 Then Cancel = True
    End If
End Sub

Private Sub vsDisease_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 2 Then Cancel = True
End Sub

Private Sub vsDisease_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset

    With vsDisease(Index)

        If Val(.TextMatrix(Row, 1)) <> 0 Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", 0, , True, False, .ColData(2))
        Else
            'D-ICD-10��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D,B", 0, Decode(cbo�����Ա�.ListIndex, 1, "��", 2, "Ů"), True, True, .ColData(2))
        End If
        If Not rsTmp Is Nothing Then
            Call SetDiseaseInput(Index, Row, rsTmp)
            Call DiseaseEnterNextCell(Index, True)
        End If
    End With
End Sub

Private Sub vsDisease_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim strTemp As String

    With vsDisease(Index)
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
            Call vsDisease_KeyPress(Index, KeyCode)
        End If
    End With
End Sub

Private Sub vsDisease_KeyPress(Index As Integer, KeyAscii As Integer)
    With vsDisease(Index)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call DiseaseEnterNextCell(Index)
        Else
            If .Col = 2 Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDisease_CellButtonClick(Index, .Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDisease_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDisease_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDisease(Index).EditSelStart = 0
    vsDisease(Index).EditSelLength = zlCommFun.ActualLen(vsDisease(Index).EditText)
End Sub

Private Sub vsDisease_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim str�Ա� As String, strInput As String
    Dim vPoint As POINTAPI, int������� As Integer
    
    With vsDisease(Index)
        If Col = 2 Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call DiseaseEnterNextCell(Index)
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call DiseaseEnterNextCell(Index)
            Else
                strInput = UCase(.EditText)
                If Val(.TextMatrix(Row, 1)) <> 0 Then
                    '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                    If zlCommFun.IsCharChinese(strInput) Then
                        strSql = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
                    Else
                        strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                    End If
                    strSql = _
                        " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                        " From �������Ŀ¼ A,������ϱ��� B" & _
                        " Where A.ID=B.���ID And A.���=1" & _
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
                    strSql = _
                        " Select ID,ID as ��ĿID,����,����,����," & IIf(gint���� = 0, "����", "����� as ����") & ",˵��" & _
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
                    Call SetDiseaseInput(Index, Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call DiseaseEnterNextCell(Index, True)
                End If
            End If
            mblnReturn = False
        End If
    End With
End Sub

Private Sub DiseaseEnterNextCell(Index As Integer, Optional ByVal blnNewRow As Boolean)
    With vsDisease(Index)
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

Private Sub SetDiseaseInput(Index As Integer, ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim i As Long
    Dim intCount As Integer

    With vsDisease(Index)
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
                .TextMatrix(lngRow, 2) = "[" & rsInput!���� & "]" & Nvl(rsInput!����)
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
        strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                 " From ���ű� A,������Ա B,��������˵�� C" & _
                 " Where A.ID=B.����ID And B.��ԱID=[1]" & _
                 " And A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'  " & _
                 " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                 " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                 " Order by A.����"
    Else
        'ȫԺ·��Ȩ��
        '���ݷ������Ʋ��ң��ҵ��ͼ��أ��Ҳ������Զ����أ�����Ա�ֶ�����
        strSql = "Select Distinct A.ID,A.����,A.����,A.����" & _
                 " From ���ű� A,��������˵�� C" & _
                 " Where A.ID=C.����ID And C.������� IN(2,3) And C.��������='�ٴ�'" & _
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
