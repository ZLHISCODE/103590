VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExamineEdit 
   BackColor       =   &H00808080&
   Caption         =   "���˷��������༭"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   10665
   Icon            =   "frmExamineEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10665
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsExist 
      Height          =   2370
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      _cx             =   12462
      _cy             =   4180
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
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExamineEdit.frx":1601A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VB.CommandButton cmd���� 
      Caption         =   "����ģ��(&T)"
      Height          =   350
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����Ϊģ��(&S)"
      Height          =   350
      Left            =   8040
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   609
      TabWidthStyle   =   1
      TabFixedWidth   =   2290
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ȫ��(&0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�г�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�в�ҩ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picLineX 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8415
      TabIndex        =   7
      Top             =   2760
      Width           =   8415
   End
   Begin VB.PictureBox picLineY 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   4440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   2880
      Width           =   45
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   2550
      ScaleHeight     =   405
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   2955
      Width           =   5775
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         Height          =   350
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   350
         Left            =   1125
         TabIndex        =   4
         Top             =   30
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2775
         TabIndex        =   3
         Top             =   55
         Width           =   2955
      End
      Begin VB.Image ImgFind 
         Height          =   240
         Left            =   2475
         Picture         =   "frmExamineEdit.frx":1615C
         Top             =   85
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2865
      Left            =   2520
      TabIndex        =   8
      Top             =   3375
      Width           =   7065
      _cx             =   12462
      _cy             =   5054
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
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExamineEdit.frx":164E6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3285
      Left            =   0
      TabIndex        =   9
      Top             =   2955
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   5794
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":165BF
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16719
            Key             =   "Exp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16873
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16CC5
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":17117
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":1756F
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":179C3
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":17E17
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":1826B
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":190BD
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6975
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   635
      SimpleText      =   $"frmExamineEdit.frx":19F0F
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExamineEdit.frx":19F56
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13732
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
            Enabled         =   0   'False
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
   Begin MSComctlLib.ListView lvwģ�� 
      Height          =   2895
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������Ŀ����"
      Height          =   180
      Left            =   0
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl��Ŀ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ģ����Ŀ�б�"
      Height          =   180
      Left            =   0
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblģ�� 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ŀģ���б�"
      Height          =   180
      Left            =   0
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frmExamineEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ColAdded
    ѡ�� = 0: ʹ������: ���: ����: ����: ���: ����: ��λ: ˵��: ������: ����ʱ��: ID
End Enum

Private Enum ColAdd
    ѡ�� = 0: ���: ����: ����: ���: ����: ��λ: ˵��: ID
End Enum

Private mrsExistItem As New ADODB.Recordset
Private mint���� As Integer
Private mlng����ID As Long, mlng��ҳID As Long, mlng���� As Long
Private mlngCount As Long
Private mblnDel As Boolean
Private mblnDelPrv As Boolean
Private mblnģ�� As Boolean

Private Function Exist����(lng���� As Long, str���� As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select 1 from ������Ŀģ�� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����)
    If Not rsTemp.EOF Then Exist���� = True: Exit Function
    
    strSQL = "select 1 from ������Ŀģ�� where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
    If Not rsTemp.EOF Then Exist���� = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReadExistsTemplet()
'��ȡ�Ѵ��ڵ�ģ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lvwItem As ListItem
    
On Error GoTo errHandle
    
    If mblnģ�� Then
        strSQL = "Select distinct(����) ����,���� From ������Ŀģ�� Order By ����"
    Else
        strSQL = "Select distinct(����) ����,���� From ������Ŀģ�� A,����֧����Ŀ B Where A.��ĿID = B.�շ�ϸĿID And B.Ҫ������ = 1 And B.���� = [1] Order By ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����)
    lvwģ��.ListItems.Clear
    
    If mblnģ�� = True Then
        Set lvwItem = lvwģ��.ListItems.Add(, "_ADD", "����������Ŀģ��", "Write", "Write")
    End If
    '����30028 by lesfeng 2010-06-01 �������Ϊ�����
    While Not rsTemp.EOF
        Set lvwItem = lvwģ��.ListItems.Add(, "_" & rsTemp!����, IIf(IsNull(rsTemp!����), "", rsTemp!����), "Item", "Item")
        rsTemp.MoveNext
    Wend
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub View״̬()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mblnģ�� = False Then
        strSQL = "Select סԺ��, ���� From ������Ϣ Where ����id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        
        stbThis.Panels(2).Text = "סԺ��:" & rsTemp!סԺ�� & " ����:" & rsTemp!���� & "   �ò����ܹ��趨��" & mlngCount & "��������¼��"
    Else
        If lvwģ��.SelectedItem Is Nothing Then
            stbThis.Panels(2).Text = "����ģ����ѡ��ģ���б��е�[���������Ŀģ��],ѡ�������Ŀ����[����]"
        ElseIf UCase(Mid(lvwģ��.SelectedItem.Key, 2)) = "ADD" Then
            stbThis.Panels(2).Text = "����ģ��,ѡ�������Ŀ����[����]"
        Else
            stbThis.Panels(2).Text = "ģ��:" & lvwģ��.SelectedItem.Text & "  ��ģ���ܹ��趨��" & mlngCount & "��������¼��"
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
    Dim lngRow As Long, lngCount As Long
    Dim strTemp As String
    Dim blnTemp As Boolean, blnTrans As Boolean
    Dim lng���� As Long
    Dim str���� As String
        
    If mblnDel = True Then Unload Me: Exit Sub
    
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.ѡ��) = flexChecked Then
            blnTemp = True
        End If
    Next lngRow
    
    If blnTemp = False Then
        MsgBox "��ѡ��Ҫ���ӵ�������Ŀ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnģ�� = True Then
        If UCase(Mid(lvwģ��.SelectedItem.Key, 2)) = "ADD" Then
            str���� = frmTempletEdit.EditTemplet(Me)
            If str���� = "" Then Exit Sub
            lng���� = Mid(str����, 1, InStr(str����, ",") - 1)
            str���� = Mid(str����, InStr(str����, ",") + 1)
            If Exist����(lng����, str����) = True Then
                MsgBox "�ñ����Ѿ����ڲ�������!", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            lng���� = Mid(lvwģ��.SelectedItem.Key, 2)
            str���� = lvwģ��.SelectedItem.Text
        End If
    End If
    
On Error GoTo errHandle
    gcnOracle.BeginTrans: blnTrans = True
    
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.ѡ��) = flexChecked Then
            strTemp = vsList.TextMatrix(lngRow, ColAdd.ID) & "," & strTemp
            lngCount = lngCount + 1
        End If
        If lngCount = 100 Then
            If mblnģ�� = False Then
                gstrSQL = "Zl_����������Ŀ_Insert(" & mlng����ID & "," & mlng��ҳID & ",'" & strTemp & "','" & gstrUserName & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                gstrSQL = "zl_������Ŀģ��_Insert(" & lng���� & ",'" & str���� & "','" & strTemp & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            lngCount = 0
            strTemp = ""
        End If
    Next lngRow
    
    If strTemp <> "" Then
        If mblnģ�� = False Then
            gstrSQL = "Zl_����������Ŀ_Insert(" & mlng����ID & "," & mlng��ҳID & ",'" & strTemp & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            gstrSQL = "zl_������Ŀģ��_Insert(" & lng���� & ",'" & str���� & "','" & strTemp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    vsList.Redraw = flexRDNone
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.ѡ��) = flexChecked Then
            vsList.Cell(flexcpChecked, lngRow, ColAdd.ѡ��) = 2
            vsList.RowHidden(lngRow) = True
        End If
    Next lngRow
    
    vsList.Redraw = flexRDDirect
    If mblnģ�� = False Then
        Call ReadExistsItem(mlng����ID, mlng��ҳID)
    Else
        If UCase(Mid(lvwģ��.SelectedItem.Key, 2)) = "ADD" Then
            lvwģ��.ListItems.Add , "_" & lng����, str����, "Item", "Item"
            Set lvwģ��.SelectedItem = lvwģ��.ListItems.Item("_" & lng����)
            Call ReadTempletItem(lng����)
        Else
            Call ReadTempletItem(lng����)
        End If
    End If
    
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    Dim strTemp As String
    Dim lngRow As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnTemp As Boolean, blnTrans As Boolean
    Dim lng���� As Long
    Dim i As Integer
    Dim strMsg As String
    
    '�ж��Ƿ�ѡ����Ҫɾ������Ŀ
    For lngRow = 1 To vsExist.Rows - 1
      If vsExist.Cell(flexcpChecked, lngRow, ColAdded.ѡ��) = 1 Then
        blnTemp = True
        Exit For
      End If
    Next lngRow
    
    If blnTemp = True Then
        If MsgBox("ȷ��ɾ����������Ŀ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        MsgBox "��ѡ��Ҫɾ������Ŀ!", vbInformation, Me.Caption
        vsExist.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    
'����ɾ��״̬
    If mblnģ�� = False Then
        
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.ѡ��) = 1 Then
                strTemp = strTemp & "," & vsExist.TextMatrix(vsExist.Row, ColAdded.ID)
            End If
        Next lngRow
        'by lesfeng 2009-12-30 �����  ���˷��ü�¼ --��סԺ���ü�¼ ����ֻ��סԺ
        'by lesfeng 2010-03-06 ���ܰ�
        strSQL = "Select A.�շ�ϸĿid, B.���� " & _
                 "From סԺ���ü�¼ A, �շ���ĿĿ¼ B " & _
                 "Where A.�շ�ϸĿid = B.ID And InStr([1], ',' || A.�շ�ϸĿid || ',') > 0" & _
                    " And A.����id = [2] And A.��ҳid = [3]" & _
                    " And (B.վ��=[4] Or B.վ�� is Null)"
    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "," & strTemp & ",", mlng����ID, mlng��ҳID, gstrNodeNo)
        
        If rsTemp.RecordCount > 0 Then
           Do While Not rsTemp.EOF
              If i > 5 Then strMsg = strMsg & "...": Exit Do
              strMsg = strMsg & ",[" & rsTemp!���� & "]"
              i = i + 1
           Loop
           strMsg = Mid(strMsg, 2)
           If MsgBox("�ò����Ѿ�����" & strMsg & Chr(13) & Chr(10) & "�ķ�����Ϣ,�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.ѡ��) = 1 Then
                gstrSQL = "Zl_����������Ŀ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & vsExist.TextMatrix(lngRow, ColAdded.ID) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next lngRow
        gcnOracle.CommitTrans: blnTrans = False
    Else
        '����ģ��״̬
        lng���� = Mid(lvwģ��.SelectedItem.Key, 2)
        
        gcnOracle.BeginTrans: blnTrans = True
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.ѡ��) = 1 Then
                gstrSQL = "ZL_������Ŀģ��_DELETE(" & lng���� & "," & vsExist.TextMatrix(lngRow, ColAdded.ID) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next lngRow
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    If mblnģ�� = False Then
        Call ReadExistsItem(mlng����ID, mlng��ҳID)
    Else
        Call ReadTempletItem(lng����)
        If mrsExistItem.RecordCount = 0 Then
            lvwģ��.ListItems.Remove lvwģ��.SelectedItem.Key
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long, lngCount As Long, blnTrans As Boolean
    Dim strTemp As String
    Dim lng���� As Long
    Dim str���� As String
    Dim strTabKey As String

    
    If vsExist.Rows < 1 Then
        MsgBox "û�б���Ϊģ���������Ŀ!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    str���� = frmTempletEdit.EditTemplet(Me)
    If str���� = "" Then Exit Sub
    lng���� = Mid(str����, 1, InStr(str����, ",") - 1)
    str���� = Mid(str����, InStr(str����, ",") + 1)
    If Exist����(lng����, str����) = True Then
        MsgBox "�ñ����Ѿ����ڲ�������!", vbInformation, gstrSysName
        Exit Sub
    End If
    
On Error GoTo errHandle
    gcnOracle.BeginTrans: blnTrans = True
    strTabKey = tabClass.SelectedItem.Key
    mrsExistItem.Filter = 0
    
    While Not mrsExistItem.EOF
        
        strTemp = mrsExistItem!ID & "," & strTemp
        lngCount = lngCount + 1
        
        If lngCount = 100 Then
            gstrSQL = "zl_������Ŀģ��_Insert(" & lng���� & ",'" & str���� & "','" & strTemp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            lngCount = 0
            strTemp = ""
        End If
        
        mrsExistItem.MoveNext
    Wend
    
    If strTemp <> "" Then
        gstrSQL = "zl_������Ŀģ��_Insert(" & lng���� & ",'" & str���� & "','" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '����30020 by lesfeng 2010-06-01 �������Ϊ�����
    If Not mrsExistItem.EOF Then
        If str���� <> "" Then
            lvwģ��.ListItems.Add , "_" & lng����, str����, "Item", "Item"
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    Call tabClass_Click
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    If lvwģ��.ListItems.Count = 0 Then
        MsgBox "û�����������ڵ�ǰ���������ģ��,��������ģ��!", vbInformation, gstrSysName
        Exit Sub
    End If
    lvwģ��.Visible = True
    lvwģ��.Move cmd����.Left + cmd����.Width - lvwģ��.Width, cmd����.Top + cmd����.Height, lvwģ��.Width, 3000
    lvwģ��.Height = lvwģ��.ListItems.Item(1).Height * (lvwģ��.ListItems.Count + 1)
    lvwģ��.ZOrder
    lvwģ��.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If mblnDel = False Then
            If Me.ActiveControl.Name = "vsList" Then
                If vsList.Rows > 1 Then
                    vsList.Editable = flexEDKbdMouse
                    If KeyCode = vbKeyR Then
                        vsList.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsList.Rows - 1, ColAdd.ѡ��) = 2
                    ElseIf KeyCode = vbKeyA Then
                        vsList.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsList.Rows - 1, ColAdd.ѡ��) = 1
                    End If
                    vsList.Editable = flexEDNone
                End If
            ElseIf Me.ActiveControl.Name = "vsExist" Then
                If vsExist.Rows > 1 Then
                    vsExist.Editable = flexEDKbdMouse
                    If KeyCode = vbKeyR Then
                        vsExist.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsExist.Rows - 1, ColAdd.ѡ��) = 2
                    ElseIf KeyCode = vbKeyA Then
                        vsExist.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsExist.Rows - 1, ColAdd.ѡ��) = 1
                    End If
                    vsExist.Editable = flexEDNone
                End If
                
            End If
        Else
            If vsExist.Rows > 1 Then
                vsExist.Editable = flexEDKbdMouse
                If KeyCode = vbKeyR Then
                    vsExist.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsExist.Rows - 1, ColAdd.ѡ��) = 2
                ElseIf KeyCode = vbKeyA Then
                    vsExist.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsExist.Rows - 1, ColAdd.ѡ��) = 1
                End If
                vsExist.Editable = flexEDNone
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picLineX.Width = Me.Width
    If mblnģ�� = False Then
        vsExist.Move 0, tabClass.Height - 20, Me.Width - 100, picLineX.Top - tabClass.Height + 50
        tabClass.Left = 0
        tabClass.Top = 30
        If mblnDel = False Then
            tvwMain_S.Move 0, picLineX.Top + picLineX.Height, picLineY.Left, Me.ScaleHeight - picLineX.Top - picLineX.Height - stbThis.Height
            picLineY.Top = tvwMain_S.Top
            picLineY.Height = tvwMain_S.Height
            pic.Move picLineY.Left + picLineY.Width, tvwMain_S.Top, Me.Width - picLineY.Left - picLineY.Width, pic.Height
            vsList.Move pic.Left, tvwMain_S.Top + pic.Height + 20, pic.Width - 100, tvwMain_S.Height - pic.Height - 20
            tvwMain_S.Visible = True
            txtFind.Width = Me.Width - pic.Left - txtFind.Left - 200
            If cmdSave.Visible = True Then
                cmdSave.Move Me.ScaleWidth - cmdSave.Width - 100, 30, cmdSave.Width, 300
                cmd����.Move cmdSave.Left - cmd����.Width - 30, 30, cmd����.Width, 300

            Else
                cmd����.Move Me.ScaleWidth - cmd����.Width - 100, 30, cmd����.Width, 300
            End If
            tabClass.Width = cmd����.Left - 100
            lvwģ��.ColumnHeaders.Item(1).Width = lvwģ��.Width - 100
            cmdSave.ZOrder
            cmd����.ZOrder
            vsExist.ZOrder
        Else
            picLineY.Visible = False
            picLineX.Visible = False
            tvwMain_S.Visible = False
            vsList.Visible = False
            txtFind.Visible = False
            ImgFind.Visible = False
            stbThis.Visible = False
            vsExist.Height = Me.ScaleHeight - IIf(tabClass.Visible = True, tabClass.Height, 0) - pic.Height
            pic.Top = vsExist.Height + vsExist.Top
            pic.Width = cmdDelete.Width + cmdAdd.Width
            pic.Left = Me.Width - pic.Width - 200
            vsExist.ZOrder
        End If
    Else
        picLineY.Height = Me.ScaleHeight - stbThis.Height
        lvwģ��.Visible = True
        cmd����.Visible = False
        cmdSave.Visible = False
        lblģ��.Left = 30
        lblģ��.Top = 30
        lvwģ��.Move 0, lblģ��.Height + lblģ��.Top, picLineY.Left, picLineX.Top - lblģ��.Height - lblģ��.Top
        If tabClass.Visible = False Then
            lbl��Ŀ.Left = picLineY.Left + picLineY.Width + 30
            lbl��Ŀ.Top = 30
            vsExist.Move picLineY.Left + picLineY.Width, lbl��Ŀ.Height + lbl��Ŀ.Top, Me.ScaleWidth - picLineY.Left - picLineY.Width, picLineX.Top - lbl��Ŀ.Height - lbl��Ŀ.Top
        Else
            lbl��Ŀ.Left = picLineY.Left + picLineY.Width
            lbl��Ŀ.Top = 30
            tabClass.Top = lbl��Ŀ.Height + lbl��Ŀ.Top
            vsExist.Move picLineY.Left + picLineY.Width, tabClass.Top + tabClass.Height - 30, Me.ScaleWidth - picLineY.Left - picLineY.Width, picLineX.Top - tabClass.Height + 30 - lbl��Ŀ.Height - lbl��Ŀ.Top
            tabClass.Left = picLineY.Left + picLineY.Width
            tabClass.Width = vsExist.Width
        End If
        lbl����.Left = 30
        lbl����.Top = picLineX.Top + picLineX.Height + 30
        tvwMain_S.Move 0, lbl����.Top + lbl����.Height, picLineY.Left, Me.ScaleHeight - picLineX.Top - picLineX.Height - stbThis.Height - lbl����.Height - 30
        picLineY.Top = Me.ScaleTop
        pic.Move picLineY.Left + picLineY.Width, tvwMain_S.Top, Me.Width - picLineY.Left - picLineY.Width, pic.Height
        vsList.Move pic.Left, tvwMain_S.Top + pic.Height + 20, pic.Width - 100, tvwMain_S.Height - pic.Height - 20
        tvwMain_S.Visible = True
        txtFind.Width = Me.Width - pic.Left - txtFind.Left - 200
        stbThis.Visible = True
        lvwģ��.ColumnHeaders.Item(1).Width = lvwģ��.Width - 100
        vsExist.ZOrder
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Set mrsExistItem = Nothing
'    mblnDel = False
'    mblnDelPrv = False
    mlngCount = 0
    
End Sub

Private Sub ImgFind_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub lvwģ��_DblClick()
    Dim strSQL As String, blnTrans As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    
On Error GoTo errHandle
   
    If mblnģ�� = False Then
        If MsgBox("�ò��˵�������Ŀ�Ƿ����ģ��" & lvwģ��.SelectedItem.Text & "�������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then lvwģ��.Visible = False: Exit Sub
        
        strSQL = "Select A.��Ŀid From ������Ŀģ�� A ,����֧����Ŀ B Where A.��Ŀid = B.�շ�ϸĿID And B.���� = [4] And A.���� = [1] And B.Ҫ������ = 1 " & _
                 "Minus " & _
                 " Select ��ĿID From ����������Ŀ A Where A.����id = [2] And ��ҳID = [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(lvwģ��.SelectedItem.Key, 2), mlng����ID, mlng��ҳID, mlng����)
        
        gcnOracle.BeginTrans: blnTrans = True
        While Not rsTemp.EOF
            strTemp = rsTemp!��Ŀid & "," & strTemp
            lngCount = lngCount + 1
            
            If lngCount = 100 Then
                gstrSQL = "Zl_����������Ŀ_Insert(" & mlng����ID & "," & mlng��ҳID & ",'" & strTemp & "','" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                lngCount = 0
                strTemp = ""
            End If
            
            rsTemp.MoveNext
        Wend
        
        If strTemp <> "" Then
             gstrSQL = "Zl_����������Ŀ_Insert(" & mlng����ID & "," & mlng��ҳID & ",'" & strTemp & "','" & UserInfo.���� & "')"
             Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        Call ReadExistsItem(mlng����ID, mlng��ҳID)
        vsList.Tag = ""
        If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
        lvwģ��.Visible = False
    End If
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwģ��_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mblnģ�� = True Then
        If UCase(Mid(lvwģ��.SelectedItem.Key, 2)) <> "ADD" Then
            If lvwģ��.Tag <> lvwģ��.SelectedItem.Key Then
                Call ReadTempletItem(Mid(lvwģ��.SelectedItem.Key, 2))
                vsList.Tag = ""
                If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
                lvwģ��.Tag = lvwģ��.SelectedItem.Key
            End If
        Else
            Set mrsExistItem = Nothing
            tabClass.Visible = False
            vsList.Tag = ""
            If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
            Call Form_Resize
            lvwģ��.Tag = ""
            vsExist.Rows = 1
            View״̬
        End If
    End If
End Sub

Private Sub lvwģ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If mblnģ�� = False Then
        If KeyCode = vbKeyEscape Then
            lvwģ��.Visible = False
        End If
    End If
End Sub

Private Sub lvwģ��_KeyPress(KeyAscii As Integer)
    If mblnģ�� = False Then
        If KeyAscii = vbKeyReturn Then
            Call lvwģ��_DblClick
        End If
    End If
End Sub

Private Sub picLineX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With picLineX
        If .Top + Y < 1000 Then Exit Sub
        If .Top + Y > Me.ScaleHeight - 2000 Then Exit Sub
        .Top = .Top + Y
    End With
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub picLineY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With picLineY
        If .Left + X < 2000 Then Exit Sub
        If .Left + X > Me.ScaleWidth - 3500 Then Exit Sub
        
        .Move .Left + X
    End With
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub InitForm()
'��ʼ������
    With vsExist
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, ColAdded.ѡ��) = ""
        .TextMatrix(0, ColAdded.ʹ������) = "ʹ������"
        .TextMatrix(0, ColAdded.���) = "���"
        .TextMatrix(0, ColAdded.����) = "����"
        .TextMatrix(0, ColAdded.����) = "����"
        .TextMatrix(0, ColAdded.���) = "���"
        .TextMatrix(0, ColAdded.����) = "����"
        .TextMatrix(0, ColAdded.��λ) = "��λ"
        .TextMatrix(0, ColAdded.˵��) = "˵��"
        .TextMatrix(0, ColAdded.������) = "������"
        .TextMatrix(0, ColAdded.����ʱ��) = "����ʱ��"
        .TextMatrix(0, ColAdded.ID) = "ID"

        .Cell(flexcpAlignment, 0, ColAdded.���, 0, .Cols - 1) = 4
        .ColWidth(ColAdded.ѡ��) = 240
        .ColWidth(ColAdded.ʹ������) = 650
        .ColWidth(ColAdded.���) = 650
        .ColWidth(ColAdded.����) = 1100
        .ColWidth(ColAdded.����) = 2000
        .ColWidth(ColAdded.���) = 1800
        .ColWidth(ColAdded.����) = 2000
        .ColWidth(ColAdded.��λ) = 500
        .ColWidth(ColAdded.˵��) = 3000
        .ColWidth(ColAdded.������) = 900
        .ColWidth(ColAdded.����ʱ��) = 900
        .ColWidth(ColAdded.ID) = 0
        If mblnDel = False Then
            .ColHidden(ColAdded.ѡ��) = True
        End If
        .ColHidden(ColAdded.ID) = True
    End With
    
    
    vsExist.ColHidden(ColAdded.ID) = True
    If mblnģ�� Then
        vsExist.ColHidden(ColAdded.ʹ������) = True
        vsExist.ColHidden(ColAdded.������) = True
        vsExist.ColHidden(ColAdded.����ʱ��) = True
    End If
    
    With vsList
        .Rows = 2
        .Cols = 9
        .TextMatrix(0, ColAdd.ѡ��) = ""
        .TextMatrix(0, ColAdd.���) = "���"
        .TextMatrix(0, ColAdd.����) = "����"
        .TextMatrix(0, ColAdd.����) = "����"
        .TextMatrix(0, ColAdd.���) = "���"
        .TextMatrix(0, ColAdd.����) = "����"
        .TextMatrix(0, ColAdd.��λ) = "��λ"
        .TextMatrix(0, ColAdd.˵��) = "˵��"
        .TextMatrix(0, ColAdd.ID) = "ID"
        
        .Cell(flexcpAlignment, 0, ColAdded.���, 0, ColAdded.˵��) = 4
        .ColWidth(ColAdd.���) = 650
        .ColWidth(ColAdd.����) = 1100
        .ColWidth(ColAdd.����) = 1700
        .ColWidth(ColAdd.���) = 1300
        .ColWidth(ColAdd.����) = 1500
        .ColWidth(ColAdd.��λ) = 500
        .ColWidth(ColAdd.˵��) = 1700
        .ColWidth(ColAdd.ID) = 0
        
        .ColHidden(ColAdd.ID) = True
        .Cell(flexcpChecked, 1, ColAdd.ѡ��) = 1
    End With
End Sub

Private Function FillTree() As Boolean
'����:װ���շ������շ�ϸĿ�����з��ൽtvwMain_S
    '�����������ڵ�����������KEYֵ��һ���ַ������ڶ�λ��������
    Dim i As Long
    Dim objNode As Node
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    LockWindowUpdate tvwMain_S.hWnd
    tvwMain_S.Nodes.Clear
    tvwMain_S.Sorted = False
    
    '��ʾ����
    strSQL = " Select ��, ����, A.ID, �ϼ�id, ����" & _
             " From (Select Level As ��, 0 As ����, ID, �ϼ�id, '[' || ���� || ']' || ���� As ����" & _
             "        From �շѷ���Ŀ¼ A" & _
             "        Start With �ϼ�id Is Null" & _
             "        Connect By Prior ID = �ϼ�id) A," & _
             "      (Select Distinct (ID) ID" & _
             "        From �շѷ���Ŀ¼ A" & _
             "        Start With ID In (Select Distinct A.����id From �շ���ĿĿ¼ A,����֧����Ŀ D " & _
             "                          Where A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1 And (A.վ��=[1] Or A.վ�� is Null))" & _
             "        Connect By Prior �ϼ�id = ID) B" & _
             " Where a.ID = B.ID" & _
             " Union"
              
    strSQL = strSQL & _
             " Select 0 As ��,����, To_Number('99999999' || ����) As ID, -null As �ϼ�id," & _
             "        Chr(13) || Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ', 7, '��������') As ����" & _
             " From ���Ʒ���Ŀ¼ " & _
             " Where Instr(',1,2,3,7,', ',' || ���� || ',') > 0" & _
             " Union"
    
    strSQL = strSQL & _
             " Select ��, ����, A.ID, �ϼ�id, ����" & _
             " From (Select Level As ��, ����, ID As ID, Nvl(�ϼ�id, To_Number('99999999' || ����)) As �ϼ�id," & _
             "               '[' || ���� || ']' || ���� As ����" & _
             "        From ���Ʒ���Ŀ¼" & _
             "        Where Instr(',1,2,3,7,', ',' || ���� || ',') > 0" & _
             "        Start With �ϼ�id Is Null" & _
             "        Connect By Prior ID = �ϼ�id) A," & _
             "      (Select Distinct ID" & _
             "        From ���Ʒ���Ŀ¼" & _
             "        Start With ID In (Select Distinct (B.����id) ����id" & _
             "                          From �շ���ĿĿ¼ A, ������ĿĿ¼ B, ҩƷ��� C,����֧����Ŀ D" & _
             "                          Where A.ID = C.ҩƷid And B.ID = C.ҩ��id AND A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1" & _
             "                                 And (A.վ��=[1] Or A.վ�� is Null))" & _
             "        Connect By Prior �ϼ�id = ID) B" & _
             " Where a.ID = B.ID"

    On Error GoTo errHandle
    'by lesfeng 2010-03-06 ���ܰ�
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!�ϼ�ID) Then
            Set objNode = tvwMain_S.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!����, "RootS", "Exp")
        Else
            Set objNode = tvwMain_S.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, rsTmp!����, "RootS", "Exp")
        End If
        objNode.Tag = rsTmp!���� '��ŷ�������:0-��ҩƷ������,1-����ҩ,2-�г�ҩ,3-�в�ҩ,7-��������
        objNode.ExpandedImage = "Exp"
        rsTmp.MoveNext
    Next
    If tvwMain_S.Nodes.Count > 0 Then
        tvwMain_S.Nodes(1).Expanded = True
        If tvwMain_S.Nodes(1).Children > 0 Then
            tvwMain_S.Nodes(1).Child.Selected = True
        Else
            tvwMain_S.Nodes(1).Selected = True
        End If
        tvwMain_S.SelectedItem.EnsureVisible
        Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    End If
    On Error Resume Next
    If Not tvwMain_S.Nodes.Item("_999999991") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999991").Children = 0 Then tvwMain_S.Nodes.Remove "_999999991"
    End If
    If Not tvwMain_S.Nodes.Item("_999999992") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999992").Children = 0 Then tvwMain_S.Nodes.Remove "_999999992"
    End If
    If Not tvwMain_S.Nodes.Item("_999999993") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999993").Children = 0 Then tvwMain_S.Nodes.Remove "_999999993"
    End If
    If Not tvwMain_S.Nodes.Item("_999999997") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999997").Children = 0 Then tvwMain_S.Nodes.Remove "_999999997"
    End If
    FillTree = True
    Screen.MousePointer = 0
    LockWindowUpdate 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Function FillList(str���� As String, str���� As String) As Boolean
    Dim strSQL  As String, str��� As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim int��� As Integer
    Dim bln���� As Boolean
    Dim lng���� As Long
    
    If mblnģ�� = True Then
        If lvwģ��.SelectedItem Is Nothing Then
            lng���� = 0
        ElseIf UCase(Mid(lvwģ��.SelectedItem.Key, 2)) = "ADD" Then
            lng���� = 0
        Else
            lng���� = Mid(lvwģ��.SelectedItem.Key, 2)
        End If
    End If
    
    Select Case str����
        Case 1
            str��� = 5
        Case 2
            str��� = 6
        Case 3
            str��� = 7
        Case 7
            str��� = 4
    End Select
    
    bln���� = InStr(str����, "99999") = 0
    Screen.MousePointer = 11
    If mblnģ�� = False Then
        If str���� = 0 Then
             strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                      " From �շ���ĿĿ¼ A," & _
                      "      (Select A.ID" & _
                      "        From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                      "        Where A.����id In (Select ID From �շѷ���Ŀ¼ Start With ID = [1] Connect By Prior ID = �ϼ�id) And" & _
                      "              A.������� In (2, 3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                      "              A.ID = D.�շ�ϸĿID And D.���� = [4] And D.Ҫ������ = 1" & _
                      "              And (A.վ��=[5] Or A.վ�� is Null)" & _
                      "        Minus" & _
                      "        Select ��ĿID From ����������Ŀ A Where A.����id = [2] And ��ҳID = [3]) B, �շ���Ŀ��� C" & _
                      " Where A.ID = B.ID And A.��� = C.����" & _
                      " Order By A.���, A.����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����, mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
        Else
           If bln���� = True Then
                str���� = CStr(CLng(str����))
                strSQL = " Select C.���� �������,A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��,A.ID" & _
                         " From �շ���ĿĿ¼ A, " & _
                         " (SELECT A.ID " & _
                         "   FROM �շ���ĿĿ¼ A,�շ���Ŀ��� B, ҩƷ��� D, ������ĿĿ¼ E,����֧����Ŀ F" & _
                         "   Where A.��� = B.���� And A.ID = D.ҩƷid And D.ҩ��id = E.ID And" & _
                         "      E.����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = [1] Connect By Prior ID = �ϼ�id) And" & _
                         "      A.������� In (2, 3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                         "      A.���=[2] AND A.ID = F.�շ�ϸĿID And F.���� = [5] And F.Ҫ������ = 1" & _
                         "      And (A.վ��=[6] Or A.վ�� is Null) And (E.վ��=[6] Or E.վ�� is Null)" & _
                         "   Minus" & _
                         "   Select ��Ŀid From ����������Ŀ A Where A.����id = [3] And ��ҳID = [4]) B, �շ���Ŀ��� C" & _
                         " Where  A.ID = B.ID And A.��� = C.����" & _
                         " Order By A.����"
                'by lesfeng 2010-03-06 ���ܰ�
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(str����), str���, mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
            Else
                strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                         " From �շ���ĿĿ¼ A, " & _
                         "      (Select A.ID" & _
                         "       From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                         "       Where A.������� In (2, 3) And" & _
                         "            (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                         "             A.��� = [1] And A.ID = D.�շ�ϸĿID And D.���� = [4] And D.Ҫ������ = 1" & _
                         "             And (A.վ��=[5] Or A.վ�� is Null)" & _
                         "       Minus" & _
                         "       Select ��Ŀid From ����������Ŀ Where ����id = [2] And ��ҳID = [3]) B, �շ���Ŀ��� C" & _
                         "  Where A.ID = B.ID And A.��� = C.����" & _
                         " Order By A.����"
                'by lesfeng 2010-03-06 ���ܰ�
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���, mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
            End If
        End If
    Else
        If str���� = 0 Then
             strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                      " From �շ���ĿĿ¼ A," & _
                      "      (Select Distinct A.ID" & _
                      "        From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                      "        Where A.����id In (Select ID From �շѷ���Ŀ¼ Start With ID = [1] Connect By Prior ID = �ϼ�id) And" & _
                      "              A.������� In (2, 3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                      "              A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1" & _
                      "             And (A.վ��=[3] Or A.վ�� is Null)" & _
                      "        Minus" & _
                      "        Select ��ĿID From ������Ŀģ�� A Where A.���� = [2]) B, �շ���Ŀ��� C" & _
                      " Where A.ID = B.ID And A.��� = C.����" & _
                      " Order By A.���, A.����"
             'by lesfeng 2010-03-06 ���ܰ�
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����, lng����, gstrNodeNo)
        Else
           If bln���� = True Then
                str���� = CStr(CLng(str����))
                strSQL = " Select C.���� �������,A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��,A.ID" & _
                         " From �շ���ĿĿ¼ A, " & _
                         " (SELECT Distinct A.ID " & _
                         "   FROM �շ���ĿĿ¼ A,�շ���Ŀ��� B, ҩƷ��� D, ������ĿĿ¼ E,����֧����Ŀ F" & _
                         "   Where A.��� = B.���� And A.ID = D.ҩƷid And D.ҩ��id = E.ID And" & _
                         "      E.����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = [1] Connect By Prior ID = �ϼ�id) And" & _
                         "      A.������� In (2, 3) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                         "      A.���=[2] AND A.ID = F.�շ�ϸĿID And F.Ҫ������ = 1" & _
                         "      And (A.վ��=[4] Or A.վ�� is Null) And (E.վ��=[4] Or E.վ�� is Null)" & _
                         "   Minus" & _
                         "   Select ��Ŀid From ������Ŀģ�� A Where A.���� = [3]) B, �շ���Ŀ��� C" & _
                         " Where  A.ID = B.ID And A.��� = C.����" & _
                         " Order By A.����"
                'by lesfeng 2010-03-06 ���ܰ�
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(str����), str���, lng����, gstrNodeNo)
            Else
                strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                         " From �շ���ĿĿ¼ A, " & _
                         "      (Select Distinct A.ID" & _
                         "       From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                         "       Where A.������� In (2, 3) And" & _
                         "            (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                         "             A.��� = [1] And A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1" & _
                         "           And (A.վ��=[3] Or A.վ�� is Null)" & _
                         "       Minus" & _
                         "       Select ��Ŀid From ������Ŀģ�� Where ���� = [2]) B, �շ���Ŀ��� C" & _
                         "  Where A.ID = B.ID And A.��� = C.����" & _
                         " Order By A.����"
                'by lesfeng 2010-03-06 ���ܰ�
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���, lng����, gstrNodeNo)
            End If
        End If
    End If
On Error GoTo errHandle
    
    vsList.Rows = 1
    lngRow = 1
    While Not rsTemp.EOF
        With vsList
            .Rows = lngRow + 1
            .TextMatrix(lngRow, ColAdd.����) = rsTemp!����
            .TextMatrix(lngRow, ColAdd.���) = rsTemp!�������
            If str��� <> rsTemp!������� Then
                str��� = rsTemp!�������
                int��� = int��� + 1
            End If
            .TextMatrix(lngRow, ColAdd.����) = rsTemp!����
            .TextMatrix(lngRow, ColAdd.���) = rsTemp!��� & ""
            .TextMatrix(lngRow, ColAdd.��λ) = rsTemp!���㵥λ & ""
            .TextMatrix(lngRow, ColAdd.����) = rsTemp!���� & ""
            .TextMatrix(lngRow, ColAdd.˵��) = rsTemp!˵�� & ""
            .TextMatrix(lngRow, ColAdd.ID) = rsTemp!ID
        End With
        lngRow = lngRow + 1
        rsTemp.MoveNext
    Wend
    If int��� = 0 Or int��� = 1 Then
        vsList.ColHidden(ColAdd.���) = True
    Else
        vsList.ColHidden(ColAdd.���) = False
    End If
    
    If str���� = 0 Then
        vsList.ColHidden(ColAdd.����) = True
    Else
        vsList.ColHidden(ColAdd.����) = False
    End If
    
    If vsList.Rows > 1 Then
        vsList.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsList.Rows - 1, ColAdd.ѡ��) = 2
    End If
    Screen.MousePointer = 0
    vsList.Editable = flexEDKbdMouse
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub tabClass_Click()
    If tabClass.SelectedItem.Index <> 1 Then
        mrsExistItem.Filter = "���='" & tabClass.SelectedItem.Tag & "'"
    Else
        mrsExistItem.Filter = 0
    End If
    Set vsExist.DataSource = mrsExistItem
    If tabClass.SelectedItem.Index <> 1 Then
        vsExist.ColHidden(ColAdded.���) = True
    Else
        vsExist.ColHidden(ColAdded.���) = False
    End If
    vsExist.ColHidden(ColAdded.ID) = True
    
    If InStr("�в�ҩ,�г�ҩ,����ҩ,����", tabClass.SelectedItem.Tag) = 0 Then
        vsExist.ColHidden(ColAdded.����) = True
    Else
         vsExist.ColHidden(ColAdded.����) = False
    End If
    
    If mblnģ�� Then
        vsExist.ColHidden(ColAdded.ʹ������) = True
        vsExist.ColHidden(ColAdded.������) = True
        vsExist.ColHidden(ColAdded.����ʱ��) = True
    End If
    
    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.ѡ��) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.ѡ��, vsExist.Rows - 1, ColAdded.ѡ��) = 2
    End If
    
    vsExist.ColAlignment(ColAdded.����) = flexAlignLeftCenter
    vsExist.Tag = tabClass.SelectedItem.Index
End Sub

Private Sub tvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mblnģ�� Then
        If lvwģ��.Visible Then lvwģ��.Visible = False
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If vsList.Tag <> Node.Key Then
        Call FillList(Node.Tag, Mid(Node.Key, 2))
        vsList.Tag = Node.Key
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFind.Text) = "" Then Exit Sub
        If UCase(txtFind.Text) <> UCase(txtFind.Tag) Then
            Call FindItem(Trim(txtFind.Text))
            txtFind.Tag = txtFind.Text
            vsList.Tag = ""
        End If
    End If
End Sub

Private Sub vsExist_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = ColAdded.ʹ������ Then
        On Error GoTo errHandle
        With vsExist
            gstrSQL = "Zl_����������Ŀ_Update(" & mlng����ID & "," & mlng��ҳID & "," & .TextMatrix(Row, ColAdded.ID) & "," & Val(.TextMatrix(Row, Col)) & ")"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Sub

Private Sub vsExist_EnterCell()
    If vsExist.Col = ColAdded.ѡ�� Then
        vsExist.Editable = flexEDKbdMouse
    ElseIf vsExist.Col = ColAdded.ʹ������ Then
        vsExist.Editable = flexEDKbdMouse
    Else
        vsExist.Editable = flexEDNone
    End If
End Sub

Private Sub vsExist_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsExist.Row > 0 Then
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            If vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.ѡ��) = 1 Then
                If mblnDelPrv = True Then
                    Call cmdDelete_Click
                End If
            Else
                vsExist.TextMatrix(vsExist.Row, ColAdded.ʹ������) = ""
                Call vsExist_AfterEdit(vsExist.Row, ColAdded.ʹ������)
            End If
        End If
    End If
End Sub

Private Sub vsExist_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn And Col = ColAdded.ʹ������ Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsExist_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If vsExist.Row < vsExist.Rows - 1 Then
            vsExist.Row = vsExist.Row + 1
            vsExist.Col = ColAdded.ʹ������
            vsExist.ShowCell vsExist.Row, vsExist.Col
        Else
            cmdDelete.SetFocus
        End If
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
        vsExist.Editable = flexEDKbdMouse
        vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.ѡ��) = IIf(vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.ѡ��) = 1, 2, 1)
        vsExist.Editable = flexEDNone
    End If
End Sub

Private Sub vsExist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mblnģ�� Then
        If lvwģ��.Visible Then lvwģ��.Visible = False
    End If
End Sub

Private Sub vsList_EnterCell()
    If vsList.Col = ColAdd.ѡ�� Then
        vsList.Editable = flexEDKbdMouse
    Else
        vsList.Editable = flexEDNone
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If vsList.Row < vsList.Rows - 1 Then
            vsList.Row = vsList.Row + 1
            Do Until vsList.RowHidden(vsList.Row) = False
                vsList.Row = vsList.Row + 1
            Loop
            vsList.Col = ColAdd.ѡ��
            vsList.ShowCell vsList.Row, vsList.Col
        End If
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
        vsList.Editable = flexEDKbdMouse
        vsList.Cell(flexcpChecked, vsList.Row, ColAdd.ѡ��) = IIf(vsList.Cell(flexcpChecked, vsList.Row, ColAdd.ѡ��) = 1, 2, 1)
        vsList.Editable = flexEDNone
    End If
End Sub

Private Function ReadExistsItem(lng����ID As Long, lng��ҳID As Long) As Boolean
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
    
    strSQL = " Select NULL,B.ʹ������,C.���� ���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��,B.������,TRUNC(B.����ʱ��) ����ʱ��,A.ID" & _
             " From �շ���ĿĿ¼ A,����������Ŀ B, �շ���Ŀ��� C " & _
             " Where A.��� = C.���� And A.ID=B.��ĿID AND B.����ID=[1] AND B.��ҳID=[2]" & _
             "       And (A.վ��=[3] Or A.վ�� is Null)" & _
             " order by ���,����"
On Error GoTo errHandle
    'by lesfeng 2010-03-06 ���ܰ�
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, gstrNodeNo)
    mlngCount = mrsExistItem.RecordCount
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        If mblnDel = True Then
            ReadExistsItem = False
            Exit Function
        Else
            With vsExist
                .Cell(flexcpAlignment, 0, ColAdded.���, 0, .Cols - 1) = 4
                .ColWidth(ColAdded.ѡ��) = 240
                .ColWidth(ColAdded.���) = 650
                .ColWidth(ColAdded.����) = 1100
                .ColWidth(ColAdded.����) = 1700
                .ColWidth(ColAdded.���) = 1300
                .ColWidth(ColAdded.����) = 1500
                .ColWidth(ColAdded.��λ) = 500
                .ColWidth(ColAdded.˵��) = 1700
                .ColWidth(ColAdded.ID) = 0
            End With
        End If
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!��� <> strOld Then
            strClass = strClass & "," & mrsExistItem!���
            strOld = mrsExistItem!���
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tabClass.Visible = True
        tabClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '��Alt��ݼ������޷�����
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i))
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tabClass.Visible = False
    End If
    vsExist.ColWidth(ColAdded.ID) = 0
    vsExist.ColHidden(ColAdded.ID) = True
    
    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.ѡ��) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.ѡ��, vsExist.Rows - 1, ColAdded.ѡ��) = 2
    End If

    vsExist.ColAlignment(ColAdded.����) = flexAlignLeftCenter
    If vsExist.Tag <> "" Then
        If vsExist.Tag < tabClass.Tabs.Count Then
            Set tabClass.SelectedItem = tabClass.Tabs.Item(Int(vsExist.Tag))
        End If
        Call tabClass_Click
    End If
    
    If vsExist.Col = 1 Then
        vsExist.Editable = flexEDKbdMouse
    End If
    Call View״̬
    Call Form_Resize
    vsExist.Editable = flexEDKbdMouse
    ReadExistsItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ReadExistsItem = False
    Call SaveErrLog
End Function

Private Sub ReadTempletItem(lng���� As Long)
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
             
             
    strSQL = "Select NULL,NULL,C.���� ���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��,Null ������,Null ����ʱ��,A.Id " & _
             "From �շ���ĿĿ¼ A,������Ŀģ�� B, �շ���Ŀ��� C " & _
             "Where A.��� = C.���� And A.ID = B.��ĿID And B.���� = [1] " & _
             "      And (A.վ��=[2] Or A.վ�� is Null)" & _
             "Order By ���,����"
             
On Error GoTo errHandle
    'by lesfeng 2010-03-06 ���ܰ�
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����, gstrNodeNo)
    mlngCount = mrsExistItem.RecordCount
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        If mblnDel = True Then
            Unload Me
            Exit Sub
        Else
            With vsExist
                .Cell(flexcpAlignment, 0, ColAdded.���, 0, .Cols - 1) = 4
                .ColWidth(ColAdded.ѡ��) = 240
                .ColWidth(ColAdded.���) = 650
                .ColWidth(ColAdded.����) = 1100
                .ColWidth(ColAdded.����) = 1700
                .ColWidth(ColAdded.���) = 1300
                .ColWidth(ColAdded.����) = 1500
                .ColWidth(ColAdded.��λ) = 500
                .ColWidth(ColAdded.˵��) = 1700
            End With
        End If
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!��� <> strOld Then
            strClass = strClass & "," & mrsExistItem!���
            strOld = mrsExistItem!���
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tabClass.Visible = True
        tabClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '��Alt��ݼ������޷�����
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i))
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tabClass.Visible = False
    End If
    
    vsExist.ColHidden(ColAdded.ID) = True
    vsExist.ColHidden(ColAdded.ʹ������) = True
    vsExist.ColHidden(ColAdded.������) = True
    vsExist.ColHidden(ColAdded.����ʱ��) = True

    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.ѡ��) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.ѡ��, vsExist.Rows - 1, ColAdded.ѡ��) = 2
    End If
    
    vsExist.ColAlignment(ColAdded.����) = flexAlignLeftCenter
    If vsExist.Tag <> "" Then
        If vsExist.Tag < tabClass.Tabs.Count Then
            Set tabClass.SelectedItem = tabClass.Tabs.Item(Int(vsExist.Tag))
        End If
        Call tabClass_Click
    End If
    
    If vsExist.Col = 1 Then
        vsExist.Editable = flexEDKbdMouse
    End If
    Call View״̬
    Call Form_Resize
    vsExist.Editable = flexEDKbdMouse
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
'    Resume
End Sub

Private Sub FindItem(strWhere As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim str��� As String
    Dim int��� As Integer, lngRow As Long, lng���� As Long
    
    If mblnģ�� = True Then
        If UCase(Mid(lvwģ��.SelectedItem.Key, 2)) = "ADD" Then Exit Sub
        lng���� = Mid(lvwģ��.SelectedItem.Key, 2)
    End If
    
On Error GoTo errHandle
    Screen.MousePointer = 11
    If mblnģ�� = False Then
        If IsNumeric(Trim(strWhere)) Then
            
            strWhere = gstrLike & strWhere & "%"
        
            strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select A.ID" & _
                     "       From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                     "       Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND A.������� IN (2,3) And  " & _
                     "             A.ID = D.�շ�ϸĿID And D.���� = [4] And D.Ҫ������ = 1 And A.���� Like [1]" & _
                     "             And (A.վ��=[5] Or A.վ�� is Null)" & _
                     "       Minus" & _
                     "       Select A.��Ŀid From ����������Ŀ A Where ����id =[2] AND ��ҳID = [3]) B, �շ���Ŀ��� C" & _
                     " Where A.ID = B.ID And A.��� = C.����" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strWhere, mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
            
        ElseIf zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
        'ֻ����ĸʱ��ѯ����
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select �շ�ϸĿid" & _
                     "        From �շ���Ŀ����" & _
                     "        Where " & IIf(gbytCode + 1 = 3, "", "���� = [1] And") & " ���� Like [2]" & _
                     "        Group By �շ�ϸĿid" & _
                     "        Minus" & _
                     "        Select A.��Ŀid From ����������Ŀ A Where ����id = [3] And ��ҳID = [4]) B, �շ���Ŀ��� C,����֧����Ŀ D" & _
                     " Where A.ID = B.�շ�ϸĿid And A.��� = C.���� And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                     "       A.������� In (2, 3) And A.ID = D.�շ�ϸĿID And D.���� = [5] And D.Ҫ������ = 1" & _
                     "       And (A.վ��=[6] Or A.վ�� is Null)" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gbytCode + 1, UCase(strWhere), mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
        Else
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select �շ�ϸĿid" & _
                     "        From �շ���Ŀ����" & _
                     "        Where ����=1 AND ���� Like [1]" & _
                     "        Group By �շ�ϸĿid" & _
                     "        Minus" & _
                     "        Select A.��Ŀid From ����������Ŀ A Where ����id = [2] And ��ҳID = [3]) B, �շ���Ŀ��� C,����֧����Ŀ D" & _
                     " Where A.ID = B.�շ�ϸĿid And A.��� = C.���� And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                     "       A.������� In (2, 3) And A.ID = D.�շ�ϸĿID And D.���� = [4] And D.Ҫ������ = 1" & _
                     "       And (A.վ��=[5] Or A.վ�� is Null)" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWhere), mlng����ID, mlng��ҳID, mlng����, gstrNodeNo)
        End If
    Else
         If IsNumeric(Trim(strWhere)) Then
            
            strWhere = gstrLike & strWhere & "%"
        
            strSQL = " Select C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select Distinct A.ID" & _
                     "       From �շ���ĿĿ¼ A,����֧����Ŀ D" & _
                     "       Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND A.������� IN (2,3) And  " & _
                     "             A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1 And A.���� Like [1]" & _
                     "             And (A.վ��=[3] Or A.վ�� is Null)" & _
                     "       Minus" & _
                     "       Select A.��Ŀid From ������Ŀģ�� A Where ���� =[2] ) B, �շ���Ŀ��� C" & _
                     " Where A.ID = B.ID And A.��� = C.����" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strWhere, lng����, gstrNodeNo)
            
        ElseIf zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
        'ֻ����ĸʱ��ѯ����
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select Distinct C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select �շ�ϸĿid" & _
                     "        From �շ���Ŀ����" & _
                     "        Where " & IIf(gbytCode + 1 = 3, "", "���� = [1] And") & " ���� Like [2]" & _
                     "        Group By �շ�ϸĿid" & _
                     "        Minus" & _
                     "        Select A.��Ŀid From ������Ŀģ�� A Where ���� = [3]) B, �շ���Ŀ��� C,����֧����Ŀ D" & _
                     " Where A.ID = B.�շ�ϸĿid And A.��� = C.���� And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                     "       A.������� In (2, 3) And A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1" & _
                     "       And (A.վ��=[4] Or A.վ�� is Null)" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gbytCode + 1, UCase(strWhere), lng����, gstrNodeNo)
        Else
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select Distinct C.���� �������, A.���, A.����, A.����, A.���, A.����, A.���㵥λ, A.˵��, A.ID" & _
                     " From �շ���ĿĿ¼ A," & _
                     "      (Select �շ�ϸĿid" & _
                     "        From �շ���Ŀ����" & _
                     "        Where ����=1 AND ���� Like [1]" & _
                     "        Group By �շ�ϸĿid" & _
                     "        Minus" & _
                     "        Select A.��Ŀid From ������Ŀģ�� A Where ����= [2] ) B, �շ���Ŀ��� C,����֧����Ŀ D" & _
                     " Where A.ID = B.�շ�ϸĿid And A.��� = C.���� And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & _
                     "       A.������� In (2, 3) And A.ID = D.�շ�ϸĿID And D.Ҫ������ = 1" & _
                     "       And (A.վ��=[3] Or A.վ�� is Null)" & _
                     " ORDER BY ���,����"
            'by lesfeng 2010-03-06 ���ܰ�
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWhere), lng����, gstrNodeNo)
        End If
    End If
    vsList.Rows = 1
    lngRow = 1
    While Not rsTmp.EOF
        With vsList
            .Rows = lngRow + 1
            .TextMatrix(lngRow, ColAdd.����) = rsTmp!����
            .TextMatrix(lngRow, ColAdd.���) = rsTmp!�������
            If str��� <> rsTmp!������� Then
                str��� = rsTmp!�������
                int��� = int��� + 1
            End If
            .TextMatrix(lngRow, ColAdd.����) = rsTmp!����
            .TextMatrix(lngRow, ColAdd.���) = rsTmp!��� & ""
            .TextMatrix(lngRow, ColAdd.��λ) = rsTmp!���㵥λ & ""
            .TextMatrix(lngRow, ColAdd.����) = rsTmp!���� & ""
            .TextMatrix(lngRow, ColAdd.˵��) = rsTmp!˵�� & ""
            .TextMatrix(lngRow, ColAdd.ID) = rsTmp!ID
        End With
        lngRow = lngRow + 1
        rsTmp.MoveNext
    Wend
    If int��� = 0 Or int��� = 1 Then
        vsList.ColHidden(ColAdd.���) = True
    Else
        vsList.ColHidden(ColAdd.���) = False
    End If
    If vsList.Rows > 1 Then
        vsList.Cell(flexcpChecked, 1, ColAdd.ѡ��, vsList.Rows - 1, ColAdd.ѡ��) = 2
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
     Screen.MousePointer = 0
End Sub

Public Sub ExamineEdit(lng����ID As Long, lng��ҳID As Long, lng���� As Long, Optional blnDel As Boolean = False, Optional blnģ�� As Boolean = False)
    
    RestoreWinState Me, App.ProductName
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng���� = lng����
    mblnDel = blnDel
    mblnDelPrv = InStr(frmManageExamine.mstrPrivs, "ɾ��������Ŀ")
    mblnģ�� = blnģ��
    
    Call InitForm
    
    If mblnDel = False Then
        Call FillTree
        cmdDelete.Caption = "ɾ��(&D)"
        cmdAdd.Caption = "����(&A)"
        Me.Caption = "���˷���������Ŀ�༭(��ǰ�û���" & UserInfo.���� & ")"
    Else
        cmdDelete.Caption = "ȷ��(&O)"
        cmdAdd.Caption = "ȡ��(&C)"
        Me.Caption = "���˷���������Ŀ�༭(��ǰ�û���" & UserInfo.���� & ")"
        
    End If
    If mblnģ�� = False Then
        lvwģ��.Visible = False
        cmd����.Visible = mblnDel = False
        lblģ��.Visible = False
        lbl��Ŀ.Visible = False
        lbl����.Visible = False
        Me.BackColor = &H8000000F
        cmdSave.Visible = InStr(frmManageExamine.mstrPrivs, "ģ�����") And mblnDel = False
        If mblnDelPrv = False Then
            cmdDelete.Visible = False
            cmdAdd.Left = cmdDelete.Left
            ImgFind.Left = cmdAdd.Left + cmdAdd.Width + 50
            txtFind.Left = ImgFind.Left + ImgFind.Width + 50
        End If
        
        If ReadExistsItem(mlng����ID, mlng��ҳID) = False Then Exit Sub
        Call ReadExistsTemplet
        If mblnDel = True And vsExist.Rows < 2 Then
            MsgBox "�ò���û�����÷���������Ŀ!", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Me.BackColor = &H808080
        Me.Caption = "������Ŀģ��༭(��ǰ�û���" & UserInfo.���� & ")"
        lblģ��.Visible = True
        lbl��Ŀ.Visible = True
        lbl����.Visible = True
        Call ReadExistsTemplet
    End If
    Call Form_Resize
    frmExamineEdit.Show 1, frmManageExamine
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mblnģ�� Then
        If lvwģ��.Visible Then lvwģ��.Visible = False
    End If
End Sub

