VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLabMainSendSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frmLabMainSendSample.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7455
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtSend 
      Height          =   300
      Left            =   2745
      TabIndex        =   12
      ToolTipText     =   "��дһ�η����ı걾�����������ʾȫ������"
      Top             =   4560
      Width           =   480
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2628
      TabIndex        =   8
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   350
      Left            =   6165
      TabIndex        =   7
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   270
      TabIndex        =   6
      ToolTipText     =   "Ctrl+A"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&R)"
      Height          =   350
      Left            =   1449
      TabIndex        =   5
      ToolTipText     =   "Ctrl+R"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "�Զ����"
      Height          =   350
      Left            =   4986
      TabIndex        =   4
      ToolTipText     =   "�������̺ţ����ź����ô˹��ܣ�"
      Top             =   4935
      Width           =   1100
   End
   Begin VB.CheckBox Chk�ѷ��� 
      Caption         =   "��ʾ�ѷ���"
      Height          =   180
      Left            =   390
      TabIndex        =   3
      Top             =   4590
      Width           =   1260
   End
   Begin VB.TextBox txt�̺� 
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   4545
      Width           =   765
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   6195
      TabIndex        =   1
      Top             =   4560
      Width           =   765
   End
   Begin VB.CommandButton cmdDele 
      Caption         =   "������"
      Height          =   350
      Left            =   3807
      TabIndex        =   0
      ToolTipText     =   "�������δ���ͱ��"
      Top             =   4935
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgSample 
      Height          =   4455
      Left            =   150
      TabIndex        =   9
      Top             =   75
      Width           =   7125
      _cx             =   12568
      _cy             =   7858
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      ForeColorSel    =   -2147483632
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.StatusBar stbInfo 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   14
      Top             =   5385
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13097
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtxtTmp 
      Height          =   255
      Left            =   4020
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmLabMainSendSample.frx":000C
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���η���       ��"
      Height          =   180
      Left            =   1980
      TabIndex        =   13
      Top             =   4620
      Width           =   1530
   End
   Begin VB.Image img�ѷ� 
      Height          =   240
      Left            =   2940
      Picture         =   "frmLabMainSendSample.frx":0091
      Top             =   4875
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNull 
      Height          =   255
      Left            =   3975
      Top             =   4905
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lbl�̺� 
      Caption         =   "�̺�"
      Height          =   180
      Left            =   3930
      TabIndex        =   11
      Top             =   4605
      Width           =   435
   End
   Begin VB.Label lbl��ʼ���� 
      Caption         =   "��ʼ����"
      Height          =   180
      Left            =   5415
      TabIndex        =   10
      Top             =   4590
      Width           =   765
   End
End
Attribute VB_Name = "frmLabMainSendSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrId() As String
Private mfrmMain As frmLabMain
Private Enum mCol
    ID = 0: ���: �ѷ���: ѡ��: �걾��: �̺�: ����: ����id: ����: ����ʱ��
End Enum
Private mlngSelect As Long '��ѡ��걾��
Private mlngSend As Long    '�ѷ��͵�
Private mlngNoSend As Long  'δ���͵�


Public Sub ShowME(ByRef strIDList() As String, ByVal frmMain As frmLabMain)
    mstrId = strIDList
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain

End Sub

Private Sub RefreshData()
        Dim lngCount As Long
        Dim strSQL As String, rsTmp As ADODB.Recordset
        Dim str���� As String, strIDs As String
        Dim iRow As Integer, lngSeq As Long
        On Error GoTo errHandle
100     txt���� = 1
102     txt�̺� = 0
104     mlngSelect = 0
106     mlngNoSend = 0
108     mlngSend = 0
        lngSeq = 0
110     cmdSend.Enabled = False
112     With vfgSample

114         .Rows = 2: .Cols = 10: .FixedRows = 1: .FixedCols = 0
116         .Clear

118         .TextMatrix(0, mCol.ID) = "id":         .ColWidth(mCol.ID) = 0
            .TextMatrix(0, mCol.���) = "���":     .ColWidth(mCol.���) = 600
120         .TextMatrix(0, mCol.ѡ��) = " ":        .ColWidth(mCol.ѡ��) = 300
122         .TextMatrix(0, mCol.�ѷ���) = " ": .ColWidth(mCol.�ѷ���) = 300

124         .TextMatrix(0, mCol.�걾��) = "�걾��": .ColWidth(mCol.�걾��) = 1200
126         .TextMatrix(0, mCol.�̺�) = "�̺�":     .ColWidth(mCol.�̺�) = 1200
128         .TextMatrix(0, mCol.����) = "����":     .ColWidth(mCol.����) = 1200

130         .TextMatrix(0, mCol.����id) = "����id": .ColWidth(mCol.����id) = 0
132         .TextMatrix(0, mCol.����) = "����": .ColWidth(mCol.����) = 0
134         .TextMatrix(0, mCol.����ʱ��) = "����ʱ��": .ColWidth(mCol.����ʱ��) = 1800
136         .Editable = flexEDKbdMouse

138         For lngCount = 0 To .Cols - 1
140             .FixedAlignment(lngCount) = flexAlignCenterCenter
142             If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
            Next

144         For iRow = LBound(mstrId) To UBound(mstrId)
146             strIDs = mstrId(iRow)
148             If strIDs <> "" Then
150                 If Left$(strIDs, 1) = "," Then strIDs = Mid$(strIDs, 2)
                
152                 strSQL = "Select /*+ Rule */ A.Rowid,A.id,A.�걾���,A.����,A.�Ƿ���,A.����id,A.����,A.����ʱ��" & vbNewLine & _
                            "From ����걾��¼ A, Table(Cast(f_Num2list([1]) As ZLTOOLS.t_Numlist)) B" & vbNewLine & _
                            "Where A.ID = B.Column_Value Order by  Lpad(A.�걾���,9,'0'),A.����ʱ�� "
154                 Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
156                 Do Until rsTmp.EOF
158                     .TextMatrix(.Rows - 1, mCol.ID) = Val("" & rsTmp!ID)
160                     .TextMatrix(.Rows - 1, mCol.�걾��) = Trim("" & rsTmp!�걾���)
        
162                     str���� = Trim("" & rsTmp!����)
164                     If InStr(str����, ",") > 0 Then
166                         .TextMatrix(.Rows - 1, mCol.�̺�) = Split(str����, ",")(0)
168                         .TextMatrix(.Rows - 1, mCol.����) = Split(str����, ",")(1)
170                         If Val(Split(str����, ",")(1)) > Val(txt����) Then txt���� = Split(str����, ",")(1)
172                         If Split(str����, ",")(0) <> Trim(txt�̺�) And Trim(Split(str����, ",")(0)) <> "" Then txt�̺� = Split(str����, ",")(0)
                        End If
174                     If Val("" & rsTmp!�Ƿ���) = 1 Then
176                         .Cell(flexcpPicture, .Rows - 1, mCol.�ѷ���) = img�ѷ�.Picture
178                         .Cell(flexcpPictureAlignment, .Rows - 1, mCol.�ѷ���) = flexPicAlignLeftCenter
180                         If Chk�ѷ���.Value = 1 Then
182                             .RowHidden(.Rows - 1) = False
184                             mlngSend = mlngSend + 1 '�ѷ��ͼ���
                                lngSeq = lngSeq + 1
                                .TextMatrix(.Rows - 1, mCol.���) = lngSeq
                            Else
186                             .RowHidden(.Rows - 1) = True
                            End If
                        
                        Else
                            lngSeq = lngSeq + 1
                            .TextMatrix(.Rows - 1, mCol.���) = lngSeq
188                         .TextMatrix(.Rows - 1, mCol.�ѷ���) = ""
190                         .Cell(flexcpPicture, .Rows - 1, mCol.�ѷ���) = imgNull.Picture
192                         .Cell(flexcpPictureAlignment, .Rows - 1, mCol.�ѷ���) = flexPicAlignLeftCenter
194                         mlngNoSend = mlngNoSend + 1 ' δ���ͼ���
                        End If
        
196                     .TextMatrix(.Rows - 1, mCol.����id) = Val("" & rsTmp!����id)
198                     .TextMatrix(.Rows - 1, mCol.����) = Val("" & rsTmp!����)
200                     .TextMatrix(.Rows - 1, mCol.����ʱ��) = Format("" & rsTmp!����ʱ��, "yyyy-MM-dd hh:mm:ss")
202                     .Rows = .Rows + 1
204                     rsTmp.MoveNext
                    Loop
                    
206                 If .Rows > .FixedRows Then .Rows = .Rows - 1
                End If
            Next
        End With
208     Call refreshStb
210     If vfgSample.Rows > vfgSample.FixedRows Then vfgSample.Cell(flexcpChecked, vfgSample.FixedRows, mCol.ѡ��, vfgSample.Rows - 1, mCol.ѡ��) = flexUnchecked     'ȫ����Ϊδѡ

        Exit Sub
errHandle:
'        WriteToLog "SendSample.refreshData," & CStr(Erl()) & "��," & Err.Description
212     If ErrCenter() = 1 Then
214         Resume
        End If
End Sub
Private Sub refreshStb()
    stbInfo.Panels(1).Text = "δ���ͱ걾:" & mlngNoSend & " �ѷ��ͱ걾:" & mlngSend & " ��ѡ��:" & mlngSelect
    If mlngSelect > 0 Then cmdSend.Enabled = True
    
End Sub
Private Sub Chk�ѷ���_Click()
    Dim lngRow As Long
    Dim lngSeq As Long
    mlngSend = 0
    lngSeq = 0
    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            If Chk�ѷ���.Value = 1 Then
                .RowHidden(lngRow) = False
                lngSeq = lngSeq + 1
                .TextMatrix(lngRow, mCol.���) = lngSeq
                If .Cell(flexcpPicture, lngRow, mCol.�ѷ���) <> imgNull.Picture Then
                    mlngSend = mlngSend + 1
                End If
            Else
                If .Cell(flexcpPicture, lngRow, mCol.�ѷ���) <> imgNull.Picture Then
                    .RowHidden(lngRow) = True
                Else
                    lngSeq = lngSeq + 1
                    .TextMatrix(lngRow, mCol.���) = lngSeq
                End If
            End If
        Next
        
        refreshStb
    End With
End Sub

Private Sub cmdAuto_Click()
    Dim str�̺� As String, lng���� As Long
    Dim blnAdd As Boolean, lng��ʼ�� As Long
    Dim lngRow As Long
    rtxtTmp.Text = ""
    With vfgSample
        str�̺� = Trim(txt�̺�)
        lng���� = Val(txt����)
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.�ѷ���) = imgNull.Picture Then
                If Trim(.TextMatrix(lngRow, mCol.�̺�)) <> "" And Trim(.TextMatrix(lngRow, mCol.����)) <> "" Then
                    rtxtTmp.Text = rtxtTmp.Text & "|" & .TextMatrix(lngRow, mCol.�̺�) & "," & .TextMatrix(lngRow, mCol.����)
                End If
            Else
                If Trim(.TextMatrix(lngRow, mCol.�̺�)) <> "" And Trim(.TextMatrix(lngRow, mCol.����)) <> "" Then
                    rtxtTmp.Text = rtxtTmp.Text & "|" & .TextMatrix(lngRow, mCol.�̺�) & "," & .TextMatrix(lngRow, mCol.����)
                End If
            End If
        Next

        lng��ʼ�� = .FixedRows

        For lngRow = lng��ʼ�� To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.�ѷ���) = imgNull.Picture Then
                If Trim(.TextMatrix(lngRow, mCol.�̺�)) = "" And Trim(.TextMatrix(lngRow, mCol.����)) = "" Then
                    blnAdd = False
                    Do While Not blnAdd
                        If InStr(rtxtTmp.Text & "|", "|" & str�̺� & "," & lng���� & "|") <= 0 Then
                            .TextMatrix(lngRow, mCol.�̺�) = str�̺�
                            .TextMatrix(lngRow, mCol.����) = lng����
                            Call vfgSample_AfterEdit(lngRow, mCol.����)
                            blnAdd = True
                            rtxtTmp.Text = rtxtTmp.Text & "|" & str�̺� & "," & lng����
                        Else
                            lng���� = lng���� + 1
                        End If
                    Loop
                End If
            End If
        Next

    End With
End Sub

Private Sub cmdDele_Click()
    Dim lngRow As Long

    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpPicture, lngRow, mCol.�ѷ���) = imgNull.Picture Then
               .TextMatrix(lngRow, mCol.�̺�) = ""
               .TextMatrix(lngRow, mCol.����) = ""
            End If
        Next
    End With

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    Dim lngRow As Long
    Dim str����ʱ�� As String, lng���� As Long, str�걾�� As String, lng����id As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngSendMax As Long '���͸���
    
    
    cmdAll.Enabled = False
    cmdAuto.Enabled = False
    cmdClear.Enabled = False
    cmdDele.Enabled = False
    cmdExit.Enabled = False
    cmdSend.Enabled = False
    Chk�ѷ���.Enabled = False
    
    If mlngSelect <= 0 Then
        
        Exit Sub
    End If
    lngSendMax = Val(txtSend.Text)
    If lngSendMax < 0 Then lngSendMax = 0
    If lngSendMax > mlngSelect Then lngSendMax = 0
    
    With vfgSample
        Call .Select(.Row, .COL)
'        WriteToLog "---> ���η��Ϳ�ʼ --->"
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngRow, mCol.ѡ��) = flexChecked Then
                str����ʱ�� = Format(CDate(.TextMatrix(lngRow, mCol.����ʱ��)), "yyyy-MM-dd")
                lng���� = Val(.TextMatrix(lngRow, mCol.����))
                str�걾�� = Val(.TextMatrix(lngRow, mCol.�걾��))
                lng����id = Val(.TextMatrix(lngRow, mCol.����id))

                SendSample mfrmMain.WinsockC, mfrmMain.WinsockC.LocalIP, lng����id, str����ʱ��, str�걾��, "", False, lng����
                strSQL = "Select �Ƿ���,��������,����,����ʱ��,���� From ����걾��¼ Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, mCol.ID)))
                
                Do Until rsTmp.EOF
'                    WriteToLog "���ͣ�" & str�걾�� & " " & rsTmp!���� & " " & rsTmp!�������� & " " & rsTmp!���� & " " & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                    If Val("" & rsTmp!�Ƿ���) = 1 Then
                        .Cell(flexcpPicture, lngRow, mCol.�ѷ���) = img�ѷ�.Picture
                        .Cell(flexcpPictureAlignment, lngRow, mCol.�ѷ���) = flexPicAlignLeftCenter
                    Else
                        .Cell(flexcpPicture, lngRow, mCol.�ѷ���) = imgNull.Picture
                        .Cell(flexcpPictureAlignment, lngRow, mCol.�ѷ���) = flexPicAlignLeftCenter
                    End If
                    rsTmp.MoveNext
                Loop
            End If
        Next
'        WriteToLog "<--- ���η��ͽ��� <---"
    End With

    cmdAll.Enabled = True
    cmdAuto.Enabled = True
    cmdClear.Enabled = True
    cmdDele.Enabled = True
    cmdExit.Enabled = True
    cmdSend.Enabled = True
    Chk�ѷ���.Enabled = True
End Sub

Private Sub Form_Load()
    Call RefreshData
End Sub

Private Sub cmdAll_Click()
    Dim lngRow As Long
    Dim lngCount As Long
    mlngSelect = 0
    With vfgSample
        For lngRow = .FixedRows To .Rows - 1
            'If Not (Trim(.TextMatrix(lngRow, mCol.�̺�)) = "" Or Trim(.TextMatrix(lngRow, mCol.����)) = "") Then
            If .RowHidden(lngRow) = False Then
                .Cell(flexcpChecked, lngRow, mCol.ѡ��) = flexChecked
                mlngSelect = mlngSelect + 1
            End If
            'End If
        Next
    
        refreshStb
    End With
End Sub

Private Sub cmdClear_Click()
    vfgSample.Cell(flexcpChecked, 1, mCol.ѡ��, vfgSample.Rows - 1, mCol.ѡ��) = flexUnchecked
    mlngSelect = 0
    refreshStb
End Sub

Private Sub vfgSample_AfterEdit(ByVal Row As Long, ByVal COL As Long)
    Dim strSQL As String
    Dim str���� As String, str�̺� As String

    With vfgSample
        str���� = Trim(.TextMatrix(Row, mCol.����))
        str�̺� = Trim(.TextMatrix(Row, mCol.�̺�))
        If Not (str���� = "" And str�̺� = "") Then
            strSQL = "ZL_����걾��¼_����(" & Val(.TextMatrix(Row, mCol.ID)) & ",'" & str�̺� & "," & str���� & "')"
        Else
            strSQL = "ZL_����걾��¼_����(" & Val(.TextMatrix(Row, mCol.ID)) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End With
End Sub

Private Sub vfgSample_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If InStr("," & mCol.���� & "," & mCol.�̺� & ",", "," & COL & ",") <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub vfgSample_Click()
    With vfgSample
        If .MouseCol = mCol.ѡ�� Then
            'If Trim(.TextMatrix(.Row, mCol.����)) = "" Or Trim(.TextMatrix(.Row, mCol.�̺�)) = "" Then Exit Sub
            mlngSelect = mlngSelect + IIf(.Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked, 1, -1)
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = IIf(.Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked, flexChecked, flexUnchecked)
            
        End If
        
        Call refreshStb
    End With
End Sub

Private Sub vfgSample_EnterCell()
    With vfgSample
        Dim blnCancle As Boolean
        Call vfgSample_BeforeEdit(.Row, .COL, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
End Sub

Private Sub vfgSample_LeaveCell()
    With vfgSample
        Dim blnCancle As Boolean
        Call vfgSample_BeforeEdit(.Row, .COL, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgSample_ValidateEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Dim intRow As Integer
    Dim str���� As String
    Dim str�̺�  As String
    With vfgSample
        If IsNumeric(.EditText) = False And .EditText <> "" Then Cancel = True
        Select Case COL

            Case mCol.����

                str�̺� = Trim(.TextMatrix(Row, mCol.�̺�))
                str���� = Trim(.EditText)

            Case mCol.�̺�

                str�̺� = Trim(.EditText)
                str���� = Trim(.TextMatrix(Row, mCol.����))
        End Select

        If str�̺� <> "" And str���� <> "" Then
            For intRow = .FixedRows To .Rows - 1
                If intRow <> Row Then
                    If Trim(.TextMatrix(intRow, mCol.�̺�)) <> "" And Trim(.TextMatrix(intRow, mCol.����)) <> "" Then
                        If Trim(.TextMatrix(intRow, mCol.�̺�)) = str�̺� And Trim(.TextMatrix(intRow, mCol.����)) = str���� Then
                            Cancel = True
                            .Cell(flexcpChecked, Row, mCol.ѡ��) = flexUnchecked
                       End If
                    End If
                End If
            Next
        End If

    End With

End Sub






