VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmProcDiff 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "���̵���"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9660
   Icon            =   "frmDiff.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   9660
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin XtremeSyntaxEdit.SyntaxEdit txtLeft 
      Height          =   2295
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
      _Version        =   983043
      _ExtentX        =   3201
      _ExtentY        =   4048
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      ReadOnly        =   -1  'True
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin XtremeSyntaxEdit.SyntaxEdit txtRight 
      Height          =   2175
      Left            =   6720
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
      _Version        =   983043
      _ExtentX        =   2778
      _ExtentY        =   3836
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox pctBottom 
      Height          =   1335
      Left            =   720
      ScaleHeight     =   1275
      ScaleWidth      =   2115
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2175
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   6360
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiff.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiff.frx":6C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiff.frx":700B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfIcon 
      Height          =   3135
      Left            =   6000
      TabIndex        =   16
      Top             =   840
      Width           =   255
      _cx             =   450
      _cy             =   5530
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   180
      RowHeightMax    =   180
      ColWidthMin     =   220
      ColWidthMax     =   220
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      PicturesOver    =   -1  'True
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   4560
      Top             =   1200
   End
   Begin VB.PictureBox pctOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   4440
      ScaleHeight     =   3255
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   1920
      Width           =   615
      Begin VB.Label lblsta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "F7"
         ForeColor       =   &H00404000&
         Height          =   180
         Index           =   7
         Left            =   360
         TabIndex        =   13
         Top             =   510
         Width           =   180
      End
      Begin VB.Label lblsta 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "F8"
         ForeColor       =   &H00404000&
         Height          =   180
         Index           =   8
         Left            =   360
         TabIndex        =   12
         Top             =   2310
         Width           =   180
      End
      Begin VB.Image imgDown 
         Height          =   240
         Left            =   75
         Picture         =   "frmDiff.frx":73EB
         Top             =   2280
         Width           =   240
      End
      Begin VB.Image imgUp 
         Height          =   240
         Left            =   75
         Picture         =   "frmDiff.frx":DC3D
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "������˳�(&Q)"
      Height          =   350
      Left            =   6960
      TabIndex        =   2
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "��������(&S)"
      Height          =   350
      Left            =   5400
      TabIndex        =   1
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "��һ������(&P)"
      Height          =   350
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ������(&N)"
      Height          =   350
      Left            =   1680
      TabIndex        =   5
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8400
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Image imgTitle 
      Height          =   690
      Left            =   240
      Picture         =   "frmDiff.frx":1448F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���¹��̱༭"
      Height          =   180
      Left            =   6720
      TabIndex        =   18
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�û��䶯�����޸ĺۼ�"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   960
      Width           =   2160
   End
   Begin VB.Label lblPgs 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����60�����̴��������,��ǰΪ��35������"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   3510
   End
   Begin VB.Label lblsta 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�뽫�û��䶯�������޸ĵĴ�����������µĲ�Ʒ��׼�����У��Ӷ��õ����µ��û��䶯���̡�"
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   360
      Width           =   7560
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ɫ"
      ForeColor       =   &H0000C000&
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ɫ"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "��ɫ"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   4
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   360
   End
   Begin VB.Label lblsta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ʾ�����Ĵ���,    ��ʾɾ���Ĵ���,    ��ʾ�޸Ĵ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   4500
   End
End
Attribute VB_Name = "frmProcDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnEdit As Boolean '�Ƿ����༭

Private mcolDiff1 As New Collection '����ı�����촦����
Private mcolDiff2 As New Collection '�Ҳ��ı�����촦�ļ���

Private marrIds() As String '���д����ID����
Private mlngIdx As Long '��ǰ����ID

Private mstrTipLeft As String   '�����ʾ�����
Private mstrTipright As String  '�ұ���ʾ�����

Private mstrCurUser As String   '��ǰ�û�
Private mstrCurStat As String   '�޸�˵��

Private mblnChanged As Boolean
Private mlngRows As Long    '�Ҳ��ı��������
Private mlngCurRow As Long '�Ҳ��ı���ĵ�ǰ��
Private mlngCurCol As Long '�Ҳ��ı���ĵ�ǰ��

Private mintLast As Integer '���һ����ȡ��������ı��ռ� 1-��߿ؼ�  2-�ұ߿ռ�

Private mlngLeftRow As Long
Private mlngRightRow As Long

Private Const lngRowHeight = 180
Private Enum ��ɫ
    ��ɫ = &HFFFFFF
    ����ɫ = &HC9C9CD
    ��ɫ = &H106E2A
    ��ɫ = &H0&
    ��ɫ = &H4040FF
    ��ɫ = vbBlue
End Enum

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdContinue_Click()
    Dim strErr As String, blnSuccess As Boolean
    
    
    If Not frmProcEditor.ShowMe(mstrCurUser, mstrCurStat) Then Exit Sub
    
    blnSuccess = SaveProc(strErr)
    If blnSuccess Then
        mblnEdit = True
        If mlngIdx < UBound(marrIds) Then     '����޸Ĺ���û�����һ��,�ͼ��������¸�
            mlngIdx = mlngIdx + 1
            Call LoadDiff
        End If
    Else
       MsgBox "���̱���ʧ��" & vbNewLine & strErr
    End If
    
    
End Sub

Private Sub cmdNext_Click()
    
    If mlngIdx = UBound(marrIds) Then
        MsgBox "��ǰ�Ѿ������һ�����̡�", , gstrSysName
        Exit Sub
    End If
    
    mlngIdx = mlngIdx + 1
    Call LoadDiff
End Sub

Private Sub cmdPrevious_Click()
    
    
    If mlngIdx = 0 Then
        MsgBox "��ǰ�Ѿ��ǵ�һ�����̡�", , gstrSysName
        Exit Sub
    End If
    
    mlngIdx = mlngIdx - 1
    Call LoadDiff
End Sub

Private Sub cmdQuit_Click()
    Dim strErr As String, blnSuccess As Boolean
    
    If Not frmProcEditor.ShowMe(mstrCurUser, mstrCurStat) Then Exit Sub
    
    blnSuccess = SaveProc(strErr)
    If blnSuccess Then
        mblnEdit = True
        Unload Me
    Else
       MsgBox "���̱���ʧ��" & vbNewLine & strErr
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 118 Then   '����F7
        imgUp_Click
    ElseIf KeyCode = 119 Then   '����F8
        imgDown_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcolDiff1 = Nothing
    Set mcolDiff2 = Nothing
End Sub

Public Function ShowMe(ByVal arrIds As Variant, ByVal lngIdx As Long) As Boolean
    
    marrIds = arrIds
    mlngIdx = lngIdx
     
    ShowFlash "���ڶԱ�...."
    Call LoadDiff
    ShowFlash ""
    
    Me.Show 1
    ShowMe = mblnEdit
End Function

Private Sub Form_Resize()
    On Error Resume Next

    txtLeft.Width = Me.ScaleWidth / 2 - 600
    txtRight.Width = Me.ScaleWidth / 2 - 600
    
    pctOpt.Left = txtLeft.Left + txtLeft.Width
    
    vsfIcon.Top = txtLeft.Top
    vsfIcon.Left = pctOpt.Left + pctOpt.Width
    
    txtRight.Left = vsfIcon.Left + vsfIcon.Width
    lblRight.Left = vsfIcon.Left
    
    cmdCancel.Left = txtRight.Left + txtRight.Width - cmdCancel.Width
    cmdCancel.Top = Me.ScaleHeight - cmdCancel.Height - 120
    
    cmdQuit.Left = cmdCancel.Left - 60 - cmdQuit.Width
    cmdQuit.Top = cmdCancel.Top
    
    cmdContinue.Left = cmdQuit.Left - 60 - cmdContinue.Width
    cmdContinue.Top = cmdCancel.Top
    
    cmdNext.Top = cmdCancel.Top
    cmdPrevious.Top = cmdCancel.Top
    
    txtLeft.Height = cmdCancel.Top - txtLeft.Top - 60
    txtRight.Height = cmdCancel.Top - txtRight.Top - 60
    vsfIcon.Height = txtRight.Height
    pctOpt.Top = txtLeft.Top + txtLeft.Height / 2 - pctOpt.Height / 2
    
    lblPgs.Top = cmdNext.Top + cmdNext.Height / 2 - lblPgs.Height / 2
    lblPgs.Left = cmdNext.Left + cmdNext.Width + 60
    
    pctBottom.Move txtLeft.Left, txtLeft.Top, txtLeft.Width, txtLeft.Height
End Sub


Private Sub lblsta_Click(Index As Integer)
    If Index = 7 Then
        imgUp_Click
    ElseIf Index = 8 Then
        imgDown_Click
    End If
End Sub

Private Sub imgUp_Click()
    '��λ��һ����ͬ
    Dim i As Long
    
    If mintLast = 2 Then
        '�Ҳ�ռ�,֮�������񼴿�
        For i = txtRight.CurrPos.Row - 1 To 0 Step -1
            If vsfIcon.Rows = 0 Then Exit Sub
            If vsfIcon.Cell(flexcpData, i, 0) <> 0 Then
                txtRight.CurrPos.Row = i
                txtRight.TopRow = i - (txtLeft.Height / lngRowHeight / 2)
                Exit Sub
            End If
        Next
    Else
        '���ؼ� ,��Ҫ��������
        For i = txtLeft.CurrPos.Row - 1 To 1 Step -1
            If GetValueFromCol(mcolDiff1, "_" & i) <> "" And GetValueFromCol(mcolDiff1, "_" & i - 1) = "" Then
                txtLeft.CurrPos.Row = i
                txtLeft.TopRow = i - (txtLeft.Height / lngRowHeight / 2)
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub imgDown_Click()
    '��λ��һ����ͬ
    Dim i As Long
    
    If mintLast = 2 Then
        '�Ҳ�ռ�,֮�������񼴿�
        For i = txtRight.CurrPos.Row + 1 To txtRight.RowsCount - 1
            If vsfIcon.Rows <= 0 Then Exit Sub
            If vsfIcon.Cell(flexcpData, i, 0) <> 0 Then
                txtRight.CurrPos.Row = i
                txtRight.TopRow = i - (txtLeft.Height / lngRowHeight / 2)
                Exit Sub
            End If
        Next
    Else
        '���ؼ� ,��Ҫ��������
        For i = txtLeft.CurrPos.Row + 1 To txtRight.RowsCount - 1
            If GetValueFromCol(mcolDiff1, "_" & i) <> "" And GetValueFromCol(mcolDiff1, "_" & i - 1) = "" Then
                txtLeft.CurrPos.Row = i
                txtLeft.TopRow = i - (txtLeft.Height / lngRowHeight / 2)
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub Timer_Timer()
    vsfIcon.TopRow = txtRight.TopRow
End Sub

Private Function FindOldText(ByVal lngRowNum As Long, ByVal colText As Collection) As String
    '�ҵ�ԭ�����ı�
    Dim i As Long, strResult As String, strErr As String
    Dim strTmp As String, strAfter As String, strBefore As String
    
    strResult = GetValueFromCol(colText, "_" & lngRowNum)
    If strResult = "" Or strResult = "�տ�" Then Exit Function
    
    '�����������
    i = lngRowNum - 1
    Do While i <> -1
        strTmp = GetValueFromCol(colText, "_" & i, strErr)
        If strTmp <> "�տ�" And strTmp <> "" And strErr = "" Then
            strBefore = IIf(strBefore = "", strTmp, strTmp & vbNewLine & strBefore)
            i = i - 1
        Else
            i = -1
        End If
    Loop
    
    '�����������
    i = lngRowNum + 1
    Do While i <> -1
        strTmp = GetValueFromCol(colText, "_" & i, strErr)
        If strTmp <> "�տ�" And strTmp <> "" And strErr = "" Then
            strAfter = IIf(strAfter = "", strTmp, strAfter & vbNewLine & strTmp)
            i = i + 1
        Else
            i = -1
        End If
    Loop
    
    FindOldText = IIf(strBefore = "", "", strBefore & vbNewLine) & strResult & IIf(strAfter = "", "", vbNewLine & strAfter)
End Function


Private Sub GetProcText(strTxt1 As String, strTxt2 As String, strTxt3 As String, strTxt4 As String, strSta As String, strUser As String)
    '����ID���ع��̵���Ϣ
    'strTxt1 - �ϴα�׼����  strTxt2 - ���ݿ��еĹ���  strTxt3 - ���α�׼���� strTxt4 - �����Զ����� strSta - �޸�˵�� strUser -�޸���
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strID As String, strName As String
    Dim strTmp As String
    
    On Error GoTo errH
    
    strID = Split(marrIds(mlngIdx), ":")(0)
    strName = Split(marrIds(mlngIdx), ":")(1)
    If strID = "" Then
        '���ռ����̺�,�޷���ȡID,Ҫ�������ƶ�ȡһ��ID
        strID = GetProcIdByName(strName)

        If strID = 0 Then
            MsgBox "��ȡ����IDʧ�ܣ���ˢ�����ԡ�", gstrSysName
            Exit Sub
        End If
    End If
    
    '��ȡ�汾
    strSQL = "Select ����ǰ�汾,������汾,�޸���Ա,˵�� From zlProcedure where ID = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�汾", strID)
    mstrTipLeft = "����ǰ���½ű��ļ�" & rsTmp!����ǰ�汾
    mstrTipright = "���������½ű��ļ�" & rsTmp!������汾
    strSta = rsTmp!˵�� & ""
    strUser = rsTmp!�޸���Ա & ""
    
    strSQL = "Select ���� ,���� From zlproceduretext Where ����ID=[1]  Order by ����,���"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡ�����ı�", strID)
    
    '��׼����
    rsTmp.Filter = "���� = 1"
    If rsTmp.RecordCount = 0 Then Exit Sub
    Do While Not rsTmp.EOF
        strTxt1 = IIf(strTxt1 = "", rsTmp!����, strTxt1 & vbNewLine & rsTmp!����)
        rsTmp.MoveNext
    Loop
    
    '��ǰ���ݿ����
    strTxt2 = LoadBaseProcs(UCase(strName))
    
    '������׼����
    rsTmp.Filter = "���� = 4"
    Do While Not rsTmp.EOF
        strTmp = rsTmp!���� & ""
        
        strTxt3 = IIf(strTxt3 = "", strTmp, strTxt3 & vbNewLine & strTmp)
        rsTmp.MoveNext
    Loop
    
    '�����Զ�����
    rsTmp.Filter = "���� = 3"
    Do While Not rsTmp.EOF
        strTmp = rsTmp!���� & ""
        strTxt4 = IIf(strTxt4 = "", strTmp, strTxt4 & vbNewLine & strTmp)
        rsTmp.MoveNext
    Loop
    
    If strTxt3 = "" Then
        MsgBox "��ѡ�����ڱ��������в��ᱻ�޸ģ�����������", , gstrSysName
        Exit Sub
    End If
    
    Exit Sub
errH:
    MsgBox Err.Description, , gstrSysName
End Sub

Private Sub LoadDiff()
    '���ع����漰���ı��Ĳ�ͬ
    Dim strTxt1 As String, strTxt2 As String
    Dim strTxt3 As String, strTxt4 As String
    Dim strSta As String, strUser As String
    Dim i As Long, j As Long
    
    ShowFlash "���ڼ��ضԱȽ��..."
    
    lblPgs.Caption = "��" & mlngIdx + 1 & "/" & UBound(marrIds) + 1 & "������"
    
    ''strTxt1 - �ϴα�׼����  strTxt2 - ���ݿ��еĹ���  strTxt3 - ���α�׼���� strTxt4 - �����Զ�����
    GetProcText strTxt1, strTxt2, strTxt3, strTxt4, strSta, strUser
    mstrCurUser = IIf(mstrCurUser = "", strUser, mstrCurUser)
    mstrCurStat = strSta
    
    '��׼���̺ͱ䶯���̶Ա�
    CompareIt strTxt1, strTxt2: MergeDiff strTxt1, strTxt2
    MergeDiffInto1SynEdit strTxt1, strTxt2, txtLeft, mcolDiff1
    
    '��������̺��������̶Ա�
    vsfIcon.Rows = 0
    If strTxt4 = "" Then
        txtRight.Text = strTxt3  '�ո��ռ���,û�б����Զ�����
    Else
        CompareIt strTxt3, strTxt4: MergeDiff strTxt3, strTxt4
        MergeDiffInto1SynEdit strTxt3, strTxt4, txtRight, mcolDiff2, False
        vsfIcon.Rows = txtRight.RowsCount + 200 '����ཨ��һЩ��,��ֹĩβ�����ͱ༭�����ܶ���
        vsfIcon.TopRow = txtRight.TopRow
        
        '�ڷ����޸ĵ������ͼ��
        For i = 1 To txtRight.RowsCount
            If GetValueFromCol(mcolDiff2, "_" & i + j) <> "" Then
                If GetValueFromCol(mcolDiff2, "_" & i + j - 1) <> "" And txtRight.RowText(i - 1) = "" And txtRight.RowText(i) = "" Then
                    'ɾ�� -ɾ�����пո�
                    vsfIcon.RemoveItem i
                    txtRight.RemoveRow i
                    vsfIcon.Cell(flexcpData, i - 1, 0) = vsfIcon.Cell(flexcpData, i - 1, 0) & vbNewLine & GetValueFromCol(mcolDiff2, "_" & i + j) & ""
                    j = j + 1
                    i = i - 1
                ElseIf GetValueFromCol(mcolDiff2, "_" & i + j - 1) = "" Then
                    vsfIcon.Cell(flexcpData, i, 0) = GetValueFromCol(mcolDiff2, "_" & i + j) & "" '��Ϊ�кſ��ܷ����仯,���԰Ѳ��챣����RowData��
                    
                    'ͼƬ1-ɾ��  2-�޸� 3-����
                    If GetValueFromCol(mcolDiff2, "_" & i + j) = "�տ�" And txtRight.RowText(i) <> "" Then
                        vsfIcon.Cell(flexcpPicture, i, 0) = imgIcon.ListImages(3).Picture
                    ElseIf GetValueFromCol(mcolDiff2, "_" & i + j) <> "�տ�" And txtRight.RowText(i) = "" Then
                        vsfIcon.Cell(flexcpPicture, i, 0) = imgIcon.ListImages(1).Picture
                    Else
                        vsfIcon.Cell(flexcpPicture, i, 0) = imgIcon.ListImages(2).Picture
                    End If
                Else
                    vsfIcon.Cell(flexcpData, i, 0) = GetValueFromCol(mcolDiff2, "_" & i + j)
                End If
            End If
        Next
    End If
    
    mlngCurRow = 0: mlngCurCol = 0
    mlngRows = 0
    ShowFlash ""
End Sub

Private Function SaveProc(Optional ByRef strErr As String) As Boolean
    '����:���ı����еĹ��̱��������ݿ�
    Dim i As Long, strSQL As String
    Dim strID As String, strName As String
    Dim arrProc() As String
    
    On Error GoTo errH
    strID = Split(marrIds(mlngIdx), ":")(0)
    strName = Split(marrIds(mlngIdx), ":")(1)
    
    If strID = "" Then
        strID = GetProcIdByName(strName)
        If strID = "" Then
            strErr = "��ȡ����IDʧ�ܣ���ˢ�����ԡ�"
            Exit Function
        End If
    End If
    
    gcnOracle.BeginTrans
    '�Ȱ�֮ǰ���ı�ɾ��
    strSQL = "Delete From zlProcedureText Where ����ID = " & strID & " And ���� = 3"
    gcnOracle.Execute strSQL
    
    'ȥ���ı��Ŀ���
    strSQL = txtRight.Text
    Do While InStr(1, strSQL, vbNewLine & vbNewLine) > 0
        strSQL = Replace(strSQL, vbNewLine & vbNewLine, vbNewLine)
    Loop
    If Right(strSQL, 2) = vbNewLine Then
        strSQL = Left(strSQL, Len(strSQL) - 2)
    End If
    arrProc = Split(strSQL, vbNewLine)
    
    '�����޸ĺ�ı䶯����
    strSQL = "Insert Into zlProcedureText(����ID,����,���,����) "
    For i = 0 To UBound(arrProc)
        strSQL = strSQL & vbNewLine & "Select " & strID & ",3," & i & ",'" & Replace(arrProc(i), "'", "''") & "' From Dual Union All "
    Next
    strSQL = strSQL & vbNewLine & "Select  1,1,1,'1' From Dual where 1=0"   '��������β
    gcnOracle.Execute strSQL
    
    '����״̬
    strSQL = "Update zlProcedure Set ״̬ = 3 ,  �޸���Ա = '" & mstrCurUser & "',  ˵�� = '" & mstrCurStat & "'," & vbNewLine & _
                " �޸�ʱ�� =  Sysdate  , �ϴ��޸���Ա = �޸���Ա, �ϴ��޸�ʱ�� = �޸�ʱ��  where ID = " & strID
    gcnOracle.Execute strSQL
    
    gcnOracle.CommitTrans
    SaveProc = True
    Exit Function
errH:
    If InStr(1, UCase(Err.Description), "ORA") Then
        gcnOracle.RollbackTrans
    End If
    strErr = Err.Description
End Function

Private Sub txtLeft_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim lngMinBound As Long, lngMaxBound As Long
    Dim strTmp As String

    lngMinBound = (txtLeft.Font.Size + 3) * (txtLeft.CurrPos.Row - txtLeft.TopRow - 1.5)
    lngMaxBound = (txtLeft.Font.Size + 3) * (txtLeft.CurrPos.Row - txtLeft.TopRow + 1.5)

    If y > lngMinBound And y < lngMaxBound Then
        strTmp = FindOldText(txtLeft.CurrPos.Row, mcolDiff1)
        If strTmp <> "" Then
            strTmp = RPAD(strTmp, 70)    '�����Rpad��Ϊ�˱�֤��Ϣ����ȫ��ʾ
        End If
        If strTmp = "" Then
            ShowTipInfo pctBottom.hwnd, ""
        Else
            ShowTipInfo pctBottom.hwnd, strTmp, , , , mstrTipLeft, True
        End If
    Else
        ShowTipInfo pctBottom.hwnd, ""
    End If
End Sub

Private Sub vsfIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strTip As String, strTmp As String
    Dim i As Long
    
    On Error Resume Next
    
    With vsfIcon
        If .MouseRow = -1 Then Exit Sub
        strTip = .Cell(flexcpData, .MouseRow, 0)
        
        If strTip = "" Then
            ShowTipInfo vsfIcon.hwnd, ""
        Else
            strTip = IIf(strTip = "�տ�", "", strTip)
            For i = .MouseRow + 1 To .Rows - 1
                strTmp = .Cell(flexcpData, i, 0)
                If strTmp <> "" Then
                    strTip = strTip & vbNewLine & IIf(strTmp = "�տ�", "", strTmp)
                Else
                    Exit For
                End If
            Next
            If ConvertStr(strTip) <> "" Then
                If txtRight.RowText(.MouseRow) = "" Then
                    ShowTipInfo vsfIcon.hwnd, RPAD(strTip, 70), , , , "�޸�����:ɾ��    " & mstrTipright
                Else
                    ShowTipInfo vsfIcon.hwnd, RPAD(strTip, 70), , , , "�޸�����:�޸�    " & mstrTipright
                End If
            Else
                ShowTipInfo vsfIcon.hwnd, RPAD("������", 70), , , , "�޸�����:����     " & mstrTipright
            End If
        End If
        
    End With
End Sub

Private Sub txtRight_KeyDown(KeyCode As Integer, Shift As Integer)
'��Ϊ���뵼��ɾ������������,���Ҫ�ڼ��̰���ʱ��¼�仯֮ǰ���к�
    mblnChanged = True
    mlngRows = txtRight.RowsCount
    mlngCurCol = txtRight.CurrPos.Col
    mlngCurRow = txtRight.CurrPos.Row
End Sub

Private Sub txtRight_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
     Dim i As Long, j As Long
     Dim arrTmp(2) As String, strTmp As String
     Dim blnExit As Boolean
     
    txtRight.SetRowBkColor mlngRightRow, ��ɫ
    txtRight.SetRowBkColor nNewRow, ����ɫ
    mlngRightRow = nNewRow
     
    '��������
    TxtLink 2
    
    '���������仯,ɾ��������������
     With txtRight
        If vsfIcon.Rows = 0 Then Exit Sub
        If Not mblnChanged Then Exit Sub
        If .RowsCount = mlngRows Then Exit Sub  '����δ��
        
        mblnChanged = False
        
        If .RowsCount > mlngRows Then    '��������
        
            If mlngCurCol = 1 And nNewCol = 1 Then  '��һ�仰�Ŀ�ͷ���»س�,���������ж�
                For i = 1 To .RowsCount - mlngRows
                    vsfIcon.AddItem "", mlngCurRow - 1
                Next
            Else '�����λ�ð��»س� ,�����кŲ���
                For i = 1 To .RowsCount - mlngRows
                    vsfIcon.AddItem "", mlngCurRow + 1
                Next
            End If
            mlngRows = .RowsCount
        End If
        
        If .RowsCount < mlngRows Then   'ɾ������
            
            If mlngCurCol = 1 And nNewCol = 1 Then  '��һ�仰�Ŀ�ͷ���»���,���������ж�
            For i = 1 To mlngRows - .RowsCount
                    vsfIcon.RemoveItem mlngCurRow - 1
                Next
            Else '�����λ�ð��»��� ,�����кŲ���
                For i = 1 To mlngRows - .RowsCount
                    vsfIcon.RemoveItem mlngCurRow + 1
                Next
            End If
            mlngRows = .RowsCount
        End If
     End With
     

End Sub

Private Sub txtLeft_CurPosChanged(ByVal nNewRow As Long, ByVal nNewCol As Long)
    '���к�,������ı���ѡ�������
    Dim i As Long, j As Long
    Dim arrTmp(2) As String, strTmp As String
    
    TxtLink 1
    
    txtLeft.SetRowBkColor mlngLeftRow, ��ɫ
    txtLeft.SetRowBkColor nNewRow, ����ɫ
    mlngLeftRow = nNewRow

End Sub


Private Sub txtLeft_GotFocus()
    mintLast = 1
End Sub

Private Sub txtRight_GotFocus()
    mintLast = 2
End Sub

Private Sub txtLeft_LostFocus()
    txtLeft.SetRowBkColor mlngLeftRow, ��ɫ
End Sub

Private Sub txtRight_LostFocus()
    txtRight.SetRowBkColor mlngRightRow, ��ɫ
End Sub

Private Sub TxtLink(ByVal intType As Integer)
    '����:����ؼ���,���ݿؼ��ĵ�ǰ��,ʹ��һ���ؼ��Զ���λ
    'intType=1 �����߿ؼ�   intType= 2 ����ұ߿ؼ�
    Dim i As Long, strTmp As String
    Dim lngCurRowLeft As Long, lngCurRowRight As Long
    Dim arrTmp(2) As String, lngPageRows As Long
    Dim lngUpLenth As Long, lngDownLenth As Long
    
    lngCurRowLeft = txtLeft.CurrPos.Row
    lngCurRowRight = txtRight.CurrPos.Row
    lngPageRows = txtLeft.Height / lngRowHeight
    
    If txtLeft.RowsCount < lngPageRows Then Exit Sub    '�������С��һҳ���������,�Ͳ���Ҫ����
    If ConvertStr(txtLeft.RowText(lngCurRowLeft)) = ConvertStr(txtRight.RowText(lngCurRowRight)) Then Exit Sub  '��ǰ�о��˳�
    
    If intType = 1 Then
        '������ؼ�
        With txtLeft
            '�����ϲ���
            For i = lngCurRowLeft To 3 Step -1
                If .RowText(i) <> "" And GetValueFromCol(mcolDiff1, "_" & i) = "" And _
                .RowText(i - 1) <> "" And GetValueFromCol(mcolDiff1, "_" & i - 1) = "" And _
                .RowText(i - 2) <> "" And GetValueFromCol(mcolDiff1, "_" & i - 2) = "" Then
                    arrTmp(0) = ConvertStr(.RowText(i - 2))
                    arrTmp(1) = ConvertStr(.RowText(i - 1))
                    arrTmp(2) = ConvertStr(.RowText(i))
                    lngUpLenth = lngCurRowLeft - i
                    Exit For
                End If
            Next
            
            '�����²���
            If arrTmp(0) = "" Or lngUpLenth > lngPageRows / 2 Then
                For i = lngCurRowLeft To txtLeft.RowsCount - 3
                    If .RowText(i) <> "" And GetValueFromCol(mcolDiff1, "_" & i) = "" And _
                    .RowText(i + 1) <> "" And GetValueFromCol(mcolDiff1, "_" & i + 1) = "" And _
                    .RowText(i + 2) <> "" And GetValueFromCol(mcolDiff1, "_" & i + 2) = "" Then
                        arrTmp(0) = ConvertStr(.RowText(i))
                        arrTmp(1) = ConvertStr(.RowText(i + 1))
                        arrTmp(2) = ConvertStr(.RowText(i + 2))
                        lngDownLenth = i - lngCurRowLeft
                        Exit For
                    End If
                Next
            End If
            
            If lngUpLenth <> 0 And lngDownLenth <> 0 Then   '�Ϸ��·��������Ϊ0,����Ҫȡ����϶̵�
                If lngUpLenth > lngDownLenth Then
                    arrTmp(0) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth))
                    arrTmp(1) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth + 1))
                    arrTmp(2) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth + 2))
                Else
                    arrTmp(0) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth - 2))
                    arrTmp(1) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth - 1))
                    arrTmp(2) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth))
                End If
            End If
            
            If arrTmp(0) <> "" Then
                strTmp = arrTmp(0) & arrTmp(1) & arrTmp(2)
                If InStr(1, ConvertStr(txtRight.Text), strTmp) > 0 Then
                    For i = 1 To txtRight.RowsCount - 3
                        If ConvertStr(txtRight.RowText(i)) = arrTmp(0) And ConvertStr(txtRight.RowText(i + 1)) = arrTmp(1) _
                            And ConvertStr(txtRight.RowText(i + 2)) = arrTmp(2) Then
                            txtRight.TopRow = i - (.CurrPos.Row - .TopRow)
                            Exit For
                        End If
                    Next
                End If
            End If
        End With
        
    Else
        '����Ҳ�ؼ�
        With txtRight
            '�����ϲ���
            For i = lngCurRowRight To 3 Step -1
                If .RowText(i) <> "" And .RowText(i - 1) <> "" And .RowText(i - 2) <> "" Then
                    arrTmp(0) = ConvertStr(.RowText(i - 2))
                    arrTmp(1) = ConvertStr(.RowText(i - 1))
                    arrTmp(2) = ConvertStr(.RowText(i))
                    lngUpLenth = lngCurRowRight - i
                    Exit For
                End If
            Next
            
            '�����²���
            If arrTmp(0) = "" Or lngUpLenth > lngPageRows / 2 Then
                For i = lngCurRowLeft To .RowsCount - 3
                    If .RowText(i) <> "" And .RowText(i + 1) <> "" And .RowText(i + 2) <> "" Then
                        arrTmp(0) = ConvertStr(.RowText(i))
                        arrTmp(1) = ConvertStr(.RowText(i + 1))
                        arrTmp(2) = ConvertStr(.RowText(i + 2))
                        lngDownLenth = i - lngCurRowRight
                        Exit For
                    End If
                Next
            End If
            
            If lngUpLenth <> 0 And lngDownLenth <> 0 Then   '�Ϸ��·��������Ϊ0,����Ҫȡ����϶̵�
                If lngUpLenth > lngDownLenth Then
                    arrTmp(0) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth))
                    arrTmp(1) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth + 1))
                    arrTmp(2) = ConvertStr(.RowText(lngCurRowLeft + lngDownLenth + 2))
                Else
                    arrTmp(0) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth - 2))
                    arrTmp(1) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth - 1))
                    arrTmp(2) = ConvertStr(.RowText(lngCurRowLeft - lngUpLenth))
                End If
            End If
            
            If arrTmp(0) <> "" Then
                strTmp = arrTmp(0) & arrTmp(1) & arrTmp(2)
                If InStr(1, ConvertStr(txtLeft.Text), strTmp) > 0 Then
                    For i = 1 To txtLeft.RowsCount - 3
                        If ConvertStr(txtLeft.RowText(i)) = arrTmp(0) And ConvertStr(txtLeft.RowText(i + 1)) = arrTmp(1) _
                            And ConvertStr(txtLeft.RowText(i + 2)) = arrTmp(2) Then
                            txtLeft.TopRow = i - (.CurrPos.Row - .TopRow)
                            Exit For
                        End If
                    Next
                End If
            End If
        End With
    End If
    
End Sub
