VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileRetry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8310
   Icon            =   "frmTendFileRetry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8310
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraMain 
      Height          =   5190
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   8295
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��"
         Height          =   330
         Left            =   7185
         TabIndex        =   3
         Top             =   4770
         Width           =   900
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��"
         Height          =   330
         Left            =   6165
         TabIndex        =   2
         Top             =   4770
         Width           =   900
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFile 
         Height          =   4575
         Left            =   15
         TabIndex        =   1
         Top             =   105
         Width           =   8160
         _cx             =   14393
         _cy             =   8070
         Appearance      =   2
         BorderStyle     =   0
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
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
         Begin VB.CheckBox chkChoose 
            Height          =   165
            Left            =   660
            TabIndex        =   4
            Top             =   60
            Width           =   180
         End
      End
      Begin MSComctlLib.ImageList imgList 
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
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTendFileRetry.frx":06EA
               Key             =   "���µ�"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTendFileRetry.frx":0DFC
               Key             =   "��¼��"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTendFileRetry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Enum mCol
    f��־ = 0: fID: ѡ��: f�ļ�����: f��ʼʱ��: f����ʱ��: �����ļ�: �����ļ�id
End Enum
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintBaby As Integer
Private mblnSave As Boolean           '�Ƿ�����

Public Function ShowMe(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBabby As Integer, ByVal strPrivs As String) As Boolean
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintBaby = intBabby
    mblnSave = False
    Call zlRefData
    Me.Show 1
    ShowMe = mblnSave
End Function

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim intRow As Integer
    Dim lngID As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    Err = 0
    On Error GoTo ErrHand
    '------------------------------------------------------------------------------------------------------------------
    '�����ļ�ˢ��
    
    With vsfFile
        .Clear
        .Rows = 1
        .Cols = 8
        .FixedCols = 1
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .TextMatrix(0, mCol.f��־) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f�ļ�����) = "�ļ�����"
        .TextMatrix(0, mCol.f��ʼʱ��) = "��ʼʱ��"
        .TextMatrix(0, mCol.f����ʱ��) = "����ʱ��"
        .TextMatrix(0, mCol.�����ļ�) = "�����¼��"
        .TextMatrix(0, mCol.�����ļ�id) = "�����ļ�id"
        
        
        .ColDataType(mCol.ѡ��) = flexDTBoolean
        .ColWidth(mCol.f��־) = 270: .ColWidth(mCol.ѡ��) = 500
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f�ļ�����) = 2200: .ColWidth(mCol.f��ʼʱ��) = 1800
        .ColWidth(mCol.f����ʱ��) = 1800: .ColWidth(mCol.�����ļ�) = 2200: .ColWidth(mCol.�����ļ�id) = 0
    End With
    
    intRow = vsfFile.FixedRows
    '--------------------------------------------------------------------------------------------------------------
    strSQL = " Select A.ID,A.�ļ�����, B.���� AS �ļ���Դ,B.����,A.��ʼʱ��,A.����ʱ��,A.������,A.����ʱ��,A.�鵵��,C.�ļ����� AS �����ļ�,C.ID AS �����ļ�ID,B.���� " & _
              " From ���˻����ļ� A,�����ļ��б� B,���˻����ļ� C" & _
              " Where A.��ʽID=B.ID And A.����ID=[1] And A.��ҳID=[2] And A.Ӥ��=[3] And A.����ID=C.ID(+) and B.����=0" & _
              " Order by B.����,A.��ʼʱ��"
    Call SQLDIY(strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ʾָ�����˵Ļ����ļ��б�", mlng����ID, mlng��ҳID, mintBaby)
    rsTemp.Filter = 0
    rsTemp.Sort = "��ʼʱ��"
    With Me.vsfFile
        Do While Not rsTemp.EOF
            vsfFile.Rows = vsfFile.Rows + 1
            i = vsfFile.Rows
            If Val(.TextMatrix(i - 1, mCol.fID)) > 0 Then .AddItem ""
            
            lngID = Val(NVL(rsTemp!ID, 0))
            .TextMatrix(i - 1, mCol.fID) = lngID
            Set .Cell(flexcpPicture, i - 1, mCol.f��־) = imgList.ListImages("��¼��").Picture
            .TextMatrix(i - 1, mCol.f�ļ�����) = NVL(rsTemp!�ļ�����)
            .TextMatrix(i - 1, mCol.f��ʼʱ��) = Format(NVL(rsTemp!��ʼʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(i - 1, mCol.f����ʱ��) = Format(NVL(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
            .TextMatrix(i - 1, mCol.�����ļ�) = NVL(rsTemp!�����ļ�)
            .TextMatrix(i - 1, mCol.�����ļ�id) = NVL(rsTemp!�����ļ�id)
            
            rsTemp.MoveNext
        Loop
    End With
    With chkChoose
        .Value = 0
        .Top = vsfFile.Top - (vsfFile.RowHeight(0) - .Height) / 2
        .Left = vsfFile.ColWidth(mCol.f��־) + (vsfFile.ColWidth(mCol.ѡ��) - .Width) / 2 - vsfFile.Left
    End With
    
    'ѡ����
    Call vsfFile.Select(intRow, mCol.fID)
    
    zlRefData = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub Allcheck()
    Dim i As Integer
    If chkChoose.Value = Checked Then
        For i = 1 To vsfFile.Rows - 1
            vsfFile.Cell(flexcpChecked, i, mCol.ѡ��, i, mCol.ѡ��) = flexChecked
        Next
    Else
        For i = 1 To vsfFile.Rows - 1
            vsfFile.Cell(flexcpChecked, i, mCol.ѡ��, i, mCol.ѡ��) = flexUnchecked
        Next
    End If
End Sub

Private Sub chkChoose_Click()
    Call Allcheck
End Sub

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngFileID As Long
    Dim lng����ID As Long
    Dim intCountRetry As Integer
    Dim lngLoop As Long, j As Long
    Dim strFileID As String
    Dim blnCheck As Boolean
    
    lngFileID = 0
    lng����ID = 0
    intCountRetry = 0
    strFileID = ""
    
    On Error GoTo ErrHand
    
    blnCheck = False
    For lngLoop = 0 To vsfFile.Rows - 1
        If vsfFile.Cell(flexcpChecked, lngLoop, mCol.ѡ��) = flexChecked Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then Exit Sub
    
    If MsgBox("���㽫��Ե�ǰѡ�еļ�¼���ļ��Լ���ؼ�¼���ļ���������������,���һ���ݲ���:�����ļ�ҳ�����,�Ե�ǰѡ�м�¼���ļ���֮��ļ�¼" & _
            "���ļ�����ҳ������,�˲������������¼���ļ��Ĵ�ӡ��Ϣ��" & vbCrLf & "�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Screen.MousePointer = 11
        For lngLoop = 0 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, lngLoop, mCol.ѡ��) = flexChecked Then
                lngFileID = Val(vsfFile.TextMatrix(lngLoop, mCol.fID))
                lng����ID = Val(vsfFile.TextMatrix(lngLoop, mCol.�����ļ�id))
                If lngFileID > 0 And Not InStr(1, strFileID & ",", "," & lngFileID & ",") > 0 Then
                    If frmTendFilePreview.AnaliseData(Me, lngFileID, mlng����ID, mlng��ҳID) Then intCountRetry = intCountRetry + 1
                    strFileID = strFileID & "," & lngFileID
                    If lng����ID <> 0 And Not chkChoose.Value = Checked And Not InStr(1, strFileID & ",", "," & lng����ID & ",") > 0 Then
                        For j = 0 To vsfFile.Rows - 1
                            If frmTendFilePreview.AnaliseData(Me, lng����ID, mlng����ID, mlng��ҳID) Then intCountRetry = intCountRetry + 1
                            strFileID = strFileID & "," & lng����ID
                            lng����ID = CheckContinue(lng����ID)
                            If lng����ID = 0 Then Exit For
                        Next j
                    End If
                End If
            End If
        Next lngLoop
        For lngLoop = 0 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, lngLoop, mCol.ѡ��) = flexChecked Then
                Exit For
            End If
        Next lngLoop
        For j = lngLoop To vsfFile.Rows - 1
            lngFileID = Val(vsfFile.TextMatrix(j, mCol.fID))
            If lngFileID > 0 Then
                gstrSQL = "Zl_���˻����ӡ_Batchretrypage(" & lngFileID & ",'1;1',0)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "ҳ������")
            End If
        Next
        
        Screen.MousePointer = 0
        mblnSave = True
        MsgBox "������" & intCountRetry & "�ݼ�¼���ļ���", vbInformation, gstrSysName
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Form_Resize()
    fraMain.Move 0, -90, Me.ScaleWidth, Me.ScaleHeight + 90
    vsfFile.Move 15, 105, vsfFile.Width, fraMain.Height - 105 - cmdOK.Height - 120 * 2
    cmdOK.Move 6015, vsfFile.Height + vsfFile.Top + 120
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 165, cmdOK.Top
End Sub

Private Function CheckContinue(ByVal FileID As Long) As Long
    '���� : ��������id��������id,û�з���0
    Dim i As Integer
    Dim lng����ID As Long
    
    lng����ID = 0
    For i = 0 To vsfFile.Rows - 1
        If Val(vsfFile.TextMatrix(i, mCol.fID)) = FileID Then
            lng����ID = Val(vsfFile.TextMatrix(i, mCol.�����ļ�id))
            Exit For
        End If
    Next i
    CheckContinue = lng����ID
    
End Function

Private Sub vsfFile_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If NewLeftCol <= mCol.ѡ�� Then
        chkChoose.Visible = True
    Else
        chkChoose.Visible = False
    End If

End Sub

Private Sub vsfFile_AfterUserResize(ByVal ROW As Long, ByVal COL As Long)
    If COL <= mCol.ѡ�� Then
        With chkChoose
            .Value = 0
            .Top = vsfFile.Top - (vsfFile.RowHeight(0) - .Height) / 2
            .Left = vsfFile.ColWidth(mCol.f��־) + (vsfFile.ColWidth(mCol.ѡ��) - .Width) / 2 - vsfFile.Left
        End With
    End If
End Sub

Private Sub vsfFile_Click()
    Call vsfFile_DblClick
End Sub

Private Sub vsfFile_DblClick()
    Dim i As Integer
    Dim blnAllChoose As Boolean
    If vsfFile.COL <> mCol.ѡ�� Then Exit Sub
    If vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.ѡ��) = flexUnchecked Then
        vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.ѡ��, vsfFile.ROW, mCol.ѡ��) = flexChecked
        blnAllChoose = True
        For i = 1 To vsfFile.Rows - 1
            If vsfFile.Cell(flexcpChecked, i, mCol.ѡ��, i, mCol.ѡ��) <> flexChecked Then
                blnAllChoose = False
                Exit For
            End If
        Next
        If blnAllChoose = True Then chkChoose.Value = Checked
    Else
        If chkChoose.Value = Checked Then
            chkChoose.Value = Unchecked
            For i = 1 To vsfFile.Rows - 1
                If i <> vsfFile.ROW Then
                    vsfFile.Cell(flexcpChecked, i, mCol.ѡ��, i, mCol.ѡ��) = flexChecked
                End If
            Next
        Else
           vsfFile.Cell(flexcpChecked, vsfFile.ROW, mCol.ѡ��, vsfFile.ROW, mCol.ѡ��) = flexUnchecked
        End If
        
    End If
End Sub



