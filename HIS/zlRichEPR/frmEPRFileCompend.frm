VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~1.OCX"
Object = "*\A..\zlRichEditor\zlRichEdit.vbp"
Begin VB.Form frmEPRFileContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "�����ļ����"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   2175
      ScaleHeight     =   2985
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   3420
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4170
         Left            =   0
         TabIndex        =   1
         Top             =   810
         Width           =   7500
         _cx             =   13229
         _cy             =   7355
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   2
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   2970
         TabIndex        =   3
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
   End
   Begin zlRichEditor.Editor edtThis 
      Height          =   2580
      Left            =   315
      TabIndex        =   4
      Top             =   225
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4551
      WithViewButtonas=   0   'False
      ShowRuler       =   0   'False
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRFileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum zlEnumCompendParentKind     '��ٸ�����
    cprEmCPKFileDefine = 0              '�ļ���������
    cprEmCPKModelEssay = 1              '��������
End Enum

'-----------------------------------------------------
'�����¼�
'-----------------------------------------------------
Public Event DblClick()                                                 '����˫�������¼�

'-----------------------------------------------------
'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Me.edtThis.Copy
    End Select
End Sub

Private Sub Form_Load()
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.edtThis
        .Left = Me.ScaleLeft + 90: .Width = Me.ScaleWidth - 2 * .Left
        .Top = Me.ScaleTop + 90: .Height = Me.ScaleHeight - 2 * .Top
    End With
    With Me.picTab
        .Left = Me.ScaleLeft + 90: .Width = Me.ScaleWidth - 2 * .Left
        .Top = Me.ScaleTop + 90: .Height = Me.ScaleHeight - 2 * .Top
    End With
End Sub

Private Sub picTab_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move Me.picTab.ScaleLeft, Me.picTab.ScaleTop + 120, Me.picTab.ScaleWidth
    Me.lblSubhead.Move Me.picTab.ScaleLeft + 210, Me.lblTitle.Top + Me.lblTitle.Height + 120
    Me.vfgThis.Move Me.picTab.ScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, Me.picTab.ScaleWidth - 210 * 2
    Me.vfgThis.Height = Me.picTab.ScaleHeight - Me.vfgThis.Top - 210
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, x As Single, y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Popup.ShowPopup
    End With
End Sub

'-----------------------------------------------------
'���幫������
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngParentId As Long, Optional bytParentKind As zlEnumCompendParentKind = cprEmCPKFileDefine)
    '���ܣ���ʾָ���ļ�/���ĵ����ݣ�
    Dim strTemp As String, strZipFile As String
    Dim rsTemp As New ADODB.Recordset
    
    Me.edtThis.Visible = True
    Me.picTab.Visible = False
    
    If lngParentId = 0 Then Me.edtThis.ReadOnly = False: Me.edtThis.NewDoc: Me.edtThis.ReadOnly = True: Exit Sub
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    Me.edtThis.Freeze
    If lngParentId > 0 Then
        '����ҳ���ʽ
        Dim mEPRFileInfo As New cEPRFileDefineInfo
        If bytParentKind = cprEmCPKFileDefine Then
            '�����ļ���ʽ
            gstrSQL = "Select b.ID, a.��ʽ From ����ҳ���ʽ a, �����ļ��б� b " & _
                " Where  a.���� = b.���� And a.��� = b.ҳ�� And b.ID = [1]"
        Else
            '�������ĸ�ʽ
            gstrSQL = "Select C.ID, a.��ʽ From ����ҳ���ʽ a, �����ļ��б� b, ��������Ŀ¼ c " & _
                " Where  c.�ļ�ID = b.ID And b.���� = a.���� And b.ҳ�� = a.��� And c.ID = [1]"
        End If
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If Not rsTemp.EOF Then
            mEPRFileInfo.��ʽ = "" & rsTemp!��ʽ
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.��ʽ
        End If
        Set mEPRFileInfo = Nothing
    End If
    Me.edtThis.UnFreeze
    Me.edtThis.ResetWYSIWYG
    If bytParentKind = cprEmCPKFileDefine Then
        '�����ļ���ʽ
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select l.����,l.���� From �����ļ��б� l Where l.Id = [1]"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If Val("" & rsTemp!����) < 0 And rsTemp!���� <> 6 Then
            With Me.edtThis
                .Freeze
                .Text = vbCrLf & Space(4) & "���ļ�Ϊ�����ʽ���������������ʽ..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "����": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
                .UnFreeze
            End With
        ElseIf rsTemp!���� <> 3 Then
            strZipFile = zlBlobRead(1, lngParentId)
            If Len(strZipFile) > 0 Then
                If gobjFSO.FileExists(strZipFile) Then
                    strTemp = zlFileUnzip(strZipFile)
                    If gobjFSO.FileExists(strTemp) Then
                        Me.edtThis.OpenDoc strTemp
                        gobjFSO.DeleteFile strTemp, True
                    End If
                    gobjFSO.DeleteFile strZipFile, True
                End If
            End If
        Else
            Me.edtThis.Visible = False
            Me.picTab.Visible = True
            
            Dim lngCurColor As Long, strCurFont As String, objFont As StdFont
            Me.lblTitle.Caption = "": Me.lblSubhead.Caption = ""
            Me.vfgThis.Redraw = flexRDNone
            Me.vfgThis.Clear: Me.vfgThis.MergeCells = flexMergeFree
            Me.vfgThis.MergeRow(0) = True: Me.vfgThis.MergeRow(1) = True
            
            gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
                " Order By d.�������"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Select Case "" & !Ҫ������
                    Case "��ͷ����"
                        If Val("" & !�����ı�) = 1 Then
                            Me.vfgThis.RowHidden(0) = True
                        Else
                            Me.vfgThis.RowHidden(0) = False
                        End If
                    Case "������":  Me.vfgThis.Cols = Val("" & !�����ı�)
                    Case "��С�и�": Me.vfgThis.RowHeightMin = Val("" & !�����ı�)
                    Case "�ı�����"
                        strCurFont = "" & !�����ı�
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                        End With
                        Set Me.vfgThis.Font = objFont
                        Set Me.lblSubhead.Font = Me.vfgThis.Font
                        
                    Case "�ı���ɫ": Me.vfgThis.ForeColor = Val("" & !�����ı�)
                    Case "�����ɫ": Me.vfgThis.GridColor = Val("" & !�����ı�): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                    
                    Case "�����ı�": Me.lblTitle.Caption = "" & !�����ı�
                    Case "��������"
                        strCurFont = "" & !�����ı�
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                        End With
                        Set Me.lblTitle.Font = objFont
                        Me.lblTitle.AutoSize = False
                    End Select
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
                " Order By d.�������"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubhead.Caption = ""
                Do While Not .EOF
                    Me.lblSubhead.Caption = Me.lblSubhead.Caption & " " & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
                    .MoveNext
                Loop
                Me.lblSubhead.Caption = Trim(Me.lblSubhead.Caption)
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
                " Order By d.�������"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.TextMatrix(!�����д� - 1, !������� - 1) = "" & !�����ı�
                    Me.vfgThis.FixedAlignment(!������� - 1) = flexAlignCenterCenter
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ" & _
                " From �����ļ��ṹ d, �����ļ��ṹ p" & _
                " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
                " Order By d.�������, d.�����д�"
            Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.ColWidth(!������� - 1) = Val("" & !��������)
                    .MoveNext
                Loop
            End With
            Me.vfgThis.Redraw = flexRDDirect
                    
            '---------------------------------------------------
            Call picTab_Resize
        End If
    Else
        '������������
        strZipFile = zlBlobRead(3, lngParentId)
        If Len(strZipFile) > 0 Then
            If gobjFSO.FileExists(strZipFile) Then
                strTemp = zlFileUnzip(strZipFile)
                If gobjFSO.FileExists(strTemp) Then
                    Me.edtThis.OpenDoc strTemp
                    gobjFSO.DeleteFile strTemp, True
                End If
                gobjFSO.DeleteFile strZipFile, True
            End If
        End If
    End If
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
