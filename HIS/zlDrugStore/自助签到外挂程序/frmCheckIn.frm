VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCheckIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ǩ��ϵͳ"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12180
   Icon            =   "frmCheckIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12180
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Frame frmLineBruchButtom 
      Height          =   120
      Left            =   0
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
      Width           =   12165
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   20
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer TimerAuto 
      Interval        =   60000
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtCard 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   330
      Width           =   4095
   End
   Begin VB.Frame frmLineTop 
      Height          =   120
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   11805
   End
   Begin VB.PictureBox picBrush 
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   120
      ScaleHeight     =   6975
      ScaleWidth      =   12015
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   12015
      Begin VB.Frame frmLineBrushTop 
         Height          =   120
         Left            =   -120
         TabIndex        =   10
         Top             =   720
         Width           =   12045
      End
      Begin VB.CommandButton cmdCheckIn 
         Caption         =   "ǩ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   9
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdCheckInAll 
         Caption         =   "ȫ��ǩ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   8
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "��һ�Ŵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   7
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "��һ�Ŵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10080
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4155
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   11895
         _cx             =   20981
         _cy             =   7329
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCheckIn.frx":030A
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
         ExplorerBar     =   7
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
         VirtualData     =   0   'False
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
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   8640
         TabIndex        =   19
         Top             =   6360
         Width           =   3225
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "��������ǩ��-->"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   6405
         Width           =   3465
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         TabIndex        =   16
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   15
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblRecipeComment 
         AutoSize        =   -1  'True
         Caption         =   "�ܼ�2�Ŵ���δǩ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   8040
         TabIndex        =   14
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label lblBill 
         AutoSize        =   -1  'True
         Caption         =   "��Ʊ�ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1740
      End
      Begin VB.Label lblNo 
         AutoSize        =   -1  'True
         Caption         =   "�����ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   12
         Top             =   1140
         Width           =   1740
      End
   End
   Begin VB.PictureBox picUnBrush 
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7215
      ScaleWidth      =   11895
      TabIndex        =   3
      Top             =   1080
      Width           =   11895
      Begin VB.Label lblCommen 
         AutoSize        =   -1  'True
         Caption         =   "��ӭʹ������ǩ��ϵͳ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   3240
         TabIndex        =   4
         Top             =   6480
         Width           =   4050
      End
      Begin VB.Image imgHos 
         Height          =   6075
         Left            =   120
         Picture         =   "frmCheckIn.frx":03F2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   15015
      End
   End
   Begin VB.Label lblCard 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��ˢ���￨"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1665
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mrsRecipeList As ADODB.Recordset
Dim mstrCardPasswordRule As String          '�����Ĺ���
Private mlngҩ��id As Long
Private mBlnBegin As Boolean

Private Sub cmdBack_Click()
    picBrush.Visible = False
    picUnBrush.Visible = True
    Me.cmdBack.Visible = False
    Me.frmLineBruchButtom.Visible = False
    Me.txtCard.Text = ""
End Sub

Private Sub cmdCheckIn_Click()
    Dim strSQL As String
    
    On Error GoTo errRow
    strSQL = "Zl_δ��ҩƷ��¼_��ҩȷ��("
    strSQL = strSQL & "'" & mrsRecipeList!no & "',"
    strSQL = strSQL & mrsRecipeList!���� & ","
    strSQL = strSQL & mlngҩ��id & ","
    strSQL = strSQL & "1,"
    strSQL = strSQL & "'auto')"
    Call zldatabase.ExecuteProcedure(strSQL, "cmdCheckIn_Click")
    
    If mrsRecipeList.RecordCount = 1 Then
        Me.picBrush.Visible = False
        Me.picUnBrush.Visible = True
        Me.lblCommen.Caption = "ǩ���ɹ���"
        Me.txtCard.Text = ""
        Set mrsRecipeList = Nothing
    Else
        Me.lblMsg.Caption = "����[" & mrsRecipeList!no & "]ǩ���ɹ�"

        Call mrsRecipeList.Delete(adAffectCurrent)
        Me.lblRecipeComment.Caption = " �ܼ�" & mrsRecipeList.RecordCount & "�Ŵ���δǩ��"
        
        If mrsRecipeList.RecordCount = 1 Then
            Me.cmdNext.Visible = False
            Me.cmdPrevious.Visible = False
            Me.cmdCheckInAll.Visible = False
        End If
        
        If Me.cmdNext.Enabled Then
            Call cmdNext_Click
            Me.cmdPrevious.Enabled = False
        Else
            Call cmdPrevious_Click
            Me.cmdNext.Enabled = False
        End If
    End If
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCheckInAll_Click()
    Dim strSQL As String
    Dim i As Integer
    
    On Error GoTo errRow
    With mrsRecipeList
        .MoveFirst
        For i = 1 To .RecordCount
            strSQL = "Zl_δ��ҩƷ��¼_��ҩȷ��("
            strSQL = strSQL & "'" & !no & "',"
            strSQL = strSQL & !���� & ","
            strSQL = strSQL & mlngҩ��id & ","
            strSQL = strSQL & "1,"
            strSQL = strSQL & "'auto')"
            Call zldatabase.ExecuteProcedure(strSQL, "cmdCheckInAll_Click")
            
            .MoveNext
        Next
    End With
    
    Me.picBrush.Visible = False
    Me.picUnBrush.Visible = True
    Me.lblCommen.Caption = "ǩ���ɹ���"
    Me.txtCard.Text = ""
    Me.cmdBack.Visible = False
    Set mrsRecipeList = Nothing
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdNext_Click()
    Me.cmdPrevious.Enabled = True
    
    With mrsRecipeList
        If Not .EOF Then .MoveNext
        If Not .EOF Then
            Me.lblBill.Caption = "��Ʊ�ţ�" & !����
            Me.lblNo.Caption = "�����ţ�" & !no
            
            Call GetDetail(!no)
            
            .MoveNext
            If .EOF Then
                Me.cmdNext.Enabled = False
            End If
            .MovePrevious
        End If
    End With
End Sub

Private Sub cmdPrevious_Click()
    Me.cmdNext.Enabled = True
     
    With mrsRecipeList
        If Not .BOF Then .MovePrevious
        If Not .BOF Then
            Me.lblBill.Caption = "��Ʊ�ţ�" & !����
            Me.lblNo.Caption = "�����ţ�" & !no
            Call GetDetail(!no)
            
            .MovePrevious
            If .BOF Then
                Me.cmdPrevious.Enabled = False
            End If
            .MoveNext
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mBlnBegin And KeyAscii = 13 Then
        Exit Sub
    Else
        If mBlnBegin Then
            txtCard.Text = ""
            mBlnBegin = False
        End If
        Call txtCard_KeyPress(KeyAscii)
    End If
    
End Sub

Private Sub Form_Load()
    Me.Caption = "����ǩ��ϵͳ" & "(" & gstrStockName & ")"
        
End Sub

Private Sub IniRecord()
    '��ʼ�����ݼ�
    Set mrsRecipeList = New ADODB.Recordset
    
    With mrsRecipeList
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�Ա�", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "no", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adDouble, 2, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function GetList(ByVal strText As String) As String
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    On Error GoTo errRow
    
    IniRecord
    strSQL = "Select a.����,a.no,b.����, b.�Ա�, b.����, d.����, e.����" & vbNewLine & _
            "From δ��ҩƷ��¼ a, ������Ϣ b, Ʊ�ݴ�ӡ���� c, Ʊ��ʹ����ϸ d, ����ҽ�ƿ���Ϣ e" & vbNewLine & _
            "Where a.No = c.No And a.����id = b.����id And a.����id = e.����id And c.Id = d.��ӡid And c.�������� = 1 And d.Ʊ�� = 1 And e.״̬ = 0 And" & vbNewLine & _
            "      Nvl(�Ŷ�״̬, 0) = 0 And e.���� = [1] and a.�ⷿid=[2]"
    
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "GetList", strText, mlngҩ��id)
    
    With rsTemp
        Do While Not .EOF
            mrsRecipeList.AddNew
            mrsRecipeList!���� = !����
            mrsRecipeList!���� = !����
            mrsRecipeList!no = !no
            mrsRecipeList!�Ա� = !�Ա�
            mrsRecipeList!���� = !����
            mrsRecipeList!���� = !����
            mrsRecipeList!���� = !����
            mrsRecipeList.Update
            
            .MoveNext
        Loop
    End With
    
    If rsTemp.RecordCount > 0 Then
        mrsRecipeList.MoveFirst
        With mrsRecipeList
            lblRecipeComment.Caption = "�ܼ�" & .RecordCount & "�Ŵ���δǩ��"
            If .RecordCount > 1 Then
                Me.cmdCheckInAll.Visible = True
                Me.cmdNext.Visible = True
                Me.cmdPrevious.Visible = True
                cmdPrevious.Enabled = False
            Else
                Me.cmdCheckInAll.Visible = False
                Me.cmdNext.Visible = False
                Me.cmdPrevious.Visible = False
            End If
        
            If Not .EOF Then
                lblName.Caption = !����
                lblSex.Caption = !�Ա�
                lblAge.Caption = !����
                lblBill.Caption = "��Ʊ�ţ�" & !����
                lblNo.Caption = "�����ţ�" & !no
                
                GetList = !no
                Exit Function
            
            End If
        End With
    Else
        Me.lblCommen.Caption = "��û����Ҫǩ���Ĵ�����"
        TimerAuto.Enabled = True
        GetList = ""
    End If
    Exit Function
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDetail(ByVal strNO As String)
    Dim strSQL As String
    Dim rsTemp As Recordset
    Dim dblMoney As Double
    Dim i As Integer
    
    On Error GoTo errRow
    strSQL = "Select a.No, a.���, b.���� ҩƷ����, b.���, b.���㵥λ ��λ, a.���ۼ� ����, a.ʵ������ ����, c.ʵ�ս�� ���" & vbNewLine & _
            "From ҩƷ�շ���¼ a, �շ���ĿĿ¼ b, ������ü�¼ c" & vbNewLine & _
            "Where a.ҩƷid = b.Id And a.����id = c.Id And (Mod(a.��¼״̬, 3) = 0 Or a.��¼״̬ = 1) and a.no=[1] and a.�ⷿid=[2]" & vbNewLine & _
            "Order By a.���"
    Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "GetDetail", strNO, mlngҩ��id)
    
    If Not rsTemp.EOF Then
        With Me.vsfList
            .Rows = rsTemp.RecordCount + 2
            
            For i = 1 To rsTemp.RecordCount
                .TextMatrix(i, .ColIndex("���")) = rsTemp!���
                .TextMatrix(i, .ColIndex("ҩƷ����")) = rsTemp!ҩƷ����
                .TextMatrix(i, .ColIndex("���")) = rsTemp!���
                .TextMatrix(i, .ColIndex("��λ")) = rsTemp!��λ
                .TextMatrix(i, .ColIndex("����")) = Format(rsTemp!����, "0.00#")
                .TextMatrix(i, .ColIndex("����")) = rsTemp!����
                .TextMatrix(i, .ColIndex("���")) = Format(rsTemp!���, "0.00#")
                dblMoney = Val(rsTemp!���) + dblMoney
                rsTemp.MoveNext
            Next
            
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "���С�ƣ�" & dblMoney
            .MergeCells = flexMergeFree
            .MergeRow(.Rows - 1) = True
            
        End With
    End If
    Exit Sub
errRow:
    If errcenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.lblCard.Left = (Me.ScaleWidth - (Me.lblCard.Width + Me.txtCard.Width)) / 2
    Me.txtCard.Left = Me.lblCard.Left + Me.lblCard.Width + 150
    Me.cmdBack.Left = Me.ScaleWidth - Me.cmdBack.Width - 200
    
    frmLineTop.Width = Me.ScaleWidth
    picUnBrush.Move picUnBrush.Left, picUnBrush.Top, Me.ScaleWidth - 200, Me.ScaleHeight - frmLineTop.Top - 500
    imgHos.Move imgHos.Left, imgHos.Top, picUnBrush.Width - 100, picUnBrush.Height - 1200
    Me.lblCommen.Move (Me.ScaleWidth - Me.lblCommen.Width) / 2, (picUnBrush.Height - imgHos.Height) / 2 + imgHos.Height
    
    picBrush.Move (Me.ScaleWidth - Me.picBrush.Width) / 2, picBrush.Top, picBrush.Width, Me.ScaleHeight - frmLineTop.Top + 200
    lblRecipeComment.Left = picBrush.Width - lblRecipeComment.Width - 200
    frmLineBrushTop.Width = picBrush.Width - 50
    Me.cmdPrevious.Left = picBrush.Width - cmdPrevious.Width - 100
    Me.cmdNext.Left = Me.cmdPrevious.Left - Me.cmdNext.Width - 100
    
    vsfList.Move vsfList.Left, vsfList.Top, frmLineBrushTop.Width, picUnBrush.Height - frmLineBrushTop.Top - 2200
    
    lblTo.Top = (picBrush.Height - vsfList.Height - vsfList.Top - lblTo.Height) / 2 + vsfList.Height + vsfList.Top - 200
    Me.cmdCheckIn.Top = lblTo.Top - 100
    Me.cmdCheckInAll.Top = lblTo.Top - 100
    lblMsg.Left = picBrush.Width - lblMsg.Width - 200
    lblMsg.Top = lblTo.Top + 200
    
    frmLineBruchButtom.Move frmLineTop.Left, Me.picBrush.Top + Me.cmdCheckIn.Top - 300, frmLineTop.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsRecipeList = Nothing
End Sub

Private Sub picUnBrush_KeyPress(KeyAscii As Integer)
    If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    txtCard.Text = txtCard.Text & Chr(KeyAscii)
    If Len(txtCard.Text) = txtCard.MaxLength - 1 And KeyAscii <> 8 Then
        txtCard.Text = txtCard.Text & Chr(KeyAscii)
        txtCard.SelStart = Len(txtCard.Text)
        KeyAscii = 0
    End If
End Sub

Private Sub TimerAuto_Timer()
    Me.lblCommen.Caption = "��ӭʹ������ǩ��ϵͳ"
    TimerAuto.Enabled = False
    Me.txtCard.Text = ""
    picUnBrush.Visible = True
    picBrush.Visible = False
    Me.cmdBack.Visible = False
End Sub

Private Sub txtCard_GotFocus()
    txtCard.SelStart = 0
    txtCard.SelLength = Len(txtCard.Text)
End Sub

Public Sub ShowMe(ByVal lngҩ��id As Long, ByVal strType As String)
    mlngҩ��id = lngҩ��id
    
    Me.lblCard.Caption = "��ˢ" & strType
    Me.Show
End Sub


Private Sub txtCard_KeyPress(KeyAscii As Integer)
     Dim strNO As String
    
    If KeyAscii = 13 And Not mBlnBegin Then
        If Me.txtCard.Text <> "" Then
            TimerAuto.Enabled = True
            strNO = GetList(txtCard.Text)
            Call SetPass(txtCard.Text)
            If strNO <> "" Then
                Call GetDetail(strNO)
                picBrush.Visible = True
                picUnBrush.Visible = False
                Me.cmdBack.Visible = True
                Me.frmLineBruchButtom.Visible = True
            Else
                picBrush.Visible = False
                picUnBrush.Visible = True
                Me.cmdBack.Visible = False
                Me.frmLineBruchButtom.Visible = False
            End If
        End If
        
        txtCard.SelStart = 0
        txtCard.SelLength = Len(txtCard.Text)
        mBlnBegin = True
    Else
        If InStr(":��;��?��''||" & Chr(22) & Chr(32), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub

Private Sub SetPass(ByVal strText As String)
    Dim count As Integer
    Dim intX As Integer
    Dim intY As Integer
    
    txtCard.Tag = strText
    intX = (Len(strText) - 3) / 2
    
    If intX < 2 Then
        txtCard.Text = Mid(txtCard.Text, 1, 1) & String(3, "*") & Mid(txtCard.Text, 5)
    Else
        txtCard.Text = Mid(txtCard.Text, 1, intX) & String(3, "*") & Mid(txtCard.Text, intX + 4)
    End If
   
End Sub
