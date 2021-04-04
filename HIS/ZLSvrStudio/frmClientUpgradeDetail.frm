VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClientUpgradeDetail 
   Caption         =   "������֤����"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "frmClientUpgradeDetail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   11010
   StartUpPosition =   1  '����������
   Tag             =   "������֤����"
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   9000
      TabIndex        =   1
      Top             =   4800
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _cx             =   17171
      _cy             =   7858
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmClientUpgradeDetail.frx":6852
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
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "��֤���:"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   810
   End
End
Attribute VB_Name = "frmClientUpgradeDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum UpValidate
    UpC_��� = 0
    UpC_ϵͳ
    UpC_ģ��
    UpC_��֤��ʼʱ��
    UpC_��֤����ʱ��
    UpC_��������
    UpC_��֤���
    UpC_��֤˵��
End Enum
Private Enum UpResult
    Rs_��֤�ɹ� = 0
    Rs_��֤�쳣
    Rs_��֤��ʱ
    Rs_��֤����
End Enum
Private mstrClient As String
Private mobjTip  As clsTipSwap
Private mlngLastRow As Long
Public Sub ShowMe(ByVal strClient As String, ByVal strIP As String)
    mstrClient = strClient
    Me.Caption = Me.Tag & "(" & strClient & " " & strIP & ")"
    If LoadData Then Me.Show 1
End Sub

Private Function LoadData() As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim arrNum(Rs_��֤����) As Long '������֤���������

    On Error GoTo errH
    gstrSQL = "Select c.���� ϵͳ, b.���� ģ��, To_Char(a.��֤��ʼʱ��, 'yyyy-mm-dd hh:mi') ��֤��ʼʱ��, To_Char(a.��֤����ʱ��, 'yyyy-mm-dd hh:mi') ��֤����ʱ��," & vbNewLine & _
            "       b.���� ��������, Decode(a.��֤���, 1, '��֤�ɹ�', 0, '��֤�쳣', 2, '��֤��ʱ') ��֤���, a.��֤˵��" & vbNewLine & _
            "From Zlclientvertify a, Zlprograms b, Zlsystems c" & vbNewLine & _
            "Where a.ϵͳ = b.ϵͳ And a.ģ�� = b.��� And b.ϵͳ = c.��� And a.�ͻ��� =[1]"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, gstrSQL, Me.Caption, mstrClient)
    If rsTemp.RecordCount = 0 Then
        MsgBox "��" & mstrClient & "���ͻ�������û��������֤����֤��ɵ���ϸ��Ϣ��"
        LoadData = False
        Exit Function
    Else
        With vsfDetail
            .Rows = rsTemp.RecordCount + 1
            For i = 1 To rsTemp.RecordCount
                .TextMatrix(i, UpC_���) = i
                .TextMatrix(i, UpC_ϵͳ) = rsTemp!ϵͳ
                .TextMatrix(i, UpC_ģ��) = rsTemp!ģ��
                .TextMatrix(i, UpC_��֤��ʼʱ��) = rsTemp!��֤��ʼʱ��
                .TextMatrix(i, UpC_��֤����ʱ��) = rsTemp!��֤����ʱ��
                .TextMatrix(i, UpC_��֤���) = rsTemp!��֤���
                .TextMatrix(i, UpC_��֤˵��) = rsTemp!��֤˵�� & ""
                .TextMatrix(i, UpC_��������) = rsTemp!��������
                If rsTemp!��֤��� = "��֤�ɹ�" Then
                    arrNum(Rs_��֤�ɹ�) = arrNum(Rs_��֤�ɹ�) + 1
                ElseIf rsTemp!��֤��� = "��֤�쳣" Then
                    arrNum(Rs_��֤�쳣) = arrNum(Rs_��֤�쳣) + 1
                Else
                    arrNum(Rs_��֤��ʱ) = arrNum(Rs_��֤��ʱ) + 1
                End If
                rsTemp.MoveNext
            Next
            arrNum(Rs_��֤����) = .Rows - 1
        End With
        lblResult.Caption = "��֤������б��й���ʾ" & arrNum(Rs_��֤����) & "�����ݣ���֤�ɹ�" & arrNum(Rs_��֤�ɹ�) & "������֤�쳣" & arrNum(Rs_��֤�쳣) & "������֤��ʱ" & arrNum(Rs_��֤��ʱ) & "����"
        LoadData = True
    End If
    Exit Function
errH:
    MsgBox err.Description, vbCritical, Me.Caption
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 13000
    Me.Height = 7400
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    vsfDetail.Move 0, 0, Me.Width - 240, Me.Height - cmdClose.Height - 800
    cmdClose.Move Me.Width - cmdClose.Width - 400, vsfDetail.Top + vsfDetail.Height + 100
    lblResult.Move 50, cmdClose.Top + 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjTip = Nothing
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTip As String
    Dim strTitle As String
    
    If mobjTip Is Nothing Then
        Set mobjTip = New clsTipSwap
        Set mobjTip.ParentControl = vsfDetail
        mobjTip.Icon = TTIconInfo
        mobjTip.Style = TTBalloon
        mobjTip.Create
    End If
    If mlngLastRow <> NewRow Then
        mlngLastRow = NewRow
        If NewRow >= vsfDetail.FixedRows Then
            strTip = vsfDetail.TextMatrix(NewRow, UpC_��֤˵��)
            strTitle = "ϵͳ��" & vsfDetail.TextMatrix(NewRow, UpC_ϵͳ) & " ģ�飺" & vsfDetail.TextMatrix(NewRow, UpC_ģ��)
            If strTip = "" Then strTip = "<��˵����Ϣ>"
            mobjTip.TipText = SwapText(strTip)
            mobjTip.Title = strTitle
        Else
            mobjTip.TipText = ""
            mobjTip.Title = ""
        End If
    End If
End Sub

Private Sub vsfDetail_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    With vsfDetail
        For i = 1 To .Rows - 1
            .TextMatrix(i, UpC_���) = i
        Next
    End With
End Sub

Private Sub vsfDetail_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mobjTip Is Nothing Then
        mobjTip.TipText = ""
        mobjTip.Title = ""
    End If
End Sub

Private Sub vsfDetail_BeforeSort(ByVal Col As Long, Order As Integer)
    If Col <> UpC_��֤��� Then Order = 0

End Sub

Private Function SwapText(ByVal strTxt As String) As String
    
    Dim strReturn As String, strTmp As String, i As Integer
    strReturn = strTxt
    If InStr(strTxt, ";") > 0 Then
        strReturn = SwapWord(strReturn, ";")
    End If
    If InStr(strTxt, "��") > 0 Then
        strReturn = SwapWord(strReturn, "��")
    End If
    If InStr(strTxt, ".") > 0 Then
        strReturn = SwapWord(strReturn, ".")
    End If
    If InStr(strTxt, "��") > 0 Then
        strReturn = SwapWord(strReturn, "��")
    End If
    
    If strReturn = strTxt Then
        strReturn = swapLine("����" & strTxt)
    End If
    '--
    strReturn = Replace(strReturn, " ", "")
    strReturn = Replace(strReturn, "��", "")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    '���ڻ���
    Dim varLine As Variant
    
    varLine = Split(strReturn, "[CR]")
    For i = 0 To UBound(varLine)
        strTmp = strTmp & swapLine("����" & varLine(i)) & vbNewLine
    Next
    
    If strTmp <> "" Then
        strReturn = strTmp
    End If
    '--���������Ŀ���
    strReturn = ClearLine(strReturn)
    SwapText = strReturn
End Function

Private Function ClearLine(strTxt) As String
    Dim i As Integer
    Dim varLine As Variant
    Dim strReturn As String
    varLine = Split(strTxt, vbNewLine)
    For i = 0 To UBound(varLine)
        If InStr(",.;?!])}%>���������������ݣ�����������", Mid(varLine(i), 1, 1)) > 0 Then
            strReturn = Mid(strReturn, 1, Len(strReturn) - 4) & Mid(varLine(i), 1, 1) & "[CR]" & Mid(varLine(i), 2) & "[CR]"
        Else
            strReturn = strReturn & varLine(i) & "[CR]"
        End If
    Next
    
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]��[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR];[CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR].[CR]", "[CR]")
    
    strReturn = Replace(strReturn, "[CR][CR]", "[CR]")
    strReturn = Replace(strReturn, "[CR]", vbNewLine)
    ClearLine = strReturn
End Function

Private Function SwapWord(ByVal strTxt As String, strWord As String) As String
    Dim varLine As Variant
    Dim strReturn As String
    Dim i As Integer
    Dim strTxtTmp As String
    
    strTxtTmp = strTxt
    If Mid(strTxt, Len(strTxt), 1) = strWord Then
        strTxtTmp = Mid(strTxt, 1, Len(strTxt) - 1)
    End If
    
    If InStr(strTxtTmp, strWord) > 0 Then
        varLine = Split(strTxtTmp, strWord)
        For i = 0 To UBound(varLine)
            If varLine(i) <> "" Then
                'varLine(i) = swapLine("����" & varLine(i))
                If varLine(i) & strWord <> strWord Then
                    strReturn = strReturn & varLine(i) & strWord & "[CR]"
                End If
            End If
        Next
    End If
    'If Mid(strTxtTmp, Len(strTxtTmp), 1) <> strWord Then strReturn = Mid(strReturn, 1, Len(strReturn) - 1)
    If strReturn <> "" Then
        SwapWord = strReturn
    Else
        SwapWord = strTxt
    End If
End Function

Private Function swapLine(ByVal strTxt As String) As String
    Dim strTmp As String
    strTmp = strTxt
    
    If Len(strTxt) > 18 Then
        swapLine = Mid(strTmp, 1, 18) & vbNewLine
        strTmp = Mid(strTmp, 19)
        swapLine = swapLine & swapLine(strTmp)
    Else
        swapLine = strTxt
    End If
End Function

