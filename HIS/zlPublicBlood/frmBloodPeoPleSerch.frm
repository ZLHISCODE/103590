VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmBloodPeoPleSerch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ա��ѯ"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   Icon            =   "frmBloodPeoPleSerch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   300
      Left            =   3180
      TabIndex        =   3
      Top             =   2625
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   300
      Left            =   4290
      TabIndex        =   2
      Top             =   2610
      Width           =   1000
   End
   Begin zlIDKind.PatiIdentify PatiIdentify 
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBloodPeoPleSerch.frx":08CA
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   0
      ShowSortName    =   -1  'True
      DefaultCardType =   "���￨"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFSerch 
      Height          =   1890
      Left            =   60
      TabIndex        =   1
      Top             =   555
      Width           =   5310
      _cx             =   9366
      _cy             =   3334
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
   End
End
Attribute VB_Name = "frmBloodPeoPleSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String    '
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mRsPeople As ADODB.Recordset

Private Sub CMDcancel_Click()
    mstrKey = ""
    Unload Me
End Sub

Private Sub CMDok_Click()
    Dim lngi As Long
    mstrKey = ""
    For lngi = 1 To VSFSerch.Rows - 1
        If Abs(Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("ѡ��")))) = 1 And Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("����id"))) > 0 Then
            mstrKey = Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("����id"))) & "-" & Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("��ҳID"))) & "-" & Val(VSFSerch.TextMatrix(lngi, VSFSerch.ColIndex("����")))
        End If
    Next
    If mstrKey = "" Then
        MsgBox "��ѡ��Ҫ��ӵĲ��ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    Unload Me
End Sub

Public Function SerchPeople(frmMain As Object, lngModule As Long) As String
    Dim strCardKind As String
    mstrKey = ""
    '��ʼ��Patidentify�ؼ�
    Call CreateSquareCardObject(frmMain, 2200, lngModule)
    strCardKind = "��|����|0|0|0|0|0|0;ס|סԺ��|0|0|0|0|0|0;��|�����|0|0|0|0|0|0;��|���￨|0|0|8|0|0|0;��|�������֤|0|0|0|0|0|0;IC|IC��|1|0|0|0|0|0"
    If Not gobjCardSquare Is Nothing Then
        strCardKind = gobjCardSquare.zlGetIDKindStr(strCardKind)
    End If
    '���������Nothing,���������壬������ر�ʱ�ᴥ��active�¼���Ӧ���Ƕ��ˢ��ε��ø÷��������⣩
    Call PatiIdentify.zlInit(Nothing, 2200, , gcnOracle, gstrDBUser, gobjCardSquare, strCardKind)
    PatiIdentify.AutoSize = True
'    PI1.ShowPropertySet = True
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
    '��ʼ�����
    Call initvsf
    
    Me.Show 1, frmMain
    SerchPeople = mstrKey
End Function

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    '���ܣ���ȡ���š���������
    Dim strSQL As String
    Dim lngPatiID As Long
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.����id
    End If
    If lngPatiID = 0 Then
        Call FindPeople(PatiIdentify.Text)
    Else
            strSQL = _
                " Select Distinct b.����id, b.��ҳid, b.����, b.�Ա�, b.����, b.��Ժ���� As ����, b.סԺ��, '' As �Һŵ�,0 ����" & vbNewLine & _
                " From ѪҺ�շ���¼ d, ѪҺ��Ѫ��¼ c, ������ҳ b, ������Ϣ a" & vbNewLine & _
                " Where d.�䷢id = c.Id And Mod(d.��¼״̬, 3) = 1 And d.����� Is Not Null And c.����id = b.����id And c.��ҳid = b.��ҳid And" & vbNewLine & _
                "      b.����id = a.����id And b.��ҳid = a.��ҳid And a.��Ժ = 1 And a.����ID=[1]" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct b.����id, a.Id As ��ҳid, a.����, a.�Ա�, a.����, '' As ����, 0 As סԺ��, a.No As �Һŵ�,1 ����" & vbNewLine & _
                " From ѪҺ�շ���¼ d, ѪҺ��Ѫ��¼ c, ����ҽ����¼ b, ���˹Һż�¼ a" & vbNewLine & _
                " Where d.�䷢id = c.Id And Mod(d.��¼״̬, 3) = 1 And d.����� Is Not Null And c.����id = b.Id And b.����id = a.����id And b.�Һŵ� = a.No And" & vbNewLine & _
                "      b.������� = 'K' And a.ִ��״̬ = 2 And a.��¼���� = 1 And a.��¼״̬ = 1 And a.����id = [1]"
        Set mRsPeople = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", lngPatiID)
        Call mclsVsf.LoadGrid(mRsPeople)
    End If
    If mRsPeople.EOF = True Then
        VSFSerch.TextMatrix(1, VSFSerch.ColIndex("ѡ��")) = -1 'Ĭ���������ݳ�ʼΪѡ��״̬
    End If
End Sub


Private Sub initvsf()
    Set mclsVsf = New clsVsf
    With mclsVsf
        Call .Initialize(Me.Controls, VSFSerch, True, True)
        Call .ClearColumn
        Call .AppendColumn("ѡ��", 400, flexAlignLeftCenter, flexDTBoolean, "", "ѡ��", True)
        Call .AppendColumn("����id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
        Call .AppendColumn("��ҳid", 0, flexAlignRightCenter, flexDTString, "", "", True, , , True)
        Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("�Ա�", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("סԺ��", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("�Һŵ�", 800, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����", 800, flexAlignLeftCenter, flexDTLong, "", , False, False, False, True)
        
        .AppendRows = False
        .SysHidden(.ColIndex("����id")) = True
        .SysHidden(.ColIndex("��ҳid")) = True

        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
        
    End With
End Sub

Private Sub FindPeople(strfind As String)
    Dim strSQL As String
    If strfind = "" Then Exit Sub
    strSQL = _
        " Select Distinct b.����id, b.��ҳid, b.����, b.�Ա�, b.����, b.��Ժ���� As ����, b.סԺ��, '' As �Һŵ�,0 ����" & vbNewLine & _
        " From ѪҺ�շ���¼ d, ѪҺ��Ѫ��¼ c, ������ҳ b, ������Ϣ a" & vbNewLine & _
        " Where d.�䷢id = c.Id And Mod(d.��¼״̬, 3) = 1 And d.����� Is Not Null And c.����id = b.����id And c.��ҳid = b.��ҳid And" & vbNewLine & _
        "      b.����id = a.����id And b.��ҳid = a.��ҳid And a.��Ժ = 1 And a.���� Like [1]" & vbNewLine & _
        " Union All" & vbNewLine & _
        " Select Distinct b.����id, a.Id As ��ҳid, a.����, a.�Ա�, a.����, '' As ����, 0 As סԺ��, a.No As �Һŵ�,1 ����" & vbNewLine & _
        " From ѪҺ�շ���¼ d, ѪҺ��Ѫ��¼ c, ����ҽ����¼ b, ���˹Һż�¼ a, ������Ϣ e" & vbNewLine & _
        " Where d.�䷢id = c.Id And Mod(d.��¼״̬, 3) = 1 And d.����� Is Not Null And c.����id = b.Id And b.����id = a.����id And b.�Һŵ� = a.No And" & vbNewLine & _
        "      b.������� = 'K' And a.ִ��״̬ = 2 And a.��¼���� = 1 And a.��¼״̬ = 1 And a.����id = e.����id And e.���� Like [1]"
    Set mRsPeople = gobjDatabase.OpenSQLRecord(strSQL, "������Ϣ", strfind & "%")
    Call mclsVsf.LoadGrid(mRsPeople)
End Sub

Private Sub VSFSerch_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub VSFSerch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngi As Long
    If Col <> VSFSerch.ColIndex("ѡ��") Then Cancel = True: Exit Sub
    If Val(VSFSerch.TextMatrix(Row, VSFSerch.ColIndex("����id"))) = 0 Then Cancel = True: Exit Sub
    For lngi = VSFSerch.FixedRows To VSFSerch.Rows - 1
        If lngi <> Row Then
            VSFSerch.TextMatrix(lngi, Col) = 0
        End If
    Next
End Sub

Private Sub VSFSerch_DblClick()
    Dim lngRow As Long
    If VSFSerch.Row >= VSFSerch.FixedRows And VSFSerch.Row < VSFSerch.Rows Then
        If Val(VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("����id"))) = 0 Then Exit Sub
        For lngRow = VSFSerch.FixedRows To VSFSerch.Rows - 1
            If lngRow <> VSFSerch.Row Then
                VSFSerch.TextMatrix(lngRow, VSFSerch.ColIndex("ѡ��")) = 0
            End If
        Next
        If Abs(Val(VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("ѡ��")))) = 0 Then
            VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("ѡ��")) = 1
        Else
            VSFSerch.TextMatrix(VSFSerch.Row, VSFSerch.ColIndex("ѡ��")) = 0
        End If
    End If
End Sub
