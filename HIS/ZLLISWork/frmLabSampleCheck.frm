VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabSampleCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ͼ�걾�˶�"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14355
   Icon            =   "frmLabSampleCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14355
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtGotoSample 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6210
      TabIndex        =   25
      Top             =   7620
      Width           =   3150
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   14280
      TabIndex        =   14
      Top             =   0
      Width           =   14310
      Begin VB.TextBox txtSampleCode 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   915
         TabIndex        =   0
         Top             =   248
         Width           =   3150
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ɨ������"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ͼ����"
         Height          =   180
         Index           =   1
         Left            =   5220
         TabIndex        =   20
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ͼ�ʱ��"
         Height          =   180
         Index           =   3
         Left            =   10365
         TabIndex        =   19
         Top             =   330
         Width           =   720
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   6210
         TabIndex        =   18
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   8610
         TabIndex        =   17
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblInto 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   11265
         TabIndex        =   16
         Top             =   330
         Width           =   90
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   8550
         X2              =   9780
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6135
         X2              =   7365
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   11265
         X2              =   13530
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�ͼ���"
         Height          =   180
         Index           =   2
         Left            =   7860
         TabIndex        =   15
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdQuet 
      Caption         =   "�˳�(&Q)"
      Height          =   360
      Left            =   13245
      TabIndex        =   4
      Top             =   7590
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�˶�(&D)"
      Height          =   360
      Left            =   11865
      TabIndex        =   3
      Top             =   7590
      Width           =   1035
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   8115
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   6630
      Left            =   -45
      ScaleHeight     =   6600
      ScaleWidth      =   14355
      TabIndex        =   1
      Top             =   840
      Width           =   14385
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6285
         Index           =   2
         Left            =   120
         ScaleHeight     =   6285
         ScaleWidth      =   7050
         TabIndex        =   12
         Top             =   45
         Width           =   7050
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   6165
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   270
            Width           =   7065
            _cx             =   12462
            _cy             =   10874
            Appearance      =   3
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
            BackColorBkg    =   -2147483636
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���ǼǱ걾(0)"
            Height          =   180
            Index           =   4
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   1170
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4230
         Index           =   1
         Left            =   7275
         ScaleHeight     =   4230
         ScaleWidth      =   7005
         TabIndex        =   9
         Top             =   45
         Width           =   7005
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   4140
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   285
            Width           =   7065
            _cx             =   12462
            _cy             =   7302
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ɨ��걾(0)"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.Frame fraNS 
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   7185
         MousePointer    =   7  'Size N S
         TabIndex        =   6
         Top             =   4335
         Width           =   7095
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   7260
         ScaleHeight     =   1815
         ScaleWidth      =   7065
         TabIndex        =   5
         Top             =   4545
         Width           =   7065
         Begin VSFlex8Ctl.VSFlexGrid vsfList 
            Height          =   1545
            Index           =   2
            Left            =   30
            TabIndex        =   7
            Top             =   225
            Width           =   6960
            _cx             =   12277
            _cy             =   2725
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
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
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����ѵǼǻ��Ѻ��ձ걾(0)"
            Height          =   180
            Index           =   6
            Left            =   45
            TabIndex        =   8
            Top             =   45
            Width           =   2250
         End
      End
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   5685
      TabIndex        =   24
      Top             =   7680
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "˫��""���Ǽ�""����е������п��Խ�������ӵ�""��ɨ��""�����,˫��""��ɨ��""����е������п��Խ������˻ص�""���Ǽ�""�����"
      ForeColor       =   &H00004000&
      Height          =   465
      Left            =   90
      TabIndex        =   23
      Top             =   7560
      Width           =   5310
   End
End
Attribute VB_Name = "frmLabSampleCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnUse As Boolean                       '��ǰ�����Ƿ�ʹ��
Private mstrPrivs As String
Private mlngBatch As Long
Private mlngSampleCount As Long                  '�����걾����
Private mObjSelectVSF As VSFlexGrid              '������VSF�ؼ�
Private mstrFind As String
'Private WithEvents mfrmFind As frmLabSampleCheckFind

Public Sub ShowME(Objfrm As Object)
    Me.Show vbModal, Objfrm
End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 6 Then
'        Call cmdFind_Click
'    End If
'End Sub
'
'Private Sub mfrmFind_Finded(ByVal blnFind As Boolean, ByVal strVale As String)
'    '��λ:
'    Dim varTmp As Variant, strSampleCode As String
'    If blnFind Then
'        varTmp = Split(strVale, ",")
'        strSampleCode = varTmp(0)
'        Call findSample(strSampleCode)
''        Call RptItem_SelectionChanged
'    End If
'End Sub

Private Sub saveSample()
    '��ȡ�걾��
    Dim i As Integer
    Dim strSampleIDs As String
    
    With Me.vsfList(1)
        For i = 1 To .Rows - 1
            strSampleIDs = strSampleIDs & .TextMatrix(i, .ColIndex("ҽ��ID")) & ","
        Next
    End With
    Call SaveRegister(strSampleIDs, Me.vsfList(1))
End Sub

Private Function SaveRegister(ByVal strSampleIDs As String, objVsf As VSFlexGrid) As Boolean
    'ǩ�ձ걾       strSampleCodes-����ı걾��,��","�ָ�
    Dim var_Tmp As Variant
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intTimeLimit As Integer         '�ͼ�ʱ�޵�λ����
    Dim blnTimeLimit As Boolean         '�Ƿ񳬹��ͼ�ʱ�� true = ����
    Dim strAdvice As String
    Dim blnShowMsg As Boolean
    Dim blnSave As Boolean              '�Ƿ�ǿ��ͨ��
    Dim i As Integer
    
'    On Error GoTo ErrHand
    var_Tmp = Split(strSampleIDs, ",")
    blnShowMsg = True
    blnSave = False
    For i = 0 To UBound(var_Tmp) - 1
        If Chk���۷���(Me, CStr(var_Tmp(i)), 0) = False Then
            MsgBox var_Tmp(i) & "û�л���", vbInformation, "��ʾ"
            Exit Function
        End If
    Next
    
    If mblnUse = True Or mlngBatch = 0 Then
        '�õ�һ���µ�����
        strSQL = "select ����ҽ������_��������.nextval from dual "
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
        mlngBatch = rsTmp(0)
        mblnUse = False
    End If
    
    With objVsf
        
        For i = 1 To .Rows - 1
            '�����Ƿ񳬹��ɼ�ʱ��
            strSQL = "select �ͼ�ʱ�� from ������Ŀѡ�� where ������Ŀid = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(.TextMatrix(i, .ColIndex("������ĿID"))))
            If rsTmp.EOF = True Then
                intTimeLimit = 0
            Else
                intTimeLimit = Val(Nvl(rsTmp("�ͼ�ʱ��")))
            End If
            
            If IsDate(.TextMatrix(i, .ColIndex("����ʱ��"))) = False And intTimeLimit > 0 Then
                blnTimeLimit = True
            Else
                If IsDate(.TextMatrix(i, .ColIndex("����ʱ��"))) = True Then
                    If DateDiff("n", .TextMatrix(i, .ColIndex("����ʱ��")), zlDatabase.Currentdate) > intTimeLimit _
                        And intTimeLimit > 0 Then
                        '�����ͼ�ʱ��
                        blnTimeLimit = True
                    End If
                Else
                    If intTimeLimit > 0 Then
                        blnTimeLimit = True
                    End If
                End If
            End If
            
            If blnTimeLimit = True Then
                '��ʱ�����鿴�Ƿ���Ȩ�ޣ���Ȩ��ʱֻ��ʾ
                If InStr(mstrPrivs, "ǿ��ͨ���ͼ�ʱ��") > 0 Then
                    If blnShowMsg = True Then
                        '��ʾ
                        If MsgBox("��������ʱ��Ϊ��" & .TextMatrix(i, .ColIndex("����ʱ��")) & "��" & vbCrLf & _
                                "�ѳ�������ʱ��" & intTimeLimit & "����,�ͼ��ӳ٣�" & vbCrLf & _
                                "����ǿ��ͨ���ͼ�ʱ��Ȩ��" & vbCrLf & _
                                "�Ƿ�ǿ��ͨ��?", vbQuestion + vbYesNo) = vbYes Then
                            blnSave = True
                        End If
                        blnShowMsg = False
                    End If
                    If blnSave = True Then
                        strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("ҽ��ID"))
                        Call vsfDataToVsfData(.TextMatrix(i, .ColIndex("����")), objVsf, Me.vsfList(2))
                        .RowHidden(i) = True
                    End If
                Else
                    '�ܾ��Ǽ�
                    If blnShowMsg = True Then
                        MsgBox ("��������ʱ��Ϊ��" & .TextMatrix(i, .ColIndex("����ʱ��")) & "��" & vbCrLf & _
                                "�ѳ�������ʱ��" & intTimeLimit & "����,������Ǽǣ�")
                        blnShowMsg = False
                    End If
                End If

            ElseIf .TextMatrix(i, .ColIndex("����ʱ��")) = "" Then
                '����ǿ�ƵǼ�δ�����걾
                If InStr(mstrPrivs, "ǿ�ƵǼ�δ�����걾") > 0 Then
                    '��ʾ
                    If blnShowMsg = True Then
'                        If MsgBox("��ǰ��" & .TextMatrix(i, .ColIndex("������Ŀ")) & "��δ����!", vbInformation + vbQuestion) = vbYes Then
'                            blnSave = True
'                        End If
'                        If blnSave = True Then
                            strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("ҽ��ID"))
'                        End If
                        blnShowMsg = False
                    End If
                Else
                    '�ܾ��Ǽ�
                    If blnShowMsg = True Then
                        MsgBox "��ǰ��" & .TextMatrix(i, .ColIndex("������Ŀ")) & "��δ����,������Ǽǣ�", vbInformation
                        blnShowMsg = False
                    End If
                End If
            Else
                strAdvice = strAdvice & "|" & .TextMatrix(i, .ColIndex("ҽ��ID"))
                Call vsfDataToVsfData(.TextMatrix(i, .ColIndex("����")), objVsf, Me.vsfList(2))
                .RowHidden(i) = True
            End If
        Next
    End With
    Call RemoveHiddenItem(objVsf)
    Call showNum
    
    strSQL = "Zl_����ҽ������_SampleInput('" & Mid(strAdvice, 2) & "','" & UserInfo.���� & "','" & mlngBatch & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure strSQL, gstrSysName
    mblnUse = True
    
    If strAdvice <> "" Then
        Call WriterCheckSampleToLIS(Mid(strAdvice, 2), UserInfo.����, mlngBatch)
    End If
    SaveRegister = True
    Exit Function
ErrHand:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub WriterCheckSampleToLIS(strAdvices As String, strName As String, strBatchNO As Long)
    '����   ��ǩ����Ϣд��LIS
    Dim strErr As String
    If Not mobjLisInsideComm Is Nothing Then
        If mobjLisInsideComm.SampleCheckinInfoWrite(strAdvices, strName, strBatchNO, strErr) = False Then
            MsgBox "д��ǩ����Ϣ��LIS���뵥����!" & vbCrLf & strErr
        End If
    End If
End Sub

'Private Sub cmdFind_Click()
'    Dim strFindSQL As String, strFindFiled As String
'    If mfrmFind Is Nothing Then Set mfrmFind = New frmLabSampleCheckFind
'    strFindSQL = "select ���� from (Select Distinct a.ҽ��id, a.�������� ����, b.����, b.�걾��λ As �걾, " & _
'                 " b.ҽ������ ������Ŀ, a.�ͼ���, c.���� �ͼ����, b.������Ŀid, a.����ʱ��, a.�걾�ͳ�ʱ�� �ͼ�ʱ��," & _
'                 " a.�걾�������� , a.������, a.����ʱ�� From ����ҽ������ A, ����ҽ����¼ B, ���ű� C" & _
'                 " Where a.ҽ��id = b.Id And a.ִ�в���id = c.Id And a.ִ��״̬ In (0) And" & _
'                 " a.�걾�������� In (Select �걾�������� From ����ҽ������ Where �������� =100008332023)) where" & _
'                 " ���� like [1] or ���� like [1] or �걾 LIKE [1] or ������Ŀ like [1]"
'    Call mfrmFind.ShowFind(strFindSQL)
'End Sub

Private Sub cmdQuet_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call saveSample
End Sub

Private Sub Form_Load()
     
    Call vsfSeting(Me.vsfList(0), 0)
    Call vsfSeting(Me.vsfList(1), 1)
    Call vsfSeting(Me.vsfList(2), 2)
    mstrPrivs = gstrPrivs       '��ʹ��Ȩ��
    
    If mobjLisInsideComm Is Nothing Then
        Dim strErr As String
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        '��ʼ��LIS�ӿڲ���
        If Not mobjLisInsideComm Is Nothing Then
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If
    
End Sub

Private Sub vsfSeting(ByVal objVsf As VSFlexGrid, Optional Index As Integer)
    Dim intFontSize As Integer
    Dim lbl As Label, lblInto As Label
    
    intFontSize = 11
    With objVsf
        .Clear
        .FixedCols = 0
        .Cols = 15
        .Rows = 1
        .ExtendLastCol = True
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShow
        .ColWidth(0) = 1
        .TextMatrix(0, 1) = "����": .ColKey(1) = "����": .ColWidth(.ColIndex("����")) = 2000: .Cell(flexcpAlignment, 0, .ColIndex("����")) = flexAlignCenterCenter
        .TextMatrix(0, 2) = "����": .ColKey(2) = "����": .ColWidth(.ColIndex("����")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("����")) = flexAlignCenterCenter
        .TextMatrix(0, 3) = "�Ա�": .ColKey(3) = "�Ա�": .ColWidth(.ColIndex("�Ա�")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("�Ա�")) = flexAlignCenterCenter
        .TextMatrix(0, 4) = "�걾": .ColKey(4) = "�걾": .ColWidth(.ColIndex("�걾")) = 1000: .Cell(flexcpAlignment, 0, .ColIndex("�걾")) = flexAlignCenterCenter
        .TextMatrix(0, 5) = "������Ŀ": .ColKey(5) = "������Ŀ": .Cell(flexcpAlignment, 0, .ColIndex("������Ŀ")) = flexAlignCenterCenter
        .TextMatrix(0, 6) = "�ͼ���": .ColKey(6) = "�ͼ���": .Cell(flexcpAlignment, 0, .ColIndex("�ͼ���")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�ͼ���")) = True
        .TextMatrix(0, 7) = "�ͼ����": .ColKey(7) = "�ͼ����": .Cell(flexcpAlignment, 0, .ColIndex("�ͼ����")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�ͼ����")) = True
        .TextMatrix(0, 8) = "�ͼ�ʱ��": .ColKey(8) = "�ͼ�ʱ��": .Cell(flexcpAlignment, 0, .ColIndex("�ͼ�ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�ͼ�ʱ��")) = True
        .TextMatrix(0, 9) = "������": .ColKey(9) = "������": .Cell(flexcpAlignment, 0, .ColIndex("������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("������")) = True
        .TextMatrix(0, 10) = "����ʱ��": .ColKey(10) = "����ʱ��": .Cell(flexcpAlignment, 0, .ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
        .TextMatrix(0, 11) = "ҽ��ID": .ColKey(11) = "ҽ��ID": .Cell(flexcpAlignment, 0, .ColIndex("ҽ��ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("ҽ��ID")) = True
        .TextMatrix(0, 12) = "������ĿID": .ColKey(12) = "������ĿID": .Cell(flexcpAlignment, 0, .ColIndex("������ĿID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("������ĿID")) = True
        .TextMatrix(0, 13) = "����ʱ��": .ColKey(13) = "����ʱ��": .Cell(flexcpAlignment, 0, .ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
        .TextMatrix(0, 14) = "����": .ColKey(14) = "����": .ColWidth(.ColIndex("����")) = 1200: .Cell(flexcpAlignment, 0, .ColIndex("����")) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 1) = 3 '������ж���
        .BackColorBkg = vbWhite
        .FontSize = intFontSize
    End With
    For Each lbl In Me.lbl
        lbl.FontSize = intFontSize
    Next
    For Each lblInto In Me.lblInto
        lblInto.FontSize = intFontSize
    Next
    Me.txtSampleCode.FontSize = intFontSize
    Me.txtSampleCode.Height = Me.lbl(0).Height
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrHand
    Me.txtSampleCode.Left = Me.lbl(0).Left + Me.lbl(0).Width + 100
    Me.lblInto(0).Move Me.lbl(1).Left + Me.lbl(1).Width + 100, Me.lbl(1).Top
'    With Me.Line1(0)
'        .X1 = Me.lblInto(0).Left
'        .X2 = Me.lblInto(0).Left + Me.lblInto(0).Width
'        .Y1 = Me.lbl(1).Top + Me.lbl(1).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(1).Move Me.lbl(2).Left + Me.lbl(2).Width + 100, Me.lblInto(0).Top
'    With Me.Line1(1)
'        .X1 = Me.lblInto(1).Left
'        .X2 = Me.lblInto(1).Left + Me.lblInto(1).Width
'        .Y1 = Me.lbl(2).Top + Me.lbl(2).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(2).Move Me.lbl(3).Left + Me.lbl(3).Width + 100, Me.lblInto(0).Top
'    With Me.Line1(2)
'        .X1 = Me.lblInto(2).Left
'        .X2 = Me.lblInto(2).Left + Me.lblInto(2).Width
'        .Y1 = Me.lbl(3).Top + Me.lbl(3).Height + 50
'        .Y2 = .Y1
'    End With
    Me.lblInto(1).Left = Me.lbl(2).Left + Me.lbl(2).Width + 100
    Me.lblInto(2).Left = Me.lbl(3).Left + Me.lbl(3).Width + 100
    Me.StatusBar.Panels(1).Width = Me.Width
    
    Me.picTop.Move 0, 0, Me.Width
    Me.picMain.Move 0, Me.picTop.Height, Me.picTop.Width
    Exit Sub
ErrHand:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnUse = False
    mlngBatch = 0
End Sub



Private Sub fraNS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.fraNS.Top = Me.fraNS.Top + Y
        Call picMain_Resize
    End If
End Sub

Private Sub pic_Resize(Index As Integer)
    On Error GoTo ErrHand
    Select Case Index
        Case 0
            Me.lbl(6).Move 50, 50
            Me.vsfList(2).Move 50, Me.lbl(6).Top + Me.lbl(6).Height + 50, Me.pic(Index).Width - 150, Me.pic(Index).Height - Me.lbl(6).Height - 100
        Case 1
            Me.lbl(5).Move 50, 50
            Me.vsfList(1).Move 50, Me.lbl(5).Top + Me.lbl(5).Height + 50, Me.pic(Index).Width - 150, Me.pic(Index).Height - Me.lbl(5).Height - 100
        Case 2
            Me.lbl(4).Move 50, 50
            Me.vsfList(0).Move 50, Me.lbl(4).Top + Me.lbl(4).Height + 50, Me.pic(Index).Width, Me.pic(Index).Height - Me.lbl(4).Height - 100
    End Select
    Exit Sub
ErrHand:
    
End Sub

Private Sub picMain_Resize()
    On Error GoTo ErrHand
    Me.pic(2).Move 0, 0, Me.picMain.Width / 2 - 10, Me.picMain.Height
    Me.pic(1).Move Me.picMain.Width / 2 + 10, 0, Me.picMain.Width / 2 - 60, Me.fraNS.Top
    Me.pic(0).Move Me.picMain.Width / 2 + 10, Me.fraNS.Top + Me.fraNS.Height, Me.picMain.Width / 2 - 60, Me.picMain.Height - Me.fraNS.Top - Me.fraNS.Height
    Exit Sub
ErrHand:
    
End Sub

Private Sub txtGotoSample_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer
    Dim intStart As Integer

    If KeyAscii = 13 Then
        If Val(txtGotoSample.Tag) = 0 Then
            intStart = 1
        Else
            intStart = Val(txtGotoSample.Tag) + 1
        End If
        With Me.vsfList(0)
            For intRow = intStart To .Rows - 1
                If (.TextMatrix(intRow, .ColIndex("����")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
                If (.TextMatrix(intRow, .ColIndex("������Ŀ")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
            Next
            If intRow >= .Rows - 1 Then
                txtGotoSample.Tag = 0
            End If
        End With
        With Me.vsfList(1)
            For intRow = intStart To .Rows - 1
                If (.TextMatrix(intRow, .ColIndex("����")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
                If (.TextMatrix(intRow, .ColIndex("������Ŀ")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
            Next
            If intRow >= .Rows - 1 Then
                txtGotoSample.Tag = 0
            End If
        End With
        With Me.vsfList(2)
            For intRow = intStart To .Rows - 1
                If (.TextMatrix(intRow, .ColIndex("����")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
                If (.TextMatrix(intRow, .ColIndex("������Ŀ")) Like "*" & Me.txtGotoSample.Text & "*") = True Then
                    .Row = intRow
                    txtGotoSample.Tag = .Row
                    .ShowCell intRow, 1
                    Call selectAll(txtGotoSample)
                    Exit For
                End If
            Next
            If intRow >= .Rows - 1 Then
                txtGotoSample.Tag = 0
            End If
        End With
        txtGotoSample.SetFocus
        Call selectAll(txtGotoSample)
    End If
End Sub

Private Sub txtSampleCode_GotFocus()
   Call selectAll(Me.txtSampleCode)
End Sub

Private Sub selectAll(ByVal objTxt As TextBox)
    objTxt.SelStart = 0
    objTxt.SelLength = Len(objTxt.Text)
End Sub

Private Sub txtSampleCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Call setVSFData(Trim(Me.txtSampleCode.Text))
            Call selectAll(Me.txtSampleCode)
    End Select
End Sub


Private Sub findSample(ByVal strSampleCode As String)
    Dim i As Long
    
    With Me.vsfList(0)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                Me.vsfList(1).Select 0, 1
                Me.vsfList(2).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
    With Me.vsfList(1)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                Me.vsfList(0).Select 0, 1
                Me.vsfList(2).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
    With Me.vsfList(2)
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                Me.vsfList(0).Select 0, 1
                Me.vsfList(1).Select 0, 1
                .Select i, 1, i, 4
                .ShowCell i, 1
            End If
        Next
    End With
End Sub

Private Function setVSFData(ByVal strSampleCode As String) As Boolean
    '����               �����ݵ�VSF
    'strSampleCode      ɨ�������
    Dim strSampleCodesLeft As String '���VSF����������
    Dim strSampleCodesRight As String '�ұ�VSF����������
    Dim strSampleCodesYDJ As String     '�ѵǼǻ��Ѻ�������
    Dim var_Tmp As Variant
    Dim rsData As Recordset
    Dim i As Integer, j As Integer
    
    '��֤TAT�Ƿ�ʱ
    If getTATTime(strSampleCode) = False Then
        Exit Function
    End If
    
    Set rsData = ReadData(strSampleCode)
    mlngSampleCount = rsData.RecordCount
    
    If rsData.EOF = True Then
        MsgBox "���벻��ȷ���߱�����ȫ���Ǽ�,����    ", vbInformation, "��ʾ"
        Exit Function
    End If
    
    With Me.vsfList(0)
        For i = 1 To .Rows - 1
            strSampleCodesLeft = strSampleCodesLeft & .TextMatrix(i, .ColIndex("����")) & ","
        Next
    End With
    With Me.vsfList(1)
        For i = 1 To .Rows - 1
            strSampleCodesRight = strSampleCodesRight & .TextMatrix(i, .ColIndex("����")) & ","
        Next
    End With
    With Me.vsfList(2)
        For i = 1 To .Rows - 1
            strSampleCodesYDJ = strSampleCodesYDJ & .TextMatrix(i, .ColIndex("����")) & ","
        Next
    End With

    If InStr(strSampleCodesLeft, strSampleCode & ",") > 0 And InStr(strSampleCodesRight, strSampleCode & ",") = 0 Then
        '�����VSF��ɨ���������뵽�ұ�VSF
               
        With Me.vsfList(0)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                    
                    
                    With Me.vsfList(1)
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = vsfList(0).TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = vsfList(0).TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = vsfList(0).TextMatrix(i, .ColIndex("�Ա�")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�Ա�")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("�걾")) = vsfList(0).TextMatrix(i, .ColIndex("�걾")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�걾")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = vsfList(0).TextMatrix(i, .ColIndex("������Ŀ")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������Ŀ")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("�ͼ���")) = vsfList(0).TextMatrix(i, .ColIndex("�ͼ���")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ���")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("�ͼ����")) = vsfList(0).TextMatrix(i, .ColIndex("�ͼ����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ����")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��")) = vsfList(0).TextMatrix(i, .ColIndex("�ͼ�ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ�ʱ��")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("������")) = vsfList(0).TextMatrix(i, .ColIndex("������")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = vsfList(0).TextMatrix(i, .ColIndex("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = vsfList(0).TextMatrix(i, .ColIndex("ҽ��ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("ҽ��ID")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("������ĿID")) = vsfList(0).TextMatrix(i, .ColIndex("������ĿID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������ĿID")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = vsfList(0).TextMatrix(i, .ColIndex("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = vsfList(0).TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                        Me.lblInto(0).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("�ͼ����"))
                        Me.lblInto(1).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("�ͼ���"))
                        Me.lblInto(2).Caption = vsfList(1).TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��"))
                    End With
                    .RowHidden(i) = True
                    Call Form_Resize
                End If
            Next
            Call RemoveHiddenItem(Me.vsfList(0))
'            Me.lbl(4).Caption = "���ǼǱ걾(" & Me.vsfList(0).Rows - 1 & ")"
'            Me.lbl(5).Caption = "��ɨ��걾(" & Me.vsfList(1).Rows - 1 & ")"
            Call showNum
            Me.StatusBar.Panels(1).Text = "�����걾����:" & mlngSampleCount & "��"
        End With
    ElseIf InStr(strSampleCodesRight, strSampleCode & ",") = 0 And InStr(strSampleCodesYDJ, strSampleCode & ",") = 0 Then
'        '������
'        If Me.vsfList(0).Rows > 1 Then
'            If MsgBox("�����벻���ڱ�����,�Ƿ������������ɨ��������?    ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                Exit Function
'            End If
'        End If
        '��ʼ�����
'        Call vsfSeting(vsfList(0), 0)
'        Call vsfSeting(vsfList(1), 1)
'        Call vsfSeting(vsfList(2), 2)
        '������
        For i = 1 To rsData.RecordCount
            If IIf(IsNull(rsData("������")), "", rsData("������")) = "" Then    'δ�Ǽ�
                With Me.vsfList(0)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData("����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData("����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsData("�Ա�"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�Ա�")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�걾")) = rsData("�걾"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�걾")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsData("������Ŀ"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������Ŀ")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ���")) = IIf(IsNull(rsData("�ͼ���")), "", rsData("�ͼ���")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ���")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ����")) = rsData("�ͼ����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��")) = rsData("�ͼ�ʱ��"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ�ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = IIf(IsNull(rsData("������")), "", rsData("������")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = IIf(IsNull(rsData("����ʱ��")), "", rsData("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = IIf(IsNull(rsData("ҽ��ID")), "", rsData("ҽ��ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("ҽ��ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������ĿID")) = IIf(IsNull(rsData("������ĿID")), "", rsData("������ĿID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������ĿID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = IIf(IsNull(rsData("����ʱ��")), "", rsData("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(rsData("����")), "", rsData("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                End With
            ElseIf IIf(IsNull(rsData("������")), "", rsData("������")) <> "" Then   '�ѵǼ�
                With Me.vsfList(2)
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData("����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = rsData("����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsData("�Ա�"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�Ա�")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�걾")) = rsData("�걾"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�걾")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsData("������Ŀ"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������Ŀ")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ���")) = IIf(IsNull(rsData("�ͼ���")), "", rsData("�ͼ���")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ���")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ����")) = rsData("�ͼ����"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��")) = rsData("�ͼ�ʱ��"): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ�ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = IIf(IsNull(rsData("������")), "", rsData("������")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = IIf(IsNull(rsData("����ʱ��")), "", rsData("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = IIf(IsNull(rsData("ҽ��ID")), "", rsData("ҽ��ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("ҽ��ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������ĿID")) = IIf(IsNull(rsData("������ĿID")), "", rsData("������ĿID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������ĿID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = IIf(IsNull(rsData("����ʱ��")), "", rsData("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(rsData("����")), "", rsData("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                End With
            End If
            rsData.MoveNext
        Next
        Call showNum
        Me.StatusBar.Panels(1).Text = "�����걾����:" & mlngSampleCount & "��"
        Call setVSFData(strSampleCode)
    ElseIf InStr(strSampleCodesRight, strSampleCode & ",") > 0 Then
        MsgBox "�������Ѿ�����ɨ��������   ", vbInformation
        With Me.vsfList(1)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                    .Select i, 1
                    .ShowCell i, 1
                End If
            Next
        End With
        Exit Function
    ElseIf InStr(strSampleCodesYDJ, strSampleCode & ",") > 0 Then
        MsgBox "�������ѵǼǻ��Ѻ���   ", vbInformation
        With Me.vsfList(2)
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                    .Select i, 1
                    .ShowCell i, 1
                End If
            Next
        End With
        Exit Function
    End If
    
End Function

Private Sub showNum()
    Me.lbl(4).Caption = "���ǼǱ걾(" & Me.vsfList(0).Rows - 1 & ")"
    Me.lbl(5).Caption = "��ɨ��걾(" & Me.vsfList(1).Rows - 1 & ")"
    Me.lbl(6).Caption = "�����ѵǼǻ��Ѻ��ձ걾(" & Me.vsfList(2).Rows - 1 & ")"
End Sub

Private Function ReadData(ByVal strSampleCode As String) As Recordset
    '����               ����ɨ�������Ӧ���������µ���������
    'strSampleCode      ɨ�������
    Dim strSQL As String
    Dim rsSampleCodes As Recordset
    Dim strSampleCodes
        
    strSQL = "Select Distinct a.ҽ��ID,a.�������� ����, b.����, b.�Ա�, b.�걾��λ As �걾," & _
             " b.ҽ������ ������Ŀ,a.�ͼ���,c.���� �ͼ����,b.������Ŀid,a.����ʱ��," & _
             " a.�걾�ͳ�ʱ�� �ͼ�ʱ��, a.�걾��������,a.������,a.����ʱ��,Decode(b.������־, 1, '����', '') As ����" & _
             " From ����ҽ������ A, ����ҽ����¼ B,���ű� C" & _
             " Where a.ҽ��id = b.Id and a.ִ�в���id=c.id and  a.ִ��״̬ in (0)" & _
             " And a.�걾�������� In (Select �걾�������� From ����ҽ������ Where �������� = [1])"
    Set rsSampleCodes = zlDatabase.OpenSQLRecord(strSQL, "�����������", strSampleCode)
        
    Set ReadData = rsSampleCodes
End Function

Private Sub vsfList_Click(Index As Integer)
    Select Case Index
        Case 0
            Set mObjSelectVSF = Me.vsfList(0)
        Case 1
            Set mObjSelectVSF = Me.vsfList(1)
        Case 2
            Set mObjSelectVSF = Me.vsfList(2)
    End Select
End Sub

Private Sub vsfList_DblClick(Index As Integer)
    Dim strSampleCode As String
    
    With vsfList(Index)
        If .MouseRow > 0 Then
            strSampleCode = .TextMatrix(.MouseRow, .ColIndex("����"))
        End If
        Select Case Index
            Case 0
                If getTATTime(.TextMatrix(.MouseRow, .ColIndex("����"))) = False Then
                    Exit Sub
                End If
                Call vsfDataToVsfData(strSampleCode, Me.vsfList(0), Me.vsfList(1))
                Call RemoveHiddenItem(Me.vsfList(0))
            Case 1
                Call vsfDataToVsfData(strSampleCode, Me.vsfList(1), Me.vsfList(0))
                Call RemoveHiddenItem(Me.vsfList(1))
        End Select
    End With
    Call showNum
    Me.StatusBar.Panels(1).Text = "�����걾����:" & mlngSampleCount & "��"
End Sub

Private Sub vsfDataToVsfData(ByVal strSampleCode As String, objVSFFrom As VSFlexGrid, objVSFTo As VSFlexGrid)
    '�����ݴ�һ��VSFת�Ƶ���һ��VSF
    'strSampleCode-����ƥ�������
    'indexFrom-������Դ��VSF����
    'indexTo-Ҫ������ݵ�VSF����
    
    Dim i As Long
    With objVSFFrom
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
                With objVSFTo
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objVSFFrom.TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objVSFFrom.TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = objVSFFrom.TextMatrix(i, .ColIndex("�Ա�")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�Ա�")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�걾")) = objVSFFrom.TextMatrix(i, .ColIndex("�걾")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�걾")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = objVSFFrom.TextMatrix(i, .ColIndex("������Ŀ")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������Ŀ")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ���")) = objVSFFrom.TextMatrix(i, .ColIndex("�ͼ���")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ���")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ����")) = objVSFFrom.TextMatrix(i, .ColIndex("�ͼ����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ����")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��")) = objVSFFrom.TextMatrix(i, .ColIndex("�ͼ�ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("�ͼ�ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������")) = objVSFFrom.TextMatrix(i, .ColIndex("������")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = objVSFFrom.TextMatrix(i, .ColIndex("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("ҽ��ID")) = objVSFFrom.TextMatrix(i, .ColIndex("ҽ��ID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("ҽ��ID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("������ĿID")) = objVSFFrom.TextMatrix(i, .ColIndex("������ĿID")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("������ĿID")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = objVSFFrom.TextMatrix(i, .ColIndex("����ʱ��")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����ʱ��")) = flexAlignLeftCenter
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = objVSFFrom.TextMatrix(i, .ColIndex("����")): .Cell(flexcpAlignment, .Rows - 1, .ColIndex("����")) = flexAlignLeftCenter

                End With
                .RowHidden(i) = True
                Exit For
            End If
        Next
    End With
'    Call RemoveHiddenItem(objVSFFrom)
End Sub

Private Sub RemoveHiddenItem(objVsf As VSFlexGrid)
    Dim i As Long
begin:
    With objVsf
        For i = 1 To .Rows - 1
            If .RowHidden(i) = True Then
                .RemoveItem i
                GoTo begin
            End If
        Next
    End With
End Sub

Private Sub VSFList_RowColChange(Index As Integer)
    With Me.vsfList(Index)
        If .Rows > 1 Then
            If .TextMatrix(1, 1) <> "" Then
                Me.lblInto(0).Caption = .TextMatrix(.RowSel, .ColIndex("�ͼ����"))
                Me.lblInto(1).Caption = .TextMatrix(.RowSel, .ColIndex("�ͼ���"))
                Me.lblInto(2).Caption = .TextMatrix(.RowSel, .ColIndex("�ͼ�ʱ��"))
                Call Form_Resize
            End If
        End If
    End With
End Sub

Private Function getTATTime(ByVal strSampleCode As String) As Boolean
    '���TAT��ʱ
    Dim strSex As String    '�Ա�
    Dim strDept As String   '�������
    Dim strItem As String   '������Ŀ   ��ĿID1,��Ŀ����1,����ʱ��1,����1;��ĿID2,��Ŀ����12,����ʱ��2,����2........
    Dim Record As ReportRecord
    Dim intMsg As Integer
    Dim strMsgShow As String
    Dim strMsgShowStop As String
    Dim strMsgNoTime As String
    Dim strTATItems As String
    Dim var_Tmp As Variant
    Dim var_Tmp1 As Variant
    Dim strShowBef As String
    Dim strItemCode As String
    Dim strItemCodeReplace As String
    Dim i As Integer, j As Integer
    Dim strErr As String

    If mobjLisInsideComm Is Nothing Then
        Set mobjLisInsideComm = CreateObject("zl9LisInsideComm.clsLisInsideComm")
        If Not mobjLisInsideComm Is Nothing Then
            '��ʼ��LIS�ӿڲ���
            If mobjLisInsideComm.InitComponentsHIS(glngSys, glngModul, gcnOracle, strErr) = False Then
                If strErr <> "" Then
                    MsgBox "��ʼ��LIS�ӿ�ʧ�ܣ�" & vbCrLf & strErr
                End If
                Set mobjLisInsideComm = Nothing
            End If
        End If
    End If


    '��ȡ�����Ա���������
    With Me.vsfList(0)
        '��ȡ��ĿID,��Ŀ����,����ʱ��,����
        strItem = ""
        For i = 1 To .Rows - 1
            strSex = .TextMatrix(i, .ColIndex("�Ա�"))
            strDept = .TextMatrix(i, .ColIndex("�ͼ����"))
            If .TextMatrix(i, .ColIndex("����")) = strSampleCode Then
'                If .TextMatrix(i, .ColIndex("�ͼ�ʱ��")) <> "" Then
                    strItem = strItem & ";" & .TextMatrix(i, .ColIndex("������ĿID")) & "," & .TextMatrix(i, .ColIndex("������Ŀ")) & _
                                                "," & .TextMatrix(i, .ColIndex("�ͼ�ʱ��")) & "," & IIf(.TextMatrix(i, .ColIndex("����")) = "����", 1, 0) & _
                                                "," & .TextMatrix(i, .ColIndex("ҽ��ID")) & ",," & .TextMatrix(i, .ColIndex("����"))
'                Else
'                    strMsgNoTime = strMsgNoTime & .TextMatrix(i, .ColIndex("������Ŀ")) & vbCrLf
'                End If
            End If
        Next
        If strMsgNoTime <> "" Then MsgBox strMsgNoTime & "δ�ͼ�,����ǩ��   ", vbInformation, Me.Caption
        If strItem <> "" Then strItem = Mid(strItem, 2)
    End With

    '���TAT�Ƿ�ʱ
    On Error GoTo errold
    strTATItems = mobjLisInsideComm.GetTatTimeShow(2, strItem, strDept, "", "", strSex, intMsg, strShowBef, , UserInfo.����)
    If strTATItems <> "" Then
        var_Tmp = Split(strTATItems, ";")
        Do While UBound(Split(var_Tmp(0), ",")) < 9
            '����9��Ԫ�ص������ƴ��һ��0
            strTATItems = ""
            For i = LBound(var_Tmp) To UBound(var_Tmp)
                strTATItems = strTATItems & ";" & var_Tmp(i) & ",0"
            Next
            If strTATItems <> "" Then strTATItems = Mid(strTATItems, 2)
            var_Tmp = Split(strTATItems, ";")
        Loop
        
        '��ȡ������Ŀ������
        For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 Then
                strItemCode = strItemCode & "," & Split(var_Tmp(i), ",")(6)
            End If
        Next
        
'        strIDs = ""
        
       For i = LBound(var_Tmp) To UBound(var_Tmp)
            If Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '�ѳ�ʱֻ��ʾ
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", Split(var_Tmp(i), ",")(7) & "����") & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 1 And InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) > 0 And Split(var_Tmp(i), ",")(2) <> "" Then
                '����ͬ������Ŀ��
                strMsgShow = strMsgShow & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", "") & "����ͬ�����ֹ��Ŀ,���ܼ���" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(8) <> "0" And Split(var_Tmp(i), ",")(2) = "" Then
                'û��ǰһ��ʱ��ڵ��
                strMsgShowStop = strMsgShowStop & Split(var_Tmp(i), ",")(1) & "δ�ͼ�,����ǩ��" & vbCrLf
            ElseIf Split(var_Tmp(i), ",")(7) <> 0 And Split(var_Tmp(i), ",")(9) = 2 And Split(var_Tmp(i), ",")(2) <> "" Then
                '��ʱ����ֹ��
                strMsgShowStop = strMsgShowStop & Replace(Replace(Split(var_Tmp(i), ",")(8), "[��Ŀ]", Split(var_Tmp(i), ",")(1)), "[��ʱ]", Split(var_Tmp(i), ",")(7) & "����") & vbCrLf
'            Else
'                '��ͬ��Ŀͬ�����ʱ��,����һ����Ŀ��ʱ,�����и��������Ŀ�������ͼ�
'                If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) <= 0 Then
'                    strIDs = strIDs & "," & Split(var_Tmp(i), ",")(4) & "," & Split(var_Tmp(i), ",")(5)
'                End If
            End If
        Next
        
'        If strIDs <> "" Then
'            strIDs = Mid(strIDs, 2)
'        End If
        
        '������Ϊ��ʾʱ,�������ʱ,���ͼ����й�ѡ����Ŀ,���˷�,��ֻ�ͼ�Ϊ��ʱ�ı걾
        If strMsgShow <> "" Then
            If MsgBox(strMsgShow & "�Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                getTATTime = False
                Exit Function
            Else
                '�����,�����������й�ѡ����Ŀ
'                strIDs = ""
'                For i = LBound(var_Tmp) To UBound(var_Tmp)
'                    If InStr(strItemCode, "," & Split(var_Tmp(i), ",")(6)) > 0 Then
'                        MsgBox Split(var_Tmp(i), ",")(1) & "����ͬ�����ֹ��Ŀ,���ܼ���", vbInformation, Me.Caption
'                        getTATTime = False
'                        Exit Function
'                    End If
'                Next
            End If
        End If
        If strMsgShowStop <> "" Then
            MsgBox strMsgShowStop, vbInformation, Me.Caption
            getTATTime = False
            Exit Function
        End If
        
    End If
    getTATTime = True
    
    Exit Function
errold:
    getTATTime = True
    
    
    Exit Function
ErrHand:
    MsgBox Err.Description, vbInformation, Me.Caption
    Err.Clear

End Function





