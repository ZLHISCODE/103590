VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "Vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDiagnotor 
   Caption         =   "����ɸ������"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7170
   Icon            =   "frmDiagnotor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7170
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6420
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiagnotor.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7594
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
   Begin VSFlex8Ctl.VSFlexGrid hgdResult 
      Height          =   3585
      Left            =   15
      TabIndex        =   9
      Top             =   2835
      Width           =   7095
      _cx             =   12515
      _cy             =   6324
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDiagnotor.frx":115C
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCodex 
      Height          =   1635
      Left            =   3150
      TabIndex        =   18
      Top             =   2700
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   2884
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   2745
      Left            =   -3840
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   330
      Visible         =   0   'False
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   4842
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClass 
      Caption         =   "��ҽ������Ŀ(&N)"
      Height          =   350
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1500
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   2415
      Width           =   5625
   End
   Begin VB.CommandButton cmdClass 
      Caption         =   "��ҽ������Ŀ(&W)"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   510
      Width           =   1500
   End
   Begin VB.TextBox txtDays 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4230
      TabIndex        =   7
      Text            =   "30"
      Top             =   1650
      Width           =   540
   End
   Begin VB.OptionButton optSource 
      Caption         =   "ָ��ʱ���ڵĲ�������(&2)         ������"
      Height          =   225
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   1695
      Width           =   3735
   End
   Begin VB.OptionButton optSource 
      Caption         =   "���˱��ξ��ﲡ������(&1)"
      Height          =   225
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Top             =   1365
      Value           =   -1  'True
      Width           =   2700
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   375
      Width           =   5625
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "ɸ������(&F)"
      Height          =   350
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   1500
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   5925
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   5925
      TabIndex        =   10
      Top             =   120
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -375
      Top             =   1455
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnotor.frx":1185
            Key             =   "CLASS"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdItems 
      Height          =   765
      Left            =   4755
      TabIndex        =   19
      Top             =   2625
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label lblClass 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�ڿ�"
      Height          =   300
      Index           =   1
      Left            =   1665
      TabIndex        =   3
      Top             =   945
      Width           =   3930
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      Caption         =   "��ɸ�����������������¼�������:"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   2580
      Width           =   2790
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   6285
      Picture         =   "frmDiagnotor.frx":171F
      Top             =   2205
      Width           =   480
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ID��1356    ��������ѧ�    �Ա�Ů    ���䣺65"
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   165
      TabIndex        =   14
      Top             =   120
      Width           =   4680
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      Caption         =   "�������ݷ�Χ(&R)"
      Height          =   180
      Left            =   165
      TabIndex        =   4
      Top             =   1380
      Width           =   1350
   End
   Begin VB.Label lblClass 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ѫ�ܼ���"
      Height          =   300
      Index           =   0
      Left            =   1665
      TabIndex        =   1
      Top             =   555
      Width           =   3930
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuPreview 
         Caption         =   "Ԥ��(V)"
      End
      Begin VB.Menu mnuPopuPrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPopuExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuPopuSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuCopy 
         Caption         =   "����(C)"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmDiagnotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiId As Long      '���м���ɸ��Ĳ���ID
Private mblnInpati As Boolean   '�Ƿ�סԺ���ˣ�����Ϊ���ﲡ��
Private mlngPageId As Long      '���м���ɸ��Ĳ�����ҳID��סԺ����ʹ�ã�
Private mstrRegist As String    '���м���ɸ��ĹҺŵ��ţ����ﲡ��ʹ�ã�

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String, aryTemp() As String

Const conCol������ As Integer = 0
Const conCol���ID As Integer = 1
Const conCol������� As Integer = 2
Const conCol���ͱ�֤ As Integer = 3
Const conCol���ɳ̶� As Integer = 4
Const conCol�������� As Integer = 5

Private WithEvents objParentForm As Form
Attribute objParentForm.VB_VarHelpID = -1

Public Sub ShowMe(ByVal bytModal As Byte, ByVal frmParent As Object, _
    ByVal lngPatiId As Long, ByVal blnInpati As Boolean, _
    Optional ByVal lngPageId As Long = 1, Optional ByVal strRegist As String)
    '---------------------------------------------
    '���ܣ������ϼ�����Ҫ�󣬵��ò��˼���ɸ����򣬹�ҽ���ο�
    '��Σ�frmParent-�����壻
    '      blnModal-�Ƿ�ģ̬��ʾ��ͨ�����ϼ�����һ�£���
    '      lngPatiId-Ҫ��ʾ�Ĳ���ID��
    '      blnInpati-�Ƿ�סԺ���ˣ�����Ϊ���ﲡ�ˣ�
    '      lngPageId-Ҫ��ʾ��סԺ������ҳID��
    '      strRegist-Ҫ��ʾ�����ﲡ�˹Һŵ��ţ�
    '---------------------------------------------
    mlngPatiId = lngPatiId: mblnInpati = blnInpati
    mlngPageId = lngPageId: mstrRegist = strRegist
    
    gstrSql = "select ����ID,�����,סԺ��,����,�Ա�,����" & _
            " from ������Ϣ" & _
                " where ����id=" & lngPatiId
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.Title, Me.Caption, gstrSql): .Open gstrSql, gcnOracle: Call SQLTest
        If .RecordCount > 0 Then
            Me.lblPati.Caption = "����ID��" & lngPatiId & _
                    Space(3) & "������" & !���� & _
                    Space(3) & "�Ա�" & IIf(IsNull(!�Ա�), "", !�Ա�) & _
                    Space(3) & "���䣺" & IIf(IsNull(!����), "", !����)
            Me.lblPati.Tag = !����
        Else
            MsgBox "ָ�����˲����ڣ�", vbExclamation, gstrSysName: Unload Me: Exit Sub
        End If
    End With
    
    '��ȡϰ����Ͽ�Ŀ
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\����ɸ���Ŀ\", "��ҽ", "")
        gstrSql = "select ID,����,����" & _
                " From ������Ϸ���" & _
                " Where ��� =1 and ID in (" & IIf(Trim(strTemp) = "", 0, strTemp) & ")" & _
                " order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.lblClass(0).Caption = "": Me.lblClass(0).Tag = ""
        Do While Not .EOF
            Me.lblClass(0).Caption = Me.lblClass(0).Caption & "; [" & !���� & "]" & !����
            Me.lblClass(0).Tag = Me.lblClass(0).Tag & "," & !ID
            .MoveNext
        Loop
        If Me.lblClass(0).Caption <> "" Then
            Me.lblClass(0).Caption = Mid(Me.lblClass(0).Caption, 3)
            Me.lblClass(0).Tag = Trim(Mid(Me.lblClass(0).Tag, 2))
        End If
        
        strTemp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\����ɸ���Ŀ\", "��ҽ", "")
        gstrSql = "select ID,����,����" & _
                " From ������Ϸ���" & _
                " Where ��� =2 and ID in (" & IIf(Trim(strTemp) = "", 0, strTemp) & ")" & _
                " order by ����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.lblClass(1).Caption = "": Me.lblClass(1).Tag = ""
        Do While Not .EOF
            Me.lblClass(1).Caption = Me.lblClass(1).Caption & "; [" & !���� & "]" & !����
            Me.lblClass(1).Tag = Me.lblClass(1).Tag & "," & !ID
            .MoveNext
        Loop
        If Me.lblClass(1).Caption <> "" Then
            Me.lblClass(1).Caption = Mid(Me.lblClass(1).Caption, 3)
            Me.lblClass(1).Tag = Trim(Mid(Me.lblClass(1).Tag, 2))
        End If
    End With
    
    '��ʾ����
    On Error Resume Next
    Set objParentForm = frmParent
    Me.Show bytModal, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClass_Click(Index As Integer)
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select ID,�ϼ�ID,����,����,����" & _
                " From ������Ϸ���" & _
                " Where ��� = " & IIf(Index = 0, 1, 2) & _
                " start with �ϼ�ID is null" & _
                " connect by prior ID=�ϼ�ID"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "CLASS")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "CLASS")
            End If
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error GoTo 0
    With Me.tvwClass
        .Tag = Index
        aryTemp = Split(Me.lblClass(Index).Tag, ",")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            .Nodes("_" & aryTemp(intCount)).Selected = True
            .SelectedItem.Checked = True
            .SelectedItem.EnsureVisible
        Next
        .Left = Me.lblClass(Index).Left: .Width = Me.lblClass(Index).Width
        .Top = Me.lblClass(Index).Top + Me.lblClass(Index).Height
        .ZOrder 0
        .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If Trim(Me.lblClass(0).Tag) = "" And Trim(Me.lblClass(1).Tag) = "" Then
        MsgBox "δѡ�񼲲����ɸ��ķ�Χ��", vbExclamation, gstrSysName
        Me.cmdClass(0).SetFocus
        Exit Sub
    End If
    
    '��ȡ�Ѿ��������ļ���Ŀ¼
    Me.stbThis.Panels(2).Text = "����ȷ��ɸ�鷶Χ..."
    Err = 0: On Error GoTo ErrHand
    With rsTemp
        gstrSql = "select distinct I.���,I.ID,I.����,I.����,I.����,I.�ٴ�,X.�����,X.������" & _
                " from (select distinct ID" & _
                "       from ������Ϸ��� C" & _
                "       start with ���=1 and ID in (" & IIf(Me.lblClass(0).Tag = "", "0", Me.lblClass(0).Tag) & ")" & _
                "               or ���=2 and ID in (" & IIf(Me.lblClass(1).Tag = "", "0", Me.lblClass(1).Tag) & ")" & _
                "       connect by prior ID=�ϼ�id) C," & _
                "      ����������� R,�������Ŀ¼ I,������Ϲ��� X" & _
                " Where C.ID=R.����ID and R.���ID=I.ID and I.ID = X.���ID" & _
                " order by I.���,I.����"
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
    End With
    Me.hgdCodex.Clear: Set Me.hgdCodex.Recordset = rsTemp
    
    Dim lngDegree As Long   '������
    Dim strGists As String  '��������
    
    '�����������ɸ��
    Me.hgdResult.Rows = Me.hgdResult.FixedRows
    For intCount = 0 To Me.hgdCodex.Rows - 1
        Me.stbThis.Panels(2).Text = "���ڽ���" & Me.hgdCodex.TextMatrix(intCount, 3) & "����..."
        '��ȡ��Ҫ�жϵ�����
        gstrSql = "select X.��ĿID,I.������,I.Ӣ����,I.����,X.��ϵʽ,X.����ֵ,X.���ɶ�" & _
                " from ������Ϲ��� X,����������Ŀ I" & _
                " where X.��ĿID=I.ID " & _
                "       and X.���ID=" & Me.hgdCodex.TextMatrix(intCount, 1) & _
                "       and nvl(�����,0)=" & Val(Me.hgdCodex.TextMatrix(intCount, 6)) & _
                " order by X.������"
        With rsTemp
            If .State = adStateOpen Then .Close
            Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
        End With
        Me.hgdItems.Clear: Set Me.hgdItems.Recordset = rsTemp
        
        lngDegree = 0: strGists = ""
        For intRow = 0 To Me.hgdItems.Rows - 1
            '��ȡָ����Ŀ������¼������
            gstrSql = "select S.��������" & _
                    " from ���˲��������� S" & _
                    " where S.������ID=" & Me.hgdItems.TextMatrix(intRow, 0) & _
                    "       and S.����ID=(" & _
                    "           select max(S.����ID)" & _
                    "           from (select S.����ID from ���˲��������� S where S.������ID=" & Me.hgdItems.TextMatrix(intRow, 0) & ") S," & _
                    "                (select C.ID,C.������¼ID " & _
                    "                 from ���˲�����¼ L,���˲������� C" & _
                    "                 where L.ID=C.������¼ID" & _
                    "                       and L.����ID=" & mlngPatiId
            If Me.optSource(0).Value = True And mblnInpati = True Then
                gstrSql = gstrSql & "       and L.��ҳid=" & mlngPageId & ") C" & _
                    "           where C.ID=S.����ID)"
            ElseIf Me.optSource(0).Value = True And mblnInpati = False Then
                gstrSql = gstrSql & "       and L.�Һŵ�='" & mstrRegist & "') C" & _
                    "           where C.ID=S.����ID)"
            Else
                gstrSql = gstrSql & "       and L.��д����>=trunc(sysdate-" & Val(Me.txtDays.Text) & ")) C" & _
                    "           where C.ID=S.����ID)"
            End If
            With rsTemp
                If .State = adStateOpen Then .Close
                Call SQLTest(App.ProductName, Me.Caption, gstrSql): .Open gstrSql, gcnOracle, adOpenStatic, adLockReadOnly: Call SQLTest
                If Not .EOF Then
                    strTemp = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
                    '���������ж�
                    If zlVerifyValue(strTemp, Val(Me.hgdItems.TextMatrix(intRow, 3)), Me.hgdItems.TextMatrix(intRow, 4), Me.hgdItems.TextMatrix(intRow, 5)) Then
                        lngDegree = lngDegree + Val(Me.hgdItems.TextMatrix(intRow, 6))
                        strGists = strGists & vbCrLf & Me.hgdItems.TextMatrix(intRow, 1) & _
                                IIf(Trim(Me.hgdItems.TextMatrix(intRow, 2)) = "", "", "(" & Me.hgdItems.TextMatrix(intRow, 2) & ")") & _
                                "Ϊ" & strTemp
                        strGists = strGists & "��" & Me.hgdItems.TextMatrix(intRow, 4) & Me.hgdItems.TextMatrix(intRow, 5) & "��"
                    End If
                End If
            End With
        Next
        If strGists <> "" Then strGists = Mid(strGists, 3)
        If Val(Me.hgdCodex.TextMatrix(intCount, 4)) <> 0 And lngDegree >= Val(Me.hgdCodex.TextMatrix(intCount, 4)) _
           Or Val(Me.hgdCodex.TextMatrix(intCount, 5)) <> 0 And lngDegree >= Val(Me.hgdCodex.TextMatrix(intCount, 5)) Then
            With Me.hgdResult
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, conCol������) = .Rows - 1
                .TextMatrix(.Rows - 1, conCol���ID) = Me.hgdCodex.TextMatrix(intCount, 1)
                .TextMatrix(.Rows - 1, conCol�������) = Me.hgdCodex.TextMatrix(intCount, 3)
                .TextMatrix(.Rows - 1, conCol���ͱ�֤) = Me.hgdCodex.TextMatrix(intCount, 7)
                .TextMatrix(.Rows - 1, conCol��������) = strGists
            If Val(Me.hgdCodex.TextMatrix(intCount, 5)) <> 0 And lngDegree >= Val(Me.hgdCodex.TextMatrix(intCount, 5)) Then
                .TextMatrix(.Rows - 1, conCol���ɳ̶�) = "�ٴ�"
            Else
                .TextMatrix(.Rows - 1, conCol���ɳ̶�) = "����"
            End If
            End With
        End If
    Next
    With Me.hgdResult
        If .Rows > .FixedRows Then
            Call .AutoSize(conCol��������)
        Else
            .Rows = .FixedRows + 1
            .TextMatrix(.Rows - 1, conCol�������) = "δɸ�鵽����..."
        End If
        .SetFocus
    End With
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    '���ý���Ԫ����̬
    With Me.hgdResult
        .Rows = 2: .Cols = 6
        .ColWidth(conCol������) = 280
        .TextMatrix(0, conCol���ID) = "���ID": .ColWidth(conCol���ID) = 0
        .TextMatrix(0, conCol�������) = "���": .ColWidth(conCol�������) = 2500
        .TextMatrix(0, conCol���ͱ�֤) = "����/��֤": .ColWidth(conCol���ͱ�֤) = 1300
        .TextMatrix(0, conCol���ɳ̶�) = "���ɶ�": .ColWidth(conCol���ɳ̶�) = 700
        .TextMatrix(0, conCol��������) = "��������": .ColWidth(conCol��������) = 3600
        For intCol = 0 To .Cols - 1
            .FixedAlignment(intCol) = flexAlignCenterCenter
        Next
    End With
    
End Sub

Private Sub Form_Resize()
    Dim lngStatus As Single
    
    If WindowState = 1 Then Exit Sub
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    Err = 0: On Error Resume Next
    Me.cmdClose.Left = Me.ScaleWidth - Me.cmdClose.Width - 90
    Me.cmdHelp.Left = Me.cmdClose.Left
    Me.imgLogo.Left = Me.cmdClose.Left + (Me.cmdClose.Width - Me.imgLogo.Width) / 2
        
    Me.fraLine(0).Width = Me.cmdClose.Left - Me.fraLine(0).Left - 180
    Me.fraLine(1).Width = Me.cmdClose.Left - Me.fraLine(1).Left - 180
    
    Me.lblClass(0).Width = Me.cmdClose.Left - Me.lblClass(0).Left - 180
    Me.lblClass(1).Width = Me.cmdClose.Left - Me.lblClass(1).Left - 180
    
    With Me.hgdResult
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - lngStatus
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\����ɸ���Ŀ\", "��ҽ", Me.lblClass(0).Tag)
    Call SaveSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\����ɸ���Ŀ\", "��ҽ", Me.lblClass(1).Tag)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub hgdResult_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call Me.hgdResult.AutoSize(conCol��������)
End Sub

Private Sub hgdResult_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Call PopupMenu(Me.mnuPopu, 2)
End Sub

Private Sub mnuPopuCopy_Click()
    strTemp = ""
    With Me.hgdResult
        For intRow = .FixedRows To .Rows - 1
            For intCol = 0 To .Cols - 1
                If intCol = .Cols - 1 Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & vbCrLf
                Else
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & vbTab
                End If
            Next
        Next
    End With
    If strTemp <> "" Then
        VB.Clipboard.Clear
        VB.Clipboard.SetText strTemp
    End If
End Sub

Private Sub mnuPopuExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuPopuPreview_Click()
    Call zlRptPrint(2)
End Sub

Private Sub mnuPopuPrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub objParentForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub optSource_Click(Index As Integer)
    If Me.optSource(1).Value = True Then
        Me.txtDays.Enabled = True: Me.txtDays.SetFocus
    Else
        Me.txtDays.Enabled = False
    End If
End Sub

Private Sub tvwClass_DblClick()
    intCount = Val(Me.tvwClass.Tag)
    Me.lblClass(intCount).Caption = "": Me.lblClass(intCount).Tag = ""
    For Each objNode In Me.tvwClass.Nodes
        If objNode.Checked = True Then
            Me.lblClass(intCount).Caption = Me.lblClass(intCount).Caption & "; " & objNode.Text
            Me.lblClass(intCount).Tag = Me.lblClass(intCount).Tag & "," & Mid(objNode.Key, 2)
        End If
    Next
    If Me.lblClass(intCount).Caption <> "" Then
        Me.lblClass(intCount).Caption = Mid(Me.lblClass(intCount).Caption, 3)
        Me.lblClass(intCount).Tag = Trim(Mid(Me.lblClass(intCount).Tag, 2))
    End If
    Me.cmdClass(Val(Me.tvwClass.Tag)).SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    Me.tvwClass.Visible = False
End Sub

Private Sub txtDays_GotFocus()
    Me.txtDays.SelStart = 0: Me.txtDays.SelLength = 100
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Function zlVerifyValue(strVerify As String, bytType As Byte, strFormula As String, strAskValue) As Boolean
    '-------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������������ʽ
    '��Σ� strVerify-���жϵ���ֵ
    '       bytType-��ֵ����
    '       strFormula-��ϵʽ������˵����
    '       strAskValue-Ҫ�����ֵ��Χ��
    '���Σ���ȷ����true�����򷵻�false
    '-------------------------------------------------
    zlVerifyValue = False
    
    Select Case Val(bytType)
    Case 0  '��ֵ
        Select Case Trim(strFormula)
        Case "����"
            If Val(strVerify) = Val(strAskValue) Then zlVerifyValue = True
        Case "������"
            If Val(strVerify) <> Val(strAskValue) Then zlVerifyValue = True
        Case "����"
            If Val(strVerify) > Val(strAskValue) Then zlVerifyValue = True
        Case "С��"
            If Val(strVerify) < Val(strAskValue) Then zlVerifyValue = True
        Case "����"
            If Val(strVerify) <= Val(strAskValue) Then zlVerifyValue = True
        Case "����"
            If Val(strVerify) >= Val(strAskValue) Then zlVerifyValue = True
        Case "����"
            aryTemp = Split(strAskValue, ",")
            If UBound(aryTemp) = 1 Then
                If Val(strVerify) >= Val(aryTemp(0)) And Val(strVerify) <= Val(aryTemp(1)) Then zlVerifyValue = True
                If Val(strVerify) >= Val(aryTemp(1)) And Val(strVerify) <= Val(aryTemp(0)) Then zlVerifyValue = True
            End If
        Case "����"
            strAskValue = Replace(strAskValue, Space(1), "")
            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") > 0 Then zlVerifyValue = True
        Case "������"
            strAskValue = Replace(strAskValue, Space(1), "")
            If InStr(1, "," & strAskValue & ",", "," & Val(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
    Case 1  '����
        Select Case Trim(strFormula)
        Case "����"
            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        Case "������"
            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
        Case "����"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) > 0 Then zlVerifyValue = True
        Case "������"
            If InStr(1, Trim(strVerify), Trim(strAskValue)) = 0 Then zlVerifyValue = True
        Case "����"
            strAskValue = Replace(strAskValue, Space(1), "")
            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") > 0 Then zlVerifyValue = True
        Case "������"
            strAskValue = Replace(strAskValue, Space(1), "")
            If InStr(1, "," & strAskValue & ",", "," & Trim(strVerify) & ",") = 0 Then zlVerifyValue = True
        End Select
    Case 2  '����
        strVerify = Format(strVerify, "YYYY-MM-DD")
        Select Case Trim(strFormula)
        Case "����"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) = Trim(strAskValue) Then zlVerifyValue = True
        Case "������"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) <> Trim(strAskValue) Then zlVerifyValue = True
        Case "����"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) > Trim(strAskValue) Then zlVerifyValue = True
        Case "����"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) < Trim(strAskValue) Then zlVerifyValue = True
        Case "������"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) <= Trim(strAskValue) Then zlVerifyValue = True
        Case "������"
            strAskValue = Format(strAskValue, "YYYY-MM-DD")
            If Trim(strVerify) >= Trim(strAskValue) Then zlVerifyValue = True
        Case "����"
            aryTemp = Split(strAskValue, ",")
            If UBound(aryTemp) = 1 Then
                aryTemp(0) = Format(aryTemp(0), "YYYY-MM-DD")
                aryTemp(1) = Format(aryTemp(1), "YYYY-MM-DD")
                If Trim(strVerify) >= Trim(aryTemp(0)) And Trim(strVerify) <= Trim(aryTemp(1)) Then zlVerifyValue = True
                If Trim(strVerify) >= Trim(aryTemp(1)) And Trim(strVerify) <= Trim(aryTemp(0)) Then zlVerifyValue = True
            End If
        End Select
    Case 3  '�߼�
        Select Case Trim(strFormula)
        Case "��"
            If Val(strVerify) = 1 Then zlVerifyValue = True
        Case "��"
            If Val(strVerify) = 0 Then zlVerifyValue = True
        End Select
    Case Else
    End Select
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:��¼���ӡ
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    On Error Resume Next
    objPrint.Title.Text = "��" & Me.lblPati.Tag & "����������ɸ��"
    objPrint.Title.Font.Size = 11
    Set objPrint.Body = Me.hgdResult
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub


