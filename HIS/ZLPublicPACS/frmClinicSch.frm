VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicSch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ٴ����ԤԼ"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   Icon            =   "frmClinicSch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame frmSchedInfo 
      Caption         =   "�˼�������ԤԼ��ԤԼ��Ϣ���£�"
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   9015
      Begin VB.Label lblSchedInfo 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   8655
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "�����"
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "���� >>"
      Height          =   375
      Index           =   5
      Left            =   8030
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "����"
      Height          =   375
      Index           =   3
      Left            =   5640
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "����"
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "<< ����"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ԤԼ"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   7200
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSchSegment 
      Height          =   1530
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   9015
      _cx             =   15901
      _cy             =   2699
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSchDate 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   9015
      _cx             =   15901
      _cy             =   4895
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
      Rows            =   1
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox cboSchDevice 
         Height          =   300
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblSchTime 
         Caption         =   "ԤԼʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lblSchDate 
         Caption         =   "ԤԼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   671
         Width           =   1335
      End
      Begin VB.Label lblSchDevice 
         Caption         =   "ԤԼ�豸"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   263
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "ԤԼʱ��Σ�"
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "ԤԼ���ڣ�"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   671
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "ԤԼ�豸��"
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   263
         Width           =   975
      End
      Begin VB.Label lblOrder 
         Caption         =   "ҽ������"
         Height          =   675
         Left            =   1320
         TabIndex        =   4
         Top             =   671
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "ҽ�����ݣ�"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   671
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "����"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   263
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "����������"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   263
         Width           =   975
      End
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Caption         =   "2019��6��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   1500
      Width           =   2055
   End
End
Attribute VB_Name = "frmClinicSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long            '�򿪴���ʱ��ҽ��ID
Private mblnCanSchedule As Boolean      '�ҵ�ԤԼʱ����豸
Private mblnRefreshDevice As Boolean    '�Ƿ�ˢ��ԤԼ�豸
Private mblnOK As Boolean               'ԤԼ�ɹ�

'XML���ص�����
Private mstrOrderID As String
Private mstr������ĿID As String
Private mstr������Ŀ���� As String
Private mstrҽ������ As String
Private mstrԤԼ�豸���� As String
Private mstrԤԼ�豸ID As String
Private mstrԤԼ���� As String
Private mstrԤԼ��ʼʱ�� As String
Private mstrԤԼ����ʱ�� As String
Private mrsTimes As ADODB.Recordset
    
'�������
Private Enum schDateColTitle
    col_SchDate_��һ = 0
    col_SchDate_�ܶ� = 1
    col_SchDate_���� = 2
    col_SchDate_���� = 3
    col_SchDate_���� = 4
    col_SchDate_���� = 5
    col_SchDate_���� = 6
End Enum

'���ʱ���
Private Enum schTimeSegColTitle
    col_SchTimeSeg_��� = 0
    col_SchTimeSeg_�豸 = 1
    col_SchTimeSeg_��ʼʱ�� = 2
    col_SchTimeSeg_����ʱ�� = 3
End Enum

Public Function zlShowMe(objParent As Object, blnShowModal As Boolean, lngAdviceID As Long, blnModify As Boolean) As Boolean
'-----------------------------------------------------------
'����:��ʾ�ٴ����ԤԼ����
'���:  objParent -- ������
'       blnShowModal -- �Ƿ�ģʽ����
'       lngAdviceID -- ҽ��ID
'       blnModify -- �޸�ԤԼ
'����:
'-----------------------------------------------------------
    
    On Error GoTo err
    
    mblnOK = False
    mlngAdviceID = lngAdviceID
            
    If refreshDate(Format(Now, "YYYY-MM-DD")) = False Then
        If mblnCanSchedule = False Then
            zlShowMe = True
        Else
            zlShowMe = False
        End If
        Unload Me
        Exit Function
    End If
    
    Call loadPatInfo             '�Ȳ�ѯ���ں���ػ��߻�����Ϣ
    
    If blnModify = True Then
        frmSchedInfo.Visible = True
        Call loadSchedInfo
        Me.Caption = "�޸��ٴ����ԤԼ"
    Else
        frmSchedInfo.Visible = False
        vsfSchSegment.Height = vsfSchSegment.Height + frmSchedInfo.Height
    End If
    
    mblnRefreshDevice = True
    
    Call Show(IIf(blnShowModal, 1, 0), objParent)
    
    zlShowMe = mblnOK
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cboSchDevice_Click()
On Error GoTo err
    If mblnRefreshDevice = True Then
        mstrԤԼ�豸ID = cboSchDevice.ItemData(cboSchDevice.ListIndex)
        cmdChangeDate_Click (1) 'ˢ�½��������
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeDate_Click(index As Integer)
    Dim dtDate As Date
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Integer
    Dim j As Integer
    Dim strSelDate As String
On Error GoTo err
    strSelDate = ""
    'index =0-���£�1-���죻2-���죻3-���죻4-����죻5-����
    Select Case index
    Case 0:     '����
        dtDate = Format(mstrԤԼ��ʼʱ��, "YYYY-MM-DD")
        If GetOtherMonth(False, dtDate) = False Then
            Exit Sub
        End If
        
    Case 1:  '����
        dtDate = Now
    Case 2:     '����
        dtDate = Now + 1
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 3:     '����
        dtDate = Now + 2
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 4:     '�����
        dtDate = Now + 3
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 5:     '����
        dtDate = Format(mstrԤԼ��ʼʱ��, "YYYY-MM-DD")
        If GetOtherMonth(True, dtDate) = False Then
            Exit Sub
        End If
    End Select
    
    Call refreshDate(dtDate, CLng(mstrԤԼ�豸ID), strSelDate)
    
    If index = 1 Or index = 2 Or index = 3 Or index = 4 Then
        If Format(dtDate, "YYYY-MM-DD") <> Format(mstrԤԼ����, "YYYY-MM-DD") Then
            MsgBox IIf(index = 1, "����", IIf(index = 2, "����", IIf(index = 3, "����", "�����"))) _
                & "ԤԼ������ֻ�ܴ�" & mstrԤԼ���� & "��ʼԤԼ��", vbOKOnly, "���ԤԼ"
        End If
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strIn As String
    Dim strOut As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strResult As String
    
    On Error GoTo err
    
    strIn = "<IN><ADVICEID>" & mstrOrderID & "</ADVICEID><MACHINEID>" _
        & mstrԤԼ�豸ID & "</MACHINEID><SCHBEGINTIME>" & mstrԤԼ��ʼʱ�� _
        & "</SCHBEGINTIME><SCHENDTIME>" & mstrԤԼ����ʱ�� & "</SCHENDTIME></IN>"
    strOut = gobjComLib.zlDatabase.CallProcedure("zl_Ӱ��ԤԼ_ScheduleInsert", Me.Caption, strIn, Empty)
    
    '�������ص�XML��
    If strOut = "" Then
        MsgBox "����ԤԼ��Ϣ����������ѡ��ʱ��κ��ٴ�ԤԼ��", vbOKOnly, "���ԤԼ"
        'ˢ��ԤԼ����
        Call refreshDate(CDate(Format(mstrԤԼ����, "YYYY-MM-DD")))
    End If
    '  --�ɹ���
    '  --<OUTPUT>
    '  --  <RESULT>true</RESULT>
    '  --</OUTPUT>
    '
    '  --ʧ�ܣ�
    '  --<OUTPUT>
    '  --  <RESULT>false</RESULT>
    '  --  <ERROR>
    '  --    <MSG>��ϸ������ʾ</MSG>
    '  --  </ERROR>
    '  --</OUTPUT>
    Set objXml = CreateObject("zl9ComLib.clsXML")
    Call objXml.OpenXMLDocument(strOut)
    Call objXml.GetSingleNodeValue("RESULT", strResult)
    If strResult = "true" Then
        mblnOK = True
        'ԤԼ�ɹ�����ӡԤԼ�����˳�
        Call PrintSchedule(Me, mlngAdviceID)
        Unload Me
    Else
        'ԤԼʧ�ܣ���ʾ��ˢ���б�
        Call objXml.GetSingleNodeValue("MSG", strResult)
        If InStr(strResult, "[ZLSOFT]") > 0 Then
            strResult = Split(strResult, "[ZLSOFT]")(1)
        End If
        MsgBox "����ԤԼ��Ϣ����������ѡ��ʱ��κ��ٴ�ԤԼ��" & vbCrLf & vbCrLf _
            & "������Ϣ��" & strResult, vbOKOnly, "���ԤԼ"
        'ˢ��ԤԼ����
        Call refreshDate(CDate(Format(mstrԤԼ����, "YYYY-MM-DD")))
    End If
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadPatInfo()
'-----------------------------------------------------------
'����:���ػ��߻�����Ϣ
'���:
'����:
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo err
    
    lngIndex = -1
    
    strSQL = "select a.����,nvl(a.Ӥ��,0) as Ӥ��,a.ҽ������,a.����ID,a.��ҳID from ����ҽ����¼ a where id =[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ϣ", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        If rsTemp!Ӥ�� <> 0 Then
            strSQL = "Select Decode(a.Ӥ������, Null, b.���� || '֮��' || Trim(To_Char(a.���, '9')), a.Ӥ������) As Ӥ������, " _
                    & " From ������������¼ A, ������Ϣ B Where a.����id = [ 1 ] And a.��ҳid = [ 2 ] " _
                    & " And a.����id = b.����id And a.��� = [ 3 ]"
            Set rsBaby = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ϣ", rsTemp!����ID, rsTemp!��ҳID, rsTemp!Ӥ��)
            lblName.Caption = rsBaby!Ӥ������
        Else
            lblName.Caption = rsTemp!����
        End If
        lblOrder.Caption = rsTemp!ҽ������
    End If
    
    strSQL = "select b.id ,b.�豸���� from Ӱ��ԤԼ��Ŀ A ,Ӱ��ԤԼ�豸 B,����ҽ����¼ C,Ӱ��ԤԼ���� D WHERE c.id=[1] " _
            & " and c.������Ŀid = a.������Ŀid and a.ԤԼ�豸id = b.id and b.�Ƿ�����=1 and B.ID=D.ԤԼ�豸ID(+) and D.�Ƿ�����=1 "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ�豸", mlngAdviceID)
    
    If rsTemp.RecordCount > 1 Then
        cboSchDevice.Clear
        While rsTemp.EOF = False
            cboSchDevice.AddItem rsTemp!�豸����
            cboSchDevice.ItemData(cboSchDevice.NewIndex) = rsTemp!ID
            If rsTemp!ID = mstrԤԼ�豸ID Then
                lngIndex = cboSchDevice.ListCount - 1
            End If
            rsTemp.MoveNext
        Wend
        If lngIndex <> -1 Then
            cboSchDevice.ListIndex = lngIndex
        Else
            cboSchDevice.ListIndex = 0
        End If
        
        cboSchDevice.Visible = True
        lblSchDevice.Visible = False
    Else
        cboSchDevice.Visible = False
        lblSchDevice.Visible = True
    End If
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadSchedInfo()
'-----------------------------------------------------------
'����:�����Ѿ�ԤԼ����Ϣ
'���:
'����:
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err

    strSQL = "SELECT a.ԤԼ�豸����,a.��������,a.ԤԼ��ʼʱ��,a.ԤԼ����ʱ�� FROM Ӱ��ԤԼ��¼ a where a.ҽ��ID = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ԤԼ��Ϣ", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        lblSchedInfo = "ԤԼ�豸��" & rsTemp!ԤԼ�豸���� & "     ԤԼ���ڣ�" & Format(rsTemp!ԤԼ��ʼʱ��, "YYYY-MM-DD") _
            & "      ԤԼʱ�䣺" & Format(rsTemp!ԤԼ��ʼʱ��, "hh:mm:ss") & " - " & Format(rsTemp!ԤԼ����ʱ��, "hh:mm:ss")
    Else
        lblSchedInfo = "��"
    End If

    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadSchDate(strSelDate As String)
'-----------------------------------------------------------
'����:����ԤԼ����
'���:  strSelDate -- ��Ҫѡ�е����ڣ����Ϊ�գ���ѡ��mstrԤԼ��ʼʱ��
'����:
'-----------------------------------------------------------
    Dim i As Integer
    Dim dtԤԼ���� As Date
    Dim lng1���ܼ� As Long
    Dim lng������ As Long
    Dim lng�ܼ� As Long
    Dim lngRow As Long
    Dim dt���� As Date
    Dim dtѡ������ As Date
    
    On Error GoTo err
    
    If mblnCanSchedule = False Then Exit Sub
    
    '����ԤԼ������
    With vsfSchDate
        .Rows = 1
        .Cols = 1
        .Rows = 6
        .Cols = 7
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .ColWidthMin = 1280
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .SelectionMode = flexSelectionFree
        .AllowSelection = False

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        .TextMatrix(0, 0) = "��һ"
        .TextMatrix(0, 1) = "�ܶ�"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "����"
        .TextMatrix(0, 6) = "����"
        
        '��ȡ��ǰ�µĵ�һ�������ڼ�
        dtԤԼ���� = Format(CDate(mstrԤԼ��ʼʱ��), "YYYY-MM-DD")
        dt���� = dtԤԼ���� - Day(dtԤԼ����) + 1
        If strSelDate <> "" And Format(dtѡ������, "YYYY-MM-DD") > Format(strSelDate, "YYYY-MM-DD") Then
            dtѡ������ = Format(strSelDate, "YYYY-MM-DD")
        Else
            dtѡ������ = dtԤԼ����
        End If
        
        lng1���ܼ� = Weekday(dt����, vbMonday)
        
        lng������ = Day(DateSerial(Year(dtԤԼ����), Month(dtԤԼ����) + 1, 0))
        
        lng�ܼ� = lng1���ܼ�
        lngRow = 1
        For i = 1 To lng������
            .TextMatrix(lngRow, lng�ܼ� - 1) = i
            '�������>=ԤԼ���ڣ�����ԤԼ��������������ʾ�Ѿ�ԤԼ��������������
            If DateCanSch(dt����, CLng(mstrԤԼ�豸ID)) = True Then
                .Cell(flexcpBackColor, lngRow, lng�ܼ� - 1) = &HFFFFC0  ' &HC0FFC0
            Else
                .Cell(flexcpBackColor, lngRow, lng�ܼ� - 1) = vbBlack
            End If
            
            If dt���� = dtѡ������ And .Cell(flexcpBackColor, lngRow, lng�ܼ� - 1) <> vbBlack Then
                .Select lngRow, lng�ܼ� - 1
            End If
            
            dt���� = dt���� + 1
            lng�ܼ� = lng�ܼ� + 1
            If lng�ܼ� > 7 Then
                lngRow = lngRow + 1
                lng�ܼ� = 1
            End If
            If lngRow = 6 Then
                .Rows = 7
            End If
        Next i
        
        .Rows = lngRow + 1  '�����������
        .RowHeightMin = IIf(lngRow = 5, 450, 382)
        'ѡ�� dtԤԼ����
        .Refresh

    End With
    
    '����������ʾ
    lblDate.Caption = Format(dtԤԼ����, "YYYY��MM��")
    Call loadSchSegment
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub loadSchSegment()
'-----------------------------------------------------------
'����:����ԤԼ����
'���:  rsTimes -- ԤԼʱ������ݼ�
'����:
'-----------------------------------------------------------
    Dim lngRow As Long
    Dim dtBeginTime As Date
    Dim dtEndTime As Date
    
    On Error GoTo err

    If mrsTimes.EOF = True Then
        vsfSchSegment.Rows = 1
        Exit Sub
    End If
    
    lngRow = 1
    '����ԤԼʱ���
    With vsfSchSegment
        .Rows = mrsTimes.RecordCount
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .ColWidthMin = 500
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(col_SchTimeSeg_�豸) = 1500
        .ColWidth(col_SchTimeSeg_��ʼʱ��) = 2000
        .ColWidth(col_SchTimeSeg_����ʱ��) = 2000
        .TextMatrix(0, col_SchTimeSeg_���) = "���"
        .TextMatrix(0, col_SchTimeSeg_�豸) = "ԤԼ�豸"
        .TextMatrix(0, col_SchTimeSeg_��ʼʱ��) = "��ʼʱ���"
        .TextMatrix(0, col_SchTimeSeg_����ʱ��) = "����ʱ���"
        While mrsTimes.EOF = False
            If mrsTimes!node_name = "SEGBEGINTIME" Then
                dtBeginTime = Format(mrsTimes!node_value, "HH:MM:SS")
                mrsTimes.MoveNext
                dtEndTime = Format(mrsTimes!node_value, "HH:MM:SS")
                If dtBeginTime >= Format(mstrԤԼ��ʼʱ��, "HH:MM:SS") Then
                    .TextMatrix(lngRow, col_SchTimeSeg_���) = lngRow
                    .TextMatrix(lngRow, col_SchTimeSeg_�豸) = mstrԤԼ�豸����
                    .TextMatrix(lngRow, col_SchTimeSeg_��ʼʱ��) = dtBeginTime
                    .TextMatrix(lngRow, col_SchTimeSeg_����ʱ��) = dtEndTime
                    If dtBeginTime = Format(mstrԤԼ��ʼʱ��, "HH:MM:SS") Then
                        .Cell(flexcpChecked, lngRow, col_SchTimeSeg_�豸) = 1
                    Else
                        .Cell(flexcpChecked, lngRow, col_SchTimeSeg_�豸) = 2
                    End If
                    lngRow = lngRow + 1
                End If
            End If
            mrsTimes.MoveNext
        Wend
        .Rows = lngRow
    End With
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ChangeSchDate(dtDate As Date, Optional lngSchDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'����:�ı�ԤԼ����
'���:  dtDate -- �����ԤԼ������
'       lngSchDeviceID -- ԤԼ�豸ID�������ָ���豸����0
'����: True -- �ɹ��� False -- ʧ��
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsTimes As ADODB.Recordset
    Dim strInXML As String
    Dim strOutXML As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strError As String
    
    Dim strOrderIDOld As String
    Dim str������ĿIDOld As String
    Dim str������Ŀ����Old As String
    Dim strҽ������Old As String
    Dim strԤԼ�豸����Old As String
    Dim strԤԼ�豸IDOld As String
    Dim strԤԼ��ʼʱ��Old As String
    Dim strԤԼ����ʱ��Old As String
    Dim strԤԼ����Old As String
    
    On Error GoTo err
    
    
    Set rsTimes = Nothing
        
    If GetSchDate(dtDate, strOutXML, lngSchDeviceID) = False Then
        mblnCanSchedule = False
        Exit Function
    End If
    
    '����XML��
'  --  <OUTPUT>
'  --    <ERROR>
'  --      <MSG>������Ϣ</MSG>
'  --    </ERROR>
'  --    <SCHINFO>
'  --      <ADVICEID>ҽ��ID</ADVICEID>
'  --      <CHECKID>������ĿID</CHECKID>
'  --      <CHECKNAME>������Ŀ����</CHECKNAME>
'  --      <ADVICEDOC>ҽ������</ADVICEDOC>
'  --      <MACHINENAME>ԤԼ�豸����</MACHINENAME>
'  --      <MACHINEID>ԤԼ�豸ID</MACHINEID>
'  --      <SCHBEGINTIME>��ʼʱ���</SCHBEGINTIME>
'  --      <SCHENDTIME>����ʱ���</SCHENDTIME>
'  --    </SCHINFO>
'  --    <SCHTIMES>
'  --      <SEGBEGINTIME>��ʼʱ���1</SEGBEGINTIME>
'  --      <SEGENDTIME>����ʱ���1</SEGENDTIME>
'  --    </SCHTIMES>
'  --    <SCHTIMES>
'  --      <SEGBEGINTIME>��ʼʱ���2</SEGBEGINTIME>
'  --      <SEGENDTIME>����ʱ���2</SEGENDTIME>
'  --    </SCHTIMES>
'  --  </OUTPUT>
    Set objXml = CreateObject("zl9ComLib.clsXML")
    Call objXml.OpenXMLDocument(strOutXML)
    
    Call objXml.GetSingleNodeValue("ADVICEID", mstrOrderID)
    Call objXml.GetSingleNodeValue("CHECKID", mstr������ĿID)
    Call objXml.GetSingleNodeValue("CHECKNAME", mstr������Ŀ����)
    Call objXml.GetSingleNodeValue("ADVICEDOC", mstrҽ������)
    Call objXml.GetSingleNodeValue("MACHINENAME", mstrԤԼ�豸����)
    Call objXml.GetSingleNodeValue("MACHINEID", mstrԤԼ�豸ID)
    Call objXml.GetSingleNodeValue("SCHBEGINTIME", mstrԤԼ��ʼʱ��)
    Call objXml.GetSingleNodeValue("SCHENDTIME", mstrԤԼ����ʱ��)
    
    Call objXml.GetMultiNodeRecord("OUTPUT/SCHTIMES", mrsTimes)
        
    mstrԤԼ���� = Format(mstrԤԼ��ʼʱ��, "YYYY-MM-DD")
    
    '��ѯԤԼʱ�������ȡ������Ϣ
    If mstrOrderID = "" Then
        Call objXml.GetSingleNodeValue("MSG", strError)
        If strError <> "�޿��õ�ԤԼ�豸��" Then
            MsgBox "��ȡԤԼ�������ִ���" & strError, vbOKOnly, "���ԤԼ"
        Else
            MsgBox "�޿��õ�ԤԼ�豸��", vbOKOnly, "���ԤԼ"
        End If
        mblnCanSchedule = False
        ChangeSchDate = False
        mstrOrderID = strOrderIDOld
        mstr������ĿID = str������ĿIDOld
        mstr������Ŀ���� = str������Ŀ����Old
        mstrҽ������ = strҽ������Old
        mstrԤԼ�豸���� = strԤԼ�豸����Old
        mstrԤԼ�豸ID = strԤԼ�豸IDOld
        mstrԤԼ��ʼʱ�� = strԤԼ��ʼʱ��Old
        mstrԤԼ����ʱ�� = strԤԼ����ʱ��Old
        mstrԤԼ���� = strԤԼ����Old
    Else
        mblnCanSchedule = True
        ChangeSchDate = True
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub loadSchInfo()
'-----------------------------------------------------------
'����:����ԤԼ��Ϣ
'���:
'����:
'-----------------------------------------------------------

    On Error GoTo err
    mblnRefreshDevice = False
    '��дԤԼ��Ϣ
    If mblnCanSchedule = True Then
        lblSchDevice.Caption = mstrԤԼ�豸����
        lblSchDate.Caption = Format(mstrԤԼ��ʼʱ��, "YYYY-MM-DD")
        lblSchTime.Caption = Format(mstrԤԼ��ʼʱ��, "hh:mm:ss") & " -- " & Format(mstrԤԼ����ʱ��, "hh:mm:ss")
    Else
        lblSchDevice.Caption = ""
        lblSchDate.Caption = ""
        lblSchTime.Caption = ""
    End If
    mblnRefreshDevice = True
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchDate_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Dim strDate As String
    On Error GoTo err
    '������첻��ԤԼ�����ֹ����
    strDate = vsfSchDate.TextMatrix(NewRowSel, NewColSel)
    If strDate = "" Then
        Cancel = True
    ElseIf vsfSchDate.Cell(flexcpBackColor, NewRowSel, NewColSel) = vbBlack Then
            Cancel = True
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDate As String
On Error GoTo err
    With vsfSchDate
        If .RowSel >= 1 Then
            strDate = .TextMatrix(.RowSel, .ColSel)
            If strDate <> "" Then
                Call ChangeSchDate(CDate(Format(mstrԤԼ��ʼʱ��, "YYYY-MM") & "-" & Format(strDate, "00")), CLng(mstrԤԼ�豸ID))
                Call loadSchSegment
                Call loadSchInfo
            End If
        End If
    End With
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchSegment_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    On Error GoTo err
    With vsfSchSegment
        If Row >= 1 And Col = col_SchTimeSeg_�豸 Then
            If .Cell(flexcpChecked, Row, col_SchTimeSeg_�豸) = 1 Then
                'ȡ�������е�ѡ��
                For i = 1 To .Rows - 1
                    If i <> Row Then
                        .Cell(flexcpChecked, i, col_SchTimeSeg_�豸) = 2
                    End If
                Next i
                '����ԤԼʱ�����ʾ
                mstrԤԼ��ʼʱ�� = mstrԤԼ���� & " " & .TextMatrix(Row, col_SchTimeSeg_��ʼʱ��)
                mstrԤԼ����ʱ�� = mstrԤԼ���� & " " & .TextMatrix(Row, col_SchTimeSeg_����ʱ��)
                Call loadSchInfo
            ElseIf .Cell(flexcpChecked, Row, col_SchTimeSeg_�豸) = 2 Then
                .Cell(flexcpChecked, Row, col_SchTimeSeg_�豸) = 1
            End If
            .Refresh
        End If
    End With
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Function GetOtherMonth(ByVal blnNextMonth As Boolean, ByRef dtDate As Date) As Boolean
'-----------------------------------------------------------
'����:�л����»����£��л������µ�һ�죬�������µ�һ��
'���:  blnNextMonth -- �Ƿ����£�True - ���£�False - ����
'       dtDate --����Ρ����Ρ� ��Ҫ�л�������
'����: True -- �ɹ��� False -- ʧ��
'-----------------------------------------------------------
    Dim dtNewDate As Date
    
    On Error GoTo err
    
    If blnNextMonth = True Then '��һ�µĵ�һ��
        dtNewDate = DateAdd("m", 1, dtDate)
        dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-01")
    Else    '��һ�µ����һ��
        dtNewDate = DateAdd("m", -1, dtDate)
        
        dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-01")
        '���һ�� dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-" & Day(DateSerial(Year(dtNewDate), Month(dtNewDate) + 1, 0)))
    End If
        
    dtDate = dtNewDate
    GetOtherMonth = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function refreshDate(dtDate As Date, Optional lngSchDeviceID As Long = 0, Optional strSelDate As String = "") As Boolean
'-----------------------------------------------------------
'����:ˢ������
'���:  dtDate --��Ҫ�л�������
'       lngSchDeviceID -- ԤԼ�豸ID�������ָ���豸����0
'       strSelDate -- ��Ҫѡ�е����ڣ����Ϊ�գ���ѡ��dtDate
'����: True -- �ɹ��� False -- ʧ��
'-----------------------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
    
    blnResult = ChangeSchDate(dtDate, lngSchDeviceID)
    
    If blnResult = False Then
        refreshDate = False
        Exit Function
    End If
    
    Call loadSchDate(strSelDate)
    Call loadSchInfo
    refreshDate = True
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSchDate(ByVal dtDate As Date, ByRef strOutXML As String, Optional lngShcDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'����:�����������ڣ�ȷ�������ԤԼ����
'���:  dtDate --��Ҫ�л�������
'       strOutXML -- ��ѯ����ֵ
'       lngShcDeviceID -- ԤԼ�豸ID�������ָ��ԤԼ�豸����0
'����: True -- �ɹ���False -- ʧ��
'-----------------------------------------------------------
    Dim strInXML As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '��ѯ����Ŀ�ԤԼ�豸��ʱ��
    If lngShcDeviceID = 0 Then
        strInXML = "<IN><ADVICEID>" & mlngAdviceID & "</ADVICEID>" & _
            "<BEGINTIME>" & dtDate & "</BEGINTIME></IN>"
    Else
        strInXML = "<IN><ADVICEID>" & mlngAdviceID & "</ADVICEID>" & _
            "<BEGINTIME>" & dtDate & "</BEGINTIME><MACHINEID>" & _
            lngShcDeviceID & "</MACHINEID></IN>"
    End If
    
'    strSQL = "select zl_Test_GetSchTimes(xmltype('" & strInXML & "')) as outData from dual"
'    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ԤԼ���ں��豸")
    
    strOutXML = gobjComLib.zlDatabase.CallProcedure("zl_Ӱ��ԤԼ_GetScheduleTimes", Me.Caption, strInXML, Empty)
    
    If strOutXML = "" Then
        GetSchDate = False
        Exit Function
    End If
    
    GetSchDate = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Private Function DateCanSch(dtDate As Date, Optional lngDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'����:�жϵ����Ƿ����ԤԼ
'���:  dtDate -- ����
'       lngSchDeviceID -- ԤԼ�豸ID�������ָ���豸����0
'����: True -- ��ԤԼ��False -- ����ԤԼ
'-----------------------------------------------------------
    Dim strOutXML As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strSchDate As String
    
    On Error GoTo err
    
    DateCanSch = False
    
    '���С�ڽ��죬ֱ�ӷ���False����ԤԼ
    If Format(Now, "YYYY-MM-DD") > Format(dtDate, "YYYY-MM-DD") Then
        Exit Function
    End If
    
    If GetSchDate(dtDate, strOutXML, lngDeviceID) = True Then
        Set objXml = CreateObject("zl9ComLib.clsXML")
        Call objXml.OpenXMLDocument(strOutXML)
        Call objXml.GetSingleNodeValue("SCHBEGINTIME", strSchDate)
        
        If Format(strSchDate, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
            DateCanSch = True
        End If
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfSchSegment_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_SchTimeSeg_�豸 Then
        Cancel = True
    End If
End Sub
