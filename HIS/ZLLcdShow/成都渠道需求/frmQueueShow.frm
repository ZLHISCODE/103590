VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQueueShow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer timerLCD 
      Interval        =   2000
      Left            =   7200
      Top             =   1920
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgQueuingData 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
      _cx             =   20770
      _cy             =   13150
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483627
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmQueueShow.frx":0000
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
   Begin VB.Label labInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5835
      TabIndex        =   2
      Top             =   8880
      Width           =   105
   End
   Begin VB.Label labTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5850
      TabIndex        =   1
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "frmQueueShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngQueuingItemsCount As Long    '�ŶӶ�����ʾ������
Private mlngQueuingDocItemsCount As Long    '�ŶӶ���ÿ��ҽ������ʾ������
Private mlngLedLoopQueryTime As Long '��ѯ���ʱ�䳤��
Private mlngQueuingingColor As Long     '�Ŷ�����ɫ
Private mlngCallingColor As Long        '��������ɫ
Private mlngEmergColor As Long      '����������ʾɫ
Private mstrGreeting As String
Private mlngTitleColor As Long
Private mlngVisitColor As Long
Private mlngCalledColor As Long

Private mstr��������() As String
Private mint��Ч���� As Integer
Private mstr�������� As String, mstrҽ������ As String, mstrExcludeData As String

Private mintViewDataType As Integer '������ʾ����
Private mblnComeBackFirst As Boolean    '���ﲡ���Ƿ������Ŷ�
Private mstrDelString As String '��Ҫɾ�����ַ���ʹ�á�,���ŷָ�
Private mlngRoomsCount As Long      'ÿҳ��ʾ�Ŀ�������
Private mlngCurPageIndex As Long   '��ǰ��ʾ��ҳ����
Private mlngPageSwitchTime As Long  'ҳ���л�ʱ�䳤��
Private mblnStartSwitch As Boolean  '�Ƿ�ʼ�л�ҳ��

Private mlngQueueId() As Long      '�������ҵĵ�ǰ��ʾ����ID
Private mstrQueueKey() As String
Private mlngBackColor As Long      '������ɫ����

'�ж������Ƿ�Ϊ��
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Function zlShowMe(cnOracle As ADODB.Connection, str��������() As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", _
    Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0, _
    Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ�Ŷ����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
    '         strCur��������-��ǰ��������
    '         lngCurҵ��ID-ҵ��ID
    '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�Ŷӵ�ָ��ҵ��ID
    '         intViewDataType������ʾ����(��ҽ��վ��"���ﷶΧ"������)��0��ʾ��ǰ�����µ��������ݣ�
    '                                      1��ʾ����Ϊ��ǰ���ң�����ҽ���������ڵ�ǰҽ������������Ϊ�պ�ҽ��Ϊ�յ�����
    '                                      2��ʾ����Ϊ��ǰ���ң���ҽ���������ڵ�ǰҽ��������
    '                                      3��ʾ��ǰҽ��������
    '         blnComebackFirst���ﲡ���Ƿ������Ŷ�
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mstr�������� = str��������
    mstr�������� = str����
    mstrҽ������ = strҽ��
    mstrExcludeData = strExcludeData
    
    '��intViewDataType=1���������Ϊ0
    If intViewDataType = 1 Then
        mintViewDataType = 0
    Else
        mintViewDataType = intViewDataType
    End If
    mblnComeBackFirst = blnComeBackFirst
    Call GetDepartKey(mstr��������, mstrQueueKey)
    Me.Show
End Function

Public Function zlSetPara(str��������() As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", _
    Optional strExcludeData As String, Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ�Ŷ����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
     '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�Ŷӵ�ָ��ҵ��ID
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    mstr�������� = str��������
    mstr�������� = str����
    mstrҽ������ = strҽ��
    mstrExcludeData = strExcludeData
    mblnComeBackFirst = blnComeBackFirst
    Call GetDepartKey(mstr��������, mstrQueueKey)
    
End Function

Private Sub MultiRoomsDisplay()
'**************************************************************************
'��ʾ������ҵĽк���Ϣ
'**************************************************************************
    Dim i As Integer, j As Integer
    Dim intCurPageRoomIndex As Integer '���浱ǰ��Ļҳ�Ŀ�������
    Dim blnSwitchPage As Boolean
    Dim blnAllowRoll As Boolean    '�жϵ�ǰ����ʾ�������Ƿ���Ҫ�����������������С�ڵ�ǰ�ܹ���ʾ�ļ�¼�������򲻹���
    Dim strCurRoomKey As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsSource As ADODB.Recordset   '��������
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String, strCurExcludeData() As String
    Dim strDoc As String, intPati As Integer
    
    err = 0: On Error GoTo errHandle
    If SafeArrayGetDim(mstr��������) > 0 Then
        j = 0
        
        strFilter = ""
        strValue = ""
        strUninTable = ""
        
        For i = 1 To UBound(mstr��������)
            If j > 10 Then
                strFilter = strFilter & " Or A.�������� ='" & mstr��������(i) & "'"
            Else
                If gobjCommFun.ActualLen(strValue) > 2000 Then
                     strValues(j) = Mid(strValue, 2)
                     strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
                     strValue = "": j = j + 1
                End If
                strValue = strValue & "," & mstr��������(i)
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as �������� From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
        
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
       strUninTable = " Select  Column_Value as �������� From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) "
    End If
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
    
    '�����ݿ��в�ѯ�����Ŷ���� ��a.�Ŷӱ�� || to_char(a.�ŶӺ���,'FM0000') As ���� ȥ���Ŷӱ��;������ǿ��ȥ�������ִ�п���
    '0:�Ŷ��У�1:�����У�2��������(����)��3����ͣ��4����ɾ��6�����7���Ѻ���
    '������(20150715):üɽ������ҽԺҪ���������ھ������ʾ
    strSQL = "Select /*+ Rule*/  to_Number(a.ID) as ID, a.��������, b.���� as ��������, to_char(a.�ŶӺ���,'FM0000') As ����,to_number(a.�ŶӺ���) �ŶӺ���," & _
             "a.��������, a.ҽ������, a.����,To_Char(m.����ʱ��,'HH24:MI') As ����ʱ��,m.����,m.ԤԼ,m.NO,m.�Ա�,m.����,r.רҵ����ְ��, " & _
             "decode(a.�Ŷ�״̬,0,'������',1,'������',2,'������',3,'��ͣ',4,'��ɾ���',6,'����',7,'�Ѻ���') as �Ŷ�״̬, to_Number(a.����) as ����, a.�Ŷ�ʱ��, a.����ʱ��, to_Number(�������) as �������, to_Number(a.ҵ������) as ҵ������, a.ҵ��ID, " & _
              IIf(mblnComeBackFirst, "to_Number(Nvl(a.�������, 9999999999)) as ���������", "0 as ���������") & _
             " From �ŶӽкŶ��� a, ���ű� b ,���˹Һż�¼ m,��Ա�� r, (" & strUninTable & ") E " & _
                IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
             " Where a.����id = b.Id And a.ҵ��ID=m.ID And a. ��������=E.��������  And (a.�Ŷ�״̬ in (0, 1, 7) Or (a.�Ŷ�״̬=2 and m.ִ��״̬<>0)) and nvl(m.��¼��־,0)=0 and �Ŷ�ʱ�� <= trunc(sysdate + 1) - 1/24/60/60 And ҵ������=0" & _
                IIf(mintViewDataType = 1, " and  ((a.����=C.Column_Value and a.ҽ������ is null) or a.ҽ������=D.Column_Value or (a.���� is null and a.ҽ������ is null))", "") & _
                IIf(mintViewDataType = 2, " and ((a.����=C.Column_Value or a.ҽ������=D.Column_Value) ", "") & _
                IIf(mintViewDataType = 3, " and a.ҽ������=D.Column_Value", "") & " and r.����=a.ҽ������" & _
             " Order By ҽ������,�Ŷ�״̬ desc,a.���� desc, ���������, to_number(a.�ŶӺ���) "

    Set rsSource = gobjDatabase.OpenSQLRecord(strSQL, "��ʾ�Ŷ����", mstr��������, mstrҽ������, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    Set rsTemp = gobjDatabase.CopyNewRec(rsSource)
    
    If rsTemp.EOF Then
        vfgQueuingData.Cell(flexcpText, 2, 0, vfgQueuingData.Rows - 1, vfgQueuingData.Cols - 1) = ""
        Exit Sub
    End If
     
    While Not rsTemp.EOF
        If InStr(1, mstrExcludeData, rsTemp!ҵ������ & ":" & rsTemp!ҵ��ID) > 0 Then
            rsTemp.Delete
        End If
        rsTemp.MoveNext
    Wend
    
    'ע���������,��������ʵʱ��ʾ��ʱ�򣬸���ʵ�ʵ����ֳ��ȵ����еĿ��
'    Call vfgQueuingData.AutoSize(0, 2)

    '��ʾ��ǰ���Һ��������
    If SafeArrayGetDim(mstrQueueKey) <= 0 Then
        Exit Sub
    End If
    
    '���mstrQueueKey�Ŀ���ȫ��Ϊ�գ���ִ��
    For i = 1 To UBound(mstrQueueKey)
        If Trim(mstrQueueKey(i)) <> "" Then
            GoTo Start
        End If
    Next i
    
    Exit Sub
    
Start:
    
    '��ҳ�潫����ȫ����ʾ���ټ��������ʾ��ʱ��
    If mblnStartSwitch Then
        If mlngPageSwitchTime > 0 Then
            mlngPageSwitchTime = mlngPageSwitchTime - 1
            Exit Sub
        Else
            mblnStartSwitch = False
        End If
    End If
    
    blnSwitchPage = True

    '���ú�����б��ϲ���ʽ,����ͷ�ϲ�
    vfgQueuingData.MergeCellsFixed = flexMergeFree
    vfgQueuingData.MergeCells = flexMergeNever
    
    '��������Ҫ��ʾ�Ŀ������ƣ�����ȡ��Ϣ��ʾ
    For i = (mlngCurPageIndex - 1) * mlngRoomsCount + 1 To mlngCurPageIndex * mlngRoomsCount
        
        'ȡ�õ�ǰҳ���Ӧ�Ŀ���������
        intCurPageRoomIndex = i - (mlngCurPageIndex - 1) * mlngRoomsCount
        
        '����LCD��ʾ�Ŀ�������
        If i <= UBound(mstrQueueKey) Then
            strCurRoomKey = mstrQueueKey(i)
            vfgQueuingData.Cell(flexcpText, 0, (intCurPageRoomIndex - 1) * 5 + 1) = strCurRoomKey
           '��ʾ����ʱ,����һ�кϲ���ʾ
            Call vfgMergeRowCol(0, (intCurPageRoomIndex - 1) * 5 + 1)
        Else
            strCurRoomKey = ""
            vfgQueuingData.Cell(flexcpText, 0, (intCurPageRoomIndex - 1) * 5 + 1) = ""
            '��ʾ����ʱ,����һ�кϲ���ʾ
            Call vfgMergeRowCol(0, (intCurPageRoomIndex - 1) * 5 + 1)
        End If
        
        If Trim(strCurRoomKey) <> "" Then
                        
            '���˳���ǰ��������Ҫ��ʾ�ĺ�������
            'rsTemp.Filter = "�Ŷ�״̬='�Ŷ���' and ��������='" & strCurRoomKey & "' and ID>" & mlngQueueId(intCurPageRoomIndex)
            'rsTemp.Filter = "��������='" & strCurRoomKey & "' and ID>" & mlngQueueId(intCurPageRoomIndex)
            'rsTemp.Sort = "ҽ������,�Ŷ�״̬,���� desc, ��������� asc, �Ŷ�ʱ�� asc, �ŶӺ��� asc"                          '�˴�������ʾ˳��
            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
            
            '�����¼����С�ڵ�ǰ�ܹ���ʾ�ļ�¼�������������������blnAllowRoll��ֱ������Ϊtrue
            blnAllowRoll = IIf(rsTemp.RecordCount <= mlngQueuingItemsCount, False, True)  'blnAllowRoll = True
        
            If Not rsTemp.EOF Then
                '��ʵ�ʼ�¼��С��ÿҳ�ܹ���ʾ�ļ�¼ʱ���򲻽��й���
                If blnAllowRoll Then
                    mlngQueueId(intCurPageRoomIndex) = rsTemp!ID
                    
                    '���û�н�����������ҳ
                    blnSwitchPage = False
                End If
                
                j = 0
                '��ʾָ�������ļ�¼�������¼���ݲ�����ʾ������ʾ������
                While j < mlngQueuingItemsCount
                    If Not rsTemp.EOF Then
                        If strDoc <> rsTemp!ҽ������ Then       '����һ����ʾҽ��
                            Call SetRoomsData(1, j + 2, intCurPageRoomIndex, "", "", "", "", ""): j = j + 1
                            Call SetRoomsData(1, j + 2, intCurPageRoomIndex, " ҽ����" & Nvl(rsTemp!ҽ������) & "  " & Nvl(rsTemp!רҵ����ְ��), " ҽ����" & Nvl(rsTemp!ҽ������) & "  " & Nvl(rsTemp!רҵ����ְ��), " ҽ����" & Nvl(rsTemp!ҽ������) & "  " & Nvl(rsTemp!רҵ����ְ��), " ҽ����" & Nvl(rsTemp!ҽ������) & "  " & Nvl(rsTemp!רҵ����ְ��), " ҽ����" & Nvl(rsTemp!ҽ������) & "  " & Nvl(rsTemp!רҵ����ְ��))
                            Call vfgMergeRowCol(j + 2, (intCurPageRoomIndex - 1) * 5 + 1)
                            j = j + 1: strDoc = rsTemp!ҽ������: intPati = 1
                        End If
                        If intPati - 1 < mlngQueuingDocItemsCount Then
                            Call SetRoomsData(2, j + 2, intCurPageRoomIndex, Nvl(rsTemp!�ŶӺ���), Nvl(rsTemp!��������), Nvl(rsTemp!�Ա�), Nvl(rsTemp!����), Nvl(rsTemp!�Ŷ�״̬) & " " & IIf(Nvl(rsTemp!����) = "1", "����", IIf(Nvl(rsTemp!ԤԼ) = "1", "ԤԼ", "")))
                            intPati = intPati + 1
                            If Nvl(rsTemp!����) = "1" Then
                                Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngEmergColor)
                            Else
                                Select Case Nvl(rsTemp!�Ŷ�״̬)
                                    Case "������"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngCallingColor)
                                    Case "������"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngQueuingingColor)
                                    Case "������"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngVisitColor)
                                    Case "�Ѻ���"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngCalledColor)
                                    Case Else
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngQueuingingColor)
                                End Select
                            End If
                            
                            j = j + 1
                        End If
                        rsTemp.MoveNext
                    Else
                        Call SetRoomsData(2, j + 2, intCurPageRoomIndex, "", "", "", "", "")
                        j = j + 1
                    End If
                    DoEvents
                Wend
            Else
                For j = 0 To mlngQueuingItemsCount - 2
                    Call SetRoomsData(2, j + 3, intCurPageRoomIndex, "", "", "", "", "")
                    DoEvents
                Next j
            End If
        Else
            '���û�ж�Ӧ�Ŀ��ҿ���ʾ����ʹ����Ϊ��
            For j = 0 To mlngQueuingItemsCount - 1
                Call SetRoomsData(2, j + 3, intCurPageRoomIndex, "", "", "", "", "")
                DoEvents
            Next j
        End If
        DoEvents
    Next i
           
    '���blnSwitchPageΪ�棬�������һ��ҳ�����ʾ
    If blnSwitchPage Then
        mlngCurPageIndex = mlngCurPageIndex + 1
        
        Dim intPageCount As Integer
        '����ҳ������
        If UBound(mstrQueueKey) Mod mlngRoomsCount <> 0 Then
            intPageCount = Int(UBound(mstrQueueKey) / mlngRoomsCount) + 1
        Else
            intPageCount = UBound(mstrQueueKey) / mlngRoomsCount
        End If
        
        '�ж��Ƿ��Ѿ���ʾ������ҳ������ǣ���������ʾ��һҳ
        If mlngCurPageIndex > intPageCount Then
            mlngCurPageIndex = 1
        End If
        
        For i = 1 To UBound(mlngQueueId)
            mlngQueueId(i) = -1
        Next i
        
        mblnStartSwitch = True
        mlngPageSwitchTime = 3  '��������ҳ�������ʾ��ʱ�䳤��
    End If
    
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    
    Exit Sub
errHandle:
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strReg As String
    
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    
    mlngRoomsCount = CLng(GetSetting("ZLSOFT", strReg, "ҳ����ʾ��", "1"))
    
    mlngLedLoopQueryTime = Val(GetSetting("ZLSOFT", strReg, "LED��ѯʱ��", "2"))
    timerLCD.Interval = mlngLedLoopQueryTime * 1000
    mlngQueuingItemsCount = Val(GetSetting("ZLSOFT", strReg, "�ŶӼ�¼��ʾ��", "6"))
    mlngQueuingDocItemsCount = Val(GetSetting("ZLSOFT", strReg, "�ŶӼ�¼��ʾ��", "6"))
    mlngQueuingingColor = GetSetting("ZLSOFT", strReg, "�Ŷ�����ɫ", vbGreen)
    mlngCallingColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    mlngVisitColor = GetSetting("ZLSOFT", strReg, "��������ɫ", vbGreen)
    mlngCalledColor = GetSetting("ZLSOFT", strReg, "�Ѻ�����ɫ", vbGreen)
     
    mlngEmergColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbRed)
    mstrGreeting = GetSetting("ZLSOFT", strReg, "ף����", "ף�����տ�����")
    mlngTitleColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbRed)
    
    mlngCurPageIndex = 1
    mlngPageSwitchTime = 3 '10��
    mblnStartSwitch = False
    
    Me.BackColor = vbBlack
    
    '���������С
    ReDim mlngQueueId(mlngRoomsCount)
    ReDim mstrQueueKey(mlngRoomsCount)
    
    For i = 1 To mlngRoomsCount
      mlngQueueId(i) = -1
      mstrQueueKey(i) = ""
    Next i
    
    Call GetDepartKey(mstr��������, mstrQueueKey)
     
    '������ʾ����
    Call SetFaceFont
    '������ʾλ��
    Call SetFacePostion
    '���ñ�����ɫ
    Call SetBackColor
    
    Call InitFace(mlngRoomsCount, mlngQueuingItemsCount)
End Sub

Private Sub Form_Resize()
    Call InitFace(mlngRoomsCount, mlngQueuingItemsCount)
End Sub

Public Sub SetFaceFont()
'************************************************************************************
'���ý�����ʾ��������ʽ
'************************************************************************************
    Dim strReg As String
    Dim curFontSize As Currency
    
    On Error Resume Next
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\�Ŷӽк�\Һ������"

    vfgQueuingData.Font.Name = GetSetting("ZLSOFT", strReg, "����", "����")
    vfgQueuingData.Font.Bold = GetSetting("ZLSOFT", strReg, "����", "False")
    vfgQueuingData.Font.Italic = GetSetting("ZLSOFT", strReg, "б��", "False")
    
    curFontSize = GetSetting("ZLSOFT", strReg, "�ֺ�", "14")
    vfgQueuingData.Font.Size = curFontSize * 4.5 / 5
    
    
    labTitle.Font.Name = vfgQueuingData.Font.Name
    labTitle.Font.Bold = vfgQueuingData.Font.Bold
    labTitle.Font.Italic = vfgQueuingData.Font.Italic
    labTitle.Font.Size = curFontSize + 1
    labTitle.Caption = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "") & "���ﲡ�˺���һ����"
    labTitle.ForeColor = mlngTitleColor
    
    labInfo.Font.Name = vfgQueuingData.Font.Name
    labInfo.Font.Bold = vfgQueuingData.Font.Bold
    labInfo.Font.Italic = vfgQueuingData.Font.Italic
    labInfo.Font.Size = curFontSize + 0.5
    labInfo.ForeColor = mlngTitleColor
End Sub

Public Sub SetFacePostion()
'************************************************************************************
'
'���ý������ʾλ��
'
'
'************************************************************************************
    Dim strReg As String
    
    On Error Resume Next
        
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    
    '������ʾ����
    Me.Left = GetSetting("ZLSOFT", strReg, "�Ŷ���Ļ��", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "�Ŷ���Ļ��", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "���", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "�߶�", "768") * Screen.TwipsPerPixelY
    
End Sub

Private Sub InitFace(ByVal lngRoomsCount As Long, ByVal lngWaitItemsCount As Long)
'************************************************************************************
'��ʼ��������ʾ
'lngRoomsCount: ÿ��Ļ�ܹ���ʾ�Ŀ�������
'lngWaitItemsCount�����������ʾ��������
'************************************************************************************
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '�����Ŷ���ʾ-------------------------------------
    labTitle.Top = Round(Me.ScaleHeight * 0.01)
    labTitle.Left = 0
    'labTitle.Height = Round(Me.ScaleHeight * 0.042)
    labTitle.Width = Me.ScaleWidth
    
    'vfgQueuingData.Top = Round(Me.ScaleHeight * 0.052)
    vfgQueuingData.Top = labTitle.Height + 100
    vfgQueuingData.Left = 100
    vfgQueuingData.Width = Me.ScaleWidth - 300
    'vfgQueuingData.Height = Round(Me.ScaleHeight * 0.907)
    vfgQueuingData.Height = Me.ScaleHeight - labTitle.Height - labInfo.Height - 200

    'labInfo.Top = Round(Me.ScaleHeight * 0.961)
    labInfo.Top = labTitle.Height + vfgQueuingData.Height + 100
    labInfo.Left = 0
    labInfo.Width = Me.ScaleWidth
    'labInfo.Height = Round(Me.ScaleHeight * 0.039)
    
    vfgQueuingData.Cols = lngRoomsCount * 5
    vfgQueuingData.Rows = Int(vfgQueuingData.Height / vfgQueuingData.Cell(flexcpHeight, 0, 0))
    
    '�Զ����ÿ���ʾ������
    mlngQueuingItemsCount = vfgQueuingData.Rows - 2
    vfgQueuingData.ForeColor = mlngQueuingingColor
    vfgQueuingData.Enabled = False
    
    '���ú�����ʾ��ʽ--------------------------------------
    For i = 0 To lngRoomsCount - 1
        vfgQueuingData.ColWidth(i * 4 + 0 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.75)
        vfgQueuingData.ColWidth(i * 4 + 1 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 1.1)
        vfgQueuingData.ColWidth(i * 4 + 2 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.85)
        vfgQueuingData.ColWidth(i * 4 + 3 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.9)
        vfgQueuingData.ColWidth(i * 4 + 4 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 1.4)
        
        '���ÿ��ҵ���ʾ��ĿΪ���ж���
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 0 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 1 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 2 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 3 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 4 + i) = flexAlignCenterCenter

        '����������������ʾ��ĿΪ���ж���
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 0 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 1 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 2 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 3 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 4 + i) = flexAlignCenterCenter
        
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 0 + i) = "  ���  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 1 + i) = "  ����  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 2 + i) = "  �Ա�  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 3 + i) = "  ����  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 4 + i) = "  ����״̬  "
    Next i
    
    Call DrawBorder
        
    '��ʾ��Ϣ--------------------------------------
    If Split(mstrGreeting & "|", "|")(1) <> "" Then
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & " " & Format(Now, "yyyy-mm-dd hh:mm") & " ����" & GetTodayNum & " " & Split(mstrGreeting & "|", "|")(1)
    Else
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & " " & Format(Now, "yyyy-mm-dd hh:mm") & " ����" & GetTodayNum
    End If
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Function GetTodayNum()
    On Error Resume Next
    
    Select Case Weekday(Date, vbMonday)
        Case 1: GetTodayNum = "һ"
        Case 2: GetTodayNum = "��"
        Case 3: GetTodayNum = "��"
        Case 4: GetTodayNum = "��"
        Case 5: GetTodayNum = "��"
        Case 6: GetTodayNum = "��"
        Case 7: GetTodayNum = "��"
    End Select
End Function

Public Sub timerLCD_Timer()
        On Error GoTo errHandle
        Dim blnTimer As Boolean

        blnTimer = timerLCD.Enabled
        timerLCD.Enabled = False
        
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & "  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  ����" & GetTodayNum & "  " & Split(mstrGreeting & "|", "|")(1)
        
        Call MultiRoomsDisplay
        
        timerLCD.Enabled = blnTimer
    Exit Sub
errHandle:
    Call gobjComLib.SaveErrLog
    
    timerLCD.Enabled = blnTimer
End Sub

Private Function setTextWidth(strText As String, iLen As Integer, intWay As Integer) As String
'**************************************************************************
'
'�����ı����ȴﵽ�ƶ����ȣ�������㣬�򲹳�ո�
'
'strText����Ҫ���õ��ı���
'
'iLen���ı�����
'
'intWay�����뷽��
'
'**************************************************************************
    
    On Error GoTo errHandle
    
    If Len(strText) >= iLen Then
        setTextWidth = Mid(strText, 1, iLen)
        Exit Function
    End If
    
    Select Case intWay
      Case 1
        setTextWidth = Space(iLen - Len(strText)) & strText
      Case 2
        setTextWidth = strText & Space(iLen - Len(strText))
      Case 3
        setTextWidth = Space((iLen - Len(strText)) - Int((iLen - Len(strText)) / 2)) & strText & Space(Int((iLen - Len(strText)) / 2))
    End Select
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Private Sub GetDepartKey(str��������() As String, strDepartKey() As String)
'**************************************************************************
'
'ȡ���Ŷӽк�ϵͳ���漰���Ŀ�������
'
'str��������()���к�ϵͳ�еĶ�����������
'
'strDepartKey()�������������
'
'**************************************************************************

    Dim strSQL As String
    Dim i As Integer
    Dim rsDepart As ADODB.Recordset
    Dim strDepartId As String
    
    On Error GoTo errHandle
    
    If SafeArrayGetDim(str��������) <= 0 Then
        Exit Sub
    End If
    
    If UBound(str��������) <= 0 Then
        Exit Sub
    End If
    
    'ȡ����Ҫ�����Ŀ���ID
    strDepartId = ""
    For i = 1 To UBound(str��������)
        If Trim(str��������(i)) <> "" Then
            If Trim(strDepartId) <> "" Then strDepartId = strDepartId & ","
            strDepartId = strDepartId & Mid(str��������(i) & ":", 1, InStr(1, str��������(i) & ":", ":") - 1)
        End If
    Next i
    
    If Trim(strDepartId) = "" Then
      Exit Sub
    End If
    
    strSQL = "select /*+ Rule*/ distinct ����, id from ���ű� a, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) b where a.id =b.Column_Value order by Id"
    Set rsDepart = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ������Ϣ", strDepartId)

    
    If rsDepart.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ReDim strDepartKey(rsDepart.RecordCount)
    
    For i = 1 To rsDepart.RecordCount
        strDepartKey(i) = rsDepart!����
        rsDepart.MoveNext
    Next i
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub vfgMergeRowCol(ByVal lngrow As Long, ByVal lngcol As Long)
'2010-07-09 ZHQ ǿ�Ƶ�һ��ָ����ǰ�����н��кϲ�������ڽ���������Ԫ��������дΪ��ͬ����
'               ֱ���Ե�(0,lngcol)����Ϊ׼���кϲ�
    
    Dim strTemp As String
    strTemp = vfgQueuingData.TextMatrix(lngrow, lngcol)
    If strTemp = "" Then strTemp = " "
    
    With vfgQueuingData
        .TextMatrix(lngrow, lngcol - 1) = strTemp
        .TextMatrix(lngrow, lngcol) = strTemp
        .TextMatrix(lngrow, lngcol + 1) = strTemp
        .TextMatrix(lngrow, lngcol + 2) = strTemp
        .TextMatrix(lngrow, lngcol + 3) = strTemp
        .MergeRow(lngrow) = True
        .MergeCol(lngcol - 1) = True
        .MergeCol(lngcol) = True
        .MergeCol(lngcol + 1) = True
        .MergeCol(lngcol + 2) = True
        .MergeCol(lngcol + 3) = True
        .MergeCells = flexMergeRestrictRows
    End With
End Sub

Private Sub DrawBorder()
'**************************************************************************
'���Ʊ�ͷ�߿�
'**************************************************************************

    Dim i As Integer
    
    On Error GoTo errHandle
    
    For i = 0 To mlngRoomsCount - 1
        With vfgQueuingData
            .Select 0, i * 5, 0, i * 5 + 4
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    For i = 0 To vfgQueuingData.Cols - 1
        With vfgQueuingData
            .Select 1, i, 1, i
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    For i = 0 To mlngRoomsCount - 1
        With vfgQueuingData
            .Select 2, i * 5, .Rows - 1, i * 5 + 4
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub SetRoomsData(ByVal Align As Integer, ByVal intRowIndex As Integer, ByVal intRoomIndex As Integer, _
    ByVal strName As String, ByVal strSex As String, ByVal strAge, ByVal strState As String, ByVal strDocter As String)
'**************************************************************************
'������ʾ����
'intRowIndex����ǰ������ʾ��
'intRoomIndex����ǰ��ʾ������������1��ʼ
'strTime�� ʱ�䣬strName��������strSex��
'**************************************************************************
    On Error GoTo errHandle
        
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 1) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 2) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 3) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 4) = False
        
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5) = strName
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 1) = strSex
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 2) = strAge
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 3) = strState
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 4) = strDocter
        
        If Align = 1 Then                       '�����
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignLeftCenter
        ElseIf Align = 3 Then                   '�Ҷ���
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignRightCenter
        Else                                    '���ж���
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignCenterCenter
        End If
        
    Exit Sub
errHandle:
    err.Clear
End Sub

Private Sub SetRoomsColor(ByVal intRowIndex As Integer, ByVal intRoomIndex As Integer, fcolor As Long)
'**************************************************************************
'�����ض��е���ʾɫ
'**************************************************************************
    On Error GoTo errHandle
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 1) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 2) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 3) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 4) = fcolor
        
    Exit Sub
errHandle:
    err.Clear
End Sub

Public Sub SetBackColor()
'************************************************************************************
'���ý��汳��ɫ
'************************************************************************************
    Dim strReg As String
    Dim mlngBackColor As Long
    
    On Error Resume Next
    '��ע����У���ȡ��ʾ����
    strReg = "����ģ��\�Ŷӽк�\Һ������"
    mlngBackColor = GetSetting("ZLSOFT", strReg, "������ɫ", vbBlack)
    
    With vfgQueuingData
        .BackColor = mlngBackColor
        .BackColorAlternate = mlngBackColor
        .BackColorBkg = mlngBackColor
        .SheetBorder = mlngBackColor
    End With
    labTitle.BackColor = mlngBackColor
    labInfo.BackColor = mlngBackColor
End Sub

