VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRMP150 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "frmRMP150.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkFile 
      Alignment       =   1  'Right Justify
      Caption         =   "�����ļ���ʽ"
      Height          =   255
      Left            =   270
      TabIndex        =   6
      ToolTipText     =   "ѡ����Ŀ���뻹�ǰ������ļ�����"
      Top             =   3090
      Width           =   1380
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "����"
      Height          =   225
      Left            =   2490
      TabIndex        =   5
      ToolTipText     =   "�������ļ�����Ѿ���������򲻵��룻���ϣ�Ҫ����ѡ�ļ����룬�������ļ��Ƿ��Ѿ��������"
      Top             =   3405
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   705
      TabIndex        =   3
      Top             =   3360
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   111607811
      CurrentDate     =   39658
   End
   Begin MSScriptControlCtl.ScriptControl vbsCalce 
      Left            =   5295
      Top             =   3090
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   1
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   0
      Top             =   3270
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2940
      Left            =   30
      TabIndex        =   2
      Top             =   45
      Width           =   6300
      _cx             =   11112
      _cy             =   5186
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   3405
      Width           =   450
   End
End
Attribute VB_Name = "frmRMP150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFilePath As String                      '�ļ�·��
Private mdtStart As Date
Private mdtEnd As Date
Private mlngSampleNo As Long
Private mobjFile As New Scripting.FileSystemObject  '�ļ�����
Private mstrReturn() As String '�淵�ؽ��

Enum mCol
    ѡ�� = 0:    ��Ŀ���: Ӣ��: ��Ŀ����: ����ʱ��: ��ʼ�걾��: С��: CutOff��ʽ: ���Թ�ʽ: ������ʽ: ��Сֵ: �ļ���
End Enum
'---��дINI�ļ���API����
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2 'ע�ͣ����ƶ�����
Private Const SWP_NOSIZE = &H1 'ע�ͣ����ı䴰��ߴ�
Private Const HWND_TOPMOST = -1         'ע�ͣ�����������ǰ��
Private Const HWND_NOTOPMOST = -2       'ע�ͣ����岻����ǰ��


Public Function ShowMe(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    mstrFilePath = mobjFile.GetParentFolderName(strFile)
    If Len(mstrFilePath) = 0 Then mstrFilePath = App.Path
    mdtStart = dtStart
    mdtEnd = dtEnd
    mlngSampleNo = lngSampleNO
    ReDim mstrReturn(0) As String
    Me.Show vbModal
    ShowMe = mstrReturn
End Function

Private Sub ImpFile()
    Dim objStream As TextStream, strLine As String, str���� As String
    Dim strFileName As String, strDataName As String
    Dim strImpedFile As String, strImpFile As String
    Dim intCount As Integer
    
    Dim lngFileCount As Long, lngLoop As Long, strTmpFile As String, bln�ѵ��� As Boolean
    strFileName = mstrFilePath & "\RMPDAT.IDX"
    
    If chkFile = 0 Then
        '--------------------------------------------------------------------------------
        '---  �Զ�����
        '--------------------------------------------------------------------------------
        '�������ļ� ��ȡ�ѵ����ļ��б�
        
        If chkReplace.Value <> 0 Then
            '���Ƿ�ʽ���걾���声
            With vfgList
                For intCount = .FixedRows To .Rows - 1
                    If .TextMatrix(intCount, mCol.ѡ��) = 1 Then
                        WriteToIni App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), .TextMatrix(intCount, mCol.Ӣ��), 0
                    End If
                Next
            End With
        Else
            strLine = ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "Count")
            If Val(strLine) > 0 Then
                For intCount = 1 To Val(strLine)
                    strImpedFile = strImpedFile & "," & ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "File" & intCount)
                Next
            End If
        End If
        
        '�� RMPDAT.IDX�ļ�,�ҳ�Ҫ������ļ�,���浽�������ļ��б�
        Set objStream = mobjFile.OpenTextFile(strFileName, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If InStr(strLine, ".DAT=UNDEFINED,") > 0 And InStr(strLine, "RMP") > 0 Then
                str���� = Split(strLine, ",")(2)
                str���� = Right(str����, 2) & "/" & Mid(str����, 1, 5)
                strDataName = Replace(Split(strLine, ",")(0), "=UNDEFINED", "")
                
                If Format(CDate(str����), "yyyy-MM-dd") = Format(dtpDate.Value, "yyyy-MM-dd") _
                   And InStr(strImpedFile, "," & strDataName) <= 0 Then
                    
                    strImpFile = strImpFile & "," & strDataName & "|" & Format(CDate(str����), "yyyy-MM-dd")
                    
                End If
            End If
        Loop
    Else
        With vfgList
            For intCount = .FixedRows To .Rows - 1
                If .TextMatrix(intCount, mCol.ѡ��) = 1 Then
                    
                    If chkReplace.Value = 0 Then
                        '����Ƿ��ѵ����
                        bln�ѵ��� = False
                        lngFileCount = Val(ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "Count"))
                        If lngFileCount > 0 Then
                           For lngLoop = 1 To lngFileCount
                               strTmpFile = ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "File" & lngLoop)
                               If strTmpFile = .TextMatrix(intCount, mCol.�ļ���) Then
                                   bln�ѵ��� = True
                                   Exit For
                               End If
                           Next
                        End If
                        If Not bln�ѵ��� Then strImpFile = strImpFile & "," & .TextMatrix(intCount, mCol.�ļ���) & "|" & Format(dtpDate.Value, "yyyy-MM-dd")
                    Else
                        strImpFile = strImpFile & "," & .TextMatrix(intCount, mCol.�ļ���) & "|" & Format(dtpDate.Value, "yyyy-MM-dd")
                    End If
                End If
            Next
        End With
    End If
    '������
    If strImpFile <> "" Then Call ImpItem(strImpFile)
End Sub

Private Sub ImpItem(ByVal strImpFiles As String)
    '������
    Dim varItem As Variant, strFileName As String, intRow As Integer
    Dim objStream As TextStream
    
    Dim dblBC_Total As Double, dblBC As Double, intBC_Count As Integer                  '�հ׿�
    Dim dblNC_Total As Double, dblNC As Double, intNC_Count As Integer                  '���Կ�
    Dim dblPC_Total As Double, dblPC As Double, intPC_Count As Integer                  '���Կ�
    Dim iCount As Integer '�ܼ�¼��
    Dim iRow As Integer, strType As String, strXY As String, dblOD As Double
    Dim i��� As Integer
    Dim arrData(3, 1 To 8, 1 To 12) As String                '(0=���;1=ԭʼOD:2=OD;3=����)
    Dim lngResultCount As Long
    Dim strIni�걾�� As String, str���� As String, str��Ŀ As String
    '--- ��ȡҪ�������Ŀ�͹�ʽ
    Dim strҪ������Ŀ As String, intС�� As Integer, strCoutOff��ʽ As String, str���Թ�ʽ As String, str�����Թ�ʽ As String
    Dim dblCoutOffֵ As Double, bln���� As Boolean, bln������ As Boolean, int��ʼ�걾�� As String, strLine As String
    Dim strOD As String, dbl���Կ���Сֵ As Double, str���Խ�� As String
    Dim intFileCount As Integer, str��ʽ As String
    varItem = Split(Mid(strImpFiles, 2), ",")
    lngResultCount = -1
    For intRow = LBound(varItem) To UBound(varItem)
        strFileName = varItem(intRow)
        
        str���� = Split(strFileName, "|")(1)
        strFileName = Split(strFileName, "|")(0)
        i��� = 0
        '��ȡ����
        str��Ŀ = ReadFromIni(mstrFilePath & "\" & strFileName, "Test Log", "Test Class")
        str��Ŀ = Mid(str��Ŀ, InStr(str��Ŀ, " ") + 1)
        intС�� = 0: strCoutOff��ʽ = "": str���Թ�ʽ = "": int��ʼ�걾�� = 0
        intFileCount = Val(ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & str����, "Count"))
        With vfgList
            For iRow = .FixedRows To .Rows - 1
                If str��Ŀ = .TextMatrix(iRow, mCol.Ӣ��) And .TextMatrix(iRow, mCol.ѡ��) = 1 Then
                    intС�� = Val(.TextMatrix(iRow, mCol.С��))
                    strCoutOff��ʽ = .TextMatrix(iRow, mCol.CutOff��ʽ)
                    str���Թ�ʽ = .TextMatrix(iRow, mCol.���Թ�ʽ)
                    str�����Թ�ʽ = .TextMatrix(iRow, mCol.������ʽ)
                    int��ʼ�걾�� = Val(ReadFromIni(App.Path & "\RMP150.ini", str����, str��Ŀ))
                    dbl���Կ���Сֵ = Val(.TextMatrix(iRow, mCol.��Сֵ))
                    Exit For
                End If
            Next
        End With
        
        If strCoutOff��ʽ <> "" Then
            iCount = Val(ReadFromIni(mstrFilePath & "\" & strFileName, "Test Log", "Nof Results"))
            
            dblBC_Total = 0: intBC_Count = 0
            dblNC_Total = 0: intNC_Count = 0
            dblPC_Total = 0: intPC_Count = 0
            
            dblBC = 0: dblNC = 0: dblPC = 0
            If iCount > 0 Then
                For iRow = 1 To iCount
                    
                    strType = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "Liquid Type") '����
                    strXY = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "Position Name") '����
                    strOD = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "OD Values")    '�����
                    
                   
                    
                    If strOD <> "" Then
                        dblOD = Val(strOD)
                        
                        If strType = "blk" Then                             '�հ׿�
                            dblBC_Total = dblBC_Total + dblOD
                            intBC_Count = intBC_Count + 1
                        ElseIf strType = "nc" Then                          '���Կ�
                            dblNC_Total = dblNC_Total + dblOD
                            intNC_Count = intNC_Count + 1
                        ElseIf strType = "pc" Then                          '���Կ�
                            dblPC_Total = dblPC_Total + dblOD
                            intPC_Count = intPC_Count + 1
                        ElseIf strType = "smp" Then                         'ODֵ
                            
                            If dblBC = 0 Then dblBC = dblBC_Total / intBC_Count ' / 1077936128 'intBC_Count
                            If dblNC = 0 Then dblNC = (dblNC_Total / intNC_Count - dblBC) ' / 1073741824
                            If dblPC = 0 Then dblPC = (dblPC_Total / intPC_Count - dblBC) '/ 1073741824
                            
                            dblOD = dblOD - dblBC
                            
                            If dblNC < dbl���Կ���Сֵ Then dblNC = dbl���Կ���Сֵ
                            If dblOD < dbl���Կ���Сֵ Then dblOD = dbl���Կ���Сֵ
                            
                            '���ݹ�ʽ���㣬������
                            str��ʽ = strCoutOff��ʽ
                            str��ʽ = Replace(str��ʽ, "[NC]", dblNC)
                            str��ʽ = Replace(str��ʽ, "[PC]", dblPC)
                            'strCoutOff��ʽ = Replace(strCoutOff��ʽ, "[BC]", dblBC)
                            
                            dblCoutOffֵ = vbsCalce.Eval(str��ʽ)
                                
                            '���Թ�ʽ
                            str��ʽ = str���Թ�ʽ
                            str��ʽ = Replace(str��ʽ, "[NC]", dblNC)
                            str��ʽ = Replace(str��ʽ, "[PC]", dblPC)
                            'str���Թ�ʽ = Replace(str���Թ�ʽ, "[BC]", dblBC)
                            str��ʽ = Replace(str��ʽ, "[OD]", dblOD)
                            
                            bln���� = vbsCalce.Eval(str��ʽ)
                            
                            '�����Թ�ʽ
                            str��ʽ = str�����Թ�ʽ
                            str��ʽ = Replace(str��ʽ, "[NC]", dblNC)
                            str��ʽ = Replace(str��ʽ, "[PC]", dblPC)
                            'str�����Թ�ʽ = Replace(str�����Թ�ʽ, "[BC]", dblBC)
                            str��ʽ = Replace(str��ʽ, "[OD]", dblOD)
                            
                            bln������ = vbsCalce.Eval(str��ʽ)
                            If bln���� Then
                                str���Խ�� = "����"
                            Else
                                If bln������ Then
                                    str���Խ�� = "������"
                                Else
                                    str���Խ�� = "����"
                                End If
                            End If
                            
                            i��� = i��� + 1
                            If mlngSampleNo = -1 Then
                                lngResultCount = lngResultCount + 1
                                ReDim Preserve mstrReturn(lngResultCount)
                                mstrReturn(lngResultCount) = str���� & "|" & int��ʼ�걾�� + i��� & "| |Ѫ��|0|" & str��Ŀ & "|" & str���Խ�� & "^" & Format(dblOD, "0." & String(intС��, "0")) & "^" & Format(dblCoutOffֵ, "0." & String(intС��, "0")) & "^" & Format(dblOD / dblCoutOffֵ, "0." & String(intС��, "0"))
                            Else
                                If int��ʼ�걾�� + lngResultCount = mlngSampleNo Then
                                    lngResultCount = lngResultCount + 1
                                    ReDim Preserve mstrReturn(lngResultCount)
                                    mstrReturn(lngResultCount) = str���� & "|" & int��ʼ�걾�� + i��� & "| |Ѫ��|0|" & str��Ŀ & "|" & str���Խ�� & "^" & Format(dblOD, "0." & String(intС��, "0")) & "^" & Format(dblCoutOffֵ, "0." & String(intС��, "0")) & "^" & Format(dblOD / dblCoutOffֵ, "0." & String(intС��, "0"))
                                End If
                            End If
                            WriteLog strFileName, "��Ŀ���ͣ�" & strType & " λ�ã�" & strXY & " ԭʼOD=" & strOD, "BC=" & Format(dblBC, "0.000") & ",NC=" & Format(dblNC, "0.000") & ",PC=" & Format(dblPC, "0.000") & ",OD=" & Format(dblOD, "0.000")
                        End If
                        
                        
                    End If 'if strOd<>""
                Next
            End If  'iCount > 0
            WriteToIni App.Path & "\RMP150.ini", str����, str��Ŀ, int��ʼ�걾�� + i���
        
            intFileCount = intFileCount + 1
            WriteToIni App.Path & "\RMP150.ini", "ImpFile " & str����, "Count", intFileCount
            WriteToIni App.Path & "\RMP150.ini", "ImpFile " & str����, "File" & intFileCount, strFileName

        End If 'strCoutOff��ʽ <> ""
    Next
    
End Sub

Private Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''дINI�ļ�
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Private Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''��INI�ļ�
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function

Private Sub chkFile_Click()
    If chkFile.Value = 0 Then
        Call ShowItemList
    Else
        Call ShowFileList
    End If
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call ImpFile
    Unload Me
End Sub

Private Sub dtpDate_Change()
    Dim str�걾�� As String, intRow As Integer
    With vfgList
        For intRow = .FixedRows To .Rows - 1
            str�걾�� = ReadFromIni(App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), .TextMatrix(intRow, mCol.Ӣ��))
            .TextMatrix(intRow, mCol.��ʼ�걾��) = Val(str�걾��) + 1
        Next
    End With
End Sub

Private Sub Form_Load()

     
     dtpDate.MinDate = mdtStart
     dtpDate.MaxDate = mdtEnd
     dtpDate.Value = mdtStart
     Call ShowItemList
    '�ö���ʾ����
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mCol.��ʼ�걾�� Then
        vfgList.TextMatrix(Row, Col) = CLng(Val(vfgList.TextMatrix(Row, Col)))
        WriteToIni App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), vfgList.TextMatrix(Row, mCol.Ӣ��), CLng(Val(vfgList.TextMatrix(Row, Col))) - 1
    End If
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.��ʼ�걾�� Then Cancel = True
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .MouseRow <> 0 Then
            If .TextMatrix(.MouseRow, 0) = 1 Then
                .TextMatrix(.MouseRow, 0) = 0
            Else
                .TextMatrix(.MouseRow, 0) = 1
            End If
        End If
    End With
End Sub

Private Sub ShowItemList()
    '�������ļ�
    Dim iCount As Integer, iLoop As Integer
    Dim strItem As String, strItems As String, varItem As Variant
    Dim str�걾�� As String
    vfgList.Rows = 2: vfgList.Cols = 12
    vfgList.TextMatrix(0, mCol.ѡ��) = "": vfgList.ColWidth(mCol.ѡ��) = 300: vfgList.ColDataType(mCol.ѡ��) = flexDTBoolean
    vfgList.TextMatrix(0, mCol.��Ŀ���) = "���": vfgList.ColWidth(mCol.��Ŀ���) = 450
    vfgList.TextMatrix(0, mCol.Ӣ��) = "Ӣ��": vfgList.ColWidth(mCol.Ӣ��) = 1000
    vfgList.TextMatrix(0, mCol.��Ŀ����) = "��Ŀ����": vfgList.ColWidth(mCol.��Ŀ����) = 2500
    vfgList.TextMatrix(0, mCol.��ʼ�걾��) = "��ʼ�걾��": vfgList.ColWidth(mCol.��ʼ�걾��) = 1200: vfgList.ColDataType(mCol.��ʼ�걾��) = flexDTLong
    vfgList.TextMatrix(0, mCol.С��) = "С��": vfgList.ColWidth(mCol.С��) = 0
    vfgList.TextMatrix(0, mCol.CutOff��ʽ) = "CutOff��ʽ": vfgList.ColWidth(mCol.CutOff��ʽ) = 0
    vfgList.TextMatrix(0, mCol.���Թ�ʽ) = "���Թ�ʽ": vfgList.ColWidth(mCol.���Թ�ʽ) = 0
    vfgList.TextMatrix(0, mCol.������ʽ) = "������ʽ": vfgList.ColWidth(mCol.������ʽ) = 0
    vfgList.TextMatrix(0, mCol.��Сֵ) = "��Сֵ": vfgList.ColWidth(mCol.��Сֵ) = 0
    vfgList.TextMatrix(0, mCol.����ʱ��) = "ʱ��": vfgList.ColWidth(mCol.����ʱ��) = 0
    vfgList.TextMatrix(0, mCol.�ļ���) = "�ļ���": vfgList.ColWidth(mCol.�ļ���) = 0
     
     
     strItems = ReadFromIni(App.Path & "\RMP150.ini", "Base", "Item")
     
     If Len(strItems) <= 0 Then '��ΪĬ��ֵ
         Exit Sub
     Else
         varItem = Split(strItems, "|")
         For iLoop = LBound(varItem) To UBound(varItem)
             strItem = ReadFromIni(App.Path & "\RMP150.ini", varItem(iLoop), "Info")
             '���|Ӣ��|����|С��λ��|CoutOff��ʽ|���Թ�ʽ|������ʽ|���Զ�����Сֵ
             With vfgList
                 .TextMatrix(.Rows - 1, mCol.ѡ��) = 1
                 .TextMatrix(.Rows - 1, mCol.��Ŀ���) = Split(strItem, "|")(0)
                 .TextMatrix(.Rows - 1, mCol.Ӣ��) = varItem(iLoop)
                 .TextMatrix(.Rows - 1, mCol.��Ŀ����) = Split(strItem, "|")(1)
                 
                  str�걾�� = ReadFromIni(App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), varItem(iLoop))
                 .TextMatrix(.Rows - 1, mCol.��ʼ�걾��) = Val(str�걾��) + 1
                 
                 .TextMatrix(.Rows - 1, mCol.С��) = Split(strItem, "|")(2)
                 .TextMatrix(.Rows - 1, mCol.CutOff��ʽ) = Split(strItem, "|")(3)
                 .TextMatrix(.Rows - 1, mCol.���Թ�ʽ) = Split(strItem, "|")(4)
                 .TextMatrix(.Rows - 1, mCol.������ʽ) = Split(strItem, "|")(5)
                 .TextMatrix(.Rows - 1, mCol.��Сֵ) = Split(strItem, "|")(6)
                 .Rows = .Rows + 1
             End With
             
         Next
         If vfgList.Rows > 2 Then vfgList.Rows = vfgList.Rows - 1
         vfgList.Editable = flexEDKbdMouse
     End If
     
End Sub
Private Sub ShowFileList()
    '����������ʾ�������ļ��б�
    Dim objStream As TextStream
    Dim strIDX As String, strLine As String, str���� As String, strʱ�� As String, str�걾��  As String
    Dim strӢ�� As String, strDataName As String, strItem As String, strImpFile As String
    Dim strDate As String '����
    Dim lngFileCount As Long, lngLoop As Long, strTmpFile As String
    strIDX = mstrFilePath & "\RMPDAT.IDX"
    
    vfgList.Rows = 2: vfgList.Cols = 12
    vfgList.TextMatrix(0, mCol.ѡ��) = "": vfgList.ColWidth(mCol.ѡ��) = 300: vfgList.ColDataType(mCol.ѡ��) = flexDTBoolean
    vfgList.TextMatrix(0, mCol.��Ŀ���) = "���": vfgList.ColWidth(mCol.��Ŀ���) = 450
    vfgList.TextMatrix(0, mCol.Ӣ��) = "Ӣ��": vfgList.ColWidth(mCol.Ӣ��) = 1000
    vfgList.TextMatrix(0, mCol.��Ŀ����) = "��Ŀ����": vfgList.ColWidth(mCol.��Ŀ����) = 1500
    vfgList.TextMatrix(0, mCol.��ʼ�걾��) = "��ʼ�걾��": vfgList.ColWidth(mCol.��ʼ�걾��) = 1200: vfgList.ColDataType(mCol.��ʼ�걾��) = flexDTLong
    vfgList.TextMatrix(0, mCol.С��) = "С��": vfgList.ColWidth(mCol.С��) = 0
    vfgList.TextMatrix(0, mCol.CutOff��ʽ) = "CutOff��ʽ": vfgList.ColWidth(mCol.CutOff��ʽ) = 0
    vfgList.TextMatrix(0, mCol.���Թ�ʽ) = "���Թ�ʽ": vfgList.ColWidth(mCol.���Թ�ʽ) = 0
    vfgList.TextMatrix(0, mCol.������ʽ) = "������ʽ": vfgList.ColWidth(mCol.������ʽ) = 0
    vfgList.TextMatrix(0, mCol.��Сֵ) = "��Сֵ": vfgList.ColWidth(mCol.��Сֵ) = 0
    vfgList.TextMatrix(0, mCol.����ʱ��) = "ʱ��": vfgList.ColWidth(mCol.����ʱ��) = 900
    vfgList.TextMatrix(0, mCol.�ļ���) = "�ļ���": vfgList.ColWidth(mCol.�ļ���) = 0
    
    strDate = Format(dtpDate.Value, "yyyy-MM-dd")
    Set objStream = mobjFile.OpenTextFile(strIDX, ForReading)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, ".DAT=UNDEFINED,") > 0 And InStr(strLine, "RMP") > 0 Then
            str���� = Split(strLine, ",")(2)
            str���� = Right(str����, 2) & "/" & Mid(str����, 1, 5)
            strʱ�� = Split(strLine, ",")(3)
            strӢ�� = Split(strLine, ",")(1)
            
            strDataName = Replace(Split(strLine, ",")(0), "=UNDEFINED", "")
            
            If Format(CDate(str����), "yyyy-MM-dd") = strDate Then
                With vfgList
                    .TextMatrix(.Rows - 1, mCol.ѡ��) = 0
                    .TextMatrix(.Rows - 1, mCol.Ӣ��) = Mid(strӢ��, InStr(strӢ��, " ") + 1)
                    strItem = ReadFromIni(App.Path & "\RMP150.ini", Mid(strӢ��, InStr(strӢ��, " ") + 1), "Info")
                    .TextMatrix(.Rows - 1, mCol.����ʱ��) = strʱ��
                     str�걾�� = ReadFromIni(App.Path & "\RMP150.ini", strDate, .TextMatrix(.Rows - 1, mCol.Ӣ��))
                    .TextMatrix(.Rows - 1, mCol.��ʼ�걾��) = Val(str�걾��) + 1
                    .TextMatrix(.Rows - 1, mCol.��Ŀ���) = Split(strItem, "|")(0)
                    .TextMatrix(.Rows - 1, mCol.��Ŀ����) = Split(strItem, "|")(1)
                    .TextMatrix(.Rows - 1, mCol.С��) = Split(strItem, "|")(2)
                    .TextMatrix(.Rows - 1, mCol.CutOff��ʽ) = Split(strItem, "|")(3)
                    .TextMatrix(.Rows - 1, mCol.���Թ�ʽ) = Split(strItem, "|")(4)
                    .TextMatrix(.Rows - 1, mCol.������ʽ) = Split(strItem, "|")(5)
                    .TextMatrix(.Rows - 1, mCol.��Сֵ) = Split(strItem, "|")(6)
                    .TextMatrix(.Rows - 1, mCol.�ļ���) = strDataName
                    

                    .Rows = .Rows + 1
                End With
                
            End If
        End If
    Loop
    If vfgList.Rows > 2 Then vfgList.Rows = vfgList.Rows - 1
    
End Sub
