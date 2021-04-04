VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################

Public Enum CellProperty
    AligmentV = 1                           '��ֱ����
    AligmentH = 2                           'ˮƽ����
    FontName = 3
    FontSize = 4
    ForeColor = 5
    WordWarp = 6
    MuliLine = 7
    AutoFit = 8
    Data = 9
    OutPutV = 10
End Enum

Private mrsGrIDFormat As ADODB.Recordset
Private mrsGrIDData As ADODB.Recordset
Private mrsGrIDRowFormat As ADODB.Recordset
Private mrsGrIDColFormat As ADODB.Recordset

Private mclsLisReportLib As clsLisReportLib

'######################################################################################################################
Public Function InitGrID(ByVal bytAligment As Byte, ByVal lngForeColor As Long, ByVal lngBackColor As Long, ByVal strFontName As String, ByVal sglFontNameSize As Single, Optional ByVal strGrIDLine As String = "111111", Optional ByVal bytAligmentV As Byte = 2) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʼ�����������
    '������ rsData��- ��Ԫ�����ݣ�rsRowFormat -  �и�ʽ˵����rsColFormat - �и�ʽ˵����
    '���أ�
    '******************************************************************************************************************
    
    Set mrsGrIDFormat = New ADODB.Recordset
    With mrsGrIDFormat
        .Fields.Append "������", adVarChar, 12
        .Fields.Append "ǰ��ɫ", adVarChar, 20
        .Fields.Append "����ɫ", adVarChar, 20
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "��С", adVarChar, 10
        .Fields.Append "ˮƽ����", adVarChar, 2
        .Fields.Append "��ֱ����", adVarChar, 2
        .Open
    End With
        
    Set mrsGrIDColFormat = New ADODB.Recordset
    With mrsGrIDColFormat
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�п�", adVarChar, 30
        .Fields.Append "����", adTinyInt
        .Fields.Append "�߿�", adTinyInt
        .Open
    End With
    
    Set mrsGrIDRowFormat = New ADODB.Recordset
    With mrsGrIDRowFormat
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�и�", adVarChar, 30
        .Fields.Append "����", adTinyInt
        .Fields.Append "�߿�", adTinyInt
        .Open
    End With
    
    Set mrsGrIDData = New ADODB.Recordset
    With mrsGrIDData
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "�к�", adTinyInt
        .Fields.Append "����", adVarChar, 1000
        .Fields.Append "ǰ��ɫ", adVarChar, 20
        .Fields.Append "����ɫ", adVarChar, 20
        .Fields.Append "����", adVarChar, 20
        .Fields.Append "��С", adVarChar, 10
        .Fields.Append "ˮƽ����", adVarChar, 2
        .Fields.Append "��ֱ����", adVarChar, 2
        .Fields.Append "�Զ�����", adTinyInt
        .Fields.Append "�������", adTinyInt
        .Fields.Append "�Զ���Ӧ", adTinyInt
        .Fields.Append "��ת�Ƕ�", adTinyInt   '0-����;1-���ϵ�������;2-���ϵ���������ת90��;3-���µ�������;4-���µ�����ת90��
        .Open
    End With

    mrsGrIDFormat.AddNew
    mrsGrIDFormat("ˮƽ����").Value = bytAligment
    mrsGrIDFormat("��ֱ����").Value = bytAligmentV
    mrsGrIDFormat("ǰ��ɫ").Value = lngForeColor
    mrsGrIDFormat("����ɫ").Value = lngBackColor
    mrsGrIDFormat("����").Value = strFontName
    mrsGrIDFormat("��С").Value = sglFontNameSize
    mrsGrIDFormat("������").Value = Mid(strGrIDLine, 1, 1) & ";" & Mid(strGrIDLine, 2, 1) & ";" & Mid(strGrIDLine, 3, 1) & ";" & Mid(strGrIDLine, 4, 1) & ";" & Mid(strGrIDLine, 5, 1) & ";" & Mid(strGrIDLine, 6, 1)
    
    InitGrID = True
    
End Function

Public Property Let Cell(ByVal Property As CellProperty, ByVal intRow As Integer, ByVal intCol As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "�к�=" & intRow & " And �к�=" & intCol
    
    If mrsGrIDData.RecordCount > 0 Then
        Select Case Property
        Case CellProperty.AligmentH
            mrsGrIDData("ˮƽ����").Value = Val(vData)
        Case CellProperty.AligmentV
            mrsGrIDData("��ֱ����").Value = Val(vData)
        Case CellProperty.FontName
            mrsGrIDData("����").Value = Trim(vData)
        Case CellProperty.FontSize
            mrsGrIDData("��С").Value = Val(vData)
        Case CellProperty.ForeColor
            mrsGrIDData("ǰ��ɫ").Value = Val(vData)
        Case CellProperty.WordWarp
            mrsGrIDData("�Զ�����").Value = Val(vData)
        Case CellProperty.MuliLine
            mrsGrIDData("�������").Value = Val(vData)
        Case CellProperty.AutoFit
            mrsGrIDData("�Զ���Ӧ").Value = Val(vData)
        Case CellProperty.Data
            mrsGrIDData("����").Value = vData
        Case CellProperty.OutPutV
            mrsGrIDData("��ת�Ƕ�").Value = Val(vData)
        End Select
    Else
        Select Case Property
        Case CellProperty.Data
    
            mrsGrIDData.AddNew
            mrsGrIDData("�к�").Value = intRow
            mrsGrIDData("�к�").Value = intCol
            mrsGrIDData("����").Value = vData
            mrsGrIDData("����").Value = mrsGrIDFormat("����").Value
            mrsGrIDData("��С").Value = mrsGrIDFormat("��С").Value
            mrsGrIDData("ǰ��ɫ").Value = mrsGrIDFormat("ǰ��ɫ").Value
            mrsGrIDData("����ɫ").Value = mrsGrIDFormat("����ɫ").Value
            mrsGrIDData("ˮƽ����").Value = mrsGrIDFormat("ˮƽ����").Value
            mrsGrIDData("��ֱ����").Value = mrsGrIDFormat("��ֱ����").Value
            mrsGrIDData("�Զ�����").Value = 0
            mrsGrIDData("�������").Value = 0
            mrsGrIDData("�Զ���Ӧ").Value = 0
            mrsGrIDData("��ת�Ƕ�").Value = 0
        End Select
    End If
    

        
End Property

Public Property Let RowCell(ByVal Property As CellProperty, ByVal intRow1 As Integer, ByVal intRow2 As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "�к�>=" & intRow1 & " And �к�<=" & intRow2
    
    If mrsGrIDData.RecordCount > 0 Then
        mrsGrIDData.MoveFirst
        Do While Not mrsGrIDData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGrIDData("ˮƽ����").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGrIDData("��ֱ����").Value = Val(vData)
            Case CellProperty.FontName
                mrsGrIDData("����").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGrIDData("��С").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGrIDData("ǰ��ɫ").Value = Val(vData)
            End Select
            
            mrsGrIDData.MoveNext
        Loop
    End If
    
End Property

Public Property Let ColCell(ByVal Property As CellProperty, ByVal intCol1 As Integer, ByVal intCol2 As Integer, ByVal vData As Variant)

    mrsGrIDData.Filter = ""
    mrsGrIDData.Filter = "�к�>=" & intCol1 & " And �к�<=" & intCol2
    
    If mrsGrIDData.RecordCount > 0 Then
        mrsGrIDData.MoveFirst
        Do While Not mrsGrIDData.EOF
        
            Select Case Property
            Case CellProperty.AligmentH
                mrsGrIDData("ˮƽ����").Value = Val(vData)
            Case CellProperty.AligmentV
                mrsGrIDData("��ֱ����").Value = Val(vData)
            Case CellProperty.FontName
                mrsGrIDData("����").Value = Trim(vData)
            Case CellProperty.FontSize
                mrsGrIDData("��С").Value = Val(vData)
            Case CellProperty.ForeColor
                mrsGrIDData("ǰ��ɫ").Value = Val(vData)
            End Select
            
            mrsGrIDData.MoveNext
        Loop
    End If
    
End Property


Public Function SetGrIDRowFormat(ByVal strRowHeight As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strRowHeight, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDRowFormat.AddNew
        mrsGrIDRowFormat("�к�").Value = intLoop + 1
        mrsGrIDRowFormat("�и�").Value = mclsLisReportLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGrIDRowFormat("�߿�").Value = 1
        mrsGrIDRowFormat("����").Value = 0
    Next
    
    SetGrIDRowFormat = True
End Function

Public Function SetGrIDRowHeight(ByVal intRow As Integer, ByVal sglHeight As Single) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "�к�=" & intRow + 1
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("�и�").Value = mclsLisReportLib.GetTwipsY(sglHeight)
    End If
    mrsGrIDRowFormat.Filter = ""

    SetGrIDRowHeight = True
    
End Function

Public Function SetGrIDRowLineType(ByVal intRow As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "�к�=" & intRow
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("�߿�").Value = bytLineType
    End If
    
    SetGrIDRowLineType = True
End Function

Public Function SetGrIDRowLineStyle(ByVal intRow As Integer, ByVal bytLineStyle As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDRowFormat.Filter = "�к�=" & intRow
    If mrsGrIDRowFormat.RecordCount > 0 Then
        mrsGrIDRowFormat("����").Value = bytLineStyle
    End If
    
    SetGrIDRowLineStyle = True
End Function

Public Function SetGrIDColFormat(ByVal strColWidth As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strColWidth, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDColFormat.AddNew
        mrsGrIDColFormat("�к�").Value = intLoop + 1
        mrsGrIDColFormat("�п�").Value = mclsLisReportLib.GetTwipsY(Val(varTmp(intLoop)))
        mrsGrIDColFormat("�߿�").Value = 1
        mrsGrIDRowFormat("����").Value = 0
    Next
    
    SetGrIDColFormat = True
    
End Function

Public Function SetGrIDColLineType(ByVal intCol As Integer, ByVal bytLineType As Byte) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mrsGrIDColFormat.Filter = ""
    mrsGrIDColFormat.Filter = "�к�=" & intCol
    If mrsGrIDColFormat.RecordCount > 0 Then
        mrsGrIDColFormat("�߿�").Value = bytLineType
    End If
    
    SetGrIDColLineType = True
End Function


Public Function Width() As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGrIDColFormat.Filter = ""
    intCols = mrsGrIDColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGrIDColFormat.Filter = ""
        mrsGrIDColFormat.Filter = "�к�=" & intLoop
        If mrsGrIDColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGrIDColFormat("�п�").Value)

        End If
    Next
    
    Width = mclsLisReportLib.GetCentimetersX(sglWidth)
End Function

Public Function SumGrIDColWidth(Optional ByVal intCol As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglWidth As Single
    Dim intCols As Integer
    
    sglWidth = 0
    mrsGrIDColFormat.Filter = ""
    intCols = mrsGrIDColFormat.RecordCount
    For intLoop = 1 To intCols

        mrsGrIDColFormat.Filter = ""
        mrsGrIDColFormat.Filter = "�к�=" & intLoop
        If mrsGrIDColFormat.RecordCount > 0 Then
'
            sglWidth = sglWidth + Val(mrsGrIDColFormat("�п�").Value)
            
            If intLoop = intCol Then
                SumGrIDColWidth = sglWidth
                Exit Function
            End If

        End If
    Next
    
    SumGrIDColWidth = sglWidth
End Function

Public Function Height(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGrIDRowFormat.Filter = ""
    intRows = mrsGrIDRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "�к�=" & intLoop
        If mrsGrIDRowFormat.RecordCount > 0 Then
            
            sglHeight = sglHeight + mclsLisReportLib.GetCentimetersY(Val(mrsGrIDRowFormat("�и�").Value))
            
            If intLoop = intRow Then
                Height = sglHeight
                Exit Function
            End If
            
        End If
    Next
    
    Height = sglHeight
End Function

Public Function SumGrIDRowHeight(Optional ByVal intRow As Integer) As Single
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim sglHeight As Single
    Dim intRows As Integer
    
    sglHeight = 0
    
    mrsGrIDRowFormat.Filter = ""
    intRows = mrsGrIDRowFormat.RecordCount
    For intLoop = 1 To intRows

        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "�к�=" & intLoop
        If mrsGrIDRowFormat.RecordCount > 0 Then
'
            sglHeight = sglHeight + Val(mrsGrIDRowFormat("�и�").Value)
            
            If intLoop = intRow Then
                SumGrIDRowHeight = sglHeight
                Exit Function
            End If

        End If
    Next
    
    SumGrIDRowHeight = sglHeight
End Function

Public Function SetGrIDColData(ByVal intCol As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDData.AddNew
        mrsGrIDData("�к�").Value = intLoop + 1
        mrsGrIDData("�к�").Value = intCol
        mrsGrIDData("����").Value = CStr(varTmp(intLoop))
        mrsGrIDData("����").Value = mrsGrIDFormat("����").Value
        mrsGrIDData("��С").Value = mrsGrIDFormat("��С").Value
        mrsGrIDData("ǰ��ɫ").Value = mrsGrIDFormat("ǰ��ɫ").Value
        mrsGrIDData("����ɫ").Value = mrsGrIDFormat("����ɫ").Value
        mrsGrIDData("ˮƽ����").Value = mrsGrIDFormat("ˮƽ����").Value
        mrsGrIDData("��ֱ����").Value = mrsGrIDFormat("��ֱ����").Value
        mrsGrIDData("�Զ�����").Value = 0
        mrsGrIDData("�������").Value = 0
        mrsGrIDData("�Զ���Ӧ").Value = 0
        mrsGrIDData("��ת�Ƕ�").Value = 0
'        mrsGrIDData("ͼ���ļ�").Value = ""
    Next
    
    SetGrIDColData = True
    
End Function

Public Function SetGrIDRowData(ByVal intRow As Integer, ByVal strData As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim varTmp As Variant

    varTmp = Split(strData, ";")
    
    For intLoop = 0 To UBound(varTmp)
        mrsGrIDData.AddNew
        mrsGrIDData("�к�").Value = intRow
        mrsGrIDData("�к�").Value = intLoop + 1
        mrsGrIDData("����").Value = CStr(varTmp(intLoop))
        mrsGrIDData("����").Value = mrsGrIDFormat("����").Value
        mrsGrIDData("��С").Value = mrsGrIDFormat("��С").Value
        mrsGrIDData("ǰ��ɫ").Value = mrsGrIDFormat("ǰ��ɫ").Value
        mrsGrIDData("����ɫ").Value = mrsGrIDFormat("����ɫ").Value
        mrsGrIDData("ˮƽ����").Value = mrsGrIDFormat("ˮƽ����").Value
        mrsGrIDData("��ֱ����").Value = mrsGrIDFormat("��ֱ����").Value
        mrsGrIDData("�Զ�����").Value = 0
        mrsGrIDData("�������").Value = 0
        mrsGrIDData("�Զ���Ӧ").Value = 0
        mrsGrIDData("��ת�Ƕ�").Value = 0
'        mrsGrIDData("ͼ���ļ�").Value = ""
    Next
    
    SetGrIDRowData = True
    
End Function

Friend Function FormatData(Optional ByVal strKind As String = "����") As Boolean
    '******************************************************************************************************************
    '���ܣ� ������������
    '������ mrsGrIDData��- ��Ԫ�����ݣ�mrsGrIDRowFormat -  �и�ʽ˵����mrsGrIDColFormat - �и�ʽ˵����
    '       gobjRect - ������������Ͻ���������½�����
    '���أ�
    '******************************************************************************************************************
    Dim strTmp As String
    Dim intRows As Integer
    Dim intCols As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim X0 As Single
    Dim Y0 As Single
    Dim X1 As Single
    Dim Y1 As Single
    Dim varGrIDLine As Variant
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim objSvrFont As USERFONT
    Dim objSvrRect As USERRECT
    
    On Error GoTo errHand
    
    Call SaveFont(gobjFont, objSvrFont)
    Call SaveRect(gobjRect, objSvrRect)
    
    X0 = gobjRect.X0
    Y0 = gobjRect.Y0
    X1 = gobjRect.X1
    Y1 = gobjRect.Y1
    
    mrsGrIDRowFormat.Filter = ""
    mrsGrIDColFormat.Filter = ""
    
    intRows = mrsGrIDRowFormat.RecordCount
    intCols = mrsGrIDColFormat.RecordCount
    
    '�Ȼ�������
    '------------------------------------------------------------------------------------------------------------------
    varGrIDLine = Split(mrsGrIDFormat("������").Value, ";")
    
    '����Ϻ���
    If Val(varGrIDLine(0)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y0
        Call AppendPrintData(strKind, "����")
    End If
    
    '���������
    If Val(varGrIDLine(1)) > 0 Then
        gobjRect.X0 = X1
        gobjRect.Y0 = Y0
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
    
    End If
    
    '����º���
    If Val(varGrIDLine(2)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y1
        gobjRect.X1 = X1
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
    
    End If
    
    '���������
    If Val(varGrIDLine(3)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.X1 = X0
        gobjRect.Y1 = Y1
        Call AppendPrintData(strKind, "����")
            
    End If
    
    
    '�ڿ򣺺���
    gobjFont.LineWidth = 1
    If Val(varGrIDLine(4)) > 0 Then
        gobjRect.Y0 = Y0
        gobjRect.X0 = X0
        gobjRect.X1 = X1
        For intRow = 1 To intRows - 1
            mrsGrIDRowFormat.Filter = ""
            mrsGrIDRowFormat.Filter = "�к�=" & intRow
            If mrsGrIDRowFormat.RecordCount > 0 Then
                gobjRect.Y0 = gobjRect.Y0 + Val(mrsGrIDRowFormat("�и�").Value)
                gobjRect.Y1 = gobjRect.Y0
                If Val(mrsGrIDRowFormat("�߿�").Value) <> 0 Then
                    gobjFont.LineWidth = Val(mrsGrIDRowFormat("�߿�").Value)
                    gobjFont.LineStyle = Val(mrsGrIDRowFormat("����").Value)
                    Call AppendPrintData(strKind, "����")
                End If
            End If
        Next
    End If
        
    '�ڿ�����
    If Val(varGrIDLine(5)) > 0 Then
        gobjRect.X0 = X0
        gobjRect.Y0 = Y0
        gobjRect.Y1 = Y1
        For intCol = 1 To intCols - 1
            mrsGrIDColFormat.Filter = ""
            mrsGrIDColFormat.Filter = "�к�=" & intCol
            If mrsGrIDColFormat.RecordCount > 0 Then
                gobjRect.X0 = gobjRect.X0 + Val(mrsGrIDColFormat("�п�").Value)
                gobjRect.X1 = gobjRect.X0
                gobjFont.LineWidth = Val(mrsGrIDColFormat("�߿�").Value)
                gobjFont.LineStyle = Val(mrsGrIDRowFormat("����").Value)
                Call AppendPrintData(strKind, "����")
            End If
        Next
    End If
    gobjFont.LineWidth = 1
    
    '��д����
    '------------------------------------------------------------------------------------------------------------------
    gobjRect.Y1 = Y0
    For intRow = 1 To intRows
        
        gobjRect.Y0 = gobjRect.Y1
        gobjRect.X0 = X0
        
        mrsGrIDRowFormat.Filter = ""
        mrsGrIDRowFormat.Filter = "�к�=" & intRow
        
        If mrsGrIDRowFormat.RecordCount > 0 Then

            gobjRect.Y1 = gobjRect.Y0 + Val(mrsGrIDRowFormat("�и�").Value)
            gobjRect.X1 = X0
            
            For intCol = 1 To intCols
                
                mrsGrIDColFormat.Filter = ""
                mrsGrIDColFormat.Filter = "�к�=" & intCol
                    
                If mrsGrIDColFormat.RecordCount > 0 Then
                
                    gobjRect.X0 = gobjRect.X1
                    gobjRect.X1 = gobjRect.X0 + Val(mrsGrIDColFormat("�п�").Value)
                                        
                    mrsGrIDData.Filter = ""
                    mrsGrIDData.Filter = "�к�=" & intRow & " And �к�=" & intCol
                    If mrsGrIDData.RecordCount > 0 Then
     
                        strTmp = mrsGrIDData("����").Value
                    
                        gobjFont.Name = mrsGrIDData("����").Value
                        gobjFont.Size = mrsGrIDData("��С").Value
                        gobjFont.BackColor = Val(mrsGrIDData("����ɫ").Value)
                        gobjFont.ForeColor = Val(mrsGrIDData("ǰ��ɫ").Value)

                        gobjFont.Bold = False
                        gobjFont.Underline = False
                        Call SetDraw(gobjFont)
                        
                        '<Image>c:\a<Image>"
                        intPos1 = InStr(strTmp, "<Image>")
                        intPos2 = InStr(strTmp, "</Image>")
                        If intPos1 > 0 And intPos2 > 0 And intPos1 < intPos2 Then
                            strTmp = Mid(strTmp, intPos1 + 7, intPos2 - intPos1 - 7)
                            If Dir(strTmp) <> "" And strTmp <> "" Then
                                Call AppendPrintData(strKind, "ͼ��", Val(mrsGrIDData("ˮƽ����").Value), False, strTmp)
                            End If
                        Else
                        
                            If Val(mrsGrIDData("ˮƽ����").Value) = 1 Then gobjRect.X0 = gobjRect.X0 + mclsLisReportLib.GetTwipsX(0.05)
                            If Val(mrsGrIDData("��ֱ����").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 + mclsLisReportLib.GetTwipsY(0.05)
                        
                            Call AppendPrintData(strKind, "�ı�", Val(mrsGrIDData("ˮƽ����").Value), (Val(mrsGrIDData("�Զ�����").Value) = 1), strTmp, Val(mrsGrIDData("��ֱ����").Value), (Val(mrsGrIDData("�������").Value) = 1), , (Val(mrsGrIDData("�Զ���Ӧ").Value) = 1), , , Val(mrsGrIDData("��ת�Ƕ�").Value))
                            
                        
                            If Val(mrsGrIDData("ˮƽ����").Value) = 1 Then gobjRect.X0 = gobjRect.X0 - mclsLisReportLib.GetTwipsX(0.05)
                            If Val(mrsGrIDData("��ֱ����").Value) = 1 Then gobjRect.Y0 = gobjRect.Y0 - mclsLisReportLib.GetTwipsY(0.05)
                        
                        End If

                        
                    End If
                End If
            Next
        End If
    Next
    
    Call SaveFont(objSvrFont, gobjFont)
    Call SaveRect(objSvrRect, gobjRect)
    
    FormatData = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call SaveFont(objSvrFont, gobjFont)
End Function

Private Sub Class_Initialize()
    Set mclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mclsLisReportLib = Nothing
End Sub