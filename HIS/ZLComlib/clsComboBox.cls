VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#���ܣ�ComboBox��صĴ������жϵȹ�������
'�������򣺺�����������Ϊͨ���׶���Ӣ�ļ��ɣ����ü�Cbo�ؼ���
'�ϼ��ࣺclsControl
'##############################################################################

Public Sub SetListWidth(ByVal cboControl As Variant, ByVal lngWidth As Long)
'���ܣ�����Combo�ؼ������б��Ŀ���
'�˴��Ŀ������������б��Ŀ��ȣ���������TWIPΪ��λ
    If IsObject(cboControl) Then
        Call gobjComLib.zlControl.CboSetWidth(cboControl.hWnd, lngWidth)
    Else
        Call gobjComLib.zlControl.CboSetWidth(cboControl, lngWidth)
    End If
End Sub

Public Sub SetListWidthAuto(cboControl As Object, Optional ByVal dblMaxRate As Double)
'���ܣ��Զ�����Combo�ؼ������б��������������б�����
'������cboControl=�Զ����������б��Ŀؼ�
'          intMaxRate=���ޱ��ʣ��Զ����ú�������б��Ŀ��ȷ�ΧΪcboControl.Width<= �����б�����<=cboControl.Width*intMaxRate
'                               <2��ʾ���޿���
'˵�����ú������ڼ���ComboBox�����б�������
    Dim i As Long
    Dim lngTmp As Long, lngMaxLen As Long
    Dim lngWidth As Long
    Dim lngOneChar As Long
    
    For i = 0 To cboControl.ListCount - 1
        lngTmp = gobjComLib.zlStr.ActualLen(cboControl.List(i))
        If lngTmp > lngMaxLen Then
            lngMaxLen = lngTmp
        End If
    Next
    If lngMaxLen <> 0 Then
        If dblMaxRate < 1 Then dblMaxRate = 0
        lngOneChar = gobjComLib.zlControl.OneCharWidth(cboControl.Font)
        lngWidth = CLng((lngMaxLen + 4) * lngOneChar)   '�������,��4������Ϊ��������߽��ֵ
        If lngWidth < cboControl.Width Then '������С���ȣ���Ϊ��С����
            lngWidth = cboControl.Width
        ElseIf dblMaxRate <> 0 And lngWidth > cboControl.Width * dblMaxRate Then '�����趨������ʣ���Ϊ������
            lngWidth = CLng(cboControl.Width * dblMaxRate)
        End If
    Else
        lngWidth = cboControl.Width 'Ĭ����ؼ�����һ��
    End If
    Call SetListWidth(cboControl.hWnd, lngWidth)
End Sub

Public Sub SetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'���ܣ������ı�������Combo�ؼ��ĵ�ǰֵ
'������cboControl  ׼�����õ�ComboBox�ؼ�
'      strText     ������ı���
'      blnAfter    ��ʾ�ڷָ���֮ǰ��֮��ȡֵ�����û�зָ�������ȡ֮��
'      strSplit    �ָ�����ͨ��Ϊ-
    Call gobjComLib.zlControl.CboSetText(cboControl, strText, blnAfter, strSplit)
End Sub

Public Sub SeekIndex(objCbo As Object, varFind As Variant, Optional Keep As Boolean, Optional ByVal blnOnlyExact As Boolean)
'���ܣ����ַ�����ComboBox�в�������
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
'          blnOnlyExact:ֻ��ȷƥ�䣬�ò���ֻ����ַ�
    Dim i As Integer
    i = FindIndex(objCbo, varFind, blnOnlyExact)
    If i <> -1 Then
        objCbo.ListIndex = i
    ElseIf Not Keep Then
        objCbo.ListIndex = -1
    End If
End Sub

Public Sub SetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'���ܣ�����Combo�ؼ���Indexֵ
'Ϊһ��Combo�ؼ�ѡ���б�����ֲ�������Click�¼�
    Call gobjComLib.zlControl.CboSetIndex(hWnd_combo, lngIndex)
End Sub

Public Sub SetListHeight(cboControl As Object, ByVal lngHeight As Long)
'���ܣ�����Combo�ؼ��������б��ĸ߶�
'�˴��Ŀ������������б��ĸ߶ȣ���������TWIPΪ��λ
    Call gobjComLib.zlControl.CboSetHeight(cboControl, lngHeight)
End Sub

Public Function MatchIndex(ByVal lnghwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�ƥ�������(����ƥ�������)
'˵����ֻ��������,��δ�����������뽫�ú�����KeyPress�¼��е��á�
     MatchIndex = gobjComLib.zlControl.CboMatchIndex(lnghwnd, KeyAscii, sngInterval)
End Function

Public Function Locate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'�������ã�ʹ��Cbo.SeekIndex����
'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Locate = gobjComLib.zlControl.CboLocate(cboObj, strValue, blnItem)
End Function

Public Sub LoadFromList(ByRef arrCbo As Variant, ByVal varList As Variant, Optional ByVal intDefault As Integer = -1, Optional ByVal strSplit As String = ",")
'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String��������Զ��ŷָ���ַ���
'      arrCbo=һ��ComboBox���󣬻���ComboBox���������
'      intDefaut=ȱʡ����
'     strSplit=�����ַ���ʱ����Ŀ�ָ���
    Dim i As Long, j As Long
    Dim arrTmp As Variant
    Dim arrCboTmp As Variant
    If TypeName(arrCbo) = "Variant()" Then
        arrCboTmp = arrCbo
    Else
        arrCboTmp = Array(arrCbo)
    End If
    
    If TypeName(varList) = "String" Then
        arrTmp = Split(varList, strSplit)
        For i = 0 To UBound(arrCboTmp)
            arrCboTmp(i).Clear
            For j = 0 To UBound(arrTmp)
                arrCboTmp(i).AddItem arrTmp(j)
            Next
            arrCboTmp(i).ListIndex = intDefault 'ȱʡΪδѡ��
        Next
    Else
        For i = 0 To UBound(arrCboTmp)
             arrCboTmp(i).Clear
            For j = 0 To UBound(varList)
                arrCboTmp(i).AddItem varList(j)
            Next
            arrCboTmp(i).ListIndex = intDefault 'ȱʡΪδѡ��
        Next
    End If
End Sub

Public Function ListWidth(ByVal hWnd_combo As Long) As Long
'���ܣ� ȡ�� Combo �����Ŀ���,��TWIPΪ��λ
    Dim lRetVal As Long
    lRetVal = SendMessage(hWnd_combo, &H15F, 0, 0)
    If lRetVal <> -1 Then
        ListWidth = lRetVal * Screen.TwipsPerPixelX
    Else
        ListWidth = 0
    End If
End Function

Public Function ItemExists(objControl As Object, ByVal strList As String, ByVal lngData As Long, Optional ByVal intCheckType As Integer = 2) As Boolean
'���ܣ��ж�ָ����Ŀ���б����Ƿ��Ѿ�����
'˵������ͬ��ĿָText��ItemData����ͬ
'          intCheckType=2-�����Ŀ��Text��ItemData�Ƿ���ͬ,0-�������Ŀ��Text�Ƿ���ͬ��1-�������Ŀ��ItemData�Ƿ���ͬ
    Dim i As Long
    Dim blnFind As Boolean
    
    For i = 0 To objControl.ListCount - 1
        blnFind = True
        If intCheckType <> 1 Then
            If objControl.List(i) = strList Then
                blnFind = False
            End If
        End If
        If intCheckType <> 0 And blnFind Then
            If objControl.ItemData(i) = lngData Then
                blnFind = False
            End If
        End If
        If blnFind Then
            ItemExists = True: Exit For
        End If
    Next
End Function

Public Function FindIndex(objCbo As Object, varData As Variant, Optional ByVal blnOnlyExact As Boolean) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
'blnOnlyExact:ֻ��ȷƥ�䣬�ò���ֻ����ַ�
    Dim strType As String, i As Integer
    
    FindIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    FindIndex = i: Exit Function
                ElseIf gobjComLib.zlStr.NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    FindIndex = i: Exit Function
                End If
            Next
            If Not blnOnlyExact Then
                '��ģ������
                For i = 0 To objCbo.ListCount - 1
                    If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                        FindIndex = i: Exit Function
                    End If
                Next
            End If
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    FindIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Sub AppendText(cboControl As Object, KeyAscii As Integer)
'���ܣ���ComboBoxʵ������������Զ���ɵĹ���
'˵������Combox.KeyPress�¼��е���
    Call gobjComLib.zlControl.CboAppendText(cboControl, KeyAscii)
End Sub


Public Sub AddData(objSource As Object, ByVal rsDataSource As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'����:װ��������ָ�������������������е���������
'����:objSource   Ҫװ�ؼ�¼����������ؼ�
'     rsDataSource     ��¼������,ֻ��������������,��һ������������,�ڶ����Ƕ�Ӧ�Ĺؼ���,��id
'     blnClear    װ��ʱ�Ƿ����ԭ�е���������,ȱʡΪTrue
'����:
    Call gobjComLib.zlControl.CboAddData(objSource, rsDataSource, blnClear)
End Sub

