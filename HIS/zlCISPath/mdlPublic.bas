Attribute VB_Name = "mdlPublic"
Option Explicit
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function ClientToScreen Lib "user32" (ByVal Hwnd As Long, lpPoint As POINTAPI) As Long


Public Const ETO_OPAQUE = 2
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'Shell------------------------------------------------
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'公共方法--------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetParTable(ByVal strPar As String, ByVal strParTable As String, ByRef strTableOut As String, ByRef intMaxIdx As Integer) As Variant
'功能：对于动态内存表的绑定参数超长的处理
'参数：strPar 参数串，strParTable 内存表形式要传入
'返回：一个字符串数组，10个元素
    Dim n As Long, p As Long
    Dim varPar(0 To 9) As String
    Dim strTable As String, strThis As String
    Dim intNum As Integer '参数号
    
    For n = 0 To 9
        varPar(n) = ""
    Next
    
    p = InStr(strParTable, "[") + 1
    intNum = Mid(strParTable, p, 1)
    
    n = 0
    Do While True
        If Len(strPar) < 4000 Then
            p = Len(strPar) + 1
        Else
            p = InStrRev(Mid(strPar, 1, 4000), ",")
        End If
        
        strThis = Mid(strPar, 1, p - 1)
        
        If n > 9 Then
            strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "'" & strThis & "'")
        Else
            varPar(n) = strThis
            If n = 0 Then
                strTable = strParTable
                intMaxIdx = intNum
            Else
                intMaxIdx = (n + intNum)
                strTable = strTable & vbNewLine & " Union All " & Replace(strParTable, "[" & intNum & "]", "[" & intMaxIdx & "]")
            End If
        End If
        
        n = n + 1
        
        strPar = Mid(strPar, p + 1)
        
        If strPar = "" Then Exit Do
    Loop
    
    strTableOut = strTable
    GetParTable = varPar
    
End Function


Public Function CreateNode(ByVal TabNumber As Integer, _
    ByVal Parent As IXMLDOMNode, _
    Optional ByVal Node_Name As String, _
    Optional ByVal Node_Type As tagDOMNodeType = NODE_ELEMENT, _
    Optional ByVal Node_Value As String = "") As IXMLDOMNode
    
    Dim New_Node As IXMLDOMNode
    
    '字符缩进值设置（不影响数据），只影响阅读美观度
    Parent.appendChild Parent.ownerDocument.createTextNode(vbCrLf & String(TabNumber, vbKeyTab))   '创建文本节点
    
    '创建新节点
    Set New_Node = Parent.ownerDocument.CreateNode(Node_Type, Node_Name, "")
    
    '设置文本值
    New_Node.Text = Node_Value
    
    '添加到父节点
    Parent.appendChild New_Node

    Set CreateNode = New_Node
End Function

Public Function GetNodeValue(ByVal CurNode As IXMLDOMNode, ByVal SubNodeName As String, Optional ByVal DefaultValue As String) As String
    Dim NodeTMP As IXMLDOMNode
    
    Set NodeTMP = CurNode.selectSingleNode(".//" & SubNodeName)
    If NodeTMP Is Nothing Then
        GetNodeValue = DefaultValue
    Else
        GetNodeValue = NodeTMP.Text
    End If
End Function
