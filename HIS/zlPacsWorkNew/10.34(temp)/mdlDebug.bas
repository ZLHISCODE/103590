Attribute VB_Name = "mdlDebug"
Option Explicit

Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)


Public Sub OutputDebug(ByVal strMethob As String, objErr As ErrObject)
    OutputDebugString "[" & App.ProductName & "]" & strMethob & "£º" & objErr.Description
End Sub


Public Sub RaiseErr(objErr As ErrObject)
    Call err.Raise(objErr.Number, objErr.Source, objErr.Description, objErr.HelpFile, objErr.HelpContext)
End Sub
