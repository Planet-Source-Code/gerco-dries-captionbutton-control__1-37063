Attribute VB_Name = "MDebug"
Option Explicit


Public Sub Log(obj As Object, sText As String)
#If DEBUGMODE Then
    
    Dim h As Integer
    h = FreeFile
    Open App.Path & "\Debug.log" For Append As h
        #If DEBUGLOG Then
            Print #h, Timer & ": " & Hex$(ObjPtr(obj)) & " >> " & sText
        #End If
        #If DEBUGIDE Then
            Debug.Print Timer & ": " & Hex$(ObjPtr(obj)) & " >> "; sText
        #End If
    Close h
#End If

End Sub


