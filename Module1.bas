Attribute VB_Name = "Module1"
' --- Globals
Dim objHTTP As Object
Dim url, CeresTypeC As String
Dim lResolve, lConnect, lSend, lReceive As Long
Dim PromoNo, VersionLT, User, Json, result, PromoIndicator, CurStatus As String
Dim FirstColon, WordTracking, Diff, LastCol As Integer
Dim GetMsgBox As Variant
Dim strData() As String
Dim MyData As String


Sub UpdateIdealQtyFromSheet()

' ----------------------------------------------------------------------
' Loop through the work sheet filling variable Json
' ----------------------------------------------------------------------
    
' --- Set idx table to work sheet: Hard code to 1000 for now.  Set delimiter to pipe and insert CR at end of each row.  Skip non field fields
    Application.StatusBar = "Get Spread Sheet details'"
    Set IdealQtys = Range("a2:p1000")
    Json = Empty
    For i = 1 To 1000
        For k = 1 To 17
            If k <> 3 And k <> 4 And k <> 5 Then
                If k = 17 Then
                    Json = Json & Chr(10)
                Else
                    ChkValue = IdealQtys(i, k).Value
                    Json = Json & IdealQtys(i, k) & "|"
                End If
            End If
        Next
    Next
    
    Application.StatusBar = "Call the Web Service to update the table with the false/no auto EOD "
    
    Call ExecuteThePost("false")
    
    Application.DisplayAlerts = True
    Application.StatusBar = False
    
    
End Sub

    
Sub ExecuteThePost(RunBranch)
' ----------------------------------------------------------------------
'  Routine to call WS VBA style
' ----------------------------------------------------------------------

' --- Get the user name from the env on this machine
    UserName = Left(Environ$("UserName"), 4)

' --- Show the progress in the status line
    Application.StatusBar = "Set up call"
    
' --- Set up the call URL with the Query String set to User (RoA)
    url = "http://172.16.189.20/cgi-bin/IdealQuant/UpdateIdealQty.cgi?" & UserName
    
' --- Set further web call details VB style
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
 ' --- Timeout values are in milli-seconds
    lResolve = 12000
    lConnect = 12000
    lSend = 12000
    lReceive = 12000
            
' --- Set to synchronus
    objHTTP.Open "POST", url, False
    
    Application.StatusBar = "DO the Web Call: Set up Request ..."
    objHTTP.SetRequestHeader "Content-type", "application/json"
    
    objHTTP.SetTimeouts lResolve, lConnect, lSend, lReceive
    'Json = "This is a test"
    
' --- Send the data
    Application.StatusBar = "DO the Web Call: Send Request ..."
    objHTTP.Send (Json)
    result = objHTTP.ResponseText
   
    ShowResult = MsgBox("Table Updated: " & result, vbOKOnly, "IDEAL QTY: UPDATE DONE")
End Sub





