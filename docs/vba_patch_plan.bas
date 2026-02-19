' Draft VBA patch (reference) for delivery mode dropdown support
Option Explicit

Private Function GetDeliveryMode() As String
    Dim mode As String
    mode = Trim$(CStr(mSh.Range("B10").Value)) ' example cell linked to dropdown

    Select Case UCase$(mode)
        Case "CIP", "DDP", "CIP&DDP"
            GetDeliveryMode = UCase$(mode)
        Case Else
            GetDeliveryMode = "CIP"
    End Select
End Function

Sub GO_DISPATCH_BY_MODE()
    Dim mode As String
    mode = GetDeliveryMode()

    ' source file should be selected once before this point
    Select Case mode
        Case "CIP"
            RunCIP
        Case "DDP"
            RunDDP
        Case "CIP&DDP"
            RunCIP
            RunDDP
            SaveReportMarker "CIP&DDP"
    End Select
End Sub

Private Function BuildSpecNameByMode(ByVal baseSpecName As String, ByVal mode As String) As String
    If mode = "DDP" Then
        BuildSpecNameByMode = baseSpecName & "_DDP"
    Else
        BuildSpecNameByMode = baseSpecName
    End If
End Function

Private Sub SaveReportMarker(ByVal marker As String)
    ' add marker into report row/column to keep audit of combined runs
    ' minimal required behavior: marker must appear in report line for the run
End Sub
