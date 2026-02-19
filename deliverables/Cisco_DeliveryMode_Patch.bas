Attribute VB_Name = "Cisco_DeliveryMode_Patch"
Option Explicit

' =========================
' Cisco delivery mode patch
' =========================
' 1) Add a dropdown on MACROS sheet with values: CIP,DDP,CIP&DDP
' 2) Link that dropdown to cell MACROS!B10 (or change DELIVERY_MODE_CELL)
' 3) Replace original GO_CISCO with this one (or bind GO button to GO_CISCO)

Private Const DELIVERY_MODE_CELL As String = "B10"
Private Const COMMENT_CELL As String = "B9"

Private Function GetDeliveryMode() As String
    Dim modeText As String
    modeText = UCase$(Trim$(CStr(mSh.Range(DELIVERY_MODE_CELL).Value)))

    Select Case modeText
        Case "CIP", "DDP", "CIP&DDP"
            GetDeliveryMode = modeText
        Case Else
            GetDeliveryMode = "CIP"   ' default mode
    End Select
End Function

Private Function BuildComment(ByVal baseComment As String, ByVal modeMarker As String) As String
    Dim cleanBase As String
    cleanBase = Trim$(baseComment)

    If cleanBase = "" Then
        BuildComment = modeMarker
    Else
        BuildComment = cleanBase & "_" & modeMarker
    End If
End Function

Private Function PickSourceFilePath(ByVal initialFolder As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = initialFolder
        .AllowMultiSelect = False
        .ButtonName = "OK"
        .Filters.Clear
        .Filters.Add Description:="Файлы Microsoft Excel", Extensions:="*.xls; *.xlsx; *.xlsm"
        If .Show = 0 Then
            PickSourceFilePath = ""
        Else
            PickSourceFilePath = .SelectedItems(1)
        End If
    End With
End Function

Private Sub OpenSourceBook(ByVal fullPath As String)
    If sB Is Nothing Then
        Set sB = Workbooks.Open(fullPath)
    Else
        On Error Resume Next
        If sB.Name <> "" Then
            sB.Close False
        End If
        On Error GoTo 0
        Set sB = Workbooks.Open(fullPath)
    End If
End Sub

Private Sub RunCiscoForCurrentCountry(ByVal mode As String)
    Select Case mSh.[B3]
        Case "KZ"
            Call Macros_Cisco_KZ

        Case "UZ", "TM", "AZ", "TJ"
            If mode = "DDP" Then
                Call Macros_Cisco_shUZddp
            Else
                Call Macros_Cisco_UZnew
            End If

        Case "GR", "AM"
            Call Macros_Cisco_GR
    End Select
End Sub

Public Sub GO_CISCO()
    Dim thisB As Workbook
    Dim sourcePath As String
    Dim selectedMode As String
    Dim originalComment As String

    Set thisB = ActiveWorkbook
    selectedMode = GetDeliveryMode()
    originalComment = CStr(mSh.Range(COMMENT_CELL).Value)

    sourcePath = PickSourceFilePath(thisB.Path)
    If sourcePath = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo FailHandler

    Select Case selectedMode
        Case "CIP"
            mSh.Range(COMMENT_CELL).Value = originalComment
            OpenSourceBook sourcePath
            RunCiscoForCurrentCountry "CIP"

        Case "DDP"
            mSh.Range(COMMENT_CELL).Value = BuildComment(originalComment, "DDP")
            OpenSourceBook sourcePath
            RunCiscoForCurrentCountry "DDP"

        Case "CIP&DDP"
            ' CIP file with combined marker (for report traceability)
            mSh.Range(COMMENT_CELL).Value = BuildComment(originalComment, "CIP&DDP")
            OpenSourceBook sourcePath
            RunCiscoForCurrentCountry "CIP"

            ' DDP file from the same source file, with _DDP suffix + combined marker
            mSh.Range(COMMENT_CELL).Value = BuildComment(originalComment, "CIP&DDP_DDP")
            OpenSourceBook sourcePath
            RunCiscoForCurrentCountry "DDP"
    End Select

CleanExit:
    mSh.Range(COMMENT_CELL).Value = originalComment
    On Error Resume Next
    If Not sB Is Nothing Then sB.Close False
    Set sB = Nothing
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

FailHandler:
    MsgBox "Ошибка выполнения GO_CISCO: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
