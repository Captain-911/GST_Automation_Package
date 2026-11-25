Attribute VB_Name = "GSTAutomation"
Option Explicit

' =====================================================================================
' GST Returns & Notices — Bulk Download Automation
' -------------------------------------------------------------------------------------
' Import this module: Alt+F11 > right-click VBAProject > Import File > this .bas file.
' Then assign DownloadGSTR1 / DownloadGSTR3B / DownloadBoth / DownloadNotices to the
' four Form Control buttons on the "Clients" sheet. See "Setup & Instructions" sheet.
'
' HOW IT WORKS
'   For each selected row on "Clients": checks Consent = Y, writes a small job file,
'   shells out to gst_portal_bot.py (WAITS for it to finish), reads back a result file,
'   writes Status/Timestamp/Files back into the row, and appends to "Run Log".
'   One browser session / one CAPTCHA per row, run strictly one row at a time.
' =====================================================================================

Private Const COL_SRNO As Long = 1
Private Const COL_NAME As Long = 2
Private Const COL_GSTIN As Long = 3
Private Const COL_LOGIN As Long = 4
Private Const COL_PASSWORD As Long = 5
Private Const COL_FY As Long = 6
Private Const COL_CONSENT As Long = 7
Private Const COL_CONSENT_DATE As Long = 8
Private Const COL_DL_FOLDER As Long = 9
Private Const COL_LAST_ACTION As Long = 10
Private Const COL_LAST_RUN As Long = 11
Private Const COL_STATUS As Long = 12
Private Const COL_FILES As Long = 13
Private Const HEADER_ROW As Long = 6

Public Sub DownloadGSTR1()
    RunAutomationForSelection "GSTR1"
End Sub

Public Sub DownloadGSTR3B()
    RunAutomationForSelection "GSTR3B"
End Sub

Public Sub DownloadBoth()
    RunAutomationForSelection "BOTH"
End Sub

Public Sub DownloadNotices()
    RunAutomationForSelection "NOTICES"
End Sub

' -------------------------------------------------------------------------------------
' Core driver: loops every data row intersected by the current selection
' -------------------------------------------------------------------------------------
Private Sub RunAutomationForSelection(ByVal action As String)
    Dim ws As Worksheet, wsSettings As Worksheet, wsLog As Worksheet
    Dim rngRow As Range, cell As Range
    Dim successCount As Long, failCount As Long, blockedCount As Long
    Dim processedRows As Object
    Set processedRows = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Sheets("Clients")
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    Set wsLog = ThisWorkbook.Sheets("Run Log")

    If ws.Name <> ActiveSheet.Name Then
        MsgBox "Please select client row(s) on the 'Clients' sheet first, then click this button again.", vbExclamation
        Exit Sub
    End If

    If Selection.Row < HEADER_ROW + 1 Then
        MsgBox "Please click on a client data row (row " & (HEADER_ROW + 1) & " or below) before clicking this button.", vbExclamation
        Exit Sub
    End If

    Dim pythonExe As String, scriptPath As String, jobFolder As String, defaultDl As String
    pythonExe = Trim(wsSettings.Range("B4").Value)
    scriptPath = Trim(wsSettings.Range("B5").Value)
    jobFolder = Trim(wsSettings.Range("B6").Value)
    defaultDl = Trim(wsSettings.Range("B7").Value)

    If pythonExe = "" Or scriptPath = "" Or jobFolder = "" Then
        MsgBox "Please fill in the 'Settings' sheet (Python path, script path, job folder) before running.", vbCritical
        Exit Sub
    End If
    If Right(jobFolder, 1) <> "\" Then jobFolder = jobFolder & "\"
    If Dir(jobFolder, vbDirectory) = "" Then MkDir jobFolder

    For Each cell In Selection.Rows
        Dim r As Long
        r = cell.Row
        If r >= HEADER_ROW + 1 And Not processedRows.Exists(r) Then
            processedRows.Add r, True
            Dim partyName As String, gstin As String, loginId As String, pwd As String
            Dim fy As String, consent As String, dlFolder As String

            partyName = Trim(ws.Cells(r, COL_NAME).Value)
            gstin = Trim(ws.Cells(r, COL_GSTIN).Value)
            loginId = Trim(ws.Cells(r, COL_LOGIN).Value)
            pwd = Trim(ws.Cells(r, COL_PASSWORD).Value)
            fy = Trim(ws.Cells(r, COL_FY).Value)
            consent = UCase(Trim(ws.Cells(r, COL_CONSENT).Value))
            dlFolder = Trim(ws.Cells(r, COL_DL_FOLDER).Value)
            If dlFolder = "" Then dlFolder = defaultDl

            If gstin = "" And loginId = "" Then
                ' blank row - skip silently
            ElseIf consent <> "Y" Then
                blockedCount = blockedCount + 1
                ws.Cells(r, COL_STATUS).Value = "BLOCKED - Consent not marked Y"
                LogRun wsLog, partyName, gstin, action, fy, "BLOCKED", "Consent Obtained column is not 'Y'"
            ElseIf fy = "" Then
                blockedCount = blockedCount + 1
                ws.Cells(r, COL_STATUS).Value = "BLOCKED - Financial Year not selected"
                LogRun wsLog, partyName, gstin, action, fy, "BLOCKED", "Pick a Financial Year first"
            Else
                Dim jobFile As String, resultFile As String
                jobFile = jobFolder & "job_" & gstin & "_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
                resultFile = jobFile & ".result.txt"

                WriteJobFile jobFile, gstin, loginId, pwd, fy, action, partyName, dlFolder

                ws.Cells(r, COL_STATUS).Value = "RUNNING - a Chrome window will open, enter the CAPTCHA there..."
                Application.ScreenUpdating = True
                DoEvents

                Dim cmd As String
                cmd = """" & pythonExe & """ """ & scriptPath & """ """ & jobFile & """"
                Dim shellObj As Object
                Set shellObj = CreateObject("WScript.Shell")
                Dim rc As Long
                rc = shellObj.Run(cmd, 1, True)   ' 1 = normal window, True = wait for completion

                Dim statusTxt As String, msgTxt As String, filesTxt As String
                ReadResultFile resultFile, statusTxt, msgTxt, filesTxt

                ws.Cells(r, COL_LAST_ACTION).Value = action
                ws.Cells(r, COL_LAST_RUN).Value = Now
                ws.Cells(r, COL_STATUS).Value = statusTxt & " - " & msgTxt
                ws.Cells(r, COL_FILES).Value = filesTxt

                If UCase(statusTxt) = "SUCCESS" Then
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If

                LogRun wsLog, partyName, gstin, action, fy, statusTxt, msgTxt & " | " & filesTxt
            End If
        End If
    Next cell

    MsgBox "Run complete." & vbCrLf & _
           "Succeeded: " & successCount & vbCrLf & _
           "Failed: " & failCount & vbCrLf & _
           "Blocked (no consent / no FY): " & blockedCount, vbInformation
End Sub

' -------------------------------------------------------------------------------------
' Writes a simple KEY=VALUE job file for the Python script to read
' -------------------------------------------------------------------------------------
Private Sub WriteJobFile(ByVal path As String, ByVal gstin As String, ByVal loginId As String, _
                          ByVal pwd As String, ByVal fy As String, ByVal action As String, _
                          ByVal partyName As String, ByVal dlFolder As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "GSTIN=" & gstin
    Print #f, "LOGIN_ID=" & loginId
    Print #f, "PASSWORD=" & pwd
    Print #f, "FY=" & fy
    Print #f, "ACTION=" & action
    Print #f, "PARTY_NAME=" & partyName
    Print #f, "DOWNLOAD_FOLDER=" & dlFolder
    Close #f
End Sub

' -------------------------------------------------------------------------------------
' Reads back the result file Python writes: line1=STATUS, line2=MESSAGE, rest=file paths
' -------------------------------------------------------------------------------------
Private Sub ReadResultFile(ByVal path As String, ByRef statusOut As String, _
                            ByRef msgOut As String, ByRef filesOut As String)
    statusOut = "FAILED"
    msgOut = "No result file was written — check the Python console window for an error."
    filesOut = ""

    If Dir(path) = "" Then Exit Sub

    Dim f As Integer, lineTxt As String, lineNo As Long
    Dim filesList As String
    f = FreeFile
    Open path For Input As #f
    lineNo = 0
    Do While Not EOF(f)
        Line Input #f, lineTxt
        lineNo = lineNo + 1
        Select Case lineNo
            Case 1: statusOut = lineTxt
            Case 2: msgOut = lineTxt
            Case Else
                If filesList = "" Then
                    filesList = lineTxt
                Else
                    filesList = filesList & "; " & lineTxt
                End If
        End Select
    Loop
    Close #f
    filesOut = filesList
End Sub

Private Sub LogRun(ByVal wsLog As Worksheet, ByVal partyName As String, ByVal gstin As String, _
                    ByVal action As String, ByVal fy As String, ByVal status As String, ByVal message As String)
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 4 Then nextRow = 4
    wsLog.Cells(nextRow, 1).Value = Now
    wsLog.Cells(nextRow, 2).Value = partyName
    wsLog.Cells(nextRow, 3).Value = gstin
    wsLog.Cells(nextRow, 4).Value = action
    wsLog.Cells(nextRow, 5).Value = fy
    wsLog.Cells(nextRow, 6).Value = status
    wsLog.Cells(nextRow, 7).Value = message
End Sub
