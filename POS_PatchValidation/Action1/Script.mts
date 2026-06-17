'===========================================================================
' POS LANE HOTFIX / PATCH VALIDATION 
'===========================================================================

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

Dim excelPath
excelPath = "C:\Users\BanuchandhiranM\Lanes\POS_Lanes_TestData.xlsx"			'---> change the path as per the local 

Dim laneAddress, userId, userPassword, domainName
Call ReadLoginData(excelPath, laneAddress, userId, userPassword, domainName)

' CLEANUP & LAUNCH
SystemUtil.CloseProcessByName "javaw.exe"
SystemUtil.CloseProcessByName "java.exe"
SystemUtil.CloseProcessByName "excel.exe"
Wait 2

WshShell.Run "javaw -jar ""C:\Users\BanuchandhiranM\Downloads\TRCConsole.jar""", 1, False  '---> change the path as per the local 
Wait 10

AIUtil.SetContext Window("regexpwndtitle:=IBM Endpoint Manager.*", "regexpwndclass:=SunAwtFrame")


'===========================================================================
' MAIN FLOW
'===========================================================================

Call openConnection(laneAddress)
Call userAuthenticate(userId, userPassword, domainName)
Call runHotFixPowerShell()
Call pullClipboardToLocal()
Call validateHotFixes(excelPath)
Call closeConsole()

'===========================================================================
' FUNCTION: READ LOGIN DATA FROM EXCEL 
'===========================================================================
Function ReadLoginData(excelPath, ByRef laneAddress, ByRef userId, ByRef userPassword, ByRef domainName)

    Dim objExcel, objWorkbook, objSheet
    Dim lastRow, i, paramName, paramValue

    Set objExcel     = CreateObject("Excel.Application")
    objExcel.Visible = False

    Set objWorkbook = objExcel.Workbooks.Open(excelPath)
    Set objSheet    = objWorkbook.Sheets("LoginData")

    lastRow = objSheet.UsedRange.Rows.Count

    ' Column A = ParameterName | Column B = Value
    For i = 2 To lastRow
        paramName  = Trim(objSheet.Cells(i, 1).Value)
        paramValue = Trim(objSheet.Cells(i, 2).Value)

        Select Case LCase(paramName)
            Case "laneaddress"  : laneAddress  = paramValue
            Case "userid"       : userId        = paramValue
            Case "userpassword" : userPassword  = paramValue
            Case "domainname"   : domainName    = paramValue
        End Select
    Next

    objWorkbook.Close False
    objExcel.Quit

    Reporter.ReportEvent micDone, "LoginData", _
        "Lane: " & laneAddress & " | User: " & userId & " | Domain: " & domainName

End Function


'===========================================================================
' FUNCTION: OPEN CONNECTION
'===========================================================================
Function openConnection(laneAddress)

    AIUtil("down_triangle", micAnyText, micFromTop, 1).Click
    Wait 2
    AIUtil("down_triangle", micAnyText, micFromTop, 1).Click

    AIUtil("text_box", "Address").RightClick
    Wait 2
    WshShell.SendKeys "{DOWN}"
    WshShell.SendKeys "{DOWN}"
    WshShell.SendKeys "{ENTER}"
    Wait 2

    AIUtil("text_box", "Address").Type laneAddress
    AIUtil("button", "Active").Click
    Wait 5

    Reporter.ReportEvent micDone, "Connection", "Connected to lane: " & laneAddress

End Function


'===========================================================================
' FUNCTION: LOGIN
'===========================================================================
Function userAuthenticate(userId, userPassword, domainName)

    AIUtil("combobox", "User ID").Type userId
    AIUtil("text_box", "Password").Type userPassword
    AIUtil("combobox", "Domain").Click
    Wait 2

    Set combo = AIUtil("combobox", "Domain")
    If Trim(combo.GetValue) <> Trim(domainName) Then
        combo.Click
        Wait 1
    End If

    AIUtil("button", "Login").Click
    Wait 8

    Reporter.ReportEvent micDone, "Login", "Logged in: " & userId

End Function


'===========================================================================
' FUNCTION: RUN POWERSHELL -> GET-HOTFIX -> SAVE TO NOTEPAD ON REMOTE
'===========================================================================
Function runHotFixPowerShell()

    ' Activate remote viewer
    Window("regexpwndtitle:=IBM Endpoint Manager.*").Activate
    Wait 2
    Window("regexpwndtitle:=IBM Endpoint Manager.*").Click 332, 10
    Wait 3

    ' Open Task Manager
    WshShell.SendKeys "{ENTER}"
    Wait 1
    WshShell.SendKeys "{DOWN}"
    Wait 1
    WshShell.SendKeys "{ENTER}"
    Wait 5

    ' Task Manager -> File -> Run New Task
    AIUtil.FindTextBlock("Task Manager").Click
    Wait 4
    AIUtil.FindTextBlock("File").Click
    Wait 2
    WshShell.SendKeys "{DOWN}"
    Wait 1
    WshShell.SendKeys "{ENTER}"
    Wait 3

    ' Launch PowerShell
    WshShell.SendKeys "powershell"
    WshShell.SendKeys "{ENTER}"
    Wait 7

    ' Run Get-HotFix and save output to temp file on remote machine
    Dim psCmd
    psCmd = "Get-HotFix | Select-Object HotFixID, InstalledOn | Format-Table -AutoSize | Out-String -Width 4096 | Out-File 'C:\temp\hotfix_id.txt' -Encoding ascii"

    WshShell.SendKeys psCmd
    WshShell.SendKeys "{ENTER}"
    Wait 12

    Reporter.ReportEvent micDone, "PowerShell", "Get-HotFix executed. hotfix_id.txt saved on remote."

    ' Open hotfix.txt in Notepad on remote
    WshShell.SendKeys "C:\temp\hotfix_id.txt"    
    WshShell.SendKeys "{ENTER}"
    Wait 5

    Reporter.ReportEvent micDone, "Notepad", "hotfix.txt opened in Notepad on remote."

    '-----------------------------------------------------------------------
    ' SCREENSHOT: Capture Notepad showing HotFixIDs on remote screen
    '-----------------------------------------------------------------------
    Wait 3

    Dim screenshotPath
    screenshotPath = "C:\Temp\HotFix_Notepad_Screenshot.png"   '---> change the path as per the local 

    Dim fsoSnap
    Set fsoSnap = CreateObject("Scripting.FileSystemObject")
    If Not fsoSnap.FolderExists("C:\Temp") Then
        fsoSnap.CreateFolder "C:\Temp"
    End If
    Set fsoSnap = Nothing

    ' Capture the full remote viewer window as screenshot
    Window("regexpwndtitle:=IBM Endpoint Manager.*").CaptureBitmap screenshotPath, True

    Reporter.ReportEvent micDone, "Screenshot Captured", _
        "Notepad with HotFixIDs captured to: " & screenshotPath

    ' Attach screenshot to UFT HTML report
    Reporter.ReportEvent micDone, "HotFix Notepad View", _
        "<img src='" & screenshotPath & "' width='800'/>"

End Function

'===========================================================================
' FUNCTION:  PULL CLIPBOARD TO LOCAL
'===========================================================================
Function pullClipboardToLocal()

    ' Click inside Notepad content area on remote screen
    Window("regexpwndtitle:=IBM Endpoint Manager.*").Click 400, 300
    Wait 2

    ' Select All text in remote Notepad
    WshShell.SendKeys "^a"
    Wait 2

    ' Copy to remote clipboard
    WshShell.SendKeys "^c"
    Wait 5

    ' Click the Remote Control Viewer toolbar to trigger clipboard sync
 Window("IBM Endpoint Manager for Remote Control  (US90L01-09956)").Click 527,11
    Wait 10

    ' Navigate down twice and Enter to confirm clipboard transfer
    WshShell.SendKeys "{DOWN}"
    Wait 1
    WshShell.SendKeys "{DOWN}"
    Wait 1
    WshShell.SendKeys "{ENTER}"
    Wait 5

    Reporter.ReportEvent micDone, "Clipboard Pull", "Remote clipboard pulled to local machine."

    ' Read from local clipboard and save to file
    Dim html, clipboardText
    Set html      = CreateObject("htmlfile")
    clipboardText = html.ParentWindow.ClipboardData.GetData("Text")

    If Trim(clipboardText) = "" Then
        Reporter.ReportEvent micFail, "Clipboard Empty", "No text received from remote clipboard."
        Exit Function
    End If

    ' Save to local temp file
    Dim fso, txtFile
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists("C:\Temp") Then
        fso.CreateFolder "C:\Temp"
    End If

    Set txtFile = fso.CreateTextFile("C:\Temp\hotfix_local.txt", True)
    txtFile.Write clipboardText
    txtFile.Close

    Set fso  = Nothing
    Set html = Nothing

    Reporter.ReportEvent micDone, "Clipboard Saved", "HotFix list saved to C:\Temp\hotfix_local.txt"

End Function


'===========================================================================
' FUNCTION: VALIDATE HOTFIX IDs AGAINST EXCEL
'===========================================================================
Function validateHotFixes(excelPath)

    ' Read local hotfix file
    Dim fso, txtFile, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists("C:\Temp\hotfix_local.txt") Then
        Reporter.ReportEvent micFail, "File Not Found", "C:\Temp\hotfix_local.txt does not exist."
        Exit Function
    End If

    Set txtFile  = fso.OpenTextFile("C:\Temp\hotfix_local.txt", 1)
    fileContent  = txtFile.ReadAll
    txtFile.Close
    Set fso = Nothing

    If Trim(fileContent) = "" Then
        Reporter.ReportEvent micFail, "File Empty", "hotfix_local.txt is empty."
        Exit Function
    End If

    Reporter.ReportEvent micDone, "File Read", "hotfix_local.txt loaded successfully."

    ' -----------------------------------------------------------------------
    ' STEP 1: Kill any lingering Excel process before opening
    ' -----------------------------------------------------------------------
    SystemUtil.CloseProcessByName "excel.exe"
    Wait 3

    ' -----------------------------------------------------------------------
    ' STEP 2: Open Excel with Visible = True to avoid silent save failures
    ' -----------------------------------------------------------------------
    Dim objExcel, objWorkbook, objSheet
    Set objExcel            = CreateObject("Excel.Application")
    objExcel.Visible        = False    ' Visible so save is not blocked silently
    objExcel.DisplayAlerts  = False

    Set objWorkbook = objExcel.Workbooks.Open(excelPath)
    Set objSheet    = objWorkbook.Sheets("PatchValidation")

    Reporter.ReportEvent micDone, "Excel Opened", "PatchValidation sheet loaded."

    ' -----------------------------------------------------------------------
    ' STEP 3: Find last used row in Column A (HotfixID)
    ' -----------------------------------------------------------------------
    Dim lastRow
    lastRow = 1
    Do While Trim(objSheet.Cells(lastRow + 1, 1).Value) <> ""
        lastRow = lastRow + 1
    Loop

    Reporter.ReportEvent micDone, "Row Count", "Total data rows found: " & (lastRow - 1)

    ' -----------------------------------------------------------------------
    ' STEP 4: Loop and validate each HotFix ID
    ' -----------------------------------------------------------------------
    Dim i, hotFixID, installedOn, statusValue

    For i = 2 To lastRow

        hotFixID    = Trim(objSheet.Cells(i, 1).Value)   ' Column A : HotfixID
        installedOn = Trim(objSheet.Cells(i, 2).Value)   ' Column B : InstalledOn (reference only)
        statusValue = "FAIL"

        If hotFixID <> "" Then

            If InStr(LCase(fileContent), LCase(hotFixID)) > 0 Then
                statusValue = "PASS"
                Reporter.ReportEvent micPass, "PASS | " & hotFixID, _
                    "HotFix " & hotFixID & " found on lane. InstalledOn: " & installedOn
            Else
                statusValue = "FAIL"
                Reporter.ReportEvent micFail, "FAIL | " & hotFixID, _
                    "HotFix " & hotFixID & " NOT found on lane. InstalledOn: " & installedOn
            End If

            ' Write PASS/FAIL to Column C (Status)
            objSheet.Cells(i, 3).Value = statusValue

            ' Force cell update immediately after each write
            objExcel.Calculate

        End If

    Next

    ' -----------------------------------------------------------------------
    ' STEP 5: Save using SaveAs to same path (avoids OneDrive/lock issues)
    ' xlOpenXMLWorkbook = 51 (saves as .xlsx without prompt)
    ' -----------------------------------------------------------------------
    On Error Resume Next
    objWorkbook.SaveAs excelPath, 51
    If Err.Number <> 0 Then
        Reporter.ReportEvent micFail, "Save Failed (SaveAs)", _
            "SaveAs error: " & Err.Description & ". Trying Save..."
        Err.Clear
        objWorkbook.Save
        If Err.Number <> 0 Then
            Reporter.ReportEvent micFail, "Save Failed", _
                "Both Save attempts failed: " & Err.Description
        End If
    Else
        Reporter.ReportEvent micDone, "Excel Saved", "Results written to PatchValidation sheet successfully."
    End If
    On Error GoTo 0

    ' -----------------------------------------------------------------------
    ' STEP 6: Close Excel cleanly
    ' -----------------------------------------------------------------------
    objWorkbook.Close False   ' False = do NOT prompt save again
    objExcel.Quit
    Wait 2

    Set objSheet    = Nothing
    Set objWorkbook = Nothing
    Set objExcel    = Nothing

    Reporter.ReportEvent micDone, "Patch Validation Complete", _
        "HotFix validation done. PASS/FAIL written to Column C of PatchValidation sheet."

End Function

'===========================================================================
' FUNCTION: CLOSE REMOTE CONSOLE
'===========================================================================
Function closeConsole()

    On Error Resume Next

    Window("regexpwndtitle:=IBM Endpoint Manager.*").Activate
    Wait 2

    AIUtil("close", micAnyText, micFromBottom, 1).Click
    Wait 2

    AIUtil("close", micAnyText, micFromBottom, 1).Click
    Wait 2

    AIUtil("close", micAnyText, micFromBottom, 1).Click
    Wait 2

    If Window("regexpwndtitle:=IBM Endpoint Manager.*").Exist(3) Then

        SystemUtil.CloseProcessByName "javaw.exe"
        SystemUtil.CloseProcessByName "java.exe"

    End If

    On Error GoTo 0

End Function

'==================================================================================================== @@ hightlight id_;_1248956_;_script infofile_;_ZIP::ssf13.xml_;_
