'===========================================================================
' POS LANE CASH FLOW VALIDATION 
'===========================================================================

Dim WshShell
Set WshShell = CreateObject("WScript.Shell")

Dim excelPath
excelPath = "C:\Users\BanuchandhiranM\Git\POS_Lanes_Automation\POS_Lanes_TestData.xlsx"   '--> Change path as per local setup

Dim laneAddress
Dim userId
Dim userPassword
Dim domainName

Dim userCode

Dim securityCode

Dim departmentButton
Dim miscButton
Dim departmentArray
Dim iterator

Dim priceBelowLimit
Dim priceOverLimit_1
Dim actualPrice_1

Dim dobForTobacco
Dim tabaccoPriceOverLimit
Dim tabaccoActualPrice

Dim priceOverLimit_2
Dim actualPrice_2

Dim creditButton
Dim couponType
Dim couponName
Dim couponAmount
Dim totalButton

Dim objWord
Dim objDoc

Dim excel
Dim workbook
Dim sheet
Dim data


SystemUtil.CloseProcessByName "EXCEL.EXE"
Wait 1
'================================================
' READ TEST DATA
'================================================
Set excel = CreateObject("Excel.Application")
Set workbook = excel.Workbooks.Open(excelPath)
Set sheet = workbook.Sheets("CashFlow")

'------Create Dictionary------
Set data = CreateObject("Scripting.Dictionary")
row = 2   ' Start from row 2 (skip header)
Do While sheet.Cells(row, 1).Value <> ""
    key = Trim(sheet.Cells(row, 1).Value)
    value = sheet.Cells(row, 2).Value
    data.Add key, value
    row = row + 1
Loop

'------Assign Variables------
userCode				= data("userCode")
securityCode			= data("securityCode")
priceOverLimit_1		= data("priceOverLimit_1")
actualPrice_1			= data("actualPrice_1")
tabaccoPriceOverLimit	= data("tabaccoPriceOverLimit")
tabaccoActualPrice		= data("tabaccoActualPrice")
priceOverLimit_2		= data("priceOverLimit_2")
actualPrice_2			= data("actualPrice_2")
priceBelowLimit			= data("priceBelowLimit")
dobForTobacco			= data("dobForTobacco")
departmentButton		= data("departmentButton")
miscButton				= data("miscButton")
creditButton			= data("creditButton")
couponType			= data("couponType")
couponName			= data("couponName")
couponAmount			= data("couponAmount")
totalButton				= data("totalButton")

'------Convert Array String to Actual Array------
rawArray = data("departmentArray")
cleanStr = Replace(rawArray, "Array(", "")
cleanStr = Replace(cleanStr, ")", "")
cleanStr = Replace(cleanStr, """", "")
departmentArray = Split(cleanStr, ",")

Call ReadLoginData(workbook, laneAddress, userId, userPassword, domainName)

'===========================================================================
' MAIN FLOW
'===========================================================================
Set laneSheet = workbook.Sheets("Lanes")

lastRow = laneSheet.Cells(laneSheet.Rows.Count,1).End(-4162).Row

For r = 2 To lastRow

Reporter.ReportEvent micDone, _
	    "Lane - " & laneAddress & " - Start", _
	   "Lane - " & laneAddress

SystemUtil.CloseProcessByName "javaw.exe"
SystemUtil.CloseProcessByName "java.exe"

Wait 2

WshShell.Run "javaw -jar ""C:\Users\BanuchandhiranM\Downloads\TRCConsole.jar""", 1, False	'--> Change path as per local setup

Wait 10

AIUtil.SetContext Window("regexpwndtitle:=IBM Endpoint Manager.*", "regexpwndclass:=SunAwtFrame")

    laneAddress = Trim(laneSheet.Cells(r,1).Value)

    If laneAddress <> "" Then

        Call initializeWordDocument(laneAddress)

        Call openConnection(laneAddress)
        
        Call userAuthenticate(userId, userPassword, domainName)
        
        Call TakeScreenshotToWord ("After User Authentication")
        
        Call signOn(userCode, securityCode)
        
        Call TakeScreenshotToWord ("After Sign On")
        
        Call productCheckin(departmentArray, iterator, miscButton, departmentButton, dobForTobacco, priceBelowLimit, tabaccoPriceOverLimit, tabaccoActualPrice, priceOverLimit_1, actualPrice_1, priceOverLimit_2, actualPrice_2)

couponName = GetCouponForDepartment(departmentArray)

If couponName <> "" Then

	Call discountOnProduct(creditButton, couponType, couponName, couponAmount)

Else

    Reporter.ReportEvent micDone, "Coupon Flow", "No eligible department found. Coupon skipped."

End If

	Call productCheckout(totalButton) 
	
	Call finalProductCheckout()
	
	Call takeScreenshotToWord("End")
	
	Call closeConsole()
	
	Call finalizeWordDocument(laneAddress)
	
Reporter.ReportEvent micDone, _
    "Lane - " & laneAddress & " Completed", _
   "Lane - " & laneAddress & " completed successfully."

	
	 End If

Next

workbook.Close False
excel.Quit

Set sheet = Nothing
Set workbook = Nothing
Set excel = Nothing



'-----------------------------------------------------------------------------------------------------------------------------------

'================================================
' INITIALIZE WORD DOCUMENT
'================================================
Function initializeWordDocument(laneAddress)

    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True

    Set objDoc = objWord.Documents.Add()


    objDoc.Content.InsertAfter "POS Cash Flow Transaction - Test Evidence :"& laneAddress
    objDoc.Content.InsertParagraphAfter
    objDoc.Content.InsertAfter "Execution Date : " & Date
    objDoc.Content.InsertParagraphAfter
   objDoc.Content.InsertAfter "Execution Time : " & Hour(Now) & ":" & Minute(Now)
    objDoc.Content.InsertParagraphAfter
    objDoc.Content.InsertParagraphAfter

End Function
'================================================
' SAVE AND CLOSE WORD DOCUMENT
'================================================
Function finalizeWordDocument(laneAddress)

Dim safeLane

safeLane = Trim(laneAddress)

savePath = "C:\Users\BanuchandhiranM\Lanes\CashFlow_" & _
           safeLane & "_" & _
           Year(Now) & Month(Now) & Day(Now) & "_" & _
           Hour(Now) & Minute(Now) & Second(Now) & _
           ".docx"
    'MsgBox savePath

    objDoc.SaveAs savePath
    objDoc.Close
    objWord.Quit

    Set objDoc = Nothing
    Set objWord = Nothing
    
End Function

'================================================
' TAKE SCREENSHOT
'================================================
Function takeScreenshotToWord(stepName)
    folderPath = "C:\Users\BanuchandhiranM\Lanes\Screenshots\"  '----------->Change the file path
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder(folderPath)
    End If
    
    'Create file name
    fileName = stepName & "_" & Replace(Replace(Now, ":", "_"), "/", "_") & ".png"
    fullPath = folderPath & fileName
    'Take Screenshot
    
    Desktop.CaptureBitmap fullPath, True
    
    ' Always work with same range at end
    Set rng = objDoc.Range
    ' move to end
    rng.Collapse 0  
    'Insert Image first
    objDoc.InlineShapes.AddPicture fullPath, False, True, rng
    'Move to end again
    rng.Collapse 0
    'Insert Step name after image
    rng.InsertAfter vbNewLine & stepName & vbNewLine
    ' Optional spacing
    rng.Collapse 0
    rng.InsertAfter vbNewLine
End Function

'================================================
' OPEN CONNECTION
'================================================
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

'================================================
' LOGIN
'================================================
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

'================================================
' SIGN ON
'================================================
Function signOn(userCode, securityCode)
	WshShell.SendKeys "{ENTER}"
	wait(2)
        Set WshShell = CreateObject("WScript.Shell")
        AIUtil("button", "Sign On").Click
        WshShell.SendKeys userCode
        wait(2)
        WshShell.SendKeys "{ENTER}"
        wait(2)
        WshShell.SendKeys securityCode
        wait(2)
        WshShell.SendKeys "{ENTER}"
        wait(5)
End Function
'================================================
' PRODUCT CHECKIN
'================================================
Function productCheckin(departmentArray, iterator, miscButton, departmentButton, dobForTobacco, priceBelowLimit, tabaccoPriceOverLimit, tabaccoActualPrice, priceOverLimit_1, actualPrice_1, priceOverLimit_2, actualPrice_2)
        Set WshShell = CreateObject("WScript.Shell")
        Dim dept
        Dim dobValidated
        dobValidated = False
        For iterator = 0 To UBound(departmentArray)
                dept = Trim(departmentArray(iterator)) 
                
                Call departmentNavigation(miscButton, departmentButton)


	'------ALCOHOL------
               
               If dept = "Beer" Or dept = "Wine" Then

    AIUtil.FindTextBlock(dept).Click

    If dobValidated = False Then

        TakeScreenshotToWord("Before entering DOB for department : " & dept)
        AIUtil.FindTextBlock("Clear").Click
        wait(2)
        WshShell.SendKeys dobForTobacco
        Wait(3)

        TakeScreenshotToWord("DOB Entered for department : " & dept)

        WshShell.SendKeys "{ENTER}"
        Wait(3)

        dobValidated = True

    End If

    'Call priceBelowLimitCheck(priceBelowLimit, dept)

    'AIUtil.FindTextBlock(dept).Click
    'Call priceOverLimitCheck(priceOverLimit_1, dept)

    'AIUtil.FindTextBlock(dept).Click
   
   Call setActualPrice(actualPrice_1, dept)

    '------TOBACCO------
    
    ElseIf dept = "Tobacco" or dept = "Cigarettes" Then
                        
                        AIUtil.FindTextBlock(dept).Click
                        wait(3)
                       If dobValidated = False Then
                       AIUtil.FindTextBlock("Clear").Click
                        wait(2)
    
    TakeScreenshotToWord("Before entering DOB for department : " & dept)

    WshShell.SendKeys dobForTobacco
    Wait(3)

    TakeScreenshotToWord("DOB Entered for department : " & dept)

    WshShell.SendKeys "{ENTER}"
    Wait(3)

    dobValidated = True

End If
                        wait(3)
    
    'Call priceBelowLimitCheck(priceBelowLimit, dept)
    'AIUtil.FindTextBlock(dept).Click
    'Call priceOverLimitCheck(tabaccoPriceOverLimit, dept)
    'AIUtil.FindTextBlock(dept).Click

Call setActualPrice(tabaccoActualPrice, dept)
                
 '------MEAT / NON-FOOD------
 
 ElseIf dept = "Meat" or dept = "Non-Food" Then
                      '  AIUtil.FindTextBlock(dept).Click
                       ' wait(2)
                       ' Call priceBelowLimitCheck(priceBelowLimit, dept)
                       ' AIUtil.FindTextBlock(dept).Click
                       ' Call priceOverLimitCheck(priceOverLimit_2, dept)
                       ' AIUtil.FindTextBlock(dept).Click
                       ' Call setActualPrice(actualPrice_2, dept)

 
 '------DEFAULT------

Else
                     AIUtil.FindTextBlock(dept).Click
                     wait(2)
                     Call priceBelowLimitCheck(priceBelowLimit, dept)
                      wait(2)
                     AIUtil.FindTextBlock(dept).Click
                     Call priceOverLimitCheck(priceOverLimit_1, dept)
                      wait(2)
                        AIUtil.FindTextBlock(dept).Click
                        Call setActualPrice(actualPrice_1, dept)
                End If
        Next
End Function

'================================================
' PRICE BELOW LIMIT CHECK
'================================================
Function priceBelowLimitCheck(priceBelowLimit, dept)
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys priceBelowLimit
        wait(5)
        WshShell.SendKeys "{ENTER}"
        TakeScreenshotToWord ("Price Below Limit Check for department " & dept)
        wait(5)
        WshShell.SendKeys "{ENTER}"
        wait(5)
        AIUtil.FindTextBlock("[No").Click
        Wait(2)
        Call departmentNavigation(miscButton, departmentButton)
End Function

'================================================
' PRICE OVER LIMIT CHECK
'================================================
Function priceOverLimitCheck(priceOverLimit, dept)
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys priceOverLimit
        wait(5)
        WshShell.SendKeys "{ENTER}"
        TakeScreenshotToWord ("Price Over Limit Check for department " & dept)
        wait(5)
        WshShell.SendKeys "{ENTER}"
        wait(5)
        AIUtil.FindTextBlock("[No").Click
        wait(2)
        Call departmentNavigation(miscButton, departmentButton)
End Function

'================================================
' SET ACTUAL PRICE
'================================================
Function setActualPrice(actualPrice, dept)
        Set WshShell = CreateObject("WScript.Shell")
        WshShell.SendKeys actualPrice
        wait(5)
        WshShell.SendKeys "{ENTER}"
        TakeScreenshotToWord ("Set Actual Price for department " & dept)
        wait(5)
        WshShell.SendKeys "{ENTER}"
        wait(5)
End Function
'================================================
' DEPARTMENT NAVIGATION
'================================================
Function departmentNavigation(miscButton, departmentButton)
        AIUtil("button", miscButton).Click
        TakeScreenshotToWord ("After Clicking Misc Button")
        wait(2)
        AIUtil.FindTextBlock(departmentButton).Click
        TakeScreenshotToWord ("After Clicking Department Button")
        wait(2)
End Function
'================================================
' DISCOUNT ON PRODUCT
'================================================
Function  discountOnProduct(creditButton, couponType, couponName,couponAmount)
        Set WshShell = CreateObject("WScript.Shell")
        TakeScreenshotToWord ("Before Applying Discount")
        AIUtil("button", creditButton).Click
        wait(3)
        TakeScreenshotToWord ("After Clicking Credit Button")
        AIUtil("button", couponType).Click
        wait(3)
        AIUtil.FindTextBlock(couponName, micFromTop, 1).Click
        wait(3)
        WshShell.SendKeys couponAmount
        TakeScreenshotToWord ("After Entering Coupon Amount")
        wait(5)
        WshShell.SendKeys "{ENTER}"
        wait(5)
        TakeScreenshotToWord ("After Applying Discount")
End Function
'================================================
' PRODUCT CHECKOUT
'================================================
Function productCheckout(totalButton)
        Set WshShell = CreateObject("WScript.Shell")
        AIUtil("button",totalButton).Click
        TakeScreenshotToWord ("After Entering Total Button")
        wait(2)
        WshShell.SendKeys "{ESC}"
        wait(5)
        AIUtil.FindTextBlock("[Yes").Click
        TakeScreenshotToWord ("After Accepting The Pop up")
        wait(2)
        AIUtil.FindTextBlock("Cash").Click
        TakeScreenshotToWord ("After Entering Cash Button")
        wait(2)
End Function
'================================================
' FINAL PRODUCT CHECKOUT
'================================================
Function finalProductCheckout()
        AIUtil.FindText("Exact").Click
        wait(5)
        TakeScreenshotToWord ("After Entering Exact Button")
        wait(10)
        TakeScreenshotToWord ("Before Sign Off")
        AIUtil("button", "Sign off").Click
        TakeScreenshotToWord ("After Sign Off")
wait (3)
End Function
'================================================
' CLOSE REMOTE CONSOLE
'================================================
Function closeConsole()


 On Error Resume Next

    Window("regexpwndtitle:=IBM Endpoint Manager.*").Activate
    Wait 2

     AIUtil("close").Click
     Wait 2
     AIUtil("button", "Yes").Click
     Wait 2

    If Window("regexpwndtitle:=IBM Endpoint Manager.*").Exist(3) Then

        SystemUtil.CloseProcessByName "javaw.exe"
        SystemUtil.CloseProcessByName "java.exe"

    End If

    On Error GoTo 0

End Function

Function GetCouponForDepartment(departmentArray)

    Dim i, dept

    GetCouponForDepartment = ""

    For i = 0 To UBound(departmentArray)

        dept = UCase(Trim(departmentArray(i)))

        Select Case dept

            Case "GROCERY"
                GetCouponForDepartment = "SC-Grocery"
                Exit Function

            Case "DELI"
                GetCouponForDepartment = "SC-Deli"
                Exit Function

            Case "MEAT"
                GetCouponForDepartment = "SC-Meat"
                Exit Function

            Case "NON-FOOD"
                GetCouponForDepartment = "SC-Non Food"
                Exit Function

            Case "PRODUCE"
                GetCouponForDepartment = "SC-Produce"
                Exit Function

            Case "HEALTH & BEAUTY"
                GetCouponForDepartment = "SC-Health & Beauty"
                Exit Function

            Case "PREPARED FOOD"
                GetCouponForDepartment = "SC-Prepared Food"
                Exit Function

            Case "FROZEN FOOD", "DAIRY", "FROZEN FOOD / DAIRY"
                GetCouponForDepartment = "SC-Frozen Food / Dairy"
                Exit Function

        End Select

    Next

End Function
Function ReadLoginData(workbook, ByRef laneAddress, ByRef userId, ByRef userPassword, ByRef domainName)
	
	    
	    Dim objLoginSheet
	    Dim lastRow
	    Dim i
	    Dim paramName
	    Dim paramValue
	
	    
	    Set objLoginSheet    = workbook.Sheets("LoginData")
	
	    lastRow = objLoginSheet.UsedRange.Rows.Count
	
	      For i = 2 To lastRow
	
	        paramName  = Trim(objLoginSheet.Cells(i, 1).Value)
	        paramValue = Trim(objLoginSheet.Cells(i, 2).Value)
	
	        Select Case LCase(paramName)
	           ' Case "laneaddress"  : laneAddress  = paramValue
	            Case "userid"       		: userId        = paramValue
	            Case "userpassword" : userPassword  = paramValue
	            Case "domainname"   : domainName    = paramValue
	        End Select
	
	    Next
	
	    	
	    Reporter.ReportEvent micDone, "LoginData", _
	        "Lane: " & laneAddress & " | User: " & userId & " | Domain: " & domainName
	
	End Function'
