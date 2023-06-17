WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set "john" @@ hightlight id_;_1884151152_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "6481e20185abbebf6d8c" @@ hightlight id_;_1884147696_;_script infofile_;_ZIP::ssf4.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_1884147984_;_script infofile_;_ZIP::ssf3.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select "Frankfurt" @@ hightlight id_;_1884161616_;_script infofile_;_ZIP::ssf6.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select "Denver" @@ hightlight id_;_1884134832_;_script infofile_;_ZIP::ssf8.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("datePicker").SetDate "10-Jun-2023" @@ hightlight id_;_1962369152_;_script infofile_;_ZIP::ssf9.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select "Business" @@ hightlight id_;_1962379088_;_script infofile_;_ZIP::ssf11.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select "2" @@ hightlight id_;_1962379472_;_script infofile_;_ZIP::ssf13.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1962380960_;_script infofile_;_ZIP::ssf14.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,1 @@ hightlight id_;_1953984440_;_script infofile_;_ZIP::ssf15.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_1970170232_;_script infofile_;_ZIP::ssf16.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set "john" @@ hightlight id_;_1954191888_;_script infofile_;_ZIP::ssf19.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_1954199040_;_script infofile_;_ZIP::ssf20.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfProgressBar("progBar").WaitProperty "value", 100, 10000
WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 140 completed").Check CheckPoint("Order 140 completed") @@ hightlight id_;_1950190592_;_script infofile_;_ZIP::ssf21.xml_;_
'客製化檢核點
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTicketsCombo").Output CheckPoint("機票張數") @@ hightlight id_;_1944073968_;_script infofile_;_ZIP::ssf21.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("$207.60").Output CheckPoint("機票金額") @@ hightlight id_;_1950211600_;_script infofile_;_ZIP::ssf21.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfObject("$415.20").Output CheckPoint("機票總額")
'設定變數
strTicketNumber = DataTable("機票張數_out")
strTicketPrice = DataTable("機票金額_out")
'使用CStr函數將checkTotalPrice變數型態轉為String
checkTotalPrice = CStr(DataTable("機票張數_out")*DataTable("機票金額_out"))
'使用Replace函數將strTotalPrice變數做字串處理
strTotalPrice = Replace(Replace(DataTable("機票總額_out"),"0",""),"$","")

If checkTotalPrice = strTotalPrice Then
	Reporter.ReportEvent micPass, "訂單金額檢核","正確。機票張數:" & strTicketNumber & "機票金額:" & strTicketPrice & "機票總額:" & checkTotalPrice
Else	
	Reporter.ReportEvent micFail, "訂單金額檢核","錯誤。機票張數:" & strTicketNumber & "機票金額:" & strTicketPrice & "機票總額:" & checkTotalPrice
End If
