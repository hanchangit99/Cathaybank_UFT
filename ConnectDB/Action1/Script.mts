'Create ADODB connection object
Set objConnection = CreateObject("ADODB.Connection")
'Create Recordsetobject
Set objRecordSet = CreateObject("ADODB.Recordset")
'Connect to DB using provider and server
objConnection.open "provider=sqloledb;Server=nimbusclient\SQLEXPRESS;User Id=admin; Password=P@ssw0rd; Database=testdb;Trusted_Connection=Yes"
'Write the SQL Query
sqlQuery="Select TOP (1) [CustomerName],[FlightNumber],[NumberOfTickets],[Class],[DepartureDate],[DepartureCity],[ArrivalCity],[Price] from [testdb].[dbo].[Input$]"
'Execute the query
objRecordSet.open sqlQuery, objConnection
'Display output
value = objRecordSet.fields.Item(0)
msgbox Value
objRecordSet.Close
objConnection.Close
Set objConnection = Nothing
Set objRecordSet = Nothing

strSQL = "provider=sqloledb;Server=nimbusclient\SQLEXPRESS;User Id=admin; Password=P@ssw0rd; Database=testdb;Trusted_Connection=Yes"
sheetname = "dbo.input$"
ExecuteSQL strSQL,sheetname
wait 1

Sub ExecuteSQL(strSQL, sheetname)
 Dim strConn, objConn
 Set objConn = CreateObject("ADODB.Connection")
 On error resume next
 
 ' 設定DB連線字串
 strConn = GetMPBConnStr
 
 objConn.Open strConn
 If Err.Number <> 0 Then
  Reporter.ReportEvent micWarning, "資料庫連線失敗", "錯誤代碼：" & Err.Number & vbcrlf & vbcrlf & _
                "錯誤描述：" & Err.Description
  Err.Clear
 End If
 
 Set objConn1 = CreateObject("ADODB.Recordset")
 objConn1.Open strSQL, strConn
 If Err.Number <> 0 Then
  Reporter.ReportEvent micWarning, "SQL 指令執行失敗", "錯誤代碼：" & Err.Number & vbcrlf & vbcrlf & _
                   "錯誤描述：" & Err.Description
  Err.Clear
 Else 
  
  If sheetname <> "" Then
   For i = 0  to objConn1.Fields.Count              ' sql select 欄位的數量
    If CheckSheetExist(sheetname) Then
     DataTable.GetSheet(sheetname).AddParameter objConn1.Fields(i).name,""
    Else
     DataTable.AddSheet(sheetname).AddParameter objConn1.Fields(i).name,""  '增加DataTable worksheet 和 欄位name
    End If
   Next
   
   intRecords = 0  '第幾列變數
   While not objConn1.EOF    '讀檔宜到最後一筆
    intRecords = intRecords +1
    For i = 0  to objConn1.Fields.Count
     DataTable.GetSheet(sheetname).SetCurrentRow intRecords  '移到datatable worksheet  第X列
     DataTable(objConn1.Fields(i).name,sheetname) =  objConn1.Fields(i).value  '將返回值寫入DataTable worksheet 欄位中
    Next
    objConn1.MoveNext '往下讀一筆
   Wend
  End If
           
 End If
 Print Environment("ActionName") & " > " & "Function Call:ExecuteMPBSQL"
 Print "strSQL: " & strSQL
 
 objConn.Close
 objConn1.Close
 
 'Set objRecordSet = Nothing
 Set objConn = Nothing
 Set objConn1 = Nothing
 On error goto 0
End Sub
