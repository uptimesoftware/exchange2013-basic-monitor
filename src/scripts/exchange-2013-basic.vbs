On Error Resume Next
' Sameple JSON format output. 
' 
' {
'   "hostname": "%COMPUTERNAME%",
'   "plugin": "Windows Exchange 2013 Agent side rexec plugin",
'   "version": "7.3.0",
'   "currentTime": "%DATE:~-4%-%DATE:~3,2%-%DATE:~0,2%T%TIME:~0,8%",
'   "status": "200",
'   "result": {
'      "SMTPBytesSentPerSecond": "128", 
'      "SMTPBytesReceivedPerSecond": "256",
'      "SMTPMessagesSentPerSecond": "512",
'      "SMTPInboundMessagesReceivedPerSecond": "640",
'      "SMTPAverageBytesPerInboundMessage": "768",
'      "SMTPInboundConnections": "896",
'      "SMTPOutboundConnections": "1024",
'      "CurrentWebmailUsers": "5",
'      "WebmailUserLogonsPerSecond": "10",
'      "RPCAveragedLatency": "15",
'      "RPCOperationsPerSecond": "20",
'      "RPCRequests": "25",
'      .......
'      .......
'   }
' }

'On up.time Agent console :
'Command Name : exchange2013monitor
'Path to script : cmd /c "cscript "C:\Program Files (x86)\uptime software\up.time agent\scripts\exchange-2013-basic.vbs" //Nologo //T:0"

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

'There should be a way to combine formJSON, formNestedJSON, formEndLineJSON functions into one. Fix it later.
'take fieldName and fieldValue parameters and form JSON.
Private Function formJSON(fieldName, fieldValue)
  Dim lineInJSONFormat
  lineInJSONFormat = chr(34) & fieldName & chr(34) & ": " & chr(34) & fieldValue & chr(34) & ","
  'set function's return value
  formJSON = lineInJSONFormat
END Function

'take fieldName and curly bracket parameters and form nested JSON.
Private Function formNestedJSON(fieldName)
  Dim lineInJSONFormat
  lineInJSONFormat = chr(34) & fieldName & chr(34) & ": " & "{"
  'set function's return value
  formNestedJSON = lineInJSONFormat
END Function

'take fieldName and curly bracket parameters and form nested JSON.
Private Function formEndLineJSON(fieldName, fieldValue)
  Dim lineInJSONFormat
  lineInJSONFormat = chr(34) & fieldName & chr(34) & ": " & chr(34) & fieldValue & chr(34)
  'set function's return value
  formEndLineJSON = lineInJSONFormat
END Function

'Get name of host
Dim wshNetwork
Set wshNetwork = CreateObject("WScript.Network")
strComputerName = wshNetwork.ComputerName

Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\CIMV2")

'Outputing JSON begins
WScript.Echo "{"
WScript.Echo formJSON("Computer", strComputerName)
WScript.Echo formJSON("plugin", "Windows Exchange 2013 Agent side rexec plugin")
WScript.Echo formJSON("version", "7.3.0")
WScript.Echo formJSON("currentTime", Now)
WScript.Echo formJSON("status", "200")
'Nested JSON begins
WScript.Echo formNestedJSON("result")

'MSExchangeTransportSMTPReceive'
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeTransportSMTPReceive_MSExchangeTransportSMTPReceive WHERE Name = ""_total""", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
  WScript.Echo formJSON("SMTPAverageBytesPerInboundMessage", objItem.AveragebytesPermessage)
  WScript.Echo formJSON("SMTPBytesReceivedPerSecond", objItem.BytesReceivedPersec)
  WScript.Echo formJSON("SMTPInboundConnections", objItem.ConnectionsCurrent)
  WScript.Echo formJSON("SMTPInboundMessagesReceivedPerSecond", objItem.MessagesReceivedPersec)
Next

'MSExchangeTransportSmtpSend'
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeTransportSmtpSend_MSExchangeTransportSmtpSend WHERE Name = ""_total""", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
  WScript.Echo formJSON("SMTPBytesSentPerSecond", objItem.BytesSentPersec)
  WScript.Echo formJSON("SMTPOutboundConnections", objItem.ConnectionsCurrent)
  WScript.Echo formJSON("SMTPMessagesSentPerSecond", objItem.MessagesSentPersec)
Next

'MSExchangeOWA
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeOWA_MSExchangeOWA", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
  WScript.Echo formJSON("CurrentWebmailUsers", objItem.CurrentUsers)
  WScript.Echo formJSON("WebmailUserLogonsPerSecond", objItem.LogonsPersec)
Next

'MSExchangeRpcClientAccess
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeRpcClientAccess_MSExchangeRpcClientAccess", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colItems
  WScript.Echo formJSON("RPCAveragedLatency", objItem.RPCAveragedLatency)
  WScript.Echo formJSON("RPCClientsBytesRead", objItem.RPCClientsBytesRead)
  WScript.Echo formJSON("RPCClientsBytesWritten", objItem.RPCClientsBytesWritten)
  WScript.Echo formJSON("RPCDispatchTaskActiveThreads", objItem.RPCdispatchtaskactivethreads)
  WScript.Echo formJSON("RPCDispatchTaskQueueLength", objItem.RPCdispatchtaskqueuelength)
  WScript.Echo formJSON("RPCOperationsPersec", objItem.RPCOperationsPersec)
  WScript.Echo formJSON("RPCRequests", objItem.RPCRequests)
Next

'Assistants - Per Assistant'
'Using "msexchangemailboxassistants-total" instance
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeAssistantsPerAssistant_MSExchangeAssistantsPerAssistant WHERE Name = ""msexchangemailboxassistants-total""", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
  WScript.Echo formJSON("AverageEventProcessingTimeInSecondsPerAssistant", objItem.AverageEventProcessingTimeInSeconds)
  WScript.Echo formJSON("AverageEventQueueTimeInSecondsPerAssistant", objItem.AverageEventQueueTimeInSeconds)
  WScript.Echo formJSON("EventsInQueuePerAssistant", objItem.EventsinQueue)
  WScript.Echo formJSON("EventsProcessedPerAssistant", objItem.EventsProcessed)
  WScript.Echo formJSON("EventsProcessedPerSecondPerAssistant", objItem.EventsProcessedPersec)
Next

'Assistants - Per Database'
'Using "msexchangemailboxassistants-total" instance
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_MSExchangeAssistantsPerDatabase_MSExchangeAssistantsPerDatabase WHERE Name = ""msexchangemailboxassistants-total""", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem In colItems
  WScript.Echo formJSON("AverageEventProcessingTimeInSecondsPerDatabase", objItem.AverageEventProcessingTimeInseconds)
  WScript.Echo formJSON("AverageMailboxProcessingTimeInSecondsPerDatabase", objItem.AverageMailboxProcessingTimeInseconds)
  WScript.Echo formJSON("EventsInQueuePerDatabase", objItem.Eventsinqueue)
  WScript.Echo formJSON("MailboxesProcessedPerDatabase", objItem.MailboxesProcessed)
  WScript.Echo formEndLineJSON("MailboxesProcessedPerSecondPerDatabase", objItem.MailboxesprocessedPersec)
Next
WScript.Echo "}"
WScript.Echo "}"
'End of outputing JSON.