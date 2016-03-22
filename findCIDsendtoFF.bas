Attribute VB_Name = "Module21"
Dim ele
Dim ie As Object
Dim objShell As Object
Dim objWindow As Object
Dim objItem As Object
Dim entry
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim cidInt  As Integer

Dim x, y
Dim dialedNumber As String

Sub findCIDsendtoFF()
Worksheets("FraudNotification").Range("L2").Select
'clear contents of CID'
Worksheets("FraudNotification").Columns(12).ClearContents
'assign header for CIDs (L1)'
Worksheets("FraudNotification").Cells(1, 12).Value = "CIDs (query: *customerkey, *uslog)"
Worksheets("FraudNotification").Cells(1, 12).Font.Bold = True

'Determine if a specific instance of IE is already open.
    Set objShell = CreateObject("Shell.Application")
    IE_count = objShell.Windows.Count
    For x = 0 To (IE_count - 1)
        On Error Resume Next    ' sometimes more web pages are counted than are open
        my_url = objShell.Windows(x).Document.Location
        my_title = objShell.Windows(x).Document.Title

        'You can use my_title of my_url, whichever you want
        If my_title Like "Kibana 3 - Logstash Search" Or my_url Like "*log.j2noc.com*" Then   'identify the existing web page
            Set ie = objShell.Windows(x)
            Exit For
        Else
        End If
    Next
    

If ie.READYSTATE = READYSTATE.READYSTATE_COMPLETE Then
For y = 2 To Range("A1").End(xlDown).Row
     Dim msgString As String
     Dim msgArray() As String
     Dim cidArray() As String
     
     'populate search text field'
     dialedNumber = Worksheets("FraudNotification").Cells(y, 1)
     'MsgBox ("dialed number: " + dialedNumber)'
     PopulateTextField (dialedNumber)
     
     'click the magnifying glass'
     'ClickSearch'
     ie.Document.parentwindow.execscript "document.querySelectorAll('form ul li a')[1].click()"
     Application.Wait (Now + TimeValue("0:00:01"))
     'Application.Wait (Now + TimeValue("0:00:01"))'
     Do While ie.Busy: DoEvents: Loop
     
     'START OF IF STATEMENT READY STATE!!!!!!'
     'expand/parse entry and insert into excel'
     If ie.READYSTATE = READYSTATE.READYSTATE_COMPLETE Then
            For Each entry In ie.Document.All.tags("table") 'for each result in kibana'
                If InStr(entry.innerHTML, "tbody") > 0 Then
                For i = 0 To 1 'takes 1 results (tbody); (determines how many sample entries for phone number)'
                    For j = 0 To 1 'expands j number of entries (tr)'
                        entry.Children(i).Children(j).Click
                        Do While ie.Busy: DoEvents: Loop
                                       'tbody'      'tr'        'td'      'table'       'tbody'     'tr'        'td'
                        msgString = entry.Children(i).Children(j).Children(0).Children(1).Children(1).Children(10).Children(2).innerText 'get message attribute innertext'
                        msgArray = Split(msgString, "queue->m_iCustomerKey ", 2) 'parsing CID'
                        cidArray = Split(msgArray(1), """", 2) 'still parsing CID'
                        'MsgBox ("CID: " + cidArray(0))'
                    Next 'for loop (tr)'
                Next 'for loop (tbody)'
                End If
            Next 'for loop (table)'
            
            ActiveCell.Value = Val(cidArray(0))
            ActiveCell.Offset(1, 0).Select
            
            'erase string and arrays'
            msgString = ""
            Erase msgArray
            Erase cidArray
     End If
     Do While ie.Busy: DoEvents: Loop
     'END OF IF STATEMENT READY STATE!!!!!!'
Next y
End If 'end outer ready state'
    
End Sub

Sub PopulateTextField(number)
Dim counter
counter = 0
For Each ele In ie.Document.All.tags("input")
    ele.Value = "*" + number 'the text you want is here'
    Do While ie.Busy: DoEvents: Loop
    If counter = 8 Then Exit For
    counter = counter + 1
Next
End Sub

'click search button (21st <i></i> is search button)'
Private Sub ClickSearch()
Dim counter
counter = 0
For Each x In ie.Document.All.tags("i")
    If counter = 21 Then
        x.Click
        Do While ie.Busy: DoEvents: Loop
    End If
    Do While ie.Busy: DoEvents: Loop
    If counter = 21 Then Exit For
    counter = counter + 1
Next
End Sub
