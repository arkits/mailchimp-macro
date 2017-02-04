'MailChimp Macro
'Based on MailChimp's RESTFUL API
'By ArKits - arkits@outlook.com

Attribute VB_Name = "Module1"
Global Const apikey = "123456789012345678901234567890-us2"
Global Const strListId = "12345678"
Global Const strMailChimpURL = "http://us2.api.mailchimp.com/1.3/" 'Match location with the apikey'

Function isSubscribed(ByVal strEmailAddress)
'Checks to see if an email address is subscribed
'By default, MailChimp throws an HTTP error to indicate the subscriber's status.
	Dim objXMLHTTP As Object
	Dim strResponseText As String
	Dim strError As String
	Dim strStatus As String
	Dim strURL As String

	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

	strURL = strMailChimpURL & "?method=listMemberInfo&output=xml&apikey=" & apikey & "&id=" & strListId & _
	"&email_address=" & strEmailAddress

	objXMLHTTP.Open "GET", strURL, False
	objXMLHTTP.send

	strResponseText = objXMLHTTP.responsetext

	If InStr(1, strResponseText, "status") > 0 Then
		If strStatus = "Unsubscribed" Then
			isSubscribed = "Unsubscribed"
		Else
			isSubscribed = "Already subscribed"
		End If
	Else
		isSubscribed = "Not registered"
	End If
	Set objXMLHTTP = Nothing
	Exit Function
	errorHandler:
	isSubscribed = "ERR : " & Err.Description

End Function


Function listSubscribe(ByVal strEmailAddress, Optional ByVal strFirstName = "", Optional ByVal strLastName = "", Optional ByVal strPhone = "")
'Subscribes an email address to the list (with additional details)
	On Error GoTo errorHandler

	Dim objXMLHTTP As Object
	Dim strResponseText As String
	Dim strError As String
	Dim strURL As String
	Dim strIsSubscribed As String

'Checking and handling whether a person is already in the system'
	strIsSubscribed = isSubscribed(strEmailAddress)
	If strIsSubscribed <> "Not registered" Then
		listSubscribe = strIsSubscribed
		Exit Function
	End If


	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

	strURL = strMailChimpURL & "?method=listSubscribe&output=xml&apikey=" & apikey & "&id=" & strListId & _
	"&email_address=" & strEmailAddress & _
	"&merge_vars=" & _
	"&merge_vars[FNAME]=" & strFirstName & _
	"&merge_vars[LNAME]=" & strLastName & _
	"&merge_vars[phone]=" & strPhone & _
	"&send_welcome=false" & _
	"&double_optin=false"

	objXMLHTTP.Open "GET", strURL, False
	objXMLHTTP.send

	strResponseText = objXMLHTTP.responsetext

	If InStr(1, strResponseText, "error") > 0 Then
		MsgBox "If you can read this, then AK messed up listSubscribe..." & strResponseText
	Else

	End If

	Set objXMLHTTP = Nothing
	Exit Function
	errorHandler:
	listSubscribe = "ERR : " & Err.Description
End Function



Function DeleteRows()
'Deleting the row which have "TRUE" in colums 14 to 16

	Dim i As Integer
	For i = 14 To 16
		Columns(i).Select
		Dim rngRange As Range
		Dim lngNumRows, lngFirstRow, lngLastRow, lngCurrentRow As Long
		Dim lngCompareColumn As Long
		Set rngRange = Selection.CurrentRegion
		lngNumRows = rngRange.Rows.Count
		lngFirstRow = rngRange.Row
		lngLastRow = lngFirstRow + lngNumRows - 1
		lngCompareColumn = ActiveCell.Column
    'For faster execution, turn off the screen refresh
		Application.ScreenUpdating = False
    'For each row, check to see if the comparison column is true. If so, delete it.
		For lngCurrentRow = lngLastRow To lngFirstRow Step -1
			If (Cells(lngCurrentRow, lngCompareColumn).Text = "TRUE") Then _
			Rows(lngCurrentRow).Delete
			Next lngCurrentRow
			Next i

'Turn the screen updates back on
		Application.ScreenUpdating = True

End Function

Sub test()

			Call DeleteRows
			Dim LastRow As Long
			With ActiveSheet
				LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
			End With
			Dim j As Integer
			For j = 2 To LastRow
				Email = ActiveSheet.Cells(j, 1)
				Name = ActiveSheet.Cells(j, 3)
				Call listSubscribe(Email, Name)
				Next j

				MsgBox "Yay it worked!"
				
	End Sub
