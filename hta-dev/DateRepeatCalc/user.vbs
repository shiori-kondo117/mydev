Private Sub window_onload()
	Call window.ResizeTo(340,480)
End Sub
Private Sub cmdExec()
	If (F_DATE.value = "" Or _
		F_REPEAT.value = "") Then
		msgbox "日付と回数を入力して下さい。"
		Exit Sub
	End If

	intDay = 7

	If (F_DAY.value <> "") Then
		intDay = F_DAY.value
	End If

	strDate=F_DATE.value
	strHTML = "<table border=""1"" cellpadding=""2"" cellspacing=""2"">" & vbCrLf
	strHTML = strHTML & "<tr>" & vbCrLf
	strHTML = strHTML & "<th>ＮＯ</th>" & vbCrLf
	strHTML = strHTML & "<th>日付</th>" & vbCrLf
	strHTML = strHTML & "</tr>" & vbCrLf
	For i=1 to F_REPEAT.value
		strHTML = strHTML & "<tr>" & vbCrLf
		strHTML = strHTML & "<td>" & i & "</td>" & vbCrLf
		strHTML = strHTML & "<td>" & strDate & "</td>" & vbCrLf
		strDate = DateAdd("d", intDay, strDate)
		strHTML = strHTML & "</tr>" & vbCrLf
	Next
	strHTML = strHTML & "</table>" & vbCrLf
	ID_RESULT.innerHTML = strHTML
End Sub
