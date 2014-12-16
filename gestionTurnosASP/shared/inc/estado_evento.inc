<%' on error goto 0
Dim l_estadonro
function eventocerrado
	Dim l_rsx
	Dim l_sqlx
	
	Set l_rsx = Server.CreateObject("ADODB.RecordSet")
	l_sqlx = " SELECT estevenro "
	l_sqlx = l_sqlx & " FROM cap_evento "
	l_sqlx = l_sqlx & " WHERE evenro = " & l_evenro
	rsOpen l_rsx, cn, l_sqlx, 0 
	if not l_rsx.eof AND (l_rsx("estevenro")= 6 OR l_rsx("estevenro")= 4) then
		eventocerrado = -1
	else
		eventocerrado = 0
	end if
	l_estadonro = l_rsx("estevenro")
	l_rsx.Close
end function
%>
