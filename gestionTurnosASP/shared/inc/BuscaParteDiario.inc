<%
function buscarParte(ternro,fecha)
	Dim l_cadena
	Dim l_rsbp
	Dim l_sql
	Set l_rsbp = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = "SELECT gti_cabparte.gcpnro,gti_cabparte.gcpdesabr "
	l_sql = l_sql & "FROM gti_detturtemp INNER JOIN gti_cabparte ON gti_cabparte.gcpnro = gti_detturtemp.gcpnro "
	l_sql = l_sql & "WHERE gti_detturtemp.ternro=" & ternro 
	l_sql = l_sql & " AND gttempdesde <=" & cambiafecha(fecha,"YMD",true)
	l_sql = l_sql & " AND gttemphasta >=" & cambiafecha(fecha,"YMD",true)
	
	rsOpen l_rsbp, cn, l_sql, 0 
	if not l_rsbp.eof then
		l_cadena = l_rsbp("gcpnro")
		l_cadena = l_cadena & ": " & l_rsbp("gcpdesabr")
	else
		l_cadena= "vacio"
	end if
	l_rsbp.close
	Set l_rsbp = nothing
	
	buscarParte= l_cadena
	
end function

%>
