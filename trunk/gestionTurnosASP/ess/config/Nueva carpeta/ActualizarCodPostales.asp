<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs
Dim l_rs2

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")

l_sql = " SELECT * FROM tkt_localidad "
l_sql = l_sql & " INNER JOIN tkt_provincia ON tkt_provincia.pronro = tkt_localidad.pronro "
l_sql = l_sql & " WHERE (loccodpos is null or loccodpos = '') " '  and locnro = 27
'l_sql = l_sql & " and (locnro >= 1 and locnro <=500 ) "
'l_sql = l_sql & " and (locnro >= 501 and locnro <=1000 ) "
'l_sql = l_sql & " and (locnro >= 1001 and locnro <=1500 ) "
'l_sql = l_sql & " and (locnro >= 1501 and locnro <=2000 ) "
'l_sql = l_sql & " and (locnro >= 2001 and locnro <=3000 ) "
'l_sql = l_sql & " and (locnro >= 3001 and locnro <=4000 ) "
'l_sql = l_sql & " and (locnro >= 4001 and locnro <=5000 ) "
'l_sql = l_sql & " and (locnro >= 5001 and locnro <=6000 ) "
'l_sql = l_sql & " and (locnro >= 6001 and locnro <=7000 ) "
'l_sql = l_sql & " and (locnro >= 7001 and locnro <=9000 ) "
'l_sql = l_sql & " and (locnro >= 9001 and locnro <=11000 ) "
l_sql = l_sql & " and (locnro >= 11001 and locnro <=13000 ) "
'l_sql = l_sql & " and (locnro >= 13001 and locnro <=15000 ) "
'l_sql = l_sql & " and (locnro >= 15001 and locnro <=17000 ) "
'l_sql = l_sql & " and (locnro >= 17001 and locnro <=19000 ) "
'l_sql = l_sql & " and (locnro >= 19001 and locnro <=21000 ) "
'l_sql = l_sql & " and (locnro >= 21001 and locnro <=23000 ) "
rsOpen l_rs, cn, l_sql, 0 
do while not l_rs.eof 
		
	l_sql = " SELECT cpostal FROM locaoncca "
	l_sql = l_sql & " WHERE locaoncca.localidad = '" & l_rs("locdes") & "'"
	l_sql = l_sql & "   AND locaoncca.provincia = '" & l_rs("prodes") & "'"
	rsOpen l_rs2, cn, l_sql, 0 
	if not l_rs2.eof then 
		l_sql = "UPDATE tkt_localidad"
		l_sql = l_sql & " SET loccodpos = '" & l_rs2("cpostal") & "'"
		l_sql = l_sql & " WHERE locnro = " & l_rs("locnro")
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	end if
	l_rs2.close
	l_rs.movenext
loop

Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');</script>"
%>

