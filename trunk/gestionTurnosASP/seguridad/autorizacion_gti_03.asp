<%  Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
' Declaracion de Variables locales  -------------------------------------
Dim l_tipo
Dim l_cm
Dim l_sql

Dim l_cystipnro
Dim l_cystipact 
Dim l_cystipsis
Dim l_cystipmsg
Dim l_cystipmail
Dim l_cystipnombre
Dim l_cystipprogdesc
Dim l_cystipprogdet 
Dim l_cystipprogweb 
Dim l_cystipaccion 

' variables de campos que no se usan y no pueden insertarse nulas...


' traer valores del form de alta/modificacion -------------------------------------

l_tipo = request("tipo")

l_cystipnro			= Request.form("cystipnro")
l_cystipact			= Request.form("cystipact")
l_cystipsis			= 0
l_cystipmsg			= Request.form("cystipmsg")
l_cystipmail		= Request.form("cystipmail")
l_cystipnombre		= Request.form("cystipnombre")
l_cystipprogdesc	= Request.form("cystipprogdesc")
l_cystipprogdet		= Request.form("cystipprogdet")
l_cystipprogweb		= Request.form("cystipprogweb")
l_cystipaccion		= Request.form("cystipaccion")


' trasnformar valor de checkboxes en valores logicos --------------------------

IF l_cystipmsg = "on" then
	l_cystipmsg = -1
else
	l_cystipmsg = 0
end if

IF l_cystipmail = "on" then
	l_cystipmail = -1
else
	l_cystipmail = 0
end if

IF l_cystipact = "on" then
	l_cystipact = -1
else
	l_cystipact = 0
end if

	
set l_cm = Server.CreateObject("ADODB.Command")
if l_tipo = "A" then 
	l_sql = "insert into cystipo "
	l_sql = l_sql & "(cystipact , cystipsis, cystipmsg, cystipmail, cystipnombre, "
	l_sql = l_sql & "cystipprogdesc, cystipprogdet , cystipprogweb , cystipaccion) "
	l_sql = l_sql & " values (" & l_cystipact  & "," 
	l_sql = l_sql & l_cystipsis & ", " & l_cystipmsg & ","
	l_sql = l_sql & l_cystipmail & ", '" & l_cystipnombre & "','"
	l_sql = l_sql & l_cystipprogdesc   & "','" &  l_cystipprogdet & "','" &  l_cystipprogweb & "'," & l_cystipaccion  & ")"
else
	l_sql = "update cystipo set "
	l_sql = l_sql & "cystipact		= "  & l_cystipact       & ","
	l_sql = l_sql & "cystipsis		= "  & l_cystipsis       & ","
	l_sql = l_sql & "cystipmsg		= "  & l_cystipmsg       & ","
	l_sql = l_sql & "cystipmail		= "  & l_cystipmail      & ","
	l_sql = l_sql & "cystipnombre	= '" & l_cystipnombre    & "',"
	l_sql = l_sql & "cystipprogdesc	= '" & l_cystipprogdesc  & "',"
	l_sql = l_sql & "cystipprogdet	= '" & l_cystipprogdet   & "',"
	l_sql = l_sql & "cystipprogweb	= '" & l_cystipprogweb   & "',"
	l_sql = l_sql & "cystipaccion 	= "  & l_cystipaccion    
	l_sql = l_sql & " where cystipnro = " & l_cystipnro

end if

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql

cmExecute l_cm, l_sql, 0

Set cn = Nothing
Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.opener.ifrm.location = 'autorizacion_gti_01.asp';window.close();</script>"
%>
