<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

dim l_idvisita

Dim l_practicaid
Dim l_solicitadopor
Dim l_precio
Dim l_turnoid
Dim l_idobrasocial
Dim l_idpracticarealizada

Set l_rs = Server.CreateObject("ADODB.RecordSet")

function BuscarPrecio(idobrasocial, idpractica )
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT precio "
	l_sql = l_sql & " FROM listaprecioscabecera "
	l_sql = l_sql & " INNER JOIN listapreciosdetalle ON listapreciosdetalle.idlistaprecioscabecera = listaprecioscabecera.id "
	l_sql = l_sql & " WHERE flag_activo = -1 " 
	l_sql = l_sql & " AND idobrasocial = " & idobrasocial
	l_sql = l_sql & " AND idpractica = " & idpractica
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		BuscarPrecio = l_rs("precio")
	else
		BuscarPrecio = 0
	end if
	l_rs.Close
	Set l_rs = Nothing
end function


l_tipo 		               = request.querystring("tipo")
l_idvisita                 = request.Form("idvisita")
l_practicaid 			   = request.Form("practicaid")
l_solicitadopor			   = request.Form("idrecursoreservable")
l_turnoid				   = request.Form("idvisita")
l_idpracticarealizada       = request.Form("idpracticarealizada") 

response.write l_idvisita & " - " & l_practicaid & " - " & l_solicitadopor


set l_cm = Server.CreateObject("ADODB.Command")

 l_sql = "SELECT clientespacientes.idobrasocial "
 l_sql = l_sql & " FROM visitas "
 l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "	  
 l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "	  	  
 l_sql = l_sql & " WHERE visitas.id= " & l_turnoid
 rsOpen l_rs, cn, l_sql, 0
 if  not l_rs.eof then
 	l_idobrasocial = l_rs("idobrasocial")
 else
 	l_idobrasocial = 0
 end if


if l_tipo = "A" then
	l_precio = BuscarPrecio(l_idobrasocial, l_practicaid )
	l_sql = "INSERT INTO practicasrealizadas (idvisita , idpractica , idsolicitadapor , precio ) "
	l_sql = l_sql & " VALUES ( " & l_idvisita & ","  & l_practicaid & "," & l_solicitadopor & "," & l_precio & ")"	
else
	l_precio = request.Form("precio")
	l_sql = "INSERT INTO practicasrealizadas (idvisita , idpractica , idsolicitadapor , precio ) "
	l_sql = l_sql & " VALUES ( " & l_idvisita & ","  & l_practicaid & "," & l_solicitadopor & "," & l_precio & ")"	
	
		l_sql = "UPDATE practicasrealizadas"
		l_sql = l_sql & " SET idpractica = " & l_practicaid
		l_sql = l_sql & " , idsolicitadapor = " & l_solicitadopor		
		l_sql = l_sql & " , precio = " & l_precio
  	    l_sql = l_sql & " WHERE id = " & l_idpracticarealizada	
		
end if 	

response.write l_sql & "<br>"

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing
if l_tipo = "A" then
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
else
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.parent.ifrm.location.reload();window.parent.close();</script>"
end if
%>

