<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 


on error goto 0

Dim l_tipo
Dim l_cm
Dim l_sql
Dim l_rs

dim l_idvisita
dim l_motivo
dim l_cantturnossimult  
dim l_cantsobreturnos     
dim l_turnoid



'l_tipo 		               = request.querystring("tipo")
l_idvisita                 = request.Form("idvisita")
'l_motivo                   = request.Form("motivo")
'l_turnoid                  = request.Form("turnoid")

' l_domicilio      = request.Form("domicilio")
'l_idobrasocial      = request.Form("legape")




'if len(l_legfecing) = 0 then
'	l_legfecing = "null"
'else 
'	l_legfecing = cambiafecha(l_legfecing,"YMD",true)	
'end if 
'if len(l_legfecnac) = 0 then
'	l_legfecnac = "null"
'else 
'	l_legfecnac = cambiafecha(l_legfecnac,"YMD",true)	
'end if 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")


 l_sql = "SELECT * "
 l_sql = l_sql & " FROM practicasrealizadas "
 l_sql = l_sql & " WHERE idvisita = " & l_idvisita
 rsOpen l_rs, cn, l_sql, 0
 do while not l_rs.eof 

	l_sql = "DELETE FROM pagos "
	l_sql = l_sql & " WHERE idpracticarealizada = " & l_rs("id")
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0

	l_rs.movenext
loop
l_rs.close

l_sql = "DELETE FROM practicasrealizadas "
l_sql = l_sql & " WHERE idvisita = " & l_idvisita
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0


l_sql = "DELETE FROM visitas "
l_sql = l_sql & " WHERE id = " & l_idvisita
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0
Set l_cm = Nothing

Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
%>

