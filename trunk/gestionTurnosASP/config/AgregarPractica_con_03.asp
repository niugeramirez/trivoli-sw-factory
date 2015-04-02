<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/sqls.inc"-->
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
Dim l_practicarealizada

Set l_rs = Server.CreateObject("ADODB.RecordSet")


' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado(tabla)
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id",tabla)
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()

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

function BuscarmediopagoOS( )
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM mediosdepago "
	l_sql = l_sql & " WHERE flag_obrasocial = -1 " 
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		BuscarmediopagoOS = l_rs("id")
	else
		BuscarmediopagoOS = 0
	end if
	l_rs.Close
	Set l_rs = Nothing
end function

l_tipo 		               = request.querystring("tipo")
l_idvisita                 = request.Form("idvisita")
l_practicaid 			   = request.Form("practicaid")
l_solicitadopor			   = request.Form("idrecursoreservable")

l_idpracticarealizada      = request.Form("idpracticarealizada") 
l_precio 				   = request.Form("precio2")

set l_cm = Server.CreateObject("ADODB.Command")

 l_sql = "SELECT isnull(clientespacientes.idobrasocial,0) idobrasocial,  isnull(obrassociales.flag_particular,0) flag_particular, CONVERT(VARCHAR(10),visitas.fecha, 101) AS FechaVisita "
 l_sql = l_sql & " FROM visitas "
 l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "	  
 l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "	  	  
 l_sql = l_sql & " WHERE visitas.id= " & l_idvisita
 rsOpen l_rs, cn, l_sql, 0
 if  not l_rs.eof then
 	l_idobrasocial = l_rs("idobrasocial")
 else
 	l_idobrasocial = 0
 end if


if l_tipo = "A" then
	'l_precio = BuscarPrecio(l_idobrasocial, l_practicaid )
	l_sql = "INSERT INTO practicasrealizadas (idvisita , idpractica , idsolicitadapor , precio ,created_by,creation_date,last_updated_by,last_update_date ) "
	l_sql = l_sql & " VALUES ( " & l_idvisita & ","  & l_practicaid & "," & l_solicitadopor & "," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"	

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0	
	

		'Response.write "<script>alert('Operación " & l_rs("flag_particular") & "Realizada.');</script>"
				
		' Si tiene Obra Social registro el Pago
		if l_rs("flag_particular") = 0 then
			l_practicarealizada = codigogenerado("practicasrealizadas")				
			
			l_sql = "INSERT INTO pagos "
			l_sql = l_sql & "( idpracticarealizada, idmediodepago, idobrasocial, fecha , importe ,created_by,creation_date,last_updated_by,last_update_date) "
			l_sql = l_sql & "VALUES (" & l_practicarealizada  & "," & BuscarmediopagoOS( ) & "," & l_rs("idobrasocial") & ",'" & l_rs("FechaVisita") & "'," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0 
		
		end if
	  
else
	
	
		l_sql = "UPDATE practicasrealizadas"
		l_sql = l_sql & " SET idpractica = " & l_practicaid
		l_sql = l_sql & " , idsolicitadapor = " & l_solicitadopor		
		l_sql = l_sql & " , precio = " & l_precio
		l_sql = l_sql & "    ,last_updated_by = '" &session("loguinUser") & "'"
		l_sql = l_sql & "    ,last_update_date = GETDATE()" 		
  	    l_sql = l_sql & " WHERE id = " & l_idpracticarealizada	

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0		
		
end if 	

response.write l_sql & "<br>"

Set l_cm = Nothing
if l_tipo = "A" then
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
else
	Response.write "<script>alert('Operación Realizada.');window.parent.opener.parent.ifrm.location.reload();window.parent.close();</script>"
end if
%>

