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

Dim l_id

Dim l_calfec
Dim l_descripcion

Dim l_idrecursoreservable

Dim l_pacienteid

Dim l_turnos
Dim l_lista
Dim i
Dim l_nuevavisita
Dim l_precio
Dim l_practicarealizada
Dim l_noasistio

l_turnos 		           = request("cabnro")
l_noasistio                = request("cabnro2")


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
	l_sql = l_sql & " and listaprecioscabecera.empnro = " & Session("empnro")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		BuscarPrecio = Replace(l_rs("precio"), ",", ".")
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
	l_sql = l_sql & " and empnro = " & Session("empnro")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		BuscarmediopagoOS = l_rs("id")
	else
		BuscarmediopagoOS = 0
	end if
	l_rs.Close
	Set l_rs = Nothing
end function

'l_tipo 		           = request.querystring("tipo")
'l_idrecursoreservable  = request.Form("idrecursoreservable")
'l_pacienteid     	   = request.Form("pacienteid")
'l_calfec               = request("calfec")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

set l_cm = Server.CreateObject("ADODB.Command")

'--------------
'  Asistio
'--------------
	 l_lista= Split(l_turnos,",")
	 i = 1
	 do while i <= UBound(l_lista)
	 
	 
	  l_sql = "SELECT turnos.*, CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) AS FechaVisita , calendarios.idrecursoreservable idrecursoreservable "
	  l_sql = l_sql & " , isnull(clientespacientes.idobrasocial,0) idobrasocial "	  
	  l_sql = l_sql & " ,  isnull(obrassociales.flag_particular,0) flag_particular  "	  
	  l_sql = l_sql & " ,  isnull(turnos.idmedicoderivador,0) idsolicitadapor  "	  
	  l_sql = l_sql & " FROM turnos "
	  l_sql = l_sql & " INNER JOIN calendarios ON turnos.idcalendario = calendarios.id "
	  l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "	  
	  l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "	  	  
	  l_sql = l_sql & " WHERE turnos.id= " & l_lista(i)
	  l_sql = l_sql & " and turnos.empnro = " & Session("empnro")

	  'Response.write "<script>alert('Operación"& l_sql &" Realizada.');</script>"		  

	  rsOpen l_rs, cn, l_sql, 0
	  do while not l_rs.eof 
	  
	  'Response.write "<script>alert('Operación"& l_rs("idsolicitadapor") &" Realizada.');</script>"	 
	  
		l_sql = "INSERT INTO visitas "
		l_sql = l_sql & "(idturno, idpaciente, idrecursoreservable, fecha ,created_by,creation_date,last_updated_by,last_update_date,empnro ) "
		l_sql = l_sql & "VALUES (" & l_lista(i) & "," & l_rs("idclientepaciente") & "," & l_rs("idrecursoreservable") & ",'" & l_rs("FechaVisita") & "'"&",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0  		

		'Ingreso la lista de empleados a la tabla
		l_nuevavisita = codigogenerado("visitas")		
		l_precio = buscarprecio(l_rs("idobrasocial") , l_rs("idpractica") )			
		
		l_sql = "INSERT INTO practicasrealizadas "
		l_sql = l_sql & "(idvisita, idpractica, idsolicitadapor, precio ,created_by,creation_date,last_updated_by,last_update_date,empnro ) "
		l_sql = l_sql & "VALUES (" & l_nuevavisita & "," & l_rs("idpractica") & "," & l_rs("idsolicitadapor") & "," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0  		

		' Si tiene Obra Social registro el Pago (solo si tiene precio, para no generar informacion innecesaria)
		if l_rs("flag_particular") = 0 and l_precio <> 0 then
			l_practicarealizada = codigogenerado("practicasrealizadas")	
			
			l_sql = "INSERT INTO pagos "
			l_sql = l_sql & "( idpracticarealizada, idmediodepago, idobrasocial, fecha , importe ,created_by,creation_date,last_updated_by,last_update_date,empnro) "
			l_sql = l_sql & "VALUES (" & l_practicarealizada  & "," & BuscarmediopagoOS( ) & "," & l_rs("idobrasocial") & ",'" & l_rs("FechaVisita") & "'," & l_precio &",'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0 

		end if
	  
		l_rs.movenext
	  loop
	  l_rs.close
	 
	 
		
		i = i + 1
	 loop	


'--------------
'  NO Asistio
'--------------
	 l_lista= Split(l_noasistio,",")
	 i = 1
	 do while i <= UBound(l_lista)
	 
	 
	  l_sql = "SELECT turnos.*, CONVERT(VARCHAR(10), calendarios.fechahorainicio, 101) AS FechaVisita , calendarios.idrecursoreservable idrecursoreservable "
	  l_sql = l_sql & " , isnull(clientespacientes.idobrasocial,0) idobrasocial "	  
	  l_sql = l_sql & " ,  isnull(obrassociales.flag_particular,0) flag_particular  "	  
	  l_sql = l_sql & " ,  isnull(turnos.idmedicoderivador,0) idsolicitadapor  "	  
	  l_sql = l_sql & " FROM turnos "
	  l_sql = l_sql & " INNER JOIN calendarios ON turnos.idcalendario = calendarios.id "
	  l_sql = l_sql & " INNER JOIN clientespacientes ON clientespacientes.id = turnos.idclientepaciente "	  
	  l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "	  	  
	  l_sql = l_sql & " WHERE turnos.id= " & l_lista(i)
	  l_sql = l_sql & " and turnos.empnro = " & Session("empnro")

	  'Response.write "<script>alert('Operación"& l_sql &" Realizada.');</script>"		  

	  rsOpen l_rs, cn, l_sql, 0
	  do while not l_rs.eof 
	  
	  'Response.write "<script>alert('Operación"& l_rs("idsolicitadapor") &" Realizada.');</script>"	 
	  
		l_sql = "INSERT INTO visitas "
		l_sql = l_sql & "(idturno, idpaciente, idrecursoreservable, fecha, flag_ausencia ,created_by,creation_date,last_updated_by,last_update_date,empnro ) "
		l_sql = l_sql & "VALUES (" & l_lista(i) & "," & l_rs("idclientepaciente") & "," & l_rs("idrecursoreservable") & ",'" & l_rs("FechaVisita") & "'"&",-1,'"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE(),'"& session("empnro") &"')"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0  		
	  
		l_rs.movenext
	  loop
	  l_rs.close
	 
		i = i + 1
	 loop		 

Set l_cm = Nothing
Response.write "<script>alert('Operación Realizada.');window.parent.parent.opener.ifrm.location.reload();window.parent.parent.close();</script>"
%>

