<% Option Explicit

Dim l_tipo
Dim l_registrar

l_registrar        = request.Form("registrar_envio")
l_tipo             = request.Form("tipo")

if l_registrar <> "on" then
	if l_tipo <> "T" then
		if l_tipo = "O" then
			Response.AddHeader "Content-Disposition", "attachment;filename=VisitasOblig.txt" 
		else
			Response.AddHeader "Content-Disposition", "attachment;filename=VisitasVol.txt" 
		end if
	else
		Response.AddHeader "Content-Disposition", "attachment;filename=VisitasTod.txt" 
	end if

	Response.ContentType = "application/octet-stream"
end if

 %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<% 
Dim l_rs
Dim l_sql
Dim l_orden
Dim l_idos
Dim l_fechadesde
Dim l_fechahasta
Dim l_factura
Dim l_nro_obrasocial
Dim l_cod_practica
Dim l_precio_practica
Dim l_matricula
Dim l_dia
Dim l_mes
Dim l_oblig
Dim l_sql2
Dim l_cm
Dim l_sql3

l_orden = " ORDER BY  visitas.fecha"

Function PrecioPractica(id_practica , id_obrasocial )
dim l_rs
dim l_sql

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  precio AS precio "
l_sql = l_sql & " FROM listaprecioscabecera "
l_sql = l_sql & " INNER JOIN listapreciosdetalle ON listapreciosdetalle.idlistaprecioscabecera = listaprecioscabecera.id "
l_sql = l_sql & " WHERE flag_activo = -1 "
l_sql = l_sql & " AND idobrasocial = " & id_obrasocial
l_sql = l_sql & " AND idpractica = " & id_practica
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	PrecioPractica = Replace(l_rs("precio"), ",", ".")
else
	PrecioPractica = 0
end if
l_rs.close

end Function

Function PrecioPracticaTxt(id_practica , id_obrasocial )
dim l_rs
dim l_sql



Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = "SELECT  precio*100 AS precio "
l_sql = l_sql & " FROM listaprecioscabecera "
l_sql = l_sql & " INNER JOIN listapreciosdetalle ON listapreciosdetalle.idlistaprecioscabecera = listaprecioscabecera.id "
l_sql = l_sql & " WHERE flag_activo = -1 "
l_sql = l_sql & " AND idobrasocial = " & id_obrasocial
l_sql = l_sql & " AND idpractica = " & id_practica
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	PrecioPracticaTxt = l_rs("precio")
else
	PrecioPracticaTxt = 0
end if
l_rs.close

end Function

l_fechadesde       = request.Form("fechadesde")
l_fechahasta       = request.Form("fechahasta")
l_idos             = request.Form("idos")
l_factura          = request.Form("factura")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")

if l_registrar = "on" then
	l_sql3 = " DELETE FROM facturacionobrassociales WHERE factura = '" & request.Form("factura") & "'"

	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql3
	cmExecute l_cm, l_sql3, 0
end if

l_sql = "SELECT  recursosreservables.nro_matricula, visitas.fecha, CONVERT(VARCHAR(10), visitas.fecha, 101) fechavisita, codigos.codigo, clientespacientes.nro_obra_social, practicasrealizadas.idpractica, clientespacientes.idobrasocial" 
l_sql = l_sql & ", obrassociales.descripcion nombreos, clientespacientes.apellido, clientespacientes.nombre, practicas.descripcion nombrepractica, recursosreservables.descripcion nombremedico, ISNULL(clientespacientes.afiliado_obligatorio,'N')"
l_sql = l_sql & " FROM visitas "
l_sql = l_sql & " INNER JOIN practicasrealizadas ON practicasrealizadas.idvisita = visitas.id "
l_sql = l_sql & " LEFT JOIN recursosreservables ON recursosreservables.id = visitas.idrecursoreservable "
l_sql = l_sql & " LEFT JOIN clientespacientes ON clientespacientes.id = visitas.idpaciente "
l_sql = l_sql & " LEFT JOIN obrassociales ON obrassociales.id = clientespacientes.idobrasocial "
l_sql = l_sql & " LEFT JOIN practicas ON practicas.id = practicasrealizadas.idpractica "
l_sql = l_sql & " LEFT JOIN codigospracticas codigos ON codigos.idpractica = practicasrealizadas.idpractica AND codigos.idobrasocial = clientespacientes.idobrasocial"
l_sql = l_sql & " WHERE  visitas.fecha  >= " & cambiafecha(l_fechadesde,"YMD",true) 
l_sql = l_sql & " AND  visitas.fecha <= " & cambiafecha(l_fechahasta,"YMD",true) 
l_sql = l_sql & " AND  isnull(visitas.flag_ausencia,0) <> -1" 
l_sql = l_sql & " AND visitas.empnro = " & Session("empnro")   

if l_tipo <> "T" then
	if l_tipo = "O" then
		l_sql = l_sql & " AND clientespacientes.afiliado_obligatorio = 'S'"
	else
		l_sql = l_sql & " AND (clientespacientes.afiliado_obligatorio IS NULL OR clientespacientes.afiliado_obligatorio <> 'S')"
	end if
end if

if l_idos <> "0" then
	l_sql = l_sql &" AND exists ( select (ospago.id) from pagos LEFT JOIN obrassociales ospago ON ospago.id = pagos.idobrasocial where pagos.idpracticarealizada = practicasrealizadas.id and ospago.id IN "& l_idos & ")" 
end if
l_sql = l_sql & " " & l_orden

rsOpen l_rs, cn, l_sql, 0 

do until l_rs.eof	
	l_precio_practica = PrecioPracticaTxt(l_rs("idpractica") , l_rs("idobrasocial") )
	
	l_precio_practica = "0000000000" & l_precio_practica
	
	l_nro_obrasocial = l_rs("nro_obra_social")
	
	if IsNull(l_nro_obrasocial) then
		l_nro_obrasocial = "0000000000000000"
	else
		l_nro_obrasocial = "0000000000000000" & l_rs("nro_obra_social")
	end if
	
	l_nro_obrasocial = Right(l_nro_obrasocial,16)
	
	l_cod_practica = l_rs("codigo")
	
	if IsNull(l_cod_practica) then
		l_cod_practica = "000000000"
	else
		l_cod_practica = "000000000" & l_rs("codigo")
	end if
	
	l_cod_practica = Right(l_cod_practica,9)
	
	l_factura = "00000000" & l_factura
	
	l_matricula = l_rs("nro_matricula")
	
	l_matricula = "000000" & l_matricula
	
	l_dia = "00" & Day(l_rs("fecha"))
	l_mes = "00" & Month(l_rs("fecha"))
	
	if l_registrar <> "on" then
		Response.Write Right(l_factura,8)  										' Factura
		Response.Write "01" 													' Tipo Profesional Prestador
		Response.Write "B" 														' Codigo Provincia Matricula Prestador
		Response.Write Right(l_matricula,6) 									' Matricula Prestador
		Response.Write Space(20) 												' Apellido Prestador
		Response.Write Space(20) 												' Nombre Prestador
		Response.Write Space(2) 												' Tipo Profesional Prescriptor
		Response.Write Space(1) 												' Codigo Provincia Matricula Prescriptor
		Response.Write Space(6) 												' Matricula Prescriptor
		Response.Write Space(20) 												' Apellido Prescriptor
		Response.Write Space(20) 												' Nombre Prescriptor
		Response.Write Space(8) 						    					' Fecha Prescripcion
		Response.Write Mid(l_nro_obrasocial,1,3) 								' Nodo Cuenta
		Response.Write Mid(l_nro_obrasocial,4,8) 								' Numero Cuenta
		Response.Write Mid(l_nro_obrasocial,12,2) 								' Numero Adherente
		Response.Write Right(l_nro_obrasocial,3) 								' Numero Efector
		Response.Write Year(l_rs("fecha")) & Right(l_mes,2) & Right(l_dia,2) 	' Fecha Realizacion
		Response.Write "AMB" 													' Origen Practica
		Response.Write Mid(l_cod_practica,1,3) 									' Tipo Nomenclador
		Response.Write Right(l_cod_practica,6) 									' Codigo Practica
		Response.Write "0000000001" 											' Cantidad Consultas
		Response.Write Right(l_precio_practica,10) 								' Valor Unitario
		Response.Write "TOD" 													' Terminador
		
		Response.Write vbCrLf
	else
		l_oblig = l_rs("afiliado_obligatorio")
		
		if l_oblig = "" then
			l_oblig = "N"
		end if
		
		l_sql2 = "INSERT INTO facturacionobrassociales"
		l_sql2 = l_sql2 & " (factura, fecha, obrasocial, paciente, practica, medico, afiliado_obligatorio, monto, empnro,created_by,creation_date,last_updated_by,last_update_date)"
		l_sql2 = l_sql2 & " VALUES (" 
		l_sql2 = l_sql2 &  "'" & request.Form("factura") & "','" & l_rs("fechavisita") & "','" & l_rs("nombreos") & "','" & l_rs("apellido") & " " & l_rs("nombre") &  "','" & l_rs("nombrepractica") &  "','" & l_rs("nombremedico")  & "','" & l_oblig  & "'," & PrecioPractica(l_rs("idpractica") , l_rs("idobrasocial") ) & ",'" & session("empnro") & "','"&session("loguinUser")&"',GETDATE(),'"&session("loguinUser")&"',GETDATE())"
	
		l_cm.activeconnection = cn
		
		l_cm.CommandText = l_sql2
		
		cmExecute l_cm, l_sql2, 0
	end if
	
	l_rs.MoveNext
loop 

l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing

if l_registrar = "on" then
	Response.write "<script>alert('Operacion Realizada.');window.parent.opener.ifrm.location.reload();window.parent.close();</script>"
end if
%>