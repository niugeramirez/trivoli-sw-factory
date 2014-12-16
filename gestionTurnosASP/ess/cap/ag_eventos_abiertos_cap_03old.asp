<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/util.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo		: ag_eventos_cerrados_cap_03.asp
Descripcion	: Consulta de Eventos por empleados
Autor		: Lisandro Moro
Fecha		: 25/03/2004
-----------------------------------------------------------------------------
-->
<% 
'on error goto 0

Dim l_cm
Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_tipo
Dim l_orden

Dim l_ternro
Dim l_empleg
Dim l_codproc
Dim l_evenro
Dim l_donde
Dim l_canjus
Dim l_valorevenro
'Dim l_usrname
'Dim	l_usrmail
Dim l_emailleg

Dim l_reportaa
dim l_empemail
dim l_empleado
dim l_evedesabr
Dim l_eveorigen
Dim l_eveforeva

l_tipo    = request("tipo")
l_ternro = request("ternro")
l_evenro = request("evenro")
l_canjus = request("canjus")
'Guardo los datos en la BD
   
Set l_rs = Server.CreateObject("ADODB.RecordSet")	
Set l_cm = Server.CreateObject("ADODB.Command")
l_cm.activeconnection = Cn	

l_sql = "SELECT orden FROM cap_candidato"
l_sql = l_sql & "  WHERE evenro = " & l_evenro
l_sql = l_sql & " order by orden desc "
rsOpen l_rs, cn, l_sql, 0 	
if l_rs.eof then
	l_orden = 1
else
	l_rs.MoveFirst
	l_orden = l_rs("orden") + 1
end if
l_rs.Close

l_sql = " SELECT eveorigen, eveforeva"
l_sql = l_sql & " FROM cap_evento "
l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.eof then
	l_eveforeva = l_rs("eveforeva")	
	if isnull(l_rs("eveorigen"))   then 
		l_eveorigen = 0
	else 
		l_eveorigen = l_rs("eveorigen")	
	end if
end if 
l_rs.Close



if l_tipo <> "Q" then
'Me fijo si el empleado ya existe
	l_sql = "SELECT * FROM cap_candidato WHERE ternro=" & l_ternro & " AND "	
	l_sql = l_sql & " cap_candidato.evenro = " & l_evenro
	rsOpen l_rs, cn, l_sql, 0 	
	if l_rs.eof then
		select case l_tipo
			case "C"
				l_donde = 0
			case "P"
				l_donde = 1
		end select
		'Inserto el dato en la BD
		
		if l_eveforeva = 1 then
			l_valorevenro = l_evenro
		else 
			l_valorevenro = l_eveorigen
		end if 		
		
	    l_sql = "INSERT INTO cap_candidato "
		l_sql = l_sql & "(evenro, ternro, selcannro, conf, canfecini, canfecfin, canentnro, recdip, fecrecdip, confpart, invitado, orden, invext, cancanths, canjus )"
		l_sql = l_sql & " values ("& l_valorevenro & ", " & l_ternro  & ", 0, " & l_donde & ", null, null, 0, 0, null, 0,0," & l_orden & ",0,0,'' )"
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
		'Caledarios
		'----------------------------------------------------------------
		l_rs.close
		if l_tipo = "P" then
			l_sql = " select calnro "
			l_sql = l_sql & " from cap_evento "
			l_sql = l_sql & " inner join cap_calendario on cap_calendario.evenro = cap_evento.evenro "
			l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
			rsOpen l_rs, cn, l_sql, 0
			if not l_rs.eof then
			   l_rs.MoveFirst
			end if
			Do While Not l_rs.eof
				l_sql = "INSERT INTO cap_part_cal "
				l_sql = l_sql & "(calnro, ternro) "
				l_sql = l_sql & " VALUES (" & l_rs(0) & "," & l_ternro & ")"
				l_cm.activeconnection = Cn
				l_cm.CommandText = l_sql
				cmExecute l_cm, l_sql, 0
				l_rs.MoveNext
			Loop
			l_rs.Close
			
			' actualizo la cantidad de participantes
			l_sql = "UPDATE cap_evento " 
			l_sql = l_sql & " SET evecanrealalu = evecanrealalu + 1"
			l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
			
		end if
		'----------------------------------------------------------------		
	else
		'YA existe para el evento
		%><script>
			alert("Ya existe el empleado en el evento.");
			window.close();
		</script><%
		response.end
	end if
'l_rs.close
else 'Quitar del evento
	'Primero busco si es participante o candidato
	
	l_sql = "SELECT conf FROM cap_candidato "
	l_sql = l_sql & " WHERE ternro=" & l_ternro 
	l_sql = l_sql & " AND cap_candidato.evenro = " & l_evenro
	rsOpen l_rs, cn, l_sql, 0 	
	if not l_rs.eof then 
		if l_rs("conf") = 1 then
		' actualizo la cantidad de participantes
			l_sql = "UPDATE cap_evento " 
			l_sql = l_sql & " SET evecanrealalu = evecanrealalu - 1"
			l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
	end if
	l_rs.close

	l_sql = "DELETE FROM cap_candidato "
	l_sql = l_sql & " WHERE ternro = " & l_ternro
	l_sql = l_sql & " AND cap_candidato.evenro = " & l_evenro 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
	
	l_sql = " select calnro "
	l_sql = l_sql & " from cap_evento "
	l_sql = l_sql & " inner join cap_calendario on cap_calendario.evenro = cap_evento.evenro "
	l_sql = l_sql & " WHERE cap_evento.evenro = " & l_evenro 
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	   l_rs.MoveFirst
	end if
	Do While Not l_rs.eof
		l_sql = " DELETE FROM cap_part_cal "
		l_sql = l_sql & " WHERE calnro = " & l_rs(0)
		l_sql = l_sql & " AND ternro = "  & l_ternro
'		l_sql = l_sql & "(calnro, ternro) "
'		l_sql = l_sql & " VALUES (" & l_rs(0) & "," & l_ternro & ")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_rs.MoveNext
	Loop
	l_rs.Close
end if

'### Envio un email en todos los casos ###
l_sql = "SELECT ternro, ternom2, empreporta, terape, terape2, empleg, ternom "
l_sql = l_sql & " FROM empleado "
l_sql = l_sql & " WHERE empleado.ternro = " & l_ternro

l_rs.Open l_sql, cn
if l_rs.eof then
    response.write "<script>window.close();</script>"
	response.end
else 
  l_empleg = l_rs("empleg")
  l_ternro = l_rs("ternro")
  l_reportaa = l_rs("empreporta")
  l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2") & " (" & l_empleg & ")"
end if
l_rs.close

l_sql = "SELECT empemail FROM empleado WHERE empleado.empleg = " & l_reportaa
l_rs.Open l_sql, cn
if not l_rs.eof then
	l_empemail = l_rs("empemail")
else
   	response.write "<script>window.close();</script>"
	response.end
end if
l_rs.close

l_sql = "SELECT evedesabr FROM cap_evento WHERE evenro= " & l_evenro
l_rs.Open l_sql, cn
if not l_rs.eof then
	l_evedesabr = l_rs("evedesabr")
end if

cn.Close
Set l_cm = Nothing
%>
<html>
	<head></head>
	<body onload="JavaScript: validar()">
		<form action="../cgi-bin/postmail.exe" method="post" name="FormVar" id="FormVar">
			<input type="text" name="Host" value="<%= cEmailHost %>">
			<input type="text" name="Port" value="<%= cEmailPort %>">
			<input type="text" name="Userid" value="<%= cEmailUserid %>">
			<input type="text" name="Date" value="<%= FormatDateTime(now,2) %>">
			<input type="text" name="FromAddress" value="<%= cEmailFromAddress %>">
			<input type="text" name="FromName" value="<%= cEmailFromName %>">
			<input type="text" name="ReplyTo" value="<%= cEmailReplyTo %>">
			<input type="text" name="ToAddress" value="<%= l_empemail %>">
		<% select case l_tipo %>
				<%	case "C" %>
			<input type="text" name="Subject" value="Alta a eventos de capacitación">
			<input type="text" name="Body" value="El funcionario <%= l_empleado  %> se ingreso como Candidato a un pedido del evento de <%= l_evedesabr %> de capacitación en el Módulo de Autogestión.">
				<% 	case "P" %>
			<input type="text" name="Subject" value="Alta a eventos de capacitación">
			<input type="text" name="Body" value="El funcionario <%= l_empleado  %> se ingreso como Participante a un pedido del evento de <%= l_evedesabr %> de capacitación en el Módulo de Autogestión.">
				<% 	case "Q" %>
			<input type="text" name="Subject" value="Baja a eventos de capacitación">
			<input type="text" name="Body" value="El funcionario <%= l_empleado  %> se dio de Baja a un pedido del evento de <%= l_evedesabr %> de capacitación en el Módulo de Autogestión.">
		<% end select %>
		</form>
	</body>
<script>
function validar(){
	//alert('<%'= cEmailHost&"-"&cEmailPort&"-"&cEmailUserid&"-"&FormatDateTime(now,2)&"-"&cEmailFromAddress&"-"&cEmailFromName&"-"&cEmailReplyTo&"-"&l_empemail%>');
	v = window.open('','enviamail', 'top=2000, left=2000');
	document.FormVar.target="enviamail"
	document.FormVar.submit()
	v.close();
	alert('Operacion realizada.');
	<% If l_tipo = "Q" then %>
	window.opener.location.reload();
	window.close();
	<% Else  %>
	window.opener.location.reload();
	<% End If %>
	window.close();
}
	
</script>
</html>
