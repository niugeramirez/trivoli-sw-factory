<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 



Dim l_tipo

Dim l_rs
Dim l_sql

Dim l_id
Dim l_dni


Dim texto

texto = ""
l_tipo		    = request.QueryString("tipo")
l_id            = request.QueryString("id")
l_dni 	= request.QueryString("dni")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

if l_dni = "" then
	texto = ""
else
	'Verifico que no este repetida la descripción o el código externo
	l_sql = "SELECT * "
	l_sql = l_sql & " FROM clientespacientes "
	l_sql = l_sql & " WHERE dni=" & l_dni 
	if l_tipo = "M" then
		l_sql = l_sql & " AND id <> " & l_id
	end if
	l_sql  = l_sql  & " AND clientespacientes.empnro = " & Session("empnro")
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
	    texto =  "Ya existe otro Paciente con el Nro de DNI ingresado."
	end if 
	l_rs.close
end if	
%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<%= texto %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

