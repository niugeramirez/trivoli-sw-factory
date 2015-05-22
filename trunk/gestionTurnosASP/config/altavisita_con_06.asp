<% Option Explicit %>

<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<% 

Dim l_pacienteid

Dim l_rs
Dim l_sql

Dim texto

l_pacienteid = request("pacienteid")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT * "
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " WHERE id=" & l_pacienteid
rsOpen l_rs, cn, l_sql, 0
texto = ""
if not l_rs.eof then
	if l_rs("dni") = "" or l_rs("nrohistoriaclinica") = "0" or l_rs("nrohistoriaclinica") = "" then
    	texto =  "El Paciente seleccionado no tiene DNI o Nro de Historia Clinica cargado. Ir a la opcion Pacientes para completar esta informacion"
	end if
end if 
l_rs.close
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