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
Dim l_nrohistoriaclinica


Dim texto

texto = ""
l_tipo		    	 = request.QueryString("tipo")
l_id            	 = request.QueryString("id")
l_dni 				 = request.QueryString("dni")
l_nrohistoriaclinica = request.QueryString("nrohistoriaclinica") 

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripci�n o el c�digo externo
l_sql = "SELECT * "
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " WHERE dni=" & l_dni 
l_sql = l_sql & " AND dni <> 0 " 
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
l_sql = l_sql & " and clientespacientes.empnro = " & Session("empnro")   
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Paciente con el Nro de DNI ingresado."
end if 
l_rs.close

'Verifico que no este repetida el nro de historia clinica
l_sql = "SELECT * "
l_sql = l_sql & " FROM clientespacientes "
l_sql = l_sql & " WHERE nrohistoriaclinica='" & l_nrohistoriaclinica & "'" 
l_sql = l_sql & " AND nrohistoriaclinica <> '0'"
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
l_sql = l_sql & " and clientespacientes.empnro = " & Session("empnro")   
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Paciente con el Nro de Historia Clinica ingresado."
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

