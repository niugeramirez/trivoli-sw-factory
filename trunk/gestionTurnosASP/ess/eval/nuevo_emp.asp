<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_empleg
Dim l_terape
Dim l_ternom
Dim l_nuevoleg
Dim l_tabla

l_empleg = request.querystring("empleg")
l_tabla = request.querystring("tabla")

if len(trim(l_tabla))=0 then
	l_tabla= "empleado"
end if
%>
<%

if trim(l_empleg)<>"" then
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = "SELECT  ternro, empleg, terape, ternom FROM "& l_tabla &" where empleg = " & l_empleg & " "
	rsOpen l_rs, cn, l_sql, 0
	
	on error resume next
	
	if not l_rs.eof then
		l_ternro = l_rs("ternro")
		l_nuevoleg = l_rs("empleg")
		l_terape = l_rs("terape")
		l_ternom = l_rs("ternom")
	else
		l_nuevoleg = "0"
		l_ternro = "0"
	end if
	l_rs.Close
	set l_rs=nothing
else
	l_nuevoleg = "0"
	l_ternro = "0"
end if
Set cn = Nothing

if err then
	l_nuevoleg = "0"
	l_ternro = "0"
end if

on error goto 0
%>

<script>
window.opener.nuevoempleado(<%= l_ternro %>,<%= l_nuevoleg %>,"<%= l_terape %>","<%= l_ternom %>");
window.close();
</script>
