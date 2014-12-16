<% Option Explicit %>
<!--#include virtual="/intranet/shared/inc/sec.inc"-->
<!--#include virtual="/intranet/shared/inc/const.inc"-->
<!--#include virtual="/intranet/shared/db/conn_db.inc"-->
<% 



Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_id
Dim l_descripcion


Dim texto

texto = ""
l_tipo		    = request.QueryString("tipo")
l_id            = request.QueryString("id")
l_descripcion 	= request.QueryString("descripcion")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo
l_sql = "SELECT descripcion"
l_sql = l_sql & " FROM practicas "
l_sql = l_sql & " WHERE descripcion='" & l_descripcion & "'"
if l_tipo = "M" then
	l_sql = l_sql & " AND id <> " & l_id
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otra Práctica con esa Descripción."
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

