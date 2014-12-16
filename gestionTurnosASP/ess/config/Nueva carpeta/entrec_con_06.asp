<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: entrec_con_06.asp
'Descripción: Abm de Entregadores y recibidores
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_entnro
Dim l_entcod
Dim l_entdes
Dim l_foco

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_entnro    = request.QueryString("entnro")
l_entcod 	= request.QueryString("entcod")
l_entdes 	= request.QueryString("entdes")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetida la descripción o el código externo

l_sql = "SELECT entcod"
l_sql = l_sql & " FROM tkt_entrec "
l_sql = l_sql & " WHERE entcod='" & Trim(l_entcod) & "' AND entact=-1"
if l_tipo = "M" then
	l_sql = l_sql & " AND entnro <> " & l_entnro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe otro Entregador/Recibidor con ese Código."
	l_foco = "parent.document.datos.entcod.focus();"
else
	l_rs.close
	l_sql = "SELECT entdes "
	l_sql = l_sql & " FROM tkt_entrec "
	l_sql = l_sql & " WHERE entdes='" & l_entdes & "'"
	if l_tipo = "M" then
		l_sql = l_sql & " AND entnro <> " & l_entnro
	end if
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		texto =  "Ya existe otro Entregador/Recibidor con esa Descripción."
		l_foco = "parent.document.datos.entdes.focus();"
	end if
end if 
l_rs.close
%>

<script>
<% 
 if texto <> "" then
%>
   parent.invalido('<%= texto %>','<%= l_foco %>')
<% else%>
   parent.valido();
<% end if%>
</script>

<%
Set l_rs = Nothing
%>

