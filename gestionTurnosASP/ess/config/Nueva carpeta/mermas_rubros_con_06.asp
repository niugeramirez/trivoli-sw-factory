<% Option Explicit %>
<!--#include virtual="/ticket/shared/inc/sec.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<% 

'Archivo: mermas_rubros_con_06.asp
'Descripción: ABM de Tipos de Mermas para rubros
'Autor : Alvaro Bayon
'Fecha: 16/02/2005

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_tipmernro
Dim l_lugnro
Dim l_rubnro

Dim texto

texto = ""
l_tipo		= request.QueryString("tipo")
l_tipmernro = request.QueryString("tipmernro")
l_lugnro 	= request.QueryString("lugnro")
l_rubnro 	= request.QueryString("rubnro")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que no este repetido el rubro para este lugar
l_sql = "SELECT tipmernro"
l_sql = l_sql & " FROM tkt_tipomerma "
l_sql = l_sql & " WHERE rubnro=" & l_rubnro & " AND lugnro = " & l_lugnro
if l_tipo = "M" then
	l_sql = l_sql & " AND tipmernro <> " & l_tipmernro
end if
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
    texto =  "Ya existe este Rubro en el lugar especificado."
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

