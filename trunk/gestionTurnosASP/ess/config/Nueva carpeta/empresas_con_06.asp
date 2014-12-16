<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 

'Archivo: empresas_con_06.asp
'Descripción: ABM de Empresas
'Autor : Gustavo Manfrin
'Fecha: 12/09/2006

Dim l_tipo
Dim l_rs
Dim l_sql

Dim l_locloccod

Dim texto

texto = ""
l_locloccod 	= request.QueryString("locloccod")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Verifico que exista la localidad local.
l_sql = "SELECT locnro FROM tkt_localidad "
l_sql = l_sql & " WHERE loccod ='" & l_locloccod & "'"
rsOpen l_rs, cn, l_sql, 0
if l_rs.eof then
    texto =  "La localidad local no existe."
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

