<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
on error goto 0

'Archivo: embarque_con_07.asp
'Autor : Gustavo Manfrin
'Fecha: 20/09/2006

Dim l_vencorcod
Dim l_vencordes

Dim l_rs
Dim l_sql

l_vencorcod = request.QueryString("vencorcod")

'=====================================================================================
Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
l_sql = "SELECT vencordes "
l_sql = l_sql & "FROM tkt_vencor "
l_sql = l_sql & " WHERE vencorcod = '" & l_vencorcod & "'"
rsOpen l_rs, cn, l_sql, 0
if  not(l_rs.eof) then
	l_vencordes = l_rs(0)
else
	l_vencordes = ""
end if	
l_rs.close
Set l_rs = nothing
%>

<script>
	parent.actualizar_vendedor('<%= l_vencordes %>')
</script>