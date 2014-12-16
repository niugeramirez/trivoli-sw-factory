<% Option Explicit %>
<!--#include virtual="/ticket/shared/db/conn_db.inc"-->
<!--#include virtual="/ticket/shared/inc/const.inc"-->
<% 
'Archivo: depositos_con_04.asp
'Descripción: ABM de Depósitos
'Autor : Alvaro Bayon
'Fecha: 11/02/2005

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_cabnro
	
l_cabnro = request.querystring("cabnro")
set l_cm = Server.CreateObject("ADODB.Command")
l_sql = " DELETE FROM tkt_tipomerma WHERE tipmernro = " & l_cabnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

set l_cm = nothing
cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
	window.opener.ifrm.location.reload();
	window.close();
</script>
	




