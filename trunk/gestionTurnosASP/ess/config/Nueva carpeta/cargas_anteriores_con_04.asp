<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<% 
'Archivo: cargas_anteriores_con_04.asp
'Descripción: Abm de cargas anteriores
'Autor : Gustavo Manfrin
'Fecha: 07/08/2006

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_carconnro
Dim l_pronro
	
l_carconnro = request.querystring("cabnro")
Set l_rs = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")
	l_cm.activeconnection = Cn

l_sql = " DELETE FROM tkt_cargasconf WHERE carconnro = " & l_carconnro
l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
	window.opener.ifrm.location.reload();
	window.close();
</script>
	




