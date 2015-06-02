<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<% 
'Archivo: recursosreservables_con_04.asp
'Descripción: Script Baja Medicos
'Autor : Trivoli
'Fecha: 31/05/2015

'on error goto 0
Dim l_cm
Dim l_rs
Dim l_sql
Dim l_id
	
l_id = request.querystring("cabnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_cm = Server.CreateObject("ADODB.Command")

l_sql = "DELETE FROM recursosreservables  WHERE id = " & l_id

l_cm.activeconnection = Cn
l_cm.CommandText = l_sql
cmExecute l_cm, l_sql, 0

cn.Close
Set cn = Nothing
%>
<script>
	alert('Operación Realizada.');
    window.opener.parent.ifrm.location.reload();
	window.close();
</script>