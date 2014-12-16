<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/adovbs.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/asistente.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
window.resizeTo(150,150);
</script>
<%
'========================================================================================
'Archivo	: etapa_cabecera_eva_01.asp
'Descripción: grabar cambio de etapa de evacab
'Autor		: CCRossi
'Fecha		: 31-05-2004
'========================================================================================


'ADO
 Dim l_cm
 Dim l_sql
 Dim l_rs
 
'parametros
 Dim l_evacabnro
 Dim l_evaetanro
 
 l_evacabnro  = Request.Form("evacabnro")
 l_evaetanro  = Request.Form("evaetanro")
 
'uso local
 l_sql = " UPDATE evacab  "
 l_sql = l_sql & " SET evaetanro = " & l_evaetanro
 l_sql = l_sql & " WHERE evacabnro = " & l_evacabnro
 set l_cm = Server.CreateObject("ADODB.Command")  
 l_cm.activeconnection = Cn
 l_cm.CommandText = l_sql
 cmExecute l_cm, l_sql, 0
 
 Set cn = Nothing
 Set l_cm = Nothing
 response.write "<script>alert('Operación Realizada.');opener.window.location.reload();window.close();</script>"
%> 