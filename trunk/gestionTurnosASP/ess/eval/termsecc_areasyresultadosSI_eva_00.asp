<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: termsecc_areasyresultadosSI_eva_00.asp
'Descripción	: verifica si se cargaron resultados para areas y competencias
'Autor			: 01-09-2005 
'Fecha			: LAmadio

'================================================================================

on error goto 0
Dim l_rs
Dim l_sql

dim l_terminarsecc
dim l_evacabnro
dim l_evaseccnro
dim l_evatevnro
dim l_evldrnro
dim l_habCalifArea

  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro") 
  l_evacabnro  = Request.QueryString("evacabnro")
  l_evatevnro  = Request.QueryString("evatevnro")
  'l_habCalifArea = Request.QueryString("habCalifArea")
  
' -------------------------------------------------------------------
' BODY --------------------------------------------------------------
' -------------------------------------------------------------------

l_terminarsecc = "SI" 
'response.write "terminar " &l_terminarsecc & "<br>"

' se fija si se evaluaron todas las comptencias
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaresultado.evatrnro "   ',evadetevldor.evldrnro 
l_sql = l_sql & " FROM evaresultado "
l_sql = l_sql & " WHERE evldrnro="& l_evldrnro
l_sql = l_sql & "    AND evaresultado.evatrnro IS NULL "
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then 
	l_terminarsecc = "NO"
end if 
l_rs.Close
set l_rs = Nothing


'response.write "terminar1 " &l_terminarsecc & "<br>"     --  and  cint(l_evatevnro) <> cint(caconsejado)
	'response.write l_terminarsecc
	' response.write l_sql & "<br>"
if l_terminarsecc = "SI" and ( cint(l_evatevnro) <> cint(cautoevaluador))  then  
		'response.write "terminar2 " &l_terminarsecc & "<br>"
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evaarea.evatrnro "
		l_sql = l_sql & " FROM evaarea "
		l_sql = l_sql & " WHERE evldrnro="& l_evldrnro
		l_sql = l_sql & "   AND evaarea.evatrnro IS NULL "
		'response.write l_sql
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.EOF then 
			l_terminarsecc = "NO"
		end if
		l_rs.Close
		set l_rs = Nothing
end if

response.write l_terminarsecc
'response.write l_sql
'response.write "terminar 3" &l_terminarsecc & "<br>"
if l_terminarsecc = "NO" then
	Response.write "<script>parent.document.datos.terminarsecc2.value='NO';window.close();</script>"
else
	Response.write "<script>parent.document.datos.terminarsecc2.value='SI'; window.close();</script>" 'xx  reload???
	'Response.write "<script>parent.document.location.reload(); window.close();</script>" ' cicla --!!!!!!!
end if
%>
