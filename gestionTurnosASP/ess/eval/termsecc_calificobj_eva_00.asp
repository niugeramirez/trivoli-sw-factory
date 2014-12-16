<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'================================================================================
'Archivo		: termsecc_areasyresultados_eva_00.asp
'Descripción	: verifica si se cargaron resultados para objetivos
'Autor			: 19-04-2005 
'Fecha			: LAmadio

'================================================================================
'     xxxxxxxxxxxx
on error goto 0
Dim l_rs
Dim l_sql

dim l_terminarsecc
dim l_evacabnro
dim l_evaseccnro
dim l_evatevnro
dim l_evldrnro
'dim l_habCalifGral

  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro") 
  l_evacabnro  = Request.QueryString("evacabnro")
  l_evatevnro  = Request.QueryString("evatevnro")
  'l_habCalifGral = Request.QueryString("habCalifGral")
  
' -------------------------------------------------------------------
' BODY --------------------------------------------------------------
' -------------------------------------------------------------------
' me fijo si se evaluaron los objetivos.
l_terminarsecc = "SI"
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT evaluaobj.evatrnro, evadetevldor.evldrnro "
l_sql = l_sql & " FROM evaluaobj "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evaluaobj.evldrnro "
l_sql = l_sql & " WHERE evadetevldor.evacabnro="& l_evacabnro
l_sql = l_sql & "    AND evadetevldor.evaseccnro="& l_evaseccnro ' calific obj 
l_sql = l_sql & "	 AND evadetevldor.evatevnro="& l_evatevnro
l_sql = l_sql & "    AND evaluaobj.evatrnro IS NULL "
'response.write l_sql & "<br>"
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.EOF then 
	l_terminarsecc = "NO"
end if
l_rs.Close
set l_rs = Nothing

response.write l_terminarsecc 
	'response.write l_terminarsecc & " -----------"
	'response.write l_sql & "<br>"
	'response.write l_habCalifGral & "<br>"
	'response.write "evento: " & l_evatevnro &" +++++ " & cautoevaluador

	' if (cint(l_evatevnro) <> cint(cautoevaluador))then response.write "distintoss" end if
if (l_terminarsecc = "SI") and (  cint(l_evatevnro) <> cint(cautoevaluador)  and  cint(l_evatevnro) <> cint(caconsejado))  then  'and l_habCalifArea="SI"
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT evagralobj.evatrnro "
		l_sql = l_sql & " FROM evagralobj "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro = evagralobj.evldrnro "
		l_sql = l_sql & " WHERE evadetevldor.evacabnro="& l_evacabnro
		l_sql = l_sql & "   AND ( evadetevldor.evatevnro="& cevaluador  & " OR evadetevldor.evatevnro=" & cconsejero & " ) "
		l_sql = l_sql & "   AND evagralobj.evatrnro IS NULL "
		'response.write l_sql & "<br>"
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.EOF then 
			l_terminarsecc = "NO"
		end if
		l_rs.Close
		set l_rs = Nothing
end if

'response.write l_terminarsecc 
if l_terminarsecc= "NO" then
	Response.write "<script>parent.document.datos.terminarsecc2.value='NO';window.close();</script>"
else
	Response.write "<script>parent.document.datos.terminarsecc2.value='SI';</script>"
	Response.write "<script>window.close();</script>"
end if
%>
