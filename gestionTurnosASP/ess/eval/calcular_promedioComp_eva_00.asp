<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<script>
var xc = screen.availWidth;
var yc = screen.availHeight;
window.moveTo(xc,yc);	
</script>
<%
'=====================================================================================
'Archivo  : calcular_promediocomp_eva_00.asp
'Objetivo : calcular promedio competencias que NO sean potenciales
'Fecha	  : 30-11-2004
'Autor	  : CCRossi
'Modificacion: 
'=====================================================================================

dim l_Error
Dim l_rs_oblig
dim l_rs
dim l_sql

dim l_suma
dim l_cantidad
dim l_promedio

'parametros de entrada
dim l_evldrnro
dim l_evaseccnro

l_evldrnro   = Request.QueryString("evldrnro")
l_evaseccnro = Request.QueryString("evaseccnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_Error = 0


'=========================================================================================
' Calcular Suma Total
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT SUM(evatrvalor) AS suma "
 l_sql = l_sql & " FROM  evaresultado "
 l_sql = l_sql & " INNER JOIN evatipresu    ON evatipresu.evatrnro = evaresultado.evatrnro   "
 l_sql = l_sql & " INNER JOIN evaseccfactor ON evaseccfactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & " INNER JOIN evafactor		ON evafactor.evafacnro = evaresultado.evafacnro "
 l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
 l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro 
 l_sql = l_sql & " AND   evaresultado.evldrnro    = " & l_evldrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then 
	l_suma = l_rs("suma")
 end if
 l_rs.Close

'response.write l_maximo & "<br>"
  l_sql = "SELECT COUNT(evaseccfactor.evafacnro) AS cantidad "
  l_sql = l_sql & " FROM evaseccfactor "
  l_sql = l_sql & " INNER JOIN evafactor ON evafactor.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & "		AND evafactor.evafacpot = 0 "
  l_sql = l_sql & " WHERE evaseccfactor.evaseccnro = " & l_evaseccnro
  l_sql = l_sql & " AND EXISTS "
  l_sql = l_sql & " (SELECT * FROM evaresultado "
  l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evaresultado.evatrnro "
  l_sql = l_sql & " AND evatipresu.evatrnro IS NOT NULL "
  l_sql = l_sql & " WHERE evaresultado.evafacnro = evaseccfactor.evafacnro "
  l_sql = l_sql & " AND   evaresultado.evldrnro = " & l_evldrnro
  l_sql = l_sql & ")"
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then 
	l_cantidad = l_rs("cantidad")
  end if
  l_rs.Close

  if trim(l_cantidad)<>"" and not isnull(l_cantidad) and l_cantidad<>0 then
	l_promedio = cdbl(l_suma) / cdbl(l_cantidad)
	
  end if
  %>
	<script>
		
		window.returnValue='<%= round(l_promedio,2)%>';
		window.close();
	</script>




%>