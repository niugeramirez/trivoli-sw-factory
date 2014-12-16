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
'Archivo  : calcular_totallinea_eva_00.asp
'Objetivo : calcular promedio competencias que NO sean potenciales
'Fecha	  : Nov-2005
'Autor	  : CCRossi
'Modificacion:   - LA - arreglo de calculos
'=====================================================================================

dim l_Error

dim l_rs
dim l_sql

dim l_evatrvalor
dim l_totallinea
dim l_promedio

'parametros de entrada
dim l_evatrnro
dim l_evafacpor

l_evatrnro   = Request.QueryString("evatrnro")
l_evafacpor = Request.QueryString("ponderacion")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_Error = 0


'=========================================================================================
' Calcular Suma Total
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = " SELECT evatrvalor FROM evatipresu WHERE evatrnro= " & l_evatrnro
 rsOpen l_rs, cn, l_sql, 0
 if not l_rs.eof then 
	l_evatrvalor = l_rs("evatrvalor")
 end if
 l_rs.Close

'response.write l_maximo & "<br>"

if  isnull(l_evafacpor) or trim(l_evafacpor)="" then
l_totallinea = 0  '	cdbl(l_evatrvalor) si es nulo, se toma como que el porcentaje asignado a esa competencia es 0
else
l_totallinea = cdbl(l_evatrvalor) * cdbl(l_evafacpor) / 100
end if
  %>
	<script>
		window.returnValue='<%= round(l_totallinea,2)%>';
		window.close();
	</script>



