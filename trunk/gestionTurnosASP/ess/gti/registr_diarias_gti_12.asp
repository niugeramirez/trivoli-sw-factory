<%Option Explicit %>
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : registr_diarias_12.asp
Descripcion    : Controla que no se carguen registraciones repetidas
Creador        : Cristian Tetaseca
Fecha Creacion : 07/01/2004
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0
const l_valornulo = "null"

dim l_tipo

dim l_rs
dim l_rs1
dim l_sql

Dim l_regnro
Dim l_regfecha
Dim l_reghora 
Dim l_regentsal
Dim l_relnro
Dim l_ternro
Dim l_hayerror

l_tipo 		= request.querystring("tipo")
l_regnro 	= request.Form("regnro")
if l_regnro = "" then 
	l_regnro = "0"
end if
l_ternro 	= request.Form("ternro") 
l_relnro 	= request.Form("relnro")
l_regfecha = request.Form("regfecha")
l_regentsal = request.Form("regentsal")

if l_relnro = "" then
	l_relnro = l_valornulo
end if

if l_regentsal = "D" then
	l_regentsal = l_valornulo
else 
	l_regentsal = "'" & l_regentsal & "'"
end if
l_reghora = request.Form("reghora1") & request.Form("reghora2")

Set l_rs = Server.CreateObject("ADODB.RecordSet")	

'Busco los datos del tercero

l_sql = "SELECT * "
l_sql = l_sql & " FROM gti_registracion"
l_sql = l_sql & " WHERE regnro <> "   & l_regnro
l_sql = l_sql & "   AND ternro = "     & l_ternro
l_sql = l_sql & "   AND relnro = "   & l_relnro
l_sql = l_sql & "   AND reghora = "   & l_reghora
if l_regentsal = l_valornulo then 
	l_sql = l_sql & "   AND regentsal IS NULL "
else 
	l_sql = l_sql & "   AND regentsal = "   & l_regentsal
end if 
l_sql = l_sql & "   AND regfecha = " & cambiafecha(l_regfecha,"YMD",true)

rsOpen l_rs, cn, l_sql, 0 

l_hayerror = (not l_rs.eof)

l_rs.close

set l_rs = nothing

if l_hayerror then
%>
<script>
 parent.DatosIncorrectos('Registracion Existente.');
</script>
<%
else
%>
<script>
 parent.DatosCorrectos();
</script>
<%
end if
%>

