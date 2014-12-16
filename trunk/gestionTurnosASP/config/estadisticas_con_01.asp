<% Option Explicit
'if request.querystring("excel") then
'	Response.AddHeader "Content-Disposition", "attachment;filename=Estadisticas.xls" 
'	Response.ContentType = "application/vnd.ms-excel"
'end if
 %>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 53
Const l_cantcols = 10
Const l_empresa = "Cámara Portuaria y Marítima de Bahía Blanca"

Dim l_indice
Dim CAS(50)
Dim INF(50)

Dim l_porcentaje
Dim l_rs
Dim l_rs2
Dim l_buqdes
Dim l_canbuq
Dim l_totton

Dim l_sql
dim primero
dim ultimo

Dim l_enHP9000

Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_total 

dim l_fecini
dim l_fecfin
dim l_feciniHP
dim l_fecfinHP
dim l_anulado

dim l_movcod
dim l_operacion

dim l_lugar

dim l_repelegido
Dim l_indice_mercaderia
Dim i

'Variable usadas para imprimir los Totales
dim l_nroope


'---------------
' rep4
Dim l_indmer
Dim l_totfil
Dim TOTCAS(100)
Dim l_totcaston
Dim y
Dim l_expdes
Dim x
Dim l_existe
Dim l_ColMer
Dim l_TotMerExp
Dim l_TotTotMerExp
'---------------

'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")

l_repelegido  = request.querystring("repnro")

'l_anioini = "01/01/" & year(l_fecfin)



Dim l_fila


sub fin_encabezado
%>
</table>	
<%
end sub 'finencabezado


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<%
if request.querystring("excel") <> "true" then
%>

<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">


<%
end if
%>


<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<table>
    <tr>
        <td align="center" width="30%">&nbsp;</td>					
        <th align="center" width="20%" nowrap>Servicio Local</th>			
	    <th align="center" width="20%" nowrap >Legajos Atendidos</th>	
        <td align="center" width="30%">&nbsp;</td>								 
    </tr>
<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_nropagina = 1



Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT distinct(ser_legajo.legpar1), ser_servicio.serdes, count(*) "
l_sql = l_sql & " FROM ser_legajo "
l_sql = l_sql & " INNER JOIN ser_servicio ON ser_legajo.legpar1 = ser_servicio.sercod "
l_sql = l_sql & " WHERE ser_legajo.legfecing >= " & cambiafecha(l_fecini,"YMD",true)
l_sql = l_sql & " AND ser_legajo.legfecing <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " group by ser_legajo.legpar1 , ser_servicio.serdes"

rsOpen l_rs, cn, l_sql, 0

do while not l_rs.eof
%>
		<tr>
        <td align="center" width="30%">&nbsp;</td>							
		<td align="center" width="20%" nowrap ><%= l_rs(1) %></td>			
		<td align="center" width="20%" nowrap><%= l_rs(2) %></td>					
        <td align="center" width="30%">&nbsp;</td>							
		</tr>		
<%
	l_rs.movenext
loop
%>
		<tr>
		<td align="center" colspan="4">
	  	  <iframe frameborder="0" name="ifrmgra21" scrolling="No" src="gra_21.asp?fecini=<%= l_fecini %>&fecfin=<%= l_fecfin %>" width="600" height="300"></iframe> 
		</td>
		</tr>  		
<%
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1



set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

