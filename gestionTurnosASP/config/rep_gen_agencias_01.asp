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

Dim l_desdes
Dim l_indice_destino
Dim ArrDesNro(100)
Dim ArrDesDes(100)
Dim MatMerDes(100, 100)
Dim l_TotMerDes
Dim l_TotTotMerDes

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

Dim l_indice_exportadora

Dim ArrExpNro(50)
Dim ArrExpDes(50)
			
Dim ArrMerNro(50)
Dim ArrMerDes(50)

Dim MatMerExp(50,50)


Dim l_fila

Function NombreTipoOperacion(nro)

select case nro
case 1
	NombreTipoOperacion = "Carga"
case 2
	NombreTipoOperacion = "Descarga"
case 3
	NombreTipoOperacion = "Exportación"
case 4
	NombreTipoOperacion = "Importación"
end select

end Function



Function NombreMes(nro)

select case nro
case 1
	NombreMes = "Ene"
case 2
	NombreMes = "Feb"
case 3
	NombreMes = "Mar"
case 4
	NombreMes = "Abr"
case 5
	NombreMes = "May"
case 6
	NombreMes = "Jun"
case 7
	NombreMes = "Jul"
case 8
	NombreMes = "Ago"
case 9
	NombreMes = "Sep"
case 10
	NombreMes = "Oct"
case 11
	NombreMes = "Nov"
case 12
	NombreMes = "Dic"

end select

end Function


sub Inicializar_Arreglo(Arr, Lim, Valor)

	for x = 1 to Lim
		Arr(x) = Valor
	next

end sub	


sub encabezado_detatebuqage(titulo)


%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
						&nbsp;
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>
	    <tr>
	        <td align="center" width="30%">&nbsp;</td>					
	        <th align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" nowrap>Agencias</th>			
   		    <th align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;" nowrap >Buques Atendidos</th>	
	        <td align="center" width="30%">&nbsp;</td>								 
	    </tr>

<%
end sub

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

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_nropagina = 1

encabezado_detatebuqage("Detalle de Atención Buques por Agencia") 

l_sql = " SELECT distinct(buq_agencia.agedes), count(*) "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " group by buq_agencia.agedes "

rsOpen l_rs, cn, l_sql, 0

do while not l_rs.eof
%>
		<tr>
        <td align="center" width="30%">&nbsp;</td>							
		<td align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" nowrap ><%= l_rs(0) %></td>			
		<td align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;" nowrap><%= l_rs(1) %></td>					
        <td align="center" width="30%">&nbsp;</td>							
		</tr>		
<%
	l_rs.movenext
loop
%>
		<tr>
		<td align="center" colspan="4">
	  	  <iframe frameborder="0" name="ifrmgra21" scrolling="No" src="gra_21.asp?anioini=<%= l_fecini %>&fecfin=<%= l_fecfin %>" width="600" height="300"></iframe> 
		</td>
		</tr>  		
<%
response.write "</table><p style='page-break-before:always'></p>"
l_rs.close


set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

