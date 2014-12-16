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

Dim l_cadena
Dim l_cadena2
Dim l_cadena3
Dim l_cadena4
Dim l_cadena5
Dim l_cadena6
Dim l_cadena7
Dim l_cadena80


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


'Obtengo los parametros
l_buqdes 	  = request.querystring("qbuqdes")

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


sub encabezado_expbuq(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
				<!--
					<tr>
				       	<td colspan="8">&nbsp;
						</td>				
					</tr>												
					<tr>
				       	<td width="5%" style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="1"><img src="/serviciolocal/shared/images/puerto.gif" border="0"></td>
				       	<td width="5%" style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="1">Cámara Portuaria y Marítima <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; de Bahía Blanca</td>
				       	<td width="90%"style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="6">&nbsp;</td>													
						</td>				
					</tr>												

					<tr>
				       	<td colspan="8">&nbsp;
						</td>				
					</tr>					
					-->
					<tr>
						<td align="center" width="100%" colspan="7">
							<b><%= titulo%></b> <%= l_buqdes %>
						</td>
						<!--
				       	<td align="right" nowrap width="5%" > 
							P&aacute;gina: <%'= l_nropagina%>
						</td>				
						-->
					</tr>
					<tr>
				       	<td nowrap colspan="8">&nbsp;
						</td>				
					</tr>									
				</table>
			</td>				
		</tr>
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Buque</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Comenzó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Terminó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Toneladas</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Mercadería</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Sitio</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Agencia</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Destino</th>																
	    </tr>		
<%
end sub 'encabezado

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
</script>
	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_nropagina = 1

'l_nropagina = 1
encabezado_expbuq("Detalle del Buque:" ) 
l_nrolinea = 6

'l_encabezado = true
'l_corte = false
'l_total = 0


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 3 "  ' EXPORTACION
'l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
'l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqdes = '" & l_buqdes & "'"
l_sql = l_sql & " ORDER BY buq_buque.buqdes, buq_buque.buqfechas, buq_buque.buqfecdes " 
rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		if l_nrolinea > l_Max_Lineas_X_Pag then
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina = l_nropagina + 1
			encabezado_expbuq("Exportación - Detalle de Buques") 
			l_nrolinea = 6
		end if
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%" nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;&nbsp;&nbsp;<%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("buqfechas") %></td>
			<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("conton") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("merdes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("sitdes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("agedes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs("desdes") %></td>
	    </tr>
		<%
		l_nrolinea = l_nrolinea + 1
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop





l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >&nbsp;</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;</td>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Total Toneladas</td>				
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><b><%= l_totton %></b></td>
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;">&nbsp;</td>			
</tr>
<%

l_nrolinea = l_nrolinea + 1
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

