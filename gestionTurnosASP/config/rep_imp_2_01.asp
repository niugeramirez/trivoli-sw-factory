<% Option Explicit %>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 53
Const l_cantcols = 10
Const l_empresa = "Cámara Portuaria y Marítima de Bahía Blanca"


Dim l_porcentaje
Dim l_rs
Dim l_rs2
Dim l_buqdes
Dim l_canbuq
Dim l_totton

Dim l_sql
dim primero
dim ultimo


Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_total 

dim l_fecini
dim l_fecfin

Dim ArrMerNro(100)
Dim ArrMerDes(100)
Dim l_indice
Dim i
Dim l_feciniant
Dim l_fecfinant
Dim l_valor

'Variable usadas para imprimir los Totales
dim l_nroope
dim l_anioini
dim l_sitnro

' Imprime los Totales


'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_sitnro      = request.querystring("qsitnro")

'l_repelegido  = request.querystring("repnro")

l_anioini = "01/01/" & year(l_fecfin)


Dim l_indice_exportadora

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

sub encabezado_expbuq(titulo)
%>
	<table style="width:300" cellpadding="0" cellspacing="0" >
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td align="left" width="100%" colspan="7">
							<b>* <%= titulo%></b> 
						</td>
				       	<td align="right" nowrap width="5%" > 
							<!--P&aacute;gina: <%'= l_nropagina%> -->
						</td>				
					</tr>
					<!--
					<tr>
						<td align="left" width="100%" colspan="7">
							<%'= l_fecini  %>&nbsp;-&nbsp;<%'= l_fecfin %>
						</td>
				       	<td align="right" nowrap width="5%" > 
							&nbsp;
						</td>										
					</tr>
					<tr>
				       	<td nowrap colspan="8">&nbsp;
						</td>				
					</tr>
					-->														
				</table>
			</td>				
		</tr>
<%
end sub 'encabezado


sub encabezado_sitio(titulo)

Dim l_nombre_sitio

l_sql = " SELECT buq_sitio.sitdes "
l_sql = l_sql & " FROM buq_sitio  "
l_sql = l_sql & " WHERE buq_sitio.sitnro = " & l_sitnro
rsOpen l_rs, cn, l_sql, 0
if not l_Rs.eof then
	l_nombre_sitio = l_rs(0)
else
	l_nombre_sitio = "Todos"
end if 
l_rs.close

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" >
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td align="left" width="100%" colspan="7">
							<b>Sitio: <%= 	l_nombre_sitio %></b>&nbsp;&nbsp;&nbsp;<%= l_fecini %> &nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap width="5%" > 
							<!--P&aacute;gina: <%'= l_nropagina%> -->
						</td>				
					</tr>
				</table>
			</td>				
		</tr>
<%
end sub 'encabezado


sub fin_encabezado
%>
</table>	
<%
end sub 'finencabezado


%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

encabezado_sitio("")

encabezado_expbuq("PRODUCTO") 

l_sql = " SELECT buq_mercaderia.merdes , sum(conton), buq_mercaderia.mernro "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_mercaderia ON buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
if l_sitnro <> "" then
	l_sql = l_sql & "  AND buq_contenido.sitnro = " & l_sitnro
end if
l_sql = l_sql & " group by buq_mercaderia.merdes, buq_mercaderia.mernro "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0
'response.write l_sql

l_totton = 0
%>
	<tr>
		<th align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Producto</th>
		<th align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%
dIM l_cadena4
l_indice = 0
l_cadena4 = ""
do until l_rs.eof
	ArrMerNro(l_indice) = l_rs(2)
	ArrMerDes(l_indice) = l_rs(0)
	l_indice = l_indice + 1
%>
	<tr>
		<td align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(0) %></td>
		<td align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>		
    </tr>
<%
	l_cadena4 = l_cadena4 & l_rs(0) & "-" & l_rs(1) & ","
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="150" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<tr>
	<td align="center" colspan="6">
  	  <iframe frameborder="0" name="ifrmgra15" scrolling="No" src="gra_15.asp?cadena=<%= l_cadena4 %>" width="600" height="200"></iframe> 
	</td>
</tr> 			
<%
fin_encabezado
l_rs.close

encabezado_expbuq("DESTINO") 

l_sql = " SELECT buq_destino.desdes , sum(conton) "
l_sql = l_sql & " FROM buq_buque " 
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_destino ON buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
if l_sitnro <> "" then
	l_sql = l_sql & "  AND buq_contenido.sitnro = " & l_sitnro
end if
l_sql = l_sql & " group by buq_destino.desdes "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Sitio</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

encabezado_expbuq("COMPARACION MES/AÑO ANTERIOR") 

if month(l_fecini) < 10 then
	l_feciniant = "01/0" & month(l_fecini) & "/" & year(l_fecini) - 1
else
	l_feciniant = "01/" & month(l_fecini) & "/" & year(l_fecini) - 1
end if
l_fecfinant = cdate("01/0" & month(l_fecini) + 1 & "/" & year(l_fecini) - 1) -1
'response.write l_feciniant & " - " & l_fecfinant
%>
<tr>
	<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></th>
	<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= month(l_fecfinant) %>/<%=  year(l_fecfinant) %></b></th>
	<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= month(l_fecfin) %>/<%= year(l_fecfin) %></b></th>
</tr>
<%
Dim l_cadena6
Dim l_cadena7
l_cadena6 = ""
l_cadena7 = ""
for i = 0 to l_indice - 1
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= ArrMerDes(i) %></td>	
<%
	l_sql = " SELECT sum(conton) "
	l_sql = l_sql & " FROM buq_buque " 
	l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_feciniant,"YMD",true)  
	l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfinant,"YMD",true) 
	l_sql = l_sql & " AND buq_buque.tipopenro = 4  "
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_sitnro
	l_sql = l_sql & " and buq_contenido.mernro = " & ArrMerNro(i)
	rsOpen l_rs, cn, l_sql, 0
	if l_Rs.eof then
		l_valor = 0
	else
		if isnull(l_rs(0)) then
			l_valor = 0
		else
			l_valor = l_rs(0)
		end if 	
	end if 
	l_rs.close
	
	l_cadena6 = l_cadena6 & left(ArrMerDes(i),7) & "-" & l_valor & ","
	
	%>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_valor %></td>
	<%	

	l_sql = " SELECT sum(conton) "
	l_sql = l_sql & " FROM buq_buque " 
	l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
	l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_buque.tipopenro = 4  "
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_sitnro
	l_sql = l_sql & " and buq_contenido.mernro = " & ArrMerNro(i)
	rsOpen l_rs, cn, l_sql, 0
	if l_Rs.eof then
		l_valor = 0
	else
		l_valor = l_rs(0)
	end if 
	%>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_valor %></td>
    </tr>
	<%
	l_rs.close
	
	l_cadena7 = l_cadena7 & left(ArrMerDes(i),7) & "-" & l_valor & ","
	
next
%>
<tr>
	<td align="center" colspan="6">
  	  <iframe frameborder="0" name="ifrmgra18" scrolling="No" src="grafico_comparativo_mercaderia.asp?cadena=<%= l_cadena6 %>&cadena2=<%= l_cadena7 %>" width="720" height="350"></iframe> 
	</td>
</tr> 			
<%
response.end



l_sql = " SELECT buq_destino.desdes , sum(conton) "
l_sql = l_sql & " FROM buq_buque " 
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_destino ON buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 4 "
if l_sitnro <> "" then
	l_sql = l_sql & "  AND buq_contenido.sitnro = " & l_sitnro
end if
l_sql = l_sql & " group by buq_destino.desdes "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Sitio</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close
%>

<tr>
	<td align="center" colspan="6">
  	  <iframe frameborder="0" name="ifrmgra18" scrolling="No" src="gra_18.asp?cadena=<%= l_cadena6 %>&cadena2=<%= l_cadena7 %>" width="720" height="350"></iframe> 
	</td>
</tr> 			
<%
response.end

encabezado_expbuq("EXPORTADORA") 

l_sql = " SELECT buq_exportadora.expdes , sum(conton) "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_contenido ON buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " INNER JOIN buq_exportadora ON buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 3 "
l_sql = l_sql & " group by buq_exportadora.expdes "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Exportadora</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close


encabezado_expbuq("AGENCIA") 

l_sql = " SELECT buq_agencia.agedes , sum(conton) "
l_sql = l_sql & " FROM buq_buque  "
l_sql = l_sql & " INNER JOIN buq_agencia ON buq_agencia.agenro = buq_buque.agenro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & "  AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & "  AND buq_buque.tipopenro = 3 "
l_sql = l_sql & " group by buq_agencia.agedes "
l_sql = l_sql & " order by 2 desc "

rsOpen l_rs, cn, l_sql, 0

l_totton = 0
%>
	<tr>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Exportadora</th>
		<th align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">Toneladas</th>
    </tr>
<%	
do until l_rs.eof
%>
	<tr>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(0) %></td>
		<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_rs(1) %></td>
    </tr>
<%
	l_totton = l_totton + l_rs(1)
	l_rs.MoveNext
loop
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b>Total</b></td>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
fin_encabezado
l_rs.close

response.end



l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Cantidad de Buques</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><b><%= l_canbuq %></b></td>
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

