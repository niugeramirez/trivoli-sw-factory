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


sub encabezado_expcasdes(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table cellpadding="0" cellspacing="0">
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
						<!--
				       	<td align="right" nowrap > 
							P&aacute;gina: <%'= l_nropagina%>
						</td>				
						-->
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
					<tr>
				       	<td nowrap colspan="3">&nbsp;
						</td>				
					</tr>					
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Empresas</th>			
			<%  
			'if l_encabezado = true then
				l_sql = " SELECT * "
				l_sql = l_sql & " FROM buq_buque "
				l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
				l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
				l_sql = l_sql & " inner join buq_destino on buq_destino.desnro = buq_contenido.desnro "
				l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
				l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
				l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
				l_sql = l_sql & " ORDER BY  buq_destino.desdes "
				rsOpen l_rs, cn, l_sql, 0 
				
				'response.write l_sql
				
				if not l_rs.eof then
					l_desdes = ""
				end if
				
				
				l_indice_destino = 1
				l_indice_mercaderia = 1
				do while not l_rs.eof
							
					if l_desdes <> l_rs("desdes") then
						ArrDesNro(l_indice_destino) = l_rs("desnro")
						ArrDesDes(l_indice_destino) = l_rs("desdes")
						l_desdes = l_rs("desdes")
						l_indice_destino = l_indice_destino + 1
					end if
					
					l_existe = false
					for x = 1 to l_indice_mercaderia - 1
						if l_rs("mernro") = ArrMerNro(x) then
							l_existe = true
							l_ColMer = x
						end if 
					next
					if l_existe = false then
						ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
						ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
						l_ColMer = l_indice_mercaderia
						l_indice_mercaderia = l_indice_mercaderia + 1
					end if 
				
					MatMerDes(l_ColMer , l_indice_destino -1) = MatMerDes(l_ColMer , l_indice_destino -1) + l_rs("conton")
	
					l_rs.MoveNext
				loop
				l_rs.Close
			'end if
				
			for x = 1 to l_indice_mercaderia - 1
				%>			  
			   <th align="center" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " ><%= ArrMerDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " >Toneladas</th>					
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


encabezado_expcasdes("Total Exportado por Destino - Cereales, Aceites y Subproductos") 
l_nrolinea = 6

dim ran_descas_acu_nom1
dim ran_descas_acu_val1
dim ran_descas_acu_nom2
dim ran_descas_acu_val2
dim ran_descas_acu_nom3
dim ran_descas_acu_val3

ran_descas_acu_nom1 = 0
ran_descas_acu_val1 = 0
ran_descas_acu_nom2 = 0
ran_descas_acu_val2 = 0
ran_descas_acu_nom3 = 0
ran_descas_acu_val3 = 0

for x = 1 to l_indice_destino - 1
'	if l_nrolinea > l_Max_Lineas_X_Pag then
'		response.write "si" & "-" & l_nrolinea & "-" & l_Max_Lineas_X_Pag & "<br>"
'		response.write "</table><p style='page-break-before:always'></p>"
'		l_nropagina = l_nropagina + 1
'		l_encabezado = false
'		encabezado_expcasdes("Total Exportado por Destino - Cereales, Aceites y Subproductos") 
'		l_nrolinea = 6

'	else	
'	response.write "no" & "<br>"
'	end if
	%>
	<tr>
		<td nowrap align="center" width="10%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrDesDes(x) %></td>			
	<%
	l_TotMerDes = 0
	for y = 1 to l_indice_mercaderia - 1
		if MatMerDes(y,x) = "" then
		%>
			<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " >&nbsp;</td>			
		<%
		else
		%>
			<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " ><%= MatMerDes(y,x) %></td>			
		<%
		end if
		
		l_TotMerDes = l_TotMerDes + MatMerDes(y,x)
	next
	
	'---------------------mayor
	if l_TotMerDes >= ran_descas_acu_val1 then 
		ran_descas_acu_nom3 = ran_descas_acu_nom2
		ran_descas_acu_val3 = ran_descas_acu_val2
		
		ran_descas_acu_nom2 = ran_descas_acu_nom1
		ran_descas_acu_val2 = ran_descas_acu_val1
		
		ran_descas_acu_nom1 = ArrDesDes(x)
		ran_descas_acu_val1 = l_TotMerDes

	else
		if l_TotMerDes >= ran_descas_acu_val2 then
			ran_descas_acu_nom3 = ran_descas_acu_nom2
			ran_descas_acu_val3 = ran_descas_acu_val2
			
			ran_descas_acu_nom2 = ArrDesDes(x)
			ran_descas_acu_val2 = l_TotMerDes

		else 
			if l_TotMerDes >= ran_descas_acu_val3 then
				ran_descas_acu_nom3 = ArrDesDes(x)
				ran_descas_acu_val3 = l_TotMerDes
			end if
		end if
	end if	
	'---------------------fin mayor	
	
	%>
		<td align="right" width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_TotMerDes %></td>			
	</tr>	
	<%
	l_nrolinea = l_nrolinea + 1	
next
%>
	<tr>
		<td align="center" width="5%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Total</td>			
<%

'response.end

' Totales
l_TotTotMerDes = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerDes = 0
	for x = 1 to l_indice_destino - 1
		l_TotMerDes = l_TotMerDes + MatMerDes(i,x)
	next
%>	
	<td align="right"  width="7%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><% if l_TotMerDes = 0 then response.write "" else response.write l_TotMerDes end if  %></td>			
<%	
	l_TotTotMerDes = l_TotTotMerDes + l_TotMerDes
next
%>	
	<td  align="right" width="10%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= l_TotTotMerDes %></td>			
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

