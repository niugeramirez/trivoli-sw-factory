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

'Variable usadas para imprimir los Totales
dim l_nroope

' Imprime los Totales


dim l_rep1
dim l_rep2
dim l_rep3
dim l_rep4
dim l_rep5
dim l_rep6 ' pendiente
dim l_rep7
dim l_rep8
dim l_rep9

dim l_rep10

dim l_rep11

dim l_rep12
dim l_rep13
dim l_rep14
Dim l_rep15

Dim l_rep16
Dim l_rep17


Dim l_rep18

dim l_rep19
Dim l_rep20
dim l_rep21
Dim l_rep22

Dim ArrTotDes(100)


'---------------
' rep2
'---------------


'---------------
' rep3
Dim l_merdes
Dim l_indice_mercaderia
Dim MatMesMer(50,50)
Dim l_Mes
Dim l_TotMesMer
Dim l_TotTotMesMer
'---------------

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

'---------------
' rep5
Dim l_anioini
'---------------

'---------------
' rep7
Dim l_total_toneladas

Dim l_desdes
Dim ArrDesNro(50)
Dim ArrDesDes(50)
Dim MatMerDes(50,50)
Dim l_indice_destino
Dim l_TotMerDes
Dim l_TotTotMerDes
'---------------

'---------------
' rep8
Dim l_TotMes
Dim l_TotMer
Dim l_TotTotMerMes

Dim ArrMerMes(130, 12)
'---------------

'---------------
' rep9
Dim l_sitdes
Dim l_indice_sitio
Dim ArrSitDes(50)
Dim ArrSitNro(50)
Dim MatSitMer(20,100)

'---------------
' rep10
Dim l_indice_terminal
Dim ArrTerDes(150)
Dim MatMerTer(100,50)

Dim TotCol(100)
Dim TotFil(100)
Dim TotFilCol

'---------------
' rep11
Dim MatMesDes(50,50)
Dim MatMesExp(50,50)
'---------------


'---------------
' rep12
Dim l_FilExp
Dim TotFil2(100)
'---------------


'---------------
' rep15
Dim ArrTipMer(4,100)
Dim l_totfil15
'---------------

'---------------
' rep18
Dim ArrTipMes(4,12)
Dim l_Empresa_Aux
'---------------

Dim l_rep7111

'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")

l_repelegido  = request.querystring("repnro")

l_anioini = "01/01/" & year(l_fecfin)

if l_repelegido = 0 then
	l_rep1 = true
	l_rep2 = true
	l_rep3 = true
	l_rep4 = true
	l_rep5 = true
	l_rep6 = false ' Pendiente
	l_rep7 = true ' igual al 5 pero por destino 
	l_rep8 = true
	l_rep9 = true
	l_rep10 = true
	l_rep11 = false ' Exportacion de Pescado por ahora no se imprime
	
	l_rep12 = false
	
	l_rep13 = true
	l_rep14 = true
	l_rep15 = true
	
	l_rep18 = true
	
	l_rep19 = true
	l_rep20 = true
	l_rep21 = true
	l_rep22 = false

end if

select case l_repelegido
case 1
	l_rep1 = true
case 2
	l_rep2 = true
case 3
	l_rep3 = true
case 4
	l_rep4 = true
case 5
	l_rep5 = true
case 6
	l_rep6 = true
case 7
	l_rep7 = true
case 8
	l_rep8 = true	
case 9
	l_rep9 = true	
case 10
	l_rep10 = true	
case 11
	l_rep11 = true	
case 12
	l_rep12 = true	
case 13
	l_rep13 = true	
case 14
	l_rep14 = true	
case 15
	l_rep15 = true
case 16
	l_rep16 = true	
case 17
	l_rep17 = true	
case 18
	l_rep18 = true	
case 19
	l_rep19 = true	
case 20
	l_rep20 = true	
case 21
	l_rep21 = true	
case 22
	l_rep22 = true
end select	

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



sub totales()
	%>
	<tr>
		<td colspan="<%= l_cantcols %>">&nbsp;</td>
	</tr>
	<tr>
		<td align="right"colspan="4"><b>Total de movimientos faltantes en HP9000 :</b></td>
		<td align="center"><b><%= l_Total %></b></td>
		<td align="center" colspan="<%= l_cantcols - 5 %>">&nbsp;</td>					
	</tr>
	<tr>
		<td colspan="<%= l_cantcols %>">&nbsp;</td>
	</tr>		
	<%
end sub 'totales


sub Portada()
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																	
					<tr>
				       	<td bgcolor="#FFFFFF" width="5%" style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="1">&nbsp;<!--<img src="/serviciolocal/shared/images/puerto.gif" border="0">--></td>
				       	<td bgcolor="#FFFFFF" width="5%" style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="1">&nbsp;<!--Cámara Portuaria y Marítima <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; de Bahía Blanca--></td>
				       	<td bgcolor="#FFFFFF" width="90%"style=" border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " align="left" nowrap colspan="6">&nbsp;</td>													
						</td>				
					</tr>												
					<tr>
				       	<td colspan="8" bgcolor="#FFFFFF">&nbsp;
						</td>				
					</tr>
					<tr>
						<td style="FONT-SIZE: 13pt;" align="center" width="100%" colspan="8" bgcolor="#FFFFFF">
							<b><%= NombreMes(month(l_fecfin))  %>&nbsp;-&nbsp;<%= year(l_fecfin) %></b>
						</td>
					</tr>										
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>										
					<tr>					
						<td align="center" width="100%" colspan="8" bgcolor="#FFFFFF">
							<b><img src="/serviciolocal/shared/images/portada.gif" border="0"></b> 
						</td>
					</tr>
					<% 'response.end %>
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																	
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																	
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																																
					<tr>
				       	<td align="center" bgcolor="#FFFFFF" colspan="8" style="FONT-SIZE: 16pt;">INDICE
						</td>				
					</tr>					
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																																
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																																										
					<tr>
				       	<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" colspan="7">&nbsp;Exportaciones</td>
						<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" nowrap colspan="8">1</td>						
						</td>				
					</tr>					
					<tr>
				       	<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" colspan="7">&nbsp;Importaciones</td>
						<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" nowrap colspan="8">8</td>						
						</td>				
					</tr>										
					<tr>
				       	<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" colspan="7">&nbsp;Removido</td>
						<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" nowrap colspan="8">9</td>						
						</td>				
					</tr>
					<tr>
				       	<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" nowrap colspan="7">&nbsp;Cargas Generales</td>
						<td bgcolor="#FFFFFF" style="FONT-SIZE: 12pt;" nowrap colspan="8">15</td>						
						</td>				
					</tr>																				
					
					<tr>
				       	<td bgcolor="#FFFFFF" nowrap colspan="8">&nbsp;
						</td>				
					</tr>																	
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																																																	
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																																																	
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																			
					<tr>
				       	<td bgcolor="#FFFFFF" colspan="8">&nbsp;
						</td>				
					</tr>																													
				</table>
			</td>				
		</tr>
<%
			response.write "</table><p style='page-break-before:always'></p>"
end sub


sub encabezado_expbuq(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
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
					<tr>
						<td align="center" width="100%" colspan="7">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap width="5%" > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
						<td align="center" width="100%" colspan="7">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap width="5%" > 
							&nbsp;
						</td>										
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

sub encabezado_impbuq(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="8">
				<table cellpadding="0" cellspacing="0">
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
					<tr>
						<td align="center" width="100%" colspan="7">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap colspan="1"  width="5%"> 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
						<td align="center" width="100%" colspan="7">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap colspan="1" > 
							&nbsp;
						</td>								
					</tr>
					<tr>
				       	<td nowrap colspan="8">&nbsp;
						</td>				
					</tr>														
				</table>
			</td>				
		</tr>
		<tr>
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Buque</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Comenzó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Terminó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Toneladas</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Mercadería</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Sitio</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Agencia</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Procedencia</th>				
	    </tr>		
<%
end sub 'encabezado

sub encabezado_impmer(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0" >
		<tr>
			<td align="center" colspan="14">
				<table>
					<tr>
				       	<td nowrap colspan="3">&nbsp;
						</td>				
					</tr>																
					<% if l_encabezado = true then %>					
					<tr>
				       	<td nowrap colspan="3"><%= l_Empresa %>
						</td>				
					</tr>								
					<% End If %>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap >
							<% if l_encabezado = true then %>					
								P&aacute;gina: <%= l_nropagina%>
							<% End If %>			
							&nbsp;					
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap > 
						&nbsp;
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
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Mercadería</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">ENE</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">FEB</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">MAR</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">ABR</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">MAY</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">JUN</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">JUL</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">AGO</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">SEP</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">OCT</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">NOV</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">DIC</th>												
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">TON</th>			
	    </tr>		
<%
end sub 'encabezado

sub encabezado_exppes(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
					<tr>
				       	<td nowrap><%= l_Empresa %>
						</td>
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>

				       	<td align="right" nowrap > &nbsp;
						</td>			
					</tr>
				</table>
			</td>				
		</tr>
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Destino</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">ENE</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">FEB</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">MAR</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">ABR</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">MAY</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">JUN</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">JUL</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">AGO</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">SEP</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">OCT</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">NOV</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">DIC</th>												
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">TON</th>			
	    </tr>		
<%
end sub 'encabezado


sub encabezado_expcas(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0">
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
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
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_exportadora.expdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_expdes = ""
			end if
			
			
			l_indice_exportadora = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_expdes <> l_rs("expdes") then
					ArrExpNro(l_indice_exportadora) = l_rs("expnro")
					ArrExpDes(l_indice_exportadora) = l_rs("expdes")
					l_expdes = l_rs("expdes")
					l_indice_exportadora = l_indice_exportadora + 1
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
			
				MatMerExp(l_ColMer , l_indice_exportadora -1) = MatMerExp(l_ColMer , l_indice_exportadora -1) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_mercaderia - 1
			%>			  
			  <th align="center" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " ><%= ArrMerDes(x) %></th>
			  <!-- <th style="FONT-SIZE: 9pt; layout-flow : vertical-ideographic;" align="center" ><%'= ArrMerDes(x) %></td>-->
			<%
			next
			%>			  
			   <th align="center" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " >Toneladas</th>					
 		    </tr>	
<%
end sub

sub encabezado_parter(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
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
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Terminal</th>			
			<%
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
				l_merdes = ""
			end if
			
			Inicializar_Arreglo TotCol, 50 , 0
			Inicializar_Arreglo TotFil, 50 , 0
			TotFilCol = 0
			
			l_indice_terminal = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_merdes <> l_rs("merdes") then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_merdes = l_rs("merdes")
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_terminal - 1
					if l_rs("sitter") = ArrTerDes(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					'ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrTerDes(l_indice_terminal) = l_rs("sitter")
					l_ColMer = l_indice_terminal
					l_indice_terminal = l_indice_terminal + 1
				end if
			
				TotCol(l_indice_mercaderia -1) = TotCol(l_indice_mercaderia -1) + l_rs("conton")
				TotFil(l_ColMer) = TotFil(l_ColMer) + l_rs("conton")
				TotFilCol = TotFilCol + + l_rs("conton")
			
				MatMerTer(l_indice_mercaderia -1, l_ColMer ) = MatMerTer(l_indice_mercaderia -1, l_ColMer ) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			for x = 1 to l_indice_mercaderia - 1
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrMerDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Toneladas</th>					
 		    </tr>	
<%
end sub


sub encabezado_movgen(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="12">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b>-&nbsp; <%= l_anioini %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
						&nbsp;
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>
	    <tr>
	        <th colspan="7" align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Cereales, Aceites y Subproductos</th>			
		</tr>							
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Empresas</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
				l_merdes = ""
			end if
			
			Inicializar_Arreglo ArrMerNro, 50 , 0
			Inicializar_Arreglo ArrSitNro, 50 , 0
			Inicializar_Arreglo TotCol, 50 , 0
			Inicializar_Arreglo TotFil, 50 , 0
			TotFilCol = 0
			for i = 1 to 20
				for j = 1 to 100
					MatSitMer(i,j) = 0
				next
			next
			
			l_indice_sitio = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_merdes <> l_rs("merdes") then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_merdes = l_rs("merdes")
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_sitio - 1
					if l_rs("sitnro") = ArrSitNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrSitNro(l_indice_sitio) = l_rs("sitnro")
					ArrSitDes(l_indice_sitio) = l_rs("sitdes")
					l_ColMer = l_indice_sitio
					l_indice_sitio = l_indice_sitio + 1
				end if 
			
				MatSitMer(l_ColMer , l_indice_mercaderia -1) = MatSitMer(l_ColMer , l_indice_mercaderia -1) + l_rs("conton")
				
				TotCol(l_indice_sitio -1) = TotCol(l_indice_sitio -1) + l_rs("conton")
				TotFil(l_indice_mercaderia - 1) = TotFil(l_indice_mercaderia - 1) + l_rs("conton")

				TotFilCol = TotFilCol + l_rs("conton")
				
				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_sitio - 1
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= ArrSitDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</th>					
 		    </tr>	
<%
			'response.end
end sub

sub encabezado_movgeninf(titulo)
%>
	    <tr>
	        <th colspan="12" align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Inflamables</th>			
		</tr>					
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Producto</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'INF' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
				l_merdes = ""
			end if
			
			Inicializar_Arreglo ArrMerNro, 50 , 0
			Inicializar_Arreglo ArrSitNro, 50 , 0
			Inicializar_Arreglo TotCol, 50 , 0
			Inicializar_Arreglo TotFil, 50 , 0
			TotFilCol = 0
			for i = 1 to 20
				for j = 1 to 100
					MatSitMer(i,j) = 0
				next
			next
			
			l_indice_sitio = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_merdes <> l_rs("merdes") then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_merdes = l_rs("merdes")
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_sitio - 1
					if l_rs("sitnro") = ArrSitNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrSitNro(l_indice_sitio) = l_rs("sitnro")
					ArrSitDes(l_indice_sitio) = l_rs("sitdes")
					l_ColMer = l_indice_sitio
					l_indice_sitio = l_indice_sitio + 1
				end if 
			
				MatSitMer(l_ColMer , l_indice_mercaderia -1) = MatSitMer(l_ColMer , l_indice_mercaderia -1) + l_rs("conton")
				
				TotCol(l_indice_sitio -1) = TotCol(l_indice_sitio -1) + l_rs("conton")
				TotFil(l_indice_mercaderia - 1) = TotFil(l_indice_mercaderia - 1) + l_rs("conton")

				TotFilCol = TotFilCol + l_rs("conton")
				
				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_sitio - 1
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= ArrSitDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</th>					
 		    </tr>	
<%

end sub


sub encabezado_movgenotr(titulo)
%>
	    <tr>
	        <th colspan="12" align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Otras Cargas</th>			
		</tr>					
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Producto</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'OTR' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			if not l_rs.eof then
				l_merdes = ""
			end if
			
			Inicializar_Arreglo ArrMerNro, 50 , 0
			Inicializar_Arreglo ArrSitNro, 50 , 0
			Inicializar_Arreglo TotCol, 50 , 0
			Inicializar_Arreglo TotFil, 50 , 0
			TotFilCol = 0
			for i = 1 to 20
				for j = 1 to 100
					MatSitMer(i,j) = 0
				next
			next
			
			l_indice_sitio = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_merdes <> l_rs("merdes") then
					ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
					ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
					l_merdes = l_rs("merdes")
					l_indice_mercaderia = l_indice_mercaderia + 1
				end if
				
				l_existe = false
				for x = 1 to l_indice_sitio - 1
					if l_rs("sitnro") = ArrSitNro(x) then
						l_existe = true
						l_ColMer = x
					end if 
				next
				if l_existe = false then
					ArrSitNro(l_indice_sitio) = l_rs("sitnro")
					ArrSitDes(l_indice_sitio) = l_rs("sitdes")
					l_ColMer = l_indice_sitio
					l_indice_sitio = l_indice_sitio + 1
				end if 
			
				MatSitMer(l_ColMer , l_indice_mercaderia -1) = MatSitMer(l_ColMer , l_indice_mercaderia -1) + l_rs("conton")
				
				TotCol(l_indice_sitio -1) = TotCol(l_indice_sitio -1) + l_rs("conton")
				TotFil(l_indice_mercaderia - 1) = TotFil(l_indice_mercaderia - 1) + l_rs("conton")

				TotFilCol = TotFilCol + l_rs("conton")
				
				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_sitio - 1
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= ArrSitDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 >Total</th>					
 		    </tr>	
<%

end sub


sub encabezado_expcasanio(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table cellpadding="0" cellspacing="0">
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
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
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_exportadora.expdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_expdes = ""
			end if
			
			
			l_indice_exportadora = 1
			l_indice_mercaderia = 1
			do while not l_rs.eof
						
				if l_expdes <> l_rs("expdes") then
					ArrExpNro(l_indice_exportadora) = l_rs("expnro")
					ArrExpDes(l_indice_exportadora) = l_rs("expdes")
					l_expdes = l_rs("expdes")
					l_indice_exportadora = l_indice_exportadora + 1
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
			
				MatMerExp(l_ColMer , l_indice_exportadora -1) = MatMerExp(l_ColMer , l_indice_exportadora -1) + l_rs("conton")

				l_rs.MoveNext
			loop
			l_rs.Close
			
			
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


sub encabezado_expcasdes(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table cellpadding="0" cellspacing="0">
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
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
				l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
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

sub encabezado_porparsitcas(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="15">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%" rowspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"  >Producto</th>			
			<%  
			l_sql = " SELECT * "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
			l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			'l_sql = l_sql & "			and buq_mercaderia.mernro = 13 "
			l_sql = l_sql & " ORDER BY  buq_sitio.sitdes "
			rsOpen l_rs, cn, l_sql, 0 
			
			'response.write l_sql
			
			if not l_rs.eof then
				l_sitdes = ""
			end if
			
			l_indice_sitio = 1
			l_indice_mercaderia = 1
			TotFilCol = 0
			do while not l_rs.eof
						
				if l_sitdes <> l_rs("sitdes") then
					ArrSitNro(l_indice_sitio) = l_rs("sitnro")
					ArrSitDes(l_indice_sitio) = l_rs("sitdes")
					l_sitdes = l_rs("sitdes")
					l_indice_sitio = l_indice_sitio + 1
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
			
				MatSitMer(l_indice_sitio -1, l_ColMer ) = MatSitMer(l_indice_sitio -1, l_ColMer ) + l_rs("conton")
				
				TotCol(l_indice_sitio -1) = TotCol(l_indice_sitio -1) + l_rs("conton")
				TotFil(l_ColMer) = TotFil(l_ColMer) + l_rs("conton")
				
				TotFilCol = TotFilCol + l_rs("conton")
				l_rs.MoveNext
			loop
			l_rs.Close
			
			
			for x = 1 to l_indice_sitio - 1
			%>			  
			   <th align="center" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrSitDes(x) %></th>
			<%
			next
			%>			  
			   <th align="center" rowspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Total</th>					
 		    </tr>	
 		    <tr>						
			<%
			for x = 1 to l_indice_sitio - 1
			%>			  
			   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Ton</th>
	 		   <th align="center" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">%</th>			   
			<%
			next
			%>			  
 		    </tr>				
<%
end sub


sub encabezado_expinf(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
					<tr>
				       	<td nowrap colspan="3"><%= l_Empresa %>
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%= l_anioini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>
		<%
		l_sql = " SELECT  * "
		l_sql = l_sql & " FROM buq_buque "
		l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
		l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
		l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "		
		l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
		l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
		l_sql = l_sql & " AND buq_mercaderia.tipmerdes = 'INF' "
		l_sql = l_sql & " AND buq_buque.tipopenro = 3 "
		l_sql = l_sql & " Order by buq_mercaderia.mernro "
		
		l_indice_mercaderia = 1
		l_indice_exportadora = 1
		l_merdes = ""
		
		Inicializar_Arreglo ArrMerNro, 50 , 0
		Inicializar_Arreglo ArrMerDes, 50 , ""
		Inicializar_Arreglo TotCol, 50 , 0
		Inicializar_Arreglo TotFil, 50 , 0
		Inicializar_Arreglo ArrExpNro, 50 , 0
		Inicializar_Arreglo ArrExpDes, 50 , ""
		
		TotFilCol = 0
		
		rsOpen l_rs, cn, l_sql, 0
		do while not l_rs.eof
			if l_merdes <> l_rs("merdes") then
				ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
				ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
				l_merdes = l_rs("merdes")
				l_indice_mercaderia =  	l_indice_mercaderia + 1
			end if
			
			l_existe = false
			for x = 1 to l_indice_exportadora - 1
				if l_rs("expnro") = ArrExpNro(x) then
					l_existe = true
					l_FilExp = x
				end if 
			next
			if l_existe = false then
				ArrExpNro(l_indice_exportadora) = l_rs("expnro")
				ArrExpDes(l_indice_exportadora) = l_rs("expdes")
				l_FilExp = l_indice_exportadora
				l_indice_exportadora = l_indice_exportadora + 1
			end if 
			
			ArrMerMes(l_indice_mercaderia - 1, month(l_rs("buqfechas"))  ) = ArrMerMes(l_indice_mercaderia - 1 , month(l_rs("buqfechas"))  )  + l_rs("conton")			
			TotCol(l_indice_mercaderia -1) = TotCol(l_indice_mercaderia -1) + cdbl(l_rs("conton"))
			TotFil(month(l_rs("buqfechas"))) = TotFil(month(l_rs("buqfechas"))) + cdbl(l_rs("conton"))

			MatMerExp(l_indice_mercaderia -1, l_FilExp ) = MatMerExp(l_indice_mercaderia -1, l_FilExp ) + l_rs("conton")
			TotFil2(l_FilExp) = TotFil2(l_FilExp) + cdbl(l_rs("conton"))
				
			TotFilCol = TotFilCol + l_rs("conton")
			
			
			l_rs.movenext
		loop
		l_rs.close
		%>
		<tr>
			<th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mes</th>
		<%
		for i = 1 to l_indice_mercaderia - 1
		%>
		  <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= ArrMerDes(i) %></th>
		<%
		next
		%>
			<th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</th>
		  </tr>		  							
		<%

end sub





sub encabezado_expcasaniodestino(titulo)

	l_anioini = "01/01/" & year(l_fecfin)
%>
	<table style="width:99%">
		<tr>
			<td align="center" colspan="<%= l_cantcols %>">
				<table>
					<tr>
				       	<td nowrap colspan="3">Cámara Portuaria y Marítima de Bahía Blanca
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							Período:&nbsp;&nbsp;01/01/<%= year(l_fecfin)  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>

	    <tr>
	        <th align="center" width="10%">Destino</th>			
			<%  
			l_sql = " SELECT distinct(merdes), merord, buq_mercaderia.mernro "
			l_sql = l_sql & " FROM buq_buque "
			l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
			l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "

			l_sql = l_sql & " AND buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
			l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)	
			l_sql = l_sql & " WHERE  buq_mercaderia.tipmerdes = 'CAS' "
			l_sql = l_sql & " ORDER BY  buq_mercaderia.merord "
			rsOpen l_rs, cn, l_sql, 0 
			l_indice = 0
			do while not l_rs.eof
				CAS(l_indice) = l_rs("mernro")
			%>			  
			   <th align="center" ><%= l_rs("merdes") %></th>					
			<%
				 l_indice = l_indice + 1
				 l_rs.movenext
			 loop
			 %>
			 <th align="center" width="10%">Toneladas</th>
			 <th align="center" width="10%">%</th>			 
 		    </tr>	
<%
end sub


sub encabezado_cabmarcar(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;
						</td>				

						<td align="center" width="100%">
							<%= l_fecini  %>&nbsp;-&nbsp;<%= l_fecfin %>
						</td>
				       	<td align="right" nowrap > &nbsp;
						</td>					
					</tr>
				</table>
			</td>				
		</tr>

		<tr>
	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Buque</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Comenzó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Terminó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Toneladas</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mercadería</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Sitio</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Agencia</th>		
	    </tr>		
<%
end sub 'encabezado



sub encabezado_detcargasitio(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table>
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
				       	<td colspan="20">&nbsp;
						</td>				
					</tr>				
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
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
<%
end sub


sub encabezado_reminf(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
						&nbsp;
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub


sub encabezado_remimpexp(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							<% if l_Empresa_Aux <> "" then%>							
								P&aacute;gina: <%= l_nropagina%>
							<% End If %>								
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
						&nbsp;
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub

sub encabezado_detatebuqage(titulo)


%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
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


sub encabezado_MovBuqSitMes(titulo)

%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="20">
				<table>
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<!--
							Período:&nbsp;&nbsp;<%'= l_fecini  %>&nbsp;-&nbsp;<%'= l_fecfin %>
							-->
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub

sub encabezado_MovBuqSitMes2(titulo)

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
				       	<td align="right" nowrap > &nbsp;
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<!--
							Período:&nbsp;&nbsp;<%'= l_fecini  %>&nbsp;-&nbsp;<%'= l_fecfin %>
							-->
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>


<%
end sub


sub encabezado_podio(titulo)


%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table cellpadding="0" cellspacing="0" border="0">
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
					<tr>
				       	<td nowrap>&nbsp;
						</td>				
						<td align="center" width="100%" style="FONT-SIZE: 10pt">
							<b><%= titulo%></b> 
						</td>
				       	<td align="right" nowrap > 
							P&aacute;gina: <%= l_nropagina%>
						</td>				
					</tr>
					<tr>
   			         	<td nowrap>&nbsp;&nbsp;&nbsp;
						</td>				
						<td align="center" width="100%">
							<%'= l_fecini  %>&nbsp;&nbsp;<%'= l_fecfin %>
						</td>
				       	<td align="right" nowrap >&nbsp;
						</td>										
					</tr>
				</table>
			</td>				
		</tr>
		<!--
	    <tr>
	        <td align="center" width="30%">&nbsp;</td>					
	        <th align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" nowrap>Agencias</th>			
   		    <th align="center" width="20%" style="border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;" nowrap >Buques Atendidos</th>	
	        <td align="center" width="30%">&nbsp;</td>								 
	    </tr>
		-->

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

<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">


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


'--------------------------------
' Llamo a la Portada del Programa
'--------------------------------
Portada()


if l_rep1 = true then

'l_nropagina = 1
encabezado_expbuq("Exportación - Detalle de Buques") 
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
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " ORDER BY buq_buque.buqfechas, buq_buque.buqfecdes " 
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
				<td align="left" width="10%" nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%=l_rs("buqdes")%></td>			
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
end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep4 = true then 

encabezado_expcas("Exportación de Cereales, Aceites y Subproductos") 
l_nrolinea = 6

dim ran_empcas_nom1
dim ran_empcas_val1
dim ran_empcas_nom2
dim ran_empcas_val2
dim ran_empcas_nom3
dim ran_empcas_val3

dim ran_mercas_nom1
dim ran_mercas_val1
dim ran_mercas_nom2
dim ran_mercas_val2
dim ran_mercas_nom3
dim ran_mercas_val3

ran_empcas_nom1 = 0
ran_empcas_val1 = 0
ran_empcas_nom2 = 0
ran_empcas_val2 = 0
ran_empcas_nom3 = 0
ran_empcas_val3 = 0

ran_mercas_nom1 = 0
ran_mercas_val1 = 0
ran_mercas_nom2 = 0
ran_mercas_val2 = 0
ran_mercas_nom3 = 0
ran_mercas_val3 = 0

for x = 1 to l_indice_exportadora - 1
		if l_nrolinea > l_Max_Lineas_X_Pag then
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina = l_nropagina + 1
			encabezado_expcas("Exportación de Cereales, Aceites y Subproductos") 
			l_nrolinea = 6
		end if
	%>
	<tr>
		<td nowrap align="center" width="5%" style="FONT-SIZE: 7pt; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= left(ArrExpDes(x),15) %></td>			
	<%
	l_TotMerExp = 0
	for y = 1 to l_indice_mercaderia - 1
		if MatMerExp(y,x) = "" then
		%>
			<td align="right"  width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " >&nbsp;</td>			
		<%
		else
		%>
			<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " ><%= MatMerExp(y,x) %></td>			
		<%
		end if 
		
		l_TotMerExp = l_TotMerExp + MatMerExp(y,x)
	next
	'---------------------mayor
	if l_TotMerExp >= ran_empcas_val1 then 
		ran_empcas_nom3 = ran_empcas_nom2
		ran_empcas_val3 = ran_empcas_val2
		
		ran_empcas_nom2 = ran_empcas_nom1
		ran_empcas_val2 = ran_empcas_val1
		
		ran_empcas_nom1 = ArrExpDes(x)
		ran_empcas_val1 = l_TotMerExp


	else
		if l_TotMerExp >= ran_empcas_val2 then
			ran_empcas_nom3 = ran_empcas_nom2
			ran_empcas_val3 = ran_empcas_val2
			
			ran_empcas_nom2 = ArrExpDes(x)
			ran_empcas_val2 = l_TotMerExp

		else 
			if l_TotMerExp >= ran_empcas_val3 then
				ran_empcas_nom3 = ArrExpDes(x)
				ran_empcas_val3 = l_TotMerExp
			end if
		end if
	end if	
	'---------------------fin mayor
	
	%>
		<td align="right" width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_TotMerExp %></td>			
	</tr>	
	<%
	l_nrolinea = l_nrolinea + 1
next

%>
	<tr>
		<td align="center" width="5%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; " >Total</td>			
<%

' Totales
l_TotTotMerExp = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerExp = 0
	for x = 1 to l_indice_exportadora - 1
		l_TotMerExp = l_TotMerExp + MatMerExp(i,x)
	next
	
	'---------------------mayor
	if l_TotMerExp >= ran_mercas_val1 then 
		ran_mercas_nom3 = ran_mercas_nom2
		ran_mercas_val3 = ran_mercas_val2
		
		ran_mercas_nom2 = ran_mercas_nom1
		ran_mercas_val2 = ran_mercas_val1
		
		ran_mercas_nom1 = ArrMerDes(i)
		ran_mercas_val1 = l_TotMerExp


	else
		if l_TotMerExp >= ran_mercas_val2 then
			ran_mercas_nom3 = ran_mercas_nom2
			ran_mercas_val3 = ran_mercas_val2
			
			ran_mercas_nom2 = ArrMerDes(i)
			ran_mercas_val2 = l_TotMerExp

		else 
			if l_TotMerExp >= ran_mercas_val3 then
				ran_mercas_nom3 = ArrMerDes(i)
				ran_mercas_val3 = l_TotMerExp
			end if
		end if
	end if	
	'---------------------fin mayor	
	
	%>	
	<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><% if l_TotMerExp = 0 then response.write "" else response.write l_TotMerExp end if  %></td>			
<%	
	l_TotTotMerExp = l_TotTotMerExp + l_TotMerExp
next
%>	
	<td  align="right" width="5%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= l_TotTotMerExp %></td>			
</tr>	
<%
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep5 = true then 

encabezado_expcasanio("Exportación de Cereales, Aceites y Subproductos") 
l_nrolinea = 6

dim ran_empcas_acu_nom1
dim ran_empcas_acu_val1
dim ran_empcas_acu_nom2
dim ran_empcas_acu_val2
dim ran_empcas_acu_nom3
dim ran_empcas_acu_val3

dim ran_mercas_acu_nom1
dim ran_mercas_acu_val1
dim ran_mercas_acu_nom2
dim ran_mercas_acu_val2
dim ran_mercas_acu_nom3
dim ran_mercas_acu_val3

ran_empcas_acu_nom1 = 0
ran_empcas_acu_val1 = 0
ran_empcas_acu_nom2 = 0
ran_empcas_acu_val2 = 0
ran_empcas_acu_nom3 = 0
ran_empcas_acu_val3 = 0

ran_mercas_acu_nom1 = 0
ran_mercas_acu_val1 = 0
ran_mercas_acu_nom2 = 0
ran_mercas_acu_val2 = 0
ran_mercas_acu_nom3 = 0
ran_mercas_acu_val3 = 0

for x = 1 to l_indice_exportadora - 1
	if l_nrolinea > l_Max_Lineas_X_Pag then
		response.write "</table><p style='page-break-before:always'></p>"
		l_nropagina = l_nropagina + 1
		encabezado_expcasanio("Exportación de Cereales, Aceites y Subproductos") 
		l_nrolinea = 6
	end if
	%>
	<tr>
		<td nowrap align="center" width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= left(ArrExpDes(x),15) %></td>			
	<%
	l_TotMerExp = 0
	for y = 1 to l_indice_mercaderia - 1
		if MatMerExp(y,x) = "" then
		%>
			<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " >&nbsp;</td>			
		<%
		else
		%>
			<td align="right"  width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; " ><%= MatMerExp(y,x) %></td>			
		<%
		end if
		
		l_TotMerExp = l_TotMerExp + MatMerExp(y,x )
	next
	'---------------------mayor
	if l_TotMerExp >= ran_empcas_acu_val1 then 
		ran_empcas_acu_nom3 = ran_empcas_acu_nom2
		ran_empcas_acu_val3 = ran_empcas_acu_val2
		
		ran_empcas_acu_nom2 = ran_empcas_acu_nom1
		ran_empcas_acu_val2 = ran_empcas_acu_val1
		
		ran_empcas_acu_nom1 = ArrExpDes(x)
		ran_empcas_acu_val1 = l_TotMerExp


	else
		if l_TotMerExp >= ran_empcas_acu_val2 then
			ran_empcas_acu_nom3 = ran_empcas_acu_nom2
			ran_empcas_acu_val3 = ran_empcas_acu_val2
			
			ran_empcas_acu_nom2 = ArrExpDes(x)
			ran_empcas_acu_val2 = l_TotMerExp

		else 
			if l_TotMerExp >= ran_empcas_acu_val3 then
				ran_empcas_acu_nom3 = ArrExpDes(x)
				ran_empcas_acu_val3 = l_TotMerExp
			end if
		end if
	end if	
	'---------------------fin mayor
		
	%>
		<td align="right" width="5%" style="FONT-SIZE: 7pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_TotMerExp %></td>			
	</tr>	
	<%
	l_nrolinea = l_nrolinea + 1	
next

%>
	<tr>
		<td align="center" width="5%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Total</td>			
<%

' Totales
l_TotTotMerExp = 0
for i = 1 to l_indice_mercaderia - 1
	l_TotMerExp = 0
	for x = 1 to l_indice_exportadora - 1
		l_TotMerExp = l_TotMerExp + MatMerExp(i,x)
	next
	
	'---------------------mayor
	if l_TotMerExp >= ran_mercas_acu_val1 then 
		ran_mercas_acu_nom3 = ran_mercas_acu_nom2
		ran_mercas_acu_val3 = ran_mercas_acu_val2
		
		ran_mercas_acu_nom2 = ran_mercas_acu_nom1
		ran_mercas_acu_val2 = ran_mercas_acu_val1
		
		ran_mercas_acu_nom1 = ArrMerDes(i)
		ran_mercas_acu_val1 = l_TotMerExp


	else
		if l_TotMerExp >= ran_mercas_acu_val2 then
			ran_mercas_acu_nom3 = ran_mercas_acu_nom2
			ran_mercas_acu_val3 = ran_mercas_acu_val2
			
			ran_mercas_acu_nom2 = ArrMerDes(i)
			ran_mercas_acu_val2 = l_TotMerExp

		else 
			if l_TotMerExp >= ran_mercas_acu_val3 then
				ran_mercas_acu_nom3 = ArrMerDes(i)
				ran_mercas_acu_val3 = l_TotMerExp
			end if
		end if
	end if	
	'---------------------fin mayor
%>	
	<td align="right"  width="7%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><% if l_TotMerExp = 0 then response.write "" else response.write l_TotMerExp end if  %></td>			
<%	
	l_TotTotMerExp = l_TotTotMerExp + l_TotMerExp
next
%>	
	<td  align="right" width="10%" style="FONT-SIZE: 7pt;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= l_TotTotMerExp %></td>			
</tr>	
<%
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep7 = true then 

'l_encabezado = true
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
end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************
' PODIO 

encabezado_podio("Ranking Resúmen del Mes y Acumulado")

%>
<tr>
   <td align="center" colspan="6" style="FONT-SIZE: 10pt;" >Exportación Cereales, Aceites y Subproductos por Empresa</td>
</tr>
<tr>
   <td align="center" >&nbsp;</td>					
   <th align="center" colspan="2" >Del Mes</th>					
   <td align="center">&nbsp;</td>					   
   <th align="center" colspan="2">Del Año</th>
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>					
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; ">&nbsp;1 - <%= ran_empcas_nom1 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_empcas_val1 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>					
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; ">&nbsp;1 - <%= ran_empcas_acu_nom1 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_empcas_acu_val1 %></td>					   
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;">&nbsp;2 - <%= ran_empcas_nom2 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_empcas_val2 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;">&nbsp;2 - <%= ran_empcas_acu_nom2 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_empcas_acu_val2 %></td>					      
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;3 - <%= ran_empcas_nom3 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= ran_empcas_val3 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;3 - <%= ran_empcas_acu_nom3 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= ran_empcas_acu_val3 %></td>					      
</tr>

<tr>
	<td align="center" colspan="3">
  	  <iframe frameborder="0" name="ifrmgra21" scrolling="No" src="gra_30.asp?nom1=<%= ran_empcas_nom1 %>&val1=<%= ran_empcas_val1 %>&nom2=<%= ran_empcas_nom2 %>&val2=<%= ran_empcas_val2 %>&nom3=<%= ran_empcas_nom3 %>&val3=<%= ran_empcas_val3 %>" width="350" height="150"></iframe> 
	</td>
	<td align="center" colspan="3">
  	  <iframe frameborder="0" name="ifrmgra22" scrolling="No" src="gra_31.asp?nom1=<%= ran_empcas_acu_nom1 %>&val1=<%= ran_empcas_acu_val1 %>&nom2=<%= ran_empcas_acu_nom2 %>&val2=<%= ran_empcas_acu_val2 %>&nom3=<%= ran_empcas_acu_nom3 %>&val3=<%= ran_empcas_acu_val3 %>" width="350" height="150"></iframe> 
	</td>	
</tr> 

<tr>
   <td align="center" colspan="6" style="FONT-SIZE: 10pt;" >&nbsp;</td>
</tr>
<tr>
   <td align="center" colspan="6" style="FONT-SIZE: 10pt;" >Exportación Cereales, Aceites y Subproductos por Producto</td>
</tr>
<tr>
   <td align="center" >&nbsp;</td>					
   <th align="center" colspan="2" >Del Mes</th>					
   <td align="center">&nbsp;</td>					   
   <th align="center" colspan="2">Del Año</th>
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>					
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; ">&nbsp;1 - <%= ran_mercas_nom1 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_mercas_val1 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>					
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; ">&nbsp;1 - <%= ran_mercas_acu_nom1 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_mercas_acu_val1 %></td>					   
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;">&nbsp;2 - <%= ran_mercas_nom2 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_mercas_val2 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;">&nbsp;2 - <%= ran_mercas_acu_nom2 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_mercas_acu_val2 %></td>					      
</tr>
<tr>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;3 - <%= ran_mercas_nom3 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= ran_mercas_val3 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;3 - <%= ran_mercas_acu_nom3 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= ran_mercas_acu_val3 %></td>					      
</tr>

<tr>
	<td align="center" colspan="3">
  	  <iframe frameborder="0" name="ifrmgra23" scrolling="No" src="gra_40.asp?nom1=<%= ran_mercas_nom1 %>&val1=<%= ran_mercas_val1 %>&nom2=<%= ran_mercas_nom2 %>&val2=<%= ran_mercas_val2 %>&nom3=<%= ran_mercas_nom3 %>&val3=<%= ran_mercas_val3 %>" width="350" height="150"></iframe> 
	</td>
	<td align="center" colspan="3">
  	  <iframe frameborder="0" name="ifrmgra24" scrolling="No" src="gra_41.asp?nom1=<%= ran_mercas_acu_nom1 %>&val1=<%= ran_mercas_acu_val1 %>&nom2=<%= ran_mercas_acu_nom2 %>&val2=<%= ran_mercas_acu_val2 %>&nom3=<%= ran_mercas_acu_nom3 %>&val3=<%= ran_mercas_acu_val3 %>" width="350" height="150"></iframe> 
	</td>	
</tr> 


<tr>
   <td align="center" colspan="6" style="FONT-SIZE: 10pt;" >&nbsp;</td>
</tr>
<tr>
   <td align="center" colspan="6" style="FONT-SIZE: 10pt;" >Exportación Cereales, Aceites y Subproductos por Destino</td>
</tr>
<tr>
   <td align="center" >&nbsp;</td>					
   <td align="center" colspan="2" >&nbsp;</td>					
   <td align="center">&nbsp;</td>					   
   <th align="center" colspan="2">Del Año</th>
</tr>
<tr>
   <td align="left" >&nbsp;&nbsp;</td>					
   <td align="left" >&nbsp;<%'= ran_mercas_nom1 %></td>					
   <td align="center" ><%'= ran_mercas_val1 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>					
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-top-color: #000000; border-top-style: solid; border-top-width: 1px; ">&nbsp;1 - <%= ran_descas_acu_nom1 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_descas_acu_val1 %></td>					   
</tr>
<tr>
   <td align="left" >&nbsp;&nbsp;</td>
   <td align="left" >&nbsp;<%'= ran_descas_acu_nom2 %></td>					
   <td align="center" ><%'= ran_descas_acu_val2 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;">&nbsp;2 - <%= ran_descas_acu_nom2 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px;"><%= ran_descas_acu_val2 %></td>					      
</tr>
<tr>
   <td align="left" >&nbsp;&nbsp;</td>
   <td align="left" >&nbsp;<%'= ran_mercas_nom3 %></td>					
   <td align="center" ><%'= ran_mercas_val3 %></td>					   

   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;&nbsp;</td>
   <td align="left" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">&nbsp;3 - <%= ran_descas_acu_nom3 %></td>					
   <td align="center" style="FONT-SIZE: 8pt;border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= ran_descas_acu_val3 %></td>					      
</tr>

<tr>
	<td align="center" colspan="3">
	<!--
  	  <iframe frameborder="0" name="ifrmgra25" scrolling="No" src="gra_50.asp?nom1=<%= ran_mercas_nom1 %>&val1=<%= ran_mercas_val1 %>&nom2=<%= ran_mercas_nom2 %>&val2=<%= ran_mercas_val2 %>&nom3=<%= ran_mercas_nom3 %>&val3=<%= ran_mercas_val3 %>" width="350" height="150"></iframe> 
	  -->
	</td>
	<td align="center" colspan="3">
  	  <iframe frameborder="0" name="ifrmgra26" scrolling="No" src="gra_51.asp?nom1=<%= ran_descas_acu_nom1%>&val1=<%= ran_descas_acu_val1 %>&nom2=<%= ran_descas_acu_nom2 %>&val2=<%= ran_descas_acu_val2 %>&nom3=<%= ran_descas_acu_nom3 %>&val3=<%= ran_descas_acu_val3 %>" width="350" height="150"></iframe> 
	</td>	
</tr> 
<%
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************
' IMPORTACION
'***************************************************************************************************************************

if l_rep2 = true then

'l_nropagina = l_nropagina + 1
encabezado_impbuq("Importación - Detalle de Buques") 
l_nrolinea = 6


l_sql = " SELECT buqdes, buqfecdes, buqfechas, buq_mercaderia.merdes, buq_sitio.sitdes, buq_agencia.agedes, buq_destino.desdes, sum(conton) tons "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 4 "  ' IMPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " group by buqdes, buqfecdes, buqfechas, buq_mercaderia.merdes, buq_sitio.sitdes, buq_agencia.agedes, buq_destino.desdes "
l_sql = l_sql & " order by buqfechas, buqfecdes "


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
			encabezado_impbuq("Importación - Detalle de Buques") 
			l_nrolinea = 6
		end if
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%" nowrap style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%" nowrap style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;</td>
			   <%
  			   end if
			 %>

			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("buqfechas") %></td>			
			<td align="right" width="5%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("tons") %></td>
			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("agedes") %></td>			
			<td align="center" width="10%" style="border-left-color: #000000; border-left-style: solid; border-left-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;"><%= l_rs("desdes") %></td>			
	    </tr>
		<%
		l_nrolinea = l_nrolinea + 1
		l_totton = l_totton + l_rs("tons")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Cantidad de Buques</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total Toneladas</td>				
	<td align="center" width="10%" colspan="3" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><b><%= l_totton %></b></td>
</tr>
<%
l_nrolinea = l_nrolinea + 1

end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep3 = true then 

l_encabezado = false
encabezado_impmer("Importación - Detalle de Mercaderías") 
l_nrolinea = l_nrolinea + 6


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 4 "  ' IMPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " ORDER BY buq_mercaderia.merdes " 

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end

if not l_rs.eof then
	l_merdes = ""
end if

l_indice_mercaderia = 1

do until l_rs.eof

	if l_merdes <> l_rs("merdes") then
		ArrMerNro(l_indice_mercaderia) = l_rs("mernro")
		ArrMerDes(l_indice_mercaderia) = l_rs("merdes")
		
		MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) = MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) + l_rs("conton")		
		
		l_merdes = l_rs("merdes")
		l_indice_mercaderia = l_indice_mercaderia + 1
	else 
		MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) = MatMesMer(month(l_rs("buqfechas")) , l_indice_mercaderia) + l_rs("conton")						
	end if

	l_rs.MoveNext
loop
l_rs.Close

for i = 1 to l_indice_mercaderia - 1

		if l_nrolinea > l_Max_Lineas_X_Pag then
			response.write "</table><p style='page-break-before:always'></p>"
			l_encabezado = true
			l_nropagina = l_nropagina + 1
			encabezado_impmer("Importación - Detalle de Mercaderías") 
			l_nrolinea = 6
		end if
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(i) %></td>			
<%	
	l_TotMesMer = 0
	for l_Mes = 1 to 12
		if MatMesMer(l_Mes, i) = "" then
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; ">&nbsp;</td>			
		<%
		else
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= MatMesMer(l_Mes, i) %></td>			
		<%
		end if
		
		l_TotMesMer = l_TotMesMer + MatMesMer(l_Mes, i ) 
	next
%>
	<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><%= l_TotMesMer %></td>			
</tr>	
<%
	l_nrolinea = l_nrolinea + 1
next

%>
<tr>	
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>			
<%

' Totales
l_TotTotMesMer = 0
for l_Mes = 1 to 12
	l_TotMesMer = 0
	for i = 1 to l_indice_mercaderia - 1
		l_TotMesMer = l_TotMesMer + MatMesMer(l_Mes, i )
	next
%>	
	<td align="right"  width="7%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><% if l_TotMesMer = 0 then response.write "&nbsp;" else response.write l_TotMesMer end if  %></td>			
<%	
	l_TotTotMesMer = l_TotTotMesMer + l_TotMesMer
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "><%= l_TotTotMesMer %></td>			
</tr>	
<%
l_nrolinea = l_nrolinea + 1

response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1
end if


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************
' REMOVIDO
'***************************************************************************************************************************

if l_rep13 = true then

encabezado_cabmarcar("Cabotaje Marítimo Nacional - Removido Salidas - Cargas") 
l_nrolinea = 6

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 1 "  ' CARGAS
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " order by buq_buque.buqfechas "

rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("buqfechas") %></td>			
			<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("conton") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("agedes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 >Cantidad de Buques</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"	><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total Toneladas</td>				
	<td colspan="2" align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"	><b><%= l_totton %></b></td>
</tr>
<%
response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1
end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************


if l_rep14 = true then
encabezado_cabmarcar("Cabotaje Marítimo Nacional - Removido Entradas - Descargas") 
l_nrolinea = 6

'l_nrolinea = 1
'l_nropagina = 1
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

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 2 "  ' DESCARGAS
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

l_sql = l_sql & " order by buq_buque.buqfechas "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end
if not l_rs.eof then
	l_buqdes = ""
end if

l_canbuq = 0
l_totton = 0
do until l_rs.eof
		%>
		<tr>
			<% if l_buqdes <> l_rs("buqdes") then
			   %>
				<td align="left" width="10%"  nowrap><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				l_canbuq = l_canbuq + 1
			   else
			   %>
				<td align="left" width="10%"  nowrap>&nbsp;</td>			
			   <%
  			   end if
			 %>

			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("buqfechas") %></td>			
			<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("conton") %></td>
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("merdes") %></td>			
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("sitdes") %></td>			
			<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_rs("agedes") %></td>			
	    </tr>
		<%
		l_totton = l_totton + l_rs("conton")
		l_buqdes = l_rs("buqdes")
		
	l_rs.MoveNext
loop
l_rs.Close

%>
<tr>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Cantidad de Buques</td>			
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"	><b><%= l_canbuq %></b></td>
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total Toneladas</td>				
	<td align="center" width="10%" colspan="2" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"	><b><%= l_totton %></b></td>
</tr>
<%
end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep15 = true then 

'l_nrolinea = 1
'l_nropagina = 1
'l_encabezado = true
'l_corte = false
'l_total = 0


l_sql = "  SELECT distinct(sitdes) ,buq_contenido.sitnro "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND buq_sitio.sittip = 'Combustibles' " 
l_sql = l_sql & " AND ( buq_sitio.sitnro = 9 or buq_sitio.sitnro = 10 or buq_sitio.sitnro = 16)" 

rsOpen l_rs, cn, l_sql, 0
do while not l_rs.eof
	response.write "</table><p style='page-break-before:always'></p>"
	l_nropagina = l_nropagina + 1
	encabezado_reminf("Removido Inflamables") 

	' Inicializo 
	for i = 1 to 4
		for j = 1 to 100
			ArrTipMer(i,j) = 0
		next 
	next

	Inicializar_Arreglo TotCol, 50 , 0
	Inicializar_Arreglo TotFil, 50 , 0	
	
	'-----------------------------------------------------
	' Toneladas del Mes
	'-----------------------------------------------------
	
	l_sql = " SELECT  * "
	l_sql = l_sql & " FROM buq_buque "
	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
	l_sql = l_sql & " Order by buq_mercaderia.mernro "
	
	l_indice_mercaderia = 1
	l_merdes = ""
	TotFilCol = 0
	rsOpen l_rs2, cn, l_sql, 0
	do while not l_rs2.eof
		if l_merdes <> l_rs2("merdes") then
			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
			l_merdes = l_rs2("merdes")
			l_indice_mercaderia =  	l_indice_mercaderia + 1
		end if
		
		TotCol(l_rs2("tipopenro")) = TotCol(l_rs2("tipopenro")) + l_rs2("conton")
		TotFil(l_indice_mercaderia - 1) = TotFil(l_indice_mercaderia - 1) + l_rs2("conton")
		
		TotFilCol = TotFilCol + l_rs2("conton")
		
		ArrTipMer(l_rs2("tipopenro"), l_indice_mercaderia - 1 ) = ArrTipMer(l_rs2("tipopenro"), l_indice_mercaderia - 1 )  + l_rs2("conton")
		l_rs2.movenext
	loop
	l_rs2.close
	
	
	%>
	  <tr>
		  <th colspan="6" align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_rs("sitdes") %>&nbsp;&nbsp;&nbsp;&nbsp;<%= l_fecini %>&nbsp;-&nbsp;<%= l_fecfin %></th>							
      </tr>		  
	  <tr>
		  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Producto</td>
	<%
	for i = 1 to 4
	%>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= NombreTipoOperacion(i) %></td>							
	<%
	next
	%>	
		  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 >Total</td>								
	  </tr>		  							
	<%
	l_cadena4 = ""
	for j = 1 to l_indice_mercaderia - 1
	%>
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(j) %></td>
	<%
		for i = 1 to 4
			%>
			  <td align="right"  width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><% if  ArrTipMer(i,j) = 0 then response.write "&nbsp;" else response.write ArrTipMer(i,j) end if %></td>
			<%
		next
	%>
	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(j) %></td>
	</tr>
	<%
	
		l_cadena4 = l_cadena4 & ArrMerDes(j) & "-" & TotFil(j) & ","
	
	next
	
	%>
	<tr>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>									
	<%	
	
	' Totales
	for i = 1 to 4
	%>
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= TotCol(i) %></td>			
	<%
	next
%>
	  <td align="right"  width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>
	</tr> 		  
	<tr>
		<td align="center" colspan="6">
	  	  <iframe frameborder="0" name="ifrmgra15" scrolling="No" src="gra_15.asp?cadena=<%= l_cadena4 %>" width="600" height="200"></iframe> 
		</td>
	</tr>  			
<%	
	
	'-----------------------------------------------------
	' Toneladas del Año
	'-----------------------------------------------------
	
	' Inicializo 
	for i = 1 to 4
		for j = 1 to 100
			ArrTipMer(i,j) = 0
		next 
	next

	Inicializar_Arreglo TotCol, 50 , 0
	Inicializar_Arreglo TotFil, 50 , 0
	
	l_sql = " SELECT  * "
	l_sql = l_sql & " FROM buq_buque "
	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
	l_sql = l_sql & " Order by buq_mercaderia.mernro "
	
	l_indice_mercaderia = 1
	l_merdes = ""
	TotFilCol = 0 
	rsOpen l_rs2, cn, l_sql, 0
	do while not l_rs2.eof
		if l_merdes <> l_rs2("merdes") then
			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
			l_merdes = l_rs2("merdes")
			l_indice_mercaderia =  	l_indice_mercaderia + 1
		end if
		
		TotCol(l_rs2("tipopenro")) = TotCol(l_rs2("tipopenro")) + l_rs2("conton")
		TotFil(l_indice_mercaderia - 1) = TotFil(l_indice_mercaderia - 1) + l_rs2("conton")
		
		TotFilCol = TotFilCol + l_rs2("conton")
		
		ArrTipMer(l_rs2("tipopenro"), l_indice_mercaderia - 1 ) = ArrTipMer(l_rs2("tipopenro"), l_indice_mercaderia - 1 )  + l_rs2("conton")
		l_rs2.movenext
	loop
	l_rs2.close
	
	
	%>
	  <tr>
		  <td colspan="20" align="center" width="5%">&nbsp;</td>
      </tr>		  	
	  <tr>
		  <th colspan="20" align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_rs("sitdes") %>&nbsp;&nbsp;&nbsp;<%= l_anioini %>&nbsp;-&nbsp;<%= l_fecfin %></th>							
      </tr>		  
	  <tr>
		  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Producto</td>							
	<%
	for i = 1 to 4
	%>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= NombreTipoOperacion(i) %></td>							
	<%
	next
	%>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>								
	  </tr>		  							
	<%
	l_cadena5 = ""
	for j = 1 to l_indice_mercaderia - 1
	%>
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(j) %></td>							
	<%
		for i = 1 to 4
			%>
			  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><% if  ArrTipMer(i,j) = 0 then response.write "&nbsp;" else response.write ArrTipMer(i,j) end if %></td>							
			<%
		next
	%>

	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(j) %></td>								
	</tr>
	<%
		l_cadena5 = l_cadena5 & ArrMerDes(j) & "-" & TotFil(j) & ","
	next

	%>
	<tr>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>									
	<%	
	
	' Totales
	for i = 1 to 4
	%>
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= TotCol(i) %></td>			
	<%
	next
	%>
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td></tr>			
	</tr>		
	<tr>
		<td align="center" colspan="6">
	  	  <iframe frameborder="0" name="ifrmgra15b" scrolling="No" src="gra_15b.asp?cadena=<%= l_cadena5 %>" width="720" height="200"></iframe> 
		</td>
	</tr> 		
	<%

	l_rs.movenext
loop
l_rs.close
response.write "</table><p style='page-break-before:always'></p>"
end if 



'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep18 = true then 

'l_nrolinea = 1
'l_nropagina = 1
'l_encabezado = true
'l_corte = false
'l_total = 0

l_sql = "  SELECT distinct(sitdes) ,buq_contenido.sitnro "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND buq_sitio.sittip = 'Combustibles' " 
l_sql = l_sql & " AND ( buq_sitio.sitnro = 11 or buq_sitio.sitnro = 12)" 

rsOpen l_rs, cn, l_sql, 0
l_cadena6 = ""
l_cadena7 = ""

l_Empresa_Aux = l_Empresa

do while not l_rs.eof
	l_nropagina =  l_nropagina + 1
	encabezado_remimpexp("P. Rosales - Removido, Importación y Exportación - Petroleo Crudo") 
	l_nrolinea = 6

	l_Empresa_Aux = ""
	'response.end
	
	' Inicializo 
	for i = 1 to 4
		for j = 1 to 12
			ArrTipMes(i,j) = 0
		next 
	next

	Inicializar_Arreglo TotCol, 50 , 0
	Inicializar_Arreglo TotFil, 50 , 0
	TotFilCol = 0
	
	l_sql = " SELECT  * "
	l_sql = l_sql & " FROM buq_buque "
	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
	l_sql = l_sql & " Order by buq_mercaderia.mernro "
	
	'l_indice_mercaderia = 1
	rsOpen l_rs2, cn, l_sql, 0
	do while not l_rs2.eof
'		if l_merdes <> l_rs2("merdes") then
'			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
'			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
'			l_merdes = l_rs2("merdes")
'			l_indice_mercaderia =  	l_indice_mercaderia + 1
'		end if
		
		TotCol(l_rs2("tipopenro")) = TotCol(l_rs2("tipopenro")) + l_rs2("conton")
		TotFil(month(l_rs2("buqfechas"))) = TotFil(month(l_rs2("buqfechas"))) + l_rs2("conton")
		TotFilCol = TotFilCol + l_rs2("conton")
		
		
		ArrTipMes(l_rs2("tipopenro"), month(l_rs2("buqfechas")) ) = ArrTipMes(l_rs2("tipopenro"), month(l_rs2("buqfechas")) )  + l_rs2("conton")
		l_rs2.movenext
	loop
	l_rs2.close	
	
	%>
	  <tr>
		  <th colspan="20" align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_rs("sitdes") %></th>
      </tr>		  
	  <tr>
		  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mes</td>
		  
	<%
	for i = 1 to 4
	%>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px;"	><%= NombreTipoOperacion(i) %></td>							
	<%
	next
	%>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "	>Total</td>								
	  </tr>		  							
	<%
	
	for j = 1 to month(l_fecfin)
	%>
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= NombreMes(j) %></td>							
	<%
		for i = 1 to 4
			%>
			  <td align="right"  width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; "><% if  ArrTipMes(i,j) = 0 then response.write "0&nbsp;" else response.write ArrTipMes(i,j) & "&nbsp;" end if %></td>
			<%
		next
	%>
	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px;"><%= TotFil(j) %>&nbsp;</td>

	</tr>
	<%
		if l_rs("sitnro") = 11 then
			l_cadena6 = l_cadena6 & NombreMes(j) & "-" & TotFil(j) & ","
		else
			l_cadena7 = l_cadena7 & NombreMes(j) & "-" & TotFil(j) & ","
		end if 
	next

	%>
	<tr>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>									
	<%	
	
	' Totales
	for i = 1 to 4
	%>
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "	><%= TotCol(i) %>&nbsp;</td>			
	<%
	next
	%>
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; "	><%= TotFilCol %>&nbsp;</td></tr>			
	<%
	
	'response.write "</table><p style='page-break-before:always'></p>"
	
	l_rs.movenext
loop
l_rs.close
	%>
	<tr>
		<td align="center" colspan="6">
	  	  <iframe frameborder="0" name="ifrmgra18" scrolling="No" src="gra_18.asp?cadena=<%= l_cadena6 %>&cadena2=<%= l_cadena7 %>" width="720" height="350"></iframe> 
		</td>
	</tr> 			
	<%

response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************
' DETALLE CARGA POR SITIOS
'***************************************************************************************************************************


if l_rep8 = true then 

encabezado_detcargasitio("Detalle de Cargas por Sitio") 
l_nrolinea = 6

l_sql = "  SELECT distinct(sitdes) ,buq_contenido.sitnro "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
rsOpen l_rs, cn, l_sql, 0

Dim l_cantsitios 
l_cantsitios = 0

do while not l_rs.eof

	' Inicializo 
	for i = 1 to 100
		for j = 1 to 12
			ArrMerMes(i,j) = 0
		next 
	next

	'response.write l_rs(0) & " - "
	
	l_sql = " SELECT  * "
	l_sql = l_sql & " FROM buq_buque "
	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true)
	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
	l_sql = l_sql & " Order by buq_mercaderia.mernro "

'	response.write l_sql
	
	l_indice_mercaderia = 1
	l_merdes = ""
	rsOpen l_rs2, cn, l_sql, 0
	do while not l_rs2.eof
		if l_merdes <> l_rs2("merdes") then
			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
			l_merdes = l_rs2("merdes")
			l_indice_mercaderia =  	l_indice_mercaderia + 1
		end if
		ArrMerMes(l_indice_mercaderia - 1, month(l_rs2("buqfechas"))  ) = ArrMerMes(l_indice_mercaderia - 1 , month(l_rs2("buqfechas"))  )  + l_rs2("conton")
		l_rs2.movenext
	loop
	l_rs2.close
	%>
	  <tr>
		  <th colspan="<%= l_indice_mercaderia + 1 %>" align="left" width="5%" nowrap><%= l_rs("sitdes") %></th>							
		  <td colspan="<%= 20 - l_indice_mercaderia - 1 %>"  >&nbsp;</td>			
      </tr>		  
	  <tr>
		  <td align="center" width="5%" style="FONT-SIZE: 8pt; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Mes</td>							
	<%
	l_cadena80 = ""
	for i = 1 to l_indice_mercaderia - 1
	%>
	  <td align="center" width="5%" style="FONT-SIZE: 8pt; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(i) %></td>
	<%
	next
	%>
		<td align="right" width="5%" style="FONT-SIZE: 8pt; border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>	
		<td colspan="<%= 20 - l_indice_mercaderia - 1 %>"  >&nbsp;</td>			
	  </tr>		  							
	<%
	
	for j = month(l_fecfin) to month(l_fecfin)
	%>
	<tr>
	  <td align="center" width="5%" style="FONT-SIZE: 8pt; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= NombreMes(j) %></td>							
	<%
		l_TotMes = 0
		for i = 1 to l_indice_mercaderia - 1
			%>
			  <td align="right" width="5%" style="FONT-SIZE: 8pt; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><% if  ArrMerMes(i,j) = 0 then response.write "&nbsp;" else response.write ArrMerMes(i,j) end if %></td>							
			<%
			l_TotMes = l_TotMes + ArrMerMes(i,j)
			l_cadena80 = l_cadena80 & ArrMerDes(i) & "-" &ArrMerMes(i,j) & ","
		next
		
	%>
	  <td align="right" width="5%" style="FONT-SIZE: 8pt; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= l_TotMes %></td>								
 	  <td colspan="<%= 20 - l_indice_mercaderia - 1 %>"  >&nbsp;</td>		  
	</tr>
	<%
	next
	%>

	<tr>
		<td align="left" colspan="20">
	  	  <iframe frameborder="0" name="ifrmgra80" scrolling="No" src="gra_80.asp?cadena=<%= l_cadena80 %>" width="500" height="230"></iframe> 
		</td>
	</tr>	
	<tr>
		<td colspan="20">&nbsp;</td>
	</tr>
	<%
	l_cantsitios = l_cantsitios + 1
	if l_cantsitios = 3 then
		l_cantsitios = 0
		l_nropagina = l_nropagina + 1
		response.write "</table><p style='page-break-before:always'></p>"
		encabezado_detcargasitio("Detalle de Cargas por Sitio") 
		l_nrolinea = 6
	end if	
	l_rs.movenext
loop
l_rs.close

response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1
end if 



'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************
' Backup


'if l_rep8 = true then 
'
'encabezado_detcargasitio("Detalle de Cargas por Sitio") 
'l_nrolinea = 6
'
'l_sql = "  SELECT distinct(sitdes) ,buq_contenido.sitnro "
'l_sql = l_sql & " FROM buq_buque "
'l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
'l_sql = l_sql & " inner join buq_sitio on buq_sitio.sitnro = buq_contenido.sitnro "
'l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
'l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
'rsOpen l_rs, cn, l_sql, 0
'do while not l_rs.eof
'
'	' Inicializo 
'	for i = 1 to 100
'		for j = 1 to 12
'			ArrMerMes(i,j) = 0
'		next 
'	next
'
'	'response.write l_rs(0) & " - "
'	
'	l_sql = " SELECT  * "
'	l_sql = l_sql & " FROM buq_buque "
'	l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
'	l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
'	l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
'	l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
'	l_sql = l_sql & " AND buq_contenido.sitnro = " & l_rs(1)
'	l_sql = l_sql & " Order by buq_mercaderia.mernro "
'	
'	l_indice_mercaderia = 1
'	l_merdes = ""
'	rsOpen l_rs2, cn, l_sql, 0
'	do while not l_rs2.eof
'		if l_merdes <> l_rs2("merdes") then
'			ArrMerNro(l_indice_mercaderia) = l_rs2("mernro")
'			ArrMerDes(l_indice_mercaderia) = l_rs2("merdes")
'			l_merdes = l_rs2("merdes")
'			l_indice_mercaderia =  	l_indice_mercaderia + 1
'		end if
'		ArrMerMes(l_indice_mercaderia - 1, month(l_rs2("buqfechas"))  ) = ArrMerMes(l_indice_mercaderia - 1 , month(l_rs2("buqfechas"))  )  + l_rs2("conton")
'		l_rs2.movenext
'	loop
'	l_rs2.close
'	
'	
'	%>
	<!--
	  <tr>
		  <th colspan="3" align="left" width="5%" nowrap><%'= l_rs("sitdes") %></th>							
      </tr>		  
	  <tr>
		  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Mes</td>							
	-->		  
	<%
'	for i = 1 to l_indice_mercaderia - 1
	%>
	<!--
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%'= ArrMerDes(i) %></td>							
	  -->
	<%
'	next
	%><!--
		<td align="right" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>	
	  </tr>		  							
	  -->
	<%
	
'	for j = 1 to month(l_fecfin)
	%><!--
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%'= NombreMes(j) %></td>							
	  -->
	<%
'		l_TotMes = 0
'		for i = 1 to l_indice_mercaderia - 1
'			%><!--
			  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%' if  ArrMerMes(i,j) = 0 then response.write "&nbsp;" else response.write ArrMerMes(i,j) end if %></td>							
			  -->
			<%
'			l_TotMes = l_TotMes + ArrMerMes(i,j)
'		next
	%><!--
	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%'= l_TotMes %></td>								
	</tr>
	-->
	<%
'	next
	%><!--
	<tr>		
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>										
		-->	  
	<%
'	l_TotTotMerMes = 0
'	for i = 1 to l_indice_mercaderia - 1
'		l_TotMer = 0
'		for j = 1 to month(l_fecfin)
'			l_TotMer = l_TotMer + ArrMerMes(i,j)
'		next
'		l_TotTotMerMes = l_TotTotMerMes + l_TotMer
		%><!--
		  <td align="right" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%'= l_TotMer %></td>							
		  -->
		<%
'	next
	%><!--
	  <td align="right" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%'= l_TotTotMerMes %></td>							
	</tr>	  	
	<tr>
		<td>&nbsp;</td>
	</tr>-->
	<%
'	l_rs.movenext
'loop
'l_rs.close

'response.write "</table><p style='page-break-before:always'></p>"

'end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep9 = true then 

encabezado_porparsitcas("Porcentaje de Participación por Sitios - CAS") 
l_nrolinea = 6

l_cadena = ""

for x = 1 to l_indice_mercaderia - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrMerDes(x) %></td>			
	<%
	for y = 1 to l_indice_sitio - 1
	%>
		<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >&nbsp;<%= MatSitMer(y,x) %></td>			
		<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><% if TotFil(x) <> 0 then response.write formatnumber((MatSitMer(y,x) * 100) / TotFil(x),2) else response.write "&nbsp;" end if %></td>		
	<%	
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(x) %></td>			
	</tr>		
	<%
	
	l_cadena = l_cadena & ArrMerDes(x) & "-" & TotFil(x) & ","
	
next

%>
	<tr>
		<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >Total</td>			
<%

' Totales
l_cadena2 = ""
for i = 1 to l_indice_sitio - 1
	if TotFilCol <> 0 then 
		l_porcentaje = formatnumber( (TotCol(i) * 100) / TotFilCol,2) 
		l_cadena2 = l_cadena2 & ArrSitDes(i) & "-" & l_porcentaje & "@"
	else 
		l_porcentaje = "&nbsp;" 
	end if
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= TotCol(i) %></td>			
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= l_porcentaje  %></td>		
<%
next

%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFilCol %></td>			
</tr>	
<tr>
	<td colspan="12">
	<table cellpadding="0" cellspacing="0" border="0">
		<tr>					
			<td align="left" width="50%">
		  	  <iframe frameborder="0" name="ifrmgra09" scrolling="No" src="gra_09b.asp?cadena=<%= l_cadena2 %>" width="100%" height="250"></iframe> 
			</td>	
		</tr>	
		<tr>
			<td align="left" width="50%">
		  	  <iframe frameborder="0" name="ifrmgra09" scrolling="No" src="gra_09.asp?cadena=<%= l_cadena %>" width="100%" height="300"></iframe> 
			</td>
		</tr>
	</table>
	</td>
</tr>  		

<%
response.write "</table><p style='page-break-before:always'></p>"
end if 

'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep10 = true then 

l_nropagina = l_nropagina + 1
encabezado_parter("Participación por Terminal") 
l_nrolinea = 6

l_cadena3 = ""
for x = 1 to l_indice_terminal - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrTerDes(x) %></td>			
	<%
	for y = 1 to l_indice_mercaderia - 1
		if MatMerTer(y,x) = "" then
		%>
			<td align="right"  width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" >&nbsp;</td>			
		<%
		else
		%>
			<td align="right"  width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= MatMerTer(y,x) %></td>			
		<%
		end if
		
	next
	%>
		<td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(x) %></td>			
	</tr>	
	<%
	l_cadena3 = l_cadena3 & ArrTerDes(x) & "-" & TotFil(x) & ","
next

%>
<tr>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">Total</td>				
<%

for j = 1 to l_indice_mercaderia - 1
%>
	<td align="right"  width="7%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotCol(j) %></td>			
<%
next
%>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFilCol %></td>				
</tr>	

<tr>
	<td colspan="18">
	<table cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="left" width="100%">
		  	  <iframe frameborder="0" name="ifrmgra10" scrolling="No" src="gra_10.asp?cadena=<%= l_cadena3 %>" width="100%" height="300"></iframe> 
			</td>
		</tr>
	</table>
	</td>
</tr>  		

<%
response.write "</table><p style='page-break-before:always'></p>"
end if 




'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep11 = true then 

encabezado_exppes("Exportación de Pescado Congelado") 

'response.end

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_destino on buq_destino.desnro = buq_contenido.desnro "
l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " WHERE  buq_buque.tipopenro = 3 "  ' EXPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " AND buq_contenido.mernro = 35 " ' Pescado Congelado
l_sql = l_sql & " ORDER BY buq_destino.desdes " 

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end

if not l_rs.eof then
	l_desdes = ""
end if

l_indice_destino = 1
l_indice_exportadora = 1

Inicializar_Arreglo ArrDesNro, 50 , 0
Inicializar_Arreglo ArrDesDes, 50 , 0
Inicializar_Arreglo ArrExpNro, 50 , 0
Inicializar_Arreglo ArrExpDes, 50 , 0
Inicializar_Arreglo TotCol, 50 , 0
Inicializar_Arreglo TotFil, 50 , 0
Inicializar_Arreglo TotFil2, 50 , 0
TotFilCol = 0

do until l_rs.eof

	if l_desdes <> l_rs("desdes") then
		ArrDesNro(l_indice_destino) = l_rs("desnro")
		ArrDesDes(l_indice_destino) = l_rs("desdes")
		l_desdes = l_rs("desdes")
		l_indice_destino = l_indice_destino + 1
	end if
	
	l_existe = false
	for x = 1 to l_indice_exportadora - 1
		if l_rs("expnro") = ArrExpNro(x) then
			l_existe = true
			l_ColMer = x
		end if 
	next
	if l_existe = false then
		ArrExpNro(l_indice_exportadora) = l_rs("expnro")
		ArrExpDes(l_indice_exportadora) = l_rs("expdes")
		l_ColMer = l_indice_exportadora
		l_indice_exportadora = l_indice_exportadora + 1
	end if 	
	

	MatMesDes(month(l_rs("buqfechas")) , l_indice_destino -1) = MatMesDes(month(l_rs("buqfechas")) , l_indice_destino -1) + l_rs("conton")			
	TotFil(l_indice_destino -1) = TotFil(l_indice_destino -1) + l_rs("conton")
	TotCol(month(l_rs("buqfechas"))) = TotCol(month(l_rs("buqfechas"))) + l_rs("conton")
	
	MatMesExp(month(l_rs("buqfechas")) , l_indice_exportadora -1) = MatMesExp(month(l_rs("buqfechas")) , l_indice_exportadora -1) + l_rs("conton")			
	TotFil2(l_indice_exportadora -1) = TotFil2(l_indice_exportadora -1) + l_rs("conton")
	
	TotFilCol = TotFilCol + l_rs("conton")
	
	l_rs.MoveNext
loop
l_rs.Close

'---------------------------------
' por destino
'---------------------------------
for i = 1 to l_indice_destino - 1
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrDesDes(i) %></td>			
<%	
	for l_Mes = 1 to 12
		if MatMesDes(l_Mes, i) = "" then
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;</td>			
		<%
		else
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= MatMesDes(l_Mes, i ) %></td>			
		<%
		end if 
	next
%>
	<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(i) %></td>			
</tr>	
<%
next
%>
<tr>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>				
<%

for l_Mes = 1 to 12
%>
	<td align="right"  width="7%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotCol(l_Mes) %></td>			
<%
next
%>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>				
</tr>	
<%


'---------------------------------
' por exportadora
'---------------------------------
%>
	    <tr>
	        <td align="center" colspan="14">Exportadora</td>
	    </tr>		

	    <tr>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Exportadora</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>ENE</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>FEB</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>MAR</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>ABR</th>		
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>MAY</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>JUN</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>JUL</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>AGO</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>SEP</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>OCT</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>NOV</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>DIC</th>												
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>TON</th>			
	    </tr>		
<%
for i = 1 to l_indice_exportadora - 1
%>
<tr>
	<td align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrExpDes(i) %></td>			
<%	
	for l_Mes = 1 to 12
		if MatMesExp(l_Mes, i ) = "" then
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;">&nbsp;</td>			
		<%
		else
		%>
		<td align="right"  width="7%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= MatMesExp(l_Mes, i ) %></td>			
		<%
		end if
	next
%>
	<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil2(i) %></td>			
</tr>	
<%
next
%>
<tr>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>				
<%

for l_Mes = 1 to 12
%>
	<td align="right"  width="7%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotCol(l_Mes) %></td>			
<%
next
%>
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>				
</tr>	
<%

response.write "</table><p style='page-break-before:always'></p>"
end if





'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep12 = true then 

encabezado_expinf("Exportación Inflamables") 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

for j = 1 to month(l_fecfin)
%>
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= NombreMes(j) %></td>							
<%
	for i = 1 to l_indice_mercaderia - 1
	%>
	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><% if  ArrMerMes(i,j) = 0 then response.write "&nbsp;" else response.write ArrMerMes(i,j) end if %></td>							
	<%
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(j) %></td>			
	</tr>		
<%
next
%>
	<tr>
		<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 >Total</td>			
<%
' Totales
for i = 1 to l_indice_mercaderia - 1
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= TotCol(i) %></td>			
<%
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>			
</tr>	

<tr>
	<td colspan="14">&nbsp;
	</td>
</tr>

<tr>
	<th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mes</th>
<%
for i = 1 to l_indice_mercaderia - 1
%>
  <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= ArrMerDes(i) %></th>
<%
next
%>
	<th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</th>
  </tr>		  							
<%

for j = 1 to l_indice_exportadora - 1
%>
	<tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrExpdes(j) %></td>							
<%
	for i = 1 to l_indice_mercaderia - 1
	%>
	  <td align="right" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><% if  MatMerExp(i,j) = 0 then response.write "&nbsp;" else response.write MatMerExp(i,j) end if %></td>							
	<%
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil2(j) %></td>			
	</tr>		
<%
next
%>
<tr>
	<td align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 >Total</td>			
<%
' Totales
for i = 1 to l_indice_mercaderia - 1
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	 ><%= TotCol(i) %></td>			
<%
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>			
</tr>	
<%

l_totcaston = 0
response.write "</table><p style='page-break-before:always'></p>"

end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep19 = true then 

l_nropagina = l_nropagina + 1
encabezado_MovBuqSitMes("Movimientos de Buques por Sitio") 
l_nrolinea = 6

'l_nrolinea = 1
'l_nropagina = 1
'l_encabezado = true
'l_corte = false
'l_total = 0

dim ran_buqsit_nom1
dim ran_buqsit_val1
dim ran_buqsit_nom2
dim ran_buqsit_val2
dim ran_buqsit_nom3
dim ran_buqsit_val3

dim ran_buqsit_acu_nom1
dim ran_buqsit_acu_val1
dim ran_buqsit_acu_nom2
dim ran_buqsit_acu_val2
dim ran_buqsit_acu_nom3
dim ran_buqsit_acu_val3

dim ran_buqtip_nom1
dim ran_buqtip_val1
dim ran_buqtip_nom2
dim ran_buqtip_val2
dim ran_buqtip_nom3
dim ran_buqtip_val3

dim ran_buqtip_acu_nom1
dim ran_buqtip_acu_val1
dim ran_buqtip_acu_nom2
dim ran_buqtip_acu_val2
dim ran_buqtip_acu_nom3
dim ran_buqtip_acu_val3

ran_buqsit_nom1 = 0 
ran_buqsit_val1 = 0
ran_buqsit_nom2 = 0
ran_buqsit_val2 = 0
ran_buqsit_nom3 = 0
ran_buqsit_val3 = 0

ran_buqsit_acu_nom1 = 0
ran_buqsit_acu_val1 = 0
ran_buqsit_acu_nom2 = 0
ran_buqsit_acu_val2 = 0
ran_buqsit_acu_nom3 = 0
ran_buqsit_acu_val3 = 0

ran_buqtip_nom1 = 0
ran_buqtip_val1 = 0
ran_buqtip_nom2 = 0
ran_buqtip_val2 = 0
ran_buqtip_nom3 = 0
ran_buqtip_val3 = 0

ran_buqtip_acu_nom1 = 0
ran_buqtip_acu_val1 = 0
ran_buqtip_acu_nom2 = 0
ran_buqtip_acu_val2 = 0
ran_buqtip_acu_nom3 = 0
ran_buqtip_acu_val3 = 0

Dim ArrSitMes(100,100)
Dim i
Dim j
Dim k

for i = 1 to 100
	for j = 1 to 12
		ArrSitMes(i,j) = 0
	next
next

l_sql = " SELECT buq_buque.buqnro, sitnro, buqfechas "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

'l_sql = l_sql & " AND buq_contenido.sitnro = 1 "
l_sql = l_sql & " group by buq_buque.buqnro, sitnro, buqfechas "
rsOpen l_rs, cn, l_sql, 0

dim valor
Dim l_buquenumero
dim l_sitionumero

do while not l_rs.eof

	ArrSitMes(l_rs("sitnro") , month(l_rs("buqfechas"))  ) = ArrSitMes(l_rs("sitnro") , month(l_rs("buqfechas"))  ) + 1
	
	l_rs.movenext
loop
l_rs.close


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_sitio "
l_sql = l_sql & " order by buq_sitio.sitnro "

rsOpen l_rs, cn, l_sql, 0
%>
	<tr>
        <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mes</th>
<%

Dim ArrNomSit(50)
Dim ArrTotSit(50)

i = 0
do while not l_rs.eof
%>
    <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_rs("sitdes") %></th>
<%
	i = i + 1
	ArrNomSit(i) = l_rs("sitdes")
	l_rs.movenext
loop
l_rs.close
%>
	</tr>
<%
l_cadena2 = ""
for j = 1 to month(l_fecfin)
%>
  <tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= NombreMes(j) %></td>							
<%

	for k = 1 to i
	
		if j = month(l_fecfin) then
			l_cadena2 = l_cadena2 & ArrNomSit(k) & "-" & ArrSitMes(k,j) & "@"
		'response.write ArrSitMes(k,j) & "<br>"	
		end if
		'response.write l_cadena2
		
		%>
		  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrSitMes(k,j) %></td>							
		<%
		ArrTotSit(k) = ArrTotSit(k) + ArrSitMes(k,j)
		
		if j = month(l_fecfin) then
		
				'---------------------mayor
				if ArrSitMes(k,j) >= ran_buqsit_val1 then 
				
					ran_buqsit_nom3 = ran_buqsit_nom2
					ran_buqsit_val3 = ran_buqsit_val2
					
					ran_buqsit_nom2 = ran_buqsit_nom1
					ran_buqsit_val2 = ran_buqsit_val1
					
					ran_buqsit_nom1 = ArrNomSit(k)
					ran_buqsit_val1 = ArrSitMes(k,j)


				else
					if ArrSitMes(k,j) >= ran_buqsit_val2 then
						ran_buqsit_nom3 = ran_buqsit_nom2
						ran_buqsit_val3 = ran_buqsit_val2
						
						ran_buqsit_nom2 = ArrNomSit(k)
						ran_buqsit_val2 = ArrSitMes(k,j)
			
					else 
						if ArrSitMes(k,j) >= ran_buqsit_val3 then
							ran_buqsit_nom3 = ArrNomSit(k)
							ran_buqsit_val3 = ArrSitMes(k,j)
						end if
					end if
				end if	
				'---------------------fin mayor		
		end if		
	next
	
%>
	</tr>
<%
next

%>
  <tr>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><b>Tot</b></td>
<%
'l_cadena2 = ""
for k = 1 to i
%>
	<td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= ArrTotSit(k) %></td>
<%
	'l_cadena2 = l_cadena2 & ArrSitDes(k) & "-" & ArrTotSit(k) & "@"
	
	'---------------------mayor
	if ArrTotSit(k) >= ran_buqsit_acu_val1 then 
	
		ran_buqsit_acu_nom3 = ran_buqsit_acu_nom2
		ran_buqsit_acu_val3 = ran_buqsit_acu_val2
		
		ran_buqsit_acu_nom2 = ran_buqsit_acu_nom1
		ran_buqsit_acu_val2 = ran_buqsit_acu_val1
		
		ran_buqsit_acu_nom1 = ArrNomSit(k)
		ran_buqsit_acu_val1 = ArrTotSit(k)


	else
		if ArrTotSit(k) >= ran_buqsit_acu_val2 then
			ran_buqsit_acu_nom3 = ran_buqsit_acu_nom2
			ran_buqsit_acu_val3 = ran_buqsit_acu_val2
			
			ran_buqsit_acu_nom2 = ArrNomSit(k)
			ran_buqsit_acu_val2 = ArrTotSit(k)

		else 
			if ArrTotSit(k) >= ran_buqsit_acu_val3 then
				ran_buqsit_acu_nom3 = ArrNomSit(k)
				ran_buqsit_acu_val3 = ArrTotSit(k)
			end if
		end if
	end if	
	'---------------------fin mayor
next
%>
  </tr>
<tr>
	<td colspan="12">
	<table cellpadding="0" cellspacing="0" border="0">
		<tr>					
			<td align="left" width="50%">
		  	  <iframe frameborder="0" name="ifrmgra19" scrolling="No" src="gra_19a.asp?cadena=<%= l_cadena2 %>" width="100%" height="250"></iframe> 
			</td>	
		</tr>	
	</table>
	</td>
</tr>  		  
  <tr>
	  <td colspan="20">&nbsp;</td>							    
  </tr>	  
<%

encabezado_MovBuqSitMes2("Clase y Cantidad de Buques") 

'l_nrolinea = 1
'l_nropagina = 1
'l_encabezado = true
'l_corte = false
'l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

Dim ArrTipBuqMes(100,100)

for i = 1 to 100
	for j = 1 to 12
		ArrTipBuqMes(i,j) = 0
	next
next

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)

'l_sql = l_sql & " AND buq_buque.tipbuqnro = 7 "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql

do while not l_rs.eof

	ArrTipBuqMes(l_rs("tipbuqnro") , month(l_rs("buqfechas"))  ) = ArrTipBuqMes(l_rs("tipbuqnro") , month(l_rs("buqfechas"))  ) + 1
	
	'response.write l_rs("tipbuqnro") & " ---" & month(l_rs("buqfechas")) & "<br>"
	
	'response.write ArrTipBuqMes(2,1) & "<br>"
	
	l_rs.movenext
loop
l_rs.close

'response.write ArrTipBuqMes(2,1)


l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_tipobuque "
l_sql = l_sql & " order by buq_tipobuque.tipbuqnro "
rsOpen l_rs, cn, l_sql, 0
%>
	<tr>
        <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Mes</th>
<%

Dim ArrNomTipBuq(50)
Dim ArrTotTipBuq(50)
Dim ArrTotTipBuqMes(50)

i = 0
do while not l_rs.eof
%>
    <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_rs("tipbuqdes") %></th>
<%
	i = i + 1
	ArrNomTipBuq(i) = l_rs("tipbuqdes")
	l_rs.movenext
loop
l_rs.close
%>
        <th align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Totales</th>
	</tr>
<%

for j = 1 to month(l_fecfin)
%>
  <tr>
	  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= NombreMes(j) %></td>							
<%

	for k = 1 to i
		%>
		  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrTipBuqMes(k,j) %></td>							
		<%
		ArrTotTipBuq(k) = ArrTotTipBuq(k) + ArrTipBuqMes(k,j)
		ArrTotTipBuqMes(j) = ArrTotTipBuqMes(j) + ArrTipBuqMes(k,j)
		
		if j = month(l_fecfin) then
		
				'---------------------mayor
				if ArrTipBuqMes(k,j) >= ran_buqtip_val1 then 
				
					ran_buqtip_nom3 = ran_buqtip_nom2
					ran_buqtip_val3 = ran_buqtip_val2
					
					ran_buqtip_nom2 = ran_buqtip_nom1
					ran_buqtip_val2 = ran_buqtip_val1
					
					ran_buqtip_nom1 = ArrNomTipBuq(k)
					ran_buqtip_val1 = ArrTipBuqMes(k,j)
					


				else
					if ArrTipBuqMes(k,j) >= ran_buqtip_val2 then
						ran_buqtip_nom3 = ran_buqtip_nom2
						ran_buqtip_val3 = ran_buqtip_val2
						
						ran_buqtip_nom2 = ArrNomTipBuq(k)
						ran_buqtip_val2 = ArrTipBuqMes(k,j)
			
					else 
						if ArrTipBuqMes(k,j) >= ran_buqtip_val3 then
							ran_buqtip_nom3 =  ArrNomTipBuq(k)
							ran_buqtip_val3 = ArrTipBuqMes(k,j)
						end if
					end if
				end if	
				'---------------------fin mayor		
		end if				
		
	next
		%>
		  <td align="center" width="5%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrTotTipBuqMes(j) %></td>							
		<%
%>
	</tr>
<%
next


%>
  <tr>
	  <td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><b>Tot</b></td>							  
<%
Dim l_tottot
l_tottot = 0

for k = 1 to i
%>
	<td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= ArrTotTipBuq(k) %></td>							
<%
	'---------------------mayor
	if ArrTotTipBuq(k) >= ran_buqtip_acu_val1 then 
	
		ran_buqtip_acu_nom3 = ran_buqtip_acu_nom2
		ran_buqtip_acu_val3 = ran_buqtip_acu_val2
		
		ran_buqtip_acu_nom2 = ran_buqtip_acu_nom1
		ran_buqtip_acu_val2 = ran_buqtip_acu_val1
		
		ran_buqtip_acu_nom1 = ArrNomTipBuq(k)
		ran_buqtip_acu_val1 = ArrTotTipBuq(k)


	else
		if ArrTotTipBuq(k) >= ran_buqtip_acu_val2 then
			ran_buqtip_acu_nom3 = ran_buqtip_acu_nom2
			ran_buqtip_acu_val3 = ran_buqtip_acu_val2
			
			ran_buqtip_acu_nom2 = ArrNomTipBuq(k)
			ran_buqtip_acu_val2 = ArrTotTipBuq(k)

		else 
			if  ArrTotTipBuq(k) >= ran_buqtip_acu_val3 then
				ran_buqtip_acu_nom3 = ArrNomTipBuq(k)
				ran_buqtip_acu_val3 = ArrTotTipBuq(k)
			end if
		end if
	end if	
	'---------------------fin mayor


	l_tottot = l_tottot + ArrTotTipBuq(k) 
next
%>
	<td align="center" width="5%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= l_tottot %></td>							
  </tr>
<%


response.write "</table><p style='page-break-before:always'></p>"
end if


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep20 = true then 

l_nropagina = l_nropagina + 1
encabezado_movgen("Movimiento General")
l_nrolinea = 6

l_encabezado = true
l_corte = false
l_total = 0

for x = 1 to l_indice_mercaderia - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;" ><%= ArrMerDes(x) %></td>			
	<%
	for y = 1 to l_indice_sitio - 1
	%>
		<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= MatSitMer(y,x) %></td>			
	<%
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(x) %></td>			
	</tr>	
	<%
next
%>
	<tr>	
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>			
<%
' Totales
for i = 1 to l_indice_sitio - 1
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotCol(i) %></td>			
<%
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>			
</tr>	
<%

encabezado_movgeninf("")

for x = 1 to l_indice_mercaderia - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(x) %></td>			
	<%
	for y = 1 to l_indice_sitio - 1
	%>
		<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= MatSitMer(y,x) %></td>			
	<%
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(x) %></td>			
	</tr>	
	<%
next
%>
	<tr>	
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>			
<%
' Totales
for i = 1 to l_indice_sitio - 1
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotCol(i) %></td>			
<%
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>			
</tr>	
<%

encabezado_movgenotr("")

for x = 1 to l_indice_mercaderia - 1
	%>
	<tr>
		<td nowrap align="center" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= ArrMerDes(x) %></td>			
	<%
	for y = 1 to l_indice_sitio - 1
	%>
		<td align="right"  width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= MatSitMer(y,x) %></td>			
	<%
	next
	%>
		<td align="right" width="10%" style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><%= TotFil(x) %></td>			
	</tr>	
	<%
next
%>
	<tr>	
	<td align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	>Total</td>			
<%
' Totales
for i = 1 to l_indice_sitio - 1
%>
	<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotCol(i) %></td>			
<%
next
%>	
	<td  align="right" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px;border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"	><%= TotFilCol %></td>			
</tr>	
<%

response.write "</table><p style='page-break-before:always'></p>"
end if 


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

if l_rep21 = true then 

l_nropagina = l_nropagina + 1
encabezado_detatebuqage("Detalle de Atención Buques por Agencia") 
l_nrolinea = 6

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT distinct(buq_agencia.agedes), count(*) "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
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
	  	  <iframe frameborder="0" name="ifrmgra21" scrolling="No" src="gra_21.asp?anioini=<%= l_anioini %>&fecfin=<%= l_fecfin %>" width="600" height="300"></iframe> 
		</td>
		</tr>  		
<%
response.write "</table><p style='page-break-before:always'></p>"
l_rs.close
end if




response.end %>


<tr>
   <th align="center" colspan="6">Exportación CAS por Destino</th>
</tr>
<tr>
   <th align="center" colspan="3" >Del Mes</th>					
   <th align="center" colspan="3">Del Año</th>
</tr>
<tr>   
   <td align="center">1</td>					
   <td align="center"><%= ran_descas_acu_nom1 %></td>					
   <td align="center"><%= ran_descas_acu_val1 %></td>					   
</tr>
<tr>
   <td align="center">2</td>					
   <td align="center"><%= ran_descas_acu_nom2 %></td>					
   <td align="center"><%= ran_descas_acu_val2 %></td>					   
</tr>
<tr>   
   <td align="center">3</td>					
   <td align="center"><%= ran_descas_acu_nom3 %></td>					
   <td align="center"><%= ran_descas_acu_val3 %></td>					   
</tr>


<tr>
   <th align="center" colspan="6">Movimientos de Buques por Sitio</th>
</tr>
<tr>
   <th align="center" colspan="3" >Del Mes</th>					
   <th align="center" colspan="3">Del Año</th>
</tr>
<tr>   
  <td align="center">1</td>					
   <td align="center"><%= ran_buqsit_nom1 %></td>
   <td align="center"><%= ran_buqsit_val1 %></td>
   <td align="center">1</td>					
   <td align="center"><%= ran_buqsit_acu_nom1 %></td>					
   <td align="center"><%= ran_buqsit_acu_val1 %></td>					   
</tr>
<tr>
  <td align="center">2</td>					
   <td align="center"><%= ran_buqsit_nom2 %></td>
   <td align="center"><%= ran_buqsit_val2 %></td>
   <td align="center">2</td>					
   <td align="center"><%= ran_buqsit_acu_nom2 %></td>					
   <td align="center"><%= ran_buqsit_acu_val2 %></td>					   
</tr>
<tr>   
  <td align="center">3</td>					
   <td align="center"><%= ran_buqsit_nom3 %></td>
   <td align="center"><%= ran_buqsit_val3 %></td>
   <td align="center">3</td>					
   <td align="center"><%= ran_buqsit_acu_nom3 %></td>					
   <td align="center"><%= ran_buqsit_acu_val3 %></td>					   
</tr>

<tr>
   <th align="center" colspan="6">Clases de Buques</th>
</tr>
<tr>
   <th align="center" colspan="3" >Del Mes</th>					
   <th align="center" colspan="3">Del Año</th>
</tr>
<tr>   
  <td align="right" >1</td>					
   <td align="center"><%= ran_buqtip_nom1 %></td>
   <td align="center"><%= ran_buqtip_val1 %></td>
   <td align="right">1</td>					
   <td align="center"><%= ran_buqtip_acu_nom1 %></td>					
   <td align="center"><%= ran_buqtip_acu_val1 %></td>					   
</tr>
<tr>
  <td align="right">2</td>					
   <td align="center"><%= ran_buqtip_nom2 %></td>
   <td align="center"><%= ran_buqtip_val2 %></td>
   <td align="right">2</td>					
   <td align="center"><%= ran_buqtip_acu_nom2 %></td>					
   <td align="center"><%= ran_buqtip_acu_val2 %></td>					   
</tr>
<tr>   
  <td align="right">3</td>					
   <td align="center"><%= ran_buqtip_nom3 %></td>
   <td align="center"><%= ran_buqtip_val3 %></td>
   <td align="right">3</td>					
   <td align="center"><%= ran_buqtip_acu_nom3 %></td>					
   <td align="center"><%= ran_buqtip_acu_val3 %></td>					   
</tr>

<%

response.end 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_total = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT distinct(buq_agencia.agedes), count(*) "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE buq_buque.buqfechas >= " & cambiafecha(l_anioini,"YMD",true)
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
	  	  <iframe frameborder="0" name="ifrmgra21" scrolling="No" src="gra_21.asp?anioini=<%= l_anioini %>&fecfin=<%= l_fecfin %>" width="600" height="300"></iframe> 
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

