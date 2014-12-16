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

l_anioini = "01/01/" & year(l_fecfin)

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
							<b><%= titulo%></b> 
						</td>
						<!--
				       	<td align="right" nowrap width="5%" > 
							P&aacute;gina: <%'= l_nropagina%>
						</td>				
						-->
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

<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>
function Buque(buqdes){

	param = "qfecini=<%= l_fecini %>&qfecfin=<%= l_fecfin %>&qbuqdes=" + buqdes ;
	
   	abrirVentana("rep_imp_detalle_buques_03.asp?" + param ,'',780,580);	
	//parent.frames.ifrm.focus();
	//window.print();	
}
</script>

	
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

l_nropagina = 1

'l_nropagina = 1
encabezado_expbuq("Importación - Detalle de Buques") 
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

l_sql = l_sql & " WHERE  buq_buque.tipopenro = 4 "  ' IMPORTACION
l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
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
				<td align="left" width="10%" nowrap style="border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-left-color: #000000; border-left-style: solid; border-left-width: 1px;"><a href="Javascript:Buque('<%= l_rs("buqdes") %>');"><img alt="Ver Detalle del Buque" src="/serviciolocal/shared/images/cal.gif" border="0"></a>&nbsp;&nbsp;&nbsp;<%=l_rs("buqdes")%></td>			
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


'***************************************************************************************************************************
'***************************************************************************************************************************
'***************************************************************************************************************************

set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

