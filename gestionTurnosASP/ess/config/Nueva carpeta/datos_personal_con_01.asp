<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_vessel_tender_con_01.asp
Autor: Raul Chinestra
Creacion: 01/02/2008
Descripcion: Reporte de Vessel Tender
 -----------------------------------------------------------------------------
-->
<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 50
Const l_cantcols = 19

Dim l_rs
Dim l_rs2

Dim l_sql

Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_total 
dim l_costo_total

dim l_fecini
dim l_fecfin
dim l_quatypnro
Dim l_pordes


dim l_nomcomdesabr
dim l_titulo_quality
dim l_titulo_Country
dim l_titulo_Port


Dim l_totpornom
Dim l_totcomnom
Dim l_totvesnom
Dim l_totporloa
Dim l_totcomloa
Dim l_totvesloa

dim l_nrohas
dim l_medtra
dim l_lugnro
dim l_cieter
dim l_pronro

'Variable usadas para imprimir los Totales
dim l_TotBruto
dim l_TotTara
dim l_TotKilosMermas
dim l_TotNeto
dim l_TotkilosProc
dim l_TotDifer
dim l_nroope

dim l_nompro
dim l_PesoBruto
dim l_PesoTara

dim l_valorbase
dim l_ValorObservado
dim l_humedos
dim l_secos

l_humedos = 0
l_secos = 0

dim l_NetoxHum
dim l_empini
dim l_proini
dim l_totemppro

' Imprime el encabezado de cada pagina
sub encabezado(titulo)
	%>
		<table>
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">
			<table cellpadding="0" cellspacing="0">
				<tr>
			       	<td width="20%">&nbsp;</td>				
					<td align="center" width="80%"><b><%= titulo%></b></td>
			       	<td align="right" valign="top"  nowrap width="20%">Página: <%= l_nropagina%></td>				
				</tr>
				<tr>
			       	<td width="20%">&nbsp;</td>				
					<td align="center" width="80%"><b>Commodity:</b>&nbsp;<%= l_titulo_quality%></td>
			       	<td align="right" valign="top"  nowrap width="20%">&nbsp;</td>				
				</tr>
				<tr>
			       	<td width="20%">&nbsp;</td>				
			       	<td align="center" width="80%">
						<b>Country:</b>&nbsp;<%= l_titulo_Country %>&nbsp;&nbsp;&nbsp;&nbsp;
						<b>Port:</b>&nbsp;<%= l_titulo_Port %>&nbsp;&nbsp;&nbsp;&nbsp;						
						<b>From:</b>&nbsp;<%= l_fecini %> <b>To:</b>&nbsp; <%= l_fecfin %>
					</td>
			       	<td align="right" valign="top"  nowrap width="20%">&nbsp;</td>									
				</tr>								
			</table>
			</td>				
		</tr>		
	<%
end sub 'encabezado


sub totalesEmpresa()
	%>
	<tr>
		<td align="left" colspan="<%= l_cantcols%>">&nbsp;</td>					
	</tr>	
	<tr>
		<td align="center" colspan="6"><b>Totales por Empresa</b></td>					
		<td align="center" colspan="7">&nbsp;</td>					
	</tr>	
	</td></tr>		
	<tr>
		<td align="left" colspan="2"><b>Empresa</b></td>					
		<td align="left" colspan="2"><b>Producto</b></td>
		<td align="center" colspan="2"><b>Kilos</b></td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>
		<td align="center">&nbsp;</td>	
		</tr>	
	
	
	
	<%
	dim l_rs
	dim l_sql
	
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	
	l_sql = " select movnro,empnro,empdes,prodes "
	l_sql = l_sql & " from tkt_movimiento "
	l_sql = l_sql & " inner join tkt_cartaporte on tkt_cartaporte.carpornro = tkt_movimiento.carpornro "
	
		' Condiciones para contemplar los filtros ingresados
		
		' Fecha desde - Fecha Hasta
		if l_fecini <> "" and l_fecfin <> "" then 	
			l_sql = l_sql & " WHERE ( tkt_movimiento.movfec >= " & cambiafecha(l_fecini,"YMD",true) 
			l_sql = l_sql & "   AND tkt_movimiento.movfec <= " & cambiafecha(l_fecfin,"YMD",true) & " ) "
		else 	
			l_sql = l_sql & " WHERE 1 = 1  " ' Se coloco esta condicion para que aparezca una sola vez la sentencia WHERE
		end if
		
		' Nro Operación Desde - Nro Operación Hasta
		if l_nrodes <> "" and l_nrohas <> "" then
			l_sql = l_sql & " AND ( tkt_movimiento.movnro >= " & l_nrodes 
			l_sql = l_sql & "   AND tkt_movimiento.movnro <= " & l_nrohas & " ) "
		end if
	
		' Camion - Vagon - Ambos
		select case l_medtra
			case "C" ' Camion
				l_sql = l_sql & " AND ( tkt_cartaporte.carpormed = 'C' ) "
			case "V" ' Vagon
				l_sql = l_sql & " AND ( tkt_cartaporte.carpormed = 'V' ) "
			case "A" ' Ambos
			
		end select
		
		' Lugar
		if l_lugnro <> 0 then
			l_sql = l_sql & " AND ( tkt_cartaporte.lugorinro = " & l_lugnro & " ) "
		end if
	
		' Producto
		if l_pronro <> 0 then
			l_sql = l_sql & " AND ( tkt_cartaporte.pronro = " & l_pronro & " ) "
		end if
		
		l_sql = l_sql & " order by tkt_cartaporte.prodes,tkt_cartaporte.empdes "	
		rsOpen l_rs, cn, l_sql, 0 
		
		if not l_rs.eof then	
			l_empini = l_rs("empdes")
			l_proini = l_rs("prodes")
		end if	
		l_totemppro = 0
	
		do while not l_rs.eof
		
			if l_rs("empdes") <> l_empini and l_rs("prodes") <> l_proini then
		%>
					<tr>
						<td align="left"colspan="2"><%= l_empini %></td>					
						<td align="left" colspan="2"><%= l_proini %></td>
						<td align="center" colspan="2"><%= l_totemppro  %></td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>
						<td align="center">&nbsp;</td>		
						</tr>	
				
		<%		l_empini = l_rs("empdes")
				l_proini = l_rs("prodes")
				l_totemppro = 0
			end if
			l_totemppro = l_totemppro + clng(Neto (l_rs("movnro")))
			l_rs.movenext
		loop
		%>
			<tr>
				<td align="left"colspan="2"><%= l_empini %></td>					
				<td align="left" colspan="2"><%= l_proini %></td>
				<td align="center" colspan="2"><%= l_totemppro  %></td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>
				<td align="center">&nbsp;</td>		
			</tr>	
		<%	
		
		
		
		l_rs.close
		%>
		
	</td></tr>		
	<%
end sub 'totalesEmpresa

' Imprime el encabezado del Vessel
sub encab_Vessel()
	%>
		<tr>
			<td align="center" colspan="<%= l_cantcols%>" style="border-left-style: solid; border-left-width: 1px; border-left-color: Silver; border-top-color: Silver; border-top-style: solid; border-top-width: 1px; border-bottom-color: Silver; border-bottom-style: solid; border-bottom-width: 1px; border-right-color: Silver; border-right-style: solid; border-right-width: 1px;">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
			       	<td width="10%" align="right"><b>Vessel:</b></td>				
					<td align="left" width="30%"><%= l_rs("vesdesabr")%></td>
			       	<td align="right" nowrap width="10%"><b>Trip:</b></td>				
			       	<td align="left" nowrap width="20%"><%= l_rs("vestri") %></td>									
			       	<td align="right" nowrap width="10%"><b>Eta:</b></td>				
			       	<td align="left" nowrap width="20%"><%= l_rs("veseta") %></td>														
				</tr>
				<tr>
			       	<td width="10%" align="right" nowrap><b>Vessel Volume:</b></td>				
					<td align="left"><%= cdbl(l_rs("vesquantity")) / 1000 %></td>
			       	<td align="right" nowrap><b>Load Area:</b></td>				
			       	<td align="left" nowrap colspan="3"><%= l_rs("loaaredes") %></td>														
				</tr>
				<tr>
			       	<td colspan="2">&nbsp;</td>				
			       	<td align="right" nowrap><b>Disch Area:<b></td>				
			       	<td align="left" nowrap colspan="3"><%= l_rs("disaredes") %></td>														
				</tr>
			</table>
			</td>				
		</tr>		
	<%
end sub 'encabezado' Imprime el encabezado de cada pagina


' Imprime el encabezado de la Commodity
sub encab_Commodity()
	%>
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">&nbsp;</td>
		</tr>	
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
			       	<td width="10%" align="right"><b>Commodity:</b></td>				
					<td align="left" width="30%"><%= l_rs("quadesabr")%></td>
			       	<td align="right" nowrap width="10%"><b>Freight:</b></td>				
			       	<td align="left" nowrap width="20%"><%= l_rs("vesprototfre") %></td>									
			       	<td align="right" nowrap width="10%"><b>Product Volume:</b></td>				
			       	<td align="left" nowrap width="20%"><%= cdbl(l_rs("vesproqua")) / 1000 %></td>														
				</tr>
			</table>
			</td>				
		</tr>	
	<%
end sub

' Imprime el encabezado del Port
sub encab_Port()
	%>
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">&nbsp;</td>
		</tr>	
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
			       	<td width="10%" align="right" nowrap><b>Port:</b></td>				
					<td align="left"><%= l_rs("pordes") %></td>
			       	<td align="right" nowrap><b>Port Volume:</b></td>				
			       	<td align="left" nowrap colspan="3"><%= cdbl(l_rs("vesproporqua")) / 1000 %></td>														
				</tr>
			</table>
			</td>				
		</tr>	
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">&nbsp;</td>
		</tr>				
		<tr>
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Ctr Date</b></td>				
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Contract</b></th>				
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Client</b></th>				
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Ctr Vol</b></th>									
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>S/A</b></th>					
	       	<th style="font-size: 07pt;"  align="center" nowrap width="10%"><b>Shipmt</b></th>	
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Arrivalt</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Terms</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>F/P</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Price</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Premium</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Fob Parity</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Tendered</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>Loaded</b></th>					
	       	<th style="font-size: 07pt;" align="center" nowrap width="10%"><b>B/l/Date</b></th>								
		</tr>				
	<%
end sub

' Imprime el encabezado de la Commodity
sub mostrar_Datos()
	%>
<tr>
      	<td style="font-size: 07pt;" align="center" nowrap width="10%" ><%= l_rs("confec") %></td>				
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("ctrnum") %></td>				
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("clidesabr") %></td>				
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= cdbl(l_rs("conquantity")) / 1000 %></td>									
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("conshiarr") %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("conshiini") %></td>	
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("conarrini") %></td>					
      	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("terdes") %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("conprefla") %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= FormatNumber(l_rs("conflapri"),2) %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= FormatNumber(l_rs("conpre"),2) %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("confobpar") %></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><%= cdbl(l_rs("convolnom")) / 1000 %></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><%= cdbl(l_rs("convolloa")) / 1000 %></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%"><%= l_rs("nomsalblfec") %></td>					
</tr>		
	<%
		l_totpornom = l_totpornom + cdbl(l_rs("convolnom")) / 1000
		l_totporloa = l_totporloa + cdbl(l_rs("convolloa")) / 1000
		
		l_totcomnom = l_totcomnom + cdbl(l_rs("convolnom")) / 1000
		l_totcomloa = l_totcomloa + cdbl(l_rs("convolloa")) / 1000
	
		l_totvesnom = l_totvesnom + cdbl(l_rs("convolnom")) / 1000
		l_totvesloa = l_totvesloa + cdbl(l_rs("convolloa")) / 1000	
	
end sub 'encabezado' Imprime el encabezado de cada pagina


' Imprime los totales de Port
sub mostrar_Totales_Port()
	%>
<tr>
      	<td style="font-size: 07pt;" align="center" nowrap width="10%" colspan="10" >&nbsp;</td>	
      	<td style="font-size: 07pt;" align="right" nowrap width="10%" colspan="2"><b>Total Port:</b></td>						
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totpornom %></b></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totporloa %></b></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%">&nbsp;</td>					
</tr>		
	<%
		l_totpornom = 0
		l_totporloa = 0
	
end sub 'Imprime los totales de Port

' Imprime los totales de Commodity
sub mostrar_Totales_Commodity()
	%>
<tr>
      	<td style="font-size: 07pt;" align="center" nowrap width="10%" colspan="10" >&nbsp;</td>	
      	<td style="font-size: 07pt;" align="right" nowrap width="10%" colspan="2" ><b>Total Commodity:</b></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totcomnom %></b></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totcomloa %></b></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%">&nbsp;</td>					
</tr>		
	<%
		l_totcomnom = 0
		l_totcomloa = 0
		
end sub 'Imprime los totales de Commodity

' Imprime los totales de Vessel
sub mostrar_Totales_Vessel()
	%>
<tr>
      	<td style="font-size: 07pt;" align="center" nowrap width="10%" colspan="10" >&nbsp;</td>				
      	<td style="font-size: 07pt;" align="right" nowrap width="10%" colspan="2"><b>Total Vessel:</b></td>		
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totvesnom %></b></td>					
       	<td style="font-size: 07pt;" align="right" nowrap width="10%"><b><%= l_totvesloa %></b></td>					
       	<td style="font-size: 07pt;" align="center" nowrap width="10%">&nbsp;</td>					
</tr>		
	<%
		l_totvesnom = 0
		l_totvesloa = 0
		
end sub 'Imprime los totales de Vessel


sub encabezado(titulo)
	%>
		<table>
		<tr>
			<td align="center" colspan="<%= l_cantcols%>">
			<table cellpadding="0" cellspacing="0">
				<tr>
			       	<td width="20%">&nbsp;</td>				
					<td style="font-size: 12pt;" align="center" width="80%"><b><%= titulo%></b></td>
			       	<td align="right" valign="top"  nowrap width="20%">Página: <%= l_nropagina%></td>				
				</tr>
			</table>
			</td>				
		</tr>		
	<%
end sub


sub mostrar_datos1

	%>
		<tr>
			<td align="left" colspan="<%= l_cantcols %>"><b>Company:</b>&nbsp;<%= l_rs("comdesabr")%></td>
	    </tr>		
		<tr>
			<td align="left" colspan="<%= l_cantcols %>"><b>Port:</b>&nbsp;<%= l_rs("pordes")%></td>
	    </tr>				
		<tr>
			<td align="left" colspan="<%= l_cantcols %>">&nbsp;&nbsp;&nbsp;&nbsp;<% if l_rs("conpursal") = "P" then response.write "<b>Purchases</b>" else response.write "<b>Sales</b>" end if %></td>
	    </tr>						
		<tr>
			<th nowrap>Date</th>
			<th nowrap>Contract</th>
			<th nowrap>Broker</th>
			<th nowrap>Brik_nmbr</th>
			<th nowrap>Client</th>
			<th nowrap>Quantity</th>		
			<th nowrap>Tol</th>		
			<th nowrap>Shipmt</th>		
			<th nowrap>Price</th>				
			<th nowrap>Balance</th>		
			<th nowrap>Nomination Ref.</th>
			<th nowrap>Date</th>
			<th nowrap>Vessel</th>		
			<th nowrap>Eta</th>		
			<th nowrap>Port</th>
			<th nowrap>Nominated</th>
			<th nowrap>Loaded</th>					
			<th nowrap>B/l</th>
			<th nowrap>Remarks</th>								
		</tr>		
		<!-- Imprimo la LINEA 1 -->
		<tr>
			<td align="center" width="10%" nowrap><%= l_rs("confec")      %></td>
			<td align="center" width="10%" nowrap><%= l_rs("ctrnum")      %></td>
			<td align="left"   width="10%" nowrap><%= l_rs("brodes")      %></td>
			<td align="center" width="10%" nowrap><%= l_rs("bronum")      %></td>
			<td align="center" width="10%" nowrap><%= l_rs("clidesabr")   %></td>
			<td align="center" width="10%" nowrap><%= l_rs("conquantity") %></td>		
			<td align="center" width="10%" nowrap><%= l_rs("contol")      %></td>
			<td align="center" width="10%" nowrap><%= l_rs("conshiini")   %></td>				
			<td align="center" width="10%" nowrap><%'= l_KilosProc %></td>	 <!-- Neto de Origen -->			
			<td align="center" width="10%" nowrap><%'= l_Difer %></td>							
			<td align="center" width="10%" nowrap><%'= ValorObservado (l_rs("movnro"),1) %></td> <!-- % Hum -->
			<td align="center" width="10%" nowrap>&nbsp;</td> <!-- Gr. -->
			<td align="left" width="10%" nowrap><%' if l_cieter = 0 then response.write l_rs("concod") else response.write "" end if%></td>
	    </tr>	

	<%	
	
end sub 'mostrar datos

'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_quatypnro	  = request.querystring("qquatypnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

' Inicializo las variables totalizadoras

l_totpornom = 0
l_totporloa = 0
		
l_totcomnom = 0
l_totcomloa = 0
	
l_totvesnom = 0
l_totvesloa = 0
	
'response.write "l_fecini= " & l_fecini & "<br>"
'response.write "l_fecfin= " & l_fecfin & "<br>"
'response.write "l_quanro= " & l_quanro & "<br>"
'response.write "l_counro= " & l_counro & "<br>"
'response.write "l_pornro= " & l_pornro & "<br>"
'response.write "l_comnro= " & l_comnro & "<br>"

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

	l_sql = " select *, for_vessel.vesdesabr, loadarea.aredes loaaredes, discarea.aredes disaredes "
	l_sql = l_sql & " from for_vessel "
	l_sql = l_sql & " inner join for_vesselproduct on for_vesselproduct.vesnro = for_vessel.vesnro "
	l_sql = l_sql & " inner join for_vesselproductport on for_vesselproductport.vespronro = for_vesselproduct.vespronro "	
    l_sql = l_sql & " inner join for_quality on for_quality.quanro    = for_vesselproduct.quanro "
    l_sql = l_sql & " inner join for_area loadarea on loadarea.arenro    = for_vessel.areloanro "
    l_sql = l_sql & " inner join for_area discarea on discarea.arenro    = for_vessel.aredisnro "
    l_sql = l_sql & " inner join for_port on for_port.pornro = for_vesselproductport.pornro "

    l_sql = l_sql & " inner join for_nominationsale  on for_nominationsale.charter = for_vesselproductport.charter "
    l_sql = l_sql & " inner join for_contract on for_contract.connro = for_nominationsale.connro "
    l_sql = l_sql & " inner join for_client on for_client.clinro = for_contract.clinro "
    l_sql = l_sql & " inner join for_term on for_term.ternro = for_contract.ternro "

    l_sql = l_sql & " WHERE for_vesselproductport.vesproporare = 'D' "
	
	' Condiciones para contemplar los filtros ingresados
	
	' Fecha desde - Fecha Hasta
	if l_fecini <> "" and l_fecfin <> "" then 	
		l_sql = l_sql & " AND ( for_vessel.veseta >= " & cambiafecha(l_fecini,"YMD",true) 
		l_sql = l_sql & "   AND for_vessel.veseta <= " & cambiafecha(l_fecfin,"YMD",true) & " ) "
	end if
	
	' Commodity
	if l_quatypnro <> 0 then
		l_sql = l_sql & " AND ( for_quality.quatypnro = " & l_quatypnro & " ) "
	end if
	
	' Ordeno la consulta
	l_sql = l_sql & " order by for_vessel.vesdesabr, for_quality.quadesabr, for_contract.confec "

	'response.write l_sql
	
	'response.end
	
	Dim l_vesdesabr
	Dim l_quadesabr
	
	rsOpen l_rs, cn, l_sql, 0 
	if not l_rs.eof then 
	
		l_vesdesabr = l_rs("vesdesabr")
		l_quadesabr = l_rs("quadesabr")
		l_pordes    = l_rs("pordes")

		encabezado "VESSELS REPORT"
		encab_Vessel
		encab_Commodity
		encab_Port
		
		do while not l_rs.eof
		
			if l_vesdesabr <> l_rs("vesdesabr") then
			
				' Imprimo los Totales
				mostrar_Totales_Port
				mostrar_Totales_Commodity
				mostrar_Totales_Vessel
				
                response.write "</table><p style='page-break-before:always'></p><table>"
				l_nrolinea = 1
				
				l_vesdesabr = l_rs("vesdesabr")
				l_quadesabr = l_rs("quadesabr")
				l_pordes    = l_rs("pordes")
				
				encab_Vessel
				encab_Commodity
				encab_Port
			else
				if l_quadesabr <> l_rs("quadesabr") then
				
					l_quadesabr = l_rs("quadesabr")
					l_pordes    = l_rs("pordes")
					' Imprimo los totales de Port y Commodity
					mostrar_Totales_Port
					mostrar_Totales_Commodity
					' Imprimo los Encabezados
					encab_Commodity
					encab_Port
					
				else
					if l_pordes <> l_rs("pordes") then
						l_pordes    = l_rs("pordes")
						' Imprimo el total por Port
						mostrar_Totales_Port
						' Imprimo el encabezado del Port
						encab_Port
					end if
				
				end if
				
			end if
			mostrar_Datos
		
		'response.end

'			if l_nomcomdesabr <> l_rs("comdesabr") then
'				l_nomcomdesabr = l_rs("comdesabr")
'                response.write "</table><p style='page-break-before:always'></p>"
'				encabezado "Fob Contract Nominations"
'				l_nrolinea = 1
'			end if

'			mostrar_datos
			'l_total = 	l_total + 1
			l_rs.MoveNext			
			'l_nrolinea = l_nrolinea + 1	
		
			'if l_nrolinea > l_Max_Lineas_X_Pag then 
				'l_corte = true		
				'l_encabezado = true
				'l_nropagina	= l_nropagina + 1
			'else 
				'l_encabezado = false
			'end if
		
		loop 'end loop l_rs	
		
		' Imprimo los Totales del ultimo Vessel
		mostrar_Totales_Port
		mostrar_Totales_Commodity
		mostrar_Totales_Vessel
		
	
	else
	%>
	 <table><tr><td>No Existen datos para el filtro seleccionado.</td></tr></table>

	<%	
	end if
	
l_rs.Close
set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</table>
</body>
</html>

