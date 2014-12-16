<% Option Explicit%>

<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 53
Const l_cantcols = 10
Const l_empresa = "Cámara Portuaria y Marítima <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; de Bahía Blanca"


Dim l_rs
Dim l_rs2
Dim l_sql

Dim l_encabezado
Dim l_corte 
Dim l_nropagina
Dim l_nrolinea

dim l_fecini
dim l_fecfin
dim l_expnro
dim l_desexportadora
dim l_buqdes
dim l_totton
dim l_primeravez
dim l_expdes
dim l_primeravezexp

'Obtengo los parametros
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_expnro      = request.querystring("expnro")

'response.write l_expnro
'response.end

sub encabezado_expbuq(titulo)
%>
	<table style="width:99%" cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td align="center" colspan="14">
				<table border="0" cellpadding="0" cellspacing="0" >
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
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Planilla</th>			
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Comenzó</th>
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Terminó</th>
	        <th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Mercadería</th>			
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Toneladas</th>		
			<th align="center" width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; border-right-color: #000000; border-right-style: solid; border-right-width: 1px; border-bottom-color: #000000; border-bottom-style: solid; border-bottom-width: 1px; ">Destino</th>																
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

<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
</head>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

' Busco la exportadora
'if l_expnro = 0 then
'	l_desexportadora = "Todas"
'else 
'
'	l_sql = " SELECT * "
'	l_sql = l_sql & " FROM buq_exportadora "
'	l_sql = l_sql & " where expnro = " & l_expnro
'	rsOpen l_rs, cn, l_sql, 0
'	if not l_rs.eof then
'		l_desexportadora = l_rs("expdes")
'	end if
'	l_rs.close
'end if


l_nropagina = 1
'l_nropagina = 1
'encabezado_expbuq("Listado de Buques - Exportadora: " & l_desexportadora)  
'l_nrolinea = 6

'l_encabezado = true
'l_corte = false
'l_total = 0

l_sql = " SELECT * "
l_sql = l_sql & " FROM buq_buque "
l_sql = l_sql & " inner join buq_contenido on buq_contenido.buqnro = buq_buque.buqnro "
l_sql = l_sql & " inner join buq_mercaderia on buq_mercaderia.mernro = buq_contenido.mernro "
l_sql = l_sql & " inner join buq_exportadora on buq_exportadora.expnro = buq_contenido.expnro "
l_sql = l_sql & " left join buq_destino on buq_destino.desnro = buq_contenido.desnro "
'l_sql = l_sql & " inner join buq_agencia on buq_agencia.agenro = buq_buque.agenro "

l_sql = l_sql & " WHERE  (buq_contenido.expnro = " & l_expnro & " OR " & l_expnro & " = 0 ) "

l_sql = l_sql & " AND  buq_buque.buqfechas >= " & cambiafecha(l_fecini,"YMD",true) 
l_sql = l_sql & " AND buq_buque.buqfechas <= " & cambiafecha(l_fecfin,"YMD",true)
l_sql = l_sql & " ORDER BY buq_contenido.expnro , buq_buque.buqfecdes "

rsOpen l_rs, cn, l_sql, 0

'response.write l_sql
'response.end

if l_rs.eof then
	'l_buqdes = ""
	%>
    <tr>	
		<td align="left" colspan="10">No existen Buques para el filtro seleccionado</td>
    </tr>
	<%
else

	'l_canbuq = 0
	
	l_totton = 0
	l_buqdes = ""
	l_primeravez = true
	l_primeravezexp = true
	l_expdes = ""
	do until l_rs.eof

		if l_nrolinea > l_Max_Lineas_X_Pag then
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina = l_nropagina + 1
			encabezado_expbuq("Listado de Buques - Exportadora: " & l_rs("expdes") )  
			l_nrolinea = 6
		end if
		
		if l_expdes <> l_rs("expdes") then
				if l_primeravezexp = true then 
					l_primeravezexp = false
				else
				%>
				<tr>						
					<td align="left" width="10%" nowrap >&nbsp;</td>									
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>							
					<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; "><b><%= l_totton %></b>&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>							
			    </tr>		
				<tr>						
					<td align="left" width="10%" nowrap >&nbsp;</td>									
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>							
					<td align="right"  width="10%" >&nbsp;</td>
					<td align="center" width="10%" >&nbsp;</td>							
			    </tr>
				<%
				response.write "</table><p style='page-break-before:always'></p>"
				l_nropagina = l_nropagina + 1
				end if 

				l_totton = 0
				l_buqdes = l_rs("buqdes")
				l_nrolinea = l_nrolinea + 1
				l_totton = l_totton + l_rs("conton")
	
				encabezado_expbuq("Listado de Buques - Exportadora: " & l_rs("expdes") )  
				l_nrolinea = 6
				l_expdes = l_rs("expdes")

				%>
				<tr>			   
					<td align="left" width="10%" nowrap ><%=l_rs("buqdes")%></td>			
					<td align="center" width="10%" ><%= l_rs("buqnro") %></td>
					<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
					<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>
					<td align="center" width="10%" ><%= l_rs("merdes") %></td>
					<td align="right"  width="10%" ><%= l_rs("conton") %>&nbsp;</td>
					<td align="center" width="10%" >&nbsp;<%= l_rs("desdes") %></td>
			    </tr>				
				<%
		else
		%>
			<% if l_buqdes <> l_rs("buqdes") then
					if l_primeravez = true then 
						l_primeravez = false
					else
						%>
						<tr>						
							<td align="left" width="10%" nowrap >&nbsp;</td>									
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>							
							<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; " ><b><%= l_totton %></b>&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>							
					    </tr>		
						<tr>						
							<td align="left" width="10%" nowrap >&nbsp;</td>									
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>							
							<td align="right"  width="10%" >&nbsp;</td>
							<td align="center" width="10%" >&nbsp;</td>							
					    </tr>						
						<%
						l_totton = 0
					end if
			   %>
			<tr>			   
				<td align="left" width="10%" nowrap ><%=l_rs("buqdes")%></td>			
			   <%
			    l_buqdes = l_rs("buqdes")
				'l_canbuq = l_canbuq + 1
			   else
			   %>
			<tr>			   
				<td align="left" width="10%"  nowrap >&nbsp;</td>			
			   <%
  			   end if
			 %>
			<td align="center" width="10%" ><%= l_rs("buqnro") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfecdes") %></td>
			<td align="center" width="10%" ><%= l_rs("buqfechas") %></td>
			<td align="center" width="10%" ><%= l_rs("merdes") %></td>
						
			<td align="right"  width="10%" ><%= l_rs("conton") %>&nbsp;</td>

			<td align="center" width="10%" >&nbsp;<%= l_rs("desdes") %></td>
	    </tr>
		<%
		l_nrolinea = l_nrolinea + 1
		l_totton = l_totton + l_rs("conton")
		
		end if
		
		l_rs.MoveNext
	loop
	%>
	<tr>						
		<td align="left" width="10%" nowrap >&nbsp;</td>									
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>							
		<td align="right"  width="10%" style="border-top-color: #000000; border-top-style: solid; border-top-width: 1px; "><b><%= l_totton %></b>&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>							
    </tr>		
	<tr>						
		<td align="left" width="10%" nowrap >&nbsp;</td>									
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>							
		<td align="right"  width="10%" >&nbsp;</td>
		<td align="center" width="10%" >&nbsp;</td>							
	 </tr>	
	<%

end if 

'response.write l_nrolinea 

l_rs.Close

l_nrolinea = l_nrolinea + 1
'response.write "</table><p style='page-break-before:always'></p>"
l_nropagina = l_nropagina + 1




set l_rs = Nothing
cn.Close
set cn = Nothing
%>
</body>
</html>

