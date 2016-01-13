<% Option Explicit
response.buffer = true
 %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--#include virtual="/turnos/shared/inc/a_texto.inc"-->
<!--#include virtual="/turnos/shared/inc/fnticket.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_log_seg_01.asp
Autor: Raul Chinestra
Creacion: 29/06/2006
Descripcion: Reporte de Log
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 50
Const l_cantcols = 7

Dim l_rs
Dim l_rs2

Dim l_sql

Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

dim l_tiplognro
dim l_logusr
dim l_fecini
dim l_fecfin
Dim l_carpornum
Dim l_movcod
Dim l_logdetest
Dim l_cont
Dim l_valorlog
Dim l_det
Dim l_ing
Dim l_rec

' Imprime los Totales

sub totales()
%>
<tr>
	<td align="center"><b>Total</b></td>
	<td align="center">&nbsp;</td>
	<td align="center">&nbsp;</td>		
	<td align="center"><b><%= l_nroope %></b></td>
</tr>	
</td></tr>		
<%
end sub 'totales

' Imprime el encabezado de la Cámara
sub encabcamara()
' Imprime 3 lineas
%>
<tr><td colspan="<%= l_cantcols%>">
	<table>
	<tr>
		<td align="center" colspan="<%= l_cantcols%>">
		<table>
			<tr>
		       	<td align="right"><b>Cámara:</b></td>				
				<td align="left" width="80%">
					<b><%= l_rs("camdes")%></b>
				</td>
		       	<td align="right" valign="top"  nowrap width="20%"> 
					<b>Fecha: <%= date %></b>
				</td>								
			</tr>					
			<tr><td colspan="3">&nbsp;
			</td></tr>			
			<tr><td colspan="3">Los siguientes Camiones son los que se envían a vuestra Cámara para su análisis:
			</td>
			</tr>
		</table>
		</td>				
	</tr>
	</table>
</td></tr>		
<%
end sub 'encabezado



' Imprime el encabezado de cada pagina
sub encabezado(titulo)
%>
	<table>
	<tr>
		<td align="center" colspan="<%= l_cantcols%>">
		<table>
			<tr>
		       	<td width="35%"  align="left"> <%= Empresa %>
					&nbsp;
				</td>				
				<td align="center" width="80%">
					<b><%= titulo%></b><br>
					<%= l_fecini %> -  <%= l_fecfin %>
				</td>
		       	<td align="right" valign="top"  nowrap width="20%"> 
					P&aacute;gina: <%= l_nropagina%>
				</td>				
			</tr>
		</table>
		</td>				
	</tr>		
	<% 	if l_det = "true" then 
		'--------------------------------------------------
		' Imprimir en forma Detallada el log
		'--------------------------------------------------
	%>
		<tr>
			<th nowrap width="10%" >Fecha</th>
			<th nowrap width="5%" >Hora</th>
			<th nowrap width="20%" >Usuario</th>
			<th nowrap width="20%" >Tipo Log</th>
			<th nowrap width="5%" >I = Ingresado <br>
								   R = Rechazado</th>				
			<th nowrap width="20%" >Carta de Porte</th>
			<th nowrap width="20%" >Movimiento</th>		
		</tr>
	<% Else
		'--------------------------------------------------
		' Imprimir en forma Normal el log
		'--------------------------------------------------
	%>
		<tr>
			<th nowrap width="10%" >Fecha</th>
			<th nowrap width="5%" >Hora</th>
			<th nowrap width="20%" >Usuario</th>
			<th nowrap width="20%" >Tipo Log</th>
			<th nowrap width="20%" >Ingresados</th>
			<th nowrap width="20%" >Rechazados</th>
			<th nowrap width="20%" >Total</th>		
		</tr>	
	<% End If %>
<%

end sub 'encabezado

sub mostrar_datos
	l_cont = l_cont + 1 ' Variable usada para realizar el Flush
	
	if l_det = "true" then
	
		'--------------------------------------------------
		' Imprimir en forma Detallada el log
		'--------------------------------------------------
	
		if l_valorlog <> l_rs("lognro") then
			l_valorlog = l_rs("lognro")
	%>
			<tr>
				<td align="center" nowrap valign="top"><%= l_rs("logfec") %></td>
				<td align="center" valign="top" ><%= mid(l_rs("loghor"),1,2) %>:<%= mid(l_rs("loghor"),3,2) %></td>		
				<td align="left"   nowrap valign="top"><%= l_rs("logusr") %></td>
				<td align="center" nowrap valign="top"><%= l_rs("tiplogdes") %></td>
				<td align="center" nowrap valign="top"><%= l_rs("logdetest") %></td>		
				<td align="center" nowrap valign="top"><%= l_rs("carpornum")%></td>
				<td align="center" nowrap valign="top"><%= l_rs("movcod")%></td>
		    </tr>	
		<%  else %>
			<tr>
				<td align="center" nowrap valign="top"><%'=  %></td>
				<td align="center" valign="top" ><%'= mid(l_rs("loghor"),1,2) %><%'= mid(l_rs("loghor"),3,2) %></td>		
				<td align="left"   nowrap valign="top"><%'= l_rs("logusr") %></td>
				<td align="center" nowrap valign="top"><%'= l_rs("tiplogdes") %></td>
				<td align="center" nowrap valign="top"><%= l_rs("logdetest") %></td>		
				<td align="center" nowrap valign="top"><%= l_rs("carpornum")%></td>
				<td align="center" nowrap valign="top"><%= l_rs("movcod")%></td>
		    </tr>	
		<% End If %>			
		
	<% else
		'--------------------------------------------------
		' Imprimir en forma Normal el log
		'--------------------------------------------------
	%>		
		<tr>
			<td align="center" nowrap valign="top"><%= l_rs("logfec") %></td>
			<td align="center" valign="top" ><%= mid(l_rs("loghor"),1,2) %>:<%= mid(l_rs("loghor"),3,2) %></td>		
			<td align="left"   nowrap valign="top"><%= l_rs("logusr") %></td>
			<td align="center" nowrap valign="top"><%= l_rs("tiplogdes") %></td>
			<td align="center" nowrap valign="top"><%= l_rs("loging") %></td>		
			<td align="center" nowrap valign="top"><%= l_rs("logrec")%></td>
			<% 	if isnull(l_rs("loging")) then 
					l_ing = 0 
				else 
					l_ing = clng(l_rs("loging")) 
				end if
				if isnull(l_rs("logrec")) then 
					l_rec = 0 
				else 
					l_rec = clng(l_rs("logrec")) 
				end if%>
			<td align="center" nowrap valign="top"><%= l_ing + l_rec %></td>			

	    </tr>		

	<% end if 
	
	if l_cont > 1000 then
		response.flush
		l_cont = 0
	end if

end sub 'mostrar datos


'Obtengo los parametros
l_tiplognro	  = request.querystring("qtiplognro")
l_logusr 	  = request.querystring("qlogusr")
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_carpornum	  = request.querystring("qcarpornum")
l_movcod	  = request.querystring("qmovcod")
l_logdetest	  = request.querystring("qlogdetest")
l_det         = request.querystring("qdet")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%

l_encabezado = true
l_corte = false

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

	l_sql = " SELECT *"
	l_sql = l_sql & " FROM tkt_log "
	l_sql = l_sql & " INNER JOIN tkt_tipolog on tkt_tipolog.tiplognro = tkt_log.tiplognro  "
	l_sql = l_sql & " INNER JOIN user_per ON user_per.iduser = tkt_log.logusr "
	
	if l_det = "true" then 
		'---------------------------------------------------------
		' Busco el detalle de la Cabecera del Log
		'---------------------------------------------------------
		l_sql = l_sql & " INNER JOIN tkt_logdet ON tkt_logdet.lognro = tkt_log.lognro "	
	end if
	
	l_sql = l_sql & "  WHERE tkt_log.logfec >= " & cambiafecha(l_fecini,"YMD",true) 
	l_sql = l_sql & "    AND tkt_log.logfec <= " & cambiafecha(l_fecfin,"YMD",true)
	
	'--------------------------------
	' Tipo de Log
	'--------------------------------
	if l_tiplognro <> "0" then
		l_sql = l_sql & " AND ( tkt_log.tiplognro = " & l_tiplognro & " ) "
	end if 
	
	'--------------------------------
	' Usuario
	'--------------------------------
	if l_logusr <> "0" then
		l_sql = l_sql & " AND ( tkt_log.logusr = '" & l_logusr & "' ) "
	end if
	
	if l_det = "true" then 
		'----------------------------------------------------------------------------------
		' En el caso de que el Usuario selecciono Detallado puedo filtrar por los campos 
		' Carta de Porte , Movimiento y Estado
		'----------------------------------------------------------------------------------
	
		'--------------------------------
		' Carta de Porte
		'--------------------------------
		if l_carpornum <> "" then
			l_sql = l_sql & " AND ( tkt_logdet.carpornum like '%" & l_carpornum & "%' ) "
		end if
	
		'--------------------------------
		' Movimiento
		'--------------------------------
		if l_movcod <> "" then
			l_sql = l_sql & " AND ( tkt_logdet.movcod like '%" & l_movcod & "%' ) "
		end if
		
		'--------------------------------
		' Estado
		'--------------------------------
		if l_logdetest <> "0" then
			l_sql = l_sql & " AND ( tkt_logdet.logdetest = '" & l_logdetest & "' ) "
		end if	
		
	end if
	
	if l_det = "true" then 
		' Ordenado por Lognro y Logdetest	
		l_sql = l_sql & " ORDER BY tkt_log.lognro, logdetest "
	else
		l_sql = l_sql & " ORDER BY tkt_log.lognro "
	end if
	rsOpen l_rs, cn, l_sql, 0 
	
if not l_rs.eof then

	l_nropagina = 1
	encabezado "Reporte de Log"
	l_nrolinea = l_nrolinea + 1
	l_cont = 0
	l_valorlog = 0
	do while not l_rs.eof

		' Controla que haya lugar para imprimir el encabezado de la camara y el producto en esta pagina.
		if l_nrolinea + 5 > l_Max_Lineas_X_Pag then 
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina	= l_nropagina + 1
			encabezado ""
			l_nrolinea = 2
		end if

		mostrar_datos
		l_nrolinea = l_nrolinea + 1
		' Controla que haya lugar para imprimir el encabezado de la camara y el producto en esta pagina.
		if l_nrolinea > l_Max_Lineas_X_Pag then 
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina	= l_nropagina + 1
			encabezado ""
			l_nrolinea = 2
		end if
		
		'l_total = 	l_total + 1
		l_rs.MoveNext			
	
	loop 'end loop l_rs	
	'totales
	if l_nrolinea > l_Max_Lineas_X_Pag then 
		response.write "</table><p style='page-break-before:always'></p>"
		l_nropagina	= l_nropagina + 1
		encabezado ""
		l_nrolinea = 2
	end if
else 
%>
	<table>
	<tr>
		<td align="left" colspan="<%= l_cantcols%>">
			No Existen Log para el filtro seleccionado.
		</td>
	</tr>	
	</table>
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

