<% Option Explicit
if request.querystring("excel") then
	Response.AddHeader "Content-Disposition", "attachment;filename=Auditoria.xls" 
	Response.ContentType = "application/vnd.ms-excel"
end if
response.buffer = true
 %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fecha.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/a_texto.inc"-->
<!--#include virtual="/trivoliSwimming/shared/inc/fnticket.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo: rep_auditoria_seg_01.asp
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

dim l_acnro
dim l_logusr
dim l_fecini
dim l_fecfin
dim l_campos
dim l_cont

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
	<tr>
		<th nowrap width="10%" >Fecha</th>
		<th nowrap width="5%" >Hora</th>
		<th nowrap width="20%" >Usuario</th>
		<th nowrap width="20%" >Acción</th>
		<th nowrap width="20%" >Valor Actual</th>
		<th nowrap width="20%" >Valor Anterior</th>		
	</tr>	

<%
end sub 'encabezado

sub mostrar_datos
	l_cont = l_cont + 1 ' Variable usada para realizar el Flush
	%>		
	<tr>
		<td align="center" nowrap valign="top"><%= l_rs("aud_fec") %></td>
		<td align="center" valign="top" ><%=l_rs("aud_hor") %></td>		
		<td align="left"   nowrap valign="top"><%= l_rs("iduser") %></td>
		<td align="left" nowrap valign="top"><%= l_rs("aud_des") %></td>
		<% if l_rs("aud_actual") = "-1" and l_rs("aud_ant")= "0" then %>
			<td align="center" nowrap valign="top">Si</td>		
			<td align="center" nowrap valign="top">No</td>
		<%else%>
			<% if l_rs("aud_actual") = "0" and l_rs("aud_ant")= "-1" then %>
				<td align="center" nowrap valign="top">No</td>		
				<td align="center" nowrap valign="top">Si</td>
			<%else%>
				<td align="center" nowrap valign="top"><%= l_rs("aud_actual") %></td>		
				<td align="center" nowrap valign="top"><%= l_rs("aud_ant")%></td>
			<%end if%>
		<%end if%>
    </tr>		
	<%
	if l_cont > 1000 then
		response.flush
		l_cont = 0
	end if

end sub 'mostrar datos


'Obtengo los parametros
l_acnro		  = request.querystring("qacnro")
l_logusr 	  = request.querystring("qlogusr")
l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_campos 	  = request.querystring("qcampos")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>

<% If not request.querystring("excel") then %>
	<link href="/trivoliSwimming/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<% End If %>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%

l_encabezado = true
l_corte = false

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

	l_sql = " SELECT * FROM auditoria "
	'l_sql = l_sql & " INNER JOIN aud_campo ON aud_campo.aud_campnro = rep_auditoria.aud_campnro "
	
	l_sql = l_sql & "  WHERE auditoria.aud_fec >= " & cambiafecha(l_fecini,"YMD",true) 
	l_sql = l_sql & "    AND auditoria.aud_fec <= " & cambiafecha(l_fecfin,"YMD",true)	
	
	if l_acnro <> "0" then 
		'---------------------------------------------------------
		' Filtro por las Acciones
		'---------------------------------------------------------
		l_sql = l_sql & " AND acnro = " & l_acnro
	end if
	
	'--------------------------------
	' Usuario
	'--------------------------------
	if l_logusr <> "0" then
		l_sql = l_sql & " AND ( auditoria.iduser = '" & l_logusr & "' ) "
	end if
	
	'--------------------------------
	' Campos
	'--------------------------------	
	if l_campos <> "" then
		l_sql = l_sql & " AND auditoria.aud_campnro IN (" & l_campos & " ) "
	end if	
	
	l_sql = l_sql & " ORDER BY auditoria.aud_fec, auditoria.aud_hor"
	rsOpen l_rs, cn, l_sql, 0 
	
if not l_rs.eof then

	l_nropagina = 1
	encabezado "Reporte de Auditoría"
	l_nrolinea = l_nrolinea + 1
	l_cont = 0
	do while not l_rs.eof

		' Controla que haya lugar para imprimir el encabezado de la camara y el producto en esta pagina.
		if not request.querystring("excel") and (l_nrolinea + 5 > l_Max_Lineas_X_Pag) then 
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina	= l_nropagina + 1
			encabezado ""
			l_nrolinea = 2
		end if

		mostrar_datos
		l_nrolinea = l_nrolinea + 1
		' Controla que haya lugar para imprimir el encabezado de la camara y el producto en esta pagina.
		if not request.querystring("excel") and l_nrolinea > l_Max_Lineas_X_Pag then 
			response.write "</table><p style='page-break-before:always'></p>"
			l_nropagina	= l_nropagina + 1
			encabezado ""
			l_nrolinea = 2
		end if
		
		'l_total = 	l_total + 1
		l_rs.MoveNext			
	
	loop 'end loop l_rs	
	'totales
	if not request.querystring("excel") and l_nrolinea > l_Max_Lineas_X_Pag then 
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
			No Existen Auditorías para el filtro seleccionado.
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

