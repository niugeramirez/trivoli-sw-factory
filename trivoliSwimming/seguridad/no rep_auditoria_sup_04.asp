<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--#include virtual="/turnos/shared/inc/fecha.inc"-->
<!--
Archivo        : rep_auditoria_sup_04.asp
Descripción    : Reporte - Auditoria - Salida html
Autor          : JMH
Fecha Creacion : 20/01/2005
Modificado     : 
	26/07/2005 - Scarpa D. - Se agrego la columan empleado
-->

<% 
on error goto 0

Const l_Max_Lineas_X_Pag = 45
const l_nro_col = 9

Dim l_rs
Dim l_rs2
Dim l_rs3
Dim l_rs4
Dim l_sql
Dim l_filtro

dim l_linea
Dim l_nrolinea
Dim l_nropagina

Dim l_encabezado
Dim l_corte 

Dim l_estrdabrant

Dim l_orden
dim l_caudnro
dim l_acnro
dim l_iduser
Dim l_fechadesde
Dim l_fechahasta
Dim l_empnro

dim l_usuarios
dim l_acciones
dim l_confaud
dim l_empresas

Dim l_totalestr
Dim l_titulofiltro
Dim l_bpronro

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
Set l_rs4 = Server.CreateObject("ADODB.RecordSet")

'OBTENGO EL PARAMETRO
l_bpronro = request("bpronro")

'------------------------------ FUNCIONES AUXILIARES ----------------------------------------

' Imprime el encabezado de cada pagina
sub encabezado%>
	<tr>
		<td align="center" colspan="<%=l_nro_col%>">
		<table>
			<tr>
		       	<td width="10%"> 
					&nbsp;
				</td>				
				<td align="center" width="80%">
					<b> AUDITORIA </b>
				</td>
		       	<td align="right" width="10%"> 
					P&aacute;gina: <%= l_nropagina%>
				</td>				
			</tr>
		</table>
		</td>				
	</tr>
	<tr>
		<td nowrap colspan="<%=l_nro_col%>" align="center">
			<%=l_titulofiltro%>
		</td>
	</tr>
	<tr>
		<th nowrap align="center"><b>Fecha</b></th>
		<th nowrap align="center"><b>Hora</b></th>
		<th nowrap align="center"><b>Usuario</b></th>
		<th nowrap align="center"><b>Acci&oacute;n</b></th>
		<th nowrap align="center"><b>Descripci&oacute;n</b></th>
		<th nowrap align="center"><b>Campo</b></th>
		<th nowrap align="center"><b>Val.Actual</b></th>
		<th nowrap align="center"><b>Val.Anterior</b></th>
		<th nowrap align="center"><b>Empleado</b></th>
	</tr>

<%
end sub 'encabezado

'---------------------------------- FIN FUNCIONES AUXILIARES ------------------------------

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/rhprox2/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
l_nrolinea = 1
l_nropagina = 0
l_encabezado = true
l_corte = false

l_sql = " SELECT * FROM rep_auditoria "
l_sql = l_sql & " INNER JOIN aud_campo ON aud_campo.aud_campnro = rep_auditoria.aud_campnro "
l_sql = l_sql & " WHERE bpronro= " & l_bpronro
l_sql = l_sql & " ORDER BY rep_auditoria.aud_fec, rep_auditoria.aud_hor"

rsOpen l_rs, cn, l_sql, 0 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_estrdabrant=""

response.write "<table style='border: 0px solid black;'>"

do until l_rs.eof
	
	if l_encabezado then 
		if l_corte then
            response.write "<tr><td nowrap style='page-break-before:always;background:white;' colspan='" & l_nro_col & "'><br></td></tr>"
			l_nrolinea = 1
		end if 		
		
		encabezado 
		l_nrolinea = l_nrolinea+4
		
	end if	
	
	' mostrarDatos 

	l_totalestr = l_totalestr + 1
	l_linea = "<tr>"
	l_linea = l_linea &   "<td nowrap align=""center"">" &l_rs("aud_fec")& "</td>"
	l_linea = l_linea &   "<td nowrap align=""center"">" &l_rs("aud_hor")&"</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_iduser")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("acc_desc")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_des")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_campdesabr")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_actual")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_ant")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("empleado")& "&nbsp;</td>"	
	l_linea = l_linea & "</tr>"
	
	response.write l_linea
	l_nrolinea  = l_nrolinea + 1
	
	
	if l_nrolinea > l_Max_Lineas_X_Pag then 
		l_corte = true
		l_encabezado = true
		l_nropagina	= l_nropagina + 1
	else 
		l_encabezado = false
	end if

	l_rs.MoveNext
loop

' Se imprime el total de registros
if l_totalestr = 0 then
	%>
	<tr>
		<td align="center" colspan="<%=l_nro_col%>"><b>No se encontraron datos</b></td>
	</tr>
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

