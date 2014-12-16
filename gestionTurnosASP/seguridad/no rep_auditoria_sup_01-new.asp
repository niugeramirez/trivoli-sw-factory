<% Option Explicit %>
<!--#include virtual="/rhprox2/shared/db/conn_db.inc"-->
<!--#include virtual="/rhprox2/shared/inc/fecha.inc"-->
<%
'--------------------------------------------------------------------------------------
'Archivo        : rep_auditoria_sup_01.asp
'Descripción    : Reporte - Auditoria
'Autor          : CCRossi
'Fecha Creacion : 28-04-2004
'Modificado     
'--------------------------------------------------------------------------------------
Const l_Max_Lineas_X_Pag = 47
const l_nro_col = 7

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

Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
Set l_rs3 = Server.CreateObject("ADODB.RecordSet")
Set l_rs4 = Server.CreateObject("ADODB.RecordSet")

'-----------------------------------------------------------------------------------------------------------
' Imprime el encabezado de cada pagina
'-----------------------------------------------------------------------------------------------------------
sub encabezado%>
	<table>
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
		<td nowrap align="center"><b>Fecha</b></td>
		<td nowrap align="center"><b>Hora</b></td>
		<td nowrap align="center"><b>Usuario</b></td>
		<td nowrap align="center"><b>Acci&oacute;n</b></td>
		<td nowrap align="center"><b>Descripci&oacute;n</b></td>
		<td nowrap align="center"><b>Val.Actual</b></td>
		<td nowrap align="center"><b>Val.Anterior</b></td>
	</tr>

<%
end sub 'encabezado


'-----------------------------------------------------------------------------------------------------------
'										   B O D Y 
'-----------------------------------------------------------------------------------------------------------
l_titulofiltro = request("tfiltro")
l_filtro 	= request("filtro")

l_acciones  = request("acciones")
l_usuarios  = request("usuarios")
l_confaud  = request("confaud")
l_empresas  = request("empresas")
l_acnro		= request("acnro")
l_iduser 	= request("iduser")
l_caudnro 	= request("caudnro")
l_empnro 	= request("empnro")

l_fechadesde= request("fechadesde")
l_fechahasta= request("fechahasta")
l_orden		= request("orden")

'Response.Write(l_filtro) & "<br>"
'Response.Write(l_fechadesde) & "<br>"
'Response.Write(l_fechahasta) & "<br>"
'Response.Write("acciones "& l_acciones) & "<br>"
'Response.Write("acnro "&l_acnro) & "<br>"
'Response.Write("usuarios "&l_usuarios) & "<br>"
'Response.Write("iduser "&l_iduser) & "<br>"
'Response.Write("confuaud "&l_confaud) & "<br>"
'Response.Write("caudnro "&l_caudnro) & "<br>"
'Response.Write("orden "&l_orden) & "<br>"
'Response.Write("empresas "&l_empresas) & "<br>"
'Response.Write("estrnro "&l_estrnro) & "<br>"
'Response.End

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/rhprox2/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
l_sql = "SELECT  auditoria.aud_fec, auditoria.aud_hor, auditoria.iduser, auditoria.aud_des, " 
l_sql = l_sql & " accion.acnro,accion.acdesc, "
l_sql = l_sql & " auditoria.aud_actual , auditoria.aud_ant,"
l_sql = l_sql & " estructura.estrdabr "
l_sql = l_sql & " FROM auditoria "
l_sql = l_sql & " INNER JOIN accion ON accion.acnro  = auditoria.acnro "  
if trim(l_acnro) <>"" and l_acnro <> "0" then
	l_sql = l_sql & " AND accion.acnro =" & l_acnro
end if
l_sql = l_sql & " INNER JOIN confaud ON confaud.caudnro  = auditoria.caudnro "  
if trim(l_caudnro) <>"" and l_caudnro <> "0" then
	l_sql = l_sql & " AND confaud.caudnro =" & l_caudnro
end if
l_sql = l_sql & " INNER JOIN user_per ON user_per.iduser  = auditoria.iduser "  
if trim(l_iduser) <>"" and l_iduser <> "0" then
	l_sql = l_sql & " AND user_per.iduser ='" & l_iduser & "'"
end if
l_sql = l_sql & " INNER JOIN empresa ON empresa.empnro  = auditoria.empnro "  
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro  = empresa.estrnro "  
l_sql = l_sql & " WHERE (0=0)"
l_sql = l_sql & " AND " & l_filtro 
if trim(l_empnro) <>"" and l_empnro <> "0" then
	l_sql = l_sql & " AND auditoria.empnro =" & l_empnro
end if
if trim(l_fechadesde) <> "" then
l_sql = l_sql & " AND auditoria.aud_fec >= " & cambiafecha(l_fechadesde,"","")
end if
if trim(l_fechahasta) <> "" then
l_sql = l_sql & " AND auditoria.aud_fec <= " & cambiafecha(l_fechahasta,"","")
end if
l_sql = l_sql & " ORDER BY  estrdabr, "&l_orden
'Response.Write(l_sql)
'Response.End
rsOpen l_rs, cn, l_sql, 0 

l_nrolinea = 1
l_nropagina = 1
l_encabezado = true
l_corte = false
l_estrdabrant=""
do until l_rs.eof
	
	if l_encabezado then 
		if l_corte then
			response.write "</table><p style='page-break-before:always'></p>"
			l_nrolinea = 1
		end if 		
		
		encabezado 
		l_nrolinea = l_nrolinea+4
		
	end if	
	
	' mostrarDatos 
	
	if trim(l_estrdabrant) <> trim(l_rs("estrdabr")) then%>
		<tr> 
			<td><b>Empresa:</b></td>
			<td colspan=<%=l_nro_col%>><b><%=l_rs("estrdabr")%></b></td>
	   </tr>
		<% 
	    l_nrolinea	 = l_nrolinea + 1
		l_estrdabrant= l_rs("estrdabr")
	end if
	
	l_totalestr = l_totalestr + 1
	l_linea = "<tr>"
	l_linea = l_linea &   "<td nowrap align=""center"">" &l_rs("aud_fec")& "</td>"
	l_linea = l_linea &   "<td nowrap align=""center"">" &l_rs("aud_hor")&"</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("iduser")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("acdesc")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_des")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_actual")& "</td>"	
	l_linea = l_linea &   "<td nowrap align=""left"">" & l_rs("aud_ant")& "</td>"	
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
	<table>
	<tr>
		<td align="center" colspan="<%=l_nro_col%>"><b>No se encontraron datos</b></td>
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
