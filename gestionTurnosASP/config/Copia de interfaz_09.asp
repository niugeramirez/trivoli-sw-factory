<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : interfaz_09.asp
Descripcion    : 
Creador        : Raul Chinestra 
Fecha Creacion : 05/03/2008
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_rs2
Dim l_sql
Dim l_fs
Dim l_cm

Dim l_path
Dim l_modelo
Dim l_archivo
Dim l_separador
Dim l_encabezado
Dim l_arch
Dim l_arr
Dim l_str
Dim l_i
Dim l_mostrar
Dim l_storage
Dim l_seed
Dim l_maipro
Dim l_quanro

Dim l_fila
Dim l_todos

Dim l_mostrarEncab

Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
set l_cm = Server.CreateObject("ADODB.Command")

'l_modelo  = request("modelo")
'l_archivo = request("archivo")
'l_mostrar = request("mostrar")

'l_mostrar = (l_mostrar = "1")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/serviciolocal/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Interfaces de Precios Indices - buques</title>
</head>
<script src="/serviciolocal/shared/js/fn_sel_multiple.js"></script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%


'--------------------------------------------------------------------------------------------------------------------------
' IMPORTACION DE PRECIOS INDICES
'--------------------------------------------------------------------------------------------------------------------------

  l_path = "c:\loc.csv"

  Set l_fs = CreateObject("Scripting.FileSystemObject")
  Set l_arch = l_fs.OpenTextFile(l_path,1,false) 
  

  l_fila = 0
  l_separador = ";"
  
  response.write "<table><tr><th>Localidad</th></tr>"
  
  Do While l_arch.AtEndOfStream <> true
  
    l_str=l_arch.readline
	l_arr = split(l_str,l_separador)
	
	'Verifico que no este dado de ALTA
'	l_sql = "SELECT  promap "
'	l_sql = l_sql & " FROM for_product "
'	l_sql = l_sql & " WHERE promap = " & l_arr(1)
'	rsOpen l_rs, cn, l_sql, 0
'	if not l_rs.eof then

		l_sql = "INSERT INTO int_localidad "
		l_sql = l_sql & " (locdes, pronro)"
		l_sql = l_sql & " VALUES ('" & l_arr(0) & "'," & l_arr(1) & ")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		l_fila = l_fila + 1
		
'	else
'		response.write "<tr><td>El PRODUCTO No existe. " & l_arr(1) & "</td></tr>"
'	end if
	'l_rs.close

 	if l_fila>=25 then
		response.flush
	end if

  Loop
  response.write "<tr><th>La Cantidad Importada de LOCALIDADES es = " & l_fila & "</th></tr>"


  l_arch.close
  Set l_cm = Nothing
  response.end 
  
%>

<form name="datos" method="post" action="interfaz_03.asp">
<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="listatodos" value="<%= l_todos%>">
<input type="Hidden" name="modelo" value="<%= l_modelo%>">
<input type="Hidden" name="archivo" value="<%= l_archivo%>">
<input type="Hidden" name="param" value="">
</form>
</body>
</html>
