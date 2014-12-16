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
Dim l_valor

Dim l_fila
Dim l_todos

Dim l_mostrarEncab
Dim l_buque
Dim l_tipoope
Dim l_destino

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
<title>Interfaces </title>
</head>
<script src="/serviciolocal/shared/js/fn_sel_multiple.js"></script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<%
  
  response.write "<table><tr><th>Buques</th></tr>"
  
  l_sql = "SELECT * "
  l_sql = l_sql & " FROM buq_mercaderia  "
  rsOpen l_rs, cn, l_sql, 0 
  do while not l_rs.eof
  
  	l_sql = " select year(buqfechas), sum(conton) "
	l_sql = l_sql & " from buq_contenido "
	l_sql = l_sql & " inner join buq_buque on buq_contenido.buqnro = buq_buque.buqnro "
	l_sql = l_sql & " where mernro = " & l_rs("mernro")
	l_sql = l_sql & " group by year(buqfechas) "
    rsOpen l_rs2, cn, l_sql, 0 
	do while not l_rs2.eof 

	   response.write "<tr><td>" & l_rs("mernro") & "</td>"
   	   response.write "    <td>" & l_rs2(0) & "</td>"
   	   response.write "    <td>" & l_rs2(1) & "</td></tr>"
	   
'		l_sql = "INSERT INTO buq_acumulado "
'		l_sql = l_sql & " (acuanio, acumes, acutip, acucod, acutot)"
'		l_sql = l_sql & " VALUES (" & l_rs2(0) & ",13,1" & l_rs("mernro") & "," & l_rs2(1) & ")"
'		l_cm.activeconnection = Cn
'		l_cm.CommandText = l_sql
'		cmExecute l_cm, l_sql, 0
	
		l_rs2.movenext
	loop
	l_rs2.close 
  
  	l_rs.movenext
  loop
  l_rs.close
      
	

'	
'    	l_valor = codigogenerado()
'		l_buque = l_arr(0)
'		
'	end if
'
'  Loop
  
'  cn.CommitTrans 
  
'  response.write "<tr><th>La Cantidad Importada de Buques es = " & l_fila & "</th></tr>"

  Set l_cm = Nothing
  
%>

</body>
</html>
