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


Dim l_dervul
Dim l_mednro
Dim l_legpar1
Dim l_legpar2
Dim l_legpar3
Dim l_aux
Dim l_fecing
Dim l_fecnac

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

function fsql_seqvalue(ByVal campo, ByVal tabla)
  Dim auxi
  Select Case l_base

        Case "2" 
			auxi = "select @@IDENTITY as " & campo & " "

  End Select
  fsql_seqvalue = auxi
end function

' ------------------------------------------------------------------------------------------------------------------
' codigogenerado() :
' ------------------------------------------------------------------------------------------------------------------
function codigogenerado()
	Dim l_rs
	Dim l_sql
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	l_sql = fsql_seqvalue("next_id","cap_evento")
	rsOpen l_rs, cn, l_sql, 0
	codigogenerado=l_rs("next_id")
	l_rs.Close
	Set l_rs = Nothing
end function 'codigogenerado()

'Al operar sobre varias tablas debo iniciar una transacción
 cn.BeginTrans
 


  l_path = "c:\_\PRUEBA_10.csv"

  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  Set l_fs = CreateObject("Scripting.FileSystemObject")
  Set l_arch = l_fs.OpenTextFile(l_path,1,false) 
  

  l_fila = 0
  l_separador = ";"
  
  response.write "<table><tr><th>Portadas</th></tr>"
  
  l_buque = ""
  
  Do While l_arch.AtEndOfStream <> true
  
    l_str=l_arch.readline
	l_arr = split(l_str,l_separador)

		l_aux = l_arr(0)
		l_legpar1 =  mid (l_aux,1,instr(l_arr(0), "/") -1 )
		l_aux = mid(l_aux,instr(l_aux, "/") + 1 , len(l_aux) )
		l_legpar2 =  mid(l_aux,1,instr(l_aux, "/") -1 )
		l_aux = mid(l_aux,instr(l_aux, "/") + 1 , len(l_aux) )
		l_legpar3 =  mid(l_aux,1,len(l_aux) )
		
		'response.write "-1-" & l_legpar1 & "<br>"
		'response.write "-2-" & l_legpar2 & "<br>"
		'response.write "-3-" & l_legpar3 & "<br>"
		'response.end
	
		if l_arr(5) = "" then
			l_fecnac = "null"
		else
			l_fecnac = cambiafecha(l_arr(5),"YMD",true)
		end if
		
		l_sql = "SELECT pronro "
		l_sql = l_sql & " FROM ser_problematica where prodes = '" & l_arr(7) & "'"
		rsOpen l_rs, cn, l_sql, 0 
		if l_rs.eof then
			l_dervul = 0
		else	
			l_dervul = l_rs("pronro")
		end if
		l_rs.close
		
		l_sql = "SELECT mednro "
		l_sql = l_sql & " FROM ser_medida where meddes = '" & l_arr(17) & "'"
		rsOpen l_rs, cn, l_sql, 0 
		if l_rs.eof then
			l_mednro = 0
		else	
			l_mednro = l_rs("mednro")
		end if
		l_rs.close		
	
		l_sql = "INSERT INTO ser_legajo "
		l_sql = l_sql & " (legpar1, legpar2, legpar3, legfecing, legape, legnom, legdni, legfecnac, legdom, pronro,  legsex, legapenommad, legdommad, legdnimad,  legtelmad, legapenompad, legdompad, legdnipad, legtelpad, mednro)"
		l_sql = l_sql & " VALUES ('" & l_legpar1 & "','" & l_legpar2 & "','" & l_legpar3 & "'," & cambiafecha(l_arr(1),"YMD",true) & ",'" & l_arr(2) & "','" & l_arr(3) & "','" & l_arr(4) & "'," & l_fecnac  & ",'" & l_arr(6) & "'," & l_dervul & ",'" & l_arr(8) & "','" & l_arr(9) & "','" & l_arr(10) & "','" & l_arr(11) & "','" & l_arr(12) & "','" & l_arr(13) & "','" & l_arr(14) & "','" & l_arr(15) & "','" & l_arr(16) & "'," & l_mednro & ")"
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
		
	l_fila = l_fila + 1

 	if l_fila>=25 then
		response.flush
	end if

  Loop
  
  cn.CommitTrans 
  
  response.write "<tr><th>La Cantidad Importada de Legajos = " & l_fila & "</th></tr>"


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
