<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : gen_select_emp_01.asp
Creador        : Scarpa D.
Fecha Creacion : 27/11/2003
Descripcion    : Modulo que se encarga de mostrar los tipos de filtros
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

'Variables base de datos
 Dim l_rs
 Dim l_rs2 
 Dim l_sql
 
 Dim l_todos
 Dim l_seltipnro 
 Dim l_dato
 
 l_seltipnro = request("seltipnro")

%>
<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">
<html>

<head>
<link href="/turnos/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<meta http-equiv="Content-Type" http-equiv="refresh" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%></title>
</head>
<script src="/turnos/shared/js/fn_sel_multiple.js"></script>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script>

function SeleccionarFila(fila,selnro)
{
 if (fila.className == "SelectedRow"){
     fila.className = "MouseOutRow"; 
     eliminarDeLista(selnro);   
     parent.cambioSQL(selnro,'');      
 }else{
     if (fila.asp != ''){
        abrirVentana(fila.asp + '?selnro=' + selnro,'',400,100);
     }else{
        fila.className = "SelectedRow";
        agregarALista(selnro);   		
        parent.cambioSQL(selnro,fila.sql);		
     }
 }
}

function cambioSQL(selnro,sql){
  var obj = document.getElementById(selnro);

  if (obj){
      obj.className = "SelectedRow";	
      agregarALista(selnro);   			  	          
      parent.cambioSQL(selnro,sql);  	  
  } 
}

function deSistema(objid){
  var fila = document.getElementById(objid);
  
  return fila.sistema;
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="" method="post">
<table >
    <tr>
        <th width="5%" nowrap>Codigo</th>
        <th nowrap>Descripci&oacute;n</th>
    </tr>

<%

	Set l_rs = Server.CreateObject("ADODB.RecordSet")
	Set l_rs2 = Server.CreateObject("ADODB.RecordSet")	
	
	l_sql =         " SELECT * "
	l_sql = l_sql & " FROM seleccion "
	l_sql = l_sql & " WHERE seltipnro = " & l_seltipnro & " OR selglobal = -1 "
	l_sql = l_sql & " ORDER BY seldesabr "

	rsOpen l_rs, cn, l_sql, 0 
	
	l_todos = ""

	do until l_rs.eof
	
	   if CInt(l_rs("selclase")) = 3 then
	      l_sql = " SELECT * FROM sel_ter WHERE selnro = " & l_rs("selnro")
 	      rsOpen l_rs2, cn, l_sql, 0 		  
		  
		  l_dato = "0"
	      do until l_rs2.eof		  
		     l_dato = l_dato & "," &  l_rs2("ternro")
		     l_rs2.moveNext
		  loop
		  l_rs2.close
	   else
	      l_dato = l_rs("selsql")
	   end if
	
	%>
	    <tr id="<%=l_rs("selnro")%>" sql="<%= l_dato %>" asp="<%= l_rs("selprog")%>" sistema="<%= l_rs("selsist")%>" onclick="javascript:SeleccionarFila(this,<%=l_rs("selnro")%>);">
			<td width="5%"><%= l_rs("selnro")%></td>
	        <td><%= l_rs("seldesabr")%></td>
	    </tr>
	<%
	
	 	 if l_todos="" then
	        l_todos = l_rs("selnro") 
	   	 else
	        l_todos = l_todos & "," & l_rs("selnro") 
	  	 end if
	
	
		l_rs.MoveNext
	loop
	l_rs.Close
	set l_rs = Nothing

%>
</table>

<input type="Hidden" name="cabnro" value="0">
<input type="Hidden" name="sistema" value="0">
<input type="Hidden" name="listanro" value="">
<input type="Hidden" name="listatodos" value="<%= l_todos%>">
</form>

<script>
  setearObjDatos(document.datos.listanro, document.datos.listatodos);
</script>

<%
cn.Close
set cn = Nothing
%>
</body>
</html>
