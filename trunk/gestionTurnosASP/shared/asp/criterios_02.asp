<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : criterios_02.asp
Descripcion    : Modulo que se encarga de administrar los criterios de filtrado
Creador        : Scarpa D.
Fecha Creacion : 28/11/2003
Modificacion   :
-----------------------------------------------------------------------------
-->
<% 
on error goto 0

Dim l_rs
Dim l_sql
Dim l_tipo

Dim l_selnro
Dim l_selclase
Dim l_selsist
Dim l_selglobal
Dim l_seldesabr
Dim l_seldesext
Dim l_selsql
Dim l_selasp
Dim l_seltipnro

Dim l_seleccion

l_seleccion = request("seleccion")

%>
<% 

l_tipo     = request("tipo")
l_selnro   = request("selnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Criterios filtrado - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function Validar_Formulario(){
  var errores = 0;
  
  if (document.datos.seldesabr.value == ""){
     errores = 1;
	 alert('La descripción no puede ser nula.')
  }
  
  if (errores == 0){
     if (document.datos.selclase.value == 1){
	    document.datos.sql.value = document.ifrmclase.datos.sql.value;
	 }else{
	    if (document.datos.selclase.value == 2){
	 	    document.datos.asp.value = document.ifrmclase.datos.asp.value;
	    }	 
	 } 
  }
	
  if (errores == 0){
     abrirVentanaH('','voculta',600,400);
     document.datos.target = 'voculta';
     document.datos.action = 'criterios_03.asp?Tipo=<%= l_tipo %>';
	 document.datos.submit(); 
  }	   
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto)
{
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function cambioClase(clase){
  if (clase == 3){
     document.ifrmclase.scrolling="yes";
  }else{
     document.ifrmclase.scrolling="no";  
  }
  document.datos.target = 'ifrmclase'
  document.datos.action = 'criterios_05.asp?clase=' + clase
  document.datos.submit(); 
}

</script>
<%
select Case l_tipo
	Case "A","TA","LE":

      l_selclase  = "1"
      l_selsist   = "0"
      l_selglobal = "0"
      l_seldesabr = ""
      l_seldesext = ""
      l_selsql    = ""
      l_selasp    = ""
      l_seltipnro = "0"

	Case "M","TM":
        l_sql = "SELECT * "
        l_sql = l_sql & " FROM seleccion "
		l_sql = l_sql & " WHERE selnro = " & l_selnro

		rsOpen l_rs, cn, l_sql, 0 

		if not l_rs.eof then
		      l_selclase  = l_rs("selclase")
		      l_selsist   = l_rs("selsist")
		      l_selglobal = l_rs("selglobal")
		      l_seldesabr = l_rs("seldesabr")
		      l_seldesext = l_rs("seldesext")
		      l_selsql    = l_rs("selsql")
		      l_selasp    = l_rs("selprog")
		      l_seltipnro = l_rs("seltipnro")
		end if
		l_rs.Close
		
        l_sql = "SELECT * "
        l_sql = l_sql & " FROM sel_ter "
		l_sql = l_sql & " WHERE selnro = " & l_selnro

		rsOpen l_rs, cn, l_sql, 0 

		l_seleccion = "0"
		do until l_rs.eof 
		   l_seleccion = l_seleccion & "," & l_rs("ternro")
		   l_rs.moveNext
		loop
		l_rs.Close
		
end select

%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">
<form name="datos" action="" method="post" target="voculta">
<input type="hidden" name="selnro" value="<%= l_selnro%>">
<input type="hidden" name="seleccion" value="<%= l_seleccion%>">
<input type="hidden" name="sql" value="<%= l_selsql%>">
<input type="hidden" name="asp" value="<%= l_selasp%>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2" align="left" colspan="2" >Datos del Criterio</td>
	<td class="th2" align="right" colspan="2">
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
  	<tr>
	    <td align="right" nowrap height="10%"><b>Criterio:</b></td>
		<td colspan="3" align="left" nowrap height="90%"><input type="text" name="TextSelNro" size="10" class="deshabinp" maxlength="10" value="<%= l_selnro %>" readonly></td>
	</tr>
	<tr>
	    <td align="right" nowrap><b>Descripci&oacute;n:</b></td>
		<td colspan="3" align="left" nowrap><input type="text" name="seldesabr" size="30" maxlength="60" value="<%= l_seldesabr %>"></td>
	</tr>
	<tr>
	    <td align="right" nowrap><b>Desc.Ext.:</b></td>
		<td colspan="3" align="left" nowrap><textarea name="seldesext" rows="3" cols="40" maxlength="200"><%=trim(l_seldesext)%></textarea></td>
	</tr>
<tr>
	<td align="right">
	<b>Tipo Criterio:</b>
	</td>
	<td>
	  	<select name="seltipnro" size="1">
		<%	l_sql = "SELECT * "
			l_sql  = l_sql  & " FROM tiptableros "
			l_sql  = l_sql  & " ORDER BY tiptabdesc "

			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof		%>	
			<option value="<%= l_rs("tiptabnro") %>" <% if CInt(l_rs("tiptabnro")) = CInt(l_seltipnro) then response.write "selected" end if %>> 
			<%= l_rs("tiptabdesc") %> </option>
		<%		l_rs.Movenext
			loop
			l_rs.Close %>	
		</select>
	</td>
	<td align="right">
	<b>Modelo:</b>
	</td>
	<td >	
      <select name="selclase" onchange="javascript:cambioClase(this.value);">
	  <%if l_tipo <> "LE" then%>
     	  <%if (l_tipo = "A")  OR (l_tipo = "TA") then%>
	      <option value="1" <% if CInt(l_selclase) = 1 then response.write "selected" end if %>>Consulta SQL
	      <option value="2" <% if CInt(l_selclase) = 2 then response.write "selected" end if %>>Programa 
		  <%else%>
	      <option value="1" <% if CInt(l_selclase) = 1 then response.write "selected" end if %>>Consulta SQL
	      <option value="2" <% if CInt(l_selclase) = 2 then response.write "selected" end if %>>Programa 
	      <option value="3" <% if CInt(l_selclase) = 3 then response.write "selected" end if %>>Lista Empleados	  	  		  
		  <%end if%>
	  <%else
	    l_selclase = 3
	  %>
	  <option value="3" <% if CInt(l_selclase) = 3 then response.write "selected" end if %>>Lista Empleados	  	  
	  <%end if %>	  
	  </select>
	</td>
</tr>
<tr>
	<td align="right">	
	</td>
	<td>
    <input type="checkbox" value="1" name="selsist" <% if CInt(l_selsist) = -1 then response.write "checked" end if %>><b>&nbsp;Sistema</b>
	</td>
	<td align="right">	
	</td>
	<td>
    <input type="checkbox" value="1" name="selglobal" <% if CInt(l_selglobal) = -1 then response.write "checked" end if %>><b>&nbsp;Global</b>
	</td>
</tr>  
<tr>
	<td align="center" colspan="4">	
	  <iframe <%if l_tipo <> "LE" then%>scrolling="no"<%end if%> name="ifrmclase" src="" width="98%" height="150" frameborder="0"></iframe> 
	</td>
</tr>  
<tr>
    <td align="right" class="th2" colspan="4">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</form>
<%
set l_rs = nothing
%>

<script>
cambioClase(<%= l_selclase %>);
</script>
</body>
</html>