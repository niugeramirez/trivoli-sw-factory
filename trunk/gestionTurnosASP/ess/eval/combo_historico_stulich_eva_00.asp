<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--
Archivo    : combo_historico_stulich_eva_00.asp
Autor      : Leticia Amadio
Creacion   : /06/2005
Descripcion: Muestra los reportes de stulich ya generados.
Modificacion: 05-05-2006 - LA. - Cambio la funcion cint a cdbl en bpronro

-->
<% 
on error goto 0

 Dim l_rs
 Dim l_rs2
 Dim l_sql
 Dim l_disabled
 Dim l_ancho
 Dim l_str
 
' Variables
 Dim l_evento
 Dim l_estrnro1
 Dim l_estrnro2
 'Dim l_tenro3
 Dim l_estrnro3
 Dim l_consejero
 Dim l_aconsejado 
 Dim l_bpronro
 Dim l_bpronroaux
 
  l_bpronro  = Request.QueryString("bpronro") 
' l_evento  = Request.QueryString("evento") ...

  ' response.write " bpronro " & l_bpronro & "<br>"
 'l_bpronro = 2133
  ' response.write " bpronro " & l_bpronro 
 l_ancho     = request("ancho")
 l_disabled  = request("disabled")
 
 if l_ancho = "" then
	l_ancho = 100
 end if  
 
 l_user = Session("Username")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<html>
<head>
	<title>Untitled</title>
	
<script>

function actualiza(){
  var bpronro;
 /* var evento;
  var estrnro1
  var estrnro2
  var estrnro3
  var consejero
  var aconsejado */
 
 	/*
  evento 	 = document.datos.historico[document.datos.historico.selectedIndex].evento;

  estrnro1	 = document.datos.historico[document.datos.historico.selectedIndex].estrnro1;
  estrnro2	 = document.datos.historico[document.datos.historico.selectedIndex].estrnro2;
  estrnro3	 = document.datos.historico[document.datos.historico.selectedIndex].estrnro3;
  consejero	 = document.datos.historico[document.datos.historico.selectedIndex].consejero;
  aconsejado = document.datos.historico[document.datos.historico.selectedIndex].aconsejado; */
  
  bpronro    = document.datos.historico[document.datos.historico.selectedIndex].bpronro;
      //  alert(bpronro);
  
 /*' if (proaprob == 0 && pronro != -1)
	todospro = -1;
  else			
	todospro = 0; 
	*/
	parent.cambioCS(bpronro,-1);
	// parent.cambioCS(evento, estrnro1, estrnro2, estrnro3,aconsejado,consejero,bpronro,-1);

}

</script>	
	
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" scroll=no bgcolor="#808080">
<% ' response.write  l_user & "<br>" %>
<form name="datos">
<%
Set l_rs = Server.CreateObject("ADODB.RecordSet") 
Set l_rs2 = Server.CreateObject("ADODB.RecordSet") 

l_bpronroaux = ""

l_sql = " SELECT DISTINCT bpronro, grupo,depto, categ, Fecha, Hora, evento, ternro, consejero, ternro, apeynom, todosgrupo, todosdepto, todoscateg, todosaconsej,todosconsej " 'tenro2, , estrnro3  tenro3
l_sql = l_sql & " FROM rep_stulich "
l_sql = l_sql & " ORDER BY Fecha DESC, Hora DESC"

rsOpen l_rs, cn, l_sql, 0
'	response.write l_sql
%>
<select <%= l_disabled %> class="<% if l_disabled <> "" then response.write "deshabinp" else response.write "habinp" end if %>" name="historico" size="1" style="width:100%" width="100%" onChange="javascript:actualiza();">
<option value="" bpronro="">&laquo;Seleccione una opci&oacute;n&raquo;</option>
<!--  <option value="" evento="" estrnro1="" estrnro2="" estrnro3="" consejero="" aconsejado="" bpronro="">&laquo;Seleccione una opci&oacute;n&raquo;</option>  "width:<%= l_ancho%>" <%= l_ancho%>-->
<%	
do until l_rs.eof
	if l_bpronroaux  <> l_rs("bpronro") then
		l_sql = " SELECT evaevenro , evaevedesabr FROM evaevento  WHERE evaevenro =" & l_rs("evento")
		rsOpen l_rs2, cn, l_sql, 0
		'l_str  = " Evento: " 
		l_str  = " ("& l_rs2("evaevenro") &") - " & l_rs2("evaevedesabr") & " -"
		l_rs2.Close
		
		if l_rs("todosgrupo")= -1 then 
			l_str = l_str & " - por Grupo " ' Todos Grupos
		else 
			l_str = l_str & " - para Grupo: " & l_rs("grupo")
		end if 
		
		if l_rs("todosdepto")= -1 then 
			l_str = l_str & " - por Departamento " 
		else 
			l_str = l_str & " - para Depto: " & l_rs("depto")
		end if 
		
		if l_rs("todoscateg")= -1 then 
			l_str = l_str & " - por Categoría " 
		else 
			l_str = l_str & " - para Categ: " & l_rs("categ")
		end if 
		
		if l_rs("todosconsej") <>  -1 then 
			l_str = l_str & " - Consej: " & l_rs("consejero")
		end if 
		
		if l_rs("todosaconsej") <>  -1 then 
			l_str = l_str & " - Aconsej: " & l_rs("apeynom")
		end if 
	
	'l_str = l_str & " - " & l_rs("Fecha")
    'l_str = l_str & " " & l_rs("Hora")
	
	'l_str = l_str & " - " & l_rs("bpronro")
	%>	

	<option evento="<%=l_rs("evento")%>" estrnro1="<%=l_rs("grupo")%>" estrnro2="<%=l_rs("depto")%>" estrnro3="<%=l_rs("categ")%>" consejero="<%=l_rs("consejero")%>" aconsejado="<%=l_rs("ternro")%>" bpronro="<%=l_rs("bpronro")%>" value="1" <%if cdbl(l_bpronro) = cdbl(l_rs("bpronro")) then%> selected<% else %> v=0 <%end if%>> 	
		<%'=l_rs("bpronro")%><%= l_str & "&nbsp;&nbsp;&nbsp;"%>
	</option>
	<%
		l_bpronroaux = l_rs("bpronro")
	 
	end if
	l_rs.Movenext
loop
			%>
				<!--
					<option evento="<%'=l_rs("evento")%>" estrnro1="<%'=l_rs("grupo")%>" estrnro2="<%'=l_rs("depto")%>" estrnro3="<%'=l_rs("categ")%>" consejero="<%'=l_rs("consejero")%>" aconsejado="<%'=l_rs("ternro")%>" bpronro="<%'=l_rs("bpronro")%>" value="1"> 	
					<% 'if (CStr(l_rs("empresa")) = CStr(l_empresa)) AND (CStr(l_rs("pliqnro")) = CStr(l_pliqnro)) AND (CStr(l_rs("pronro")) = CStr(l_pronro)) AND (CStr(l_rs("proaprob")) = CStr(l_proaprob)) then response.write "selected" end if%>
				-->
<%
			l_rs.Close
%>	
	  </select>
</form>
<script>
 actualiza();
</script>
</body>
</html>
