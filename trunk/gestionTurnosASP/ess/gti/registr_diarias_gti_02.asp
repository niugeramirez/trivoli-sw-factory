<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
-----------------------------------------------------------------------------
Archivo        : registr_diarias_gti_02.asp
Descripcion    : Modulo que se encarga de mostrar los datos de una registracion
Modificacion   :
   18/09/2003 - Scarpa D. - Coordinacion con el tablero del empleado
   01/10/2003 - Scarpa D. - Permitir seleccionar un empleado
   08/10/2003 - Scarpa D. - Cambio en el mecanismo usado para abrir el 03
   10/10/2003 - Scarpa D. - Si la registracion es nula la muestra como desconocida   
   14/10/2003 - Scarpa D. - Reloj, mostrar el cod. interno atras.	   
   15/10/2003 - Scarpa D. - Poder seleccionar empleado inactivos
   20/10/2003 - Scarpa D. - Correccion al cargar los datos.
   27/10/2003 - Scarpa D. - Actualizar el tablero cuando se cierra la ventana
   15/12/2004 - Fernando Favre - Se valida el campo Hora Desde con la funcion validanumero
   21-08-2007 - Diego Rosso - Se agrego src="blanc.asp" para https
-----------------------------------------------------------------------------
-->
<% 
Dim l_regnro
Dim l_regfecha
Dim l_reghora 
Dim l_regestado
Dim l_regentsal
Dim l_regmanual
Dim l_relnro
Dim l_ternro

Dim l_empleado
Dim l_empleg
Dim l_reloj

Dim l_tipo
Dim l_sql
Dim l_rs
Dim l_fechadesde
Dim l_fechahasta
Dim l_datos

Set l_rs = Server.CreateObject("ADODB.RecordSet") 'locreo aqui para poder usarlo en las consultas de abajo

l_tipo = request.querystring("tipo")
l_fechadesde = request.querystring("fechadesde")
l_fechahasta = request.querystring("fechahasta")

l_ternro = l_ess_ternro
l_empleg = l_ess_empleg

l_sql = " SELECT * FROM empleado WHERE ternro=" & l_ternro
rsOpen l_rs, cn, l_sql, 0 
l_empleado = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
l_empleg   = leg
l_rs.close

%>
<html>
<head>
<link href="../<%=c_estilo%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Registraciones - Gesti&oacute;n de Tiempos - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>
<script src="/serviciolocal/shared/js/fn_buscar_emp.js"></script>
<script src="/serviciolocal/shared/js/fn_help_emp.js"></script>
<script>

function Tecla(num){
  if (num==13) {
     if (document.datos.empleg.value != document.datos.emplegant.value){
         document.datos.emplegant.value = document.datos.empleg.value;
         buscar_emp_todos_porLeg(document.datos.empleg.value);
		 return false;
	 }	
  }
  return num;
}

function nuevoempleado(ternro,empleg,terape,ternom)
{
 if (empleg != 0){
	 document.datos.empleg.value = empleg;
	 document.datos.ternro.value = ternro;
	 document.datos.empleado.value = terape + ','+ ternom;
     document.datos.emplegant.value = document.datos.empleg.value;	 
 }else{
	 document.location.reload();
 }
}

function buscarEmpleado(){
   if (document.datos.empleg.value != document.datos.emplegant.value){
       document.datos.emplegant.value = document.datos.empleg.value;
	   buscar_emp_todos_porLeg(document.datos.empleg.value);
   }	
}

function Validar_Formulario()
{
if (document.datos.ternro.value == "") 
	alert("Debe ingresar un Empleado.");
else
if (validarfecha(document.datos.regfecha)) 
{
if (!validanumero(document.datos.reghora1,2,0)|| (document.datos.reghora1.value<0)|| (document.datos.reghora1.value>23) 
		|| !validanumero(document.datos.reghora2,2,0)|| (document.datos.reghora2.value<0)|| (document.datos.reghora2.value>59)
		|| (document.datos.reghora1.value.length != 2) || (document.datos.reghora2.value.length != 2)) 
			alert("Debe ingresar la Hora Desde o esta mal ingresada.");
else {
	document.datos.target = "valida";
	document.datos.action = "registr_diarias_gti_12.asp"
	document.datos.submit();
}
}
}

function DatosCorrectos(){
	abrirVentanaH('','vent_oculta',200,200);
	document.datos.cantingresos.value = parseInt(document.datos.cantingresos.value) + 1;  
	document.datos.target = "vent_oculta";
	document.datos.action = "registr_diarias_gti_03.asp?tipo=<%= l_tipo%>";
	document.datos.submit();
}

function DatosIncorrectos(msg){
	alert(msg);
}

function radioclick(radio){
	if (radio=="E") {
		document.datos.regentsal.value= "E";
	}
	else{ 
		if (radio=="S") {
			document.datos.regentsal.value= "S";
		}
		else {
			document.datos.regentsal.value= "D";
		}
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

function cerrarVent(){
  if (parseInt(document.datos.cantingresos.value) > 0){
     window.opener.actualizar();
  }  
  window.close();  
}

function Horario_Habitual(){
	with (document.datos) {
		if (relnro.value == "")
			alert('Debe seleccionar un reloj.');
		else
			valida.location.href = "registr_diarias_gti_13.asp?ternro=" + ternro.value + "&relnro=" + relnro.value + "&fecha=" + regfecha.value;
	}
}
</script>

<% 

select Case l_tipo
	Case "A","TA":
		l_regnro = ""
		l_regfecha = CDate(Day(Date) & "/" & Month(Date) & "/" & Year(Date))
		l_reghora = ""
		l_regestado = ""
		l_regentsal = "D"
		l_regmanual = ""
		l_relnro = ""
		
		l_sql = " SELECT relnro FROM his_estructura " 
        l_sql = l_sql & " INNER JOIN Alcance_Testr ON his_estructura.tenro = Alcance_Testr.tenro " 
        l_sql = l_sql & " INNER JOIN estructura ON his_estructura.estrnro = estructura.estrnro " 
        l_sql = l_sql & " INNER JOIN gti_rel_estr ON gti_rel_estr.estrnro = his_estructura.estrnro " 
        l_sql = l_sql & " WHERE (estructura.estrest = -1) AND (tanro = 7) AND (ternro = " & l_ternro & ") AND "
        l_sql = l_sql & " (htetdesde <= " & cambiafecha(l_regfecha,"YMD",true) & ") AND " 
        l_sql = l_sql & " ((" & cambiafecha(l_regfecha,"DMY",true) & " <= htethasta) or (htethasta is null))" 
        l_sql = l_sql & " ORDER BY alcance_testr.alteorden DESC, his_estructura.htetdesde Desc "  		
		
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_relnro = l_rs("relnro")
		end if
		l_rs.Close		
	Case "M":
		l_regnro = request("cabnro")
		l_sql = "SELECT  ternro, regfecha, reghora, relnro, regestado, regentsal, regmanual "
		l_sql = l_sql & "FROM gti_registracion "
		l_sql = l_sql & " WHERE  regnro="&l_regnro
		rsOpen l_rs, cn, l_sql, 0
		if not l_rs.eof then
			l_regfecha = l_rs("regfecha")
			l_reghora = l_rs("reghora")
			l_regestado = l_rs("regestado")
			if isNull(l_rs("regentsal")) then
			   l_regentsal = "D"
			else
			   l_regentsal = l_rs("regentsal")
			end if
			l_regmanual = l_rs("regmanual")
			l_relnro = l_rs("relnro")
		end if
		l_rs.Close
end select
%>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0">

<form name="datos" 	target="vent_oculta" action="registr_diarias_gti_03.asp?tipo=<%= l_tipo%>" method="post">
<input type="Hidden" name="regnro" value="<%= l_regnro %>">
<input type="Hidden" name="fechadesde" value="<%= l_fechadesde %>">
<input type="Hidden" name="fechahasta" value="<%= l_fechahasta %>">
<input type="Hidden" name="emplegant" value="<%= l_empleg %>">
<input type="Hidden" name="cantingresos" value="0">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
    <th colspan="4">Datos de Registraciones </th>
  </tr>
<tr>
	<td align="right">
		<b>Empleado:</b>
	</td>
	<td align="left" colspan="3">
		<input type="Hidden" name="ternro" value="<%= l_ternro %>">	
		
		<input type="text" tabindex="1" name="empleg" size="8" maxlength="8" readonly="true" value="<%= l_empleg %>">	
        <input type="Text" name="empleado" size="30"  readonly="true" value="<%= l_empleado %>">	  		
	</td>
</tr>
<tr>
    <td align="right" height="20"><b>Fecha:</b></td>
	<td>
		<input tabindex="3" type="text" name="regfecha" size="10" maxlength="10" value="<%= l_regfecha %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.regfecha);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>	
	</td>
	<td rowspan = 3>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
							<td>
						    <input type="Hidden" name="regentsal" value="<%= l_regentsal %>">
							<input type="radio" name="entsal"  value="E" 
							<% if l_regentsal = "E" then
									response.write "Checked" 
							   end if %>  onclick="Javascript:radioclick('E');">
							Entrada 
							</td>														   
						</tr>
						<tr>
							<td>
							<input type="radio" name="entsal"  value="S" 
							<% if l_regentsal = "S" then
									response.write "Checked" 
							   end if %>  onclick="Javascript:radioclick('S');">
							Salida 
							</td>														   
						</tr>
						<tr>
							<td>
							<input type="radio" name="entsal"  value="D" 
							<% if l_regentsal = "D" then
									response.write "Checked" 
							   end if %>  onclick="Javascript:radioclick('D');">
							Desconocido 
							</td>														   
						</tr>
					</table>
	
	
	</td>

</tr>
<tr>
	<td align="right"><b>Hora:</b></td>
	<td>
		<input type="text" name="reghora1" size="2" maxlength="2" tabindex="4" value="<%= mid(l_reghora,1,2) %>">
		<b>:</b>
		<input type="text" name="reghora2" size="2" maxlength="2" tabindex="5" value="<%= mid(l_reghora,3,2) %>">
	</td>
</tr>
</td>
	<td align="right"><b>[Reloj]</b></td>
	<td>
			<select name="relnro" size="1" tabindex="6">
			<%	l_sql = "SELECT relnro, reldabr "
				l_sql  = l_sql  & "FROM gti_reloj "
				rsOpen l_rs, cn, l_sql, 0
				do until l_rs.eof	 
					l_reloj=l_rs("relnro")%>	
					<option value= <%= l_reloj %> > 
					<%= l_rs("reldabr") & "&nbsp;(" & l_reloj & ")" %> </option>
					<%			l_rs.Movenext
				loop
				l_rs.Close
				%>	
			</select>
			<script>
				document.datos.relnro.value = "<%= l_relnro %>";
			</script>
							
  	</td>
	</tr>
<tr>
    <td align="right" class="th2" colspan = 4>
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
        <a class=sidebtnABM href="Javascript:window.close();">Cancelar</a>		   
	</td>
</tr>
</table>
<iframe src="blanc.asp" name=valida width="500" height="500"></iframe>
</form>
</body>
</html>
<%
set l_rs= nothing 	

cn.close

set cn= nothing 	
%>
