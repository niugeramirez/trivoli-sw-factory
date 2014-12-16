<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->

<% 
'--------------------------------------------------------------------------
'Archivo       : carga_datosadm_eva_00.asp
'Descripcion   : Carga los datos de evaadmdatos (Alta y modif)
'Creacion      : 22-12-2004
'Autor         : Leticia Amadio.
'Modificado		: Leticia Amadio - 13-10-2005 - Adecuacion a Autogestion
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'				24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'--------------------------------------------------------------------------

on error goto 0

' Variables
' de uso local  
  Dim l_existe  
  Dim l_gerente
  Dim l_revisor
  Dim l_evacabnro
  
  Dim l_horas
  Dim l_fechareunion
  Dim l_basereunion
  Dim l_evldrnro
  Dim l_tipo 
  
dim l_sinfecha
dim l_evatevnro
  
  dim l_emplegactual
  dim l_empleg
' de base de datos 
  Dim l_sql
  Dim l_rs 
  Dim l_rs1
  Dim l_cm 
  
' parametros de entrada-----------------------
  l_evldrnro = Request.QueryString("evldrnro")
  l_empleg = Session("empleg")
  
Set l_rs  = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")

'Buscar el nro de cabecera de esta evaluacion
 l_sql = "SELECT evacabnro, empleg "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro=evadetevldor.evaluador "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
   l_evacabnro= l_rs1("evacabnro")
   l_emplegactual = l_rs1("empleg")
 end if 
 l_rs1.close
 
' Crear registros de evadatosadm
   l_sql = "SELECT * "
   l_sql = l_sql & " FROM  evadatosadm "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evadatosadm.evldrnro" 
   l_sql = l_sql & "		AND evadetevldor.evacabnro = " & l_evacabnro
   rsOpen l_rs1, cn, l_sql, 0
   if l_rs1.EOF then
    l_existe		= "no"
	l_horas			= ""
	l_fechareunion	= "" 
	l_basereunion	= "1"
  else
  	l_existe		= "si"
	l_horas			= l_rs1("horas")
	l_fechareunion	= l_rs1("fechareunion")
	l_basereunion	= l_rs1("basereunion")
   end if
   l_rs1.Close

'Buscar el rold e evaluador que esta entrando
 l_sql = "SELECT evatevnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
   l_evatevnro= l_rs1("evatevnro")
 end if 
 l_rs1.close
 
 l_sinfecha="NO"
 if l_evatevnro <> cevaluador then
	l_sinfecha="SI"
 end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Vistos de Evaluaci&oacute;n - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_numeros.js"></script>

<script>
function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha; 
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto) {
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}



function Validar_Formulario(tipo){
var errores = 0;
var deshabilitado;

		
if ((errores == 0) && ( (!(validanumero(document.datos.horas, 4,0)) )  || (document.datos.horas.value == "")  )){          
    alert("Las horas imputadas deben ser un número entero (hasta 4 digitos).");
	document.datos.horas.focus();	
	errores++;   
}

if ((errores == 0) && ( document.datos.horas.value < 0 ) ){
    alert("El número debe ser positivo.");
	document.datos.horas.focus();	
	errores++;   
}


if ((errores == 0) && (document.datos.fechareunion.value !== "")){
	if ((errores == 0) && (!validarfecha(document.datos.fechareunion))){
   		document.datos.fechareunion.focus();	
   		errores++;   
	}
}



<%if trim(l_empleg)<>"" and (l_empleg=l_emplegactual)  then%>
	if ((errores == 0) && (document.datos.fechareunion.value == "")){
	   	alert("Debe ingresar una fecha.");
   		document.datos.fechareunion.focus();
	   	errores++;   
	}

	<% if l_sinfecha="NO" then%>
		if ((errores == 0) && (!validarfecha(document.datos.fechareunion))){
	   	document.datos.fechareunion.focus();	
   		errores++;   
		}
	<% end if%>

<%end if%>

 // document.datos.basereunion[document.datos.basereunion.selectedIndex].value

if (errores == 0){  
				//document.datos.target = "valida"
	deshabilitado = false
	if (document.datos.basereunion.disabled != false) {
	 	document.datos.basereunion.disabled = false
		deshabilitado = true
	}
	
	document.datos.target = "grabar"
	document.datos.action = "grabar_datosadm_eva_00.asp?Tipo="+tipo+'&evldrnro=<%=l_evldrnro%>'	
	document.datos.submit(); 
	
	if (tipo == 'M') {
		document.datos.grabado.value='M';
	} else {
		document.datos.grabado.value='G';
	}
	
	if (deshabilitado) {
		document.datos.basereunion.disabled = true
	}
}

}


</script>
<style>
.blanc
{
	font-size: 10;
	border-style: none;
	background : transparent;
}
.rev
{
	font-size: 10;
	border-style: none;
}
</style>
</head>

<body>

<form name="datos" action=""  method="post">
<input type="Hidden" name="terminarsecc" value="SI">

<%
     ' buscar datos de proyecto relacionado con evldrnro
l_sql = "SELECT  estrdabr,  proygerente, proyrevisor, evaproyfdd,  evaclinom, evadetevldor.evatevnro   " 
l_sql = l_sql & " FROM evaproyecto "
l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evaproyecto.estrnro  "  
   	' nombre del cliente 
l_sql = l_sql & " INNER JOIN evaengage  ON evaengage.evaengnro = evaproyecto.evaengnro "
l_sql = l_sql & " INNER JOIN evacliente ON evacliente.evaclinro = evaengage.evaclinro "
  	' 
l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evaevenro = evaevento.evaevenro "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
   
l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
   
rsOpen l_rs1, cn, l_sql, 0


' Selecciona el nombre del gerente y del revisor.
if not l_rs1.eof then
	l_sql= " SELECT terape, terape2,ternom, ternom2 "
	l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs1("proygerente")
  	rsOpen l_rs, cn, l_sql, 0
	l_gerente = l_rs("terape") & " " &  l_rs("terape2") & " " & l_rs("ternom") &  " "  & l_rs("ternom2")
	l_rs.Close
	
	l_sql= " SELECT terape, terape2,ternom, ternom2 "
	l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs1("proyrevisor")
  	rsOpen l_rs, cn, l_sql, 0
	l_revisor = l_rs("terape") & " " &  l_rs("terape2") & " " & l_rs("ternom") &  " "  & l_rs("ternom2")
	l_rs.Close
end if
%>

<table border="0" cellpadding="0" cellspacing="0">
<tr style="border-color :CadetBlue;">
	<th colspan="3" align="left" class="th2"> Carga de Datos Administrativos </th>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Linea de Servicio: </b></td>
 	<td><%= l_rs1("estrdabr")%>&nbsp;</td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Gerente: </b></td>
	<td><%=l_gerente %></td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Revisor: </b></td>
	<td> <%= l_revisor%></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Per&iacute;odo de Revisi&oacute;n: </b></td>
	<td>Desde el &nbsp;  <%=l_rs1("evaproyfdd")%></td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Cliente: </b></td>
	<td> <%= l_rs1("evaclinom")%>&nbsp;</td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Horas Imputadas: </b></td>
	<td> <input type="text" name="horas" value="<%=l_horas%>" size="6" maxlength="4">&nbsp;</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Reuni&oacute;n: </b></td>
	<td> 
	<% if l_rs1("evatevnro") <> cevaluador then %>
		<input class="rev" style="background : #e0e0de;" type="text" name="fechareunion" value="<%=l_fechareunion%>" size="10" readonly>&nbsp;
	<% else %> 
		<input type="text" name="fechareunion" value="<%=l_fechareunion%>" size="10">&nbsp;
		<a href="Javascript:Ayuda_Fecha(document.datos.fechareunion)"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a></td>
	<% end if %>
	</td>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="right"><b>Base para la Revisi&oacute;n: </b></td>
	<td>
	<% if l_rs1("evatevnro") <> cevaluador then %>
		<select name="basereunion" size="1" disabled> 
	<% else %> 
		<select name="basereunion" size="1">
	<% end if %>
	
		<%select case l_basereunion 
			case 1: %>	
				<option value="1" selected>Moderada</option>
				<option value="2">Urgente</option>
		<% case 2: %>
				<option value="1">Moderada</option>
				<option value="2" selected>Urgente</option>
		<% end select%>	
		</select>
	</td> 
	<td>&nbsp;</td>
</tr>
<tr>
	<td>&nbsp; </td> <td>&nbsp;</td>
	<% if l_existe = "si" then      ' lo de l_tipo = "M" no me funciono!!!  %>	
		<td valign=top><a href=# onclick="Javascript:Validar_Formulario('M');">Modificar</a>
	<%else %>
		<td valign=top><a href=# onclick="Javascript:Validar_Formulario('A');">Grabar</a>
	<%end if%>	
		<input type="text" readonly disabled name="grabado" size="1">
		</td>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>

<%
l_rs1.Close
%>
</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>

</body>
</html>
