<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'--------------------------------------------------------------------------
'Archivo       : ver_datosadm_eva_00.asp
'Descripcion   : Muestra los datos de evaadmdatos (readonly)
'Creacion      : 23 - 12- 2004
'Autor         : Leticia Amadio.
'Modificacion  : 
'            13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'			 24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'--------------------------------------------------------------------------
on error goto 0

' Variables
  'Dim l_existe  
 Dim l_evacabnro
 Dim l_horas
 Dim l_fechareunion
 Dim l_basereunion
 Dim l_evldrnro

Dim l_gerente
Dim l_revisor 
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' parametros de entrada---------------------------------------  
  l_evldrnro = Request.QueryString("evldrnro")
  
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
Set l_rs = Server.CreateObject("ADODB.RecordSet")

'Buscar el nro de cabecera de esta evaluacion
 l_sql = "SELECT evacabnro "
 l_sql = l_sql & " FROM  evadetevldor "
 l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
 rsOpen l_rs1, cn, l_sql, 0
 if not l_rs1.EOF then
   l_evacabnro= l_rs1("evacabnro")
 end if 
 l_rs1.close

' Crear registros de evadatosadm
   l_sql = "SELECT * "
   l_sql = l_sql & " FROM  evadatosadm "
   l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evldrnro=evadatosadm.evldrnro" 
   l_sql = l_sql & "		AND evadetevldor.evacabnro = " & l_evacabnro
   rsOpen l_rs1, cn, l_sql, 0
 if l_rs1.EOF then
    ' l_existe = "no"
	l_horas = ""
	l_fechareunion = "" 
	l_basereunion  = "1"
 else
  	'l_existe = "si"
	l_horas = l_rs1("horas")
	l_fechareunion = l_rs1("fechareunion")
	l_basereunion  = l_rs1("basereunion")
 end if
	
   l_rs1.Close

%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Carga de Vistos de Evaluaci&oacute;n - Evaluaci&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Nuevo_Dialogo(w_in, pagina, ancho, alto){
 return w_in.showModalDialog(pagina,'', 'center:yes;dialogWidth:' + ancho.toString() + ';dialogHeight:' + alto.toString() + ';');
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">

<form name="datos">
<input type="Hidden" name="terminarsecc" value="SI">

<%   ' buscar datos de proyecto relacionado con evldrnro
   l_sql = "SELECT  estrdabr,  proygerente, proyrevisor, evaproyfdd,  evaclinom   " 
   
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


<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr style="border-color :CadetBlue;">
	<th colspan="3" align="left" class="th2">Datos Administrativos</th>
</tr>
<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Linea de Servicio: </b></td>
	<td><%= l_rs1("estrdabr")%>&nbsp;</td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Gerente: </b></td>
	<td><%= l_gerente %></td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Revisor: </b></td>
	<td> <%= l_revisor %></td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Per&iacute;odo de Revisi&oacute;n: </b></td>
	<td>Desde el &nbsp;  <%=l_rs1("evaproyfdd")%></td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Cliente: </b></td>
	<td> <%= l_rs1("evaclinom")%>&nbsp;</td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Horas Imputadas: </b></td>
	<td> <%=l_horas%>&nbsp;</td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Reuni&oacute;n: </b></td>
	<td> <%=l_fechareunion%>&nbsp;</td> <td>&nbsp;</td></tr>
<tr>
	<td align="right"><b>Base para la Revisi&oacute;n: </b></td>
	<td><select name="basereunion" size="1" readonly disabled>
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
	<td>&nbsp;</td></tr>


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
