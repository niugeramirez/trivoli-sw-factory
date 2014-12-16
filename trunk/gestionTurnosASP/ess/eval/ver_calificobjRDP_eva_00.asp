<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: ver_califobjrdp
'Descripción	: Cargar resultado calif gral de objetivos RDP
'Autor			:  -05-2005
'Fecha			: L Amadio
'Modificado		: 21-07-2005 - L.A. - permitir evaluar RDP si existe al menos una RDE cerrada
'				:  03-08-2005 - L.A. - Cambiar cod de proyecto por cod de evento.
'            		13-10-2005 - Leticia Amadio -  Adecuacion a Autogestion
'					24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
'================================================================================

on error goto 0

' Variables
 
' de uso local  
  dim l_evatrnro 
  dim l_evacabnro 
  dim l_ternro  
  dim l_evaluador 
  dim  l_evaevenro
  dim l_evatevnro
  dim l_proyectos
  dim l_cantproys
  dim l_cantproysaprob
  dim l_proys
  dim l_datos 
  dim l_sincerrarRDE 
  
  dim l_gerente
  dim l_socio
  dim l_horas
  dim l_objResu
  dim l_objGrabar
  dim l_objGrabar2
  
  dim l_evaproynro
   
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs2
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  Dim l_evapernro  ' VERRR si se pasa siempre este parametro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")
  l_evapernro  = Request.QueryString("evapernro")
  
  l_evaproynro =  Request.QueryString("evaproynro")

  ' busca ternro del Evaluado (Aconsejado) y el evacabnro
  l_ternro=""
  Set l_rs = Server.CreateObject("ADODB.RecordSet")
  l_sql = "SELECT evacab.evacabnro, empleado,  evaevento.evaevenro, evadetevldor.evatevnro "
  l_sql = l_sql & " FROM evacab "
  l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
  l_sql = l_sql & " INNER JOIN evaevento ON evaevento.evaevenro = evacab.evaevenro "
  l_sql = l_sql & " WHERE evadetevldor.evldrnro = " & l_evldrnro
  rsOpen l_rs, cn, l_sql, 0
  if not l_rs.eof then
	l_evacabnro  = l_rs("evacabnro")
	l_ternro   = l_rs("empleado")
	l_evaevenro = l_rs("evaevenro")
	l_evatevnro = l_rs("evatevnro")
  end if
  l_rs.close
  set l_rs=nothing


' ______________________________________________________________
'  Busco si existe un reg de esta evaluacion en evagralobj 
l_evatrnro  = ""
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT *  FROM  evagralobj"
l_sql = l_sql & " WHERE evagralobj.evldrnro="& l_evldrnro
rsOpen l_rs1, cn, l_sql, 0 
if l_rs1.EOF then
	set l_cm = Server.CreateObject("ADODB.Command")  
	l_sql = "INSERT INTO evagralobj "
	l_sql = l_sql & " (evldrnro, evatrnro) "
	l_sql = l_sql & " VALUES (" & l_evldrnro & ",NULL)" 
	l_cm.activeconnection = Cn
	l_cm.CommandText = l_sql
	cmExecute l_cm, l_sql, 0
else
	l_evatrnro= l_rs1("evatrnro")
end if
l_rs1.Close
set l_rs1=nothing


' ______________________________________________________________________________________________________________
' buscar todos los proyectos en que participo el empleado (igual periodo y estrnro que evento RDP - )
l_proyectos = "0" 
l_cantproys = 0 
l_proys="" 
l_cantproysaprob = 0 
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT DISTINCT evaproyecto.evaproynro, cabaprobada, evento.evaevenro "
l_sql = l_sql & " FROM evaevento "
l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evapernro = evaevento.evaperact "
l_sql = l_sql & " INNER JOIN evaevento evento ON evento.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
l_sql = l_sql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro "
l_sql = l_sql & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro  AND evatip_estr.tenro ="& cdepartamento 
l_sql = l_sql & " WHERE  evaevento.evaevenro ="& l_evaevenro &" AND evacab.empleado="&l_ternro
		' evacab.cabaprobada= -1 AND 
rsOpen l_rs1, cn, l_sql, 0 
do while not l_rs1.eof 
	if l_rs1("cabaprobada") = -1 then 
		l_proyectos = l_proyectos & "," & l_rs1("evaproynro") 
		l_cantproysaprob = l_cantproysaprob + 1 
	else 
		l_proys = l_proys &  " - " & l_rs1("evaevenro")  ' l_rs1("evaproynro")
	end if 
	l_cantproys = l_cantproys +1 
l_rs1.MoveNext 
loop 
l_rs1.close 

l_proyectos = Split(l_proyectos,",") 

l_sincerrarRDE="NO" 
if l_cantproys <> l_cantproysaprob then 
	l_sincerrarRDE="SI" 
end if 

' HARDCORE ------------
'l_sincerrarRDE="NO"   




' ________________________________________________________
'  busca los resultados para la calific gral de objetivos 
sub datosObj (objResu, objGrabar,objGrabar2,evatrnro)
objResu=""
objGrabar=""
objGrabar2="" 

'BUSCAR la descripcion del resultado  ----------------------------
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  evatipresu.evatrnro, evatipresu.evatrvalor, evatipresu.evatrdesabr "
l_sql = l_sql & " FROM evatipresu  "
l_sql = l_sql & " WHERE evatrtipo=2 ORDER BY evatrvalor "
rsOpen l_rs1, cn, l_sql, 0 

objResu= objResu & "<select name=""evatrnro"" disabled>"
objResu = objResu & "<option value=>&nbsp;&nbsp; Sin Evaluar</option>"
do while not l_rs1.eof
	objResu= objResu & "<option value="&l_rs1("evatrnro")&">"& "&nbsp;&nbsp;&nbsp;"& l_rs1("evatrdesabr")&"</option>"
 	l_rs1.MoveNext
loop 
objResu = objResu & "	</select> "
l_rs1.Close
set l_rs1 = nothing
	
if evatrnro = ""  or isNull(evatrnro) then 
	objResu= objResu & " <script>document.datos.evatrnro.value='';</script>"
else
  objResu= objResu & " <script>document.datos.evatrnro.value="& evatrnro &";</script>"
end if

'objGrabar= "<a href=# onclick=""if (Controlar('texto',document.datos.evatrnro)) {grabar.location='grabar_calificobjRDP_eva_00.asp?evldrnro="&l_evldrnro&"&evapernro="&l_evapernro&"&evatrnro='+document.datos.evatrnro.value; document.datos.grabado.value='G'; }"">Grabar</a>"
'objGrabar2 = " <input size=1 class=""rev"" type=""text""  style=""background : #e0e0de;"" readonly disabled name=""grabado""> "

objGrabar= "<a href=# onclick="" return false; "">Grabar</a>"
objGrabar2 = " <input size=1 class=""rev"" type=""text""  style=""background : #e0e0de;"" readonly disabled name=""grabado""> "
end sub
%>

<html>
<head>
<link href="../<%=c_estiloTabla  %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<style>
.rev {
	font-size: 10;
	border-style: none;
}
</style>
</head>

<script>

function Controlar(texto,valor){
	if (texto.value==""){
		alert('Ingrese un Objetivo.');
		texto.focus();
		return false;
	}
	else
		if (valor.value==""){
			alert('Seleccione un resultado.');
			valor.focus();
			return false;
		}
		else
			return true;
}	

function Blanquear(texto){
 var  aux;
 aux = replaceSubstring(texto,"'","")
 aux = replaceSubstring(aux,"´","")
 return aux;
}

function replaceSubstring(inputString, fromString, toString) {
   // Goes through the inputString and replaces every occurrence of fromString with toString
   var temp = inputString;
   if (fromString == "") {
      return inputString;
   }
   if (toString.indexOf(fromString) == -1) { // If the string being replaced is not a part of the replacement string (normal situation)
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } else { // String being replaced is part of replacement string (like "+" being replaced with "++") - prevent an infinite loop
      var midStrings = new Array("~", "`", "_", "^", "#");
      var midStringLen = 1;
      var midString = "";
      // Find a string that doesn't exist in the inputString to be used
      // as an "inbetween" string
      while (midString == "") {
         for (var i=0; i < midStrings.length; i++) {
            var tempMidString = "";
            for (var j=0; j < midStringLen; j++) { tempMidString += midStrings[i]; }
            if (fromString.indexOf(tempMidString) == -1) {
               midString = tempMidString;
               i = midStrings.length + 1;
            }
         }
      } // Keep on going until we build an "inbetween" string that doesn't exist
      // Now go through and do two replaces - first, replace the "fromString" with the "inbetween" string
      while (temp.indexOf(fromString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(fromString));
         var toTheRight = temp.substring(temp.indexOf(fromString)+fromString.length, temp.length);
         temp = toTheLeft + midString + toTheRight;
      }
      // Next, replace the "inbetween" string with the "toString"
      while (temp.indexOf(midString) != -1) {
         var toTheLeft = temp.substring(0, temp.indexOf(midString));
         var toTheRight = temp.substring(temp.indexOf(midString)+midString.length, temp.length);
         temp = toTheLeft + toString + toTheRight;
      }
   } // Ends the check to see if the string being replaced is part of the replacement string or not
   return temp; // Send the updated string back to the user
} // Ends the "replaceSubstring" function


</script>
<body leftmargin="0" topmargin="0" rightmargin="0">
<form name="datos">
<input type="Hidden" name="terminarsecc" value="--">
<input type="Hidden" name="terminarsecc2" value="">
<% datosObj l_objResu, l_objGrabar,l_objGrabar2, l_evatrnro %>
<table border="0" cellpadding="0" cellspacing="1" width="100%">
<% if l_sincerrarRDE="SI" then %>
<tr height="20">
	<td colspan="8" align="left" width="25%">
		<% if l_cantproysaprob > 0 then  %>
		<b> AVISO:</b> El empleado no tiene todas sus RDE's cerradas. <br>
		<% else %>
		<b> AVISO:</b> No se permite calificar la secci&oacute;n, dado que el empleado no tiene ninguna RDE cerrada. <br>		  
		<% end if%>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; Eventos sin RDE's cerradas: <%=l_proys%>
	</td>
</tr>
<tr><td colspan="8" align="center">&nbsp;</td> </tr>
<% end if %>
<tr height="20"><td colspan="8" align="center">&nbsp;</td> </tr>
<tr height="20">
	<td colspan="8">
	<table><tr>
		<td align="right" width="25%"><b>Calificaci&oacute;n General de Objetivos</b></td> 
		<td  width="25%"><%=l_objResu %> &nbsp;</td>
		<td  align="left" width="25%"><%=l_objGrabar %> <br> &nbsp;<%=l_objGrabar2 %></td>
		<td>&nbsp;</td>
	</tr></table>
	</td>
</tr>
<tr>
	<th class="th2">Evento</th>
	<th class="th2">Engagement </th>
	<th class="th2">Cliente </th>
	<th class="th2">Gerente </th>
	<th class="th2">Socio </th>
	<th class="th2">Horas Imputadas </th>
	<th class="th2">Desde</th> 
	<th class="th2">Calificaci&oacute;n Gral.</th>
</tr>
<%
 	' buscar lista de engagement en la que participo el empleado,(con = periodo - estrnro) y con RDE cerrada--------------
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT evaengage.evaengnro, evaengdesabr, evaclinom, proygerente, proysocio, evaproyfdd, evaproyecto.evaproynro, evento.evaevenro"
 l_sql = l_sql & " FROM evaengage "
 l_sql = l_sql & " INNER JOIN evacliente  ON evacliente.evaclinro  = evaengage.evaclinro "
 l_sql = l_sql & " INNER JOIN evaproyecto ON evaproyecto.evaengnro = evaengage.evaengnro "
 l_sql = l_sql & " INNER JOIN evaevento evento ON evento.evaproynro = evaproyecto.evaproynro "
 l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
 l_sql = l_sql & " INNER JOIN evaevento ON evaproyecto.evapernro = evaevento.evaperact "
 l_sql = l_sql & " INNER JOIN evatipoeva ON evatipoeva.evatipnro = evaevento.evatipnro "
 l_sql = l_sql & " INNER JOIN evatip_estr ON evatip_estr.evatipnro = evatipoeva.evatipnro AND evatip_estr.estrnro=evaproyecto.estrnro  AND evatip_estr.tenro ="& cdepartamento 
 l_sql = l_sql & " WHERE evacab.cabaprobada=-1  AND evaevento.evaevenro =" &l_evaevenro &" AND evacab.empleado="&l_ternro
 rsOpen l_rs, cn, l_sql, 0 
 
 if not l_rs.eof then 
 	
	do while not l_rs.eof 
		
		Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
		' buscar nom de gerente y socio
		l_sql= " SELECT terape, terape2,ternom, ternom2 "
		l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs("proygerente")
  		rsOpen l_rs1, cn, l_sql, 0 
		l_gerente = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
		l_rs1.Close 
		l_sql= " SELECT terape, terape2,ternom, ternom2 "
		l_sql = l_sql & " FROM tercero  WHERE ternro= " & l_rs("proysocio")
  		rsOpen l_rs1, cn, l_sql, 0
		l_socio = l_rs1("terape") & " " &  l_rs1("terape2") & " " & l_rs1("ternom") &  " "  & l_rs1("ternom2")
		l_rs1.Close
		
		l_sql = "SELECT evadetevldor.evldrnro, horas  " 
		l_sql = l_sql & " FROM evaproyecto " 
		l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & " INNER JOIN evadatosadm ON evadatosadm.evldrnro = evadetevldor.evldrnro "
		l_sql = l_sql & " WHERE evacab.empleado =" & l_ternro & " AND evaproyecto.evaproynro=" & l_rs("evaproynro") & " AND evadetevldor.evatevnro=" & cevaluador
		rsOpen l_rs1, cn, l_sql, 0 
		if l_rs1.eof then 
			l_horas = "--"
		else
			l_horas = l_rs1("horas")
		end if
		l_rs1.Close
		
		' buscar - Calificacion gral de OBJ (evldrnro - evatrnro ) 
		l_sql = "SELECT evadetevldor.evldrnro, evagralobj.evatrnro, evatrdesabr " 
		l_sql = l_sql & " FROM evaproyecto " 
		l_sql = l_sql & " INNER JOIN evacab ON evacab.evaproynro = evaproyecto.evaproynro "
		l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro = evacab.evacabnro "
		l_sql = l_sql & " INNER JOIN evagralobj ON evagralobj.evldrnro = evadetevldor.evldrnro "
		l_sql = l_sql & " INNER JOIN evatipresu ON evatipresu.evatrnro = evagralobj.evatrnro "
	  	l_sql = l_sql & " WHERE evacab.empleado = " & l_ternro & " AND evaproyecto.evaproynro=" & l_rs("evaproynro") & " AND evadetevldor.evatevnro="& cevaluador
		rsOpen l_rs1, cn, l_sql, 0 
		
%>
		<tr>
			<td><%= l_rs("evaevenro")%></td>
			<td><%= l_rs("evaengdesabr")%></td>
			<td><%= l_rs("evaclinom")%></td>
			<td><%= l_gerente %> </td>
			<td><%= l_socio %></td>
			<td><%= l_horas%></td>
			<td><%= l_rs("evaproyfdd")%></td> 
			<% if not l_rs1.eof then  %>
				<td align="right">
					<%= l_rs1("evatrdesabr")%> &nbsp;
					<a href=# onclick="Javascript:abrirVentana('detalle_califobjRDE_eva_00.asp?evaproynro=<%= l_rs("evaproynro")%>&ternro=<%=l_ternro%>','',550,300,',scrollbars=yes')" title="Detalle de Calificación de Objetivos RDE"> ++</a>
				</td>
			<% else %>
				<td align="center"> -- </td>
			<% end if%>
		</tr>
<%  	l_rs1.Close
	l_rs.MoveNext
	loop 

else  %>
   <tr height="20">
	 	<td colspan="8" align="center"> No existen proyectos asociados a el per&iacute;odo.</td>
   </tr>
<%
end if
  
l_rs.close
set l_rs=nothing
%>
</form>		
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">
</iframe>
<iframe name="terminarsecc" src="termsecc_calificobj_eva_00.asp?evacabnro=<%=l_evacabnro%>&evaseccnro=<%=l_evaseccnro%>&evldrnro=<%=l_evldrnro%>&evatevnro=<%=l_evatevnro%>" style="visibility:hidden;width:0;height:0"><! -- &habCalifGral=<%'=l_habCalifGral%> -->
</iframe>
</body>
</html>

<%
cn.close
%>
