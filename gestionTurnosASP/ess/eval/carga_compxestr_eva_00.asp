<%Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
' _______________________________________________________________________________________
' archivo	: carga_compxestr_eva_00	
' autor		: CCR
' fecha		: 03-11-2005	
' Modificado: Ene-06 CCR Que se vea la interpretacion de resultados
'			: 04-07-2006 - LA. - agregar la funcion UNESCAPE a las Conductas Observables
'			 			  - LA - arreglo de calculos
'             11-11-2006 - LA. -  Adecuacion a Autogestion
' 			  04-10-2006 - LA. - Si se Termino la Seccion mostrar las competencias desde la evaluacion,sino mostrarlas desde la Estructuras.
' 								 Sino se termino la seccion mantener las nuevas competencias y eliminar las aneriores
'			  24/05/07 - Diego Rosso - Se agrego src="blanc.asp" para que funcione con https.
' _______________________________________________________________________________________

' Variables
dim l_cols
l_cols=7

' de uso local  
  Dim l_evafacnro
  Dim l_evatitdesabr
  Dim l_observables
  Dim l_interpretaciones
  Dim l_usarporcen
  Dim l_evldorcargada
  dim l_totallinea
  dim l_estrnros ' es la lista
  Dim l_factores
  dim l_ternro  
  Dim l_hayalgo  
  Dim l_cantidad  
  
' de base de datos  
  Dim l_sql
  Dim l_rs
  Dim l_rs1
  Dim l_cm

' de parametros de entrada---------------------------------------
  Dim l_evaseccnro
  Dim l_evldrnro
  
' parametros de entrada---------------------------------------  
  l_evaseccnro = Request.QueryString("evaseccnro")
  l_evldrnro   = Request.QueryString("evldrnro")
  
 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_cantidad = 0 

'Buscar el TERNRO del evaluado - y ver si termino la seccion
l_sql = "SELECT empleado, evldorcargada FROM evacab "
l_sql = l_sql & " INNER JOIN evadetevldor ON evadetevldor.evacabnro=evacab.evacabnro "
l_sql = l_sql & " AND evadetevldor.evldrnro= " & l_evldrnro
'response.write l_sql 
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_ternro= l_rs("empleado")
	l_evldorcargada = l_rs("evldorcargada")	
end if
l_rs.Close

'Buscar el si usa porcentaje
l_sql = "SELECT usarporcen FROM evasecc WHERE evaseccnro = " & l_evaseccnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.eof then
	l_usarporcen= l_rs("usarporcen")
end if
l_rs.Close
if l_usarporcen = 0 then
	l_cols = l_cols -2
end if

' Si no se termino la Seccion para el evaluador, busca las competencias a Evaluar.
' Si se termino la Seccion entonces busca las competencias ya evaluadas.
if cint(l_evldorcargada) <> -1 then
	
	'Buscar los Tipos de Estructura cargadas en restric (en la seccion)
	l_sql = "SELECT evaseccestr.tenro FROM evaseccestr "
	l_sql = l_sql & " WHERE evaseccestr.evaseccnro = " & l_evaseccnro
	rsOpen l_rs, cn, l_sql, 0
	' response.write l_sql
	l_estrnros = "0"
	do while not l_rs.eof 
		l_sql = "SELECT estrnro FROM his_estructura WHERE htethasta IS NULL "
		l_sql = l_sql & " AND ternro= " & l_ternro
		l_sql = l_sql & " AND tenro = " & l_rs("tenro")
		rsOpen l_rs1, cn, l_sql, 0
		do while not l_rs1.eof 
			l_estrnros = l_estrnros & "," & l_rs1("estrnro")
			l_rs1.MoveNext
		loop
		l_rs1.Close
		l_rs.MoveNext
	loop
	l_rs.Close

	set l_cm = Server.CreateObject("ADODB.Command")  
	'Crear registros de evaresultado para los facnro y el evldrnro
	l_sql = "SELECT DISTINCT evadescomp.evafacnro FROM evadescomp "
	l_sql = l_sql & " WHERE evadescomp.estrnro IN ( " & l_estrnros & ")"
	rsOpen l_rs, cn, l_sql, 0
		
	l_factores = 0
	
	do while not l_rs.eof
		l_evafacnro = l_rs("evafacnro")
		l_factores = l_factores &","& l_rs("evafacnro")
		
		l_sql = "SELECT * FROM  evaresultado "
		l_sql = l_sql & " WHERE evaresultado.evldrnro   = " & l_evldrnro
		l_sql = l_sql & " AND   evaresultado.evafacnro  = " & l_rs("evafacnro")
		rsOpen l_rs1, cn, l_sql, 0
		'response.write l_sql & "<br>"
		if l_rs1.EOF then
			l_sql = "INSERT INTO evaresultado "
			l_sql = l_sql & " (evldrnro, evafacnro, evaresudesc) "
			l_sql = l_sql & " VALUES (" & l_evldrnro & "," & l_rs("evafacnro") & ",'')"
			l_cm.activeconnection = Cn
			l_cm.CommandText = l_sql
			cmExecute l_cm, l_sql, 0
		end if
		l_rs1.Close
		
		l_rs.MoveNext
	loop
	l_rs.Close
			
	' Borrar los evaaresultados de Puestos anteriores o definiciones anteriores.
	l_sql = "SELECT evafacnro FROM  evaresultado "
	l_sql = l_sql & " WHERE evaresultado.evldrnro =" & l_evldrnro
	l_sql = l_sql & "   AND evaresultado.evafacnro  NOT IN ( " & l_factores & ")"
	rsOpen l_rs1, cn, l_sql, 0
	do while not l_rs1.eof
		l_sql = "DELETE FROM evaresultado "
		l_sql = l_sql & " WHERE evaresultado.evldrnro="& l_evldrnro
		l_sql = l_sql & "   AND evaresultado.evafacnro=" & l_rs1("evafacnro")
		l_cm.activeconnection = Cn
		l_cm.CommandText = l_sql
		cmExecute l_cm, l_sql, 0
	l_rs1.MoveNext
	loop
	l_rs1.Close
'else
end if


'CONTROL DE EVALUADOR LOGEADO =================================================================
dim l_empleg
dim l_evaluador ' guarda el empleg del evaluador del evadetevldor, para comparar con el logeado.
dim l_mostrar '1 o 0 si tiene que mostrar la observacion. 

l_empleg = Session("empleg")
if trim(l_empleg)="" then
   l_empleg = Request.QueryString("empleg")
end if	
 
'buscar la evacab
l_sql = "SELECT empleado.empleg FROM  evadetevldor "
l_sql = l_sql & " INNER JOIN empleado ON empleado.ternro = evadetevldor.evaluador "
l_sql = l_sql & " WHERE evldrnro   = " & l_evldrnro
rsOpen l_rs, cn, l_sql, 0
if not l_rs.EOF then
   l_evaluador = l_rs("empleg")
end if
l_rs.close
 
'Response.Write l_empleg & "<br>" & l_evaluador
l_mostrar = "0"
if trim(l_empleg)<>"" and not isNull(l_empleg) then
   if trim(l_empleg) = trim(l_evaluador) then
   	l_mostrar = "1"
   else	
   	l_mostrar = "0"
   end if
else
	l_mostrar = "1"
end if

' _______________________________________________________________________________
' Procedimiento que busca las conductas observables asociada a las competencias  
' _______________________________________________________________________________
Sub conductasObservables(observables,estrnros, evafacnro)
	l_sql = "SELECT evadescomp.evadcdes, estructura.estrdabr "
	l_sql = l_sql & " FROM evadescomp "
	l_sql = l_sql & " INNER JOIN estructura ON estructura.estrnro = evadescomp.estrnro "
	l_sql = l_sql & " WHERE evadescomp.evafacnro = " & evafacnro
	l_sql = l_sql & " AND   evadescomp.estrnro IN (" & estrnros & ")"
	rsOpen l_rs1, cn, l_sql, 0
	observables=""
	do while not l_rs1.eof
		observables = observables & l_rs1("estrdabr") & " - "& l_rs1("evadcdes")& "\n"
		l_rs1.MoveNext
	loop
	l_rs1.Close
	
end sub

'==============================================================================================
%>

<html>
<head>
<link href="../<%=c_estiloTabla %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Competencias por Estructuras  - Gesti&oacute;n de Desempe&ntilde;o - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
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

<script>
function Controlar(resu){
	if ((resu.value=="")||(resu.value=="0")){
		alert('Seleccione un resultado.');
		resu.focus();
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


function Promedio(){
	
	//alert(cantidad);
//	var r = showModalDialog('calcular_promediocompxEstr_eva_00.asp?evldrnro=<%=l_evldrnro%>&cantidad='+ cantidad, '','dialogWidth:55;dialogHeight:55'); 
	var r = showModalDialog('calcular_promediocompxEstr_eva_00.asp?evldrnro=<%=l_evldrnro%>&usarporcen=<%=l_usarporcen%>', '','dialogWidth:55;dialogHeight:55'); 
	var arr = r.split(",");
	document.datos.promedio.value=arr[0];
	document.datos.totalponderado.value=arr[1];
}

function CalcularTotalLinea(codigo,ponderacion,campo){
	var r = showModalDialog('calcular_totallinea_eva_00.asp?evatrnro='+ codigo+'&ponderacion='+ponderacion, '','dialogWidth:55;dialogHeight:55'); 
	campo.value=r;
}
</script>


<body leftmargin="0" topmargin="0" rightmargin="0" height="100%" width="100%" bottommargin="0" onload="Promedio();">
<form name="datos">

<table border="0" cellpadding="1" cellspacing="1" >
<%
  'BUSCAR evaresultados para MODIFICAR resultados ----------------------------
   l_sql = "SELECT DISTINCT evaresultado.evldrnro, evaresultado.evafacnro,  evatitdesabr ,"
   l_sql = l_sql & " evaresultado.evatrnro, evaresultado.evaresudesc, evaresultado.evaresuejem, "
   l_sql = l_sql & " evaresultado.evarespor, evaresultado.evarestot, "
   l_sql = l_sql & " evafactor.evafacdesabr, evafactor.evafacdesext, evatrvalor,evatitulo.evatitdesabr " 'evadescomp.tenro,evadescomp.estrnro,
   if cint(l_evldorcargada) <> -1 then
   l_sql = l_sql & " ,evafacpor "
   else
   l_sql = l_sql & " ,evarespor "
   end if
   l_sql = l_sql & " FROM evaresultado "
   l_sql = l_sql & " LEFT JOIN evatipresu     ON evaresultado.evatrnro = evatipresu.evatrnro "
   l_sql = l_sql & " INNER JOIN evafactor     ON evafactor.evafacnro = evaresultado.evafacnro "
   l_sql = l_sql & " INNER JOIN evatitulo     ON evatitulo.evatitnro = evafactor.evatitnro "
   if cint(l_evldorcargada) <> -1 then
   l_sql = l_sql & " INNER JOIN evadescomp    ON evadescomp.evafacnro=evafactor.evafacnro  "
   end if
   l_sql = l_sql & " WHERE evaresultado.evldrnro    = " & l_evldrnro
   if cint(l_evldorcargada) <> -1 then
   l_sql = l_sql & " AND   evadescomp.estrnro IN (" & l_estrnros &")"
    end if
   l_sql = l_sql & " ORDER BY evatitdesabr "
	
   l_evatitdesabr=""
   l_hayalgo = "NO"
   'response.write l_sql & "<br>"
   rsOpen l_rs, cn, l_sql, 0
   if not l_rs.eof then
		l_hayalgo = "SI"
   else
   	%><tr style="height:5">
		<th colspan="<%=l_cols%>">No hay Competencias cargadas.<br>Verifique que estén configuradas los Tipos de Estructura, y <BR> las competencias tengan asociadas estos tipos de estructuras<br> o que los empleados tengan asociadas las competencias.</th>
	</tr>
   <%		
   end if
   do while not l_rs.eof 
   if trim(l_evatitdesabr) <> trim(l_rs("evatitdesabr")) then%>
		<tr style="height:5">
			<th align=left class="th2"><%=l_rs("evatitdesabr")%></th>
			<th colspan="<%=l_cols - 1%>" class="th2"></th>
		</tr>
		<tr style="height:5">
			<td><b>Descripci&oacute;n</b></td>
			<td><b>Puntuaci&oacute;n</b></td>
			<% if l_usarporcen=-1 then%>
			<td><b>Ponderaci&oacute;n</b></td>
			<td><b>Total Ponderado</b></td>
			<%end if%>
			<td align=center><b>Observaciones</b></td>
			<td>&nbsp;</td>
			<%if cint(l_evldorcargada) <> -1 then%>
			<td><b>Observables</b></td>
			<%end if%>
		</tr>
		<%l_evatitdesabr = l_rs("evatitdesabr")
	end if%>
	
	<tr style="height:10">
		<td valign="top">
			<%if trim(l_rs("evafacdesext"))="" or isnull(l_rs("evafacdesext")) then%> <%=l_rs("evafacdesabr")%> 
			<%else%>
			<%=l_rs("evafacdesext")%> 
			<%end if%>
		</td>	
		<td nowrap valign="top">
			<%'BUSCAR los resultados  ----------------------------
			l_sql = "SELECT  evatipresu.evatrnro, evaresudes, evatrdesabr, evatrtipo, evatrletra, evatrvalor,evatrrdesext FROM evatipresu "
			l_sql = l_sql & " INNER JOIN evaresu ON evaresu.evatrnro= evatipresu.evatrnro " ' Competencias
			l_sql = l_sql & " WHERE evaresu.evaseccnro= " & l_evaseccnro
			l_sql = l_sql & "   AND evaresu.evafacnro= " & l_rs("evafacnro")
			l_sql = l_sql & " order by evatrvalor "
			rsOpen l_rs1, cn, l_sql, 0
			%>
			<% if l_usarporcen=-1 then%>
			 <select name="evatrnro<%=l_rs("evafacnro")%>" onchange="CalcularTotalLinea(this.value,document.datos.evafacpor<%=l_rs("evafacnro")%>.value,document.datos.totallinea<%=l_rs("evafacnro")%>)">
			<% else  %>
			 <select name="evatrnro<%=l_rs("evafacnro")%>">
			<% end if %>
				<option value=0> 0&nbsp;&nbsp; Sin Evaluar</option>
				<%l_interpretaciones=""
				 do while not l_rs1.eof
					if trim(l_rs1("evaresudes")) <>"" and not isnull(l_rs1("evaresudes")) then
					l_interpretaciones = l_interpretaciones & l_rs1("evatrdesabr") &": "& l_rs1("evaresudes") & "\n"
					end if%>
				<option value=<%=l_rs1("evatrnro")%>><%=l_rs1("evatrvalor")%>&nbsp;&nbsp;&nbsp;<%=l_rs1("evatrdesabr")%></option>
				<%l_rs1.MoveNext
				loop 
				l_rs1.Close
				%>
			</select>
			
			<script>document.datos.evatrnro<%=l_rs("evafacnro")%>.value='<%=l_rs("evatrnro")%>'</script>
			
			<%if trim(l_interpretaciones)="" then%>
				<a href=# onclick="alert('No hay Interpretaciones cargadas para estos resultados.')">?</a></td>
			<%else%>	
				<a href=# onclick="alert('<%=unescape(l_interpretaciones)%>')">?</a></td>
			<%end if%>	
		</td>
		
		<input type="hidden" name="evaresuejem<%=l_rs("evafacnro")%>">
		<% if l_usarporcen=-1 then%>
		<td valign="top">
			<% if cint(l_evldorcargada) <> -1 then %>		
			<INPUT class="rev" readonly name="evafacpor<%=l_rs("evafacnro")%>" size="3" value="<%=l_rs("evafacpor")%>">
			<%else%>	
			<INPUT class="rev" readonly name="evafacpor<%=l_rs("evafacnro")%>" size="3" value="<%=l_rs("evarespor")%>">
			<%end if%>	
		</td>
		<td valign="top">
			<%l_totallinea=""
			if not isNull(l_rs("evarestot")) then
				if l_rs("evarestot") <> "" then
					l_totallinea = l_rs("evarestot")
				end if
			end if
			%>
			<INPUT class="rev" readonly name="totallinea<%=l_rs("evafacnro")%>" size="3" value="<%=l_totallinea%>">
		</td>
		<%else%>
		<INPUT type="hidden" name="totallinea<%=l_rs("evafacnro")%>">
		<INPUT type="hidden" name="evafacpor<%=l_rs("evafacnro")%>" value="100">
		<%end if%>
		
		<td valign="top">
			<%if l_mostrar="1" then%>
			<textarea name="evaresudesc<%=l_rs("evafacnro")%>" cols=20 rows=1><%=trim(l_rs("evaresudesc"))%></textarea>
			<%else%>
			<input type="hidden" name="evaresudesc<%=l_rs("evafacnro")%>" size=5 value="<%=trim(l_rs("evaresudesc"))%>">
			No habilitado.
			<%end if%>
		</td>
		<td nowrap valign="top">
			<a href=# onclick="if (Controlar(document.datos.evatrnro<%=l_rs("evafacnro")%>)) {  grabar.location='grabar_compxestr_eva_00.asp?evafacnro=<%=l_rs("evafacnro")%>&evldrnro=<%=l_evldrnro%>&evaresudesc='+escape(Blanquear(document.datos.evaresudesc<%=l_rs("evafacnro")%>.value))+'&evaresuejem='+escape(Blanquear(document.datos.evaresuejem<%=l_rs("evafacnro")%>.value))+'&evatrnro='+document.datos.evatrnro<%=l_rs("evafacnro")%>.value+'&evarespor='+document.datos.evafacpor<%=l_rs("evafacnro")%>.value+'&evarestot='+document.datos.totallinea<%=l_rs("evafacnro")%>.value;document.datos.grabado<%=l_rs("evafacnro")%>.value='G';}">Grabar</a>
			<br>
			<input class="rev" type="text" style="background : #e0e0de;" readonly disabled name="grabado<%=l_rs("evafacnro")%>" size="1">
		</td>
			
		<% ' Si ya se evaluo no se muestran las conductas observables.
		   if cint(l_evldorcargada) <> -1 then
				conductasObservables l_observables,l_estrnros, l_rs("evafacnro")
				if trim(l_observables)="" then%>
				<td valign=top align=center><a href=# onclick="alert('No hay definidas Conductas Observables \n para las Estructuras del Empleado \n y la Competencia.')">?</a></td>
				<%else%>	
				<td valign=top align=center><a href=# onclick="alert('<%=unescape(l_observables)%>')">?</a></td>
				<%end if%>	
		<% end if%>	
		
		</tr>
		<%
		
		l_rs.Movenext
		loop
		l_rs.Close%>

	<!-- Promedio ----------------------------------->
	<%if l_hayalgo="SI" then%>
    <tr style="height:10">
		<% if l_usarporcen=-1 then%>
		<td align=left></td>
		<td align=left></td>
		<td align=right><b>Total</td>
		<td align=left>
		<input style="background:#e0e0de;" readonly class="blanc" type="text" name="totalponderado" size=5></td>
		</td>
		<%else%>
		<td align=left><input type="hidden" name="totalponderado" size=5></td>
		<td align=left></td>
		<td align=left></td>
		<td align=left></td>
		<%end if%>
		<td align=center colspan="<%=l_cols-4%>"></td>
	</tr>
        <tr style="height:10">
		<td align=left></td>
		<td align=left></td>
		<td align=right><b>Promedio</b></td>
		<td align=left>
		<input style="background:#e0e0de;" readonly class="blanc" type="text" name="promedio" size=5></td>
		<td align=center colspan="<%=l_cols-4%>"></td>
	</tr>
    <%else%>
    <input type="hidden" name="promedio">
    <input type="hidden" name="totalponderado">
    <%end if%>

</form>	
</table>

<iframe src="blanc.asp" name="grabar" style="visibility:hidden;width:0;height:0">

</iframe>

</body>
</html>
