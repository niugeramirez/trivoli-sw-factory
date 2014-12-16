<%Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
'Archivo		: emp_est_formales_adp_02.asp
'Descripción	: relacion empleado estudios formales
'Autor			: lisandro moro
'Fecha			: 09/08/2003 
'Modificado:
' A.Bayon - 15-09-2003 - Se recibe como parámetro el nivel						
'		 				Se agrega el nivel como parte de la llave para la selección
'						Fecha desde y hasta no obligatorias
'						Comparación en las fechas (si no están vacías)
' A.Bayon - 16-09-2003 - Corrección de error al seleccionar título
' CCRossi - 23/09/2003 - Cambiar la manera de actualizar titulos e Institucion en el onchange 
' CCRossi - 29/09/2003 - Institucion debe poder seleccionarse. No relacionada directamente al titulo
' CCRossi - 29/09/2003 - Carrera deshabilitado si Completada=true. Sino opcional.
' CCRossi - 02/10/2003 - "completo" en lugar de "completado"
'						 Titulo no obligatorio
' CCRossi - 03/10/2003 - Carrera no obligatoria
' CCRossi - 06-10-2003 - Sacar carrera si es completo
' CCRossi - 07-10-2003 - poner tilde de completo a la derecha del nivel estudio
'	y la cantidad de materias que se habilite si es INCOMPLETO
'   Corregir el tema de la institucion, que se permita elegir a pesar de no elegir un titulo
' CCRossi - 10-10-2003- faltaba titdesabr del ORDER en el select. En Informix pincha eso.
' CCRossi - 17-10-2003- Siempre permite cargar la Institucion
' Modificado - 25-02-2004 - Scarpa D. - Se agrego la opcion de estudio actual
' Modificado - 05-03-2004 - Scarpa D. - Se agrego el campo descripcion
'			 - 17-10-2005 - Leticia A. - Adaptacion a Autogestion - arreglo el Alta de Est. F
'============================================================================================ 

on error goto 0

'Variables
'de parametro de entrada
Dim l_tipo
Dim l_ternro

'de base de datos
dim l_rs
dim l_rs1
dim l_sql
  
'variables
Dim l_titnro
Dim l_instnro
Dim l_carredunro
Dim l_nivnro
Dim l_capcomp
Dim l_capcantmat
Dim l_capestact
Dim l_capanocur
Dim l_capfecdes
Dim l_capfechas
Dim l_caprango
Dim l_capprom
Dim l_capactual
Dim l_futdesc

' parametros entrada
  l_tipo    = Request.QueryString("tipo")
  l_ternro  = l_ess_ternro

' BODY ===========================================================================
  select Case l_tipo
	Case "A":
		l_nivnro = 0
		l_titnro = 0
		l_instnro = 0
		l_carredunro = 0
		l_capcomp    = 0
		l_capcantmat = ""
		l_capestact  = 0
		l_capanocur  = ""
		l_capfecdes  = ""
		l_capfechas  = ""
		l_caprango   = ""
		l_capprom    = ""
	Case "M":
		l_nivnro 	  = Request.QueryString("nivnro")
		l_titnro 	  = Request.QueryString("titnro")
		l_instnro	  = Request.QueryString("instnro")
		l_carredunro  = Request.QueryString("carredunro")
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_sql = "SELECT   ternro, nivnro, titnro, instnro, carredunro, capfecdes, capfechas, capcomp, capanocur, capcantmat, capestact, caprango, capprom "
		l_sql = l_sql & " FROM  cap_estformal "
		l_sql = l_sql & " WHERE ternro   = " & l_ternro 
		if l_titnro <> "" then
		l_sql = l_sql & " AND titnro     = " & l_titnro 
		end if
		if l_instnro <> "" then
		l_sql = l_sql & " AND instnro    = " & l_instnro 
		end if
		if l_nivnro <> "" then
		l_sql = l_sql & " AND nivnro     = " & l_nivnro
		end if
		if l_carredunro <> "" then
		l_sql = l_sql & " AND carredunro = " & l_carredunro
		end if
		
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_carredunro= l_rs("carredunro")
			l_capcomp   = l_rs("capcomp")
			
			if isNull(l_rs("capcantmat")) then
			   l_capcantmat = ""
			else
				l_capcantmat= l_rs("capcantmat")
			end if
			if isNull(l_rs("capanocur")) then
				l_capanocur = ""
			else
				l_capanocur = l_rs("capanocur")
			end if
			l_capestact = l_rs("capestact")
			l_capfecdes = l_rs("capfecdes")
			l_capfechas = l_rs("capfechas")
			if isnull(l_rs("caprango")) then
				l_caprango = ""
			else
				l_caprango  = l_rs("caprango")
			end if
			if IsNull(l_rs("capprom")) then
				l_capprom   = ""
			else
				l_capprom   = l_rs("capprom")
			end if
			if isNull(l_titnro) or l_titnro="" then
				l_titnro=0
			end if
		end if
		l_rs.Close
		set l_rs = nothing
end select


response.write "<script languaje='javascript'>" & vbCrLf
response.write "function CargarInstitucion(vez){ " & vbCrLf

dim l_institucion

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  titnro,instnro, titdesabr "
l_sql = l_sql & " FROM titulo "
l_sql = l_sql & " ORDER BY titdesabr "
rsOpen l_rs1, cn, l_sql, 0 

response.write "document.datos.unica.value = false;" & vbCrLf

do while not l_rs1.eof
		response.write "if (document.datos.titnro.value == " & l_rs1(0) & ") {" & vbCrLf
			if l_rs1(1)<> "" and l_rs1(1)<> "0" then
				response.write "document.datos.instnro.value = '"&l_rs1(1)&"';" & vbCrLf
				response.write "document.datos.instnro.disabled = true;" & vbCrLf
				response.write "document.datos.unica.value = true;" & vbCrLf
			end if
		response.write "};" & vbCrLf
l_rs1.MoveNext
loop
l_rs1.Close
set l_rs1 = nothing

response.write "if (document.datos.titnro.value == 0) {" & vbCrLf
response.write "if (vez =='0') {" & vbCrLf
response.write "	document.datos.instnro.value=0;" & vbCrLf
response.write "    document.datos.instnro.disabled=false;" & vbCrLf
response.write "	};" & vbCrLf
response.write "    document.datos.instnro.disabled=false;" & vbCrLf
response.write "	};" & vbCrLf
response.write "else {" & vbCrLf
response.write "		if (document.datos.instnro.value == '0') " & vbCrLf
response.write "			document.datos.instnro.disabled=false;" & vbCrLf
response.write "		else {" & vbCrLf
response.write "			if (document.datos.unica.value=='true') {" & vbCrLf
response.write "				document.datos.instnro.disabled=true } ;" & vbCrLf
response.write "			else {" & vbCrLf
response.write "				document.datos.instnro.disabled=false } ;" & vbCrLf
response.write "			};" & vbCrLf
response.write "	};" & vbCrLf

if l_tipo="M" and l_instnro<>"" then
response.write "if (document.datos.titnro.value!=='0' && vez=='1') {" & vbCrLf
response.write "	document.datos.instnro.value="&l_instnro&";" & vbCrLf
response.write "	if (document.datos.unica.value=='true') {" & vbCrLf
response.write "		document.datos.instnro.disabled=true;" & vbCrLf
response.write "	};" & vbCrLf
response.write "};" & vbCrLf
end if

response.write "};" & vbCrLf
response.write "</script>" & vbCrLf

'----------------------------------------------------------------------------------
response.write "<script languaje='javascript'>" & vbCrLf
response.write "function CargarTitulo(){ " & vbCrLf

response.write " var obj = document.datos.titnro ;" & vbCrLf &_ 
		   "     var i;" & vbCrLf &_
		   "     var long = obj.length;" & vbCrLf &_
		   "      for (i=long;i>=0;i--){" & vbCrLf &_
    	   "        obj.remove(i) ;" & vbCrLf &_
    	   "      }" & vbCrLf 

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  titnro,titdesabr, nivnro, instnro "
l_sql = l_sql & " FROM titulo "
l_sql = l_sql & " ORDER BY titdesabr "
rsOpen l_rs1, cn, l_sql, 0 
'''''''''''''''''''
		response.write "newOp = new Option();" & vbCrLf
		response.write "newOp.value = '0';" & vbCrLf
		response.write "newOp.text = '<<Seleccione una opción>>';" & vbCrLf	
		response.write "obj.add(newOp);" & vbCrLf
'''''''''''''''''''


do while not l_rs1.eof
		response.write "if (document.datos.nivnro.value == " & l_rs1(2) & ") {" & vbCrLf
		response.write "newOp = new Option();" & vbCrLf
		response.write "newOp.value = " & l_rs1(0) & ";" & vbCrLf
		response.write "newOp.text = '" & l_rs1(1) &"("&l_rs1(0)&")"& "';" & vbCrLf	
		response.write "obj.add(newOp);" & vbCrLf
		response.write "};" & vbCrLf
l_rs1.MoveNext
loop
response.write "if (obj.length == 0 ) { " & vbCrLf
response.write "   newOp = new Option();" & vbCrLf
response.write "   newOp.value = '0' ;" & vbCrLf
response.write "   newOp.text = ' ';" & vbCrLf	
response.write "   obj.add(newOp);" & vbCrLf
response.write "   document.datos.titnro.value='0';" & vbCrLf
response.write "   obj.disabled=true;" & vbCrLf
response.write "   document.datos.instnro.disabled=false;" & vbCrLf
response.write "  }" & vbCrLf
response.write "else" & vbCrLf
response.write "  {" & vbCrLf
response.write "		obj.disabled=false;" & vbCrLf
response.write "		newOp = new Option();" & vbCrLf
response.write "		newOp.value = '0';" & vbCrLf
response.write "		newOp.text = ' ';" & vbCrLf	
response.write "		obj.add(newOp);" & vbCrLf
response.write "		document.datos.instnro.disabled=false;" & vbCrLf
response.write "   }" & vbCrLf

response.write " document.datos.titnro.value = "&l_titnro& ";" & vbCrLf

response.write "if (document.datos.capcomp.checked == false ) { " & vbCrLf
response.write "		document.datos.titnro.value='0';" & vbCrLf
response.write "		document.datos.titnro.disabled=true;" & vbCrLf
response.write "		document.datos.instnro.disabled=false;" & vbCrLf
response.write "   }" & vbCrLf

response.write "};" & vbCrLf
response.write "</script>" & vbCrLf


l_rs1.Close
set l_rs1 = nothing

%>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><% If l_tipo = "A" then%>Alta<% Else  %>Modificaci&oacute;n<% End If %> de Datos Formales - Administraci&oacute;n de Personal - RHPro &reg;</title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script>
function Ayuda_Fecha(txt)
{
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}

function Validar_Formulario(){

// document.datos.caprango.value = Trim(document.datos.caprango.value);
// document.datos.capprom.value = Trim(document.datos.capprom.value);
// document.datos.futdesc.value = Trim(document.datos.futdesc.value);

if (document.datos.capcomp.checked)
	document.datos.capcomp.value = -1;
else 
	document.datos.capcomp.value = 0;
	
if (document.datos.capestact.checked)
	document.datos.capestact.value = -1;
else
	document.datos.capestact.value = 0;

if (document.datos.nivnro.value=="0"){
	alert("Debe seleccionar un Nivel");
	document.datos.nivnro.focus();
}
else
if ((isNaN(document.datos.capanocur.value))||(document.datos.capanocur.value < 0)){
	alert("Debe ingresar un año válido de cuatro dígitos.");
	document.datos.capanocur.focus();
	document.datos.capanocur.select();
}
else 
if ((isNaN(document.datos.capcantmat.value))||(document.datos.capcantmat.value < 0)){
	alert("Debe ingresar una cantidad de materias válida.");
	document.datos.capcantmat.focus();
	document.datos.capcantmat.select();
}
else 
if ((document.datos.capfecdes.value!="") && (!validarfecha(document.datos.capfecdes))){
	document.datos.capfecdes.focus();
	document.datos.capfecdes.select();
}
else
if ((document.datos.capfechas.value!="")&& (!validarfecha(document.datos.capfechas))){
	document.datos.capfechas.focus();
	document.datos.capfechas.select();
}
else
if ((document.datos.capfecdes.value!="")&&(document.datos.capfechas.value!="")&&(!menorque(document.datos.capfecdes.value,document.datos.capfechas.value))) {
	alert("La Fecha Desde no puede ser mayor que la Fecha Hasta.");
	document.datos.capfecdes.focus();
	document.datos.capfecdes.select();
}
else
if (document.datos.nivnro.value==0){
	alert("Debe seleccionar un Nivel de Estudio");
	document.datos.nivnro.focus();
	document.datos.nivnro.select();
} else {
	document.datos.instnroaux.value = document.datos.instnro.value;
	document.datos.submit();
	}
}

function Carrera(valor){
	
	if (valor==-1){
		valor =true;
		}
	
	if (valor == true) 
		{
		document.datos.carredunro.value ="0";
		document.datos.carredunro.disabled =true;
		document.datos.capcantmat.disabled =true;
		document.datos.capcantmat.value ="0";
		document.datos.titnro.disabled=false;
		if (document.datos.unica.value=false || document.datos.titnro.value==0)
			document.datos.instnro.disabled=false;
		}
	else
		{
		document.datos.carredunro.disabled =false;
		document.datos.titnro.value = "0";
		document.datos.titnro.disabled=true;
		document.datos.instnro.disabled=false;
		document.datos.capcantmat.disabled =false;
		}
}

</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0"  onload="<%if l_tipo="M" then%>CargarTitulo();CargarInstitucion(1);Carrera(<%=l_capcomp%>);<%else%>document.datos.capcomp.checked=false;Carrera(false);<%end if%>" onunload="Javascript:window.opener.location.reload();">

<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%" >
	<tr>
		<th colspan="3" align="left"><% If l_tipo = "A" then%>Alta<% Else  %>Modificaci&oacute;n<% End If %> de Estudios Formales</th>
		<th align="right">&nbsp;</th>	
	</tr>

<form name="datos" action="ag_emp_est_formales_adp_03.asp" method="post" >
	<input type="hidden" name="tipo"			value="<%= l_tipo %>">
	<input type="Hidden" name="ternro"			value="<%= l_ternro %>">
	<input type="Hidden" name="nivnroant"		value="<%= l_nivnro %>">
	<input type="Hidden" name="titnroant"		value="<%= l_titnro %>">
	<input type="Hidden" name="instnroant"		value="<%= l_instnro %>">
	<input type="Hidden" name="carredunroant"	value="<%= l_carredunro %>">
	<input type="Hidden" name="instnroaux"		value="<%= l_instnro %>">
	<input type="Hidden" name="unica">
	<tr>
		<td align="right">Nivel de estudio:</td>
		<td>
		<select name="nivnro" size="1" onChange="CargarTitulo();" style="width:250px;">
			<option value=""></option>
			<%Set l_rs = Server.CreateObject("ADODB.RecordSet") 
  		    l_sql = "SELECT nivnro, nivdesc, nivsist, nivcodext, nivobligatorio,  nivproximo "
			l_sql  = l_sql  & "FROM nivest "
			rsOpen l_rs, cn, l_sql, 0
				do until l_rs.eof%>	
					<option value="<%=l_rs("nivnro")%>"><%=l_rs("nivdesc")%></option>
				<%l_rs.Movenext
				loop
			l_rs.Close%>	
		</select>
		<script>document.datos.nivnro.value = '<%=l_nivnro%>';</script>
		
		</td>
		<td align="right">Completo&nbsp;
		<input type="Checkbox" name="capcomp" onclick="Carrera(this.checked);" value="<%=l_capcomp %>"<% If l_tipo = "M" and l_capcomp = -1 then %> checked<% End If %> ></td>
		<td align="right">Cantidad de Materias&nbsp;
		<input type="Text" size="3" name="capcantmat" value="<%= l_capcantmat %>" ></td>
	</tr>

	<tr>
    <td align="right">T&iacute;tulo:</td>
	<td colspan=3>
	<select name=titnro onchange="CargarInstitucion(0);">
	</select>
	<script>document.datos.titnro.value = '<%=l_titnro%>';</script>
	</td>
	</tr>

	<td align="right">Instituci&oacute;n:</td>
	<td colspan="3">
		<select name=instnro>
			<option value="0"></option>
			<%Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT  institucion.instnro,instdes  "
			l_sql = l_sql & " FROM institucion "
			l_sql = l_sql & " WHERE instedu = -1 "
			rsOpen l_rs1, cn, l_sql, 0 
			do until l_rs1.eof%>
				<option value="<%=l_rs1("instnro")%>"><%=l_rs1("instdes")%></option> 
			<%l_rs1.MoveNext
			loop
			l_rs1.Close
			set l_rs1 = nothing
			%>
		</select>
		<script>document.datos.instnro.value = '<%=l_instnro%>';</script>
	</td>
	</tr>
	<tr>
		<td align="right">Carrera Educ.:</td>
		<td colspan="3">
			<select name="carredunro" size="1">
			<option value="0"></option>
			<%Set l_rs = Server.CreateObject("ADODB.RecordSet") 
  		    l_sql = "SELECT carredunro, carredudesabr, carredudesext "
			l_sql  = l_sql  & "FROM cap_carr_edu "
			rsOpen l_rs, cn, l_sql, 0
				do until l_rs.eof%>	
					<option value=<%=l_rs("carredunro")%>><%=l_rs("carredudesabr")%> </option>
				<%l_rs.Movenext
				loop
			l_rs.Close%>
		</select>
		</td>
		<script>document.datos.carredunro.value = '<%=l_carredunro%>';</script>
	</tr>
	<tr>
		<td align="right">Estudia Actualmente</td><td style="padding-left:0px;"><input type="Checkbox" name="capestact" value="<%= l_capestact %>" <% If l_tipo = "M" and l_capestact = 1 then %> checked<% End If %>></td>
		<td align="right">Año cursado</td><td><input type="Text" size="3" name ="capanocur" value="<%= l_capanocur %>"></td>
	</tr>
	<tr>
		<td align="right">Fecha Desde:</td><td>
		<input  type="text" name="capfecdes" size="10" maxlength="10" value="<%= l_capfecdes %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.capfecdes);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td align="right">Fecha Hasta:</td><td>
		<input  type="text" name="capfechas" size="10" maxlength="10" value="<%= l_capfechas %>">
		<a href="Javascript:Ayuda_Fecha(document.datos.capfechas);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
	</td>
	</tr>
	<tr>
		<td align="right">Rango:</td><td colspan="3"><input type="Text" size="30" maxlength="30" name="caprango" value="<%= l_caprango %>"></td>
	</tr>
	<tr>
		<td align="right">Promedio:</td><td colspan="3"><input type="Text" size="30" maxlength="30" name="capprom" value="<%= l_capprom %>"></td>
	</tr>	
	</form>
<tr>
    <td colspan="4" align="right" class="th2">
		<a class=sidebtnABM href="javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>

</body>
</html>
