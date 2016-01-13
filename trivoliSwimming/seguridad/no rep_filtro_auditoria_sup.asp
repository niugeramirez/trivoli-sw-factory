<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%
'Archivo	: rep_filtro_auditoria_sup.asp
'Descripción: Filtra datos. 
'NOTA		: La pagina que llame a esta, tiene que declarar una variable opant en Jscript
'Autor		: CCRossi- 
'Fecha		: 28-04-2004
'Modificado	: 
'             20-01-2005 - JMH - Se quito del filtro la empresa y se permite selección 
'                                múltiples de configuración
'			  20-07-2005 - Fapitalle N. - Se agrega el filtro por empleados
'			  22-07-2005 - Fapitalle N. - Modifico el pasaje de param
'---------------------------------------------------------------------------------------
%>

<script>
var vectoracciones = new Array();
var vectorusuarios = new Array();
var vectorconfauds = new Array();
var vectorempresas = new Array();
</script>
<%
on error goto 0
Dim l_rs 
Dim l_rs1	 
Dim l_sql 

Dim l_opant

Dim l_fechadesde
Dim l_fechahasta
Dim l_acciones
Dim l_acnro
Dim l_usuarios
Dim l_iduser
Dim l_caudnro
Dim l_empresas
Dim l_empnro
Dim l_campos

Dim l_orden
Dim l_ordenado
Dim l_empsel
Dim l_seleccion

Dim l_empleg
Dim l_seltipnro
Dim l_pronro

l_seltipnro = 1

'--------------------------------------------------------------------------------------------------------------------------
' Inicializacion de recupero de los valores del filtro anterior

' opant = String donde los parametros estan separados por ?

l_opant = request("opant")
'Response.Write(l_opant)
'Response.End

if l_opant <> "" then
	l_opant = Split(request("opant"), "?", -1, 1)
else
	redim l_opant(11)
	l_opant(0) = ""
	l_opant(1) = ""
	l_opant(2) = ""
	l_opant(3) = ""
	l_opant(4) = ""
	l_opant(5) = ""
	l_opant(6) = ""
	l_opant(7) = ""
	l_opant(8) = ""
	l_opant(9) = ""
	l_opant(10) = ""
	l_opant(11) = ""
end if


Set l_rs = Server.CreateObject("ADODB.RecordSet")

' acciones -----------------------------------------------------------------
if l_opant(0)="" then
	l_acciones = "-1"
else
	l_acciones = l_opant(0)
end if

' acnro ------------------------------------------------------------
if l_opant(1)="" then
	l_acnro = "0"
else 
	l_acnro = l_opant(1)
end if

' usuarios ------------------------------------------------------------------
if l_opant(2)="" then
	l_usuarios = "-1"
else
	l_usuarios = l_opant(2)
end if

'iduser ------------------------------------------------------------
if l_opant(3)="" then
	l_iduser = "0"
else 
	l_iduser = l_opant(3)
end if

' caudnro  ------------------------------------------------------------
if l_opant(4)="" then
	l_caudnro = "0"
else 
	l_caudnro = l_opant(4)
end if

'fechadesde--------------------------------------------------------------
if l_opant(5)="" then
	l_fechadesde = Date
else
	l_fechadesde = l_opant(5)
end if
'fechahasta--------------------------------------------------------------
if l_opant(6)="" then
	l_fechahasta = Date
else
	l_fechahasta = l_opant(6)
end if

' orden---------------------------------------------------------------
if l_opant(7)="" then
	l_orden = "Fecha"
else
	l_orden = l_opant(8)
end if

' ordenado------------------------------------------------------------
if l_opant(8)="" then
	l_ordenado = "Asc"
else
	l_ordenado = l_opant(9)
end if

' empsel----------0-todos---1 algunos---2 uno-------------------------
if l_opant(9)="" then
	l_empsel = 0
else
	l_empsel = l_opant(9)
end if

''----------------------------------------------------------------------------------------------------------------------------
' FUNCION para Cargar Cangiguración de auditoría
'----------------------------------------------------------------------------------------------------------------------------
response.write "<script src='/rhprox2/shared/js/fn_fechas.js'></script>" & vbCrLf
response.write "<script languaje='javascript'>" & vbCrLf
response.write "function CargarConfig(confnro){ " & vbCrLf

'aca voy a tener que armar lista de los pliqnro entre el desde y el hasta elegidos
Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT caudnro, cauddes  "
l_sql = l_sql & " FROM confaud "
rsOpen l_rs1, cn, l_sql, 0

response.write " var obj = document.datos.caudnro ;" & vbCrLf &_ 
		   "     var i;" & vbCrLf &_
		   "     var long = obj.length;" & vbCrLf &_
		   "      for (i=long;i>=0;i--){" & vbCrLf &_
    	   "        obj.remove(i) ;" & vbCrLf &_
    	   "      }" & vbCrLf 


response.write " var i ;" & vbCrLf
dim entro
entro =0

do while not l_rs1.eof

		response.write "			newOp = new Option();" & vbCrLf
		response.write "			newOp.value = " & l_rs1(0) & ";" & vbCrLf
		response.write "			newOp.text = '" & l_rs1(1) & "';" & vbCrLf	
		response.write "			obj.add(newOp);" & vbCrLf
l_rs1.MoveNext
loop

response.write " if (obj.length > 0) { " & vbCrLf
response.write "   newOp = new Option();" & vbCrLf
response.write "   newOp.value = '0' ;" & vbCrLf
response.write "   newOp.text = 'Todos';" & vbCrLf	
response.write "   obj.add(newOp,0);" & vbCrLf
response.write " } " & vbCrLf


response.write "if (obj.length == 0 ) { " & vbCrLf
response.write "   newOp = new Option();" & vbCrLf
response.write "   newOp.value = '0' ;" & vbCrLf
response.write "   newOp.text = '';" & vbCrLf	
response.write "   obj.add(newOp);" & vbCrLf
response.write "  }" & vbCrLf
response.write "else { " & vbCrLf

if l_opant(4) <> "" then
	'Marcar seleccionados lo que viene de opant(4)------------------
	dim l_aux
	l_aux = split(l_opant(4),",")
	dim i
	
	response.write "   for (i = 0; i < document.datos.caudnro.options.length; i++) {" & vbCrLf 
			i = 0
			do while i<= Ubound(l_aux)
				if inStr(l_opant(4),"0,0") <> 0 then
					response.write " if (document.datos.caudnro.options[i].value =='0') { ;" & vbCrLf 
					response.write "	document.datos.caudnro.options[i].selected =true; } " & vbCrLf 
				else
					if 	l_aux(i) <> "0" then
					response.write "       if (document.datos.caudnro.options[i].value =='"&l_aux(i)&"') { ;" & vbCrLf 
					response.write "			document.datos.caudnro.options[i].selected =true; } " & vbCrLf 					
					end if
				end if	
			i = i + 1
			loop
	response.write "   }" & vbCrLf 
	
response.write "  }" & vbCrLf
else response.write "	document.datos.caudnro.options[0].selected =true; } " & vbCrLf 
end if
response.write "};" & vbCrLf
response.write "</script>" & vbCrLf

l_rs1.Close
set l_rs1 = nothing

'----------------------------------------------------------------------------------------------------------------------------

function estaEnLista(conf,lista)
Dim arr
Dim i
Dim salida 

  arr = split(lista,",")
  salida = false
  
  for i = 0 to UBound(arr) 
     if CInt(arr(i)) = CInt(conf) then
	    salida = true
	 end if
  next

  estaEnLista = salida
end function 'estaEnLista(proceso,lista)

%>

<html>
<head>
<link href="/rhprox2/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Filtro - Auditor&iacute;a - RHPro &reg;</title>
</head>
<script src="/rhprox2/shared/js/fn_windows.js"></script>
<script src="/rhprox2/shared/js/fn_confirm.js"></script>
<script src="/rhprox2/shared/js/fn_ayuda.js"></script>
<script src="/rhprox2/shared/js/fn_fechas.js"></script>
<script src="/rhprox2/shared/js/fn_buscar_emp.js"></script>
<script src="/rhprox2/shared/js/fn_help_emp.js"></script>
<script>

var campos = opener.datos.campos.value;

function filtrar()
{
    var tex;
	var tex2;
	var fechadesde;
	var fechahasta;
	var accciones;
	var acnro;
	var usuarios;
	var iduser;
	var confaud;
	var caudnro;
	var empresas;
	var empnro;
	
	var emptipo;
	var empleados;
	
	var orden	 = " auditoria.aud_fec ";
	var ordenpor = "Fecha";
	var ordenado = "Asc";
	var estado;
	var opfiltroant;   // Variable para preservar los valores predeterminados del filtro 
	var titulofiltro;  // Variable para preservar los valores seleccionados del filtro. Estan representados como un string  
	tex = "";
	cant = 0;
	
	tex = tex + " (0=0)"; 
	
	if (document.datos.campos.value == ""){
		alert('Debe Seleccionar los campos.');
		return;}
		
	if (document.datos.orden[1].checked){
		ordenpor = "Usuario";
		orden= " auditoria.iduser ";}
	if (document.datos.ordenado[1].checked){
		orden= orden + " desc ";
		ordenado = "Des";}
	
	if (!validarfecha(document.datos.fechadesde) || !validarfecha(document.datos.fechahasta)){}
	else{
		fechadesde	= document.datos.fechadesde.value;
		fechahasta	= document.datos.fechahasta.value;
		
		if (document.datos.acciones.checked) {
			acciones	= -1;
			}
		else { acciones	= 0;
		       if (document.datos.acnro.value == "0"){
		          alert('Debe Seleccionar la acción.');
 		          return;}
			 }
		acnro		= document.datos.acnro.value;

		if (document.datos.usuarios.checked) {
			usuarios	= -1;
			}
		else { usuarios	= 0;
		       if (document.datos.iduser.value == "0"){
		          alert('Debe Seleccionar el usuario.');
 		          return;}
			  }
		iduser		= document.datos.iduser.value;

		caudnro		= document.datos.vcaudnro.value;
		
		if (document.datos.emptipo[0].checked){
			emptipo = 0;
			document.datos.empleados.value = "0";
			opener.datos.empleados.value = "0";
			}
		if (document.datos.emptipo[1].checked){
			if (document.datos.empleados.value != ""){
				emptipo = 1;
				opener.datos.empleados.value = document.datos.empleados.value;
				}
			else{
				alert("Seleccione el filtro de empleados");
				return;
				}
			}
		if (document.datos.emptipo[2].checked){
			if (document.datos.empleg.value != ""){
				emptipo = 2;
				document.datos.empleados.value = document.datos.empleg.value;
				opener.datos.empleados.value = document.datos.empleg.value;
				}
			else{
				alert("Seleccione el empleado");
				return;
				}
			}
		
		opfiltroant = acciones +"?" 
		opfiltroant = opfiltroant + acnro +"?";
		opfiltroant = opfiltroant + usuarios +"?";
		opfiltroant = opfiltroant + iduser+"?";
		opfiltroant = opfiltroant + caudnro+"?";
		opfiltroant = opfiltroant + document.datos.fechadesde.value+"?";
		opfiltroant = opfiltroant + document.datos.fechahasta.value+"?";
		opfiltroant = opfiltroant + ordenpor+"?"+ordenado+"?";
		opfiltroant = opfiltroant + emptipo;

		opener.opant = opfiltroant;
		opener.titulofiltro = titulofiltro;
		
		opener.datos.campos.value = document.datos.campos.value;

		opener.Filtrar(tex,acciones,acnro,usuarios,iduser,caudnro,fechadesde,fechahasta,orden,emptipo);
		window.close();
		
	}
}

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/rhprox2/shared/js/calendar.html', 16, 15);

 if (jsFecha == null) txt.value = ''
 else txt.value = jsFecha;
}


function Acciones(valor){
	if (valor==true)
		{
		document.datos.acnro.disabled =true;
		document.datos.acnro.value ="0";
		}
	else
		document.datos.acnro.disabled =false;
}
function Usuarios(valor){
	if (valor==true)
		{
		document.datos.iduser.disabled =true;
		document.datos.iduser.value ="0";
		}
	else
		document.datos.iduser.disabled =false;
}
function Confaud(valor){
	if (valor==true)
		{
		document.datos.caudnro.disabled =true;
		document.datos.caudnro.value ="0";
		}
	else
		document.datos.caudnro.disabled =false;
}
function Empresas(valor){
	if (valor==true)
		{
		document.datos.empnro.disabled =true;
		document.datos.empnro.value ="0";
		}
	else
		document.datos.empnro.disabled =false;
}

function habilitar(){
	
	<%if l_acciones="-1" then%>
		document.datos.acnro.disabled =true;
		document.datos.acnro.value ="0";
	<%else%>
		document.datos.acnro.disabled =false;
	<%end if%>
	
	<%if l_usuarios="-1" then%>
		document.datos.iduser.disabled =true;
		document.datos.iduser.value ="0";
	<%else%>
		document.datos.iduser.disabled =false;
	<%end if%>
	
	<%if l_empsel=2 then %>
		document.datos.empleg.value = opener.datos.empleados.value;
        buscar_emp_todos_porLeg(document.datos.empleg.value);
	<%end if%>

	document.datos.caudnro.disabled =false;
}

function actualiza(){
    var lista="0";
	var i;
	
	for (i = 0; i < document.datos.caudnro.options.length; i++) 
		{
			if (document.datos.caudnro.options[i].selected){
				lista = lista + "," + document.datos.caudnro.options[i].value; 
			}			
		} 
	 
    document.datos.vcaudnro.value = lista;
}

function elegirCampos(){

	abrirVentana('rep_auditoria_sup_06.asp?campos='+document.datos.campos.value,'',650,500);

}

CentrarVentana(580,470);
//parent.dialogHeight = 22;
//window.resizeTo(300,180)

function empleadotipo(tipo){
  switch (tipo){
    case 1:
       // Un Empleado 
       document.datos.empleg.className   = "habinp";   
       document.datos.boton.disabled     = 0;      
       document.datos.empleg.disabled    = 0;         
	   document.all.btnseleccion.className = "sidebtnDSB";
	   break;		  		
    case 2:
	   // Filtro	
       document.datos.empleg.className   = "deshabinp";   
       document.datos.boton.disabled     = 1;      
       document.datos.empleg.disabled    = 1;         
	   document.all.btnseleccion.className = "sidebtnSHW";
	   break;		  	
    case 3:
	   // Todos	
       document.datos.empleg.className   = "deshabinp";   
       document.datos.boton.disabled     = 1;      
       document.datos.empleg.disabled    = 1;         
	   document.all.btnseleccion.className = "sidebtnDSB";
	   break;		  
  }
}

function Tecla(num){
  if (num==13) {
        buscar_emp_todos_porLeg(document.datos.empleg.value)
		return false;
  }
  return num;
}

function selectEmpleados(){
   if (document.datos.emptipo[1].checked){
     abrirVentana('../shared/asp/gen_select_emp_v2_00.asp?seltipnro=<%= l_seltipnro%>&srcdatos=opener.datos.empleados','',700,570);
   }
}


function nuevoempleado(ternro,legajo,apellido,nombre){
if (legajo==0){
	//document.datos.ternro.value   = "";
	document.datos.empleg.value   = "";
	document.datos.empleado.value = "";
}else{
	//document.datos.ternro.value   = ternro;
	document.datos.empleg.value   = legajo;
	document.datos.empleado.value = apellido + ", " + nombre;
}
}

</script>

<body onload="javascript:habilitar();CargarConfig(-1);" leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" scroll=no>
<form name="datos" method="post">

<input type="hidden" name="empleados">
<script> document.datos.empleados.value = opener.datos.empleados.value </script>
<input type="hidden" name="campos">
<script> document.datos.campos.value = opener.datos.campos.value </script>

<input type="hidden" name="vcaudnro" value="<%= l_caudnro%>">
<input type="hidden" name="vcampos" value="">

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
  <tr>
    <td class="th2" >Filtrar</td>
	<td class="th2" align="center">
		  <a class=sidebtnHLP href="javascript:elegirCampos();">Campos</a>
	</td>
	<td class="th2" align="right">
		  <a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
  </tr>
<tr>
	<td align="right"><input type="Checkbox" name="acciones" onclick="Acciones(this.checked);" <% If l_acciones ="-1" then %> checked<% End If %> ></td>
	<td align="left" colspan="2" ><b>Todas las Acciones</b></td>
</tr>	

<tr>
	<td align="right"><b>Acciones:</b></td>
	<td colspan="2">
		<select name=acnro style="width=435">
			<option value="0">&laquo;Seleccione una opción&raquo;</option>
			<%Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT acnro, acdesc  "
			l_sql = l_sql & " FROM accion "
			rsOpen l_rs1, cn, l_sql, 0 
			do until l_rs1.eof%>
				<option value="<%=l_rs1("acnro")%>"><%=l_rs1("acdesc")%></option> 
				<script>
					vectoracciones[<%=l_rs1("acnro")%>]='<%=l_rs1("acdesc")%>';
				</script>
			<%l_rs1.MoveNext
			loop
			l_rs1.Close
			set l_rs1 = nothing
			%>
		</select>
		<script>document.datos.acnro.value = '<%=l_acnro%>';</script>
	</td>
</tr>
<tr>
	<td align="right"><input type="Checkbox" name="usuarios" onclick="Usuarios(this.checked);" <% If l_usuarios ="-1" then %> checked<% End If %> ></td>
	<td align="left" colspan="2" ><b>Todos los Usuarios</b></td>
</tr>	
<tr>
	<td align="right"><b>Usuarios:</b></td>
	<td colspan="2">
		<select name=iduser style="width=435">
			<option value="0">&laquo;Seleccione una opción&raquo;</option>
			<%Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
			l_sql = "SELECT iduser, usrnombre  "
			l_sql = l_sql & " FROM user_per "
			rsOpen l_rs1, cn, l_sql, 0 
			do until l_rs1.eof%>
				<option value="<%=l_rs1("iduser")%>"><%=l_rs1("iduser")%>&nbsp;-&nbsp;<%=l_rs1("usrnombre")%></option> 
				<script>
					vectorusuarios['<%=l_rs1("iduser")%>']='<%=l_rs1("iduser")%>&nbsp;<%=l_rs1("usrnombre")%>';
				</script>
			<%l_rs1.MoveNext
			loop
			l_rs1.Close
			set l_rs1 = nothing
			%>
		</select>
		<script>document.datos.iduser.value = '<%=l_iduser%>';</script>
	</td>
</tr>
<tr>
	<td align="right"><b>Configuraciones:</b></td>
	<td colspan="2">
		<select multiple name="caudnro" style="width=435" onChange="javascript:actualiza();"></select>
	</td>
</tr>
<tr>
    <td align="right"><b>Fecha Desde:</b></td>
	<td colspan="2"><input type="Text" name="fechadesde" size="10" value="<%=l_fechadesde%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde)"><img src="/rhprox2/shared/images/cal.gif" border="0"></a>
	</td>
</tr>
<tr>
    <td align="right"><b>Fecha Hasta:</b></td>
	<td colspan="2"><input type="Text" name="fechahasta" size="10" value="<%=l_fechahasta%>">
	<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta)"><img src="/rhprox2/shared/images/cal.gif" border="0"></a>
	</td>
</tr>
<tr>
    <td align="right" rowspan="2"><b>Ordenar:</b></td>
	<td>
		<input type="radio" name="orden"  value="E" <%if l_orden = "Fecha" then%>Checked<%end if%>>
		Por&nbsp;fecha
	</td>														   
	<td>
		<input type="radio" name="ordenado"  value="E" <%if l_ordenado = "Asc" then%>Checked<%end if%>>
		Ascendente
	</td>														   
</tr>
<tr>
	<td>
		<input type="radio" name="orden"  value="E" <%if l_orden = "Usuario" then%>Checked<%end if%>>
		Por&nbsp;Usuario 
	</td>														   
	<td>
		<input type="radio" name="ordenado"  value="E" <%if l_ordenado = "Des" then%>Checked<%end if%>>
		Descendente 
	</td>														   
</tr> 
	<tr>
		<td valign="top" align="right" rowspan="3"><b>Empleados:</b></td>
		<td valign="top" align="left" colspan="2">
			<input type="radio" name="emptipo"  value="3" onclick="Javascript:empleadotipo(3);" 
 		    <% if l_empsel=0 then %>			
			   checked
			<% end if%>
			>
			Todos 
		</td>										
	</tr>
	<tr>
		<td valign="top" align="left">
			<input type="radio" name="emptipo"  value="2"  onclick="Javascript:empleadotipo(2);"
 		    <% if l_empsel=1 then %>			
			   checked
			<% end if%>
			>
			Filtro
		</td>
		<td align="left">
		 <a class="<% If l_empsel=1 then %>sidebtnSHW<% Else  %>sidebtnDSB<% End If %>" name="btnseleccion" onclick="Javascript:selectEmpleados();" href="#" >Selecci&oacute;n</a>
		</td>
	</tr>
	<tr>
		<td valign="top" align="left">
		    <input type="radio" name="emptipo" value="1" onclick="Javascript:empleadotipo(1);"
 		    <% if l_empsel=2 then %>			
			   checked
			<% end if%>
			>
		    Un Empleado
		</td>
		<td align="left" valign="center" nowrap>
				<input type="text" name="empleg" size="10" <% If l_empsel<>2 then %>class="deshabinp" <% Else  %>value="<%= l_seleccion%>" <% End If %> maxlength="10" onKeyPress="return Tecla(event.keyCode);" onchange="javascript:buscar_emp_todos_porLeg(this.value);">
				<input type="Button" name="boton" value="?" onclick="JavaScript:help_emp_todos();">
				<input type="Text" name="empleado" size="35" style="width=225" readonly class="deshabinp" value="">
		</td>
	</tr>
<tr>
    <td align="right" class="th2" colspan="3">
		<a class=sidebtnABM href="Javascript:if (filtrar()) {window.close();}">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>		
	</td>
</tr>
</form>
</body>
</html>
