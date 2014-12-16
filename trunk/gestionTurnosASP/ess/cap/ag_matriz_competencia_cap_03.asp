<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->

<!--
Archivo:ag_matriz_competencia_cap_03.asp
Descripción: 
Autor : Raul Chinestra

-->
<%
'on error goto 0

' c_IndiceLado: Indica en que columna del arreglo Datos esta el indicador de seleccionado o no
' La primer columna es la 0

' c_Clave1, c_Clave2: Indican los limites inf. sup. de los campos que seran claves
' c_Campos1, c_Campos2: Indican los limites inf. sup. de los campos que se mostraran en los select
' c_Par1, c_Par2: Indican los limites inf. sup. de los atributos de la relación
' c_separador: el separador entre campos que se verá
 Const c_IndiceLado = 2
 Const c_Clave1 = 0
 Const c_Clave2 = 0
 Const c_Campos1 = 0
 Const c_Campos2 = 2
 
 Const c_Par1 = 2
 Const c_Par2 = 2
 Const c_separador = " - "

' Declaracion de variables
' Filtro
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
 Dim l_EtiquetasOrDer
 Dim l_EtiquetasOrIz
 Dim l_FuncionesOrDer
 Dim l_FuncionesOrIz
 Dim l_FuncionesOrSql
 
 ' Filtro
 ' Se detallan los campos como subíndices del arreglo datos
 l_etiquetas = "C&oacute;digo:;Descripción:;Porcentaje:"
 l_Campos    = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1];datos[objeto[i].subindice][2]"
 l_Tipos     = "N;T;N"
 
 ' Orden
 l_EtiquetasOrIz = "C&oacute;digo:;Descripción::"
 l_FuncionesOrIz = "Menor_Codigo;Menor_p1"
 
 l_EtiquetasOrDer = "C&oacute;digo:;Descripción:;Porcentaje:"
 l_FuncionesOrDer = "Menor_Codigo;Menor_Descripcion;Menor_p1;Menor_p2"


 ' ADO
 Dim l_rs
 Dim l_rs2
 Dim l_rs3
 Dim l_sql
 Dim l_sql2
 Dim l_sql3

 
 ' Locales
 Dim l_indice
 Dim l_indice1
 Dim l_indice2
 Dim l_indice3

 Dim l_existe
 Dim l_modnro
 Dim l_moddesabr
 Dim l_porc
 Dim l_total
 Dim l_evento
 Dim l_curso
 Dim l_codigo
 Dim l_eveorigen

 Dim l_concabr

 Dim l_pasnro
 Dim txtCadenaNula
 
 Dim l_ternro
 Dim l_puesto
 Dim l_empleado
 Dim l_participante

 l_ternro	  = Request.QueryString("ternro")
 l_puesto     = Request.QueryString("puesto")
 
 l_puesto = 1224
 'l_participante = Request.QueryString("parnro")
 
 'response.end
 
'------------------------------------------------------------------------------------------------------------------------------------
dim l_rs0
dim l_sql0
Set l_rs0 = Server.CreateObject("ADODB.RecordSet")
l_sql0 = "SELECT terape, terape2, ternom, ternom2, empleg "  
l_sql0 = l_sql0 & " FROM empleado "
l_sql0 = l_sql0 & " WHERE ternro = " & l_ternro
rsOpen l_rs0, cn, l_sql0, 0 
if not l_rs0.eof then
 	l_empleado = l_rs0("empleg") & " - " & l_rs0("terape") & " " & l_rs0("terape2") & " ," & l_rs0("ternom") & " " & l_rs0("ternom2")
end if 
l_rs0.close
set l_rs0 = Nothing
'------------------------------------------------------------------------------------------------------------------------------------
 
  ' SQL con todos los datos. Utilizar LEFT JOIN o una manera de poder distinguir cuales estan seleccionados de los que no.
 ' Un elemento debe ser una de las llaves en la tabla relación
 ' Su existencia o no indica el registro aparece en la lista izquierda o derecha

'==== Realizo las consultas ====

'---- Todas las competencias-----------------------------

l_sql = l_sql & " SELECT evafactor.evafacnro, evafacdesabr, evafacdesext"
l_sql = l_sql & " FROM evafactor "

Set l_rs = Server.CreateObject("ADODB.RecordSet")

'----- Competencias segun el puesto -----------------------------------------------------

l_sql2 = l_sql2 & " SELECT evafactor.evafacnro, evafacdesabr, evafacdesext, evafactor.evafacnro " 
l_sql2 = l_sql2 & " FROM evadescomp "
l_sql2 = l_sql2 & " INNER JOIN evafactor ON evafactor.evafacnro = evadescomp.evafacnro "
l_sql2 = l_sql2 & " WHERE evadescomp.tenro = 4  AND  evadescomp.estrnro =" & l_puesto

Set l_rs2 = Server.CreateObject("ADODB.RecordSet")

 response.write "<script language='Javascript'>" & vbCrLf
 response.write "datos = new Array();" & vbCrLf
 response.write "datos1 = new Array();" & vbCrLf
 response.write "datos2 = new Array();" & vbCrLf
 response.write "datos3 = new Array();" & vbCrLf

 response.write "var jsIndiceLado = " & c_IndiceLado & ";" & vbCrLf 
 response.write "var jsClave1 = " & c_Clave1 & ";" & vbCrLf 
 response.write "var jsClave2 = " & c_Clave2 & ";" & vbCrLf 
 response.write "var jsCampos1 = " & c_Campos1 & ";" & vbCrLf 
 response.write "var jsCampos2 = " & c_Campos2 & ";" & vbCrLf 
 response.write "var jsPar1 = " & c_Par1 & ";" & vbCrLf 
 response.write "var jsPar2 = " & c_Par2 & ";" & vbCrLf 
 response.write "var jsSeparador = '" & c_separador & "';" & vbCrLf 

 txtCadenaNula = hacerCadenaNula()
 rsOpen l_rs, cn, l_sql, 0
 l_indice = 0
 response.write "function Objetivo(){" & vbCrLf
 response.write "datos = datos1;" & vbCrLf
 response.write "}" & vbCrLf
 
 do until l_rs.eof
 	'if l_rs(c_IndiceLado) <> "" then
	'	l_existe = "true"
	'else
		l_existe = "false"
	'end if

 	response.write "datos1[" & l_indice & "]=[" & l_rs(0) & " , '" & l_rs(1) & "'," & 0 & ", " & l_existe & "];" & vbCrLf
	' Agregar en datos todos los campos que se muestran en el doble browse
    l_indice = l_indice + 1
	l_rs.MoveNext
 loop
response.write "var tope1 = " & l_indice & ";" & vbCrLf

 txtCadenaNula = hacerCadenaNula()
 rsOpen l_rs2, cn, l_sql2, 0
 l_indice = 0
 do until l_rs2.eof
 '	if l_rs2(c_IndiceLado) <> "" then
	'	l_existe = "true"
'	else
		l_existe = "false"
	'end if
 	response.write "datos2[" & l_indice & "]=[" & l_rs2(0) & " , '" & l_rs2(1) & "'," & 0 & ", " & l_existe & "];" & vbCrLf
	'Agregar en datos todos los campos que se muestran en el doble browse
	l_indice = l_indice + 1
	l_rs2.MoveNext
 loop
 response.write "var tope2 = " & l_indice & ";" & vbCrLf

 response.write "</script>"

 l_rs.close
 'l_rs2.close

set l_rs = nothing
set l_rs2 = nothing

 function hacerCadenaNula()
 	Dim i, j
	Dim cadena
	j = c_Par2 - c_Par1
	cadena = ""
	for i = 1 to j
		cadena = cadena & ";"
	next
	hacerCadenaNula = cadena
 end function

%>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_doblebrowse_param.js"></script>
<script>
// Funciones de orden	========================================
var jsOrdenDerecho 		= "Menor_Codigo";
var jsOrdenIzquierdo 	= "Menor_Codigo";
var jsAscenDerecho 		= true;
var jsAscenIzquierdo 	= true;

function Menor_Codigo(a,b){
	return (a[0] < b[0]);
}

function Menor_Descripcion(a,b){
	return (parseInt(a[1]) < parseInt(b[1]));
}

function Menor_p1(a,b){
	return (a[2] < b[2]);
}

function Menor_p2(a,b){
	return (a[3] < b[3]);
}
//============================================================

function VentanaParams(destino)
//	Abre la ventana para cargar los atributos de la relación
{   
	// Si los params son nulos entonces no los paso a la ventana
	//*alert(datos[destino[destino.selectedIndex].subindice][2]);
	if (typeof(datos[destino[destino.selectedIndex].subindice][jsPar1]) == 'undefined'){
		var r = showModalDialog('ag_matriz_competencia_cap_04.asp?radio=' + document.all.radio.value + '&porcen=' + datos[destino[destino.selectedIndex].subindice][2], '','dialogWidth:18;dialogHeight:10');
	}else{
		var r = showModalDialog('ag_matriz_competencia_cap_04.asp?radio=' + document.all.radio.value + '&porcen=' + datos[destino[destino.selectedIndex].subindice][2], '','dialogWidth:18;dialogHeight:10');
		}
	return(r);
}

function UnoParam(nselfil,selfil, lado){
	if (!((nselfil.length == 0) && (lado == true))){
		if (nselfil.selectedIndex == -1){
			alert('Debe Seleccionar un Elemento');
		}else{
		var r = VentanaParams(nselfil);
		if (typeof(r) != 'undefined'){
			//if (!(r==0)) 
			//{ 
				Uno(nselfil,selfil, lado,r);
			//}	
		}
	}
}}

function TodosParam(nselfil,selfil, lado){	
	if (!((nselfil.length == 0) && (lado == true))){
		Selall(nselfil);
		var r = VentanaParams(nselfil);
		if (typeof(r) != 'undefined'){
			if (!(r==0)) 
			{ 
				Todos(nselfil,selfil, lado,r);
			}	
	}}
}
function Selall(objeto){
	for (i=0;i<objeto.length;i++)
		objeto[i].selected = true;
}

function Modificar(destino, lado){
if (!((selfil.length == 0) && (lado == true))){
	var r = VentanaParams(destino);
	if (typeof(r) != 'undefined'){
		if (!(r==0)) 
			ModificarValores(destino, r);
		}
	}
}

function Aceptar(){
	Guarda();
	var cadena1 = ',';
	var cadena2 = ',';
	var cadena3 = ',';
	//0 y 2 son las posiciones del numero de mod y el valor del %
	var i;
	var j;
	for (i=0;i<=tope1-1;i++){
		if (datos1[i][jsIndiceLado]){
		    cadena1 = cadena1 + datos1[i][0] + ',' + datos1[i][2]  + ','
		}
	}
	
	for (i=0;i<=tope2-1;i++){
		if (datos2[i][jsIndiceLado]){
		   cadena2 = cadena2 + datos2[i][0] + ',' + datos2[i][2] + ','
	    }
	}
	
	
	abrirVentanaH('ag_matriz_competencia_cap_05.asp?ternro=<%=l_ternro%>&grabar1=' + cadena1 + '&grabar2=' + cadena2, '50','50');
	//window.close();
}
function Actualiza(Quien){
	switch (Quien){
		case '1':
				document.all.objetivo.checked = true;
				document.all.factor.checked = false;
				document.all.radio.value = 1;
				Guarda();
				document.datos.titulo.value = "Objetivo";
				datos = datos1;
				Cargar();
				filtro();
			break;
		case '2':
				document.all.objetivo.checked = false;
				document.all.factor.checked = true;
				document.all.radio.value = 2;
				Guarda()
				document.datos.titulo.value = "Factor";
				datos = datos2;
				Cargar();
				filtro();
			break;
	}

}

function Guarda(){
	switch (document.datos.titulo.value){
		case 'Objetivo':
				datos1 = datos;
			return;
		case 'Factor':
				datos2 = datos;
			return;
	}
}
</script>
<script language="VBScript">
function filtro()
	Select case document.datos.titulo.value
		case "Objetivo"		
			document.datos.filtro_etiqueta.value = "C&oacute;digo:;Descripción:;Porcentaje:"
			document.datos.filtro_campo.value  = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1];datos[objeto[i].subindice][2]"
			document.datos.filtro_tipo.value     = "N;T;N"
			 
			document.datos.orden_etiquetasIz.value = "C&oacute;digo:;Descripción::"
			document.datos.orden_funcionesIz.value = "Menor_Codigo;Menor_p1"
			document.datos.orden_etiquetasDer.value = "C&oacute;digo:;Descripción:;Porcentaje:"
			document.datos.orden_funcionesDer.value = "Menor_Codigo;Menor_Descripcion;Menor_p1;Menor_p2"
			 
		case "Factor"
			document.datos.filtro_etiqueta.value = "C&oacute;digo:;Descripción:;Porcentaje:"
			document.datos.filtro_campo.value    = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1];datos[objeto[i].subindice][2]"
			document.datos.filtro_tipo.value	 = "N;T;N"
			
			document.datos.orden_etiquetasIz.value = "C&oacute;digo:;Descripción:"
			document.datos.orden_funcionesIz.value = "Menor_Codigo;Menor_p1"
 			document.datos.orden_etiquetasDer.value = "C&oacute;digo:;Descripción:;Porcentaje:"
			document.datos.orden_funcionesDer.value = "Menor_Codigo;Menor_Descripcion;Menor_p1;Menor_p2"
						
	End Select
end function
</script>
<html>
<head>
<link href="../<%= session("estilo")%>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Alta de Competencias</title>
</head>

<form name="datos">
<input type=hidden name=modnro value=<%=l_modnro%>>
<input type=hidden name=concnro value=<%=l_modnro%>>

<input  type="Hidden" name="filtro_etiqueta" value="<%=l_etiquetas%>">
<input  type="Hidden" name="filtro_campo" value="<%= l_campos %>">
<input  type="Hidden" name="filtro_tipo" value="<%= l_tipos %>">

<input  type="Hidden" name="orden_etiquetasIz" value="<%=l_EtiquetasOrIz%>">
<input  type="Hidden" name="orden_funcionesIz" value="<%= l_FuncionesOrIz%>">

<input  type="Hidden" name="orden_etiquetasDer" value="<%=l_EtiquetasOrDer%>">
<input  type="Hidden" name="orden_funcionesDer" value="<%= l_FuncionesOrDer%>">

<input type=hidden name=titulo value="Objetivo">
</form>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:Objetivo();Cargar();">
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr>
		<td class="th2" colspan="2">
			<script>document.write(document.title);</script>
		</td>
		<td align="right" class="barra" >
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	 </tr>
	 <tr>
 		<td colspan="3" align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="0" height="0">
				<tr>
					<td align="right" width="30%"><b>Empleado: </b></td>
					<td align="left" width="0"><input border="0" readonly type="Text" class="deshabinp" style="width:350" value="<%= l_empleado %>" ></td>
				</tr> 
			</table>
		</td>
	</tr>
	<tr>
 		<td colspan="3" >
			<table border="0">
				<tr>
					<td width="200"></td>
					<td width="0" nowrap><b>Todas las competencias:</b><input type="Radio" name="objetivo" checked onclick="Actualiza('1')" ></td>
					<td width="0" nowrap><b>Competencias según el puesto:</b><input type="Radio" name="factor" onclick="Actualiza('2')"></td>
					<td width="200"></td>
					<td width="0" nowrap><input type="Hidden" name="radio" value="1" ></td>
				</tr>
			</table>
		</td>
 	</tr>
	<tr>
		<td align="center">
			<a class=sidebtnSHW href="javascript:abrirVentana('filtro_doblebrowse.asp?campos=' + document.datos.filtro_campo.value +'&tipos=' + document.datos.filtro_tipo.value +'&etiquetas=' + document.datos.filtro_etiqueta.value +'&lado=false','',250,160);">Filtro</a>
			<a class=sidebtnSHW href="javascript:abrirVentana('orden_doblebrowse.asp?Etiquetas=' + document.datos.orden_etiquetasIz.value + '&Funciones=' + document.datos.orden_funcionesIz.value + '&lado=false','',250,160);;">Orden</a>
		</td>
		<td>&nbsp;</td>
		<td align="center">
			<a class=sidebtnSHW href="javascript:abrirVentana('filtro_doblebrowse.asp?campos=' + document.datos.filtro_campo.value +'&tipos=' + document.datos.filtro_tipo.value +'&etiquetas=' + document.datos.filtro_etiqueta.value +'&lado=true','',250,160);">Filtro</a>
			<a class=sidebtnSHW href="javascript:abrirVentana('orden_doblebrowse.asp?Etiquetas=' + document.datos.orden_etiquetasDer.value + '&Funciones=' + document.datos.orden_funcionesDer.value + '&lado=true','',250,160);;">Orden</a>
			<a id="Modi" class=sidebtnSHW href="javascript:Modificar(selfil,true);">Modifica</a>
		</td>
	</tr>
	<tr>
		<td align=center><b>No Seleccionados</b><br><div align="right">
			Visibles:&nbsp;
			<input type="Text" size="6" name="nfiltro" class="hidden" value=0>
			Total:&nbsp;
			<input type="Text" size="6" name="ntotal" class="hidden" value=0></div>
			<select class="doblebrowse" style="width:270px" size=20 name=nselfil ondblclick="UnoParam(nselfil,selfil, true);"></select>
		</td>
		<td align=center width=40>
		<!--	    <a class=sidebtnSHW href="javascript:TodosParam(nselfil,selfil, true);">>></a>-->
			<a class=sidebtnSHW href="javascript:UnoParam(nselfil,selfil, true);">></a>
			<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil, false, '<%=txtCadenaNula%>');"><</a>
		<!--		<a class=sidebtnSHW href="javascript:Todos(selfil,nselfil, false,'<%'=txtCadenaNula%>');"><<</a>-->
		</td>
		<td align=center><b>Seleccionados</b><br><div align="right">
			Visibles:&nbsp;
			<input type="Text" size="6" name="filtro" class="hidden" value=0>
			Total:&nbsp;
			<input type="Text" size="6" name="total" class="hidden" value=0></div>		
		    <select class="doblebrowse" size=20 style="width:270px" name=selfil ondblclick="Uno(selfil,nselfil, false, ';');"></select>
		</td>
	</tr>
	<!-- <tr>
		<td align="center">
			<a class=sidebtnSHW href="javascript:InvertirSeleccion(nselfil);">Invertir Selecci&oacute;n</a>
		</td>
		<td>&nbsp;</td>
		<td align="center">
			<a class=sidebtnSHW href="javascript:InvertirSeleccion(selfil);">Invertir Selecci&oacute;n</a>
		</td>
	 </tr>-->
	<tr>
    	<td colspan="3" align="right" class="th2">
         	<a class=sidebtnABM href="javascript:Aceptar()">Aceptar</a>
			<a class=sidebtnABM href="javascript:window.close()">Cancelar</a>
         	<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 
		</td>
	</tr>
</table>
</html>

