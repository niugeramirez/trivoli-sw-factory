<% Option Explicit %>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<!--
Archivo: ag_especializaciones_cap_02.asp
Descripcion: especializaciones
Autor: Lisandro Moro
Fecha: 29/03/2004
Modificado:
-->
<%
 on error goto 0
 
' c_IndiceLado: Indica en que columna del arreglo Datos esta el indicador de seleccionado o no
' La primer columna es la 0

' c_Clave1, c_Clave2: Indican los limites inf. sup. de los campos que seran claves
' c_Campos1, c_Campos2: Indican los limites inf. sup. de los campos que se mostraran en los select
' c_Par1, c_Par2: Indican los limites inf. sup. de los atributos de la relación
' c_separador: el separador entre campos que se verá
 
 Const c_IndiceLado = 3
 Const c_Clave1 = 0
 Const c_Clave2 = 0
 Const c_Campos1 = 0
 Const c_Campos2 = 2
 
 Const c_Par1 = 2
 Const c_Par2 = 3
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
 l_etiquetas = "C&oacute;digo:;Eespecialización:;Nivel:"
 l_Campos    = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1];datos[objeto[i].subindice][2]"
 l_Tipos     = "N;T;T"
 
 ' Orden
 l_EtiquetasOrIz = "C&oacute;digo:;Especialización::"
 l_FuncionesOrIz = "Menor_Codigo;Menor_p1"
 
 l_EtiquetasOrDer = "C&oacute;digo:;Especialización:;Nivel:"
 l_FuncionesOrDer = "Menor_Codigo;Menor_Descripcion;Menor_p1"

 ' ADO
 Dim l_rs
 Dim l_sql
 
 ' Locales
 Dim l_indice
 Dim l_indice1
 Dim l_indice2
 Dim l_indice3

 Dim l_existe
 Dim l_moddesabr
 Dim l_porc
 Dim l_evenro
 Dim l_total
 Dim l_Cont
 
 Dim l_ternro
 Dim l_empleg
 Dim l_nombre
 
 Dim l_concabr
 
 Dim l_pasnro
 Dim txtCadenaNula
 Dim l_estadoesp
 
 l_ternro	= l_ess_ternro
 l_empleg	= l_ess_empleg

' l_evenro	= Request.QueryString("cabnro")
 
'------------------------------------------------------------------------------------------------------------------------------------
dim l_rs0
dim l_sql0
Set l_rs0 = Server.CreateObject("ADODB.RecordSet")

l_sql0 = "SELECT terape,terape2,ternom,ternom2 FROM empleado"  
l_sql0 = l_sql0 & " WHERE ternro= " & l_ternro
rsOpen l_rs0, cn, l_sql0, 0 

l_nombre = ""
if not l_rs0.eof then
   l_nombre = l_rs0("terape") & " " & l_rs0("terape2") & ", " & l_rs0("ternom") & " " & l_rs0("ternom2")
end if

l_rs0.close

l_sql0 = "SELECT espnro, espdesabr "  
l_sql0 = l_sql0 & " FROM especializacion "
rsOpen l_rs0, cn, l_sql0, 0 

'if not l_rs0.eof then
'end if 
'l_rs0.close
'set l_rs0 = Nothing
'------------------------------------------------------------------------------------------------------------------------------------
 
 ' SQL con todos los datos. Utilizar LEFT JOIN o una manera de poder distinguir cuales estan seleccionados de los que no.
 ' Un elemento debe ser una de las llaves en la tabla relación
 ' Su existencia o no indica el registro aparece en la lista izquierda o derecha

'==== Realizo las consultas ====

 response.write "<script language='Javascript'>" & vbCrLf
 
 response.write "var jsIndiceLado = " & c_IndiceLado & ";" & vbCrLf 
 response.write "var jsClave1 = " & c_Clave1 & ";" & vbCrLf 
 response.write "var jsClave2 = " & c_Clave2 & ";" & vbCrLf 
 response.write "var jsCampos1 = " & c_Campos1 & ";" & vbCrLf 
 response.write "var jsCampos2 = " & c_Campos2 & ";" & vbCrLf 
 response.write "var jsPar1 = " & c_Par1 & ";" & vbCrLf 
 response.write "var jsPar2 = " & c_Par2 & ";" & vbCrLf 
 response.write "var jsSeparador = '" & c_separador & "';" & vbCrLf 

 response.write "datos = new Array();" & vbCrLf

l_rs0.MoveFirst
l_Cont = 1
do until l_rs0.eof
	l_sql =  " SELECT eltoana.eltananro, eltanadesabr, espnivdesabr ,espnivel.espnivnro, especemp.eltananro, especemp.espestrrhh "
	l_sql = l_sql & " FROM eltoana "
	l_sql = l_sql & " LEFT JOIN especemp on especemp.eltananro = eltoana.eltananro "
	l_sql = l_sql & " AND especemp.ternro = " & l_ternro
	l_sql = l_sql & " LEFT JOIN espnivel ON espnivel.espnivnro = especemp.espnivnro" ''''''''''<<<<
	l_sql = l_sql & " WHERE eltoana.espnro = " & l_rs0("espnro")
	Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 	response.write "datos" & l_Cont & " = new Array();" & vbCrLf

	txtCadenaNula = hacerCadenaNula()
 	rsOpen l_rs, cn, l_sql, 0
	l_indice = 0
 	response.write "function Espec" & l_Cont & "(){" & vbCrLf
 	response.write "datos = datos" & l_Cont & ";" & vbCrLf
 	response.write "}" & vbCrLf
	 do until l_rs.eof

 		if not isNull(l_rs("espestrrhh")) then
			if clng(l_rs("espestrrhh")) = -1 then
				l_estadoesp = -1
			else
				l_estadoesp = 0			
			end if
		else 
			l_estadoesp = 0
		end if

	 	if l_rs(c_IndiceLado) <> "" then
			l_existe = "true"
		else
			l_existe = "false"
		end if
		'response.write "datos" & l_Cont & "[" & l_indice & "]=[" & l_rs(0) & " , '" & l_rs(1) & "','" & l_rs(2) & "','" & l_rs(3) & "'," & l_existe & "];" & vbCrLf
		response.write "datos" & l_Cont & "[" & l_indice & "]=[" & l_rs(0) & " , '" & l_rs(1) & "','" & l_rs(2) & "','" & l_rs(3) & "'," & l_estadoesp & "," & l_existe & "];" & vbCrLf
		' Agregar en datos todos los campos que se muestran en el doble browse
		l_indice = l_indice + 1
		l_rs.MoveNext
	 loop
	response.write "var tope" & l_Cont & " = " & l_indice & ";" & vbCrLf
	l_rs0.MoveNext
	l_Cont = l_Cont + 1
Loop
response.write "</script>"

 l_rs.close
 set l_rs = nothing
 
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
	//alert(datos[destino[destino.selectedIndex].subindice][jsPar1]);
	if (typeof(datos[destino[destino.selectedIndex].subindice][jsPar1]) == 'undefined'){
		var r = showModalDialog('ag_especializaciones_cap_04.asp?concnro=' + document.datos.concnro.value + '&titulo=' +document.datos.titulo.value + '&valor='+datos[destino[destino.selectedIndex].subindice][2] + '&valormax='+datos[destino[destino.selectedIndex].subindice][3] + '&valorsum='+datos[destino[destino.selectedIndex].subindice][4], '','dialogWidth:30;dialogHeight:13');
	}else{
		var r = showModalDialog('ag_especializaciones_cap_04.asp?nivel='+datos[destino[destino.selectedIndex].subindice][3] , '','dialogWidth:30;dialogHeight:13');
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
			if (!(r==0)) 
			{ 
				Uno(nselfil,selfil, lado,r);
			}	
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

function Modificar(destino, lado)
{
	var r = VentanaParams(destino);
	if (typeof(r) != 'undefined'){
		if (!(r==0)) 
			ModificarValores(destino, r);
	}
}

function Elegir(valor){
	switch (valor){
	case "0":
		datos = "";
		Cargar();
		break;
	<%	l_rs0.MoveFirst
	l_Cont = 1
	do until l_rs0.eof		%>
	case "<%= l_rs0(0) %>":
		datos = datos<%= l_Cont %>;
		Cargar();
		break;
<%  l_Cont = l_Cont + 1
	l_rs0.MoveNExt
	Loop %>
	}
}

function Aceptar(){
	var i;
	<%	l_rs0.MoveFirst
	l_Cont = 1
	do until l_rs0.eof %>
		var cadena<%= l_Cont %> = ',';
		for (i=0;i<=tope<%= l_Cont %> -1;i++){
			if (datos<%= l_Cont %>[i][jsIndiceLado]){
			    cadena<%= l_Cont %> = cadena<%= l_Cont %> + datos<%= l_Cont %>[i][0] + ',' + datos<%= l_Cont %>[i][3] + ','
			}
		}
<%  l_Cont = l_Cont + 1
	l_rs0.MoveNExt
	Loop %>
// cargo todas las cadenas :-)
<%	l_rs0.MoveFirst
	l_Cont = 1
	do until l_rs0.eof%>
		document.datos.Grabar.value += cadena<%= l_Cont %> + ";"
<%  l_Cont = l_Cont + 1
	l_rs0.MoveNExt
	Loop %>
	
	//alert(document.datos.Grabar.value);
	abrirVentanaH('ag_especializaciones_cap_03.asp?ternro=<%= l_ternro %>&grabar=' + document.datos.Grabar.value, '','');
	//window.close();
}

</script>
<html>
<head>
<link href="../<%= c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Asignaci&oacute;n de Especializaciones</title>
</head>

<form name="datos">
<input type=hidden name=ternro value=<%=l_ternro%>>
<input type=hidden name=empleg value=<%=l_empleg%>>
<input type=hidden name=Grabar value=";">


<input  type="Hidden" name="filtro_etiqueta" value="<%=l_etiquetas%>">
<input  type="Hidden" name="filtro_campo" value="<%= l_campos %>">
<input  type="Hidden" name="filtro_tipo" value="<%= l_tipos %>">

<input  type="Hidden" name="orden_etiquetasIz" value="<%=l_EtiquetasOrIz%>">
<input  type="Hidden" name="orden_funcionesIz" value="<%= l_FuncionesOrIz%>">

<input  type="Hidden" name="orden_etiquetasDer" value="<%=l_EtiquetasOrDer%>">
<input  type="Hidden" name="orden_funcionesDer" value="<%= l_FuncionesOrDer%>">

<input type=hidden name=titulo value="Objetivo">
</form>
<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:Cargar();//Objetivo();" >
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	<tr>
		<td class="th2" colspan="2">
			&nbsp;
		</td>
		<td align="right" class="th2" >
			&nbsp;
		</td>
	 </tr>
	 <tr>
 		<td colspan="3" align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="0" height="0">
				<tr>
					<td align="right" width="30%"><b>Empleado: </b></td>
					<td align="left" width="0">
						<input border="0" readonly type="Text" class="deshabinp" style="width:50" value="<%= l_empleg %>" >
						<input border="0" readonly type="Text" class="deshabinp" style="width:330" value="<%= l_nombre %>" >
					</td>
				</tr> 
				<tr>
					<td align="right"><b>Especialización: </b></td>
					<td align="left"  width="0">
					<select name=espnro style="width:385px" size="1" onchange="Elegir(document.all.espnro.value);">
					<option value=0 selected><< Seleccione una Opción >></option>
					<%	l_rs0.MoveFirst
					do until l_rs0.eof		%>	
						<option value= <%= l_rs0("espnro") %> > 
						<%= l_rs0("espdesabr") %> (<%=l_rs0("espnro")%>) </option>
						<% l_rs0.Movenext
					loop
					l_rs0.Close %>	
					</select>
				
					</td>
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
			<a class=sidebtnSHW href="javascript:Modificar(selfil,true);">Modifica</a>
		</td>
	</tr>
	<tr>
		<td align=center><b>No Seleccionados</b><br><div align="right">
			Visibles:&nbsp;
			<input type="Text" size="6" name="nfiltro" class="hidden" value=0>
			Total:&nbsp;
			<input type="Text" size="6" name="ntotal" class="hidden" value=0></div>
			<select class="doblebrowse" multiple style="width:270px" size=23 name=nselfil ondblclick="UnoParam(nselfil,selfil, true);"></select>
		</td>
		<td align=center width=40>
		    <a class=sidebtnSHW href="javascript:TodosParam(nselfil,selfil, true);">>></a>
			<a class=sidebtnSHW href="javascript:UnoParam(nselfil,selfil, true);">></a>
			<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil, false, '<%=txtCadenaNula%>');"><</a>
			<a class=sidebtnSHW href="javascript:Todos(selfil,nselfil, false,'<%'=txtCadenaNula%>');"><<</a>
		</td>
		<td align=center><b>Seleccionados</b><br><div align="right">
			Visibles:&nbsp;
			<input type="Text" size="6" name="filtro" class="hidden" value=0>
			Total:&nbsp;
			<input type="Text" size="6" name="total" class="hidden" value=0></div>		
		    <select class="doblebrowse" multiple size=23 style="width:270px" name=selfil ondblclick="Uno(selfil,nselfil, false, ';');"></select>
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

