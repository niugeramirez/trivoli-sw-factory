<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<%

on error goto 0

'Archivo: permisos_seg_00.asp
'Descripción: Permisos para el usuario
'Autor : Alvaro Bayon
'Fecha: 07/03/2005
 
' c_IndiceLado: Indica en que columna del arreglo Datos esta el indicador de seleccionado o no
' La primer columna es la 0

' c_Clave1, c_Clave2: Indican los limites inf. sup. de los campos que seran claves
' c_Campos1, c_Campos2: Indican los limites inf. sup. de los campos que se mostraran en los select
 Const c_IndiceLado = 2
 Const c_Clave1 = 0
 Const c_Clave2 = 0
 Const c_Campos1 = 0
 Const c_Campos2 = 1 
' Declaracion de variables
' Filtro
 Dim l_Etiquetas  ' Son los nombres que deben aparecer en la ventana para que el usuario seleccione
 Dim l_Campos     ' Son los campos de la base que apareceran en la clausula where, que deben estar asociados a las etiquetas
 Dim l_Tipos      ' Son los tipos de datos que tienen los campos (N=Numerico, T=Texto y F=Fecha)
 
' Orden
 Dim l_EtiquetasOr
 Dim l_FuncionesOr
 Dim l_FuncionesOrSql
 
 ' Filtro
 ' Se detallan los campos como subíndices del arreglo datos
 l_etiquetas = "C&oacute;digo:;Descripción:"
 l_Campos    = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1]"
 l_Tipos     = "N;T"
 
 ' Orden
 l_EtiquetasOr = "C&oacute;digo:;Descripción:"
 l_FuncionesOr = "Menor_Codigo;Menor_Descripcion"
 
 ' ADO
 Dim l_rs
 Dim l_sql
 
 ' Locales
 Dim l_indice
 Dim l_existe
 Dim l_iduser
 Dim l_usrnombre
 Dim l_tconnro
 Dim l_perfil
 Dim l_autorizado
 Dim i
 Dim l_arreglo
 
 l_iduser	 = Trim(Request.QueryString("iduser"))
 
'------------------------------------------------------------------------------------------------------------------------------------
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 l_sql = "SELECT iduser, usrnombre, perfnom"
 l_sql = l_sql & " FROM user_per"
 l_sql = l_sql & " INNER JOIN perf_usr ON user_per.perfnro = perf_usr.perfnro "
 l_sql = l_sql & " WHERE iduser = '" & l_iduser & "'"
 
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
	l_usrnombre = l_rs("usrnombre")
	l_perfil = l_rs("perfnom")
 end if
 l_rs.close

'------------------------------------------------------------------------------------------------------------------------------------
 
 ' SQL con todos los datos. Utilizar LEFT JOIN o una manera de poder distinguir cuales estan seleccionados de los que no.
 ' Un elemento debe ser una de las llaves en la tabla relación
 ' Su existencia o no indica el registro aparece en la lista izquierda o derecha
 l_sql =  "SELECT menumstr.menunro, menumstr.menuname, menumstr.menuaccess, tkt_usu_men_pla.menunro"
 l_sql = l_sql & " FROM menumstr LEFT JOIN tkt_usu_men_pla ON tkt_usu_men_pla.menunro = menumstr.menunro "
 l_sql = l_sql & " AND tkt_usu_men_pla.iduser = '" & l_iduser & "'"

 'La primera vez se tiene en cuenta el orden definido en la consulta
 l_sql = l_sql & " ORDER BY menumstr.menunro"
 
 rsOpen l_rs, cn, l_sql, 0
 l_indice = 0
 
 response.write "<script language='Javascript'>" & vbCrLf
 response.write "datos = new Array();" & vbCrLf
 response.write "var jsIndiceLado = " & c_IndiceLado & ";" & vbCrLf 
 response.write "var jsClave1 = " & c_Clave1 & ";" & vbCrLf 
 response.write "var jsClave2 = " & c_Clave2 & ";" & vbCrLf 
 response.write "var jsCampos1 = " & c_Campos1 & ";" & vbCrLf 
 response.write "var jsCampos2 = " & c_Campos2 & ";" & vbCrLf 
 
 do until l_rs.eof
		
		'Si ya fue seleccionada entonces la paso a la derecha
	 	if l_rs(3) <> "" then
			l_existe = "true"
		else
			l_existe = "false"
		end if
	 	response.write "datos[" & l_indice & "]=[" & l_rs(0) & " , '" & l_rs(1) & "', " & l_existe & "];" & vbCrLf
		' Agregar en datos todos los campos que se muestran en el doble browse
		l_indice = l_indice + 1
	l_rs.MoveNext
 loop
 response.write "</script>"
 l_rs.close
 set l_rs = nothing
 
%>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_doblebrowse.js"></script>

<script>
var jsOrdenDerecho 		= "Menor_Codigo";
var jsOrdenIzquierdo 	= "Menor_Codigo";
var jsAscenDerecho 		= true;
var jsAscenIzquierdo 	= true;

function Menor_Codigo(a,b){
	return (a[0] < b[0]);
}

function Menor_Descripcion(a,b){
	return (a[1] < b[1]);
}



function Aceptar(){
	var cadena = ',';
	var i;
	var long = selfil.length-1;
	for (i=0;i<=long;i++){
	    cadena = cadena + selfil[i].value + ','   ;
	}
	//alert(cadena)
	abrirVentanaH('permisos_seg_01.asp?iduser='+document.datos.iduser.value + '&grabar=' + cadena, '','','');
	window.close();
}

</script>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Selecci&oacute;n de Permisos para un Usuario</title>
</head>

<form name="datos">
<input type=hidden name=iduser value=<%=l_iduser%>>
</form>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:Cargar();" >
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
 <tr>
    <td colspan="2" class="th2">
		<script>document.write(document.title);</script>
	</td>
	<td align="right" class="barra" rowspan="2">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
 </tr>
 <tr style="border-color :CadetBlue;">
	<td colspan="2" align="left" colspan="2" class="barra">
		Usuario:&nbsp;<%=l_usrnombre%>
	</td>
 </tr>
 <tr>
	<td colspan="3" align="center">
		<b>Menúes</b>
	</td>
 </tr>
 <tr>
	<td align="center">
		<a class=sidebtnSHW href="javascript:abrirVentana('filtro_doblebrowse.asp?campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&lado=false','',250,160);">Filtro</a>
		<a class=sidebtnSHW href="javascript:abrirVentana('orden_doblebrowse.asp?Etiquetas=<%=l_EtiquetasOr%>&Funciones=<%= l_FuncionesOr%>&lado=false','',250,160);;">Orden</a>
	</td>
	<td>&nbsp;</td>
	<td align="center">
		<a class=sidebtnSHW href="javascript:abrirVentana('filtro_doblebrowse.asp?campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&lado=true','',250,160);">Filtro</a>
		<a class=sidebtnSHW href="javascript:abrirVentana('orden_doblebrowse.asp?Etiquetas=<%=l_EtiquetasOr%>&Funciones=<%= l_FuncionesOr%>&lado=true','',250,160);;">Orden</a>
	</td>
 </tr>
 <tr>
	<td align=center><b>No Seleccionados</b><br><div align="right">
		Visibles:&nbsp;
		<input type="Text" size="6" name="nfiltro" class="hidden" value=0 readonly><br>
		Total:&nbsp;
		<input type="Text" size="6" name="ntotal" class="hidden" value=0 readonly></div>
    	<select class="doblebrowse" style="width:270px" size=23 name=nselfil ondblclick="Uno(nselfil,selfil, true);" multiple></select>
    </td>
    <td align=center width=40>
	    <a class=sidebtnSHW href="javascript:Todos(nselfil,selfil, true);">>></a>
		<a class=sidebtnSHW href="javascript:Uno(nselfil,selfil, true);">></a>
		<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil, false);"><</a>
		<a class=sidebtnSHW href="javascript:Todos(selfil,nselfil, false);"><<</a>
    </td>
    <td align=center><b>Seleccionados</b><br><div align="right">
		Visibles:&nbsp;
		<input type="Text" size="6" name="filtro" class="hidden" value=0 readonly><br>
		Total:&nbsp;
		<input type="Text" size="6" name="total" class="hidden" value=0 readonly></div>		
	    <select class="doblebrowse" size=23 style="width:270px" name=selfil ondblclick="Uno(selfil,nselfil, false);" multiple></select>
    </td>
 </tr>
 <tr>
	<td align="center">
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(nselfil);">Invertir Selecci&oacute;n</a>
	</td>
	<td>&nbsp;</td>
	<td align="center">
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(selfil);">Invertir Selecci&oacute;n</a>
	</td>
 </tr>



 <tr>
    <td colspan="3" align="right" class="th2">
		<a class=sidebtnABM href="javascript:Aceptar()">Guardar</a>
	</td>
 </tr>
 <iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</table>
</html>
