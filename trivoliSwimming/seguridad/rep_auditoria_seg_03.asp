<% Option Explicit%>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<!--
Archivo: rep_auditoria_seg_03.asp
Descripción: Seleccion de campos
Autor : Raul Chinestra
Fecha: 18/07/2006
-->
<%

' c_IndiceLado: Indica en que columna del arreglo Datos esta el indicador de seleccionado o no
' La primer columna es la 0

' c_Clave1, c_Clave2: Indican los limites inf. sup. de los campos que seran claves
' c_Campos1, c_Campos2: Indican los limites inf. sup. de los campos que se mostraran en los select
 Const c_IndiceLado = 3
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
 l_etiquetas = "Código:;Descripción:"
 l_Campos    = "datos[objeto[i].subindice][0];datos[objeto[i].subindice][1]"
 l_Tipos     = "N;T"
 
 ' Orden
 l_EtiquetasOr = "Código:;Descripción:"
 l_FuncionesOr = "Menor_Codigo;Menor_Descripcion"
 
 ' ADO
 Dim l_rs
 Dim l_sql
 
 ' Locales
 Dim l_indice
 Dim l_existe
 Dim l_concnro
 Dim l_concabr
 Dim l_tconnro
 Dim l_pasnro
 dim l_lista
 DIm l_camposT
 
 l_camposT	 = Request.QueryString("campos")
 
'------------------------------------------------------------------------------------------------------------------------------------
 
 l_sql =  "SELECT aud_campnro, aud_campdesabr, aud_camptabla  "
 l_sql = l_sql & " FROM aud_campo "
 l_sql = l_sql & " ORDER BY aud_campo.aud_campnro " 
 
 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 
 rsOpen l_rs, cn, l_sql, 0
 l_indice = 0
 
 response.write "<script language='Javascript'>" & vbCrLf
 response.write "datos = new Array();" & vbCrLf
 response.write "var jsIndiceLado = " & c_IndiceLado & ";" & vbCrLf 
 response.write "var jsClave1 = " & c_Clave1 & ";" & vbCrLf 
 response.write "var jsClave2 = " & c_Clave2 & ";" & vbCrLf 
 response.write "var jsCampos1 = " & c_Campos1 & ";" & vbCrLf 
 response.write "var jsCampos2 = " & c_Campos2 & ";" & vbCrLf 
 
 l_lista = "," & l_camposT & ","
 
 do until l_rs.eof
 	if inStr(l_lista,"," & l_rs("aud_campnro") & ",") > 0 then
		l_existe = "true"
	else
		l_existe = "false"
	end if
 	response.write "datos[" & l_indice & "]=[" & l_rs(0) & " , '" & l_rs(1) & "', '" & l_rs(2) & "', " & l_existe & "];" & vbCrLf
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
	var cadena = '';
	var i;
	var long = selfil.length-1;
	for (i=0;i<=long;i++){
	    if (i < long)
		cadena = cadena + selfil[i].value + ','   ;
		else cadena = cadena + selfil[i].value
	}
	/*
	if (cadena == ''){
	   alert('Debe seleccionar al menos un campo.');
	   return;
	   }
   */
	opener.datos.campos.value = cadena;
	window.close();
}

</script>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Selección de Campos</title>
</head>

<form name="datos">
<input type=hidden name=concnro value=<%=l_concnro%>>
</form>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:Cargar();" >
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
 <tr>
    <td colspan="2" class="th2">
		<script>document.write(document.title);</script>
	</td>
	<td align="right" class="barra" >
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
 </tr>
  <tr>
	<td colspan="3" align="center">
		<b>Selección de Campos</b>
	</td>
 </tr>
 <tr>
	<td align="center">
		<a class=sidebtnSHW href="javascript:abrirVentana('../shared/asp/filtro_doblebrowse.asp?campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&lado=false','',250,160);">Filtro</a>
		<a class=sidebtnSHW href="javascript:abrirVentana('../shared/asp/orden_doblebrowse.asp?Etiquetas=<%=l_EtiquetasOr%>&Funciones=<%= l_FuncionesOr%>&lado=false','',250,160);;">Orden</a>
	</td>
	<td>&nbsp;</td>
	<td align="center">
		<a class=sidebtnSHW href="javascript:abrirVentana('../shared/asp/filtro_doblebrowse.asp?campos=<%= l_campos%>&tipos=<%=l_tipos%>&etiquetas=<%=l_etiquetas%>&lado=true','',250,160);">Filtro</a>
		<a class=sidebtnSHW href="javascript:abrirVentana('../shared/asp/orden_doblebrowse.asp?Etiquetas=<%=l_EtiquetasOr%>&Funciones=<%= l_FuncionesOr%>&lado=true','',250,160);;">Orden</a>
	</td>
 </tr>
 <tr>
	<td align=center><b>No Seleccionados</b><br><div align="right">
		Visibles:&nbsp;
		<input type="Text" size="6" name="nfiltro" class="hidden" value=0 readonly=""><br>
		Total:&nbsp;
		<input type="Text" size="6" name="ntotal" class="hidden" value=0 readonly=""></div>
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
		<input type="Text" size="6" name="filtro" class="hidden" value=0 readonly=""><br>
		Total:&nbsp;
		<input type="Text" size="6" name="total" class="hidden" value=0 readonly=""></div>		
	    <select class="doblebrowse" size=23 style="width:270px" name=selfil ondblclick="Uno(selfil,nselfil, false);" multiple></select>
    </td>
 </tr>
 <tr>
	<td align="center">
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(nselfil);">Invertir Selección</a>
	</td>
	<td>&nbsp;</td>
	<td align="center">
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(selfil);">Invertir Selección</a>
	</td>
 </tr>



 <tr>
    <td colspan="3" align="right" class="th2">
		<a class=sidebtnABM href="javascript:Aceptar()">Aceptar</a>
	</td>
 </tr>
 <iframe name="valida" style="visibility=hidden;" src="" width="100%" height="100%"></iframe> 
</table>
</html>
