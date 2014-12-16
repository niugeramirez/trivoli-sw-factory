<% Option Explicit %>
<!--#include virtual="/Ticket/shared/db/conn_db.inc"-->
<!--
Archivo: cupos_con_00.asp 
Descripción: Asignacion de los Cupos que se desea bajar en la planta
Autor : Raul Chinestra	
-->
<%
 
' c_IndiceLado: Indica en que columna del arreglo Datos esta el indicador de seleccionado o no
' c_Clave1, c_Clave2: Indican los limites inf. sup. de los campos que seran claves
' c_Campos1, c_Campos2: Indican los limites inf. sup. de los campos que se mostraran en los select
 Const c_IndiceLado = 1
 Const c_Clave1 = 0
 Const c_Clave2 = 0
 Const c_Campos1 = 0
 Const c_Campos2 = 0
 
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
 l_etiquetas = "C&oacute;digo:;Desc. Abreviada:"
 l_Campos    = "datos[objeto[i].subindice][1];datos[objeto[i].subindice][2]"
 l_Tipos     = "T;T"
 
 ' Orden
 l_EtiquetasOr = "C&oacute;digo:;Desc. Abreviada:"
 l_FuncionesOr = "Menor_Codigo;Menor_Descripcion"
 
 ' ADO
 Dim l_rs
 Dim l_sql
 
 ' Locales
 Dim l_indice
 Dim l_existe
 Dim l_codigo
 Dim l_rs2
 Dim auto
 
 ' Ticket
 Dim l_cupos
 
' l_evenro	 = Request.QueryString("cabnro")
' l_codigo   = Request.QueryString("codigo")

 Set l_rs = Server.CreateObject("ADODB.RecordSet")
 Set l_rs2 = Server.CreateObject("ADODB.RecordSet")
 
'------------------------------------------------------------------------------------------------------------------------------------

 l_sql = "SELECT concup "  
 l_sql = l_sql & " FROM tkt_config "
 rsOpen l_rs, cn, l_sql, 0 
 if not l_rs.eof then
 	if isnull(l_rs("concup")) or l_rs("concup") = "" then
		l_cupos = "'-1'" ' En los casos en que no tenga ningun lugar, se setea en -1
	else
		dim l_zonas		
		dim l_a
		l_zonas = split(l_rs("concup"), ",") 
		for l_a = 0 to UBound(l_zonas)
			l_cupos = l_cupos & "'" & l_zonas(l_a) & "',"
		next
		l_cupos = left(l_cupos, len(l_cupos) - 1 )
	end if
 else
	 l_cupos = "'-1'" ' En los casos en que no tenga ningun lugar, se setea en -1
 end if
 l_rs.close 
 
'------------------------------------------------------------------------------------------------------------------------------------
 
 ' SQL con todos los datos. Utilizar LEFT JOIN o una manera de poder distinguir cuales estan seleccionados de los que no.

 l_sql =  " select distinct(tkt_lugar.lugzon) "
 l_sql = l_sql & " from tkt_lugar "
 l_sql = l_sql & " where tkt_lugar.lugzon <> '' and tkt_lugar.lugzon IN ( " & l_cupos & " ) "
 l_sql = l_sql & " order by tkt_lugar.lugzon "
 'response.write l_sql
' response.end
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
	l_existe = "true"
	response.write "datos[" & l_indice & "]=['" & l_rs(0) & "', " & l_existe & "];" & vbCrLf
	' Agregar en datos todos los campos que se muestran en el doble browse
	l_indice = l_indice + 1
	l_rs.MoveNext
 loop

 l_rs.close
 
' Asigno los lugares del lado izquierdo del Browse 
 l_sql =  " select distinct(tkt_lugar.lugzon) "
 l_sql = l_sql & " from tkt_lugar "
 l_sql = l_sql & " where tkt_lugar.lugzon <> '' and tkt_lugar.lugzon NOT IN ( " & l_cupos & " ) "
 l_sql = l_sql & " order by tkt_lugar.lugzon "
 rsOpen l_rs, cn, l_sql, 0
  do until l_rs.eof
	l_existe = "false"
	response.write "datos[" & l_indice & "]=[ '" & l_rs(0) & "', " & l_existe & "];" & vbCrLf
	' Agregar en datos todos los campos que se muestran en el doble browse
	l_indice = l_indice + 1
	l_rs.MoveNext
 loop
 
 response.write "</script>"
 l_rs.close
 set l_rs = nothing
 
%>
<script src="/ticket/shared/js/fn_windows.js"></script>
<script src="/ticket/shared/js/fn_confirm.js"></script>
<script src="/ticket/shared/js/fn_ayuda.js"></script>
<script src="/ticket/shared/js/fn_doblebrowse.js"></script>
<script>
var jsOrdenDerecho 		= "Menor_Codigo";
var jsOrdenIzquierdo 	= "Menor_Codigo";
var jsAscenDerecho 		= true;
var jsAscenIzquierdo 	= true;

function Menor_Codigo(a,b){
	return (a[1] < b[1]);
}

function Menor_Descripcion(a,b){
	return (a[2] < b[2]);
}

function Aceptar(){
	var cadena = '';
	var i;
	var long = selfil.length-1;
	for (i=0;i<=long;i++){
	    cadena = cadena + selfil[i].value + "," ;
	}
	cadena = cadena.substr(0, cadena.length-1)
    abrirVentanaH('zonas_con_01.asp?grabar=' + cadena, '','','');	 
}

</script>
<html>
<head>
<link href="/ticket/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Selecci&oacute;n de Zonas Comerciales </title>
</head>

<form name="datos">
<input type=hidden name=evenro value=<%=l_evenro%>>
</form>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:Cargar();" >
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
 <tr>
    <td colspan="3" class="th2">
        <table border="0" cellpadding="0" cellspacing="0" width="100%" height="25">
            <tr>
              	<td align="left" class="barra">&nbsp;
            	</td>
            	<td align="right" class="barra">
            		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
            	</td>
            </tr>
        </table>
	</td>
 </tr>
 <tr style="border-color :CadetBlue;">
 </tr>
 <tr>
	<td colspan="3" align="center">
		<b>Zonas Comerciales</b>
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
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(nselfil);">Invertir Selecci&oacute;n</a>
	</td>
	<td>&nbsp;</td>
	<td align="center">
		<a class=sidebtnSHW href="javascript:InvertirSeleccion(selfil);">Invertir Selecci&oacute;n</a>
	</td>
 </tr>
 <tr>
    <td colspan="3" align="right" class="th2" valign="middle"  height="25">
       <iframe name="valida" style="visivility=hidden;" src="" width="0" height="0"></iframe> 
     		<a class=sidebtnABM href="javascript:Aceptar()">Aceptar</a>
    		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
 </tr>

</table>
</html>
