<% Option Explicit %>
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->

<script>
function Salvar(cadena){
     abrirVentanaH('autorizacion_gti_06.asp?cystipnro='+document.all.cystipnro.value + '&grabar=' + cadena,150,150);
}
</script>

<%' ----------------------------------------------------------------------------------
Sub DobleBrowse(l_sql, l_sql2 )

dim l_rs
Dim l_texto
Dim l_signo
Dim l_pje

Dim l_indice
Dim l_cadenafiltro
Dim l_indiceder
Dim l_indiceizq

response.write "<script languaje='javascript'>" & vbCrLf

response.write "function CargarArreglo() {" & vbCrLf

Set l_rs = Server.CreateObject("ADODB.RecordSet")
rsOpen l_rs, cn, l_sql, 0 
l_indice = 0
l_indiceder = 0
l_indiceizq = 0

do until l_rs.eof
	l_texto =  l_rs(1)
    response.write "newOp = new Option();" & vbCrLf
	response.write "newOp.value ='" & l_rs(0) & "';" & vbCrLf
    response.write "newOp.text = '" & l_texto & "';" & vbCrLf
	
    response.write "nselfil.options[" & l_indiceizq & "] = newOp;" & vbCrLf
	l_indiceizq = l_indiceizq + 1.		

	l_indice = l_indice + 1
	l_rs.MoveNext
loop
l_rs.Close

rsOpen l_rs, cn, l_sql2, 0 

l_indice = 0
l_indiceder = 0

do until l_rs.eof
' muestra codigo hora,descripcion hora, signo y porcentaje
	
	l_texto =  l_rs(1) 
    response.write "newOp = new Option();" & vbCrLf
	response.write "newOp.value = '"  & l_rs(0) & "';" & vbCrLf
    response.write "newOp.text  = '" & l_texto & "';" & vbCrLf
	
    response.write "selfil.options[" & l_indiceder & "] = newOp;"  & vbCrLf
	l_indiceder = l_indiceder + 1.		
    l_indice = l_indice + 1
	on error resume next
	l_rs.MoveNext
loop


response.write "}" & vbCrLf


response.write "function Uno(fuente,destino){" & vbCrLf &_
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione un Usuario'); }" & vbCrLf &_
		   "    else {" & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value= fuente[fuente.selectedIndex].value;" & vbCrLf &_
    	   " // ponerle los valores de los arrays -----------------------------" & vbCrLf &_
    	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf

response.write "function UnoDerecha(fuente,destino){" & vbCrLf &_
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione un Usuario'); }" & vbCrLf &_
		   "    else {" & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value= fuente[fuente.selectedIndex].value;" & vbCrLf &_
    	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf


response.write "function Aceptar(){" & vbCrLf &_
    	   "    var cadena = ',';" & vbCrLf &_
    	   "    var i;" & vbCrLf &_
    	   "    var long = selfil.length-1;" & vbCrLf &_
    	   "    for (i=0;i<=long;i++){" & vbCrLf &_
    	   "        cadena = cadena + selfil[i].value + ','   ;" & vbCrLf &_
    	   "    }" & vbCrLf &_
    	   "    Salvar(cadena);" & vbCrLf &_
    	   "    window.close();" & vbCrLf &_
    	   "}" & vbCrLf



response.write "</script>" & vbCrLf

l_rs.Close
cn.Close


End Sub
' ----------------------------------------------------------------------------------

dim l_obj
dim l_cystipnro
dim l_cystipdesabr

dim l_rs
dim l_sql
dim l_sql2
dim l_seleccion

l_cystipnro = Request.QueryString("cystipnro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT cystipnro, cystipnombre "
l_sql = l_sql & "FROM cystipo "
l_sql = l_sql & "WHERE cystipnro = " & l_cystipnro
l_rs.MaxRecords = 1
rsOpen l_rs, cn, l_sql, 0 
if l_rs.EOF then
	l_cystipdesabr = ""
else	
	l_cystipdesabr = l_rs("cystipnombre")
end if
l_rs.Close

'--------------------- Doble browse -------------------------------------


' se envia el select para armar el lado DERECHO   del browse ---------------
l_sql2 = "SELECT  iduser,usrnombre FROM cysfincirc INNER JOIN user_per "
l_sql2 = l_sql2 & " ON cysfincirc.userid = user_per.iduser "
l_sql2 = l_sql2 & " WHERE cysfincirc.cystipnro = " & l_cystipnro
' se envia el select para armar el lado IZQUIERDO del browse --------------
l_sql = "SELECT  iduser,usrnombre "
l_sql = l_sql & " FROM user_per WHERE NOT EXISTS "
l_sql = l_sql & " (SELECT * " 
l_sql = l_sql & " FROM cysfincirc WHERE cysfincirc.cystipnro = " & l_cystipnro & " AND "
l_sql = l_sql & " cysfincirc.userid = user_per.iduser)"

DobleBrowse l_sql, l_sql2 

%>
<html>
<head>
<link href="/turnos/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Fin de Firmas</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
</head>

<table border="0" cellpadding="0" cellspacing="0">
<td colspan="2" align="left" class="barra">Autorizaci&oacute;n:<%=l_cystipnro%>&nbsp;<%=l_cystipdesabr%></td>
<form name=datos>
<input type=hidden name=cystipnro value=<%=l_cystipnro%>>
<input type=hidden name=cadena>
</form>
</tr>
</table>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:CargarArreglo();" >

<input type=hidden name=radio>
<input type=hidden name=fi-grupo value="">
<table width=100% border=0>
<tr>
    <td><b>Usuarios No Seleccionados</b><br>
    <select class="doblebrowse" size=20 name=nselfil onDblClick="javascript:UnoDerecha(nselfil,selfil);"></select> 
    </td>
	<td align=left width=40>
<!--	<a class=sidebtnSHW href="javascript:Todos(nselfil,selfil);">>></a> -->
	<a class=sidebtnSHW href="javascript:UnoDerecha(nselfil,selfil);">></a></a>
	<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil);"><</a></a></a>
<!--		<a class=sidebtnSHW href="javascript:Todos(selfil,nselfil);"><<</a> -->
    </td>
        <td><b>Usuarios Seleccionados</b><br>
    <select class="doblebrowse" size=20 name=selfil onDblClick="javascript:Uno(selfil,nselfil);"></select>
    </td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
		<a class=sidebtnABM href="javascript:Aceptar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>
</html>
