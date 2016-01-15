<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<script>

function Salvar(obj){
  
   window.open('usuarios_seg_06.asp?grabar=' + obj.value+'&userid='+userid.value, '','','');
}

</script>

<%' ----------------------------------------------------------------------------------

Sub DobleBrowse(sql, sql2)

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
l_rs.Open l_sql, cn
l_indice = 0
l_indiceder = 0
l_indiceizq = 0

do until l_rs.eof
	l_texto =  l_rs(0) & " " & l_rs(1)
    response.write "obj = " & l_obj & ";" & vbCrLf
    response.write "newOp = new Option();" & vbCrLf
	response.write "newOp.value = " & l_rs(0) & ";" & vbCrLf
    response.write "newOp.text = '" & l_texto & "';" & vbCrLf
	
    response.write "nselfil.options[" & l_indiceizq & "] = newOp;" & vbCrLf
	l_indiceizq = l_indiceizq + 1.		

	l_indice = l_indice + 1
	l_rs.MoveNext
loop
l_rs.Close

if len(trim(l_sql2)) <> 0 then
	l_rs.Open l_sql2, cn
	l_indice = 0
	l_indiceder = 0
	do until l_rs.eof
	' muestra codigo estruc,descripcion
	
		l_texto =  l_rs(0)& " " & l_rs(1)
		response.write "obj = " & l_obj & ";" & vbCrLf
		response.write "newOp = new Option();" & vbCrLf
		response.write "newOp.value = "  & l_rs(0) & ";" & vbCrLf
		response.write "newOp.text  = '" & l_texto & "';" & vbCrLf
	
		response.write "selfil.options[" & l_indiceder & "] = newOp;"  & vbCrLf
		l_indiceder = l_indiceder + 1.		
		l_indice = l_indice + 1
		on error resume next
		l_rs.MoveNext
	loop
l_rs.Close
'cn.Close
end if

response.write "}" & vbCrLf


response.write "function Uno(fuente,destino){" & vbCrLf &_
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione una autorizacion'); }" & vbCrLf &_
		   "    else {" & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value= fuente[fuente.selectedIndex].value;" & vbCrLf &_
       	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf

response.write "function UnoDerecha(fuente,destino){" & vbCrLf &_
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione una autorizacion'); }" & vbCrLf &_
		   "    else {" & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value= fuente[fuente.selectedIndex].value;" & vbCrLf &_
    	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf


response.write "function Aceptar(obj){" & vbCrLf &_
    	   "    var cadena = ',';" & vbCrLf &_
    	   "    var i;" & vbCrLf &_
    	   "    var long = selfil.length-1;" & vbCrLf &_
    	   "    for (i=0;i<=long;i++){" & vbCrLf &_
    	   "        cadena = cadena + selfil[i].value + ','   ;" & vbCrLf &_
    	   "    }" & vbCrLf &_
    	   "    obj.value = cadena;" & vbCrLf &_
    	   "    Salvar(obj);" & vbCrLf &_
    	   "    window.close();" & vbCrLf &_
    	   "}" & vbCrLf



response.write "function Todos(fuente,destino){"  & vbCrLf &_
    	   "x=fuente.length;"  & vbCrLf &_
		   "    for (i=1;i<=x;i++){" & vbCrLf &_
           "var opcion = new Option();" & vbCrLf &_
           "opcion.value= fuente[0].value;" & vbCrLf &_
           "opcion.text  = fuente[ 0].text;" & vbCrLf &_
           "fuente.remove(0);" & vbCrLf &_
           "destino.add(opcion);" & vbCrLf &_
   			"}" & vbCrLf &_
	   "}" 



response.write "</script>" & vbCrLf
End Sub
' ----------------------------------------------------------------------------------

dim l_obj
dim l_userid
dim l_usrnombre

dim l_rs
dim l_rs1
dim l_rs2
dim l_sql
dim l_sql1
dim l_sql2
Dim l_cm

dim l_seleccion

l_userid	= Request.QueryString("userid")
l_obj		= Request.QueryString("obj")

'--------------------- Doble browse -------------------------------------

l_obj = "opener." & Request.QueryString("obj")

Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
' se envia el select para armar el lado DERECHO   del browse ---------------
l_sql2 = "SELECT cysfincirc.cystipnro, cystipo.cystipnombre"
l_sql2 = l_sql2 & " FROM cysfincirc "
l_sql2 = l_sql2 & " INNER JOIN cystipo ON cystipo.cystipnro = cysfincirc.cystipnro "
l_sql2 = l_sql2 & " WHERE cysfincirc.userid = '" & l_userid & "'"

' se envia el select para armar el lado IZQUIERDO del browse --------------
l_sql = "SELECT cystipo.cystipnro, cystipo.cystipnombre "
l_sql = l_sql & " FROM cystipo "
l_sql = l_sql & " WHERE NOT EXISTS "
l_sql = l_sql & " (SELECT * FROM  cysfincirc WHERE "
l_sql = l_sql & " cysfincirc.cystipnro = cystipo.cystipnro"
l_sql = l_sql & "  AND "
l_sql = l_sql & " cysfincirc.userid = '" & l_userid & "')"

DobleBrowse l_sql, l_sql2 

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT iduser,usrnombre  "
l_sql = l_sql & " FROM user_per"
l_sql = l_sql & " INNER JOIN perf_usr ON user_per.perfnro = perf_usr.perfnro "
l_sql = l_sql & " WHERE user_per.iduser = '" & l_userid & "'"
l_rs.Maxrecords = 1
rsOpen l_rs, cn, l_sql, 0 

if not l_rs.EOF then
	l_usrnombre = l_rs("usrnombre")
end if%>
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Fin de Firma</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
</head>


<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:CargarArreglo();" >

<input type=hidden name=radio>
<input type=hidden name=fi-grupo value="">

<input type=hidden name="seleccion"  value=<%=l_seleccion%>>
<input type=hidden name="userid"  value=<%=l_userid%>>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td class="th2" colspan=4>
	<script>document.write(document.title);</script>
	<br>
	Usuario: <%=l_usrnombre%>
	</td>
</tr>


<table width=100% border=0>
<tr>
    <td><b>Autorizaciones No Seleccionadas</b><br>
    <select class="doblebrowse" size=20 name=nselfil ondblclick="javascript:UnoDerecha(nselfil,selfil);"></select>
    </td>
	<td align=left width=40>
	<a class=sidebtnSHW href="javascript:Todos(nselfil,selfil);">>></a></a>
	<a class=sidebtnSHW href="javascript:UnoDerecha(nselfil,selfil);">></a></a>
	<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil);"><</a></a></a>
	<a class=sidebtnSHW href="javascript:Todos(selfil,nselfil);"><<</a></a>
    </td>
    <td><b>Autorizaciones Seleccionadas</b><br>
    <select class="doblebrowse" size=20 name=selfil ondblclick="javascript:Uno(selfil,nselfil);"></select>
    </td>
</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td align="right" class="th2">
	<a class=sidebtnABM href="javascript:Aceptar(seleccion)">Aceptar</a>
	<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>
</table>


</body>
</html>
