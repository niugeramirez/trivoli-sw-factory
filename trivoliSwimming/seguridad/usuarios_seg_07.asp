<% Option Explicit %>
<!--#include virtual="/trivoliSwimming/shared/db/conn_db.inc"-->

<script>

function Salvar(obj){
   //alert(iduser.value);
   window.open('usuarios_seg_08.asp?tenro='+tenro.value + '&iduser=' + iduser.value + '&grabar=' + obj, '','','');
}

function Recargar(){
	document.location = 'usuarios_seg_07.asp?tenro=' + tenro.value + '&userid=' + iduser.value+'&obj=seleccion' ;
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
rsOpen l_rs, cn, l_sql, 0 		
l_indice = 0
l_indiceder = 0
l_indiceizq = 0

do until l_rs.eof
	l_texto =  l_rs(0) & " " & l_rs(1)
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
	rsOpen l_rs, cn, l_sql2, 0 		
	l_indice = 0
	l_indiceder = 0
	do until l_rs.eof
	' muestra codigo estruc
	
		l_texto =  l_rs(0) & " " & l_rs(1) 
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
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione una estructura'); }" & vbCrLf &_
		   "    else {" & vbCrLf &_
    	   "    var opcion = new Option();" & vbCrLf &_
    	   "    opcion.value = fuente[fuente.selectedIndex].value;" & vbCrLf &_
       	   "    opcion.text  = fuente[fuente.selectedIndex].text;" & vbCrLf &_
    	   "    fuente.remove(fuente.selectedIndex);" & vbCrLf &_
    	   "    destino.add(opcion);" & vbCrLf &_
    	   "    destino[destino.length-1].focus();" & vbCrLf &_
    	   "}" & vbCrLf &_
    	   "}" & vbCrLf

response.write "function UnoDerecha(fuente,destino){" & vbCrLf &_
		   "    if (fuente.selectedIndex == -1) { alert ('Seleccione una estructura'); }" & vbCrLf &_
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
    	   "    cadena = cadena + tenro.value + ','   ;" & vbCrLf &_
    	   "   //alert(cadena);" & vbCrLf &_
    	   "    for (i=0;i<=long;i++){" & vbCrLf &_
    	   "        cadena = cadena + selfil[i].value + ','   ;" & vbCrLf &_
    	   "   //alert(cadena);" & vbCrLf &_
    	   "    }" & vbCrLf &_
    	   "    obj.value = cadena;" & vbCrLf &_
    	   "    Salvar(cadena);" & vbCrLf &_
    	   "    window.close();" & vbCrLf &_
    	   "}" & vbCrLf



response.write "</script>" & vbCrLf
End Sub
' ----------------------------------------------------------------------------------

dim l_obj
dim l_tenro
dim l_iduser

dim l_rs
dim l_rs1
dim l_rs2
dim l_sql
dim l_sql1
dim l_sql2

dim l_seleccion

l_iduser     = Request.QueryString("userid")
l_tenro      = Request.QueryString("tenro")
l_obj        = Request.QueryString("obj")


'--------------------- Doble browse -------------------------------------

l_obj = "opener." & Request.QueryString("obj")

if len(trim(l_tenro)) = 0 then
	' encontrar el primer registro de estructura
	l_sql1 = "SELECT tipoestructura.tenro, "
	l_sql1 = l_sql1 & " tipoestructura.tedabr "
	l_sql1 = l_sql1 & " FROM tipoestructura "
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_rs1.Maxrecords = 50
	rsOpen l_rs1, cn, l_sql1, 0 
	l_rs1.MoveFirst
	if not l_rs1.eof then	
		l_tenro = l_rs1("tenro")
	end if
	l_rs1.Close
end if

' se envia el select para armar el lado DERECHO   del browse ---------------
l_sql2 = "SELECT  usupuedever.estrnro , "
l_sql2 = l_sql2 & "  estructura.estrdabr  "
l_sql2 = l_sql2 & " FROM usupuedever "
l_sql2 = l_sql2 & " INNER JOIN estructura     ON estructura.estrnro   = usupuedever.estrnro "
l_sql2 = l_sql2 & " INNER JOIN tipoestructura ON tipoestructura.tenro = usupuedever.tenro  "
l_sql2 = l_sql2 & " WHERE usupuedever.iduser = '" & l_iduser
l_sql2 = l_sql2 & "' AND   usupuedever.tenro   = " & l_tenro

' se envia el select para armar el lado IZQUIERDO del browse --------------
l_sql = "SELECT  estructura.estrnro, estructura.estrdabr "
l_sql = l_sql & " FROM estructura WHERE "
l_sql = l_sql & " estructura.tenro= " & l_tenro
l_sql = l_sql & " AND  " 
l_sql = l_sql & " NOT EXISTS (SELECT *  " 
l_sql = l_sql & " FROM usupuedever   "
l_sql = l_sql & " WHERE usupuedever.tenro    = estructura.tenro "
l_sql = l_sql & " AND   usupuedever.estrnro  = estructura.estrnro "
l_sql = l_sql & " AND   usupuedever.iduser  = '" & l_iduser & "')"

DobleBrowse l_sql, l_sql2 
%>
<html>
<head>
<link href="/trivoliSwimming/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Estructuras del Usuario</title>
<script src="/trivoliSwimming/shared/js/fn_windows.js"></script>
<script src="/trivoliSwimming/shared/js/fn_confirm.js"></script>
</head>


<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0" onload="Javascript:CargarArreglo();" >

<input type=hidden name=radio>
<input type=hidden name=fi-grupo value="">

<input type=hidden name="iduser" value="<%=l_iduser%>">
<input type=hidden name="seleccion" value=<%=l_obj%>>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td class="th2" colspan=2>
	<script>document.write(document.title);</script>
	</td>
</tr>
<tr>
    <td colspan=2>
	Estructuras para el Usuario:<b> <%=l_iduser%></b><br>
	</td>
</tr>

<tr>
	<%
	' BUSCAR Estructuras PARA EL <SELECT>
	l_sql1 = "SELECT tipoestructura.tenro, "
	l_sql1 = l_sql1 & " tipoestructura.tedabr "
	l_sql1 = l_sql1 & " FROM tipoestructura "
	%>

    <td align="right"><b>Estructuras:</b></td>
	<td align="left">
	<select name=tenro onchange="javascript:Recargar();">
	<%
	Set l_rs1 = Server.CreateObject("ADODB.RecordSet")
	l_rs1.Maxrecords = 50
	rsOpen l_rs1, cn, l_sql1, 0 

	l_rs1.MoveFirst
	   do while not l_rs1.eof%>
			<option value=<%=l_rs1("tenro")%>><%=l_rs1("tedabr")%></option>
		<%l_rs1.MoveNext
		loop
		l_rs1.Close
		'set l_rs1 = nothing%>
		</select>
		<script>tenro.value='<%=l_tenro%>'</script>
	</td>
</tr>

</table>
<table width=100% border=0>

<tr>
    <td><b>Estructuras No Seleccionadas</b><br>
    <select class="doblebrowse" size=20 name=nselfil ondblclick="javascript:UnoDerecha(nselfil,selfil);"></select>
    </td>
	<td align=left width=40>
	<!--a class=sidebtnSHW href="javascript:Todos(nselfil,selfil);">>></a></a-->
	<a class=sidebtnSHW href="javascript:UnoDerecha(nselfil,selfil);">></a></a>
	<a class=sidebtnSHW href="javascript:Uno(selfil,nselfil);"><</a></a></a>
	<!--a class=sidebtnSHW href="javascript:Todos(selfil,nselfil);"><<</a></a-->
    </td>
    <td><b>Estructuras Seleccionadas</b><br>
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
