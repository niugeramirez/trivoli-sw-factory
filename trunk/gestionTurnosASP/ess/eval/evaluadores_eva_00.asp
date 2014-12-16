<% Option Explicit %>
<!--#include virtual="/ess/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const_eva.inc"-->
<%
'================================================================================
'Archivo		: evaluadores_eva_00.asp
'Descripción	: Form de asignacion MANUAL de evaluadores
'Autor			: CCRossi
'Fecha			: 19-05-2004
'Modificado		:
'================================================================================

'ADO
Dim l_sql
Dim l_rs

'Local
dim l_nombre 
dim l_nombrerevisor

'Parametros
Dim l_ternro
Dim l_evaevenro

l_ternro    = request.querystring("ternro")
l_evaevenro = request.querystring("evaevenro")

Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT terape, terape2, ternom, ternom2 FROM empleado WHERE empleado.ternro = " & l_ternro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_nombre = l_rs("terape")
	if trim(l_rs("terape2"))<>"" then
		l_nombre = l_nombre & " " & l_rs("terape2")
	end if	
	if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & ","
	end if	
	if trim(l_rs("ternom"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom"))
	end if	
	if trim(l_rs("ternom2"))<>"" then
		l_nombre = l_nombre & " " & trim(l_rs("ternom2"))
	end if	
end if
l_rs.close
set l_rs=nothing
		
%>
<html>
<head>
<link href="../<%=c_estilo %>" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Asignaci&oacute;n Manual de <%if ccodelco=-1 then%>Roles<%else%>Evaluadores<%end if%> - Gesti&oacute;n de Desempeño - RHPro &reg;</title>
</head>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
function Validar_Formulario()
{
	var i;
	var lista;
	var vacio;
	//alert(document.ifrelem.datos.cheq1.checked);
	//alert(document.ifrelem.datos.elem1.value);
		
	var formElements = document.datos.elements;
	i=2;
	lista='';
	vacio=0;
	while (i<formElements.length-1)
	{
		if (lista=='') 
			lista = lista + formElements[i].value;
		else	
			lista = lista + ','+ formElements[i].value;
		
		lista = lista + ','+ formElements[i+1].value;
		if (formElements[i+1].value=="")
			vacio=1;
		i = i + 3;
	}	
	if (vacio==1)
		alert('Asigne todos los Roles.');
	else{	
		var r = showModalDialog('evaluadores_eva_01.asp?ternro=<%=l_ternro%>&evaevenro=<%=l_evaevenro%>&lista='+lista, '','dialogWidth:50;dialogHeight:50'); 
		window.opener.ifrm.location.reload();
		window.close();
	}
	
}

function Teclarev(num,empleg,campo1,campo2){
   if (num==13) {
		buscarrevisor(empleg,campo1, campo2)
		return false;
  }
  return num;
}

function buscarrevisor(esto,campo1, campo2){
if (isNaN(esto)){
	esto = "";
	<%if ccodelco=-1 then%>
		alert("El número ingresado no es correcto.");
	<%else%>	
		alert("El legajo ingresado no es correcto.");
	<%end if%>	
	}
else {
	if (esto=="")
	{
		<%if ccodelco=-1 then%>
		alert("El número ingresado no es correcto.");
		<%else%>	
		alert("El legajo ingresado no es correcto.");
		<%end if%>	
	}
	else	
		abrirVentanaH('nuevo_rev.asp?empleg='+esto+'&campo1='+campo1+'&campo2='+campo2,'',200,100);
	}
}

</script>

<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="">
<form name="datos" action="evaluadores_eva_01.asp" method="post" >

<input type="Hidden" name="ternro"	  value="<%= l_ternro %>">
<input type="Hidden" name="evaevenro" value="<%= l_evaevenro %>">


<table cellspacing="0" cellpadding="0" border="0" width="100%" height="5%">
<tr>
    <td class="th2"></td>
	<td class="th2" align="right">		  
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
</table>

<table cellspacing="0" cellpadding="0" border="0" width="100%" height="96%">

    	<%
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = " SELECT distinct evatevdesabr, evatipevalua.evatevnro, evaluador, empleado.empleg, terape, terape2, ternom, ternom2 FROM evadetevldor INNER JOIN evacab ON evacab.evacabnro = evadetevldor.evacabnro  INNER JOIN evaoblieva ON evaoblieva.evatevnro = evadetevldor.evatevnro AND evaoblieva.evaseccnro = evadetevldor.evaseccnro "
l_sql = l_sql & " LEFT JOIN empleado ON empleado.ternro=evadetevldor.evaluador  INNER JOIN evatipevalua ON evatipevalua.evatevnro =evadetevldor.evatevnro WHERE evacab.evaevenro = " & l_evaevenro & " AND   evacab.empleado = " & l_ternro
rsOpen l_rs, cn, l_sql, 0 
if l_rs.eof then%>
		<tr>
			<td align="center" colspan=2><b>No se encuentran Roles a los cuales asignar Evaluadores.
			<br><br>
			<b>Si el empleado ya ha sido relacionado con anterioridad utilice la Opci&oacute;n de Configuraci&oacute;n<br>
			de secciones y verifique la Asignaci&oacute;n de Roles.</b>
			<br><br>
			Sino Acepte para grabar en la ventana de Empleados del Evento y <br> luego ingrese a Asignaci&oacute;n de Evaluadores.<br><br>
		</td>
		</tr>
		<tr>
			<td align="center" colspan=2><b></b></td>
		</tr>
		<%
		else
		do while not l_rs.eof 
			l_nombrerevisor = l_rs("terape")
			if trim(l_rs("terape2"))<>"" then
				l_nombrerevisor = l_nombrerevisor & " " & l_rs("terape2")
			end if	
			if trim(l_rs("ternom"))<>"" or trim(l_rs("ternom2"))<>"" then
				l_nombrerevisor = l_nombrerevisor & ","
			end if	
			if trim(l_rs("ternom"))<>"" then
				l_nombrerevisor = l_nombrerevisor & " " & trim(l_rs("ternom"))
			end if	
			if trim(l_rs("ternom2"))<>"" then
				l_nombrerevisor = l_nombrerevisor & " " & trim(l_rs("ternom2"))
			end if	
			%>
		<tr>
		<td align="right"><b><%=l_rs(0)%>:</b></td>
		<td>
		
<input type="hidden" name="evatevnro<%=l_rs("evatevnro")%>" value="<%=l_rs("evatevnro")%>">
<input type="text" name="rempleg<%=l_rs("evatevnro")%>" value="<%=l_rs("empleg")%>" onKeyPress="return Teclarev(event.keyCode,this.value,'document.datos.rempleg<%=l_rs("evatevnro")%>','document.datos.revisor<%=l_rs("evatevnro")%>')" onChange="buscarrevisor(this.value,'document.datos.rempleg<%=l_rs("evatevnro")%>','document.datos.revisor<%=l_rs("evatevnro")%>');" size="8" class="rev" > 

<a onclick="JavaScript:window.open('help_emp_00.asp?campo1=document.datos.rempleg<%=l_rs("evatevnro")%>&campo2=document.datos.revisor<%=l_rs("evatevnro")%>','new','toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no,width=700,height=400');" onmouseover="window.status='Buscar Empleado por Apellido'" onmouseout="window.status=' '" style="cursor:hand;">
   <img align="absmiddle" src="/serviciolocal/shared/images/profile.gif" alt="Ayuda Empleados" border="0">
</a>
<input class="rev" name="revisor<%=l_rs("evatevnro")%>" value="<%=l_nombrerevisor%>" style="background : #e0e0de;" readonly type="text" size="35" maxlength="35" >
		</td>
		</tr>
		<%l_rs.MoveNext
		loop
		end if
		l_rs.Close
		set l_rs=nothing%>


<tr height=42>
    <td  colspan="2" valign=top align="right" class="th2">
		<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
	</td>
</tr>

</table>
<iframe name="valida" style="visibility=hidden;" src="blanc.asp" width="100%" height="100%"></iframe> 
</form>
</body>
</html>
