<%Option Explicit%>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->

<!--
-----------------------------------------------------------------------------
Archivo        : hab_emp_ess_00.asp
Descripcion    : Pagina utilizada para habilitar al empleado a entrar a Autogestion
Creador        : GdeCos
Fecha Creacion : 14/4/2005
modificado	   : 21/09/2006 - Martin Ferraro - Se agrego actualizacion del password
		   	   : 24/08/2007 - FGZ - Se agregó el control de nulo para el campo empessactivo
		   	   : 30/08/2007 - Martin Ferraro - el src="" daba error en httpS
-----------------------------------------------------------------------------
-->
<%
on error goto 0

Dim l_rs
Dim l_sql
Dim l_ternro
Dim l_empessactivo
Dim l_perfnro

l_ternro = request.querystring("ternro")

l_empessactivo = ""
l_perfnro = 0

Set l_rs = Server.CreateObject("ADODB.RecordSet")	

if not l_ternro = "" then
	l_sql = "SELECT empessactivo,perfnro FROM empleado WHERE ternro = " & l_ternro
	rsOpen l_rs, cn, l_sql, 0
	if not l_rs.eof then
		if isNull(l_rs("empessactivo")) then
			l_empessactivo = 0
		else
			l_empessactivo = Cint(l_rs("empessactivo"))
		end if

		if isNull(l_rs("perfnro")) then
		   l_perfnro      = 0
		else
		   l_perfnro      = CLng(l_rs("perfnro"))		
		end if
	end if
	 l_rs.Close
end if

%>
<html>
<head>
<title>buques - Heidt & Asociados S.A.</title>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script>

function ok2(){
		
		//if((document.FormVar.newpass.value.length != 0)||(document.FormVar.newpassrep.value.length != 0)){
		if(document.FormVar.tg.checked){

			if (document.FormVar.newpass.value.length == 0) 
			{
				alert("Debe ingresar la Contraseña.");		
				document.FormVar.newpass.focus();
				return;
			}
			if (document.FormVar.newpass.value.length < <%= MinCarPass_ess %>) 
			{
				alert("La Contraseña debe tener mas de <%= MinCarPass_ess - 1 %> caracteres.");		
				document.FormVar.newpass.focus();
				return;
			}
			if (document.FormVar.newpass.value != document.FormVar.newpassrep.value){
		      alert("La Nueva contraseña y la Repetición no coinciden.");
			  document.FormVar.newpass.focus();
			  return;
		  	}
		}

		document.FormVar.action = "hab_emp_ess_01.asp";

		document.FormVar.target = "ifrmx";
		document.FormVar.method = "POST";
		document.FormVar.submit();
}

function CambioEstado(){
	if (document.FormVar.estado.value == 0){
		document.FormVar.tg.disabled = true;
		document.FormVar.tg.checked = false;
	}
	else{
		document.FormVar.tg.disabled = false;
		document.FormVar.tg.checked = false;
	}
	CambioTg();
}

function CambioTg(){
	if (!document.FormVar.tg.checked){
		document.FormVar.newpass.value = '';
		document.FormVar.newpass.disabled = true;
		document.FormVar.newpass.className = "deshabinp";
		document.FormVar.newpassrep.value = '';
		document.FormVar.newpassrep.disabled = true;
		document.FormVar.newpassrep.className = "deshabinp";
	}
	else{
		document.FormVar.newpass.value = '';
		document.FormVar.newpass.disabled = false;
		document.FormVar.newpass.className = "habinp";
		document.FormVar.newpassrep.value = '';
		document.FormVar.newpassrep.disabled = false;
		document.FormVar.newpassrep.className = "habinp";
	}
}
window.resizeTo(370,245);
</script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="FormVar" method="post">
<input type="Hidden" name="ternro" value="<%= l_ternro %>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" background="" align="center">
	<tr>
	  <td class="th2" colspan="2" height="1">
	     <table cellpadding="0" cellspacing="0" border="0" width="100%">
		    <tr>
			    <td class="th2" height="10">Ingreso a Autogesti&oacute;n</td>
				<td class="th2" align="right">		  
					<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
				</td>
			</tr>
		 </table>
	  </td>
	</tr>
	<tr nowrap>
		<td align="right" width="30%">
			<b>Estado&nbsp;Actual&nbsp;:</b>
		</td>		
		<td align="left">
		   <select name="estado" size="1" style="width:110px" onchange="CambioEstado()">
					<option value="1" <% if l_empessactivo < 0 then %> selected <% end if %>>Habilitado</option>
					<option value="0" <% if l_empessactivo = 0 then %> selected <% end if %>>Inhabilitado</option>
		   </select>
		</td>		
	</tr>		
   <tr>
	<td align="right" nowrap><b>Perfil&nbsp;:</b></td>
	<td colspan="1">
		<select name="perfnro" size="1" style="width: 200px;">
			<% 
			l_sql = "SELECT * "
			l_sql = l_sql & "FROM perf_usr ORDER BY perfnom "
			
			rsOpen l_rs, cn, l_sql, 0
			do until l_rs.eof
			 	%>
				<option value="<%= l_rs("perfnro") %>" <%if Clng(l_rs("perfnro")) = l_perfnro then response.write "selected" end if%>><%= l_rs("perfnom") %></option>
				<%
				l_rs.MoveNext
			loop
			l_rs.close
			%>
		</select>
	</td>
</tr>

<tr>
    <td align="right"><b>Cambiar&nbsp;Contraseña :</td>
	<td>
		<input type="Checkbox" name="tg" onclick="CambioTg();">
	</td>
</tr>


<tr>
    <td align="right"><b>Nueva&nbsp;Contraseña :</td>
	<td>
		<input type="Password" name="newpass" size="20" maxlength="20">
	</td>
</tr>

<tr>
    <td align="right"><b>Repetir&nbsp;Contraseña :</b></td>
	<td>
		<input type="Password" name="newpassrep" size="20" maxlength="20">
	</td>
</tr>
	
	<tr>
	    <td height="20" colspan=2>
			&nbsp;
		</td>
    </tr>
	<tr>
	    <td  colspan="2" align="right" class="th2" height="1">
			<a class=sidebtnABM href="Javascript:ok2();">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
		</td>
	</tr>
</table>
<iframe name="ifrmx" src="blanc.asp" style="" width="0" height="0"></iframe>
</form>
<script>CambioEstado();</script>
</body>
</html>
