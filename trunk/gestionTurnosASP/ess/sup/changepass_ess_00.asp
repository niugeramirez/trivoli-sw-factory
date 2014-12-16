<%Option Explicit%>
<!--#include virtual="/serviciolocal/ess/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/ess/shared/inc/accesoESS.inc"-->
<%
on error goto 0

Dim l_sql
Dim l_rs

Dim l_empleg
Dim l_apellido

Dim l_es_MSS
 
Set l_rs  = Server.CreateObject("ADODB.RecordSet")

l_sql = " SELECT * FROM empleado WHERE ternro=" & l_ess_ternro

rsOpen l_rs, cn, l_sql, 0 

l_apellido = l_rs("terape") & " " & l_rs("terape2") & ", " & l_rs("ternom") & " " & l_rs("ternom2")
l_empleg   = l_ess_empleg

l_rs.close

l_es_MSS = (CStr(Session("empleg")) <> CStr(l_ess_empleg))

%>
<!--
-----------------------------------------------------------------------------
Archivo        : changepass_ess_00.asp
Descripcion    : Pagina utilizada para cambiar el password del usuario en Autogestion
Creador        : GdeCos
Fecha Creacion : 1/4/2005
-----------------------------------------------------------------------------
-->

<html>
<head>
<title>buques - Heidt & Asociados S.A.</title>
<link href="../<%= c_estilo%>" rel="StyleSheet" type="text/css">
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>

function ok(){
	t=event.keyCode;
	event.returnValue = true;
	if (t==13){
		ok2();
	}
}

function ok2(){
		document.cookie = "usr=;expires=now";
		<%if not l_es_MSS then%>
		if (document.FormVar.passOld.value == ""){
			alert('Debe ingresar la contraseña actual.');
			return;
		}
		<%end if%>
		if (document.FormVar.passNew.value == ""){
			alert('Debe ingresar la contraseña nueva.');
			return;
		}
		if (document.FormVar.passNewRep.value == ""){
			alert('Debe repetir la nueva contraseña.');
			return;
		}
		if (document.FormVar.passNewRep.value != document.FormVar.passNew.value){
			alert('La nueva contraseña no coincide con la confirmación.');
			return;
		}

		document.FormVar.action = "changepass_ess_01.asp?empleg=<%= request.querystring("empleg") %>";

		document.FormVar.target = "ifrmx2";
		document.FormVar.method = "POST";
		document.FormVar.submit();
}

</script>
<body>
<form name="FormVar" method="post">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<th colspan="2">Cambiar Contraseña</th>
		</tr>
	<tr>
	    <td colspan=2> 
		<table align="center" style="width:60%; border-color:gray ; border-width: 1 ; border-style:solid ;">
			 <tr> 
			   	<td height="5" colspan=2>
			  	    &nbsp;
				</td>
			 </tr>
			<%if not l_es_MSS then%>
			<tr nowrap>
				<td align="right">
					<b>Contraseña Actual : </b>
				</td>		
				<td align="left">
			        <input name="passOld" type="password" size="19" onKeyPress="ok();">
				</td>		
			</tr>	
			<%else%>
			<tr nowrap>
				<td align="right">
					<b>Empleado&nbsp;:</b>
				</td>		
				<td align="left">
			        <%= l_apellido%>
				</td>		
			</tr>					
			<%end if%>
			<tr nowrap>
				<td align="right">
					<b>Nueva Contraseña : </b>
				</td>		
				<td align="left">
			        <input name="passNew" type="password" size="19" onKeyPress="ok();">
				</td>		
			</tr>		
			<tr nowrap>
				<td align="right">
					<b>Confirmar Contraseña : </b>
				</td>		
				<td align="left">
			        <input name="passNewRep" type="password" size="19" onKeyPress="ok();">
				</td>		
			</tr>		
			 <tr> 
			   	<td height="5" colspan=2>
			  	    &nbsp;
				</td>
			 </tr>
			  <tr> 
			    <td align="center" colspan=2> 	
					 <a class=sidebtnABM href="#" onclick="Javascript:ok2()" style="height=20px">&nbsp;&nbsp;&nbsp;Cambiar Contraseña &nbsp;&nbsp;&nbsp;</a>
					&nbsp;		
				</td>
			  </tr>
			 <tr> 
			   	<td height="5" colspan=2>
			  	    &nbsp;
				</td>
			 </tr>
		</table>
		</td>
  </tr>
  <tr> 
    <td height="20" colspan=2>
		&nbsp;
	</td>
  </tr>
</table>
<iframe name="ifrmx2" src="blanc.asp" style="" width="500" height="500"></iframe>
</form>
</body>
</html>
