<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<% 
Dim l_tipo
Dim l_texto1
Dim l_texto2
Dim l_texto3
Dim l_titulo
Dim l_error
dim l_empresa

l_empresa="desa"

l_tipo = request("tipo")

if l_tipo <> "pass" then
	l_texto1 = "Usuario:"
	l_texto2 = "Contraseña:"
	l_texto3 = "Planta:"
	l_titulo = "Identificación del Usuario"
	l_error  = ""
else
	l_texto1 = "Contraseña:"
	l_texto2 = "Nueva:"
	l_texto3 = "Confirma:"
	l_titulo = "Usuario: " & session("loguinUser")
	l_error  = "Cambiar contraseña"
end if
 %>
<html>
<head>
<title><%= Session("Titulo")%>Ticket - Oleaginosa Moreno S.A.</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="/serviciolocal/shared/css/tables4.css" rel="StyleSheet" type="text/css">
<style type="text/css">
<!--
a {
	text-decoration: none;
}
.texto{
	font : Arial;
	font-size : 17px;
	color : Navy;
	text-align : right;
	border : Black;
	border: none;
	background-color : transparent;
	border:none;
}
.blanc{
	font:16px;
	border:1px solid Navy;
	background-color : transparent;
	border:none;
/*	border:1px solid #87CEFA;*/
	border:1px solid Navy;
	width:90%;
	color:Navy;
	z-index:100;
}
.blancpass{
	font:16px;
/*	border:1px solid Blue;*/
/*	background-color : transparent;*/
	border:1px solid #87CEFA;
/*	border:none;*/
	width:118px;
	color: Navy;
	height: 22px;
}

-->
</style>
</head>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script>
var usertmp;
var loguser;
var passuser;
var domuser;
function tip(texto){
	document.FormVar.titulo.value = texto;
}
function help(){
	abrirVentana(document.FormVar.ayuda.value,'',750,550);
}

function ok(){
	t=event.keyCode;
	event.returnValue = true;
	if (t==13){
		ok2();
	}
}

function ok2(){
	<% If l_tipo <> "pass" then %>
		document.cookie = "usr=;expires=now";
		if (document.FormVar.seguridad.value == -1){
			document.FormVar.usr.value = domuser;
		}else{
			document.FormVar.usr.value = document.FormVar.usr2.value;
		}
		if (document.FormVar.empresa.value == 'desa'){
			document.FormVar.seg_nt.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
			document.FormVar.seguridad.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
			document.FormVar.base.value  = document.FormVar.basex[document.FormVar.basex.selectedIndex].bases;
		}else{
			document.all.seg_nt.value = document.FormVar.basex.seg;
			document.all.seguridad.value = document.FormVar.basex.seg;
			document.all.base.value = document.FormVar.basex.bases;
		}

		if (document.FormVar.usr2.value == ""){
			alert('Debe ingresar un usuario.');
			return;
		}
		document.FormVar.action = '../shared/db/defaultM.asp';
	<% Else  %>
		document.FormVar.action = '../seguridad/cambio_clave_seg_01.asp';
	<% End If %>
		document.FormVar.target = "ifrmx";
		document.FormVar.method = "POST";
		//alert(document.FormVar.usr.value +"="+ document.FormVar.pass.value +"="+ document.FormVar.base.value +"="+ document.FormVar.seg_nt.value+"="+ document.FormVar.menu.value+"="+ document.FormVar.debug.value);
		document.FormVar.submit();
}

function cancel(){
	<% If l_tipo <> "pass" then %>
		window.close();
	<% Else  %>
		window.location = "lanzador2M.asp";
	<% End If %>
}

function web(){
	abrirVentana(document.FormVar.inicio.value,'',750,550);
}

function email(){

}
function inicio(){
	<% if l_tipo <> "pass" then %>
		document.FormVar.usr2.focus();
	<% else %>
		document.FormVar.usrpass.focus();
	<% End If %>
	document.all.ifrm2.src = "lanzador.asp?menu=html&tipo=<%= l_tipo %>";
//	window.history.lenght=0;
	document.clear();
}

function baseopt(){
	<% If l_tipo <> "pass" then %>
		if (document.all.usr2.value != usertmp){
			loguser = document.FormVar.usr2.value;
			passuser = document.FormVar.pass.value;
		}
		if (document.FormVar.basex[document.FormVar.basex.selectedIndex].seg == -1){
			document.FormVar.usr2.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].user;
			document.FormVar.pass.value = "";
			document.FormVar.usr2.disabled = true;
			document.FormVar.pass.disabled = true;
		}else{
			document.FormVar.usr2.value = loguser;
			document.FormVar.pass.value = passuser;
			document.FormVar.usr2.disabled = false;
			document.FormVar.pass.disabled = false;
		}
	<% Else  %>
		document.FormVar.user.value = domuser;
	<% End If %>
	document.FormVar.base.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].bases;
	document.FormVar.seg_nt.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
	document.FormVar.seguridad.value = document.FormVar.basex[document.FormVar.basex.selectedIndex].seg;
	document.FormVar.debug.value = document.FormVar.debug.value;
}
function logout(arg){
	if (arg == "0"){
		//document.ifrm.location = "lanzador/logout.asp?arg=0";
	}else{
		//abrirVentanaH('logout.asp?arg=1','',10,10);
	}
}
</script>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="inicio();" bgcolor="#ffffff" onunload="Javascript:logout(1);">
<form name="FormVar" method="post">
	<input type="hidden" name="menu" value="html">
	<input type="hidden" name="base" value=0>
	<input type="hidden" name="seg_nt" value=0>
	<input type="hidden" name="usr" value="">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" >
  <tr> 
  	<td width="50%"></td>
    <td align="center" valign="middle">
	<fieldset>
	<!--<legend>Acceso menu</legend>-->
		<table border="0" class="tabla">
			<tr>
				<th colspan="2">
					<!--texto descriptivo en la barra sup-->
					<input class="texto" tabindex="-1" style="text-align:center;color:white;width:300px;font-weight:bold;font-size:16px;" name="texto" readonly value="<%= l_titulo %>">
				</th>
			</tr>
			<tr>
				<td colspan="2">
					<!-- ayuda sup izquierda -->
					<input name="titulo" tabindex="-1" class="blanc" style="border:none;" readonly type="Text" align="absmiddle" size="40" border="0" >
				</td>
			</tr>
			<tr>
				<td align="right">
					<!--texto 1-->
				    <input class="texto" tabindex="-1" style="width:120px;" readonly value="<%= l_texto1 %>">
				</td>
				<td>
				<% If l_tipo <> "pass" then %>
					<input name="usr2" class="blanc" type="Text" size="18" border="0"  onKeyPress="ok();">	
				<% Else  %>
					<input name="usrpass" class="blancpass" type="password" size="18" border="0"  onKeyPress="ok();">				
				<% End If %>
				</td>
			</tr>
			<tr>
				<td  align="right">
					<!--texto 2-->
					<input class="texto" tabindex="-1" style="width:120px;" readonly value="<%= l_texto2 %>">
				</td>
				<td>
				<% If l_tipo <> "pass" then %>
					<input name="pass" type="password" size="19" class="blanc"  onKeyPress="ok();">
				<% Else  %>
					<input name="usrpassnuevo" type="password"  size="19" class="blancpass"  onKeyPress="ok();">
				<% End If %>
				</td>
			</tr>
			<tr>
				<td  align="right">
					<!--texto 3-->
				    <input id="texto3" tabindex="-1" class="texto" style="width:120px;" readonly value="<%= l_texto3 %>">
				</td>
				<td>
				<% If l_tipo <> "pass" then %>
					<% If l_empresa = "desa" then %>
 						<div id="combobox" >
					   		<select name="basex" class="blanc" style="border:none;width:120px;font:15px" onchange="baseopt()"> 
							<option value=0 seg=-1 bases=2 user="">Manera</option>							
<!--							<option value=1 seg=0 bases=1 user="">M.America</option>    ¿ lo hereda de lanzador.asp ? -->
							</select>
						</div>
					<%Else%>	
						<input name=basex readonly seg="" bases="" user="" value='<%= l_empresa %>' tabindex=-1 type=Text  style='width:140px;border:none;text-align:left;font-family:Arial;font-size:17px;color:Blue;background-color:transparent;'>
					<%End If%>	
				<% Else  %>
					<input name="usrconfirm" type="password" size="19" class="blancpass"  onKeyPress="ok();">
				<% End If %>
				</td>
			</tr>
			<tr>
				<td align="right">
				</td>
				<td  align="right">
					<a class="sidebtnABM" href="Javascript:ok2()">Aceptar</a>
					&nbsp;&nbsp;&nbsp;
					<a class="sidebtnABM" href="Javascript:window.close()">Cancelar</a>
					&nbsp;
				</td>
			</tr>
			<tr>
				<td colspan="2" class="barra">
					<!-- ayuda o errores -->
					<input name="desc" tabindex="-1" class="blanc" style="width:240px;border:none;" ReadOnly type="Text" align="absmiddle" size="40" border="0">
				</td>
			</tr>
		</table>
	</fieldset>	
		<input type="hidden" name="inicio" value="">
		<input type="hidden" name="email2" value="">
		<input type="hidden" name="ayuda" value="">
		<input type="hidden" name="empresa" value="">
		<input type="hidden" name="seguridad" value=0>
		<input type="hidden" name="conexion" value=0>
		<input type="hidden" name="debug" value=0>
		<input type="hidden" name="tipo" value=<%= l_tipo %>>
		<iframe name="ifrm2" src="" style="visibility: hidden;" width="0" height="0"></iframe> 
		<iframe name="ifrmx" src="" style="visibility: hidden;" width="0" height="0"></iframe> <!---->
	</form>
	</td>
  	<td width="50%"></td>
  </tr>
</table>
<script>
	//alert(document.ifrm2.document.all.value);
</script>
<map name="Map">
  <area id="aceptar" shape="circle" coords="442,246,18" href="javascript:ok2();" onMouseOver="tip('Aceptar');" onMouseOut="tip('');">
  <area id="cancelar" shape="circle" coords="381,300,19" href="javascript:cancel();" onMouseOver="tip('Cancelar');" onMouseOut="tip('');">
  <area id="ayuda" shape="circle" coords="58,247,19" href="javascript:help();" onMouseOver="tip('Ayuda');" onMouseOut="tip('');">
  <area id="home" shape="circle" coords="75,290,18" href="#" onClick="web();" onMouseOver="tip('Home');" onMouseOut="tip('');">
  <area id="email" shape="circle" coords="118,298,18" href="#"  onMouseOver="tip('email');" onMouseOut="tip('');">
</map>
</body>
</html>
