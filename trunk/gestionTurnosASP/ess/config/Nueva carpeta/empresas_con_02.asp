<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<% 
'Archivo: empresas_con_02.asp
'Descripción: Abm de Cámaras
'Autor : Lisandro Moro
'Fecha: 08/02/2005

'Datos del formulario
'empnro	empcod	empdes	empdir	sitimpnro	locnro	empcaj	empnrocaj	empcai	empfecven	empfecini	empticpre	empremsuc	empremnro	empcarpor1suc	empcarpor1nro	empcarpor2suc	empcarpor2nro	empact
'on error goto 0

Dim l_empnro
Dim l_empcod
Dim l_empdes
Dim l_empdir
Dim l_sitimpdes
Dim l_locdes
Dim l_empcaj
Dim l_empnrocaj
Dim l_empcai
Dim l_empfecven
Dim l_empfecini
Dim l_empticpre
Dim l_empremsuc
Dim l_empremnro
Dim l_empcarpor1suc
Dim l_empcarpor1nro
Dim l_empcarpor2suc
Dim l_empcarpor2nro
Dim l_empcarpor3suc
Dim l_empcarpor3nro

Dim l_empcau1nro
Dim l_empfecvto1
Dim l_empcau2nro
Dim l_empfecvto2
Dim l_empcau3nro
Dim l_empfecvto3
Dim l_empesp
Dim l_dirloc
Dim l_locloc
Dim l_loclocdes
Dim l_locloccod

'ADO
Dim l_sql
Dim l_rs

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Empresas - Ticket</title>
</head>
<style type="text/css">
.none{
	padding : 0;
	padding-left : 0;
}
</style>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function validar(){
	if (isNaN(document.datos.empremnro.value)){
		document.datos.empremnro.select();
		alert('Debe ingresar un valor numérico en el Remito .');
		document.datos.empremnro.focus();
		return;
	}	

	if (!(document.datos.empcarpor1suc.value == "")) {	
  	if (isNaN(document.datos.empcarpor1suc.value)){
		document.datos.empcarpor1suc.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte 1 Sucursal.');
		document.datos.empcarpor1suc.focus();
		return;
	}	
	if (document.datos.empcarpor1nro.value ==""){
		document.datos.empcarpor1nro.select();
		alert('Debe ingresar un valor en \n Carta de Porte 1 Número.');
		document.datos.empcarpor1nro.focus();
		return;
	}	
	if (isNaN(document.datos.empcarpor1nro.value)){
		document.datos.empcarpor1nro.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte 1 Número.');
		document.datos.empcarpor1nro.focus();
		return;
	}	
	
	if (document.datos.empcau1nro.value == "") {
  		alert("Debe ingresar el Nro de CAC de C.Porte 1");
  		document.datos.empcau1nro.focus();
		return;
	}
	
	if (document.datos.empfecvto1.value == "") {
  		alert("La Fecha de Vencimiento de C.Porte 1 No debe ser Vacía");
  		document.datos.empfecvto1.focus();
		return;
	}

	if (!validarfecha(document.datos.empfecvto1)) {
  		document.datos.empfecvto1.focus();
		return;
	}
	}

	if (!(document.datos.empcarpor2suc.value == "")) {	
  	if (isNaN(document.datos.empcarpor2suc.value)){
		document.datos.empcarpor1suc.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte 2 Sucursal.');
		document.datos.empcarpor2suc.focus();
		return;
	}	
		if (document.datos.empcarpor2nro.value ==""){
		document.datos.empcarpor2nro.select();
		alert('Debe ingresar un valor en \n Carta de Porte 2 Número.');
		document.datos.empcarpor2nro.focus();
		return;
	}	
	
	if (isNaN(document.datos.empcarpor2nro.value)){
		document.datos.empcarpor2nro.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte 2 Número.');
		document.datos.empcarpor2nro.focus();
		return;
	}	
	
	if (document.datos.empcau2nro.value == "") {
  		alert("Debe ingresar el Nro de CAC de C.Porte 2");
  		document.datos.empcau2nro.focus();
		return;
	}
	
	if (document.datos.empfecvto2.value == "") {
  		alert("La Fecha de Vencimiento C.Porte 2 No debe ser Vacía");
  		document.datos.empfecvto2.focus();
		return;
	}

	if (!validarfecha(document.datos.empfecvto2)) {
  		document.datos.empfecvto2.focus();
		return;
	}
	}
	
	if (!(document.datos.empcarpor3suc.value == "")) {	
  	if (isNaN(document.datos.empcarpor3suc.value)){
		document.datos.empcarpor3suc.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte Ferro Sucursal.');
		document.datos.empcarpor3suc.focus();
		return;
	}	
	if (document.datos.empcarpor3nro.value ==""){
		document.datos.empcarpor3nro.select();
		alert('Debe ingresar un valor en \n Carta de Porte Ferro Número.');
		document.datos.empcarpor3nro.focus();
		return;
	}	
	if (isNaN(document.datos.empcarpor3nro.value)){
		document.datos.empcarpor3nro.select();
		alert('Debe ingresar un valor numérico en \n Carta de Porte Ferro Número.');
		document.datos.empcarpor3nro.focus();
		return;
	}	
	
	if (document.datos.empcau3nro.value == "") {
  		alert("Debe ingresar el Nro de CAC en C.P.Ferro");
  		document.datos.empcau3nro.focus();
		return;
	}
	
	if (document.datos.empfecvto3.value == "") {
  		alert("La Fecha de Vencimiento C.P.Ferro No debe ser Vacía");
  		document.datos.empfecvto3.focus();
		return;
	}

	if (!validarfecha(document.datos.empfecvto3)) {
  		document.datos.empfecvto3.focus();
		return;
	}
	}
	
	if (!(document.datos.locloccod.value == "")) {
	    var d=document.datos;
		document.ifrm.location = "empresas_con_06.asp?locloccod="+document.datos.locloccod.value;
		}
	else{
		valido();
		}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.locloccod.focus();
}

function Ayuda_Fecha(txt){
 var jsFecha = Nuevo_Dialogo(window, '/serviciolocal/shared/js/calendar.html', 16, 15);
 if (jsFecha == null){
 	//txt.value = '';
 }else{
 	txt.value = jsFecha;
 	//DiadeSemana(jsFecha);
	}
}

</script>
<% 
Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_empnro = request.querystring("cabnro")
l_sql = "SELECT empnro, empcod, empdes, empdir, sitimpdes, tkt_localidad.locdes, empcaj, empnrocaj, empcai "
l_sql = l_sql & " ,empfecven ,empfecini ,empticpre ,empremsuc ,empremnro ,empcarpor1suc ,empcarpor1nro "
l_sql = l_sql & " ,empcarpor2suc, empcarpor2nro, empcarpor3suc, empcarpor3nro "
l_sql = l_sql & " ,empcau1nro ,empcau2nro, empcau3nro, empfecvto1, empfecvto2, empfecvto3, empesp "
l_sql = l_sql & " ,dirloc , locloc, local.loccod locloccod, local.locdes loclocdes "
l_sql = l_sql & " FROM tkt_empresa "
l_sql = l_sql & " LEFT JOIN tkt_sitimp ON tkt_empresa.sitimpnro = tkt_sitimp.sitimpnro "
l_sql = l_sql & " LEFT JOIN tkt_localidad ON tkt_empresa.locnro = tkt_localidad.locnro "
l_sql = l_sql & " LEFT JOIN tkt_localidad local ON tkt_empresa.locloc = local.locnro "
l_sql  = l_sql  & " WHERE empnro = " & l_empnro
rsOpen l_rs, cn, l_sql, 0 
if not l_rs.eof then
	l_empnro = l_rs("empnro")
	l_empcod = l_rs("empcod")
	l_empdes = l_rs("empdes")
	l_empdir = l_rs("empdir")
	l_sitimpdes = l_rs("sitimpdes")
	l_locdes = l_rs("locdes")
	l_empcaj = l_rs("empcaj")
	l_empnrocaj = l_rs("empnrocaj")
	l_empcai = l_rs("empcai")
	l_empfecven = l_rs("empfecven")
	l_empfecini = l_rs("empfecini")
	l_empticpre = l_rs("empticpre")
	l_empremsuc = l_rs("empremsuc")
	l_empremnro = l_rs("empremnro")
	l_empcarpor1suc = l_rs("empcarpor1suc")
	l_empcarpor1nro = l_rs("empcarpor1nro")
	l_empcarpor2suc = l_rs("empcarpor2suc")
	l_empcarpor2nro = l_rs("empcarpor2nro")
	l_empcarpor3suc = l_rs("empcarpor3suc")
	l_empcarpor3nro = l_rs("empcarpor3nro")
	l_empcau1nro = l_rs("empcau1nro")
	l_empfecvto1 = l_rs("empfecvto1")
	l_empcau2nro = l_rs("empcau2nro")
	l_empfecvto2 = l_rs("empfecvto2")
	l_empcau3nro = l_rs("empcau3nro")
	l_empfecvto3 = l_rs("empfecvto3")
	l_empesp = l_rs("empesp") 
    l_dirloc = l_rs("dirloc") 
    l_locloc = l_rs("locloc") 
    l_locloccod = l_rs("locloccod") 	
    l_loclocdes = l_rs("loclocdes") 	
end if
l_rs.Close
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0"  onload="Javascript:document.datos.empcai.focus();">
<form name="datos" action="empresas_con_03.asp" method="post" target="ifrm">
	<input type="Hidden" name="empnro" value="<%= l_empnro %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
<tr>
    <td class="th2"  nowrap>Empresas</td>
	<td class="th2" align="right">
		<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
	</td>
</tr>
<tr>
	<td colspan="2" height="100%">
		<table border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td width="5%"></td>
				<td>
					<table cellspacing="0" cellpadding="0" border="0">
						<tr>
						    <td align="right" nowrap width = "20%"><b>Código:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empcod" size="12" maxlength="20" value="<%= l_empcod %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Descripción:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empdes" size="50" maxlength="50" value="<%= l_empdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Dirección Fiscal:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empdir" size="50" maxlength="50" value="<%= l_empdir %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Localidad:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="locdes" size="50" maxlength="50" value="<%= l_locdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Situación Impositiva:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="sitimpdes" size="50" maxlength="50" value="<%= l_sitimpdes %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Caja:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empcaj" size="20" maxlength="20" value="<%= l_empcaj %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Número Caja:</b></td>
							<td>
								<input type="text" readonly class="deshabinp" name="empnrocaj" size="20" maxlength="20" value="<%= l_empnrocaj %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Fecha Inicio:</b></td>
 							<td align="left">
								<input type="text" class="habinp" name="empfecini" size="10" maxlength="10" value="<%= l_empfecini %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Nro. Ticket Preimp.:</b></td>
							<td>
								<input type="text" class="habinp" name="empticpre" size="10" maxlength="9" value="<%= l_empticpre %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Último Nro. Remito:</b></td>
							<td class="none">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
								<td align="left">
								<input type="text" class="habinp" name="empremsuc" size="4" maxlength="4" value="<%= l_empremsuc %>">
								<b>-</b>
								<input type="text"  class="habinp" name="empremnro" size="8" maxlength="8" value="<%= l_empremnro %>">
								</td>
								<td><b>CAI:</b></td>
								<td>
									<input type="text"  class="habinp" name="empcai" size="15" maxlength="15" value="<%= l_empcai %>">									
								</td>
								<td><b> F.Vto:</b></td>
								<td>
									<input type="text"  class="habinp" name="empfecven" size="10" maxlength="10" value="<%= l_empfecven %>">
								    <!--<a href="Javascript:Ayuda_Fecha(document.datos.empfecvto1);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>-->
								</td>
								</tr>
 								</table>								
							</td>							
							
						</tr>
						<tr>
						    <td align="right" nowrap><b>Último Nro. C. Porte C. 1:</b></td>
							<td class="none">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
								<td align="left">
								<input type="text" class="habinp" name="empcarpor1suc" size="4" maxlength="4" value="<%= l_empcarpor1suc %>">
								<b>-</b>
								<input type="text"  class="habinp" name="empcarpor1nro" size="8" maxlength="8" value="<%= l_empcarpor1nro %>">
								</td>
								<td><b>CAC:</b></td>
								<td>
									<input type="text"  class="habinp" name="empcau1nro" size="15" maxlength="15" value="<%= l_empcau1nro %>">									
								</td>
								<td><b> F.Vto:</b></td>
								<td>
									<input type="text"  class="habinp" name="empfecvto1" size="10" maxlength="10" value="<%= l_empfecvto1 %>">
								    <!--<a href="Javascript:Ayuda_Fecha(document.datos.empfecvto1);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>-->
								</td>
								</tr>
 								</table>								
							</td>							
						</tr>
						<tr>
						    <td align="right" nowrap><b>Último Nro. C. Porte C. 2:</b></td>
							<td class="none">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
								<td align="left">
								<input type="text" class="habinp" name="empcarpor2suc" size="4" maxlength="4" value="<%= l_empcarpor2suc %>">
								<b>-</b>
								<input type="text"  class="habinp" name="empcarpor2nro" size="8" maxlength="8" value="<%= l_empcarpor2nro %>">
								</td>
								<td><b>CAC:</b></td>
								<td>
									<input type="text"  class="habinp" name="empcau2nro" size="15" maxlength="15" value="<%= l_empcau2nro %>">									
								</td>
								<td><b> F.Vto:</b></td>
								<td>
									<input type="text"  class="habinp" name="empfecvto2" size="10" maxlength="10" value="<%= l_empfecvto2 %>">
								    <!--<a href="Javascript:Ayuda_Fecha(document.datos.empfecvto1);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>-->
								</td>
								</tr>
 								</table>								
							</td>							
						</tr>	
						<tr>
						    <td align="right" nowrap><b>Último Nro. C. P. Ferro :</b></td>
							<td class="none">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
								<td align="left">
								<input type="text" class="habinp" name="empcarpor3suc" size="4" maxlength="4" value="<%= l_empcarpor3suc %>">
								<b>-</b>
								<input type="text"  class="habinp" name="empcarpor3nro" size="8" maxlength="8" value="<%= l_empcarpor3nro %>">
								</td>
								<td><b>CAC:</b></td>
								<td>
									<input type="text"  class="habinp" name="empcau3nro" size="15" maxlength="15" value="<%= l_empcau3nro %>">									
								</td>
								<td><b> F.Vto:</b></td>
								<td>
									<input type="text"  class="habinp" name="empfecvto3" size="10" maxlength="10" value="<%= l_empfecvto3 %>">
								    <!--<a href="Javascript:Ayuda_Fecha(document.datos.empfecvto1);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>-->
								</td>
								</tr>
 								</table>								
							</td>							
						</tr>	
						
						<tr>
   					        <td height="100%" align="right"><b>Especial:</b></td>
							<td height="100%">
								<input type="Checkbox"  name="empesp" <% If l_empesp = -1 then %>Checked<% end if %>>
							</td>
						</tr>		
						<tr>
						    <td align="right" nowrap><b>Dirección local:</b></td>
							<td>
								<input type="text" name="dirloc" size="50" maxlength="50" value="<%= l_dirloc %>">
							</td>
						</tr>
						<tr>
						    <td align="right" nowrap><b>Localidad:</b></td>
							<td align="left" >
							<table>
							<tr>
							<td>
								<input type="text" name="locloccod" size="6" maxlength="5" value="<%= l_locloccod %>">
							</td>
							<td>
								<input type="text" readonly class="deshabinp" name="loclocdes" size="37" maxlength="37" value="<%= l_loclocdes %>">
							</td>
							</tr>
							</table>
					    	</td>
						</tr>
											
					</table>
				</td>
				<td width="5%"></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
    <td colspan="2" align="right" class="th2">
		<iframe name="ifrm" style="visibility=hidden;" src="" width="0" height="0"></iframe> 
		<a class=sidebtnABM href="Javascript:validar()">Aceptar</a>
		<a class=sidebtnABM href="Javascript:window.close()">Salir</a>
	</td>
</tr>
</table>
</form>
<%
set l_rs = nothing
cn.Close
set cn = nothing
%>
</body>
</html>
