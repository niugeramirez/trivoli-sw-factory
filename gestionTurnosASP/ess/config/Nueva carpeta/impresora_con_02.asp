<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<% 
'Archivo: impresoras_con_02.asp
'Descripción: ABM de Impresoras
'Autor : Lisandro Moro
'Fecha: 26/09/2005
'Modificado: 

'on error goto 0


'Datos del formulario
Dim l_impnro
Dim l_impnom
Dim l_impnomcom
Dim l_impmat

'ADO
Dim l_tipo
Dim l_sql
Dim l_rs

l_tipo = request.querystring("tipo")

%>
<html>
<head>
<link href="/serviciolocal/shared/css/tables3.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title><%= Session("Titulo")%>Impresoras - Ticket</title>
</head>
<OBJECT id=cd classid=clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="847">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CancelError" VALUE="0">
	<PARAM NAME="Color" VALUE="0">
	<PARAM NAME="Copies" VALUE="1">
	<PARAM NAME="DefaultExt" VALUE="">
	<PARAM NAME="DialogTitle" VALUE="Buscar impresora">
	<PARAM NAME="FileName" VALUE="">
	<PARAM NAME="Filter" VALUE="">
	<PARAM NAME="FilterIndex" VALUE="0">
	<PARAM NAME="Flags" VALUE="0">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="FontName" VALUE="">
	<PARAM NAME="FontSize" VALUE="8">
	<PARAM NAME="FontStrikeThru" VALUE="0">
	<PARAM NAME="FontUnderLine" VALUE="0">
	<PARAM NAME="FromPage" VALUE="0">
	<PARAM NAME="HelpCommand" VALUE="0">
	<PARAM NAME="HelpContext" VALUE="0">
	<PARAM NAME="HelpFile" VALUE="">
	<PARAM NAME="HelpKey" VALUE="">
	<PARAM NAME="InitDir" VALUE="">
	<PARAM NAME="Max" VALUE="0">
	<PARAM NAME="Min" VALUE="0">
	<PARAM NAME="MaxFileSize" VALUE="260">
	<PARAM NAME="PrinterDefault" VALUE="0">
	<PARAM NAME="ToPage" VALUE="0">
	<PARAM NAME="Orientation" VALUE="1">
</OBJECT>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_valida.js"></script>
<script>
function Validar_Formulario(){
	if (Trim(document.datos.impnom.value) == ""){
		alert("Debe ingresar el nombre de la impresora.");
		document.datos.impnom.focus();
	}else if(Trim(document.datos.impnomcom.value) == ""){
		alert("Debe ingresar el nombre compartido de la impresora.");
		document.datos.impnomcom.focus();
	}else{
		var d=document.datos;
		document.valida.location = "impresora_con_06.asp?tipo=<%= l_tipo%>&impnro="+d.impnro.value + "&impnom="+d.impnom.value + "&impnomcom="+d.impnomcom.value + "&impmat="+d.impmat.checked;
	}	
}

function valido(){
	document.datos.submit();
}

function invalido(texto){
	alert(texto);
	document.datos.impnom.focus();
}

alert(document.all.cd.Name + 'Under construction');
function AbrirDialogo(){
	var prn, a, b;
	try{
		document.all.cd.showPrinter();
		//prn = new ActiveXObject(document.all.cd.showPrinter);
		//set prn = getObject(document.all.cd);
		//prn.showPrinter();
		getPrinterSharedName("licho");
		alert(document.all.cd.printer);
		for (a in document.all.cd){
			//alert(document.all.cd[a].name  + ' - ' + a);
			for (b in a){
				alert(a + '-' + b);
			}
			//alert(a);
		}
		//prn = new ActiveXObject(document.all.cd.showPrinter());

		//prn.showPrinter()
		//alert(prn.DeviceName);
	}catch(e){
		alert(e.description + ' - Esta terminal no se encuentra configurada.');
	}
}
</script>

<script language="VBScript">
	sub AbrirDialogo2()
		'''Set prn = CreateObject("Scripting.Printers")
		''document.all.cd.showPrinter()
		''msgbox  printer.deviceName,,"cuac"
		''//prn = document.all.cd.showPrinter()
		'for each a in prn
		'	msgbox a.name,,a.value
		'next
	end sub


</script>
<% 
select Case l_tipo
	Case "A":
		l_impnro = ""
		l_impnom = ""
		l_impnomcom = ""
		l_impmat = 0
	Case "M":
		Set l_rs = Server.CreateObject("ADODB.RecordSet")
		l_impnro = request.querystring("cabnro")
		
		l_sql = "SELECT  impnom, impnomcom ,impmat "
		l_sql = l_sql & " FROM tkt_impresora "
		l_sql  = l_sql  & " WHERE impnro = " & l_impnro
		rsOpen l_rs, cn, l_sql, 0 
		if not l_rs.eof then
			l_impnom = l_rs("impnom")
			l_impnomcom = l_rs("impnomcom")
			l_impmat = l_rs("impmat")
		end if
		l_rs.Close
end select
%>
<body leftmargin="0" rightmargin="0" topmargin="0" bottommargin="0" onload="JavaScript:document.datos.impnom.focus();">
<form name="datos" action="impresora_con_03.asp?tipo=<%= l_tipo %>" method="post" target="valida">
<input type="Hidden" name="impnro" value="<%= l_impnro %>">
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="100%">
	<tr>
	    <td class="th2" nowrap>Impresoras</td>
		<td class="th2" align="right">
			<a class=sidebtnHLP href="Javascript:ayuda('<%= Request.ServerVariables("SCRIPT_NAME")%>');">Ayuda</a>
		</td>
	</tr>
	<tr>
		<td colspan="2" height="100%">
			<table border="0" cellspacing="0" cellpadding="0"  >
				<tr>
					<td width="50%"></td>
					<td>
						<table cellspacing="0" cellpadding="0" border="0">
							<tr>
							    <td align="right" nowrap><b>Nombre Impresora:</b></td>
								<td>
									<input type="text" name="impnom" size="55" maxlength="80" value="<%= l_impnom %>">
								</td>
								<td>
									<a class=sidebtnABM href="Javascript:AbrirDialogo();">Buscar</a>
								</td>
							</tr>
							<tr>
							    <td align="right" nowrap><b>Nombre Compartido:</b></td>
								<td>
									<input type="text" name="impnomcom" size="55" maxlength="80" value="<%= l_impnomcom %>">
								</td>
								<td></td>
							</tr>
							<tr>
							    <td align="right"><b>Matricial:</b></td>
								<td>
									<input type="Checkbox" name="impmat" <%if l_impmat = "-1" then%>checked<%end if%>>
								</td>
								<td></td>
							</tr>
						</table>
					</td>
					<td width="50%"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
    	<td colspan="2" align="right" class="th2">
			<a class=sidebtnABM href="Javascript:Validar_Formulario()">Aceptar</a>
			<a class=sidebtnABM href="Javascript:window.close()">Cancelar</a>
			<iframe name="valida" style="visibility=hidden;" src="" width="0" height="0"></iframe> 
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
