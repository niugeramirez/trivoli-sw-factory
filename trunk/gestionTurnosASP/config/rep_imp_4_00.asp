<% Option Explicit %>
<!--#include virtual="/serviciolocal/shared/db/conn_db.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sec.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/const.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/fecha.inc"-->
<!--#include virtual="/serviciolocal/shared/inc/sqls.inc"-->
<html>
<head>
<link href="/serviciolocal/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title><%= Session("Titulo")%> Estadísticas </title>
<script src="/serviciolocal/shared/js/fn_windows.js"></script>
<script src="/serviciolocal/shared/js/fn_confirm.js"></script>
<script src="/serviciolocal/shared/js/fn_ayuda.js"></script>
<script src="/serviciolocal/shared/js/fn_fechas.js"></script>
<script src="/serviciolocal/shared/js/fn_ay_generica.js"></script>
<script>

<% on error goto 0
Dim l_rs
Dim l_sql

Dim l_fecini
Dim l_fecfin
Dim l_mernro2

'if month(date()) < 10 then
'	l_fecdes = "01/0" & month(date()) & "/" & year(date())
'	l_fecfin = "31/0" & month(date()) & "/" & year(date())
'else
'	l_fecdes = "01/" & month(date()) & "/" & year(date())
'	l_fecfin = "31/" & month(date()) & "/" & year(date())
'end if 

'	l_fecdes = "01/07" & "/2008" 
'	l_fecfin = "31/07" & "/2008" 

l_fecini 	  = request.querystring("qfecini")
l_fecfin 	  = request.querystring("qfecfin")
l_mernro2 	  = request.querystring("qmernro")

%>

function Imprimir(){
	parent.frames.ifrm.focus();
	window.print();	
}

function Actualizar(destino){

	var param;
	//Fechas	
	if ((document.datos.fecini.value == "") && (document.datos.fecfin.value == "" )) {
  		alert("Debe ingresar las Fechas Desde y Hasta");
  		document.datos.fecini.focus();
		return;
	}

	if ((document.datos.fecini.value == "") && (document.datos.fecfin.value != "" )) {
  		alert("Debe ingresar la Fecha Desde");
  		document.datos.fecini.focus();
		return;
	}
	
	if ((document.datos.fecfin.value == "") && (document.datos.fecini.value != "" )) {	
  		alert("Debe ingresar la Fecha Hasta");
  		document.datos.fecfin.focus();
		return;
	}
	
	if ((document.datos.fecini.value != "") && (document.datos.fecfin.value != "" )) {
	
			if (!validarfecha(document.datos.fecini)) {
		  		document.datos.fecini.focus();
				return;
			}	
			
			if (!validarfecha(document.datos.fecfin)) {
		  		document.datos.fecfin.focus();
				return;
			}	
			
			if (!(menorque(document.datos.fecini.value,document.datos.fecfin.value))) {
				alert("La Fecha Desde debe ser menor o igual que la Fecha Hasta.");
				document.datos.fecini.focus();
				return;
			}	  
	}
	
	if (document.datos.mernro.value == 0) {
  		alert("Debe Seleccionar un Producto");
  		document.datos.mernro.focus();
		return;
	}				
	
	param = "qfecini=" + document.all.fecini.value + "&qfecfin=" + document.all.fecfin.value + "&qmernro=" + document.all.mernro.value;
	
	if (destino== "exel")
    	abrirVentana("rep_exp_buques_con_01.asp?" + param + "&excel=true",'execl',250,150);
	else
		document.ifrm.location = "rep_imp_4_01.asp?" + param;			
	
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

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" 
<% if l_fecini <> "" then %>
	onload="Javascript:Actualizar('ifrm');" >
<% Else  %>
	onload="Javascript:document.datos.fecini.focus();" >
<% End If %>
<form name="datos">
<input  type="hidden" name="mernro2" value="<%= l_mernro2 %>" >
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" colspan="5">
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
			<!--
			<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>		  
			-->
			<!--<a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a> -->
			&nbsp;
			
		</td>
	</tr>
	<tr>
		<td align="right" size="10%">
			<b>Fecha Desde:</b>
		</td>
		<td>
			<input  type="text" name="fecini" size="10" maxlength="10" value="<% if l_fecini <> "" then response.write l_fecini else response.write "" end if %>" >
			<a href="Javascript:Ayuda_Fecha(document.datos.fecini);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td align="right" size="10%">
			<b>Fecha Hasta:</b>
		</td>	
		<td>
			<input  type="text" name="fecfin" size="10" maxlength="10" value="<% if l_fecfin <> "" then response.write l_fecfin else response.write "" end if %>">
			<a href="Javascript:Ayuda_Fecha(document.datos.fecfin);"><img src="/serviciolocal/shared/images/cal.gif" border="0"></a>
		</td>
		<td align="right"><b>Producto:</b></td>
		<td><select name="mernro" size="1" style="width:120;">
			<option value=0 selected>&nbsp;</option>
			<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
		 		  l_sql = "SELECT  * "
				  l_sql  = l_sql  & " FROM buq_mercaderia "
     			  l_sql  = l_sql  & " ORDER BY merdes "
 				  rsOpen l_rs, cn, l_sql, 0
				do until l_rs.eof		%>	
					<option value= <%= l_rs("mernro") %> > 
					<%= l_rs("merdes") %> </option>
					<%	l_rs.Movenext
				loop
			l_rs.Close %>
			</select>
			<script>document.datos.mernro.value= document.datos.mernro2.value </script>
			</td>	
	</tr>
	<tr valign="top" height="100%">
		<td colspan="6" align="center" width="100%">
      		<iframe name="ifrm" scrolling="Yes" src="" width="100%" height="100%"></iframe>
      	</td>
	</tr>
</table>
</form>	
</body>
</html>
