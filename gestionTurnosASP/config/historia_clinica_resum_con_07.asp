<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Vista Preliminar Historia Clinica</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script src="/turnos/shared/js/fn_ay_generica.js"></script>


<script>

<% on error goto 0
Dim l_rs
Dim l_sql
Dim l_id

l_id = request("cabnro")
%>

function Imprimir(){
	document.ifrm.focus();
	window.print();	
}

function Actualizar(destino){

	var param;
	//Fechas	
	
	
	
	
	param = "&id=" + document.all.id.value; // + document.all.repnro.value;
	
	if (destino== "exel")
    	abrirVentana("historia_clinica_resum_con_08.asp?" + param + "&excel=true",'execl',250,150);
	else
		document.ifrm.location = "historia_clinica_resum_con_08.asp?" + param;			
	
}

function Ayuda_Fecha(txt){
	var jsFecha = Nuevo_Dialogo(window, '/turnos/shared/js/calendar.html', 16, 15);
	if (jsFecha == null){
		//txt.value = '';
	}else{
		txt.value = jsFecha;
		//DiadeSemana(jsFecha);
	}
}



</script>

</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fecini.focus();" >
<form name="datos">
<input type="hidden" name="id" value="<%= l_id %>">
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="center" class="barra" >
			<a class=sidebtnSHW href="Javascript:Imprimir()"><img  src="/turnos/shared/images/eliminarvisita_24.png" border="0" title="Imprimir"></a> 
			&nbsp;
			
		</td>
	</tr>
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<input type="hidden" name="filtro" value="">
				</table>
			</td>
		</tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="historia_clinica_resum_con_08.asp?id=<%= l_id %>" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
</table>

</form>	
</body>
</html>
