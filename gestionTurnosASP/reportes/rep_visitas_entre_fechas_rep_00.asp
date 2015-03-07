<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Visitas entre Fechas</title>
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

l_id = 0
%>

function Imprimir(){
	document.ifrm.focus();
	window.print();	
}

function Actualizar(destino){

	var param;
	//Fechas	
	
	
	if (document.datos.fechadesde.value == "")  {
  		alert("Debe ingresar la Fecha Desde ");
  		document.datos.fechadesde.focus();
		return;
	}
	
	if (document.datos.fechahasta.value == "")  {
  		alert("Debe ingresar la Fecha Hasta ");
  		document.datos.fechahasta.focus();
		return;
	}	
	
	if (!menorque(document.datos.fechadesde.value ,document.datos.fechahasta.value)){
		alert('La fecha hasta debe ser mayor que la fecha desde.');
  		document.datos.fechadesde.focus();
		return;
	}		

	/*
	if (document.datos.fecini.value != "") {
	
			if (!validarfecha(document.datos.fechadesde)) {
		  		document.datos.fechadesde.focus();
				return;
			}	
	}
	*/

	if (document.datos.pacienteid.value == "0")  {
  		alert("Debe ingresar el Paciente ");
  		document.datos.pacienteid.focus();
		return;
	}	
	
	param = "qfechadesde=" + document.all.fechadesde.value + "&qfechahasta=" + document.all.fechahasta.value + "&idpaciente=" + document.all.pacienteid.value; // + document.all.repnro.value;
	
	if (destino== "exel")
    	abrirVentana("rep_visitas_entre_fechas_rep_01.asp?" + param + "&excel=true",'execl',250,150);
	else
		document.ifrm.location = "rep_visitas_entre_fechas_rep_01.asp?" + param;			
	
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
<table border="0" cellpadding="0" cellspacing="0" height="100%">
	<tr style="border-color :CadetBlue;">
		<td align="left" class="barra" nowrap>
			<!--<a class=sidebtnSHW href="Javascript:window.close();">Salir</a>--></td>
		<td align="right" class="barra" >
			<a class=sidebtnSHW href="Javascript:Actualizar('ifrm')">Actualizar</a>		  
			<!--<a class=sidebtnSHW href="Javascript:Imprimir()">Imprimir</a>	-->	  
			<a class=sidebtnSHW href="Javascript:Actualizar('exel')">Excel</a> 
			&nbsp;
			
		</td>
	</tr>
		<tr>
			<td align="center" colspan="2">
				<table border="0">
					<input type="hidden" name="filtro" value="">

					<tr>
						<td align="right"><b>Fecha Desde: </b></td>
						<td><input  type="text" name="fechadesde" size="10" maxlength="10" value="<%'= l_fecha%>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechadesde);"><img src="/turnos/shared/images/calendar1.png" border="0"></a>
						</td>
						<td align="right"><b>Fecha Hasta: </b></td>
						<td><input  type="text" name="fechahasta" size="10" maxlength="10" value="<%'= l_fecha%>" >
							<a href="Javascript:Ayuda_Fecha(document.datos.fechahasta);"><img src="/turnos/shared/images/calendar1.png" border="0"></a>
						</td>						
						<td  align="right" nowrap><b>Paciente: </b></td>
						<td colspan="3"><select name="pacienteid" size="1" style="width:200;">
								<option value=0 selected>Seleccione un Paciente</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM clientespacientes "
								l_sql  = l_sql  & " ORDER BY apellido "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("apellido") %>&nbsp;<%= l_rs("nombre") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
						</td>							
					</tr>	

				</table>
			</td>
		</tr>
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
</table>
</form>	
</body>
</html>
