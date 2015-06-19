<% Option Explicit %>
<!--#include virtual="/turnos/shared/inc/sec.inc"-->
<!--#include virtual="/turnos/shared/inc/const.inc"-->
<!--#include virtual="/turnos/shared/db/conn_db.inc"-->
<% 

on error goto 0

  Dim l_rs
  Dim l_sql
  
  Dim l_hd
  Dim l_md
  Dim l_hh
  Dim l_mh
  
%>
<html>
<head>
<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">
<title>Agenda</title>
<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>


<!-- Inicio TOOLTIP -->	
<script type="text/javascript" src="../js/jquery-1.6.js" ></script>
<script type="text/javascript" src="../js/atooltip.jquery.js"></script>
<!-- Final TOOLTIP -->		


<!-- Inicio MULTIPLE SELECCION -->	
<script src="../js/jquery.min.js"></script>
<script src="../js/jquery.sumoselect.js"></script>
<link href="../js/sumoselect.css" rel="stylesheet" />

<script type="text/javascript">
    $(document).ready(function () {
        window.asd = $('.SlectBox').SumoSelect({ csvDispCount: 3 });
        window.test = $('.testsel').SumoSelect({okCancelInMulti:true });
        window.testSelAll = $('.testSelAll').SumoSelect({okCancelInMulti:true, selectAll:true });
        window.testSelAll2 = $('.testSelAll2').SumoSelect({selectAll:true });
    });
</script>
<style type="text/css">
    body{font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;color:#444;font-size:13px;}
    p,div,ul,li{padding:0px; margin:0px;}
</style>
<!-- Final MULTIPLE SELECCION -->



<script>

function Buscar(){
	var tieneotro;
	var estado;
	document.datos.filtro.value = "";
	tieneotro = "no";
	estado = "si";

escribe();
	
	if (document.datos.id.value == "0"){
		alert("Debe ingresar un Medico.");
		document.datos.id.focus();
		return;
	}	
	
	
	if (document.datos.id.value != 0){
		if (tieneotro == "si"){
			document.datos.filtro.value += " AND calendarios.idrecursoreservable = " + document.datos.id.value + "";
		}else{
			document.datos.filtro.value += " calendarios.idrecursoreservable = " + document.datos.id.value + "";
		}
		tieneotro = "si";
	}	

	if (estado == "si"){
		window.ifrm.location = 'Agenda_con_01.asp?id=' + document.datos.id.value + "&hd="+document.datos.hd.value + "&md="+ document.datos.md.value + "&hh="+ document.datos.hh.value + "&mh="+ document.datos.mh.value;
	}
}


function Limpiar(){

	document.datos.fechadesde.value = "<%= date() - 1 %>";
	document.datos.fechahasta.value = "<%= date() %>";

	document.datos.sernro.value     = 0;
	document.datos.pronro.value     = 0;	
/*	
	document.datos.pronro.value     = 0;
	document.datos.ternro.value     = 0;
	document.datos.pornro.value     = 0;
	document.datos.clinro.value     = 0;	
	document.datos.ctrnum.value     = 0;	
	document.datos.txtctrnum.value  = "";	
*/	
	window.ifrm.location = 'buques_con_01.asp';
}

function Contenido(){ 
	if (document.ifrm.datos.cabnro.value == 0) {
		alert("Debe seleccionar un Buque")
		return;
	}		
	else {
		abrirVentana("contenidos_con_00.asp?buqnro=" + document.ifrm.datos.cabnro.value,'',780,580);	
	}		
	
}

function fnctrnum(valor){
	if (valor == 0){
		document.datos.txtctrnum.disabled = true;
		document.datos.txtctrnum.className = "deshabinp";
	}else{
		document.datos.txtctrnum.disabled = false;
		document.datos.txtctrnum.className = "habinp";
	}
}

function TotalVolumen(valor){
	document.datos.totvol.value =  valor;
}


function escribe() {
		alert('1');
         lista = document.all.Med;
		 alert('2');
         opciones = lista.options

         //escribir = document.getElementById("respuesta")
         //escribir.innerHTML = ""
         for (i=0;i<opciones.length;i++) {
              if (opciones[i].selected == true ) {
                 grupos = opciones[i].text
                 //escribir.innerHTML += grupos + "<br/>"
				 alert(grupos)
                 }
              }
         }



</script>
</head>
<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0" onload="Javascript:document.datos.fechadesde.focus();">
      <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr style="border-color :CadetBlue;">
          <td align="left" class="barra">&nbsp;</td>
          <td nowrap align="right" class="barra">
		  		  
          <%' call MostrarBoton ("opcionbtn", "Javascript:abrirVentana('legajos_con_02.asp?Tipo=A','',800,500);","Alta")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:eliminarRegistro(document.ifrm,'legajos_con_04.asp?cabnro=' + document.ifrm.datos.cabnro.value);","Baja")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:abrirVentanaVerif('legajos_con_02.asp?Tipo=M&cabnro=' + document.ifrm.datos.cabnro.value,'',800,500);","Modifica")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("sidebtnSHW", "Javascript:llamadaexcel();","Excel")%>
		  &nbsp;&nbsp;
          <%' call MostrarBoton ("opcionbtn", "Javascript:Contenido();","Contenido")%>		  
		  		  
          <%' call MostrarBoton ("opcionbtn", "Javascript:orden('../../config/contracts_con_01.asp','',490,300);","Orden")%>		  
		  <!--
		  <a class=sidebtn href="Javascript:orden('../../config/contracts_con_01.asp');">Orden</a>
		  -->
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		  
		  </td>
        </tr>		
		<tr>
			<td align="center" colspan="2">
				<table border="0">

					<form name="datos" action="#">
					<input type="hidden" name="filtro" value="" title="dsf">
					
					
<select name="Med" multiple="multiple" placeholder="Todos los Medicos" onchange="console.log($(this).children(':selected').length)" class="testSelAll">
<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
l_sql = "SELECT  * "
l_sql  = l_sql  & " FROM recursosreservables  "
l_sql  = l_sql  & " ORDER BY descripcion "
rsOpen l_rs, cn, l_sql, 0
do until l_rs.eof		%>	
<option value= "<%= l_rs("id") %>" > 
<%= l_rs("descripcion") %> </option>
<%	l_rs.Movenext
loop
l_rs.Close %>
								
<!--       <option selected value="volvo">Volvo</option>
       <option value="saab">Saab</option>
       <option disabled="disabled" value="mercedes">Mercedes</option>
       <option value="audi">Audi</option>
       <option selected value="bmw">BMW</option>
       <option value="porsche">Porche</option>
       <option value="ferrari">Ferrari</option>
       <option value="mitsubishi">Mitsubishi</option> -->


</select>					
					
					
					<tr>

						<td  align="right" nowrap><b>M&eacute;dico: </b></td>
						<td colspan="3"><select name="id" size="1" style="width:200;">
								<option value=0 selected>Seleccionar un M&eacute;dico</option>
								<%Set l_rs = Server.CreateObject("ADODB.RecordSet")
								l_sql = "SELECT  * "
								l_sql  = l_sql  & " FROM recursosreservables  "
								l_sql  = l_sql  & " ORDER BY descripcion "
								rsOpen l_rs, cn, l_sql, 0
								do until l_rs.eof		%>	
								<option value= <%= l_rs("id") %> > 
								<%= l_rs("descripcion") %> </option>
								<%	l_rs.Movenext
								loop
								l_rs.Close %>
							</select>
							<script>document.datos.id.value= "0"</script>
						</td>	

						<td  align="right" nowrap><b>Desde: </b></td>
						<td ><select name="hd" size="1" style="width:50;">
								<%
								l_hd = 0  
								do while clng(l_hd) < 24 %>
								<option value= <%= right("0" & l_hd, 2) %>> <%= right("0" & l_hd, 2) %> </option>
								<%	l_hd = clng(l_hd) + 1
								loop
								%>
							</select>							
						    <b>:</b>
							<select name="md" size="1" style="width:50;">
								<%
								l_md = 0  
								do while clng(l_md) < 60 %>
								<option value= <%= right("0" & l_md, 2) %>> <%= right("0" & l_md, 2) %> </option>
								<%	l_md = clng(l_md) + 15
								loop
								%>
							</select>							
						</td>			
						<td  align="right" nowrap><b>Hasta: </b></td>
						<td ><select name="hh" size="1" style="width:50;">
								<%
								l_hh = 0  
								do while clng(l_hh) < 24 %>
								<option value= <%= right("0" & l_hh, 2) %>> <%= right("0" & l_hh, 2) %> </option>
								<%	l_hh = clng(l_hh) + 1
								loop
								%>
							</select>	
							<script>document.datos.hh.value="23"</script>							
						<b>:</b>
						   <select name="mh" size="1" style="width:50;">
								<%
								l_mh = 0  
								do while clng(l_mh) < 60 %>
								<option value= <%= right("0" & l_mh, 2) %>> <%= right("0" & l_mh, 2) %> </option>
								<%	l_mh = clng(l_mh) + 15
								loop
								%>
							</select>							
						</td>							
						<td ><a class="sidebtnABM" href="Javascript:Buscar();" ><img class="normaltip"  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar"></a></td>
						<td ><a class="sidebtnABM" href="Javascript:Limpiar();"><img class="normaltip" src="/turnos/shared/images/Limpiar_24.png" border="0" title="Limpiar"></a></td>
						<td >&nbsp;&nbsp;</td>
								
					</tr>
			

				</table>
			</td>
		</tr>		
		
        <tr valign="top" height="100%">
          <td colspan="2" style="" width="100%">
      	  <iframe scrolling="yes" name="ifrm" src="" width="100%" height="100%"></iframe> 
	      </td>
        </tr>		
			</form>		
      </table>
</body>

<script>
	//Buscar();
</script>
</html>
