<% Option Explicit %>
<% 
'Archivo: recursosreservables_con_00.asp
'Descripción: Administración de Médicos
'Autor : Trivoli
'Fecha: 31/05/2015

' Son las listas de parametros a pasarle a los programas de filtro y orden
' En las mismas se deberan poner los valores, separados por un punto y coma

on error goto 0

Dim l_rs
Dim l_sql
  %>

<html>
<head>
<style  type="text/css">
    .title {background-color:#009999;
            color:white;
            text-align:left;
            padding:5px;
            font-weight:bold;}
    .colWidth25{width:25%}
</style>

<link href="/turnos/ess/shared/css/tables_gray.css" rel="StyleSheet" type="text/css">

<title>Administracion de Medicos</title>

<script src="/turnos/shared/js/fn_windows.js"></script>
<script src="/turnos/shared/js/fn_confirm.js"></script>
<script src="/turnos/shared/js/fn_ayuda.js"></script>
<script src="/turnos/shared/js/fn_fechas.js"></script>
<script>

function Buscar(){
	document.datos.filtro.value = "";

	// Apellido
	if (document.datos.inpapellido.value != 0){
		document.datos.filtro.value += " recursosreservables.descripcion like '*" + document.datos.inpapellido.value + "*'";
	}		
    
    window.ifrm.location = 'recursosreservables_con_01.asp?asistente=0&filtro=' + document.datos.filtro.value;
}

function Limpiar(){
	window.ifrm.location = 'recursosreservables_con_01.asp';
}
</script>
</head>

<body leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
    <form name="datos">
    <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%">
        <tr>
            <td align="left">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td class="title">
                            Administracion de Medicos
                        </td>
                    </tr>
                </table>
    		</td>
        </tr>
        <tr>
			<td>
                <input type="hidden" name="filtro" value="">
				<table border="0" width="100%">
                    <colgroup>
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                        <col class="colWidth25">
                    </colgroup>
                    <tbody>
				    <tr>
					    <td><b>Apellido: </b></td>
						<td><input  type="text" name="inpapellido" size="21" maxlength="21" value="" ></td>
					    <td></td>
                        <td align="center">
                            <a class="sidebtnABM" href="Javascript:Buscar();" ><img  src="/turnos/shared/images/Buscar_24.png" border="0" title="Buscar">
                            <a class="sidebtnABM" href="Javascript:Limpiar();" ><img  src="/turnos/shared/images/Limpiar_24.png" border="0" title="Limpiar">
                            <a class="sidebtnABM" href="Javascript:abrirVentana('recursosreservables_con_02.asp?Tipo=A','',650,250);" ><img  src="/turnos/shared/images/Agregar_24.png" border="0" title="Agregar Medico">
                        </td>
                    </tr>
					</tbody>
                </table>
			</td>
		</tr>		
		<tr valign="top" height="100%">
            <td>
      	        <iframe scrolling="yes" name="ifrm" src="recursosreservables_con_01.asp" width="100%" height="100%"></iframe> 
	        </td>
        </tr>		
	</table>
    </form>
</body>

<script>
	Buscar();
</script>
</html>
