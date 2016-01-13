

function eliminarRegistro(obj,donde,sistema)
{
	if (obj.datos.cabnro.value == 0)
		{
		alert("Debe seleccionar un registro para realizar la operación.");
		}
	else
		{
        if ((sistema != null) &&
            (ifrm.jsSelRow.cells(sistema).innerText.toUpperCase() == 'SI'))
            alert('Registro del sistema. No se lo puede eliminar.');
		else
			if (confirm('¿ Desea eliminar el registro seleccionado ?') == true)
 				{
				abrirVentanaH(donde,"",250,120);
  				}
		}
}

function operarRegistro(obj, donde, width, height)
{
	if (obj.datos.cabnro.value == 0)
		{
		alert("Debe selecionar un registro para realizar la operación.");
		}
	else
		{
		abrirVentana(donde,"",width,height);
		}
}
//		window.open("cgi-bin/project1.exe?Tipo=1&Tarea=" + document.tareas.datos.cabnro.value + "&dir=" + "Server.URLEncode(cDirTareas & "\tarea") " + document.tareas.datos.cabnro.value +"\\<%= cDirOriginales %>","","toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=yes,width=300,height=120");
function bajarArchivo(obj, donde, width, height)
{
	if (obj.datos.cabnro.value == 0)
		{
		alert("Debe seleccionar una tarea, para bajar los archivos.");
		}
	else
		{
		if (confirm('¿ Desea bajar los archivos ?') == true)
 			{
			abrirVentanaH(donde,"",250,120);
			}
		}
}

