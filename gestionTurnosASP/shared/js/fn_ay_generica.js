
function verificacodigo(codigo, descrip, campocod, campodest, tabla,where)
{
	if (where == null)
		where = "";
	if (codigo.value != "")
	{
		if (isNaN(codigo.value)) {
		     descrip.value = '';
		     codigo.value = '';
		     alert('Código incorrecto')
		     codigo.focus();
		     codigo.select();	 
		}
		else {
		 var jsNuevo = Nuevo_DialogoH(window, "verifica_codigo.asp?codigo=" + codigo.value+"&campocod=" + campocod +"&campodest=" 
		 								+ campodest +"&tabla=" + tabla+"&where="+where, 1, 1);
		 
		 if ((jsNuevo == null) ||
		     (jsNuevo == "")) 
		   {
		     descrip.value = '';
		     codigo.value = '';
		     alert('Código inexistente')
		     codigo.focus();
		     codigo.select();	 
		   }
		 else 
		   {
		 	 descrip.value = jsNuevo.substr(0);
		   } 
   		 }
	}
	else
		descrip.value = "";
}

function ayudacodigo(objcodigo, objdescrip, campocod, campodest, tabla, where, titcols, titulo)
{
	 var jsNuevo = Nuevo_Dialogo(window, "dialog_frame.asp?titulo=" + titulo + "&pagina=ayuda_codigo.asp&titcols=" + titcols + "&tabla=" + tabla + "&where=" + where +"&campocod=" + campocod +"&campodest=" + campodest, 25, 25);
	 
	 if ((jsNuevo != null) &&
	     (jsNuevo != "")) 
	   {
	     var c = jsNuevo.substr(0, jsNuevo.indexOf("__,"));
	     var d = jsNuevo.substr(jsNuevo.indexOf("__,") + 3, jsNuevo.length);
		 
	     objcodigo.value  = c;
		 objdescrip.value = d;
	     //objcodigo.focus();
	     objcodigo.select();	 
	   }
}

function verificacodigochar(codigo, descrip, campocod, campodest, tabla)
{
	if (codigo.value != "")
	{
		var jsNuevo = Nuevo_DialogoH(window, "verifica_codigoChar.asp?codigo=" + codigo.value+"&campocod=" + campocod +"&campodest=" + campodest +"&tabla=" + tabla, 1, 1);
		 
		 if ((jsNuevo == null) ||
		     (jsNuevo == "")) 
		   {
		     descrip.value = '';
		     codigo.value = '';
		     alert('Código inexistente')
		     codigo.focus();
		     codigo.select();	 
		   }
		 else 
		   {
		 	 descrip.value = jsNuevo.substr(0);
		   }  
	}
	else
		descrip.value = "";
}

function ayudacodigochar(objcodigo, objdescrip, campocod, campodest, tabla, where, titcols, titulo)
{
	 var jsNuevo = Nuevo_Dialogo(window, "dialog_frame.asp?titulo=" + titulo + "&pagina=ayuda_codigoChar.asp&titcols=" + titcols + "&tabla=" + tabla + "&where=" + where +"&campocod=" + campocod +"&campodest=" + campodest, 25, 25);
	 
	 if ((jsNuevo != null) &&
	     (jsNuevo != "")) 
	   {
	     var c = jsNuevo.substr(0, jsNuevo.indexOf("__,"));
	     var d = jsNuevo.substr(jsNuevo.indexOf("__,") + 3, jsNuevo.length);
		 
	     objcodigo.value  = c;
		 objdescrip.value = d;
	     objcodigo.focus();
	     objcodigo.select();	 
	   }
}
