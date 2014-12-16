// Para utilizar esta libreria hay que agregar la libreria fn_windows.js en cliente de este modulo.

// Descripción: Considera solamente los empleados activos en la ayuda para encontrar un empleado
function buscar_emp_activo_porLeg(legajo){
   if ( isNaN(legajo) || legajo.toString(10).search('E')!=-1 || legajo.toString(10).search('e')!=-1){
     alert('Debe ingresar un número para el legajo.');
   }else{
	  abrirVentanaH('buscar_emp_01.asp?empest=-1' + '&legajo='+ legajo + '&apellido=','',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no' );
   }
}


// Descripción: Considera solamente los empleados inactivos en la ayuda para encontrar un empleado
function buscar_emp_inactivos_porLeg(legajo){
   if ( isNaN(legajo) || legajo.toString(10).search('E')!=-1 || legajo.toString(10).search('e')!=-1){
     alert('Debe ingresar un número para el legajo.');
   }else{
	 abrirVentanaH('buscar_emp_01.asp?empest=0' + '&legajo='+ legajo + '&apellido=','',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
   }
}

//Descripción: Considera todos los empleados en la ayuda para encontrar un empleado
function buscar_emp_todos_porLeg(legajo){
   if ( isNaN(legajo) || legajo.toString(10).search('E')!=-1 || legajo.toString(10).search('e')!=-1){
     alert('Debe ingresar un número para el legajo.');
   }else{
	 abrirVentanaH('buscar_emp_01.asp?empest=' + '&legajo='+ legajo + '&apellido=','',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
   }
}

// Descripción: Considera solamente los empleados activos en la ayuda para encontrar un empleado
function buscar_emp_activo_porApe(apellido){
	abrirVentanaH('buscar_emp_01.asp?empest=-1' + '&legajo=&apellido=' + apellido,'',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no' );
}

// Descripción: Considera solamente los empleados inactivos en la ayuda para encontrar un empleado
function buscar_emp_inactivos_porApe(apellido){
	abrirVentanaH('buscar_emp_01.asp?empest=0' + '&legajo=&apellido=' + apellido,'',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
}

//Descripción: Considera todos los empleados en la ayuda para encontrar un empleado
function buscar_emp_todos_porApe(apellido){
	abrirVentanaH('buscar_emp_01.asp?empest=' + '&legajo=&apellido=' + apellido,'',100,100,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
}

