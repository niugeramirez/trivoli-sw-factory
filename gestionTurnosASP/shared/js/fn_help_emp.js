// Para utilizar esta libreria hay que agregar la libreria fn_windows.js en cliente de este modulo.

// Descripción: Considera solamente los empleados activos en la ayuda para encontrar un empleado
function help_emp_activos(){
	abrirVentana('help_emp_01.asp?empest=-1','',700,400,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no' );
}


// Descripción: Considera solamente los empleados inactivos en la ayuda para encontrar un empleado
function help_emp_inactivos(){
	abrirVentana('help_emp_01.asp?empest=0','',700,400,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
}


//Descripción: Considera todos los empleados en la ayuda para encontrar un empleado
function help_emp_todos(){
	abrirVentana('help_emp_01.asp?empest=','',700,400,'toolbar=no,location=no,directories=no,satus=no,menubar=no,scrollbars=no,resizable=no');
}

 
 