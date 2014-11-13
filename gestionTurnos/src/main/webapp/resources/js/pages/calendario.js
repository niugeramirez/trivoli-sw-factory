function controller($scope, $http) {
	
	// Pagina solicitada al Backend
    $scope.nroPaginaPacientes = 0;

    // Estado Actual de la Vista
    $scope.estado = 'busy';

    // Ultima Accion solictada por el Usuario
    $scope.ultimaAccion = '';

    // URL base de la Vista
    $scope.url = "/gestionTurnos/protected/calendarios/";

    // Flags diversos que manejan la interacción del Usuario con la Vista
    $scope.errorSubmit = false;
    $scope.errorAccesoIlegal = false;
    $scope.mostrarMensajesUsuario = false;
    $scope.mostrarErrorValidacion = false;
    $scope.mostrarErrorValidacionOS = false;
    
    
    $scope.pagina={};
    $scope.pagina.mensajeAccion = "";
    $scope.pagina.mensajeBusqueda = '';
    
    //filtros
    $scope.filtroPaciente={};
    $scope.filtroPaciente.dni					=	'';
    $scope.filtroPaciente.nombre				=	'';
    $scope.filtroPaciente.apellido				=	'';
    //estos campos no son filtros pero se usan para inicializacion en la creacion
    $scope.filtroPaciente.nroHistoriaClinica	=	'';
    $scope.filtroPaciente.obraSocial			=	{};
    $scope.filtroPaciente.telefono				=	'';

    // Objetos JSON
    $scope.calendariosActuales 	= 	[]; 
    $scope.calendarioActual 	= 	null;   
    $scope.turnosActuales 		= 	[];
    $scope.recursosActuales 	= 	{};
    $scope.recursoActual		= 	{};    
    $scope.pacientes			=	[];
    $scope.obrasSociales 		= 	{};
    
    // Filtro de Busqueda   
    fecha = 	new Date(2014,8,9);
    //fecha = 	new Date();
    $scope.fechaActual = 	fecha.getDate().toString() + '-'+
    						(fecha.getMonth()+1)  + '-'+
    						fecha.getFullYear().toString();
        

    

    
    // Definición de Funciones del Controlador de la Página de Administración de Turnos
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.Inicializar = function () {
    	// Se obtiene la URL actual
        var url = $scope.url;
        //Se setea el modal de carga AJAX
		loadingId = "#loadingModalRecursos";
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
		$scope.startDialogAjaxRequest(loadingId);           
        
		var config = {};		
		config.params = {};
		config.params.nroPagina = 0;
		
		//Parametros propios de esta llamada AJAX    
		//NO HAY PARAMETROS PROPIOS EN ESTA LLAMADA

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxRecursos(data,loadingId);    
                $scope.listarCalendario();
            })
            .error(function () {
            	$scope.errorAjax();
            });
        
    };
/************************************************************************************************************************************************************************/    
    // Funcion que recupera del backend todos los Recursos
    $scope.listarCalendario = function () {
    	// Se obtiene la URL actual
        var url = $scope.url +  $scope.recursoActual.id;
        //Se setea el modal de carga AJAX
		loadingId = "#loadingModalCalendario";
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
		$scope.startDialogAjaxRequest(loadingId);        

        // Se fijan los parámetros de la llamada al servicio Rest 
		var config = {};		
		config.params = {};
		config.params.nroPagina = 0;
		
		//Parametros propios de esta llamada AJAX        
        config.params.idRecurso = $scope.recursoActual.id;
        config.params.fechaTurnos = $scope.fechaActual;
        

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxCalendario(data,loadingId);    
                $scope.buscarTurnos();
            })
            .error(function () {
            	$scope.errorAjax();
            });
        
    };
/************************************************************************************************************************************************************************/

    // Funcion que recupera del backend todos los Recursos
    $scope.buscarTurnos = function () {    
	    
		if	($scope.calendarioActual!=null) {
			// Se obtiene la URL actual
			var url = $scope.url + $scope.calendarioActual.recurso.id +'/'+ $scope.calendarioActual.id;
			//Se setea el modal de carga AJAX
			loadingId = "#loadingModalTurno";
			
	        // Se fija la Ultima Accion Solicitada por el Usuario
	        $scope.ultimaAccion = 'list';

	        // Se abre el dialogo Loading .....
			$scope.startDialogAjaxRequest(loadingId);        

	        // Se fijan los parámetros de la llamada al servicio Rest 			
			var config = {};		
			config.params = {};
			config.params.nroPagina = 0;
			
			//Parametros propios de esta llamada AJAX
			config.params.idCalendario = $scope.calendarioActual.id;
			config.params.idRecurso = $scope.calendarioActual.recurso.id;
	  
	        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados asincronicamente) 			
			$http.get(url, config)
			    .success(function (data) {			    	
			    	 $scope.finishAjaxTurnos(data,loadingId); 
			    })
			    .error(function() {
			    	$scope.errorAjax();
			    });    		   
		}
	};

/************************************************************************************************************************************************************************/    
    $scope.buscarPacientes = function () {
	    //TODO ver la posibilidad de definir la pagina de busqueda de paciente como una pagina aparte con su JS separado y que "devuelva" el paciente seleccionado de modo de re usar por ejemplo en el ABM de pacientes                
	
	    $scope.ultimaAccion = 'search';
	
	    var url = $scope.url +  'pacientes';
	    //Se setea el modal de carga AJAX
	    //TODO hacer que el loading ajax funcione para los dialogos modales
	    loadingId = "#loadingModalPacientes";
	             
	     // Se abre el dialogo Loading .....
	     $scope.startDialogAjaxRequest(loadingId);    
	
	    //Parametros propios de esta llamada AJAX
	    var config = {};
	    config.params = {};
	    config.params.nroPagina = $scope.nroPaginaPacientes;
	     if($scope.filtroPaciente){
	         config.params.filtroDNI = $scope.filtroPaciente.dni;                    
	         config.params.filtroNombre = $scope.filtroPaciente.nombre;
	         config.params.filtroApellido = $scope.filtroPaciente.apellido;            
	     }
	    
	
	    $http.get(url, config)
	        .success(function (data) {
	                     $scope.finishAjaxPacientes(data,loadingId);
	        })
	        .error(function() {
	                $scope.errorAjax();
	        });
    };
   
    /************************************************************************************************************************************************************************/
    $scope.guardarPaciente = function (form) {
        
    	//TODO la primitiva required no funciona para el tag <SELECT>, con el chequeo del objeto bindeado zafamos pero el span connel mensaje de requerida para la OS no
    	//queda dinamico (es decir, no desaparece ni bien se completa el select). el todo es buscar una solucion mas general y elegante
    	if (!form.$valid || !$scope.pacienteActual.obraSocial.id) {
            $scope.mostrarErrorValidacion = true;
            
            if (!$scope.pacienteActual.obraSocial.id){
            	$scope.mostrarErrorValidacionOS = true;
            }
            else{
            	$scope.mostrarErrorValidacionOS = false;	
            }
            
            return;
        }    	
    
    	//Para mejorar la seleccion de un paciente cuando el mismo es recien creado, chequeo si hay filtros seteadod, si no hay nada fuerzo el filtro dni de modo que
    	//luego d ela creacion solo se muestre este resultado de busqueda y el usuario lo pueda seleccionar facilmente
    	if ($scope.modoEditCreate=='create'
    			&& !$scope.filtroPaciente.dni
    			&& !$scope.filtroPaciente.nombre
    			&& !$scope.filtroPaciente.apellido
    		){
    		$scope.filtroPaciente.dni=$scope.pacienteActual.dni;
    	}
    	
        $scope.ultimaAccion = 'update';

        var url = $scope.url +  'pacientes';
	    //TODO hacer que el loading ajax funcione para los dialogos modales
	    loadingId = "#loadingModalPacientes";
	             
	     // Se abre el dialogo Loading .....
	     $scope.startDialogAjaxRequest(loadingId);    
	
	    //Parametros propios de esta llamada AJAX
	    var config = {};
	    config.params = {};
	    config.params.nroPagina = $scope.nroPaginaPacientes;
	     if($scope.filtroPaciente){
	         config.params.filtroDNI = $scope.filtroPaciente.dni;                    
	         config.params.filtroNombre = $scope.filtroPaciente.nombre;
	         config.params.filtroApellido = $scope.filtroPaciente.apellido;            
	     }
	
	      $http.put(url,$scope.pacienteActual, config)
	        .success(function (data) {
	                     $scope.finishAjaxUpdatePacientes(data,loadingId);
	        })
	        .error(function(data, status, headers, config) {
	                $scope.errorAjaxQuickEditCreate(status,data);	                
	        });
    };    
/************************************************************************************************************************************************************************/    
	    // Funcion que recupera del backend todos los Recursos
	    $scope.buscarObrasSociales = function () {
		   
	    	$scope.ultimaAccion = 'search';
		
		   var url = $scope.url +  'obrasSociales';
		   //Se setea el modal de carga AJAX
		   //TODO hacer que el loading ajax funcione para los dialogos modales
		   loadingId = "#loadingModalObrasSociales";
		
		   // Se abre el dialogo Loading .....
		   $scope.startDialogAjaxRequest(loadingId);    
		
			//Parametros propios de esta llamada AJAX
		   var config = {};
		   config.params = {};
		   config.params.nroPagina = 0;	       
		
		   $http.get(url, config)
		       .success(function (data) {
		       		$scope.finishAjaxObrasSociales(data,loadingId);
		       })
		       .error(function() {
		    	   $scope.errorAjax();
		       });	        
	    };	   
	
/************************************************************************************************************************************************************************/
    
	 $scope.finishAjaxTurnos = function (data,loadingId) {  
    	
        $scope.turnosActuales   = data.registros;

        $scope.finishAjaxGral(loadingId,data);
    };
    
/************************************************************************************************************************************************************************/
    
	 $scope.finishAjaxObrasSociales = function (data,loadingId) {  
  	
		$scope.obrasSociales = data.registros;		     
		
		//El binding con el modelo en el selct donde se usan estos datos funciona correctamente, salvo que no inicializa correctamente el selct con el modelo actual
		//esto es porque las referencias del array son distintas a pesar de que apunta a objetos iguales (con las mismas propiedades)
		//Por tal motivo hago este bucle donde modifico el modelo de modo que apunte al objeto del array correcto
		//TODO buscar una manera gral de bindear con el modelo siempre que se utilicen objetos y arrays
		for (i in $scope.obrasSociales) {
			
			if ($scope.obrasSociales[i].id == $scope.pacienteActual.obraSocial.id){
				$scope.pacienteActual.obraSocial = $scope.obrasSociales[i];
			}
		}  		
		
		$scope.finishAjaxGral(loadingId,data);
	 };
	 /************************************************************************************************************************************************************************/
	    
	 $scope.finishAjaxUpdatePacientes = function (data,loadingId) {  

		$("#pacienteQuickEditCreate").modal('hide');
		 $scope.errorSubmit = false;
		 
		$scope.finishAjaxPacientes(data,loadingId);
   };
   /************************************************************************************************************************************************************************/
    
	 $scope.finishAjaxPacientes = function (data,loadingId) {  
   	
       //lleno la tabla que se va a mostrar
	   $scope.pacientes   = data.registros;
       
	   //seteos generales para mostrar controles
       if (data.cantPaginas > 0) {
           $scope.estadoPacientes = 'list'; 
       }
       else {
    	   $scope.estadoPacientes = 'noresult';
       }       
       
       //Seteo de visualizacion de filtros (uso variables donde repito el valor del modelo para que no se actualizen con la edicion)
       $scope.mostrarFiltroDNI = $scope.filtroPaciente.dni;  
       $scope.mostrarFiltroApellido = $scope.filtroPaciente.apellido;
       $scope.mostrarFiltroNombre = $scope.filtroPaciente.nombre;
       
       $scope.paginaPacientes = {paginaActual: $scope.nroPaginaPacientes, cantPaginas: data.cantPaginas, totalRegistros : data.totalRegistros};
       
       $scope.finishAjaxGral(loadingId,data);
   };
/************************************************************************************************************************************************************************/
     
    $scope.finishAjaxRecursos = function (data,loadingId) {   	
        
    	$scope.recursosActuales = data.registros;
        $scope.recursoActual =data.registros[0];        
        
        $scope.finishAjaxGral(loadingId,data);
    };    

/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxCalendario = function (data,loadingId) { 	     
        
        if (data.totalRegistros > 0) {            
        	$scope.calendarioActual = data.registros[0].calendario;   
        }    
        else	{
        	$scope.calendarioActual = null;     
        }
        $scope.calendariosActuales = data.registros;
        
        $scope.finishAjaxGral(loadingId,data);
    };   
/************************************************************************************************************************************************************************/
    
    $scope.finishAjaxGral = function (loadingId,data) { 	             
        
        $scope.estado = $scope.estadoAnterior;
        $(loadingId).modal('hide');      
        $scope.ultimaAccion = '';
        $scope.pagina.mensajeAccion = data.mensajeAccion;
        $scope.pagina.mensajeBusqueda = data.mensajeBusqueda;           
    };      
    /************************************************************************************************************************************************************************/
    
    $scope.startDialogAjaxRequest = function (loadingId) {
        //TODO ver la posibilidad que quede el modal de ajax unificado para todas las llamadas
    	//TODO ver la posibilidad de usar distinto Scopes para los distintos niveles de llamadas AJAX
    	//TODO ver la unificacion de codigo javascript a nivel general (una libreria q se llame desde todos los JS)

    	//$scope.mostrarErrorValidacion = false; //esto no se si va
        $(loadingId).modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
       
    /************************************************************************************************************************************************************************/     
    $scope.seleccionarCalendario = function (registroActual) {
     
        // Se copia el objeto JSON seleccionado en la grilla al registro actual
        $scope.calendarioActual = angular.copy(registroActual.calendario);                              
        $scope.buscarTurnos();
        
    };
    /************************************************************************************************************************************************************************/    

    $scope.errorAjaxQuickEditCreate = function (status,data) {
        
    	//TODO Crear una manera de mostrar los mensajes de error uniforme para toda la pagina 
        $scope.estado = $scope.estadoAnterior;

        $scope.errorSubmit = true;
        $scope.ultimaAccion = '';
        $scope.mensajeError = data;
    };
    
    /************************************************************************************************************************************************************************/
    
	$scope.errorAjax = function () {  
        $scope.estado = 'error';     
   };

/************************************************************************************************************************************************************************/    
   $scope.cambiarPaginaPacientes = function (pagina) {
       
	   $scope.nroPaginaPacientes = pagina;
       $scope.buscarPacientes();
       
   };   
/************************************************************************************************************************************************************************/
   
   $scope.resetearBusqueda = function(filtro){
       if (filtro == 'dni') {
    	   $scope.mostrarFiltroDNI = '';
    	   $scope.filtroPaciente.dni = '';
       }

       if (filtro == 'apellido') {
	       $scope.mostrarFiltroApellido = '';
	       $scope.filtroPaciente.apellido= '';
       }
       
       if (filtro == 'nombre') {
	       $scope.mostrarFiltroNombre= '';
	       $scope.filtroPaciente.nombre= '';     
       }
       
       $scope.nroPaginaPacientes=0;
       $scope.buscarPacientes();
       
   };
   /************************************************************************************************************************************************************************/
   $scope.exitQuickEditCreate = function (modalId) {
       
       $scope.pacienteActual = {};
       $scope.errorSubmit = false;
       $scope.mostrarErrorValidacion = false;
       $scope.mostrarErrorValidacionOS = false;
       $("#pacienteQuickEditCreate").modal('hide');
   };   
   /************************************************************************************************************************************************************************/
   $scope.exit = function (modalId) {
       $(modalId).modal('hide');
       
       $scope.pacientes = [];
       $scope.paginaPacientes = {};
       
	   $scope.mostrarFiltroDNI = '';
	   $scope.filtroPaciente.dni = '';
	   $scope.mostrarFiltroApellido = '';
	   $scope.filtroPaciente.apellido= '';
	   $scope.mostrarFiltroNombre= '';
	   $scope.filtroPaciente.nombre= '';     
	   $scope.nroPaginaPacientes=0;
	   
	   $scope.errorSubmit = false;
   };
   /************************************************************************************************************************************************************************/
   
       $scope.quickEditCreatePaciente = function (paciente,modo) {
        
    	$scope.modoEditCreate = modo;
    	
        // Se copia el objeto JSON seleccionado en la grilla al registro actual
        if (modo=='create'){
        	$scope.pacienteActual = angular.copy($scope.filtroPaciente);
        } else {
        	$scope.pacienteActual = angular.copy(paciente);
        }
        	
        
        //TODO hacer que se oculten los modal de busqueda y que solo quede activo el de edicion/creacion
        $("#pacienteSearchParameters").modal('hide');
        $("#subPacienteSearchResult").modal('hide');
        $("#busquedaGral").modal('hide');
        
        $scope.buscarObrasSociales();
    };    
   /************************************************************************************************************************************************************************/    
   // Codigo de Inicializacion del Controlador de la Página de Administración de Obras Sociales
    $scope.Inicializar();        
       
}
