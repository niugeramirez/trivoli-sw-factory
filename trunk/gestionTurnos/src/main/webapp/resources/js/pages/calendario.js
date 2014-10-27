function controller($scope, $http) {
	// Se define el Modelo de la Página de Administración de Obras Sociales
$scope.button = 'red';	
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
    
    
    $scope.pagina={};
    $scope.pagina.mensajeAccion = "";
    $scope.pagina.mensajeBusqueda = '';
    
    $scope.filtroPaciente={};
    $scope.filtroPaciente.DNI='';
    $scope.filtroPaciente.nombre='';
    $scope.filtroPaciente.apellido='';

    // Objetos JSON
    $scope.calendariosActuales 	= 	[]; 
    $scope.calendarioActual 	= 	null;   
    $scope.turnosActuales 		= 	[];
    $scope.recursosActuales 	= 	{};
    $scope.recursoActual		= 	{};    
    $scope.pacientes			=	[];
    
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
	            config.params.filtroDNI = $scope.filtroPaciente.DNI;
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
    
	 $scope.finishAjaxTurnos = function (data,loadingId) {  
    	
        $scope.turnosActuales   = data.registros;

        $scope.finishAjaxGral(loadingId,data);
    };
/************************************************************************************************************************************************************************/
    
	 $scope.finishAjaxPacientes = function (data,loadingId) {  
   	
       $scope.pacientes   = data.registros;
       
       if (data.cantPaginas > 0) {
           $scope.estadoPacientes = 'list'; 
       }
       else {
    	   $scope.estadoPacientes = 'noresult';
       }       
       
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
    	$scope.mostrarErrorValidacion = false;
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
    
	$scope.errorAjax = function () {  
        $scope.estado = 'error';     
   };

/************************************************************************************************************************************************************************/    
   $scope.cambiarPaginaPacientes = function (pagina) {
       
	   $scope.nroPaginaPacientes = pagina;
       $scope.buscarPacientes();
       
   };   
/************************************************************************************************************************************************************************/    
   // Codigo de Inicializacion del Controlador de la Página de Administración de Obras Sociales
    $scope.Inicializar();    
    


       
}
