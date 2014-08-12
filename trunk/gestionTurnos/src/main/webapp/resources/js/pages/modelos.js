function modelosController($scope, $http) {
	// Se define el Modelo de la Página de Administración de Modelos
	
	// Pagina solicitada al Backend
    $scope.nroPagina = 0;

    // Estado Actual de la Vista
    $scope.estado = 'busy';

    // Ultima Accion solictada por el Usuario
    $scope.ultimaAccion = '';

    // URL base de la Vista
    $scope.url = "/gestionTurnos/protected/modelos/";

    // Flags diversos que manejan la interacción del Usuario con la Vista
    $scope.errorSubmit = false;
    $scope.errorAccesoIlegal = false;
    $scope.mostrarMensajesUsuario = false;
    $scope.mostrarErrorValidacion = false;
    $scope.mostrarMensajeBusqueda = false;
    $scope.mostrarBotonBuscar = false;
    $scope.mostrarBotonCrear = false;

    // Objeto JSON que almacena el Modelo actual
    $scope.modelo = {};

    // Filtro de Busqueda
    $scope.filtroDescripcion = "";

    
    
    
    // Definición de Funciones del Controlador de la Página de Administración de Modelos
    
    // Funcion que recupera del backend todos los Modelos
    $scope.listarTodosLosModelos = function () {
    	// Se obtiene la URL actual
        var url = $scope.url;
        
        // Se fija la Ultima Accion Solicitada por el Usuario
        $scope.ultimaAccion = 'list';

        // Se abre el dialogo Loading .....
        $scope.startDialogAjaxRequest();

        // Se fijan los parámetros de la llamada al servicio Rest (Página Solicitada por el Usuario)
        var config = {params: {nroPagina: $scope.nroPagina}};

        // Se realiza un requerimiento HTTP a través de un método GET esperando 2 posibles resultados (callbacks invocados
        // asincronicamente) 
        $http.get(url, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, null, false);
            })
            .error(function () {
                $scope.estado = 'error';
                $scope.mostrarBotonCrear = false;
            });
    };
    
    
    // Funcion que Actualiza un Modelo en el Backend
    $scope.updateModelo = function (updateModeloForm) {
        if (!updateModeloForm.$valid) {
            $scope.mostrarErrorValidacion = true;
            return;
        }

        $scope.ultimaAccion = 'update';

        var url = $scope.url + $scope.modelo.id;

        $scope.startDialogAjaxRequest();

        var config = {};

        $scope.agregarFiltroBusqueda(config);

        $http.put(url, $scope.modelo, config)
            .success(function (data) {
                $scope.finishAjaxCallOnSuccess(data, "#editarModeloDialog", false);
            })
            .error(function(data, status, headers, config) {
                $scope.handleErrorInDialogs(status);
            });
    };
    
    $scope.startDialogAjaxRequest = function () {
        $scope.mostrarErrorValidacion = false;
        $("#loadingModal").modal('show');
        $scope.estadoAnterior = $scope.estado;
        $scope.estado = 'busy';
    };
    
    
    $scope.handleErrorInDialogs = function (status) {
        $("#loadingModal").modal('hide');
        
        $scope.estado = $scope.estadoAnterior;

        // Acceso No Permitido
        if(status == 403){
            $scope.errorAccesoIlegal = true;
            return;
        }

        $scope.errorSubmit = true;
        $scope.ultimaAccion = '';
    };
    
    $scope.finishAjaxCallOnSuccess = function (data, modalId, isPagination) {
    	// Se muestran los datos en la Grilla de Modelos
        $scope.populateTable(data);
        
        $("#loadingModal").modal('hide');

        if(!isPagination){
            if(modalId){
                $scope.exit(modalId);
            }
        }

        $scope.ultimaAccion = '';
    };
    
    $scope.populateTable = function (data) {
    	// Se diferencian los casos de Respuesta con datos vs. Respuesta vacía
        if (data.cantPaginas > 0) {
            $scope.estado = 'list';

            // Se define la Pagina de Datos a mostrar
            $scope.pagina = {registros: data.registros, paginaActual: $scope.nroPagina, cantPaginas: data.cantPaginas, totalRegistros : data.totalRegistros};

            if($scope.pagina.cantPaginas <= $scope.pagina.paginaActual){
                $scope.nroPagina = $scope.pagina.cantPaginas - 1;
                $scope.pagina.paginaActual = $scope.pagina.cantPaginas - 1;
            }

            $scope.mostrarBotonCrear = true;
            $scope.mostrarBotonBuscar = true;
        } else {
            $scope.estado = 'noresult';
            $scope.mostrarBotonCrear = true;

            if(!$scope.filtroDescripcion){
                $scope.mostrarBotonBuscar = false;
            }
        }

        if (data.mensajeAccion || data.mensajeBusqueda) {
            $scope.mostrarMensajesUsuario = $scope.ultimaAccion != 'search';

            $scope.pagina.mensajeAccion = data.mensajeAccion;
            $scope.pagina.mensajeBusqueda = data.mensajeBusqueda;
        } else {
            $scope.mostrarMensajesUsuario = false;
        }
    };
    
     
    $scope.modeloSeleccionado = function (modelo) {
        var modeloSeleccionado = angular.copy(modelo);
        
        // Se copia el objeto JSON seleccionado en la grilla al Modelo actual
        $scope.modelo = modeloSeleccionado;
    };
    
    
    $scope.resetModelo = function(){
        $scope.modelo = {};
    };
    
    
    $scope.exit = function (modalId) {
        $(modalId).modal('hide');
        
        $scope.resetModelo();
        
        $scope.errorSubmit = false;
        $scope.errorAccesoIlegal = false;
        $scope.mostrarErrorValidacion = false;
    };
    
    
    $scope.resetearBusqueda = function(){
        $scope.filtroDescripcion = "";
        $scope.nroPagina = 0;
        $scope.listarTodosLosModelos();
        $scope.mostrarMensajeBusqueda = false;
    };
    
    
    $scope.cambiarPagina = function (pagina) {
        $scope.nroPagina = pagina;

        if($scope.filtroDescripcion){
            $scope.buscarModelos($scope.filtroDescripcion, true);
        } else{
            $scope.listarTodosLosModelos();
        }
    };
    
    // Codigo de Inicializacion del Controlador de la Página de Administración de Modelos
    $scope.listarTodosLosModelos();
}