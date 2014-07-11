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
}