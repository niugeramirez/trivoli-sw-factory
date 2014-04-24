function LocationController($scope, $location) {
    if($location.$$absUrl.lastIndexOf('/recursos') > 0){
        $scope.activeURL = 'recursos';
    } else{
        $scope.activeURL = 'home';
    }
}