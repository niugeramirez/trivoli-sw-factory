var app = angular.module('turnosApp', []);

app.directive('datepicker', function() {
    return {
        restrict: 'A',
        require : 'ngModel',
        link : function (scope, element, attrs, ngModelCtrl) {
            $(function(){
                element.datepicker({
                	dateFormat: "dd-mm-yy"
						,dayNamesMin: [ "Do", "Lu", "Ma", "Mie", "Jue", "Vie", "Sa" ] 
						,monthNames: [ "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" ]
						,showAnim: "slideDown"	
						,showOn: "both"
						,showOtherMonths: true
						,selectOtherMonths: true
						,	
					onSelect:function (date) {
                        scope.$apply(function () {
                            ngModelCtrl.$setViewValue(date);
                        });
                    }
                });
            });
        }
    };
});

