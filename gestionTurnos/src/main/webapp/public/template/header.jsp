<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ taglib prefix="fn" uri="http://java.sun.com/jsp/jstl/functions" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<div>
    <h3 class="text-muted">
        <spring:message code='header.message'/>
    </h3>
    
    <div class="navbar navbar-default" role="navigation">
        <div class="container-fluid">
          <div class="navbar-header">
            <button type="button" class="navbar-toggle" data-toggle="collapse" data-target=".navbar-collapse">
              <span class="sr-only">Desplegar Navegacion</span>
              <span class="icon-bar"></span>
              <span class="icon-bar"></span>
              <span class="icon-bar"></span>
            </button>
            <a class="navbar-brand" href="<c:url value="/"/>" title='<spring:message code="header.home"/>'>
            	<p><spring:message code="header.home"/></p>
            </a>
          </div>
		  <div class="navbar-collapse collapse">
		    <ul class="nav navbar-nav" ng-controller="LocationController">
		    	<li ng-class="{'active': activeURL == 'recursos', '': activeURL != 'recursos'}">
		        	<a title='<spring:message code="header.recursos"/>' href="<c:url value='/protected/recursos'/>">
		            	<p><spring:message code="header.recursos"/></p>
		            </a>
		        </li>
		        <li ng-class="{'active': activeURL == 'modelos', '': activeURL != 'modelos'}">
		        	<a title='<spring:message code="header.modelos"/>' href="<c:url value='/protected/modelos'/>">
		            	<p><spring:message code="header.modelos"/></p>
		            </a>
		        </li>
		        <li ng-class="{'active': activeURL == 'obrasSociales', '': activeURL != 'obrasSociales'}">
		        	<a title='<spring:message code="header.obrasSociales"/>' href="<c:url value='/protected/obrasSociales'/>">
		            	<p><spring:message code="header.obrasSociales"/></p>
		            </a>
		        </li>
		        <li ng-class="{'active': activeURL == 'calendarios', '': activeURL != 'calendarios'}">
		        	<a title='<spring:message code="header.turnos"/>' href="<c:url value='/protected/calendarios'/>">
		            	<p><spring:message code="header.turnos"/></p>
		            </a>
		        </li>			        		        
			</ul>
			<ul class="nav navbar-nav navbar-right">
		        <li>
		        	<a href="<c:url value='/logout' />" title='<spring:message code="header.logout"/>'>
		            	<p><spring:message code="header.logout"/>&nbsp;(${usuarioActual.nombreCompleto})</p>
		            </a>
		        </li>			
			</ul>
		  </div>
		</div>
	</div>
</div>