<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<%@ taglib prefix="c" uri="http://java.sun.com/jsp/jstl/core" %>
<%@ taglib prefix="spring" uri="http://www.springframework.org/tags" %>

<div class="login">
	<div class="row-fluid clearfix" >
	    <div class="col-md-10 col-md-offset-1 well-lg">
	        <h1 class="welcome"><spring:message code='project.name'/></h1>
	    </div>
	</div>
	
	<div class="row-fluid clearfix">
		<div class="col-md-4 col-md-offset-4 well" ng-controller="loginController">
			<div class="alert alert-danger" ng-class="{'': displayLoginError == true, 'none': displayLoginError == false}">
	            <spring:message code="login.error" />
	        </div>
	 		<form class="form-signin" method="post" action="j_spring_security_check">
	            <h2 class="form-signin-heading text-center"><spring:message code="login.header" /></h2>
	            <input name="j_username" id="j_username" type="text" class="form-control" placeholder="<spring:message code='login.username'/>">
	            <input name="j_password" id="j_password" type="password"  class="form-control" placeholder="<spring:message code='login.password'/>">
	            <button type="submit" name="submit" class="btn btn-lg btn-primary btn-block"><spring:message code="login.signIn" /></button>
	        </form>
	 	</div>
	</div>       
</div>
<script src="<c:url value='/resources/js/pages/login.js' />"></script>