/**
 * 
 */
package ar.com.trivoli.gestionturnos.login.controller;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

/**
 * @author posadas
 * 
 *         Controlador Directorio Raiz de la Aplicación
 */
@Controller
@RequestMapping(value = "/")
public class RootController {

	@RequestMapping(method = { RequestMethod.GET, RequestMethod.POST,
			RequestMethod.DELETE, RequestMethod.PUT }, produces = "application/json")
	public ResponseEntity<?> doGetAjax() {
		return new ResponseEntity<Object>(HttpStatus.FORBIDDEN);
	}

	@RequestMapping(method = { RequestMethod.GET, RequestMethod.POST,
			RequestMethod.DELETE, RequestMethod.PUT })
	public String redirect() {
		// Se redirecciona a la Home Page ubicada en un area segura de la
		// Aplicación
		return "redirect:/protected/home";
	}
}
