/**
 * 
 */
package ar.com.trivoli.gestionturnos.login.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

/**
 * @author posadas
 * 
 *         Controlador Home Page de la Aplicación
 */
@Controller
@RequestMapping(value = "/protected/home")
public class HomeController {

	@RequestMapping(method = RequestMethod.GET)
	public ModelAndView welcome() {
		// Se retorna la Vista Home
		return new ModelAndView("home");
	}
}
