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
 *         Controlador Login Page de la Aplicación
 */
@Controller
@RequestMapping("/login")
public class LoginController {

	@RequestMapping(method = { RequestMethod.GET, RequestMethod.POST,
			RequestMethod.DELETE, RequestMethod.PUT })
	public ModelAndView doGet() {
		// Se retorna la Vista Login
		return new ModelAndView("login");
	}
}
