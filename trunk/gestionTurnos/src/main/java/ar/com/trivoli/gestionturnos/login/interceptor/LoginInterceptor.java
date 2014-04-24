/**
 * 
 */
package ar.com.trivoli.gestionturnos.login.interceptor;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.web.servlet.handler.HandlerInterceptorAdapter;

import ar.com.trivoli.gestionturnos.usuario.model.Usuario;
import ar.com.trivoli.gestionturnos.usuario.service.UsuarioService;

/**
 * @author posadas
 * 
 */
public class LoginInterceptor extends HandlerInterceptorAdapter {
	@Autowired
	private UsuarioService usuarioService;

	@Override
	public boolean preHandle(HttpServletRequest request,
			HttpServletResponse response, Object handler) throws Exception {
		HttpSession session = request.getSession();

		Usuario usuario = (Usuario) session.getAttribute("usuarioActual");
		if (usuario == null) {
			Authentication auth = SecurityContextHolder.getContext()
					.getAuthentication();
			String email = auth.getName();
			usuario = usuarioService.findByUsuario(email);
			session.setAttribute("usuarioActual", usuario);
		}

		return super.preHandle(request, response, handler);
	}
}
