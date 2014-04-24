/**
 * 
 */
package ar.com.trivoli.gestionturnos.usuario.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import ar.com.trivoli.gestionturnos.usuario.model.Usuario;
import ar.com.trivoli.gestionturnos.usuario.repository.UsuarioRepository;

/**
 * @author posadas
 * 
 *         Servicio de Usuarios de la Aplicacion
 */
@Service
public class UsuarioService {

	@Autowired
	private UsuarioRepository usuarioRepository;

	public Usuario findByUsuario(String usuario) {
		return usuarioRepository.findByUsuario(usuario);
	}
}
