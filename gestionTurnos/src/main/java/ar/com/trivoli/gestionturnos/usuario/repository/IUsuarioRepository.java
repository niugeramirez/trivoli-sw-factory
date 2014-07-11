/**
 * 
 */
package ar.com.trivoli.gestionturnos.usuario.repository;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.usuario.model.Usuario;

/**
 * @author posadas
 * 
 *         Repositorio de Usuarios de la Aplicacion
 * 
 */
public interface IUsuarioRepository extends
		PagingAndSortingRepository<Usuario, Integer> {

	Usuario findByUsuario(String usuario);
}
