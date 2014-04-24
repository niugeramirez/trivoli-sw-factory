/**
 * 
 */
package ar.com.trivoli.gestionturnos.recursos.repository;

import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.recursos.model.Recurso;

/**
 * @author posadas
 * 
 *         Repositorio de Recursos
 */
public interface RecursoRepository extends
		PagingAndSortingRepository<Recurso, Integer> {
	Page<Recurso> findByDescripcionLike(Pageable pageable, String descripcion);
}
