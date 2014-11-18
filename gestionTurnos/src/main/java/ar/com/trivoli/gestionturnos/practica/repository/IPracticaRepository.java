/**
 * 
 */
package ar.com.trivoli.gestionturnos.practica.repository;

import org.springframework.data.repository.PagingAndSortingRepository;

import ar.com.trivoli.gestionturnos.practica.model.Practica;


/**
 * @author ramirez
 *
 */
public interface IPracticaRepository extends
			PagingAndSortingRepository<Practica, Integer>{

}
