/**
 * 
 */
package ar.com.trivoli.gestionturnos.paciente.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import ar.com.trivoli.gestionturnos.common.model.ListaEntidadDTO;
import ar.com.trivoli.gestionturnos.paciente.model.Paciente;
import ar.com.trivoli.gestionturnos.paciente.repository.IPacienteRepository;

/**
 * @author ramirez
 *
 */
@Service
@Transactional
public class PacienteService {
	
	@Autowired
	private IPacienteRepository pacienteRepository;

	/************************************************************************************************************************************************************************/
	private Sort ordenPredeterminado() {
		return new Sort(Sort.Direction.ASC, "nombre");
	}
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public ListaEntidadDTO<Paciente> recuperarTodos(int nroPagina,
			int registrosPorPagina) {
		PageRequest pageRequest = new PageRequest(nroPagina,
				registrosPorPagina, ordenPredeterminado());

		Page<Paciente> resultado = pacienteRepository.findAll(pageRequest);

		// Se determina si la pagina requerida es posterior a la Ultima pagina
		// disponible
		if (resultado.getTotalElements() > 0
				&& nroPagina > (resultado.getTotalPages() - 1)) {
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,
					ordenPredeterminado());

			resultado = pacienteRepository.findAll(pageRequest);
		}

		return new ListaEntidadDTO<Paciente>(	resultado.getTotalPages(),
												resultado.getTotalElements(),
												resultado.getContent());
	}
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public ListaEntidadDTO<Paciente> recuperarPorComienzoDni(	int nroPagina
																,int registrosPorPagina
																,String dni
															) 
		{
		
		PageRequest pageRequest = new PageRequest(nroPagina,registrosPorPagina, ordenPredeterminado());

		Page<Paciente> resultado = pacienteRepository.findByDniStartingWith(pageRequest,dni);

		// Se determina si la pagina requerida es posterior a la Ultima pagina disponible
		if (resultado.getTotalElements() > 0	&& nroPagina > (resultado.getTotalPages() - 1))
		{
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,	ordenPredeterminado());

			resultado = pacienteRepository.findByDniStartingWith(pageRequest,dni);
		}

		return new ListaEntidadDTO<Paciente>(	resultado.getTotalPages(),
												resultado.getTotalElements(),
												resultado.getContent());
	}
	/************************************************************************************************************************************************************************/
	@Transactional(readOnly = true)
	public ListaEntidadDTO<Paciente> recuperarPorComienzoDniApellidoNombreLike(	int nroPagina
																,int registrosPorPagina
																,String dni
																,String apellido
																,String nombre																
															) 
		{
		
		PageRequest pageRequest = new PageRequest(nroPagina,registrosPorPagina, ordenPredeterminado());

		Page<Paciente> resultado = pacienteRepository.findByDniStartingWithAndApellidoLikeAndNombreLike(pageRequest,dni,"%" +apellido+"%" ,"%" +nombre+"%" );

		// Se determina si la pagina requerida es posterior a la Ultima pagina disponible
		if (resultado.getTotalElements() > 0	&& nroPagina > (resultado.getTotalPages() - 1))
		{
			int ultimaPagina = resultado.getTotalPages() - 1;

			pageRequest = new PageRequest(ultimaPagina, registrosPorPagina,	ordenPredeterminado());

			resultado = pacienteRepository.findByDniStartingWithAndApellidoLikeAndNombreLike(pageRequest,dni,apellido,nombre);
		}

		return new ListaEntidadDTO<Paciente>(	resultado.getTotalPages(),
												resultado.getTotalElements(),
												resultado.getContent());
	}	
	/************************************************************************************************************************************************************************/
}
