import React, { useState, useEffect } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import axios from 'axios';

// Reemplazar con la URL de Invocación ($default) copiada de AWS API Gateway
const API_GATEWAY_URL = "https://iwnne47f0f.execute-api.us-east-1.amazonaws.com";
function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const activeAccount = instance.getActiveAccount() || accounts[0];
  const correoUsuario = activeAccount?.username || activeAccount?.idTokenClaims?.preferred_username || '';

  // Estados generales
  const [respuestaApi, setRespuestaApi] = useState(null);
  const [tokenDecodificado, setTokenDecodificado] = useState(null);
  const [errorMsg, setErrorMsg] = useState('');

  // Estados para formularios de negocio
  const [titulo, setTitulo] = useState('');
  const [descripcion, setDescripcion] = useState('');
  const [categoriaId, setCategoriaId] = useState(1);
  const [prioridadId, setPrioridadId] = useState(3);

  // Estados para actualización
  const [solicitudId, setSolicitudId] = useState('');
  const [nuevoEstado, setNuevoEstado] = useState('ASIGNADA');

  // Roles de usuario
  const rolesUsuario = tokenDecodificado?.roles || activeAccount?.idTokenClaims?.roles || accounts[0]?.idTokenClaims?.roles || [];
  const esCliente = rolesUsuario.includes('ROLE_CLIENTE');
  const esOperadorOAdmin = rolesUsuario.includes('ROLE_OPERADOR') || rolesUsuario.includes('ROLE_ADMIN');
  const puedeActualizar = esOperadorOAdmin || !esCliente;

  const obtenerSiguientesEstados = (estado) => {
    switch (estado) {
      case 'CREADA':
        return ['ASIGNADA', 'CANCELADA'];
      case 'ASIGNADA':
        return ['EN_PROCESO'];
      case 'EN_PROCESO':
        return ['RESUELTA'];
      case 'RESUELTA':
        return ['CERRADA'];
      case 'CERRADA':
      case 'CANCELADA':
        return [];
      default:
        return ['ASIGNADA', 'CANCELADA'];
    }
  };

  const solicitudSeleccionada = Array.isArray(respuestaApi)
    ? respuestaApi.find(s => String(s.id || s.solicitudId) === String(solicitudId))
    : (respuestaApi && String(respuestaApi.id || respuestaApi.solicitudId) === String(solicitudId) ? respuestaApi : null);

  const estadoActual = solicitudSeleccionada?.estado || 'CREADA';
  const siguientesEstados = obtenerSiguientesEstados(estadoActual);
  const estaDeshabilitadoActualizar = estadoActual === 'CERRADA' || estadoActual === 'CANCELADA' || siguientesEstados.length === 0;

  useEffect(() => {
    const siguientes = obtenerSiguientesEstados(estadoActual);
    if (siguientes.length > 0 && !siguientes.includes(nuevoEstado)) {
      setNuevoEstado(siguientes[0]);
    }
  }, [estadoActual, nuevoEstado]);

  const handleLogin = () => {
    setErrorMsg('');
    instance.loginRedirect(loginRequest).catch(e => console.error(e));
  };

  const handleLogout = () => {
    instance.logoutRedirect({ postLogoutRedirectUri: "http://localhost:3000" });
  };

  // Función helper para obtener el token JWT
  const obtenerToken = async () => {
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    });

    const token = response.accessToken;

    // Decodificar claims (Demostración requerida)
    const payloadBase64 = token.split('.')[1];
    const decodedJson = JSON.parse(atob(payloadBase64));
    setTokenDecodificado(decodedJson);

    return token;
  };

  // 1. GET: Consultar APIs
  const llamarApiGet = async (endpoint) => {
    try {
      setErrorMsg('');
      const token = await obtenerToken();
      const res = await axios.get(`${API_GATEWAY_URL}${endpoint}`, {
        headers: { Authorization: `Bearer ${token}` }
      });
      let data = res.data;
      if (endpoint === '/api/bff/solicitudes' && esCliente && !esOperadorOAdmin) {
        if (Array.isArray(data)) {
          data = data.filter(s => s.usuarioSolicitante === correoUsuario);
        } else if (data && Array.isArray(data.content)) {
          data = {
            ...data,
            content: data.content.filter(s => s.usuarioSolicitante === correoUsuario)
          };
        }
      }
      setRespuestaApi(data);
    } catch (error) {
      console.error(error);
      setErrorMsg(error.response?.data?.message || "Error al realizar la consulta o sin autorización.");
    }
  };

  // 2. POST: Crear nueva solicitud
  const crearSolicitud = async (e) => {
    e.preventDefault();
    try {
      setErrorMsg('');
      const token = await obtenerToken();
      const correoUsuarioActivo = activeAccount?.username || activeAccount?.idTokenClaims?.preferred_username;
      const payload = {
        titulo,
        descripcion,
        categoriaId: Number(categoriaId),
        prioridadId: Number(prioridadId),
        usuarioSolicitante: correoUsuarioActivo
      };

      const res = await axios.post(`${API_GATEWAY_URL}/api/bff/solicitudes`, payload, {
        headers: { Authorization: `Bearer ${token}` }
      });

      alert("¡Solicitud creada con éxito!");
      setRespuestaApi(res.data);
      setTitulo('');
      setDescripcion('');
    } catch (error) {
      console.error(error);
      setErrorMsg("Error al crear la solicitud.");
    }
  };

  // 3. PUT: Actualizar / Editar Solicitud por ID
  const actualizarSolicitud = async (e) => {
    e.preventDefault();
    try {
      setErrorMsg('');
      const token = await obtenerToken();
      const payload = {
        titulo: solicitudSeleccionada?.titulo || titulo,
        descripcion: solicitudSeleccionada?.descripcion || descripcion,
        categoriaId: Number(solicitudSeleccionada?.categoriaId || categoriaId),
        prioridadId: Number(solicitudSeleccionada?.prioridadId || prioridadId),
        estado: nuevoEstado
      };

      const res = await axios.put(`${API_GATEWAY_URL}/api/bff/solicitudes/${solicitudId}`, payload, {
        headers: { Authorization: `Bearer ${token}` }
      });

      alert("Solicitud actualizada con éxito");
      setRespuestaApi((prev) => {
        if (Array.isArray(prev)) {
          return prev.map((s) => (String(s.id || s.solicitudId) === String(solicitudId) ? { ...s, ...res.data, estado: nuevoEstado } : s));
        }
        return res.data;
      });
    } catch (error) {
      console.error(error);
      setErrorMsg("Error al actualizar la solicitud o falta de permisos.");
    }
  };



  return (
    <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif', maxWidth: '900px', margin: '0 auto' }}>
      <h1>MesaTech Cloud - Sistema de Gestión de Soporte</h1>

      {!isAuthenticated ? (
        <div style={{ padding: '20px', border: '1px solid #ccc', borderRadius: '8px' }}>
          <h3>Bienvenido</h3>
          <p>Debes autenticarte mediante Microsoft Entra ID para acceder al sistema.</p>
          <button onClick={handleLogin} style={{ padding: '10px 20px', cursor: 'pointer' }}>
            Iniciar Sesión con Entra ID
          </button>
        </div>
      ) : (
        <div>
          {/* BARRA DE USUARIO */}
          <div style={{ background: '#e9ecef', padding: '15px', borderRadius: '8px', marginBottom: '20px' }}>
            <p style={{ margin: 0 }}>
              Usuario Activo: <strong>{accounts[0]?.name}</strong> ({accounts[0]?.username})
            </p>
            <p style={{ margin: '5px 0 0 0', fontSize: '0.9em', color: '#555' }}>
              Roles en JWT: <strong>{rolesUsuario.join(', ') || 'Sin roles asignados'}</strong>
            </p>
            <button onClick={handleLogout} style={{ marginTop: '10px', cursor: 'pointer' }}>
              Cerrar Sesión
            </button>
          </div>

          {/* ERRORES */}
          {errorMsg && (
            <div style={{ background: '#f8d7da', color: '#721c24', padding: '10px', borderRadius: '5px', marginBottom: '20px' }}>
              <strong>Mensaje del Sistema / Error:</strong> {errorMsg}
            </div>
          )}

          {/* SECCIÓN 1: ACCIONES DE LECTURA (APIs) */}
          <div style={{ marginBottom: '25px' }}>
            <h3>1. Consultas Rápidas (AWS API Gateway)</h3>
            <button onClick={() => llamarApiGet('/api/bff/solicitudes')}>
              Listar Solicitudes (/api/bff/solicitudes)
            </button>
            <button onClick={() => llamarApiGet('/api/bff/categorias')} style={{ marginLeft: '8px' }}>
              Ver Categorías (/api/bff/categorias)
            </button>
            <button onClick={() => llamarApiGet('/api/bff/prioridades')} style={{ marginLeft: '8px' }}>
              Ver Prioridades (/api/bff/prioridades)
            </button>
          </div>

          {/* SECCIÓN 2: FORMULARIO CREAR SOLICITUD */}
          <div style={{ border: '1px solid #ddd', padding: '15px', borderRadius: '8px', marginBottom: '25px' }}>
            <h3>2. Crear Solicitud de Soporte</h3>
            <form onSubmit={crearSolicitud}>
              <div style={{ marginBottom: '10px' }}>
                <label>Título: </label><br />
                <input type="text" value={titulo} onChange={(e) => setTitulo(e.target.value)} required style={{ width: '100%', padding: '8px' }} />
              </div>
              <div style={{ marginBottom: '10px' }}>
                <label>Descripción: </label><br />
                <textarea value={descripcion} onChange={(e) => setDescripcion(e.target.value)} required style={{ width: '100%', padding: '8px' }} />
              </div>
              <div style={{ display: 'flex', gap: '15px', marginBottom: '10px' }}>
                <div>
                  <label>Categoría: </label>
                  <select value={categoriaId} onChange={(e) => setCategoriaId(Number(e.target.value))}>
                    <option value={1}>Soporte Hardware/Software</option>
                    <option value={2}>Redes e Infraestructura</option>
                    <option value={3}>Cuentas y Accesos</option>
                  </select>
                </div>
                <div>
                  <label>Prioridad: </label>
                  <select value={prioridadId} onChange={(e) => setPrioridadId(Number(e.target.value))}>
                    <option value={1}>BAJA</option>
                    <option value={2}>MEDIA</option>
                    <option value={3}>ALTA</option>
                  </select>
                </div>
              </div>
              <button type="submit" style={{ padding: '8px 15px', background: '#0d6efd', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                Enviar Solicitud
              </button>
            </form>
          </div>

          {/* SECCIÓN 3: EDITAR / ACTUALIZAR SOLICITUD */}
          {puedeActualizar && (
            <div style={{ border: '1px solid #ddd', padding: '15px', borderRadius: '8px', marginBottom: '25px' }}>
              <h3>3. Actualizar Solicitud por ID (PUT)</h3>
              <form onSubmit={actualizarSolicitud} style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                <input
                  type="text"
                  placeholder="ID Solicitud"
                  value={solicitudId}
                  onChange={(e) => setSolicitudId(e.target.value)}
                  required
                  style={{ padding: '8px' }}
                />
                <select
                  value={nuevoEstado}
                  onChange={(e) => setNuevoEstado(e.target.value)}
                  disabled={estaDeshabilitadoActualizar}
                  style={{ padding: '8px' }}
                >
                  {siguientesEstados.length > 0 ? (
                    siguientesEstados.map((est) => (
                      <option key={est} value={est}>
                        {est}
                      </option>
                    ))
                  ) : (
                    <option value={estadoActual}>{estadoActual}</option>
                  )}
                </select>
                <button
                  type="submit"
                  disabled={estaDeshabilitadoActualizar}
                  style={{
                    padding: '8px 15px',
                    background: estaDeshabilitadoActualizar ? '#6c757d' : '#198754',
                    color: 'white',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: estaDeshabilitadoActualizar ? 'not-allowed' : 'pointer'
                  }}
                >
                  Actualizar Solicitud
                </button>
              </form>
            </div>
          )}

          {/* SECCIÓN 4: VISUALIZACIÓN DE CLAIMS Y RESPUESTA DE API */}
          {tokenDecodificado && (
            <div style={{ background: '#f8f9fa', padding: '15px', borderRadius: '8px', marginBottom: '20px', border: '1px solid #e9ecef' }}>
              <h4>Claims Relevantes del JWT (Demostración de Identidad):</h4>
              <p><strong>Emisor (iss):</strong> {tokenDecodificado.iss}</p>
              <p><strong>Audiencia (aud):</strong> {tokenDecodificado.aud}</p>
              <p><strong>Roles (roles):</strong> {JSON.stringify(tokenDecodificado.roles || [])}</p>
            </div>
          )}

          {respuestaApi && (
            <div style={{ background: '#212529', color: '#f8f9fa', padding: '15px', borderRadius: '8px' }}>
              <h4 style={{ margin: '0 0 10px 0', color: '#0dcaf0' }}>Respuesta desde Backend (vía AWS API Gateway):</h4>
              <pre style={{ overflowX: 'auto', margin: 0 }}>{JSON.stringify(respuestaApi, null, 2)}</pre>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export default App;