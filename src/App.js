import React, { useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import axios from 'axios';

// Reemplazar con la URL de Invocación ($default) copiada de AWS API Gateway
const API_GATEWAY_URL = "https://iwnne47f0f.execute-api.us-east-1.amazonaws.com/$default";
function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  // Estados generales
  const [respuestaApi, setRespuestaApi] = useState(null);
  const [tokenDecodificado, setTokenDecodificado] = useState(null);
  const [errorMsg, setErrorMsg] = useState('');

  // Estados para formularios de negocio
  const [titulo, setTitulo] = useState('');
  const [descripcion, setDescripcion] = useState('');
  const [categoria, setCategoria] = useState('Soporte');
  const [prioridad, setPrioridad] = useState('ALTA');

  // Estados para actualización
  const [solicitudId, setSolicitudId] = useState('');

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
      setRespuestaApi(res.data);
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
      const payload = { titulo, descripcion, categoria, prioridad };

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
      const payload = { titulo, descripcion, categoria, prioridad };

      const res = await axios.put(`${API_GATEWAY_URL}/api/bff/solicitudes/${solicitudId}`, payload, {
        headers: { Authorization: `Bearer ${token}` }
      });

      alert("Solicitud actualizada con éxito");
      setRespuestaApi(res.data);
    } catch (error) {
      console.error(error);
      setErrorMsg("Error al actualizar la solicitud o falta de permisos.");
    }
  };

  // Extraer roles del token
  const rolesUsuario = tokenDecodificado?.roles || [];

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
                      <select value={categoria} onChange={(e) => setCategoria(e.target.value)}>
                        <option value="Soporte">Soporte Hardware/Software</option>
                        <option value="Redes">Redes e Infraestructura</option>
                        <option value="Accesos">Cuentas y Accesos</option>
                      </select>
                    </div>
                    <div>
                      <label>Prioridad: </label>
                      <select value={prioridad} onChange={(e) => setPrioridad(e.target.value)}>
                        <option value="BAJA">BAJA</option>
                        <option value="MEDIA">MEDIA</option>
                        <option value="ALTA">ALTA</option>
                      </select>
                    </div>
                  </div>
                  <button type="submit" style={{ padding: '8px 15px', background: '#0d6efd', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                    Enviar Solicitud
                  </button>
                </form>
              </div>

              {/* SECCIÓN 3: EDITAR / ACTUALIZAR SOLICITUD */}
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
                  <button type="submit" style={{ padding: '8px 15px', background: '#198754', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                    Actualizar Solicitud
                  </button>
                </form>
              </div>

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