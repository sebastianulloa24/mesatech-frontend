import React, { useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import axios from 'axios';

// Reemplaza con la URL base de tu AWS API Gateway HTTP API
const API_GATEWAY_URL = "https://xxxxxx.execute-api.us-east-1.amazonaws.com";

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

  // Estados para cambio de estado
  const [solicitudId, setSolicitudId] = useState('');
  const [nuevoEstado, setNuevoEstado] = useState('EN_PROCESO');

  const handleLogin = () => {
    setErrorMsg('');
    instance.loginRedirect(loginRequest).catch(e => console.error(e));
  };

  const handleLogout = () => {
    instance.logoutRedirect({ postLogoutRedirectUri: "http://localhost:3000" });
  };

  // Función helper para obtener el token de forma silenciosa
  const obtenerToken = async () => {
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account: accounts[0]
    });

    const token = response.accessToken;

    // Decodificar claims (Demostración requerida en la evaluación)
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

  // 2. POST: Crear nueva solicitud (Cliente / Opcional otros)
  const crearSolicitud = async (e) => {
    e.preventDefault();
    try {
      setErrorMsg('');
      const token = await obtenerToken();
      const payload = { titulo, descripcion, categoria, prioridad };

      const res = await axios.post(`${API_GATEWAY_URL}/v1/solicitudes`, payload, {
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

  // 3. PUT/PATCH: Actualizar Estado (Operador / Administrador)
  const actualizarEstado = async (e) => {
    e.preventDefault();
    try {
      setErrorMsg('');
      const token = await obtenerToken();

      const res = await axios.patch(`${API_GATEWAY_URL}/v1/solicitudes/${solicitudId}/estado`,
          { estado: nuevoEstado },
          { headers: { Authorization: `Bearer ${token}` } }
      );

      alert("Estado actualizado con éxito");
      setRespuestaApi(res.data);
    } catch (error) {
      console.error(error);
      setErrorMsg("Regla de negocio no válida o falta de permisos (Ej: Pasar a RESUELTA sin estar EN_PROCESO).");
    }
  };

  // Extraer los roles del token para condicionar la interfaz
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
                <button onClick={() => llamarApiGet('/v1/solicitudes/mias')}>Mis Solicitudes (/v1/solicitudes/mias)</button>
                <button onClick={() => llamarApiGet('/v1/solicitudes')} style={{ marginLeft: '8px' }}>Todas las Solicitudes (/v1/solicitudes)</button>
                <button onClick={() => llamarApiGet('/v1/catalogo')} style={{ marginLeft: '8px' }}>Ver Catálogo (/v1/catalogo)</button>
                <button onClick={() => llamarApiGet('/v2/solicitudes')} style={{ marginLeft: '8px', backgroundColor: '#e2e3e5' }}>Versión v2 (/v2/solicitudes)</button>
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

              {/* SECCIÓN 3: CAMBIAR ESTADO DE SOLICITUD (Operador / Admin) */}
              <div style={{ border: '1px solid #ddd', padding: '15px', borderRadius: '8px', marginBottom: '25px' }}>
                <h3>3. Actualizar Estado de Solicitud (Operador / Admin)</h3>
                <form onSubmit={actualizarEstado} style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                  <input
                      type="text"
                      placeholder="ID Solicitud"
                      value={solicitudId}
                      onChange={(e) => setSolicitudId(e.target.value)}
                      required
                      style={{ padding: '8px' }}
                  />
                  <select value={nuevoEstado} onChange={(e) => setNuevoEstado(e.target.value)} style={{ padding: '8px' }}>
                    <option value="ASIGNADA">ASIGNADA</option>
                    <option value="EN_PROCESO">EN_PROCESO</option>
                    <option value="RESUELTA">RESUELTA</option>
                    <option value="CERRADA">CERRADA</option>
                    <option value="CANCELADA">CANCELADA</option>
                  </select>
                  <button type="submit" style={{ padding: '8px 15px', background: '#198754', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>
                    Cambiar Estado
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