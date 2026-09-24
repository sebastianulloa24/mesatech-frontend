import React, { useState } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import axios from 'axios';

// Reemplazar con la URL base de tu AWS API Gateway HTTP API
const API_GATEWAY_URL = "https://xxxxxx.execute-api.us-east-1.amazonaws.com";

function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [respuestaApi, setRespuestaApi] = useState(null);
  const [tokenDecodificado, setTokenDecodificado] = useState(null);

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch(e => console.error(e));
  };

  const handleLogout = () => {
    instance.logoutRedirect({ postLogoutRedirectUri: "http://localhost:3000" });
  };

  const llamarApiGateway = async (endpoint) => {
    try {
      // Obtención silenciosa del token JWT
      const response = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0]
      });

      const token = response.accessToken;

      // Decodificación de claims (mide verificación de roles/identidad)
      const payloadBase64 = token.split('.')[1];
      const decodedJson = JSON.parse(atob(payloadBase64));
      setTokenDecodificado(decodedJson);

      // Petición hacia AWS API Gateway
      const apiRes = await axios.get(`${API_GATEWAY_URL}${endpoint}`, {
        headers: {
          Authorization: `Bearer ${token}`
        }
      });

      setRespuestaApi(apiRes.data);
    } catch (error) {
      console.error("Error al consumir la API:", error);
      alert("Error en la solicitud o rechazo de autorización.");
    }
  };

  return (
      <div style={{ padding: '20px', fontFamily: 'Arial, sans-serif' }}>
        <h1>MesaTech Cloud - Plataforma de Soporte</h1>

        {!isAuthenticated ? (
            <div>
              <p>Estado: <strong>No Autenticado</strong></p>
              <button onClick={handleLogin}>Iniciar Sesión con Entra ID</button>
            </div>
        ) : (
            <div>
              <p>Estado: <strong>Autenticado</strong></p>
              <p>Usuario: <strong>{accounts[0]?.name}</strong> ({accounts[0]?.username})</p>
              <button onClick={handleLogout}>Cerrar Sesión</button>

              <hr />

              <h3>Consumo de Endpoints (vía AWS API Gateway):</h3>
              <button onClick={() => llamarApiGateway('/v1/solicitudes/mias')}>
                Ver Mis Solicitudes (/v1/solicitudes/mias)
              </button>
              <button onClick={() => llamarApiGateway('/v1/solicitudes')} style={{ marginLeft: '10px' }}>
                Ver Todas las Solicitudes (/v1/solicitudes)
              </button>
              <button onClick={() => llamarApiGateway('/v1/catalogo')} style={{ marginLeft: '10px' }}>
                Ver Catálogo (/v1/catalogo)
              </button>

              {tokenDecodificado && (
                  <div style={{ marginTop: '20px', background: '#f4f4f4', padding: '15px', borderRadius: '5px' }}>
                    <h4>Claims del Access Token (Demostración de Roles):</h4>
                    <p><strong>Issuer (iss):</strong> {tokenDecodificado.iss}</p>
                    <p><strong>Audience (aud):</strong> {tokenDecodificado.aud}</p>
                    <p><strong>Roles asignados:</strong> {JSON.stringify(tokenDecodificado.roles || [])}</p>
                  </div>
              )}

              {respuestaApi && (
                  <div style={{ marginTop: '20px' }}>
                    <h4>Respuesta Backend:</h4>
                    <pre>{JSON.stringify(respuestaApi, null, 2)}</pre>
                  </div>
              )}
            </div>
        )}
      </div>
  );
}

export default App;