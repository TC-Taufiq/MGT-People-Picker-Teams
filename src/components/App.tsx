// https://fluentsite.z22.web.core.windows.net/quick-start
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  tokens
} from '@fluentui/react-components';
import { useEffect } from 'react';
import { HashRouter as Router, Navigate, Route, Routes } from 'react-router-dom';
import { app } from '@microsoft/teams-js';
import { useTeamsUserCredential } from '@microsoft/teamsfx-react';
import Tab from './Tab';
import { TeamsFxContext } from './Context';

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { loading, theme, themeString, teamsUserCredential } = useTeamsUserCredential({
    clientId: '094e7487-00f3-4ba3-af05-691f44a37abd',
    initiateLoginEndpoint: 'https://4ad7-183-82-30-115.ngrok-free.app/auth-start.html'
  });
  useEffect(() => {
    loading &&
      app.initialize().then(() => {
        // Hide the loading indicator.
        app.notifySuccess();
      });
  }, [loading]);
  return (
    <TeamsFxContext.Provider value={{ theme, themeString, teamsUserCredential }}>
      <FluentProvider
        theme={
          themeString === 'dark'
            ? teamsDarkTheme
            : themeString === 'contrast'
            ? teamsHighContrastTheme
            : {
                ...teamsLightTheme,
                colorNeutralBackground3: '#FFFFFF'
              }
        }
        style={{ background: tokens.colorNeutralBackground3 }}
      >
        <Router>
          {!loading && (
            <Routes>
              <Route path="/tab" element={<Tab />} />
              <Route path="*" element={<Navigate to={'/tab'} />}></Route>
            </Routes>
          )}
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}
