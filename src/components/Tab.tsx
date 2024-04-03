import { useContext } from "react";
import { TeamsFxContext } from "./Context";
import React from "react";
import { Providers, ProviderState, Agenda, Person, applyTheme,People,PeoplePicker,Todo, PersonCard, ViewType } from "@microsoft/mgt-react";
import { Button } from "@fluentui/react-components";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import {
  TeamsUserCredential,
  TeamsUserCredentialAuthConfig,
} from "@microsoft/teamsfx";

const authConfig: TeamsUserCredentialAuthConfig = {
  clientId: '094e7487-00f3-4ba3-af05-691f44a37abd',
  initiateLoginEndpoint: 'https://4ad7-183-82-30-115.ngrok-free.app/auth-start.html',
};

const scopes = [
  'People.Read',
  'User.Read', 
  'User.Read.All',
  'Group.ReadWrite.All'
];
const credential = new TeamsUserCredential(authConfig);
const provider = new TeamsFxProvider(credential, scopes);

Providers.globalProvider = provider;

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [consentNeeded, setConsentNeeded] = React.useState<boolean>(false);

  const [activeDiv, setActiveDiv] = React.useState(null);

  React.useEffect(() => {
    const init = async () => {
      try {
        await credential.getToken(scopes);
        Providers.globalProvider.setState(ProviderState.SignedIn);
      } catch (error) {
        setConsentNeeded(true);
      }
    };

    init();
  }, []);

  const consent = async () => {
    setLoading(true);
    await credential.login(scopes);
    Providers.globalProvider.setState(ProviderState.SignedIn);
    setLoading(false);
    setConsentNeeded(false);
  };

  const toggleVisibility = (divName:any) => {
    setActiveDiv(divName === activeDiv ? null : divName);
  };

  React.useEffect(() => {
    applyTheme(themeString === "default" ? "light" : "dark");
  }, [themeString]);

  return (
    <div>
      {consentNeeded && (
        <>
        <div className="loginbtn">
          <Button appearance="primary" disabled={loading} onClick={consent}>
            Login
          </Button>
        </div>
        </>
      )}
      {!consentNeeded && (
        <>
        <div className="mgtDetails">
          <Person personQuery="me"  />
          <div>
            <h3 className="cursorPointer" onClick={() => toggleVisibility('Agenda')}><i className="arrow right"></i>Agenda</h3> 
            <div style={{ display: activeDiv === 'Agenda' ? 'block' : 'none' }}> 
            <Agenda></Agenda> 
            </div>
          </div>

          <div>
            <h3 className="cursorPointer" onClick={() => toggleVisibility('PeoplePicker')}><i className="arrow right"></i>PeoplePicker</h3> 
            <div style={{ display: activeDiv === 'PeoplePicker' ? 'block' : 'none' }}> 
            <PeoplePicker></PeoplePicker>
          </div>
          </div>

          <div>
            <h3 className="cursorPointer" onClick={() => toggleVisibility('Todo')}><i className="arrow right"></i>ToDo</h3> 
            <div style={{ display: activeDiv === 'Todo' ? 'block' : 'none' }}>
              <Todo></Todo> 
            </div>
          </div>
            
          <div>
            <h3 className="cursorPointer" onClick={() => toggleVisibility('Person')}><i className="arrow right"></i>Person Card</h3> 
            <div style={{ display: activeDiv === 'Person' ? 'block' : 'none' }}>
            <Person personQuery="me" /> 
            </div>
          </div>

          <div>
            <h3 className="cursorPointer" onClick={() => toggleVisibility('People')}><i className="arrow right"></i>Person</h3> 
            <div style={{ display: activeDiv === 'People' ? 'block' : 'none' }}>
            <People></People>
            </div>
          </div>
        </div>
        </>
      )}
    </div>
  );
}