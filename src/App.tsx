import { Providers, ProviderState } from '@microsoft/mgt-element';
import { Person, Login } from '@microsoft/mgt-react';
import React, { useState, useEffect } from 'react';
import { Client } from '@microsoft/microsoft-graph-client';
import './App.css';

function useIsSignedIn(): [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);
  
  useEffect(() => {
    const updateState = () => {
      const provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

function useAvail(): [string] {
  const [avail, setAvail] = useState('');
  
  useEffect(() => {
    const fetchAvail = async () => {
      const provider = Providers.globalProvider;
      if (provider.state === ProviderState.SignedIn) {
        const graphClient = Client.initWithMiddleware({authProvider: Providers.globalProvider})
        const result = await graphClient
          .api('/me/presence')
          .version('beta')
          .get();
        setAvail(result.availability);
      } else {
        setAvail('unknown');
      }
    }
    Providers.onProviderUpdated(fetchAvail);
  }, [])

  return [avail];
}

function App() {
  const [isSignedIn] = useIsSignedIn();
  const [avail] = useAvail();
  
  return (
    <div className="app">
      <Login showPresence={true} loginView='full'/>
      {isSignedIn && 
        <p>{avail}</p>
      }
      {/* {<Person personQuery='me' showPresence={true} view={6}/>} */}
    </div>
  );
}

export default App;
