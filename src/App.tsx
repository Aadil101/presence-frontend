import { Providers, ProviderState } from '@microsoft/mgt-element';
import { Login, ThemeToggle } from '@microsoft/mgt-react';
import React, { useEffect, useState } from 'react';
import { Client } from '@microsoft/microsoft-graph-client';
import { FluentProvider, teamsLightTheme, teamsDarkTheme, Text } from "@fluentui/react-components";
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

function usePresence(): [string, string] {
  const [availability, setAvailability] = useState('');
  const [activity, setActivity] = useState('');
  
  useEffect(() => {
    const fetchAvail = async () => {
      const provider = Providers.globalProvider;
      if (provider.state === ProviderState.SignedIn) {
        const graphClient = Client.initWithMiddleware({authProvider: Providers.globalProvider})
        const result = await graphClient
          .api('/me/presence')
          .version('beta')
          .get();
        setAvailability(result.availability);
        setActivity(result.activity);
      } else {
        setAvailability('');
        setActivity('');
      }
    }
    Providers.onProviderUpdated(fetchAvail);
  }, [])

  return [availability, activity];
}

function App() {
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [isSignedIn] = useIsSignedIn();
  const [availability, activity] = usePresence();

  return (
    <FluentProvider theme={isDarkMode ? teamsDarkTheme : teamsLightTheme}>
      <div className="app">
        <Login showPresence={true} loginView='full'/>
        {isSignedIn && availability && activity && 
          <>
            <Text>Availability: {availability}</Text>
            <Text>Activity: {activity}</Text>
          </>
        }
        <ThemeToggle darkmodechanged={(e) => setIsDarkMode(e.detail)}>Dark Mode</ThemeToggle>
      </div>
    </FluentProvider>
  );
}

export default App;
