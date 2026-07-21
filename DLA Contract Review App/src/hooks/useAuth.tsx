import { createContext, useContext, useEffect, useMemo, useState } from 'react';
import type { ReactNode } from 'react';
import { UiPath, UiPathError } from '@uipath/uipath-typescript';
import { isUiPathConfigured, uipathConfig } from '../config';

type AuthContextValue = {
  sdk: UiPath;
  isAuthenticated: boolean;
  isConfigured: boolean;
  isLoading: boolean;
  error: string | null;
  login: () => Promise<void>;
  logout: () => void;
};

const AuthContext = createContext<AuthContextValue | undefined>(undefined);

function clearOAuthSession(clientId: string) {
  sessionStorage.removeItem(`uipath_sdk_user_token-${clientId}`);
  sessionStorage.removeItem('uipath_sdk_oauth_context');
  sessionStorage.removeItem('uipath_sdk_code_verifier');
}

function messageFromError(error: unknown, fallback: string) {
  if (error instanceof UiPathError || error instanceof Error) return error.message;
  return fallback;
}

export function AuthProvider({ children }: { children: ReactNode }) {
  const isConfigured = isUiPathConfigured();
  const config = useMemo(() => uipathConfig, []);
  const [sdk, setSdk] = useState(() => new UiPath(config));
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(isConfigured);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    async function initialize() {
      if (!isConfigured) {
        setIsLoading(false);
        return;
      }

      setIsLoading(true);
      setError(null);
      try {
        await sdk.initialize();
        setIsAuthenticated(sdk.isAuthenticated());
      } catch (authError) {
        setError(messageFromError(authError, 'Authentication failed'));
        setIsAuthenticated(false);
      } finally {
        setIsLoading(false);
      }
    }

    void initialize();
  }, [config.clientId, isConfigured, sdk]);

  async function login() {
    if (!isConfigured) {
      setError('UiPath OAuth is not configured for this deployment.');
      return;
    }

    setIsLoading(true);
    setError(null);
    try {
      const nextSdk = new UiPath(config);
      setSdk(nextSdk);
      await nextSdk.initialize();
      setIsAuthenticated(nextSdk.isAuthenticated());
    } catch (authError) {
      setError(messageFromError(authError, 'Login failed'));
      setIsAuthenticated(false);
    } finally {
      setIsLoading(false);
    }
  }

  function logout() {
    clearOAuthSession(config.clientId);
    setSdk(new UiPath(config));
    setIsAuthenticated(false);
    setError(null);
  }

  return (
    <AuthContext.Provider value={{ sdk, isAuthenticated, isConfigured, isLoading, error, login, logout }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (!context) throw new Error('useAuth must be used inside AuthProvider');
  return context;
}
