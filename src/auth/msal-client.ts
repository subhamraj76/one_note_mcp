import {
  PublicClientApplication,
  Configuration,
  AuthenticationResult,
  AccountInfo,
  InteractionRequiredAuthError,
  LogLevel
} from '@azure/msal-node';
import { promises as fs } from 'fs';
import * as path from 'path';
import * as http from 'http';
import * as url from 'url';

// Microsoft Graph PowerShell public client ID - no app registration needed
const CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const AUTHORITY = 'https://login.microsoftonline.com/common';
const REDIRECT_URI = 'http://localhost:3000';
const SCOPES = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/Notes.ReadWrite',
  'offline_access',
  'openid',
  'profile'
];

export interface AuthConfig {
  cacheDir: string;
}

export class MsalClient {
  private pca: PublicClientApplication | null = null;
  private config: AuthConfig;
  private cacheFilePath: string;
  private useSecureCache: boolean = true;
  private onWarning?: (message: string) => void;

  constructor(config: AuthConfig, onWarning?: (message: string) => void) {
    this.config = config;
    this.cacheFilePath = path.join(config.cacheDir, 'onenote-mcp-cache.json');
    this.onWarning = onWarning;
  }

  private async ensureCacheDir(): Promise<void> {
    try {
      await fs.mkdir(this.config.cacheDir, { recursive: true });
    } catch (error) {
      // Directory might already exist
    }
  }

  private async createCachePlugin(): Promise<any> {
    await this.ensureCacheDir();

    // Try to use secure cache with native encryption
    try {
      const { 
        PersistenceCreator, 
        PersistenceCachePlugin,
        DataProtectionScope 
      } = await import('@azure/msal-node-extensions');

      const persistence = await PersistenceCreator.createPersistence({
        cachePath: this.cacheFilePath,
        dataProtectionScope: DataProtectionScope.CurrentUser,
        serviceName: 'OneNoteMCP',
        accountName: 'onenote-mcp-cache',
        usePlaintextFileOnLinux: false
      });

      this.useSecureCache = true;
      return new PersistenceCachePlugin(persistence);
    } catch (error) {
      // Fall back to plaintext cache
      this.useSecureCache = false;
      
      const warningMessage = 
        'OneNote MCP: Secure token storage unavailable. Using plaintext cache. ' +
        'Your tokens will be stored in plain text. For better security on Linux, install libsecret.';
      
      if (this.onWarning) {
        this.onWarning(warningMessage);
      } else {
        console.warn(warningMessage);
      }

      // Create a simple file-based cache plugin
      return this.createPlaintextCachePlugin();
    }
  }

  private createPlaintextCachePlugin(): any {
    const cacheFilePath = this.cacheFilePath;
    
    return {
      beforeCacheAccess: async (cacheContext: any) => {
        try {
          const data = await fs.readFile(cacheFilePath, 'utf-8');
          cacheContext.tokenCache.deserialize(data);
        } catch (error) {
          // File doesn't exist yet, that's OK
        }
      },
      afterCacheAccess: async (cacheContext: any) => {
        if (cacheContext.cacheHasChanged) {
          await fs.writeFile(cacheFilePath, cacheContext.tokenCache.serialize(), 'utf-8');
        }
      }
    };
  }

  async initialize(): Promise<void> {
    const cachePlugin = await this.createCachePlugin();

    const config: Configuration = {
      auth: {
        clientId: CLIENT_ID,
        authority: AUTHORITY
      },
      cache: {
        cachePlugin
      },
      system: {
        loggerOptions: {
          loggerCallback: (level: LogLevel, message: string) => {
            if (level === LogLevel.Error) {
              console.error(message);
            }
          },
          piiLoggingEnabled: false,
          logLevel: LogLevel.Error
        }
      }
    };

    this.pca = new PublicClientApplication(config);
  }

  async getAccessToken(): Promise<string> {
    if (!this.pca) {
      await this.initialize();
    }

    const accounts = await this.pca!.getTokenCache().getAllAccounts();
    
    if (accounts.length > 0) {
      // Try silent acquisition first
      try {
        const result = await this.pca!.acquireTokenSilent({
          account: accounts[0],
          scopes: SCOPES
        });
        return result.accessToken;
      } catch (error) {
        if (!(error instanceof InteractionRequiredAuthError)) {
          throw error;
        }
        // Fall through to interactive auth
      }
    }

    // Interactive authentication required
    return this.acquireTokenInteractive();
  }

  private async acquireTokenInteractive(): Promise<string> {
    return new Promise((resolve, reject) => {
      let server: http.Server | null = null;
      let timeoutId: NodeJS.Timeout | null = null;
      let isResolved = false;

      const cleanup = () => {
        if (timeoutId) {
          clearTimeout(timeoutId);
          timeoutId = null;
        }
        if (server) {
          server.close();
          server = null;
        }
      };

      const resolveOnce = (value: string) => {
        if (!isResolved) {
          isResolved = true;
          cleanup();
          resolve(value);
        }
      };

      const rejectOnce = (error: Error) => {
        if (!isResolved) {
          isResolved = true;
          cleanup();
          reject(error);
        }
      };

      // 5 minute timeout for user to complete auth
      timeoutId = setTimeout(() => {
        rejectOnce(new Error('Authentication timed out. Please try again.'));
      }, 5 * 60 * 1000);

      server = http.createServer(async (req, res) => {
        const parsedUrl = url.parse(req.url || '', true);
        
        if (parsedUrl.pathname === '/') {
          const code = parsedUrl.query.code as string;
          const error = parsedUrl.query.error as string;
          const errorDescription = parsedUrl.query.error_description as string;

          if (error) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(`
              <!DOCTYPE html>
              <html>
                <head><title>Authentication Failed</title></head>
                <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;">
                  <div style="text-align: center;">
                    <h1 style="color: #d32f2f;">❌ Authentication Failed</h1>
                    <p>${errorDescription || error}</p>
                    <p>You can close this window.</p>
                  </div>
                </body>
              </html>
            `);
            rejectOnce(new Error(`Authentication failed: ${errorDescription || error}`));
            return;
          }

          if (code) {
            try {
              const result = await this.pca!.acquireTokenByCode({
                code,
                scopes: SCOPES,
                redirectUri: REDIRECT_URI
              });

              res.writeHead(200, { 'Content-Type': 'text/html' });
              res.end(`
                <!DOCTYPE html>
                <html>
                  <head><title>Authentication Successful</title></head>
                  <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;">
                    <div style="text-align: center;">
                      <h1 style="color: #4caf50;">✅ Authentication Successful</h1>
                      <p>You are now signed in to OneNote MCP.</p>
                      <p>You can close this window and return to VS Code.</p>
                    </div>
                  </body>
                </html>
              `);

              resolveOnce(result.accessToken);
            } catch (error: any) {
              res.writeHead(200, { 'Content-Type': 'text/html' });
              res.end(`
                <!DOCTYPE html>
                <html>
                  <head><title>Authentication Error</title></head>
                  <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;">
                    <div style="text-align: center;">
                      <h1 style="color: #d32f2f;">❌ Authentication Error</h1>
                      <p>${error.message}</p>
                      <p>You can close this window.</p>
                    </div>
                  </body>
                </html>
              `);
              rejectOnce(error);
            }
          } else {
            res.writeHead(400, { 'Content-Type': 'text/plain' });
            res.end('Missing authorization code');
          }
        }
      });

      server.on('error', (error: any) => {
        if (error.code === 'EADDRINUSE') {
          rejectOnce(new Error('Port 3000 is already in use. Please free it and try again.'));
        } else {
          rejectOnce(error);
        }
      });

      server.listen(3000, async () => {
        // Build the authorization URL
        const authUrl = await this.pca!.getAuthCodeUrl({
          scopes: SCOPES,
          redirectUri: REDIRECT_URI,
          prompt: 'select_account'
        });

        // Open the browser for authentication
        try {
          const open = (await import('open')).default;
          await open(authUrl);
        } catch (error) {
          console.log('\n===========================================');
          console.log('Please open the following URL in your browser:');
          console.log(authUrl);
          console.log('===========================================\n');
        }
      });
    });
  }

  async getAccount(): Promise<AccountInfo | null> {
    if (!this.pca) {
      await this.initialize();
    }
    
    const accounts = await this.pca!.getTokenCache().getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  async logout(): Promise<void> {
    if (!this.pca) {
      return;
    }

    const accounts = await this.pca.getTokenCache().getAllAccounts();
    for (const account of accounts) {
      await this.pca.getTokenCache().removeAccount(account);
    }

    // Also delete the cache file
    try {
      await fs.unlink(this.cacheFilePath);
    } catch (error) {
      // File might not exist
    }
  }

  isSecureCacheEnabled(): boolean {
    return this.useSecureCache;
  }
}
