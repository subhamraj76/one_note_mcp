import * as vscode from 'vscode';
import * as path from 'path';

let outputChannel: vscode.OutputChannel;

export function activate(context: vscode.ExtensionContext) {
  outputChannel = vscode.window.createOutputChannel('OneNote MCP');
  outputChannel.appendLine('OneNote MCP Extension activated');

  const didChangeEmitter = new vscode.EventEmitter<void>();

  // Register the MCP server definition provider
  const mcpProvider = vscode.lm.registerMcpServerDefinitionProvider('onenoteMcp', {
    onDidChangeMcpServerDefinitions: didChangeEmitter.event,

    provideMcpServerDefinitions: async () => {
      const cacheDir = await getCacheDirectory(context);
      
      if (!cacheDir) {
        vscode.window.showErrorMessage(
          'OneNote MCP: No workspace folder available. Please open a folder first.'
        );
        return [];
      }

      const serverPath = context.asAbsolutePath(path.join('dist', 'server.js'));
      
      outputChannel.appendLine(`Server path: ${serverPath}`);
      outputChannel.appendLine(`Cache directory: ${cacheDir}`);

      return [
        new vscode.McpStdioServerDefinition(
          'OneNote MCP Server',
          'node',
          [serverPath],
          {
            ONENOTE_MCP_CACHE_DIR: cacheDir
          }
        )
      ];
    },

    resolveMcpServerDefinition: async (definition) => {
      outputChannel.appendLine('Resolving MCP server definition...');
      return definition;
    }
  });

  context.subscriptions.push(mcpProvider);
  context.subscriptions.push(didChangeEmitter);

  // Register login command - forces re-authentication
  const loginCommand = vscode.commands.registerCommand('onenote-mcp.login', async () => {
    const cacheDir = await getCacheDirectory(context);
    
    if (cacheDir) {
      const cacheFile = path.join(cacheDir, 'onenote-mcp-cache.json');
      
      // Delete existing cache to force re-auth
      try {
        await vscode.workspace.fs.delete(vscode.Uri.file(cacheFile));
        outputChannel.appendLine('Existing token cache cleared for re-authentication');
      } catch (error) {
        // File might not exist, that's OK
      }
    }
    
    // Notify that server definitions changed (forces reconnect which triggers auth)
    didChangeEmitter.fire();
    
    vscode.window.showInformationMessage(
      'OneNote MCP: Ready to sign in. Use any OneNote tool now and a browser window will open for authentication.'
    );
  });

  // Register logout command
  const logoutCommand = vscode.commands.registerCommand('onenote-mcp.logout', async () => {
    const cacheDir = await getCacheDirectory(context);
    
    if (cacheDir) {
      const cacheFile = path.join(cacheDir, 'onenote-mcp-cache.json');
      
      try {
        await vscode.workspace.fs.delete(vscode.Uri.file(cacheFile));
        vscode.window.showInformationMessage('OneNote MCP: Successfully signed out. Token cache cleared.');
        outputChannel.appendLine('Token cache cleared');
      } catch (error) {
        // File might not exist
        vscode.window.showInformationMessage('OneNote MCP: No active session to sign out from.');
      }
    }
    
    // Notify that server definitions changed (forces reconnect)
    didChangeEmitter.fire();
  });

  // Register auth status command
  const statusCommand = vscode.commands.registerCommand('onenote-mcp.status', async () => {
    const cacheDir = await getCacheDirectory(context);
    
    if (!cacheDir) {
      vscode.window.showWarningMessage('OneNote MCP: No workspace open. Cannot determine auth status.');
      return;
    }
    
    const cacheFile = path.join(cacheDir, 'onenote-mcp-cache.json');
    
    try {
      const cacheUri = vscode.Uri.file(cacheFile);
      const cacheData = await vscode.workspace.fs.readFile(cacheUri);
      const cacheJson = JSON.parse(Buffer.from(cacheData).toString('utf-8'));
      
      // Check if there are any accounts in the cache
      let accountName = 'Unknown';
      let hasValidToken = false;
      
      if (cacheJson.Account) {
        const accounts = Object.values(cacheJson.Account) as any[];
        if (accounts.length > 0) {
          accountName = accounts[0].username || accounts[0].name || 'Unknown';
          hasValidToken = true;
        }
      }
      
      if (hasValidToken) {
        vscode.window.showInformationMessage(
          `OneNote MCP: ✅ Authenticated as ${accountName}`,
          'Sign Out'
        ).then(selection => {
          if (selection === 'Sign Out') {
            vscode.commands.executeCommand('onenote-mcp.logout');
          }
        });
        outputChannel.appendLine(`Auth status: Authenticated as ${accountName}`);
      } else {
        vscode.window.showWarningMessage(
          'OneNote MCP: ⚠️ No valid session found. Use a OneNote tool to sign in.',
          'Sign In Now'
        ).then(selection => {
          if (selection === 'Sign In Now') {
            vscode.commands.executeCommand('onenote-mcp.login');
          }
        });
        outputChannel.appendLine('Auth status: No valid session');
      }
    } catch (error) {
      vscode.window.showWarningMessage(
        'OneNote MCP: ⚠️ Not signed in. Use a OneNote tool to sign in.',
        'Sign In Now'
      ).then(selection => {
        if (selection === 'Sign In Now') {
          vscode.commands.executeCommand('onenote-mcp.login');
        }
      });
      outputChannel.appendLine('Auth status: No cache file found');
    }
  });

  context.subscriptions.push(loginCommand);
  context.subscriptions.push(logoutCommand);
  context.subscriptions.push(statusCommand);
  context.subscriptions.push(outputChannel);

  outputChannel.appendLine('OneNote MCP Extension ready');
}

async function getCacheDirectory(context: vscode.ExtensionContext): Promise<string | undefined> {
  const workspaceFolders = vscode.workspace.workspaceFolders;

  if (!workspaceFolders || workspaceFolders.length === 0) {
    // Fall back to extension's global storage
    return context.globalStorageUri.fsPath;
  }

  if (workspaceFolders.length === 1) {
    // Single folder workspace - use .vscode in that folder
    return path.join(workspaceFolders[0].uri.fsPath, '.vscode');
  }

  // Multi-root workspace - prompt user to select
  const selected = await vscode.window.showWorkspaceFolderPick({
    placeHolder: 'Select a workspace folder to store OneNote MCP token cache'
  });

  if (selected) {
    return path.join(selected.uri.fsPath, '.vscode');
  }

  // User cancelled - fall back to global storage
  return context.globalStorageUri.fsPath;
}

export function deactivate() {
  if (outputChannel) {
    outputChannel.appendLine('OneNote MCP Extension deactivated');
    outputChannel.dispose();
  }
}
