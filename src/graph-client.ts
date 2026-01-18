import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

export interface GraphClientConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

export class GraphClientManager {
  private client: Client;

  constructor(config: GraphClientConfig) {
    const credential = new ClientSecretCredential(
      config.tenantId,
      config.clientId,
      config.clientSecret
    );

    this.client = Client.initWithMiddleware({
      authProvider: {
        getAccessToken: async () => {
          const token = await credential.getToken('https://graph.microsoft.com/.default');
          return token.token;
        },
      },
    });
  }

  getClient(): Client {
    return this.client;
  }
}
