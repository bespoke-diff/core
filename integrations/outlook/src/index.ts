import { integrationCreate } from './account-create';
import { handleSchedule } from './schedule';
import {
  IntegrationCLI,
  IntegrationEventPayload,
  IntegrationEventType,
  Spec,
} from '@redplanethq/sdk';
import { getTools, callTool } from './mcp';

export async function run(eventPayload: IntegrationEventPayload) {
  switch (eventPayload.event) {
    case IntegrationEventType.SETUP:
      return await integrationCreate(eventPayload.eventBody);

    case IntegrationEventType.SYNC:
      return await handleSchedule(eventPayload.config, eventPayload.state);

    case IntegrationEventType.GET_TOOLS: {
      const tools = await getTools();
      return tools;
    }

    case IntegrationEventType.CALL_TOOL: {
      const integrationDefinition = eventPayload.integrationDefinition;

      if (!integrationDefinition) {
        return null;
      }

      const config = eventPayload.config as any;
      const { name, arguments: args } = eventPayload.eventBody;

      const result = await callTool(
        name,
        args,
        integrationDefinition.config.clientId,
        integrationDefinition.config.clientSecret,
        config?.redirect_uri,
        config
      );

      return result;
    }

    default:
      return { message: `The event payload type is ${eventPayload.event}` };
  }
}

// CLI implementation that extends the base class
class OutlookCLI extends IntegrationCLI {
  constructor() {
    super('outlook', '1.0.0');
  }

  protected async handleEvent(eventPayload: IntegrationEventPayload): Promise<any> {
    return await run(eventPayload);
  }

  protected async getSpec(): Promise<Spec> {
    return {
      name: 'Outlook Email',
      key: 'outlook',
      description:
        'Connect your workspace to Outlook. Monitor emails, send messages, and manage your email workflow with Microsoft Outlook',
      icon: 'outlook',
      mcp: {
        type: 'cli',
      },
      schedule: {
        frequency: '*/15 * * * *',
      },
      auth: {
        OAuth2: {
          token_url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
          authorization_url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
          scopes: [
            'https://graph.microsoft.com/Mail.Read',
            'https://graph.microsoft.com/Mail.ReadWrite',
            'https://graph.microsoft.com/Mail.Send',
            'https://graph.microsoft.com/User.Read',
            'offline_access',
          ],
          scope_identifier: 'scope',
          scope_separator: ' ',
          token_params: {
            prompt: 'consent',
          },
          authorization_params: {
            prompt: 'consent',
            response_mode: 'query',
          },
        },
      },
    };
  }
}

// Define a main function and invoke it directly.
// This works after bundling to JS and running with `node index.js`.
function main() {
  const outlookCLI = new OutlookCLI();
  outlookCLI.parse();
}

main();
