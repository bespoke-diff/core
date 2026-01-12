import { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import {
  getGraphClient,
  OutlookConfig,
  OutlookMessage,
  parseEmailContent,
} from '../utils';

// Schema definitions for Outlook MCP tools
const SendEmailSchema = z.object({
  to: z.array(z.string()).describe('List of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body content'),
  htmlBody: z.string().optional().describe('HTML version of the email body'),
  cc: z.array(z.string()).optional().describe('List of CC recipients'),
  bcc: z.array(z.string()).optional().describe('List of BCC recipients'),
  importance: z
    .enum(['low', 'normal', 'high'])
    .optional()
    .default('normal')
    .describe('Email importance level'),
});

const ReadEmailSchema = z.object({
  messageId: z.string().describe('ID of the email message to retrieve'),
});

const SearchEmailsSchema = z.object({
  query: z.string().describe('Search query text to find in email subject or body'),
  folder: z
    .string()
    .optional()
    .describe('Folder to search in (inbox, sentitems, drafts, deleteditems). Defaults to all folders'),
  maxResults: z.number().optional().default(10).describe('Maximum number of results to return'),
});

const ListEmailsSchema = z.object({
  folder: z
    .enum(['inbox', 'sentitems', 'drafts', 'deleteditems', 'archive'])
    .optional()
    .default('inbox')
    .describe('Mail folder to list emails from'),
  maxResults: z.number().optional().default(10).describe('Maximum number of results to return'),
  unreadOnly: z.boolean().optional().describe('Only return unread emails'),
});

const ReplyEmailSchema = z.object({
  messageId: z.string().describe('ID of the email message to reply to'),
  body: z.string().describe('Reply body content'),
  replyAll: z.boolean().optional().default(false).describe('Reply to all recipients'),
});

const ForwardEmailSchema = z.object({
  messageId: z.string().describe('ID of the email message to forward'),
  to: z.array(z.string()).describe('List of recipient email addresses to forward to'),
  comment: z.string().optional().describe('Optional comment to include with the forwarded email'),
});

const DeleteEmailSchema = z.object({
  messageId: z.string().describe('ID of the email message to delete'),
});

const MarkAsReadSchema = z.object({
  messageId: z.string().describe('ID of the email message'),
  isRead: z.boolean().default(true).describe('Mark as read (true) or unread (false)'),
});

const MoveEmailSchema = z.object({
  messageId: z.string().describe('ID of the email message to move'),
  destinationFolder: z
    .enum(['inbox', 'archive', 'deleteditems', 'drafts', 'junkemail'])
    .describe('Destination folder'),
});

const ListFoldersSchema = z.object({}).describe('List all mail folders');

const CreateDraftSchema = z.object({
  to: z.array(z.string()).describe('List of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body content'),
  htmlBody: z.string().optional().describe('HTML version of the email body'),
  cc: z.array(z.string()).optional().describe('List of CC recipients'),
  bcc: z.array(z.string()).optional().describe('List of BCC recipients'),
  importance: z
    .enum(['low', 'normal', 'high'])
    .optional()
    .default('normal')
    .describe('Email importance level'),
});

/**
 * Get list of available tools
 */
export async function getTools() {
  const tools = [
    {
      name: 'send_email',
      description: 'Sends a new email via Outlook',
      inputSchema: zodToJsonSchema(SendEmailSchema),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
    },
    {
      name: 'create_draft',
      description: 'Creates a draft email in Outlook',
      inputSchema: zodToJsonSchema(CreateDraftSchema),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false },
    },
    {
      name: 'read_email',
      description: 'Retrieves the content of a specific email',
      inputSchema: zodToJsonSchema(ReadEmailSchema),
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    },
    {
      name: 'search_emails',
      description: 'Searches for emails using text query',
      inputSchema: zodToJsonSchema(SearchEmailsSchema),
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    },
    {
      name: 'list_emails',
      description: 'Lists emails from a specific folder',
      inputSchema: zodToJsonSchema(ListEmailsSchema),
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    },
    {
      name: 'reply_email',
      description: 'Replies to an email',
      inputSchema: zodToJsonSchema(ReplyEmailSchema),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
    },
    {
      name: 'forward_email',
      description: 'Forwards an email to specified recipients',
      inputSchema: zodToJsonSchema(ForwardEmailSchema),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false },
    },
    {
      name: 'delete_email',
      description: 'Deletes an email (moves to deleted items)',
      inputSchema: zodToJsonSchema(DeleteEmailSchema),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true },
    },
    {
      name: 'mark_as_read',
      description: 'Marks an email as read or unread',
      inputSchema: zodToJsonSchema(MarkAsReadSchema),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true },
    },
    {
      name: 'move_email',
      description: 'Moves an email to a different folder',
      inputSchema: zodToJsonSchema(MoveEmailSchema),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true },
    },
    {
      name: 'list_folders',
      description: 'Lists all mail folders',
      inputSchema: zodToJsonSchema(ListFoldersSchema),
      annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true },
    },
  ];

  return tools;
}

// Map folder names to well-known folder IDs
const folderIdMap: Record<string, string> = {
  inbox: 'inbox',
  sentitems: 'sentitems',
  drafts: 'drafts',
  deleteditems: 'deleteditems',
  archive: 'archive',
  junkemail: 'junkemail',
};

/**
 * Call a specific tool
 */
export async function callTool(
  name: string,
  args: Record<string, any>,
  client_id: string,
  client_secret: string,
  redirect_uri: string,
  credentials: Record<string, string>
) {
  const config: OutlookConfig = {
    access_token: credentials.access_token,
    refresh_token: credentials.refresh_token,
    client_id,
    client_secret,
    expires_at: credentials.expires_at ? parseInt(credentials.expires_at) : undefined,
  };

  const client = await getGraphClient(config);

  try {
    switch (name) {
      case 'send_email': {
        const validatedArgs = SendEmailSchema.parse(args);

        const message = {
          subject: validatedArgs.subject,
          body: {
            contentType: validatedArgs.htmlBody ? 'html' : 'text',
            content: validatedArgs.htmlBody || validatedArgs.body,
          },
          toRecipients: validatedArgs.to.map(email => ({
            emailAddress: { address: email },
          })),
          ccRecipients: validatedArgs.cc?.map(email => ({
            emailAddress: { address: email },
          })),
          bccRecipients: validatedArgs.bcc?.map(email => ({
            emailAddress: { address: email },
          })),
          importance: validatedArgs.importance,
        };

        await client.post('/me/sendMail', {
          message,
          saveToSentItems: true,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email sent successfully to: ${validatedArgs.to.join(', ')}`,
            },
          ],
        };
      }

      case 'create_draft': {
        const validatedArgs = CreateDraftSchema.parse(args);

        const message = {
          subject: validatedArgs.subject,
          body: {
            contentType: validatedArgs.htmlBody ? 'html' : 'text',
            content: validatedArgs.htmlBody || validatedArgs.body,
          },
          toRecipients: validatedArgs.to.map(email => ({
            emailAddress: { address: email },
          })),
          ccRecipients: validatedArgs.cc?.map(email => ({
            emailAddress: { address: email },
          })),
          bccRecipients: validatedArgs.bcc?.map(email => ({
            emailAddress: { address: email },
          })),
          importance: validatedArgs.importance,
        };

        const response = await client.post('/me/messages', message);

        return {
          content: [
            {
              type: 'text',
              text: `Draft created successfully with ID: ${response.data.id}`,
            },
          ],
        };
      }

      case 'read_email': {
        const validatedArgs = ReadEmailSchema.parse(args);

        const response = await client.get(`/me/messages/${validatedArgs.messageId}`, {
          params: {
            $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,webLink,importance,hasAttachments',
          },
        });

        const message: OutlookMessage = response.data;
        const from = message.from?.emailAddress?.address || 'Unknown';
        const fromName = message.from?.emailAddress?.name || from;
        const to = message.toRecipients?.map(r => r.emailAddress.address).join(', ') || '';
        const { textContent, htmlContent } = parseEmailContent(message);
        const body = textContent || htmlContent.replace(/<[^>]*>/g, '');

        return {
          content: [
            {
              type: 'text',
              text: `Subject: ${message.subject}\nFrom: ${fromName} <${from}>\nTo: ${to}\nDate: ${message.receivedDateTime}\nImportance: ${message.importance}\n\n${body}`,
            },
          ],
        };
      }

      case 'search_emails': {
        const validatedArgs = SearchEmailsSchema.parse(args);

        let endpoint = '/me/messages';
        if (validatedArgs.folder) {
          const folderId = folderIdMap[validatedArgs.folder.toLowerCase()] || validatedArgs.folder;
          endpoint = `/me/mailFolders/${folderId}/messages`;
        }

        const response = await client.get(endpoint, {
          params: {
            $search: `"${validatedArgs.query}"`,
            $top: validatedArgs.maxResults,
            $select: 'id,subject,from,receivedDateTime,isRead',
            $orderby: 'receivedDateTime desc',
          },
        });

        const messages = response.data.value || [];

        if (messages.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: 'No emails found matching your search.',
              },
            ],
          };
        }

        const results = messages.map((msg: any) => {
          const from = msg.from?.emailAddress?.address || 'Unknown';
          return `ID: ${msg.id}\nSubject: ${msg.subject}\nFrom: ${from}\nDate: ${msg.receivedDateTime}\nRead: ${msg.isRead ? 'Yes' : 'No'}\n`;
        });

        return {
          content: [
            {
              type: 'text',
              text: `Found ${messages.length} emails:\n\n${results.join('\n')}`,
            },
          ],
        };
      }

      case 'list_emails': {
        const validatedArgs = ListEmailsSchema.parse(args);
        const folderId = folderIdMap[validatedArgs.folder] || validatedArgs.folder;

        let filter = '';
        if (validatedArgs.unreadOnly) {
          filter = 'isRead eq false';
        }

        const response = await client.get(`/me/mailFolders/${folderId}/messages`, {
          params: {
            $top: validatedArgs.maxResults,
            $select: 'id,subject,from,receivedDateTime,isRead,importance',
            $orderby: 'receivedDateTime desc',
            ...(filter && { $filter: filter }),
          },
        });

        const messages = response.data.value || [];

        if (messages.length === 0) {
          return {
            content: [
              {
                type: 'text',
                text: `No emails found in ${validatedArgs.folder}.`,
              },
            ],
          };
        }

        const results = messages.map((msg: any) => {
          const from = msg.from?.emailAddress?.address || 'Unknown';
          return `ID: ${msg.id}\nSubject: ${msg.subject}\nFrom: ${from}\nDate: ${msg.receivedDateTime}\nRead: ${msg.isRead ? 'Yes' : 'No'}\n`;
        });

        return {
          content: [
            {
              type: 'text',
              text: `${messages.length} emails in ${validatedArgs.folder}:\n\n${results.join('\n')}`,
            },
          ],
        };
      }

      case 'reply_email': {
        const validatedArgs = ReplyEmailSchema.parse(args);

        const endpoint = validatedArgs.replyAll
          ? `/me/messages/${validatedArgs.messageId}/replyAll`
          : `/me/messages/${validatedArgs.messageId}/reply`;

        await client.post(endpoint, {
          comment: validatedArgs.body,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Reply sent successfully${validatedArgs.replyAll ? ' to all recipients' : ''}.`,
            },
          ],
        };
      }

      case 'forward_email': {
        const validatedArgs = ForwardEmailSchema.parse(args);

        await client.post(`/me/messages/${validatedArgs.messageId}/forward`, {
          comment: validatedArgs.comment || '',
          toRecipients: validatedArgs.to.map(email => ({
            emailAddress: { address: email },
          })),
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email forwarded successfully to: ${validatedArgs.to.join(', ')}`,
            },
          ],
        };
      }

      case 'delete_email': {
        const validatedArgs = DeleteEmailSchema.parse(args);

        await client.delete(`/me/messages/${validatedArgs.messageId}`);

        return {
          content: [
            {
              type: 'text',
              text: `Email ${validatedArgs.messageId} deleted successfully.`,
            },
          ],
        };
      }

      case 'mark_as_read': {
        const validatedArgs = MarkAsReadSchema.parse(args);

        await client.patch(`/me/messages/${validatedArgs.messageId}`, {
          isRead: validatedArgs.isRead,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email marked as ${validatedArgs.isRead ? 'read' : 'unread'}.`,
            },
          ],
        };
      }

      case 'move_email': {
        const validatedArgs = MoveEmailSchema.parse(args);
        const destinationId = folderIdMap[validatedArgs.destinationFolder] || validatedArgs.destinationFolder;

        await client.post(`/me/messages/${validatedArgs.messageId}/move`, {
          destinationId,
        });

        return {
          content: [
            {
              type: 'text',
              text: `Email moved to ${validatedArgs.destinationFolder}.`,
            },
          ],
        };
      }

      case 'list_folders': {
        const response = await client.get('/me/mailFolders', {
          params: {
            $top: 100,
            $select: 'id,displayName,totalItemCount,unreadItemCount',
          },
        });

        const folders = response.data.value || [];

        const results = folders.map((folder: any) =>
          `Name: ${folder.displayName}\nID: ${folder.id}\nTotal: ${folder.totalItemCount}\nUnread: ${folder.unreadItemCount}\n`
        );

        return {
          content: [
            {
              type: 'text',
              text: `Found ${folders.length} folders:\n\n${results.join('\n')}`,
            },
          ],
        };
      }

      default:
        return {
          content: [
            {
              type: 'text',
              text: `Unknown tool: ${name}`,
            },
          ],
        };
    }
  } catch (error: any) {
    const errorMessage = error.response?.data?.error?.message || error.message;
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${errorMessage}`,
        },
      ],
    };
  }
}
