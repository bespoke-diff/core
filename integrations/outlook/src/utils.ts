import axios, { AxiosInstance } from 'axios';

export interface OutlookConfig {
  access_token: string;
  refresh_token: string;
  client_id: string;
  client_secret: string;
  expires_at?: number;
}

export interface EmailContent {
  textContent: string;
  htmlContent: string;
}

export interface OutlookMessage {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  toRecipients: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  receivedDateTime: string;
  sentDateTime: string;
  body: {
    contentType: string;
    content: string;
  };
  webLink: string;
  isRead: boolean;
  importance: string;
  hasAttachments: boolean;
}

export interface OutlookUser {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
}

const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';
const TOKEN_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

/**
 * Refresh the access token using the refresh token
 */
export async function refreshAccessToken(config: OutlookConfig): Promise<{
  access_token: string;
  refresh_token: string;
  expires_in: number;
}> {
  const params = new URLSearchParams({
    client_id: config.client_id,
    client_secret: config.client_secret,
    refresh_token: config.refresh_token,
    grant_type: 'refresh_token',
  });

  const response = await axios.post(TOKEN_URL, params.toString(), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
  });

  return {
    access_token: response.data.access_token,
    refresh_token: response.data.refresh_token || config.refresh_token,
    expires_in: response.data.expires_in,
  };
}

/**
 * Create an authenticated Microsoft Graph API client
 */
export async function getGraphClient(config: OutlookConfig): Promise<AxiosInstance> {
  let accessToken = config.access_token;

  // Check if token is expired or about to expire (within 5 minutes)
  if (config.expires_at) {
    const now = Date.now();
    const expiresAt = typeof config.expires_at === 'string'
      ? parseInt(config.expires_at)
      : config.expires_at;

    if (now >= expiresAt - 5 * 60 * 1000) {
      // Token is expired or about to expire, refresh it
      try {
        const refreshed = await refreshAccessToken(config);
        accessToken = refreshed.access_token;
      } catch (error) {
        console.error('Failed to refresh token:', error);
        // Continue with existing token, might still work
      }
    }
  }

  return axios.create({
    baseURL: GRAPH_API_BASE,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  });
}

/**
 * Get user profile information
 */
export async function getUserProfile(config: OutlookConfig): Promise<OutlookUser> {
  const client = await getGraphClient(config);
  const response = await client.get('/me');
  return response.data;
}

/**
 * Parse email content from Outlook message
 */
export function parseEmailContent(message: OutlookMessage): EmailContent {
  const body = message.body;

  if (body.contentType === 'html') {
    return {
      htmlContent: body.content,
      textContent: body.content.replace(/<[^>]*>/g, '').trim(),
    };
  }

  return {
    textContent: body.content,
    htmlContent: '',
  };
}

/**
 * Format email sender for display
 */
export function formatEmailSender(from: OutlookMessage['from']): string {
  if (from.emailAddress.name) {
    return from.emailAddress.name;
  }
  return from.emailAddress.address;
}

/**
 * Search for emails with a filter
 */
export async function searchEmails(
  config: OutlookConfig,
  filter: string,
  maxResults: number = 50,
  orderBy: string = 'receivedDateTime desc'
): Promise<OutlookMessage[]> {
  const client = await getGraphClient(config);

  const params: Record<string, string> = {
    $top: maxResults.toString(),
    $orderby: orderBy,
    $select: 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,body,webLink,isRead,importance,hasAttachments',
  };

  if (filter) {
    params.$filter = filter;
  }

  const response = await client.get('/me/messages', { params });
  return response.data.value || [];
}

/**
 * Get messages from a specific folder
 */
export async function getMessagesFromFolder(
  config: OutlookConfig,
  folderId: string,
  maxResults: number = 50,
  filter?: string
): Promise<OutlookMessage[]> {
  const client = await getGraphClient(config);

  const params: Record<string, string> = {
    $top: maxResults.toString(),
    $orderby: 'receivedDateTime desc',
    $select: 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,body,webLink,isRead,importance,hasAttachments',
  };

  if (filter) {
    params.$filter = filter;
  }

  const response = await client.get(`/me/mailFolders/${folderId}/messages`, { params });
  return response.data.value || [];
}

/**
 * Get a single message by ID
 */
export async function getMessage(
  config: OutlookConfig,
  messageId: string
): Promise<OutlookMessage> {
  const client = await getGraphClient(config);
  const response = await client.get(`/me/messages/${messageId}`, {
    params: {
      $select: 'id,subject,from,toRecipients,receivedDateTime,sentDateTime,body,webLink,isRead,importance,hasAttachments',
    },
  });
  return response.data;
}

/**
 * Send an email
 */
export async function sendEmail(
  config: OutlookConfig,
  to: string[],
  subject: string,
  body: string,
  options?: {
    cc?: string[];
    bcc?: string[];
    isHtml?: boolean;
    importance?: 'low' | 'normal' | 'high';
  }
): Promise<void> {
  const client = await getGraphClient(config);

  const message = {
    subject,
    body: {
      contentType: options?.isHtml ? 'html' : 'text',
      content: body,
    },
    toRecipients: to.map(email => ({
      emailAddress: { address: email },
    })),
    ccRecipients: options?.cc?.map(email => ({
      emailAddress: { address: email },
    })),
    bccRecipients: options?.bcc?.map(email => ({
      emailAddress: { address: email },
    })),
    importance: options?.importance || 'normal',
  };

  await client.post('/me/sendMail', {
    message,
    saveToSentItems: true,
  });
}

/**
 * Create a draft email
 */
export async function createDraft(
  config: OutlookConfig,
  to: string[],
  subject: string,
  body: string,
  options?: {
    cc?: string[];
    bcc?: string[];
    isHtml?: boolean;
    importance?: 'low' | 'normal' | 'high';
  }
): Promise<OutlookMessage> {
  const client = await getGraphClient(config);

  const message = {
    subject,
    body: {
      contentType: options?.isHtml ? 'html' : 'text',
      content: body,
    },
    toRecipients: to.map(email => ({
      emailAddress: { address: email },
    })),
    ccRecipients: options?.cc?.map(email => ({
      emailAddress: { address: email },
    })),
    bccRecipients: options?.bcc?.map(email => ({
      emailAddress: { address: email },
    })),
    importance: options?.importance || 'normal',
  };

  const response = await client.post('/me/messages', message);
  return response.data;
}

/**
 * Reply to an email
 */
export async function replyToEmail(
  config: OutlookConfig,
  messageId: string,
  body: string,
  replyAll: boolean = false
): Promise<void> {
  const client = await getGraphClient(config);

  const endpoint = replyAll
    ? `/me/messages/${messageId}/replyAll`
    : `/me/messages/${messageId}/reply`;

  await client.post(endpoint, {
    comment: body,
  });
}

/**
 * Forward an email
 */
export async function forwardEmail(
  config: OutlookConfig,
  messageId: string,
  to: string[],
  comment?: string
): Promise<void> {
  const client = await getGraphClient(config);

  await client.post(`/me/messages/${messageId}/forward`, {
    comment: comment || '',
    toRecipients: to.map(email => ({
      emailAddress: { address: email },
    })),
  });
}

/**
 * Mark email as read/unread
 */
export async function markAsRead(
  config: OutlookConfig,
  messageId: string,
  isRead: boolean = true
): Promise<void> {
  const client = await getGraphClient(config);

  await client.patch(`/me/messages/${messageId}`, {
    isRead,
  });
}

/**
 * Delete an email (move to deleted items)
 */
export async function deleteEmail(
  config: OutlookConfig,
  messageId: string
): Promise<void> {
  const client = await getGraphClient(config);
  await client.delete(`/me/messages/${messageId}`);
}

/**
 * Move email to a folder
 */
export async function moveEmail(
  config: OutlookConfig,
  messageId: string,
  destinationFolderId: string
): Promise<OutlookMessage> {
  const client = await getGraphClient(config);

  const response = await client.post(`/me/messages/${messageId}/move`, {
    destinationId: destinationFolderId,
  });

  return response.data;
}

/**
 * List mail folders
 */
export async function listFolders(config: OutlookConfig): Promise<Array<{
  id: string;
  displayName: string;
  parentFolderId: string;
  totalItemCount: number;
  unreadItemCount: number;
}>> {
  const client = await getGraphClient(config);
  const response = await client.get('/me/mailFolders', {
    params: {
      $top: 100,
    },
  });
  return response.data.value || [];
}
