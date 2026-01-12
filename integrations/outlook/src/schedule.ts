import {
  getGraphClient,
  parseEmailContent,
  formatEmailSender,
  OutlookConfig,
  OutlookMessage,
} from './utils';
import TurndownService from 'turndown';

interface OutlookSettings {
  lastSyncTime?: string;
  lastUserEventTime?: string;
  emailAddress?: string;
}

interface OutlookActivityCreateParams {
  text: string;
  sourceURL: string;
}

/**
 * Creates an activity message based on Outlook data
 */
function createActivityMessage(params: OutlookActivityCreateParams) {
  return {
    type: 'activity',
    data: {
      text: params.text,
      sourceURL: params.sourceURL,
    },
  };
}

/**
 * Gets default sync time (24 hours ago)
 */
function getDefaultSyncTime(): string {
  return new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
}

/**
 * Initialize Turndown service for HTML to Markdown conversion
 */
const turndownService = new TurndownService({
  headingStyle: 'atx',
  codeBlockStyle: 'fenced',
  emDelimiter: '*',
});

// Remove style, script, and other unwanted elements
turndownService.remove(['style', 'script', 'noscript', 'iframe', 'object', 'embed']);

/**
 * Clean and convert email content to markdown
 */
function cleanEmailContent(htmlContent: string, textContent: string): string {
  // If we have HTML content, convert it to markdown
  if (htmlContent) {
    const markdown = turndownService.turndown(htmlContent);
    return markdown
      .replace(/\n\n+/g, '\n\n') // Remove excessive line breaks
      .replace(/\s+/g, ' ') // Normalize spaces
      .trim();
  }

  // Otherwise use text content and clean it
  return textContent
    .replace(/\r/g, '')
    .replace(/\n\n+/g, '\n\n')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Format ISO date for Microsoft Graph API filter
 */
function toGraphApiDateFilter(isoDate: string): string {
  // Microsoft Graph API uses ISO 8601 format for datetime filters
  return new Date(isoDate).toISOString();
}

/**
 * Fetch and process received emails from inbox
 */
async function processReceivedEmails(
  client: any,
  lastSyncTime: string,
  emailAddress: string
): Promise<any[]> {
  const activities = [];
  const afterDate = toGraphApiDateFilter(lastSyncTime);

  try {
    // Query for received emails after lastSyncTime
    // Using importance filter to get important emails like Gmail does
    const response = await client.get('/me/mailFolders/inbox/messages', {
      params: {
        $filter: `receivedDateTime ge ${afterDate} and importance eq 'high'`,
        $top: 50,
        $orderby: 'receivedDateTime desc',
        $select: 'id,subject,from,toRecipients,receivedDateTime,body,webLink,importance',
      },
    });

    const messages: OutlookMessage[] = response.data.value || [];

    // If no high importance emails, get recent inbox emails instead
    if (messages.length === 0) {
      const recentResponse = await client.get('/me/mailFolders/inbox/messages', {
        params: {
          $filter: `receivedDateTime ge ${afterDate}`,
          $top: 25,
          $orderby: 'receivedDateTime desc',
          $select: 'id,subject,from,toRecipients,receivedDateTime,body,webLink,importance',
        },
      });
      messages.push(...(recentResponse.data.value || []));
    }

    for (const message of messages) {
      try {
        const sender = formatEmailSender(message.from);
        const subject = message.subject || '(No subject)';
        const { textContent, htmlContent } = parseEmailContent(message);

        // Clean and convert email content to markdown
        const cleanedContent = cleanEmailContent(htmlContent, textContent);

        // Skip if no meaningful content
        if (!cleanedContent || cleanedContent.length < 10) {
          continue;
        }

        // Use the webLink from Outlook
        const sourceURL = message.webLink;

        // Format activity text with full email content as markdown
        const text = `## Email from ${sender}

**Subject:** ${subject}

${cleanedContent}`;

        activities.push(
          createActivityMessage({
            text,
            sourceURL,
          })
        );
      } catch (error) {
        // Silently ignore errors for individual messages
        console.error('Error processing received email:', error);
      }
    }
  } catch (error) {
    console.error('Error fetching received emails:', error);
  }

  return activities;
}

/**
 * Fetch and process sent emails
 */
async function processSentEmails(
  client: any,
  lastSyncTime: string,
  emailAddress: string
): Promise<any[]> {
  const activities = [];
  const afterDate = toGraphApiDateFilter(lastSyncTime);

  try {
    // Query for sent emails after lastSyncTime
    const response = await client.get('/me/mailFolders/sentitems/messages', {
      params: {
        $filter: `sentDateTime ge ${afterDate}`,
        $top: 50,
        $orderby: 'sentDateTime desc',
        $select: 'id,subject,from,toRecipients,sentDateTime,body,webLink',
      },
    });

    const messages: OutlookMessage[] = response.data.value || [];

    for (const message of messages) {
      try {
        const recipients = message.toRecipients
          .map(r => r.emailAddress.name || r.emailAddress.address)
          .join(', ');
        const subject = message.subject || '(No subject)';
        const { textContent, htmlContent } = parseEmailContent(message);

        // Clean and convert email content to markdown
        const cleanedContent = cleanEmailContent(htmlContent, textContent);

        // Skip if no meaningful content
        if (!cleanedContent || cleanedContent.length < 10) {
          continue;
        }

        // Use the webLink from Outlook
        const sourceURL = message.webLink;

        // Format activity text with full email content as markdown
        const text = `## Sent to ${recipients}

**Subject:** ${subject}

${cleanedContent}`;

        activities.push(
          createActivityMessage({
            text,
            sourceURL,
          })
        );
      } catch (error) {
        // Silently ignore errors for individual messages
        console.error('Error processing sent email:', error);
      }
    }
  } catch (error) {
    console.error('Error fetching sent emails:', error);
  }

  return activities;
}

export const handleSchedule = async (
  config?: Record<string, string>,
  state?: Record<string, string>
) => {
  try {
    // Check if we have a valid access token
    if (!config?.access_token) {
      return [];
    }

    // Get settings or initialize if not present
    let settings = (state || {}) as OutlookSettings;

    // Default to 24 hours ago if no last sync time
    const lastSyncTime = settings.lastSyncTime || getDefaultSyncTime();

    // Create Microsoft Graph client
    const outlookConfig: OutlookConfig = {
      access_token: config.access_token,
      refresh_token: config.refresh_token || '',
      client_id: config.client_id || '',
      client_secret: config.client_secret || '',
      expires_at: config.expires_at ? parseInt(config.expires_at) : undefined,
    };

    const client = await getGraphClient(outlookConfig);

    // Get user profile to get email address
    if (!settings.emailAddress) {
      try {
        const profileResponse = await client.get('/me');
        settings.emailAddress = profileResponse.data.mail || profileResponse.data.userPrincipalName;
      } catch (error) {
        console.error('Error fetching user profile:', error);
      }
    }

    // Collect all messages
    const messages = [];

    // Process received emails
    const receivedActivities = await processReceivedEmails(
      client,
      lastSyncTime,
      settings.emailAddress || 'user'
    );
    messages.push(...receivedActivities);

    // Process sent emails
    const sentActivities = await processSentEmails(
      client,
      lastSyncTime,
      settings.emailAddress || 'user'
    );
    messages.push(...sentActivities);

    // Update last sync time
    const newSyncTime = new Date().toISOString();

    // Add state message for saving settings
    messages.push({
      type: 'state',
      data: {
        ...settings,
        lastSyncTime: newSyncTime,
        lastUserEventTime: newSyncTime,
      },
    });

    return messages;
  } catch (error) {
    console.error('Error in handleSchedule:', error);
    return [];
  }
};
