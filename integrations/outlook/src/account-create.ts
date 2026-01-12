import axios from 'axios';

export async function integrationCreate(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  data: any
) {
  const { oauthResponse, oauthParams } = data;

  // Fetch user information using the access token from Microsoft Graph API
  let userEmail = null;
  let userId = null;
  let displayName = null;

  try {
    const userInfoResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
      headers: {
        Authorization: `Bearer ${oauthResponse.access_token}`,
      },
    });

    userEmail = userInfoResponse.data.mail || userInfoResponse.data.userPrincipalName;
    userId = userInfoResponse.data.id;
    displayName = userInfoResponse.data.displayName;
  } catch (error) {
    console.error('Error fetching user info from Microsoft Graph:', error);
  }

  // Calculate expiration timestamp
  const expiresAt = oauthResponse.expires_at ||
    (oauthResponse.expires_in
      ? Date.now() + oauthResponse.expires_in * 1000
      : undefined);

  // For Outlook OAuth2, we need to store the tokens and user info
  const integrationConfiguration = {
    access_token: oauthResponse.access_token,
    refresh_token: oauthResponse.refresh_token,
    client_id: oauthResponse.client_id,
    client_secret: oauthResponse.client_secret,
    token_type: oauthResponse.token_type,
    expires_in: oauthResponse.expires_in,
    expires_at: expiresAt,
    scope: oauthResponse.scope,
    userEmail: userEmail,
    userId: userId,
    displayName: displayName,
    redirect_uri: oauthParams.redirect_uri || null,
  };

  const payload = {
    settings: {},
    accountId: integrationConfiguration.userEmail || integrationConfiguration.userId,
    config: integrationConfiguration,
  };

  return [
    {
      type: 'account',
      data: payload,
    },
  ];
}
