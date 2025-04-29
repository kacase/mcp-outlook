import { Configuration, LogLevel } from '@azure/msal-node';

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/configuration.md
 * 
 * Required environment variables:
 * - CLIENT_ID: The Application (client) ID from your Azure AD app registration
 * - AUTHORITY: Your Azure AD tenant ID
 */
export const msalConfig: Configuration = {
    auth: {
        clientId: process.env.CLIENT_ID!,
        authority: `https://login.microsoftonline.com/${process.env.AUTHORITY!}`,   
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel: any, message: any, containsPii: any) {
                // Show log messages if DEBUG is enabled
                if (process.env.DEBUG && !containsPii) {
                    console.log(`MSAL: ${message}`);
                }
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Error,
        },
    },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
    scopes: [
        'User.Read',
        'Calendars.Read',
        'Calendars.ReadWrite',
        'Mail.Send',
        'Mail.ReadWrite',
        'Mail.Read',
    ],
};