/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import { PublicClientApplication, InteractionRequiredAuthError, AuthenticationResult } from '@azure/msal-node';
import open from 'open';
import { msalConfig, loginRequest } from './authConfig.js';

// Before running the sample, you will need to replace the values in src/authConfig.js

// Open browser to sign user in and consent to scopes needed for application
const openBrowser = async (url: string) => {
    // You can open a browser window with any library or method you wish to use - the 'open' npm package is used here for demonstration purposes.
    open(url);
};

const tokenRequest = {
    ...loginRequest,
    openBrowser,
    successTemplate: '<h1>Successfully signed in!</h1> <p>You can close this window now.',
    errorTemplate:
        '<h1>Oops! Something went wrong</h1> <p>Navigate back to the application and check the console for more information.</p>',
};

/**
 * Initialize a public client application. For more information, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/docs/initialize-public-client-application.md
 */
export const pca = new PublicClientApplication(msalConfig);

// Cache to store the authentication result
let authResultCache: AuthenticationResult | null = null;

/**
 * Get the access token either from cache or by authenticating
 */
export const getAccessToken = async (): Promise<string> => {
    // If we have a cached token, return it
    if (authResultCache && authResultCache.expiresOn && authResultCache.expiresOn > new Date()) {
        return authResultCache.accessToken;
    }
    
    // Otherwise, get a new token
    const authResult = await acquireToken();
    if (!authResult) {
        throw new Error('Failed to acquire access token');
    }
    
    // Cache the result
    authResultCache = authResult;
    return authResult.accessToken;
};

/**
 * Check if the user is authenticated
 */
export const isAuthenticated = async (): Promise<boolean> => {
    const accounts = await pca.getTokenCache().getAllAccounts();
    return accounts.length > 0;
};

/**
 * Acquire an access token
 */
export const acquireToken = async (): Promise<AuthenticationResult | undefined> => {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length === 1) {
        const silentRequest = {
            account: accounts[0],
            scopes: loginRequest.scopes,
        };

        return pca.acquireTokenSilent(silentRequest).catch((e) => {
            if (e instanceof InteractionRequiredAuthError) {
                return pca.acquireTokenInteractive(tokenRequest);
            }
            throw e;
        });
    } else if (accounts.length > 1) {
        // Multiple accounts found
        throw new Error('Multiple accounts found. Please select an account to use.');
    } else {
        return pca.acquireTokenInteractive(tokenRequest);
    }
};
