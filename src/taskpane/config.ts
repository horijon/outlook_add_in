/*
 * Configuration settings for the GovStream Email Assistant add-in
 */

// Environment flag - set to 'development' for local testing, 'production' for deployment
export const ENVIRONMENT = 'development'; // or 'production'
// export const ENVIRONMENT = 'production'; // or 'development'

// Debug mode - set to true to enable console logs
export const DEBUG = true;

// API Endpoints
export const API_URLS = {
  production: {
    base: 'https://api.govstream.ai',
    auth: 'https://api.govstream.ai/auth/login',
    emailProcess: 'https://api.govstream.ai/email-assistant/process-email'
  },
  development: {
    base: 'http://localhost:3020',
    auth: 'http://localhost:3020/auth/login',
    emailProcess: 'http://localhost:3020/email-assistant/process-email'
  }
};

// Get the current API endpoints based on environment
export function getApiEndpoints() {
  if (DEBUG) {
    console.log(`Using ${ENVIRONMENT} environment`);
  }
  return ENVIRONMENT === 'production' ? API_URLS.production : API_URLS.development;
}

/*
 * Server-side Token Validation Guide
 * 
 * The token sent from the add-in is a JWT issued by Microsoft Exchange.
 * To validate it on your server:
 * 
 * 1. Verify the token signature using Microsoft's public keys
 *    - Keys available at: https://login.microsoftonline.com/common/discovery/keys
 * 
 * 2. Check the token hasn't expired (exp claim)
 * 
 * 3. Validate the audience (aud claim) matches your expected value
 *    - For Exchange tokens, this is usually "https://outlook.office.com/"
 * 
 * 4. Extract user information from claims:
 *    - User email: typically in the "upn" or "unique_name" claim
 *    - User ID: typically in the "oid" claim
 * 
 * 5. Use this validated identity to authenticate the user in your system
 * 
 * For more information, see:
 * https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/validate-an-identity-token
 */ 
