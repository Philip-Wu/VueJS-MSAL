// ----------------------------------------------------------------------------
// Copyright (c) Ben Coleman, 2021
// Licensed under the MIT License.
//  Modified by Philip Wu
//
// Drop in MSAL.js 2.x service wrapper & helper for SPAs
//   v2.1.0 - Ben Coleman 2019
//   Updated 2021 - Switched to @azure/msal-browser
// Copied from: https://github.com/benc-uk/msal-graph-vue/blob/master/src/services/auth.js
// ----------------------------------------------------------------------------

import * as msal from '@azure/msal-browser'
import store from '../store/index'

const config = {
    auth: {
        tenantId: 'e37d725c-ab5c-4624-9ae5-f0533e486437',
        redirectUri: process.env.VUE_APP_AZURE_AUTH_REDIRECT_URI,
        authority: 'https://login.microsoftonline.com/e37d725c-ab5c-4624-9ae5-f0533e486437',
        // authority: 'https://login.microsoftonline.com/common'  
        clientId: 'da47894d-ea78-4935-a83a-4a5bc4efb0fc',
    },
    cache: {
      cacheLocation: 'localStorage'
    },
    // Only uncomment when you *really* need to debug what is going on in MSAL
     system: {
      logger: new msal.Logger(
        (logLevel, msg) => { console.log(msg) },
        {
          level: msal.LogLevel.Verbose
        }
      )
    } 
  }

// MSAL object used for signing in users with MS identity platform
let msalApp

export default {
  waitingOnAccessToken: false,
  accessToken: null,
  accessTokenCallbacks: [], // functions to call back when accessToken is set

  //
  // Configure with clientId or empty string/null to set in "demo" mode
  //
  async configure() {
    console.log('authAzure.configure()');
    // Can only call configure once
    if (msalApp) {
        console.log('msalApp already exists');
        return
    } 


    console.log('### Azure AD sign-in: enabled\n', config)

    // Create our shared/static MSAL app object
    msalApp = new msal.PublicClientApplication(config)
  },

  //
  // Return the configured client id
  //
  clientId() {
    if (!msalApp) {
      return null
    }

    return msalApp.clientId
  },

  //
  // Login a user with a popup
  //
  async login(scopes = ['user.read', 'openid', 'profile']) {
    //scopes = this.defaultScope()    
    if (!msalApp) {
      return
    }

    //const LOGIN_SCOPES = ['user.read', 'openid', 'profile', 'email']
    await msalApp.loginPopup({
      scopes,
      prompt: 'select_account'
    })
  },

  //
  // Logout any stored user
  //
  logout() {
    if (!msalApp) {
      return
    }

    //msalApp.logoutPopup()
    msalApp.logoutRedirect()
  },

  //
  // Call to get user, probably cached and stored locally by MSAL
  //
  user() {
    if (!msalApp) {
      return null
    }

    const currentAccounts = msalApp.getAllAccounts()
    
    if (!currentAccounts || currentAccounts.length === 0) {
      //  console.log('no currentAccounts');
      // No user signed in
      return null
    } else if (currentAccounts.length > 1) {
      return currentAccounts[0]
    } else {
      return currentAccounts[0]
    }
  },

  isLoggedIn() {
      return this.user() != null
  },

  /**
   * As of June 2021.
   * This is super important. Otherwise, we get an error about 'Invalid signature' due to receiving an access token in v1.0 format.
   * The backend API will attempt to validate the token using a 2.0 endpoint, which is not suited for an v1.0 access token. So to force
   * v2.0 accessToken, we use the ./default. This was not mentioned in any formal documentation.
   * @returns 
   */
  defaultScope() {
    return [config.auth.clientId + '/.default']
    //return ['User.Read', 'profile', 'openid', 'api://da47894d-ea78-4935-a83a-4a5bc4efb0fc/Custom.API']
    //return ['api://da47894d-ea78-4935-a83a-4a5bc4efb0fc/Custom.API']
  },

  //
  // Call through to acquireTokenSilent or acquireTokenPopup
  //
  async acquireToken(/*scopes = ['user.read']*/) {
    this.waitingOnAccessToken = true
    // Override any scope
    let scopes = this.defaultScope()
    if (!msalApp) {
      return null
    }

    // Set scopes for token request
    const accessTokenRequest = {
      scopes,
      account: this.user()
    }

    let tokenResp
    try {
      // 1. Try to acquire token silently
      tokenResp = await msalApp.acquireTokenSilent(accessTokenRequest)
      console.log('### MSAL acquireTokenSilent was successful')
    } catch (err) {
      // 2. Silent process might have failed so try via popup
      tokenResp = await msalApp.acquireTokenPopup(accessTokenRequest)
      console.log('### MSAL acquireTokenPopup was successful')
    } finally {
        this.waitingOnAccessToken = false
    }

    // Just in case check, probably never triggers
    if (!tokenResp.accessToken) {
      throw new Error("### accessToken not found in response, that's bad")
    }

    this.accessToken = tokenResp.accessToken

    // Execute waiting call backs
    if (this.accessTokenCallbacks.length > 0) {
        this.accessTokenCallbacks.forEach( callback => {
            console.log('executing callback: '+callback)
            callback()
        })
        this.accessTokenCallbacks = []  // reset
    }

    return tokenResp.accessToken
  },

  //
  // Clear any stored/cached user
  //
  clearLocal() {
    if (msalApp) {
      for (let entry of Object.entries(localStorage)) {
        let key = entry[0]
        if (key.includes('login.windows')) {
          localStorage.removeItem(key)
        }
      }
    }
  },

  //
  // Check if we have been setup & configured
  //
  isConfigured() {
    return msalApp != null
  },

  /**
   * Initializes app authentication state. Attempt to acquire an access token if already logged in.
   */
  init() {

    this.configure();
    console.log("authAzure user: ")
    console.log(this.user());
    
    if (this.user()) {
      console.log('already signed in.');
      // Check if the current ID token is still valid based on expiration date
      if (this.checkIdToken()) {
        // set accessToken
        console.log('Acquiring accessing token ...')
        this.acquireToken().then (accessToken => {
            this.accessToken = accessToken      
            console.log('accessToken: '+ this.accessToken);      
            store.dispatch('auth/loginSuccess');                          
        });
      }
    } else {
      console.log('not signed in');
    }      
  },

  /**
   * Check if the cached ID token is still valid. If not, then clear the old token. 
   * User will be asked to reauthenticate.
   * @returns 
   */
  checkIdToken() {
    console.log('check Id token expiration')
    let user = this.user();
    if (user) {
        const expirationDateSecs = user.idTokenClaims.exp
        const expDate = new Date(expirationDateSecs * 1000)
        console.log('expDateSecs: '+expDate);

        if ((new Date()).getTime() >= expirationDateSecs * 1000) {
            console.log('IdToken expired. Clearing internal cache')
            this.clearLocal()
            return false            
        } else {
            console.log('ID token is still valid')
            return true
        }    
    } else {
        return false
    }
  },

  /**
   * Sign in and store the accessToken
   */
  appSignIn() {

    if (!msalApp) {
        return null
    }

    console.log('appSignIn')
    this.login().then( () => {
        if ( this.user()) {
            console.log('user signed in');
            // Automaticaly assign accessToken
            this.acquireToken().then (accessToken => {
                this.accessToken = accessToken      
                console.log('accessToken: '+ this.accessToken);      
                store.dispatch('auth/loginSuccess');             
              });                    
        } else {
            console.error('Failed to sign in');
            store.dispatch('auth/loginFailure');                        
        }     
    });

}  

}
