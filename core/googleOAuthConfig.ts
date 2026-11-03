/**
 * The app's Google OAuth "Desktop app" client, used by "Sign in with Google" (Settings → Online Works List).
 * A desktop client's id/secret are not confidential (Google documents them as embeddable in installed apps).
 * Create one at console.cloud.google.com → APIs & Services → Credentials, enable the Google Sheets API, and paste
 * both values here. Left blank, the Settings tile says Google sign-in isn't set up yet.
 */
export const GOOGLE_OAUTH_CLIENT_ID = '71347636667-3q9v1mb9ifbic0deh34cpv1rtk8v39hm.apps.googleusercontent.com'
export const GOOGLE_OAUTH_CLIENT_SECRET = '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'
