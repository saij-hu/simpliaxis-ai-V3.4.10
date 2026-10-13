/*
 * Simpliaxis Firebase Web App configuration.
 *
 * Paste the CONFIG object from Firebase Console → Project settings → Your apps → Web app.
 * This is browser-side Firebase configuration, not a service-account private key.
 * Protect the database with Firestore Security Rules and Firebase Authentication.
 * Leave values blank to keep the existing LocalStorage/IndexedDB-only fallback mode.
 */
window.SIMPLIAXIS_FIREBASE_CONFIG = window.SIMPLIAXIS_FIREBASE_CONFIG || {
  apiKey: "",
  authDomain: "",
  projectId: "",
  storageBucket: "",
  messagingSenderId: "",
  appId: ""
};
