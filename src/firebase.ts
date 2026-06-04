import { initializeApp } from "firebase/app";
import { getAuth } from "firebase/auth";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from "firebase/firestore";
import firebaseConfig from "../firebase-applet-config.json";

// Check if environment variables are provided as override
const useEnvConfig = import.meta.env.VITE_FIREBASE_PROJECT_ID ? true : false;

const config = useEnvConfig ? {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
} : firebaseConfig;

const app = initializeApp(config);

export const auth = getAuth(app);

// Active database configuration
const firestoreDatabaseId = useEnvConfig ? undefined : (firebaseConfig as any).firestoreDatabaseId;

// Configured with persistent local cache to store documents inside the browser's IndexedDB.
// Multiple tab manager allows sharing the cache across tabs securely and safely without duplicating read quotas.
export const db = initializeFirestore(app, {
  experimentalForceLongPolling: true,
  localCache: persistentLocalCache({
    tabManager: persistentMultipleTabManager()
  })
}, firestoreDatabaseId);

// Export standard default database reference for cross-database token writing
export const defaultDb = firestoreDatabaseId && firestoreDatabaseId !== "(default)"
  ? initializeFirestore(app, { experimentalForceLongPolling: true }, "(default)")
  : db;

export default app;
