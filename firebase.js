import { initializeApp } from 'firebase/app';
import { getDatabase } from 'firebase/database';
import { getAuth, GoogleAuthProvider } from 'firebase/auth';

const firebaseConfig = {
  apiKey:            "AIzaSyB3no4PBjmXmkSk9dLnQPiPyPJdEIOf2_4",
  authDomain:        "spl-dpr.firebaseapp.com",
  databaseURL:       "https://spl-dpr-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId:         "spl-dpr",
  storageBucket:     "spl-dpr.firebasestorage.app",
  messagingSenderId: "639178076230",
  appId:             "1:639178076230:web:1d009bf9979472340dda47",
  measurementId:     "G-MZ11XK0LY2"
};

const firebaseApp = initializeApp(firebaseConfig);
export const db = getDatabase(firebaseApp);

// ── Authentication (Google sign-in) ──────────────────────────────────────────
export const auth = getAuth(firebaseApp);
export const googleProvider = new GoogleAuthProvider();
googleProvider.setCustomParameters({ prompt: 'select_account' });
