// /*  firebase-auth.js
//     — single‑session support (one device at a time)           */

// import {
//   getAuth,
//   signInWithEmailAndPassword,
//   signOut
// } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-auth.js";
// import {
//   getFirestore,
//   doc,
//   getDoc,
//   setDoc,
//   updateDoc,
//   deleteDoc
// } from "https://www.gstatic.com/firebasejs/9.23.0/firebase-firestore.js";

// let auth, db;

// /*  called once from login.js after default app is initialised */
// export function initAuthAndDb(app) {
//   auth = getAuth(app);
//   db   = getFirestore(app);
// }

// /* ────────────────────────────────────────────────
//    LOGIN  (returns true  / false)                  */
// export async function loginUser(email, password) {
//   const status = document.getElementById("status") || { textContent: "" };

//   try {
//     const cred = await signInWithEmailAndPassword(auth, email, password);
//     const uid  = cred.user.uid;

//     const sessRef = doc(db, "sessions", uid);
//     const snap    = await getDoc(sessRef);
//     const localId = localStorage.getItem("sessionId") || "";

//     /* block sign‑in from second device */
//     if (snap.exists() && snap.data().sessionId !== localId) {
//       status.textContent = "❌ Already signed‑in on another device.";
//       await signOut(auth);
//       return false;
//     }

//     /* create/reuse session */
//     const sessionId = localId || crypto.randomUUID();
//     await setDoc(sessRef, { sessionId, timestamp: Date.now() });

//     /* persist locally */
//     localStorage.setItem("uid",        uid);
//     localStorage.setItem("email",      email);
//     localStorage.setItem("sessionId",  sessionId);

//     status.textContent = "✅ Login successful";
//     return true;

//   } catch (err) {
//     console.error("Login error:", err);
//     status.textContent = "❌ " + (err.code || "Login failed.");
//     return false;
//   }
// }

// /*  session check used by taskpane.js */
// export async function isSessionValid() {
//   const uid = localStorage.getItem("uid");
//   const id  = localStorage.getItem("sessionId");
//   if (!uid || !id || !db) return false;

//   try {
//     const s = await getDoc(doc(db, "sessions", uid));
//     return s.exists() && s.data().sessionId === id;
//   } catch {
//     return false;
//   }
// }

// /* ────────────────────────────────────────────────
//    LOCAL cleanup after user requests logout
//    (taskpane.js sends the email, then calls this)  */
// export async function logoutRequestLocal() {
//   const uid = localStorage.getItem("uid");
//   if (!uid) return;

//   try {
//     await deleteDoc(doc(db, "sessions", uid));
//     localStorage.clear();
//     await signOut(auth);
//   } catch (err) {
//     console.error("Error in logoutRequestLocal:", err);
//   }
// }


/////////////////////////

/*  firebase-auth.js
    — single‑session support (one device at a time)           */

// src/firebase-auth.js

import { getAuth, signInWithEmailAndPassword, signOut }
  from "https://www.gstatic.com/firebasejs/11.10.0/firebase-auth.js";

import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  deleteDoc
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js";

let auth, db; // These will be initialized once the default Firebase app is ready

/* Called once from login.js (or wherever your main app init happens) after default app is initialised */
export function initAuthAndDb(app) {
  auth = getAuth(app);
  db = getFirestore(app);
}

/* ────────────────────────────────────────────────
   LOGIN (returns true / false)
   ──────────────────────────────────────────────── */
export async function loginUser(email, password) {
  const status = document.getElementById("status") || {
    textContent: ""
  }; // Fallback for status div if not present
  if (!auth || !db) {
    console.error("Firebase Auth or Firestore not initialized in firebase-auth.js during login attempt.");
    status.textContent = "❌ Internal error: Firebase not ready.";
    return false;
  }

  try {
    const cred = await signInWithEmailAndPassword(auth, email, password);
    const uid = cred.user.uid;

    const sessRef = doc(db, "sessions", uid); // Use uid as the document ID for sessions
    const snap = await getDoc(sessRef);
    const localId = localStorage.getItem("sessionId") || "";

    /* block sign‑in from second device */
    if (snap.exists() && snap.data().sessionId !== localId) {
      status.textContent = "❌ Already signed‑in on another device.";
      await signOut(auth); // Sign out the new attempted login
      return false;
    }

    /* create/reuse session */
    const sessionId = localId || crypto.randomUUID();
    await setDoc(sessRef, {
      sessionId,
      timestamp: Date.now(),
      email: email
    });

    /* persist locally */
    localStorage.setItem("uid", uid);
    localStorage.setItem("email", email);
    localStorage.setItem("sessionId", sessionId);

    status.textContent = "✅ Login successful";
    return true;

  } catch (err) {
    console.error("Login error:", err);
    status.textContent = "❌ " + (err.code || "Login failed.");
    return false;
  }
}

/* session check used by taskpane.js */
export async function isSessionValid() {
  const uid = localStorage.getItem("uid");
  const id = localStorage.getItem("sessionId");

  if (!uid || !id || !db || !auth) {
    return false;
  }

  try {
    const s = await getDoc(doc(db, "sessions", uid));
    return s.exists() && s.data().sessionId === id;
  } catch (err) {
    console.error("Error checking session validity:", err);
    return false;
  }
}

export async function logoutRequestLocal() {
  const uid = localStorage.getItem("uid");

  try {
    if (uid && db) {
      const sessionRef = doc(db, "sessions", uid);
      await deleteDoc(sessionRef); // Delete the session document from Firestore
    }

    localStorage.clear(); // Clear all local storage items
    if (auth && auth.currentUser) {
      await signOut(auth); // Sign out from Firebase authentication
    }
  } catch (err) {
    console.error("Error during logout cleanup:", err);
    // Attempt to clear local storage and sign out even if Firestore deletion fails
    localStorage.clear();
    if (auth && auth.currentUser) {
      await signOut(auth);
    }
  }
}
