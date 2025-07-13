/*  login.js  —  loads config, initialises Firebase,
                 signs user in, redirects to taskpane */

import {
  initializeApp,
  deleteApp,
  getApps,
  getApp
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js";

import { loginUser, initAuthAndDb } from "../firebase-auth.js";   //  ← path fixed

/* ── get full Firebase config from Firestore (temp unnamed app) */
async function fetchConfig() {
  const tmp = initializeApp({ projectId: "data-transfer-44eb5" }, "tmp-cfg");
  const cfgSnap = await getDoc(doc(getFirestore(tmp), "config", "firebase"));
  if (!cfgSnap.exists()) throw new Error("Firebase config not found in Firestore");
  const cfg = cfgSnap.data();
  await deleteApp(tmp);
  return cfg;
}

/* ── main login flow */
async function doLogin(email, password) {
  const status = document.getElementById("status");

  status.textContent = "🔄 Loading…";
  const cfg = await fetchConfig();

  /* safe default‑app init */
  let app;
  if (getApps().length === 0) app = initializeApp(cfg);
  else app = getApp();

  initAuthAndDb(app);                    // wire auth + db

  status.textContent = "🔐 Signing in…";
  if (!(await loginUser(email, password))) return;

  /* read redirect from Firestore */
  const urlsSnap = await getDoc(doc(getFirestore(app), "config", "urls"));
  const taskpane = urlsSnap.data()?.taskpane;
  if (!taskpane) throw new Error("taskpane URL missing in Firestore");

  window.location.href = taskpane;       // 🚀
}

/* ── attach handler after DOM ready */
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("loginBtn").addEventListener("click", () => {
    const email = document.getElementById("emailInput").value.trim();
    const pass = document.getElementById("passwordInput").value.trim();
    const status = document.getElementById("status");

    if (!email || !pass) { status.textContent = "❌ Enter both fields."; return; }

    doLogin(email, pass).catch(err => {
      console.error(err);
      status.textContent = "❌ " + err.message;
    });
  });
});
