import { initializeApp, deleteApp } from "https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc
} from "https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js";

async function start() {
  try {
    // 1. Init temporary app
    const tempApp = initializeApp({ projectId: "data-transfer-44eb5" }, "temp");
    const db = getFirestore(tempApp);

    // 2. Get full Firebase config
    const configSnap = await getDoc(doc(db, "config", "firebase"));
    if (!configSnap.exists()) throw new Error("❌ Firebase config not found.");
    const firebaseConfig = configSnap.data();

    // 3. Delete temp app
    await deleteApp(tempApp);

    // 4. Initialize real app
    const realApp = initializeApp(firebaseConfig);
    const realDb = getFirestore(realApp);

    // 5. Get login page URL
    const urlSnap = await getDoc(doc(realDb, "config", "urls"));
    if (!urlSnap.exists()) throw new Error("❌ URLs config not found.");
    const loginUrl = urlSnap.data().login;
    if (!loginUrl) throw new Error("❌ Login URL not found in Firestore.");

    // 6. Redirect to login page
    window.location.href = loginUrl;

  } catch (err) {
    document.body.innerHTML = `<h3 style="color:red">Startup failed:</h3><pre>${err.message}</pre>`;
    console.error("Startup error:", err);
  }
}

start();
