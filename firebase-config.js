import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";
// 🔴 REPLACE THIS with your actual Firebase Config
  const firebaseConfig = {
    apiKey: "AIzaSyDvtVFsGd37aAWNoLSgAQWxhVa7-gNc1Y8",
    authDomain: "lqservicetracker.firebaseapp.com",
    databaseURL: "https://lqservicetracker-default-rtdb.europe-west1.firebasedatabase.app",
    projectId: "lqservicetracker",
    storageBucket: "lqservicetracker.firebasestorage.app",
    messagingSenderId: "716349093814",
    appId: "1:716349093814:web:f01b3f17da7c8d0a2f7317",
    measurementId: "G-DZNMPDJ91K"
  };

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Initialize Firestore with offline persistence enabled
export const db = initializeFirestore(app, {
  localCache: persistentLocalCache({tabManager: persistentMultipleTabManager()})
});
