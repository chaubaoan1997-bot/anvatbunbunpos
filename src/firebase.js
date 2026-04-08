import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
    apiKey: "AIzaSyBm9ShdWPPQVkPCyHvLDgURXwAo_25UJIc",
    authDomain: "amvatbunbun2.firebaseapp.com",
    projectId: "amvatbunbun2",
    storageBucket: "amvatbunbun2.firebasestorage.app",
    messagingSenderId: "474398447690",
    appId: "1:474398447690:web:8837400c12cecba9cd6eb3",
    measurementId: "G-NLJYSL0SWH"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);