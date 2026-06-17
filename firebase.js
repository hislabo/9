import { initializeApp } from 'https://www.gstatic.com/firebasejs/12.14.0/firebase-app.js';
    import { getAnalytics } from 'https://www.gstatic.com/firebasejs/12.14.0/firebase-analytics.js';
    import { getFirestore, doc, getDoc, setDoc } from 'https://www.gstatic.com/firebasejs/12.14.0/firebase-firestore.js';
    import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from 'https://www.gstatic.com/firebasejs/12.14.0/firebase-auth.js';

    const firebaseConfig = {
      apiKey: 'AIzaSyDuNzClGuAC3ct7B5Ide0gvfbIjxkLNg0E',
      authDomain: 'qlxn-7e126.firebaseapp.com',
      projectId: 'qlxn-7e126',
      storageBucket: 'qlxn-7e126.firebasestorage.app',
      messagingSenderId: '821322697656',
      appId: '1:821322697656:web:637f0f4d92417f5fa4dd4d',
      measurementId: 'G-XW42HCG5RD'
    };

    const app = initializeApp(firebaseConfig);
    getAnalytics(app);
    const db = getFirestore(app);
    const auth = getAuth(app);

    window.firebaseDb = db;
    window.firebaseAuth = auth;
    window.firebaseFirestoreImports = { doc, getDoc, setDoc };
    window.firebaseAuthImports = { signInWithEmailAndPassword, signOut, onAuthStateChanged };
