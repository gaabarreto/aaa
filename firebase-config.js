/* =========================================================================
   Configuración de Firebase — REEMPLAZA estos valores con los de TU proyecto.

   Cómo obtenerlos:
   1. Ve a https://console.firebase.google.com y crea un proyecto (gratis).
   2. Firestore Database (menú izquierdo) → Crear base de datos → modo producción.
   3. En "Reglas" pega el contenido de firestore.rules (incluido) y publica.
   4. Configuración del proyecto (ícono de engranaje) → Tus apps → </> (Web)
      → registra una app → copia el objeto "firebaseConfig" que te muestra
      y pégalo aquí abajo, reemplazando los valores "TU_...".

   Si dejas los valores por defecto, la app sigue funcionando 100% local
   (cada asesor solo ve sus propios datos en su dispositivo); Administrador
   y BI no verán las capturas de los demás hasta que completes esto.
   ========================================================================= */
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyB8kCQXxq8FJjRtaY9oEF0fgktlU4mqeaM",
  authDomain: "dataprice-sura.firebaseapp.com",
  projectId: "dataprice-sura",
  storageBucket: "dataprice-sura.firebasestorage.app",
  messagingSenderId: "608887586922",
  appId: "1:608887586922:web:6819354f3e7a1b19844a13",
  measurementId: "G-JMYT8EZY4W"
};
