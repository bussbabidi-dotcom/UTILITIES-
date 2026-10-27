// Cole aqui as chaves do app Web do Firebase.
// Console: adicione um app Web, ative Authentication (e-mail e senha),
// crie Firestore e Storage, e publique as regras da pasta gadibruing.
export const firebaseConfig = {
  apiKey: "",
  authDomain: "",
  projectId: "",
  storageBucket: "",
  messagingSenderId: "",
  appId: ""
};

export const firebasePronto = Boolean(
  firebaseConfig.apiKey &&
  firebaseConfig.authDomain &&
  firebaseConfig.projectId &&
  firebaseConfig.storageBucket &&
  firebaseConfig.appId
);
