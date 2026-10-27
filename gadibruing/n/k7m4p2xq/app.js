import { firebaseConfig, firebasePronto } from "./firebase-config.js";
import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.0/firebase-app.js";
import {
  browserLocalPersistence,
  createUserWithEmailAndPassword,
  getAuth,
  onAuthStateChanged,
  setPersistence,
  signInWithEmailAndPassword,
  signOut
} from "https://www.gstatic.com/firebasejs/11.6.0/firebase-auth.js";
import {
  addDoc,
  collection,
  doc,
  getDoc,
  getFirestore,
  onSnapshot,
  orderBy,
  query,
  serverTimestamp,
  setDoc
} from "https://www.gstatic.com/firebasejs/11.6.0/firebase-firestore.js";
import {
  deleteObject,
  getDownloadURL,
  getStorage,
  ref,
  uploadBytes
} from "https://www.gstatic.com/firebasejs/11.6.0/firebase-storage.js";

const $ = (id) => document.getElementById(id);

const state = {
  user: null,
  apelido: "",
  saved: paginaInicial(""),
  draft: null,
  editing: false,
  salvando: false,
  apelidoNovo: "",
  novas: new Set(),
  recorte: null,
  arraste: null,
  mensagens: [],
  unsubPage: null,
  unsubChat: null
};

let auth;
let db;
let storage;

function paginaInicial(apelido) {
  return {
    apelido,
    nome: "gadibruing",
    frase: "aqui é só nosso.",
    fotos: []
  };
}

function clamp(valor, min, max) {
  return Math.min(max, Math.max(min, valor));
}

function emailDoApelido(apelido) {
  const base = apelido
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ".")
    .replace(/[^a-z0-9.]/g, "")
    .replace(/^\.+|\.+$/g, "");
  if (base.length < 2) {
    throw new Error("usa um apelido com letras ou números");
  }
  return `${base}@gadibruing.app`;
}

function normalizar(data) {
  const fotos = Array.isArray(data.fotos) ? data.fotos.slice(0, 6) : [];
  return {
    apelido: String(data.apelido || ""),
    nome: String(data.nome || "gadibruing").slice(0, 40),
    frase: String(data.frase || "").slice(0, 180),
    fotos: fotos.map((foto) => ({
      id: String(foto.id || crypto.randomUUID()),
      url: String(foto.url || ""),
      path: String(foto.path || ""),
      x: clamp(Number(foto.x ?? 50), 0, 100),
      y: clamp(Number(foto.y ?? 50), 0, 100),
      zoom: clamp(Number(foto.zoom ?? 1), 1, 2.5)
    })).filter((foto) => foto.url)
  };
}

function aplicarRecorte(img, foto) {
  img.style.setProperty("--x", String(foto.x ?? 50));
  img.style.setProperty("--y", String(foto.y ?? 50));
  img.style.setProperty("--z", String(foto.zoom ?? 1));
}

function renderPagina(pagina) {
  $("nome").textContent = pagina.nome || "gadibruing";
  $("frase").textContent = pagina.frase || "";
  document.title = pagina.nome || "gadibruing";
  const box = $("fotos");
  box.replaceChildren();
  for (let i = 0; i < 6; i += 1) {
    const figure = document.createElement("figure");
    const foto = pagina.fotos[i];
    if (foto?.url) {
      const img = document.createElement("img");
      img.alt = "";
      aplicarRecorte(img, foto);
      img.src = foto.url;
      figure.append(img);
    }
    box.append(figure);
  }
}

function mostrarErro(id, texto) {
  $(id).textContent = texto || "";
}

function mostrarTela(nome) {
  $("pagina").hidden = nome !== "pagina";
  $("barra").hidden = nome !== "pagina";
  $("auth").hidden = nome !== "auth";
  $("editor").hidden = nome !== "editor";
  $("enquadrar").hidden = nome !== "enquadrar";
  $("conversa").hidden = nome !== "conversa";
  $("carregando").hidden = nome !== "carregando";
  if (nome === "conversa") {
    const box = $("mensagens");
    box.scrollTop = box.scrollHeight;
  }
}

function renderEditor() {
  const pagina = state.draft;
  $("edit-nome").value = pagina.nome;
  $("edit-frase").value = pagina.frase;
  const box = $("edit-fotos");
  box.replaceChildren();
  pagina.fotos.forEach((foto, index) => {
    const linha = document.createElement("article");
    linha.className = "foto-edit";
    const mini = document.createElement("div");
    mini.className = "mini";
    const img = document.createElement("img");
    img.alt = "";
    aplicarRecorte(img, foto);
    img.src = foto.url;
    mini.append(img);
    const botoes = document.createElement("div");
    botoes.className = "botoes";
    botoes.append(
      botaoFoto("subir", index, index === 0),
      botaoFoto("descer", index, index === pagina.fotos.length - 1),
      botaoFoto("enquadrar", index, false),
      botaoFoto("apagar", index, false)
    );
    linha.append(mini, botoes);
    box.append(linha);
  });
  $("add-foto").disabled = pagina.fotos.length >= 6;
}

function botaoFoto(acao, index, desligado) {
  const botao = document.createElement("button");
  botao.type = "button";
  botao.dataset.acao = acao;
  botao.dataset.i = String(index);
  botao.textContent = acao;
  botao.disabled = desligado;
  return botao;
}

function pararOuvintes() {
  state.unsubPage?.();
  state.unsubChat?.();
  state.unsubPage = null;
  state.unsubChat = null;
}

function mensagemFirebase(erro) {
  const codigo = erro?.code || "";
  if (codigo === "auth/email-already-in-use") return "esse apelido já existe";
  if (codigo === "auth/invalid-credential" || codigo === "auth/wrong-password" || codigo === "auth/user-not-found" || codigo === "auth/invalid-email") {
    return "apelido ou senha não confere";
  }
  if (codigo === "auth/weak-password") return "a senha precisa ter pelo menos 6 caracteres";
  if (codigo === "auth/operation-not-allowed") return "ativa o login por e-mail e senha no Firebase";
  if (codigo === "auth/too-many-requests") return "muitas tentativas. espera um pouco";
  return "não deu. tenta de novo";
}

async function entrarOuCriar(criar) {
  if (!auth) return;
  mostrarErro("erro-auth", "");
  const apelido = $("apelido").value.trim();
  const senha = $("senha").value;
  if (apelido.length < 2 || apelido.length > 24) {
    mostrarErro("erro-auth", "o apelido precisa ter de 2 a 24 caracteres");
    return;
  }
  if (senha.length < 6) {
    mostrarErro("erro-auth", "a senha precisa ter pelo menos 6 caracteres");
    return;
  }
  let email;
  try {
    email = emailDoApelido(apelido);
  } catch (erro) {
    mostrarErro("erro-auth", erro.message);
    return;
  }
  try {
    if (criar) state.apelidoNovo = apelido;
    if (criar) await createUserWithEmailAndPassword(auth, email, senha);
    else await signInWithEmailAndPassword(auth, email, senha);
  } catch (erro) {
    state.apelidoNovo = "";
    mostrarErro("erro-auth", mensagemFirebase(erro));
  }
}

async function aoEntrar(user) {
  pararOuvintes();
  state.user = user;
  if (!user) {
    state.apelido = "";
    state.saved = paginaInicial("");
    mostrarTela("auth");
    return;
  }
  const refUsuario = doc(db, "usuarias", user.uid);
  const snap = await getDoc(refUsuario);
  if (!snap.exists()) {
    await setDoc(refUsuario, paginaInicial(state.apelidoNovo || "amiga"));
  }
  state.apelidoNovo = "";
  state.unsubPage = onSnapshot(refUsuario, (atual) => {
    if (!atual.exists()) return;
    const pagina = normalizar(atual.data());
    state.saved = pagina;
    state.apelido = pagina.apelido;
    if (!state.editing && !state.salvando) renderPagina(pagina);
  }, () => {
    $("aviso").hidden = false;
    $("aviso").textContent = "não consegui abrir a sua página";
  });
  const conversa = query(collection(db, "conversa"), orderBy("createdAt", "asc"));
  state.unsubChat = onSnapshot(conversa, (lista) => {
    state.mensagens = lista.docs.map((item) => ({ id: item.id, ...item.data() }));
    renderMensagens();
  }, () => {
    mostrarErro("erro-conversa", "não consegui abrir a conversa");
  });
  mostrarTela("pagina");
}

function renderMensagens() {
  const box = $("mensagens");
  const perto = box.scrollHeight - box.scrollTop - box.clientHeight < 90;
  box.replaceChildren();
  for (const mensagem of state.mensagens) {
    const item = document.createElement("article");
    const minha = mensagem.uid === state.user?.uid;
    item.className = minha ? "msg minha" : "msg";
    if (!minha) {
      const quem = document.createElement("span");
      quem.className = "quem";
      quem.textContent = mensagem.apelido || "";
      item.append(quem);
    }
    item.append(document.createTextNode(mensagem.texto || ""));
    const hora = document.createElement("time");
    hora.textContent = mensagem.createdAt?.toDate
      ? mensagem.createdAt.toDate().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" })
      : "";
    item.append(hora);
    box.append(item);
  }
  if (perto || $("conversa").hidden) box.scrollTop = box.scrollHeight;
}

function abrirEditor() {
  state.editing = true;
  state.draft = structuredClone(state.saved);
  state.novas.clear();
  mostrarErro("erro-editor", "");
  renderEditor();
  mostrarTela("editor");
}

async function fecharEditor() {
  const apagar = [...state.novas];
  state.novas.clear();
  state.editing = false;
  state.draft = null;
  state.recorte = null;
  renderPagina(state.saved);
  mostrarTela("pagina");
  await Promise.all(apagar.map((path) => deleteObject(ref(storage, path)).catch(() => {})));
}

async function concluir() {
  if (!state.draft || state.salvando) return;
  const nome = $("edit-nome").value.trim();
  const frase = $("edit-frase").value.trim();
  if (!nome) {
    mostrarErro("erro-editor", "escreve um nome");
    return;
  }
  state.draft.nome = nome;
  state.draft.frase = frase;
  mostrarErro("erro-editor", "");
  state.salvando = true;
  $("concluir").disabled = true;
  try {
    await setDoc(doc(db, "usuarias", state.user.uid), {
      apelido: state.apelido,
      nome,
      frase,
      fotos: state.draft.fotos
    }, { merge: true });
    const atuais = new Set(state.draft.fotos.map((foto) => foto.path));
    const sobrou = [...state.saved.fotos.map((foto) => foto.path), ...state.novas]
      .filter((path) => path && !atuais.has(path));
    await Promise.all(sobrou.map((path) => deleteObject(ref(storage, path)).catch(() => {})));
    state.novas.clear();
    state.saved = { ...state.saved, nome, frase, fotos: state.draft.fotos };
    state.editing = false;
    state.draft = null;
    renderPagina(state.saved);
    mostrarTela("pagina");
  } catch (erro) {
    mostrarErro("erro-editor", "não consegui salvar. tenta de novo");
  } finally {
    state.salvando = false;
    $("concluir").disabled = false;
  }
}

function carregarImagem(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(img);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("foto"));
    };
    img.src = url;
  });
}

async function prepararFoto(file) {
  const img = await carregarImagem(file);
  const escala = Math.min(1, 1600 / Math.max(img.width, img.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(img.width * escala));
  canvas.height = Math.max(1, Math.round(img.height * escala));
  canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.82));
  if (!blob) throw new Error("foto");
  return blob;
}

async function adicionarFoto(file) {
  if (!state.draft || state.draft.fotos.length >= 6) return;
  mostrarErro("erro-editor", "");
  $("add-foto").disabled = true;
  try {
    const blob = await prepararFoto(file);
    const id = crypto.randomUUID();
    const path = `usuarias/${state.user.uid}/${id}.jpg`;
    const arquivo = ref(storage, path);
    await uploadBytes(arquivo, blob, { contentType: "image/jpeg" });
    const url = await getDownloadURL(arquivo);
    state.draft.fotos.push({ id, url, path, x: 50, y: 50, zoom: 1 });
    state.novas.add(path);
    renderEditor();
    renderPagina(state.draft);
  } catch (erro) {
    mostrarErro("erro-editor", "não consegui abrir essa foto");
  } finally {
    $("add-foto").disabled = state.draft.fotos.length >= 6;
  }
}

function abrirRecorte(index) {
  const foto = state.draft.fotos[index];
  state.recorte = {
    index,
    x: foto.x,
    y: foto.y,
    zoom: foto.zoom,
    original: { x: foto.x, y: foto.y, zoom: foto.zoom }
  };
  const img = $("recorte");
  aplicarRecorte(img, state.recorte);
  img.src = foto.url;
  $("zoom").value = String(foto.zoom);
  mostrarTela("enquadrar");
}

function atualizarRecorte() {
  aplicarRecorte($("recorte"), state.recorte);
}

function prontoRecorte() {
  const foto = state.draft.fotos[state.recorte.index];
  foto.x = state.recorte.x;
  foto.y = state.recorte.y;
  foto.zoom = state.recorte.zoom;
  state.recorte = null;
  renderEditor();
  renderPagina(state.draft);
  mostrarTela("editor");
}

function cancelarRecorte() {
  state.recorte = null;
  mostrarTela("editor");
}

async function enviarMensagem(event) {
  event.preventDefault();
  const campo = $("texto");
  const texto = campo.value.trim();
  if (!texto || !state.user) return;
  campo.value = "";
  mostrarErro("erro-conversa", "");
  try {
    await addDoc(collection(db, "conversa"), {
      uid: state.user.uid,
      apelido: state.apelido,
      texto: texto.slice(0, 500),
      createdAt: serverTimestamp()
    });
  } catch (erro) {
    campo.value = texto;
    mostrarErro("erro-conversa", "não consegui enviar");
  }
}

function ligarAcoes() {
  $("form-auth").addEventListener("submit", (event) => {
    event.preventDefault();
    entrarOuCriar(false);
  });
  $("criar").addEventListener("click", () => entrarOuCriar(true));
  $("abrir-editar").addEventListener("click", abrirEditor);
  $("abrir-conversa").addEventListener("click", () => {
    mostrarErro("erro-conversa", "");
    mostrarTela("conversa");
  });
  $("fechar-conversa").addEventListener("click", () => mostrarTela("pagina"));
  $("enviar").addEventListener("submit", enviarMensagem);
  $("fechar-editor").addEventListener("click", fecharEditor);
  $("concluir").addEventListener("click", concluir);
  $("sair").addEventListener("click", async () => {
    if (!auth) return;
    if (state.editing) await fecharEditor();
    await signOut(auth);
  });
  $("add-foto").addEventListener("click", () => $("arquivo").click());
  $("arquivo").addEventListener("change", () => {
    const file = $("arquivo").files?.[0];
    $("arquivo").value = "";
    if (file) adicionarFoto(file);
  });
  $("edit-nome").addEventListener("input", () => {
    if (!state.draft) return;
    state.draft.nome = $("edit-nome").value;
    renderPagina(state.draft);
  });
  $("edit-frase").addEventListener("input", () => {
    if (!state.draft) return;
    state.draft.frase = $("edit-frase").value;
    renderPagina(state.draft);
  });
  $("edit-fotos").addEventListener("click", (event) => {
    const botao = event.target.closest("button");
    if (!botao || !state.draft) return;
    const index = Number(botao.dataset.i);
    const fotos = state.draft.fotos;
    if (botao.dataset.acao === "subir" && index > 0) {
      [fotos[index - 1], fotos[index]] = [fotos[index], fotos[index - 1]];
    } else if (botao.dataset.acao === "descer" && index < fotos.length - 1) {
      [fotos[index + 1], fotos[index]] = [fotos[index], fotos[index + 1]];
    } else if (botao.dataset.acao === "apagar") {
      fotos.splice(index, 1);
    } else if (botao.dataset.acao === "enquadrar") {
      abrirRecorte(index);
      return;
    }
    renderEditor();
    renderPagina(state.draft);
  });
  $("moldura").addEventListener("pointerdown", (event) => {
    if (!state.recorte) return;
    $("moldura").setPointerCapture(event.pointerId);
    state.arraste = {
      id: event.pointerId,
      x: event.clientX,
      y: event.clientY,
      ox: state.recorte.x,
      oy: state.recorte.y
    };
  });
  $("moldura").addEventListener("pointermove", (event) => {
    if (!state.arraste || state.arraste.id !== event.pointerId) return;
    const rect = $("moldura").getBoundingClientRect();
    state.recorte.x = clamp(state.arraste.ox + ((event.clientX - state.arraste.x) / rect.width) * 100, 0, 100);
    state.recorte.y = clamp(state.arraste.oy + ((event.clientY - state.arraste.y) / rect.height) * 100, 0, 100);
    atualizarRecorte();
  });
  $("moldura").addEventListener("pointerup", () => {
    state.arraste = null;
  });
  $("zoom").addEventListener("input", () => {
    if (!state.recorte) return;
    state.recorte.zoom = Number($("zoom").value);
    atualizarRecorte();
  });
  $("recorte-pronto").addEventListener("click", prontoRecorte);
  $("recorte-cancelar").addEventListener("click", cancelarRecorte);
}

async function iniciar() {
  ligarAcoes();
  if (!firebasePronto) return;
  $("aviso").hidden = true;
  mostrarTela("carregando");
  const app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
  storage = getStorage(app);
  auth.languageCode = "pt";
  await setPersistence(auth, browserLocalPersistence);
  onAuthStateChanged(auth, (user) => {
    aoEntrar(user).catch(() => {
      mostrarErro("erro-auth", "não consegui abrir a sua página");
      mostrarTela("auth");
    });
  });
}

iniciar();
