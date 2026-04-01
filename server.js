
const express = require("express");
const session = require("express-session");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
const SESSION_SECRET = process.env.SESSION_SECRET || "troque-essa-chave";

const DATA_DIR = path.join(__dirname, "data");
const USERS_PATH = path.join(DATA_DIR, "users.json");
const ACCOUNTS_PATH = path.join(DATA_DIR, "accounts.json");
const ORDERS_PATH = path.join(DATA_DIR, "orders.json");
const AGES_PATH = path.join(DATA_DIR, "ages.json");

function readJson(file, fallback) {
  try {
    if (!fs.existsSync(file)) return fallback;
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch {
    return fallback;
  }
}
function writeJson(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2), "utf8");
}
function seedFiles() {
  if (!fs.existsSync(USERS_PATH)) {
    writeJson(USERS_PATH, [
      { user: "admin", senha: "123456", status: "ativo", nome: "Administrador", role: "admin" },
      { user: "colab1", senha: "123456", status: "ativo", nome: "Colaborador 1", role: "colaborador" }
    ]);
  }
  if (!fs.existsSync(ACCOUNTS_PATH)) writeJson(ACCOUNTS_PATH, []);
  if (!fs.existsSync(ORDERS_PATH)) writeJson(ORDERS_PATH, []);
  if (!fs.existsSync(AGES_PATH)) writeJson(AGES_PATH, {});
}
seedFiles();

app.use(express.json({ limit: "1mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: "lax",
    secure: false,
    maxAge: 1000 * 60 * 60 * 12
  }
}));
app.use(express.static(path.join(__dirname, "public")));

function getUsers() { return readJson(USERS_PATH, []); }
function saveUsers(users) { writeJson(USERS_PATH, users); }
function getAccounts() { return readJson(ACCOUNTS_PATH, []); }
function saveAccounts(accounts) { writeJson(ACCOUNTS_PATH, accounts); }
function getOrders() { return readJson(ORDERS_PATH, []); }
function saveOrders(orders) { writeJson(ORDERS_PATH, orders); }
function getAges() { return readJson(AGES_PATH, {}); }
function saveAges(ages) { writeJson(AGES_PATH, ages); }

function currentUser(req) {
  const s = req.session?.user;
  if (!s) return null;
  const u = getUsers().find(x => x.user === s.user);
  if (!u) return null;
  return { user: u.user, nome: u.nome, status: u.status, role: u.role };
}
function requireLogin(req, res, next) {
  const u = currentUser(req);
  if (!u) return res.status(401).json({ error: "Sem sessão ativa." });
  if (u.status !== "ativo") {
    req.session.destroy(() => {});
    return res.status(403).json({ error: "Usuário bloqueado." });
  }
  req.currentUser = u;
  next();
}
function requireAdmin(req, res, next) {
  requireLogin(req, res, () => {
    if (req.currentUser.role !== "admin") return res.status(403).json({ error: "Somente admin pode fazer isso." });
    next();
  });
}
function ageKey(alias, cpf, cliente) {
  return `${alias}|${cpf || ""}|${cliente || ""}`;
}
function getAccount(alias) {
  return getAccounts().find(a => a.alias === alias);
}
async function refreshTokenIfNeeded(alias) {
  const accounts = getAccounts();
  const idx = accounts.findIndex(a => a.alias === alias);
  if (idx < 0) throw new Error("Conta não encontrada.");
  const a = accounts[idx];
  if (!a.refresh_token) return a;
  let doRefresh = true;
  if (a.expires_at) {
    const exp = new Date(a.expires_at).getTime();
    if (exp > Date.now() + 120000) doRefresh = false;
  }
  if (!doRefresh) return a;

  const auth = Buffer.from(`${a.clientId}:${a.clientSecret}`).toString("base64");
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: a.refresh_token
  });

  const resp = await fetch("https://www.bling.com.br/Api/v3/oauth/token", {
    method: "POST",
    headers: {
      "Authorization": `Basic ${auth}`,
      "enable-jwt": "1",
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body
  });
  const data = await resp.json();
  if (!resp.ok) throw new Error(data?.error?.description || data?.error || "Falha ao atualizar token.");
  a.access_token = data.access_token;
  a.refresh_token = data.refresh_token;
  a.expires_at = new Date(Date.now() + Number(data.expires_in || 0) * 1000).toISOString();
  accounts[idx] = a;
  saveAccounts(accounts);
  return a;
}
async function blingHeaders(alias) {
  const a = await refreshTokenIfNeeded(alias);
  if (!a.access_token) throw new Error(`Conta sem token: ${alias}`);
  return {
    "Authorization": `Bearer ${a.access_token}`,
    "enable-jwt": "1",
    "Accept": "application/json"
  };
}
function normalizeOrder(alias, item) {
  const ages = getAges();
  const pedidoId = String(item?.id ?? "");
  const pedidoNumero = String(item?.numero ?? item?.id ?? "");
  const cliente = item?.contato?.nome || item?.cliente?.nome || "";
  const cpf = item?.contato?.numeroDocumento || item?.contato?.cpf || item?.cliente?.numeroDocumento || "";
  const data = item?.data || item?.dataCriacao || "";
  const valor = item?.total || item?.valor || "";
  const nota = item?.notaFiscal?.numero || item?.notaFiscal?.id || "";
  const idade = ages[ageKey(alias, cpf, cliente)] ?? "";
  return { alias, pedidoId, pedidoNumero, data, cliente, cpf, idade, valor, nota };
}
function joinAddress(obj) {
  if (!obj) return "";
  const parts = [];
  const first = (...names) => names.find(n => obj[n]);
  const e = first("endereco", "logradouro", "rua");
  const n = first("numero");
  const c = first("complemento");
  const b = first("bairro");
  const city = first("municipio", "cidade");
  const uf = first("uf", "estado");
  const cep = first("cep");
  if (e) parts.push(obj[e]);
  if (n) parts.push(obj[n]);
  if (c) parts.push(obj[c]);
  if (b) parts.push(obj[b]);
  if (city || uf) parts.push([city ? obj[city] : "", uf ? obj[uf] : ""].filter(Boolean).join("/"));
  if (cep) parts.push(`CEP ${obj[cep]}`);
  return parts.join(", ");
}

app.get("/", (_req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));

app.post("/api/login-access", (req, res) => {
  const { user, senha } = req.body || {};
  const u = getUsers().find(x => x.user === user);
  if (!u) return res.status(404).json({ error: "Usuário não encontrado." });
  if (String(u.senha) !== String(senha)) return res.status(401).json({ error: "Senha inválida." });
  if (u.status !== "ativo") return res.status(403).json({ error: "Usuário bloqueado." });
  req.session.user = { user: u.user };
  return res.json({ ok: true, session: { user: u.user, nome: u.nome, status: u.status, role: u.role } });
});
app.get("/api/session-status-access", requireLogin, (req, res) => {
  res.json({ ok: true, session: req.currentUser });
});
app.post("/api/logout-access", (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

app.get("/api/users-access", requireLogin, (_req, res) => {
  const users = getUsers().map(u => ({ user: u.user, nome: u.nome, status: u.status, role: u.role }));
  res.json({ ok: true, users });
});
app.post("/api/users-access", requireAdmin, (req, res) => {
  const { user, senha, nome, role } = req.body || {};
  if (!user) return res.status(400).json({ error: "Usuário é obrigatório." });
  if (!senha) return res.status(400).json({ error: "Senha é obrigatória." });
  const users = getUsers();
  if (users.some(u => u.user === user)) return res.status(400).json({ error: "Usuário já existe." });
  users.push({ user, senha, nome: nome || user, status: "ativo", role: role || "colaborador" });
  saveUsers(users);
  res.json({ ok: true });
});
app.post("/api/users-access/status", requireAdmin, (req, res) => {
  const { user, status } = req.body || {};
  const users = getUsers();
  const u = users.find(x => x.user === user);
  if (!u) return res.status(404).json({ error: "Usuário não encontrado." });
  u.status = status;
  saveUsers(users);
  res.json({ ok: true });
});
app.post("/api/users-access/password", requireAdmin, (req, res) => {
  const { user, senha } = req.body || {};
  const users = getUsers();
  const u = users.find(x => x.user === user);
  if (!u) return res.status(404).json({ error: "Usuário não encontrado." });
  u.senha = String(senha || "");
  saveUsers(users);
  res.json({ ok: true });
});
app.post("/api/users-access/delete", requireAdmin, (req, res) => {
  const { user } = req.body || {};
  if (user === "admin") return res.status(400).json({ error: "Não é permitido apagar o admin." });
  saveUsers(getUsers().filter(u => u.user !== user));
  res.json({ ok: true });
});

app.get("/api/accounts", requireLogin, (_req, res) => {
  res.json({ accounts: getAccounts().map(a => ({ alias: a.alias, clientId: a.clientId, connected: !!a.access_token })) });
});
app.post("/api/account", requireAdmin, (req, res) => {
  const { alias, clientId, clientSecret } = req.body || {};
  if (!alias) return res.status(400).json({ error: "Alias da loja é obrigatório." });
  const accounts = getAccounts();
  const idx = accounts.findIndex(a => a.alias === alias);
  if (idx >= 0) {
    if (clientId) accounts[idx].clientId = clientId;
    if (clientSecret) accounts[idx].clientSecret = clientSecret;
  } else {
    accounts.push({ alias, clientId, clientSecret, access_token: "", refresh_token: "", expires_at: "" });
  }
  saveAccounts(accounts);
  res.json({ ok: true });
});
app.get("/api/login-bling", requireAdmin, (req, res) => {
  const alias = req.query.alias;
  const a = getAccount(alias);
  if (!a) return res.status(404).json({ error: "Conta não encontrada." });
  const state = Math.random().toString(36).slice(2) + Date.now().toString(36);
  req.session.pendingBling = { alias, state };
  const redirectUri = `${BASE_URL}/callback`;
  const url = `https://www.bling.com.br/Api/v3/oauth/authorize?response_type=code&client_id=${encodeURIComponent(a.clientId)}&state=${encodeURIComponent(state)}&redirect_uri=${encodeURIComponent(redirectUri)}`;
  res.redirect(url);
});
app.get("/callback", async (req, res) => {
  try {
    const code = req.query.code;
    const state = req.query.state;
    const pending = req.session.pendingBling;
    if (!code) throw new Error("Bling não retornou authorization code.");
    if (!pending || pending.state !== state) throw new Error("State inválido.");
    const accounts = getAccounts();
    const idx = accounts.findIndex(a => a.alias === pending.alias);
    if (idx < 0) throw new Error("Conta pendente não encontrada.");
    const a = accounts[idx];
    const auth = Buffer.from(`${a.clientId}:${a.clientSecret}`).toString("base64");
    const body = new URLSearchParams({ grant_type: "authorization_code", code: String(code) });
    const resp = await fetch("https://www.bling.com.br/Api/v3/oauth/token", {
      method: "POST",
      headers: {
        "Authorization": `Basic ${auth}`,
        "enable-jwt": "1",
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.description || data?.error || "Falha ao conectar conta.");
    a.access_token = data.access_token;
    a.refresh_token = data.refresh_token;
    a.expires_at = new Date(Date.now() + Number(data.expires_in || 0) * 1000).toISOString();
    accounts[idx] = a;
    saveAccounts(accounts);
    delete req.session.pendingBling;
    res.send(`<html><body style="font-family:Segoe UI;background:#0b1220;color:#fff;padding:30px"><h2>Conta conectada com sucesso.</h2><p>Volte para o app e clique em Sincronizar todas.</p></body></html>`);
  } catch (e) {
    res.status(500).send(`<html><body style="font-family:Segoe UI;background:#0b1220;color:#fff;padding:30px"><h2>Erro ao conectar conta.</h2><p>${String(e.message)}</p></body></html>`);
  }
});

app.post("/api/sync-all", requireAdmin, async (_req, res) => {
  try {
    const all = [];
    for (const acc of getAccounts()) {
      if (!acc.access_token) continue;
      const headers = await blingHeaders(acc.alias);
      for (let pagina = 1; pagina <= 3; pagina++) {
        const url = `https://api.bling.com.br/Api/v3/pedidos/vendas?pagina=${pagina}`;
        const resp = await fetch(url, { headers });
        const data = await resp.json();
        if (!resp.ok) throw new Error(data?.error?.description || data?.error || `Falha na conta ${acc.alias}`);
        const items = data.data || data.itens || [];
        for (const item of items) all.push(normalizeOrder(acc.alias, item));
        if (!items.length || items.length < 100) break;
      }
    }
    saveOrders(all);
    res.json({ ok: true, count: all.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.get("/api/orders", requireLogin, (_req, res) => {
  res.json({ orders: getOrders() });
});
app.get("/api/order/:alias/:pedidoId", requireLogin, async (req, res) => {
  try {
    const { alias, pedidoId } = req.params;
    const headers = await blingHeaders(alias);
    const resp = await fetch(`https://api.bling.com.br/Api/v3/pedidos/vendas/${encodeURIComponent(pedidoId)}`, { headers });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.description || data?.error || "Falha ao buscar pedido.");
    const raw = data.data || data;
    const detail = {
      alias,
      pedido: pedidoId,
      cliente: raw?.contato?.nome || "",
      cpf: raw?.contato?.numeroDocumento || raw?.contato?.cpf || "",
      data: raw?.data || "",
      valor: raw?.total || raw?.valor || "",
      nota: raw?.notaFiscal?.numero || raw?.notaFiscal?.id || "",
      endereco: raw?.transporte?.etiqueta ? joinAddress(raw.transporte.etiqueta) : "",
      itemDescricao: raw?.itens?.[0]?.descricao || raw?.itens?.[0]?.descricaoDetalhada || raw?.itens?.[0]?.codigo || "",
      subtotal: raw?.total || raw?.valor || "",
      total: raw?.total || raw?.valor || "",
      rawHint: raw?.itens?.length ? "" : "Item não veio no formato esperado."
    };
    res.json({ detail });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post("/api/age", requireLogin, (req, res) => {
  const { alias, cpf, cliente, idade } = req.body || {};
  const ages = getAges();
  ages[ageKey(alias, cpf, cliente)] = idade;
  saveAges(ages);
  const orders = getOrders().map(o => {
    if (o.alias === alias && ((o.cpf && o.cpf === cpf) || (!o.cpf && o.cliente === cliente))) o.idade = idade;
    return o;
  });
  saveOrders(orders);
  res.json({ ok: true });
});
app.post("/api/clear-bling", requireAdmin, (_req, res) => {
  saveAccounts([]);
  saveOrders([]);
  saveAges({});
  res.json({ ok: true });
});

app.listen(PORT, () => console.log(`Trevizio online rodando em ${BASE_URL}`));
