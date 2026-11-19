/* global Office, Trello, Mail */

const $ = (id) => document.getElementById(id);
const SETTINGS = { token: "trelloToken", board: "lastBoardId", list: "lastListId" };

const LABEL_COLORS = {
  green: "#4bce97", yellow: "#f5cd47", orange: "#fea362", red: "#f87168", purple: "#9f8fef",
  blue: "#579dff", sky: "#6cc3e0", lime: "#94c748", pink: "#e774bb", black: "#8590a2",
};

const state = {
  mail: null,
  labels: new Set(),
  members: new Set(),
  pickedCard: null,
};

// ---------- settings (roam with the user's mailbox) ----------
const settings = {
  get: (k) => Office.context.roamingSettings.get(k),
  set(values) {
    for (const [k, v] of Object.entries(values)) {
      if (v == null) Office.context.roamingSettings.remove(k);
      else Office.context.roamingSettings.set(k, v);
    }
    return new Promise((r) => Office.context.roamingSettings.saveAsync(() => r()));
  },
};

// ---------- views ----------
function show(view) {
  for (const v of document.querySelectorAll(".view")) v.hidden = v.id !== `view-${view}`;
}

function setStatus(msg, isError = false) {
  const el = $("status");
  el.hidden = !msg;
  el.textContent = msg || "";
  el.classList.toggle("error", isError);
}

function option(value, text) {
  const o = document.createElement("option");
  o.value = value;
  o.textContent = text;
  return o;
}

const stripPrefixes = (subject) => subject.replace(/^((re|fw|fwd)\s*:\s*)+/i, "");

function friendlyError(err) {
  if (err && err.status === 401) return "Your Trello session has expired. Sign out and connect again.";
  const msg = (err && (err.message || err.toString())) || "Something went wrong.";
  return msg.length > 200 ? msg.slice(0, 200) + "…" : msg;
}

// ---------- auth ----------
function signIn() {
  $("signin-error").hidden = true;
  const url = new URL("auth-start.html", location.href).href; // relative, so it works under a subpath
  Office.context.ui.displayDialogAsync(url, { height: 70, width: 40, promptBeforeOpen: false }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      return showSigninError(`Could not open the sign-in window (${res.error.message}).`);
    }
    const dialog = res.value;
    activeDialog = dialog;
    $("form-paste").hidden = false;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
      closeDialog();
      let msg;
      try { msg = JSON.parse(arg.message); } catch { msg = {}; }
      if (!msg.token) return showSigninError(msg.error || "Trello did not return a token.");
      await saveToken(msg.token);
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      if (arg.error !== 12006) showSigninError(`Sign-in window error (${arg.error}).`); // 12006 = user closed it
    });
  });
}

let activeDialog = null;
function closeDialog() {
  try { if (activeDialog) activeDialog.close(); } catch { /* already closed */ }
  activeDialog = null;
}

// Fallback for when the dialog can't message the pane: the user pastes the token shown in the dialog.
async function pasteToken(e) {
  e.preventDefault();
  const token = $("paste-token").value.trim();
  if (!/^[A-Za-z0-9]{32,}$/.test(token)) return showSigninError("That doesn't look like a Trello code. Copy it again from the sign-in window.");
  Trello.setToken(token);
  try {
    await Trello.me();
  } catch {
    return showSigninError("Trello didn't accept that code. Try connecting again.");
  }
  closeDialog();
  await saveToken(token);
}

async function saveToken(token) {
  $("signin-error").hidden = true;
  $("paste-token").value = "";
  $("form-paste").hidden = true;
  await settings.set({ [SETTINGS.token]: token });
  start();
}

function showSigninError(text) {
  $("signin-error").textContent = text;
  $("signin-error").hidden = false;
}

async function signOut() {
  await settings.set({ [SETTINGS.token]: null });
  Trello.setToken(null);
  show("signin");
}

// ---------- email ----------
async function loadMail() {
  state.mail = await Mail.current();
  const m = state.mail;
  if (!m) return;

  $("title").value = stripPrefixes(m.subject);
  renderDescription();

  $("row-eml").hidden = $("ex-row-eml").hidden = !m.canAttachEml;
  const n = m.attachments.length;
  const fileText = `${n} attachment${n === 1 ? "" : "s"}`;
  $("file-count").textContent = $("ex-file-count").textContent = fileText;
  $("row-files").hidden = $("ex-row-files").hidden = !(m.canAttachFiles && n > 0);
}

function renderDescription() {
  if (!state.mail) return;
  $("desc").value = Mail.description(state.mail, {
    includeBody: $("opt-body").checked,
    includeLink: $("opt-link").checked,
  });
}

// ---------- boards / lists / labels / members ----------
async function loadBoards() {
  const boards = await Trello.boards();
  boards.sort((a, b) =>
    (b.starred - a.starred) || (new Date(b.dateLastView || 0) - new Date(a.dateLastView || 0)));

  const sel = $("board");
  sel.replaceChildren();
  const groups = new Map();
  for (const b of boards) {
    const group = b.starred ? "★ Starred" : (b.organization && b.organization.displayName) || "Personal";
    if (!groups.has(group)) {
      const og = document.createElement("optgroup");
      og.label = group;
      groups.set(group, og);
      sel.append(og);
    }
    groups.get(group).append(option(b.id, b.name));
  }

  const last = settings.get(SETTINGS.board);
  if (last && boards.some((b) => b.id === last)) sel.value = last;
  await onBoardChange();
}

async function onBoardChange() {
  const boardId = $("board").value;
  if (!boardId) return;
  state.labels.clear();
  state.members.clear();

  const [lists, labels, members] = await Promise.all([
    Trello.lists(boardId), Trello.labels(boardId), Trello.members(boardId),
  ]);

  const listSel = $("list");
  listSel.replaceChildren(...lists.map((l) => option(l.id, l.name)));
  const lastList = settings.get(SETTINGS.list);
  if (lastList && lists.some((l) => l.id === lastList)) listSel.value = lastList;

  renderChips($("labels"), labels.map((l) => ({
    id: l.id, text: l.name || (l.color ? l.color : "(no name)"), color: LABEL_COLORS[(l.color || "").split("_")[0]],
  })), state.labels, "No labels on this board");

  renderChips($("members"), members.map((m) => ({ id: m.id, text: m.fullName || m.username })),
    state.members, "No members");
}

function renderChips(container, items, selected, emptyText) {
  container.replaceChildren();
  if (!items.length) {
    const s = document.createElement("span");
    s.className = "muted small";
    s.textContent = emptyText;
    container.append(s);
    return;
  }
  for (const item of items) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip";
    btn.setAttribute("aria-pressed", "false");
    if (item.color) {
      const sw = document.createElement("span");
      sw.className = "swatch";
      sw.style.background = item.color;
      btn.append(sw);
    }
    btn.append(document.createTextNode(item.text));
    btn.addEventListener("click", () => {
      const on = !selected.has(item.id);
      on ? selected.add(item.id) : selected.delete(item.id);
      btn.setAttribute("aria-pressed", String(on));
    });
    container.append(btn);
  }
}

// ---------- attachments ----------
async function attachEmailParts(cardId, { eml, files }) {
  const failures = [];
  const m = state.mail;

  if (eml && m.canAttachEml) {
    try {
      const f = await Mail.emlFile(m);
      await Trello.attachFile(cardId, f.blob, f.name);
    } catch (e) {
      failures.push("the .eml copy of the email");
    }
  }

  if (files && m.canAttachFiles) {
    for (const att of m.attachments) {
      try {
        const f = await Mail.attachmentFile(att);
        if (!f) continue;
        if (f.blob) await Trello.attachFile(cardId, f.blob, f.name);
        else await Trello.attachUrl(cardId, f.url, f.name);
      } catch (e) {
        failures.push(att.name);
      }
    }
  }

  if (!eml && !files && m.link) {
    // Nothing uploaded: still leave a clickable way back to the email.
    await Trello.attachUrl(cardId, m.link, `Email: ${m.subject}`).catch(() => {});
  }
  return failures;
}

// ---------- create card ----------
async function createCard(e) {
  e.preventDefault();
  const btn = $("btn-create");
  btn.disabled = true;
  setStatus("Creating card…");

  try {
    const due = $("due").value ? new Date($("due").value).toISOString() : undefined;
    const card = await Trello.createCard({
      idList: $("list").value,
      name: $("title").value.trim() || state.mail.subject,
      desc: $("desc").value,
      pos: $("pos").value,
      due,
      idLabels: [...state.labels],
      idMembers: [...state.members],
    });

    await settings.set({ [SETTINGS.board]: $("board").value, [SETTINGS.list]: $("list").value });

    setStatus("Uploading attachments…");
    const failures = await attachEmailParts(card.id, {
      eml: $("opt-eml").checked && !$("row-eml").hidden,
      files: $("opt-files").checked && !$("row-files").hidden,
    });

    showDone("Card created", card.name, card.shortUrl, failures);
  } catch (err) {
    setStatus(friendlyError(err), true);
  } finally {
    btn.disabled = false;
  }
}

// ---------- add to existing card ----------
let searchTimer;
function onSearchInput() {
  clearTimeout(searchTimer);
  const q = $("search").value.trim();
  if (q.length < 2) { $("results").replaceChildren(); return; }
  searchTimer = setTimeout(async () => {
    try {
      const cards = (await Trello.searchCards(q)).filter((c) => !c.closed);
      renderResults(cards);
    } catch (err) {
      setStatus(friendlyError(err), true);
    }
  }, 300);
}

function renderResults(cards) {
  const ul = $("results");
  ul.replaceChildren();
  if (!cards.length) {
    const li = document.createElement("li");
    li.className = "muted small";
    li.textContent = "No matching cards.";
    ul.append(li);
    return;
  }
  for (const c of cards) {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.type = "button";
    btn.setAttribute("aria-pressed", String(state.pickedCard && state.pickedCard.id === c.id));
    btn.append(document.createTextNode(c.name));
    const board = document.createElement("span");
    board.className = "board";
    board.textContent = c.board ? c.board.name : "";
    btn.append(board);
    btn.addEventListener("click", () => {
      state.pickedCard = c;
      for (const b of ul.querySelectorAll("button")) b.setAttribute("aria-pressed", "false");
      btn.setAttribute("aria-pressed", "true");
      $("picked-name").textContent = c.name;
      $("existing-actions").hidden = false;
    });
    li.append(btn);
    ul.append(li);
  }
}

async function addToExisting(e) {
  e.preventDefault();
  const card = state.pickedCard;
  if (!card) return;
  const btn = $("btn-attach");
  btn.disabled = true;
  setStatus("Adding to card…");

  try {
    if ($("ex-body").checked) {
      const comment = `📧 **${state.mail.subject}**\n\n` + Mail.description(state.mail, { includeBody: true, includeLink: true });
      await Trello.addComment(card.id, comment);
    }
    const failures = await attachEmailParts(card.id, {
      eml: $("ex-eml").checked && !$("ex-row-eml").hidden,
      files: $("ex-files").checked && !$("ex-row-files").hidden,
    });
    showDone("Added to card", card.name, card.shortUrl, failures);
  } catch (err) {
    setStatus(friendlyError(err), true);
  } finally {
    btn.disabled = false;
  }
}

function showDone(title, cardName, url, failures) {
  setStatus("");
  $("done-title").textContent = title;
  $("done-card").textContent = cardName;
  $("done-link").href = url;
  $("done-warn").hidden = !failures.length;
  $("done-warn").textContent = failures.length ? `Couldn't attach: ${failures.join(", ")}` : "";
  show("done");
}

// ---------- tabs ----------
function switchTab(name) {
  for (const t of document.querySelectorAll(".tab")) {
    const on = t.dataset.tab === name;
    t.classList.toggle("active", on);
    t.setAttribute("aria-selected", String(on));
  }
  for (const p of document.querySelectorAll(".tab-panel")) p.hidden = p.dataset.panel !== name;
  setStatus("");
  if (name === "existing" && !$("search").value && state.mail) {
    $("search").value = stripPrefixes(state.mail.subject).slice(0, 60);
    onSearchInput();
  }
}

// ---------- startup ----------
async function start() {
  const token = settings.get(SETTINGS.token);
  if (!token) return show("signin");
  Trello.setToken(token);
  show("loading");

  try {
    const [me] = await Promise.all([Trello.me(), loadMail(), loadBoards()]);
    $("who").textContent = `Signed in as ${me.fullName || me.username}`;
    show("main");
  } catch (err) {
    if (err.status === 401) return signOut();
    show("main");
    setStatus(friendlyError(err), true);
  }
}

function resetForNewItem() {
  state.pickedCard = null;
  $("search").value = "";
  $("results").replaceChildren();
  $("existing-actions").hidden = true;
  $("due").value = "";
  setStatus("");
  loadMail().catch((err) => setStatus(friendlyError(err), true));
}

Office.onReady(() => {
  if (!window.TRELLO_CONFIG.apiKey || window.TRELLO_CONFIG.apiKey.startsWith("YOUR_")) {
    show("signin");
    $("btn-signin").disabled = true;
    return showSigninError("Set your Trello API key in src/config.js first.");
  }

  $("btn-signin").addEventListener("click", signIn);
  $("form-paste").addEventListener("submit", pasteToken);
  $("btn-signout").addEventListener("click", signOut);
  $("board").addEventListener("change", () => onBoardChange().catch((err) => setStatus(friendlyError(err), true)));
  $("opt-body").addEventListener("change", renderDescription);
  $("opt-link").addEventListener("change", renderDescription);
  $("form-new").addEventListener("submit", createCard);
  $("form-existing").addEventListener("submit", addToExisting);
  $("search").addEventListener("input", onSearchInput);
  for (const t of document.querySelectorAll(".tab")) t.addEventListener("click", () => switchTab(t.dataset.tab));
  $("btn-again").addEventListener("click", () => { resetForNewItem(); show("main"); });

  // Pinned task pane: refresh when the user selects a different email.
  if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      resetForNewItem();
      if (!$("view-done").hidden) show("main");
    });
  }

  start();
});
