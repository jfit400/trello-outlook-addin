// Minimal Trello REST client. Key/token go in the query string (Trello's standard),
// everything else is sent as a form body so long descriptions don't blow URL limits.
const Trello = (() => {
  const API = "https://api.trello.com/1";
  let token = null;

  function setToken(t) { token = t; }

  async function request(method, path, params = {}, formData) {
    const url = new URL(API + path);
    url.searchParams.set("key", window.TRELLO_CONFIG.apiKey);
    url.searchParams.set("token", token);

    let body;
    if (formData) {
      body = formData;
    } else if (method === "GET") {
      appendParams(url.searchParams, params);
    } else {
      body = new URLSearchParams();
      appendParams(body, params);
    }

    const res = await fetch(url, { method, body });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      const err = new Error(text || `${res.status} ${res.statusText}`);
      err.status = res.status;
      throw err;
    }
    return res.json();
  }

  function appendParams(target, params) {
    for (const [k, v] of Object.entries(params)) {
      if (v === undefined || v === null || v === "" || (Array.isArray(v) && !v.length)) continue;
      target.set(k, Array.isArray(v) ? v.join(",") : String(v));
    }
  }

  return {
    setToken,

    me: () => request("GET", "/members/me", { fields: "fullName,username,avatarUrl" }),

    boards: () => request("GET", "/members/me/boards", {
      filter: "open",
      fields: "name,dateLastView,starred,idOrganization",
      organization: "true",
      organization_fields: "displayName",
    }),

    lists: (boardId) => request("GET", `/boards/${boardId}/lists`, { filter: "open", fields: "name" }),
    labels: (boardId) => request("GET", `/boards/${boardId}/labels`, { fields: "name,color", limit: 1000 }),
    members: (boardId) => request("GET", `/boards/${boardId}/members`, { fields: "fullName,username" }),

    createCard: (card) => request("POST", "/cards", card),

    addComment: (cardId, text) => request("POST", `/cards/${cardId}/actions/comments`, { text }),

    attachFile(cardId, blob, name) {
      const fd = new FormData();
      fd.append("file", blob, name);
      fd.append("name", name);
      return request("POST", `/cards/${cardId}/attachments`, {}, fd);
    },

    attachUrl: (cardId, url, name) => request("POST", `/cards/${cardId}/attachments`, { url, name }),

    searchCards: (query) => request("GET", "/search", {
      query,
      modelTypes: "cards",
      card_fields: "name,shortUrl,closed",
      card_board: "true",
      board_fields: "name",
      cards_limit: 15,
      partial: "true",
    }).then((r) => r.cards || []),
  };
})();
