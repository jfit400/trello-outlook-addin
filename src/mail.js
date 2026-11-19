// Helpers for reading the currently selected Outlook message.
const Mail = (() => {
  const DESC_LIMIT = 16000; // Trello caps descriptions at 16384 chars.

  function promisify(fn) {
    return new Promise((resolve, reject) =>
      fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))));
  }

  const supports = (v) => Office.context.requirements.isSetSupported("Mailbox", v);

  // Best-effort deep link that reopens this message in Outlook on the web.
  function webLink(item) {
    try {
      const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
      const me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
      const host = /@(outlook|hotmail|live|msn)\./.test(me) ? "https://outlook.live.com" : "https://outlook.office.com";
      return `${host}/mail/deeplink/read/${encodeURIComponent(restId)}`;
    } catch {
      return null;
    }
  }

  async function current() {
    const item = Office.context.mailbox.item;
    if (!item) return null;

    const body = await promisify((cb) => item.body.getAsync(Office.CoercionType.Text, cb)).catch(() => "");
    const from = item.from || item.sender || {};
    const attachments = (item.attachments || []).filter((a) => !a.isInline && a.attachmentType === "file");

    return {
      subject: item.subject || "(no subject)",
      fromName: from.displayName || "",
      fromEmail: from.emailAddress || "",
      date: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
      body: (body || "").replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim(),
      attachments,
      link: webLink(item),
      canAttachEml: supports("1.14") && typeof item.getAsFileAsync === "function",
      canAttachFiles: supports("1.8") && typeof item.getAttachmentContentAsync === "function",
    };
  }

  function header(mail, includeLink) {
    const lines = [];
    const who = mail.fromName && mail.fromEmail ? `${mail.fromName} <${mail.fromEmail}>` : mail.fromName || mail.fromEmail;
    if (who) lines.push(`**From:** ${who}`);
    if (mail.date) lines.push(`**Date:** ${mail.date.toLocaleString()}`);
    if (includeLink && mail.link) lines.push(`[Open email in Outlook](${mail.link})`);
    return lines.join("  \n");
  }

  function description(mail, { includeBody, includeLink }) {
    let desc = header(mail, includeLink);
    if (includeBody && mail.body) desc += `\n\n---\n\n${mail.body}`;
    return desc.length > DESC_LIMIT ? desc.slice(0, DESC_LIMIT - 20) + "\n\n…(truncated)" : desc;
  }

  function base64ToBlob(b64, type) {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return new Blob([bytes], { type });
  }

  async function emlFile(mail) {
    const b64 = await promisify((cb) => Office.context.mailbox.item.getAsFileAsync(cb));
    const safe = mail.subject.replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 80) || "email";
    return { blob: base64ToBlob(b64, "message/rfc822"), name: `${safe}.eml` };
  }

  async function attachmentFile(att) {
    const c = await promisify((cb) => Office.context.mailbox.item.getAttachmentContentAsync(att.id, cb));
    if (c.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
      return { blob: base64ToBlob(c.content, att.contentType || "application/octet-stream"), name: att.name };
    }
    if (c.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
      return { url: c.content, name: att.name };
    }
    return null; // eml / iCalendar item attachments are skipped
  }

  return { current, description, emlFile, attachmentFile };
})();
