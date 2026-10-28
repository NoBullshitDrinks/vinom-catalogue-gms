// Worker Cloudflare : reçoit les souhaits de commande du catalogue GMS et les envoie par e-mail via Mailjet.
// Variables à définir dans le Worker (Settings > Variables) :
//   MJ_APIKEY_PUBLIC, MJ_APIKEY_PRIVATE (secrets)  ·  FROM_EMAIL (expéditeur validé dans Mailjet)
//   TO_EMAIL = a.durand@vinom.fr  ·  ALLOWED_ORIGIN = https://nobullshitdrinks.github.io
export default {
  async fetch(req, env) {
    const cors = {
      "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };
    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    if (req.method !== "POST") return new Response("Méthode non autorisée", { status: 405, headers: cors });
    if (req.headers.get("Origin") !== env.ALLOWED_ORIGIN) return new Response("Origine refusée", { status: 403, headers: cors });
    let b;
    try { b = await req.json(); } catch { return new Response("JSON invalide", { status: 400, headers: cors }); }
    if (b.hp) return new Response("ok", { status: 200, headers: cors }); // champ piège rempli : robot, on ignore
    const subject = String(b.subject || "").slice(0, 200);
    const text = String(b.text || "").slice(0, 20000);
    if (!subject || !text) return new Response("Demande vide", { status: 400, headers: cors });
    const msg = {
      From: { Email: env.FROM_EMAIL, Name: "Catalogue GMS VINOM" },
      To: [{ Email: env.TO_EMAIL }],
      Subject: subject,
      TextPart: text,
    };
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(b.replyTo || "")) msg.ReplyTo = { Email: b.replyTo };
    const r = await fetch("https://api.mailjet.com/v3.1/send", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Basic " + btoa(env.MJ_APIKEY_PUBLIC + ":" + env.MJ_APIKEY_PRIVATE),
      },
      body: JSON.stringify({ Messages: [msg] }),
    });
    return new Response(r.ok ? "ok" : "Erreur Mailjet " + r.status, { status: r.ok ? 200 : 502, headers: cors });
  },
};
