// auth.js — Perfis de acesso (senhas armazenadas como hash SHA-256 com salt).
// Para cadastrar/trocar senhas, use a aba "Gerador de Senha" no painel (Admin Master).
// O salt é fixo e propositalmente público — invalida tabelas pré-computadas, não é segredo.

var GAQ_SALT = "sesc-gaq-painel-2026";

// ── SHA-256 implementação pura (funciona em http://, https:// e file://) ─────
function _gaqSha256Pure(str) {
  function utf8Encode(s) {
    var out = [], i = 0, c1, c2;
    while (i < s.length) {
      c1 = s.charCodeAt(i++);
      if (c1 < 0x80) out.push(c1);
      else if (c1 < 0x800) { out.push(0xc0 | (c1 >> 6)); out.push(0x80 | (c1 & 0x3f)); }
      else if ((c1 & 0xfc00) === 0xd800 && i < s.length && (s.charCodeAt(i) & 0xfc00) === 0xdc00) {
        c2 = s.charCodeAt(i++);
        c1 = 0x10000 + (((c1 & 0x3ff) << 10) | (c2 & 0x3ff));
        out.push(0xf0 | (c1 >> 18)); out.push(0x80 | ((c1 >> 12) & 0x3f));
        out.push(0x80 | ((c1 >> 6) & 0x3f)); out.push(0x80 | (c1 & 0x3f));
      } else { out.push(0xe0 | (c1 >> 12)); out.push(0x80 | ((c1 >> 6) & 0x3f)); out.push(0x80 | (c1 & 0x3f)); }
    }
    return out;
  }
  var msg = utf8Encode(str);
  var origLen = msg.length;
  msg.push(0x80);
  while ((msg.length % 64) !== 56) msg.push(0);
  var bitLenLow = (origLen * 8) >>> 0;
  var bitLenHigh = Math.floor(origLen / 0x20000000);
  msg.push((bitLenHigh >>> 24) & 0xff, (bitLenHigh >>> 16) & 0xff, (bitLenHigh >>> 8) & 0xff, bitLenHigh & 0xff);
  msg.push((bitLenLow >>> 24) & 0xff, (bitLenLow >>> 16) & 0xff, (bitLenLow >>> 8) & 0xff, bitLenLow & 0xff);

  var h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  var k = [
    0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
    0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
    0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
    0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
    0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
    0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
    0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
    0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2
  ];
  var w = new Array(64);
  for (var i = 0; i < msg.length; i += 64) {
    for (var j = 0; j < 16; j++) {
      w[j] = ((msg[i+j*4]<<24)|(msg[i+j*4+1]<<16)|(msg[i+j*4+2]<<8)|msg[i+j*4+3]) >>> 0;
    }
    for (var j = 16; j < 64; j++) {
      var s0 = (((w[j-15]>>>7)|(w[j-15]<<25)) ^ ((w[j-15]>>>18)|(w[j-15]<<14)) ^ (w[j-15]>>>3)) >>> 0;
      var s1 = (((w[j-2]>>>17)|(w[j-2]<<15)) ^ ((w[j-2]>>>19)|(w[j-2]<<13)) ^ (w[j-2]>>>10)) >>> 0;
      w[j] = (w[j-16] + s0 + w[j-7] + s1) >>> 0;
    }
    var a=h[0],b=h[1],c=h[2],d=h[3],e=h[4],f=h[5],g=h[6],hh=h[7];
    for (var j = 0; j < 64; j++) {
      var S1 = (((e>>>6)|(e<<26)) ^ ((e>>>11)|(e<<21)) ^ ((e>>>25)|(e<<7))) >>> 0;
      var ch = ((e & f) ^ (~e & g)) >>> 0;
      var t1 = (hh + S1 + ch + k[j] + w[j]) >>> 0;
      var S0 = (((a>>>2)|(a<<30)) ^ ((a>>>13)|(a<<19)) ^ ((a>>>22)|(a<<10))) >>> 0;
      var mj = ((a & b) ^ (a & c) ^ (b & c)) >>> 0;
      var t2 = (S0 + mj) >>> 0;
      hh = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    h[0] = (h[0]+a)>>>0; h[1] = (h[1]+b)>>>0; h[2] = (h[2]+c)>>>0; h[3] = (h[3]+d)>>>0;
    h[4] = (h[4]+e)>>>0; h[5] = (h[5]+f)>>>0; h[6] = (h[6]+g)>>>0; h[7] = (h[7]+hh)>>>0;
  }
  return h.map(function(x){ return ("00000000"+x.toString(16)).slice(-8); }).join("");
}

// SHA-256 com fallback: tenta crypto.subtle (rápido, nativo) e usa pure-JS se indisponível.
async function gaqSha256(text) {
  if (typeof crypto !== "undefined" && crypto && crypto.subtle && typeof crypto.subtle.digest === "function") {
    try {
      var buf = new TextEncoder().encode(text);
      var hash = await crypto.subtle.digest("SHA-256", buf);
      var hex = "";
      var arr = new Uint8Array(hash);
      for (var i = 0; i < arr.length; i++) hex += ("00" + arr[i].toString(16)).slice(-2);
      return hex;
    } catch (e) { /* falha → cai pro fallback abaixo */ }
  }
  return _gaqSha256Pure(text);
}

// Hash de senha com salt — sempre usa esta função, nunca chame gaqSha256 direto pra senha.
async function gaqHashSenha(senha) {
  return await gaqSha256(GAQ_SALT + ":" + (senha || ""));
}

// ── Perfis de Acesso (chave = hash SHA-256 da senha; valor = nome ou null) ──
// null = administrador (vê tudo); string = filtra pelo nome exato do responsável
var PERFIS_ACESSO = {
  "2957ee13b6fb43409ef9171ba83eaa745e94063745079997ffa4c540e65935f1": null,            // Admin (visão completa)
  "8c332e47b417c9c071869e7511c4540fe8da861df670cb4a85f0baced8adcd54": "Alessandra",
  "85e3e333bd21154fe71e90e6c3bc03f146a5f2240794619b94927da8133830c2": "Viviane",
  "4ee01be25040dc10bfc3c05d654b5bf86d50d18f5aa52c0059013bff0d57565d": "Wesley",
  "f2dc04188c90f20c854b0e32e7656570c2ee6efacc83d1c1cc63555936eecc13": "Mauricio",
  "6b72da11f42d1ad59cbc8ef85a9e0f251886a08593d4bc06708464be89cf6078": "Pamela",
  "0329e3e7bf5e48416d4ea79da127ca3091327fac1d851c132b7df4e20632bb48": "Bruna",
  "1761dde646c5197702707f75523d2edcf9087f97bf38c0c739319f121489a7a0": "Gláucia",
  "f503a285c35ca0f2596b9790c286c9a3f6cdfef79d767b07022b2fef1a3b71bc": "Inacio",
  "7c9a2aa544a9eeb7365ca6c989a5516c4cd416c94cfa83402a6389b634feea6f": "Raphael",
};

// ── Admin Master — visão total + Qualidade de Dados ──
var ADMIN_MASTER_PERFIS = {
  "50cd403e7a0a78b44c69a09baadc030fd09326945d8e45b9c38a74e51b4c8da0": { nome: "Admin Master" },
};

// ── Senha do painel de Gestão/Performance (hash) ──
var GESTAO_PASS = "418d23f2fb63e665c8c1c28c137a84b5494f88a3cd47cdc53a812b4f3a07564f";

// ── Senha do upload de base (hash) ──
var UPLOAD_PASS = "7b6fa9c245c88c2bc8fdc32e61612bad4744e34d6e53a1a03de0121260b38e0b";

// ── Acesso por Área (chave = hash da senha; valor = código exato da área) ──
// Valor pode ser uma área única ("GTI") ou várias separadas por vírgula
// (ex: "GIN,ALMOXARIFADO") — o painel filtra pela união de todas elas.
var AREA_PERFIS = {
  "6c34afc9823cc0d1ee46a66ad2f817ced10e017123c56481171c595ba446cea8": "GTI",
  "b2546ddd65f727b5fa864f309742082a569b3599e8e345fc18f5fe04295f36a1": "PLE",
  "7083cc7cde253eeadeed9d40e0c13fb09f3a3e3be6ab4bd460320f26f4849302": "ASCOM",
  "40cab534a65202df90df67d9285dd55d561f953ce972b2fa0ac6c2d08b257d93": "PSC",
  "09576c8ab9368e3e46a70fb9e80c236a3130710f745c5a0e62b2d9285e869395": "GIN,ALMOXARIFADO",
};

// ── Diretoria — chave = hash da senha; valor = { nome, areasDestaque: [...] } ──
var DIRETOR_PERFIS = {
  "08e04a673111fa37b9b8c6de47a1b176fc14e27bde69f58c590e665b74aac885": { nome: "Alan", areasDestaque: ["ASCOM", "TI", "GIN"] },
};
