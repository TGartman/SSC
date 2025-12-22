/**
 * Azure Function: graphics/generate (HTTP POST)
 *
 * What this does (end-to-end):
 * 1) Auth to Microsoft Graph with client_credentials (TENANT_ID/CLIENT_ID/CLIENT_SECRET)
 * 2) Auto-resolves:
 *    - brand root folder (SSC / MT / DWELL)
 *    - Logos folder -> picks best logo file automatically
 *    - Generated Posts folder -> auto upload target when saveToSharePoint=true
 * 3) Downloads:
 *    - product image (required) from drive/item id
 *    - logo image (auto if not provided)
 * 4) Composes a 1080x1080 / 1080x1350 / 1080x1920 PNG:
 *    - product image as background (cover)
 *    - text overlay (simple, clean)
 *    - logo bottom-right with safe padding
 * 5) Uploads to SharePoint into /<Brand>/Generated Posts/<fileName>
 *
 * Dependencies (npm i):
 *   axios
 *   sharp
 *
 * NOTE on Cochin font:
 *   Node+sharp does not embed "Cochin" unless the font is available and used via SVG.
 *   This implementation renders text via SVG and sets font-family to "Cochin, serif".
 *   For guaranteed Cochin, place a licensed Cochin .ttf in your function folder and
 *   reference it in the SVG via @font-face (see TODO in makeTextSvg()).
 */

const axios = require("axios");
const sharp = require("sharp");

const DRIVE_ID = process.env.PRODUCT_IMAGES_DRIVE_ID;

// Brand root folder itemIds in the Product-Images drive (from your Graph results)
const BRAND_ROOTS = {
  SSC: "01POEEX7TJKWFWMFT2RFBKU63QDRMVT5JR", // Southern Shutter Company
  MT: "01POEEX7TBJQTUB7USFVC2EYSQCJVC7BZF",  // Millwork Traders
  DWELL: "01POEEX7THJEFWFLPDIFFZK6CESICDKTKE" // Dwell Shutter and Blinds
};

// Optional: prefer specific logo filenames
const PREFERRED_LOGO_FILENAME = {
  SSC: "SSC-Logo.png",
  MT: "Millwork-Traders-Logo.png",
  DWELL: "Dwell-Logo.png"
};

const FORMATS = {
  square_1080: { w: 1080, h: 1080 },
  portrait_1080x1350: { w: 1080, h: 1350 },
  story_1080x1920: { w: 1080, h: 1920 }
};

module.exports = async function (context, req) {
  try {
    // ---- Validate env ----
    const TENANT_ID = process.env.TENANT_ID;
    const CLIENT_ID = process.env.CLIENT_ID;
    const CLIENT_SECRET = process.env.CLIENT_SECRET;
module.exports = async function (context, req) {
  context.res = { status: 200, body: { ok: true, note: "generate restored" } };
};

    if (!DRIVE_ID) {
      return bad(context, "Missing PRODUCT_IMAGES_DRIVE_ID in local.settings.json");
    }
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
      return serverErr(context, "Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in local.settings.json");
    }

    // ---- Parse & validate request ----
    const body = req.body || {};
    const brand = String(body.brand || "").toUpperCase();
    if (!BRAND_ROOTS[brand]) {
      return bad(context, "brand must be SSC, MT, or DWELL");
    }

    const formatKey = String(body.format || "");
    if (!FORMATS[formatKey]) {
      return bad(context, "Invalid format. Use square_1080, portrait_1080x1350, or story_1080x1920.");
    }
    const { w: width, h: height } = FORMATS[formatKey];

    const productImage = body.productImage || {};
    if (!productImage.driveId || !productImage.itemId) {
      return bad(context, "productImage.driveId and productImage.itemId are required");
    }

    // Style defaults
    const style = body.style || {};
    const safePaddingPx = Number.isFinite(style.safePaddingPx) ? style.safePaddingPx : 64;
    const logoMaxWidthPct = Number.isFinite(style.logoMaxWidthPct) ? style.logoMaxWidthPct : 16; // % of canvas width
    const logoPlacement = String(style.logoPlacement || "bottom_right"); // only bottom_right supported here
    const textAlign = String(style.textAlign || "left");

    // Output defaults
    const output = body.output || {};
    const saveToSharePoint = !!output.saveToSharePoint;
    const returnBase64 = !!output.returnBase64;
    const fileName = output.fileName || `${brand}_${formatKey}_${Date.now()}.png`;

    // ---- Auth to Graph ----
    const accessToken = await getGraphToken({ TENANT_ID, CLIENT_ID, CLIENT_SECRET });

    // ---- Auto-resolve brand folders ----
    const brandRootId = BRAND_ROOTS[brand];

    const logosFolderId = await findChildFolderId(accessToken, DRIVE_ID, brandRootId, "Logos");
    const generatedPostsFolderId = await findChildFolderId(accessToken, DRIVE_ID, brandRootId, "Generated Posts");

    if (!logosFolderId) return serverErr(context, `Could not find Logos folder under brand ${brand}`);
    if (saveToSharePoint && !generatedPostsFolderId) return serverErr(context, `Could not find Generated Posts folder under brand ${brand}`);

    // ---- Auto-pick logo if not provided ----
    let logoImage = body.logoImage || {};
    if (!logoImage.driveId) logoImage.driveId = DRIVE_ID;

    if (!logoImage.itemId) {
      const picked = await pickLogoItem(accessToken, DRIVE_ID, logosFolderId, brand);
      if (!picked?.id) return serverErr(context, `No logo files found in Logos folder for ${brand}`);
      logoImage.itemId = picked.id;
      logoImage.name = picked.name;
    }

    // ---- Download product + logo bytes from Graph ----
    const productBuf = await downloadDriveItemContent(accessToken, productImage.driveId, productImage.itemId);
    const logoBuf = await downloadDriveItemContent(accessToken, logoImage.driveId, logoImage.itemId);

    // ---- Compose output image ----
    const composedPng = await composeGraphicPng({
      width,
      height,
      productBuf,
      logoBuf,
      headline: body.headline || "",
      subhead: body.subhead || "",
      cta: body.cta || "",
      safePaddingPx,
      logoMaxWidthPct,
      logoPlacement,
      textAlign
    });

    // ---- Save to SharePoint (brand/Generated Posts) ----
    let saved = null;
    if (saveToSharePoint) {
      saved = await uploadToFolder(accessToken, DRIVE_ID, generatedPostsFolderId, fileName, composedPng);
    }

    // ---- Return response ----
    const resBody = {
      brand,
      format: formatKey,
      width,
      height,
      mimeType: "image/png",
      base64: returnBase64 ? composedPng.toString("base64") : null,
      saved
    };

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: resBody
    };
  } catch (err) {
    const status = err?.response?.status;
    const data = err?.response?.data;
    context.log.error("ERROR:", status, data || err?.message || err);

    // Bubble a friendlier message
    if (status === 401) return serverErr(context, "Request failed with status code 401 (check CLIENT_SECRET / token)");
    if (status === 403) return serverErr(context, "Request failed with status code 403 (check Graph permission Sites.ReadWrite.All + consent)");
    return serverErr(context, err?.message || "Unknown server error");
  }
};

// ------------------------ Graph Auth ------------------------

async function getGraphToken({ TENANT_ID, CLIENT_ID, CLIENT_SECRET }) {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const form = new URLSearchParams();
  form.append("grant_type", "client_credentials");
  form.append("client_id", CLIENT_ID);
  form.append("client_secret", CLIENT_SECRET);
  form.append("scope", "https://graph.microsoft.com/.default");

  const res = await axios.post(url, form.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" }
  });
  return res.data.access_token;
}

// ------------------------ Graph Helpers ------------------------

async function graphGet(accessToken, url) {
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  return res.data;
}

async function findChildFolderId(accessToken, driveId, parentItemId, folderName) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentItemId}/children?$select=id,name,folder`;
  const data = await graphGet(accessToken, url);
  const hit = (data.value || []).find(
    (x) => x.folder && String(x.name || "").toLowerCase() === folderName.toLowerCase()
  );
  return hit?.id || null;
}

async function pickLogoItem(accessToken, driveId, logosFolderId, brandKey) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${logosFolderId}/children?$select=id,name,file`;
  const data = await graphGet(accessToken, url);
  const files = (data.value || []).filter((x) => x.file);

  if (!files.length) return null;

  const preferred = PREFERRED_LOGO_FILENAME[brandKey];
  if (preferred) {
    const hit = files.find((f) => String(f.name).toLowerCase() === preferred.toLowerCase());
    if (hit) return hit;
  }

  const img = files.find((f) => /\.(png|jpg|jpeg|webp)$/i.test(String(f.name)));
  return img || files[0];
}

async function downloadDriveItemContent(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: "arraybuffer"
  });
  return Buffer.from(res.data);
}

// Upload into a folder by itemId: PUT /drives/{driveId}/items/{folderId}:/{filename}:/content
async function uploadToFolder(accessToken, driveId, folderItemId, fileName, contentBuf) {
  const safeName = String(fileName).replace(/[\\/:*?"<>|]+/g, "_");
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderItemId}:/${encodeURIComponent(
    safeName
  )}:/content`;

  const res = await axios.put(url, contentBuf, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "image/png"
    }
  });

  // res.data is the driveItem
  return {
    savedToSharePoint: true,
    driveId,
    itemId: res.data?.id,
    webUrl: res.data?.webUrl
  };
}

// ------------------------ Composition ------------------------

async function composeGraphicPng({
  width,
  height,
  productBuf,
  logoBuf,
  headline,
  subhead,
  cta,
  safePaddingPx,
  logoMaxWidthPct,
  logoPlacement,
  textAlign
}) {
  // Background: product image -> cover crop to canvas
  const bg = await sharp(productBuf)
    .resize(width, height, { fit: "cover" })
    .toBuffer();

  // Add a subtle dark gradient overlay for legible text (top-to-bottom)
  const gradientSvg = makeGradientOverlaySvg(width, height);

  // Logo: resize to max width % of canvas
  const maxLogoW = Math.round((logoMaxWidthPct / 100) * width);
  const logoPng = await sharp(logoBuf)
    .resize({ width: maxLogoW, withoutEnlargement: true })
    .png()
    .toBuffer();

  const logoMeta = await sharp(logoPng).metadata();
  const logoW = logoMeta.width || maxLogoW;
  const logoH = logoMeta.height || Math.round(maxLogoW * 0.4);

  // Logo position (bottom-right)
  const logoLeft =
    logoPlacement === "bottom_right" ? width - safePaddingPx - logoW : width - safePaddingPx - logoW;
  const logoTop = height - safePaddingPx - logoH;

  // Text SVG (headline/subhead/cta)
  const textSvg = makeTextSvg({
    width,
    height,
    headline,
    subhead,
    cta,
    safePaddingPx,
    textAlign
  });

  const out = await sharp(bg)
    .composite([
      { input: Buffer.from(gradientSvg), top: 0, left: 0 },
      { input: Buffer.from(textSvg), top: 0, left: 0 },
      { input: logoPng, top: logoTop, left: logoLeft }
    ])
    .png()
    .toBuffer();

  return out;
}

function makeGradientOverlaySvg(width, height) {
  // Dark overlay at top and bottom for readability
  return `
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
      <stop offset="0%" stop-color="rgba(0,0,0,0.55)"/>
      <stop offset="35%" stop-color="rgba(0,0,0,0.08)"/>
      <stop offset="70%" stop-color="rgba(0,0,0,0.08)"/>
      <stop offset="100%" stop-color="rgba(0,0,0,0.55)"/>
    </linearGradient>
  </defs>
  <rect x="0" y="0" width="${width}" height="${height}" fill="url(#g)"/>
</svg>`;
}

function escapeXml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function makeTextSvg({ width, height, headline, subhead, cta, safePaddingPx, textAlign }) {
  const x = safePaddingPx;
  const maxW = width - safePaddingPx * 2;

  const h1 = escapeXml(headline);
  const h2 = escapeXml(subhead);
  const h3 = escapeXml(cta);

  const anchor = textAlign === "center" ? "middle" : "start";
  const textX = textAlign === "center" ? Math.round(width / 2) : x;

  // TODO (for guaranteed Cochin):
  // 1) put a licensed Cochin TTF in your function folder, e.g. ./fonts/Cochin.ttf
  // 2) base64 it at startup and include:
  //    @font-face { font-family: 'CochinLocal'; src: url(data:font/ttf;base64,...) format('truetype'); }
  // 3) then use font-family: 'CochinLocal';
  //
  // For now: uses system Cochin if available, fallback to serif.
  return `
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
  <style>
    .h1 { font-family: Cochin, serif; font-weight: 600; font-size: 72px; fill: rgba(255,255,255,0.98); }
    .h2 { font-family: Cochin, serif; font-weight: 400; font-size: 40px; fill: rgba(255,255,255,0.92); }
    .h3 { font-family: Cochin, serif; font-weight: 600; font-size: 36px; fill: rgba(255,255,255,0.98); }
  </style>

  <!-- Text block at top -->
  <g>
    ${h1 ? `<text x="${textX}" y="${safePaddingPx + 80}" text-anchor="${anchor}" class="h1">${h1}</text>` : ""}
    ${
      h2
        ? `<text x="${textX}" y="${safePaddingPx + 140}" text-anchor="${anchor}" class="h2">${h2}</text>`
        : ""
    }
  </g>

  <!-- CTA at bottom-left (above safe area) -->
  ${
    h3
      ? `<text x="${textX}" y="${height - safePaddingPx - 30}" text-anchor="${anchor}" class="h3">${h3}</text>`
      : ""
  }

  <!-- Invisible box to help SVG sizing in some renderers -->
  <rect x="${x}" y="${safePaddingPx}" width="${maxW}" height="${height - safePaddingPx * 2}" fill="rgba(0,0,0,0)"/>
</svg>`;
}

// ------------------------ Responses ------------------------

function bad(context, message) {
  context.res = {
    status: 400,
    headers: { "Content-Type": "application/json" },
    body: { error: { code: "bad_request", message } }
  };
}

function serverErr(context, message) {
  context.res = {
    status: 500,
    headers: { "Content-Type": "application/json" },
    body: { error: { code: "server_error", message } }
  };
}