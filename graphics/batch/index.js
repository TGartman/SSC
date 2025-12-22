const axios = require("axios");
const sharp = require("sharp");

// ====== ENV ======
const DRIVE_ID = process.env.PRODUCT_IMAGES_DRIVE_ID;

// Brand root folder itemIds in Product-Images drive
const BRAND_ROOTS = {
  SSC: "01POEEX7TJKWFWMFT2RFBKU63QDRMVT5JR",
  MT: "01POEEX7TBJQTUB7USFVC2EYSQCJVC7BZF",
  DWELL: "01POEEX7THJEFWFLPDIFFZK6CESICDKTKE"
};

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

// ====== HANDLER ======
module.exports = async function (context, req) {
  try {
    const TENANT_ID = process.env.TENANT_ID;
    const CLIENT_ID = process.env.CLIENT_ID;
    const CLIENT_SECRET = process.env.CLIENT_SECRET;

    if (!DRIVE_ID) return serverErr(context, "Missing PRODUCT_IMAGES_DRIVE_ID in local.settings.json");
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
      return serverErr(context, "Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in local.settings.json");
    }

    const body = req.body || {};
    const brand = String(body.brand || "").toUpperCase();
    if (!BRAND_ROOTS[brand]) return bad(context, "brand must be SSC, MT, or DWELL");

    const formatKey = String(body.format || "square_1080");
    if (!FORMATS[formatKey]) return bad(context, "Invalid format. Use square_1080, portrait_1080x1350, or story_1080x1920.");
    const { w: width, h: height } = FORMATS[formatKey];

    // What to scan
    const area = String(body.area || "Exterior Products"); // "Exterior Products" or "Interior Products"
    if (!["Exterior Products", "Interior Products"].includes(area)) {
      return bad(context, "area must be 'Exterior Products' or 'Interior Products'");
    }

    // Limits & behavior
    const limit = Number.isFinite(body.limit) ? body.limit : 10; // safety default
    const dryRun = !!body.dryRun; // if true, does not upload
    const pickStrategy = String(body.pickStrategy || "preferLifestyle"); // preferLifestyle | catalogOnly | lifestyleOnly

    // Text options
    const headlineTemplate = String(body.headlineTemplate || "{product}");
    const subhead = String(body.subhead || "");
    const cta = String(body.cta || "");

    const style = body.style || {};
    const safePaddingPx = Number.isFinite(style.safePaddingPx) ? style.safePaddingPx : 64;
    const logoMaxWidthPct = Number.isFinite(style.logoMaxWidthPct) ? style.logoMaxWidthPct : 16;
    const logoPlacement = String(style.logoPlacement || "bottom_right");
    const textAlign = String(style.textAlign || "left");

    // Output naming
    const output = body.output || {};
    const saveToSharePoint = output.saveToSharePoint !== false; // default true
    const fileNameTemplate = String(output.fileNameTemplate || "{brand}_{product}_square_1080.png");

    const accessToken = await getGraphToken({ TENANT_ID, CLIENT_ID, CLIENT_SECRET });

    // Resolve brand folders
    const brandRootId = BRAND_ROOTS[brand];
    const logosFolderId = await findChildFolderId(accessToken, DRIVE_ID, brandRootId, "Logos");
    const generatedPostsFolderId = await findChildFolderId(accessToken, DRIVE_ID, brandRootId, "Generated Posts");
    const areaFolderId = await findChildFolderId(accessToken, DRIVE_ID, brandRootId, area);

    if (!logosFolderId) return serverErr(context, `Could not find Logos folder under brand ${brand}`);
    if (!areaFolderId) return serverErr(context, `Could not find '${area}' folder under brand ${brand}`);
    if (saveToSharePoint && !generatedPostsFolderId) return serverErr(context, `Could not find Generated Posts folder under brand ${brand}`);

    // Pick logo
    const pickedLogo = await pickLogoItem(accessToken, DRIVE_ID, logosFolderId, brand);
    if (!pickedLogo?.id) return serverErr(context, `No logo file found in Logos folder for ${brand}`);
    const logoBuf = await downloadDriveItemContent(accessToken, DRIVE_ID, pickedLogo.id);

    // List product folders under area
    const productFolders = await listChildFolders(accessToken, DRIVE_ID, areaFolderId);

    const results = [];
    for (const productFolder of productFolders.slice(0, Math.max(0, limit))) {
      const productName = productFolder.name;

      // Find "Lifestyle Images" and/or "Catalog Images"
      const lifestyleId = await findChildFolderId(accessToken, DRIVE_ID, productFolder.id, "Lifestyle Images");
      const catalogId = await findChildFolderId(accessToken, DRIVE_ID, productFolder.id, "Catalog Images");

      // Choose which folder to pull from based on strategy
      const chosenFolderId = await chooseImageFolder(accessToken, DRIVE_ID, { lifestyleId, catalogId, pickStrategy });

      if (!chosenFolderId) {
        results.push({
          product: productName,
          ok: false,
          reason: "No image folder found (Lifestyle Images / Catalog Images empty or missing)"
        });
        continue;
      }

      const imageItem = await pickFirstImageFile(accessToken, DRIVE_ID, chosenFolderId);
      if (!imageItem?.id) {
        results.push({ product: productName, ok: false, reason: "No images found in chosen folder" });
        continue;
      }

      const productBuf = await downloadDriveItemContent(accessToken, DRIVE_ID, imageItem.id);

      // Compose
      const composedPng = await composeGraphicPng({
        width,
        height,
        productBuf,
        logoBuf,
        headline: headlineTemplate.replaceAll("{product}", productName),
        subhead,
        cta,
        safePaddingPx,
        logoMaxWidthPct,
        logoPlacement,
        textAlign
      });

      const finalFileName = sanitizeFileName(
        fileNameTemplate
          .replaceAll("{brand}", brand)
          .replaceAll("{product}", productName.replace(/\s+/g, "_"))
          .replaceAll("{format}", formatKey)
      );

      let saved = null;
      if (!dryRun && saveToSharePoint) {
        saved = await uploadToFolder(accessToken, DRIVE_ID, generatedPostsFolderId, finalFileName, composedPng);
      }

      results.push({
        product: productName,
        ok: true,
        pickedImage: { name: imageItem.name, itemId: imageItem.id },
        saved
      });

      // (Optional) tiny delay to be nice to Graph
      await sleep(150);
    }

    context.res = {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        brand,
        area,
        format: formatKey,
        limit,
        dryRun,
        countAttempted: Math.min(productFolders.length, Math.max(0, limit)),
        totalProductFoldersFound: productFolders.length,
        results
      }
    };
  } catch (err) {
    const status = err?.response?.status;
    const data = err?.response?.data;
    context.log.error("ERROR:", status, data || err?.message || err);
    return serverErr(context, `Batch failed: ${status || ""} ${err?.message || "Unknown error"}`.trim());
  }
};

// ====== Graph Auth ======
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

// ====== Graph Helpers ======
async function graphGet(accessToken, url) {
  const res = await axios.get(url, { headers: { Authorization: `Bearer ${accessToken}` } });
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

async function listChildFolders(accessToken, driveId, parentItemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentItemId}/children?$select=id,name,folder`;
  const data = await graphGet(accessToken, url);
  return (data.value || []).filter((x) => x.folder);
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

async function pickFirstImageFile(accessToken, driveId, folderItemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderItemId}/children?$select=id,name,file`;
  const data = await graphGet(accessToken, url);
  const files = (data.value || []).filter((x) => x.file && /\.(png|jpg|jpeg|webp)$/i.test(String(x.name)));
  return files[0] || null;
}

async function chooseImageFolder(accessToken, driveId, { lifestyleId, catalogId, pickStrategy }) {
  // helper to check if folder has at least 1 image
  async function hasImages(folderId) {
    if (!folderId) return false;
    const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$select=name,file`;
    const data = await graphGet(accessToken, url);
    return (data.value || []).some((x) => x.file && /\.(png|jpg|jpeg|webp)$/i.test(String(x.name)));
  }

  if (pickStrategy === "lifestyleOnly") return (await hasImages(lifestyleId)) ? lifestyleId : null;
  if (pickStrategy === "catalogOnly") return (await hasImages(catalogId)) ? catalogId : null;

  // preferLifestyle (default)
  if (await hasImages(lifestyleId)) return lifestyleId;
  if (await hasImages(catalogId)) return catalogId;
  return null;
}

async function downloadDriveItemContent(accessToken, driveId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
  const res = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: "arraybuffer"
  });
  return Buffer.from(res.data);
}

async function uploadToFolder(accessToken, driveId, folderItemId, fileName, contentBuf) {
  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderItemId}:/${encodeURIComponent(
    fileName
  )}:/content`;

  const res = await axios.put(url, contentBuf, {
    headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "image/png" }
  });

  return {
    savedToSharePoint: true,
    driveId,
    itemId: res.data?.id,
    webUrl: res.data?.webUrl
  };
}

// ====== Composition (same as generate) ======
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
  const bg = await sharp(productBuf).resize(width, height, { fit: "cover" }).toBuffer();

  const gradientSvg = makeGradientOverlaySvg(width, height);

  const maxLogoW = Math.round((logoMaxWidthPct / 100) * width);
  const logoPng = await sharp(logoBuf).resize({ width: maxLogoW, withoutEnlargement: true }).png().toBuffer();
  const logoMeta = await sharp(logoPng).metadata();
  const logoW = logoMeta.width || maxLogoW;
  const logoH = logoMeta.height || Math.round(maxLogoW * 0.4);

  const logoLeft = width - safePaddingPx - logoW;
  const logoTop = height - safePaddingPx - logoH;

  const textSvg = makeTextSvg({ width, height, headline, subhead, cta, safePaddingPx, textAlign });

  return await sharp(bg)
    .composite([
      { input: Buffer.from(gradientSvg), top: 0, left: 0 },
      { input: Buffer.from(textSvg), top: 0, left: 0 },
      { input: logoPng, top: logoTop, left: logoLeft }
    ])
    .png()
    .toBuffer();
}

function makeGradientOverlaySvg(width, height) {
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

function makeTextSvg({ width, height, headline, subhead, cta, safePaddingPx, textAlign }) {
  const anchor = textAlign === "center" ? "middle" : "start";
  const textX = textAlign === "center" ? Math.round(width / 2) : safePaddingPx;

  return `
<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">
  <style>
    .h1 { font-family: Cochin, serif; font-weight: 600; font-size: 72px; fill: rgba(255,255,255,0.98); }
    .h2 { font-family: Cochin, serif; font-weight: 400; font-size: 40px; fill: rgba(255,255,255,0.92); }
    .h3 { font-family: Cochin, serif; font-weight: 600; font-size: 36px; fill: rgba(255,255,255,0.98); }
  </style>
  ${headline ? `<text x="${textX}" y="${safePaddingPx + 80}" text-anchor="${anchor}" class="h1">${escapeXml(headline)}</text>` : ""}
  ${subhead ? `<text x="${textX}" y="${safePaddingPx + 140}" text-anchor="${anchor}" class="h2">${escapeXml(subhead)}</text>` : ""}
  ${cta ? `<text x="${textX}" y="${height - safePaddingPx - 30}" text-anchor="${anchor}" class="h3">${escapeXml(cta)}</text>` : ""}
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

function sanitizeFileName(name) {
  return String(name).replace(/[\\/:*?"<>|]+/g, "_");
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// ====== Responses ======
function bad(context, message) {
  context.res = { status: 400, headers: { "Content-Type": "application/json" }, body: { error: { code: "bad_request", message } } };
}
function serverErr(context, message) {
  context.res = { status: 500, headers: { "Content-Type": "application/json" }, body: { error: { code: "server_error", message } } };
}