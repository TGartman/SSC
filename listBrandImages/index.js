// ✅ WORKING BASELINE — DO NOT BREAK
const axios = require("axios");

const DRIVE_ID = process.env.PRODUCT_IMAGES_DRIVE_ID;

// Brand root folder itemIds (top-level folders under /product-images)
const BRAND_ROOT = {
  SSC: "01POEEX7TJKWFWMFT2RFBKU63QDRMVT5JR",
  MT: "01POEEX7TBJQTUB7USFVC2EYSQCJVC7BZF",
  DWELL: "01POEEX7THJEFWFLPDIFFZK6CESICDKTKE",
};

async function getToken() {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET in local.settings.json");
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  const resp = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  return resp.data.access_token;
}

async function graphGet(token, path) {
  const url = `https://graph.microsoft.com/v1.0${path}`;
  const resp = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return resp.data;
}

async function findChildFolderByName(token, parentItemId, folderName) {
  // Lists children and returns the folder item that matches folderName (case-sensitive match to be safe).
  const data = await graphGet(
    token,
    `/drives/${DRIVE_ID}/items/${parentItemId}/children?$select=id,name,folder,file&$top=999`
  );
  const match = (data.value || []).find((x) => x.folder && x.name === folderName);
  return match || null;
}

async function listFiles(token, folderItemId, maxResults) {
  const data = await graphGet(
    token,
    `/drives/${DRIVE_ID}/items/${folderItemId}/children?$select=id,name,webUrl,file,size&$top=${Math.min(
      maxResults || 50,
      200
    )}`
  );

  const files = (data.value || [])
    .filter((x) => x.file && x.file.mimeType && x.file.mimeType.startsWith("image/"))
    .map((x) => ({
      driveId: DRIVE_ID,
      itemId: x.id,
      name: x.name,
      webUrl: x.webUrl,
      mimeType: x.file?.mimeType || null,
      sizeBytes: x.size || null,
    }));

  return files;
}

module.exports = async function (context, req) {
  try {
    const { brand, category, productLine, folderType, maxResults } = req.body || {};

    if (!brand || !BRAND_ROOT[brand]) {
      context.res = { status: 400, body: { error: { code: "bad_request", message: "brand must be SSC, MT, or DWELL" } } };
      return;
    }
    if (!category || !["exterior", "interior"].includes(category)) {
      context.res = { status: 400, body: { error: { code: "bad_request", message: "category must be exterior or interior" } } };
      return;
    }
    if (!productLine || typeof productLine !== "string") {
      context.res = { status: 400, body: { error: { code: "bad_request", message: "productLine is required" } } };
      return;
    }
    if (!folderType || !["Catalog Images", "Lifestyle Images"].includes(folderType)) {
      context.res = { status: 400, body: { error: { code: "bad_request", message: "folderType must be 'Catalog Images' or 'Lifestyle Images'" } } };
      return;
    }
    if (!DRIVE_ID) {
      context.res = { status: 500, body: { error: { code: "server_error", message: "Missing PRODUCT_IMAGES_DRIVE_ID env var" } } };
      return;
    }

    const token = await getToken();

    // Walk: Brand root -> (Exterior Products|Interior Products) -> productLine -> folderType
    const brandRootId = BRAND_ROOT[brand];
    const categoryFolderName = category === "exterior" ? "Exterior Products" : "Interior Products";

    const categoryFolder = await findChildFolderByName(token, brandRootId, categoryFolderName);
    if (!categoryFolder) {
      context.res = { status: 200, body: { images: [], note: `No folder found: ${categoryFolderName}` } };
      return;
    }

    const productLineFolder = await findChildFolderByName(token, categoryFolder.id, productLine);
    if (!productLineFolder) {
      context.res = { status: 200, body: { images: [], note: `No productLine folder found: ${productLine}` } };
      return;
    }

    const imagesFolder = await findChildFolderByName(token, productLineFolder.id, folderType);
    if (!imagesFolder) {
      context.res = { status: 200, body: { images: [], note: `No folderType folder found: ${folderType}` } };
      return;
    }

    const images = await listFiles(token, imagesFolder.id, maxResults);

    context.res = {
      status: 200,
      body: { images },
    };
  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: { error: { code: "server_error", message: err.message || "Unknown error" } },
    };
  }
};