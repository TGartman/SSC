const axios = require("axios");
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");

function getGraphClient() {
  const {
    TENANT_ID,
    CLIENT_ID,
    CLIENT_SECRET
  } = process.env;

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error("Missing TENANT_ID / CLIENT_ID / CLIENT_SECRET");
  }

  const credential = new ClientSecretCredential(
    TENANT_ID,
    CLIENT_ID,
    CLIENT_SECRET
  );

  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const token = await credential.getToken(
          "https://graph.microsoft.com/.default"
        );
        return token.token;
      }
    }
  });
}

module.exports = async function (context, req) {
  try {
    const body = req.body || {};

    const {
      brand,
      area = "Exterior Products",
      format,
      limit = 999,
      pickStrategy = "preferLifestyle",
      headlineTemplate = "{product}",
      subhead,
      cta,
      style = {},
      output = {}
    } = body;

    if (!brand || !format) {
      return context.res = {
        status: 400,
        body: {
          error: "brand and format are required"
        }
      };
    }

    const graph = getGraphClient();

    // 🔹 Brand → Drive root folders
    const brandRoots = {
      "SSC": "Southern Shutter Company",
      "MT": "Millwork Traders",
      "DWELL": "Dwell Shutter and Blinds"
    };

    const brandRootName = brandRoots[brand];
    if (!brandRootName) {
      throw new Error(`Unknown brand ${brand}`);
    }

    // 🔹 Locate brand root folder
    const rootChildren = await graph
      .api("/drives/" + process.env.PRODUCT_IMAGES_DRIVE_ID + "/root/children")
      .get();

    const brandRoot = rootChildren.value.find(f => f.name === brandRootName);
    if (!brandRoot) throw new Error("Brand root not found");

    // 🔹 Find Exterior / Interior Products
    const brandChildren = await graph
      .api(`/drives/${process.env.PRODUCT_IMAGES_DRIVE_ID}/items/${brandRoot.id}/children`)
      .get();

    const areaFolder = brandChildren.value.find(f => f.name === area);
    if (!areaFolder) throw new Error(`${area} folder not found`);

    // 🔹 List product folders
    const products = await graph
      .api(`/drives/${process.env.PRODUCT_IMAGES_DRIVE_ID}/items/${areaFolder.id}/children`)
      .get();

    const results = [];

    for (const product of products.value.slice(0, limit)) {
      const productName = product.name;

      // 🔹 Get product subfolders
      const productChildren = await graph
        .api(`/drives/${process.env.PRODUCT_IMAGES_DRIVE_ID}/items/${product.id}/children`)
        .get();

      let imageFolder =
        productChildren.value.find(f => f.name === "Lifestyle Images") ||
        productChildren.value.find(f => f.name === "Catalog Images");

      if (!imageFolder) {
        context.log(`Skipping ${productName} (no images)`);
        continue;
      }

      // 🔹 Get images
      const images = await graph
        .api(`/drives/${process.env.PRODUCT_IMAGES_DRIVE_ID}/items/${imageFolder.id}/children`)
        .get();

      if (!images.value.length) continue;

      const image = images.value[0]; // first image per product

      // 🔹 Call internal generate endpoint
const generatePayload = {
  brand,
  format,
  headline: headlineTemplate.replace("{product}", productName),
  subhead: subhead || "",
  cta: cta || "",
  productImage: {
    driveId: process.env.PRODUCT_IMAGES_DRIVE_ID,
    itemId: image.id
  },
  style: {
    logoPlacement: "bottom_right",
    safePaddingPx: 64,
    logoMaxWidthPct: 16,
    textAlign: "left",
    ...(style || {})
  },
  output: {
    returnBase64: false,
    saveToSharePoint: true,
    fileName: (output.fileNameTemplate || "{brand}_{product}_{format}.png")
      .replace("{brand}", brand)
      .replace("{product}", productName.replace(/\s+/g, "_"))
      .replace("{format}", format)
  }
};
 const response = await axios.post(
        "http://localhost:7071/graphics/generate",
        generatePayload,
        { headers: { "Content-Type": "application/json" } }
      );

      results.push({
        product: productName,
        saved: response.data.saved
      });
    }

    context.res = {
      status: 200,
      body: {
        brand,
        area,
        generated: results.length,
        results
      }
    };

  } catch (err) {
    context.log.error(err);
    context.res = {
      status: 500,
      body: {
        error: err.message
      }
    };
  }
};
