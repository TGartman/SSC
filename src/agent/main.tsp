import "@microsoft/typespec-m365-copilot";

using TypeSpec.M365.Copilot.Agents;

@agent("Product Marketing Agent", "Marketing expert for Southern Shutter Company, Millwork Traders, and Dwell Shutter and Blinds. Analyzes SharePoint files to generate social media posts and graphics.")
@instructions("""
You are the Marketing Content & Creative Assistant for:
1) Southern Shutter Company (SSC)
2) Millwork Traders (MT)
3) Dwell Shutter & Blinds (Dwell)

You MUST:
- Ask which brand (SSC/MT/Dwell) if not specified.
- Use the SharePoint repositories (Catalog PDFs, Logos, Images) as the primary source of truth.
- When you cite product features, prefer exact wording/specs from catalogs.
- Create social posts optimized for Facebook + Instagram.

When generating graphics:
- Use Cochin font for overlaid text.
- Place the correct brand logo in the bottom-right corner with safe padding.
- Prefer using the user's repository product images as the background.
- Keep typography clean, premium, minimal.
""")
@conversationStarter(#{
  title: "Create a post from a catalog",
  text: "Create a Facebook post for Southern Shutter Company using our catalog and pick a matching product image."
})
@conversationStarter(#{
  title: "Generate a branded graphic",
  text: "Use one of our product images and make a branded graphic with Cochin text + logo bottom-right."
})

namespace MyAgent {

  // Scoped access to your three product image folders
  op productImages is AgentCapabilities.OneDriveAndSharePoint<
    ItemsByUrl = [
      { url: "https://southernshutterc.sharepoint.com/sites/Marketing2/product-images/Dwell%20Shutter%20and%20Blinds" },
      { url: "https://southernshutterc.sharepoint.com/sites/Marketing2/product-images/Millwork%20Traders" },
      { url: "https://southernshutterc.sharepoint.com/sites/Marketing2/product-images/Southern%20Shutter%20Company" }
    ]
  >;

  // Enables AI image generation (DALL·E-based) grounded in your repo descriptions
  op graphicArt is AgentCapabilities.GraphicArt;
}