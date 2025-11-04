import { DefaultAzureCredential } from "@azure/identity";
import fetch from "node-fetch";

const credential = new DefaultAzureCredential();

async function test() {
  try {
    const token = await credential.getToken("https://storage.azure.com/");
    console.log("✅ Token fetched successfully!");
    console.log("Access Token (truncated):", token.token.substring(0, 40) + "...");
  } catch (err) {
    console.error("❌ Failed to get token:", err.message);
  }
}

test();
