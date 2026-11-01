import { createSigner } from "@formstr/signer";
import { APP_NAME, APP_URL } from "./constants";

export const signerInstance = createSigner({
  appName: APP_NAME,
  appUrl: typeof window !== "undefined" ? window.location.origin : APP_URL,
  storageKeyPrefix: "formstr-sheets:signer:",
});
