export function hasNostrExtension(): boolean {
  if (typeof window === "undefined") return false;
  const nostr = (window as unknown as {
    nostr?: { getPublicKey?: unknown; signEvent?: unknown };
  }).nostr;
  return typeof nostr?.getPublicKey === "function" && typeof nostr?.signEvent === "function";
}

// Extensions can inject their provider after React mounts, including after
// the user enables site access from the extension popup.
export function watchNostrExtension(onChange: (available: boolean) => void): () => void {
  const check = () => onChange(hasNostrExtension());
  check();
  const timer = window.setInterval(check, 500);
  window.addEventListener("focus", check);
  return () => {
    window.clearInterval(timer);
    window.removeEventListener("focus", check);
  };
}
