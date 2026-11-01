import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { nip19, SimplePool } from "nostr-tools";
import { encryptSecretKey } from "@formstr/signer";
import type { Signer as FormstrSigner, StoredAccount } from "@formstr/signer";
import { DEFAULT_RELAYS } from "../lib/constants";
import { fetchUserRelays } from "../lib/nostr";
import { signerInstance } from "../lib/signerInstance";
import { setLocalRelayRelays } from "../lib/localRelay";
import { mergeRelays, toNpub } from "../lib/utils";

const NIP46_PERMS = ["get_public_key", "sign_event", "nip44_encrypt", "nip44_decrypt"];

export type ActiveSigner = NonNullable<ReturnType<FormstrSigner["getActiveSigner"]>>;

interface AuthContextValue {
  ready: boolean;
  account: StoredAccount | null;
  signer: ActiveSigner | null;
  pubkey: string | null;
  npub: string | null;
  locked: boolean;
  relays: string[];
  loginWithExtension: () => Promise<void>;
  createAccount: (passphrase: string) => Promise<{ npub: string; ncryptsec: string }>;
  loginWithNcryptsec: (ncryptsec: string, passphrase: string) => Promise<void>;
  importNsec: (nsec: string, passphrase: string) => Promise<void>;
  loginWithBunker: (uri: string) => Promise<void>;
  unlockWithPassphrase: (passphrase: string) => Promise<void>;
  reconnectRemote: () => Promise<void>;
  logout: () => Promise<void>;
  requireSigner: () => ActiveSigner;
  refreshRelays: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | undefined>(undefined);

const signerPool = new SimplePool();

export function AuthProvider({ children }: { children: ReactNode }) {
  const [ready, setReady] = useState(false);
  const [account, setAccount] = useState<StoredAccount | null>(null);
  const [signer, setSigner] = useState<ActiveSigner | null>(null);
  const [relays, setRelays] = useState<string[]>(DEFAULT_RELAYS);
  const relaysForRef = useRef<string | null>(null);

  const sync = useCallback(() => {
    const activeAccount = signerInstance.getActiveAccount();
    const activeSigner = signerInstance.getActiveSigner();
    setAccount(activeAccount);
    setSigner(activeSigner);
    return { activeAccount, activeSigner };
  }, []);

  useEffect(() => {
    let mounted = true;
    const unsubscribe = signerInstance.onChange(() => {
      if (mounted) sync();
    });

    signerInstance
      .unlock({ pool: signerPool })
      .catch(() => null)
      .finally(() => {
        if (!mounted) return;
        sync();
        setReady(true);
      });

    return () => {
      mounted = false;
      unsubscribe();
    };
  }, [sync]);

  const pubkey = account?.pubkey ?? null;
  const npub = pubkey ? account?.npub || toNpub(pubkey) : null;
  const locked = !!account && !signer;

  useEffect(() => {
    if (!pubkey || !signer) {
      setRelays(DEFAULT_RELAYS);
      setLocalRelayRelays(DEFAULT_RELAYS);
      relaysForRef.current = null;
      return;
    }
    if (relaysForRef.current === pubkey) return;
    relaysForRef.current = pubkey;
    let active = true;
    void fetchUserRelays(DEFAULT_RELAYS, pubkey)
      .then((userRelays) => {
        if (!active) return;
        const merged = mergeRelays(DEFAULT_RELAYS, userRelays);
        setRelays(merged);
        setLocalRelayRelays(merged);
      })
      .catch(() => {
        if (active) {
          setRelays(DEFAULT_RELAYS);
          setLocalRelayRelays(DEFAULT_RELAYS);
        }
      });
    return () => {
      active = false;
    };
  }, [pubkey, signer]);

  const loginWithExtension = useCallback(async () => {
    await signerInstance.loginWithExtension();
    sync();
  }, [sync]);

  const createAccount = useCallback(
    async (passphrase: string) => {
      const result = await signerInstance.createAccount(passphrase);
      sync();
      return result;
    },
    [sync],
  );

  const loginWithNcryptsec = useCallback(
    async (ncryptsec: string, passphrase: string) => {
      await signerInstance.loginWithNcryptsec(ncryptsec, passphrase);
      sync();
    },
    [sync],
  );

  const importNsec = useCallback(
    async (nsec: string, passphrase: string) => {
      const decoded = nip19.decode(nsec.trim());
      if (decoded.type !== "nsec") throw new Error("Enter a valid nsec private key.");
      const ncryptsec = encryptSecretKey(decoded.data, passphrase);
      await signerInstance.loginWithNcryptsec(ncryptsec, passphrase);
      sync();
    },
    [sync],
  );

  const loginWithBunker = useCallback(
    async (uri: string) => {
      await signerInstance.loginWithBunkerUri(uri, { pool: signerPool, perms: NIP46_PERMS });
      sync();
    },
    [sync],
  );

  const unlockWithPassphrase = useCallback(
    async (passphrase: string) => {
      const activeAccount = signerInstance.getActiveAccount();
      if (!activeAccount?.ncryptsec) throw new Error("This account can't be unlocked with a passphrase.");
      await signerInstance.loginWithNcryptsec(activeAccount.ncryptsec, passphrase);
      sync();
    },
    [sync],
  );

  const reconnectRemote = useCallback(async () => {
    const activeAccount = signerInstance.getActiveAccount();
    if (!activeAccount?.nip46?.uri) throw new Error("No remote signer is saved for this account.");
    await signerInstance.loginWithBunkerUri(activeAccount.nip46.uri, {
      pool: signerPool,
      perms: NIP46_PERMS,
      clientSecretKey: undefined,
    });
    sync();
  }, [sync]);

  const logout = useCallback(async () => {
    await signerInstance.logout();
    sync();
  }, [sync]);

  const requireSigner = useCallback(() => {
    const active = signerInstance.getActiveSigner();
    if (!active) throw new Error("Please sign in first.");
    return active;
  }, []);

  const refreshRelays = useCallback(async () => {
    if (!pubkey || !signer) {
      setRelays(DEFAULT_RELAYS);
      return;
    }
    const userRelays = await fetchUserRelays(DEFAULT_RELAYS, pubkey);
    const merged = mergeRelays(DEFAULT_RELAYS, userRelays);
    setRelays(merged);
    setLocalRelayRelays(merged);
  }, [pubkey, signer]);

  const value = useMemo<AuthContextValue>(
    () => ({
      ready,
      account,
      signer,
      pubkey,
      npub,
      locked,
      relays,
      loginWithExtension,
      createAccount,
      loginWithNcryptsec,
      importNsec,
      loginWithBunker,
      unlockWithPassphrase,
      reconnectRemote,
      logout,
      requireSigner,
      refreshRelays,
    }),
    [
      ready,
      account,
      signer,
      pubkey,
      npub,
      locked,
      relays,
      loginWithExtension,
      createAccount,
      loginWithNcryptsec,
      importNsec,
      loginWithBunker,
      unlockWithPassphrase,
      reconnectRemote,
      logout,
      requireSigner,
      refreshRelays,
    ],
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
}

export function useAuth(): AuthContextValue {
  const context = useContext(AuthContext);
  if (!context) throw new Error("useAuth must be used within AuthProvider");
  return context;
}
