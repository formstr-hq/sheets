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
import type { Event } from "nostr-tools";
import { fetchEvents } from "../lib/nostr";
import { loadProfileCache, saveProfileCache } from "../lib/local";
import type { Profile } from "../lib/types";
import { useAuth } from "./AuthProvider";

interface ProfileContextValue {
  profiles: Map<string, Profile>;
  ensureProfiles: (pubkeys: string[]) => Promise<void>;
  profileFor: (pubkey: string) => Profile | null;
  displayName: (pubkey: string) => string;
}

const ProfileContext = createContext<ProfileContextValue | undefined>(undefined);

function parseProfile(event: Event): Profile | null {
  try {
    const data = JSON.parse(event.content) as Record<string, unknown>;
    return {
      pubkey: event.pubkey,
      name: typeof data.name === "string" ? data.name : undefined,
      displayName:
        typeof data.display_name === "string"
          ? data.display_name
          : typeof data.name === "string"
            ? data.name
            : undefined,
      about: typeof data.about === "string" ? data.about : undefined,
      picture: typeof data.picture === "string" ? data.picture : undefined,
      nip05: typeof data.nip05 === "string" ? data.nip05 : undefined,
      lud16: typeof data.lud16 === "string" ? data.lud16 : undefined,
      updatedAt: event.created_at,
    };
  } catch {
    return null;
  }
}

export function ProfileProvider({ children }: { children: ReactNode }) {
  const { relays, pubkey, signer } = useAuth();
  const [profiles, setProfiles] = useState<Map<string, Profile>>(() => {
    const cached = loadProfileCache() as Record<string, Profile>;
    return new Map(Object.entries(cached));
  });
  const profilesRef = useRef(profiles);
  const requestedRef = useRef(new Set<string>());

  useEffect(() => {
    profilesRef.current = profiles;
  }, [profiles]);

  const ensureProfiles = useCallback(
    async (pubkeys: string[]) => {
      const missing = Array.from(new Set(pubkeys.filter(Boolean))).filter(
        (key) => !profilesRef.current.has(key) && !requestedRef.current.has(key),
      );
      if (missing.length === 0) return;
      missing.forEach((key) => requestedRef.current.add(key));

      const chunks: string[][] = [];
      for (let i = 0; i < missing.length; i += 100) chunks.push(missing.slice(i, i + 100));

      const updates = new Map<string, Profile>();
      await Promise.all(
        chunks.map(async (authors) => {
          const events = await fetchEvents(
            relays,
            { kinds: [0], authors, limit: authors.length * 2 },
            6000,
          );
          const latest = new Map<string, Event>();
          for (const event of events) {
            const current = latest.get(event.pubkey);
            if (!current || event.created_at > current.created_at) latest.set(event.pubkey, event);
          }
          for (const event of latest.values()) {
            const profile = parseProfile(event);
            if (profile) updates.set(profile.pubkey, profile);
          }
        }),
      );

      if (updates.size === 0) return;
      setProfiles((current) => {
        const next = new Map(current);
        for (const [key, value] of updates) next.set(key, value);
        const cache: Record<string, Profile> = {};
        for (const [key, value] of next) cache[key] = value;
        saveProfileCache(cache);
        return next;
      });
    },
    [relays],
  );

  useEffect(() => {
    if (pubkey && signer) void ensureProfiles([pubkey]);
  }, [pubkey, signer, ensureProfiles]);

  const profileFor = useCallback(
    (key: string) => profilesRef.current.get(key) ?? null,
    [],
  );

  const displayName = useCallback(
    (key: string) => {
      const profile = profilesRef.current.get(key);
      return profile?.displayName || profile?.name || key.slice(0, 8);
    },
    [],
  );

  const value = useMemo(
    () => ({ profiles, ensureProfiles, profileFor, displayName }),
    [profiles, ensureProfiles, profileFor, displayName],
  );

  return <ProfileContext.Provider value={value}>{children}</ProfileContext.Provider>;
}

export function useProfiles(): ProfileContextValue {
  const context = useContext(ProfileContext);
  if (!context) throw new Error("useProfiles must be used within ProfileProvider");
  return context;
}
