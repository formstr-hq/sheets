import { afterEach, describe, expect, it, vi } from "vitest";
import { hasNostrExtension, watchNostrExtension } from "./extension";

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

function browser() {
  const target = new EventTarget();
  const mock = Object.assign(target, {
    nostr: undefined as unknown,
    setInterval,
    clearInterval,
  });
  vi.stubGlobal("window", mock);
  return mock;
}

const provider = () => ({ getPublicKey: vi.fn(), signEvent: vi.fn() });

describe("browser extension detection", () => {
  it("requires a usable NIP-07 provider", () => {
    const mock = browser();
    expect(hasNostrExtension()).toBe(false);
    mock.nostr = {};
    expect(hasNostrExtension()).toBe(false);
    mock.nostr = provider();
    expect(hasNostrExtension()).toBe(true);
  });

  it("detects late injection and provider removal while watching", () => {
    vi.useFakeTimers();
    const mock = browser();
    const changed = vi.fn();
    const stop = watchNostrExtension(changed);
    expect(changed).toHaveBeenLastCalledWith(false);
    mock.nostr = provider();
    vi.advanceTimersByTime(500);
    expect(changed).toHaveBeenLastCalledWith(true);
    mock.nostr = undefined;
    vi.advanceTimersByTime(500);
    expect(changed).toHaveBeenLastCalledWith(false);
    stop();
    expect(vi.getTimerCount()).toBe(0);
  });

  it("rechecks on returning from an extension popup and cleans up", () => {
    vi.useFakeTimers();
    const mock = browser();
    const changed = vi.fn();
    const stop = watchNostrExtension(changed);
    mock.nostr = provider();
    mock.dispatchEvent(new Event("focus"));
    expect(changed).toHaveBeenLastCalledWith(true);
    stop();
    changed.mockClear();
    mock.dispatchEvent(new Event("focus"));
    vi.advanceTimersByTime(1000);
    expect(changed).not.toHaveBeenCalled();
  });
});
