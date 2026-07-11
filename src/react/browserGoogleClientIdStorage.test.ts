import { describe, expect, it } from "vitest";
import {
  GOOGLE_CLIENT_ID_STORAGE_KEY,
  createBrowserGoogleClientIdStorage
} from "./browserGoogleClientIdStorage";

const CLIENT_ID = "123-abc.apps.googleusercontent.com";

class MemoryStorage implements Storage {
  private readonly values = new Map<string, string>();

  get length() {
    return this.values.size;
  }

  clear() {
    this.values.clear();
  }

  getItem(key: string) {
    return this.values.get(key) ?? null;
  }

  key(index: number) {
    return [...this.values.keys()][index] ?? null;
  }

  removeItem(key: string) {
    this.values.delete(key);
  }

  setItem(key: string, value: string) {
    this.values.set(key, value);
  }
}

describe("browser Google client ID storage", () => {
  it("loads, saves, and clears only the normalized public client id", async () => {
    const storage = new MemoryStorage();
    const adapter = createBrowserGoogleClientIdStorage(storage);

    await adapter.save(`  ${CLIENT_ID}  `);

    expect(storage.getItem(GOOGLE_CLIENT_ID_STORAGE_KEY)).toBe(CLIENT_ID);
    expect(storage.length).toBe(1);
    await expect(adapter.load()).resolves.toBe(CLIENT_ID);

    await adapter.clear();

    await expect(adapter.load()).resolves.toBeNull();
    expect(storage.length).toBe(0);
  });

  it("rejects non-public values without writing them", async () => {
    const storage = new MemoryStorage();
    const adapter = createBrowserGoogleClientIdStorage(storage);

    await expect(adapter.save("client-secret-value")).rejects.toMatchObject({
      code: "invalid_client"
    });
    expect(storage.length).toBe(0);
  });

  it("ignores an invalid value already present in browser storage", async () => {
    const storage = new MemoryStorage();
    storage.setItem(GOOGLE_CLIENT_ID_STORAGE_KEY, "not-a-public-client-id");

    await expect(createBrowserGoogleClientIdStorage(storage).load()).resolves.toBeNull();
  });
});
