export function normalizeNamedRangeLookup(name: string): string {
  return name.trim().normalize("NFKC").toLowerCase();
}
