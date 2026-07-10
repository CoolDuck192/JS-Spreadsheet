import type {
  GoogleSheetsServiceConfiguration,
  TokenProvider
} from "../core/workbook/services";

export type GoogleOAuthOriginAssessment =
  | { status: "eligible"; origin: string; registration: "unverified" }
  | {
      status: "blocked";
      origin: string;
      reason: "opaque" | "file" | "insecure" | "ip_literal";
    };

export type GoogleClientIdValidation =
  | { valid: true; value: string }
  | { valid: false; message: string };

export type GoogleAuthSource =
  | { kind: "provider"; provider: TokenProvider }
  | {
      kind: "client";
      clientId: string;
      source: "managed" | "stored" | "session";
      factory: (clientId: string) => TokenProvider;
    }
  | { kind: "missing" };

export type EditableGoogleClientId = Readonly<{
  value: string;
  source: "stored" | "session";
}>;

const PUBLIC_WEB_CLIENT_ID =
  /^[A-Za-z0-9][A-Za-z0-9_-]*\.apps\.googleusercontent\.com$/;

export function validateGoogleClientId(value: string): GoogleClientIdValidation {
  const normalized = value.trim();
  if (PUBLIC_WEB_CLIENT_ID.test(normalized)) {
    return { valid: true, value: normalized };
  }
  return {
    valid: false,
    message:
      "Enter a Google OAuth Web application client ID ending in .apps.googleusercontent.com."
  };
}

export function assessGoogleOAuthOrigin(origin: string): GoogleOAuthOriginAssessment {
  const candidate = origin.trim();
  if (!candidate || candidate === "null") {
    return { status: "blocked", origin: candidate || "null", reason: "opaque" };
  }

  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    return { status: "blocked", origin: candidate, reason: "opaque" };
  }

  if (parsed.protocol === "file:") {
    return { status: "blocked", origin: candidate, reason: "file" };
  }
  if (parsed.origin === "null") {
    return { status: "blocked", origin: candidate, reason: "opaque" };
  }

  const hostname = stripIpv6Brackets(parsed.hostname.toLowerCase());
  const loopback =
    hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
  if (isIpLiteral(hostname) && !loopback) {
    return { status: "blocked", origin: parsed.origin, reason: "ip_literal" };
  }
  if (parsed.protocol === "https:" || (parsed.protocol === "http:" && loopback)) {
    return { status: "eligible", origin: parsed.origin, registration: "unverified" };
  }
  return { status: "blocked", origin: parsed.origin, reason: "insecure" };
}

export function resolveGoogleAuthSource(
  configuration: GoogleSheetsServiceConfiguration | undefined,
  editableClientId: EditableGoogleClientId | null,
  deprecatedFactory: ((clientId: string) => TokenProvider) | undefined,
  builtInFactory: (clientId: string) => TokenProvider
): GoogleAuthSource {
  if (configuration?.tokenProvider) {
    return { kind: "provider", provider: configuration.tokenProvider };
  }

  const managedClientId = configuration?.clientId?.trim();
  const editableValue = editableClientId?.value.trim();
  const clientId = managedClientId || editableValue;
  if (!clientId) {
    return { kind: "missing" };
  }

  return {
    kind: "client",
    clientId,
    source: managedClientId ? "managed" : editableClientId!.source,
    factory: configuration?.tokenProviderFactory ?? deprecatedFactory ?? builtInFactory
  };
}

function stripIpv6Brackets(hostname: string): string {
  return hostname.startsWith("[") && hostname.endsWith("]")
    ? hostname.slice(1, -1)
    : hostname;
}

function isIpLiteral(hostname: string): boolean {
  if (hostname.includes(":")) {
    return true;
  }
  const octets = hostname.split(".");
  return (
    octets.length === 4 &&
    octets.every((octet) => /^\d{1,3}$/.test(octet) && Number(octet) <= 255)
  );
}
