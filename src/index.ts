import { AadHttpClient, AadHttpClientConfiguration } from '@microsoft/sp-http';

const SERVICE_URL = 'https://marketplace.novalutions.de';
/** Where tenant admins activate and manage their subscriptions; link it from the lock screen. */
export const LICENSE_MANAGE_URL = SERVICE_URL + '/subscriptions';
/** App 1 "novalutions Marketplace Portal"; package-solution.json requests License.Read by that display name */
const RESOURCE = 'api://eaf87e2c-ca7b-4331-83af-e11719c26184';

const CACHE_PREFIX = 'novalutions.license.v1.';
/** ask the service at most once per hour */
const REFRESH_MS = 60 * 60 * 1000;
/** while locked, ask again after 5 minutes so an activation or purchase takes effect quickly */
const LOCKED_REFRESH_MS = 5 * 60 * 1000;
/** keep working with the last answer for 7 days while the service is unreachable */
const GRACE_MS = 7 * 24 * 60 * 60 * 1000;

/**
 * `active` … `none` are the licensing service's own values. `notApproved`: the SharePoint admin has not approved
 * the API access (License.Read). `unknown`: the service is unreachable and there is no usable earlier answer.
 */
export type LicenseStatus = ServiceStatus | 'notApproved' | 'unknown';

type ServiceStatus = 'active' | 'trial' | 'suspended' | 'pending' | 'unsubscribed' | 'none';

/**
 * Whether the product locks for each status of the service. Without a valid license the product locks: the trial is
 * the only free use.
 */
const LOCKED_BY_STATUS: { [status in ServiceStatus]: boolean } = {
  active: false,
  trial: false,
  suspended: true,
  pending: true,
  unsubscribed: true,
  none: true
};

export interface ILicenseCheckOptions {
  /** Offer ID in the licensing service, e.g. `employee-directory`. */
  offerId: string;
  /**
   * Also count test purchases (Partner Center preview, Stripe test mode). Off by default; turn it on only in builds
   * for testing.
   */
  includeTest?: boolean;
}

export interface ILicenseResult {
  status: LicenseStatus;
  /** True when the product must lock. */
  locked: boolean;
  /** Users of all valid licenses together. Shown only, never enforced. */
  seats?: number;
  /** ISO timestamp of the latest term end. */
  termEnd?: string;
}

/** The response members checkLicense reads; `HttpClientResponse` from `@microsoft/sp-http` fits. */
export interface ILicenseHttpResponse {
  status: number;
  ok: boolean;
  json(): Promise<unknown>;
}

/** The parts of an SPFx component context checkLicense uses; `this.context` of a web part or extension fits. */
export interface ILicenseCheckContext {
  aadHttpClientFactory: {
    getClient(resourceEndpoint: string): Promise<{
      get(url: string, configuration: AadHttpClientConfiguration): Promise<ILicenseHttpResponse>;
    }>;
  };
  pageContext: {
    aadInfo?: { tenantId?: { toString(): string } };
  };
}

interface ICacheEntry {
  savedAt: number;
  result: ILicenseResult;
}

interface ILicenseResponseBody {
  status?: unknown;
  /** the service's own verdict; decides when a newer service sends a status this version does not know */
  licensed?: unknown;
  seats?: unknown;
  termEnd?: unknown;
}

export function checkLicense(context: ILicenseCheckContext, options: ILicenseCheckOptions): Promise<ILicenseResult> {
  const key = CACHE_PREFIX + options.offerId + '.' + tenantId(context) + (options.includeTest ? '.test' : '');
  const cached = readCache(key);
  // a negative age means the clock was set back; such an answer is not trusted
  const age = cached ? Date.now() - cached.savedAt : -1;
  if (cached && age >= 0 && age < (cached.result.locked ? LOCKED_REFRESH_MS : REFRESH_MS)) {
    return Promise.resolve(cached.result);
  }

  return fetchLicense(context, options).then(
    (result) => {
      // a missing approval is not cached, so an approval takes effect on the next page load
      if (result.status !== 'notApproved') {
        writeCache(key, result);
      }
      return result;
    },
    (err: unknown): ILicenseResult => {
      console.warn('novalutions license check failed', err);
      if (cached && age >= 0 && age < GRACE_MS) {
        return cached.result;
      }
      return { status: 'unknown', locked: false };
    }
  );
}

function fetchLicense(context: ILicenseCheckContext, options: ILicenseCheckOptions): Promise<ILicenseResult> {
  const url =
    SERVICE_URL + '/api/license/' + encodeURIComponent(options.offerId) + (options.includeTest ? '?includeTest=true' : '');
  // starting inside then() turns a synchronous throw of the SPFx context into a rejection
  return Promise.resolve()
    .then(() => context.aadHttpClientFactory.getClient(RESOURCE))
    .then((client) => client.get(url, AadHttpClient.configurations.v1))
    .then((response): ILicenseResult | Promise<ILicenseResult> => {
      // 403: the token lacks License.Read, i.e. the API access request is not approved
      if (response.status === 403) {
        return notApproved();
      }
      if (!response.ok) {
        throw new Error('License service answered ' + response.status);
      }
      return response.json().then(toResult);
    })
    .then(undefined, (err: unknown): ILicenseResult => {
      if (isMissingApproval(err)) {
        return notApproved();
      }
      throw err;
    });
}

/** Throws for an answer that is not a license status, so it counts as an outage. */
function toResult(json: unknown): ILicenseResult {
  const body = (json !== null && typeof json === 'object' ? json : {}) as ILicenseResponseBody;
  let status: LicenseStatus;
  if (isServiceStatus(body.status)) {
    status = body.status;
  } else if (typeof body.licensed === 'boolean') {
    status = body.licensed ? 'active' : 'none';
  } else {
    throw new Error('License service sent no license status');
  }

  const result: ILicenseResult = { status, locked: isLocked(status) };
  if (typeof body.seats === 'number') {
    result.seats = body.seats;
  }
  if (typeof body.termEnd === 'string') {
    result.termEnd = body.termEnd;
  }
  return result;
}

function notApproved(): ILicenseResult {
  return { status: 'notApproved', locked: true };
}

/**
 * Missing admin approval shows up as a token error: AADSTS65001 (no consent for License.Read) or AADSTS500011
 * (App 1 does not exist in the customer tenant yet).
 */
function isMissingApproval(err: unknown): boolean {
  const message = err instanceof Error ? err.message : String(err);
  return /AADSTS65001|AADSTS500011|consent/i.test(message);
}

function isServiceStatus(value: unknown): value is ServiceStatus {
  return typeof value === 'string' && Object.prototype.hasOwnProperty.call(LOCKED_BY_STATUS, value);
}

/**
 * `notApproved` locks, otherwise not approving License.Read would get around the check. `unknown` never locks: an
 * outage at novalutions must not lock customers out.
 */
function isLocked(status: LicenseStatus): boolean {
  if (status === 'notApproved') {
    return true;
  }
  return status === 'unknown' ? false : LOCKED_BY_STATUS[status];
}

function tenantId(context: ILicenseCheckContext): string {
  try {
    const aad = context.pageContext.aadInfo;
    if (aad && aad.tenantId) {
      return aad.tenantId.toString();
    }
  } catch {
    // fall through to the shared key
  }
  return 'tenant';
}

function storage(): Storage | undefined {
  try {
    return typeof window !== 'undefined' && window.localStorage ? window.localStorage : undefined;
  } catch {
    return undefined; // blocked storage (private mode, sandboxed iframe)
  }
}

function readCache(key: string): ICacheEntry | undefined {
  const store = storage();
  if (!store) {
    return undefined;
  }
  try {
    const raw = store.getItem(key);
    if (!raw) {
      return undefined;
    }
    const entry = JSON.parse(raw) as ICacheEntry;
    if (typeof entry.savedAt !== 'number' || !entry.result || !isServiceStatus(entry.result.status)) {
      return undefined;
    }
    // the lock decision follows this version's rules, not the stored flag
    entry.result.locked = isLocked(entry.result.status);
    return entry;
  } catch {
    return undefined;
  }
}

function writeCache(key: string, result: ILicenseResult): void {
  const store = storage();
  if (!store) {
    return;
  }
  try {
    const entry: ICacheEntry = { savedAt: Date.now(), result };
    store.setItem(key, JSON.stringify(entry));
  } catch {
    // the cache is optional (quota, blocked storage)
  }
}
