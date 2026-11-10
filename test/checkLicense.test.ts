import { AadHttpClient } from '@microsoft/sp-http';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { checkLicense } from '../src';
import { FakeService, fakeService, memoryStorage, TENANT } from './fakes/spfx';

const OFFER = 'employee-directory';
const NOW = new Date('2026-11-10T09:00:00Z');
const MINUTE = 60 * 1000;
const DAY = 24 * 60 * MINUTE;

/** sets the clock to NOW plus `ms` (negative: before NOW) */
function afterStart(ms: number): void {
  vi.setSystemTime(new Date(NOW.getTime() + ms));
}

beforeEach(() => {
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(NOW);
  vi.stubGlobal('window', { localStorage: memoryStorage() });
  vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('checkLicense', () => {
  it('unlocks an active subscription and reports seats and term end', async () => {
    const service = fakeService();
    service.answer(200, { status: 'active', licensed: true, seats: 25, termEnd: '2027-10-15T00:00:00+00:00' });

    const result = await checkLicense(service.context, { offerId: OFFER });

    expect(result).toEqual({ status: 'active', locked: false, seats: 25, termEnd: '2027-10-15T00:00:00+00:00' });
    expect(service.requests).toEqual([
      {
        resource: 'api://eaf87e2c-ca7b-4331-83af-e11719c26184',
        url: 'https://marketplace.novalutions.de/api/license/employee-directory',
        configuration: AadHttpClient.configurations.v1
      }
    ]);
  });

  it.each([
    ['trial', true, false],
    ['pending', false, true],
    ['suspended', false, true],
    ['unsubscribed', false, true],
    ['none', false, true]
  ])('reports %s (licensed: %s) with locked: %s', async (status, licensed, locked) => {
    const service = fakeService();
    service.answer(200, { status, licensed });

    const result = await checkLicense(service.context, { offerId: OFFER });

    expect(result).toEqual({ status, locked });
  });

  it('locks when the service refuses the token because License.Read is not approved (403)', async () => {
    const service = fakeService();
    service.answer(403);

    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'notApproved', locked: true });
  });

  it.each([
    'AADSTS65001: The user or administrator has not consented to use the application',
    'AADSTS500011: The resource principal named api://eaf87e2c-ca7b-4331-83af-e11719c26184 was not found in the tenant'
  ])('locks when no token can be issued: %s', async (message) => {
    const service = fakeService();
    service.fail(new Error(message));

    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'notApproved', locked: true });
  });

  it('asks the service at most once per hour', async () => {
    const service = fakeService();
    service.answer(200, { status: 'active', licensed: true, seats: 25 });
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(200, { status: 'none', licensed: false });
    afterStart(59 * MINUTE);
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false, seats: 25 });
    expect(service.requests).toHaveLength(1);

    afterStart(61 * MINUTE);
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'none', locked: true });
    expect(service.requests).toHaveLength(2);
  });

  it('keeps the cache apart per tenant', async () => {
    const licensed = fakeService(TENANT);
    licensed.answer(200, { status: 'active', licensed: true });
    await checkLicense(licensed.context, { offerId: OFFER });

    const other = fakeService('22222222-2222-2222-2222-222222222222');
    other.answer(200, { status: 'none', licensed: false });
    expect(await checkLicense(other.context, { offerId: OFFER })).toEqual({ status: 'none', locked: true });
    expect(other.requests).toHaveLength(1);
  });

  it('asks again after 5 minutes while locked, so a purchase takes effect quickly', async () => {
    const service = fakeService();
    service.answer(200, { status: 'none', licensed: false });
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(200, { status: 'active', licensed: true });
    afterStart(4 * MINUTE);
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'none', locked: true });

    afterStart(6 * MINUTE);
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false });
    expect(service.requests).toHaveLength(2);
  });

  it('asks again on every call while License.Read is not approved, so an approval takes effect at once', async () => {
    const service = fakeService();
    service.answer(403);
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(200, { status: 'active', licensed: true });
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false });
    expect(service.requests).toHaveLength(2);
  });

  describe('when the service is unreachable', () => {
    it('keeps working with the last answer for up to 7 days', async () => {
      const service = fakeService();
      service.answer(200, { status: 'active', licensed: true, seats: 25 });
      await checkLicense(service.context, { offerId: OFFER });

      service.fail(new TypeError('Failed to fetch'));
      afterStart(6 * DAY);
      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false, seats: 25 });
    });

    it('keeps a locked last answer, the outage does not unlock it early', async () => {
      const service = fakeService();
      service.answer(200, { status: 'none', licensed: false });
      await checkLicense(service.context, { offerId: OFFER });

      service.fail(new TypeError('Failed to fetch'));
      afterStart(2 * DAY);
      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'none', locked: true });
    });

    it('never locks once the last answer is older than 7 days', async () => {
      const service = fakeService();
      service.answer(200, { status: 'none', licensed: false });
      await checkLicense(service.context, { offerId: OFFER });

      service.fail(new TypeError('Failed to fetch'));
      afterStart(8 * DAY);
      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'unknown', locked: false });
    });

    it.each([
      ['a network error', (service: FakeService) => service.fail(new TypeError('Failed to fetch'))],
      ['500', (service: FakeService) => service.answer(500, { error: 'boom' })],
      ['503', (service: FakeService) => service.answer(503)],
      ['429 (rate limit)', (service: FakeService) => service.answer(429)],
      ['401 (token rejected)', (service: FakeService) => service.answer(401)],
      ['404 (offer not configured)', (service: FakeService) => service.answer(404)],
      ['200 without JSON', (service: FakeService) => service.answer(200)]
    ])('never locks on %s without an earlier answer', async (_, breakService) => {
      const service = fakeService();
      breakService(service);

      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'unknown', locked: false });
    });
  });

  describe('with an answer this version does not know', () => {
    it.each([
      [{ status: 'grace', licensed: true, seats: 10 }, { status: 'active', locked: false, seats: 10 }],
      [{ status: 'revoked', licensed: false }, { status: 'none', locked: true }]
    ])('follows the service verdict "licensed" for a new status: %o', async (body, expected) => {
      const service = fakeService();
      service.answer(200, body);

      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual(expected);
    });

    it('ignores seats and term end of the wrong type', async () => {
      const service = fakeService();
      service.answer(200, { status: 'active', licensed: true, seats: '25', termEnd: 20271015 });

      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false });
    });

    it.each([null, 'Service Unavailable', { status: 'grace' }, { message: 'hello' }])(
      'treats %o as an outage and keeps the last answer',
      async (body) => {
        const service = fakeService();
        service.answer(200, { status: 'active', licensed: true });
        await checkLicense(service.context, { offerId: OFFER });

        service.answer(200, body);
        afterStart(2 * 60 * MINUTE);
        expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false });
      }
    );
  });

  it('never rejects, even when the SPFx context throws', async () => {
    const broken = {
      aadHttpClientFactory: {
        getClient(): never {
          throw new Error('serviceScope not ready');
        }
      },
      pageContext: {}
    };

    await expect(checkLicense(broken, { offerId: OFFER })).resolves.toEqual({ status: 'unknown', locked: false });
  });

  it('still works when the browser blocks localStorage', async () => {
    vi.stubGlobal('window', {
      get localStorage(): never {
        throw new DOMException('The operation is insecure.', 'SecurityError');
      }
    });
    const service = fakeService();
    service.answer(200, { status: 'trial', licensed: true, termEnd: '2026-12-10T00:00:00+00:00' });

    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({
      status: 'trial',
      locked: false,
      termEnd: '2026-12-10T00:00:00+00:00'
    });
  });

  it('counts test purchases only with includeTest, in a cache of their own', async () => {
    const service = fakeService();
    service.answer(200, { status: 'none', licensed: false });
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(200, { status: 'active', licensed: true });
    expect(await checkLicense(service.context, { offerId: OFFER, includeTest: true })).toEqual({
      status: 'active',
      locked: false
    });
    expect(service.requests.map((request) => request.url)).toEqual([
      'https://marketplace.novalutions.de/api/license/employee-directory',
      'https://marketplace.novalutions.de/api/license/employee-directory?includeTest=true'
    ]);
  });

  it('keeps the cache apart per offer', async () => {
    const service = fakeService();
    service.answer(200, { status: 'active', licensed: true });
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(200, { status: 'none', licensed: false });
    expect(await checkLicense(service.context, { offerId: 'mailbox-desk' })).toEqual({ status: 'none', locked: true });
    expect(service.requests).toHaveLength(2);
  });

  it('locks once the hour is over when the approval of License.Read was withdrawn', async () => {
    const service = fakeService();
    service.answer(200, { status: 'active', licensed: true });
    await checkLicense(service.context, { offerId: OFFER });

    service.answer(403);
    afterStart(61 * MINUTE);
    expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'notApproved', locked: true });
  });

  describe('when the clock was set back after the last answer', () => {
    it('asks the service again instead of trusting the answer forever', async () => {
      const service = fakeService();
      service.answer(200, { status: 'none', licensed: false });
      await checkLicense(service.context, { offerId: OFFER });

      service.answer(200, { status: 'active', licensed: true });
      afterStart(-2 * DAY);
      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'active', locked: false });
    });

    it('does not lock on an outage with that answer', async () => {
      const service = fakeService();
      service.answer(200, { status: 'none', licensed: false });
      await checkLicense(service.context, { offerId: OFFER });

      service.fail(new TypeError('Failed to fetch'));
      afterStart(-2 * DAY);
      expect(await checkLicense(service.context, { offerId: OFFER })).toEqual({ status: 'unknown', locked: false });
    });
  });
});
