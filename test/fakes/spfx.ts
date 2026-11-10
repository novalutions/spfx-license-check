/**
 * Fake SPFx context: the part of a web part's `this.context` that checkLicense uses, with a scripted licensing
 * service behind `aadHttpClientFactory`. Only the system boundary is faked; checkLicense itself runs unchanged.
 */
export const TENANT = '11111111-1111-1111-1111-111111111111';

export interface RecordedRequest {
  resource: string;
  url: string;
  configuration: unknown;
}

type Behaviour = { kind: 'reply'; status: number; body: unknown } | { kind: 'error'; error: Error };

export function fakeService(tenantId: string = TENANT) {
  let behaviour: Behaviour = { kind: 'error', error: new Error('fake service: no answer scripted') };
  const requests: RecordedRequest[] = [];

  const context = {
    aadHttpClientFactory: {
      getClient(resource: string) {
        return Promise.resolve({
          get(url: string, configuration: unknown) {
            requests.push({ resource, url, configuration });
            if (behaviour.kind === 'error') {
              return Promise.reject(behaviour.error);
            }
            const { status, body } = behaviour;
            return Promise.resolve({
              status,
              ok: status >= 200 && status < 300,
              json: () =>
                body === undefined ? Promise.reject(new SyntaxError('Unexpected end of JSON input')) : Promise.resolve(body)
            });
          }
        });
      }
    },
    pageContext: { aadInfo: { tenantId: { toString: () => tenantId } } }
  };

  return {
    context,
    requests,
    /** the service answers every following request with this status and JSON body */
    answer(status: number, body?: unknown) {
      behaviour = { kind: 'reply', status, body };
    },
    /** every following request fails before an answer arrives (network or token error) */
    fail(error: Error) {
      behaviour = { kind: 'error', error };
    }
  };
}

/** In-memory Web Storage, installed as window.localStorage by the tests. */
export function memoryStorage() {
  const items: Record<string, string> = {};
  return {
    getItem: (key: string): string | null => (key in items ? items[key] : null),
    setItem: (key: string, value: string): void => {
      items[key] = String(value);
    }
  };
}

export type FakeService = ReturnType<typeof fakeService>;
