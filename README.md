# @novalutions/spfx-license-check

License check for novalutions SharePoint Framework (SPFx) products. One function asks the novalutions licensing
service for the license of the current Microsoft 365 tenant and decides whether the product locks. Cache, outage
tolerance and the lock rules are built in, so every product behaves the same.

- Calls `GET https://marketplace.novalutions.de/api/license/{offerId}` with the SPFx `AadHttpClient`: no own sign-in,
  no premium license, no secrets.
- Caches the answer per tenant in `localStorage` for one hour (5 minutes while locked).
- Keeps working with the last answer for up to 7 days when the service is unreachable. An outage never locks.
- Locks after the trial and when the SharePoint admin has not approved `License.Read`.
- ES5 output with type definitions, for SPFx 1.18 or later.

## Install

```bash
npm install @novalutions/spfx-license-check
```

`@microsoft/sp-http` (1.18 or later) is a peer dependency; every SPFx project already has it.

## Set up the web part

### 1. Request the API permission

In `config/package-solution.json`:

```json
"webApiPermissionRequests": [
  { "resource": "novalutions Marketplace Portal", "scope": "License.Read" }
]
```

SPFx matches the request by the display name of the novalutions app. The customer's SharePoint admin approves it
under SharePoint admin center > Advanced > API access. The app exists in the customer tenant once someone from the
tenant has signed in at `https://marketplace.novalutions.de` (for example on the landing page after a marketplace
purchase). If it is missing, admin consent creates it:
`https://login.microsoftonline.com/<customer-tenant-id>/adminconsent?client_id=eaf87e2c-ca7b-4331-83af-e11719c26184`.

### 2. Check the license

Pass the component's `this.context` and the product's offer ID in the licensing service:

```ts
import { checkLicense } from '@novalutions/spfx-license-check';

checkLicense(this.context, { offerId: 'employee-directory' }).then((license) => {
  if (license.locked) {
    // show the lock screen instead of the product
  }
});
```

The promise never rejects. Check the offer ID carefully: the service answers an unknown one with `404`, which counts
as an outage, so a typo never locks and only shows up as a `console.warn`.

Options:

| Option | Default | Meaning |
| --- | --- | --- |
| `offerId` | required | offer ID of the product in the licensing service |
| `includeTest` | `false` | also count test purchases (Partner Center preview, Stripe test mode); turn it on only in builds for testing |

With React, a small hook keeps the web part responsive while the first check runs:

```tsx
import * as React from 'react';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { checkLicense, ILicenseResult } from '@novalutions/spfx-license-check';

export function useLicense(context: BaseComponentContext): ILicenseResult | undefined {
  const [license, setLicense] = React.useState<ILicenseResult | undefined>(undefined);
  React.useEffect(() => {
    let cancelled = false;
    checkLicense(context, { offerId: 'employee-directory' }).then((result) => {
      if (!cancelled) setLicense(result);
    });
    return () => {
      cancelled = true;
    };
  }, [context]);
  return license; // undefined while the first check runs
}
```

Render the product while `license` is `undefined` or `locked` is `false`, and the lock screen when `locked` is `true`.

## Result

```ts
interface ILicenseResult {
  status: LicenseStatus;
  locked: boolean;
  seats?: number;   // users of all valid licenses together; show it, never enforce it
  termEnd?: string; // ISO timestamp of the latest term end, e.g. to show "trial until …"
}
```

| `status` | Meaning | `locked` | Suggested message |
| --- | --- | --- | --- |
| `active` | paid license (marketplace or direct) | no | none |
| `trial` | trial only | no | trial until `termEnd`, link to manage the subscription |
| `pending` | bought but not activated yet | yes | finish the activation |
| `suspended` | payment problem, Microsoft suspended the subscription | yes | contact your billing admin |
| `unsubscribed` | cancelled or expired | yes | renew the subscription |
| `none` | no license for this tenant | yes | buy or start a trial |
| `notApproved` | the SharePoint admin has not approved `License.Read` | yes | ask the admin to approve the API access |
| `unknown` | service unreachable and no answer from the last 7 days | no | none |

Customers manage their subscriptions at `LICENSE_MANAGE_URL` (`https://marketplace.novalutions.de/subscriptions`);
link it from the lock screen:

```ts
import { LICENSE_MANAGE_URL } from '@novalutions/spfx-license-check';
```

## Behaviour in detail

- **Cache**: one entry per offer, tenant and `includeTest` in `localStorage`, reused for one hour. A locked answer is
  reused for 5 minutes only, so a purchase or activation takes effect quickly. A missing approval is not cached at
  all. An answer saved "in the future" (the clock was set back) is not trusted.
- **Outage**: network errors, `5xx`, `429`, `401`, `404` and answers without a license status count as an outage.
  The last answer then stays valid for 7 days from when it was fetched, a locked one too, so an outage neither locks
  nor unlocks. Without such an answer the result is `unknown` and the product runs unlocked.
- **Approval**: a `403` from the service or a token error (`AADSTS65001`, `AADSTS500011`, consent) means the API
  access is not approved. The product locks, otherwise not approving would get around the check.
- **Newer service versions**: a status this version does not know follows the service's `licensed` verdict
  (`true` counts as `active`, `false` as `none`).
- **Test purchases** count only with `includeTest: true`, like in the licensing service. Leave it off in releases:
  anyone with a Stripe test-mode payment link could otherwise unlock the product with a test card.
- **Blocked storage** (private mode, sandboxed iframe): the check still works, it just asks the service every time.

## Development

```bash
npm ci
npm run typecheck   # source, tests and the fit against the real SPFx context types
npm test            # every test against the TypeScript source and against the compiled ES5 output
npm run build       # lib/ with ES5 JavaScript and type definitions, checked with es-check
```

The tests call `checkLicense` with a fake `aadHttpClientFactory`, an in-memory `localStorage` and a fake clock.

Releases to npm are made by novalutions (`npm publish --access public`; `prepack` builds and checks `lib/`).

## License

[MIT](LICENSE)
