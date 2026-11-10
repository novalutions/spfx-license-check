/**
 * Compile-only check (npm run typecheck): `this.context` of a real SPFx component fits checkLicense. The SPFx base
 * context exposes exactly these two members with these types.
 */
import type { AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import type { PageContext } from '@microsoft/sp-page-context';
import { checkLicense, ILicenseHttpResponse, ILicenseResult } from '../src';

declare const context: { readonly aadHttpClientFactory: AadHttpClientFactory; readonly pageContext: PageContext };

export const result: Promise<ILicenseResult> = checkLicense(context, { offerId: 'employee-directory' });

declare const response: HttpClientResponse;
export const fits: ILicenseHttpResponse = response;
