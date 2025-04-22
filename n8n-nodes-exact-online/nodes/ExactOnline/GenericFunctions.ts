import { OptionsWithUri } from 'request';
import restApiConfig from './ExactOnlineRestApi.json';
import xmlApiConfig from './ExactOnlineXmlApi.json';

import {
	IExecuteFunctions,
	IExecuteSingleFunctions,
	IHookFunctions,
	ILoadOptionsFunctions,
} from 'n8n-core';

import { IDataObject, IOAuth2Options, NodeApiError, NodeOperationError } from 'n8n-workflow';
import {
	EndpointConfiguration,
	EndpointFieldConfiguration,
	LoadedDivision,
	LoadedFields,
	LoadedOptions,
	MatchSet,
} from './types';

// Cached merged configuration
let mergedConfigs: EndpointConfiguration[] | null = null;

/**
 * Loads and merges configurations from both REST and XML JSON files.
 */
function getAllConfigs(): EndpointConfiguration[] {
	if (mergedConfigs) {
		return mergedConfigs;
	}

	// Ensure configs are treated as arrays of the correct type
	const restEndpoints = restApiConfig as unknown as EndpointConfiguration[];
	const xmlEndpoints = xmlApiConfig as unknown as EndpointConfiguration[];

	// Add apiType to each source array
	const typedRestEndpoints = restEndpoints.map(config => ({
		...config,
		apiType: 'rest' as 'rest', // Explicitly type as 'rest'
	}));
	const typedXmlEndpoints = xmlEndpoints.map(config => ({
		...config,
		apiType: 'xml' as 'xml', // Explicitly type as 'xml'
	}));

	// Concatenate the typed arrays
	mergedConfigs = [...typedRestEndpoints, ...typedXmlEndpoints];

	return mergedConfigs;
}

export async function exactOnlineApiRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	method: string,
	uri: string,
	body: IDataObject = {},
	qs: IDataObject = {},
	option: IDataObject = {},
	nextPageUrl = '',
	// tslint:disable-next-line:no-any
): Promise<any> {
	let options: OptionsWithUri = {
		headers: {
			'Content-Type': 'application/json',
		},
		method,
		body,
		qs,
		uri: ``, // URI will be set based on credentials
		json: true,
	};
	// @ts-ignore
	options.resolveWithFullResponse = true;

	const authenticationMethod = this.getNodeParameter(
		'authentication',
		0,
		'oAuth2', // Default to OAuth2 as it's more common
	) as string;

	let credentialType: string;
	let baseUrl: string;

	if (authenticationMethod === 'accessToken') {
		const credentials = await this.getCredentials('exactOnlineApi');
		credentialType = 'exactOnlineApi';
		baseUrl = String(credentials.url) || 'https://start.exactonline.nl';
	} else {
		const credentials = await this.getCredentials('exactOnlineApiOAuth2Api');
		credentialType = 'exactOnlineApiOAuth2Api';
		baseUrl = String(credentials.url) || 'https://start.exactonline.nl';
	}

	options.uri = `${baseUrl}${uri}`;

	if(nextPageUrl!==''){
		options.uri = nextPageUrl;
	}
	options = Object.assign({}, options, option);

	try {
		if (Object.keys(body).length === 0) {
			delete options.body;
		}

		const oAuth2Options: IOAuth2Options = {
			includeCredentialsOnRefreshOnBody: true,
		};

		let response;
		try {
			response = await this.helpers.requestWithAuthentication.call(this, credentialType, options);
		} catch (error) {
			console.warn('[ExactNode] Error: ' + error.httpCode + ' | ' + error?.response?.status);
			if (error.httpCode && parseInt(error.httpCode, 10) === 429) {
				console.warn('[ExactNode] Detected 429: waiting 60 seconds.');
				await new Promise((resolve) => setTimeout(resolve, 61000)); // Wait for 60 seconds before retrying
				console.warn('[ExactNode] Waiting done.');
				response = await this.helpers.requestWithAuthentication.call(this, credentialType, options);
			} else {
				console.warn('[ExactNode] Unknown error', error);
				throw error;
			}
		}
		return response;
	} catch (error) {
		throw new NodeApiError(this.getNode(), error);
	}
}

export async function getCurrentDivision(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
): Promise<string> {
	const responseData = await exactOnlineApiRequest.call(
		this,
		'GET',
		`/api/v1/current/Me?$select=CurrentDivision`,
	);
	return responseData.body.d.results[0].CurrentDivision;
}

export async function getData(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	uri: string,
	body: IDataObject = {},
	qs: IDataObject = {},
	option: IDataObject = {},
): Promise<IDataObject[]> {
	const responseData = await exactOnlineApiRequest.call(this, 'GET', uri, body, qs, option);
	if (responseData.body.d.results) {
		return [].concat(responseData.body.d.results);
	} else {
		return [].concat(responseData.body.d);
	}
}

export async function getAllData(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	uri: string,
	limit = 60,
	body: IDataObject = {},
	qs: IDataObject = {},
	option: IDataObject = {},
	ignoreRateLimit = false,
): Promise<IDataObject[]> {
	let returnData: IDataObject[] = [];
	let responseData;
	let nextPageUrl = '';
	do {
		if (nextPageUrl === '') {
			responseData = await exactOnlineApiRequest.call(this, 'GET', uri, body, qs, option);
		} else {
			responseData = await exactOnlineApiRequest.call(
				this,
				'GET',
				uri,
				body,
				{},
				option,
				nextPageUrl,
			);
		}

		if (responseData.body.d.results) {
			returnData = returnData.concat(responseData.body.d.results);
		} else {
			returnData = returnData.concat(responseData.body.d);
		}
		nextPageUrl = responseData.body.d.__next;

		// Check if headers exist before accessing rate limit info
		if (!ignoreRateLimit && responseData.headers && responseData.headers['x-ratelimit-minutely-remaining'] === "0") {
			// Ensure reset header exists and is a number before calculating waitTime
			const resetTimestamp = responseData.headers['x-ratelimit-minutely-reset'];
			if (resetTimestamp && typeof resetTimestamp === 'number') {
				const waitTime = resetTimestamp - Date.now();
				if (waitTime >= 0) {
					await new Promise((resolve) => setTimeout(resolve, Math.min(waitTime, 60000)));
				}
			} else if (resetTimestamp) {
				console.warn(`[ExactNode Rate Limit] Invalid x-ratelimit-minutely-reset header type: ${typeof resetTimestamp}. Value: ${resetTimestamp}`);
			}
		}

	} while ((limit === 0 || returnData.length < limit) && responseData.body.d.__next);
	if (limit !== 0) {
		return returnData.slice(0, limit);
	}
	return returnData;
}
export async function getFields(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	endpointConfig: EndpointConfiguration,
): Promise<string[]> {
	return endpointConfig.fields.map((a) => a.name);
}

export async function getMandatoryFields(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	endpointConfig: EndpointConfiguration,
): Promise<string[]> {
	return endpointConfig.fields.filter((x) => x.mandatory === true).map((a) => a.name);
}

export async function getServiceOptions(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
) {
	// Use merged config
	const allConfigs = getAllConfigs();
	return allConfigs.map((x) => x.service.toLocaleLowerCase());
}
export async function getFieldType(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	endpointConfig: EndpointConfiguration,
	fieldName: string,
): Promise<string> {
	return endpointConfig.fields.filter((a) => a.name === fieldName)[0].type ?? 'Edm.String';
}

export async function getResourceOptions(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	service: string,
) {
	// Use merged config
	const allConfigs = getAllConfigs();
	return allConfigs.filter((x) => x.service.toLocaleLowerCase() === service).map((x) => x.endpoint);
}

export async function getEndpointFieldConfig(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	service: string,
	endpoint: string,
) {
	// Use merged config, await the promise
	const endpointConfig = await getEndpointConfig.call(this, service, endpoint);
	return endpointConfig?.fields || []; // Return fields or empty array if not found
}

export async function getEndpointConfig(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	service: string,
	endpoint: string,
): Promise<EndpointConfiguration | undefined> { // Return type can be undefined
	// Use merged config
	const allConfigs = getAllConfigs();
	return allConfigs.find(
		(x) => x.service.toLocaleLowerCase() === service.toLocaleLowerCase() && x.endpoint === endpoint,
	);
}

export const toDivisionOptions = (items: LoadedDivision[]) =>
	items.map(({ Code, CustomerName, Description }) => ({
		name: `${CustomerName} : ${Description}`,
		value: Code,
	}));

export const toOptions = (items: LoadedOptions[]) =>
	items.map(({ value, name }) => ({ name, value }));

export const toFieldSelectOptions = (items: LoadedFields[]) =>
	items.map(({ name }) => ({ name, value: name }));

export const toFieldFilterOptions = (items: EndpointFieldConfiguration[]) =>
	items.map(({ name }) => ({ name, value: name }));

export const toOptionsFromStringArray = (items: string[]) =>
	items.map((x) => ({ name: x.charAt(0).toUpperCase() + x.slice(1), value: x }));

/**
 * Makes an XML request to the Exact Online XML API
 */
export async function exactOnlineXmlRequest(
	this: IExecuteFunctions | IExecuteSingleFunctions | ILoadOptionsFunctions | IHookFunctions,
	division: string,
	topic: string,
	xmlBody: string,
	// tslint:disable-next-line:no-any
): Promise<any> {
	const options: OptionsWithUri = {
		headers: {
			'Content-Type': 'application/xml',
		},
		method: 'POST',
		body: xmlBody,
		uri: '', // URI will be set based on credentials
		// @ts-ignore
		resolveWithFullResponse: true,
	};

	const authenticationMethod = this.getNodeParameter(
		'authentication',
		0,
		'oAuth2',
	) as string;

	let credentialType: string;
	let baseUrl: string;

	if (authenticationMethod === 'accessToken') {
		const credentials = await this.getCredentials('exactOnlineApi');
		credentialType = 'exactOnlineApi';
		baseUrl = String(credentials.url) || 'https://start.exactonline.nl';
	} else {
		const credentials = await this.getCredentials('exactOnlineApiOAuth2Api');
		credentialType = 'exactOnlineApiOAuth2Api';
		baseUrl = String(credentials.url) || 'https://start.exactonline.nl';
	}

	const cleanBaseUrl = baseUrl.endsWith('/') ? baseUrl.slice(0, -1) : baseUrl;
	options.uri = `${cleanBaseUrl}/docs/XMLUpload.aspx?Topic=${topic}&_Division_=${division}`;

	try {
		const oAuth2Options: IOAuth2Options = {
			includeCredentialsOnRefreshOnBody: true,
		};

		let response;
		try {
			response = await this.helpers.requestWithAuthentication.call(this, credentialType, options);
		} catch (error) {
			console.warn('[ExactNode] Error: ' + error.httpCode + ' | ' + error?.response?.status);
			if (error.httpCode && parseInt(error.httpCode, 10) === 429) {
				console.warn('[ExactNode] Detected 429: waiting 60 seconds.');
				await new Promise((resolve) => setTimeout(resolve, 61000)); // Wait for 60 seconds before retrying
				console.warn('[ExactNode] Waiting done.');
				response = await this.helpers.requestWithAuthentication.call(this, credentialType, options);
			} else {
				console.warn('[ExactNode] Unknown error', error);
				throw error;
			}
		}
		return response;
	} catch (error) {
		throw new NodeApiError(this.getNode(), error);
	}
}

/**
 * Creates XML for reconciliation based on the full eExact XSD schema.
 */
export function createReconciliationXml(matchSets: MatchSet[]): string {
	let xml = '<?xml version="1.0" encoding="utf-8"?>\n';
	xml += '<eExact xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n';
	xml += '  <MatchSets>\n';

	for (const matchSet of matchSets) {
		xml += '    <MatchSet>\n';
		xml += `      <GLAccount code="${matchSet.GLAccount}"/>\n`;

		if (matchSet.Account) {
			xml += `      <Account code="${matchSet.Account}"/>\n`;
		}

		xml += '      <MatchLines>\n';
		for (const matchLine of matchSet.MatchLines) {
			const finYearInt = parseInt(String(matchLine.finYear), 10);
			const finPeriodInt = parseInt(String(matchLine.finPeriod), 10);
			const entryInt = parseInt(String(matchLine.entry), 10);
			const amountFloat = parseFloat(String(matchLine.amountDC));

			xml += '        <MatchLine ';
			xml += `finyear="${finYearInt}" `;
			xml += `finperiod="${finPeriodInt}" `;
			xml += `journal="${matchLine.journal}" `;
			xml += `entry="${entryInt}" `;
			xml += `amountdc="${amountFloat}" `;
			xml += '/>\n';
		}
		xml += '      </MatchLines>\n';

		if (matchSet.WriteOff) {
			xml += `      <WriteOff type="${matchSet.WriteOff.type}">\n`;

			if (matchSet.WriteOff.GLAccount) {
				xml += `        <GLAccount code="${matchSet.WriteOff.GLAccount}"/>\n`;
			}

			if (matchSet.WriteOff.Description) {
				xml += `        <Description>${matchSet.WriteOff.Description}</Description>\n`;
			}
			if (matchSet.WriteOff.FinYear) {
				xml += `        <FinYear>${parseInt(String(matchSet.WriteOff.FinYear), 10)}</FinYear>\n`;
			}
			if (matchSet.WriteOff.FinPeriod) {
				xml += `        <FinPeriod>${parseInt(String(matchSet.WriteOff.FinPeriod), 10)}</FinPeriod>\n`;
			}
			if (matchSet.WriteOff.Date) {
				xml += `        <Date>${matchSet.WriteOff.Date}</Date>\n`;
			}
			if (matchSet.WriteOff.VATCorrection !== undefined) {
				xml += `        <VATCorrection>${matchSet.WriteOff.VATCorrection}</VATCorrection>\n`;
			}

			xml += '      </WriteOff>\n';
		}

		xml += '    </MatchSet>\n';
	}

	xml += '  </MatchSets>\n';
	xml += '</eExact>';
	return xml;
}
