/* eslint-disable */
import { IExecuteFunctions } from 'n8n-core';
import {
	IDataObject,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodeProperties,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';
import {
	createReconciliationXml,
	exactOnlineApiRequest,
	exactOnlineXmlRequest,
	getAllData,
	getCurrentDivision,
	getData,
	getEndpointConfig,
	getEndpointFieldConfig,
	getFields,
	getFieldType,
	getMandatoryFields,
	getResourceOptions,
	getServiceOptions,
	toDivisionOptions,
	toFieldFilterOptions,
	toFieldSelectOptions,
	toOptions,
	toOptionsFromStringArray,
} from './GenericFunctions';
import {
	EndpointConfiguration,
	EndpointFieldConfiguration,
	LoadedDivision,
	LoadedFields,
	LoadedOptions,
	MatchSet,
	ReconciledTransaction,
	WriteOff,
} from './types';

export class ExactOnline implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Exact Online',
		name: 'exactOnline',
		group: ['transform'],
		icon: 'file:exactOnline.svg',
		version: 1,
		description: 'Exact Online API node',
		defaults: {
			name: 'Exact Online',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'exactOnlineApi',
				required: true,
				displayOptions: {
					show: {
						authentication: ['accessToken'],
					},
				},
			},
			{
				name: 'exactOnlineApiOAuth2Api',
				required: true,
				displayOptions: {
					show: {
						authentication: ['oAuth2'],
					},
				},
			},
		],
		properties: [
			{
				displayName: 'Authentication',
				name: 'authentication',
				type: 'options',
				options: [
					{
						name: 'Access Token',
						value: 'accessToken',
					},
					{
						name: 'OAuth2',
						value: 'oAuth2',
					},
				],
				default: 'oAuth2',
			},
			{
				displayName: 'Division Name or ID',
				name: 'division',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getDivisions',
				},
				default: '',
				description:
					'Division to get data from. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
			},
			{
				displayName: 'Service Name or ID',
				name: 'service',
				type: 'options',
				typeOptions: {
					loadOptionsDependsOn: ['division'],
					loadOptionsMethod: 'getServices',
				},
				default: '',
				description:
					'Service to connecto to. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
			},
			{
				displayName: 'Resource Name or ID',
				name: 'resource',
				type: 'options',
				typeOptions: {
					loadOptionsDependsOn: ['service'],
					loadOptionsMethod: 'getResources',
				},
				default: '',
				description:
					'Resource to connect to. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
			},
			{
				displayName: 'Operation Name or ID',
				name: 'operation',
				type: 'options',
				typeOptions: {
					loadOptionsDependsOn: ['resource'],
					loadOptionsMethod: 'getOperations',
				},
				default: '',
				description:
					'Operation to use. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
			},
			{
				displayName: 'ID',
				name: 'id',
				type: 'string',
				default: '',
				description: 'ID of record',
				displayOptions: {
					show: {
						operation: ['get', 'put', 'delete'],
					},
				},
			},
			{
				displayName: 'Parent ID',
				name: 'parentId',
				type: 'string',
				default: '',
				required: true,
				description:
					'The GUID ID of the parent resource (e.g., BankEntry ID when using BankEntryLines)',
				displayOptions: {
					show: {
						operation: ['getAllViaParentId'],
					},
				},
			},
			{
				displayName: 'Limit',
				name: 'limit',
				type: 'number',
				typeOptions: {
					minValue: 1,
				},
				default: 60,
				description: 'Max number of results to return',
				displayOptions: {
					show: {
						operation: ['getAll', 'getAllViaParentId'],
					},
				},
			},
			{
				displayName: 'Selected Fields Are Excluded',
				name: 'excludeSelection',
				type: 'boolean',
				default: false,
				description:
					'Whether the selected fields are excluded instead of included. Select nothing to retrieve all fields.',
				displayOptions: {
					show: {
						operation: ['getAll', 'getAllViaParentId'],
					},
				},
			},
			{
				displayName: 'Disable waiting for minutely rate limit',
				name: 'ignoreRateLimit',
				type: 'boolean',
				default: false,
				description:
					'When set to true, the node will not wait for the minutely rate limit to reset which will result in 429 errors when exceeding the rate-limit.',
				displayOptions: {
					show: {
						operation: ['getAll', 'getAllViaParentId'],
					},
				},
			},
			{
				displayName: 'Fields to Get',
				name: 'selectedFields',
				type: 'multiOptions',
				typeOptions: {
					loadOptionsDependsOn: ['service', 'resource', 'operation'],
					loadOptionsMethod: 'getFields',
				},
				default: [],
				description:
					'Fields to retrieve from Exact Online. Choose from the list, or specify IDs using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
				displayOptions: {
					show: {
						operation: ['getAll', 'getAllViaParentId'],
					},
				},
			},
			{
				displayName: 'Conjunction',
				name: 'conjunction',
				type: 'options',
				options: [
					{
						name: 'And',
						value: 'and',
					},
					{
						name: 'Or',
						value: 'or',
					},
				],
				default: 'and',
				description: 'Conjunction to use in filter',
				displayOptions: {
					show: {
						operation: ['getAll'],
					},
				},
			},
			{
				displayName: 'Filter',
				name: 'filter',
				placeholder: 'Add filter',
				type: 'fixedCollection',
				typeOptions: {
					loadOptionsDependsOn: ['service', 'resource', 'operation'],
					multipleValues: true,
					sortable: true,
				},
				default: {},
				displayOptions: {
					show: {
						operation: ['getAll', 'getAllViaParentId'],
					},
				},
				options: [
					{
						name: 'filter',
						displayName: 'Filter',
						values: [
							{
								displayName: 'Field Name or ID',
								name: 'field',
								type: 'options',
								typeOptions: {
									loadOptionsMethod: 'getFieldsFilter',
								},
								default: '',
								description:
									'Field name to filter. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
							},
							{
								displayName: 'Operator',
								name: 'operator',
								type: 'options',
								options: [
									{
										name: 'Equal',
										value: 'eq',
									},
									{
										name: 'Greater Than',
										value: 'gt',
									},
									{
										name: 'Greater than or Equal',
										value: 'ge',
									},
									{
										name: 'Less Than',
										value: 'lt',
									},
									{
										name: 'Less than or Equal',
										value: 'le',
									},
									{
										name: 'Not Equal',
										value: 'ne',
									},
									{
										name: 'Contains',
										value: 'contains',
									},
								],
								default: 'eq',
								description: 'Operator to use in filter',
							},
							{
								displayName: 'Value',
								name: 'value',
								type: 'string',
								default: '',
								description:
									"Value to apply in the filter. For the 'Equal' operator, you can pass an array of strings (e.g., using an expression `{{ $json.myArray }}`) to filter for multiple values (behaves like SQL's IN).",
							},
						],
					},
				],
			},
			{
				displayName: 'Manual JSON Body',
				name: 'useManualBody',
				type: 'boolean',
				default: false,
				displayOptions: {
					show: {
						operation: ['post', 'put'],
					},
				},
			},
			{
				displayName: 'JSON Body',
				name: 'manualBody',
				type: 'string',
				default: '',
				displayOptions: {
					show: {
						operation: ['post', 'put'],
						useManualBody: [true],
					},
				},
			},
			{
				displayName: 'Field Data',
				name: 'data',
				placeholder: 'Add field data',
				type: 'fixedCollection',
				typeOptions: {
					loadOptionsDependsOn: ['service', 'resource', 'operation'],
					multipleValues: true,
					sortable: true,
				},
				default: {},
				displayOptions: {
					show: {
						operation: ['post', 'put'],
						useManualBody: [false],
					},
				},
				options: [
					{
						name: 'field',
						displayName: 'Field',
						values: [
							{
								displayName: 'Field Name or ID',
								name: 'fieldName',
								type: 'options',
								typeOptions: {
									loadOptionsMethod: 'getFieldsData',
								},
								default: '',
								description:
									'Field name to include in item. Choose from the list, or specify an ID using an expression.',
							},
							{
								displayName: 'Field Value',
								name: 'fieldValue',
								type: 'string',
								default: '',
								description: 'Value for the field to add/edit',
							},
						],
					},
				],
			},
		],
	};

	methods = {
		loadOptions: {
			async getDivisions(this: ILoadOptionsFunctions) {
				const currentDivision = await getCurrentDivision.call(this);
				const divisions = await exactOnlineApiRequest.call(
					this,
					'GET',
					`/api/v1/${currentDivision}/system/Divisions`,
				);

				return toDivisionOptions(divisions.body.d.results as LoadedDivision[]);
			},

			async getServices(this: ILoadOptionsFunctions) {
				const services = (await getServiceOptions.call(this)) as string[];

				return toOptionsFromStringArray([...new Set(services)]);
			},

			async getResources(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resources = (await getResourceOptions.call(this, service)) as string[];

				return toOptionsFromStringArray(resources);
			},

			async getOperations(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;

								const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as EndpointConfiguration;

				// If endpointConfig is not found, return empty options
				if (!endpointConfig) {
					return [];
				}

				const methods = endpointConfig.methods.map((x: string) => x.toLowerCase());
				if (methods.includes('get')) {
					methods.push('getAll');
				}

				// Conditionally add 'getAllViaParentId' if parentResource is defined
				if (endpointConfig.parentResource) {
					methods.push('getAllViaParentId');
				}

				return toOptionsFromStringArray(methods);
			},

			async getFields(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as EndpointConfiguration;
				const fields = await getFields.call(this, endpointConfig);
				return toFieldSelectOptions(fields.map((x: string) => ({ name: x })) as LoadedFields[]);
			},

			async getFieldsFilter(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as EndpointConfiguration;
				const fields = endpointConfig.fields;

				return toFieldFilterOptions(fields as EndpointFieldConfiguration[]);
			},

			async getFieldsData(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as EndpointConfiguration;
				//exclude auto generated values, these cannot be set by the user.
				const exclude = [
					'Created',
					'Creator',
					'CreatorFullName',
					'Modified',
					'Modifier',
					'ModifierFullName',
				];

				const fields = endpointConfig.fields.filter(
					(x: EndpointFieldConfiguration) => !exclude.includes(x.name),
				) as EndpointFieldConfiguration[];

				return toFieldFilterOptions(fields as EndpointFieldConfiguration[]);
			},
		},
	};

	// The function below is responsible for actually doing whatever this node
	// is supposed to do. In this case, we're just appending the `myString` property
	// with whatever the user has entered.
	// You can make async calls and use `await`.
	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		let returnData: IDataObject[] = [];
		const length = items.length;

		let responseData;
		const division = this.getNodeParameter('division', 0, '') as string;
		const service = this.getNodeParameter('service', 0, '') as string;
		const resource = this.getNodeParameter('resource', 0, '') as string;
		const operation = this.getNodeParameter('operation', 0, '') as string;
		const endpointConfig = await getEndpointConfig.call(this, service, resource); // Returns EndpointConfiguration | undefined

		if (!endpointConfig) {
			throw new NodeOperationError(
				this.getNode(),
				`Configuration not found for service '${service}' and resource '${resource}'.`,
				{ itemIndex: 0 },
			);
		}

		// Determine API type and URI based on the loaded config
		const apiType = endpointConfig.apiType; // Will be 'rest' or 'xml'
		const uri = endpointConfig.uri.replace('{division}', division);

		// REST specific fields (only needed for some REST ops)
		const excludeSelection = this.getNodeParameter('excludeSelection', 0, false) as boolean;
		const selectedFields = this.getNodeParameter('selectedFields', 0, []) as string[];
		let onlyNotSelectedFields: string[] = [];
		if (apiType === 'rest' && excludeSelection) {
			const allFields = await getFields.call(this, endpointConfig);
			onlyNotSelectedFields = allFields.filter((x) => !selectedFields.includes(x));
		}

		for (let itemIndex = 0; itemIndex < length; itemIndex++) {
			try {
				// --- Branch based on API Type --- //
				if (apiType === 'xml') {
					// --- XML Logic --- //
					if (operation === 'post') {
						let xmlBody = '';
						const useManualBodyXml = this.getNodeParameter('useManualBody', itemIndex, false) as boolean;

						if (useManualBodyXml) {
							// Manual body handling (assumes the user provides the full, correct XML string)
							xmlBody = this.getNodeParameter('manualBody', itemIndex, '') as string;
							if (!xmlBody) {
								 throw new NodeOperationError(this.getNode(), 'Manual XML Body cannot be empty when Manual JSON Body is enabled.', { itemIndex });
							}
						} else {
							// Use generic 'data.field' based on endpoint config
							const fieldDefinitions = endpointConfig.fields || [];
							const dataFields = this.getNodeParameter('data.field', itemIndex, []) as IDataObject[];

							// For FFMatch, the config defines 'MatchSets' as the primary field.
							// Find the first (and likely only) field definition from the config.
							// This assumes XML POST operations primarily deal with one complex structure field.
							const mainFieldDef = fieldDefinitions.length > 0 ? fieldDefinitions[0] : null;

							if (!mainFieldDef) {
								throw new NodeOperationError(this.getNode(), `Configuration error: No field definition found for XML endpoint '${endpointConfig.endpoint}'.`, { itemIndex });
							}

							const mainFieldName = mainFieldDef.name;
							const mainFieldData = dataFields.find(df => df.fieldName === mainFieldName);

							if (!mainFieldData || !mainFieldData.fieldValue) {
								if (mainFieldDef.mandatory) {
									throw new NodeOperationError(this.getNode(), `The mandatory '${mainFieldName}' field is missing or empty in Field Data. It should contain the data as a JSON string.`, { itemIndex });
								} else {
									// If the field isn't mandatory and not provided, what should happen?
									// For FFMatch, MatchSets IS mandatory. Assume mandatory if not specified? Or error.
									throw new NodeOperationError(this.getNode(), `The field '${mainFieldName}' is missing or empty in Field Data. It should contain the data as a JSON string.`, { itemIndex });
								}
							}

							// Specific handling for FFMatch based on the field name 'MatchSets'
							// If other XML endpoints are added, this might need a switch or factory pattern
							if (mainFieldName === 'MatchSets') {
								let matchSets: MatchSet[];
								try {
									// Assume the user provides the complex structure as a JSON string in the fieldValue
									const parsedFieldValue = JSON.parse(mainFieldData.fieldValue as string);
									// Basic validation based on config type 'Array[Object]'
									if (!Array.isArray(parsedFieldValue)) {
										throw new Error(`'${mainFieldName}' field value must be a JSON array string.`);
									}
									// TODO: Deeper validation against the structure defined in JSON config?
									matchSets = parsedFieldValue as MatchSet[];
								} catch (e) {
									throw new NodeOperationError(this.getNode(), `Invalid JSON provided in '${mainFieldName}' field value: ${e.message}`, { itemIndex });
								}

								if (matchSets.length === 0) {
									throw new NodeOperationError(this.getNode(), `No valid data found in the parsed '${mainFieldName}' field to process.`, { itemIndex });
								}

								// This function remains specific to FFMatch's structure, but the data feeding it is now generic.
								xmlBody = createReconciliationXml(matchSets);
							} else {
								// Handle future XML endpoints/field names here if needed
								throw new NodeOperationError(this.getNode(), `XML construction logic not implemented for field '${mainFieldName}'.`, { itemIndex });
							}
						}

						// Generic XML Request Part
						console.log(`Generated XML for ${endpointConfig.endpoint}:\n${xmlBody}`); // Consider removing in production
						try {
							// Use endpoint name (topic) from config
							const response = await exactOnlineXmlRequest.call(this, division, endpointConfig.endpoint, xmlBody);
							if (response.statusCode === 200) {
								returnData.push({ success: true, message: `XML operation '${endpointConfig.endpoint}' completed successfully`, response: response.body });
							} else {
								// Attempt to parse Exact Online XML error message if possible
								let errorMessage = response.body;
								try {
									// Basic check for Exact's error structure
									if (typeof response.body === 'string' && response.body.includes('<error>')) {
										const errorMatch = response.body.match(/<description>(.*?)<\/description>/);
										if (errorMatch && errorMatch[1]) {
											errorMessage = errorMatch[1];
										}
									}
								} catch (parseError) { /* Ignore parsing error, use raw body */ }
								throw new NodeOperationError(this.getNode(), `XML operation '${endpointConfig.endpoint}' failed: ${response.statusCode} - ${errorMessage}`, { itemIndex });
							}
						} catch (error) {
							// Handle NodeOperationErrors thrown above, or other request errors
							if (error instanceof NodeOperationError) throw error;
							throw new NodeOperationError(this.getNode(), `XML operation '${endpointConfig.endpoint}' failed: ${error.message}`, { itemIndex });
						}
					} else {
						// Handle other potential XML operations if added later
						throw new NodeOperationError(this.getNode(), `Operation '${operation}' not supported for XML endpoint '${endpointConfig.endpoint}'.`, { itemIndex });
					}
				} else {
					// --- REST API Logic --- //
					switch (operation) {
						case 'get': {
							const qs: IDataObject = {};
							const id = this.getNodeParameter('id', itemIndex, '') as string;
							if (id !== '') {
								qs['$filter'] = `ID eq guid'${id}'`;
								qs['$top'] = 1;
								responseData = await getData.call(this, uri, {}, qs);
								returnData = returnData.concat(responseData);
							}
							break;
						}
						case 'getAll': {
							const qs: IDataObject = {};
							const limit = this.getNodeParameter('limit', itemIndex, 0) as number;
							const conjunction = this.getNodeParameter('conjunction', itemIndex, 'and') as string;
							const filter = this.getNodeParameter('filter.filter', itemIndex, 0) as IDataObject[];
							const ignoreRateLimit = this.getNodeParameter('ignoreRateLimit', 0, false) as boolean;

							if (excludeSelection) {
								qs['$select'] = onlyNotSelectedFields.join(',');
							} else if (selectedFields.length > 0) {
								qs['$select'] = selectedFields.join(',');
							}
							const filters = [];
							if (filter.length > 0) {
								for (let filterIndex = 0; filterIndex < filter.length; filterIndex++) {
									const fieldName = filter[filterIndex].field as string;
									const operator = filter[filterIndex].operator as string;
									const rawValue = filter[filterIndex].value; // Get the raw value first
									const fieldType = await getFieldType.call(this, endpointConfig, fieldName);

									let filterSegment = '';

									// Special handling for 'eq' operator and array values for specific types
									if (operator === 'eq' && Array.isArray(rawValue)) {
										if (rawValue.length === 0) {
											// Handle empty array case - maybe filter for nothing? Or throw error?
											// Option: filter for a condition that's always false
											filterSegment = '1 eq 0';
											console.warn(
												`Filter for ${fieldName}: Received empty array for 'eq' operator. Resulting filter will likely return no results.`,
											);
										} else {
											switch (fieldType) {
												case 'Edm.Guid':
													filterSegment = rawValue
														.map((id) => `${fieldName} eq guid'${id}'`)
														.join(' or ');
													break;
												case 'Edm.String':
													filterSegment = rawValue
														.map((val) => `${fieldName} eq '${val}'`)
														.join(' or ');
													break;
												// Add cases for Edm.Int32, etc. if needed
												// case 'Edm.Int32': case 'Edm.Int64': ...
												//     filterSegment = rawValue.map(val => `${fieldName} eq ${val}`).join(' or ');
												//     break;
												default: // Use only first value as fallback
													// Default numeric/boolean assumption
													// Fallback for unsupported types with array + eq - treat as single value? Or error?
													console.warn(
														`Filter for ${fieldName}: Array value provided for 'eq' operator on unsupported type ${fieldType}. Treating as single value.`,
													);
													const firstValueStr = String(rawValue[0]);
													// Proceed as if it wasn't an array using firstValueStr
													// This part duplicates the logic below - needs refactoring or clearer error
													if (fieldType === 'Edm.Guid') {
														filterSegment = `${fieldName} eq guid'${firstValueStr}'`;
													} else if (fieldType === 'Edm.String') {
														filterSegment = `${fieldName} eq '${firstValueStr}'`;
													}
													// ... handle other types for the single value ...
													else filterSegment = `${fieldName} eq ${firstValueStr}`;
													break;
											}
										}
									} else {
										// Handle non-array values or operators other than 'eq'
										const valueStr = String(rawValue); // Ensure it's a string for standard processing
										switch (fieldType) {
											case 'Edm.String':
												const opStr =
													operator === 'contains'
														? `substringof('${valueStr}', ${fieldName})`
														: `${fieldName} ${operator} '${valueStr}'`;
												filterSegment = opStr;
												break;
											case 'Edm.Guid':
												// Only 'eq' and 'ne' typically make sense for GUIDs with single values
												if (operator === 'eq' || operator === 'ne') {
													filterSegment = `${fieldName} ${operator} guid'${valueStr}'`;
												} else {
													console.warn(
														`Unsupported operator '${operator}' for Guid field '${fieldName}'. Using 'eq'.`,
													);
													filterSegment = `${fieldName} eq guid'${valueStr}'`;
												}
												break;
											case 'Edm.DateTime':
												filterSegment = `${fieldName} ${operator} datetime'${valueStr}'`;
												break;
											case 'Edm.Boolean':
												filterSegment = `${fieldName} ${operator} ${
													valueStr.toLowerCase() === 'true'
												}`;
												break;
											case 'Edm.Int16':
											case 'Edm.Int32':
											case 'Edm.Int64':
											case 'Edm.Double':
											case 'Edm.Decimal':
											case 'Edm.Byte':
												filterSegment = `${fieldName} ${operator} ${valueStr}`;
												break;
											default: // Best guess
												console.warn(
													`Filter for ${fieldName}: Unsupported field type ${fieldType}. Attempting basic filter.`,
												);
												filterSegment = `${fieldName} ${operator} '${valueStr}'`;
												break;
										}
									}

									if (filterSegment) {
										filters.push(filterSegment);
									}
								}
								if (filters.length > 0) qs['$filter'] = filters.join(` ${conjunction} `);
							}
							responseData = await getAllData.call(this, uri, limit, {}, qs, {}, ignoreRateLimit);
							returnData = returnData.concat(responseData);
							break;
						}
						case 'getAllViaParentId': {
							const parentId = this.getNodeParameter('parentId', itemIndex, '') as string;
							const limit = this.getNodeParameter('limit', itemIndex, 60) as number;
							const ignoreRateLimit = this.getNodeParameter('ignoreRateLimit', 0, false) as boolean;
							if (!endpointConfig.parentResource) {
								throw new NodeOperationError(
									this.getNode(),
									`Operation 'getAllViaParentId' is not supported for resource '${resource}' as it lacks a defined parent resource in the configuration.`,
									{ itemIndex },
								);
							}
							if (parentId === '') {
								throw new NodeOperationError(
									this.getNode(),
									'Parent ID is required for the getAllViaParentId operation.',
									{ itemIndex },
								);
							}
							const parentResource = endpointConfig.parentResource;
							const specificUri = `/api/v1/${division}/${service}/${parentResource}(guid'${parentId}')/${resource}`;
							responseData = await getAllData.call(
								this,
								specificUri,
								limit,
								{},
								{},
								{},
								ignoreRateLimit,
							);
							returnData = returnData.concat(responseData);
							break;
						}
						case 'post': {
							let body: IDataObject = {};
							const useManualBodyRest = this.getNodeParameter(
								'useManualBody',
								itemIndex,
								false,
							) as boolean;
							if (useManualBodyRest) {
								const manualBody = this.getNodeParameter(
									'manualBody',
									itemIndex,
									{},
								) as IDataObject;
								if (!manualBody) {
									throw new NodeOperationError(
										this.getNode(),
										`Please include the fields and values for the item you want to create.`,
										{ itemIndex },
									);
								}
								body = manualBody;
							} else {
								const data = this.getNodeParameter('data.field', itemIndex, 0) as IDataObject[];
								if (!data) {
									throw new NodeOperationError(
										this.getNode(),
										`Please include the fields and values for the item you want to create.`,
										{ itemIndex },
									);
								}
								const fieldsEntered = data.map((x) => x.fieldName);
								const mandatoryFields = (await getMandatoryFields.call(
									this,
									endpointConfig,
								)) as string[];
								const mandatoryFieldsNotIncluded = mandatoryFields.filter(
									(x) => !fieldsEntered.includes(x),
								);
								if (mandatoryFieldsNotIncluded.length > 0) {
									throw new NodeOperationError(
										this.getNode(),
										`The following fields are mandatory and did not get used: '${mandatoryFieldsNotIncluded.join(
											', ',
										)}'`,
										{ itemIndex },
									);
								}
								if (data.length > 0) {
									for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
										const fieldName = data[dataIndex].fieldName as string;
										const fieldType = await getFieldType.call(this, endpointConfig, fieldName);
										const fieldValue = data[dataIndex].fieldValue as string;
										switch (fieldType) {
											case 'Edm.String':
											case 'Edm.Guid':
											case 'Edm.DateTime':
												body[`${fieldName}`] = fieldValue;
												break;
											case 'Edm.Boolean':
												body[`${fieldName}`] = fieldValue.toLocaleLowerCase() === 'true';
												break;
											case 'Edm.Int16':
											case 'Edm.Int32':
											case 'Edm.Int64':
											case 'Edm.Double':
											case 'Edm.Decimal':
											case 'Edm.Byte':
												body[`${fieldName}`] = +fieldValue;
												break;
											default:
												break;
										}
									}
								}
							}
							responseData = await exactOnlineApiRequest.call(
								this,
								'Post',
								uri,
								body,
								{},
								{ headers: { Prefer: 'return=representation' } },
							);
							returnData = returnData.concat(responseData.body.d);
							break;
						}
						case 'put': {
							const id = this.getNodeParameter('id', itemIndex, '') as string;
							if (id === '') {
								throw new NodeOperationError(
									this.getNode(),
									`Please enter an Id of a record to edit.`,
									{ itemIndex },
								);
							}
							let body: IDataObject = {};
							const useManualBodyRest = this.getNodeParameter(
								'useManualBody',
								itemIndex,
								false,
							) as boolean;
							if (useManualBodyRest) {
								const manualBody = this.getNodeParameter(
									'manualBody',
									itemIndex,
									{},
								) as IDataObject;
								if (!manualBody) {
									throw new NodeOperationError(
										this.getNode(),
										`Please include the fields and values for the item you want to edit.`,
										{ itemIndex },
									);
								}
								body = manualBody;
							} else {
								const data = this.getNodeParameter('data.field', itemIndex, 0) as IDataObject[];
								if (!data) {
									throw new NodeOperationError(
										this.getNode(),
										`Please include the fields and values for the item you want to edit.`,
										{ itemIndex },
									);
								}
								if (data.length > 0) {
									for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
										const fieldName = data[dataIndex].fieldName as string;
										const fieldType = await getFieldType.call(this, endpointConfig, fieldName);
										const fieldValue = data[dataIndex].fieldValue as string;
										switch (fieldType) {
											case 'Edm.String':
											case 'Edm.Guid':
											case 'Edm.DateTime':
												body[`${fieldName}`] = fieldValue;
												break;
											case 'Edm.Boolean':
												body[`${fieldName}`] = fieldValue.toLocaleLowerCase() === 'true';
												break;
											case 'Edm.Int16':
											case 'Edm.Int32':
											case 'Edm.Int64':
											case 'Edm.Double':
											case 'Edm.Decimal':
											case 'Edm.Byte':
												body[`${fieldName}`] = +fieldValue;
												break;
											default:
												break;
										}
									}
								}
							}
							const uriWithId = `${uri}(guid'${id}')`;
							responseData = await exactOnlineApiRequest.call(this, 'Put', uriWithId, body, {});
							if (responseData.statusCode === 204) {
								returnData = returnData.concat({ msg: 'Succesfully changed field values.' });
							} else {
								throw new NodeOperationError(this.getNode(), `Something went wrong.`, {
									itemIndex,
								});
							}
							break;
						}
						case 'delete': {
							const id = this.getNodeParameter('id', itemIndex, '') as string;
							if (id === '') {
								throw new NodeOperationError(
									this.getNode(),
									`Please enter an Id of a record to delete.`,
									{ itemIndex },
								);
							}
							const uriWithId = `${uri}(guid'${id}')`;
							responseData = await exactOnlineApiRequest.call(this, 'Delete', uriWithId, {}, {});
							if (responseData.statusCode === 204) {
								returnData = returnData.concat({ msg: 'Succesfully Deleted record.' });
							} else {
								throw new NodeOperationError(this.getNode(), `Something went wrong.`, {
									itemIndex,
								});
							}
							break;
						}
						default: {
							throw new NodeOperationError(
								this.getNode(),
								`Operation '${operation}' is not supported for REST API endpoint '${endpointConfig.endpoint}'.`,
								{ itemIndex },
							);
						}
					}
				}
			} catch (error) {
				// This node should never fail but we want to showcase how
				// to handle errors.
				if (this.continueOnFail()) {
					returnData.push({ error });
				} else {
					// Adding `itemIndex` allows other workflows to handle this error
					if (error.context) {
						// If the error thrown already contains the context property,
						// only append the itemIndex
						error.context.itemIndex = itemIndex;
						throw error;
					}
					throw new NodeOperationError(this.getNode(), error, {
						itemIndex,
					});
				}
			}
		}

		return [this.helpers.returnJsonArray(returnData)];
	}
}
