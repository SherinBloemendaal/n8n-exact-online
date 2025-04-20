/* eslint-disable */
import { IExecuteFunctions } from 'n8n-core';
import {
	IDataObject,
	ILoadOptionsFunctions,
	INodeExecutionData,
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
	endpointConfiguration,
	endpointFieldConfiguration,
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
				options: [
					{
						name: 'Delete',
						value: 'delete',
					},
					{
						name: 'Get',
						value: 'get',
					},
					{
						name: 'Get All',
						value: 'getAll',
					},
					{
						name: 'Post',
						value: 'post',
					},
					{
						name: 'POST (XML)',
						value: 'postXml',
					},
					{
						name: 'Put',
						value: 'put',
					},
					{
						name: 'Get All Via Parent ID',
						value: 'getAllViaParentId',
					},
				],
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
				description: 'The GUID ID of the parent resource (e.g., BankEntry ID)',
				displayOptions: {
					show: {
						resource: ['BankEntryLines'],
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
						operation: ['getAll'],
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
						operation: ['getAll'],
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
						operation: ['getAll'],
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
								description: 'Value to apply in the filter',
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
						operation: ['post', 'postXml', 'put'],
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
						operation: ['post', 'postXml', 'put'],
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
						operation: ['post', 'postXml', 'put'],
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
									'Field name to include in item. Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>.',
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
			{
				displayName: 'Reconciliation',
				name: 'reconciliation',
				placeholder: 'Setup reconciliation',
				type: 'fixedCollection',
				default: {},
				displayOptions: {
					show: {
						operation: ['postXml'],
						service: ['financial'],
						resource: ['MatchSets'],
					},
				},
				description:
					'Configure automatic reconciliation between transactions in Exact Online. This uses the XML API since the REST API does not support automatic reconciliation. Specify the GL account and customer account to match, and optionally configure write-off parameters.',
				options: [
					{
						name: 'matchSet',
						displayName: 'Match Set',
						values: [
							{
								displayName: 'GL Account Code',
								name: 'glAccountCode',
								type: 'string',
								default: '',
								required: true,
								description: 'GL Account code for reconciliation',
							},
							{
								displayName: 'Account Code',
								name: 'accountCode',
								type: 'string',
								default: '',
								required: false,
								description:
									'Account code for reconciliation. Required if GL Account is of type Accounts receivable or Accounts payable.',
							},
							{
								displayName: 'Include Write-Off',
								name: 'includeWriteOff',
								type: 'boolean',
								default: false,
								description: 'Whether to include write-off information',
							},
							{
								displayName: 'Write-Off Type',
								name: 'writeOffType',
								type: 'string',
								default: '4',
								description: 'Type of write-off',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
							{
								displayName: 'Write-Off GL Account',
								name: 'writeOffGLAccount',
								type: 'string',
								default: '',
								description: 'GL Account for write-off',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
							{
								displayName: 'Write-Off Description',
								name: 'writeOffDescription',
								type: 'string',
								default: '',
								description: 'Description for write-off',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
							{
								displayName: 'Write-Off Financial Year',
								name: 'writeOffFinYear',
								type: 'string',
								default: '',
								description: 'Financial year for write-off',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
							{
								displayName: 'Write-Off Financial Period',
								name: 'writeOffFinPeriod',
								type: 'string',
								default: '',
								description: 'Financial period for write-off',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
							{
								displayName: 'Write-Off Date',
								name: 'writeOffDate',
								type: 'string',
								default: '',
								description: 'Date for write-off in format YYYY-MM-DD',
								displayOptions: {
									show: {
										includeWriteOff: [true],
									},
								},
							},
						],
					},
				],
			},
			{
				displayName: 'Transactions',
				name: 'transactions',
				placeholder: 'Add transactions to reconcile',
				type: 'fixedCollection',
				typeOptions: {
					multipleValues: true,
				},
				default: {},
				description:
					'Transactions to reconcile against each other. Add at least two transactions (e.g., an invoice and a payment) to be matched in Exact Online. You need the financial year, period, journal code, entry number, and amount for each transaction. These can be obtained from the transaction entries in Exact Online.',
				displayOptions: {
					show: {
						operation: ['postXml'],
						service: ['financial'],
						resource: ['MatchSets'],
					},
				},
				options: [
					{
						name: 'transaction',
						displayName: 'Transaction',
						values: [
							{
								displayName: 'Financial Year',
								name: 'finYear',
								type: 'string',
								default: '',
								required: true,
								description: 'Financial year of the transaction',
							},
							{
								displayName: 'Financial Period',
								name: 'finPeriod',
								type: 'string',
								default: '',
								required: true,
								description: 'Financial period of the transaction',
							},
							{
								displayName: 'Journal',
								name: 'journal',
								type: 'string',
								default: '',
								required: true,
								description: 'Journal code of the transaction',
							},
							{
								displayName: 'Entry',
								name: 'entry',
								type: 'string',
								default: '',
								required: true,
								description: 'Entry number of the transaction',
							},
							{
								displayName: 'Amount (DC)',
								name: 'amountDC',
								type: 'string',
								default: '',
								required: true,
								description: 'Amount (Debit/Credit) of the transaction',
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

				// Special case for MatchSets - only offer postXml operation
				if (service.toLowerCase() === 'financial' && resource === 'MatchSets') {
					return [
						{
							name: 'POST (XML)',
							value: 'postXml',
						},
					];
				}

				// Normal handling for all other resources
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as endpointConfiguration;
				const methods = endpointConfig.methods.map((x) => x.toLowerCase());
				if (methods.includes('get')) {
					methods.push('getAll');
				}

				// Conditionally add 'getAllViaParentId' for specific resources like BankEntryLines
				if (service.toLowerCase() === 'financialtransaction' && resource === 'BankEntryLines') {
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
				)) as endpointConfiguration;
				const fields = await getFields.call(this, endpointConfig);
				return toFieldSelectOptions(fields.map((x) => ({ name: x })) as LoadedFields[]);
			},

			async getFieldsFilter(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as endpointConfiguration;
				const fields = endpointConfig.fields;

				return toFieldFilterOptions(fields as endpointFieldConfiguration[]);
			},

			async getFieldsData(this: ILoadOptionsFunctions) {
				const service = this.getNodeParameter('service', 0) as string;
				const resource = this.getNodeParameter('resource', 0) as string;
				const endpointConfig = (await getEndpointConfig.call(
					this,
					service,
					resource,
				)) as endpointConfiguration;
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
					(x) => !exclude.includes(x.name),
				) as endpointFieldConfiguration[];

				return toFieldFilterOptions(fields as endpointFieldConfiguration[]);
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
		const endpointConfig = (await getEndpointConfig.call(
			this,
			service,
			resource,
		)) as endpointConfiguration;
		const uri = endpointConfig.uri.replace('{division}', division);
		const excludeSelection = this.getNodeParameter('excludeSelection', 0, false) as boolean;
		const selectedFields = this.getNodeParameter('selectedFields', 0, []) as string[];
		let onlyNotSelectedFields: string[] = [];
		if (excludeSelection) {
			const allFields = await getFields.call(this, endpointConfig);
			onlyNotSelectedFields = allFields.filter((x) => !selectedFields.includes(x));
		}

		for (let itemIndex = 0; itemIndex < length; itemIndex++) {
			try {
				if (operation === 'get') {
					const qs: IDataObject = {};
					const id = this.getNodeParameter('id', itemIndex, '') as string;
					if (id !== '') {
						qs['$filter'] = `ID eq guid'${id}'`;
						qs['$top'] = 1;
						responseData = await getData.call(this, uri, {}, qs);
						returnData = returnData.concat(responseData);
					}
				}
				if (operation === 'getAll') {
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
							const fieldType = await getFieldType.call(this, endpointConfig, fieldName);
							const fieldValue = filter[filterIndex].value as string;
							switch (fieldType) {
								case 'string':
									// Handle 'contains' operator specifically for strings, using OData v3 substringof
									if (filter[filterIndex].operator === 'contains') {
										filters.push(`substringof('${fieldValue}', ${fieldName})`); // Correct OData v3 syntax
									} else {
										// Existing logic for other string operators (eq, ne, etc.) - needs quotes
										filters.push(`${fieldName} ${filter[filterIndex].operator} '${fieldValue}'`);
									}
									break;
								case 'boolean':
									filters.push(
										`${fieldName} ${filter[filterIndex].operator} ${
											fieldValue.toLowerCase() === 'true'
										}`,
									);
									break;
								case 'number':
									filters.push(
										`${fieldName} ${filter[filterIndex].operator} ${filter[filterIndex].value}`,
									);
									break;
								default:
									break;
							}
						}
						if (filters.length > 0) {
							qs['$filter'] = filters.join(` ${conjunction} `);
						}
					}

					responseData = await getAllData.call(this, uri, limit, {}, qs, {}, ignoreRateLimit);
					returnData = returnData.concat(responseData);
				}

				if (operation === 'getAllViaParentId') {
					const parentId = this.getNodeParameter('parentId', itemIndex, '') as string;
					const limit = this.getNodeParameter('limit', itemIndex, 60) as number;
					const ignoreRateLimit = this.getNodeParameter('ignoreRateLimit', 0, false) as boolean;

					if (parentId === '') {
						throw new NodeOperationError(
							this.getNode(),
							'Parent ID is required for the getAllViaParentId operation.',
							{ itemIndex },
						);
					}

					const parentResource = 'BankEntries';
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
				}

				if (operation === 'post') {
					let body: IDataObject = {};

					const useManualBody = this.getNodeParameter('useManualBody', itemIndex, false) as boolean;
					if (useManualBody) {
						const manualBody = this.getNodeParameter('manualBody', itemIndex, {}) as IDataObject;
						if (!manualBody) {
							throw new NodeOperationError(
								this.getNode(),
								`Please include the fields and values for the item you want to create.`,
								{
									itemIndex,
								},
							);
						}
						body = manualBody;
					} else {
						const data = this.getNodeParameter('data.field', itemIndex, 0) as IDataObject[];
						if (!data) {
							throw new NodeOperationError(
								this.getNode(),
								`Please include the fields and values for the item you want to create.`,
								{
									itemIndex,
								},
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
								{
									itemIndex,
								},
							);
						}
						if (data.length > 0) {
							for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
								const fieldName = data[dataIndex].fieldName as string;
								const fieldType = await getFieldType.call(this, endpointConfig, fieldName);
								const fieldValue = data[dataIndex].fieldValue as string;
								switch (fieldType) {
									case 'string':
										body[`${fieldName}`] = fieldValue;
										break;
									case 'boolean':
										body[`${fieldName}`] = fieldValue.toLocaleLowerCase() === 'true';
										break;
									case 'number':
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
				}

				if (operation === 'put') {
					const id = this.getNodeParameter('id', itemIndex, '') as string;
					if (id === '') {
						throw new NodeOperationError(
							this.getNode(),
							`Please enter an Id of a record to edit.`,
							{
								itemIndex,
							},
						);
					}
					let body: IDataObject = {};

					const useManualBody = this.getNodeParameter('useManualBody', itemIndex, false) as boolean;
					if (useManualBody) {
						const manualBody = this.getNodeParameter('manualBody', itemIndex, {}) as IDataObject;
						if (!manualBody) {
							throw new NodeOperationError(
								this.getNode(),
								`Please include the fields and values for the item you want to edit.`,
								{
									itemIndex,
								},
							);
						}
						body = manualBody;
					} else {
						const data = this.getNodeParameter('data.field', itemIndex, 0) as IDataObject[];
						if (!data) {
							throw new NodeOperationError(
								this.getNode(),
								`Please include the fields and values for the item you want to edit.`,
								{
									itemIndex,
								},
							);
						}

						if (data.length > 0) {
							for (let dataIndex = 0; dataIndex < data.length; dataIndex++) {
								const fieldName = data[dataIndex].fieldName as string;
								const fieldType = await getFieldType.call(this, endpointConfig, fieldName);
								const fieldValue = data[dataIndex].fieldValue as string;
								switch (fieldType) {
									case 'string':
										body[`${fieldName}`] = fieldValue;
										break;
									case 'boolean':
										body[`${fieldName}`] = fieldValue.toLocaleLowerCase() === 'true';
										break;
									case 'number':
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
				}

				if (operation === 'delete') {
					const id = this.getNodeParameter('id', itemIndex, '') as string;
					if (id === '') {
						throw new NodeOperationError(
							this.getNode(),
							`Please enter an Id of a record to delete.`,
							{
								itemIndex,
							},
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
				}

				if (operation === 'postXml') {
					// Get the reconciliation parameters
					const matchSetData = this.getNodeParameter(
						'reconciliation.matchSet',
						itemIndex,
						{},
					) as IDataObject;
					const transactionsData = this.getNodeParameter(
						'transactions.transaction',
						itemIndex,
						[],
					) as IDataObject[];

					if (!matchSetData.glAccountCode) {
						throw new NodeOperationError(
							this.getNode(),
							'Please provide GL Account Code for reconciliation',
							{ itemIndex },
						);
					}

					if (transactionsData.length < 2) {
						throw new NodeOperationError(
							this.getNode(),
							'At least two transactions are required for reconciliation',
							{ itemIndex },
						);
					}

					// Prepare match lines
					const matchLines: ReconciledTransaction[] = transactionsData.map((transaction) => ({
						finYear: transaction.finYear as string,
						finPeriod: transaction.finPeriod as string,
						journal: transaction.journal as string,
						entry: transaction.entry as string,
						amountDC: transaction.amountDC as string,
					}));

					// Prepare match set
					const matchSet: MatchSet = {
						GLAccount: matchSetData.glAccountCode as string,
						MatchLines: matchLines,
					};

					// Add Account if provided (required for receivable/payable GL accounts)
					if (matchSetData.accountCode) {
						matchSet.Account = matchSetData.accountCode as string;
					}

					// Add write-off information if needed
					if (matchSetData.includeWriteOff === true) {
						const writeOff: WriteOff = {
							type: (matchSetData.writeOffType as string) || '4',
						};

						// Only add fields if they are provided
						if (matchSetData.writeOffGLAccount) {
							writeOff.GLAccount = matchSetData.writeOffGLAccount as string;
						}

						if (matchSetData.writeOffDescription) {
							writeOff.Description = matchSetData.writeOffDescription as string;
						}

						if (matchSetData.writeOffFinYear) {
							writeOff.FinYear = matchSetData.writeOffFinYear as string;
						}

						if (matchSetData.writeOffFinPeriod) {
							writeOff.FinPeriod = matchSetData.writeOffFinPeriod as string;
						}

						if (matchSetData.writeOffDate) {
							writeOff.Date = matchSetData.writeOffDate as string;
						}

						matchSet.WriteOff = writeOff;
					}

					// Create XML for the reconciliation
					const xmlBody = createReconciliationXml([matchSet]);

					// Send XML request
					try {
						const response = await exactOnlineXmlRequest.call(this, division, 'FFMatch', xmlBody);

						if (response.statusCode === 200) {
							returnData.push({
								success: true,
								message: 'Reconciliation completed successfully',
								response: response.body,
							});
						} else {
							throw new NodeOperationError(
								this.getNode(),
								`Failed to reconcile: ${response.statusCode} ${response.body}`,
								{ itemIndex },
							);
						}
					} catch (error) {
						throw new NodeOperationError(this.getNode(), `Failed to reconcile: ${error.message}`, {
							itemIndex,
						});
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
