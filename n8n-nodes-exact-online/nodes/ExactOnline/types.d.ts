export type LoadedDivision = {
	Code: string;
	CustomerName: string;
	Description:string;
}
export type LoadedOptions = {
	value:string,
	name:string
}


export type LoadedFields = {
	name:string,
}


export type endpointFieldConfiguration = {
	name:string,
	type:string,
	webhook?:boolean,
	filter?:boolean,
	mandatory:boolean,
}

export type endpointConfiguration = {
	service:string,
	endpoint:string
	uri:string,
	doc:string,
	webhook:boolean,
	methods:string[],
	fields:endpointFieldConfiguration[],
}

export type ReconciledTransaction = {
	finYear: string | number,
	finPeriod: string | number,
	journal: string,
	entry: string | number,
	amountDC: string | number,
}

export type WriteOff = {
	type: string | number, // 0 = Balance; 1 = Discount; 3 = PaymentDifference; 4 = ExchangeRateDifference
	GLAccount?: string,
	Description?: string,
	FinYear?: string | number,
	FinPeriod?: string | number,
	Date?: string,
	VATCorrection?: boolean,
}

export type MatchSet = {
	GLAccount: string, // GLAccount code
	Account?: string, // Account code (optional if GLAccount is not receivable/payable)
	MatchLines: ReconciledTransaction[],
	WriteOff?: WriteOff,
}
