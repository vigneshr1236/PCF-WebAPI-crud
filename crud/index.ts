import { IInputs, IOutputs } from "./generated/ManifestTypes";

class EntityReference {
	id: string;
	typeName: string;
	constructor(typeName: string, id: string) {
		this.id = id;
		this.typeName = typeName;
	}
}

export class crud implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _entityName: string;
	private _id: any;
	private _entityReference: EntityReference;
	private _context: ComponentFramework.Context<IInputs>;

	private _createRecordButton: HTMLButtonElement;
	private _retrieveRecordByIdButton: HTMLButtonElement;
	private _retrieveMultipleRecordsButton: HTMLButtonElement;
	private _updateRecordButton: HTMLButtonElement;
	private _deleteRecordButton: HTMLButtonElement;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		this._context = context;
		this._entityReference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		);
		this._entityName = "account";

		this._createRecordButton = document.createElement("button");
		this._retrieveRecordByIdButton = document.createElement("button");
		this._retrieveMultipleRecordsButton = document.createElement("button");
		this._updateRecordButton = document.createElement("button");
		this._deleteRecordButton = document.createElement("button");

		this._createRecordButton.innerHTML = "create";
		this._retrieveRecordByIdButton.innerHTML = "retrieveRecordById";
		this._retrieveMultipleRecordsButton.innerHTML = "retrieveMultipleRecords";
		this._updateRecordButton.innerHTML = "updateRecord";
		this._deleteRecordButton.innerHTML = "deleteRecord";

		this._createRecordButton.addEventListener("click", this.createRecord.bind(this));
		this._retrieveRecordByIdButton.addEventListener("click", this.retrieveRecordById.bind(this));
		this._retrieveMultipleRecordsButton.addEventListener("click", this.retrieveMultipleRecords.bind(this));
		this._updateRecordButton.addEventListener("click", this.updateRecord.bind(this));
		this._deleteRecordButton.addEventListener("click", this.deleteRecord.bind(this));

		container.appendChild(this._createRecordButton);
		container.appendChild(this._retrieveRecordByIdButton);
		container.appendChild(this._retrieveMultipleRecordsButton);
		container.appendChild(this._updateRecordButton);
		container.appendChild(this._deleteRecordButton);

	}

	private createRecord(): void {
		let createRecord : any = {};
		createRecord["name"] = "CreateCompany";
		createRecord["revenue"] = 5000;
		
		var thisRef = this;
		// Invoke the Web API to creat the new record
		this._context.webAPI.createRecord(this._entityName, createRecord).then
			(
				// Callback method for successful creation of new record
				function (response: ComponentFramework.EntityReference) {
					// Get the ID of the new record created
					thisRef._id = response.id;
					console.log(response.id);
				},
				function (errorResponse: any) {
					// Error handling code here - record failed to be created
					console.log(errorResponse);
				}
			);
	}

	private retrieveRecordById() {
		try {
			this._context.webAPI.retrieveRecord(this._entityReference.typeName, this._entityReference.id).then(e => console.log(e));
		}
		catch (error) {
			console.log(error);
			return [];
		}
	}

	private async retrieveMultipleRecords() {

		let fetchXml =
			"<fetch>" +
			"<entity name='account'>" +
			"<attribute name='name' />" +
			"<attribute name='primarycontactid' />" +
			"<attribute name='telephone1' />" +
			"<attribute name='accountid' />" +
			"<order attribute='name' descending='false' />" +
			"</entity>" +
			"</fetch>";

		let query = '?fetchXml=' + encodeURIComponent(fetchXml);

		try {
			const result = await this._context.webAPI.retrieveMultipleRecords(this._entityReference.typeName, query);
			console.log(result);
		}
		catch (error) {
			console.log(error);
			return [];
		}
	}

	private updateRecord(): void {

		let updateRecord : any = {};
		updateRecord["name"] = "UpdateCompany";
		updateRecord["revenue"] = 15000;

		// Invoke the Web API to creat the new record
		this._context.webAPI.updateRecord(this._entityName, this._id, updateRecord).then
			(
				// Callback method for successful creation of new record
				function (response: ComponentFramework.EntityReference) {
					//updated
					console.log("update record" + response.id +"of entity type "+ response.entityType);
				},
				function (errorResponse: any) {
					// Error handling code here - record failed to be created
					console.log(errorResponse);
				}
			);
	}

	private deleteRecord(): void {
         var ref = this;
		// Invoke the Web API to creat the new record
		this._context.webAPI.deleteRecord(this._entityName, this._id).then
			(
				// Callback method for successful creation of new record
				function (response: ComponentFramework.EntityReference) {
					console.log(response.id+"Record Was deleted");
				},
				function (errorResponse: any) {
					// Error handling code here - record failed to be created
					console.log(errorResponse);
				}
			);
	}



	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		this._context = context;
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}