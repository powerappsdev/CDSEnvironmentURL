import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class CDSEnvironmentUrl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	
	private cdsEnvironmentUrl: string;
	private notifyOutputChanged: Function;
	private _updateFromOutput: boolean;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}


	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		this.notifyOutputChanged = notifyOutputChanged;
		this._updateFromOutput = false;
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		if (this._updateFromOutput){
			this._updateFromOutput = false;
			return;
		}

		// Add code to update control view
		let environmentUrl = this.getCdsEnvironmentUrl();
		
		//only send the output from the control if the environment url is available
		// and it has not already been sent to the output.  This will ensure that the url
		// does not fire the onchange multiple times for this control.
		if (environmentUrl && environmentUrl !== this.cdsEnvironmentUrl) {
			this.cdsEnvironmentUrl = environmentUrl;			
			this.notifyOutputChanged();
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		this._updateFromOutput = true;

		return {
			cdsEnvironmentUrl: this.cdsEnvironmentUrl
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	private getCdsEnvironmentUrl(): string {
		let result = this.getCdsEnvironmentUrlFromCdsDataSourceConfigs();

		if (result?.length){
			return result;
		}

		// possibly no _r9, find the object since MS got squirmy
		result = this.getCdsEnvironmentUrlFromSearchingForCdsDataSourceConfigs();

		if (result?.length){
			return result;
		}

		return this.getCdsEnvironmentUrlFromAdalTokenKeys() ||  'CDS Environment URL could not be found';
	}

	private getCdsEnvironmentUrlFromSearchingForCdsDataSourceConfigs(): string | undefined {
		//@ts-ignore
		if (window._r9?.currentApp?.appHostClient?._cdsDataSourceConfigs){
			return;
		}

		for (const i in window){
			//@ts-ignore
			if (window[i]?.currentApp?.appHostClient?._cdsDataSourceConfigs){
				//@ts-ignore
				const dsConfigs = window[i]?.currentApp?.appHostClient?._cdsDataSourceConfigs;
				
				for (const c in dsConfigs){
					if (dsConfigs[c].hasOwnProperty('runtimeUrl') && dsConfigs[c].runtimeUrl.includes('.dynamics.com/api/data/')){
						return dsConfigs[c].runtimeUrl.split('api/')[0];
					}
				 }
				 
				 break;
			}
		}
	}

	private getCdsEnvironmentUrlFromCdsDataSourceConfigs(): string | undefined {
		//@ts-ignore
		const cdsDsConfigs = window._r9?.currentApp?.appHostClient?._cdsDataSourceConfigs;
		
		for (const i in cdsDsConfigs){
			if (cdsDsConfigs[i].hasOwnProperty('runtimeUrl') && cdsDsConfigs[i].runtimeUrl.includes('.dynamics.com/api/data/')){
				return cdsDsConfigs[i].runtimeUrl.split('api/')[0];
			}
		 }
	}

	private getCdsEnvironmentUrlFromAdalTokenKeys(): string | undefined {
		const keys: string[] = parent.localStorage['adal.token.keys'].split('|');
	
		if (!keys.length){
			return;
		}

		const crmKeys = keys.filter(k => k.endsWith('.dynamics.com/'));

		if (!crmKeys.length){
			return;
		}

		return crmKeys[0];
	}

}