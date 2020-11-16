import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {IProps, DetailListGridControl}  from './DetailListGridControl'
import {IPropsFetch, DetailListGridControlFetch}  from './DetailListGridControlFetch'

export class SubgridFiltered implements ComponentFramework.StandardControl<IInputs, IOutputs> {


	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _detailList: HTMLDivElement;
	private _detailList2: HTMLDivElement;
	private _dataSetVersion: number;
	private _isModelApp: boolean

	private _props: IProps;
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
		// Need to track container resize so that control could get the available width. 
		// The available height won't be provided even when this is true
		context.mode.trackContainerResize(true);

		this._container = container;
		this._context = context;
		this._isModelApp = window.hasOwnProperty('getGlobalContextObject');
		this._dataSetVersion = 0;
		var thisCtrl = this;
		this._props = {
			pcfContext: this._context,
			isModelApp: this._isModelApp,
			dataSetVersion: this._dataSetVersion,
			thisCtrl: thisCtrl
		}

		// set the container to display to relative so that our Scrollable Panel does not cover up the
		// Dynamics ribbon or quick search.
		this._container.style.position = 'relative';
		this._container.setAttribute("id", "pcfcontainer");
		//this._container.style.minHeight = '50%';


		this._detailList = document.createElement("div");
		this._detailList.setAttribute("id", "detailList");
		// if data-is-scrollable is not set then grid will not show all results.
		this._detailList.setAttribute("data-is-scrollable", "true");
		this._detailList.style.float = "left"; 
		this._detailList.style.minWidth = "45%";  

		// sets the height based upon the rowSpan which is there but not included in the Mode interace when
			// the control is a subgrid.
			// Then multiple by 1.5 em which is what MS uses per row.	
			// let rowspan = (this._context.mode as any).rowSpan;
			// if (rowspan) this._detailList.style.height = `${(rowspan * 1.5).toString()}em`;

		this._container.appendChild(this._detailList);

		this._detailList2 = document.createElement("div");
		this._detailList2.setAttribute("id", "detailList");
		// if data-is-scrollable is not set then grid will not show all results.
		this._detailList2.setAttribute("data-is-scrollable", "true");
		this._detailList2.style.float = "right";
		this._detailList2.style.minWidth = "45%"; 

			// sets the height based upon the rowSpan which is there but not included in the Mode interace when
			// the control is a subgrid.
			// Then multiple by 1.5 em which is what MS uses per row.	
			
			//if (rowspan) this._detailList2.style.height = `${(rowspan * 1.5).toString()}em`;

		this._container.appendChild(this._detailList2);

		//añadir botonera inferior

		//set the paging size to 5000
		context.parameters.sampleDataSet.paging.setPageSize(5000);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		var dataSet = context.parameters.sampleDataSet;
		
		if (dataSet.loading) return;

		//if data set has additional pages retrieve them before running anything else
		if (this._isModelApp && dataSet.paging.hasNextPage) {
			dataSet.paging.loadNextPage();
			return;
		}

		//useEffect on the dataSet itself was not picking up on all the updates so pass in a dataset version
		// and update it in the props so the react control knows it was updated.
		this._props.dataSetVersion = this._dataSetVersion++;
		
		// render the DetailsList control
		ReactDOM.render(
			React.createElement(DetailListGridControl, this._props), 
				this._detailList);
		ReactDOM.render(
			React.createElement(DetailListGridControlFetch, this._props), 
				this._detailList2);
	
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

}