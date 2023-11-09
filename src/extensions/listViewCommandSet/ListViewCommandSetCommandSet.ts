import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ListViewCommandSetCommandSetStrings';
import { SPDefault } from "@pnp/nodejs";
import { spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListViewCommandSetCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ListViewCommandSetCommandSet';


const sp = spfi().using(SPDefault({
  baseUrl: 'https://wlmnt.sharepoint.com/sites/Sample',
 
}));

export default class ListViewCommandSetCommandSet extends BaseListViewCommandSet<IListViewCommandSetCommandSetProperties> {

  
  @override
  public onInit(): Promise<void> {

    
    Log.info(LOG_SOURCE, 'Initialized ListViewCommandSetCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    //const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareOneCommandA: Command = this.tryGetCommand('COMMAND_A');
    const compareOneCommandB: Command = this.tryGetCommand('COMMAND_B');
    //if (compareOneCommand) {
    if (compareOneCommandA && compareOneCommandB) {
      const listName: string = this.context.pageContext.list.title
      const studentListName: string = "Students";
      // This command should be hidden unless exactly one row is selected.
      //compareOneCommand.visible = event.selectedRows.length === 1;
      compareOneCommandA.visible = compareOneCommandB.visible = event.selectedRows.length ===1 && listName.toLowerCase() == studentListName.toLowerCase();    
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    var firstName = event.selectedRows[0].getValueByName('Title');
    var lastName = event.selectedRows[0].getValueByName('LastName');
    switch (event.itemId) {
      case 'COMMAND_A':
        //Dialog.alert(`${this.properties.sampleTextOne}`);
        this.Copy(firstName,lastName,"SectionA");
        break;
      case 'COMMAND_B':
        //Dialog.alert(`${this.properties.sampleTextTwo}`);
        this.Copy(firstName,lastName,"SectionB");
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  public Copy(firstName:string, lastName:string, listName:string){
    sp.web.lists.getByTitle(listName).items.add({
    Title: firstName + " " +lastName,
   }).then(()=>{
    Dialog.alert("Student details copied successfully");
   });
   
  }
}
