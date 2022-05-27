import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'AnchorCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnchorCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'AnchorCommandSet';

export default class AnchorCommandSet extends BaseListViewCommandSet<IAnchorCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized AnchorCommandSet');
    return Promise.resolve();
  }

  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const anchorOpen: Command = this.tryGetCommand('open_with_anchor');
    const anchorShare: Command = this.tryGetCommand('share_with_anchor');

    anchorOpen.visible = event.selectedRows.length === 1;
    anchorShare.visible = event.selectedRows.length === 1;
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    let data = {
      user: {
        name: "",
        email: ""
      },
      site: {
        url: "",
        base: "",
        title: ""
      },
      item: {
        driveId: "",
        itemId: "",
        url: ""
      }
    };

    let absoluteUrl: string;
    let relativeUrl: string;
    let docRelativeUrl: string;

    let itemInfo = [];
    switch (event.itemId) {
      case 'open_with_anchor':
        itemInfo = event.selectedRows[0].getValueByName(".spItemUrl").split('?')[0].split('/drives/')[1].split('/items/');
        data.user.name = this.context.pageContext.user.displayName;
        data.user.email = this.context.pageContext.user.email;
        data.site.url = this.context.pageContext.site.absoluteUrl;
        data.site.base = this.context.pageContext.site.serverRequestPath;
        data.site.title = this.context.pageContext.web.title;
        data.item.driveId = itemInfo[0];
        data.item.itemId = itemInfo[1];

        absoluteUrl = this.context.pageContext.site.absoluteUrl;
        relativeUrl = this.context.pageContext.site.serverRelativeUrl
        docRelativeUrl = event.selectedRows[0].getValueByName('FileRef')

        data.item.url = `${absoluteUrl}/${docRelativeUrl.substr(relativeUrl.length)}?web=1`

        Dialog.alert(`${JSON.stringify(data, null, 4)}`);
        break;
      case 'share_with_anchor':
        itemInfo = event.selectedRows[0].getValueByName(".spItemUrl").split('?')[0].split('/drives/')[1].split('/items/');
        data.user.name = this.context.pageContext.user.displayName;
        data.user.email = this.context.pageContext.user.email;
        data.site.url = this.context.pageContext.site.absoluteUrl;
        data.site.base = this.context.pageContext.site.serverRequestPath;
        data.site.title = this.context.pageContext.web.title;
        data.item.driveId = itemInfo[0];
        data.item.itemId = itemInfo[1];

        absoluteUrl = this.context.pageContext.site.absoluteUrl;
        relativeUrl = this.context.pageContext.site.serverRelativeUrl
        docRelativeUrl = event.selectedRows[0].getValueByName('FileRef')

        data.item.url = `${absoluteUrl}/${docRelativeUrl.substr(relativeUrl.length)}?web=1`

        Dialog.alert(`${JSON.stringify(data, null, 4)}`);
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
