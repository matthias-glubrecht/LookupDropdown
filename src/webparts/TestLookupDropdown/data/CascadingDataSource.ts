/*
    tslint:disable:max-line-length no-any
*/
import {
    IDropdownOption
} from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';

export interface IListInfo {
    Component: React.Component;
    ListId: string;
    DisplayField: string;
    SelectedKey?: string | number;
    FilterField?: string;
}

export class CascadingDataSource {
    private _lists: IListInfo[];

    public constructor(...lists: IListInfo[]) {
        this._lists = lists;
    }

    public addListInfo(list: IListInfo): number {
        return this._lists.push(list);
    }

    public setSelectedListItem(listIndex: number, selectedKey: string | number): void {
        if (listIndex >= 0 && listIndex < this._lists.length) {
            this._lists[listIndex].SelectedKey = selectedKey;
            if (listIndex < this._lists.length - 1) {
                this._lists[listIndex + 1].SelectedKey = undefined;
                this._lists[listIndex + 1].Component.setState({filterIndex: selectedKey, options: []});
            }
        }
    }

    public async getOptions(index: number): Promise<IDropdownOption[]> {
        if (index >= 0 && index < this._lists.length) {
            const { ListId: listId, DisplayField: displayField, FilterField: filterField } = this._lists[index];
            if (index > 0 && !!filterField && !!this._lists[index - 1].SelectedKey ) {
                // filter the data based on the previous dropdown
                const filterId: string | number = this._lists[index - 1].SelectedKey;
                const items: {id: string; [key: string]: string}[] = await sp.web.lists.getById(listId).items.filter(`${filterField}/Id eq ${filterId}`).select('Id', displayField).get();
                return items.map(item => {
                    return {
                        key: item.Id,
                        text: item[displayField]
                    };
                });
            } else if (index === 0) {
                const items: any[] = await sp.web.lists.getById(listId).items.select('Id', displayField).get();
                return items.map(item => {
                    return {
                        key: item.Id,
                        text: item[this._lists[index].DisplayField]
                    };
                });
            }
        }
        return [];
    }
}