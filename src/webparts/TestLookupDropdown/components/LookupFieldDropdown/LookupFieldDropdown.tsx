/*
    tslint:disable:no-any max-line-length
*/
import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { CascadingDataSource } from '../../data/CascadingDataSource';

export interface ILookupFieldDropdownProps {
    listId: string;
    displayField: string;
    lookupField?: string;
    dataSource: CascadingDataSource;
    onChange?: (key: string | number) => void;
}

export interface ILookupFieldDropdownState {
    filterIndex: number;
    options: IDropdownOption[];
    selectedKey: string | number;
}

export class LookupFieldDropdown extends React.Component<ILookupFieldDropdownProps, ILookupFieldDropdownState> {
    private _listIndex: number;

    constructor(props: ILookupFieldDropdownProps) {
        super(props);

        this._listIndex = this.props.dataSource.addListInfo({
            Component: this,
            ListId: props.listId,
            DisplayField: props.displayField,
            FilterField: props.lookupField
        }) - 1;

        this.state = {
            filterIndex: 0,
            options: [],
            selectedKey: undefined
        };
    }

    public async componentDidMount(): Promise<void> {
        await this.loadOptions();
    }

    public async componentDidUpdate(prevProps: Readonly<ILookupFieldDropdownProps>, prevState: Readonly<ILookupFieldDropdownState>, prevContext: any): Promise<void> {
        if (this.state.filterIndex !== prevState.filterIndex) {
            await this.loadOptions();
        }
    }

    public render(): JSX.Element {
        const { options, selectedKey } = this.state;

        return (
            <Dropdown
                options={options}
                selectedKey={selectedKey}
                onChanged={this.selectionChanged}
            />
        );
    }

    private async loadOptions(): Promise<void> {
        const options: IDropdownOption[] = await this.props.dataSource.getOptions(this._listIndex);
        this.setState({ options });
    }

    private selectionChanged = (option: IDropdownOption, index?: number) => {
        const selectedKey: string | number = option ? option.key : 0;
        this.setState({ selectedKey: selectedKey});
        this.props.dataSource.setSelectedListItem(this._listIndex, selectedKey);
        if (typeof this.props.onChange === 'function') {
            this.props.onChange(selectedKey);
        }
    }
}