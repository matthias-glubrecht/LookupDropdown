/*
    tslint:disable:max-line-length
*/
import * as React from 'react';
import styles from './TestLookupDropdown.module.scss';
import { ITestLookupDropdownProps } from './ITestLookupDropdownProps';
import { LookupFieldDropdown } from '../LookupFieldDropdown/LookupFieldDropdown';
import { ITestLookupDropdownState } from './ITestLookupDropdownState';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { CascadingDataSource } from '../../data/CascadingDataSource';

export default class TestLookupDropdown extends React.Component<ITestLookupDropdownProps, ITestLookupDropdownState> {
  private _dataSource: CascadingDataSource;
  constructor(props: ITestLookupDropdownProps) {
    super(props);
    this._dataSource = new CascadingDataSource();
  }

  public render(): React.ReactElement<ITestLookupDropdownProps> {
    return (
      <div className={styles.TestLookupDropdown} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Test Filtered Dropdown</span>
            </div>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <Label>Land</Label>
              <LookupFieldDropdown dataSource={this._dataSource} listId={this.props.list1} displayField='Title' />
              <p></p>
              <Label >Stadt</Label>
              <LookupFieldDropdown dataSource={this._dataSource} listId={this.props.list2} displayField='Title' lookupField='Land' />
              <p> </p>
              <Label >Straße</Label>
              <LookupFieldDropdown dataSource={this._dataSource} listId={this.props.list3} displayField='Title' lookupField='Stadt' />
            </div>
          </div>
        </div>
      </div >
    );
  }
}
