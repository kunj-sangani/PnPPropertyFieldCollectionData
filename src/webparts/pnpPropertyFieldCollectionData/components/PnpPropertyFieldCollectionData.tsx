import * as React from 'react';
import styles from './PnpPropertyFieldCollectionData.module.scss';
import { IPnpPropertyFieldCollectionDataProps } from './IPnpPropertyFieldCollectionDataProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PnpPropertyFieldCollectionData extends React.Component<IPnpPropertyFieldCollectionDataProps, {}> {
  public render(): React.ReactElement<IPnpPropertyFieldCollectionDataProps> {
    return (
      <div className={styles.pnpPropertyFieldCollectionData}>
        <div className={styles.container}>
          <div className={styles.row}>
            {this.props.collectionData && this.props.collectionData.map((val) => {
              return (<div><span>{val.Title}</span><span style={{ marginLeft: 10 }}>{val.Lastname}</span></div>);
            })}
          </div>
        </div>
      </div>
    );
  }
}
