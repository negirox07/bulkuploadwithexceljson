import * as React from 'react';
import styles from './Upload.module.scss';
import { IUploadProps } from './IUploadProps';
//import UploadAndCreate from './UploadAndCreate';
//import AdvancedBulkUpload from './bulkupload/AdvancedBulkUpload';
import { AdvanceBulkUpload } from './bulkuploadwithcolumns/AdvanceBulkUpload';
export default class Upload extends React.Component<IUploadProps, {}> {
  public render(): React.ReactElement<IUploadProps> {
    const {
      ListName,
      context
    } = this.props;
    return (
      <section className={styles.upload}>
        {/*         <UploadAndCreate context={context} listName={ListName} /> */}
        <AdvanceBulkUpload context={context} listName={ListName} />
 {/*        <AdvancedBulkUpload context={context} listName={ListName} listFields={['RefNo', 'Category', 'Touchpoint', 'Standard']} />
     */}  </section>
    );
  }
}
