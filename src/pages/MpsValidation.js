/* eslint-disable react/no-unused-state */
import React, { PureComponent } from 'react';
import { Upload, Button, Icon } from 'antd';
import excelUtils from '../ExcelUtils/excelUtils';

class MpsValidation extends PureComponent {
  constructor(props) {
    super(props);
    this.state = { fileList: [] };
  }

  handleFileIputChange = fileObj => {
    excelUtils.loadExcelFile(fileObj.target.value);
  };

  render() {
    const { fileList } = this.state;
    const props = {
      onRemove: file => {
        this.setState(state => {
          const index = state.fileList.indexOf(file);
          const newFileList = state.fileList.slice();
          newFileList.splice(index, 1);
          return { fileList: newFileList };
        });
      },
      beforeUpload: file => {
        excelUtils.loadExcelFile(file);
        return false;
      },
      fileList,
    };
    return (
      <div>
        <Upload {...props} accept=".xls,.xlsx">
          <Button>
            <Icon type="upload" /> Select File
          </Button>
        </Upload>
        <input type="file" onChange={this.handleFileIputChange} />
      </div>
    );
  }
}

export default MpsValidation;
