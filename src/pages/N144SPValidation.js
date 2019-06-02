/* eslint-disable no-prototype-builtins */
/* eslint-disable no-nested-ternary */
/* eslint-disable prefer-destructuring */
/* eslint-disable compat/compat */
/* eslint-disable operator-assignment */
/* eslint-disable react/destructuring-assignment */
/* eslint-disable react/jsx-boolean-value */
/* eslint-disable prefer-template */
/* eslint-disable no-param-reassign */
/* eslint-disable no-plusplus */
/* eslint-disable class-methods-use-this */
/* eslint-disable react/no-unused-state */
import React, { PureComponent } from 'react';
import { Upload, Button, Icon, Card, Row, Col, Select, Spin, Table, Modal, Input } from 'antd';
import excelUtils from '../ExcelUtils/excelUtils';
import constants from '@/common/constants';

const { Option } = Select;
const Search = Input.Search;
const {
  VendorCodeMapping,
  BRSupplyPlanColumnsDefine,
  manualColumnsDefine,
  BRHeaderColumns,
  ManualHeaderColumns,
  commodityMapping,
  summaryTable,
} = constants;
class MpsValidation extends PureComponent {
  constructor(props) {
    super(props);
    this.state = {
      fileList: [],
      BuildRabbitSupplyWorkBook: {},
      BuildRabbitSheetNames: this.getBuildRabbitSheetNameOptions(['GBSupplyPlan']),
      BRSPSheetName: 'GBSupplyPlan',
      BuildRabbitSupplyColumns: this.getBuildRabbitSheetNameOptions(BRSupplyPlanColumnsDefine),
      BuildRabbitSupplyDataSheet: [],
      BuildRabbitSupplyPlanStartColumn: this.getTodayforBR(),
      manualFileList: [],
      ManualWorkBook: {},
      ManualSheetNames: this.getBuildRabbitSheetNameOptions(['Supply Plan']),
      ManualSPSheetName: 'Supply Plan',
      ManualSupplyDataSheet: [],
      ManualSupplyPlanStartColumn: this.getTodayforManual(),
      ManualSupplyColumns: this.getBuildRabbitSheetNameOptions(manualColumnsDefine),
      loading: false,
      discrepancySheet: [],
      discrepancyColumns: [],
      showFileSelectSection: true,
      showModal: false,
      currentSheetFomat: '',
      originalDataTableSheet: [],
      originalDataTableColumns: [],
      showSummaryTable: false,
      ShipCommitDataSheet: [],
      summaryTableSheet: [],
      summaryTableColumns: [],
      BRDataReady: false,
      ShipCommitReady: false,
      ManualDataReady: false,
      originalDisSheet: [],
      showFullDisTable: true,
      partAnalysisDataSource: [],
      partAnalysisColumns: [],
      showPartAnalysisTable: false,
    };
  }

  initializeData = () => {
    this.setState({
      fileList: [],
      BuildRabbitSupplyWorkBook: {},
      BuildRabbitSheetNames: this.getBuildRabbitSheetNameOptions(['GBSupplyPlan']),
      BRSPSheetName: 'GBSupplyPlan',
      BuildRabbitSupplyColumns: this.getBuildRabbitSheetNameOptions(BRSupplyPlanColumnsDefine),
      BuildRabbitSupplyDataSheet: [],
      BuildRabbitSupplyPlanStartColumn: this.getTodayforBR(),
      manualFileList: [],
      ManualWorkBook: {},
      ManualSheetNames: this.getBuildRabbitSheetNameOptions(['Supply Plan']),
      ManualSPSheetName: 'Supply Plan',
      ManualSupplyDataSheet: [],
      ManualSupplyPlanStartColumn: this.getTodayforManual(),
      ManualSupplyColumns: this.getBuildRabbitSheetNameOptions(manualColumnsDefine),
      loading: false,
      discrepancySheet: [],
      discrepancyColumns: [],
      showFileSelectSection: true,
      showModal: false,
      currentSheetFomat: '',
      originalDataTableSheet: [],
      originalDataTableColumns: [],
      showSummaryTable: false,
      ShipCommitDataSheet: [],
      summaryTableSheet: [],
      summaryTableColumns: [],
      BRDataReady: false,
      ShipCommitReady: false,
      ManualDataReady: false,
      originalDisSheet: [],
      showFullDisTable: true,
      partAnalysisDataSource: [],
      partAnalysisColumns: [],
      showPartAnalysisTable: false,
    });
  };

  getTodayforManual = () => {
    const today = new Date();
    return today.getMonth() + 1 + '月' + today.getDate() + '日';
  };

  getTodayforBR = () => {
    const today = new Date();
    const monthArray = [
      'JAN',
      'FEB',
      'MAR',
      'APR',
      'MAY',
      'JUN',
      'JUL',
      'AUG',
      'SEP',
      'OCT',
      'NOV',
      'DEC',
    ];
    const strDate = today.getDate() + '-' + monthArray[today.getMonth()];
    const columns = BRSupplyPlanColumnsDefine.filter(item => item.indexOf(strDate) > 0);
    return columns && columns.length > 0 ? columns[0] : '';
  };

  onBuildRabbitTemplateSupplyPlanSheetChange = sheetName => {
    const { BuildRabbitSupplyWorkBook } = this.state;
    if (BuildRabbitSupplyWorkBook.SheetNames.indexOf(sheetName) >= 0) {
      const workSheet = excelUtils.excelSheetToJson(BuildRabbitSupplyWorkBook, sheetName);
      this.setState({ BRSPSheetName: sheetName, BuildRabbitSupplyDataSheet: workSheet });
    }
  };

  onManualSheetChange = sheetName => {
    const { ManualWorkBook } = this.state;
    if (ManualWorkBook.SheetNames.indexOf(sheetName) >= 0) {
      const workSheet = excelUtils.excelSheetToJson(ManualWorkBook, sheetName);
      this.setState({ ManualSPSheetName: sheetName, ManualSupplyDataSheet: workSheet });
    }
  };

  onBuildRabbitTemplateSupplyPlanColumnChange = col => {
    this.setState({
      BuildRabbitSupplyPlanStartColumn: col,
    });
  };

  onManualSupplyPlanColumnChange = col => {
    this.setState({ ManualSupplyPlanStartColumn: col });
  };

  getBuildRabbitSheetNameOptions = BuildRabbitSheetNames => {
    const options = [];
    BuildRabbitSheetNames.forEach(name => {
      options.push(
        <Option value={name} key={name}>
          {name}
        </Option>
      );
    });
    return options;
  };

  startManualComparation = () => {
    const {
      ManualSupplyPlanStartColumn,
      BuildRabbitSupplyPlanStartColumn,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
    } = this.state;

    this.setState({
      currentSheetFomat: 'manual',
    });

    const weeksColumns = BRSupplyPlanColumnsDefine.slice(
      BRSupplyPlanColumnsDefine.indexOf(BuildRabbitSupplyPlanStartColumn),
      BRSupplyPlanColumnsDefine.length
    );
    const manualColumns = manualColumnsDefine.slice(
      manualColumnsDefine.indexOf(ManualSupplyPlanStartColumn),
      manualColumnsDefine.length
    );
    let middleSheet = excelUtils.initialMiddleSheet(
      ManualSupplyDataSheet,
      ManualHeaderColumns,
      manualColumns
    );

    middleSheet = excelUtils.compareWorksheet(
      middleSheet,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
      weeksColumns,
      manualColumns,
      manualColumns
    );

    this.buildDiscrepancyTable(middleSheet);
  };

  startBRComparation = () => {
    const {
      ManualSupplyPlanStartColumn,
      BuildRabbitSupplyPlanStartColumn,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
    } = this.state;
    this.setState({ currentSheetFomat: 'BR' });
    const weeksColumns = BRSupplyPlanColumnsDefine.slice(
      BRSupplyPlanColumnsDefine.indexOf(BuildRabbitSupplyPlanStartColumn),
      BRSupplyPlanColumnsDefine.length
    );
    const manualColumns = manualColumnsDefine.slice(
      manualColumnsDefine.indexOf(ManualSupplyPlanStartColumn),
      manualColumnsDefine.length
    );
    let middleSheet = excelUtils.initialMiddleSheet(
      BuildRabbitSupplyDataSheet,
      BRHeaderColumns,
      weeksColumns
    );

    middleSheet = excelUtils.compareWorksheet(
      middleSheet,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
      weeksColumns,
      manualColumns,
      weeksColumns
    );

    this.buildDiscrepancyTable(middleSheet);
  };

  startFullComparation = () => {
    const {
      ManualSupplyPlanStartColumn,
      BuildRabbitSupplyPlanStartColumn,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
    } = this.state;
    this.setState({ currentSheetFomat: 'Full' });
    const weeksColumns = BRSupplyPlanColumnsDefine.slice(
      BRSupplyPlanColumnsDefine.indexOf(BuildRabbitSupplyPlanStartColumn),
      BRSupplyPlanColumnsDefine.length
    );
    const manualColumns = manualColumnsDefine.slice(
      manualColumnsDefine.indexOf(ManualSupplyPlanStartColumn),
      manualColumnsDefine.length
    );
    let middleSheet = excelUtils.initialFullPartMiddleSheet(commodityMapping, weeksColumns);

    middleSheet = excelUtils.compareFullWorksheet(
      middleSheet,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
      weeksColumns,
      manualColumns,
      weeksColumns
    );

    this.buildDiscrepancyTable(middleSheet);
  };

  buildDiscrepancyTable = workSheet => {
    const columns = [];
    if (!workSheet || workSheet.length <= 0) return;

    Object.keys(workSheet[0]).forEach(col => {
      if (col === 'Part Desc' || col === 'Supplier Desc') {
        columns.push({
          title: col,
          dataIndex: col,
          key: col,
          width: 400,
          textWrap: 'word-break',
        });
      } else if (col === 'Buyer') {
        columns.push({
          title: col,
          dataIndex: col,
          key: col,
          width: 300,
          textWrap: 'word-break',
        });
      } else {
        columns.push({ title: col, dataIndex: col, key: col, width: 200 });
      }
    });

    this.setState({
      discrepancySheet: workSheet,
      originalDisSheet: JSON.parse(JSON.stringify(workSheet)),
      discrepancyColumns: columns,
      loading: false,
      showFileSelectSection: false,
      showSummaryTable: false,
      showPartAnalysisTable: false,
    });
  };

  loadShipCommitNightOwlFile = file => {
    const { BuildRabbitSupplyDataSheet } = this.state;
    let workSheet = [];
    this.setState({ loading: true });
    excelUtils.loadExcelFile(file).then(workbook => {
      workSheet = excelUtils.excelSheetToJson(workbook, workbook.SheetNames[0]);
      workSheet = workSheet.filter(
        row =>
          row.Measure &&
          row.Measure === 'Commit' &&
          row['Last published:'] &&
          row['Last published:'] !== ''
      );
      workSheet = workSheet.map(row => {
        row = {
          APN: row['Apple Part No.'],
          'Apple Supplier ID': row['Supplier Code'],
          'Ship Commit Received in Build Rabbit': row['Last published:'],
        };
        return row;
      });
      if (BuildRabbitSupplyDataSheet && BuildRabbitSupplyDataSheet.length > 0) {
        BuildRabbitSupplyDataSheet.forEach(row => {
          const SUPPLIER = row.SUPPLIER;
          for (let i = 0; i < workSheet.length; i++) {
            const key = workSheet[i].APN + ':' + workSheet[i]['Apple Supplier ID'];
            if (SUPPLIER === key) {
              row.ShipCommitDate = workSheet[i]['Ship Commit Received in Build Rabbit'];
              break;
            }
          }
        });
      }

      this.setState({
        BuildRabbitSupplyDataSheet,
        ShipCommitDataSheet: workSheet,
        loading: false,
        ShipCommitReady: true,
      });
    });

    return false;
  };

  loadShipCommitFile = file => {
    const { BuildRabbitSupplyDataSheet } = this.state;
    let workSheet = [];
    this.setState({ loading: true });
    excelUtils.loadExcelFile(file).then(workbook => {
      if (workbook.SheetNames.indexOf('Sheet1') >= 0) {
        workSheet = excelUtils.excelSheetToJson(workbook, 'Sheet1');
      }

      if (BuildRabbitSupplyDataSheet && BuildRabbitSupplyDataSheet.length > 0) {
        BuildRabbitSupplyDataSheet.forEach(row => {
          const SUPPLIER = row.SUPPLIER;
          for (let i = 0; i < workSheet.length; i++) {
            const key = workSheet[i].APN + ':' + workSheet[i]['Apple Supplier ID'];
            if (SUPPLIER === key) {
              row.ShipCommitDate = workSheet[i]['Ship Commit Received in Build Rabbit'];
              break;
            }
          }
        });
      }

      this.setState({
        BuildRabbitSupplyDataSheet,
        ShipCommitDataSheet: workSheet,
        loading: false,
        ShipCommitReady: true,
      });
    });

    return false;
  };

  loadBuildRabbitTemplate = file => {
    const { BRSPSheetName, ShipCommitDataSheet } = this.state;
    let workSheet = [];
    this.setState({
      loading: true,
    });
    excelUtils.loadExcelFile(file).then(workbook => {
      if (workbook.SheetNames.indexOf(BRSPSheetName) >= 0) {
        workSheet = excelUtils.excelSheetToJson(workbook, BRSPSheetName);
      }

      workSheet.forEach(row => {
        const SUPPLIER = row.SUPPLIER;
        for (let i = 0; i < commodityMapping.length; i++) {
          if (SUPPLIER === commodityMapping[i].appleKey) {
            row.Commodity = commodityMapping[i].commodity;
            row['Supplier Desc'] = commodityMapping[i].supplier;
            break;
          }
        }
        if (ShipCommitDataSheet && ShipCommitDataSheet.length > 0) {
          for (let i = 0; i < ShipCommitDataSheet.length; i++) {
            const key =
              ShipCommitDataSheet[i].APN + ':' + ShipCommitDataSheet[i]['Apple Supplier ID'];
            if (SUPPLIER === key)
              row.ShipCommitDate = ShipCommitDataSheet[i]['Ship Commit Received in Build Rabbit'];
            break;
          }
        }
      });

      this.setState({
        BuildRabbitSupplyWorkBook: workbook,
        BuildRabbitSheetNames: this.getBuildRabbitSheetNameOptions(workbook.SheetNames),
        BuildRabbitSupplyDataSheet: workSheet,

        loading: false,
        BRDataReady: true,
      });
    });

    return false;
  };

  findOriginalData = record => {
    const {
      ManualSupplyPlanStartColumn,
      BuildRabbitSupplyPlanStartColumn,
      BuildRabbitSupplyDataSheet,
      ManualSupplyDataSheet,
      currentSheetFomat,
    } = this.state;
    const weeksColumns = BRSupplyPlanColumnsDefine.slice(
      BRSupplyPlanColumnsDefine.indexOf(BuildRabbitSupplyPlanStartColumn),
      BRSupplyPlanColumnsDefine.length
    );
    const manualColumns = manualColumnsDefine.slice(
      manualColumnsDefine.indexOf(ManualSupplyPlanStartColumn),
      manualColumnsDefine.length
    );
    let middleSheet = [];
    const supplier = record.SUPPLIER;
    if (currentSheetFomat === 'BR') {
      middleSheet = excelUtils.initialDisMiddleSheet(
        BuildRabbitSupplyDataSheet,
        BRHeaderColumns,
        weeksColumns,
        supplier
      );
      middleSheet = excelUtils.getOriginalCompareData(
        middleSheet,
        BuildRabbitSupplyDataSheet,
        ManualSupplyDataSheet,
        weeksColumns,
        manualColumns,
        weeksColumns
      );
    } else if (currentSheetFomat === 'manual') {
      middleSheet = excelUtils.initialDisMiddleSheet(
        ManualSupplyDataSheet,
        ManualHeaderColumns,
        manualColumns,
        supplier
      );
      middleSheet = excelUtils.getOriginalCompareData(
        middleSheet,
        BuildRabbitSupplyDataSheet,
        ManualSupplyDataSheet,
        weeksColumns,
        manualColumns,
        manualColumns
      );
    } else {
      middleSheet = excelUtils.initialDisMiddleSheet(
        BuildRabbitSupplyDataSheet,
        BRHeaderColumns,
        weeksColumns,
        supplier
      );
      middleSheet = excelUtils.getOriginalCompareData(
        middleSheet,
        BuildRabbitSupplyDataSheet,
        ManualSupplyDataSheet,
        weeksColumns,
        manualColumns,
        weeksColumns
      );
    }

    this.buildOriDataTable(middleSheet);
  };

  buildOriDataTable = workSheet => {
    const columns = [];
    if (!workSheet || workSheet.length <= 0) return;

    Object.keys(workSheet[0]).forEach(col => {
      if (col === 'Part Desc' || col === 'Supplier Desc') {
        columns.push({
          title: col,
          dataIndex: col,
          key: col,
          width: 400,
          textWrap: 'word-break',
        });
      } else {
        columns.push({ title: col, dataIndex: col, key: col, width: 200 });
      }
    });

    this.setState({
      originalDataTableSheet: workSheet,
      originalDataTableColumns: columns,
      loading: false,
      showModal: true,
    });
  };

  buildSummaryTable = summaryTableSheet => {
    let unikey = 0;
    const summaryTableColumns = [
      { title: 'Commodity', dataIndex: 'Commodity', key: unikey++, width: 200 },
      { title: 'Supplier', dataIndex: 'supplier', key: unikey++, width: 400 },
      { title: 'TotalCount', dataIndex: 'total', key: unikey++, width: 100 },
      { title: 'ManualCount', dataIndex: 'manual', key: unikey++, width: 100 },
      { title: 'BR Count', dataIndex: 'br', key: unikey++, width: 100 },
      { title: 'Manual Percent', dataIndex: 'manualpercent', key: unikey++, width: 100 },
      { title: 'BR Matched', dataIndex: 'brmatch', key: unikey++, width: 100 },
      { title: 'BR Dis-Matched', dataIndex: 'brDisMatch', key: unikey++, width: 100 },
      { title: 'BR Percent', dataIndex: 'brpercent', key: unikey++, width: 100 },
    ];
    this.setState({
      summaryTableSheet,
      summaryTableColumns,
      showSummaryTable: true,
      showFileSelectSection: false,
      showPartAnalysisTable: false,
    });
  };

  buildPartAnalysisTable = disTable => {
    let unikey = 0;
    const disTableColumns = [
      { title: 'SUPPLIER', dataIndex: 'SUPPLIER', key: unikey++, width: 200 },
      { title: 'Commodity', dataIndex: 'Commodity', key: unikey++, width: 200 },
      { title: 'VendorName', dataIndex: 'VendorName', key: unikey++, width: 400 },
      {
        title: 'CommodityAndSupplier Comments',
        dataIndex: 'CommodityAndSupplier',
        key: unikey++,
        width: 2000,
      },
      { title: 'Build Rabbit Comments', dataIndex: 'BuildRabbitTable', key: unikey++, width: 2000 },
      { title: 'Manual SP Comments', dataIndex: 'ManualTable', key: unikey++, width: 2000 },
      { title: 'Full Part List Comments', dataIndex: 'PartList', key: unikey++, width: 2000 },
    ];
    this.setState({
      partAnalysisDataSource: disTable,
      partAnalysisColumns: disTableColumns,
      showSummaryTable: false,
      showFileSelectSection: false,
      showPartAnalysisTable: true,
    });
  };

  partListAnalysis = () => {
    const { BuildRabbitSupplyDataSheet, ManualSupplyDataSheet } = this.state;
    const disTable = [];
    const disrow = {
      SUPPLIER: '',
      Commodity: '',
      VendorName: '',
      CommodityAndSupplier: '', // 'SUPPLIER:Commodity:Supplier',
      BuildRabbitTable: '', // 'SUPPLIER + exist in where',
      ManualTable: '', // 'SUPPLIER + exist where',
      PartList: '', // 'SUPPLIER+exist where',
    };

    function checkPart(row, dataSheet, field, message) {
      let existInDataSheet = false;
      const dataSheetKey = dataSheet[0].hasOwnProperty('SUPPLIER') ? 'SUPPLIER' : 'appleKey';
      const rowKey = row.hasOwnProperty('SUPPLIER') ? 'SUPPLIER' : 'appleKey';
      for (let i = 0; i < dataSheet.length; i++) {
        if (row[rowKey] === dataSheet[i][dataSheetKey]) {
          existInDataSheet = true;
          break;
        }
      }
      if (!existInDataSheet) {
        let existInDisTable = false;
        for (let i = 0; i < disTable.length; i++) {
          if (row[rowKey] === disTable[i].SUPPLIER) {
            existInDisTable = true;
            disTable[i][field] += message;
          }
        }
        if (!existInDisTable) {
          const newrow = JSON.parse(JSON.stringify(disrow));
          newrow.SUPPLIER = row[[rowKey]];
          newrow.Commodity = row.hasOwnProperty('Commodity') ? row.Commodity : row.commodity;
          newrow.VendorName = row.hasOwnProperty('supplier')
            ? row.supplier
            : row.hasOwnProperty('Supplier Desc')
            ? row['Supplier Desc']
            : row.Vendor;
          newrow[field] = message;
          disTable.push(newrow);
        }
      }
    }

    function checkCommoditySupplier(row) {
      const rowKey = row.hasOwnProperty('SUPPLIER') ? 'SUPPLIER' : 'appleKey';
      const rowCommodity = row.hasOwnProperty('Commodity') ? 'Commodity' : 'commodity';
      const rowSupplier = row.hasOwnProperty('supplier')
        ? 'supplier'
        : row.hasOwnProperty('Supplier Desc')
        ? 'Supplier Desc'
        : 'Vendor';

      let existInSummaryTable = false;
      for (let i = 0; i < summaryTable.length; i++) {
        if (
          row[rowCommodity] === summaryTable[i].Commodity &&
          row[rowSupplier] === summaryTable[i].supplier
        ) {
          existInSummaryTable = true;
          break;
        }
      }
      if (!existInSummaryTable) {
        let existInDisTable = false;
        for (let i = 0; i < disTable.length; i++) {
          if (row[rowKey] === disTable[i].SUPPLIER) {
            existInDisTable = true;
            disTable[i].CommodityAndSupplier +=
              row[rowCommodity] + ':' + row[rowSupplier] + 'not exist in Summary Table';
          }
        }
        if (!existInDisTable) {
          const newrow = JSON.parse(JSON.stringify(disrow));
          newrow.SUPPLIER = row[[rowKey]];
          newrow.Commodity = row[rowCommodity];
          newrow.VendorName = row[rowSupplier];
          newrow.CommodityAndSupplier =
            row[rowCommodity] + ':' + row[rowSupplier] + 'not exist in Summary Table';
          disTable.push(newrow);
        }
      }
    }

    BuildRabbitSupplyDataSheet.forEach(row => {
      checkPart(row, ManualSupplyDataSheet, 'ManualTable', 'exist in BR but not in Manual SP,');
      checkPart(row, commodityMapping, 'PartList', 'exist in BR list but not in Part List,');
      checkCommoditySupplier(row);
    });
    ManualSupplyDataSheet.forEach(row => {
      checkPart(
        row,
        BuildRabbitSupplyDataSheet,
        'BuildRabbitTable',
        'exist in Manual SP but not in BR,'
      );
      checkPart(row, commodityMapping, 'PartList', 'exist in Manual SP but not in Part List,');
      checkCommoditySupplier(row);
    });
    commodityMapping.forEach(row => {
      checkPart(row, ManualSupplyDataSheet, 'ManualTable', 'no manual SP,');
      checkPart(row, BuildRabbitSupplyDataSheet, 'BuildRabbitTable', 'no BR SP,');
      checkCommoditySupplier(row);
    });
    this.buildPartAnalysisTable(disTable);
  };

  updateSummaryTable = () => {
    const { BuildRabbitSupplyDataSheet, ManualSupplyDataSheet, discrepancySheet } = this.state;
    const mySummaryTable = JSON.parse(JSON.stringify(summaryTable));

    commodityMapping.forEach(row => {
      for (let i = 0; i < mySummaryTable.length; i++) {
        if (
          mySummaryTable[i].Commodity === row.commodity &&
          mySummaryTable[i].supplier === row.supplier
        ) {
          mySummaryTable[i].total += 1;
          break;
        }
      }
    });

    BuildRabbitSupplyDataSheet.forEach(row => {
      if (row.ShipCommitDate && row.ShipCommitDate !== '') {
        for (let i = 0; i < mySummaryTable.length; i++) {
          if (
            mySummaryTable[i].Commodity === row.Commodity &&
            mySummaryTable[i].supplier === row['Supplier Desc']
          ) {
            mySummaryTable[i].br += 1;
            for (let k = 0; k < discrepancySheet.length; k++) {
              if (discrepancySheet[k].SUPPLIER === row.SUPPLIER) {
                if (discrepancySheet[k]['DIS-COUNT'] !== 0) {
                  mySummaryTable[i].brmatch += 1;
                } else {
                  mySummaryTable[i].brDisMatch += 1;
                }
              }
            }
            break;
          }
        }
      }
    });
    ManualSupplyDataSheet.forEach(row => {
      if (row.hasOwnProperty('更新日期') && row['更新日期'] !== '') {
        for (let i = 0; i < mySummaryTable.length; i++) {
          if (
            mySummaryTable[i].Commodity === row.Commodity &&
            mySummaryTable[i].supplier === row.Vendor
          ) {
            mySummaryTable[i].manual += 1;
            break;
          }
        }
      }
    });
    for (let i = 0; i < mySummaryTable.length; i++) {
      if (mySummaryTable[i].total !== 0) {
        mySummaryTable[i].manualpercent =
          Number.parseFloat((mySummaryTable[i].manual / mySummaryTable[i].total) * 100).toFixed(2) +
          '%';
      } else {
        mySummaryTable[i].manualpercent = '100%';
      }
      if (mySummaryTable[i].br !== 0) {
        mySummaryTable[i].brpercent = `${Number.parseFloat(
          (mySummaryTable[i].brmatch / mySummaryTable[i].br) * 100
        ).toFixed(2)}%(${mySummaryTable[i].brmatch}/${mySummaryTable[i].br}/${
          mySummaryTable[i].total
        })`;
      } else {
        mySummaryTable[i].brpercent = '100%';
      }
    }
    this.buildSummaryTable(mySummaryTable);
  };

  loadBuildManualCTBReport = file => {
    const { ManualSPSheetName } = this.state;
    let workSheet = [];
    this.setState({ loading: true });
    excelUtils.loadExcelFile(file).then(workbook => {
      if (workbook.SheetNames.indexOf(ManualSPSheetName) >= 0) {
        workSheet = excelUtils.excelSheetToJson(workbook, ManualSPSheetName);

        workSheet = workSheet.map(row => {
          const appleCode = this.findAppleVendorCodeBySPName(row.Vendor);
          if (appleCode) {
            row.appleVendorCode = appleCode;
            row.SUPPLIER = row.APN + ':' + appleCode;
            const SUPPLIER = row.SUPPLIER;
            for (let i = 0; i < commodityMapping.length; i++) {
              if (SUPPLIER === commodityMapping[i].appleKey) {
                row.Commodity = commodityMapping[i].commodity;
                row.Vendor = commodityMapping[i].supplier;
                break;
              }
            }
          } else {
            row.appleVendorCode = 'NO MAPPING';
            row.SUPPLIER = row.APN + ':' + row.Vendor;
          }
          return row;
        });
      }

      this.setState({
        ManualWorkBook: workbook,
        ManualSheetNames: this.getBuildRabbitSheetNameOptions(workbook.SheetNames),
        ManualSupplyDataSheet: workSheet,
        loading: false,
        ManualDataReady: true,
      });
    });

    return false;
  };

  getTableWidth = cols => {
    let width = 300;
    cols.forEach(col => {
      width += col.width;
    });
    return width;
  };

  setModaVisible = () => {
    this.setState({ showModal: false });
  };

  exportDiscrepancyTable = () => {
    const { discrepancySheet } = this.state;

    excelUtils.jsonDataToExcel(
      discrepancySheet,
      'NPI144 Supply Discrepancy.xlsx',
      'N144SPDiscrepancy'
    );
  };

  exportPartAnalysisTable = () => {
    const { partAnalysisDataSource } = this.state;
    excelUtils.jsonDataToExcel(partAnalysisDataSource, 'NPI144 Part Analysis.xlsx', 'Sheet0');
  };

  exportSummaryTable = () => {
    const { summaryTableSheet } = this.state;
    excelUtils.jsonDataToExcel(
      summaryTableSheet,
      'NPI144 Supply Discrepancy Summary.xlsx',
      'Sheet0'
    );
  };

  filterAPN = value => {
    const { originalDisSheet, showFullDisTable } = this.state;
    if (value == null || value === '') {
      if (showFullDisTable) {
        this.showFullDiscrepancy();
      } else {
        this.filterShowOnlyDiscrepancy();
      }
      return;
    }

    const discrepancySheet = originalDisSheet.filter(row => row.APN === value);
    this.setState({
      discrepancySheet,
    });
  };

  filterShowOnlyDiscrepancy = () => {
    const { originalDisSheet } = this.state;
    const discrepancySheet = originalDisSheet.filter(
      row => row['DIS-COUNT'] > 0 || row['DIS-COUNT'] < 0
    );
    this.setState({
      discrepancySheet,
      showFullDisTable: false,
    });
  };

  showFullDiscrepancy = () => {
    const { originalDisSheet } = this.state;

    this.setState({
      discrepancySheet: JSON.parse(JSON.stringify(originalDisSheet)),
      showFullDisTable: true,
    });
  };

  findAppleVendorCodeBySPName(spName) {
    let appleVendorCode = false;
    for (let i = 0; i < VendorCodeMapping.length; i++) {
      if (VendorCodeMapping[i].SPName === spName) {
        appleVendorCode = VendorCodeMapping[i].AppleCode;
        break;
      }
    }
    return appleVendorCode;
  }

  render() {
    const {
      fileList,
      BuildRabbitSheetNames,
      BuildRabbitSupplyColumns,
      manualFileList,
      ManualSheetNames,
      ManualSupplyColumns,
      loading,
      ManualSPSheetName,
      BRSPSheetName,
      discrepancyColumns,
      discrepancySheet,
      showFileSelectSection,
      BuildRabbitSupplyPlanStartColumn,
      ManualSupplyPlanStartColumn,
      originalDataTableColumns,
      originalDataTableSheet,
      showSummaryTable,
      summaryTableColumns,
      summaryTableSheet,
      BRDataReady,
      ShipCommitReady,
      ManualDataReady,
      showFullDisTable,
      showPartAnalysisTable,
      partAnalysisDataSource,
      partAnalysisColumns,
    } = this.state;

    return (
      <div style={{ backgroundColor: 'white' }}>
        <Spin spinning={loading}>
          {showFileSelectSection ? (
            <div>
              <Card
                size="small"
                title="Build Rabbit N144 Template"
                style={{ width: '100%', marginBottom: '20px' }}
              >
                <Row>
                  <Col span={8}>
                    <Upload
                      beforeUpload={this.loadBuildRabbitTemplate}
                      fileList={fileList}
                      accept=".xls,.xlsx"
                    >
                      <Button>
                        <Icon type="upload" /> Select Build Rabbit Supply Plan
                      </Button>
                    </Upload>
                  </Col>
                  <Col span={3}>
                    <span>Supply Plan Table</span>
                  </Col>
                  <Col span={4}>
                    <Select
                      style={{ width: '100%' }}
                      value={BRSPSheetName}
                      onChange={this.onBuildRabbitTemplateSupplyPlanSheetChange}
                    >
                      {BuildRabbitSheetNames}
                    </Select>
                  </Col>
                  <Col span={3}>
                    <span>Start Weeks</span>
                  </Col>
                  <Col span={4}>
                    <Select
                      style={{ width: '100%' }}
                      onChange={this.onBuildRabbitTemplateSupplyPlanColumnChange}
                      value={BuildRabbitSupplyPlanStartColumn}
                    >
                      {BuildRabbitSupplyColumns}
                    </Select>
                  </Col>
                </Row>
              </Card>
              <Card
                size="small"
                title="Ship Commit Update(NightOwl version)"
                style={{ width: '100%', marginBottom: '20px' }}
              >
                <Row>
                  <Col span={8}>
                    <Upload
                      beforeUpload={this.loadShipCommitNightOwlFile}
                      fileList={fileList}
                      accept=".xls,.xlsx"
                    >
                      <Button>
                        <Icon type="upload" /> Select Night Owl Ship Commit
                      </Button>
                    </Upload>
                  </Col>
                </Row>
              </Card>
              <Card
                size="small"
                title="Ship Commit Update"
                style={{ width: '100%', marginBottom: '20px' }}
              >
                <Row>
                  <Col span={8}>
                    <Upload
                      beforeUpload={this.loadShipCommitFile}
                      fileList={fileList}
                      accept=".xls,.xlsx"
                    >
                      <Button>
                        <Icon type="upload" /> Select Ship Commit file (Do not use this, as the
                        report have problem)
                      </Button>
                    </Upload>
                  </Col>
                </Row>
              </Card>
              <Card
                size="small"
                title="Manual N144 CTB Report"
                style={{ width: '100%', marginBottom: '20px' }}
              >
                <Row>
                  <Col span={8}>
                    <Upload
                      beforeUpload={this.loadBuildManualCTBReport}
                      fileList={manualFileList}
                      accept=".xls,.xlsx"
                    >
                      <Button>
                        <Icon type="upload" /> Select Manual N144 CTB report
                      </Button>
                    </Upload>
                  </Col>
                  <Col span={3}>
                    <span>Supply Plan Table</span>
                  </Col>
                  <Col span={4}>
                    <Select
                      style={{ width: '100%' }}
                      value={ManualSPSheetName}
                      onChange={this.onManualSheetChange}
                    >
                      {ManualSheetNames}
                    </Select>
                  </Col>
                  <Col span={3}>
                    <span>Start Weeks</span>
                  </Col>
                  <Col span={4}>
                    <Select
                      style={{ width: '100%' }}
                      onChange={this.onManualSupplyPlanColumnChange}
                      value={ManualSupplyPlanStartColumn}
                    >
                      {ManualSupplyColumns}
                    </Select>
                  </Col>
                </Row>
              </Card>
            </div>
          ) : null}
          {BRDataReady && ShipCommitReady && ManualDataReady ? (
            <Card
              size="small"
              title="Build Rabbit Supply Compare with Manual Supply Plan"
              style={{ width: '100%', marginBottom: '20px' }}
            >
              <Row style={{ marginBottom: '20px' }}>
                <Col span={18}>
                  {!showFileSelectSection && showFullDisTable ? (
                    <Button
                      onClick={this.filterShowOnlyDiscrepancy}
                      style={{ marginRight: '10px' }}
                    >
                      Only Dis-Match
                    </Button>
                  ) : !showFileSelectSection && !showFullDisTable ? (
                    <Button onClick={this.showFullDiscrepancy} style={{ marginRight: '10px' }}>
                      Full Result
                    </Button>
                  ) : null}
                  <Button onClick={this.startFullComparation} style={{ marginRight: '10px' }}>
                    Compare Full
                  </Button>
                  <Button onClick={this.startBRComparation} style={{ marginRight: '10px' }}>
                    Compare By BR
                  </Button>
                  <Button onClick={this.startManualComparation} style={{ marginRight: '10px' }}>
                    Compare by Manual
                  </Button>
                  <Button onClick={this.initializeData} style={{ marginRight: '10px' }}>
                    Initialize
                  </Button>
                  <Button onClick={this.updateSummaryTable} style={{ marginRight: '10px' }}>
                    Summary Table
                  </Button>
                  <Button onClick={this.partListAnalysis} style={{ marginRight: '10px' }}>
                    Part Analysis
                  </Button>
                </Col>
                <Col span={3}>
                  <Search
                    placeholder="filter APN"
                    onSearch={this.filterAPN}
                    enterButton
                    width="200px"
                  />
                </Col>
                {showSummaryTable ? (
                  <Col span={3}>
                    <Button onClick={this.exportSummaryTable} style={{ marginRight: '10px' }}>
                      Export Excel
                    </Button>
                  </Col>
                ) : showPartAnalysisTable ? (
                  <Col span={3}>
                    <Button onClick={this.exportPartAnalysisTable} style={{ marginRight: '10px' }}>
                      Export Excel
                    </Button>
                  </Col>
                ) : !showFileSelectSection ? (
                  <Col span={3}>
                    <Button onClick={this.exportDiscrepancyTable} style={{ marginRight: '10px' }}>
                      Export Excel
                    </Button>
                  </Col>
                ) : null}
              </Row>
              {showSummaryTable ? (
                <Table
                  dataSource={summaryTableSheet}
                  columns={summaryTableColumns}
                  style={{ backgroundColor: 'white', fontSize: '10px' }}
                  bordered={true}
                  pagination={false}
                  size="sm"
                />
              ) : showPartAnalysisTable ? (
                <Table
                  dataSource={partAnalysisDataSource}
                  columns={partAnalysisColumns}
                  scroll={{
                    x: this.getTableWidth(discrepancyColumns),
                    y: document.body.clientHeight - 200,
                  }}
                  style={{ backgroundColor: 'white', fontSize: '10px' }}
                  bordered={true}
                  pagination={false}
                  size="sm"
                />
              ) : (
                <Table
                  dataSource={discrepancySheet}
                  columns={discrepancyColumns}
                  scroll={{
                    x: this.getTableWidth(discrepancyColumns),
                    y: document.body.clientHeight - 200,
                  }}
                  style={{ backgroundColor: 'white', fontSize: '10px' }}
                  bordered={true}
                  pagination={false}
                  size="sm"
                  onRow={record => ({
                    onDoubleClick: () => {
                      this.findOriginalData(record);
                    },
                  })}
                />
              )}
            </Card>
          ) : null}

          <Modal
            title="Discrepancy Modal"
            centered
            visible={this.state.showModal}
            onOk={() => this.setModaVisible()}
            onCancel={() => this.setModaVisible()}
            width={document.body.clientWidth * 0.9}
          >
            <Table
              dataSource={originalDataTableSheet}
              columns={originalDataTableColumns}
              scroll={{ x: this.getTableWidth(discrepancyColumns) }}
              style={{ backgroundColor: 'white', fontSize: '10px' }}
              bordered={true}
              pagination={false}
              size="sm"
            />
          </Modal>
        </Spin>
      </div>
    );
  }
}

export default MpsValidation;
