/*
 * Copyright (c) 2018, 吴汶泽 (wenzewoo@gmail.com).
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.wuwenze.poi.xlsx;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.wuwenze.poi.config.Options;
import com.wuwenze.poi.convert.ReadConverter;
import com.wuwenze.poi.exception.ExcelKitEncounterNoNeedXmlException;
import com.wuwenze.poi.exception.ExcelKitReadConverterException;
import com.wuwenze.poi.exception.ExcelKitRuntimeException;
import com.wuwenze.poi.handler.ExcelReadHandler;
import com.wuwenze.poi.pojo.ExcelErrorField;
import com.wuwenze.poi.pojo.ExcelMapping;
import com.wuwenze.poi.pojo.ExcelProperty;
import com.wuwenze.poi.util.*;
import com.wuwenze.poi.validator.Validator;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @author wuwenze
 */
public class ExcelXlsxReader extends DefaultHandler {

  private Integer mCurrentSheetIndex = -1, mCurrentRowIndex = 0, mCurrentCellIndex = 0;
  private ExcelCellType mNextCellType = ExcelCellType.STRING;
  private String mCurrentCellRef, mPreviousCellRef, mMaxCellRef;
  private SharedStringsTable mSharedStringsTable;
  private String mPreviousCellValue;
  private StylesTable mStylesTable;
  private Boolean mNextIsString = false;
  private Short mFormatIndex;
  private String mFormatString;
  private Boolean errorSheet = false;

  private final ExcelMapping mExcelMapping;
  private final ExcelReadHandler mExcelReadHandler;
  private final Class<? extends Object> mEntityClass;
  private final List<Object> mExcelRowObjectData = Lists.newArrayList();
  private final List<String> headTitleList = Lists.newArrayList();
  private final Map<String, ExcelProperty> excelPropertyMap = new HashMap<String, ExcelProperty>();
  private Integer mBeginReadRowIndex = Const.XLSX_DEFAULT_BEGIN_READ_ROW_INDEX;
  private final Object mEmptyCellValue = Const.XLSX_DEFAULT_EMPTY_CELL_VALUE;

  private final DataFormatter formatter = new DataFormatter();


  public ExcelXlsxReader(Class<? extends Object> entityClass,//
                         ExcelMapping excelMapping, //
                         ExcelReadHandler excelReadHandler) {
    this(entityClass, excelMapping, null, excelReadHandler);
  }

  public ExcelXlsxReader(Class<? extends Object> entityClass,//
                         ExcelMapping excelMapping, //
                         Integer beginReadRowIndex,//
                         ExcelReadHandler excelReadHandler) {
    mEntityClass = entityClass;
    mExcelMapping = excelMapping;
    if (null != beginReadRowIndex) {
      mBeginReadRowIndex = beginReadRowIndex;
    }
    mExcelReadHandler = excelReadHandler;
  }

  public void process(String fileName) throws ExcelKitRuntimeException {
    try {
      processAll(OPCPackage.open(fileName));
    } catch (Exception e) {
      throw new ExcelKitRuntimeException("Only .xlsx formatted files are supported.", e);
    }
  }

  public void process(InputStream in) throws ExcelKitRuntimeException {
    try {
      processAll(OPCPackage.open(in));
    } catch (Exception e) {
      throw new ExcelKitRuntimeException("Only .xlsx formatted files are supported.", e);
    }
  }

  private void processAll(OPCPackage pkg)
      throws IOException, OpenXML4JException, SAXException {
    XSSFReader xssfReader = new XSSFReader(pkg);
    mStylesTable = xssfReader.getStylesTable();
    SharedStringsTable sst = xssfReader.getSharedStringsTable();
    XMLReader parser = this.fetchSheetParser(sst);
    Iterator<InputStream> sheets = xssfReader.getSheetsData();
    while (sheets.hasNext()) {
      mCurrentRowIndex = 0;
      mCurrentSheetIndex++;
      InputStream sheet = sheets.next();
      InputSource sheetSource = new InputSource(sheet);
      parser.parse(sheetSource);
      sheet.close();
    }
    pkg.close();
  }

  public void process(String fileName, int sheetIndex) throws ExcelKitRuntimeException {
    try {
      this.processBySheet(sheetIndex, OPCPackage.open(fileName));
    } catch (Exception e) {
      throw new ExcelKitRuntimeException("Only .xlsx formatted files are supported.", e);
    }
  }

  public void process(InputStream in, int sheetIndex) throws ExcelKitRuntimeException {
    try {
      this.processBySheet(sheetIndex, OPCPackage.open(in));
    } catch (Exception e) {
      throw new ExcelKitRuntimeException("Only .xlsx formatted files are supported.", e);
    }
  }

  private void processBySheet(int sheetIndex, OPCPackage pkg)
      throws IOException, OpenXML4JException, SAXException {
    XSSFReader r = new XSSFReader(pkg);
    SharedStringsTable sst = r.getSharedStringsTable();

    XMLReader parser = fetchSheetParser(sst);

    // 根据 rId# 或 rSheet# 查找sheet
    InputStream sheet = r.getSheet(Const.SAX_RID_PREFIX + (sheetIndex + 1));
    mCurrentSheetIndex++;
    InputSource sheetSource = new InputSource(sheet);
    try {
      parser.parse(sheetSource);
    } catch (ExcelKitEncounterNoNeedXmlException e) {
      sheet = r.getSheet(Const.SAX_RID_PREFIX + (sheetIndex + 3));
      sheetSource = new InputSource(sheet);
      parser.parse(sheetSource);
    }
    sheet.close();
    pkg.close();
  }

  @Override
  public void startElement(
      String uri, String localName, String name, Attributes attributes) {
    if ("sst".equals(name) || "styleSheet".equals(name)) {
      throw new ExcelKitEncounterNoNeedXmlException();
    }
    // c => 单元格
    if (Const.SAX_C_ELEMENT.equals(name)) {
      String ref = attributes.getValue(Const.SAX_R_ATTR);
      // 前一个单元格的位置
      mPreviousCellRef = null == mPreviousCellRef ? "A"+(mCurrentRowIndex+1) : mCurrentCellRef;
      // 当前单元格的位置
      mCurrentCellRef = ref;
      // Figure out if the value is an index in the SST
      String cellType = attributes.getValue(Const.SAX_T_ELEMENT);
      String cellStyleStr = attributes.getValue(Const.SAX_S_ATTR_VALUE);
      mNextIsString = (null != cellType && cellType.equals(Const.SAX_S_ATTR_VALUE));
      // 设定单元格类型
      this.setNextCellType(cellType, cellStyleStr);
    }
    mPreviousCellValue = "";
  }


  @Override
  public void endElement(String uri, String localName, String name) {
    // Process the last contents as required.
    // Do now, as characters() may be called more than once
    if (mNextIsString) {
      int index = Integer.parseInt(mPreviousCellValue);
      mPreviousCellValue = new XSSFRichTextString(mSharedStringsTable.getEntryAt(index))
          .toString();
      mNextIsString = false;
    }

    // 处理单元格数据
    if (Const.SAX_C_ELEMENT.equals(name)) {
      String value = this.getCellValue(mPreviousCellValue.trim());
      // 空值补齐(中)
      if (!mCurrentCellRef.equals(mPreviousCellRef)) {
        int num =POIUtil.countNullCell(mCurrentCellRef, mPreviousCellRef);
        if(mPreviousCellRef.startsWith("A")&&mExcelRowObjectData.size()==0) {
             num++;
        }
        for (int i = 0; i < num; i++) {
          mExcelRowObjectData.add(mCurrentCellIndex, mEmptyCellValue);
          mCurrentCellIndex++;
        }
      }
      mExcelRowObjectData.add(mCurrentCellIndex, value);
      mCurrentCellIndex++;
    }
    // 如果标签名称为 row ，这说明已到行尾，通知回调处理当前行的数据
    else if (Const.SAX_ROW_ELEMENT.equals(name)) {
      if (mCurrentRowIndex == 0) {
        mMaxCellRef = mCurrentCellRef;
      }
      // 空值补齐(后)
      if (null != mMaxCellRef && null != mCurrentCellRef) {
        for (int i = 0; i <= POIUtil.countNullCell(mMaxCellRef, mCurrentCellRef); i++) {
          mExcelRowObjectData.add(mCurrentCellIndex, mEmptyCellValue);
          mCurrentCellIndex++;
        }
      }
      try {
        this.performVerificationAndProcessFlowRow();
      } catch (Exception e) {
        e.printStackTrace();
      } finally {
        mExcelRowObjectData.clear();
        mCurrentRowIndex++;
        mCurrentCellIndex = 0;
        mPreviousCellRef = null;
        mCurrentCellRef = null;
      }
    }
  }

  @Override
  public void characters(char[] chars, int start, int length) {
    mPreviousCellValue = mPreviousCellValue.concat(new String(chars, start, length));
  }

  private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
    XMLReader parser = XMLReaderFactory.createXMLReader(Const.SAX_PARSER_CLASS);
    mSharedStringsTable = sst;
    parser.setContentHandler(this);
    return parser;
  }

  enum ExcelCellType {
    BOOL, ERROR, FORMULA, INLINESTR, STRING, NUMBER, DATE, NULL
  }

  private void setNextCellType(String cellType, String cellStyleStr) {
    mNextCellType = ExcelCellType.STRING;
    mFormatIndex = -1;
    mFormatString = null;

    if ("b".equals(cellType)) {
      mNextCellType = ExcelCellType.BOOL;
    } else if ("e".equals(cellType)) {
      mNextCellType = ExcelCellType.ERROR;
    } else if ("inlineStr".equals(cellType)) {
      mNextCellType = ExcelCellType.INLINESTR;
    } else if ("s".equals(cellType)) {
      mNextCellType = ExcelCellType.STRING;
    } else if ("str".equals(cellType)) {
      mNextCellType = ExcelCellType.FORMULA;
    }
    if (null != cellStyleStr) {
      int styleIndex = Integer.parseInt(cellStyleStr);
      XSSFCellStyle style = mStylesTable.getStyleAt(styleIndex);
      mFormatIndex = style.getDataFormat();
      mFormatString = style.getDataFormatString();
      if ("m/d/yy".equals(mFormatString)) {
        mNextCellType = mNextCellType.DATE;
        mFormatString = "yyyy-MM-dd hh:mm:ss.SSS";
      }
      if (null == mFormatString) {
        mNextCellType = mNextCellType.NULL;
        mFormatString = BuiltinFormats.getBuiltinFormat(mFormatIndex);
      }
    }
  }

  private String getCellValue(String value) {
    String thisStr;
    if(value.trim()=="") {
      return "";
    }
    try {
      switch (mNextCellType) {
        case BOOL:
          return value.charAt(0) == '0' ? "FALSE" : "TRUE";
        case ERROR:
          return "\"ERROR:" + value + '"';
        case FORMULA:
          return '"' + value + '"';
        case INLINESTR:
          return new XSSFRichTextString(value).toString();
        case STRING:
          return String.valueOf(value);
        case NUMBER:
          if (mFormatString != null) {
            thisStr = formatter.formatRawCellContents(Double.parseDouble(value), mFormatIndex, mFormatString).trim();
          } else {
            thisStr = value;
          }

          thisStr = thisStr.replace("_", "").trim();
          break;
        case DATE:
          thisStr = formatter.formatRawCellContents(Double.parseDouble(value), mFormatIndex, mFormatString);
          // 对日期字符串作特殊处理
          thisStr = thisStr.replace(" ", "T");
          break;
        default:
          thisStr = "";
          break;
      }
    }catch (Exception e ){
      thisStr = value;
    }
    return thisStr;
  }

  private final static String CHECK_MAP_KEY_OF_VALUE = "CELL_VALUE";
  private final static String CHECK_MAP_KEY_OF_ERROR = "CELL_ERROR";

  private void performVerificationAndProcessFlowRow() throws Exception {
    if(mCurrentRowIndex ==0){
      errorSheet = false;
      headTitleList.clear(); //先清除后添加
      for (int i = 0; i < mExcelRowObjectData.size(); i++) {
        headTitleList.add(i, (String)mExcelRowObjectData.get(i));
      }
      List<ExcelProperty> propertyList = mExcelMapping.getPropertyList();
      for(ExcelProperty pro:propertyList) {
        excelPropertyMap.put(pro.getColumn(),pro);
      }
      for (int i = 0; i < headTitleList.size(); i++) { //匹配是否要导入的sheet
        ExcelProperty property = excelPropertyMap.get(headTitleList.get(i));
        if (property == null) {
          errorSheet = true;
        }
      }
    }
    if(errorSheet || mExcelRowObjectData.size() !=headTitleList.size()) {
      return;
    }
    if (mCurrentRowIndex >= mBeginReadRowIndex) {
      if (!this.rowObjectDataIsAllEmptyCellValue()) {
        Object entity = mEntityClass.newInstance();
        List<ExcelErrorField> errorFields = Lists.newArrayList();
        for (int i = 0; i < headTitleList.size(); i++) {
          //跳过空数据
          ExcelProperty property = excelPropertyMap.get(headTitleList.get(i));
          if(property==null) {
            continue;
          }
          Map<String, Object> checkAndConvertPropertyRetMap = this.checkAndConvertProperty(i, property, mExcelRowObjectData.get(i));
          Object errorFieldObject = checkAndConvertPropertyRetMap.get(
              ExcelXlsxReader.CHECK_MAP_KEY_OF_ERROR);
          if (null != errorFieldObject) {
            errorFields.add((ExcelErrorField) errorFieldObject);
          }
          if (errorFields.isEmpty()) {
            Object propertyValue = checkAndConvertPropertyRetMap.get(
                ExcelXlsxReader.CHECK_MAP_KEY_OF_VALUE);
            BeanUtil.setComplexProperty(entity, property.getName(), propertyValue);
          }
        }
        if (errorFields.isEmpty()) {
          mExcelReadHandler.onSuccess(mCurrentSheetIndex, mCurrentRowIndex, entity);
          return;
        }
        mExcelReadHandler.onError(mCurrentSheetIndex, mCurrentRowIndex, errorFields);
      }
    }
  }

  private boolean rowObjectDataIsAllEmptyCellValue() {
    int emptyObjectCount = 0;
    for (Object excelRowObjectData : mExcelRowObjectData) {
      if ((null == excelRowObjectData) //
          || excelRowObjectData.equals(mEmptyCellValue) //
          || ValidatorUtil.isEmpty((String) excelRowObjectData)) {
        emptyObjectCount++;
      }
    }
    return emptyObjectCount == mExcelRowObjectData.size();
  }

  private Map<String, Object> checkAndConvertProperty(Integer cellIndex,
      ExcelProperty property,
      Object propertyValue) {

     if(null == propertyValue || ValidatorUtil.isEmpty((String) propertyValue) || Const.XLSX_DEFAULT_EMPTY_CELL_VALUE.equals(propertyValue)) {
        // required
        Boolean required = property.getRequired();
        if (null != required && required) {
            return this.buildCheckAndConvertPropertyRetMap(cellIndex, property, propertyValue, "单元格的值必须填写");
        }

        // empty cell doesn't need to check anymore
        return this.buildCheckAndConvertPropertyRetMap(cellIndex, property, null, null);
    }

    // maxLength
    Integer maxLength = property.getMaxLength();
    if (-1 != maxLength) {
      if (String.valueOf(propertyValue).length() > maxLength) {
        return this.buildCheckAndConvertPropertyRetMap(cellIndex, property, propertyValue, "超过最大长度: " + maxLength);
      }
    }

    // dateFormat
    String dateFormat = property.getDateFormat();
    if (!ValidatorUtil.isEmpty(dateFormat)) {
      try {
        // 时间格式转换后，直接返回。
        Date parseDateValue = DateUtil.parse(dateFormat, propertyValue);
        return this.buildCheckAndConvertPropertyRetMap(cellIndex, property, parseDateValue, null);
      } catch (Exception e) {
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, propertyValue, "时间格式解析失败 [" + dateFormat + "]");
      }
    }

    // options
    Options options = property.getOptions();
    if (null != options) {
      Object[] values = options.get();
      if (null != values && values.length > 0) {
        boolean containInOptions = false;
        for (Object value : values) {
          if (propertyValue.equals(value)) {
            containInOptions = true;
            break;
          }
        }
        if (!containInOptions) {
          return this.buildCheckAndConvertPropertyRetMap(//
              cellIndex, property, propertyValue, "[" + propertyValue + "]不是规定的下拉框的值");
        }
      }
    }

    // regularExp
    String regularExp = property.getRegularExp();
    if (!ValidatorUtil.isEmpty(regularExp)) {
      if (!RegexUtil.isMatches(regularExp, propertyValue)) {
        String regularExpMessage = property.getRegularExpMessage();
        String validErrorMessage = !ValidatorUtil.isEmpty(regularExpMessage) ?
            regularExpMessage : "正则表达式校验失败 [" + regularExp + "]";
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, propertyValue, validErrorMessage);
      }
    }

    // validator
    Validator validator = property.getValidator();
    if (null != validator) {
      String validErrorMessage = validator.valid(propertyValue);
      if (null != validErrorMessage) {
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, propertyValue, validErrorMessage);
      }
    }

    // readConverterExp && readConverter (按照优先级处理)
    String readConverterExp = property.getReadConverterExp();
    ReadConverter readConverter = property.getReadConverter();
    if (!ValidatorUtil.isEmpty(readConverterExp)) {
      try {
        Object convertPropertyValue = POIUtil.convertByExp(propertyValue, readConverterExp);
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, convertPropertyValue, null);
      } catch (Exception e) {
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, propertyValue, "由于readConverterExp表达式的值不规范导致转换失败");
      }
    } else if (null != readConverter) {
      try {
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, readConverter.convert(propertyValue), null);
      } catch (ExcelKitReadConverterException e) {
        return this.buildCheckAndConvertPropertyRetMap(//
            cellIndex, property, propertyValue, e.getMessage());
      }
    }
    return this.buildCheckAndConvertPropertyRetMap(cellIndex, property, propertyValue, null);
  }

  private Map<String, Object> buildCheckAndConvertPropertyRetMap(Integer cellIndex,
                                                                 ExcelProperty property,//
                                                                 Object propertyValue, String validErrorMessage) {
    Map<String, Object> resultMap = Maps.newHashMap();
    resultMap.put(ExcelXlsxReader.CHECK_MAP_KEY_OF_VALUE, propertyValue);
    if (null != validErrorMessage) {
      resultMap.put(ExcelXlsxReader.CHECK_MAP_KEY_OF_ERROR, ExcelErrorField.builder()//
          .cellIndex(cellIndex)//
          .column(property.getColumn())//
          .name(property.getName())//
          .errorMessage(validErrorMessage)//
          .build());
    }
    return resultMap;
  }
}
