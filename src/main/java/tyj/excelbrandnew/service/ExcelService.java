package tyj.excelbrandnew.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mybatis.spring.SqlSessionTemplate;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import tyj.excelbrandnew.entity.DataInfo;
import tyj.excelbrandnew.entity.ExcelFile;
import tyj.excelbrandnew.entity.HeaderInfo;
import tyj.excelbrandnew.utils.CellStyleUtil;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.HashMap;
import java.util.Map;

@Service
public class ExcelService {

    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);

    private final SqlSessionTemplate sqlSessionTemplate;

    public ExcelService(SqlSessionTemplate sqlSessionTemplate) {
        this.sqlSessionTemplate = sqlSessionTemplate;
    }

    public void processExcel(MultipartFile file, int startHeaderRow, int endHeaderRow, int startDataRow, int endDataRow, int columnCount, String templateName) throws IOException {
        Workbook workbook = null;
        try {
            log.info("Processing Excel file: {}, startHeaderRow: {}, endHeaderRow: {}, startDataRow: {}, endDataRow: {}, columnCount: {}, templateName: {}",
                    file.getOriginalFilename(), startHeaderRow, endHeaderRow, startDataRow, endDataRow, columnCount, templateName);
            workbook = new XSSFWorkbook(file.getInputStream());

            // 先将文件信息存入excel_file表
            ExcelFile excelFile = new ExcelFile();
            excelFile.setFileName(file.getOriginalFilename());
            sqlSessionTemplate.insert("ExcelFileMapper.insertExcelFile", excelFile);
            sqlSessionTemplate.flushStatements(); // 获取生成的ID
            Long excelFileId = excelFile.getId();

            if (excelFileId == null) {
                throw new RuntimeException("获取Excel文件ID失败");
            }

            Map<String, Long> headerIdMap = new HashMap<>();

            // 遍历所有的 sheet
            for (int sheetNo = 0; sheetNo < workbook.getNumberOfSheets(); sheetNo++) {
                Sheet sheet = workbook.getSheetAt(sheetNo);

                // 处理表头信息
                for (int rowIndex = startHeaderRow; rowIndex <= endHeaderRow; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (cell != null) {
                                String columnName = cell.toString();
                                String cellStyle = CellStyleUtil.serializeCellStyle(cell.getCellStyle());

                                // 检查是否有合并单元格
                                boolean isMerged = false;
                                int mergeStartColumn = -1;
                                int mergeEndColumn = -1;
                                int mergeStartRow = -1;
                                int mergeEndRow = -1;

                                for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                                    if (mergedRegion.isInRange(rowIndex, colIndex)) {
                                        isMerged = true;
                                        mergeStartColumn = mergedRegion.getFirstColumn();
                                        mergeEndColumn = mergedRegion.getLastColumn();
                                        mergeStartRow = mergedRegion.getFirstRow();
                                        mergeEndRow = mergedRegion.getLastRow();
                                        break;
                                    }
                                }

                                // 插入表头数据
                                HeaderInfo headerInfo = new HeaderInfo();
                                headerInfo.setExcelFileId(excelFileId);
                                headerInfo.setColumnName(columnName);
                                headerInfo.setColumnIndex(colIndex);
                                headerInfo.setRowIndex(rowIndex);
                                headerInfo.setCellStyle(cellStyle);
                                headerInfo.setMerged(isMerged);
                                headerInfo.setMergeStartColumn(mergeStartColumn);
                                headerInfo.setMergeEndColumn(mergeEndColumn);
                                headerInfo.setMergeStartRow(mergeStartRow);
                                headerInfo.setMergeEndRow(mergeEndRow);

                                sqlSessionTemplate.insert("HeaderInfoMapper.insertHeaderInfo", headerInfo);
                                sqlSessionTemplate.flushStatements(); // 获取生成的ID
                                Long headerId = headerInfo.getId();

                                // 将生成的 header_id 存入 map 中，使用 rowIndex 和 colIndex 作为键
                                headerIdMap.put(sheetNo + "_" + rowIndex + "_" + colIndex, headerId);
                            }
                        }
                    }
                }

                // 处理数据
                for (int rowIndex = startDataRow; rowIndex <= endDataRow; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int colIndex = 0; colIndex < columnCount; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            if (cell != null) {
                                String cellValue = cell.toString();
                                String cellStyle = CellStyleUtil.serializeCellStyle(cell.getCellStyle());

                                // 检查是否有合并单元格
                                boolean isMerged = false;
                                int mergeStartColumn = -1;
                                int mergeEndColumn = -1;
                                int mergeStartRow = -1;
                                int mergeEndRow = -1;

                                for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                                    CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                                    if (mergedRegion.isInRange(rowIndex, colIndex)) {
                                        isMerged = true;
                                        mergeStartColumn = mergedRegion.getFirstColumn();
                                        mergeEndColumn = mergedRegion.getLastColumn();
                                        mergeStartRow = mergedRegion.getFirstRow();
                                        mergeEndRow = mergedRegion.getLastRow();
                                        break;
                                    }
                                }

                                // 获取对应的 header_id
                                Long headerId = headerIdMap.get(sheetNo + "_" + startHeaderRow + "_" + colIndex);
                                if (headerId == null) {
                                    throw new RuntimeException("找不到对应的header_id");
                                }

                                // 插入数据
                                DataInfo dataInfo = new DataInfo();
                                dataInfo.setExcelFileId(excelFileId);
                                dataInfo.setHeaderId(headerId);
                                dataInfo.setSheetNo(sheetNo);
                                dataInfo.setRowIndex(rowIndex);
                                dataInfo.setColumnIndex(colIndex);
                                dataInfo.setCellValue(cellValue);
                                dataInfo.setCellStyle(cellStyle);
                                dataInfo.setMerged(isMerged);
                                dataInfo.setMergeStartColumn(mergeStartColumn);
                                dataInfo.setMergeEndColumn(mergeEndColumn);
                                dataInfo.setMergeStartRow(mergeStartRow);
                                dataInfo.setMergeEndRow(mergeEndRow);

                                sqlSessionTemplate.insert("DataInfoMapper.insertDataInfo", dataInfo);
                                sqlSessionTemplate.flushStatements();
                            }
                        }
                    }
                }
            }

            log.info("Successfully processed Excel file: {}", file.getOriginalFilename());
        } catch (Exception e) {
            log.error("Error processing Excel file", e);
            throw new RuntimeException("处理Excel文件时出错: " + e.getMessage(), e);
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }
    }


}
