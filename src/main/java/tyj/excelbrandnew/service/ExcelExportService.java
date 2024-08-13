package tyj.excelbrandnew.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mybatis.spring.SqlSessionTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import tyj.excelbrandnew.entity.DataInfo;
import tyj.excelbrandnew.entity.HeaderInfo;
import tyj.excelbrandnew.utils.CellStyleUtil;

import java.util.List;

@Service
public class ExcelExportService {

    @Autowired
    private SqlSessionTemplate sqlSessionTemplate;

    public Workbook exportExcel(Long excelFileId) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Exported Data");

        // 获取表头信息
        List<HeaderInfo> headerInfos = sqlSessionTemplate.selectList("HeaderInfoMapper.selectHeaderByExcelFileId", excelFileId);

        for (HeaderInfo headerInfo : headerInfos) {
            Row row = sheet.getRow(headerInfo.getRowIndex());
            if (row == null) {
                row = sheet.createRow(headerInfo.getRowIndex());
            }
            Cell cell = row.createCell(headerInfo.getColumnIndex());
            cell.setCellValue(headerInfo.getColumnName());

            if (headerInfo.getCellStyle() != null) {
                CellStyle cellStyle = CellStyleUtil.deserializeCellStyle(workbook, headerInfo.getCellStyle());
                cell.setCellStyle(cellStyle);
            }

            if (headerInfo.isMerged()) {
                CellRangeAddress region = new CellRangeAddress(
                        headerInfo.getMergeStartRow(),
                        headerInfo.getMergeEndRow(),
                        headerInfo.getMergeStartColumn(),
                        headerInfo.getMergeEndColumn()
                );
                addMergedRegionIfNotExists(sheet, region);
            }
        }

        // 获取数据
        List<DataInfo> dataInfos = sqlSessionTemplate.selectList("DataInfoMapper.selectDataByExcelFileId", excelFileId);

        for (DataInfo dataInfo : dataInfos) {
            Row row = sheet.getRow(dataInfo.getRowIndex());
            if (row == null) {
                row = sheet.createRow(dataInfo.getRowIndex());
            }
            Cell cell = row.createCell(dataInfo.getColumnIndex());
            cell.setCellValue(dataInfo.getCellValue());

            if (dataInfo.getCellStyle() != null) {
                CellStyle cellStyle = CellStyleUtil.deserializeCellStyle(workbook, dataInfo.getCellStyle());
                cell.setCellStyle(cellStyle);
            }

            if (dataInfo.isMerged()) {
                CellRangeAddress region = new CellRangeAddress(
                        dataInfo.getMergeStartRow(),
                        dataInfo.getMergeEndRow(),
                        dataInfo.getMergeStartColumn(),
                        dataInfo.getMergeEndColumn()
                );
                addMergedRegionIfNotExists(sheet, region);
            }
        }

        return workbook;
    }

    private void addMergedRegionIfNotExists(Sheet sheet, CellRangeAddress region) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress existingRegion = sheet.getMergedRegion(i);
            if (existingRegion.equals(region)) {
                return; // 如果该区域已经存在，则跳过添加
            }
        }
        sheet.addMergedRegion(region);
    }
}
