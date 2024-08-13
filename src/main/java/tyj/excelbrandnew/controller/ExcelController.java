package tyj.excelbrandnew.controller;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import tyj.excelbrandnew.service.ExcelExportService;
import tyj.excelbrandnew.service.ExcelService;

import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {
    private static final Logger log = LoggerFactory.getLogger(ExcelController.class);

    @Autowired
    private ExcelService excelService;

    @Autowired
    private ExcelExportService excelExportService;

    @PostMapping("/upload")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file,
                                              @RequestParam("startHeaderRow") int startHeaderRow,
                                              @RequestParam("endHeaderRow") int endHeaderRow,
                                              @RequestParam("startDataRow") int startDataRow,
                                              @RequestParam("endDataRow") int endDataRow,
                                              @RequestParam("columnCount") int columnCount,
                                              @RequestParam("templateName") String templateName) {
        try {
            excelService.processExcel(file, startHeaderRow, endHeaderRow, startDataRow, endDataRow, columnCount, templateName);
            log.info("文件上传并处理成功");
            return ResponseEntity.ok("文件上传并处理成功");
        } catch (Exception e) {
            log.error("处理文件时出错: ", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("处理文件时出错: " + e.getMessage());
        }
    }

    @GetMapping("/download")
    public ResponseEntity<Resource> downloadExcel(@RequestParam("excelFileId") Long excelFileId) {
        try {
            Workbook workbook = excelExportService.exportExcel(excelFileId);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();

            ByteArrayResource resource = new ByteArrayResource(outputStream.toByteArray());

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"exported.xlsx\"")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .contentLength(resource.contentLength())
                    .body(resource);
        } catch (Exception e) {
            log.error("处理文件时出错: ", e);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }
}
