package tyj.excelbrandnew.entity;

import lombok.Data;

import java.sql.Timestamp;

@Data
public class ExcelFile {
    private Long id;              // 文件唯一标识符
    private String fileName;      // 文件名称
    private Timestamp uploadTime; //文件上传时间戳
    private String center;        // 中心名称
    private String department;    // 处域名称
    private String projectGroup;  // 项目组名称
}
