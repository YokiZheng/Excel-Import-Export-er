package tyj.excelbrandnew.entity;

import lombok.Data;

@Data
public class HeaderInfo {
    private Long id;              // 表头信息唯一标识符
    private Long excelFileId;     // 关联的 Excel 文件的 ID
    private String columnName;    // 列名
    private int columnIndex;      // 列索引
    private int rowIndex;         // 行索引
    private String cellStyle;     // 单元格样式（序列化后的字符串）

    private boolean isMerged;     // 是否为合并单元格
    private Integer mergeStartColumn;  // 合并单元格的起始列
    private Integer mergeEndColumn;    // 合并单元格的结束列
    private Integer mergeStartRow;     // 合并单元格的起始行
    private Integer mergeEndRow;       // 合并单元格的结束行
}
