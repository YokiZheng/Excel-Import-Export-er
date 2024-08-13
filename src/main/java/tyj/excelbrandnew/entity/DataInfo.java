package tyj.excelbrandnew.entity;

import lombok.Data;

@Data
public class DataInfo {
    private Long id;             // 数据记录唯一标识符
    private Long excelFileId;    // 关联的 Excel 文件的 ID
    private Long headerId;       // 关联的表头信息 ID
    private int sheetNo;         // 所属的 Sheet 编号
    private int rowIndex;        // 行索引
    private int columnIndex;     // 列索引
    private String cellValue;    // 单元格的值
    private String cellStyle;    // 单元格样式（序列化后的字符串）
    private boolean isMerged;    // 是否为合并单元格
    private int mergeStartColumn;  // 合并单元格的开始列
    private int mergeEndColumn;    // 合并单元格的结束列
    private int mergeStartRow;     // 合并单元格的开始行
    private int mergeEndRow;       // 合并单元格的结束行
}
