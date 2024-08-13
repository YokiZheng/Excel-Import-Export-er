package tyj.excelbrandnew.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.Color;

public class CellStyleUtil {
    public static String serializeCellStyle(CellStyle cellStyle) {
        if (cellStyle == null) {
            return "";
        }

        XSSFCellStyle xssfCellStyle = (XSSFCellStyle) cellStyle;
        StringBuilder sb = new StringBuilder();
        sb.append(xssfCellStyle.getFillForegroundColor()).append(",");
        sb.append(xssfCellStyle.getFillBackgroundColor()).append(",");
        sb.append(xssfCellStyle.getFillPattern().getCode()).append(",");
        sb.append(xssfCellStyle.getBorderBottom().getCode()).append(",");
        sb.append(xssfCellStyle.getBorderLeft().getCode()).append(",");
        sb.append(xssfCellStyle.getBorderRight().getCode()).append(",");
        sb.append(xssfCellStyle.getBorderTop().getCode()).append(",");
        sb.append(xssfCellStyle.getBottomBorderColor()).append(",");
        sb.append(xssfCellStyle.getLeftBorderColor()).append(",");
        sb.append(xssfCellStyle.getRightBorderColor()).append(",");
        sb.append(xssfCellStyle.getTopBorderColor()).append(",");
        sb.append(xssfCellStyle.getAlignment().getCode()).append(",");
        sb.append(xssfCellStyle.getVerticalAlignment().getCode()).append(",");
        sb.append(xssfCellStyle.getDataFormat()).append(",");
        sb.append(xssfCellStyle.getWrapText()).append(",");
        sb.append(xssfCellStyle.getRotation()).append(",");
        sb.append(xssfCellStyle.getIndention()).append(",");

        // 序列化字体
        XSSFFont font = xssfCellStyle.getFont();
        sb.append(font.getFontHeight()).append(",");
        sb.append(font.getFontName()).append(",");
        sb.append(font.getItalic()).append(",");
        sb.append(font.getStrikeout()).append(",");
        sb.append(font.getColor()).append(",");
        sb.append(font.getBold()).append(",");
        sb.append(font.getTypeOffset()).append(",");
        sb.append(font.getUnderline()).append(",");
        sb.append(font.getFamily()).append(",");
        XSSFColor fontColor = font.getXSSFColor();
        sb.append(fontColor == null ? "null" : fontColor.getARGBHex()).append(","); // 序列化字体颜色

        XSSFColor fillForegroundColor = xssfCellStyle.getFillForegroundXSSFColor();
        sb.append(fillForegroundColor == null ? "null" : fillForegroundColor.getARGBHex()).append(","); // 序列化填充颜色

        return sb.toString();
    }

    public static CellStyle deserializeCellStyle(Workbook workbook, String styleString) {
        if (styleString == null || styleString.isEmpty()) {
            return workbook.createCellStyle();
        }

        String[] styles = styleString.split(",");
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setFillForegroundColor(Short.parseShort(styles[0]));
        cellStyle.setFillBackgroundColor(Short.parseShort(styles[1]));
        cellStyle.setFillPattern(FillPatternType.forInt(Short.parseShort(styles[2])));
        cellStyle.setBorderBottom(BorderStyle.valueOf(Short.parseShort(styles[3])));
        cellStyle.setBorderLeft(BorderStyle.valueOf(Short.parseShort(styles[4])));
        cellStyle.setBorderRight(BorderStyle.valueOf(Short.parseShort(styles[5])));
        cellStyle.setBorderTop(BorderStyle.valueOf(Short.parseShort(styles[6])));
        cellStyle.setBottomBorderColor(Short.parseShort(styles[7]));
        cellStyle.setLeftBorderColor(Short.parseShort(styles[8]));
        cellStyle.setRightBorderColor(Short.parseShort(styles[9]));
        cellStyle.setTopBorderColor(Short.parseShort(styles[10]));
        cellStyle.setAlignment(HorizontalAlignment.forInt(Short.parseShort(styles[11])));
        cellStyle.setVerticalAlignment(VerticalAlignment.forInt(Short.parseShort(styles[12])));
        cellStyle.setDataFormat(Short.parseShort(styles[13]));
        cellStyle.setWrapText(Boolean.parseBoolean(styles[14]));
        cellStyle.setRotation(Short.parseShort(styles[15]));

        cellStyle.setIndention(Short.parseShort(styles[16]));

        // 反序列化字体
        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setFontHeight(Short.parseShort(styles[17]));
        font.setFontName(styles[18]);
        font.setItalic(Boolean.parseBoolean(styles[19]));
        font.setStrikeout(Boolean.parseBoolean(styles[20]));
        font.setColor(Short.parseShort(styles[21]));
        font.setBold(Boolean.parseBoolean(styles[22]));
        font.setTypeOffset(Short.parseShort(styles[23]));
        font.setUnderline(Byte.parseByte(styles[24]));
        font.setFamily(Byte.parseByte(styles[25]));

        if (!"null".equals(styles[26])) {
            try {
                font.setColor(new XSSFColor(decodeColor(styles[26]), null)); // 反序列化字体颜色
            } catch (NumberFormatException e) {
                System.err.println("Error parsing font color: " + styles[26]);
            }
        }
        cellStyle.setFont(font);

        if (!"null".equals(styles[27])) {
            try {
                cellStyle.setFillForegroundColor(new XSSFColor(decodeColor(styles[27]), null)); // 反序列化填充颜色
            } catch (NumberFormatException e) {
                System.err.println("Error parsing fill foreground color: " + styles[27]);
            }
        }

        return cellStyle;
    }

    // 解析颜色字符串，包括处理带有alpha通道的颜色
    private static Color decodeColor(String colorStr) {
        if (colorStr.startsWith("0x")) {
            colorStr = colorStr.substring(2);
        }
        long color = Long.parseLong(colorStr, 16);
        if (colorStr.length() == 8) { // ARGB格式
            int alpha = (int) (color >> 24) & 0xFF;
            int red = (int) (color >> 16) & 0xFF;
            int green = (int) (color >> 8) & 0xFF;
            int blue = (int) color & 0xFF;
            return new Color(red, green, blue, alpha);
        } else if (colorStr.length() == 6) { // RGB格式
            int red = (int) (color >> 16) & 0xFF;
            int green = (int) (color >> 8) & 0xFF;
            int blue = (int) color & 0xFF;
            return new Color(red, green, blue);
        } else {
            throw new NumberFormatException("Invalid color format: " + colorStr);
        }
    }
}
