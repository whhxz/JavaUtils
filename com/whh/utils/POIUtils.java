package com.whh.utils;

import com.feiniu.order.annotations.ExportTitle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * POI
 * Created by xuzhuo on 2015/9/28.
 */
public class POIUtils {
    /**
     * excel导出
     *
     * @param dataList 数据
     * @param clazz    导出类
     * @return HSSFWorkbook
     * @throws IllegalAccessException
     */
    public static HSSFWorkbook exportExcel(List<?> dataList, Class clazz){
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFSheet sheet = workbook.createSheet("data");
        //标题样式
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); // 设置单无格的边框为粗体
        cellStyle.setBottomBorderColor(HSSFColor.BLACK.index); // 设置单元格的边框颜色．
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor(HSSFColor.BLACK.index);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setRightBorderColor(HSSFColor.BLACK.index);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setTopBorderColor(HSSFColor.BLACK.index);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);//居中
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中

        //数据样式
        HSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        //标题字体
        HSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("微软雅黑");
        titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        titleFont.setFontHeightInPoints((short) 12);//设置字体大小
        cellStyle.setFont(titleFont);

        Field[] fields = clazz.getDeclaredFields();
        //遍历数据集合
        for (int i = 0; i < dataList.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            //遍历导出字段
            for (Field field : fields) {
                Annotation[] annotations = field.getAnnotations();
                //遍历注解
                for (Annotation annotation : annotations) {
                    if (annotation instanceof ExportTitle) {
                        ExportTitle title = (ExportTitle) annotation;
                        if (i == 0) {
                            //创建标题头
                            HSSFRow titleRow = sheet.getRow(0) == null ? sheet.createRow(0) : sheet.getRow(0);
                            titleRow.setHeight((short) (40 * 20));
                            HSSFCell cell = titleRow.createCell((short) title.rowNum());
                            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                            cell.setCellValue(title.value());
                            cell.setCellStyle(cellStyle);
                            sheet.setColumnWidth((short) title.rowNum(), (short) (30 * 256));
                        }
                        //填充数据
                        row.createCell((short) title.rowNum());
                        row.setHeight((short) (30 * 20));
                        HSSFCell cell = row.createCell((short) title.rowNum());
                        cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                        // 获取字段get方法
                        Object value = null;
                        try {
                            PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), clazz);
                            Method readMethod = descriptor.getReadMethod();
                            value = readMethod.invoke(dataList.get(i));
                        } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
                            e.printStackTrace();
                        }
                        if (value == null || "null".equals(value)) value = "";
                        //时间格式化
                        if (value instanceof Date) {
                            value = format.format(value);
                        }
                        cell.setCellValue(value.toString());
                        cell.setCellStyle(dataStyle);
                    }
                }
            }
        }

        return workbook;
    }

    /**
     * 合并单元格
     *
     * @param hssfWorkbook 要合并的文档
     * @param columns      要合并的列集合
     */
    public static void mergeCell(HSSFWorkbook hssfWorkbook, int[] columns) {
        HSSFSheet sheet = hssfWorkbook.getSheet("data");
        int rowLength = sheet.getLastRowNum();
        for (int column : columns) {
            String tmpValue = "";
            int tmpColumnNum = -1;
            for (int j = 0; j <= rowLength; j++) {
                String cellValue = sheet.getRow(j).getCell((short) column).getStringCellValue();
                //判断前后上下单元格是否一样
                if (!tmpValue.equals(cellValue)) {
                    //判断是否应该合并
                    if (j - 1 != tmpColumnNum) {
                        sheet.addMergedRegion(new Region(tmpColumnNum, (short) column, j - 1, (short) column));
                    }
                    tmpColumnNum = j;
                }
                tmpValue = cellValue;
            }
            //判断最后是否合并
            if (tmpColumnNum != rowLength) {
                sheet.addMergedRegion(new Region(tmpColumnNum, (short) column, rowLength, (short) column));
            }
        }
    }

    /**
     * 合并某列,同时依据该列合并其他列
     *
     * @param hssfWorkbook 要合并的文档
     * @param column       合并列基础
     * @param mergeColumns 合并的其他列
     */
    public static void mergeCellByColumn(HSSFWorkbook hssfWorkbook, int column, int[] mergeColumns) {
        HSSFSheet sheet = hssfWorkbook.getSheet("data");
        int rowLength = sheet.getLastRowNum();
        String tmpValue = "";
        int tmpColumnNum = -1;
        for (int i = 0; i <= rowLength; i++) {
            String cellValue = sheet.getRow(i).getCell((short) column).getStringCellValue();
            if (!tmpValue.equals(cellValue)) {
                if (i - 1 != tmpColumnNum) {
                    sheet.addMergedRegion(new Region(tmpColumnNum, (short) column, i - 1, (short) column));
                    for (int mergeColumn : mergeColumns) {
                        sheet.addMergedRegion(new Region(tmpColumnNum, (short) mergeColumn, i - 1, (short) mergeColumn));
                    }
                }
                tmpColumnNum = i;
            }
            tmpValue = cellValue;
        }
        //判断最后是否合并
        if (tmpColumnNum != rowLength) {
            sheet.addMergedRegion(new Region(tmpColumnNum, (short) column, rowLength, (short) column));
            for (int mergeColumn : mergeColumns) {
                sheet.addMergedRegion(new Region(tmpColumnNum, (short) mergeColumn, rowLength, (short) mergeColumn));
            }
        }
    }

    /**
     * 修改序列
     *
     * @param hssfWorkbook 文档
     * @param column       序列标准
     * @param startRowNum  开始序列
     */
    public static void editExcelIndex(HSSFWorkbook hssfWorkbook, int column, int startRowNum) {

        HSSFSheet sheet = hssfWorkbook.getSheet("data");
        int rowLength = sheet.getLastRowNum();
        List<Region> regions = new ArrayList<>();
        //获取要合并的序列
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            Region region = sheet.getMergedRegionAt(i);
            if (region.getColumnFrom() == column) {
                region.setColumnFrom((short) 0);
                region.setColumnTo((short) 0);
                regions.add(region);
            }
        }
        //合并序列并修改值
        for (Region region : regions) {
            String cellValue = sheet.getRow(region.getRowFrom()).getCell(region.getColumnFrom()).getStringCellValue();
            for (int i = region.getRowFrom(); i <= region.getRowTo(); i++) {
                sheet.getRow(i).getCell(region.getColumnFrom()).setCellValue(cellValue);
            }
            sheet.addMergedRegion(region);
        }

        //重排序列
        int upValue = 0;
        int valueNum = 0;
        for (int i = startRowNum; i <= rowLength; i++) {
            String cellValue = sheet.getRow(i).getCell((short) 0).getStringCellValue();
            if (Integer.parseInt(cellValue) != upValue) {
                sheet.getRow(i).getCell((short) 0).setCellValue(Integer.toString(++valueNum));
            }
            upValue = Integer.parseInt(cellValue);
        }
    }

    /**
     * 删除列
     *
     * @param hssfWorkbook 文档
     * @param columns      要删除列集合
     */
    public static void removeColumns(HSSFWorkbook hssfWorkbook, int[] columns) {
        HSSFSheet sheet = hssfWorkbook.getSheet("data");
        //排序
        Arrays.sort(columns);
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            int tmpMove = 0;
            HSSFRow row = sheet.getRow(i);
            short lastCellNum = row.getLastCellNum();
            for (int j = columns[0]; j <= lastCellNum; j++) {
                HSSFCell cell = row.getCell((short) j);
                if (Arrays.binarySearch(columns, j) > -1) {
                    row.removeCell(cell);
                    tmpMove++;
                    continue;
                }
                cell.setCellNum((short) (j - tmpMove));
            }
        }
    }
}
