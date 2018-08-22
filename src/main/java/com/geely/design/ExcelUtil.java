package com.longruan.hnny.util;

import com.longruan.hnny.scheduling.annotation.ExcelAnnotation;
import com.longruan.hnny.scheduling.common.Const;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFRegionUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 导出excel文件工具类
 *
 * @author huxiaolong
 * @date 2018/8/22
 */
public class ExcelUtil {

    public static <E> HSSFWorkbook exportExcel(String tableName, List<E> dataSet) {
        HSSFWorkbook hssfWorkbook = null;
        if (dataSet.size() > 0) {
            int rowIndex = 0;
            Map<String, Object> sheetInfoMap = ExcelUtil.getSheetInfo(dataSet.get(0));
            String sheetName = (String) sheetInfoMap.get("sheetName");
            String datePattern = (String) sheetInfoMap.get("datePattern");
            boolean counted = (boolean) sheetInfoMap.get("counted");
            List<HeaderNode> headerNodeList = ExcelUtil.assembleHeaderNode(dataSet.get(0));
            List<String> notExtensibleHeaderNameList = ExcelUtil.getNotExtensibleHeaderNameList(headerNodeList);
            if (!notExtensibleHeaderNameList.isEmpty()) {
                if (counted) {
                    HeaderNode headerNodeForCountColumn = new HeaderNode();
                    headerNodeForCountColumn.setHeaderName(Const.COUNT);
                    headerNodeForCountColumn.setIndex(-1);
                    headerNodeForCountColumn.setLevel(0);
                    headerNodeForCountColumn.setExtensible(false);
                    headerNodeForCountColumn.setPreNotExtensibleHeaderNodeNum(0);
                    headerNodeForCountColumn.setSubNotExtensibleHeaderNodeNum(0);
                    headerNodeList.add(0, headerNodeForCountColumn);
                    notExtensibleHeaderNameList.add(0, Const.COUNT);
                }
                int columnSum = notExtensibleHeaderNameList.size();
                hssfWorkbook = new HSSFWorkbook();
                HSSFSheet hssfSheet = hssfWorkbook.createSheet(sheetName);
                /*1.标题行设置*/
                HSSFRow hssfRow = hssfSheet.createRow(rowIndex);
                rowIndex++;
                ExcelUtil.setRowHeight(Const.RowType.TABLE_NAME, hssfRow);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, columnSum - 1);
                hssfSheet.addMergedRegion(cellRangeAddress);
                Cell cell = CellUtil.createCell(hssfRow, 0, tableName);
                ExcelUtil.setCellStyleForMergeRegion(cellRangeAddress, cell, Const.RowType.TABLE_HEAD, hssfSheet, hssfWorkbook);
                /*2.表头行设置*/
                // 可以直接调用createHeader方法，
                List<HSSFRow> headerRowList = ExcelUtil.createHeader(headerNodeList, hssfSheet, hssfWorkbook, rowIndex);
                rowIndex += headerRowList.size();
                // boolean singleLineHeader = notExtensibleHeaderNameList.size() == headerNodeList.size();
                // if (singleLineHeader) {
                //     hssfRow = hssfSheet.createRow(rowIndex);
                //     rowIndex++;
                //     ExcelUtil.setRowHeight(Const.RowType.TABLE_HEAD, hssfRow);
                //     HSSFCell[] hssfCells = new HSSFCell[columnSum];
                //     for (int i = 0; i < hssfCells.length; i++) {
                //         hssfCells[i] = hssfRow.createCell(i);
                //         hssfCells[i].setCellValue(new HSSFRichTextString(notExtensibleHeaderNameList.get(i)));
                //         ExcelUtil.setCellStyle(hssfWorkbook, hssfCells[i], Const.RowType.TABLE_HEAD);
                //     }
                // } else {
                //     List<HSSFRow> headerRowList = ExcelUtil.createHeader(headerNodeList, hssfSheet, hssfWorkbook, rowIndex);
                //     rowIndex += headerRowList.size();
                // }
                /*3.内容行设置*/
                ExcelUtil.createContentRows(rowIndex, counted, columnSum, datePattern, headerNodeList, dataSet, hssfSheet, hssfWorkbook);
            }

        }
        return hssfWorkbook;
    }

    private static <E> Map<String, Object> getSheetInfo(E e) {
        Map<String, Object> sheetInfoMap = new HashMap<>(8);
        if (e != null) {
            ExcelAnnotation excelAnnotation = (ExcelAnnotation) e.getClass().getAnnotations()[0];
            sheetInfoMap.put("sheetName", excelAnnotation.sheetName());
            sheetInfoMap.put("datePattern", excelAnnotation.datePattern());
            sheetInfoMap.put("count", excelAnnotation.count());
        }
        return sheetInfoMap;
    }

    private static <E> List<HeaderNode> assembleHeaderNode(E e) {
        List<HeaderInfo> headerInfoList = ExcelUtil.getHeaderInfo(e);
        List<HeaderNode> headerNodeList = new ArrayList<>();
        HeaderNode headerNode;
        if (headerInfoList.size() > 0) {
            for (HeaderInfo item : headerInfoList) {
                headerNode = new HeaderNode();
                headerNode.setHeaderName(item.getHeaderName());
                headerNode.setIndex(item.getIndex());
                headerNode.setLevel(item.getLevel());
                headerNode.setParentIndex(item.getParentIndex());
                headerNode.setExtensible(isExtensible(item, headerInfoList));
                headerNode.setPreNotExtensibleHeaderNodeNum(getPreNotExtensibleHeaderNodeNum(item, headerInfoList));
                headerNode.setSubNotExtensibleHeaderNodeNum(getSubNotExtensibleHeaderNodeNum(item, headerInfoList));
                headerNodeList.add(headerNode);
            }
        }
        return headerNodeList;
    }

    private static List<String> getNotExtensibleHeaderNameList(List<HeaderNode> headerNodeList) {
        List<String> notExtensibleHeaderNameList = new ArrayList<>();
        if (headerNodeList.size() > 0) {
            for (HeaderNode item : headerNodeList) {
                if (!item.isExtensible()) {
                    notExtensibleHeaderNameList.add(item.getHeaderName());
                }
            }
        }
        return notExtensibleHeaderNameList;
    }

    /**
     * 设置行高
     * @param rowType 行的类型，如表名行、表头行、内容行
     * @param hssfRow 行对象
     */
    private static void setRowHeight(String rowType, HSSFRow hssfRow) {
        if (Const.RowType.TABLE_NAME.equals(rowType)) {
            hssfRow.setHeight((short) (36 * 20));
        } else if (Const.RowType.TABLE_HEAD.equals(rowType)) {
            hssfRow.setHeight((short) (20 * 20));
        } else if (Const.RowType.TABLE_CONTENT.equals(rowType)) {
            hssfRow.setHeight((short) (20 * 20));
        } else {
            hssfRow.setHeight((short) (20 * 20));
        }
    }

    /**
     * 设置跨行区域的字体及边框样式
     * @param cellRangeAddress 跨行范围对象
     * @param cell cell对象，第一列
     * @param cellType 单元格类型
     * @param hssfSheet sheet页
     * @param hssfWorkbook excel文档对象
     */
    private static void setCellStyleForMergeRegion(CellRangeAddress cellRangeAddress, Cell cell, String cellType,
                                                   HSSFSheet hssfSheet, HSSFWorkbook hssfWorkbook) {
        // 创建样式对象
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        // 创建字体对象
        HSSFFont hssfFont = hssfWorkbook.createFont();
        // 设置字体及其高度
        ExcelUtil.setFont(cellType, hssfFont);
        // 设置单元格样式字体属性
        hssfCellStyle.setFont(hssfFont);
        // 水平居中
        hssfCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 垂直居中
        hssfCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 添加样式到单元格
        cell.setCellStyle(hssfCellStyle);
        // 设置跨行区域边框
        HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, cellRangeAddress, hssfSheet, hssfWorkbook);
        HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, cellRangeAddress, hssfSheet, hssfWorkbook);
        HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, cellRangeAddress, hssfSheet, hssfWorkbook);
        HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, cellRangeAddress, hssfSheet, hssfWorkbook);
    }

    private static List<HSSFRow> createHeader(List<HeaderNode> headerNodeList, HSSFSheet hssfSheet, HSSFWorkbook hssfWorkbook, int startIndex) {
        List<HSSFRow> headerRowList = new ArrayList<>();
        if (headerNodeList.size() > 0) {
            int deep = ExcelUtil.getHeaderDeep(headerNodeList);
            HSSFRow hssfRow;
            for (int i = 0; i < deep + 1; i++) {
                hssfRow = hssfSheet.createRow(startIndex + i);
                ExcelUtil.setRowHeight(Const.RowType.TABLE_HEAD, hssfRow);
                headerRowList.add(hssfRow);
            }
            String headerName;
            int level = 0;
            int preNotExtensibleHeaderNodeNum;
            int subNotExtensibleHeaderNodeNum;
            CellRangeAddress cellRangeAddress;
            Cell cell;
            for (HeaderNode headerNode : headerNodeList) {
                headerName = headerNode.getHeaderName();
                level = headerNode.getLevel();
                preNotExtensibleHeaderNodeNum = headerNode.getPreNotExtensibleHeaderNodeNum();
                subNotExtensibleHeaderNodeNum = headerNode.getSubNotExtensibleHeaderNodeNum();
                if (headerNode.isExtensible()) {
                    cellRangeAddress = new CellRangeAddress(startIndex + level, startIndex + level,
                            preNotExtensibleHeaderNodeNum, preNotExtensibleHeaderNodeNum + subNotExtensibleHeaderNodeNum - 1);
                } else {
                    cellRangeAddress = new CellRangeAddress(startIndex + level,
                            startIndex + deep, preNotExtensibleHeaderNodeNum, preNotExtensibleHeaderNodeNum);
                }
                hssfSheet.addMergedRegion(cellRangeAddress);
                cell = CellUtil.createCell(headerRowList.get(level), preNotExtensibleHeaderNodeNum, headerName);
                ExcelUtil.setCellStyleForMergeRegion(cellRangeAddress, cell, Const.RowType.TABLE_HEAD, hssfSheet, hssfWorkbook);
            }
        }
        return headerRowList;
    }

    private static <E> void createContentRows(int startIndex, boolean counted, int columnSum, String pattern, List<HeaderNode> headerNodeList,
                                              List<E> dataSet, HSSFSheet hssfSheet, HSSFWorkbook hssfWorkbook) {
        HSSFRow hssfRow = null;
        HSSFCell[] hssfCells = new HSSFCell[columnSum];
        for (int i = 0; i < dataSet.size(); i++) {
            hssfRow = hssfSheet.createRow(i + startIndex);
            ExcelUtil.setRowHeight(Const.RowType.TABLE_CONTENT, hssfRow);
            Object temp = dataSet.get(i);
            List<Object> fieldValueList = ExcelUtil.getRowValueList(temp, headerNodeList);
            Object value;
            int k = 0;
            if (counted) {
                hssfCells[0] = hssfRow.createCell(0);
                ExcelUtil.setCellStyle(hssfWorkbook, hssfCells[0], Const.RowType.TABLE_CONTENT);
                hssfCells[0].setCellValue(i + 1);
                hssfSheet.autoSizeColumn(0);
                k++;
            }
            for (int j = k; j < columnSum; j++) {
                hssfCells[j] = hssfRow.createCell(j);
                ExcelUtil.setCellStyle(hssfWorkbook, hssfCells[j], Const.RowType.TABLE_CONTENT);
                value = fieldValueList.get(j - k);
                if (value == null) {
                    hssfCells[j].setCellValue("");
                } else if (value instanceof Integer) {
                    hssfCells[j].setCellValue((Integer)value);
                } else if (value instanceof Double) {
                    hssfCells[j].setCellValue((Double) value);
                } else if (value instanceof Float) {
                    hssfCells[j].setCellValue((Float) value);
                } else if (value instanceof Long) {
                    hssfCells[j].setCellValue((Long) value);
                } else if (value instanceof Boolean) {
                    hssfCells[j].setCellValue((Boolean) value ? "是" : "否");
                } else if (value instanceof Date) {
                    hssfCells[j].setCellValue(new SimpleDateFormat(StringUtils.isNotBlank(pattern) ? pattern : "yyyy-mm-dd").format((Date) value));
                } else {
                    hssfCells[j].setCellValue(String.valueOf(value));
                }
                hssfSheet.autoSizeColumn(j);
            }
        }
    }

    private static <E> List<HeaderInfo> getHeaderInfo(E e) {
        List<HeaderInfo> headerInfoList = new ArrayList<>();
        if (e != null) {
            HeaderInfo headerInfo = new HeaderInfo();
            Field[] fields = e.getClass().getDeclaredFields();
            ExcelAnnotation excelAnnotation;
            for (int i = 0; i < fields.length; i++) {
                if (fields[i].getAnnotations().length != 0) {
                    excelAnnotation = (ExcelAnnotation) fields[i].getAnnotations()[0];
                    headerInfo.setHeaderName(excelAnnotation.headerName());
                    headerInfo.setIndex(excelAnnotation.index());
                    headerInfo.setLevel(excelAnnotation.level());
                    headerInfo.setParentIndex(excelAnnotation.parentIndex());
                    headerInfoList.add(headerInfo);
                }
            }
        }
        return headerInfoList;
    }

    /**
     * 判断节点是否可扩展，即对应表头下是否含有子表头
     * @param headerInfo
     * @param headerInfoList
     * @return
     */
    private static boolean isExtensible(HeaderInfo headerInfo, List<HeaderInfo> headerInfoList) {
        for (HeaderInfo item : headerInfoList) {
            if (headerInfo.getIndex() == item.getParentIndex()) {
                return true;
            }
        }
        return false;
    }

    /**
     * 获得当前节点之前（或者空间左侧）不可扩展节点的数量，
     * 前提是headerInfoList中表头单元信息对象存储的先后顺序，需遵循基于对应表头单元空间位置先左后右，从上往下（或先父后子）的原则
     * @param headerInfo
     * @param headerInfoList
     * @return
     */
    private static int getPreNotExtensibleHeaderNodeNum(HeaderInfo headerInfo, List<HeaderInfo> headerInfoList) {
        int count = 0;
        for (int i = 0; i < headerInfoList.size(); i++) {
            if (headerInfo.equals(headerInfoList.get(i))) {
                break;
            }
            if (isExtensible(headerInfo, headerInfoList)) {
                count++;
            }
        }
        return count;
    }

    /**
     * 获得当前节点下侧不可扩展节点的数量，
     * 前提是headerInfoList中表头单元信息对象存储的先后顺序，需遵循基于对应表头单元空间位置先左后右，从上往下（或先父后子）的原则
     * @param headerInfo
     * @param headerInfoList
     * @return
     */
    private static int getSubNotExtensibleHeaderNodeNum(HeaderInfo headerInfo, List<HeaderInfo> headerInfoList) {
        int count = 0;
        for (int i = 0; i < headerInfoList.size(); i++) {
            if (headerInfo.getIndex() == headerInfoList.get(i).getParentIndex()) {
                if (!isExtensible(headerInfoList.get(i), headerInfoList)) {
                    count++;
                } else {
                    count += getSubNotExtensibleHeaderNodeNum(headerInfoList.get(i), headerInfoList);
                }
            }
        }
        return count;
    }

    private static void setFont(String cellType, HSSFFont hssfFont) {
        if (hssfFont != null) {
            if (Const.RowType.TABLE_NAME.equals(cellType)) {
                // 设置字体类型
                hssfFont.setFontName("宋体");
                // 设置字体高度
                hssfFont.setFontHeightInPoints((short) 18);
            } else if (Const.RowType.TABLE_HEAD.equals(cellType)) {
                // 设置字体类型
                hssfFont.setFontName("宋体");
                // 设置字体高度
                hssfFont.setFontHeightInPoints((short) 11);
            } else if (Const.RowType.TABLE_CONTENT.equals(cellType)) {
                // 设置字体类型
                hssfFont.setFontName("宋体");
                // 设置字体高度
                hssfFont.setFontHeightInPoints((short) 10);
            } else {
                // 设置字体类型
                hssfFont.setFontName("宋体");
                // 设置字体高度
                hssfFont.setFontHeightInPoints((short) 9);
                // 设置字体颜色和下划线样式
                hssfFont.setColor(HSSFFont.COLOR_RED);
                hssfFont.setUnderline((byte) 1);
            }
        }
    }

    private static int getHeaderDeep(List<HeaderNode> headerNodeList) {
        int deep = 0;
        for (int i = 0; i < headerNodeList.size(); i++) {
            if (headerNodeList.get(i).getLevel() > deep) {
                deep = headerNodeList.get(i).getLevel();
            }
        }
        return deep;
    }

    private static List<Object> getRowValueList(Object temp, List<HeaderNode> headerNodeList) {
        List<Integer> indexList = ExcelUtil.getNotExtensibleHeaderIndexList(headerNodeList);
        Map<Integer, Object> map = new TreeMap<>();
        Field[] fields = temp.getClass().getDeclaredFields();
        ExcelAnnotation excelAnnotation;
        int index;
        boolean valueColumn;
        String fieldName;
        String getMethodName;
        Object value = null;
        for (Field field : fields) {
            if (field.getAnnotations().length != 0) {
                excelAnnotation = (ExcelAnnotation) field.getAnnotations()[0];
                index = excelAnnotation.index();
                valueColumn = ExcelUtil.isValueColumn(index, indexList);
                if (valueColumn) {
                    fieldName = field.getName();
                    getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    try {
                        Method method = temp.getClass().getMethod(getMethodName);
                        value = method.invoke(temp);
                    } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                        e.printStackTrace();
                    }
                    map.put(index, value);
                }
            }
        }
        return ExcelUtil.sortByKey(map);
    }

    /**
     * 设置单元格及内部字体样式
     * @param hssfWorkbook hssfWorkbook对象
     * @param hssfCell 单元格对象
     * @param cellType 单元格类型
     */
    private static void setCellStyle(HSSFWorkbook hssfWorkbook, HSSFCell hssfCell, String cellType) {
        // 创建样式对象
        HSSFCellStyle hssfCellStyle = hssfWorkbook.createCellStyle();
        // 创建字体对象
        HSSFFont hssfFont = hssfWorkbook.createFont();
        // 设置单元格样式字体属性
        hssfCellStyle.setFont(hssfFont);
        // 水平居中
        hssfCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 垂直居中
        hssfCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 单元格表框线类型，BORDER_THIN为实线，BORDER_DOTTED为点划线
        hssfCellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        hssfCellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        hssfCellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        hssfCellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        // 设置字体及其高度
        ExcelUtil.setFont(cellType, hssfFont);
        // 为单元格添加已设置完整的样式
        hssfCell.setCellStyle(hssfCellStyle);
    }

    private static List<Integer> getNotExtensibleHeaderIndexList(List<HeaderNode> headerNodeList) {
        List<Integer> notExtensibleHeaderIndexList = new ArrayList<>();
        for (HeaderNode item : headerNodeList) {
            if (!item.isExtensible() && item.getIndex() >= 0) {
                notExtensibleHeaderIndexList.add(item.getIndex());
            }
        }
        return notExtensibleHeaderIndexList;
    }

    private static boolean isValueColumn(int index, List<Integer> indexList) {
        for (Integer item : indexList) {
            if (index == item) {
                return true;
            }
        }
        return false;
    }

    private static List<Object> sortByKey(Map<Integer, Object> map) {
        List<Object> list = new ArrayList<>();
        Set<Integer> keySet = map.keySet();
        for (Integer key : keySet) {
            list.add(map.get(key));
        }
        return list;
    }

    static class HeaderInfo {
        private String headerName;
        private int index;
        private int level;
        private int parentIndex;

        public HeaderInfo() {
        }

        public HeaderInfo(String headerName, int index, int level, int parentIndex) {
            this.headerName = headerName;
            this.index = index;
            this.level = level;
            this.parentIndex = parentIndex;
        }

        public String getHeaderName() {
            return headerName;
        }

        public void setHeaderName(String headerName) {
            this.headerName = headerName;
        }

        public int getIndex() {
            return index;
        }

        public void setIndex(int index) {
            this.index = index;
        }

        public int getLevel() {
            return level;
        }

        public void setLevel(int level) {
            this.level = level;
        }

        public int getParentIndex() {
            return parentIndex;
        }

        public void setParentIndex(int parentIndex) {
            this.parentIndex = parentIndex;
        }
    }

    static class HeaderNode {
        private String headerName;
        private int index;
        private int level;
        private int parentIndex;
        private boolean extensible;
        private int preNotExtensibleHeaderNodeNum;
        private int subNotExtensibleHeaderNodeNum;

        public HeaderNode() {
        }

        public HeaderNode(String headerName, int index, int level, int parentIndex, boolean extensible,
                          int preNotExtensibleHeaderNodeNum, int subNotExtensibleHeaderNodeNum) {
            this.headerName = headerName;
            this.index = index;
            this.level = level;
            this.parentIndex = parentIndex;
            this.extensible = extensible;
            this.preNotExtensibleHeaderNodeNum = preNotExtensibleHeaderNodeNum;
            this.subNotExtensibleHeaderNodeNum = subNotExtensibleHeaderNodeNum;
        }

        public String getHeaderName() {
            return headerName;
        }

        public void setHeaderName(String headerName) {
            this.headerName = headerName;
        }

        public int getIndex() {
            return index;
        }

        public void setIndex(int index) {
            this.index = index;
        }

        public int getLevel() {
            return level;
        }

        public void setLevel(int level) {
            this.level = level;
        }

        public int getParentIndex() {
            return parentIndex;
        }

        public void setParentIndex(int parentIndex) {
            this.parentIndex = parentIndex;
        }

        public boolean isExtensible() {
            return extensible;
        }

        public void setExtensible(boolean extensible) {
            this.extensible = extensible;
        }

        public int getPreNotExtensibleHeaderNodeNum() {
            return preNotExtensibleHeaderNodeNum;
        }

        public void setPreNotExtensibleHeaderNodeNum(int preNotExtensibleHeaderNodeNum) {
            this.preNotExtensibleHeaderNodeNum = preNotExtensibleHeaderNodeNum;
        }

        public int getSubNotExtensibleHeaderNodeNum() {
            return subNotExtensibleHeaderNodeNum;
        }

        public void setSubNotExtensibleHeaderNodeNum(int subNotExtensibleHeaderNodeNum) {
            this.subNotExtensibleHeaderNodeNum = subNotExtensibleHeaderNodeNum;
        }
    }
}
