package com.qzing.main;

import com.qzing.po.ReportPO;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 合并两个excel
 * 表格一为报表目录，表格二为每张报表包含的维度和数据库名
 * 处理逻辑：根据表一的报表名关联表二从而合并两个excel输出带有报表目录、维度、数据库名的excel表三
 */
public class Run {
    static int single = 0;
    static int multi = 0;
    public static void main(String[] args) throws Exception {
        //读取表格二存入集合中供使用
        //读取文件
        Map<String, ReportPO> map = readExcel("excel/表格二.xls");
/*        System.out.println("----------读取报表--------");
        map.forEach((k,y)->{
            System.out.println(k);
            y.getDimension().forEach((d,t)->{
                System.out.println(d+":"+t);
            });
            if(y.getDimension().size()==1){
                single += 1;
            }else{
                multi+=y.getDimension().size();
            }

            System.out.println("--");
        });
        System.out.println("----------读取报表结束--------单维度："+single+"，多维度："+multi);*/
        //读取表格一遍历目录关联表格二生成表格三
        buildExcel3(map,"excel/表格一.xls","excel/");
    }

    /**
     * 读取表格一遍历目录关联表格二生成表格三
     * @param map 表格二的数据
     * @param excel1Path 表格一的路径
     * @param outExcel3Dir 生成表格三的目录
     */
    private static void buildExcel3(Map<String, ReportPO> map, String excel1Path, String outExcel3Dir) throws Exception {
        //读取表格一
        //读取文件
        FileInputStream fis = new FileInputStream(excel1Path);
        Workbook workbook = new HSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        CellStyle reportStyle = workbook.createCellStyle();
        reportStyle.setAlignment(HorizontalAlignment.CENTER);
        reportStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex()); // 是设置前景色不是背景色
        reportStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        workbook.setSheetName(0,"报表数据");
        //合并两张表格的数据
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            //累计新增行数
            int addRowNum = 0;

            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0);
            if(cell==null){
                continue;
            }
            String key = cell.getStringCellValue();

            //关联到数据则合并
            if(map.containsKey(key)){
                //合并报表名字 占两列
                CellRangeAddress cellRangeAddress = new CellRangeAddress(i, i, 0, 1);
                sheet.addMergedRegion(cellRangeAddress);
                sheet.getRow(i).getCell(0).setCellStyle(reportStyle);
                //插入行来新增维度
                int size = map.get(key).getDimension().size();
                int startRet = i+addRowNum+1;
                sheet.shiftRows(startRet,sheet.getLastRowNum(),size);
                addRowNum+=size;
                //插入数据
                Map dimension = map.get(key).getDimension();
                int j = 0;
                for (Object str: dimension.keySet()
                     ) {
                    Row addRow = sheet.createRow(startRet+j);
                    addRow.createCell(0).setCellValue((String)str);
                    addRow.createCell(1).setCellValue((String)dimension.get((String)str));
                    j++;
                }
                //合并后边的备注行
                if(size>1){
                    CellRangeAddress cellRangeAddress2 = new CellRangeAddress(i+1, i+size, 2, 2);
                    sheet.addMergedRegion(cellRangeAddress2);
                }


            }
        }
        // 输出为一个新的Excel，也就是动态修改完之后的excel
        String fileName = "表格三" + ".xls";
        OutputStream out = new FileOutputStream(outExcel3Dir + fileName);
        workbook.write(out);
        fis.close();
        out.flush();
        out.close();
    }

    /**
     * 把表格二存到集合中
     * @param filePath 表格二的路径
     * @return
     * @throws IOException
     */
    public static Map<String, ReportPO> readExcel(String filePath)throws IOException{
        //第一列为报表名字
        //第二列为报表维度
        //第三列为该维度对应的数据库表
        //记录重复报表名字 因为报表下线 xml还可能存留在，这批报表需要手动进一步核对
        List <String>repeatReport = new ArrayList<String>();
        //存报表 k-报表名，v-报表实体
        Map<String, ReportPO> map= new HashMap<String, ReportPO>();
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new HSSFWorkbook(fis);
        //存放多维报表的行数
        List<Integer>multiLineNum = new ArrayList<Integer>();
        Sheet sheet = workbook.getSheetAt(0);
        //读取多维报表
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            //合并开始行号
            int firstRow = range.getFirstRow();
            //合并结束行号
            int lastRow = range.getLastRow();
            //存放单张报表
            ReportPO rt = new ReportPO();
            //存放维度 k-维度名 v-数据库表名
            Map dimensionMap =  new HashMap();
            //读取报表名
            String reportName = sheet.getRow(firstRow).getCell(0).getStringCellValue();
            rt.setName(reportName);
            for (int j = firstRow; j <= lastRow; j++) {
                Row row = sheet.getRow(j);
                //读取报表维度
                String dimensionName = row.getCell(1).getStringCellValue();
                //读取数据库表名
                String dimensionTable = row.getCell(2).getStringCellValue();
                dimensionMap.put(dimensionName,dimensionTable);
                multiLineNum.add(j);
                if(map.containsKey(reportName)){
                    System.out.println("==========重复报表："+reportName);
                    repeatReport.add(reportName);
                }
            }
            rt.setDimension(dimensionMap);
            map.put(reportName,rt);
        }
        //System.out.println("------------多维报表读取结束-----------------");
        //读取单维报表
        int singleCount = 0;
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            //如果不是合并的
            if(!multiLineNum.contains(i)){
                ReportPO rp =  new ReportPO();
                Map dimensionMap =  new HashMap();
                Row row = sheet.getRow(i);
                String reportName = row.getCell(0).getStringCellValue();
                String dimension = row.getCell(1).getStringCellValue();
                String dataTable = row.getCell(2).getStringCellValue();
                dimensionMap.put(dimension,dataTable);
                rp.setName(reportName);
                rp.setDimension(dimensionMap);
                if(map.containsKey(reportName)){
                    System.out.println("==========重复报表："+reportName);
                    repeatReport.add(reportName);
                }
                map.put(reportName,rp);


                singleCount++;
            }
        }
        System.out.println("总行数："+sheet.getLastRowNum()+"，单维度数所占总行数："+singleCount+"，多维度所占行数："+multiLineNum.size());
        System.out.println("重复报表:");
        repeatReport.stream().distinct().forEach(System.out::println);
       return map;
    }







    public static void test()throws IOException {
        //读取文件
        FileInputStream fis = new FileInputStream("excel/表格一.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        //XSSFSheet sheet = workbook.getSheet("目录");
        //克隆sheet
        XSSFSheet cloneSheet = workbook.cloneSheet(5, "报表数据库表统计");
        //遍历表格
        for (int i = 0; i <cloneSheet.getLastRowNum() ; i++) {
            XSSFCell cell = cloneSheet.getRow(i).getCell(0);
            String cellValue = cell.getStringCellValue();
            System.out.println("遍历报表==========================");
            System.out.println(cellValue);
            System.out.println("遍历报表结束==========================");
        }
        //向下移动
        cloneSheet.shiftRows(1,cloneSheet.getLastRowNum(),2);
        cloneSheet.shiftRows(2+2,cloneSheet.getLastRowNum(),1);
        cloneSheet.shiftRows(3+2+1,cloneSheet.getLastRowNum(),3);
        cloneSheet.shiftRows(4+2+1+3,cloneSheet.getLastRowNum(),4);
        //合并
        CellRangeAddress cellRangeAddress = new CellRangeAddress(4 + 2 + 1 + 3 - 1, 4 + 2 + 1 + 3 - 1 + 4, 0, 0);
        cloneSheet.addMergedRegion(cellRangeAddress);
        // 输出为一个新的Excel，也就是动态修改完之后的excel
        String fileName = "报表系统报表统计-" + System.currentTimeMillis() + ".xlsx";
        OutputStream out = new FileOutputStream("excel/" + fileName);
        workbook.write(out);
        fis.close();
        out.flush();
        out.close();
    }
}
