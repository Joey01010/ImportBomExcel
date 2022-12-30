package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class Main {
    public static final int CELL_ROW = 142;
    public static final int CELL_IDX = 27;
    public static final String ROOT = System.getProperty("user.dir");
    public static final String PATH1 = "/BOM录入模板.xlsx";
    public static final String PATH2 = "/工艺卡.xlsx";
    public static final String PATH3 = "/导入数据.xlsx";

    /**
     * 获取模板
     *
     * @return
     * @throws IOException
     */
    public static List<List<String>> getTemplate() throws IOException {

        //读取模板表格数据
        File template = new File(ROOT + PATH1);
        Workbook wb = WorkbookFactory.create(template);
        Sheet tSheet = wb.getSheet("物料清单#单据头(FBillHead)");

        List<String> parentList = new ArrayList<>();
        List<String> itemList = new ArrayList<>();

        Row row = tSheet.getRow(2);
        for (int i = 0; i < CELL_ROW; i++) {
            List<String> temp = null;
            Cell cell = row.getCell(i);
            if (i < CELL_IDX)
                temp = parentList;
            else
                temp = itemList;
            if (cell != null) {
                cell.setCellType(CellType.STRING);
                temp.add(cell.getStringCellValue());
            } else {
                temp.add("");
            }
        }

        List<List<String>> result = Arrays.asList(parentList, itemList);
        wb.close();
        return result;
    }

    /**
     * 获取数据
     *
     * @return
     * @throws IOException
     */
    public static List<Bom> getData() throws IOException {
        List<Bom> result = new ArrayList<>();
        File file = new File(ROOT + PATH2);
        Workbook wb = WorkbookFactory.create(file);

        //读取Excel各表
        //逻辑：先读取每张表每一行的子项，组织子项的bom组成，子项组成为空，则搜索不为空的子项id，
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            //创建每张表存储容器
            List<Bom> parentItems = new ArrayList<>();
            List<Bom> singleChildItems = new ArrayList<>();
            List<Bom> multiChildItems = new ArrayList<>();

            //获取表格
            Sheet sheet = wb.getSheetAt(i);

            //读取每张表的子项组成
            for (int j = 3; j < sheet.getLastRowNum(); j++) {
                try {
                    //封装子项Bom
                    Bom bom = new Bom();
                    Row row = sheet.getRow(j);
                        //设置bomId
                    bom.setBomId(row.getCell(1).getStringCellValue());
                        //设置bom组成
                    List<String> list = bom.getItems();
                    List<Integer> asList = Arrays.asList(3, 4, 8);
                    for (Integer integer : asList) {
                        Cell cell = row.getCell(integer);
                        if (cell != null && cell.getStringCellValue().contains("S")) {
                            list.add(cell.getStringCellValue());
                        }
                    }
                        //设置bom回路
                    Cell cell = row.getCell(14);
                    cell.setCellType(CellType.STRING);
                    bom.setCircuit(cell.getStringCellValue());

                    //存入容器
                    if(!bom.getBomId().equals("")){
                        if(bom.getItems().size() > 0)
                            singleChildItems.add(bom);
                        else
                            multiChildItems.add(bom);
                    }

                    //筛除重复和空的BomId，然后储存进结果容器。
                    Boolean flag = true;
                    for (Bom temp : result) {
                        if (temp.getBomId().equals(bom.getBomId()) || bom.getBomId().equals("")) {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                        result.add(bom);

                } catch (Exception e) {

                }
            }

            //处理多合一看板的子项组成
            List<Bom> remove = new ArrayList<>();
            for (Bom multi : multiChildItems) {
                String circuits = multi.getCircuit();
                String[] circuitList = circuits.split("/");
                for (int j = 0; j < circuitList.length; j++) {
                    String circuit = circuitList[j];
                    if (circuit != null && !circuit.equals("")) {
                        for (int k = 0; j < singleChildItems.size(); j++) {
                            Bom single = singleChildItems.get(k);
                            Boolean flag = null;
                            if(circuitList.length > 0){
                                flag = single.getCircuit().equals(circuit);
                            } else {
                                flag = single.getCircuit().contains(circuit) && !single.getCircuit().equals(circuit);
                            }
                            //多合一的回路对应单线的回路
                            if (flag) {
                                //存储进多合一容器
                                multi.getItems().add(single);
                                //存入到删除列表中
                                remove.add(single);
                            }
                        }
                    }
                }
            }
            //从单线容器中删除重复
            singleChildItems.removeAll(remove);

            //存入产成品子项组成
            parentItems.addAll(singleChildItems);
            parentItems.addAll(multiChildItems);

            //读取产成品编号
            Row codeRow = sheet.getRow(1);
            Cell codeCell = codeRow.getCell(1);
            String cellValue = codeCell.getStringCellValue();
            String[] codes = cellValue.split("/");
            //封装产成品bom组成
            for (String code : codes) {
                Bom<Bom> bom = new Bom<>();
                bom.setBomId(code);
                bom.getItems().addAll(parentItems);
                //产成品存入结果容器
                result.add(bom);
            }

            //子项存入结果容器
            for (Bom item : parentItems) {
                if(!result.contains(item)){
                    result.add(item);
                }
            }
        }

        //测试
        System.out.println(result.size());
        for (Bom bom : result) {
            System.out.println(bom.getBomId());
            for (int i = 0; i < bom.getItems().size(); i++) {
                String code = null;
                Object o = bom.getItems().get(i);
                if(o instanceof String){
                    code = o.toString();
                } else {
                    code = ((Bom) o).getBomId();
                }
                System.out.println(i + 1 + " = " + code);
            }
            System.out.println("-------------------------------");
        }

        wb.close();
        return result;
    }

    public static void main(String[] args) throws IOException {
        List<List<String>> template = getTemplate();
        List<Bom> data = getData();
    }
}