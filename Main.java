package com.company;

import java.io.*;
import java.math.RoundingMode;
import java.util.*;

import com.sun.org.apache.bcel.internal.generic.DDIV;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.NumberFormat;


public class Main {
    public static void main(String[] args) {
        try {

            //获取指定列的值
           // int col=1;
            // readSpecifyColumns(new File("G:\\IDEA\\1111.xls"),col);

            //获取指定行的值
          //  int row=1;
            // readSpecifyRows(new File("G:\\IDEA\\1111.xls"),row);

            //读取每行每列行列的值
            // readRowsAndColums(new File("G:\\IDEA\\423.xls"));

            //自定义对第几列数据进行处理
            int col_index=8;
            hash_excel(new File("G:\\IDEA\\4255.xls"), col_index);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *  	读取指定列
     * @param file
     * @throws Exception
     */
    public static void readSpecifyColumns(File file,int col)throws Exception{
        ArrayList<String> columnList = new ArrayList<String>();
        ArrayList<String> valueList = new ArrayList<String>();
        Workbook readwb = null;
        InputStream io = new FileInputStream(file.getAbsoluteFile());
        readwb = Workbook.getWorkbook(io);
        Sheet readsheet = readwb.getSheet(0);
        int rsColumns = readsheet.getColumns();  //获取表格列数
        int rsRows = readsheet.getRows();  //获取表格行数
        for (int i = 1; i < rsRows; i++) {
            Cell cell_name = readsheet.getCell(0, col);  //第一列的值
            columnList.add(cell_name.getContents());


        }
        System.out.println(columnList);

    }


    /**
     *   	读取指定行
     * @param file
     * @throws Exception
     */
    public static void readSpecifyRows(File file,int index)throws Exception{
        ArrayList<String> columnList = new ArrayList<String>();
        Workbook readwb = null;
        InputStream io = new FileInputStream(file.getAbsoluteFile());
        readwb = Workbook.getWorkbook(io);
        Sheet readsheet = readwb.getSheet(0);
        int rsColumns = readsheet.getColumns();  //获取表格列数
        int rsRows = readsheet.getRows();  //获取表格行数
        for (int i = 1; i < rsColumns; i++) {
            Cell cell_name = readsheet.getCell(i, index);  //在这里指定行，此处需要手动更改，获取不同行的值
            columnList.add(cell_name.getContents());
        }
        System.out.println(columnList);
    }


    private static void readRowsAndColums(File file) throws BiffException, IOException {
        //1:创建workbook
        Workbook workbook=Workbook.getWorkbook(new File(String.valueOf(file)));
        //2:获取第一个工作表sheet
        Sheet sheet=workbook.getSheet(0);
        //3:获取数据
        System.out.println("行："+sheet.getRows());
        System.out.println("列："+sheet.getColumns());
        for(int i=0;i<sheet.getRows();i++){
            for(int j=0;j<sheet.getColumns();j++){
                Cell cell=sheet.getCell(j,i);
                System.out.print(cell.getContents()+" ");
            }
            System.out.println();
        }

        //最后一步：关闭资源
        workbook.close();
    }


    /**
     * 	将获取到的值写入到TXT或者xls中
     * @param file
     * @throws Exception
     */
    public static void hash_excel(File file,int row) throws Exception {
        FileWriter fWriter = null;
        PrintWriter out = null;
        String fliename = file.getName().replace(".xls", "");
        fWriter = new FileWriter(file.getParent()+ "/hashExcel.xls");//输出格式为.xls


        out = new PrintWriter(fWriter);
        InputStream is = new FileInputStream(file.getAbsoluteFile());
        Workbook wb = null;
        wb = Workbook.getWorkbook(is);
        int sheet_size = wb.getNumberOfSheets();
        Sheet sheet = wb.getSheet(0);

        System.out.println("列数："+sheet.getColumns());
        System.out.println("行数："+sheet.getRows());
        ArrayList<Double> columnList = new ArrayList<Double>();

        //保存每天对应的数据数组
        Map<String,List<Double>> myMultimap=new LinkedHashMap<>();
        //将日期和平均数对应起来
        Map<String,Double> Aver_Map=new LinkedHashMap<>();

        //日期
        String data_temp;

        //获取日期：第一列
        Cell cell_Data = sheet.getCell(0,1);
        data_temp=cell_Data.getContents();

        //保存每天的日期
        List<Double>num=new ArrayList<>();


        //获取第row列的所有数字
        for (int j = 1; j < sheet.getRows(); j++) {
            //指定第row列的数据
            Cell cell_Num = sheet.getCell(row,j);

            //获取日期
            cell_Data = sheet.getCell(0,j);

            //如果还是同一天，判断是否有数字
            if(data_temp.equals(cell_Data.getContents()))
            {
                //如果是数字
                if (cell_Num.getType() == CellType.NUMBER) {
                    NumberCell numberCell = (NumberCell) cell_Num;
                    Double numberValue = numberCell.getValue();
                    num.add(numberValue);
                    columnList.add(numberValue);
                }
            }

            //如果到了下一天，就存入数据数组信息
            else
            {
                myMultimap.put(data_temp,num);

                //临时list计算平均数
                List<Double> TempNum=num;

                //获取平均数
                double temp=0.0000000000000000000;
                //排序
                Collections.sort(TempNum);

                //不加最大值最小值
                for(int k=1;k<TempNum.size()-1;k++)
                {
                    temp+=TempNum.get(k);
                }

                //求平均数
                temp/=1.000000*(TempNum.size()-2);

                //将平均数与日期对应起来
                Aver_Map.put(data_temp,temp);

                //重新获取下一天的数值
                num=new ArrayList<>();
                data_temp=cell_Data.getContents();
                j--;
            }

            //最后一天
            if(j==sheet.getRows()-1)
            {
                myMultimap.put(data_temp,num);

                List<Double> TempNum=num;

                //获取平均数
                double temp=0.0000000000000000000;
                Collections.sort(TempNum);
                for(int k=1;k<TempNum.size()-1;k++)
                {
                    temp+=TempNum.get(k);
                }

                temp/=1.000000*(TempNum.size()-2);

                //将平均数与日期对应起来
                Aver_Map.put(data_temp,temp);

            }

        }


        int map_length= myMultimap.size();
        System.out.println("总的长度是"+map_length);


        //获取excel表
        for (int j = 0; j < sheet.getRows(); j++) {
            //指定第row列的数据

            Cell cell = sheet.getCell(row,j);

            //如果是数字
            if (cell.getType() == CellType.NUMBER) {

                //获取日期，对应的平均数
                Cell cell_temp_Data = sheet.getCell(0,j);
                String temp_data=cell_temp_Data.getContents();

                //原数据
                NumberCell numberCell = (NumberCell) cell;
                Double numberValue = numberCell.getValue();

                //根据日期获取平均数
                Double value = (Double) Aver_Map.get(temp_data);
                System.out.println(temp_data+"  "+value);

                //每个数都减去平均值
                numberValue-=(Double) 1.000000*value;
                //每个数都取以平均值
                numberValue/=(Double) 1.000000*value;

                System.out.println(numberValue+" ");
                out.println(numberValue);

            }

            //如果不是数字，比如是空格，照搬到新表格
            else {
                String cellinfo = sheet.getCell(row,j).getContents();
                out.println(cellinfo);
            }
        }





        out.flush();
        out.close();//关闭流
        fWriter.close();
        out.flush();//刷新缓存
        System.out.println("输出完成！");
    }
}