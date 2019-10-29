package com.xo.util;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.util.CollectionUtils;
import com.alibaba.excel.util.StringUtils;

import java.io.*;
import java.util.*;

/**
 * Create By FanxiaoQie
 * Email Address:xo.fanxiaoqie@qq.com
 *
 * 基于easyexcel对excel中的数据进行读写
 * 并根据配置转换成需要的格式
 * EasyExcelUtil中存储上一次读取数据的配置信息
 * 若读取配置一样，则从历史数据中获得结果
 * 手动置reReadFlag为true，强制重读
 * 设置筛选挡板也会重置重读标识
 *
 */
public class EasyExcelUtil {
    private static Sheet defaultSheet;  //默认表格“sheet”
    private static int defaultHeadLineNum=1;//默认表头行数

    private static boolean reReadFlag=true; //重新读取标识
    private static String preSheetPath="";    //上一次读取的表格路径
    private static int preSheetNo=0;          //上一次读取的表格号
    private static int preHeadLineNum=0;      //上一次读取的起始行号
    private static int preDefHeadLineNum=0;   //上一次读取的标题行号

    public static List<Object> preDatas=new ArrayList<>();
    public static List<Object> preHead=new ArrayList<>();
    public static ExcelListener preListener=new ExcelListener();


    public static List<Integer> lMask=new ArrayList<>();//lMask对Excel中的数据进行排序组装（优先）
    public static Map<String, Integer> mMask=new HashMap<>();//mMask对Excel中的数据进行匹配组装
    public static List<EasyExcelMask> dMask=new ArrayList<>();

    public static int getDefaultHeadLineNum() {
        return defaultHeadLineNum;
    }

    public static void setDefaultHeadLineNum(int defaultHeadLineNum) {
        EasyExcelUtil.defaultHeadLineNum=defaultHeadLineNum;
    }

    static {
        defaultSheet=new Sheet(1, 0);//读取全部数据，根据defaultHeadLineNum分成head和datas
//        defaultSheet.setSheetName("sheet");
        defaultSheet.setAutoWidth(Boolean.TRUE);//设置自适应宽度
    }

    public static boolean isReReadFlag() {
        return reReadFlag;
    }

    public static void setReReadFlag(boolean reReadFlag) {
        EasyExcelUtil.reReadFlag=reReadFlag;
    }

    public static String getPreSheetPath() {
        return preSheetPath;
    }

    public static void setPreSheetPath(String preSheetPath) {
        EasyExcelUtil.preSheetPath=preSheetPath;
    }

    public static int getPreSheetNo() {
        return preSheetNo;
    }

    public static void setPreSheetNo(int preSheetNo) {
        EasyExcelUtil.preSheetNo=preSheetNo;
    }

    public static int getPreHeadLineNum() {
        return preHeadLineNum;
    }

    public static void setPreHeadLineNum(int preHeadLineNum) {
        EasyExcelUtil.preHeadLineNum=preHeadLineNum;
    }

    public static int getPreDefHeadLineNum() {
        return preDefHeadLineNum;
    }

    public static void setPreDefHeadLineNum(int preDefHeadLineNum) {
        EasyExcelUtil.preDefHeadLineNum=preDefHeadLineNum;
    }

    public static void setlMask(List<Integer> lM) {
        if (lM.size() > 0) {//若设置mask参数，则对数据进行重读
            reReadFlag=true;
        }
        lMask=lM;
    }

    public static void setmMask(Map<String, Integer> mM) {
        if (mM.size() > 0) {//若设置mask参数，则对数据进行重读
            reReadFlag=true;
        }
        mMask=mM;
    }

    public static void setdMask(List<EasyExcelMask> dM) {
        if (dM.size() > 0) {//若设置mask参数，则对数据进行重读
            reReadFlag=true;
        }
        EasyExcelUtil.dMask=dM;
    }

    /**
     * 功能函数分隔线
     ********************************/
    //根据表格号选择表格
    public static List<Object> readExcelData(String filePath, int sheetNo) {
        Sheet selectSheet=new Sheet(sheetNo, 0);
        return readExcel(filePath, selectSheet).getDatas();
    }

    //根据表格号选择表格
    public static ExcelListener readExcel(String filePath, int sheetNo) {
        Sheet selectSheet=new Sheet(sheetNo, 0);
        return readExcel(filePath, selectSheet);
    }

    /**
     * 最底层的Excel读取接口
     * 参数filePath：输入文件绝对路径
     * 参数sheet：表格
     * 返回值ExcelListener
     * getDatas为List<Object>类型的内容数据
     * getHead为List<Object>类型的标题数据
     * <p>
     * 读取成功会保存历史数据，并将reReadFlag置为false,
     */
    public static ExcelListener readExcel(String filePath, Sheet sheet) {
        if (!StringUtils.hasText(filePath)) {
            return null;
        }
        sheet=sheet != null ? sheet : defaultSheet;
        InputStream fileStream=null;
        try {
            if (reReadFlag == false && filePath == preSheetPath && sheet.getSheetNo() == preSheetNo &&/**preHeadLineNum==0&&*/defaultHeadLineNum == preDefHeadLineNum) {
                return preListener;//若路径，表号,起始位置(恒为0，不判断)，标题位置都没有变化，且重读标识未开则返回上一次读取数据
            } else {
                fileStream=new FileInputStream(filePath);
                ExcelListener excelListener=new ExcelListener();
                if (lMask.size() > 0) {
                    excelListener.setlMask(lMask);
                }
                if (mMask.size() > 0) {
                    excelListener.setmMask(mMask);
                }
                if (dMask.size() > 0) {
                    excelListener.setdMask(dMask);
                }
                EasyExcelFactory.readBySax(fileStream, sheet, excelListener);

                preSheetPath=filePath;
                preSheetNo=sheet.getSheetNo();
                //preHeadLineNum = 0;
                preDefHeadLineNum=defaultHeadLineNum;
                preListener.reset(excelListener);
                return excelListener;
            }
        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException:" + e.toString());
        } finally {
            try {
                if (fileStream != null) {
                    fileStream.close();
                }
            } catch (IOException e) {
                System.out.println("excel read fail:" + e.toString());
            }
        }
        return null;
    }

    /**
     * Excel输出接口1，根据输入数据进行输出（数据输出会重新生成excel文件-待解决）
     * 参数filePath：输出文件绝对路径
     * 参数data：数据
     * 参数head：标题数据（这里只有一行）
     * 参数sheet：输出表格位置
     */
    public static void writeExcel(String filePath, List<List<Object>> data, List<String> head, Sheet sheet) {
        sheet=(sheet != null) ? sheet : defaultSheet;
        if (head != null) {
            List<List<String>> list=new ArrayList<>();
            head.forEach(h -> list.add(Collections.singletonList(h)));
            sheet.setHead(list);//输出标题数据
            //ResetReadFalg(filePath,sheet.getSheetNo());//重置重读标识
        }
        OutputStream outputStream=null;
        ExcelWriter writer=null;
        try {
            outputStream=new FileOutputStream(filePath);
            writer=EasyExcelFactory.getWriter(outputStream);
            writer.write1(data, sheet);//输出数据
            ResetReadFalg(filePath, sheet.getSheetNo());//重置重读标识
        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException[1]:" + e.toString());
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                System.out.println("excel write[1] fail::" + e.toString());
            }
        }
    }

    public static void writeExcel(String filePath, List<List<Object>> data, List<String> head, Integer sheetNo) {
        Sheet selectSheet=new Sheet(sheetNo, 0);
        selectSheet.setSheetName("aaa");
        writeExcel(filePath, data, head, selectSheet);
    }

    /**
     * Excel输出接口2，根据输入数据模型进行输出
     * 参数filePath：输出文件绝对路径
     * 参数data：输入数据模型
     * 参数sheet：输出表格位置
     */
    public static void writeExcel(String filePath, List<? extends BaseRowModel> data, Sheet sheet) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        sheet=(sheet != null) ? sheet : defaultSheet;
        sheet.setClazz(data.get(0).getClass());
        OutputStream outputStream=null;
        ExcelWriter writer=null;
        try {
            outputStream=new FileOutputStream(filePath);
            writer=EasyExcelFactory.getWriter(outputStream);
            writer.write(data, sheet);//输出数据
            ResetReadFalg(filePath, sheet.getSheetNo());//重置重读标识
        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException[2]:" + e.toString());
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }
                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                System.out.println("excel write[2] fail::" + e.toString());
            }
        }
    }

    public static void writeExcel(String filePath, List<? extends BaseRowModel> data, Integer sheetNo) {
        Sheet selectSheet=new Sheet(sheetNo, defaultHeadLineNum);

    }

    /**
     * Excel输出接口3，多表输出
     */
    public static void writeExcel(String filePath, List<SheetData> sheetDataList) {
        if (CollectionUtils.isEmpty(sheetDataList)) {
            return;
        }

        OutputStream outputStream=null;
        ExcelWriter writer=null;
        try {
            outputStream=new FileOutputStream(filePath);
            writer=EasyExcelFactory.getWriter(outputStream);
            for (SheetData sheetData : sheetDataList) {
                //Sheet sheet = sheetData.getSheet() != null ? sheetData.getSheet() : defaultSheet;
                Sheet sheet=sheetData.getSheet();
                if (sheet == null) {//若表格输出配置不正确，则该表格数据不输出
                    break;
                }
                if (!CollectionUtils.isEmpty(sheetData.getData())) {
                    sheet.setClazz((Class<? extends BaseRowModel>) sheetData.getData().getClass());
                }
                writer.write((List<? extends BaseRowModel>) sheetData.getData(), sheet);
            }

        } catch (FileNotFoundException e) {
            System.out.println("FileNotFoundException[3]:" + e.toString());
        } finally {
            try {
                if (writer != null) {
                    writer.finish();
                }

                if (outputStream != null) {
                    outputStream.close();
                }
            } catch (IOException e) {
                System.out.println("excel write[3] fail::" + e.toString());
            }
        }

    }

    /**
     * 若路径和表号与上一次配置数据一致，则重读标识置为true
     */
    public static void ResetReadFalg(String filePath, Integer sheetNo) {
        if (filePath == preSheetPath && sheetNo == preSheetNo && reReadFlag == false) {
            reReadFlag=true;//若输出文件与上一次结果文件配置相同，则重读标志需要重置
        }
    }


}
