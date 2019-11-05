package com.xo.util;

import com.alibaba.fastjson.JSON;
import org.junit.Test;
import org.junit.Before;
import org.junit.After;

import java.util.*;

/**
 * EasyExcelUtil Tester.
 *
 * @author <Authors name>
 * @version 1.0
 * @since <pre>十月 21, 2019</pre>
 */
public class EasyExcelUtilTest {

    @Before
    public void before() throws Exception {
    }

    @After
    public void after() throws Exception {
    }

    /**
     * Method: readExcelData(String filePath, int sheetNo)
     * 使用lmask对读取数据进行重新排序组装
     */
    @Test
    public void testReadExcelDatalMask() throws Exception {
        String filePath = "E://HONG/1.xlsx";
        List<Integer> lMask=new ArrayList<>();
        lMask.add(2);   //第一个数据返回excel对应表格的第3个数据
        lMask.add(1);   //第二个数据返回excel对应表格的第2个数据
        lMask.add(0);   //第三个数据返回excel对应表格的第1个数据
        EasyExcelUtil.setlMask(lMask);//匹配不成功则不生效，返回全部数据
        List<Object> objects=EasyExcelUtil.readExcelData(filePath, 1);
        //System.out.println("Result:"+ JSON.toJSONString(objects));
        objects.forEach(System.out::println);
    }

    @Test
    public void testReadExcelDatamMask() throws Exception {
        String filePath = "E://HONG/1.xlsx";
        Map<String,Integer> mMask=new HashMap<>();
        mMask.put("Name",-1);   //匹配标题列是Name的项
        mMask.put("Class",-1);  //匹配标题列是Class的项
        mMask.put("Age",-1);    //匹配标题列是Age的项
        EasyExcelUtil.setmMask(mMask);//匹配不成功则不生效，返回全部数据
        List<Object> objects=EasyExcelUtil.readExcelData(filePath, 1);
        //System.out.println("Result:"+ JSON.toJSONString(objects));
        objects.forEach(System.out::println);
    }

    @Test
    public void testReadExcelDatadMask() throws Exception {
        String filePath = "E://HONG/1.xlsx";
        List<EasyExcelMask> dMask=new ArrayList<>();
        dMask.add(new EasyExcelMask("Name","B"));   //匹配标题列是Name的项，Key值转换为B（不设第二个参数，则为标题列数据）
        dMask.add(new EasyExcelMask("Age","A"));    //匹配标题列是Age的项，Key值转换为A
        dMask.add(new EasyExcelMask("Class","C"));  //匹配标题列是Class的项，Key值转换为C
        String[] mm={"1","2","3"};
        dMask.add(new EasyExcelMask("D",mm));            //匹配标题列是{1、2、3}的项，选择{}中第一个匹配数据，Key值转换为D
        EasyExcelUtil.setdMask(dMask);
        List<Object> objects=EasyExcelUtil.readExcelData(filePath, 2);
        //System.out.println("Result:"+ JSON.toJSONString(objects));
        objects.forEach(System.out::println);
    }

    @Test
    public void testReadExcelDataLoadMask() throws Exception {
        String filePath = "E://HONG/1.xlsx";
        String path = "E://HONG/1.xml";         //从配置文件中加载挡板组装配置数据
        EasyExcelMask load=new EasyExcelMask();
        List<EasyExcelMask> dMask =load.LoadMask(path,"LoadTest");//加载根节点->LoadTest节点下的Element数据，第一个属性或MaskName为mask,第二个属性或FakeName为fakemask,Text为matchmask
        EasyExcelUtil.setdMask(dMask);
        List<Object> objects=EasyExcelUtil.readExcelData(filePath, 2);
        //objects.forEach(System.out::println);
        System.out.println("Object[0]:"+JSON.toJSON(objects.get(0)));
    }


    @Test
    /**List属性*/
    public void testReadExcelDataLoadMask2() throws Exception {
        String filePath = "E://HONG/1.xlsx";
        String path = "E://HONG/1.xml";         //从配置文件中加载挡板组装配置数据
        EasyExcelMask load=new EasyExcelMask();
        List<EasyExcelMask> dMask =load.LoadMask(path,"nodelistname");//加载根节点->LoadTest节点下的Element数据，第一个属性或MaskName为mask,第二个属性或FakeName为fakemask,Text为matchmask
        EasyExcelUtil.setdMask(dMask);
        List<Object> objects=EasyExcelUtil.readExcelData(filePath, 3);
        objects.forEach(System.out::println);
        //System.out.println("Object[0]:"+JSON.toJSON(objects.get(0)));
    }


    //@Test
    public void testWriteExcel() throws Exception {
        String filePath = "E://HONG/4.xlsx";
        List<List<Object>> data = new ArrayList<>();
        data.add(Arrays.asList("111","222","333"));
        data.add(Arrays.asList("111","222","333"));
        data.add(Arrays.asList("111","222","333"));
        List<String> head = Arrays.asList("表头1", "表头2", "表头3");
        //head.addAll(Arrays.asList("表头2", "表头2", "表头2"));
        EasyExcelUtil.writeExcel(filePath,data,head,2);
    }
} 
