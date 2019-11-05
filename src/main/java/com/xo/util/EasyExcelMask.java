package com.xo.util;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.sql.Date;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * EasyExcelMask类
 * 用于EasyExcel 标题筛选的数据结构
 * 与筛选配置xml文件匹配
 *
 * 例：
 *<MaskList>
 *     <Name1>
 *          <Element MaskName="AName" FakeName="A" Type="int" List="DList">AA,AB,AC</Element>
 *     <Name1/>
 *</MaskList>
 *      读取配置文件中MaskList根节点下的数据
 *      根据输入的Name，加载其下面的所有Element数据,组装成List<EasyExcelMask>
 *     maskStr=AName(第一个property属性,MaskName优先选择)
 *     fakemask=A (第二个property属性,FakeName优先选择)
 *     typeStr=int(Type属性，多属性且名字不规范的情况下，Type属性不要排在前两个，有可能会影响maskStr和fakemask)
 *      listStr=DList(List属性只能通过xml配置文件加载使用)
 *     matchStr=AA,AB,AC(text根据','转为list)
 * */
public class EasyExcelMask {
    private String maskStr;         //单个匹配（优先）
    private String fakeMask;        //别名 针对freemarker导出模板时，适配转换的变量名
    private String[] matchStr;      //范围匹配（若maskStr为空字符串时进行范围匹配）
    private String typeStr;         //类型属性 默认为String（空值为String类型）
    private String listStr;         //数据组装格式，按照list数据组装，与excel中上下文数据有关
    private Integer index;

    public EasyExcelMask(){
        maskStr="";
        fakeMask="";
        typeStr="";
        listStr="";
        matchStr=null;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m){
        this.maskStr=m;
        this.fakeMask="";
        this.typeStr="";
        this.listStr="";
        matchStr=null;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String[] ms ){
        this.maskStr="";
        this.fakeMask="";
        this.typeStr="";
        this.listStr="";
        matchStr=ms;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm){
        this.maskStr=m;
        this.fakeMask=fm;
        this.typeStr="";
        this.listStr="";
        matchStr=null;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String fm,String[] ms ){
        this.maskStr="";
        this.fakeMask=fm;
        this.typeStr="";
        this.listStr="";
        this.matchStr =ms;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm,String[] ms ){
        this.maskStr=m;
        this.fakeMask=fm;
        this.typeStr="";
        this.listStr="";
        this.matchStr =ms;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm,String[] ms ,String t){
        this.maskStr=m;
        this.fakeMask=fm;
        this.typeStr=t;
        this.listStr="";
        this.matchStr =ms;
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm,String ms ){
        this.maskStr=(m==null)?"":m;
        this.fakeMask=(fm==null)?"":fm;
        this.typeStr="";
        this.listStr="";
        if(ms !=null){
            String msStr[]=ms.split(",");
            this.matchStr =msStr;
        }else{
            matchStr=null;
        }
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm,String ms ,String t){
        this.maskStr=(m==null)?"":m;
        this.fakeMask=(fm==null)?"":fm;
        this.typeStr=(t==null)?"":t;
        this.listStr="";
        if(ms !=null){
            String msStr[]=ms.split(",");
            this.matchStr =msStr;
        }else{
            matchStr=null;
        }
        index=-1;//默认的对应值（异常对应值）
    }

    public EasyExcelMask(String m,String fm,String ms ,String t,String l){
        this.maskStr=(m==null)?"":m;
        this.fakeMask=(fm==null)?"":fm;
        this.typeStr=(t==null)?"":t;
        this.listStr=(l==null)?"":l;
        if(ms !=null){
            String msStr[]=ms.split(",");
            this.matchStr =msStr;
        }else{
            matchStr=null;
        }
        index=-1;//默认的对应值（异常对应值）
    }

    public String getMaskStr() {
        return maskStr;
    }

    public void setMaskStr(String maskStr) {
        this.maskStr=maskStr;
    }

    public String getFakeMask() {
        return fakeMask;
    }

    public void setFakeMask(String fakeMask) {
        this.fakeMask=fakeMask;
    }

    public String[] getMatchStr() {
        return matchStr;
    }

    public void setMatchStr(String[] matchStr) {
        this.matchStr=matchStr;
    }

    public void setTypeStr(String typeStr) {
        this.typeStr=typeStr;
    }

    public void setListStr(String listStr) {
        this.listStr=listStr;
    }

    public String getListStr() {
        return listStr;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index=index;
    }

    public void setIndex(List<Object> org) {
        int len = org.size();
        int match_len = matchStr==null?0:matchStr.length;
        if(this.maskStr!="") {//在List中查找匹配maskStr的下标，匹配到则赋值为第一个符合条件的下标
            for(int i=0;i<len;i++){
                if(org.get(i)!=null&&org.get(i).equals(this.maskStr)){
                    this.index=i;
                    return;
                }
            }
        }else if(match_len>0){//在List中查找匹配matchStr的下标，匹配到则赋值
            for(int j=0;j<match_len;j++){
                for(int i=0;i<len;i++){
                    if(org.get(i)!=null&&org.get(i).equals(this.matchStr[j])){
                        if(this.fakeMask==""){//若匹配值没有fakemask,则自动填第一个匹配值
                            this.fakeMask = this.matchStr[j];
                        }
                        this.index=i;
                        return;
                    }
                }
            }
        }
        return;
    }

    public Map<String,Object> getMaskData(List<Object> org){
        Map<String,Object> ret = new HashMap<>();
        Object value;
        if(this.index>=0&&this.index<org.size()&&this.typeStr!=null){
            if(this.typeStr.equals("Integer")||this.typeStr.equals("int")){//typeStr默认初始化为""
                try {
                    value=Integer.valueOf(org.get(this.index).toString());//转int
                }catch (Exception e){
                    //System.out.println("NumberFormat ERROR[" + ((this.maskStr==null)?"":this.maskStr) + ":" +((this.fakeMask==null)?"":this.fakeMask)+"]");//不打印异常信息
                    value=org.get(this.index);
                }
            }else if(this.typeStr.equals("Date")){
                try {
                    value=Date.valueOf(org.get(this.index).toString());//转date
                }catch (Exception e){
                    //System.out.println("DateFormat ERROR[" + ((this.maskStr==null)?"":this.maskStr) + ":" +((this.fakeMask==null)?"":this.fakeMask)+"]");//不打印异常信息
                    value=org.get(this.index);
                }
            }else{
                value=org.get(this.index);
            }
        }else{
            value ="";
        }
        if(this.fakeMask!=""){
//            ret.put(this.fakeMask,value);
            if(this.listStr.equals("")){
                ret.put(this.fakeMask,value);
            }else{//若是List数据，则组装成List<Map<>>格式
                Map<String,Object> newMap = new HashMap<>();
                List<Map<String,Object>> listData = new ArrayList<>();
                newMap.put(this.fakeMask,value);
                listData.add(newMap);
                ret.put(this.listStr,listData);
            }
            return ret;
        }else {
            if(this.maskStr!=""){
//                ret.put(this.maskStr,value);
                if(this.listStr.equals("")){
                    ret.put(this.maskStr,value);
                }else{//若是List数据，则组装成List<Map<>>格式
                    Map<String,Object> newMap = new HashMap<>();
                    List<Map<String,Object>> listData = new ArrayList<>();
                    newMap.put(this.maskStr,value);
                    listData.add(newMap);
                    ret.put(this.listStr,listData);
                }
                return ret;
            }else{//未指定别名，且未找匹配的数据的情况
                //ret.put("",value);//不输出数据
                //System.out.println("Unknown MatchMaskData");//打印信息有点多
                return ret;
            }
        }
    }

    /**
     * 使用dom4j从xml配置文件中读取挡板数据
     *  xml格式
     *  *<MaskList>
     *  *     <Name1><!--函数入参2-->
     *  *          <Element MaskName="AName" FakeName="A" >AA,AB,AC</Element>
     *  *     <Name1/>
     *  *</MaskList>
     *  *      读取配置文件中MaskList根节点下的数据
     *  *      根据输入的Name，加载其下面的所有Element数据,组装成List<EasyExcelMask>
     *  *     maskStr=AName(第一个property属性,MaskName优先选择)
     *  *     fakemask=A (第二个property属性,FakeName优先选择)
     *  *     matchStr=AA,AB,AC(text根据','转为list)
     *
     *
     * */
    public List<EasyExcelMask> LoadMask(String FilePath,String Name) throws DocumentException {
        List<EasyExcelMask> ret = new ArrayList<>();
        //打开xml文件
        SAXReader reader = new SAXReader();
        Document document = reader.read(new File(FilePath));
        if(document!=null){
            Element root = document.getRootElement();//获得根节点
            List<Element> list = document.selectNodes("//"+Name +"/Element");
            for (Element e:list) {
                String matchStr = e.getText();
                String maskStr = e.attributeValue("MaskName");
                String fakeMask = e.attributeValue("FakeName");
                String typeStr = e.attributeValue("Type");
                String listStr = e.attributeValue("List");
                int attrNum = e.attributeCount();
                if(maskStr==null&&fakeMask==null&&attrNum>=2){//不存在MaskName、FakeName属性，但有两个属性值
                    maskStr=e.attribute(0).getValue();
                    fakeMask=e.attribute(1).getValue();
                }else if(maskStr==null&&fakeMask==null&&attrNum==1) {//不存在MaskName、FakeName属性，但有一个属性值
                    maskStr=e.attribute(0).getValue();
                }else if((maskStr==null||fakeMask==null)&&attrNum>=2){//有一个不存在，但有两个以上属性值，则将另一个赋值给未有数据的项
                    if(e.attribute(0).getName().equals("MaskName")){
                        fakeMask = e.attribute(1).getValue();
                    }else if(e.attribute(0).getName().equals("FakeName")){
                        maskStr = e.attribute(1).getValue();
                    }else if(e.attribute(1).getName().equals("MaskName")){
                        fakeMask = e.attribute(0).getValue();
                    }else if(e.attribute(1).getName().equals("FakeName")){
                        maskStr=e.attribute(0).getValue();
                    }
                }
//                EasyExcelMask newNode = new EasyExcelMask(maskStr,fakeMask,matchStr);
                EasyExcelMask newNode = new EasyExcelMask(maskStr,fakeMask,matchStr,typeStr,listStr);
                ret.add(newNode);
                //System.out.println("Text:"+ JSON.toJSONString(newNode));
            }
        }
        return ret;
    }
}
