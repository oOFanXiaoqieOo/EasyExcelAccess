package com.xo.util;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.*;

/**
 * 解析监听器，
 * 每解析一行会回调invoke()方法。
 * 整个excel解析结束会执行doAfterAllAnalysed()方法
 *
 * @author:
 * @date:
 */
public class ExcelListener extends AnalysisEventListener {
    private List<Object> datas=new ArrayList<Object>();
    private List<Object> head=new ArrayList<Object>();

    private Map<String, Integer> mMask=new HashMap<>();
    private List<Integer> lMask=new ArrayList<>();
    private Integer mMaskValidNum=0;
    private Integer dMaskValidNum=0;
    private List<EasyExcelMask> dMask=new ArrayList<>();
    private boolean listFlag=false;
    private List<String> listStr=new ArrayList<>();       //列表属性的索引值
    private List<Integer> listOutIndex=new ArrayList<>();       //列表属性外的索引值，用于判断是否需要组装List数据
    private List<Object> preDataList=new ArrayList<>();       //存储上一次的结果数据
    private Map<String, Object> finalData=new HashMap<>();

    /**
     * 将非空行添加至结果集
     * object : 当前行的数据
     */
    @Override
    public void invoke(Object object, AnalysisContext context) {
        Object t_object = object;
        Map<String, Object> t_dmaskData = new HashMap<>();
        if (t_object != null) {//只处理非空行
            List<Object> orgList=new ArrayList<>();
            List<Object> emptyListStr=new ArrayList<>();
            if (t_object instanceof List<?>) {
                if (lMask.size() > 0 || mMask.size() > 0 || dMask.size() > 0) {
                    for (Object o : (List<?>) t_object) {//转换成List数据
                        orgList.add(Object.class.cast(o));
                        emptyListStr.add(null);
                    }
                }
                if (lMask.size() > 0) {//优先进行数据排序组装
                    t_object=dataListMask(orgList, lMask);
                } else if (mMask.size() > 0 && mMaskValidNum > 0) {//排序组装数据为空且匹配组装有效
                    t_object=dataListMask(orgList, mMask);
                } else if (dMask.size() > 0 && dMaskValidNum > 0) {
                    //t_object=dataListMask2(orgList, dMask);
                    t_dmaskData = dataListMask2(orgList, dMask);
                    t_object = t_dmaskData;
                }
            }
            if (context.getCurrentRowNum() >= EasyExcelUtil.getDefaultHeadLineNum()) {
                if (listFlag && listOutIndex.size() > 0&& listStr.size()>0) {//List属性存在时,且不全是List属性时
                    //当除List属性数据外 全为空或全与上一次值相同，则进行List组装；否则数据组装结束
                    if(ListUtil.IsListEqual(orgList,preDataList,listOutIndex)||ListUtil.IsListEqual(orgList,emptyListStr,listOutIndex)){
                        for(int i =0;i<listStr.size();i++){
                            String key =listStr.get(i);
                            ((List)finalData.get(key)).addAll((List)t_dmaskData.get(key));//在dataListMask2 中已对List属性数据打包成List格式
                        }
                    }else{
                        if(finalData.size()>0) {//第一个数据不输出
                            datas.add(finalData);
                        }
                        preDataList = orgList;
                        finalData = t_dmaskData;
                    }

                } else {
                    datas.add(t_object);
                }
            } else {//处理标题列信息
                if (mMask.size() > 0) {
                    MaskMaker(orgList);
                } else if (dMask.size() > 0) {
                    MaskMaker2(orgList);
                } else {
                    head.add(t_object);
                }
            }
        }
    }

    /**
     * 解析完所有数据后会调用该方法
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        //解析结束销毁不用的资源
        //System.out.println("context:"+JSON.toJSONString(context));
        EasyExcelUtil.setReReadFlag(false);
        if (listFlag && listOutIndex.size() > 0) {
            datas.add(finalData);
        }
    }

    public List<Object> getDatas() {
        return this.datas;
    }

    public List<Object> getHead() {
        return this.head;
    }

    public void setDatas(List<Object> datas) {
        this.datas.addAll(datas);
    }

    public void setHead(List<Object> head) {
        this.head.addAll(head);
    }

    public void setmMask(Map<String, Integer> mMask) {
        this.mMask.clear();
        this.mMask=mMask;
    }

    public void setlMask(List<Integer> lMask) {
        this.lMask.clear();
        this.lMask=lMask;
    }

    public void reset(ExcelListener newData) {
        this.getDatas().clear();
        this.getHead().clear();
        this.setDatas(newData.getDatas());
        this.setHead(newData.getHead());
    }

    public void resetdMask() {
        this.dMask.clear();
    }

    public void setdMask(List<EasyExcelMask> dMask) {
        this.dMask.addAll(dMask);
    }

    public void setdMask(EasyExcelMask dMask) {
        this.dMask.add(dMask);
    }

    /**
     * 封装两个List转换处理函数************************************
     * 将原有List根据Mask模板进行重组                             *
     ***********************************************************/
    //将org根据mask 的下标数据重新取出重组 下标从0开始,若下标越界则填空字符串
    public List<Object> dataListMask(List<Object> org, List<Integer> mask) {
        int len=org.size();
        int ret_len=mask.size();
        List<Object> ret=new ArrayList<>();
        for (int i=0; i < ret_len; i++) {
            int index=mask.get(i);
            if (index < len && index >= 0) {
                ret.add(org.get(index));
            } else {//若存在越界的下标，则填空字符串
                ret.add("");
            }
        }
        return ret;
    }

    //org 根据mask的value数据重新取出重组成Map数据
    //若mask中value 超出org长度范围时，数据填""
    public Map<String, Object> dataListMask(List<Object> org, Map<String, Integer> mask) {
        int len=org.size();
        Map<String, Object> ret=new HashMap<>();
        for (String key : mask.keySet()) { //根据key值遍历mask
            int index=mask.get(key);
            if (index < len && index >= 0) {
                ret.put(key, org.get(index));
            } else {//若存在越界的下标，则填空字符串
                ret.put(key, "");
            }
        }
        return ret;
    }

    //根据dMask生成重组数据
    public Map<String, Object> dataListMask2(List<Object> org, List<EasyExcelMask> mask) {
        int len=org.size();
        int len_mask=mask.size();
        Map<String, Object> ret=new HashMap<>();
        for (int i=0; i < len_mask; i++) { //遍历mask
            String maskKey = mask.get(i).getListStr();
            if(maskKey.length()>0){
                if(ret.containsKey(maskKey)){
                    //((List)ret.get(maskKey)).add(((List)mask.get(i).getMaskData(org).get(maskKey)).get(0));
                    ((Map<String,Object>)((List) ret.get(maskKey)).get(0)).putAll(((Map<String,Object>)((List)mask.get(i).getMaskData(org).get(maskKey)).get(0)));
                }else{//首次添加
                    ret.putAll(mask.get(i).getMaskData(org));
                }
            }else {
                ret.putAll(mask.get(i).getMaskData(org));
            }
        }
        return ret;
    }

    //将List中的数据 与mMask中的Key值进行匹配，若匹配到，则将value值置为该下标,mask初始value数据最好置为-1
    //存在多个匹配数据时，后面的数据生效
    void MaskMaker(List<Object> org) {
        int len=org.size();
        mMaskValidNum=0;
        for (int i=0; i < len; i++) {
            if (mMask.containsKey(org.get(i))) {
                mMask.put((String) org.get(i), i);
                mMaskValidNum++;
            }
        }
    }

    //根据dMask数据生成挡板
    //存在多个匹配数据时，前面的数据生效
    void MaskMaker2(List<Object> org) {
        //int len = org.size();
        int d_len=dMask.size();
        listFlag=false;       //初始化
        listOutIndex.clear();      //置空
        dMaskValidNum=0;
        EasyExcelMask t_mask=new EasyExcelMask();
        for (int i=0; i < d_len; i++) {
            t_mask=dMask.get(i);
            t_mask.setIndex(org);
            if (t_mask.getIndex() >= 0) {//默认初始index为-1，这里若大于等于0，说明匹配成功
                dMaskValidNum++;
                if (!t_mask.getListStr().equals("")) {//存在List属性的，有可能需要对Excel上下文数据进行组装
                    listFlag=true;
                    listStr.add(t_mask.getListStr());
                } else {
                    listOutIndex.add(i);//列表属性外的索引值
                }
            }
        }
    }


}
