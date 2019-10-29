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

    private Map<String,Integer> mMask = new HashMap<>();
    private List<Integer> lMask = new ArrayList<>();
    private Integer mMaskValidNum=0;
    private Integer dMaskValidNum=0;
    private List<EasyExcelMask> dMask= new ArrayList<>();
    /**
     * 将非空行添加至结果集
     * object : 当前行的数据
     */
    @Override
    public void invoke(Object object, AnalysisContext context) {
        if (object != null) {//只处理非空行
            List<Object> orgList=new ArrayList<>();
            if(object instanceof List<?>){
                if (lMask.size() > 0||mMask.size() > 0||dMask.size() > 0) {
                    for (Object o : (List<?>) object) {//转换成List数据
                        orgList.add(Object.class.cast(o));
                    }
                }
                if (lMask.size() > 0) {//优先进行数据排序组装
                    object = dataListMask(orgList,lMask);
                } else if (mMask.size() > 0&&mMaskValidNum>0) {//排序组装数据为空且匹配组装有效
                    object = dataListMask(orgList,mMask);
                }else if(dMask.size()>0&&dMaskValidNum>0){
                    object = dataListMask2(orgList,dMask);
                }
            }
            if (context.getCurrentRowNum() >= EasyExcelUtil.getDefaultHeadLineNum()) {
                datas.add(object);
            } else {//处理标题列信息
                if (mMask.size() > 0){
                    MaskMaker(orgList);
                }else if(dMask.size() > 0){
                    MaskMaker2(orgList);
                }else{
                    head.add(object);
                }
            }
        }
    }

    /**
     * 解析完所有数据后会调用该方法
     *
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        //解析结束销毁不用的资源
        //System.out.println("context:"+JSON.toJSONString(context));
        EasyExcelUtil.setReReadFlag(false);
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

    public void reset(ExcelListener newData){
        this.getDatas().clear();
        this.getHead().clear();
        this.setDatas(newData.getDatas());
        this.setHead(newData.getHead());
    }

    public void resetdMask(){
        this.dMask.clear();
    }
    public void setdMask(List<EasyExcelMask> dMask) {
        this.dMask.addAll(dMask);
    }

    public void setdMask(EasyExcelMask dMask) {
        this.dMask.add(dMask);
    }

    /**封装两个List转换处理函数************************************
     * 将原有List根据Mask模板进行重组                             *
     ***********************************************************/
    //将org根据mask 的下标数据重新取出重组 下标从0开始,若下标越界则填空字符串
    public List<Object> dataListMask(List<Object> org,List<Integer> mask) {
        int len=org.size();
        int ret_len=mask.size();
        List<Object> ret=new ArrayList<>();
        for (int i=0; i < ret_len; i++) {
            int index=mask.get(i);
            if (index < len&&index>=0) {
                ret.add(org.get(index));
            } else {//若存在越界的下标，则填空字符串
                ret.add("");
            }
        }
        return ret;
    }

    //org 根据mask的value数据重新取出重组成Map数据
    //若mask中value 超出org长度范围时，数据填""
    public Map<String, Object> dataListMask(List<Object> org,Map<String, Integer> mask) {
        int len= org.size();
        Map<String, Object> ret=new HashMap<>();
        for (String key : mask.keySet()) { //根据key值遍历mask
            int index=mask.get(key);
            if (index < len&&index>=0) {
                ret.put(key, org.get(index));
            } else {//若存在越界的下标，则填空字符串
                ret.put(key, "");
            }
        }
        return ret;
    }

    //根据dMask生成重组数据
    public Map<String, Object> dataListMask2(List<Object> org,List<EasyExcelMask> mask) {
        int len= org.size();
        int len_mask = mask.size();
        Map<String, Object> ret=new HashMap<>();
        for (int i=0;i<len_mask;i++) { //遍历mask
            ret.putAll(mask.get(i).getMaskData(org));
        }
        return ret;
    }

    //将List中的数据 与mMask中的Key值进行匹配，若匹配到，则将value值置为该下标,mask初始value数据最好置为-1
    //存在多个匹配数据时，后面的数据生效
    void MaskMaker(List<Object> org){
        int len = org.size();
        mMaskValidNum=0;
        for(int i = 0;i<len;i++){
            if(mMask.containsKey(org.get(i))){
                mMask.put((String) org.get(i),i);
                mMaskValidNum++;
            }
        }
    }

    //根据dMask数据生成挡板
    //存在多个匹配数据时，前面的数据生效
    void MaskMaker2(List<Object> org){
        //int len = org.size();
        int d_len = dMask.size();
        dMaskValidNum=0;
        EasyExcelMask t_mask = new EasyExcelMask();
        for(int i = 0;i<d_len;i++){
            t_mask = dMask.get(i);
            t_mask.setIndex(org);
            if(t_mask.getIndex()>=0){//默认初始index为-1，这里若大于等于0，说明匹配成功
                dMaskValidNum++;
            }
        }
    }



}
