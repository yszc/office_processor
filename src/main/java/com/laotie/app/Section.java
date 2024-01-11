package com.laotie.app;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONObject;
import com.alibaba.fastjson2.JSONWriter;

import lombok.Data;

@Data
class Section {
    private String type;
    private String prefix = "";
    private String name;
    private List<Section> children;
    private JSONObject inputAttr;

    /**
     * 创建一个子节点
     * 
     * @param type
     * @param name
     * @return
     */
    public Section(String type, String name) {
        this.type = type;
        this.name = name;
        this.children = new ArrayList<>();
    }

    public Section(){

    }

    /**
     * json 序列化
     * @return
     */
    public String toJson()  {
        return JSON.toJSONString(this);
    }

    /**
     * json序列化
     * @param pretty
     * @return
     */
    public String toJson(Boolean pretty){
        if (pretty){
            return JSON.toJSONString(this, JSONWriter.Feature.PrettyFormat);
        }
        return this.toJson();
    }

    /**
     * 通过json字符串构建
     * @param jsonResult
     * @return
     */
    public static Section fromJson(String jsonResult) {
        return JSON.parseObject(jsonResult, Section.class);
    }

    /**
     * 获取非空的结构
     * 
     * @return
     */
    public Section toNoneEmpty() {
        return filterSection(this);
    }

    /**
     * 获得所有的输入模块
     * @return
     */
    public List<JSONObject> fetchAllInputAttr(){
        return _fetchAllInputAttr(this);
    }

    /**
     * 递归获得节点的输入模块
     * @param root
     * @return
     */
    private static List<JSONObject> _fetchAllInputAttr(Section root){
        List<JSONObject> result = new ArrayList<>();

        for(Section child: root.getChildren()){
            if(child.getType() == "title"){
                List<JSONObject> childInput = _fetchAllInputAttr(child);
                result.addAll(childInput);
            }else if (child.getType() == "input" && null != child.getInputAttr()){
                result.add(child.getInputAttr());
            }
        }
        return result;
    }

    /**
     * 过滤所有空标题
     * 
     * @param root
     * @return
     */
    private static Section filterSection(Section root) {
        if (null == root.getChildren()) {
            return root;
        }
        for (Section child : root.getChildren()) {
            filterSection(child);
        }
        root.setChildren(new ArrayList<>(CollectionUtils.select(root.getChildren(),
                child -> null != child && (child.getType() == "input" || child.getChildren().size() > 0))));
        return root;
    }

}