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
    private String inputReplace;

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

    public String toJson()  {
        return JSON.toJSONString(this);
    }

    public String toJson(Boolean pretty){
        if (pretty){
            return JSON.toJSONString(this, JSONWriter.Feature.PrettyFormat);
        }
        return this.toJson();
    }

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