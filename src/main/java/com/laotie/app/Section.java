package com.laotie.app;

import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.collections4.CollectionUtils;

import com.alibaba.fastjson2.JSON;
import com.alibaba.fastjson2.JSONArray;
import com.alibaba.fastjson2.JSONObject;
import com.alibaba.fastjson2.JSONWriter;

import lombok.Data;

@Data
public class Section {
    private String type;
    private String prefix = "";
    private String name;
    private List<Section> children;
    private JSONObject inputAttr;

    private String positionTitle;
    private String positionTitleL2;
    private String positionInputName;
    private String positionVar;

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

    public Section() {

    }

    /**
     * json 序列化
     * 
     * @return
     */
    public String toJson() {
        return JSON.toJSONString(this);
    }

    /**
     * json序列化
     * 
     * @param pretty
     * @return
     */
    public String toJson(Boolean pretty) {
        if (pretty) {
            return JSON.toJSONString(this, JSONWriter.Feature.PrettyFormat);
        }
        return this.toJson();
    }

    /**
     * 通过json字符串构建
     * 
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
     * 前端表单友好化适配，将不在title下的输入框放在附件封面
     * 
     * @return
     */
    public Section toFormFriendly() {
        Section generalInfo = new Section("title", "附件封面");
        List<Integer> indexForDel = new ArrayList<>();
        for (int i = 0; i < this.children.size(); i++) {
            Section child = this.children.get(i);
            if ("input".equals(child.getType())) {
                generalInfo.children.add(child);
                indexForDel.add(i);
            }
        }
        this.children.add(0, generalInfo);
        this.children = this.children.stream().filter(section -> "title".equals(section.getType()))
                .collect(Collectors.toList());
        return this;
    }

    /**
     * 获得所有的输入模块
     * 
     * @return
     */
    public List<JSONObject> fetchAllInputAttr() {
        return fetchAllInputAttr(true);
    }

    public List<JSONObject> fetchAllInputAttr(Boolean withComplex) {
        if (withComplex) {
            return _fetchAllInputAttr(this);
        } else {
            List<JSONObject> allInputs = _fetchAllInputAttr(this);
            List<JSONObject> complexInputs = allInputs.stream()
                    .filter(input -> input.containsKey("input_type") && input.getString("input_type").equals("complex"))
                    .map(input -> _complex2List(input))
                    .flatMap(List::stream)
                    .collect(Collectors.toList());
            allInputs.removeIf(
                    input -> input.containsKey("input_type") && input.getString("input_type").equals("complex"));
            allInputs.addAll(complexInputs);
            return allInputs;

        }
    }

    public String getInputName() {
        if (!this.type.equals("input") || this.getInputAttr() == null) {
            return null;
        }
        return !this.getName().isEmpty() ? this.getName() : this.getInputAttr().getString("name");
    }

    /**
     * 递归DFS的方式获得所有的输入模块
     * 
     * 同时获得每个输入模块的前面两级标题和最上层输入模块的定位信息
     * 
     * @param root
     * @return
     */
    private static List<JSONObject> _fetchAllInputAttr(Section current) {
        List<JSONObject> result = new ArrayList<>();

        for (Section child : current.getChildren()) {
            // 二级标题，仅当第一级标题非空的时候取
            if (current.getPositionTitle() != null) {
                child.setPositionTitleL2(
                        current.getPositionTitleL2() != null ? current.getPositionTitleL2() : child.getName());
            }
            child.setPositionTitle(
                    current.getPositionTitle() != null ? current.getPositionTitle() : child.getName());
            
            if ("title".equals(child.getType())) {
                result.addAll(_fetchAllInputAttr(child));
            } else if ("input".equals(child.getType()) && child.getInputAttr() != null) {
                child.setPositionInputName(current.getPositionInputName() != null ? current.getPositionInputName()
                        : child.getInputName());
                child.setPositionVar(current.getPositionVar() != null ? current.getPositionVar()
                        : child.getInputAttr().getString("var_name"));

                JSONObject inputAttr = child.getInputAttr();
                inputAttr.put("position_title", child.getPositionTitle());
                inputAttr.put("position_title_l2", child.getPositionTitleL2());
                inputAttr.put("position_input_name", child.getPositionInputName());
                inputAttr.put("position_var", child.getPositionVar());
                result.add(inputAttr);
            }

        }

        return result;
    }

    private List<JSONObject> _complex2List(JSONObject complexObject) {
        List<JSONObject> result = new ArrayList<>();
        JSONArray rows = complexObject.getJSONArray("rows");
        for (int i = 0; i < rows.size(); i++) {
            JSONArray row = rows.getJSONArray(i);
            for (int j = 0; j < row.size(); j++) {
                JSONObject item = row.getJSONObject(j);
                // if (item.containsKey("validation")) {
                item.put("position_title", complexObject.getString("position_title"));
                item.put("position_title_l2", complexObject.getString("position_title_l2"));
                item.put("position_input_name", complexObject.getString("position_input_name"));
                item.put("position_var", complexObject.getString("position_var"));
                // }
                result.add(item);
            }
            // result.addAll(row.toJavaList(JSONObject.class));
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