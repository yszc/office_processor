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
class Section {
    private String type;
    private String prefix = "";
    private String name;
    private List<Section> children;
    private JSONObject inputAttr;

    private String positionTitle;
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
            allInputs.removeIf(input -> input.containsKey("input_type") && input.getString("input_type").equals("complex"));
            allInputs.addAll(complexInputs);
            return allInputs;

        }
    }

    /**
     * 递归获得节点的输入模块
     * 
     * @param root
     * @return
     */
    private static List<JSONObject> _fetchAllInputAttr(Section root) {
        List<JSONObject> result = new ArrayList<>();
        Deque<Section> stack = new ArrayDeque<>();
        // JSONObject postion = new JSONObject();
        
        stack.push(root);
        
        while (!stack.isEmpty()) {
            Section current = stack.pop();
            for (Section child : current.getChildren()) {
                if ("title".equals(child.getType())) {
                    if (current.getPositionTitle()!= null) {
                        child.setPositionTitle(current.getPositionTitle());
                    }else{
                        child.setPositionTitle(child.getName());
                    }

                    stack.push(child);
                } else if ("input".equals(child.getType()) && child.getInputAttr() != null) {
                    if (current.getPositionTitle()!= null) {
                        child.setPositionTitle(current.getPositionTitle());
                    }else{
                        child.setPositionTitle(child.getName());
                    }
                    if (current.getPositionInputName()!= null) {
                        child.setPositionInputName(current.getPositionInputName());
                    } else {
                        child.setPositionInputName(child.getName());
                    }
                    if (current.getPositionVar()!= null) {
                        child.setPositionVar(current.getPositionVar());
                    } else {
                        child.setPositionVar(child.getInputAttr().getString("var_name"));
                    }

                    JSONObject inputAttr = child.getInputAttr();
                    inputAttr.put("position_title", child.getPositionTitle());
                    inputAttr.put("position_input_name", child.getPositionInputName());
                    inputAttr.put("position_var", child.getPositionVar());
                    result.add(inputAttr);
                }
            }
        }
        return result;
    }
    

    private List<JSONObject> _complex2List(JSONObject complexObject) {
        List<JSONObject> result = new ArrayList<>();
        JSONArray rows = complexObject.getJSONArray("rows");
        for (int i = 0; i < rows.size(); i++) {
            JSONArray row = rows.getJSONArray(i);
            for (int j = 0; j<row.size(); j++) {
                JSONObject item = row.getJSONObject(j);
                if (item.containsKey("validation")){
                    item.put("position_title", complexObject.getString("position_title"));
                    item.put("position_input_name", complexObject.getString("position_input_name"));
                    item.put("position_var", complexObject.getString("position_var"));
                }
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