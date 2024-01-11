package com.laotie.app;
import java.util.ArrayList;
import java.util.List;
import java.util.Stack;

public class JsonExtract {
    // public static void main(String[] args) throws Exception {
    //     String text = "Some text before {\"key1\":\"value1\",\"key2\":{\"nested\":\"nested\"}} some text in between {\"key2\":\"value2\"} some text after.";

    //     // 使用 Jackson 解析 JSON
    //     ObjectMapper objectMapper = new ObjectMapper();
    //     JsonNode rootNode = objectMapper.readTree(text);

    //     // 遍历 JsonNode 对象，提取多个 JSON 结构
    //     for (JsonNode jsonNode : rootNode) {
    //         System.out.println("Extracted JSON: " + jsonNode);
    //     }
    // }

    // public static void main(String[] args) {
    //     String mixedContent = "Some text before {\"key1\":\"value1\",\"key2\":{\"nested\":\"nested\"}} some text in between {\"key2\":\"value2\"} some text after.";
    //     Pattern pattern = Pattern.compile("\\{.*?\\}");
    //     Matcher matcher = pattern.matcher(mixedContent);

    //     while (matcher.find()) {
    //         System.out.println(matcher.group());
    //     }
    // }

    public static void main(String[] args) {
        String mixedContent = "Some text before {\"key1\":\"value1\",\"key2\":{\"nested\":\"nested\"}}} some text in between {\"key2\":\"value2\"} some text after.";
        for (String json: extractJson(mixedContent)){
            System.out.println(json);
        }
    }

    /**
     * 提取json信息
     * 
     * @param content
     * @return
     */
    private static List<String> extractJson(String mixedContent) {
        List<String> result = new ArrayList<>();
        Stack<Integer> stack = new Stack<>();
        for (int i = 0; i < mixedContent.length(); i++) {
            if (mixedContent.charAt(i) == '{') {
                stack.push(i);
            } else if (mixedContent.charAt(i) == '}') {
                if (stack.isEmpty()) {
                    // System.out.println("Invalid JSON string: unmatched closing bracket at position " + i);
                    continue;
                }
                int start = stack.pop();
                String json = mixedContent.substring(start, i + 1);
                if (stack.size() == 0) {
                    result.add(json);
                }
            }
        }
        // if (!stack.isEmpty()) {
        //     System.out.println("Invalid JSON string: unmatched opening bracket at position " + stack.pop());
        // }
        return result;
    }
}
