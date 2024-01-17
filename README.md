# Word 自动化示例代码

## Word模板解析

定义了一套word模板规则，从word中提取待输入项。可以给到前端生成表单，作为输入源。模板规则是json形式提供，其中name, var_name, type属性是必填项，其他可以自定义。

其中包含word主角段落读取，标题和内容的树状结构解析，混合文本的json提取等示例代码。

## Word模板填充

支持在段落中替换简单文本，表格中替换简单文本，整段富文本替换，插入新表格操作。

其中包括段落的增删改查，表格的新建和合并单元格等操作的示例代码。

# 模板规范

## 模板占位符
模板占位符是一串json，应以单行格式插入在word模板文件中，用以下json编辑器进行格式化校验并压缩为单行：

http://www.esjson.com/

一个正确的占位符结构如下：

```Json
{
  "name": "企业名称",
  "var_name": "ent_name",
  "input_type": "text",
  "placeholder": "填写正确的企业名称",
  "validation": [
      ...
  ]
}
```

其中，**name, var_name, input_type** 是必选字段，其他属性可根据需求自由设置。

## 简单文本

```Json
{
  "name": "企业名称",
  "var_name": "ent_name",
  "input_type": "text",
  "placeholder": "填写正确的企业名称"
}
```

## 单选Radio

```Json
{
  "name": "是否有饭醉记录",
  "var_name": "is_crime",
  "input_type": "radio",
  "input_des": {
    "options":[
      {
        "name":"是",
        "value":"是"
      },
      {
        "name":"否",
        "value":"否"
      }
    ]
  }
}
```

## 富文本

```Json
{
  "name": "企业名称",
  "input_type": "WYSIWYG",
  "var_name": "ent_name",
  "placeholder": "填写正确的企业名称"
}
```

## 多行表格
```Json
{
  "name": "占比数据",
  "var_name": "z_table",
  "placeholder": "请填写占比数据",
  "input_type": "table",
  "input_des": {
    "columns": [
      {
        "name": "企业名称",
        "var_name": "z_ent_name",
        "placeholder": "填写正确的企业名称",
        "input_type": "text",
        "validation": [
          {
            "type": "required",
            "msg": "企业名称必填"
          },
          {
            "type": "regexp",
            "regexp": ".{1,50}",
            "msg": "企业名称为1-50个字符"
          },
          {
            "type": "regexp",
            "regexp": "[a-zA-Z]*",
            "msg": "企业名称只能由英文字母自称"
          }
        ]
      },
      {
        "name": "统一社会信用代码",
        "var_name": "z_credit_code",
        "placeholder": "填写正确的信用代码",
        "input_type": "text"
      },
      {
        "name": "比例(%)",
        "var_name": "z_part",
        "placeholder": "请填写比例数据",
        "input_type": "text",
        "validation": [
          {
            "type": "required",
            "msg": "占比必填"
          },
          {
            "type": "regexp",
            "regexp": "[0-9]*",
            "msg": "只能填写数字"
          }
        ]
      }
    ],
    "footer": [
      {
        "type": "const",
        "content": "合计",
        "colspan": 2
      },
      {
        "type": "sum",
        "sum_col": "z_part",
        "validation": [
          {
            "type": "range",
            "range": {
              "==": 100
            },
            "msg": "合计必须等于100"
          }
        ]
      }
    ]
  }
}
```

# 开发环境

已配置vscode容器开发环境，需要：

    vscode
    docker
    插件：Dev Containers

用vscode打开项目，如果插件正常运行会提示 Reopen in contanier。如果没有提示也可以Ctrl+Shift+P，搜索 open folder in container， 选择项目文件夹即可。

# 后记

使用了apache.poi库，可以说是非常难用，在Stackoverflow查一些问题的时候，有人说它的API very ugly, 深以为然。也是因为这个缘故，整出一些我自己都感觉 ugly 的奇技淫巧。如果你有更好的实现方式，请不吝赐教。
