# Word 自动化示例代码

## Word模板解析

定义了一套word模板规则，从word中提取待输入项。可以给到前端生成表单，作为输入源。模板规则是json形式提供，其中name, var_name, type属性是必填项，其他可以自定义。

其中包含word主角段落读取，标题和内容的树状结构解析，混合文本的json提取等示例代码。

## Word模板填充

支持在段落中替换简单文本，表格中替换简单文本，整段富文本替换，插入新表格操作。

其中包括段落的增删改查，表格的新建和合并单元格等操作的示例代码。

# 开发环境

已配置vscode容器开发环境，需要：

    vscode
    docker
    插件：Dev Containers

用vscode打开项目，如果插件正常运行会提示 Reopen in contanier。如果没有提示也可以Ctrl+Shift+P，搜索 open folder in container， 选择项目文件夹即可。

# 后记

使用了apache.poi库，可以说是非常难用，在Stackoverflow查一些问题的时候，有人说它的API very ugly, 深以为然。也是因为这个缘故，整出一些我自己都感觉 ugly 的奇技淫巧。如果你有更好的实现方式，请不吝赐教。