# 经济学刷题分类系统

这是一个基于 Streamlit 的本地程序，可以：

- 读取课件 PPT（Lecture 01, 02, ...）并提取主题关键词
- 读取题库 DOC 并自动提取选择题、选项、答案、Topic
- 按 Lecture 内容自动分类题目，无法匹配的进入"未分类"
- 提供按 Lecture 刷题与解析查看
- 提供 Final 抽测模式（从微观题中随机抽取 30 题）
- 自动评分并输出错题解析
- 自动维护错题本：答错加入、二次答对自动移除

## 运行步骤

1. 安装依赖

```bash
pip install -r requirements.txt
```

2. 启动程序

```bash
streamlit run app.py
```

3. 首次进入页面后，在左侧确认三个目录路径：

- 课件目录（PPT）
- 题库目录（选择题 DOC）
- 解析目录（Solution Manual DOC）

点击"重建数据"即可自动抽取与分类。

## 数据文件

程序会在项目目录下生成 `.data/`：

- `quiz_dataset.json`：解析与分类后的题库
- `progress.json`：做题记录与错题本

## 说明

- 程序依赖 macOS 自带 `textutil` 命令来解析 `.doc` 文本。
- Final 抽测中的"微观题"默认按章节号 `<= 18` 判断。
