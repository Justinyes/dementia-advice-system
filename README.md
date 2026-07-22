# Dementia Advice System MVP

一个用于演示“对话式画像收集 + Excel 证据库 + 规则排序 + 中文建议生成”的轻量级 Web 原型。

这是一个本地可运行的“对话式健康建议系统”原型，实现了需求文档中的核心闭环：

1. 多轮对话收集个人信息
2. 将用户表达抽取成可追溯的个体画像
3. 从 Excel 证据库检索候选干预
4. 用确定性规则做个性化打分排序
5. 输出中文建议、适配理由、风险提示和证据引用

> 注意：本系统只提供生活方式与就医提示层面的健康教育建议，不做诊断、治疗方案、药物剂量或替代医生判断。

## Features

- 对话式输入：用户用自然语言描述年龄、慢病、运动限制、睡眠、社交、听力视力等情况。
- 可追溯画像：每个画像字段保留来源片段与置信度。
- Excel 证据库：首次运行会自动生成 `data/evidence.xlsx` 样例证据表。
- 可解释排序：根据证据等级、推荐强度、画像标签匹配和风险标签计算推荐分数。
- 红旗信号处理：遇到快速下降、自伤、明显生活功能受损等信号时，优先提示就医评估，同时给出保守生活管理参考。

## 运行

```bash
pip install -r requirements.txt
python3 app.py
```

然后打开：

```text
http://127.0.0.1:8000
```

## 项目结构

```text
app.py                  # 后端服务、画像抽取、证据检索、规则排序
data/evidence.xlsx      # Excel 证据库，首次运行若不存在会自动生成样例
static/index.html       # 对话界面
static/styles.css       # 页面样式
static/app.js           # 前端交互
```

## Example

可以在页面输入：

```text
我爸 72 岁，最近半年记性比以前差一些，经常忘记刚说过的话。平时不太爱运动，膝盖有点疼，但白天愿意出门，喜欢跟邻居打麻将。有高血压，血糖正常，睡眠一般，听力和视力还没检查过，没有跌倒，也没有突然快速变差。
```

系统会生成画像、命中标签，并展示个性化推荐排序。

## Medical Disclaimer

This project is for education, research, and prototyping only. It does not provide medical diagnosis, treatment, prescription, or emergency guidance. Users should consult qualified clinicians for medical decisions, especially when red-flag symptoms appear.

## License

MIT

## 后续扩展建议

- 将 `data/evidence.xlsx` 替换为真实指南、综述、RCT 整理出的证据库。
- 把 `extract_profile_updates` 替换为 LLM 结构化抽取，但保留“原文片段 + 置信度 + 合并规则”。
- 推荐生成阶段可以接入 OpenAI API，但只把排序结果和证据摘要交给模型，避免医学幻觉。
- 增加用户登录、会话存储、审计日志和医生/照护者复核界面。
