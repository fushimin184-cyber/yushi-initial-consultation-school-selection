# 输入数据结构

推荐把每次首咨内容整理成以下 JSON 结构后再生成文档。

```json
{
  "date": "2026-09-20",
  "student_name": "XX同学",
  "intake_short": "27fall",
  "application_project": "硕士（27 Fall / 2027年秋季入学）",
  "regions": ["香港", "新加坡"],
  "background": {
    "university": "某医科大学",
    "major": "临床医学",
    "study_length": "5年制",
    "graduation_year": "2026",
    "school_note": "",
    "gpa": "加权平均分 80.27 / 100；绩点 2.62",
    "language": "备考雅思",
    "target_major": "公共卫生 / 健康科学 / 生物医学",
    "soft_background": "待补充"
  },
  "programs": [
    {
      "school_en": "The University of Hong Kong",
      "school_abbr": "HKU",
      "school_zh": "香港大学",
      "program_en": "Master of Public Health",
      "program_zh": "公共卫生硕士",
      "qs_rank": "XX",
      "academic_requirement": "按官网核验后填写",
      "ielts_requirement": "按官网核验后填写",
      "official_url": "https://...",
      "highlight": true
    }
  ],
  "timeline": {
    "exploration": [
      {"time": "全年", "task": "准备并完成雅思考试，目标……"}
    ],
    "materials": [
      {"time": "7-8月", "task": "开始撰写 PS/CV"}
    ],
    "pre_arrival": [
      {"time": "1-3月", "task": "收到 offer，确定入读意向"}
    ]
  }
}
```

## 缺失值规则
- 不知道：写“待补充”。
- 语言未考：写“待考/备考中”，并将建议目标分数和学校硬性要求区分。
- 不知道学生姓名：写“XX同学”。
- 不知道具体毕业年份但知道入学季：不要自行推断，先用用户已有信息。
