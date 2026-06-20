# Agent Handoff

这个目录只保留复现当前周表所需的最少文件。

## 你会接收到什么

- `inputs/2024级人工智能专业培养方案.pdf`
- `inputs/sustech_2026_fall_undergrad_courses.xlsx`
- `inputs/completed-courses-history.png`
- `inputs/completed-courses-current-term.png`

## 你需要生成什么

运行：

```bash
python3 scripts/build_ai_course_schedule.py
```

生成：

- `outputs/ai_major_remaining_weekly_timetable_optimized.xlsx`

这是当前唯一保留的最终产物。

## 当前已写死的前提

- 专业课程清单写死在 `scripts/build_ai_course_schedule.py`
- 已修专业课集合也写死在脚本里
- 当前已修集合来自用户规则：
  - 两张截图并集
  - 明确退掉的课不算已修

## 关键规则

1. 课程归属以培养方案为准，不以开课表里的课程类别为准。
2. `CS331` 只按“可能对应 `COE301`”处理，不能默认完全等价。
3. `AI303` 单独处理，因为时间信息不完整。
4. 周表里是否算组课，只看地点里是否含 `机房`：
   - 含 `机房`：组课 / 实验课
   - 不含 `机房`：大课
5. 周表按双节显示：`1-2`、`3-4`、`5-6`、`7-8`、`9-10`。

## 用户最好补充什么

- 最新培养方案
- 当前学期开课表
- 已修课清单或截图
- 哪些课退掉了
- 是否存在替代认定关系

## 如果要继续改

优先改两件事：

1. 把已修课集合从脚本里抽出来
2. 把培养方案课程清单从脚本里抽出来
