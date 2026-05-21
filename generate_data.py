# -*- coding: utf-8 -*-
"""
样例数据生成器
------------------------------------------------------------
模拟「从监区业务系统导出」的两张表（全部为脱敏/虚构数据）：
  1) 训练任务数据.xlsx  —— 罪犯训练任务完成情况
  2) 异常行为数据.xlsx  —— 罪犯异常行为记录

运行后文件会保存到 ./data/ 目录。
真实使用时，把这两张表换成你们系统真正导出的数据即可。
"""
import numpy as np
import pandas as pd
from pathlib import Path

SEED = 20260421                     # 固定随机种子，保证每次生成的数据一样，方便复现
np.random.seed(SEED)

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)

# ---- 基础字典 ----
SURNAMES = list("赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜")
SQUADS = ["一中队", "二中队", "三中队"]
TRAIN_CATS = {
    "体能队列": ["3公里跑", "体能考核", "队列训练"],
    "劳动技能": ["缝纫技能", "电工实操", "装配作业"],
    "教育改造": ["法制教育", "思想教育", "文化学习"],
}
N_INMATE = 60                       # 罪犯人数

# ---- 1. 生成罪犯名册（脱敏） ----
inmates = []
for i in range(1, N_INMATE + 1):
    inmates.append({
        "编号": f"F-{i:04d}",
        "姓名": np.random.choice(SURNAMES) + "某某",
        "中队": SQUADS[(i - 1) % 3],          # 三个中队各 20 人
    })

# ---- 2. 生成训练任务记录（每人 3 个类别各一条） ----
train_rows = []
for inm in inmates:
    base = np.random.normal(75, 12)            # 每个人的基础能力值
    for cat, tasks in TRAIN_CATS.items():
        plan = int(np.random.choice([16, 20, 22]))
        att_rate = float(np.clip(np.random.normal(0.90, 0.12), 0.4, 1.0))
        actual = int(round(plan * att_rate))
        att_pct = round(actual / plan * 100, 1)
        score = int(np.clip(np.random.normal(base, 8), 30, 100))
        pass_flag = "是" if (score >= 60 and att_pct >= 80) else "否"
        change = np.random.choice(["进步", "持平", "退步"], p=[0.40, 0.35, 0.25])
        train_rows.append({
            "编号": inm["编号"], "姓名": inm["姓名"], "中队": inm["中队"],
            "训练类别": cat, "任务名称": np.random.choice(tasks),
            "应参训次数": plan, "实参训次数": actual, "出勤率(%)": att_pct,
            "考核成绩": score, "是否达标": pass_flag, "较上月": change,
        })
train_df = pd.DataFrame(train_rows)

# ---- 3. 生成异常行为记录 ----
ABN_TYPES = ["违规违纪", "冲突打架", "自伤倾向", "情绪心理异常"]
DESC = {
    "违规违纪": ["私藏违禁品", "不服从管理", "违反作息纪律", "传递违规物品"],
    "冲突打架": ["与同监舍发生肢体冲突", "言语挑衅引发争执", "劳动中与他人冲突"],
    "自伤倾向": ["情绪崩溃有自伤言论", "发现自伤痕迹", "流露轻生念头"],
    "情绪心理异常": ["情绪持续低落", "睡眠严重异常", "拒绝沟通行为反常"],
}
MEASURES = ["谈话教育", "禁闭处理", "加强监管", "心理干预", "谈话教育+加强监管"]
N_ABN = 28
abn_rows = []
for _ in range(N_ABN):
    inm = inmates[np.random.randint(N_INMATE)]
    atype = np.random.choice(ABN_TYPES, p=[0.45, 0.25, 0.10, 0.20])
    # 风险等级与异常类型相关：自伤/冲突更容易高风险
    if atype == "自伤倾向":
        risk = np.random.choice(["中", "高"], p=[0.3, 0.7])
    elif atype == "冲突打架":
        risk = np.random.choice(["低", "中", "高"], p=[0.2, 0.5, 0.3])
    else:
        risk = np.random.choice(["低", "中", "高"], p=[0.5, 0.4, 0.1])
    # 高风险更可能仍在跟进
    if risk == "高":
        resolved = np.random.choice(["是", "跟进中"], p=[0.4, 0.6])
        measure = np.random.choice(["心理干预", "加强监管", "谈话教育+加强监管"])
    else:
        resolved = np.random.choice(["是", "否", "跟进中"], p=[0.7, 0.1, 0.2])
        measure = np.random.choice(MEASURES)
    abn_rows.append({
        "编号": inm["编号"], "姓名": inm["姓名"], "中队": inm["中队"],
        "发生日期": f"2026-04-{np.random.randint(1, 31):02d}",
        "异常类型": atype, "行为描述": np.random.choice(DESC[atype]),
        "风险等级": risk, "处置措施": measure, "是否化解": resolved,
    })
abn_df = pd.DataFrame(abn_rows).sort_values("发生日期").reset_index(drop=True)

# ---- 4. 保存 ----
train_path = DATA_DIR / "训练任务数据.xlsx"
abn_path = DATA_DIR / "异常行为数据.xlsx"
train_df.to_excel(train_path, index=False)
abn_df.to_excel(abn_path, index=False)

print(f"训练任务记录：{len(train_df)} 条 -> {train_path}")
print(f"异常行为记录：{len(abn_df)} 条 -> {abn_path}")
print("样例数据生成完成。")
