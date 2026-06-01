# KPI Reference — SA Edit Pattern Analysis

This document defines every KPI shown in the dashboard, including exact thresholds and how each metric is computed.

---

## 1. Edit Rate

| KPI | Definition |
|---|---|
| Total AI-generated messages | Count of all rows in the uploaded dataset |
| Total SA-edited messages | Rows where `is_change = true` in the `copy_to_chat_history` JSON field |
| Edit rate | `edited / total generated × 100` |

---

## 2. Location of Change

Measures *where* in the AI message the SA made edits. Position is calculated as the character offset in the **original Chinese AI message** (not the translation).

| Zone | Threshold |
|---|---|
| Opening | First **30%** of the AI message character length |
| Middle | **30–70%** of the AI message character length |
| Closing | Last **70–100%** of the AI message character length |

A single edited message can contribute to multiple zones. Percentages represent the share of edited messages that had at least one change in that zone.

---

## 3. Length Change

Computed on **Chinese character counts** of the original AI message vs. the sent message. Not word count.

| Label | Threshold |
|---|---|
| Shortened | Sent message is **> 25% fewer** characters than AI message |
| Expanded | Sent message is **> 25% more** characters than AI message |
| No significant change | Difference is **< 10%** in either direction |
| *(unlabelled moderate)* | 10–25% change — counted but not shown as a separate bucket |

**Avg characters** rows show the mean Chinese character count for the AI message and sent message, filtered to rows in that length bucket.

---

## 4. Content Type

All keyword detection runs on the **Chinese diff text** — specifically the characters added and removed between the AI message and the sent message. English translations are not used.

A message contributes to a category if any keyword from that category appears in the changed (added or removed) characters.

### Keyword Lists

**Greeting / Salutation**
亲爱, 您好, 你好, 小姐, 姐姐, 女士, 好久不见, 先生, 太太, 早上好, 下午好, 晚上好, 嗨, 哈喽, 久违

**Call-to-Action**
有空, 看看, 联系, 随时, 过来, 试试, 试戴, 带来, 方便, 来店, 到店, 欢迎, 预约, 安排, 拜访, 有机会, 欢迎光临

**Product References**
珠宝, 卡地亚, 系列, 腕表, 戒指, LOVE, Love, 猎豹, 玫瑰, 钻石, Juste, Clou, 手镯, 项链, 手表, 吊坠, 耳环, 胸针, 蓝气球, 三环, Trinity, 钉子, Santos, Tank, Clash, 宝石, 项圈

**Emoji**
Detected via Unicode ranges — no text keywords. Covers: emoticons (U+1F600–1F64F), misc symbols (U+1F300–1F5FF), transport (U+1F680–1F6FF), misc (U+2600–27BF).

> Note: Scarcity/Urgency, After-sales/Care, and Seasonal/Occasion categories exist in the underlying data model but are excluded from the dashboard view.

---

## 5. SA Distribution

Each SA's edit rate is computed as `edited messages / total AI-generated messages` for that SA, using the **full uploaded dataset** as the denominator (not just the edited subset).

| Bucket | Range |
|---|---|
| Never edited | Edit rate = **0%** |
| Low edit rate | **1–20%** |
| Moderate edit rate | **20–75%** |
| High edit rate | **> 75%** |

**Mean / Median edit rate** are computed across all SAs who received at least 1 AI-generated message.

---

## Notes

- All analysis is performed **client-side in the browser** — no data is sent to any server.
- Translation (Chinese → English) uses Google Translate's public endpoint and is used solely to populate the translated columns in downloaded files. No KPI is computed from translated text.
- The character-level diff uses a Levenshtein edit distance algorithm operating on individual Unicode code points.
