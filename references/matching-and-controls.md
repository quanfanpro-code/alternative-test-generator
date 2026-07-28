# 供应商匹配与控制公式

## 名称匹配分层

### 第一层：exact_key（自动）

对名称执行以下规范化后全字符串比较：
1. Unicode NFKC 标准化
2. 全角空格(U+3000) → 半角空格(U+0020)
3. 首尾空白删除
4. 连续空白折叠为单个空格
5. 英文大小写统一（ASCII 范围内）
6. 全角数字/字母 → 半角

规范化后字符串相等 **且** 供应商实体唯一时 → `exact_unique`（自动）

### 第二层：候选生成（需人工确认）

以下操作只生成 `needs_confirmation` 候选：
- 去除括号及内容（中英文括号）
- 去除"有限公司""股份有限公司""有限责任公司""分公司""厂""（普通合伙）"等后缀
- 删除标点符号后比较
- 子串包含
- Levenshtein 编辑距离 ≤ 2
- 摘要文本中包含名称片段

候选排序只为展示，不自动选择第一名。

### 第三层：unmatched

以上均不匹配 → 保留原记录在未匹配清单，不删除、不记到"其他供应商"。

## 跨实体隔离

- 同一 raw_name 在不同公司/账套下分别匹配
- 用户决定不可全局复用到其他客户

## 勾稽公式

- 借方正常余额：expected_ending = opening + debit_movement - credit_movement
- 贷方正常余额：expected_ending = opening + credit_movement - debit_movement
- difference = reported_ending - expected_ending
- opening 缺失时状态 NOT_RUN
- 默认使用精确 Decimal("0.00") 容差
