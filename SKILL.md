---
name: word-paper-revision
description: 面向论文改稿的 Word 文档修订技能。优先生成 Word 原生修订(Track Changes)效果;当用户需要修改论文、保留修订痕迹、导出带修订 DOCX、或比较原文与改文时使用。
---

# Word Paper Revision

用于对 `.docx` 论文进行带痕迹修改,优先输出 **Word 原生修订模式** 文档。

## 当前能力

脚本位置:`scripts/track_changes.py`

当前版本走 **最小增量 OOXML** 路线,尽量只改 `document.xml` 和 `settings.xml`,避免破坏 Word 原始外壳。

支持六类动作：

1. `--replace OLD NEW`
   - 把一个**整段精确匹配**的段落替换为新段落
   - 输出为 Word 的删除 + 插入修订

2. `--replace-inline OLD NEW`
   - 在正文段落中，把一段**局部文本**替换为新文本
   - 输出为「前缀正文 + 删除修订 + 插入修订 + 后缀正文」
   - **默认替换所有命中**：如果同一短语在同一段内出现多次，会连续替换并保留多处修订痕迹
   - **支持同一段内多条局部替换规则串行执行**：可连续写多个 `--replace-inline`

3. `--replace-inline-nth OLD NEW N`
   - 只替换命中的**第 N 处**局部文本
   - 当前按段落从前到后扫描，命中后停止
   - 适合"这一段里同一个词出现很多次，但我只想改第 2 处"这种场景

4. `--delete OLD`
   - 删除一个**整段精确匹配**的段落
   - 输出为 Word 的删除修订

5. `--insert-after ANCHOR TEXT`
   - 在一个**整段精确匹配**的锚点段落后插入新段落
   - 输出为 Word 的插入修订

6. `--comment TARGET TEXT`
   - 在包含 TARGET 的文本位置插入批注
   - 批注内容为 TEXT，作者使用 `--author` 指定

7. `--replace-hf OLD NEW`
   - 在页眉/页脚中替换所有命中的局部文本
   - 逻辑与 `--replace-inline` 相同，但只作用于页眉和页脚

3. `--delete OLD`
   - 删除一个**整段精确匹配**的段落
   - 输出为 Word 的删除修订

4. `--insert-after ANCHOR TEXT`
   - 在一个**整段精确匹配**的锚点段落后插入新段落
   - 输出为 Word 的插入修订

## 用法

### 1. 整段替换

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace "原段落文本" "新段落文本"
```

### 2. 段内局部替换

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline "较好的预测精度" "更高的预测精度"
```

### 3. 只替换第 n 处局部命中

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline-nth "动力响应" "振动响应" 2
```

### 4. 删除段落

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --delete "需要删除的整段文本"
```

### 5. 在指定段落后插入

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --insert-after "锚点段落文本" "新增段落文本"
```

### 6. 添加批注

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "审稿人" \
  --comment "有限元" "建议补充软件名称和版本信息" \
  --comment "动力响应" "此处应补充参考文献"
```

在包含 TARGET 文本的 run 位置插入批注气泡,支持正文段落和表格单元格。执行后会输出批注日志,显示批注位置。

### 7. 替换页眉/页脚中的文本

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-hf "Company Name" "New Company" \
  --replace-hf "2023" "2024"
```

页眉和页脚中所有段落统一应用 `--replace-hf` 规则(与 `--replace-inline` 相同的替换逻辑),替换结果写入日志。

### 8. 一次执行多个动作

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace "旧段落 A" "新段落 A" \
  --replace-inline "旧短语" "新短语" \
  --replace-inline-nth "动力响应" "振动响应" 2 \
  --replace-hf "旧机构名" "新机构名" \
  --comment "有限元" "建议补充参考文献" \
  --delete "旧段落 B" \
  --insert-after "段落 C" "新增段落 D"
```

### 7. 同一段内多处局部替换

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline "较好的预测精度" "更高的预测精度" \
  --replace-inline "较强的泛化能力" "更强的泛化能力"
```

执行后,脚本会额外输出**局部替换日志**,显示命中了第几段、命中了几处、目标模式(`all` 或 `occurrence=N`),以及使用的是 `single-run` 还是 `cross-run` 替换路径。

### 8. 替换表格单元格中的文本

```bash
python3 scripts/track_changes.py input.docx output.docx \
  --author "Yachiyo" \
  --replace-inline "有限元" "有限差分"
```

运行后,日志会显示表格命中的具体位置(如 `T1-R2-C3`):
```text
局部替换日志:
- T1-R2-C3: '有限元' -> '有限差分' 命中 1 处 [all] (single-run)
```

## 自测方法

先确认脚本能运行:

```bash
python3 scripts/track_changes.py --help
python3 -m py_compile scripts/track_changes.py
```

建议用一个简单 DOCX 做冒烟测试:

- 文档中准备 3~4 段纯文本正文
- 对一整段做 `--replace`
- 对一句话中的短语做 `--replace-inline`
- 对另一段做 `--delete`
- 对最后一段做 `--insert-after`
- 用 Word 打开后检查:
  - 是否出现修订痕迹
  - 审阅模式下能否看到插入/删除
  - 是否仍然没有"发现无法读取的内容"提示

## 当前限制

- `--replace` 和 `--delete` 仍然是 **整段精确匹配**
- `--replace-inline` / `--replace-inline-nth` 支持 **单 run 内替换** 和 **跨 run 替换** 两种模式:
  - 优先尝试在单个 run 内匹配(保持原有样式)
  - 如果目标文本跨越多个 run,自动进行跨 run 替换
  - 跨 run 替换时会合并涉及的 run,使用首个 run 的样式
  - `--replace-inline` 会在同一段内持续重扫并尽量全部替换
  - `--replace-inline-nth` 当前按段落从前到后寻找第 N 处,命中后停止
- 跨 run 替换时,如果目标文本跨越 field code(域代码,如引用、公式等),会自动跳过以避免破坏文档结构
- 正文段落、**表格单元格**、**页眉/页脚**内段落均支持替换、删除、插入、批注操作
  - 表格内的段落用 `T{table}-R{row}-C{col}` 标注(如 `T1-R2-C3` 表示第 1 个表第 2 行第 3 列的单元格)
  - 替换日志中同时显示 cell_ref 位置信息
  - 页眉/页脚替换通过 `--replace-hf` 命令单独指定
- 暂不处理:文本框、脚注尾注
- 复杂格式段落(公式、域、超链接、混合 run 样式)虽然尽量保留首个 run 样式,但不能保证完全无差异
- 目前还不支持"只替换第 N 处后继续替换第 M 处"这类更复杂的命中编排

## 适用场景

- 论文初稿 / 返修稿中的正文段落改写
- 把一句话中的措辞替换成更学术的表述,同时保留修订痕迹
- 需要给导师或合作者提交"带修订痕迹"的 Word 文件

## 后续可扩展方向

- 支持单段多处局部替换 ✅
- 支持第 N 处命中控制 `--replace-inline-nth` ✅
- 支持表格单元格段落 ✅
- 支持批注 `--comment` ✅
- 支持页眉/页脚 `--replace-hf` ✅
- 支持文本框、脚注尾注
- 为复杂段落提供红字 / 删除线回退模式
- 支持 Word 原生审阅视图下的"接受/拒绝修订"自动化
