# 清单格式（spec.json）

> 这份文件是**中间模型的格式定义**。它是整个流程的枢纽：
> 输入适配器往它里写，输出序列化器从它里读，自检拿它对账。
>
> ⚠️ **它是「唯一权威」。** `spec.md` 和标号图都是从它**渲染**出来的视图，不许直接改。

## 为什么要两层

| | 装什么 | 谁来定 |
|---|---|---|
| **层 1：在说什么** | `entities` / `relations` / `groups` | **事实** → 要用户确认 |
| **层 2：长什么样** | `layout` | **审美 + 约束** → AI 定，但"跟不跟模板"要问 |

**为什么必须分开**：

- PPT / SVG **两层都要**（没有位置画不出来）
- **Mermaid 只要层 1**（位置它自己排）
- **"这条是我推的"只能装在层 1**——它是关于**意思**的，不是关于位置的

## 三条硬规矩

### ① 每条信息都要说得出它从哪来

`provenance` 必填，取值只有六个：

| 值 | 意思 |
|---|---|
| `original` | **原件明写**（页面上、XML 里直接有） |
| `measured` | **我量出来的**（坐标、像素、计数——精确可复核） |
| `composed` | **我拼出来的**（要跨几个字段/几处证据才读得对） |
| `inferred` | **我推的**（从结构推出来的，原件没直说） |
| `domain` | **行内常识补的**（原件里没有，是外部知识，**最弱**） |
| `user` | **用户说的** |

⚠️ **除 `user` 外，每条都必须带 `evidence`**（指向原件里的什么）。留不下痕迹的结论，不许进清单。

### ② 改是「追加」，不是「覆盖」

用户说"这个空框里其实是审批"时：

```
e07.text       = ""          ← 原件读到的，保留
e07.text_user  = "审批"      ← 新增一条证据
e07.status     = "changed"
changes[]      += {from:"", to:"审批", by:"user", ...}
```

**绝不抹掉"原件是空的"**。两个原因：用户可能记错；而且交付时要能说清"我改了原件的什么"。

> 已栽过一次：用户答"是一条**折线**关系"，被记成"**关系数 = 10**"。
> 那次能发现错误，**唯一的原因是原话没被解读覆盖掉**。

### ③ 每个「意义单位」都要说清它靠哪些对象表达

层 2 的每个对象**必须**用 `for` 指回它在表达哪个实体/关系/分组。

⚠️ **不要求"一对一"**——一条关系可以画成 1 个对象，也可以画成几个拼起来。**画法是可以变的，这是自由**：

> dryrun-01 真发生过：原件用**两个**线条对象拼成一条折线，我们合并成**一条**带双向箭头的折线，
> 对象数 11 → 10，**视觉一样，而且用户没反对**。

**但"谁表达谁"必须写清楚**——不然机器算不出"输出里该有几个对象"，而那是 **Check A 唯一能机器做的一件事**：

```
输出里的对象数  ==  layout 的条目数 ？
```

⚠️ **顺带说明**：原件里**一条关系由几个对象拼成**，那是**读的证据**，记在 `evidence` 里（比如 `["p:sp#41", "p:sp#42"]`）——
**它和我们画成几个是两回事**，中间差着"画法自由"。这两个数**故意不要求相等**。

## 结构

```jsonc
{
  "spec_version": "0.2",

  "task":   { "id": "...", "round": 1, "created": "..." },

  "source": {
    "canvas": [2000, 1125],              // 层 2 用的画布
    "unit": "px@150dpi",
    "files": [
      { "path": "source/deck.pptx", "kind": "pptx_shapes",
        "slide": 5, "region": [x, y, w, h] }   // region 可省
    ]
  },

  // ───────── 层 1：在说什么 ─────────
  "entities": [
    {
      "id": "e01",
      "role": ["shape"],          // shape|container|relation_carrier|annotation|decoration
      "text": "下单",
      "text_user": null,          // 用户改过的版本（追加，不覆盖 text）
      "box_src": [752.9, 304.2, 1146.3, 521.0],  // 它在**原件**里的位置
      "group": "g01",
      "provenance": "original",
      "evidence": ["p:sp#12"],
      "status": "unconfirmed"     // unconfirmed|confirmed|changed
    }
  ],

  "relations": [
    {
      "id": "r01",
      "from": "e01", "to": "e02",
      "kind": "sequence",         // sequence|dependency|containment|association
      "direction": "forward",     // forward|backward|both|none
      "provenance": "inferred",
      "evidence": ["prst=rightArrow", "rot=5400000"],
      "status": "unconfirmed"
    }
  ],

  "groups": [
    { "id": "g01", "label": "工具服务层",
      "members": ["e01", "e02"], "provenance": "original", "evidence": ["..."] }
  ],

  // ───────── 层 2：长什么样 ─────────
  "layout": [
    {
      "id": "s01",
      "for": "e01",               // ★ 必填：这个对象在表达谁
      "kind": "roundrect",        // rect|roundrect|diamond|ellipse|line|brace|arrow|picture|text
      "prst": "roundRect",        // 预设几何名（见 187 个预设形状表）
      "box": [100.0, 200.0, 160.0, 60.0],   // line 用 p1/p2 代替
      "fill": "DCE9F7",           // null = 透明
      "line": { "color": "2E75B6", "pt": 0.5, "dash": "solid" },
      // lines = 段落（每个字符串是一段）。⚠️ 不是 DrawingML 的 "run"（段内文本块）
      "text": { "lines": ["下单"], "font_pt": 8, "face": "微软雅黑", "align": "ctr" },
      "rot": 0,
      "z": 10,
      "provenance": "measured",
      "evidence": ["p:sp#12@xfrm"]
    }
  ],

  // ───────── 台账（追加型，只增不改）─────────
  "changes": [
    { "at": "...", "round": 2, "target": "e07.text",
      "from": "开发环境", "to": "测试环境", "by": "user",
      "reason": "原件写错，用户确认修改", "evidence_ref": "u003" }
  ],

  "open_questions": [
    { "id": "q01", "about": ["r03"],
      "ask": "C10 和 C11 算一条关系还是两条？",
      "if_wrong": "关系数会在 10 和 11 之间变",
      "default": "按一条处理", "status": "asked" }
  ],

  "findings": [
    { "id": "f01", "kind": "self_contradiction",
      "what": "步骤 6 写「开发环境」，却站在测试环境上方",
      "evidence": ["..."], "handling": "原样照抄并告知用户" }
  ],

  "unsupported": [
    { "what": "自由曲线图标", "name": "Shape 42",
      "reason": "FREEFORM 画不出来", "handled": "列入清单告知用户" }
  ]
}
```

## 校验器查什么

`spec_check.py` 保证**清单自洽**（不保证清单对——那要靠 Check B）：

| # | 不变式 |
|---|---|
| 1 | 顶层字段齐全、`spec_version` 对得上 |
| 2 | **id 全局唯一**（entities / relations / groups / layout 之间也不许撞） |
| 3 | `relations.from/to` 指向存在的实体或组 |
| 4 | `groups.members` 指向存在的实体 |
| 5 | **`layout.for` 必填，且指向存在的东西** |
| 6 | **每个实体/关系/分组，至少被一个 layout 对象表达**（除非显式标 `not_drawn`） |
| 7 | 除 `user` 外，每条 `provenance` 都要有 `evidence` |
| 8 | `provenance` / `status` / `kind` / `role` 取值合法 |
| 9 | `changes` 的 `from`/`to`/`by` 齐全，`by ∈ {user, ai}` |
| 10 | `layout` 里同一条目不许同时有 `box` 和 `p1/p2`；`line` 必须用 `p1/p2` |

**第 6 条是为了逼出分歧**：**漏画了 vs 有意不画**，必须说清是哪个——
所以没被表达的实体/关系**必须显式写 `not_drawn: true`**，否则校验器报错。
