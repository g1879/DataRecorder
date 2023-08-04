`Recorder`和`Filler`都可以对 xlsx 文件单元格设置样式，本节介绍`CellStyle`对象用法。

## 🐘 导入和创建对象

```python
from DataRecorder.style import CellStyle

c = CellStyle()
```

---

## 📝 使用方法

`CellStyle`内置 7 种设置项目，分别是：

- `font`：文本格式设置
- `border`：边框格式设置
- `alignment`：对齐选项设置
- `pattern_fill`：图案填充设置，与`gradient_fill`互斥
- `gradient_fill`：渐变填充设置，与`pattern_fill`互斥
- `number_format`：数字格式设置
- `protection`：单元格保护设置

每种设置项下有若干子项目，用`set_xxxx()`方法设置。

**示例：**

设置字体颜色为红色。

```python
c.font.set_color('red') 
```

---

## 🎨 颜色格式

字体、边框、背景涉及到颜色的设置，这里支持以下几种格式：

### 🖌️ 颜色名称

以下几种颜色可以直接用名字设置

- `'white'`：白色
- `'black'`：黑色
- `'red'`：红色
- `'green'`：绿色
- `'blue'`：蓝色
- `'purple'`：紫色
- `'yellow'`：黄色
- `'orange'`：橙色

```python
from DataRecorder.style import CellStyle

c = CellStyle()
c.font.set_color('red')
```

---

### 🖌️ 颜色代码

可以用`str`或`tuple`传入十六进制、十进制的颜色代码。

```python
from DataRecorder.style import CellStyle

c = CellStyle()
c.font.set_color('FFF000')  # 十六进制代码
c.font.set_color('255,255,0')  # 十进制代码，str格式
c.font.set_color((255, 255, 0))  # 十进制代码，tuple格式
```

---

### 🖌️ 使用`Color`对象

`Color`对象是 openpyxl 内置对象，除了颜色，还可以设置透明度等。

具体使用方法见 openpyxl 文档。

```python
from DataRecorder.style import CellStyle, Color

color = Color('FFF000')  # 创建Color对象
style = CellStyle()
style.font.set_color(color)  # 用Color对象设置颜色
```

---

## ✅ `font`设置

此属性用于设置单元格字体样式。

### 📌 `font.set_name()`

此方法用于设置文本使用的字体。

|  参数名称  |       类型        | 默认值 | 说明                |
|:------:|:---------------:|:---:|-------------------|
| `name` | `str`<br>`None` | 必填  | 字体名称，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_charset()`

此方法用于设置字体编码，如何设置参考 openpyxl。

|   参数名称    |       类型        | 默认值 | 说明                        |
|:---------:|:---------------:|:---:|---------------------------|
| `charset` | `int`<br>`None` | 必填  | 字体编码，`int`格式，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_size()`

此方法用于设置字体大小。

|  参数名称  |        类型         | 默认值 | 说明                |
|:------:|:-----------------:|:---:|-------------------|
| `size` | `float`<br>`None` | 必填  | 字体大小，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_bold()`

此方法用于设置是否加粗。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_italic()`

此方法用于设置是否斜体。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_strike()`

此方法用于设置是否有删除线。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_outline()`

此方法用于设置 outline。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_shadow()`

此方法用于设置是否有阴影。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_condense()`

此方法用于设置 condense。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_extend()`

此方法用于设置 extend。

|   参数名称   |        类型        | 默认值 | 说明                      |
|:--------:|:----------------:|:---:|-------------------------|
| `on_off` | `bool`<br>`None` | 必填  | `bool`表示开关，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_color()`

此方法用于设置字体颜色。格式：`'FFFFFF'`, `'255,255,255'`, `(255, 255, 255)`, `Color`对象均可，`None`表示恢复默认。

|  参数名称   |                  类型                   | 默认值 | 说明                |
|:-------:|:-------------------------------------:|:---:|-------------------|
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 字体颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_underline()`

此方法用于设置下划线类型。可选 `'single'`, `'double'`, `'singleAccounting'`, `'doubleAccounting'`，`None`表示恢复默认。

|   参数名称   |       类型        | 默认值 | 说明                 |
|:--------:|:---------------:|:---:|--------------------|
| `option` | `str`<br>`None` | 必填  | 下划线类型，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_vertAlign()`

此方法用于设置字体上下标类型。

可选：`'superscript'`, `'subscript'`, `'baseline'`，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|----|
|  `option`  | `str`<br>`None` | 必填 | 上下标类型，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `font.set_scheme()`

此方法用于设置 scheme 类型。

可选：`'major'`, `'minor'`，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|------------------------|
|  `option`  | `str`<br>`None` | 必填 | scheme 类型，`None`表示恢复默认 |

**返回：**`None`

---

## ✅ `border`设置

此项用于设置单元格边框线形和颜色。

线形可选：`'dashDot'`, `'dashDotDot'`, `'dashed'`, `'dotted'`, `'double'`, `'hair'`, `'medium'`, `'mediumDashDot'`, `'mediumDashDotDot'`, `'mediumDashed'`, `'slantDashDot'`, `'thick'`, `'thin'`，`None`表示恢复默认。

颜色格式：`'FFFFFF'`, `'255,255,255'`, `(255, 255, 255)`, `Color`对象均可，`None`表示恢复默认。

### 📌 `border.set_start()`

此方法用于设置 start 属性。

| 参数名称 |             类型             | 默认值 | 说明 |
|:----:|:--------------------------:|:---:|---------------------|
|  `style`  |      `str`<br>`None`       | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_end()`

此方法用于设置 end 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_left()`

此方法用于设置左边框。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_right()`

此方法用于设置有边框。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_top()`

此方法用于设置上边框。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_bottom()`

此方法用于设置下边框。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_diagonal()`

此方法用于设置对角线。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_vertical()`

此方法用于设置垂直中线。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_horizontal()`

此方法用于设置水平中线。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
|  `style`  | `str`<br>`None` | 必填 | 在类型种选择，`None`表示恢复默认 |
| `color` | `str`<br>`tuple`<br>`Color`<br>`None` | 必填  | 颜色，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `border.set_outline()`

此方法用于设置 outline 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool` | 必填 | `bool`表示开关 |

**返回：**`None`

---

### 📌 `border.set_diagonalDown()`

此方法用于设置 diagonalDown 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool` | 必填 | `bool`表示开关 |

**返回：**`None`

---

### 📌 `border.set_diagonalUp()`

此方法用于设置 diagonalUp 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool` | 必填 | `bool`表示开关 |

**返回：**`None`

---

## ✅ `alignment`设置

此属性用于设置单元格对齐方式。

### 📌 `alignment.set_horizontal()`

此方法用于设置水平位置。

可选：`'general'`, `'left'`, `'center'`, `'right'`, `'fill'`, `'justify'`, `'centerContinuous'`, `'distributed'`，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `horizontal` | `str`<br>`None` | 必填 | 在选项中选择，`None`表示恢复默认 |

**返回：**`None`

---

### 📌 `alignment.set_vertical()`

此方法用于设置垂直位置。

可选：`'top'`, `'center'`, `'bottom'`, `'justify'`, `'distributed'`，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `vertical` | `str`<br>`None` | 必填 | 在选项中选择，`None`表示恢复默认。 |

**返回：**`None`

---

### 📌 `alignment.set_indent()`

此方法用于设置缩进。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `indent` | `int` | 必填 | 缩进数值，`0`到`255` |

**返回：**`None`

---

### 📌 `alignment.set_relativeIndent()`

此方法用于设置相对缩进。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `indent` | `int` | 必填 | 缩进数值，`-255`到`255` |

**返回：**`None`

---

### 📌 `alignment.set_justifyLastLine()`

此方法用于设置 justifyLastLine。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool`<br>`None` | 必填 | `bool`表示开关，`None`恢复默认 |

**返回：**`None`

---

### 📌 `alignment.set_readingOrder()`

此方法用于设置 readingOrder。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `int` | 必填 | 不小于`0`的数字 |

**返回：**`None`

---

### 📌 `alignment.set_textRotation()`

此方法用于设置文本旋转角度。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `int` | 必填 | 可输入`0`到`180` |

**返回：**`None`

---

### 📌 `alignment.set_wrapText()`

此方法用于设置是否自动换行。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool`<br>`None` | 必填 | `bool`表示开关，`None`恢复默认 |

**返回：**`None`

---

### 📌 `alignment.set_shrinkToFit()`

此方法用于设置 shrinkToFit。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool`<br>`None` | 必填 | `bool`表示开关，`None`恢复默认 |

**返回：**`None`

---

## ✅ `pattern_fill`设置

此属性用于设置图案填充方式。

与`gradient_fill`互斥，会清除已有`gradient_fill`设置。

### 📌 `pattern_fill.set_patternType()`

此方法用于设置填充类型。

可选：`'none'`, `'solid'`, `'darkDown'`, `'darkGray'`, `'darkGrid'`, `'darkHorizontal'`, `'darkTrellis'`, `'darkUp'`, `'darkVertical'`, `'gray0625'`, `'gray125'`, `'lightDown'`, `'lightGray'`, `'lightGrid'`, `'lightHorizontal'`, `'lightTrellis'`, `'lightUp'`, `'lightVertical'`, `'mediumGray'`，`None`为恢复默认

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `name` | `str` | 必填 | `bool`表示开关，`None`恢复默认 |

**返回：**`None`

---

### 📌 `pattern_fill.set_fgColor()`

此方法用于设置前景色。

格式：`'FFFFFF'`, `'255,255,255'`, `(255, 255, 255)`, `Color`对象均可，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `color` | `str`<br>`tuple`<br>`Color` | 必填  | 颜色 |

**返回：**`None`

---

### 📌 `pattern_fill.set_bgColor()`

此方法用于设置背景色。

格式：`'FFFFFF'`, `'255,255,255'`, `(255, 255, 255)`, `Color`对象均可，`None`表示恢复默认。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `color` | `str`<br>`tuple`<br>`Color` | 必填 | 颜色 |

**返回：**`None`

---

## ✅ `gradient_fill`设置

此属性用于设置渐变填充方式。

与`pattern_fill`互斥，会清除已有`pattern_fill`设置。

### 📌 `gradient_fill.set_type()`

此方法用于设置填充类型。

可选：`'linear'`, `'path'`。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `name` | `str` | 必填 | 类型名称 |

**返回：**`None`

---

### 📌 `gradient_fill.set_degree()`

此方法用于设置程度。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `float` | 必填 | degree 数值 |

**返回：**`None`

---

### 📌 `gradient_fill.set_left()`

此方法用于设置左向数值。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `float` | 必填 | left 数值 |

**返回：**`None`

---

### 📌 `gradient_fill.set_right()`

此方法用于设置右向数值。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `float` | 必填 | right 数值 |

**返回：**`None`

---

### 📌 `gradient_fill.set_top()`

此方法用于设置上向数值。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `float` | 必填 | top 数值 |

**返回：**`None`

---

### 📌 `gradient_fill.set_bottom()`

此方法用于设置下向数值。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `value` | `float` | 必填 | bottom 数值 |

**返回：**`None`

---

### 📌 `gradient_fill.set_stop()`

此方法用于设置下向数值。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `values` | `list`<br>`tuple` | 必填 | stop 数值 |

**返回：**`None`

---

## ✅ `number_format`设置

此属性用于设置单元格数字格式。

### 📌 `number_format.set_format()`

此方法用于设置数字格式。

数字格式为特定格式的字符串，如`'m/d/yy h:mm'`，具体在`openpyxl.numbers`查看。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `string` | `str`<br>`None` | 必填 | 格式字符串，为`None`时恢复默认 |

**返回：**`None`

---

## ✅ `protection`设置

此属性用于设置单元格保护设置。

### 📌 `protection.set_hidden()`

此方法用于设置单元格是否隐藏。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool` | 必填  | `bool`表示开关 |

**返回：**`None`

---

### 📌 `protection.set_locked()`

此方法用于设置单元格是否锁定。

| 参数名称 | 类型 | 默认值 | 说明 |
|:----:|:--:|:---:|---------------------|
| `on_off` | `bool` | 必填  | `bool`表示开关 |

**返回：**`None`

---