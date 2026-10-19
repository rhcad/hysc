# hysc-build

《华严经疏钞》附《会玄记》文档合编，基于Python的生成工具。

## 制作说明

- 安装 LibreOffice （Windows上需要加到 PATH，Mac/Linux不需要特殊设置）
- 安装 Python3，安装依赖库 `pip3 install -r requirements.txt`
- 运行脚本生成docx文档、`pu-used.csv`：
  - 如已创建虚拟环境，可先激活: `source ./venv/bin/activate`
  - 生成校对文档(`合编.docx`)：`python3 hysc-build.py`
  - 生成发布文档(`合编-.docx`)：`python3 hysc-build.py release=1`

## 结果文档

- 校对文档: [合编.docx](合编.docx)
- 合并统计: [pu-used.csv](pu-used.csv)
- 内容改动: [pu-change.csv](pu-change.csv)

## 匹配校对

根据生成的校对文档、匹配表 `pu-used.csv`（可用正则式 `未合|无匹配|跳变|□` 定位待修正的段落），
在 `hui/X0236_xxx.html` 中找到待修正的段落元素，对其 p 元素设置 `force` 匹配属性，然后再次生成。

- 强制匹配 `<p force="会记词,疏钞文本"><span...></span>疏：`
  - `force` 文本两个一组，逗号分隔
    - 第一个文本是X0236段的引用词（在校对文档中显示为下划线，待修正的为斜体）
    - 第二个文本是docx中的疏钞文本，需要标点一致，可根据段号定位

## 资料来源

### 疏钞

- 从 [新修華嚴經疏鈔](https://hysc.dila.edu.tw/interface.php) 下载各卷网页
  - 修正个别节点结构的结果见 `shu` 文件夹
  - `hysc.json` 是用 `python3 hysc-build.py kp_build=1` 生成的科判数据

### 会玄记

- 从 CBETA [X0236](https://cbetaonline.dila.edu.tw/zh-tw/X0236) 會玄記 下载(匯出)：HTML格式、范围40卷
- 将下载的 X0236 文件夹移到本项目下
  - `hui` 文件夹是已校正匹配的会玄记网页，改动见 `pu-change.csv`
  - `pu-origin.csv` 是用 `python3 hysc-build.py gen_origin=1` 生成的原始文本数据
  - 如需比对原始文本与`hui`文件夹，可运行 `python3 hysc-build.py diff_hui=1`
