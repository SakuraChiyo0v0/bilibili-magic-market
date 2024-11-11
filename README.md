# bilibili-magic-market

此项目旨在从 Bilibili 的魔力市场  页面抓取商品信息，并根据商品的价格更新历史数据。该脚本支持读取历史数据并进行商品状态验证，最终将更新后的商品信息保存到 Excel 文件中，便于用户查看。

市集网址:

https://mall.bilibili.com/neul-next/index.html?page=magic-market_index

## 项目功能

1. **从 Bilibili 获取商品信息**：通过 Bilibili 的 API 抓取商品的详细信息，包括商品名、价格、市场价、商品图片和链接等。
2. **历史数据校验**：支持读取历史数据（Excel 格式），并验证商品是否仍在售，若有价格变化或商品下架，则进行更新。
3. **数据保存到 Excel**：将抓取到的商品信息保存到 Excel 文件中，并优化显示效果，如折扣信息、超链接等。

## 环境要求

- Python 3.x
- `requests`：用于发送 HTTP 请求
- `pandas`：用于处理数据并保存到 Excel
- `openpyxl`：用于操作 Excel 文件
- `logging`：用于记录日志信息

## 安装依赖

在开始使用此项目之前，请确保您的环境中已经安装了相关依赖。

```
pip install requests pandas openpyxl
```

## 配置

### 请求头和 Cookie 配置

您需要根据自己的需求更新请求头信息，特别是 **cookie**，以确保程序能够成功从 Bilibili 获取数据。

### 数据文件路径

您可以根据需要修改以下文件路径：

```
data_file_path = "data.xlsx"  
# 用于存储商品的历史数据
show_file_path = "show.xlsx" 
# 用于存储优化后、便于查看的数据
```

## Item类


| key         | value              |
| ----------- | ------------------ |
| c2cItems_id | 交易ID             |
| goods_id    | 谷子ID             |
| name        | 商品名             |
| img         | 商品图片           |
| price       | 交易价格           |
| marketPrice | 市场价             |
| img         | 商品图片链接       |
| link        | 对应的商品页面链接 |

## 数据输出

程序会生成两个 Excel 文件：

1. **data.xlsx**：包含所有商品的详细信息，包括商品名、交易价格、市场价、商品图片和商品链接等。
2. **show.xlsx**：包含优化后的显示数据，去除了部分冗余列，并且加入了优惠价格和折扣信息。商品图片和商品链接将以超链接的形式显示，便于查看。

## 注意事项

1. **Cookie 配置**：请确保在 `headers` 中正确配置您的 Bilibili 用户 Cookie，否则请求可能会失败。

2. **Excel 文件被占用**：请确保在运行脚本时，Excel 文件没有被其他程序占用，以避免保存失败。

   

# 改进计划

## 1.可视化gui

由于Excel 本身不支持直接在单元格中显示网络图片 

也许会写个练手的gui 进行更好的数据展示 看热度如何了



# 已知BUG

### 1. 存在一个商品栏中有多个商品的情况 这种情况只能记录第一个商品

暂时懒得修复 代码层面只关心第一个商品 

![image-20241018043718641](C:\Users\MaFuY\Desktop\PythonProjects\bilibili-magic-market\README.assets\image-20241018043718641.png)

### ~~2.福袋不能被正常记录~~

![image-20241018061516407](C:\Users\MaFuY\Desktop\PythonProjects\bilibili-magic-market\README.assets\image-20241018061516407.png)

### ~~3.一些商品下架之后 不会主动更新~~

已完善

![image-20241018061815469](C:\Users\MaFuY\Desktop\PythonProjects\bilibili-magic-market\README.assets\image-20241018061815469.png)

### ~~4.检查记录进度会有错位~~

没啥影响 + 推荐不使用检查记录

![image-20241018134009548](C:\Users\MaFuY\Desktop\PythonProjects\bilibili-magic-market\README.assets\image-20241018134009548.png)