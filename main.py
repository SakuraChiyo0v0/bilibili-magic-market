import json
import time
import requests
import pandas as pd
from openpyxl import load_workbook
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
data_file_path = "data.xlsx"
show_file_path = "show.xlsx"
# 请求参数
nextId = None
payload = json.dumps({
    "categoryFilter": "2312",
    "priceFilters": ["5000-9000"],
    "discountFilters": ["10-100"],
    "nextId": nextId
})
headers = {
    'authority': 'mall.bilibili.com',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5,ja;q=0.4',
    'content-type': 'application/json',
    'cookie': "i-wanna-go-back=-1; rpdid=|(k|RYJRRkJY0J'uYmlJuYJY); buvid_fp_plain=undefined; buvid4=67D79775-A6FB-7B52-01B8-9990EC22C6FE70170-023092122-ANYbkYUbR8Wp0gWMOoykC9Z%2FUqM1oO3q9JKYCjXBBXic2OSqWZnFuw%3D%3D; LIVE_BUVID=AUTO1416959034209763; blackside_state=0; CURRENT_BLACKGAP=0; enable_web_push=DISABLE; header_theme_version=CLOSE; is-2022-channel=1; FEED_LIVE_VERSION=V8; b_ut=5; Hm_lvt_8d8d2f308d6e6dffaf586bd024670861=1719460143; DedeUserID=3493079625500884; DedeUserID__ckMd5=aa24e33ed38ba0a5; go-back-dyn=0; home_feed_column=5; buvid3=BED3B98D-1746-9E4D-347A-4A0F479A597912245infoc; b_nut=1727113212; _uuid=DFEEC4CE-D1D1-4922-778A-9510E3F9C4BB193571infoc; hit-dyn-v2=1; browser_resolution=1513-836; fingerprint=c12b3d6b27257de6693dd88783b1374f; PVID=1; CURRENT_QUALITY=80; CURRENT_FNVAL=4048; bili_ticket=eyJhbGciOiJIUzI1NiIsImtpZCI6InMwMyIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3MjkzNTYxNDcsImlhdCI6MTcyOTA5Njg4NywicGx0IjotMX0.ow6s-Kz8DpyWgbCgfUsitz5uy0-tHP_GshslDUqLQm4; bili_ticket_expires=1729356087; buvid_fp=c12b3d6b27257de6693dd88783b1374f; SESSDATA=35628653%2C1744690602%2C7433e%2Aa2CjBKFIQPTJFfHpYMdSUgdd_OFnuVo-66JT3n7QAOL0-YtBWbmcmmNjaywJG-JBPlNroSVktyS0RnT0Z6VUFwaUxjWEttRGtlTEFjMnFsTFBOUERJcFdmNzNITm45bEdtYktnRS1WdVFndE5RR29Qb3VMdWlTX2FXNkJBMTZNbjYyOEFvRXdsZllRIIEC; bili_jct=f0452754e4a37c241269594573349ddc; sid=4vo3gldi; bp_t_offset_3493079625500884=989311636068106240; b_lsid=77CBF87C_1929BF36334; kfcFrom=market_detail; from=market_detail; kfcSource=market_detail; msource=market_detail",
    'origin': 'https://mall.bilibili.com',
    'referer': 'https://mall.bilibili.com/neul-next/index.html?page=magic-market_index',
    'sec-ch-ua': '"Microsoft Edge";v="119", "Chromium";v="119", "Not?A_Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0'
}


class Item:
    def __init__(self, c2cItems_id, goods_id, name, price, marketPrice, img, link):
        self.c2cItems_id = c2cItems_id
        self.goods_id = goods_id
        self.name = name
        self.price = price
        self.marketPrice = marketPrice
        self.img = img
        self.link = link

    def __repr__(self):
        return (f"商品名 =『{self.name}』,"
                f"交易价格 = {self.price},市场价 = {self.marketPrice},商品图片 = {self.img},链接 = {self.link}")


def get_item_link(c2cItems_id):
    return (f"https://mall.bilibili.com/neul-next/index.html?"
            f"page=magic-market_detail&noTitleBar=1&itemsId={c2cItems_id}&from=market_index")


def is_item_valid(item, ):
    url = "https://mall.bilibili.com/mall-magic-c/internet/c2c/items/queryC2cItemsDetail?c2cItemsId=" + str(
        item.c2cItems_id)
    time.sleep(1)
    response = requests.get(url, headers=headers, data=payload)
    response.raise_for_status()  # 检查请求是否成功
    response_data = response.json()
    dropReason = response_data["data"]["dropReason"]
    saleStatus = response_data["data"]["saleStatus"]
    if dropReason is not None:
        logging.info(f"『{item.name}』已被下架 下架原因:{dropReason}")
        return False
    if saleStatus != 1:
        logging.info(f"『{item.name}』已被交易")
        return False
    return True


def data_processing(data, items_hash):
    for row in data:
        item = Item(
            c2cItems_id=row['c2cItemsId'],
            goods_id=row['detailDtoList'][0]['itemsId'],
            name=row['detailDtoList'][0]['name'],
            img="https:" + row['detailDtoList'][0]['img'],
            price=float(row['showPrice']),
            marketPrice=float(row['showMarketPrice']),
            link=get_item_link(row['c2cItemsId'])
        )
        if item.goods_id == 0:
            # 这是个盲盒 直接跳过
            continue
        # 尝试查询旧纪录
        old_item = items_hash.get(item.goods_id)
        # 如果没有记录 则添加
        if old_item is None:
            items_hash[item.goods_id] = item
            print(f"添加新商品: {item}")
            continue
        # 新item更便宜 更新对象
        if item.price < old_item.price:
            items_hash[item.goods_id] = item
            print(f"『{item.name}』发现更低价:{item.price} 旧价格:{old_item.price}")
            continue
        # 没有更便宜 但是旧纪录失效 无论如何也要替换
        if not is_item_valid(old_item):
            items_hash[item.goods_id] = item
            print(f"『{item.name}』旧纪录失效 重新获得商品:{item}")
            continue


def save_to_excel(items_hash):
    try:
        df = pd.DataFrame([
            {
                '交易ID': str(item.c2cItems_id),
                '谷子ID': item.goods_id,
                '商品名': item.name,
                '交易价格': float(item.price),
                '市场价': float(item.marketPrice),
                '商品图片': item.img,
                '链接': item.link
            }
            for item in items_hash.values()
        ])

        # 保存 DataFrame 到 Excel 文件
        df.to_excel(data_file_path, index=False)

        # 更好的显示 排除不必要信息 图片显示 链接点击
        df = df.drop(columns=['交易ID', '谷子ID'])
        df.insert(df.columns.get_loc('市场价') + 1, '优惠价格', df['市场价'] - df['交易价格'])
        # 修改 '链接' 列，将其转换为 Excel 超链接格式
        df['链接'] = df['链接'].apply(
            lambda x: f'=HYPERLINK("{x}", "点击打开")' if isinstance(x, str) and x.startswith('http') else x)
        df.to_excel(show_file_path, index=False)

        # 使用 openpyxl 调整列宽
        wb = load_workbook(show_file_path)
        ws = wb.active

        # 设置各列的宽度
        column_widths = {
            'A': 50,  # 商品名
            'B': 10,  # 交易价格
            'C': 10,  # 市场价
            'D': 10,  # 优惠价格
            'E': 80,  # 商品图片
            'F': 80,  # 链接
        }

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        # 保存工作簿
        wb.save(show_file_path)

        print(f"数据已保存到 {data_file_path}")
        print(f"您可以查看 {show_file_path} 以获取更好的体验")
    except PermissionError:
        logging.error(f"请检查您是否正在打开表格文件 这将导致无法生成表格")
    except Exception as e:
        logging.error(f"生成表格异常, 错误信息: {str(e)}")


def read_from_excel():
    try:
        df = pd.read_excel(data_file_path)
        items_hash = {}
        for _, row in df.iterrows():
            try:
                item = Item(
                    c2cItems_id=row['交易ID'],
                    goods_id=row['谷子ID'],
                    name=row['商品名'],
                    img=row['商品图片'],
                    price=float(row['交易价格']),
                    marketPrice=float(row['市场价']),
                    link=row['链接']
                )

                print(f"读取到商品: {item}")
                items_hash[item.goods_id] = item
            except Exception as e:
                logging.error(f"处理行 {row} 时出现异常, 错误信息: {str(e)}")
                continue  # 跳过当前行，继续处理下一行
        print(f"从 {data_file_path} 读取了 {len(items_hash)} 条数据")
        return items_hash
    except FileNotFoundError:
        logging.error(f"文件未找到: {data_file_path}")
        return {}
    except Exception as e:
        logging.error(f"读取 Excel 文件异常, 错误信息: {str(e)}")
        return {}


def main():
    url = "https://mall.bilibili.com/mall-magic-c/internet/c2c/v2/list"

    print("读取历史数据中...请耐心等待")
    items_hash = read_from_excel()
    print("读取历史数据完成,开始进行请求")
    print("=" * 50)
    try:
        while True:
            response = requests.post(url, headers=headers, data=payload)
            response.raise_for_status()  # 检查请求是否成功
            response_data = response.json()
            nextId = response_data["data"]["nextId"]
            if nextId is None:
                break
            data = response_data["data"]["data"]
            print("===============华丽的分割线😎华丽的分割线===============")
            data_processing(data, items_hash)
            time.sleep(3)
    except requests.RequestException as e:
        logging.error(f"网络请求异常, 错误信息: {str(e)}")
    except KeyboardInterrupt:
        logging.error("程序被用户中断")
    except SystemExit as e:
        logging.error(f"系统退出, 错误信息: {str(e)}")
    finally:
        print("=" * 50)
        print(f"总共收集到数据: {len(items_hash)}")
        save_to_excel(items_hash)


if __name__ == "__main__":
    main()
