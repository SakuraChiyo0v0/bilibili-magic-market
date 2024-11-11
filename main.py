import json
import time
import requests
import pandas as pd
import logging

from openpyxl import load_workbook


# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
data_file_path = "data.xlsx"
show_file_path = "show.xlsx"
# 请求参数
nextId = None

# 价格区间(单位分) 根据自己的需求进行更改:
# "0-2000", "3000-5000", "20000-0", "5000-10000", "2000-3000", "10000-20000", "20000-0"
payload = json.dumps({
    "categoryFilter": "2312",
    "priceFilters": ["0-2000", "3000-5000", "20000-0", "5000-10000", "2000-3000", "10000-20000", "20000-0"],
    "discountFilters": [],
    "nextId": nextId
})
headers = {
    'authority': 'mall.bilibili.com',
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6,zh-TW;q=0.5,ja;q=0.4',
    'content-type': 'application/json',
    'cookie': "i-wanna-go-back=-1; buvid_fp_plain=undefined; buvid4=67D79775-A6FB-7B52-01B8-9990EC22C6FE70170-023092122-ANYbkYUbR8Wp0gWMOoykC9Z%2FUqM1oO3q9JKYCjXBBXic2OSqWZnFuw%3D%3D; CURRENT_BLACKGAP=0; enable_web_push=DISABLE; header_theme_version=CLOSE; is-2022-channel=1; FEED_LIVE_VERSION=V8; b_ut=5; DedeUserID=3493079625500884; DedeUserID__ckMd5=aa24e33ed38ba0a5; go-back-dyn=0; buvid3=BED3B98D-1746-9E4D-347A-4A0F479A597912245infoc; b_nut=1727113212; _uuid=DFEEC4CE-D1D1-4922-778A-9510E3F9C4BB193571infoc; hit-dyn-v2=1; home_feed_column=4; rpdid=0zbfvUllao|1dKAa2Fy|1y0|3w1T4yGa; LIVE_BUVID=AUTO7017305218250394; PVID=1; CURRENT_FNVAL=16; Hm_lvt_8d8d2f308d6e6dffaf586bd024670861=1729248258,1730821757; fingerprint=5f376ec98e8600e9ab655a2f38566d8b; buvid_fp=5f376ec98e8600e9ab655a2f38566d8b; bili_ticket=eyJhbGciOiJIUzI1NiIsImtpZCI6InMwMyIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3MzE0OTMwNDAsImlhdCI6MTczMTIzMzc4MCwicGx0IjotMX0.abiiz6z2Mm5-eGrguhBcy74kn8rT3M63gpvDCCrdqhs; bili_ticket_expires=1731492980; SESSDATA=4008cc36%2C1746788446%2C3d48f%2Ab2CjAPNkXx5juMB4DK1szfTFIzViuKIRitCUy3OezUKU79rfYh_1HDQYuJ02aJ6Z5ejRMSVktjY09QNFhVanY2Q0JFX2FBMkpzWDdYLXJIMzdBYUhLNHNTQWFXSTlfazNHMk42bDZwUTZRVmhOX3FiNXZ2Nzh2S2QtUTRaUDNpYkpXYjFoT29hSlJ3IIEC; bili_jct=968f14bae2c98d2eaadc13bf3ac79620; sid=7fw91pa9; CURRENT_QUALITY=120; bp_t_offset_3493079625500884=998391038018060288; b_lsid=10510C62BA_193197D8962; browser_resolution=638-665; bsource=search_bing",
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


def is_skip_check():
    while True:
        skip_check = input("是否开启历史数据校验 (y/n): %注意 由于控制请求速度原因 开启后将需要长时间加载 建议输入n%\n").strip().lower()
        if skip_check in ['y', 'n']:
            if skip_check == 'y':
                return False
            else:
                return True
        else:
            print("输入无效，请输入 y 或 n。")


def is_item_valid(item):
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
        # 创建 DataFrame
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

        # 保存原始数据
        df.to_excel(data_file_path, index=False)

        # 优化显示，删除部分列，增加优惠信息
        df = df.drop(columns=['交易ID', '谷子ID'])
        df.insert(df.columns.get_loc('市场价') + 1, '优惠价格',
                  (df['市场价'] - df['交易价格']).apply(lambda x: f"{x:.2f}元"))
        df.insert(df.columns.get_loc('市场价') + 2, '折数',
                  (round(df['交易价格'] / df['市场价'], 2) * 10).apply(lambda x: f"{x:.1f}折"))

        # 转换为带“元”的字符串
        df['交易价格'] = df['交易价格'].apply(lambda x: f"{x:.2f}元")
        df['市场价'] = df['市场价'].apply(lambda x: f"{x:.2f}元")

        # 将图片的 URL 作为超链接，显示为“点击查看图片”
        df['商品图片'] = df['商品图片'].apply(
            lambda x: f'=HYPERLINK("{x}", "点击查看图片")' if isinstance(x, str) and x.startswith('http') else x
        )

        # 超链接格式
        df['链接'] = df['链接'].apply(
            lambda x: f'=HYPERLINK("{x}", "点击打开")' if isinstance(x, str) and x.startswith('http') else x
        )

        df.to_excel(show_file_path, index=False)

        # 调整 Excel 的列宽
        wb = load_workbook(show_file_path)
        ws = wb.active
        column_widths = {
            'A': 50,  # 商品名
            'B': 10,  # 交易价格
            'C': 10,  # 市场价
            'D': 10,  # 折数
            'E': 10,  # 优惠价格
            'F': 40,  # 商品图片（超链接）
            'G': 40  # 链接
        }
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        wb.save(show_file_path)

        print(f"数据已保存到 {data_file_path}")
        print(f"您可以查看 {show_file_path} 以获取更好的体验")

    except PermissionError:
        logging.error(f"请检查您是否正在打开表格文件 这将导致无法生成表格")
    except Exception as e:
        logging.error(f"生成表格异常, 错误信息: {str(e)}")


def read_from_excel(skip_check):
    try:
        df = pd.read_excel(data_file_path)
        # 获取 DataFrame 的总数据量
        total_rows, total_columns = df.shape
        index = 1
        print(f"总行数: {total_rows}")
        items_hash = {}
        for _, row in df.iterrows():
            try:
                print(f"({index}/{total_rows})", end=" ")
                index += 1
                item = Item(
                    c2cItems_id=row['交易ID'],
                    goods_id=row['谷子ID'],
                    name=row['商品名'],
                    img=row['商品图片'],
                    price=float(row['交易价格']),
                    marketPrice=float(row['市场价']),
                    link=row['链接']
                )
                if skip_check or is_item_valid(item):
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

    skip_check = is_skip_check()
    print("读取历史数据中...请耐心等待")
    items_hash = read_from_excel(skip_check)
    print("读取历史数据完成,开始进行请求")
    print("=" * 50)
    response_data = None
    try:
        while True:
            try:
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
            except KeyError as e:
                logging.error(f"KeyError: {e}. Response data: {response_data}")
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
