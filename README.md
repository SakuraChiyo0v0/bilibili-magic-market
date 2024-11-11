# 使用教程

市集网址:

https://mall.bilibili.com/neul-next/index.html?page=magic-market_index





跳转具体商品页面

https://mall.bilibili.com/neul-next/index.html?page=magic-market_detail&noTitleBar=1&itemsId=106354806628&from=market_index

## 返回数据

```json
{
    "c2cItemsId": 106010548233,
    "type": 1,
    "c2cItemsName": "THREEZERO 帕瓦 成品模型",
    "detailDtoList": [
        {
            "blindBoxId": 188807284,
            "itemsId": 10137476,
            "skuId": 1000253914,
            "name": "THREEZERO 帕瓦 成品模型",
            "img": "//i0.hdslb.com/bfs/mall/mall/d4/f0/d4f088984e1c54658e8d956a57996788.png",
            "marketPrice": 103000,
            "type": 0,
            "isHidden": false
        }
    ],
    "totalItemsCount": 1,
    "price": 94960,
    "showPrice": "949.6",
    "showMarketPrice": "1030",
    "uid": "34***2",
    "paymentTime": 0,
    "isMyPublish": false,
    "uspaceJumpUrl": null,
    "uface": "https://i0.hdslb.com/bfs/face/member/noface.jpg",
    "uname": "b***"
}
```

处理思路: 一个商品可能有多个在商品上上架 




| key         | value    |
| ----------- | -------- |
| c2cItems_id | 交易ID   |
| goods_id    | 谷子ID   |
| name        | 商品名   |
| img         | 商品图片 |
| price       | 交易价格 |
| marketPrice | 市场价   |
|             |          |
|             |          |
|             |          |
|             |          |
|             |          |

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