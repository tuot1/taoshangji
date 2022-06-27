# coding=utf-8
import requests
import json
import pandas as pd
import time
import os.path
import re
from config import config


class Taoshangji():
    def __init__(self, keyWord):
        """
        :keyword  关键词
        """
        self.data_save_dir = os.path.join(os.getcwd(),
                                          "商机数据.xlsx") 
        self.keyWord = keyWord
        self.marketId = None
        self.headers = {
            'authority': 'taoshangji.taobao.com',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
            'cache-control': 'max-age=0',
            'cookie': config["cookie"],
            'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="102", "Google Chrome";v="102"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'document',
            'sec-fetch-mode': 'navigate',
            'sec-fetch-site': 'none',
            'sec-fetch-user': '?1',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36',
        }      
        self.params = {
            '_csrf': config["csrf"],
            'keyWord': self.keyWord,
            'pageSize': config["PAGESIZE"],  # 因为这里可以直接且最多获取60个数据，所以直接写60
        }

        self.marketName_list = list()
        self.queryUv_list = list()
        self.itmCnt_list = list()
        self.competition_list = list()
        self.tsjIndex_list = list()
        self.marketScale_list = list()
        self.marketRecLabel_list = list()
        self.marketId_list = list()

        self.searchWordSeq_list = list()
        self.otherWordSeq_list = list()

    def file_is_exists(self):
        '''
         判断  商机数据.xlsx是否被创建，如果未创建后续就写不了数据。所以不浪费请求次数
         判断  商机数据.xlsx是否被打开，如果被打开后续页写不了数据，同上
        '''
        # 判断文件是否被创建
        if not os.path.exists(self.data_save_dir):
            ex = Exception("当前目录下不存在文件,请先在当前文件夹下鼠标右键-->新建--->商机数据.xlsx")
            raise ex
        # 判断文件是否被打开
        excelname = os.path.split(self.data_save_dir)[
            -1]  # 获取文件名
        hidefilename = os.getcwd() + r"/~$" + excelname
        # print(hidefilename)
        if os.path.exists(hidefilename):
            ex = Exception("当前商机数据.xlsx已经被打开，请关闭后再试\n 别担心，已经为您节省请求次数了哦")
            raise ex

    def get_other_data(self):
        """
        获取关键词下的其他数据
        :return: "searchByrs": "极高",  // 消费者搜索人气(同行业对比)
                 "itmCnt": "&lt;1万",  // 近30天成交商品量
                 "clickByrs": "极高",  // 消费者点击人气(同行业对比)
                  "clkRate": "较高",  // 引导商品点击率(同行业对比)
                "searchWordSeq"      // 消费者同时还会搜索
                "otherWordSeq"      // 消费者还会搜索哪些
        """
        response_other_data = requests.get('https://taoshangji.taobao.com/tsj/market/search/summary',
                                           params=self.params,
                                           headers=self.headers)
        other_ret = response_other_data.content.decode()
        try:
            other_ret_dict = json.loads(other_ret)  # 很有可能因为cookie值没有更新，导致返回的是错误的数据。而这数据刚好无法被loads处理，所以这里进行了try尝试
        except:
            print("请检查cookie，当前cookie需要更新")

        return other_ret_dict

    def get_market_data(self):
        """
        搜索关键词下的有关市场。不包括更多推荐的有关市场
        :return: 可多达60个商品数据
        """
        response_market = requests.get('https://taoshangji.taobao.com/tsj/market/search', params=self.params,
                                       headers=self.headers)
        market_ret = response_market.content.decode()
        # 这里因为不管搜索什么都有对应的数据，且暂时没发现报错。所以这里不进行异常捕获了
        market_ret_dict = json.loads(market_ret)
        return market_ret_dict

    def get_detail_data(self):
        """
        通过get_market_data请求得到marketId。有了该值就可以请求更详细的数据
        :return:
        """
        self.params["marketId"] = self.marketId
        response_detail = requests.get('https://taoshangji.taobao.com/tsj/market/card/detail', params=self.params,
                                       headers=self.headers)
        detail_ret = response_detail.content.decode()
        detail_ret_dict = json.loads(detail_ret)

        return detail_ret_dict

    def handle_data(self, get_market_data, get_other_data):
        """
        1. 从get请求中得到get_market_data的数据，进行数据提取。最后用列表保存，方便后续进行.xlsx数据存储
        另添加列索引：竞争度。 处理逻辑是根据公式降序排名
        公式：需求热度/在线商品量= 竞争度 #竞争度越高说明越好
        2. 从get请求中得到get_other_data的数据，
        :param get_response_data 关键词指数信息等
        :param get_other_data 关键词其他信息，如同时搜索的关键词等
        :return: 从get请求中得到get_other_data数据，经过提炼后返回
        """
        # 1. 60个商品数据
        market_data_list = get_market_data.get("data").get("seMarketCards")
        if market_data_list == None:
            print("get_market_data数据为空")
            return
        for data in market_data_list:
            """
                # "marketId"   关键词的ID，后续数据需求分析时需要用到该ID
                "marketName"  商品名字
                "queryUv"  需求热度   
                "itmCnt"   在线商品量
                "tsjIndex"  商机分
                "marketScale"  规模级别
                "marketRecLabel"  近30天需求量98%;近7天成交量44%
            """
            marketName = data["marketName"]
            queryUv = data["queryUv"]
            itmCnt = data["itmCnt"]
            competition = queryUv / itmCnt
            tsJIndex = data["tsjIndex"]
            marketScale = data["marketScale"]
            marketRecLabel = data["marketRecLabel"]
            self.marketName_list.append(marketName)
            self.queryUv_list.append(queryUv)
            self.itmCnt_list.append(itmCnt)
            self.competition_list.append(competition)
            self.tsjIndex_list.append(tsJIndex)
            self.marketScale_list.append(marketScale)
            self.marketRecLabel_list.append(marketRecLabel)
            # print(marketName,queryUv,itmCnt,competition,tsJIndex,marketScale,marketRecLabel):
            # 第一种方式遍历匹配找到关键词的ID，感觉不优雅
            # if self.marketNmae == self.keyWord:
            # self.marketId = data["marketId"]
            # print("找到了%s对应的marketId是：%s" % (data["marketName"], data["marketId"]))
            # else:
            #     print("没有找到%s对应的marketId" % self.keyWord)

            # 第二种方式 从感觉上就更优雅，虽然本质和上面差不多~
            marketId = data["marketId"]
            self.marketId_list.append(marketId)  # 用这种方式初始化时就要创建一下空列表
        if self.marketName_list.count(self.keyWord):
            index = self.marketName_list.index(self.keyWord)  # 如果有对应的值就获取下标
            self.marketId = self.marketId_list[index]  # 通过下标取ID
            print("找到了-->%s对应的marketId是：%s，可获取更详细数据" % (self.keyWord, self.marketId))
        else:
            print("没有找到-->%s对应的marketId" % self.keyWord)

        # 2. 搜索**关键词下，消费者同时还会搜索、消费者还会搜索哪些
        o = get_other_data.get("data")  # get取值，如果没有对应的Key，会返回None而不是报错。更好提高了程序的健壮性
        searchByrs = o.get("searchByrs")  # 消费者搜索人气(同行业对比)
        clickByrs = o.get("clickByrs")  # 消费者点击人气(同行业对比)
        clkRate = o.get("clkRate")  # 引导商品点击率(同行业对比)
        itmCnt = o.get("itmCnt")  # 近30天成交商品量
        searchWordSeq_list = o.get("searchWordSeq").split(
            ",")  # 消费者同时还会搜索  # 如果o. get("searchWordSeq")本身会返回一个空列表，此时split如果没有数据又会返回一个列表 相当于两个列表嵌套[[]]
        otherWordSeq_list = o.get("otherWordSeq").split(
            ",")  # 消费者还会搜索哪些  # 同上，如果出现双列表嵌套，此时就不能简单的用None判断了。因为[[]]里面这个列表长度会算+1
        if len(searchWordSeq_list) > 1:  # 如果有数据就进行遍历，然后正则匹配。因为正则匹配如果数据为空group会报错。所以这里直接加入if判断
            for searchWordSeq in searchWordSeq_list:
                ret = re.search(r"&quot;(.*?):\d&quot;", searchWordSeq)  # searchWordSeq = &quot;裙子:7&quot;
                self.searchWordSeq_list.append(ret.group(1))
        if len(otherWordSeq_list) > 1:  # 同上
            for otherWordSeq in otherWordSeq_list:
                ret = re.search(r"&quot;(.*?):\d&quot;", otherWordSeq)
                self.otherWordSeq_list.append(ret.group(1))
        ret_other_data = {
            "消费者搜索人气": [searchByrs],
            "消费者点击人气": [clickByrs],
            "引导商品点击率": [clkRate],
            "近30天成交商品量": [itmCnt],
            "消费者同时搜的词": self.searchWordSeq_list,
            "消费者还会搜的词": self.otherWordSeq_list
        }
        return ret_other_data

    def handle_detail_data(self, get_detail_data):
        """
        :param get_detail_data:  关键词更详细的数据
        :return:
        """
        d = get_detail_data["data"]
        ret_detail_1_data = self.hangdle_detail_1_data(d)  # 获取第一块数据

        ret_detail_2_data, ret_detail_2a_data = self.hangdle_detail_2_data(d)  # 获取第二块数据

        ret_detail_3_data = self.hangdle_detail_3_data(d)
        return ret_detail_1_data, ret_detail_2_data, ret_detail_2a_data, ret_detail_3_data

    def hangdle_detail_1_data(self, d):
        # 3.1 关键词的详细需求数据
        d_markName = d["marketName"]  # 关键词
        d_queryUv = d["queryUv"]  # 需求热度
        d_itmCnt = d["itmCnt"]  # 在线商品量
        d_marketScale = d["marketScale"]  # 成交规模
        d_competition = d_queryUv / d_itmCnt  # 竞争度
        d_tsjIndex = d["tsjIndex"]  # 商机分
        d_marketRecLabel = d["marketRecLabel"]  # 近日成交指数
        d_marketDesc = d["marketDesc"]  # 市场需求状态
        d_byrDecisionRec = d["byrDecisionRec"]  # 消费者决策因素  "风格:冷淡风,轻熟"
        d_slrCntScale = d["slrCntScale"]  # 在线商家量

        d_payAmt30dRate = d["payAmt30dRate"]  # 成交规模下百分百指数
        d_queryUvRate = d["queryUvRate"]  # 需求热度百分百指数
        d_itmCnt30dRate = d["itmCnt30dRate"]  # 在线商品量百分百指数
        d_slrCnt30dRate = d["slrCnt30dRate"]  # 在线商家量百分百指数

        # 构建detail_dataframe数据
        # 第一种 多个字典的方式构建，会导致排序很乱
        # ret_detail_data = [
        #     {"搜索词": d_markName},
        #     {"需求热度": d_queryUv},
        #     {"在线商品量": d_itmCnt},
        #     {"成交规模": d_marketScale},
        #     {"竞争度": d_competition},
        #     {"消费者决策因素": d_byrDecisionRec},
        #     {"商机分": d_tsjIndex},
        #     {"近日成交指数": d_marketRecLabel},
        #     {"市场需求状态": d_marketDesc},
        #     {"在线商家量": d_slrCntScale},
        #
        #     {"成交规模": "近30天上涨%.2f" % d_payAmt30dRate + "%"},
        #     {"需求热度": "近30天上涨%.2f" % d_queryUv + "%"},
        #     {"在线商品量": "近30天上涨%.2f" % d_itmCnt + "%"},
        #     {"在线商家量": "近30天上涨%s" % d_slrCntScale + "%"}  #
        #
        # ]
        # 第二种 和上面一样这个也是多个字典的方式构建，但这里只有两个字典
        # 因为根据分析排在第一行和第二行的数据可以分别由两个字典构建完成。这种方式适合刚好数据少且一两行就可以写下
        # 更优化的使用该方式的方法是：第一个字典写下所有的key。这是重点，因为第一个字典到时后都会转换为第一行。
        # 然后第二个字典或者第三个字典等就重复写下去，这样排序不会出现问题的
        # 如果数据多行也不要慌，通过from_dict() transpose()也很好解决的~ 这里偷懒啦~~~~~~~
        ret_detail_1_data = [
            {
                "搜索词": d_markName, "需求热度": d_queryUv, "在线商品量": d_itmCnt, "成交规模": d_marketScale, "竞争度": d_competition,
                "消费者决策因素": d_byrDecisionRec, "商机分": d_tsjIndex, "近日成交指数": d_marketRecLabel, "市场需求状态": d_marketDesc,
                "在线商家量": d_slrCntScale,
            },
            {
                "成交规模": "近30天%.2f" % d_payAmt30dRate, "需求热度": "近30天%.2f" % d_queryUvRate,
                "在线商品量": "近30天%.2f" % d_itmCnt30dRate, "在线商家量": "近30天%s" % d_slrCnt30dRate
            }

        ]
        return ret_detail_1_data

    def hangdle_detail_2_data(self, d):
        d_wdjListStr_list = d.get("wdjListStr").split(",")  # 消费者都在问
        byrDecision_str = d.get("byrDecision")  # 消费者决策因素
        re_obj = re.compile(r"\d:(.*?):(.*?);")
        byrDecisio_ret = re_obj.findall(byrDecision_str)  # [('风格', '冷淡风,轻熟,潮牌,韩版,休闲,日系,韩式'),......]

        ret_detail_2_data = {
            "消费者都在问": d_wdjListStr_list,
            "消费者决策因素": []  # 占位和‘消费者都在问’排在一起
        }

        ret_detail_2a_data = {element[0]: element[1].split(",") for element in byrDecisio_ret}  # 字典推导式+split
        return ret_detail_2_data, ret_detail_2a_data

    def hangdle_detail_3_data(self, d):
        '''
        市场售价分布 商品量(占比)
        "itmCntPriceZone": "0-60.0:0.169;60.0-120.0:0.314;120.0-180.0:0.137;180.0-240.0:0.092;240.0以上:0.289",
        销售件数(占比)
        "saleCntPriceZone": "0-60.0:0.183;60.0-120.0:0.466;120.0-180.0:0.2;180.0-240.0:0.023;240.0以上:0.129",
        '''
        d_itmCntPriceZone_list = d.get("itmCntPriceZone").split(";")  # 商品量(占比)
        d_saleCntPriceZone_list = d.get("saleCntPriceZone").split(";")  # 销售件数(占比)
        itmCntPriceZone_price_list = list()  # 价格区间
        itmCntPriceZone_goods_list = list()  # 商品量占比
        # Sales_rate_list = list()  #
        # 商品量占比
        for itm in d_itmCntPriceZone_list:
            itmCntPriceZone_price_list.append(itm.split(":")[0])
            itmCntPriceZone_goods_list.append("%3.1f%%" % round(float(itm.split(":")[1])*100, 2))  # 再不确定字符串int还是float的时候就不要使用int 否则遇到fload类型就会报错ValueError: invalid literal for int() with base 10
        ret_detail_3_data = {
            "价格区间": itmCntPriceZone_price_list,
            "商品量占比": itmCntPriceZone_goods_list,
            # "销售件数占比": [f'{float(sale.split(":")[1])*100}%' for sale in d_saleCntPriceZone_list]
            # "销售件数占比": [str(float(sale.split(":")[1])*100)+'%' for sale in d_saleCntPriceZone_list]  #
            "销售件数占比": ["%3.1f%%" % round(float(sale.split(":")[1]) * 100, 2) for sale in d_saleCntPriceZone_list]  # round四舍五入

        }

        return ret_detail_3_data

    def get_conversion_rate(self, a,b):
        '''
        公式: 销售件数/商品量 = 该商品的转化指数
        :param a: 商品量占比
        :param b: 销售件数占比
        :return: conversion_rate
        '''
        
        # ['17.0%', '31.3%', '13.5%', '9.3%', '28.9%']
        a = float(a.strip('%'))
        b = float(b.strip('%'))
        return round(b / a, 2)  # round 包裹这里的，前面加到上面去了

    def save_xlsx(self, ret_other_data=None, ret_detail_1_data=None, ret_detail_2_data=None, ret_detail_2a_data=None, ret_detail_3_data=None):
        """
        :return:
        """
        market_data = {
            "关键词": self.marketName_list,
            "需求热度": self.queryUv_list,
            "在线商品量": self.itmCnt_list,
            "成交规模": self.marketScale_list,
            "竞争度": self.competition_list,
            "商机分": self.tsjIndex_list,
            "近日成交指数": self.marketRecLabel_list,
        }
        df_market_data = pd.DataFrame(market_data)

        df_other_data = None
        # print("shuzu length",ret_other_data)
        if not ret_other_data == None:
            # df_other_data = pd.DataFrame(ret_other_data)  # 当列表长度不统一时报错
            df_other_data = pd.DataFrame.from_dict(ret_other_data, orient='index')
            df_other_data = df_other_data.transpose()
        # 详情数据1模块  市场基本数据
        df_detail_1_data = None
        if not ret_detail_1_data == None:
            df_detail_1_data = pd.DataFrame(ret_detail_1_data)
            # df_detail_data = df_detail_data_nan.dropna(axis=0, how='any', inplace=False)

            # 详情数据2模块 消费者都在问和消费者决策因素
            df_detail_2_data = pd.DataFrame.from_dict(ret_detail_2_data, orient='index')
            df_detail_2_data = df_detail_2_data.transpose()
            df_detail_2a_data = pd.DataFrame.from_dict(ret_detail_2a_data, orient='index')
            df_detail_2a_data = df_detail_2a_data.transpose()
            # 详情数据3模块 销售占比
            df_detail_3_data = pd.DataFrame(ret_detail_3_data)
            # 通过调用匿名函数重新转换为小数点相除击败：typeError: unsupported operand type(s) for /: 'str' and 'str'
            df_detail_3_data["转化指数"] = df_detail_3_data.apply(lambda row: self.get_conversion_rate(row['商品量占比'], row['销售件数占比']), axis=1)




        now_data = time.strftime('%m-%d', time.localtime())
        sheet_name_data = self.keyWord + now_data
        try:
            df_data = pd.read_excel(self.data_save_dir)
            if df_data.shape[0] == 0:  # 说明是一个空的xlsx
                with pd.ExcelWriter(self.data_save_dir) as writer:
                    df_market_data.to_excel(writer, sheet_name=sheet_name_data, index=False, header=True)
                    if df_other_data is not None:
                        df_other_data.to_excel(writer, sheet_name=sheet_name_data, startrow=df_market_data[0] + 3,
                                               index=False, header=True)
                    if df_detail_1_data is not None:
                        df_detail_1_data.to_excel(writer, sheet_name=sheet_name_data,
                                                  startrow=df_market_data.shape[0] + df_other_data.shape[0] + 6,
                                                  index=False, header=True)
                        df_detail_2_data.to_excel(writer, sheet_name=sheet_name_data,
                                                  startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                           df_detail_1_data.shape[0] + 9,
                                                  index=False, header=True)
                        df_detail_2a_data.to_excel(writer, sheet_name=sheet_name_data,
                                                   startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                            df_detail_1_data.shape[0] + 10, startcol=1,
                                                   index=False, header=True)

                        df_detail_3_data.to_excel(writer, sheet_name=sheet_name_data,
                                                   startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                            df_detail_1_data.shape[0] + df_detail_2_data.shape[0] + 15,
                                                   index=False, header=True)

                    print("------------------------成功保存数据")
            else:
                # with pd.ExcelWriter(self.data_save_dir, mode='a') as writer:  # 因为with 报错了也会接着执行, 执行逻辑不对
                #     df.to_excel(writer, sheet_name=sheet_name_data, index=False, header=True)
                writer = pd.ExcelWriter(self.data_save_dir, mode='a')
                df_market_data.to_excel(writer, sheet_name=sheet_name_data, index=False, header=True)
                print("--->有关市场数据加载完毕")
                if df_other_data is not None:  # df数据类型不能简单的使用 if 判断真假。因为它不知道我们到底要它干嘛。但用empty也不行，因为我在上面设置为None，None里面没有empty这个方法
                    df_other_data.to_excel(writer, sheet_name=sheet_name_data, startrow=df_market_data.shape[0] + 3,
                                           index=False, header=True)
                    print("--->其他数据加载完毕")
                if df_detail_1_data is not None:
                    df_detail_1_data.to_excel(writer, sheet_name=sheet_name_data,
                                              startrow=df_market_data.shape[0] + df_other_data.shape[0] + 6,
                                              index=False,
                                              header=True)
                    df_detail_2_data.to_excel(writer, sheet_name=sheet_name_data,
                                              startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                       df_detail_1_data.shape[0] + 9,
                                              index=False, header=True)
                    df_detail_2a_data.to_excel(writer, sheet_name=sheet_name_data,
                                               startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                        df_detail_1_data.shape[0] + 10, startcol=3,
                                               index=False, header=True)

                    df_detail_3_data.to_excel(writer, sheet_name=sheet_name_data,
                                                   startrow=df_market_data.shape[0] + df_other_data.shape[0] +
                                                            df_detail_1_data.shape[0] + df_detail_2_data.shape[0] + 12,
                                                   index=False, header=True)
                    print("--->商品详情模块1-2-3数据加载完毕")
                writer.save()
                print("\n===============成功追加以上数据===============")
        except PermissionError:
            ex = Exception("占用冲突，请先关闭正在使用的文件-->商机数据.xlsx")
            raise ex

    def run(self):
        self.file_is_exists()  # 判断  商机数据.xlsx是否被创建和是否被占用，如果未创建或者已被占用那么后续都写不了数据。所以不浪费请求次数

        # get请求url--->获取该关键词下消费者的搜索人气、同时搜索的关键词、还会搜索的关键词...
        get_other_data = self.get_other_data()

        # get请求url可获取多达60个商品数据
        get_market_data = self.get_market_data()
        ret_other_data = self.handle_data(get_market_data, get_other_data)

        # get请求url 获取关键词详细的需求分析.
        # 如果没有关键词的ID，也就没有详情。但有时候关键词的ID可能不在上一个响应中的 #TODO 需要改进的地方
        if not self.marketId == None:
            get_detail_data = self.get_detail_data()
            ret_detail_1_data, ret_detail_2_data, ret_detail_2a_data, ret_detail_3_data = self.handle_detail_data(get_detail_data)
            # 保存数据
            self.save_xlsx(ret_other_data, ret_detail_1_data, ret_detail_2_data, ret_detail_2a_data, ret_detail_3_data)
        else:
            self.save_xlsx(ret_other_data)


if __name__ == '__main__':
    print("-" * 30)
    print("仅初次使用需要阅读如下说明")
    print("在当前文件夹下 鼠标右键-->新建-->Microsoft Office Excel 2007 工作表")
    print("名字修改为---->商机数据.xlsx")
    print("-" * 30)
    print("*" * 30)
    print("功能:\n根据输入的关键字在庞大数据池中\n‘精准匹配’有关的商品市场\n   可一次提供多达60条商品数据")
    print("*" * 30)
    keyWord = input("请输入要查询的关键词:")
    taoshangji = Taoshangji(keyWord)
    taoshangji.run()
    print("************数据更新成功，可以打开文件查看了***************")
