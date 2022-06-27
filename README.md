# 背景-
即使是资深的运营人员在面对海量的商品并且需求变化速度超前的情况下，往往也难以确定上新选品方向，更何况个人店主、宝妈等群体更是被上新问题困扰。
在这一系列情况下淘商机应势诞生，

淘商机官方网站：

- taoshangji.taobao.com

官方对淘商机的介绍：

- "在基于全网消费者需求及货品供应情况的分析，可抢先一步找到选品运营新赛道"

我用了一段时间，感觉还可以将其美化输出一下。

- 于是，我决定自己动手，**加入新的维度数据**，帮助以下人员更简洁高效的选品
- 上新人员、个人店主、宝妈等群体--->特指我姐
- ps：实体店主、pdd店主等不管线上线下平台，感兴趣人员都可以将该数据作为参考
# 功能及数据维度介绍

- 关键词数据，可多达60条
- 消费者同时搜索的词 (比如消费者搜索'婚礼衣架'同时还会添加的关键词有哪些
- 消费者还会搜的词  (比如当消费者搜索'婚礼衣架'后接着会搜索的词
- 消费者关心的问题  (应该是根据 大家问、评论 最后词频统计得出
- 消费者决策因素  (比如'婚礼衣架'， 那么它的决策因素和产品的风格、材质、人群有很大关系
-   价格区间下各商品量占比，销售件数占比
-   商家占比
-   商品占比
-   .......更多数据维度内容自己看哈
-   保存为xlsx文件，方便自定义查询，修改
-   当下次运行时数据会追加写入新的sheet，方便操作

# 作者
- Author: Alfie
- Date: 2022.6.27 

# 更新
- 加入数据维度：竞争指数，一定程度上表明 竞争指数同关键词转化效果成正比
-  加入数据维度：转化指数， 一定程度上表明 转化指数同该价格区间转化效果成正比

# 使用步骤
1. 登录淘商机官网后提取csrf和cookie值
2. 打开config.json 将******内容替换为对应的csrf和cookie值即可
3. 运行 spider_taoshangji.py
4. 输入关键词等待数据载入保存即可。、

# 示例图片
![Image text](https://github.com/tuot1/taoshangji/blob/main/%E6%95%88%E6%9E%9C%E5%B1%95%E7%A4%BA_%E5%A9%9A%E7%A4%BC%E8%A1%A3%E6%9E%B6.png)


#### 请勿商用，请勿商用，请勿商用，
#### 数据使用，分享等请遵守淘商机官网约定协议
