# ptcg-card
> 下载ptcg繁中卡图，同时生成可打印的word文档
## 使用说明
### 环境安装
1. 到微软商店搜索python，点击安装即可
2. 下载本项目，解压，进入到目录下，在目录的地址栏输入cmd，弹出黑窗口，输入下面命令
   ```bash
   pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/
   ```
### 生成word
- 先到ptcg[香港官网](https://asia.pokemon-card.com/hk/card-search/list/)获取到需要打印的卡的编号
- 然后将编号填入到card.txt文件中，编号后面空格隔开，可以再输入一个数字，标识要插入这张卡的次数，也可以不填，默认就插入4次
- 最后双击启动等待就可以了
- 下载的卡图在tmp文件夹下，word文档会新建一个，名字是当前时间戳