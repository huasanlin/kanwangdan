# 看望单-by 林三化
**看望单自动生成程序

第一步是修改两个文件：
村.csv 或村.xslx；
看望人员.csv 或 看望人员.xslx。
把里面的人名，村名，改成本教会的。保存。
至于通讯方式，基于隐私，我没有做进去，把文件搞大不利于节省油墨纸张。哈哈。建议搞个统一的通讯录，熟悉了都是在自己的手机通讯录里直接搜索。
村后面的领访人员有几个填写几个，其他删除。
多余的行删除，样本的行不足，也可以参照上面的行格式自行添加。
看望人员后面的数字是下季度愿意出访的次数，所有人的总次数加起来应该大于一个季度，13次 x 村数量。每次有人员调动就修改这两个，否者以后无需改动。
假如你使用excel，修改的是xslx文件，csv没修改，那就删掉。只保留xslx。程序是优先从csv取数据，不删，数据是遵从csv。

第二步，运行 k1.py
如果上面的格式没改错的话，会读取上面的两个数据，生成一个 总表.xslx
打开，检查是否排的对。检查项目包括日期是否对，默认是下季度的每周六。各个村名，人名。
假设对自动生成不满意，需要对人员安排进行手动调整，直接修改这个表。改完后保存即可。
假设无法生成总表.xslx，多半原因是第一步修改的错误，重新下载样本，进行第一步的修改。尤其是csv格式，用诸如notpad++之类的编辑器打开修改，比windows默认txt打开修改要靠谱。

第三步，运行k2.py
此时会读取总表.xslx，
生成每个参与看望的人员的日程表。在新诞生的那个“个人日程表”文件夹里面。
每人有两种，一种是excel格式，供打印，送到相应同工的手里。
另一种是日程格式，.ics，直接发到同工的微信里面，打开，导入手机日程表后，有看望任务的那周周六早上7点，手机会自动提示。节约油墨纸张。好吧，我太环保了。

ok。
关于美观啊，哪些需要优化的地方。根据使用者反馈（在下面留言或加微信：109868901），再进行修改。我现在都不知道这个看望单自动生成程序，是否可以用，万一没啥用，那我过度优化细节，岂非自作多情。

初步的几种设想：

1做个在线网站，上传两个文件，自动生成结果：没钱，还是长期没工作的那种，还是长期没老婆的那种，555，不说了，越说越惨了。，这方案涉及隐私问题，自己用电脑安装python环境，自动生成吧，也就是一开始难一点。

2做个手机app实现相同功能：？？？

3做了个windows可运行的exe文件。[在右面 Releases 那里下载]，下载后解压，看使用说明。-这压缩包太大了，还是推荐自建一个python环境，下载源码运行。

随意下载测试传播使用，开源且完全免费。
