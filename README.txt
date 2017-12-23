使用方法：
1.维护user.txt
开发名字，开发所属组（ios，Android，server，FE）

2.将bugfree上的bug导出excel
导出的excel全部放入脚本统计

3.点击parsebugfree.exe执行程序


注意点：
1.导出的bug必须完全符合bugfree导出的excel格式
2.user.txt的格式也必须完全符合规定格式
3.冒烟和接口测试用例是按照，title中是否包含关键字“冒烟”或“接口”来进行统计的



版本升级：

第一版：支持常规等各种统计

第二版：增加了以下4个需求
1.BUG数统计到人
2.每日创建BUG数、解决BUG数，统计到四端
3.需求变更、不是BUG等拆分到不同的sheet页
4.统计重开BUG次数
