# 《烽火例维表格生成工具》使用说明



## 第一部分：软件简介

该软件针对烽火网管系统导出的三个表格：oa.xlsx，ocp.xlsx，osc.xlsx和用户可编辑的模板表格ops_template.xlsx 与 och_template.xlsx，实现复杂数据的自动检索、计算和归纳，生成三个例维所需的分表格（genNewOA.xls、genNewOCP.xls和genNewOSC.xls）。让用户可从繁琐的数据归档工作中解脱出来，方便查看光缆线路具体情况，提高例维的效率。

该软件的开发环境为：Python 3.7.8 + Windows 10 专业版 x64 + PyCharm Community 2020.2.3（Build #PC-202.7660.27, built on October 6, 2020，Runtime version: 11.0.8+10-b944.34 amd64）

> 该使用说明采用 Typora 基于 Markdown 语言编写，需要编辑请**事先安装 Typora 编辑器**，再使用 Typora 打开 README.md 文件，方可得到优美的格式渲染。



## 第二部分：软件界面介绍

本软件仅有一个窗口作为主界面，窗口上方显示文本，为软件简介和作者信息。另有一个命令提示符，进行程序运行的提示输出。主界面下设五个功能按钮，具体对应功能如下：

- **Generate OMS Sheet using oa.xlsx：**利用 oa.xlsx 生成 “烽火波分OMS检查（每月）” 表格
- **Generate OCH Sheet using ocp.xlsx：**利用 ocp.xlsx 生成 “烽火波分OCH光功率检查（每月）” 表格
- **Generate OPS Sheet using osc.xlsx：**利用 osc.xlsx 生成 “烽火波分OPS检查（每月）” 表格
- **Generate! ：**一次性生成三个表格（不建议使用，运行速度5分钟，可能因电脑性能原因假死）
- **Exit：**退出本软件

![image-20210809225741880](C:\Users\Marco\AppData\Roaming\Typora\typora-user-images\image-20210809225741880.png)



## 第三部分：软件功能函数介绍

本软件共计```genNewOA()```、```genNewNewOCP()```、```genNewNewOSC()```三个主要函数，分别对应三个表格的生成。```wholeProcess()```函数对应实现 “Generate!” 按钮对应的一键生成功能。



#### genNewOA()

**功能简介：**利用 oa.xlsx 生成 genNewOA.xlsx，即为 “烽火波分OMS检查（每月）” 表格。如果 VOA 小于 3dB，则自动在 ”处理建议“ 中输出 ”注意！“ 的提醒。并且在 genNewOA.xlsx 的第二个子表中，会将所有需要注意的线路（即 VOA 小于 3dB）单独列出。

**原理介绍：**

首先利用 pandas 读入 oa.xlsx 原始表格作为原始数据输入，同时利用 xlwt 创建一个新的 workbook 对象作为所生成表格的基础。

利用正则匹配，将原始表中所需要的信息合并提取出来，并去除其中的重复部分，保存进新定义的名为 ```oaDirectSlotMatch``` 的一个list中。

> 【例: 将 “惠州:惠州-惠东烽火环WDM96*10G02:25-惠州巽寮:25-08-惠州巽寮OTM-碧甲方向” 和 “OA-1821（发碧甲）[1-2C]::IN/OTS-1” 转换为：“25-08-惠州巽寮OTM-碧甲方向:OA-1821（发碧甲）:IN/OTS-1”】

利用 xlwt 写入表格标题，并写入刚刚整理好的 OA 方向/槽位信息。

通过检索，逐个搜寻每个 OA 方向/槽位对应的 IOP、OOP 和 VOA_ATT 数据，并利用 xlwt 写入表格中。

保存第一个版本的表格，再用 pandas 读入，利用 pandas 的 ```sort_values``` 函数实现对 OA 方向/槽位信息进行排序，即完成对 genNewOA.xlsx 表格的生成。

由于在实际维护工作中，当其它字段相同时（尤其是收光位置相同），一条线路的方向/槽位名中，同时出现了 PA 与 OA 的时候，需要放弃 OA 对应的那一行。观察排序后的表格发现，这种情况往往是紧贴着的，因此，设计了一个遍历，如果方向/槽位名出现了 OA ，且为 OTM 线路，那么，自动查找下一行是否为 PA ，如果是 PA 且收光位置相同，则会删除掉 OA 的对应行。

如果该功能函数无法检测到每个 OA 方向/槽位对应的 IOP、OOP 和 VOA_ATT 数据，则会在对应的单元格**留空**。



#### genNewNewOCP()

**功能简介：**利用 ocp.xlsx 生成 genNewOCP.xls，即为 “烽火波分OCH光功率检查（每月）” 表格。

**原理介绍：**

首先利用 pandas 读入 ocp.xlsx 原始表格作为原始数据输入，再读入 och_template.xlsx 作为写入的模板表，作为所生成表格的基础。

~~利用正则匹配，将原始表中所需要的信息合并提取出来，并去除其中的重复部分，保存进新定义的名为 ```ocpDirectSlotMatch``` 的一个list中。【例如，将 “惠州:惠州-本地网烽火环WDM96*10G04:24-惠州江南:24-03-惠州江南OTM-江北方向” 和 “OCP-λ50[1-1B]::TRXB-1/OCH-1” 转换为：“24-03-惠州江南OTM-江北方向:OCP-λ50:TRXB-1/OCH-1”】~~

> 注：抛弃正则匹配并去重策略改为采用模板表的原因是，每次导出的网管数据中，总会缺一两条线路，但后面的差异计算是两两进行计算，那如果缺了一个，那后面就全部乱套了。得出的差异数值也不是所需要的差异数值。

通过遍历检索，逐个搜寻每个 OCP 方向/板卡/端口 的IOP数据，写入到 och_template.xlsx 对应的 DataFrame 的 “输入光功率” 数据栏中，如果是出现 “收无光” 的情况，则写入 “收无光” ；若有具体数值，则写入浮点数数值（不带单位，便于后续计算）。

利用 pandas 的 ```sort_values``` 函数实现对OCP方向/板卡/端口信息进行排序，并保存为 genNewOCP.xls 。

排序后，用 xlrd 读入新生成的 genNewOCP.xls ，再利用 ```xlutils.copy``` 将 xlrd 对象转为 xlwt 对象，开始合并单元格并进行差异计算。如果一组互为主备的AB线路中，任意一个出现了 “收无光” 的情况，或者任一条线路无读数，则会自动在后方的差异计算的合并单元格中写入 “无法计算” ；若均为可计算的浮点数，则会自动生成AB线路 IOP 的差异的绝对值。

~~在该表格的生成过程中，由于 pandas 的排序无法自定义，因此会出现【24-01-惠州江北OTM-平山方向:OCP-λ3.6:TRXA-1/OCH-1】和【24-01-惠州江北OTM-平山方向:OCP-λ3.6:TRXA-2/OCH-1】排列在一起的现象，然而我们需要的情况是 A-1 对应 B-1；A-2 对应 B-2 的情况，因此，要把出现这种情况的列标注出来，方便人工进行核查。利用正则匹配提取A-1、B-2类型信息，然后再判断是否连续出现两个A/B，再利用 xlwt 的带格式写入，进行标注。~~

> 注：由于采用模板表的方法，因此上方画有删除线的标注策略也不使用了。

即完成对 genNewOCP.xlsx 表格的生成。

**经过修改后，由于模板表本身含有1300余行待写入的数据格，每一行都需要遍历一次数据源 DataFrame 进行搜寻，即使在刚执行函数时，会对函数进行化简，但整个函数的循环次数达到了近500万次，经过多次测试，持续时间大约在170-180秒（三分钟左右），为了避免用户误认为程序假死，在提示窗口中，会有执行进度。但依然建议执行顺序：先读取 OA 和 OSC ，最后再读取 OCP 。**



#### genNewNewOSC()

**功能简介：**利用 osc.xlsx 和 ops_template.xlsx 生成 genNewOSC.xlsx，即为 “烽火波分OPS检查（每月）” 表格，同时还会在同一个表格文件中生成一个名为 “烽火波分OPS警告” 的 Sheet ，方便看到需要维护的线路情况。

**原理介绍：**

> 注：函数名中有两个New的原因是，先前用另一套逻辑写了一个 ```genNewOSC()``` 的函数，改用逻辑后弃用。
>

首先利用 pandas 读入 osc.xlsx 原始表格作为原始数据输入，保存进 ```osc_Ori``` 变量中；再利用 pandas 读入 ops_template.xlsx 作为模板表格，保存进 ```ops_template``` 变量中。

经过观察可发现，在给定的目标表格 “烽火OTN.xlsx” 中，第二列带有 “GEx”（x为数字）的，在总表中均无采集需求。因此，为了化简原始数据，减少遍历次数，提高运行效率，利用 pandas 提供的 drop 函数对无用的行进行删除，并重置 DataFrame 'osc_Ori' 的索引，避免出现无法遍历的情况。

由于 Pandas 无法处理合并单元格，会在合并单元格的非首行处写入大量 nan 值，因此使用 ```ffill()``` 函数，取前值进行填充，使合并单元格能正常分配到对应的每一行中。

**在模板表格 ops_template.xlsx 中，大致可分为三类结点设备名称：**

- 第一类（OTM带方向类）：25-08-惠州巽寮OTM-碧甲方向 （代码中其计数器称为 ```countOTMFX``` ）
- 第二类（OTM类）：79-4-惠州柏塘OTM 或 78-6-惠州泰美OA （代码中其计数器称为 ```countOTM```）
- 第三类（ROADM类）：83-39-惠州潼湖-ROADM(OA) 或 83-39-惠州潼湖-ROADM （代码中称其计数器为 ```countROADM```）

由于三类结点设备名需要用不同的方法进行数据采集，所以数据采集和输入的流程亦分为了三类。为了能顺利制表，需要得到三类情况中各类的数目。需要注意的是，取得了前述三类的数目后，由于 OTM 类和 ROADM 类排在第二和第三，所以 OTM 类的计数器应当加上 OTM 带方向类的数目，得到其在列表中的位置；而 ROADM 类则应当加上 OTM 带方向类与 OTM 类的数目，得到其在列表中的位置。

得到三类结点设备名的数目后，就可以进行分类数据采集和写入。利用检索功能即可实现对数据的搜寻，这里采用了遍历 DataFrame 的方法进行搜寻。首先基于 “A结点设备名称” 进行搜寻，利用方向和其它元素进行定位，找到数据后写入 ```ops_template``` 的A结点衰耗列中。同理，对B结点衰耗进行写入，完成结点数据的写入。

对于各类的结点数据的搜索和定位方法，由于细节太多，将会专门独立一部分进行讲解，本部分不做赘述。

完成三类数据的写入后，利用 Pandas 将 ops_template 写入表格中并保存。再利用 xlrd 读取该表，利用```xlutils.copy``` 将 xlrd 对象转为 xlwt 对象，进行下一步处理。

对照 ops_template.xlsx 表格，可以知道在程序端，可对 “实际与理论衰耗差值” 、 “双纤衰耗差值” 、 “是否检修（K列判决）” 、 “是否检修（K&J列判决）” 依照原始数据，进行计算和写入。在 xlwt 环境下，可以顺利使用合并单元格功能进行写入，这有助于恢复被 pandas 读取后所破坏的合并单元格。



## 第四部分 软件的数据检索和定位策略

#### genNewOA()

**需求：**需要匹配出对应的信息，并且需要读取和写入该线路的 IOP、OOP 和 VOA_ATT 的数据。

**策略：**利用正则匹配，在 oa.xlsx 中的B列和C列，匹配出所需要的信息后，创建一个新的List称为 ```oaDirectSlotMatch``` ，并将其不重复地写入新表；让 ocp.xlsx 中的每一行的B列和C列与 ```oaDirectSlotMatch``` 进行匹配，如果匹配，再检索G列，如果G列的值为 IOP / OOP / VOA_ATT ，则将 oa. xlsx 的H列数据部分写入新表对应的 IOP / OOP / VOA_ATT 格中。



#### genNewNewOCP()

**需求：**匹配出对应的互为保护的线路信息，采集对应的输入光功率（IOP）值，并计算互为保护的线路的输入光功率差异值。

**策略：**读取模板表 och_template.xlsx ，利用正则匹配将模板表中的 方向/槽位 名进行分割

> 例：24-01-惠州江北OTM-平山方向:OCP_λ86:TRXA-1/OCH-1 将会被分割为：
>
> ['24-01-惠州江北OTM-平山方向', 'OCP_λ86', 'TRXA-1/OCH-1']，含有三个元素的数组

如果第一个元素，即 “24-01-惠州江北OTM-平山方向” 存在  ocp.xlsx 中的B列中，且第二个元素 ‘’OCP_λ86" 与第三个元素 'TRXA-1/OCH-1' 存在于 ocp.xlsx C列中，且又是IOP数值的，则认为完成匹配，将对应的数值填入模板表 “输入光功率” 一栏中。

因此，有时候无法找到数据的话，多数是因为模板表中的内容和网管系统导出的数据表中的内容无法对应。

> 例：24-01-惠州江北1-ROADM:OCP_λ86:TRXA-1/OCH-1 的 数据发生了缺失，是因为，在网管系统导出的数据表中，系统给的标题实际上为：“24-01-惠州江北-ROADM:OCP_λ86:TRXA-1/OCH-1” ，**和模板表对比，少了一个 “1” **，因此这个时候也人工对模板表进行修改，将 “1” 去掉，就可以正常搜寻到数据了。

**注意：**在生成的 genNewOCP.xls 表中，对有些 IOP数据缺失的 方向/槽位 ，程序会自动选取IOP_MIN进行补全。



#### genNewNewOSC()

**需求：**读取A和B结点设备信息，找到对应的输出光功率和输入光功率并填入模板表格中。再进行差异计算等后续处理。

- **第一类（OTM带方向类）：**
  - **策略**：每路过一条ops_template的数据，就会遍历一次osc_Ori原始表，如果ops_template中的A结点设备名称与原始表的B列中的名称能对应，且C列说明类型为 “OSC_W” ，且G列为'OOP_MIN'，则将H列数据写入左侧的 “输出光功率” 中；若B结点设备名称与原始表的E列中的名称能对应，C列说明类型为 “OSC_W” ，G列为 "IOP_MIN"，则写入右侧的 “输入光功率” 中（牵涉到第二个for循环）。
- **第二类（OTM类）：**
  - **策略**：观察 “烽火OTN.xlsx” 表格可发现，这里的值不仅需要取决于A结点设备名，还需要取B结点设备中的地名作为方向；同时，是用 OSC_W 还是用 OSC_E 也是按照传输方向前的标识符究竟是E还是W来决定的。因此，对于第二类的 “输出光功率” ，首先利用A结点设备名称来定位发送结点，然后提取出B结点设备名称中的地名，与原始数据表格中的C、D列进行检索。如果B地名前面加上字符“E”有匹配结果，则使用OSC_E的OOP_MIN，若B地名加上字符“W”有匹配结果，则使用OSC_W的OOP_MIN。对于输入光功率也只是反过来。
  - 但是，在开发过程中，发现有些特殊的地方，例如阿婆角OTM，目的地的信息是放在D列的，又发现，在D列放置地名的线路，不会同时出现一地有OSC_W 与 OSC_E 的情况。因此，在代码中再加入了一个 else if 条件，如果目的地名出现在了D列，则直接采用其 OOP_MIN 或 IOP_MIN 数据。
- **第三类（ROADM类）：**
  - **策略**：ROADM类中，绝大部分都是采用OSC_W的数据，因此在代码中就利用OSC_W的数据进行检索了。原理与OTM类差不多，利用A结点设备的名称进行第一步检索，然后提取出B结点设备的目的地名，在后面加上“方向”字符串，如果在C列有匹配，且G列为'OOP_MIN'，则写入 “输入光功率” 中。
  - 对于一些ROADM(OA)线路，目的地信息也是在D列的，因此，在代码中，如果在C列找不到匹配方向，就去D列找，如果能够成功匹配，则利用该条数据填入到对应的功率框中。



## 第五部分 使用注意事项和潜在Bug

- 该软件仅可在 64 位的 Windows 操作系统下运行，若需在其它系统或 32 位系统下运行，或者对代码有修改的，需要在对应环境下重新打包编译，或配置好 Python 环境后直接在Python IDE中运行代码。

- oa.xlsx、osc.xlsx、ocp.xlsx和 ops_template.xlsx 必须与软件本身在同一个目录下，必须保持表格文件名一致，软件所生成的三个表格也会放在与软件所在位置相同的目录下。（如下图所示）
  - ![image-20210809230348168](C:\Users\Marco\AppData\Roaming\Typora\typora-user-images\image-20210809230348168.png)
  
- 在使用本软件前，请在每个初始数据表格第一行加入 0-9 的编号，方便软件作为识别位依据（如下图所示）。
  - ![image-20210809230216775](C:\Users\Marco\AppData\Roaming\Typora\typora-user-images\image-20210809230216775.png)
  
- 在生成的 genNewOA.xls 表中，对于有一些 IOP / OOP / VOA_ATT 数据缺失的 方向/槽位 ，是没有办法写入数据的，程序会自动留空，需要人工选取IOP_MIN / OOP_MIN等进行手动补全。

- ~~在生成的 genNewOCP.xls 表中，出现了较多 A-1 与 A-2 连续排列的情况，但这个问题无法调整，是由于 Pandas 不支持自定义排序字段造成的。~~
  （这个问题通过使用模板表解决了）

- 在生成的 genNewOSC.xls 表中，网元名无法实现自动识别合并单元格，需要人工归纳网元名。

- 填写 ops_template.xlsx 模板表格时，要确保理论每公里衰耗值、理论光路长度、理论衰耗都有数据，避免出现意料之外的问题。

- **【一个无法修正的Bug】**
  - 83-36-惠州平安-ROADM(OA)在原始表格中，需要采集杨村的OSC_E方向，软件无法识别采集（164、165）
    83-34-惠州铁岗-ROADM(OA)在原始表格中，需要采集永汉的OSC_E方向，软件无法识别采集（174、175）
    83-45-惠州新麻榨-ROADM(OA)在原始表格中，需要采集横河的OSC_E方向，软件无法识别采集（180、181）
    83-44-惠州横河-ROADM(OA)在原始表格中，需要采集湖镇的OSC_E方向，软件无法识别采集（182、183）
    83-39-惠州潼湖-ROADM(OA)在原始表格中，需要采集陈江的OSC_E方向，软件无法识别采集（194、195）
  - 这个 bug 应当无法修正，因为，例如83-01-惠州江北-ROADM线路中，往江南3方向是既存在 OSC_E ，也存在 OSC_W 的，但表格绝大多采用 OSC_W ，因此默认把 OSC_E 抛弃了
  
- 使用“一键生成（Generate!）”功能的时候，每完成一个表会弹出一个处理结果的提示框，需要点击“好的”进行确认，才能执行下一个表的生成。

- 每个表格生成后都会通过消息框提醒生成所花费时间，一般而言，第一个表大概四秒钟左右生产，第二个表大约3分钟左右产生，第三个表则需要大概三十秒左右，共计4~5分钟左右。（生成时间取决于电脑配置）

- 推荐使用前三个分步骤按键一步步执行程序，避免程序假死。

- 程序运行时，可能鼠标指针会变为忙碌状态，程序有可能假死，但无需理会，大约一分钟后就能得到三个表格。

- 日后新增光缆时，对 ops_template.xlsx 模板表格进行增添时，需要按照第三部分所介绍的 “三大类” 分类情况，按类型，在类型的区域内进行插入，不能直接在全表最末尾进行追加。

- > **该软件仅为一个自动化辅助工具，未经大量数据用例的稳定性测试，由于现实中网管系统生成数据情况众多且复杂，代码的复用性可能会下降，因此，在发现数据异常的情况下，必须进行人工修正确认，避免出现维护事故。**



## 第六部分 软件依赖库说明

```python
# -*- coding:utf-8 -*-
__author__ = "Marco Cheung"
import re, sys, time
import pandas
import xlrd, xlwt
from PySide2 import QtCore, QtWidgets
from xlutils.copy import copy
from pyautogui import alert
```
