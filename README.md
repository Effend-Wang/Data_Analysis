# Data_Analysis
Data Analysis Program v0.3
作者: Effend Wang
版本信息: v0.3
最后编辑时间: 2019/8/19 Monday

v0.3更新内容：
1、代码框架化、模块化以易于编写及修改；
2、修复STD算法，消除了存在的计算误差；
3、CPK输出表格新增数据信息；
4、log信息中增加了module名称以易于追踪错误。

使用说明：
1、该程序目前仅用于测试，不代表最终版本，后续功能根据情况逐步添加。
2、该程序需要运行在Vista、Win7及以上系统，无法在WinXP上运行。
3、在本版本中仅能使用CPK Analysis功能，无法使用Test Coverage Analysis。
4、若在运行程序时有疑问或有希望添加的功能，可联系Effend Wang。

使用方法：
1、准备好预先处理过了的.xlsx或.xls文件，文件形式可参考example文件夹下的示例文件。
2、双击该程序后会进入cmd命令行界面，选择工作模式（使用CPK Analysis输入a。v0.2中仅能使用CPK Analysis，无法使用Test Coverage Analysis），然后将需要处理的excel表格文件拖入命令行界面并回车，程序会输出文件中存在的所有的sheet，输入需要处理的sheet名，接着分别输入需要处理的数据的起始行和起始列，如A1则为第一行，第一列。输入完成回车后会自动进行计算，计算完毕后程序自动退出。
3、在程序路径下，程序运行后将生成2个文件，一个是RunSteps.log，其中存放着程序运行时产生的日志信息；一个是Analysis_Result.xls，其中存放着分析结果数据。（注意：日志文件非覆盖写入，除非删除该文件，否则会保留之前的运行记录；分析结果数据表格为覆盖写入，因此若需保存结果，请将表格文件移出程序路径后再运行程序）
4、在日志文件中，包含了每个处理过的参数的信息。当CPK=NA或者CPK<1.33时，日志会以ERROR记录；当所计算的参数Limit有不存在或者Standard Deviation=0时，日志会以WARNING记录。因此，可通过搜索ERROR和WARNING查看存在问题的项。

程序异常：
1、若程序运行中途异常退出，在日志中可看到A Programe ERROR Occured异常信息。